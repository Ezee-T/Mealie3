from pandas.core.common import notnull

+-'''
Created on 25 Apr 2013

Update Work Order status's for COMPLETION

@author: ACE
'''


#/bin/env python
#
# $Header: /usr/local/cvsroot/mealie/entities/itemcomm.py,v 1.1 2012/02/10 14:07:50 djw Exp $
#
# MEA Lightweight Integration Environment
#
# Entity Handler for ITEM SPECIFICATIONS
#
# Issue Track 
#
#
#

########################################################################################
#
#     Import required modules
#
########################################################################################

import cx_Oracle
import xlrd
import sys
import time
import Tkinter

from common import get_unspsc_class
#from AttributeSpec import curs_target_lookup

# Values of commodity that have been processed. Only first occurrence in file is processed
processed_ic = []

########################################################################################
#
#     Class definitions
#
########################################################################################
class InterfaceableWorkOrderFromXls:

    """Class representing each work order to be loaded by MEA from a spreadsheet"""

    def __init__(self, worksheet, i, c, logger):
        
        # Values are taken from the Work Order Change Template
        try:
            self.wonum = int(worksheet.cell(i, 0).value)
        except ValueError:
            logger.debug("Wonum %s has non-numeric work order number \"%s\", setting to None" % (self.wonum, worksheet.cell(i, 0).value))
            self.level = None
            
            
        self.descr = worksheet.cell(i, 1).value
        self.worktype = worksheet.cell(i, 2).value
        self.wostatus = worksheet.cell(i, 3).value
        self.reportby = worksheet.cell(i, 4).value
        self.glacc = worksheet.cell(i, 5).value
        self.craft = worksheet.cell(i, 6).value
        self.feedback = worksheet.cell(i, 7).value
        self.changeby = worksheet.cell(i, 8).value
        self.labcode = worksheet.cell(i, 9).value
        self.labcraft = worksheet.cell(i, 10).value
        self.labpayrate = worksheet.cell(i, 12)
        self.labreghrs = worksheet.cell(i, 13).value
        self.enterby = worksheet.cell(i, 14).value
        self.enterdate = worksheet.cell(i, 15).value
        self.startdate = worksheet.cell(i, 16).value
        self.starttime = worksheet.cell(i, 17).value
        self.finishdate = worksheet.cell(i, 18).value
        self.finishtime = worksheet.cell(i, 19).value
        self.transid = worksheet.cell(i, 20).value
        self.transseq = worksheet.cell(i, 21).value

        self.errmsg = None
        self.severity = None

    def is_valid(self, c, logger):

        # Test for exceptions. An ERROR will cause the whole run to abort,
        # rolling back the transaction
        
        logger.debug("Checking wonum \"%s\" has no feedback or actuals" % (self.wonum))
        
        q_exists_feedback="""
            SELECT refwo
            FROM maximo.worklog l
            WHERE l.location = :wonum"""
        
        c.execute(q_exists_feedback, wonum=self.wonum)
        wonum_feedback = c.fetchone()
        logger.debug("Feedback for %s looked up: %s" % (self.wonum, wonum_feedback))
        
        if (not wonum_feedback):
            logger.debug("Work Order %s has no feedback, will check for Actuals information" % (self.wonum))
            
            logger.debug("Checking Actuals for Work Order: %s" % (self.wonum))
            
            q_exits_labtrans="""SELECT refwo
                FROM maximo.labtrans l
                WHERE l.location = :wonum"""
                
            c.execute(q_exits_labtrans, wonum=self.wonum)
            wonum_labtrans = c.fetchone()
            logger.debug("Feedback for %s looked up: %s" % (self.wonum, wonum_labtrans))
            
            if (not wonum_labtrans):
                logger.debug("Work Order %s has no labour Actuals" % (self.wonum))
                return True
            else:
                return False
        else:
            return False
            
        if self.wonum == "":
            logger.error("Blank work ornder number is illegal")
            return False

        if self.feedback == "":
            logger.error("Blank feedback is illegal")
            return False
        
        #if self.propvalue =="":
            #logger.error("Blank value for this attribute ______ is illegal")
            #return False
        
        return True
    
##########################################################################################
#
#     Functions
#=########################################################################################

def get_next_txn_id(c, logger):

    # Get the next available transaction id for a MEA batch
    q_get_id = """
    SELECT maximo.hul_mealie_isp_seq.NEXTVAL
    FROM DUAL"""

    c.execute(q_get_id)
    txn_id = c.fetchone()
    logger.debug("Obtained MEA Txn Id %s" % (txn_id[0]))

    return txn_id[0]

def get_next_worklog_seq(c, logger):
    
    # Get the next work log sequence number for each work order needing a work log
    q_get_worklogseq = """
    SELECT maximo.worklogseq.NEXTVAL
    FROM DUAL"""
      
    c.execute(q_get_worklogseq)
    work_log_seq = c.fetchone()
    logger.debug("Obtained new work log sequence number %d" % (work_log_seq[0]))
    
    return work_log_seq[0]

def get_next_LabTrans_seq(c, logger):
    
    #Get the next Labor Transaction Sequence number for each new work order needing Actuals information
    q_get_Labtrans = """
    SELECT maximo.labtransseq.NEXTVAL
    FROM DUAL"""
    
    c.execute(q_get_Labtrans)
    lab_trans_seq = c.fetchone()
    logger.debug("Obtained new labor transaction sequence number %d" % (lab_trans_seq[0]))
    
    return lab_trans_seq[0]

def get_nextWoStatus_seq(c, logger):
    
    # Get the next Work Order status change sequence number for each work order status change
    q_get_wostatus = """
    SELECT maximo.wostatusseq.NEXTVAL
    FROM DUAL"""
    
    c.execute(q_get_wostatus)
    wo_stat_seq = c.fetchone()
    logger.debug("Obtained new work order status change sequence number %d" % (wo_stat_seq[0]))
    
    return wo_stat_seq[0]

def get_queued_batch_ct(c, iface_name, logger):

    # Get the number of batches for THIS INTERFACE that 
    # have yet to be processed by MEA
    q_get_ct = """
    SELECT COUNT(*)
    FROM maximo.mxin_inter_trans 
    WHERE extsysname = 'EXTSYS1'
    AND ifacename = :InterfaceName"""

    time.sleep(5)       # t = seconds
    c.execute(q_get_ct, Interfacename=iface_name)
    wait_ct = c.fetchone()
    logger.info("MEA has %s unprocessed batches in the queue" % (wait_ct[0]))

    return wait_ct[0]

def write_mea_queue(c, txid, logger):
    # Insert row into MEA queue
        
    q_write_queue = """
    INSERT INTO maximo.mxin_inter_trans (
                           extsysname,
                           ifacename,
                           action,
                           transid
                           )
    VALUES ('EXTSYS1', :InterfaceName, 'AddChange', :trans_id)"""

    logger.debug("Writing MEA queue for transid %d" % txid)
    c.execute(q_write_queue, trans_id=txid)

    return True
   

def WoStatChange_handler(er):

    type_handler = {
                    "ORA": spec_handler_xls_ora
    }

    return type_handler.get(er.source_type)(er)   # Call appropriate handler for the source type



def spec_handler_xls_ora(er):

    # Prompt user for target password, if not in parameters
    # file. When used in Production NEVER put passwords in the parameter file.

    if (not er.target_pwd):
        er.target_pwd = str(raw_input("Enter password for target %s@%s: " % (er.target_user, er.target_db)))

    # Open source spreadsheet

    try:
        ifh = xlrd.open_workbook(er.source_xlsfile)
    except IOError:
        print("ERROR: Unable to open %s" % (er.source_xlsfile))
        er.logger.error("Unable to open %s" % (er.source_xlsfile))
        return False
    wsh = ifh.sheet_by_index(0)

    # Connect to MEA target database

    try:
        conn_target = cx_Oracle.connect("%s/%s@%s" % (er.target_user, er.target_pwd, er.target_db))
    except cx_Oracle.DatabaseError, exc:
        error, = exc.args
        print("ERROR: %s" % (error.message))
        er.logger.error("%s" % (error.message))
        return False
    
    # Open the necessary cursors on target database

    curs_target_mea = conn_target.cursor()
    curs_target_lookup = conn_target.cursor()
    curs_target_seq = conn_target.cursor()
    
    # Insert a row into MEA interface for this entity type 
    q_write_interface = """
    """
    
    
    #:itemspecseq, itemid, 'ITEM', 'DESCRIPTION', :itemdesc, :longdescseq, :transid, :transseq

    
    # Initialise local variables for fetch loop

    total_ct = 0       # Total no. of items
    start_row = 1      # The first row in spreadsheet that contains data (starting at 0)
    processed_ItemAttr = [] #Values of item+attribute that have already been proceseed  
    batch_queue = []   # Batch id queue for queue table

    # Iterate through all rows of data in input spreadsheet
    
    for this_row in range(start_row, wsh.nrows):

            ic = InterfaceableWorkOrderFromXls(wsh, this_row, curs_target_seq, er.logger)
            er.logger.info("Processing work order number %s" % (ic.wonum))

            if (ic.is_valid(curs_target_lookup, er.logger)):
            
                    er.logger.debug("Getting new transaction id for this load")
            
                    worklog_seq = get_next_worklog_seq(curs_target_seq, er.logger)
                    labtrans_seq = get_next_LabTrans_seq(curs_target_seq, er.logger)         
                    trans_id = get_next_txn_id(curs_target_seq, er.logger)
                    #er.logger.debug ("Item_Spec_Seq: %s" % (ic.itemnum + ic.propident))
        
                    
                    curs_target_mea.execute(q_write_interface, itemnum=ic.itemnum,
                                                                    shortdesc=ic.shortdesc, 
                                                                    #itemdesc=ic.longdesc,  #Get long description for Long Description Table
                                                                    assetattrid=ic.propident,
                                                                    #assetsequence=asset_sequence, #"Get from classspec table, where classtructureid"
                                                                    #measureunitid= measure_unit, #"Get from classspec table, where classtructureid"  
                                                                    classid=ic.classid,
                                                                    alnvalue=ic.propALNvalue, #Get values from if statement that determines if aln value or numm value
                                                                    numvalue=ic.propNUMvalue, #Get values from if statement that determines if aln value or numm value
                                                                    itemspecseq=itemspec_seq,
                                                                    #longdescseq=long_desc_seq,
                                                                    transid=ic.transid, 
                                                                    transseq=ic.transseq)
                    er.logger.debug("Wrote interface")
                    processed_ItemAttr.append(ic.itemnum+ic.propident)
                                            #total_ct += process_item(ic, curs_target_mea, er.logger)
                    total_ct += 1           
            else:
                er.logger.error("ERROR: Values for itemnum %s, attribute %s %s are invalid, aborting" % (ic.itemnum, ic.propident, ic.propname))
                print("ERROR: Values for itemnum %s, attribute %s %s are invalid, aborting" % (ic.itemnum, ic.propident, ic.propname))
                #conn_target.rollback()
                conn_target.close()         # Free the connection
                sys.exit(2)

    # (-) End of item processing loop

    # The transaction will not be committed if we are in 'test mode'
    # Otherwise we commit the whole transaction

    if (not er.test_mode):
        # Write the MEA queue
        er.logger.info("Writing MEA queue, stored txn ids are %s" % batch_queue)
        write_mea_queue(curs_target_mea, batch_queue, er.logger)
        er.logger.info("Wrote MEA queue, stored txn ids are %s" % batch_queue)
        er.logger.info("Committing changes to MEA target %s" % er.target_db)
        conn_target.commit()
        
    else:
        er.logger.info("Test Mode: Rolling back changes to %s" % er.target_db)
        #conn_target.rollback() 

    er.logger.info("Processed %s %s entities" % (total_ct, er.entity_type))

    # Wait until MEA has flushed the queue

    er.logger.info("Wait for MEA to flush its queue")
    queued_batches = get_queued_batch_ct(curs_target_mea, er.logger)
    while (queued_batches):
        er.logger.debug("MEA still has %s items in queue" % (queued_batches))
        queued_batches = get_queued_batch_ct(curs_target_mea, er.logger)

    if (not er.test_mode):
        # Tune this value according to your cron configuration
        er.logger.info("Waiting %d seconds for MEA to finish loading" % (er.mea_wait_interval))
        time.sleep(er.mea_wait_interval)

    print("Processed %s %s entities" % (total_ct, er.entity_type))
    print "See %s for full details of this run" % er.log_file

    # Close open cursors and disconnect from MEA target database

    curs_target_mea.close()
    curs_target_lookup.close()

    conn_target.close()

    return True


