'''
Created on 23 Jul 2013

@author: ACE
'''

#/bin/env python
#
# $Header:  $
#
# MEA Lightweight Integration Environment
#
# Entity Handler for PALLET ITEMS SPECIFICTAIONS
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

class InterfaceFromXLS:
    
    
    def __init__(self, worksheet, i, c, logger):
        self.itemnum = worksheet.cell(i, 0).value
        self.item_desc = worksheet.cell(i, 1).value
        self.item_transid = worksheet.cell(i, 6).value
        
    
    def is_valid(self, c, logger):
        
        #Check if description returns a number
        if self.item_desc == "":
            logger.error("Blank Description on item %s is illegal" % (self.itemnum))
        return False
        
        
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

def get_queued_batch_ct(c, logger):

    # Get the number of batches for THIS INTERFACE that 
    # have yet to be processed by MEA
    q_get_ct = """
    SELECT COUNT(*)
    FROM maximo.mxin_inter_trans 
    WHERE extsysname = 'EXTSYS1'
    AND ifacename = 'MXITEMSPECINTERFACE'"""

    time.sleep(5)       # t = seconds
    c.execute(q_get_ct)
    wait_ct = c.fetchone()
    logger.info("MEA has %s unprocessed batches in the queue" % (wait_ct[0]))

    return wait_ct[0]

def isp_handler(er):

    type_handler = {
        "XLS": spec_handler_xls_ora
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
    INSERT INTO maximo.mxtoolitem_iface (itemnum, description, itemsetid, prorate, rotating, 
                    outside, transid, transseq)
    VALUES(:itemnum, :item_desc, 'ITEMSET', '0', '0', 
                    '0', :transid, '1')"""
    
    total_ct = 0       # Total no. of items
    start_row = 1      # The first row in spreadsheet that contains data (starting at 0)
    processed_ItemAttr = [] #Values of item+attribute that have already been proceseed  

    # Iterate through all rows of data in input spreadsheet
    
    for this_row in range(start_row, wsh.nrows):
            
            ic = InterfaceFromXLS(wsh, this_row, curs_target_seq, er.logger)
            er.logger.info("Processing for item %s" % (ic.itemnum))
            
            if (ic.is_valid(curs_target_lookup, er.logger)):
                
                    er.logger.debug("Getting new transaction id for this load")
                    trans_id = get_next_txn_id(curs_target_seq, er.logger)
                    trans_seq = 1
                    er.logger.debug ("Item_Spec_Seq: %s" % (ic.itemnum + ic.propident))
                
            
                    curs_target_mea.execute(q_write_interface, itemnum=ic.itemnum,
                                            item_desc=ic.item_desc,                                                                #shortdesc=ic.shortdesc,
                                                                    #itemdesc=ic.longdesc,  #Get long description for Long Description Table
                                                                    #assetattrid=ic.propident,
                                                                    #assetsequence=asset_sequence, #"Get from classspec table, where classtructureid"
                                                                    #measureunitid= measure_unit, #"Get from classspec table, where classtructureid"  
                                                                    #classid=ic.classid,
                                                                    #alnvalue=valuealn, #Get values from if statement that determines if aln value or numm value
                                                                    #numvalue=valuenum, #Get values from if statement that determines if aln value or numm value
                                                                    #itemspecseq=itemspec_seq,
                                                                    #longdescseq=long_desc_seq,
                                                                    transid=trans_id, 
                                                                    transseq=trans_seq)
                    er.logger.debug("Wrote interface")
                    
                    
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
