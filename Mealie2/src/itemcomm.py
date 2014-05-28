#/bin/env python
#
# $Header: /usr/local/cvsroot/mealie/entities/itemcomm.py,v 1.1 2012/02/10 14:07:50 djw Exp $
#
# MEA Lightweight Integration Environment
#
# Entity Handler for ITEM COMMODITIES
#
# Issue Track 8112
#
# $Log: itemcomm.py,v $
# Revision 1.1  2012/02/10 14:07:50  djw
# 8112 Handling for ICG (Item Commodity)
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

from common import get_unspsc_class

# Values of commodity that have been processed. Only first occurrence in file is processed
processed_ic = []

########################################################################################
#
#     Class definitions
#
########################################################################################

class InterfaceableItemCommodityFromXls:

    """Class representing a Commodity Group or Commodity Code to be associated with an Item and 
    loaded by MEA from a spreadsheet"""

    def __init__(self, worksheet, i, c, logger):
        
        # Values are taken from the Item Classification Template

        # Duplicate values of the Commodity Code are skipped because a full Item Classification
        # Template would have many duplicate Properties. Thus in effect only the first one 
        # encountered is loaded to MEA

        self.itemnum = worksheet.cell(i, 0).value
        self.commodity = worksheet.cell(i, 3).value

        self.errmsg = None
        self.severity = None

    def is_valid(self, c, logger):

        # Test for exceptions. An ERROR will cause the whole run to abort,
        # rolling back the transaction

        logger.debug("Checking commodity %s for itemnum %s is valid" % (self.commodity, self.itemnum))

        if self.itemnum == "":
            logger.error("Blank itemnum is illegal")
            return False

        if self.commodity == "":
            logger.error("Blank commodity is illegal")
            return False
        
        return True

##########################################################################################
#
#     Functions
#=########################################################################################

def get_next_txn_id(c, logger):

    # Get the next available transaction id for a MEA batch
    q_get_id = """
    SELECT maximo.hul_mealie_icg_seq.NEXTVAL
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
    AND ifacename = 'MXITEMCOMMInterface'"""

    time.sleep(5)       # t = seconds
    c.execute(q_get_ct)
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
    VALUES ('EXTSYS1', 'MXITEMCOMMInterface', 'AddChange', :trans_id)"""

    logger.debug("Writing MEA queue for transid %d" % txid)
    c.execute(q_write_queue, trans_id=txid)

    return True

def process_item(ic, c, logger):
    
    # Insert a row into MEA interface for this entity type 
    q_write_interface = """
    INSERT INTO maximo.mxitemcomm_iface (
        itemnum,
        itemsetid,
        commoditygroup,
        commodity,
        transid,
        transseq)
    SELECT :itemnum,
           setid,
           :commoditygroup,
           :commodity,
           :transid,
           :transseq
    FROM maximo.sets WHERE settype = 'ITEM'"""

    # Have we already processed this Item?
    if ic.itemnum not in processed_ic:

        # Obtain new MEA transaction id for next asset 

        logger.debug("Getting new transaction id for this load")
        trans_id = get_next_txn_id(c, logger)

        trans_seq = 1          # Sequence within batch

        # Write the MEA interface table
        logger.debug("Processing interface record for itemnum=%s" % (ic.itemnum))
                
        logger.info("Writing interface: commodity=%s transid=%s transseq=%s" % (ic.commodity, trans_id, trans_seq))

        logger.debug("commoditygroup=%s" % (ic.commoditygroup))
        logger.debug("commodity=%s" % (ic.commodity))
        c.execute(q_write_interface, itemnum=ic.itemnum,
                                     commoditygroup=ic.commoditygroup,
                                     commodity=ic.commodity,
                                     transid=trans_id,
                                     transseq=trans_seq)
        logger.debug("Wrote interface")
        processed_ic.append(ic.itemnum)     # We processed it, remember it

        # Write the MEA queue
        ##write_mea_queue(c, trans_id, logger)
        logger.debug("Wrote MEA queue for transid %d" % trans_id)
        ct = 1
                
    else:
                
        logger.debug("Item %s was already processed, skipping" % (ic.itemnum))
        ct = 0
        
    return ct
    
    
def icg_handler(er):

    type_handler = {
        "XLS": icg_handler_xls_ora
    }

    return type_handler.get(er.source_type)(er)   # Call appropriate handler for the source type


def icg_handler_xls_ora(er):

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

    # Initialise local variables for fetch loop

    total_ct = 0       # Total no. of items
    start_row = 1      # The first row in spreadsheet that contains data (starting at 0)

    # Iterate through all rows of data in input spreadsheet

    for this_row in range(start_row, wsh.nrows):

        ic = InterfaceableItemCommodityFromXls(wsh, this_row, curs_target_mea, er.logger)
        er.logger.info("Processing commodity %s for item %s" % (ic.commodity, ic.itemnum))

        if (ic.is_valid(curs_target_lookup, er.logger)):
            
            # Is this a UNSPSC Class Cob de or Commodity Code, i.e Parent or Child?
            er.logger.debug("Commodity %s has an implied parent %s" % (ic.commodity, get_unspsc_class(ic.commodity)))
            
            if (ic.commodity == get_unspsc_class(ic.commodity)):
                
                # The commodity supplied is a parent so process the item as a Commodity Group

                er.logger.debug("Commodity %s is a UNSPSC Class Code" % (ic.commodity))
                ic.commoditygroup = ic.commodity
                ic.commodity = None
                
            else:
                
                # The commodity supplied is a child so process the item as a Commodity Code
                # deriving its Commodity Group
                er.logger.debug("Commodity %s is a UNSPSC Commodity Code" % (ic.commodity))
                ic.commoditygroup = get_unspsc_class(ic.commodity)

            total_ct += process_item(ic, curs_target_mea, er.logger)
                        
        else:

            er.logger.error("ERROR: Values for itemnum %s are invalid, aborting" % (ic.itemnum))
            print("ERROR: Values for itemnum %s are invalid, aborting" % (ic.itemnum))
            conn_target.rollback()
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
        conn_target.rollback() 

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
