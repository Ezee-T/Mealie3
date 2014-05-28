#/bin/env python
#
# $Header:  $
#
# MEA Lightweight Integration Environment
#
# 
#
# 
#
# $Log:  $
#
#
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

          
processed_ic = []  
class ReadItemSpecFromXLS:
    
    """ Class representing a set of Item Specifications to be loaded by MEA:SPE from a Spreadsheet"""
    
    def __init__(self, worksheet, i, c, logger):
        self.itemnum = worksheet.cell(i, 0).value
        self.clsfy = worksheet.cell(i, 1).value
        self.classname  = worksheet.cell(i, 2).value
        self.PropIdentifier = worksheet.cell(i, 9).value
        self.PropName = worksheet.cell(i, 10).value
        
        logger.debug("Get itemid for item %s" % (self.itemnum))
        item_id = """
        SELECT itemid
        FROM maximo.item i
        WHERE i.itemnum = :itemnum"""
        c.execute(item_id, itemnum = self.itemnum)
        this_itemid = c.fetchone()
        logger.debug("Itemid for Item %s lookup: %s" % (self.itemnum, this_itemid))
        if (this_itemid == None):
            self.itemid = None
        else:
            self.itemid = this_itemid[0]
            
        logger.debug("Getting classstructureid for classificationid %s" % (self.clsfy))
        q_classid="""
            SELECT classstructureid, usewithitems
            FROM maximo.classstructure c
            WHERE c.classificationid = :classificationid"""
        c.execute(q_classid, classificationid=self.clsfy)
        this_classid = c.fetchone()
        logger.debug("classstructureid for item %s looked up: %s" % (self.itemnum, this_classid))
        if (this_classid == None):
            self.classstructureid = None
            self.usewithitems = 0
        else:
            self.classstructureid = this_classid[0]
            self.usewithitems = this_classid[1]
            
            self.attr_val = []
               
        
        self.errmsg = None
        self.severity = None
        
    def is_valid(self, c, logger):
        # Testing for exceptions
        
        logger.debug("Checking whether the Classification %s and Attribute %s are valid for item %s" % (self.clsfy, self.PropIdentifier, self.itemnum))
        
        if self.clsfy == "":
            logger.error("Blank Classification is illegal")
            return False
        
        if self.PropIdentifier == "":
            logger.error("Blank Attribute is illegal")
            return False
        
        if self.itemnum == "":
            logger.error("Blank item number is illegal")
            return False
        
        logger.debug("Checking whether item %s exist" % (self.itemnum))
        
        q_itemnum_exist = """
            SELECT itemnum
            FROm maximo.item i
            WHERE i.itemnum = :itemnum"""
        c.execute(q_itemnum_exist, itemnum=self.itemnum)
        this_item = c.fetchone()
        logger.debug("Item %s looked up: %s" % (self.itemnum, this_item))
        if (this_item == None):
            logger.error("Item %s does not exist" % (self.itemnum))
            print("ERROR: Item %s does not exist" % (self.itemnum))    
        
        if (not this_item):
            return False
        
        logger.debug("Checking classification %s is legal" % (self.clsfy))

        q_exists_class="""
            SELECT classificationid
            FROM maximo.hul_mealie_classification c
            WHERE c.classificationid = :classificationid"""
        c.execute(q_exists_class, classificationid=self.clsfy)
        this_class = c.fetchone()
        logger.debug("Legal Classification %s looked up: %s" % (self.clsfy, this_class))

        if (not this_class):
            return False
        
        #logger.debug("Checking classification %s is Use With items" % (self.clsfy))
        #if (self.usewithitems != 1):
         #   logger.error("Classification %s is not set to Use With Assets" % (self.clsfy))
          #  print("ERROR: Classification %s is not set to Use With Assets" % (self.clsfy))
           # return False
        #
        return True

########################################################################################
#
#     Functions
#
########################################################################################

def get_next_txn_id(c, logger):
    
    #Get next available sequence number for batch 
    
    q_get_asp_seq = """
    SELECT maximo.itemspecseq.NEXTVAL
    FROm DUAL"""
    
    c.execute(q_get_asp_seq)
    mealie_seq = c.fetchone()
    logger.debug("Obtained MEA seq Id %s" % (mealie_seq[0]))
    
    return mealie_seq[0]

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

def write_mea_queue(c, txid, logger):
    #Insert row into Queue
    
    q_write_queue = """
    INSERT INTO maximo.mxin_inter_trans (
                           extsysname,
                           ifacename,
                           action,
                           transid
                           )
    VALUES ('EXTSYS1', 'MXITEMSPECINTERFACE', 'AddChange', :trans_id)"""
    
    logger.debug("Writing MEA queue for transid %d" % txid)
    c.execute(q_write_queue, trans_id=txid)
    
    return True

def process_itemSpec(ic, c, logger):
    
    q_write_interface = """
    INSERT INTO maximo.mxitemspec_iface (
    itemnum,
    classstructureid,
    transid,
    transseq)
    VALUES (:itemnum,
           '0188598',
           :transid,
           :transseq)"""
    
    if ic.itemnum not in processed_ic:
    # Obtain new MEA transaction id for next asset 

        logger.debug("Getting new transaction id for this load")
        trans_id = get_next_txn_id(c, logger)

        trans_seq = 1          # Sequence within batch

        # Write the MEA interface table
        logger.debug("Processing interface record for itemnum=%s" % (ic.itemnum))
                
        logger.info("Writing interface: classstructureid=%s transid=%s transseq=%s" % (ic.classstructureid, trans_id, trans_seq))

        #logger.debug("commoditygroup=%s" % (ic.commoditygroup))
        #logger.debug("commodity=%s" % (ic.commodity))
        c.execute(q_write_interface, itemnum=ic.itemnum,
                                     classstructureid = ic.classstructureid,
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

def itemSpec_handler(er):
    
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

    # Initialise local variables for fetch loop

    total_ct = 0       # Total no. of items
    start_row = 1      # The first row in spreadsheet that contains data (starting at 0)

    # Iterate through all rows of data in input spreadsheet

    for this_row in range(start_row, wsh.nrows):

        ic = ReadItemSpecFromXLS(wsh, this_row, curs_target_mea, er.logger)
        er.logger.info("Processing classificationid %s for item %s" % (ic.clsfy, ic.itemnum))

        if (ic.is_valid(curs_target_lookup, er.logger)):
            
            # Is this a UNSPSC Class Code or Commodity Code, i.e Parent or Child?
            er.logger.debug("Commodity %s has an implied parent %s" % (ic.clsfy, get_unspsc_class(ic.clsfy)))
            
            if (ic.clsfy == get_unspsc_class(ic.clsfy)):
                
                # The commodity supplied is a parent so process the item as a Commodity Group

                er.logger.debug("Commodity %s is a UNSPSC Class Code" % (ic.clsfy))
                ic.commoditygroup = ic.commodity
                ic.commodity = None
                
            else:
                
                # The commodity supplied is a child so process the item as a Commodity Code
                # deriving its Commodity Group
                er.logger.debug("Commodity %s is a UNSPSC Commodity Code" % (ic.clsfy))
                ic.commoditygroup = get_unspsc_class(ic.clsfy)

            total_ct += process_itemSpec(ic, curs_target_mea, er.logger)
                        
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

