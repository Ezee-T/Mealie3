#/bin/env python
#
# $Header: /usr/local/cvsroot/mealie/entities/measureunits.py,v 1.1 2012/02/07 09:02:37 djw Exp $
#
# MEA Lightweight Integration Environment
#
# Entity Handler for MEASURE UNITS
#
# Issue Track 8112
#
# $Log: measureunits.py,v $
# Revision 1.1  2012/02/07 09:02:37  djw
# 8112 New UOM entity handler. Tested on maxsit using dummy data. See C:\SRCCVS\maximo6\8112\hul_test_UOM.stg
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

########################################################################################
#
#     Class definitions
#
########################################################################################

class InterfaceableMeasureunitFromXls:

    """Class representing a Unit of Measure to be loaded by MEA from a spreadsheet"""

    def __init__(self, worksheet, i, c, logger):
        
        # Values are taken from the Classification Template

        # Blank lines are skipped because not every Property need have a Measure Unit
        
        # Duplicate values of the Measure Unit id are skipped because a full Classification
        # Template would have many duplicate Properties and even different Properties may have
        # the same Measure Unit. Thus in effect only the first one encountered is loaded to MEA

        self.measureunitid = worksheet.cell(i, 7).value
        self.description = worksheet.cell(i, 8).value
        self.abbreviation = worksheet.cell(i, 9).value

        self.errmsg = None
        self.severity = None

    def is_valid(self, c, logger):

        # Test for exceptions. An ERROR will cause the whole run to abort,
        # rolling back the transaction

        logger.debug("Checking measureunitid %s is valid" % (self.measureunitid))
        if self.measureunitid == "":
            logger.error("Blank measureunitid is illegal")
            return False
        if len(self.abbreviation) > 8:
            logger.error("abbreviation %s exceeds maximum allowed length of 8 characters", (self.abbreviation))
            return False
        if len(self.description) > 100:
            logger.error("description %s exceeds maximum allowed length of 100 characters", (self.description))
            return False
        return True

########################################################################################
#
#     Functions
#
########################################################################################

def get_next_txn_id(c, logger):

    # Get the next available transaction id for a MEA batch
    q_get_id = """
    SELECT maximo.hul_mealie_uom_seq.NEXTVAL
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
    AND ifacename = 'mxmeasure_iface'"""

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
    VALUES ('EXTSYS1', 'mxmeasure_iface', 'AddChange', :trans_id)"""

    logger.debug("Writing MEA queue for transid %d" % txid)
    c.execute(q_write_queue, trans_id=txid)

    return True


def uom_handler(er):

    type_handler = {
        "XLS": uom_handler_xls_ora
    }

    return type_handler.get(er.source_type)(er)   # Call appropriate handler for the source type


def uom_handler_xls_ora(er):

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

    # Insert a row into MEA interface for entity type ASSET

    q_write_interface = """
    INSERT INTO maximo.mxmeasure_iface (
        measureunitid,
        abbreviation, 
        description,
        transid,
        transseq)
    VALUES (
        :measureunitid,
        :abbreviation, 
        :description,
        :transid,
        :transseq)"""

    # Initialise local variables for fetch loop

    total_ct = 0       # Total no. of items
    start_row = 1      # The first row in spreadsheet that contains data (starting at 0)
    processed_mu = []  # Values of measureunitid that have been processed. Only first occurrence in file is processed

    # Iterate through all rows of data in input spreadsheet

    for this_row in range(start_row, wsh.nrows):

        mu = InterfaceableMeasureunitFromXls(wsh, this_row, curs_target_seq, er.logger)
        if mu.measureunitid == "":
        
            er.logger.info("Skipping specification with no UOM")
            
        else:
        
            er.logger.info("Processing measureunitid %s" % (mu.measureunitid))

            if (mu.is_valid(curs_target_lookup, er.logger)):
            
                # Have we already processed this Unit of Measure?
                if mu.measureunitid not in processed_mu:

                    # Obtain new MEA transaction id for next asset 

                    er.logger.debug("Getting new transaction id for this load")
                    trans_id = get_next_txn_id(curs_target_seq, er.logger)

                    trans_seq = 1          # Sequence within batch

                    # Write the MEA interface table
                    er.logger.debug("Processing interface record for measureunitid=%s" % (mu.measureunitid))
                
                    er.logger.info("Writing interface: measureunitid=%s transid=%s transseq=%s" % (mu.measureunitid, trans_id, trans_seq))

                    er.logger.debug("description=%s" % (mu.description))
                    er.logger.debug("abbreviation=%s" % (mu.abbreviation))

                    curs_target_mea.execute(q_write_interface, measureunitid=mu.measureunitid,
                                                               description=mu.description,
                                                               abbreviation=mu.abbreviation,
                                                               transid=trans_id,
                                                               transseq=trans_seq)
                    er.logger.debug("Wrote interface")
                    processed_mu.append(mu.measureunitid)     # We processed it, remember it

                    # Write the MEA queue
                    write_mea_queue(curs_target_mea, trans_id, er.logger)
                    er.logger.debug("Wrote MEA queue for transid %d" % trans_id)
    
                    total_ct += 1
                
                else:
                
                    er.logger.info("measureunitid %s was already processed, skipping" % (mu.measureunitid))
                
            else:

                er.logger.error("ERROR: Values for Measure Unit %s are invalid, aborting" % (mu.measureunitid))
                print("ERROR: Values for Measure Unit %s are invalid, aborting" % (mu.measureunitid))
                conn_target.rollback()
                conn_target.close()         # Free the connection
                sys.exit(2)


    # (-) End of measure unit processing loop

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
    queued_batches = get_queued_batch_ct(curs_target_seq, er.logger)
    while (queued_batches):
        er.logger.debug("MEA still has %s items in queue" % (queued_batches))
        queued_batches = get_queued_batch_ct(curs_target_seq, er.logger)

    if (not er.test_mode):
        # Tune this value according to your cron configuration
        er.logger.info("Waiting %d seconds for MEA to finish loading" % (er.mea_wait_interval))
        time.sleep(er.mea_wait_interval)

    print("Processed %s %s entities" % (total_ct, er.entity_type))
    print "See %s for full details of this run" % er.log_file

    # Close open cursors and disconnect from MEA target database

    curs_target_seq.close()
    curs_target_mea.close()
    curs_target_lookup.close()

    conn_target.close()

    return True
