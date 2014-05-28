'''
Created on 14 Jun 2012

@author: ACE
'''
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
class InterfaceableEntryFromXls:

    """Class representing a UNSPSC Dictionary Entity to be inserted from a spreadsheet"""

    def __init__(self, worksheet, i, c, logger):
        
        # Values are taken from the UNSPSC Dictionary Template

        self.escn = worksheet.cell(i, 0).value
        self.esci = worksheet.cell(i, 1).value[:100] # commodities.description is varchar2(100)
        self.attr_seq = worksheet.cell(i, 2).value
        self.espn = worksheet.cell(i, 3).value[:100]
        self.espi = worksheet.cell(i, 4).value
        
        self.errmsg = None
        self.severity = None

    def is_valid(self, c, logger):

        # Test for exceptions. An ERROR will cause the whole run to abort,
        # rolling back the transaction

        logger.debug("Validating %s %s %s %s %s" % (self.escn, self.esci, self.attr_seq, self.espn, self.espi))
        if self.escn == "":
            logger.error("ESCN must be supplied")
            return False
        if self.esci == "":
            logger.error("ESCI must be supplied")
            return False
        if self.attr_seq == "":
            logger.error("Attribute Sequence must be supplied")
            return False
        if self.espn == "":
            logger.error("ESPN must be supplied")
            return False
        if self.espi == "":
            logger.error("ESPI must be supplied")
            return False
        
        return True

########################################################################################
#
#     Functions
#
########################################################################################

def get_next_txn_id(c, logger):
    
    #Get next available sequence number for batch 
    
    q_get_cls_seq = """
    SELECT maximo.hul_classif_dict_seq.NEXTVAL
    FROm DUAL"""
    
    c.execute(q_get_cls_seq)
    mealie_seq = c.fetchone()
    logger.debug("Obtained MEA seq Id %s" % (mealie_seq[0]))
    
    return mealie_seq[0]

def write_dict(c, escn, esci, attribute_seq, espn, espi, logger):

    # Insert row into dictionary

    q_insert_dict = """
    INSERT INTO maximo.hul_classif_dict (
                           escn,
                           esci,
                           attribute_seq,
                           espn,    
                           espi
                           )
    VALUES (:escn, :esci, :attribute_seq, :espn, :espi)"""

    q_update_dict = """
    UPDATE maximo.hul_classif_dict
    SET escn = :escn, esci = UPPER(:esci), attribute_seq = :attribute_seq, espn = :espn, espi = :espi
    WHERE escn = :escn
    AND esci = :esci
    AND attribute_seq = :attribute_seq
    AND espn = :espn
    AND espi = :espi"""

    logger.debug("Inserting dictionary for %s %s %s %s" % (escn, esci, espi, espn))
    try:
        c.execute(q_insert_dict, escn=escn, esci=esci, attribute_seq=attribute_seq, espn=espn, espi=espi)
    except cx_Oracle.IntegrityError:
        logger.debug("Row already exists, attempting update")
        logger.debug("Updating dictionary for %s %s %s %s" % (esci, escn, espi, espn))
        c.execute(q_update_dict, escn=escn, esci=esci, attribute_seq=attribute_seq, espn=espn, espi=espi)

    return True


def classif_handler(er):

    type_handler = {
        "XLS": und_handler_xls_ora
    }

    return type_handler.get(er.source_type)(er)   # Call appropriate handler for the source type


def und_handler_xls_ora(er):

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

    # Connect to target database
    try:
        conn_target = cx_Oracle.connect("%s/%s@%s" % (er.target_user, er.target_pwd, er.target_db))
    except cx_Oracle.DatabaseError, exc:
        error, = exc.args
        print("ERROR: %s" %  (error.message))
        er.logger.error("%s" % (error.message))
        return False
    
    # Open the necessary cursors on target database
    curs_target = conn_target.cursor()

    # Initialise local variables for fetch loop
    total_ct = 0       # Total no. of items
    start_row = 1      # The first row in spreadsheet that contains data (starting at 0)
    
    prev_segment = None
    prev_family = None
    prev_classcode = None

    # Iterate through all rows of data in input spreadsheet

    for this_row in range(start_row, wsh.nrows):

        ie = InterfaceableEntryFromXls(wsh, this_row, curs_target, er.logger)
        er.logger.info("Processing %s %s %s %s %s" % (ie.escn, ie.esci, ie.attr_seq, ie.espn, ie.espi))

        if (ie.is_valid(curs_target, er.logger)):
            
            er.logger.debug("Class Identifier=%s" % (ie.esci))
            er.logger.debug("Property Identifier=%s" % (ie.espi))
            #trans_id = get_next_txn_id(c, logger)
            
            
            # Code type 4 represents Commodity level
            write_dict(curs_target, ie.escn, ie.esci, ie.attr_seq, ie.espn, ie.espi, er.logger)
            er.logger.debug("Processed to Dictionary")
            total_ct += 1
                
        else:
                
            #er.logger.error("ERROR: Values for Commodity Code %s are invalid, aborting" % (ie.commcode))
            #print("ERROR: Values for Commodity Code %s are invalid, aborting" % (ie.commcode))
            #conn_target.rollback()
            #conn_target.close()         # Free the connection
            sys.exit(2)

    # The transaction will not be committed if we are in 'test mode'
    # Otherwise we commit the whole transaction

    if (not er.test_mode):
        er.logger.info("Committing changes to database target %s" % er.target_db)
        conn_target.commit() 
    else:
        er.logger.info("Test Mode: Rolling back changes to %s" % er.target_db)
        conn_target.rollback() 

    er.logger.info("Processed %s %s entities" % (total_ct, er.entity_type))

    print("Processed %s %s entities" % (total_ct, er.entity_type))
    print "See %s for full details of this run" % er.log_file

    # Close open cursors and disconnect from target database
    curs_target.close()
    conn_target.close()

    return True
