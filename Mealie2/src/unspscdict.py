#/bin/env python
#
# $Header: /usr/local/cvsroot/mealie/entities/unspscdict.py,v 1.2 2012/02/09 10:02:11 djw Exp $
#
# MEA Lightweight Integration Environment
#
# Entity Handler for UNSPSC DICTIONARY
#
# Issue Track 8112
#
# $Log: unspscdict.py,v $
# Revision 1.2  2012/02/09 10:02:11  djw
# 8112 Descriptions are always in upper case
#
# Revision 1.1  2012/02/09 07:53:35  djw
# 8112 UND entity unit tested. Test data at maximo/8112/hul_test_UND.stg
# maximo/8112/hul_test_unspsc_dict.stg
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

class InterfaceableEntryFromXls:

    """Class representing a UNSPSC Dictionary Entity to be inserted from a spreadsheet"""

    def __init__(self, worksheet, i, c, logger):
        
        # Values are taken from the UNSPSC Dictionary Template

        self.segment = worksheet.cell(i, 0).value
        self.segment_desc = worksheet.cell(i, 1).value[:100] # commodities.description is varchar2(100)
        self.family = worksheet.cell(i, 2).value
        self.family_desc = worksheet.cell(i, 3).value[:100]
        self.classcode = worksheet.cell(i, 4).value
        self.classcode_desc = worksheet.cell(i, 5).value[:100]
        self.commcode = worksheet.cell(i, 7).value
        self.commcode_desc = worksheet.cell(i, 8).value[:100]
        self.definition = worksheet.cell(i, 9).value # At commodity level

        self.errmsg = None
        self.severity = None

    def is_valid(self, c, logger):

        # Test for exceptions. An ERROR will cause the whole run to abort,
        # rolling back the transaction

        logger.debug("Validating %s %s %s %s" % (self.segment, self.family, self.classcode, self.commcode))
        if self.segment == "":
            logger.error("Segment must be supplied")
            return False
        if self.segment_desc == "":
            logger.error("Segment description must be supplied")
            return False
        if self.family == "":
            logger.error("Family must be supplied")
            return False
        if self.family_desc == "":
            logger.error("Family description must be supplied")
            return False
        if self.classcode == "":
            logger.error("Class must be supplied")
            return False
        if self.classcode_desc == "":
            logger.error("Class description must be supplied")
            return False
        if self.commcode == "":
            logger.error("Commodity must be supplied")
            return False
        if self.commcode_desc == "":
            logger.error("Commodity description must be supplied")
            return False

        return True

########################################################################################
#
#     Functions
#
########################################################################################

def write_dict(c, code, codetype, description, definition, logger):

    # Insert row into dictionary

    q_insert_dict = """
    INSERT INTO maximo.hul_unspsc_dict (
                           code,
                           codetype,
                           description,
                           definition
                           )
    VALUES (:code, :codetype, UPPER(:description), :definition)"""

    q_update_dict = """
    UPDATE maximo.hul_unspsc_dict
    SET codetype = :codetype, description = UPPER(:description), definition = :definition
    WHERE code = :code"""

    logger.debug("Inserting dictionary for %d %s %s" % (codetype, code, description))
    try:
        c.execute(q_insert_dict, code=code, codetype=codetype, description=description, definition=definition)
    except cx_Oracle.IntegrityError:
        logger.debug("Row already exists, attempting update")
        logger.debug("Updating dictionary for %d %s %s" % (codetype, code, description))
        c.execute(q_update_dict, code=code, codetype=codetype, description=description, definition=definition)

    return True


def und_handler(er):

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
        print("ERROR: %s" % (error.message))
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
        er.logger.info("Processing %s %s %s %s" % (ie.segment, ie.family, ie.classcode, ie.commcode))

        if (ie.is_valid(curs_target, er.logger)):
            
            er.logger.debug("commcode=%s" % (ie.commcode))
            er.logger.debug("commcode_desc=%s" % (ie.commcode_desc))

            # Code type 1 represents Segment level
            if ie.segment != prev_segment:
                write_dict(curs_target, ie.segment, 1, ie.segment_desc, None, er.logger)
                total_ct += 1
                prev_segment = ie.segment

            # Code type 2 represents Family level
            if ie.family != prev_family:
                write_dict(curs_target, ie.family, 2, ie.family_desc, None, er.logger)
                total_ct += 1
                prev_family = ie.family

            # Code type 3 represents Class level
            if ie.classcode != prev_classcode:
                write_dict(curs_target, ie.classcode, 3, ie.classcode_desc, None, er.logger)
                total_ct += 1
                prev_classcode = ie.classcode

            # Code type 4 represents Commodity level
            write_dict(curs_target, ie.commcode, 4, ie.commcode_desc, ie.definition, er.logger)
            total_ct += 1
                
        else:
                
            er.logger.error("ERROR: Values for Commodity Code %s are invalid, aborting" % (ie.commcode))
            print("ERROR: Values for Commodity Code %s are invalid, aborting" % (ie.commcode))
            conn_target.rollback()
            conn_target.close()         # Free the connection
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
