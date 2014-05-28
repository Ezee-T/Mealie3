#!/usr/bin/env python
#
# $Header: /usr/local/cvsroot/mealie/entities/assets.py,v 1.1 2012/01/05 12:22:34 djw Exp $
#
# MEA Lightweight Integration Environment
#
# Entity Handler for ASSETS
#
# $Log: assets.py,v $
# Revision 1.1  2012/01/05 12:22:34  djw
# Migrated to new project
#
# Revision 1.19  2011/06/09 13:06:42  djw
# 7798 disabled (column J in the spreadsheet) and
# isrunning (column K in the spreadsheet) must be booleans
#
# Revision 1.18  2011/04/19 08:55:49  djw
# 7624 Parents assets and locations and asset locations must be OPERATING
#
# Revision 1.17  2011/04/15 15:44:51  djw
# 7591 Error handling for IO error when attempting to open input file
#
# Revision 1.16  2011/04/15 15:19:59  djw
# 7591 Propogate return status from handler method
#
# Revision 1.15  2011/01/05 11:58:09  djw
# MA00101 Permit a blank location in ASSET spreadsheet data
#
# Revision 1.14  2010/12/29 13:58:01  djw
# MA00062 get_next_txn_id() Moved from base/common, separate
# sequence per entity
#
# Revision 1.13  2010/12/29 12:45:48  djw
# MA00056 Give clear message on Oracle connect error
#
# Revision 1.12  2010/12/07 08:59:40  djw
# MA00063 If the orgid is not a numeric value it is set to None, then it will
# be rejected anyway with the exit status: ERROR: The following ORGIDS are
# incorrect: [None], aborting
#
# Revision 1.11  2010/11/26 14:46:58  djw
# MA00047 Now must supply one of -u -f -t as arg.
# Processing added for Full Load runs that can be
# batched in a script and will abort if MEA does
# not load all entities after a configurable
# interval mea_wait_interval
#
# Revision 1.10  2010/11/24 13:36:57  djw
# MA00049 Enhanced so that non-numeric or blank value in the level
# column A can be handled without a crash and with a meaningful
# error message in the log
#
# Revision 1.9  2010/11/23 13:51:51  djw
# MA00048 When processing spreadsheets, prompt user for target password,
# if not in parameters file.
#
# Revision 1.8  2010/11/18 13:14:59  djw
# 6924 Enhanced so that the different hierarchies themselves are processed in
# batches from highest to lowest , thus level switches occur less frequently
# Error handling made consistent. Improved debugging at level 10.
#
# Revision 1.7  2010/11/17 15:04:52  djw
# 6943 InterfaceableAssetFromXls orgid property should be integer
#
# Revision 1.6  2010/10/27 10:08:22  njm3
# Added error count
#
# Revision 1.5  2010/10/27 10:07:09  njm3
# Oops - rectified incorrect class name
#
# Revision 1.4  2010/10/27 10:02:14  njm3
# Added ORGID Validation
#
# Revision 1.3  2010/10/26 12:38:28  njm3
# Changed mealie path - one dir higher:
#     e.g. mealie.base.x is now base.x
#     (Update MEALIE_PATH in ~/.bash_profile accordingly: /usr/local/lib/python/mealie)
#
# Revision 1.2  2010/09/09 07:15:39  djw
# 6261 Now able to read source data from spreadsheet as well as Oracle database
#
# Revision 1.1  2010/04/14 12:19:35  djw
# MA00036. Initial revision
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

from common import read_orgids

########################################################################################
#
#     Class definitions
#
########################################################################################

class InterfaceableAssetFromOra:

    """Class to represent an Asset to be loaded by MEA from a resultset"""

    def __init__(self, asset_resultset):

        self.level = asset_resultset[0]
        self.ancestor = asset_resultset[1]
        self.parent = asset_resultset[2]
        self.assetnum = asset_resultset[3]
        self.location = asset_resultset[4]
        self.description = asset_resultset[5]
        self.orgid = asset_resultset[6]
        self.siteid = asset_resultset[7]
        self.status = asset_resultset[8]
        self.disabled = asset_resultset[9]
        self.isrunning = asset_resultset[10]

        self.errmsg = None
        self.severity = None

    def is_valid(self, c, logger):

        # Test for exceptions. An ERROR will cause the whole run to abort,
        # rolling back the transaction

        # The parent must already be an existing ASSET

        logger.debug("Checking parent %s of asset %s exists as an asset" % (self.parent, self.assetnum))

        q_exists_ass="""
            SELECT assetnum
            FROM maximo.asset a
            WHERE a.status = 'OPERATING'
            AND a.assetnum = :assetnum"""
        c.execute(q_exists_ass, assetnum=self.parent)
        parent_ass = c.fetchone()
        logger.debug("Asset %s looked up: %s" % (self.parent, parent_ass))

        if (not parent_ass):
            logger.debug("%s is not a valid existing Asset, will check interface table" % (self.parent))

            logger.debug("Checking parent %s of asset %s exists on interface table" % (self.parent, self.assetnum))

            q_exists_interface="""
                SELECT assetnum
                FROM maximo.mxasset_iface i
                WHERE i.assetnum = :assetnum"""
            c.execute(q_exists_interface, assetnum=self.parent)
            iface_ass = c.fetchone()
            logger.debug("Asset %s looked up: %s" % (self.parent, iface_ass))

            if (not iface_ass):
                logger.debug("%s is not an Asset previously written to interface table" % (self.parent))
                return False

        logger.debug("Checking location %s of asset %s exists as a location" % (self.location, self.assetnum))

        q_exists_loc="""
            SELECT location
            FROM maximo.locations l
            WHERE l.status = 'OPERATING'
            AND l.location = :location"""
        c.execute(q_exists_loc, location=self.location)
        ass_loc = c.fetchone()
        logger.debug("Location %s looked up: %s" % (self.location, ass_loc))

        if (not ass_loc):
            return False

        return True

class InterfaceableAssetFromXls:

    """Class to represent an Asset to be loaded by MEA from a spreadsheet"""

    def __init__(self, worksheet, i, logger):

        self.assetnum = worksheet.cell(i, 3).value
        # MA00049 If user does not provide an integer value for Level we catch the ValueError 
        #         and set level to None, it will be rejected by the is_valid() method
        try:
            self.level = int(worksheet.cell(i, 0).value)
        except ValueError:
            logger.debug("Asset %s has non-numeric level \"%s\", setting to None" % (self.assetnum, worksheet.cell(i, 0).value))
            self.level = None
      
        self.ancestor = worksheet.cell(i, 1).value
        self.parent = worksheet.cell(i, 2).value
        if (worksheet.cell(i, 4).ctype == xlrd.XL_CELL_NUMBER):
            self.location = str(int(worksheet.cell(i, 4).value))
        else:
            self.location = worksheet.cell(i, 4).value
        self.description = worksheet.cell(i, 5).value
        try: # MA00063 If non-numeric orgid, set to None
            self.orgid = int(worksheet.cell(i, 6).value) # 6943 should be integer
        except ValueError:
            logger.debug("Asset %s has non-numeric orgid \"%s\", setting to None" % (self.assetnum, worksheet.cell(i, 6).value))
            self.orgid = None
        self.siteid = worksheet.cell(i, 7).value
        self.status = worksheet.cell(i, 8).value
        self.disabled = worksheet.cell(i, 9).value
        self.isrunning = worksheet.cell(i, 10).value

        self.errmsg = None
        self.severity = None

    def is_valid(self, c, logger):

        # Test for exceptions. An ERROR will cause the whole run to abort,
        # rolling back the transaction

        # MA00049 The Level must be an integer

        logger.debug("Checking level \"%s\" of asset %s is integer" % (self.level, self.assetnum))
        if type(self.level) is not int:
            logger.error("Asset %s has non-integer level, check source data" % (self.assetnum))
            return False

        # 7798 isrunning and disabled must be booleans

        if (self.isrunning != 0 and self.isrunning != 1):
            logger.error("Asset %s has invalid isrunning value %d, check source data" % 
                (self.assetnum, self.isrunning))
            return False

        if (self.disabled != 0 and self.disabled != 1):
            logger.error("Asset %s has invalid disabled value %d, check source data" % 
                (self.assetnum, self.disabled))
            return False

        # The parent must already be an existing ASSET

        logger.debug("Checking parent %s of asset %s exists as an asset" % (self.parent, self.assetnum))

        q_exists_ass="""
            SELECT assetnum
            FROM maximo.asset a
            WHERE a.status = 'OPERATING'
            AND a.assetnum = :assetnum"""
        c.execute(q_exists_ass, assetnum=self.parent)
        parent_ass = c.fetchone()
        logger.debug("Asset %s looked up: %s" % (self.parent, parent_ass))

        if (not parent_ass):
            logger.debug("%s is not a valid existing Asset, will check interface table" % (self.parent))

            logger.debug("Checking parent %s of asset %s exists on interface table" % (self.parent, self.assetnum))

            q_exists_interface="""
                SELECT assetnum
                FROM maximo.mxasset_iface i
                WHERE i.assetnum = :assetnum"""
            c.execute(q_exists_interface, assetnum=self.parent)
            iface_ass = c.fetchone()
            logger.debug("Asset %s looked up: %s" % (self.parent, iface_ass))

            if (not iface_ass):
                logger.debug("%s is not an Asset previously written to interface table" % (self.parent))
                return False

        logger.debug("Checking location %s of asset %s exists as a location" % (self.location, self.assetnum))

        if (self.location):
            logger.debug("Asset %s has a location" % (self.assetnum))
            q_exists_loc="""
                SELECT location
                FROM maximo.locations l
                WHERE l.status = 'OPERATING'
                AND l.location = :location"""
            c.execute(q_exists_loc, location=self.location)
            ass_loc = c.fetchone()
            logger.debug("Location %s looked up: %s" % (self.location, ass_loc))

            if (not ass_loc):
                return False

        else:
            # MA00101 If the location is blank we allow it but MEA will make the asset
            #         inherit the parent's location
            logger.debug("Asset %s has no location, skipping check" % (self.assetnum))

        return True

########################################################################################
#
#     Functions
#
########################################################################################

def get_next_txn_id(c, logger):

    # Get the next available transaction id for a MEA batch
    q_get_id = """
    SELECT maximo.hul_mealie_ass_seq.NEXTVAL
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
    AND ifacename = 'MXASSETInterface'"""

    time.sleep(5)  # t = seconds
    c.execute(q_get_ct)
    wait_ct = c.fetchone()
    logger.info("MEA has %s unprocessed batches in the queue" % (wait_ct[0]))

    return wait_ct[0]

def get_entity_ct(c, logger):

    # Get the number of entities in Maximo
    q_get_ct = """
    SELECT COUNT(*)
    FROM maximo.asset"""

    c.execute(q_get_ct)
    entity_ct = int(c.fetchone()[0])
    logger.info("There are %d ASS entities in Maximo" % (entity_ct))

    return entity_ct



def write_mea_queue(c, q, logger):

    # Insert rows into MEA queue

    q_write_queue = """
    INSERT INTO maximo.mxin_inter_trans (
                           extsysname,
                           ifacename,
                           transid
                           )
    VALUES ('EXTSYS1',
               'MXASSETInterface',
               :transid)"""

    # Process each entry in the queue until it is empty

    for _i in range(len(q)):
        j = q.pop(0)
        logger.debug("Popped txn id %s. Queue now holds %s" % (j, q))
        c.execute(q_write_queue, transid = j)

    return True


def ass_handler(er):

    type_handler = {
        "ORA": ass_handler_ora_ora,
        "XLS": ass_handler_xls_ora
    }

    return type_handler.get(er.source_type)(er) # Call appropriate handler for the source type


def ass_handler_xls_ora(er):

    # MA00048 Prompt user for target password, if not in parameters
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
    # 6924 DJW Moved to before the validation block

    try:
        conn_target = cx_Oracle.connect("%s/%s@%s" % (er.target_user, er.target_pwd, er.target_db))
    except cx_Oracle.DatabaseError, exc:
        error, = exc.args
        print("ERROR: %s" % (error.message))
        er.logger.error("%s" % (error.message))
        return False
    
    # Validation of ORGIDS and GLACCOUNTS
    # 6924 DJW On full debug dump out all valid values
    orgids = read_orgids(conn_target)
    er.logger.debug("Valid orgids obtained: These are %s" % ("".join(str(orgids))))
    er.logger.debug("Declaring levels[] list")

    levels = []        # A list of the levels in this spreadsheet
    invalid_orgids = []
    orgids_err_ct = 0
    start_row = 1

    er.logger.debug("Validating spreadsheet data")

    for x in range(start_row, wsh.nrows):

        ia = InterfaceableAssetFromXls(wsh, x, er.logger)
        er.logger.debug("Instantiated InterfaceableAssetFromXls for %s" % (ia.assetnum))
        # MA00049 InterfaceableAssetFromXls.level is supposed to be an integer so no need to cast
        if (ia.level not in levels):
            levels.append(ia.level)

        if (ia.orgid not in orgids):        # Validate ORGIDS
            if (ia.orgid not in invalid_orgids): # Don't report duplicate error values
                invalid_orgids.append(ia.orgid)
            orgids_err_ct += 1

    if orgids_err_ct > 0:
        # 6924 DJW Update for consistency with existing fatal errors
        er.logger.error("ERROR: The following ORGIDS are incorrect: %s, aborting" % invalid_orgids)
        print("ERROR: The following ORGIDS are incorrect: %s, aborting" % invalid_orgids)
        conn_target.rollback()
        conn_target.close() # 6924 DJW Free the connection
        sys.exit(2)

    er.logger.debug("Spreadsheet data validated")

    levels.sort()
    er.logger.info("Levels to be processed: %s" % (levels))
        
    # Open various cursors on target database

    curs_target_mea = conn_target.cursor()
    curs_target_lookup = conn_target.cursor()
    curs_target_seq = conn_target.cursor()

    # MA00047 Determine how many entities exist before we start
    start_entity_ct = get_entity_ct(curs_target_lookup, er.logger)

    # Insert a row into MEA interface for entity type ASSET

    q_write_interface = """
    INSERT INTO maximo.mxasset_iface (
        ancestor,
        as_description,
        as_location,
        as_orgid,
        as_siteid,
        as_status,
        assetnum,
        disabled,
        isrunning,
        parent,
        transid,
        transseq)
    VALUES (
        :ancestor,
        :description,
        :location,
        :orgid,
        :siteid,
        :status,
        :assetnum,
        :disabled,
        :isrunning,
        :parent,
        :transid,
        :transseq)"""

    # Initialise local variables for fetch loop

    queued_batches = 0 
    total_ct = 0       # Total no. of items
    level_ct = 0       # No. of items in this level
    last_level = None
    trans_id = None    # Current Batch id
    batch_queue = []   # Batch id queue for queue table
    trans_seq = None   # Sequence within batch

    start_row = 1      # First row in spreadsheet that contains data (first is 0)

    # 6924 DJW Iterate through levels starting with topmost in hierarchy

    for level_to_process in levels:

        # MA00049 Do not assume at this stage that all members of levels list are numeric
        er.logger.info("Processing hierarchy level %s" % (level_to_process))

        # Iterate through rows in input spreadsheet

        for this_row in range(start_row, wsh.nrows):

            ia = InterfaceableAssetFromXls(wsh, this_row, er.logger)

            # 6924 DJW Ignore unless it is the level we want to process on this iteration
            er.logger.debug("Current level being processed: %s" % (level_to_process))
            er.logger.debug("This asset is of level : %s" % (ia.level))


            if (ia.level == level_to_process):
                er.logger.debug("We are processing level %s so process this one" % (level_to_process))

                er.logger.info("Processing assetnum %s" % (ia.assetnum))
                if (ia.is_valid(curs_target_lookup, er.logger)):

                    if (ia.level != last_level):
                        er.logger.debug("LEVEL has changed: was %s, now %s" % (last_level, ia.level))

                        if (not er.test_mode):
                            if (last_level): # Ignore the first change of level since nothing yet written

                                # Write the MEA queue
                                er.logger.info("Writing MEA queue, stored txn ids are %s" % batch_queue)
                                write_mea_queue(curs_target_mea, batch_queue, er.logger)
                                er.logger.info("Wrote MEA queue, stored txn ids are %s" % batch_queue)
                                er.logger.info("Committing changes to MEA target %s" % er.target_db)
                                conn_target.commit() 

                                er.logger.debug("Processed %s items in previous LEVEL %s" % (level_ct, last_level))
                                level_ct = 0

                                # Wait until MEA has flushed the queue

                                er.logger.info("Wait for MEA to flush its queue")
                                queued_batches = get_queued_batch_ct(curs_target_seq, er.logger)
                                while (queued_batches):
                                    er.logger.debug("MEA still has %s items in queue" % (queued_batches))
                                    queued_batches = get_queued_batch_ct(curs_target_seq, er.logger)

                        else:
                            er.logger.info("Test Mode: Will not write MEA queue")

                        last_level = ia.level

                    else:
                        er.logger.debug("LEVEL is still %s" % (ia.level))

                    er.logger.debug("Getting new transaction id for %s" % ia.assetnum)

                    # Obtain new MEA transaction id for next batch 
                    trans_id = get_next_txn_id(curs_target_seq, er.logger)

                    # Add the MEA transaction id to the queue 
                    batch_queue.append(trans_id) 

                    # Sequence is reset for new batch
                    trans_seq = 1

                    # Write the MEA interface table

                    er.logger.debug("Writing interface: assetnum %s transid %s transseq %s" % (ia.assetnum, trans_id, trans_seq))
                    curs_target_mea.execute(q_write_interface, ancestor=ia.ancestor,
                                                                              description=ia.description,
                                                                              location=ia.location,
                                                                              orgid=ia.orgid,
                                                                              siteid=ia.siteid,
                                                                              status=ia.status,
                                                                              assetnum=ia.assetnum,
                                                                              disabled=ia.disabled,
                                                                              isrunning=ia.isrunning,
                                                                              parent=ia.parent,
                                                                              transid=trans_id,
                                                                              transseq=trans_seq)
                else:

                    er.logger.error("ERROR: %s is not a valid Asset, aborting" % (ia.assetnum))
                    print("ERROR: %s is not a valid Asset, aborting" % (ia.assetnum))
                    conn_target.rollback()
                    conn_target.close() # 6924 DJW Free the connection
                    sys.exit(2)  # Will never happen if driving query is correctly ordered
    
                total_ct += 1
                level_ct += 1

            else:
                er.logger.debug("We are processing level %s so ignore level %s" % (level_to_process, ia.level))

            # 6924 DJW End of condition that it is the level we want to process on this iteration

        # End of asset processing loop

    # 6924 DJW End of level processing loop

    # The transaction will not be committed if we are in 'test mode'
    # Otherwise we commit the last level-set of batches

    if (not er.test_mode):
        # Write the MEA queue
        er.logger.info("Writing MEA queue, stored txn ids are %s" % batch_queue)
        write_mea_queue(curs_target_mea, batch_queue, er.logger)
        er.logger.info("Wrote MEA queue, stored txn ids are %s" % batch_queue)
        er.logger.info("Committing changes to MEA target %s" % er.target_db)
        conn_target.commit() 
    else:
        er.logger.info("Test Mode: Rolling back changes to %s" % er.target_db)
        conn_target.rollback() 

    er.logger.info("Processed %s %s entities" % (total_ct, er.entity_type))

    if (not er.test_mode):
        # MA00047 Tune this value according to your cron configuration
        er.logger.info("Waiting %d seconds for MEA to finish loading" % (er.mea_wait_interval))
        time.sleep(er.mea_wait_interval)

    print("Processed %s %s entities" % (total_ct, er.entity_type))
    print "See %s for full details of this run" % er.log_file

    if (er.full_mode):
        # MA00047 Determine how many assets exist after all processing is done
        end_entity_ct = get_entity_ct(curs_target_lookup, er.logger)
        if (start_entity_ct + total_ct == end_entity_ct):
            er.logger.info("OK Full Load verified, entity count increased by %d" % (total_ct))
        else:
            er.logger.error("Full Load not verified, entity count increased by %d, aborting" % (end_entity_ct - start_entity_ct))
            print("ERROR: Full Load not verified, entity count increased by %d, aborting" % (end_entity_ct - start_entity_ct))
            conn_target.close() # Free connection
            sys.exit(2) # MA00047 Some entities are committed but numbers don't add up

    # Close open cursors and disconnect from MEA target database

    curs_target_seq.close()
    curs_target_mea.close()
    curs_target_lookup.close()

    conn_target.close()

    return True

def ass_handler_ora_ora(er):

    # Prompt user for passwords, if not in parameters 
    # file. When used in Production NEVER put passwords 
    # in the parameter file. 

    if (not er.source_pwd):
        er.source_pwd = str(raw_input("Enter password for source %s@%s: " % (er.source_user, er.source_db)))

    if (not er.target_pwd):
        er.target_pwd = str(raw_input("Enter password for target %s@%s: " % (er.target_user, er.target_db)))

    # Connect to Maximo source database and open the driving cursor

    conn_source = cx_Oracle.connect("%s/%s@%s" % (er.source_user, er.source_pwd, er.source_db))
    curs_source_drv = conn_source.cursor()

    # Connect to MEA target database and open various cursors

    conn_target = cx_Oracle.connect("%s/%s@%s" % (er.target_user, er.target_pwd, er.target_db))
    curs_target_mea = conn_target.cursor()
    curs_target_lookup = conn_target.cursor()
    curs_target_seq = conn_target.cursor()

    # Driving query for entity type ASSET

    # 14-APR-2010 Dave Hudson confirms that all assets in scope have 
    #             'M' as their ultimate parent so we can safely use this
    #             as the 'START WITH' assetnum

    q_driver = """
    SELECT LEVEL,
              a.ancestor,
              a.parent,
              a.assetnum,
              a.location,
              a.description,
              a.orgid,
              a.siteid,
              a.status,
              a.disabled,
              a.isrunning
    FROM maximo.asset a
    WHERE a.location LIKE :locations
    CONNECT BY PRIOR a.assetnum = a.parent
    START WITH a.assetnum = 'M'
    ORDER BY LEVEL, a.location"""


    # Insert a row into MEA interface for entity type ASSET

    q_write_interface = """
    INSERT INTO maximo.mxasset_iface (
        ancestor,
        as_description,
        as_location,
        as_orgid,
        as_siteid,
        as_status,
        assetnum,
        disabled,
        isrunning,
        parent,
        transid,
        transseq)
    VALUES (
        :ancestor,
        :description,
        :location,
        :orgid,
        :siteid,
        :status,
        :assetnum,
        :disabled,
        :isrunning,
        :parent,
        :transid,
        :transseq)"""

    # Fetch first entity

    curs_source_drv.execute(q_driver, locations = er.entity_key + '%')

    row = curs_source_drv.fetchone()

    # Initialise local variables for fetch loop

    queued_batches = 0 
    total_ct = 0       # Total no. of items
    level_ct = 0       # No. of items in this level
    last_level = None
    trans_id = None    # Current Batch id
    batch_queue = []   # Batch id queue for queue table
    trans_seq = None   # Sequence within batch

    # Process each entity until the cursor is exhausted

    while row:
        ia = InterfaceableAssetFromOra(row)
        er.logger.info("Processing asset %s" % (ia.assetnum))
        if (ia.is_valid(curs_target_lookup, er.logger)):


            if (ia.level != last_level):
                er.logger.debug("LEVEL has changed: was %s, now %s" % (last_level, ia.level))

                if (not er.test_mode):
                    if (last_level): # Ignore the first change of level since nothing yet written

                        # Write the MEA queue
                        er.logger.info("Writing MEA queue, stored txn ids are %s" % batch_queue)
                        write_mea_queue(curs_target_mea, batch_queue, er.logger)
                        er.logger.info("Wrote MEA queue, stored txn ids are %s" % batch_queue)
                        er.logger.info("Committing changes to MEA target %s" % er.target_db)
                        conn_target.commit() 

                        er.logger.debug("Processed %s items in previous LEVEL %s" % (level_ct, last_level))
                        level_ct = 0

                        # Wait until MEA has flushed the queue

                        er.logger.info("Wait for MEA to flush its queue")
                        queued_batches = get_queued_batch_ct(curs_target_seq, er.logger)
                        while (queued_batches):
                            er.logger.debug("MEA still has %s items in queue" % (queued_batches))
                            queued_batches = get_queued_batch_ct(curs_target_seq, er.logger)

                else:
                    er.logger.info("Test Mode: Will not write MEA queue")

                last_level = ia.level

            else:
                er.logger.debug("LEVEL is still %s" % (ia.level))

            er.logger.debug("Getting new transaction id for %s" % ia.assetnum)

            # Obtain new MEA transaction id for next batch 
            trans_id = get_next_txn_id(curs_target_seq, er.logger)

            # Add the MEA transaction id to the queue 
            batch_queue.append(trans_id) 

            # Sequence is reset for new batch
            trans_seq = 1

            # Write the MEA interface table
            er.logger.debug("Writing interface: assetnum %s transid %s transseq %s" % (ia.assetnum, trans_id, trans_seq))
            curs_target_mea.execute(q_write_interface, ancestor=ia.ancestor,
                                                                      description=ia.description,
                                                                      location=ia.location,
                                                                      orgid=ia.orgid,
                                                                      siteid=ia.siteid,
                                                                      status=ia.status,
                                                                      assetnum=ia.assetnum,
                                                                      disabled=ia.disabled,
                                                                      isrunning=ia.isrunning,
                                                                      parent=ia.parent,
                                                                      transid=trans_id,
                                                                      transseq=trans_seq)

        else:

            er.logger.error("ERROR: %s is not a valid Asset, aborting" % (ia.assetnum))
            print("ERROR: %s is not a valid Asset, aborting" % (ia.assetnum))
            conn_target.rollback()
            conn_target.close() # 6924 DJW Free the connection
            sys.exit(2)  # This will never happen if driving query is correctly ordered
    
        total_ct += 1
        level_ct += 1
        row = curs_source_drv.fetchone()

    # The transaction will not be committed if we are in 'test mode'
    # Otherwise we commit the last level-set of batches

    if (not er.test_mode):
        # Write the MEA queue
        er.logger.info("Writing MEA queue, stored txn ids are %s" % batch_queue)
        write_mea_queue(curs_target_mea, batch_queue, er.logger)
        er.logger.info("Wrote MEA queue, stored txn ids are %s" % batch_queue)
        er.logger.info("Committing changes to MEA target %s" % er.target_db)
        conn_target.commit() 
    else:
        er.logger.info("Test Mode: Rolling back changes to %s" % er.target_db)
        conn_target.rollback() 

    er.logger.info("Processed %s %s entities" % (total_ct, er.entity_type))
    print("Processed %s %s entities" % (total_ct, er.entity_type))

    print "See %s for full details of this run" % er.log_file

    # Close open cursors and disconnect from source and MEA target

    curs_source_drv.close()
    curs_target_seq.close()
    curs_target_mea.close()
    curs_target_lookup.close()

    conn_source.close()
    conn_target.close()

    return True


