#!/usr/bin/env python
#
# $Header: /usr/local/cvsroot/mealie/entities/locations.py,v 1.1 2012/01/05 12:22:34 djw Exp $
#
# MEA Lightweight Integration Environment
#
# Entity Handler for LOCATIONS. Uniquely supports level by level processing
#
# $Log: locations.py,v $
# Revision 1.1  2012/01/05 12:22:34  djw
# Migrated to new project
#
# Revision 1.25  2011/11/23 06:28:08  djw
# 8093 Set new column status date to current date
#
# Revision 1.24  2011/04/19 08:55:49  djw
# 7624 Parents assets and locations and asset locations must be OPERATING
#
# Revision 1.23  2011/04/15 15:44:51  djw
# 7591 Error handling for IO error when attempting to open input file
#
# Revision 1.22  2011/04/15 15:19:59  djw
# 7591 Propogate return status from handler method
#
# Revision 1.21  2010/12/29 13:58:01  djw
# MA00062 get_next_txn_id() Moved from base/common, separate
# sequence per entity
#
# Revision 1.20  2010/12/29 12:45:48  djw
# MA00056 Give clear message on Oracle connect error
#
# Revision 1.19  2010/12/07 09:05:20  djw
# MA00063 Nesting error
#
# Revision 1.18  2010/12/07 08:59:40  djw
# MA00063 If the orgid is not a numeric value it is set to None, then it will
# be rejected anyway with the exit status: ERROR: The following ORGIDS are
# incorrect: [None], aborting
#
# Revision 1.17  2010/12/01 22:50:01  djw
# MA00061 Reject input data where systemid is not valid for the orgid as per maximo.locsystem table
#
# Revision 1.16  2010/11/26 14:46:58  djw
# MA00047 Now must supply one of -u -f -t as arg.
# Processing added for Full Load runs that can be
# batched in a script and will abort if MEA does
# not load all entities after a configurable
# interval mea_wait_interval
#
# Revision 1.15  2010/11/25 13:00:36  djw
# MA00051 Do not allow any invalid GL accounts
#
# Revision 1.14  2010/11/24 13:36:57  djw
# MA00049 Enhanced so that non-numeric or blank value in the level
# column A can be handled without a crash and with a meaningful
# error message in the log
#
# Revision 1.13  2010/11/23 13:51:51  djw
# MA00048 When processing spreadsheets, prompt user for target password,
# if not in parameters file.
#
# Revision 1.12  2010/11/17 15:04:52  djw
# 6943 InterfaceableAssetFromXls orgid property should be integer
#
# Revision 1.11  2010/11/14 04:55:54  djw
# 6924 Enhanced so that the different hierarchies themselves are processed in
# batches from highest to lowest , thus level switches occur less frequently
#
# Revision 1.10  2010/11/14 00:38:23  djw
# 6924 . Correct ref 6824 to 6924
#
# Revision 1.9  2010/11/14 00:31:50  djw
# 6924 Intermediate. Remove extra Oracle connections , pass logger handle
# to base.common.is_seg_in_org(). Error handling made consistent. Improved
# debugging at level 10.
#
# Revision 1.8  2010/10/27 06:18:43  njm3
# Made sys.exit print text more clear.
#
# Revision 1.4  2010/09/09 07:15:39  djw
# 6261 Now able to read source data from spreadsheet as well as Oracle database
#
# Revision 1.3  2010/04/14 13:20:13  djw
# Update comment
#
# Revision 1.2  2010/04/14 13:16:18  djw
# MA00036 Change driving query to 'start with' top-level DEPARTMENTS
#
# Revision 1.1  2010/04/13 13:06:30  djw
# MA00036. Initial revision
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

from common import get_all_segments
# Import for validation purposes
from common import read_orgids, read_depts, read_systemids, is_seg_in_org, is_systemid_in_org, get_glaccount_five_dg


########################################################################################
#
#     Class definitions
#
########################################################################################

class InterfaceableLocationFromOra:

    """Class to represent a Location to be loaded by MEA from a resultset"""

    def __init__(self, location_resultset):

        self.level = location_resultset[0]
        self.parent = location_resultset[1]
        self.location = location_resultset[2]
        self.description = location_resultset[3]
        self.type = location_resultset[4]
        self.glaccount = location_resultset[5]
        self.siteid = location_resultset[6]
        self.orgid = location_resultset[7]
        self.status = location_resultset[8]
        self.systemid = location_resultset[9]

        self.errmsg = None
        self.severity = None

    def is_valid(self, c, logger):

        # Test for exceptions. An ERROR will cause the whole run to abort,
        # rolling back the transaction

        # The parent must already be an existing LOCATION, or have just been
        # written to the interface table

        logger.debug("Checking parent %s of location %s exists as a location" % (self.parent, self.location))

        q_exists_loc="""
            SELECT location
            FROM maximo.locations l
            WHERE l.status = 'OPERATING'
            AND l.location = :location"""
        c.execute(q_exists_loc, location=self.parent)
        parent_loc = c.fetchone()
        logger.debug("Location %s looked up: %s" % (self.parent, parent_loc))

        if (not parent_loc):
            logger.debug("%s is not a valid existing Location, will check interface table" % (self.parent))

            logger.debug("Checking parent %s of location %s exists on interface table" % (self.parent, self.location))

            q_exists_interface="""
                SELECT location
                FROM maximo.mxoperloc_iface i
                WHERE i.location = :location"""
            c.execute(q_exists_interface, location=self.parent)
            iface_loc = c.fetchone()
            logger.debug("Location %s looked up: %s" % (self.parent, iface_loc))

            if (not iface_loc):
                logger.debug("%s is not a Location previously written to interface table" % (self.parent))
                return True

        return True

class InterfaceableLocationFromXls:

    """Class to represent a Location to be loaded by MEA from a spreadsheet"""

    def __init__(self, worksheet, i, logger):

        self.location = str(worksheet.cell(i, 2).value)
        # MA00049 If user does not provide an integer value for Level we catch the ValueError
        #         and set level to None, it will be rejected by the is_valid() method
        try:
            self.level = int(worksheet.cell(i, 0).value)
        except ValueError:
            logger.debug("Location %s has non-numeric level \"%s\", setting to None" % (self.location, worksheet.cell(i, 0).value))
            self.level = None

        self.parent = str(worksheet.cell(i, 1).value)
        self.description = worksheet.cell(i, 3).value
        self.type = worksheet.cell(i, 4).value
        self.glaccount = get_all_segments(worksheet.cell(i, 5).value)
        self.segment2 =  get_glaccount_five_dg(worksheet.cell(i, 5).value)   # Added for SEG2 validation purposes.
        self.siteid = worksheet.cell(i, 6).value
        try: # MA00063 If non-numeric orgid, set to None
            self.orgid = int(worksheet.cell(i, 7).value) # 6943 should be integer
        except ValueError:
            logger.debug("Location %s has non-numeric orgid \"%s\", setting to None" % (self.location, worksheet.cell(i, 7).value))
            self.orgid = None

        self.status = worksheet.cell(i, 8).value
        self.systemid = str(worksheet.cell(i, 9).value) # MA00061 cast to str

        self.errmsg = None
        self.severity = None

    def is_valid(self, c, logger):

        # Test for exceptions. An ERROR will cause the whole run to abort,
        # rolling back the transaction

        # MA00049 The Level must be an integer

        logger.debug("Checking level \"%s\" of location %s is integer" % (self.level, self.location))
        if type(self.level) is not int:
            logger.error("Location %s has non-integer level, check source data" % (self.location))
            return False

        # The parent must already be an existing LOCATION, or have just been
        # written to the interface table

        logger.debug("Checking parent %s of location %s exists as a location" % (self.parent, self.location))

        q_exists_loc="""
            SELECT location
            FROM maximo.locations l
            WHERE l.status = 'OPERATING'
            AND l.location = :location"""
        c.execute(q_exists_loc, location=self.parent)
        parent_loc = c.fetchone()
        logger.debug("Location %s looked up: %s" % (self.parent, parent_loc))

        if (not parent_loc):
            logger.debug("%s is not a valid existing Location, will check interface table" % (self.parent))

            logger.debug("Checking parent %s of location %s exists on interface table" % (self.parent, self.location))

            q_exists_interface="""
                SELECT location
                FROM maximo.mxoperloc_iface i
                WHERE i.location = :location"""
            c.execute(q_exists_interface, location=self.parent)
            iface_loc = c.fetchone()
            logger.debug("Location %s looked up: %s" % (self.parent, iface_loc))

            if (not iface_loc):
                logger.debug("%s is not a Location previously written to interface table" % (self.parent))
            return True

        return True

########################################################################################
#
#     Functions
#
########################################################################################

def get_next_txn_id(c, logger):

    # Get the next available transaction id for a MEA batch
    q_get_id = """
    SELECT maximo.hul_mealie_loc_seq.NEXTVAL
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
    AND ifacename = 'MXOPERLOCInterface'"""

    time.sleep(5)  # t = seconds
    c.execute(q_get_ct)
    wait_ct = c.fetchone()
    logger.info("MEA has %s unprocessed batches in the queue" % (wait_ct[0]))

    return wait_ct[0]

def get_entity_ct(c, logger):

    # Get the number of entities in Maximo
    q_get_ct = """
    SELECT COUNT(*)
    FROM maximo.locations"""

    c.execute(q_get_ct)
    entity_ct = int(c.fetchone()[0])
    logger.info("There are %d LOC entities in Maximo" % (entity_ct))

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
               'MXOPERLOCInterface',
               :transid)"""

    # Process each entry in the queue until it is empty

    for _i in range(len(q)):
        j = q.pop(0)
        logger.debug("Popped txn id %s. Queue now holds %s" % (j, q))
        c.execute(q_write_queue, transid = j)

    return True


def loc_handler(er):

    type_handler = {
        "ORA": loc_handler_ora_ora,
        "XLS": loc_handler_xls_ora
    }

    return type_handler.get(er.source_type)(er) # Call appropriate handler for the source type

def loc_handler_xls_ora(er):

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
    depts = read_depts(conn_target)
    er.logger.debug("Valid segment 2 values obtained: These are %s" % (depts))
    # MA00061 Now validate systemid against orgid (table locsystem)
    systemids = read_systemids(conn_target)
    er.logger.debug("Valid systemid values obtained: These are %s" % (systemids))

    levels = []        # A list of the levels in this spreadsheet

    invalid_orgids = []
    invalid_seg2 = []
    invalid_systemid = []

    seg2_err_ct = 0
    orgids_err_ct = 0
    systemids_err_ct = 0

    start_row = 1      # First row in spreadsheet that contains data (first is 0)

    er.logger.debug("Validating spreadsheet data")
    for x in range(start_row, wsh.nrows):

        il = InterfaceableLocationFromXls(wsh, x, er.logger)
        er.logger.debug("Instantiated InterfaceableLocationFromXls for %s" % (il.location))

        # MA00049 InterfaceableLocationFromXls.level is supposed to be an integer so no need to cast
        if (il.level not in levels):
            levels.append(il.level)

        if (il.orgid not in orgids):        # Validate ORGIDS
            if (il.orgid not in invalid_orgids): # Don't report duplicate error values
                invalid_orgids.append(il.orgid)
            orgids_err_ct += 1

        if is_seg_in_org(il.segment2, il.orgid, depts, er.logger) == False: # Validate SEGMENT 2
            invalid_seg2.append(il.segment2)
            seg2_err_ct += 1

        if is_systemid_in_org(il.systemid, il.orgid, systemids, er.logger) == False: # MA00061 Validate systemid
            if il.systemid not in invalid_systemid:
                invalid_systemid.append(il.systemid)
            systemids_err_ct += 1

    if orgids_err_ct > 0:
        # 6924 DJW Update for consistency with existing fatal errors
        er.logger.error("ERROR: The following ORGIDS are incorrect: %s, aborting" % invalid_orgids)
        print("ERROR: The following ORGIDS are incorrect: %s, aborting" % invalid_orgids)
        conn_target.rollback()
        conn_target.close() # 6924 DJW Free the connection
        sys.exit(2)

    if seg2_err_ct > 0:   # MA00051 Do not allow any invalid GL accounts
        # 6924 DJW Update for consistency with existing fatal errors
        er.logger.error("ERROR: The following GLACCOUNTS are incorrect: %s, aborting" % invalid_seg2)
        print("ERROR: The following GLACCOUNTS are incorrect: %s, aborting" % invalid_seg2)
        conn_target.rollback()
        conn_target.close() # 6924 DJW Free the connection
        sys.exit(2)

    if systemids_err_ct > 0:   # MA00061 Do not allow invalid systemid for the org
        er.logger.error("ERROR: %d systemids are incorrect: %s, aborting" % (systemids_err_ct, invalid_systemid))
        print("ERROR: %d systemids are incorrect: %s, aborting" % (systemids_err_ct, invalid_systemid))
        conn_target.rollback()
        conn_target.close() # Free the connection
        sys.exit(2)

    er.logger.debug("Spreadsheet data validated")

    levels.sort()
    er.logger.info("Levels to be processed: %s" % (levels))
  
    # Open various cursors on target database 

    curs_target_mea = conn_target.cursor()
    curs_target_lookup = conn_target.cursor()
    curs_target_seq = conn_target.cursor()

    # MA00047 Determine how many locations exist before we start
    start_entity_ct = get_entity_ct(curs_target_lookup, er.logger)

    # Insert a row into MEA interface for entity type LOCATION
    # 8093 Set new column status date to current date

    q_write_interface = """
    INSERT INTO maximo.mxoperloc_iface (
        location,
        description,
        type,
        glaccount,
        siteid,
        orgid,
        parent,
        status,
        lo13,
        systemid,
        statusdate,
        transid,
        transseq)
    VALUES (
        :location,
        :description,
        :type,
        :glaccount,
        :siteid,
        :orgid,
        :parent,
        :status,
        0,
        :systemid,
        SYSDATE,
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

            il = InterfaceableLocationFromXls(wsh, this_row, er.logger)

            # 6924 DJW Ignore unless it is the level we want to process on this iteration
            er.logger.debug("We are processing level %s, level is %s" % (level_to_process, il.level))
            if (il.level == level_to_process):
                er.logger.info("Processing location %s" % (il.location))
                if (il.is_valid(curs_target_lookup, er.logger)):
                    if (il.level != last_level):
                        er.logger.debug("LEVEL has changed: was %s, now %s" % (last_level, il.level))

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
    
                        last_level = il.level
    
                    else:
                        er.logger.debug("LEVEL is still %s" % (il.level))
    
                    er.logger.debug("Getting new transaction id for %s" % il.location)

                    # Obtain new MEA transaction id for next batch 
                    trans_id = get_next_txn_id(curs_target_seq, er.logger) 

                    # Add the MEA transaction id to the queue 
                    batch_queue.append(trans_id) 

                    # Sequence is reset for new batch
                    trans_seq = 1

                    # Write the MEA interface table
                    er.logger.debug("Writing interface: location %s transid %s transseq %s" % (il.location, trans_id, trans_seq))
                    curs_target_mea.execute(q_write_interface, location=il.location,
                                                                            description=il.description,
                                                                            type=il.type,
                                                                            glaccount=il.glaccount,
                                                                            siteid=il.siteid,
                                                                            orgid=il.orgid,
                                                                            parent=il.parent,
                                                                            status=il.status,
                                                                            systemid=il.systemid,
                                                                            transid=trans_id,
                                                                            transseq=trans_seq)
    
                else:
    
                    er.logger.error("ERROR: %s is not a valid Location, aborting" % (il.parent))
                    print("ERROR: %s is not a valid Location, aborting" % (il.parent))
                    conn_target.rollback()
                    conn_target.close() # 6924 DJW Free the connection
                    sys.exit(2)         # Will never happen if driving query is correctly ordered
   
                total_ct += 1
                level_ct += 1

            
                # 6924 DJW End of condition that it is the level we want to process on this iteration

            # End of location processing loop

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
        # MA00047 Determine how many locations exist after all processing is done
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


def loc_handler_ora_ora(er):

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

    # Driving query for entity type LOCATION

    q_driver_by_level = """
    SELECT LEVEL,
              h.parent,
              l.location,
              l.description,
              l.type,
              l.glaccount,
              l.siteid,
              l.orgid,
              l.status,
              h.systemid 
    FROM maximo.locations l,
           maximo.lochierarchy h
    WHERE l.location = h.location 
    AND l.location LIKE :locations
    AND LEVEL = :tree_level
    CONNECT BY PRIOR l.location = h.parent     
    START WITH l.location = (
        SELECT j.parent
        FROM maximo.lochierarchy j
        WHERE j.location = :location 
        ) 
    ORDER BY LEVEL, h.parent, l.location"""

    q_driver_all = """
    SELECT DISTINCT LEVEL,
                          h.parent,
                          l.location,
                          l.description,
                          l.type,
                          l.glaccount,
                          l.siteid,
                          l.orgid,
                          l.status,
                          h.systemid 
    FROM maximo.locations l,
           maximo.lochierarchy h
    WHERE l.location = h.location 
    AND l.location LIKE :locations
    CONNECT BY PRIOR l.location = h.parent     
    START WITH l.location = 'DEPARTMENTS'
    ORDER BY LEVEL, h.parent, l.location"""

    # Insert a row into MEA interface for entity type LOCATION
    # 8093 Set new column status date to current date

    q_write_interface = """
    INSERT INTO maximo.mxoperloc_iface (
        location,
        description,
        type,
        glaccount,
        siteid,
        orgid,
        parent,
        status,
        systemid,
        statusdate,
        transid,
        transseq)
    VALUES (
        :location,
        :description,
        :type,
        :glaccount,
        :siteid,
        :orgid,
        :parent,
        :status,
        :systemid,
        SYSDATE,
        :transid,
        :transseq)"""


    # Fetch first entity

    if (er.entity_level):
        curs_source_drv.execute (q_driver_by_level, location = er.entity_key,
                                                                   locations = er.entity_key + '%',
                                                                   tree_level = er.entity_level)
    else:
        curs_source_drv.execute (q_driver_all, locations = er.entity_key + '%')

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
        il = InterfaceableLocationFromOra(row)
        er.logger.info("Processing location %s" % (il.location))
        if (il.is_valid(curs_target_lookup, er.logger)):


            if (il.level != last_level):
                er.logger.debug("LEVEL has changed: was %s, now %s" % (last_level, il.level))

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

                last_level = il.level

            else:
                er.logger.debug("LEVEL is still %s" % (il.level))

            er.logger.debug("Getting new transaction id for %s" % il.location)

            # Obtain new MEA transaction id for next batch 
            trans_id = get_next_txn_id(curs_target_seq, er.logger)

            # Add the MEA transaction id to the queue 
            batch_queue.append(trans_id) 

            # Sequence is reset for new batch
            trans_seq = 1

            # Write the MEA interface table
            er.logger.debug("Writing interface: location %s transid %s transseq %s" % (il.location, trans_id, trans_seq))
            curs_target_mea.execute(q_write_interface, location=il.location,
                                                                      description=il.description,
                                                                      type=il.type,
                                                                      glaccount=il.glaccount,
                                                                      siteid=il.siteid,
                                                                      orgid=il.orgid,
                                                                      parent=il.parent,
                                                                      status=il.status,
                                                                      systemid=il.systemid,
                                                                      transid=trans_id,
                                                                      transseq=trans_seq)

        else:

            er.logger.error("ERROR: %s is not a valid Location, aborting" % (il.parent))
            print("ERROR: %s is not a valid Location, aborting" % (il.parent))
            conn_target.rollback()
            conn_target.close() # 6924 DJW Free the connection
            sys.exit(2)         # This ugliness will never happen if driving query is correctly ordered
    
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


