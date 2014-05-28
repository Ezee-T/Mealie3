#/bin/env python
#
# $Header: /usr/local/cvsroot/mealie/entities/assetspecs.py,v 1.1 2012/01/05 12:22:34 djw Exp $
#
# MEA Lightweight Integration Environment
#
# Entity Handler for ASSET SPECIFICATIONS
#
# Change Ref 7474
#
# $Log: assetspecs.py,v $
# Revision 1.1  2012/01/05 12:22:34  djw
# Migrated to new project
#
# Revision 1.17  2011/07/07 08:34:06  djw
# 7890 Fix to prevent the rounding down to the nearest integer of floating
# point values where the Maximo attribute datatype is ALN (Alphanumeric)
#
# Revision 1.16  2011/07/05 08:45:12  djw
# 7882 Verify the existence of the attribute for the classification
# in question. This validation is performed even when the value is
# blank. Add new attrid_exists() function
#
# Revision 1.15  2011/07/05 06:16:28  djw
# 7880 If the attribute is mandatory for an asset then the user must
# supply a value in the spreadsheet file
#
# Revision 1.14  2011/06/23 13:03:54  djw
# 7839 Enhanced to correctly process rows in the input file that have no attributes
# whatsoever. The correct behaviour is not to create a queue row on mxin_inter_trans
# since there are no corresponding rows in the interface table mxassetspec_iface
#
# Revision 1.13  2011/06/17 08:12:05  djw
# 7824 If the attribute is an ALN but happens to have a numeric value
# in this particular case, e.g. a serial number, then cast the
# value to integer to ensure that we do not end up with a string
# value ending in '.0', i.e. a float
#
# Revision 1.12  2011/06/10 11:03:58  djw
# 7796 The classificationid must be one that is intended to be used
# with assets. See checkbox on Classifications screen
#
# Revision 1.11  2011/06/07 10:51:22  djw
# 7793 Validate upper permitted values for ALN attribute values
#      assetspec.alnvalue is a varchar2(100) so 100 is the maximum
#      string length allowed
#
# Revision 1.10  2011/05/26 06:42:26  djw
# 7703 Enhanced to cope with a classificationid that does not
# exist on the maximo.classstructure table
#
# Revision 1.9  2011/05/16 08:49:33  djw
# 7674 Domain-constrained attributes are never allowed
#
# Revision 1.8  2011/04/19 08:55:49  djw
# 7624 Parents assets and locations and asset locations must be OPERATING
#
# Revision 1.7  2011/04/15 15:44:51  djw
# 7591 Error handling for IO error when attempting to open input file
#
# Revision 1.6  2011/04/15 15:19:59  djw
# 7591 Propogate return status from handler method
#
# Revision 1.5  2011/04/04 13:33:49  djw
# 7474 Unit test complete
#
# Revision 1.4  2011/04/01 11:59:14  djw
# 7474 Now writes new MEA trans id for each asset. Also tests that attributes
# designated on hul_mealie_attribute as NUMERIC are valid float values
#
# Revision 1.3  2011/03/31 10:15:30  djw
# 7474 Bug fix. Fixed so that duplicate rows are not being written to MEA
# queue table
#
# Revision 1.2  2011/03/30 09:00:55  djw
# 7474 First revision for unit test
#
# Revision 1.1  2011/03/29 07:06:02  djw
# 7474 Interim revision
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

#from base.common import read_orgids

########################################################################################
#
#     Class definitions
#
########################################################################################

class InterfaceableAssetspecFromXls:

    """Class representing a set of Asset Specifications to be loaded by MEA from a spreadsheet"""

    def __init__(self, worksheet, i, c, logger):

        self.assetnum = worksheet.cell(i, 0).value
        self.classificationid = worksheet.cell(i, 1).value

        logger.debug("Getting assetid for asset %s" % (self.assetnum))
        q_assid="""
            SELECT assetid
            FROM maximo.asset a
            WHERE a.status = 'OPERATING'
            AND a.assetnum = :assetnum"""
        c.execute(q_assid, assetnum=self.assetnum)
        this_assid = c.fetchone()
        logger.debug("Assetid for asset %s looked up: %s" % (self.assetnum, this_assid))
        if (this_assid == None):
            self.assetuid = None
        else:
            self.assetuid = this_assid[0]

        logger.debug("Getting classstructureid for classificationid %s" % (self.classificationid))
        q_classid="""
            SELECT classstructureid, usewithassets
            FROM maximo.classstructure c
            WHERE c.classificationid = :classificationid"""
        c.execute(q_classid, classificationid=self.classificationid)
        this_classid = c.fetchone()
        logger.debug("classstructureid for asset %s looked up: %s" % (self.assetnum, this_classid))
        if (this_classid == None):
            self.classstructureid = None
            self.usewithassets = 0
        else:
            self.classstructureid = this_classid[0]
            self.usewithassets = this_classid[1]

        self.attr_val = []

        for j in range(2, worksheet.ncols):
            if (worksheet.cell(i, j).ctype == xlrd.XL_CELL_NUMBER) and (worksheet.cell(i, j).value == int(worksheet.cell(i, j).value)):
                self.attr_val.append(int(worksheet.cell(i, j).value)) # 7824 Cast to integer to ensure no trailing '.0'
            else:
                self.attr_val.append(worksheet.cell(i, j).value)
            logger.debug("Loaded attribute value %s" % worksheet.cell(i, j).value)

        self.errmsg = None
        self.severity = None

    def is_valid(self, c, logger):

        # Test for exceptions. An ERROR will cause the whole run to abort,
        # rolling back the transaction

        # The assetnum must be a valid Asset

        logger.debug("Checking asset %s exists" % (self.assetnum))

        q_exists_ass="""
            SELECT assetnum
            FROM maximo.asset a
            WHERE a.status = 'OPERATING'
            AND a.assetnum = :assetnum"""
        c.execute(q_exists_ass, assetnum=self.assetnum)
        this_ass = c.fetchone()
        logger.debug("Asset %s looked up: %s" % (self.assetnum, this_ass))
        if (this_ass == None):
            logger.error("Asset %s does not exist" % (self.assetnum))
            print("ERROR: Asset %s does not exist" % (self.assetnum))

        if (not this_ass):
            return False

        # The classificationid must be a designated loadable classification
        # on the HUL_MEALIE_CLASSIFICATION table

        logger.debug("Checking classification %s is legal" % (self.classificationid))

        q_exists_class="""
            SELECT classificationid
            FROM maximo.hul_mealie_classification c
            WHERE c.classificationid = :classificationid"""
        c.execute(q_exists_class, classificationid=self.classificationid)
        this_class = c.fetchone()
        logger.debug("Legal Classification %s looked up: %s" % (self.classificationid, this_class))

        if (not this_class):
            return False

        # 7796 The classificationid must be one that is intended to be used
        # with assets. See checkbox on Classifications screen

        logger.debug("Checking classification %s is Use With Assets" % (self.classificationid))
        if (self.usewithassets != 1):
            logger.error("Classification %s is not set to Use With Assets" % (self.classificationid))
            print("ERROR: Classification %s is not set to Use With Assets" % (self.classificationid))
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
    SELECT maximo.hul_mealie_asp_seq.NEXTVAL
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
    AND ifacename = 'MXASSETSPECInterface'"""

    time.sleep(5)       # t = seconds
    c.execute(q_get_ct)
    wait_ct = c.fetchone()
    logger.info("MEA has %s unprocessed batches in the queue" % (wait_ct[0]))

    return wait_ct[0]

def attrid_exists(conn, c, cls, attr, logger):

    # Verify the existence of attribute attr for classification cls

    q_get_exists = """
    SELECT 1
    FROM maximo.classstructure cst
    JOIN maximo.classspec csp ON csp.classstructureid = cst.classstructureid
    WHERE cst.classificationid = :classificationid
    AND csp.assetattrid = :assetattrid"""

    c.execute(q_get_exists, classificationid=cls, assetattrid=attr)
    result = c.fetchone()
    if result == None:
        logger.debug("Lookup q_get_exists failed in attrid_exists(): classificationid %s, assetattrid %s" % (cls, attr))
        retval = False
    else:
        retval = True
    return retval

def get_attr_domain_type(conn, c, cls, attr, logger):

    # Determine the domainid of attribute a of classification c

    q_get_type = """
    SELECT csp.domainid
    FROM maximo.classstructure cst
    JOIN maximo.classspec csp ON csp.classstructureid = cst.classstructureid
    WHERE cst.classificationid = :classificationid
    AND csp.assetattrid = :assetattrid"""

    c.execute(q_get_type, classificationid=cls, assetattrid=attr)
    data_type = c.fetchone()
    if not data_type:
        logger.error("Lookup failed in get_attr_domain_type(): classificationid %s, assetattrid %s, aborting" % (cls, attr))
        print("ERROR: Lookup failed in get_attr_domain_type(): classificationid %s, assetattrid %s, aborting" % (cls, attr))
        conn.rollback()
        conn.close()         # Free the connection
        sys.exit(2)
    return data_type[0]


def get_attr_assetrequirevalue(conn, c, cls, attr, logger):

    # Get the assetrequirevalue of attribute a of classification c

    q_get_type = """ SELECT csp.assetrequirevalue
                           FROM maximo.classstructure cst
                           JOIN maximo.classspec csp ON csp.classstructureid = cst.classstructureid
                           WHERE cst.classificationid = :classificationid AND csp.assetattrid = :assetattrid"""

    c.execute(q_get_type, classificationid=cls, assetattrid=attr)
    data_type = c.fetchone()
    if not data_type:
        logger.error("Lookup failed in get_attr_assetrequirevalue(): classificationid %s, assetattrid %s, aborting" % (cls, attr))
        print("ERROR: Lookup failed in get_attr_assetrequirevalue(): classificationid %s, assetattrid %s, aborting" % (cls, attr))
        conn.rollback()
        conn.close()
        sys.exit(2)
    return data_type[0]




def get_attr_type(c, cls, attr):

    # Determine the datatype of attribute a of classification c

    q_get_type = """
    SELECT a.datatype
    FROM maximo.hul_mealie_attribute a
    JOIN maximo.hul_mealie_classification c ON c.id = a.hmc_id
    WHERE a.assetattrid = :assetattrid
    AND c.classificationid = :classificationid
    AND entity_type = 'ASP'"""

    c.execute(q_get_type, classificationid=cls, assetattrid=attr)
    data_type = c.fetchone()

    if (data_type == None):   # Not on the MAXIMO.HUL_MEALIE_ATTRIBUTE table
        return None
    else:
        return data_type[0]

def write_mea_queue(c, txid, logger):

    # Insert row into MEA queue

    q_write_queue = """
    INSERT INTO maximo.mxin_inter_trans (
                           extsysname,
                           ifacename,
                           action,
                           transid
                           )
    VALUES ('EXTSYS1', 'MXASSETSPECInterface', 'AddChange', :trans_id)"""

    logger.debug("Writing MEA queue for transid %d" % txid)
    c.execute(q_write_queue, trans_id=txid)

    return True


def asp_handler(er):

    type_handler = {
        "XLS": asp_handler_xls_ora
    }

    return type_handler.get(er.source_type)(er)   # Call appropriate handler for the source type


def asp_handler_xls_ora(er):

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
    
    er.logger.debug("Validating spreadsheet data - header data")

    # Process the attribute names in the first row of the file

    attr_name = []       # Array to store attribute names
    er.logger.debug("Loading attribute names")

    for i in range(2, wsh.ncols):
        attr_name.append(wsh.cell(0, i).value)
        er.logger.debug("%d Loaded attribute %s" % (i, wsh.cell(0, i).value))

    er.logger.debug("Loaded a total of %d attributes" % len(attr_name))

    # Open various cursors on target database

    curs_target_mea = conn_target.cursor()
    curs_target_lookup = conn_target.cursor()
    curs_target_seq = conn_target.cursor()

    # Insert a row into MEA interface for entity type ASSET

    q_write_interface = """
    INSERT INTO maximo.mxassetspec_iface (
        assetnum,
        assetuid,
        classstructureid,
        s_classstructureid,
        as_orgid,
        s_orgid,
        as_siteid,
        alnvalue,
        numvalue,
        assetattrid,
        displaysequence,
        inheritedfromitem,
        itemspecvalchanged,
        transid,
        transseq)
    VALUES (
        :assetnum,
        :assetuid,
        :classstructureid,
        :classstructureid,
        '170',
        '170',
        'PMB',
        :alnvalue,
        :numvalue,
        :assetattrid,
        :displaysequence,
        0,
        0,
        :transid,
        :transseq)"""

    # Initialise local variables for fetch loop

    #queued_batches = 0 
    total_ct = 0       # Total no. of items

    start_row = 1      # The first row in spreadsheet that contains data (start at 0)

    # Iterate through all rows of data in input spreadsheet

    for this_row in range(start_row, wsh.nrows):

        ia = InterfaceableAssetspecFromXls(wsh, this_row, curs_target_seq, er.logger)
        er.logger.info("Processing assetnum %s" % (ia.assetnum))

        if (ia.is_valid(curs_target_lookup, er.logger)):

            # Obtain new MEA transaction id for next asset 

            er.logger.debug("Getting new transaction id for this load")
            trans_id = get_next_txn_id(curs_target_seq, er.logger)

            trans_seq = 0          # Sequence within batch
            display_seq = 0        # Display sequence starts at 1 for each assetnum

            attr_ct = 0            # Count of attributes for this asset

            # Write the MEA interface table

            for attr in range(len(attr_name)):

                er.logger.debug("Processing interface record for assetnum %s attribute %s" % (ia.assetnum, attr_name[attr]))

                if attrid_exists(conn_target, curs_target_lookup, ia.classificationid, attr_name[attr], er.logger):
                    pass
                else:
                    er.logger.error("Attribute %s for classificationid %s does not exist, aborting" % (attr_name[attr], ia.classificationid))
                    print("ERROR: Attribute %s for classificationid %s does not exist, aborting" % (attr_name[attr], ia.classificationid))
                    conn_target.rollback()
                    conn_target.close()         # Free the connection
                    sys.exit(2)

                if (ia.attr_val[attr] != ""):     # Only interface non-blank cells

                    trans_seq += 1      # Sequence is reset for new batch
                    display_seq += 1    # Display sequence is incremented for each attribute

                    er.logger.info("Writing interface: assetnum=%s transid=%s transseq=%s" % (ia.assetnum, trans_id, trans_seq))

                    er.logger.debug("assetuid=%s" % (ia.assetuid))
                    er.logger.debug("classstructureid=%s" % (ia.classstructureid))
                    er.logger.debug("alnvalue=%s" % (ia.attr_val[attr]))
                    er.logger.debug("assetattrid=%s" % (attr_name[attr]))
                    er.logger.debug("displaysequence=%s" % (display_seq))

                    # 7674 At the present time domain-constrained attributes are not
                    # allowed at all
                    attr_domain_type = get_attr_domain_type(conn_target, curs_target_lookup, ia.classificationid, attr_name[attr], er.logger)
                    er.logger.debug("Attribute domain type for attribute %s is %s" % (attr_name[attr], attr_domain_type))
                    if attr_domain_type is None:
                        pass
                    else:
                        er.logger.error("Domain %s is not allowed for attribute %s on asset %s, aborting" % (attr_domain_type, attr_name[attr], ia.assetnum))
                        print("ERROR: Invalid attribute type for asset %s, aborting" % (ia.assetnum))
                        conn_target.rollback()
                        conn_target.close()         # Free the connection
                        sys.exit(2)


                    data_type = get_attr_type(curs_target_lookup, ia.classificationid, attr_name[attr])
                    er.logger.debug("Attribute %s is type %s" % (attr_name[attr], data_type))

                    if (data_type == None):
                        er.logger.error("ERROR: Invalid attribute type for asset %s, aborting" % (ia.assetnum))
                        print("ERROR: Invalid attribute type for asset %s, aborting" % (ia.assetnum))
                        conn_target.rollback()
                        conn_target.close()         # Free the connection
                        sys.exit(2)

                    if (data_type == 'ALN'):
                        aln_value = unicode(ia.attr_val[attr]) # 7793 Could be float or int if numeric cell value
                        num_value = None
                    else:           
                        aln_value = None
                        num_value = ia.attr_val[attr]
                        try:
                            _num1 = float(num_value)
                        except ValueError:
                            er.logger.debug("Asset %s has non-numeric value %s for attribute %s" % (ia.assetnum, num_value, attr_name[attr]))
                            er.logger.error("ERROR: Specifications for asset %s are invalid, aborting" % (ia.assetnum))
                            print("ERROR: Specifications for asset %s are invalid, aborting" % (ia.assetnum))
                            conn_target.rollback()
                            conn_target.close()         # Free the connection
                            sys.exit(2)

                    # 7793 Validate upper permitted values for ALN attribute values
                    # assetspec.alnvalue is a varchar2(100)
                    if (data_type == 'ALN' and len(aln_value) > 100):
                        er.logger.error("Asset %s has a value which is %d characters long for attribute %s" % (ia.assetnum, len(aln_value), attr_name[attr]))
                        er.logger.error("Value of %s is: %s" % (attr_name[attr], aln_value))
                        er.logger.error("An alphanumeric value may not be more than 100 characters long")
                        er.logger.error("ERROR: Specifications for asset %s are invalid, aborting" % (ia.assetnum))
                        print("ERROR: Specifications for asset %s are invalid, aborting" % (ia.assetnum))
                        conn_target.rollback()
                        conn_target.close()         # Free the connection
                        sys.exit(2)


                    curs_target_mea.execute(q_write_interface, assetnum=ia.assetnum,
                                                                              assetuid=ia.assetuid,
                                                                              classstructureid=ia.classstructureid,
                                                                              alnvalue=aln_value,
                                                                              numvalue=num_value,
                                                                              assetattrid=attr_name[attr],
                                                                              displaysequence=display_seq,
                                                                              transid=trans_id,
                                                                              transseq=trans_seq)
                    er.logger.debug("Wrote interface")
                    attr_ct += 1
                    er.logger.debug("Increase attribute count for %s to %d" % (ia.assetnum, attr_ct))
                else:

                    # 7880 There is no value. Ensure that the attribute is not
                    # mandatory for an asset
                    attr_assetrequirevalue = get_attr_assetrequirevalue(conn_target, curs_target_lookup, ia.classificationid, attr_name[attr], er.logger)
                    er.logger.debug("Attribute %s is required for assets?: %s" % (attr_name[attr], attr_assetrequirevalue))
                    if attr_assetrequirevalue == 1:
                        er.logger.error("No value supplied for mandatory attribute %s for asset %s, aborting" % (attr_name[attr], ia.assetnum))
                        print("ERROR: No value supplied for mandatory attribute %s for asset %s, aborting" % (attr_name[attr], ia.assetnum))
                        conn_target.rollback()
                        conn_target.close()         # Free the connection
                        sys.exit(2)

                    er.logger.debug("There is an empty string value for this attribute, skipping")
        else:

            er.logger.error("ERROR: Specifications for asset %s are invalid, aborting" % (ia.assetnum))
            print("ERROR: Specifications for asset %s are invalid, aborting" % (ia.assetnum))
            conn_target.rollback()
            conn_target.close()         # Free the connection
            sys.exit(2)

        # Write the MEA queue
        er.logger.info("Interface %d attributes for %s" % (attr_ct, ia.assetnum))
        if attr_ct:
            write_mea_queue(curs_target_mea, trans_id, er.logger)
            er.logger.debug("Wrote MEA queue for transid %d" % trans_id)
        else:
            er.logger.debug("Did not write MEA queue for transid %d" % trans_id)
    
        total_ct += 1

    # (-) End of asset processing loop

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


