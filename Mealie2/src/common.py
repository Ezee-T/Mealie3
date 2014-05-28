#!/usr/bin/env python
#
# $Header: /usr/local/cvsroot/mealie/common.py,v 1.4 2012/02/10 14:07:50 djw Exp $
#
# MEA Lightweight Integration Environment
#
# $Log: common.py,v $
# Revision 1.4  2012/02/10 14:07:50  djw
# 8112 Handling for ICG (Item Commodity)
#
# Revision 1.3  2012/02/10 10:00:54  djw
# 8112 Handling COM (Commodities) entity. Unit tested on MAXSIT
# See maximo/8112/hul_test_COM.stg and hul_test_item.xls
#
# Revision 1.2  2012/02/09 08:28:40  djw
# Add COM UOM UND to usage
#
# Revision 1.1  2012/01/05 12:21:54  djw
# Migrated to new project
#
# Revision 1.18  2011/11/18 13:51:42  djw
# 8089 GL Code can be old style (numeric) in which case Rolled Products
# is assumed and seg1 is set to 'R'.  Otherwise GL Code can be new style where
# the 'R' or 'E' prefix is specified explicitly, for example the GL Code E1699
# would be interpreted as E-01699-00-????????
#
# Revision 1.17  2011/04/15 15:16:32  djw
# 7591 Clarify usage output
#
# Revision 1.16  2011/04/15 12:50:35  djw
# 7591 Add stager enhancement
#
# Revision 1.15  2011/04/08 13:00:41  djw
# 7474 Final minor changes for ASP entity
#
# Revision 1.14  2011/03/28 12:59:31  djw
# 7474 Add ASP entity
#
# Revision 1.13  2010/12/29 13:59:46  djw
# MA00062 get_next_txn_id() moved to the pertinent entity module
#
# Revision 1.12  2010/12/01 22:50:49  djw
# MA00061 Reject input data where systemid is not valid for the orgid as per maximo.locsystem table
#
# Revision 1.11  2010/11/26 12:24:46  djw
# MA00047 Now must supply one of -u -f -t as arg.
# Processing added for Full Load runs that can be
# batched in a script and will abort if MEA does
# not load all entities after a configurable
# interval mea_wait_interval
#
# Revision 1.10  2010/11/14 00:37:17  djw
# 6924 . Correct ref 6824 to 6924
#
# Revision 1.9  2010/11/14 00:34:18  djw
# 6924 Intermediate. Remove extra Oracle connections , pass logger handle
# to is_seg_in_org().
#
# Revision 1.8  2010/11/10 14:26:01  njm3
# [6789] Output no longer prints to screen, but rather to DEBUG.
#
# Revision 1.7  2010/10/27 06:12:16  njm3
# Added 5-digit check for seg2 on get_glaccount_five_dg function.
#
# Revision 1.3  2010/09/09 07:11:39  djw
# 6261 Now able to read source data from spreadsheet as well as Oracle database
#
# Revision 1.2  2010/04/14 11:08:21  djw
# Ma00036. Validate entity type
#
# Revision 1.1  2010/04/13 13:05:20  djw
# MA00036. Initial revision
#
#
#
#

########################################################################################
#
#     Import required modules
#
########################################################################################

########################################################################################
#
#     Globals
#
########################################################################################

glprefixes=['E','R']

########################################################################################
#
#     Class definitions
#
########################################################################################

class UNSPSCDictionaryEntry:

    """8112 Class representing a UNSPSC dictionary entry (segment, family, class or commodity"""

    def __init__(self, code, c, logger):
        
        q_get_entry = """
        SELECT code, codetype, description, definition
        FROM maximo.hul_unspsc_dict
        WHERE description = :code"""
        
        c.execute(q_get_entry, code=code)
        resultset = c.fetchone()
        if resultset == None:
            self.code = None
            self.codetype = None
            self.description = None
            self.definition = None
        else:
            self.code = resultset[0]
            self.codetype = resultset[1]
            self.description = resultset[2]
            self.definition = resultset[3]


########################################################################################
#
#     Functions
#
########################################################################################

def usage():

    print "Usage: mealie [-u|-t|-f] -p pfile -e entity [-k key] [-l level] | -V"
    print "       -e entity  Process entity <entity>"
    print "                  ASP - Asset specification entities"
    print "                  ASS - Asset entities"
    print "                  COM - Commodity Group entities"
    print "                  ICG - Item Commodity entities"
    print "                  LOC - Location entities"
    print "                  UND - UNSPSC dictionary entry entities"
    print "                  UOM - Unit of Measure entities"
    print "       -k key     Process entity members with key <key>"
    print "                  (Only specify if the source is a database)"
    print "       -p pfile   Get parameters from parameter file <pfile>"
    print "       -l level   Only process entities at LEVEL <level> in the tree"
    print "       -u         Normal mode (validate source and load target database)"
    print "       -t         Test mode (validate source data but do not load target database)"
    print "       -f         Full Load mode (validate source date, all items must be newly loaded)"
    print "       -V         Show MEALIE version"

def stgusage():

    print "Usage: mealie_stg -s stage -u username -p password -d database [-l severity-level]"
    print "       -s stage            Process stage file \"<stage>.stg\""
    print "       -u username         Connect to target database as user <username>"
    print "       -p password         Password for user <username>"
    print "       -d database         Connect to target database instance <database>"
    print "       -l severity-level   Logging severity level. 10=debug, 20=info, 40=error"

def get_all_segments(seg2):

    # Take the second segment of a GL account and return the
    # full account number

    seg1 = get_glaccount_seg1(seg2)
    seg2 = get_glaccount_five_dg(seg2)
    seg3 = '00'
    seg4 = '????????'

    #print seg2
    #print type(seg2)
    #print (type(seg2) == type('str'))

    if seg2 == None:
        retval = None
    elif (type(seg2) == type('str') and len(seg2) == 0):
        retval = None
    else:
        retval = seg1 + '-' + str(int(seg2)).rjust(5,'0') + '-' + seg3 + '-' + seg4

    return retval

# If the GLACCOUNT is 4 digits, prepends a 0 to the GLACCOUNT to make it valid
# Otherwise, it returns the same 5 digit value in retval.

def get_glaccount_five_dg(seg2):
    
    padchar = '0'

    if str(seg2)[0] in glprefixes:
        seg2 = seg2[1:]
    
    if seg2 == None:
        retval = None
    elif (type(seg2) == type('str') and len(seg2) == 0):
        retval = None
    elif len(str(seg2)) == 5:     # Check if the GLACCOUNT integer is already 5 digits
        retval = seg2              # If it is, return it as-is, without the additional 0
    else:
        retval = padchar + str(int(seg2))
    
    return retval

# 8089. 18-NOV-2011. seg2 can be old style (numeric) in which case Rolled Products is assumed and
# seg1 is set to 'R'.
# Otherwise seg2 can be new style where the 'R' or 'E' prefix is specified explicitly

def get_glaccount_seg1(seg2):

    retval = 'R'    # Backward compatibility. This was previously the default

    if str(seg2)[0] in glprefixes:
        retval = seg2[0]

    return retval

# 8112. 09-Feb-2012. A UNSPSC Class Code nnnnnn00 can be derived from a UNSPSC Commodity Code nnnnnnnn 

def get_unspsc_class(commcode):

    classcode = int(commcode/100) * 100  
    return classcode


# Enable validation on xls files off ORGID and GLACCOUNT

# Get orgids from Oracle and use them to populate a list
# 6924 DJW Remove hard coded connection and pass connection object
def read_orgids(conn):
    orgids = []
    curs = conn.cursor()
    curs.execute('SELECT orgid FROM maximo.organization')
    row = curs.fetchone()
    while row:
        (orgid) = (row[0])
        orgids.append(int(orgid))
        row = curs.fetchone()
    return orgids

# MA00061 Get systemids f
r systemid verification
def read_systemids(conn):
    systemids = {}
    curs = conn.cursor()
    query = """SELECT systemid, TO_NUMBER(orgid) orgid
                   FROM maximo.locsystem"""
    curs.execute(query)
    row = curs.fetchone()
    while row:
        (systemid, orgid) = (row[0], row[1])
        if systemid in systemids:
            systemids[systemid].append(orgid)
        else:
            systemids[systemid] = [int(orgid)]
        row = curs.fetchone()
    return systemids

# Get departments for second segment verification
# 6924 DJW Remove hard coded connection and pass connection object
def read_depts(conn):
    depts = {}
    curs = conn.cursor()
    query = """SELECT compvalue, TO_NUMBER(orgid) orgid
                    FROM maximo.glcomponents 
                    WHERE glorder = 1
                    AND LENGTH(compvalue) = 5"""
    curs.execute(query)
    row = curs.fetchone()
    while row:
        (dept, orgid) = (row[0], row[1])
        if dept in depts:
            depts[dept].append(orgid)
        else:
            depts[dept] = [int(orgid)]
        row = curs.fetchone()
    return depts

def is_systemid_in_org(systemid, orgid, systemids, logger):

    # MA00061 Determine whether systemid is valid for the given
    # orgid as per the maximo.locsystem table and return boolean

    if systemid in systemids:
        orgids = systemids[systemid]
        if orgid in orgids:
            logger.debug("Systemid %s is valid for orgid %d" % (systemid, orgid))
            retval = True
        else:
            logger.debug("Systemid %s is not valid for orgid %d" % (systemid, orgid))
            retval = False
    else:
        logger.debug("Systemid %s is not valid for any orgid" % (systemid))
        retval = False
    return retval

def is_seg_in_org(seg2, orgid, depts, logger):

    # Function to determine whether seg2 is a valid dept component of 
    # the account number for a given orgid

    if seg2 in depts:
        orgids = depts[seg2]
        if orgid in orgids:
            # It may be worth while removing the following print line for readability...
            logger.debug("Segment #2 %s is a valid department in orgid %d" % (seg2, orgid))
            retval = True
        else:
            logger.debug("Segment #2 %s is not a valid department in orgid %d" % (seg2, orgid))
            retval = False
    else:
        logger.debug("Segment #2 %s is not a valid department in any orgid" % (seg2))
        retval = False
    return retval

def commodity_exists(commodity, c, logger):

    # 8112 10-Feb-2012. Verify the existence of attribute attr for classification cls

    q_get_exists = """
    SELECT 1
    FROM maximo.commodities
    WHERE commodity = TO_CHAR(:commodity)"""

    logger.debug("commodity_exists is checking the existence of commodity %s" % (commodity))
    c.execute(q_get_exists, commodity=commodity)
    result = c.fetchone()
    if result == None:
        logger.debug("Lookup q_get_exists failed in commodity_exists(): commodity %s" % (commodity))
        retval = False
    else:
        retval = True
    return retval

def update_attr_datatype(c, attr, logger):
    
    q_update_attr_datatype = """UPDATE assetattribute 
    SET datatype = 'ALN'
    WHERE assetattrid = :assetattrid"""
    
    c.execute(q_update_attr_datatype, assetattrid=attr)
    
    logger.debug("Changed Attribute datatype to ALN data-type")
      
    return

def specification_import():
    
    # Verify the existence and 
    
    
    
    
    return
    

