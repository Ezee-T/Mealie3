#!/usr/bin/env python
#
# $Header: /usr/local/cvsroot/mealie/launcher.py,v 1.3 2012/02/07 09:50:02 djw Exp $
#
# MEA Lightweight Integration Environment
#
# $Log: launcher.py,v $
# Revision 1.3  2012/02/07 09:50:02  djw
# 8112 Version 3.1.0 supports the various cataloguing entites
#
# Revision 1.2  2012/01/05 13:06:50  djw
# Update version number
#
# Revision 1.1  2012/01/05 12:21:54  djw
# Migrated to new project
#
# Revision 1.15  2011/11/23 06:31:39  djw
# 8093 Version 3.0.1 for Maximo 6.2.7
#
# Revision 1.14  2011/11/18 13:51:42  djw
# 8089 GL Code can be old style (numeric) in which case Rolled Products
# is assumed and seg1 is set to 'R'.  Otherwise GL Code can be new style where
# the 'R' or 'E' prefix is specified explicitly, for example the GL Code E1699
# would be interpreted as E-01699-00-????????
#
# Revision 1.13  2011/04/15 15:16:58  djw
# 7591 Typo in variable name
#
# Revision 1.12  2011/04/15 12:50:35  djw
# 7591 Add stager enhancement
#
# Revision 1.11  2011/04/13 13:12:05  djw
# 7592 Treat -f as if it were -u for Production because of real time updates
#
# Revision 1.10  2011/04/08 13:00:41  djw
# 7474 Final minor changes for ASP entity
#
# Revision 1.9  2011/03/28 12:59:31  djw
# 7474 Add ASP entity
#
# Revision 1.8  2010/12/29 12:55:07  djw
# MA00052. Give clear err msg if parameter file not found
#
# Revision 1.7  2010/11/26 12:24:46  djw
# MA00047 Now must supply one of -u -f -t as arg.
# Processing added for Full Load runs that can be
# batched in a script and will abort if MEA does
# not load all entities after a configurable
# interval mea_wait_interval
#
# Revision 1.6  2010/11/14 04:49:49  djw
# 6924. Performance improvement to be designated 1.1.1
#
# Revision 1.5  2010/10/26 12:38:14  njm3
# Changed mealie path - one dir higher:
#     e.g. mealie.base.x is now base.x
#     (Update MEALIE_PATH in ~/.bash_profile accordingly: /usr/local/lib/python/mealie)
#
# Revision 1.4  2010/09/09 07:11:39  djw
# 6261 Now able to read source data from spreadsheet as well as Oracle database
#
# Revision 1.3  2010/04/14 13:32:21  djw
# MA00036. Ready for testing, mealie_version 1.0.0
#
# Revision 1.2  2010/04/14 11:08:21  djw
# Ma00036. Validate entity type
#
# Revision 1.1  2010/04/13 13:05:20  djw
# MA00036. Initial revision
#
#

########################################################################################
#
#     Import required modules
#
########################################################################################

import sys
import getopt
import ConfigParser
import logging
import string
import os

from common import usage, stgusage

########################################################################################
#
#     Class definitions
#
########################################################################################

class EntityRun:
    """Class to represent an entity run to be processed"""

    def __init__(self, log_file, log_severity, logger,
                      source_db, source_user, source_pwd,
                      source_xlsfile,
                      source_type,
                      target_db, target_user, target_pwd,
                      entity_type, entity_key, entity_level,
                      mea_wait_interval,
                      normal_mode, test_mode, full_mode
                    ):

        self.log_file = log_file
        self.log_severity = log_severity
        self.logger = logger

        self.source_db = source_db
        self.source_user = source_user
        self.source_pwd = source_pwd

        self.source_xlsfile = source_xlsfile

        self.source_type = source_type

        self.target_db = target_db
        self.target_user = target_user
        self.target_pwd = target_pwd

        self.entity_type = entity_type
        self.entity_key = entity_key
        self.entity_level = entity_level

        self.mea_wait_interval = mea_wait_interval

        self.normal_mode = normal_mode
        self.test_mode = test_mode
        self.full_mode = full_mode

class StageBatch:
    """Class to represent a Stage to be processed"""

    def __init__(self, stage_file, log_severity,
                      target_db, target_user, target_pwd
                    ):

        self.stage_file = stage_file
        self.log_severity = log_severity

        self.target_db = target_db
        self.target_user = target_user
        self.target_pwd = target_pwd


########################################################################################
#
#     Functions
#
########################################################################################


########################################################################################
#
#     Main processing
#
########################################################################################

def main(argv):

    mealie_version = "3.1.0"
    mealie_banner = "MEA Lightweight Integration Engine for Maximo 6.2.7"
    build_dt = "$Date: 2012/02/07 09:50:02 $"

    param_file = None
    entity_type = None
    entity_key = None
    entity_level = None

    # Will source data just be validated without attempting to load into MEA?
    test_mode = False     # Default unless -t specified
    full_mode = False     # MA00047 Every entity processed must be new
    normal_mode = False   # MA00047

    show_version = False     

    try:
        opts, _args = getopt.getopt(argv, "e:k:p:l:tufV")
    except getopt.GetoptError:
        usage()
        sys.exit(2)

    # Parse all args passed
    for opt, arg in opts:
        if opt == '-e':
            entity_type = arg
        elif opt == '-k':
            entity_key = arg
        elif opt == '-l':
            entity_level = arg
        elif opt == '-p':
            param_file = arg
        elif opt == '-t':
            test_mode = True
        elif opt == '-f':
            full_mode = True
        elif opt == '-u':
            normal_mode = True
        elif opt == '-V':
            show_version = True

    # Now let's validate the combination of args

    # Firstly, the version can be displayed. In this case it should be the only arg

    if (show_version):
        if (entity_type or entity_key or param_file or test_mode):
            usage()
            sys.exit(2)
        else:
            print "%s, version %s. Build %s" % (mealie_banner, mealie_version, build_dt)
            sys.exit(0)

    # MA00047 We must choose to run in one of: either test mode or full mode or normal mode
    if (test_mode):
        if (full_mode or normal_mode):
            print("ERROR: Only one of the arguments -u -t -f may be specified")
            usage()
            sys.exit(2)
    elif (full_mode):
        if (test_mode or normal_mode):
            print("ERROR: Only one of the arguments -u -t -f may be specified")
            usage()
            sys.exit(2)
    elif (normal_mode):
        if (test_mode or full_mode):
            print("ERROR: Only one of the arguments -u -t -f may be specified")
            usage()
            sys.exit(2)
    else:
        print("ERROR: At least one of the arguments -u -t -f must be specified")
        usage()
        sys.exit(2)


    # Otherwise the parameter file and entity must be explicitly specified
    # with the test mode argument being optional

    if not (entity_type and param_file):
        usage()
        sys.exit(2)

    if (entity_type != 'ASS' and entity_type != 'ASP' and entity_type != 'LOC'):
        print "ERROR: Unknown entity type %s" % (entity_type)
        usage()
        sys.exit(2)

    if (full_mode) and (entity_type == 'ASP'):
        print "ERROR: -f argument is not allowed for entity %s" % (entity_type)
        usage()
        sys.exit(2)

    # Get the parameters from the parameter file

    config = ConfigParser.ConfigParser() # RawConfigParser() does not do interpolation
    config.read(param_file)

    # Configure logging

    try:
        mea_wait_interval = int(config.get('log', 'mea_wait_interval'))
    except ConfigParser.NoSectionError:
        print "ERROR: No [log] section defined in %s or %s does not exist" % (param_file, param_file)
        sys.exit(2)
    except ConfigParser.NoOptionError:
        mea_wait_interval = 180   # Default wait interval
    except ValueError:
        print("ERROR: Invalid value \"%s\" for mea_wait_interval in %s" % (config.get('log', 'mea_wait_interval'), param_file))
        sys.exit(2)

    try:
        log_file = config.get('log', 'logfile')
    except ConfigParser.NoOptionError:
        log_file = '/usr/local/log/mealie.log' # Default log filename

    try:
        log_severity = config.getint('log', 'severity')
    except ConfigParser.NoOptionError:
        log_severity = 20 # Default to INFO

    mealie_logger = logging.getLogger('MealieLogger')
    mealie_logger.setLevel(log_severity)

    fh = logging.FileHandler(log_file)
    fh.setLevel(log_severity)

    LOG_FORMAT = "%(asctime)-15s %(module)-10s %(levelname)-5s %(message)s"
    formatter = logging.Formatter(LOG_FORMAT)
    fh.setFormatter(formatter)
    
    mealie_logger.addHandler(fh)

    # Validate source and target parameters

    try:
        source_type = config.get('source', 'type')
    except ConfigParser.NoSectionError:
        print "ERROR: No [source] section defined in %s or %s does not exist" % (param_file, param_file)
        sys.exit(2)
    except ConfigParser.NoOptionError:
        print "ERROR: No source type defined in %s" % param_file
        sys.exit(2)

    # MA00047 The command line argument -f may only be specified for a spreadsheet source

    if (source_type != 'XLS' and full_mode):
        print "ERROR: -f argument is not valid for source_type %s. Type mealie -h for help" % source_type
        sys.exit(2)

    # The command line argument -k should only be, and must be, specified for a database source

    if (source_type == 'XLS' and entity_key):
        print "ERROR: -k argument should not be supplied if source_type is %s" % source_type
        sys.exit(2)

    if (source_type == 'ORA' and not entity_key):
        print "ERROR: -k argument must be supplied if source_type is %s" % source_type
        sys.exit(2)

    try:
        source_db = config.get('source', 'db')
    except ConfigParser.NoOptionError:
        if (source_type != 'ORA'):
            source_db = None
        else:
            print "ERROR: No source database defined in %s" % param_file
            sys.exit(2)

    try:
        source_user = config.get('source', 'user')
    except ConfigParser.NoOptionError:
        if (source_type != 'ORA'):
            source_user = None
        else:
            print "ERROR: No source user defined in %s" % param_file
            sys.exit(2)

    try:
        source_pwd = config.get('source', 'pwd')
    except ConfigParser.NoOptionError:
        source_pwd = None

    try:
        source_xlsfile = config.get('source', 'xlsfile')
    except ConfigParser.NoOptionError:
        if (source_type != 'XLS'):
            source_xlsfile = None
        else:
            print "ERROR: No source xlsfile defined in %s" % param_file
            sys.exit(2)


    try:
        target_db = config.get('target', 'db')
    except ConfigParser.NoSectionError:
        print "ERROR: No [target] section defined in %s" % param_file
        sys.exit(2)
    except ConfigParser.NoOptionError:
        print "ERROR: No target database defined in %s" % param_file
        sys.exit(2)

    try:
        target_user = config.get('target', 'user')
    except ConfigParser.NoOptionError:
        print "ERROR: No target user defined in %s" % param_file
        sys.exit(2)

    try:
        target_pwd = config.get('target', 'pwd')
    except ConfigParser.NoOptionError:
        target_pwd = None


    if (source_type == 'ORA'):
        if (test_mode):
            mealie_logger.info("Test Mode: Data sourced from %s will NOT be loaded" % (source_db))
            print "Test Mode: Data sourced from %s will NOT be loaded" % (source_db)
        else:
            mealie_logger.info("Data sourced from %s will be loaded into %s" % (source_db, target_db))
            print "Data sourced from %s will be loaded into %s" % (source_db, target_db)
    elif (source_type == 'XLS'):
        if (test_mode):
            mealie_logger.info("Test Mode: Data sourced from %s will NOT be loaded" % (source_xlsfile))
            print "Test Mode: Data sourced from %s will NOT be loaded" % (source_xlsfile)
        else:
            mealie_logger.info("Data sourced from %s will be loaded into %s" % (source_xlsfile, target_db))
            print "Data sourced from %s will be loaded into %s" % (source_xlsfile, target_db)
    else:
        print "ERROR: Unsupported source type %s defined in %s" % (source_type, param_file)
        sys.exit(2)

    # 7592 Treat -f as if it were -u for Production because of real time updates
    if (full_mode and string.upper(target_db) == 'MAXPROD'):
        full_mode = False
        normal_mode = True

    if (full_mode):
        mealie_logger.info("Full Load: All entities processed must be loaded by MEA after %d seconds" % mea_wait_interval)

    mealie_logger.info("Logging level set to %s" % log_severity)

    mealie_logger.info("Processing entity %s LIKE %s" % (entity_type, entity_key))

    if (entity_level):
        mealie_logger.info("Processing entity LEVEL %s" % entity_level)
    else:
        mealie_logger.info("Processing all entity LEVELs")

    er = EntityRun(log_file, log_severity, mealie_logger,
                        source_db, source_user, source_pwd,
                        source_xlsfile,
                        source_type,
                        target_db, target_user, target_pwd,
                        entity_type, entity_key, entity_level,
                        mea_wait_interval,
                        normal_mode, test_mode, full_mode)

    return er


def stage(argv):

    stage_file = None
    target_user = None
    target_pwd = None
    target_db = None
    log_severity = None

    try:
        opts, _args = getopt.getopt(argv, "s:u:p:d:l:")
    except getopt.GetoptError:
        stgusage()
        sys.exit(2)

    # Parse all args passed
    for opt, arg in opts:
        if opt == '-s':
            stage_file = arg
        elif opt == '-u':
            target_user = arg
        elif opt == '-p':
            target_pwd = arg
        elif opt == '-d':
            target_db = arg
        elif opt == '-l':
            log_severity = arg

    # Mandatory args must be supplied
    if not (stage_file and target_user and target_db):
        stgusage()
        sys.exit(2)

    stage_file = stage_file + ".stg"
    if not (os.path.isfile(stage_file)):
        print("ERROR: Stage file %s not found" % stage_file)
        stgusage()
        sys.exit(2)


    # Default to INFO
    if not (log_severity):
        log_severity = '20'

    if (log_severity != '10' and log_severity != '20' and log_severity != '40'):
        print "ERROR: Invalid log severity %s" % (log_severity)
        stgusage()
        sys.exit(2)

    sb = StageBatch(stage_file, int(log_severity), target_db, target_user, target_pwd)

    return sb

