#!/usr/bin/env python

# $Header: /usr/local/cvsroot/mealie/mealie.py,v 1.2 2012/02/10 14:07:50 djw Exp $
#
# MEA Lightweight Integration Environment
#
# $Log: mealie.py,v $
# Revision 1.2  2012/02/10 14:07:50  djw
# 8112 Handling for ICG (Item Commodity)
#
# Revision 1.1  2012/01/05 12:21:54  djw
# Migrated to new project
#
# Revision 1.7  2011/03/28 13:00:40  djw
# 7474 Add ASP entity
#
# Revision 1.6  2010/11/12 08:39:05  djw
# 6789 Move this module back into the /usr/local/bon path
#
# Revision 1.1  2010/10/26 12:39:28  njm3
# Changed mealie path - one dir higher:
#     e.g. mealie.base.x is now base.x
#     (Update MEALIE_PATH in ~/.bash_profile accordingly: /usr/local/lib/python/mealie)
#
# Revision 1.5  2010/09/01 14:42:50  djw
# Unknown. late check in from 14/4
# Adds handler for the asset entity ASS
#
# Revision 1.4  2010/04/13 11:39:57  djw
# MA00036. Use "MEALIE_PATH" to locate the packages and add some user
# friendly validation that it is set correctly
#
# Revision 1.3  2010/04/13 11:25:59  djw
# Bring code in line with PEP-8
#
# Revision 1.2  2010/04/12 10:06:56  djw
# MA00036. Call appropriate handler according to entity
#
# Revision 1.1  2010/04/12 06:04:41  djw
# MA00036 Refactor and package
# Rename mealie.py to mealieInit.py
#
# Revision 1.6  2010/04/08 14:29:54  djw
# MA00036 Change driving query for LOC to process parents and sub-parents
#
# Revision 1.5  2010/04/08 06:50:56  djw
# MA00036 Validate source type parameter
#
# Revision 1.4  2010/04/07 13:00:34  djw
# MA00036 Now batches MEA interface records where they share the
# same parent. The transid is taken from the hul_mealie_seq sequence
# and the transseq increments within each batch
#
# Revision 1.3  2010/04/06 09:00:31  djw
# MA00036 Add logging with formatter and severity
#
# Revision 1.2  2010/04/06 07:27:40  djw
# mealie.py
#
# Revision 1.1  2010/04/01 10:44:27  djw
# MA00036 Intermediate check in. Driving cursor and CLI validation
#
#
#

########################################################################################
#
#     Import required modules
#
########################################################################################

import sys
# import os

# The usual setting at Hulamin is "/usr/local/lib/python" but YMMV
# mealie_path = os.environ.get("MEALIE_PATH")

# if (mealie_path):
#     sys.path.append(mealie_path)
# else:
#     print("ERROR: Environment variable \"MEALIE_PATH\" is not defined")
#     sys.exit(2)

try:
    from launcher import main
except ImportError:
    print("ERROR: Cannot import base.launcher. Is your \"MEALIE_PATH\" set correctly?")
    sys.exit(2)

from entities.assetspecs import asp_handler
from entities.assets import ass_handler
from entities.locations import loc_handler

action_handler = {
    "ASP": asp_handler,
    "ASS": ass_handler,
    "COM": com_handler,
    "ICG": icg_handler,
    "LOC": loc_handler,
    "UND": und_handler,
    "UOM": uom_handler
}

if __name__ == "__main__":
    er = main(sys.argv[1:])                # Pass all args except name of program
    action_handler.get(er.entity_type)(er) # Call the appropriate handler for the entity
