#!/usr/bin/env python

# $Header: /usr/local/cvsroot/mealie/mealiestg.py,v 1.2 2012/01/05 14:15:57 djw Exp $
#
# MEA Lightweight Integration Environment
#
# Initialisation procedure for Stager
#
# $Log: mealiestg.py,v $
# Revision 1.2  2012/01/05 14:15:57  djw
# Simplify imports and drop the MEALIE_PATH
#
# Revision 1.1  2012/01/05 12:21:54  djw
# Migrated to new project
#
# Revision 1.1  2011/04/15 12:53:32  djw
# 7591 Initial revision
#
#
#

########################################################################################
#
#     Import required modules
#
########################################################################################

import sys
import os

from launcher import stage
from stage import stage_handler

if __name__ == "__main__":
    sb = stage(sys.argv[1:])       # Pass all args except name of program
    stage_handler(sb)              # Process the stage
