#!/usr/bin/env python
#
# $Header: /usr/local/cvsroot/mealie/stage.py,v 1.5 2012/02/10 14:07:50 djw Exp $
#
# MEA Lightweight Integration Environment
#
# Stage Handler
#
# $Log: stage.py,v $
# Revision 1.5  2012/02/10 14:07:50  djw
# 8112 Handling for ICG (Item Commodity)
#
# Revision 1.4  2012/02/10 10:00:54  djw
# 8112 Handling COM (Commodities) entity. Unit tested on MAXSIT
# See maximo/8112/hul_test_COM.stg and hul_test_item.xls
#
# Revision 1.3  2012/02/09 07:53:35  djw
# 8112 UND entity unit tested. Test data at maximo/8112/hul_test_UND.stg
# maximo/8112/hul_test_unspsc_dict.stg
#
# Revision 1.2  2012/02/07 09:02:37  djw
# 8112 New UOM entity handler. Tested on maxsit using dummy data. See C:\SRCCVS\maximo6\8112\hul_test_UOM.stg
#
# Revision 1.1  2012/01/05 12:21:54  djw
# Migrated to new project
#
# Revision 1.4  2011/04/29 10:30:56  djw
# 7637 Can now process stage files with blank lines
#
# Revision 1.3  2011/04/15 15:37:56  djw
# 7591 Handling for invalid entity type, mode in stage file
#
# Revision 1.2  2011/04/15 15:17:28  djw
# 7591 Abort whole stage if error occurs
#
# Revision 1.1  2011/04/15 12:50:35  djw
# 7591 Add stager enhancement
#
#
#

########################################################################################
#
#     Import required modules
#
########################################################################################

import re
import logging

import launcher
from entities.assetspecs import asp_handler
from entities.assets import ass_handler
from entities.commodities import com_handler
from entities.itemcomm import icg_handler
from entities.locations import loc_handler
from entities.measureunits import uom_handler
from entities.unspscdict import und_handler
from entities.classifdict import classif_handler
from entities.Specifications import class_handler
from entities.AttributeSpec import Attribute_handler
from entities.ItemSpec2 import isp_handler
from entities.item import item_handler
from entities.InvBalance import invbalance_handler
from entities.Inventory import inventory_handler
from entities.ItemSpec3 import isp_handler
from entities.PoFind import po_handler
from entities.Items_Tools import ToolItem_handler
from entities.Specifications_ClassIDComm import classcomm_handler
from entities.Commodities_ClassIDComm import CommComm_handler

########################################################################################
#
#     Class definitions
#
########################################################################################


########################################################################################
#
#     Functions
#
########################################################################################

def stage_handler(sb):

    action_handler = {
        "ASP": asp_handler,
        "ASS": ass_handler,
        "COM": com_handler,
        "ICG": icg_handler,
        "LOC": loc_handler,
        "UND": und_handler,
        "UOM": uom_handler,
        "ISP3": isp_handler,
        "ISP2": isp_handler,
        "CLD": classif_handler,
        "SPE": class_handler,
        "ATTR": Attribute_handler,
        "ITEM": item_handler,
        "INVB": invbalance_handler,
        "INV": inventory_handler,
        "POF": po_handler,
        "TOOL": ToolItem_handler,
        "CLCO": classcomm_handler
    }

    # Prompt user for target password, if specified

    if (not sb.target_pwd):
        sb.target_pwd = str(raw_input("Enter password for %s@%s: " % (sb.target_user, sb.target_db)))

    mealie_logger = logging.getLogger('MealieLogger')
    mealie_logger.setLevel(sb.log_severity)

    # Fixed values for batches
    source_type = 'XLS'
    mea_wait_interval = 180 

    # Open stage file

    stg_fh = open(sb.stage_file, 'r')

    regexp = re.compile(r"^#")

    for stg_rec in stg_fh:

        if not regexp.search(stg_rec): # Ignore comment

            if len(stg_rec.rstrip()):   # Ignore blank line

                # print stg_rec.rstrip().split(':') # Chomp and parse single stage record
                (entity_type, mode, source_xlsfile) = stg_rec.rstrip().split(':')

                # Construct log filename from data filename
                log_file = source_xlsfile.split('.')[0] + '.log'


                fh = logging.FileHandler(log_file)
                fh.setLevel(sb.log_severity)

                LOG_FORMAT = "%(asctime)-15s %(module)-10s %(levelname)-5s %(message)s"
                formatter = logging.Formatter(LOG_FORMAT)
                fh.setFormatter(formatter)

                mealie_logger.addHandler(fh)

                # Set mode
                if (mode == 'f'):
                    full_mode = True
                else:
                    full_mode = False
                if (mode == 'u'):
                    normal_mode = True
                else:
                    normal_mode = False
                if (mode == 't'):
                    test_mode = True
                else:
                    test_mode = False

                if not (full_mode or test_mode or normal_mode):
                    print("ERROR: Cannot process unknown mode \'%s\' for %s file %s" % (mode, entity_type, source_xlsfile))
                    mealie_logger.error("Cannot process unknown mode \'%s\' for %s file %s" % (mode, entity_type, source_xlsfile))
                    return False

                er = launcher.EntityRun(log_file, sb.log_severity, mealie_logger,
                                                       None, None, None,
                                                       source_xlsfile,
                                                       source_type,
                                                       sb.target_db, sb.target_user, sb.target_pwd,
                                                       entity_type, None, None,
                                                       mea_wait_interval,
                                                       normal_mode, test_mode, full_mode)

                print("Processing %s file %s" % (er.entity_type, er.source_xlsfile))
                mealie_logger.info("Processing %s file %s" % (er.entity_type, er.source_xlsfile))

                # Call appropriate handler for the entity
                handler = action_handler.get(er.entity_type)
                if (handler == None):
                    print("ERROR: Cannot process unknown entity type \'%s\' for file %s" % (entity_type, source_xlsfile))
                    mealie_logger.error("Cannot process unknown entity type \'%s\' for file %s" % (entity_type, source_xlsfile))
                    return False
                else:
                    rs = handler(er)

                mealie_logger.removeHandler(fh)

                if not (rs):
                    return False  # Abort the stage if an error occurs

                # Record processing ends

    stg_fh.close()

    return True


