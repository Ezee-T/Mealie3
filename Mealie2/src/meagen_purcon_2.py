#!/usr/bin/env python
#
# $Header: /usr/local/cvsroot/maximo6/MEA/bin/meagen_purcon.py,v 1.9 2011/04/05 09:20:03 djw Exp $
#
# Purpose
# =======
#
# Read Purchase Contract details from spreadsheet. Start at cell B3 in 
# the spreadsheet and read subsequent rows until there is an undefined
# value in the itemnum column (I). This is assumed to be the end of 
# the list of items that are to be loaded as Contract Lines
#
# Usage: meagen_item -i infile -o outfile
#
#        infile     Name of file to be processed (.CSV format). Required argument.
#        outfile    Name of script file to be generated. Required argument.
#
# Data derivation
# ===============
#
# Values that are loaded from the spreadsheet:
#
# B3         Contract number
# C3         Contract description
# D3         Contract type (PRICE, PURCHASE)
# E3         Contract revision number
# F3         Vendor code
# G3         Contract start date
# H3         Contract end date
# I3 - Gn    ( NB The contract line number is here for humans but it 
#              is ignored: auto-generate from 1 )
# J3 - Jn    Itemnum
# K3 - Kn    Item description
# L3 - Ln    Catalogue code (optional)
# M3 - Mn    Order quantity 
# N3 - Nn    Order unit 
# O3 - On    Unit cost 
#
# Constant or auto-generated values that are always loaded:
#
# Contract line number      Starts at 1 for the contract line in row 3 and 
#                           increments by 1
# Tax code                  Always 'STANDARD'
# Org id                    Always 170
#
#
#
# $Log: meagen_purcon.py,v $
# Revision 1.9  2011/04/05 09:20:03  djw
# 7532 Enhanced so that numeric data in the catalogcode column can be handled
# This was first encountered when trying to load a spreadsheet from Thivash
# Somayi. See Issue 7531 for the actual load
#
# Revision 1.8  2010/11/11 18:57:17  djw
# 6918 Contract start date and end date are permitted in cells G3 and
# H3 respectively. These are optional but if supplied must be in an
# Excel DATE format.
#
# Revision 1.7  2010/09/01 09:53:33  djw
# MA00044 Sometimes itemnum is a float, if so convert to string
#
# Revision 1.6  2010/08/31 16:43:10  djw
# MA00043 Fix to handle item descriptions that contain non-ASCII characters
#
# Revision 1.5  2010/06/09 11:22:01  djw
# 5948 Bug fix. Header keywords are updated by CVS but they are only part
# of the header of the generated script
#
# Revision 1.4  2010/06/09 11:02:36  djw
# 5941 Bug fix. Ampersand treated as start of SQL*Plus substitution var
#
# Revision 1.3  2010/06/09 08:50:06  djw
# 5962. Bug fix. Embedded commas cause data/field boundary corruption
# Now use xlrd and process a .xls directly.
# See \\hul-fps1\djw$\Maximo\Contracts\Maximo-Contract-Load.doc rev 1.1
#
# Revision 1.2  2010/05/17 13:15:14  djw
# 5939 Now takes contract type in D3
#
# Revision 1.1  2010/05/17 10:56:02  djw
# 5939. Initial revision
#
#
#

########################################################################################
#
#     Import required modules
#
########################################################################################

import sys
import re
import getopt
import xlrd
import time
import cx_Oracle
from xlwt.Worksheet import Worksheet

########################################################################################
#
#     Functions
#
########################################################################################

class InterfaceContractFromXLS:
    
    """Class Representing a contract to be loaded by MEA from a spreadsheet"""
    
    def __init__(self, worksheet, i, c, logger):
       self.contract =  worksheet.cell(i, 1).value
       self.description = worksheet.cell(i, 2).value
       self.type = worksheet.cell(i, 3).value
       self.revision = worksheet.cell(i, 4).value
       self.vendor = worksheet.cell(i, 5).value
       self.startdate = worksheet.cell(i, 6).value
       self.enddate = worksheet.cell(i, 7).value
       self.linenum = worksheet.cell(i, 8).value
       self.itemnum = worksheet.cell(i, 9).value
       self.itemdescr = worksheet.cell(i, 10).value
       self.catalog = worksheet.cell(i, 11).value
       self.itemqty = worksheet.cell(i, 12).value
       self.orderunit = worksheet.cell(i, 13).value
       self.unitcost = worksheet.cell(i, 14).value
       
    
    def is_valid(self, c, logger):
  #



    
        return True


def isp_handler(er):

    type_handler = {
        "XLS": spec_handler_xls_ora
    }

    return type_handler.get(er.source_type)(er)   # Call appropriate handler for the source type



def spec_handler_xls_ora(er):

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
    
    # Open the necessary cursors on target database
    
    curs_target_mea = conn_target.cursor()
    curs_target_lookup = conn_target.cursor()
    curs_target_seq = conn_target.cursor()
    
    # Insert a row into MEA interface for this entity type 
    
    q_write_contract = """
    INSERT INTO maximo.mxpurcon_iface (contractnum, description, revisionnum, vendor, contractlinenum, 
                    startdate, enddate, itemnum, 
                    cl_description, catalogcode, orderqty, orderunit, unitcost, chgpriceonuse, 
                    chgqtyonuse, contracttype, n_tax1code, orgid, transid, transseq) 
    VALUES ()"""
                    
                    
    # Initialise local variables for fetch loop

    total_ct = 0       # Total no. of items
    start_row = 1      # The first row in spreadsheet that contains data (starting at 0)
    
    
    # Iterate through all rows of data in input spreadsheet
    
    for this_row in range(start_row, wsh.nrows):

            ic = InterfaceContractFromXLS(wsh, this_row, curs_target_seq, er.logger)
            er.logger.info("Processing contract-line number %s for item %s" % (ic.linenum, ic.itemnum))
            
            
                    