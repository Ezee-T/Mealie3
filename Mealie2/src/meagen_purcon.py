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

########################################################################################
#
#     Functions
#
########################################################################################


def usage():

   #
   # Display usage syntax
   #

   print "Usage: meagen_purcon -i infile -o outfile"


def spoolfile_name(fullname):

   #
   # Generate and return the name of the spool file
   #

   fname, fext = fullname.split('.')
   return fname


def format_contract_dt(dt_raw, dt_type, ofh, ifh):

   #
   # Format the date as yyyymmdd
   #

   if (dt_type == xlrd.XL_CELL_EMPTY):
      retval = None      # Contract start and end date are optional so this is allowed
   elif (dt_type == xlrd.XL_CELL_DATE):
      dt_t = xlrd.xldate_as_tuple(dt_raw, ifh.datemode) # Windows/Mac mode
      retval = "".join([str(dt_t[0]).rjust(4, '0'), str(dt_t[1]).rjust(2, '0'), str(dt_t[2]).rjust(2, '0')])
   else:
      print("ERROR: Non-date value \"%s\" found in start/end date" % dt_raw)
      print("Aborting run.")
      ofh.close()
      sys.exit(2)
   return retval

def process_file(ifname, ofname):

   #
   # Main control function for processing the file
   #

   script_parta = """--
-- $Header"""

   script_partb = """
--
--
-- $Log"""

   script_partc = """
--
--
SET ECHO ON
"""

   script_partd = """
DELETE FROM maximo.mxpurcon_iface;

-- Insert new data in purchase contract interface table

"""

   script_parte = """
-- Insert in queue table

INSERT INTO maximo.mxin_inter_trans(extsysname, ifacename, transid) VALUES ('EXTSYS1', 'MXPURCONInterface', 1)
/

COMMIT
/

SPOOL OFF

"""

   cell = ()
   rec_ct = 0

   start_row = 2

   p = re.compile(r'&')

   # Open input file to be processed
   ifh = xlrd.open_workbook(ifname)
   wsh = ifh.sheet_by_index(0)

   # Open output file to be created. Will overwrite existing file
   # without warning
   ofh = open(ofname, 'w')

   ofh.write(script_parta)
   ofh.write(chr(36))
   ofh.write(script_partb)
   ofh.write(chr(36))
   ofh.write(script_partc)

   ofh.write("SPOOL %s\n" % spoolfile_name(ofname))
   ofh.write(script_partd)

   # Iterate through records in input spreadsheet until we find the end
   # of the list
   for this_row in range(start_row, wsh.nrows):

      # MA00044 Sometimes itemnum is a float, if so convert to string
      if (wsh.cell(this_row,9).ctype == xlrd.XL_CELL_NUMBER):
         itemnum = str(int(wsh.cell(this_row,9).value))
      else:
         itemnum = p.sub('&\'||\'', wsh.cell(this_row,9).value)

      itemdesc = p.sub('&\'||\'', wsh.cell(this_row,10).value)
      itemdesc = itemdesc.encode('ascii', 'ignore') # Lose chars > 255
      if (wsh.cell(this_row,11).ctype == xlrd.XL_CELL_NUMBER):
         catalogcode = str(int(wsh.cell(this_row,11).value))
      else:
         catalogcode = p.sub('&\'||\'', wsh.cell(this_row,11).value)
      orderqty = wsh.cell(this_row,12).value
      orderunit = wsh.cell(this_row,13).value
      unitcost = wsh.cell(this_row,14).value

      # 6918 DJW 11-Nov-2010 Contract start date is now in column 6 (G)
      #                      Contract end date is now in column 7 (H)

      start_dt = format_contract_dt(wsh.cell(start_row,6).value, wsh.cell(start_row,6).ctype,  ofh, ifh)
      end_dt = format_contract_dt(wsh.cell(start_row,7).value, wsh.cell(start_row,7).ctype,  ofh, ifh)

      ofh.write("INSERT INTO maximo.mxpurcon_iface (")
      ofh.write("contractnum, description, revisionnum, vendor, contractlinenum, ")
      if (start_dt != None):
         ofh.write("startdate, ")
      if (end_dt != None):
         ofh.write("enddate, ")
      ofh.write("itemnum, cl_description, catalogcode, orderqty, orderunit, ")
      ofh.write("unitcost, chgpriceonuse, chgqtyonuse, contracttype, ")
      ofh.write("n_tax1code,orgid, transid, transseq) ")
      ofh.write("VALUES('%s', " % p.sub('&\'||\'', wsh.cell(start_row,1).value))
      ofh.write("'%s', " % p.sub('&\'||\'', wsh.cell(start_row,2).value))
      ofh.write("%d, " % wsh.cell(start_row,4).value)
      ofh.write("'%s', " % wsh.cell(start_row,5).value)
      ofh.write("%d, " % int(this_row - start_row + 1))
      if (start_dt != None):
         ofh.write("TO_DATE('%s','YYYYMMDD'), " % start_dt)
      if (end_dt != None):
         ofh.write("TO_DATE('%s','YYYYMMDD'), " % end_dt)
      ofh.write("'%s', " % itemnum)
      ofh.write("UPPER('%s'), " % itemdesc)
      ofh.write("'%s', " % catalogcode)
      ofh.write("'%s', " % orderqty)
      ofh.write("UPPER('%s'), " % orderunit)
      ofh.write("%s, " % unitcost)
      ofh.write("1, 1, '%s', " % wsh.cell(start_row,3).value)
      ofh.write("'STANDARD', 170, 1, %d)" % int(this_row - start_row + 1))
      ofh.write("\n/\n")

   ofh.write(script_parte)

   # Close file
   ofh.close()



########################################################################################
#
#     Main process 
#
########################################################################################

def main(argv):

   # Parse and validate arguments then process the file

   try:
      opts, args = getopt.getopt(argv, "i:o:")
   except getopt.GetoptError:
      usage()
      sys.exit(2)

   # Initialise variables set by command line arguments
   in_file = False
   out_file = False

   # Parse command line input
   for opt, arg in opts:
      #print "opt=%s arg=%s" % (opt, arg)
      if opt == '-i':
         in_file = arg
      elif opt == '-o':
         out_file = arg

   if not (in_file and out_file):
      # One or more mandatory argument(s) not supplied
      usage()
      sys.exit(2)

   process_file(in_file, out_file)
   print "Generated loader script %s from '%s'" % (out_file, in_file)

   sys.exit(0)

if __name__ == "__main__":
   main(sys.argv[1:])    # Pass all args except name of program

