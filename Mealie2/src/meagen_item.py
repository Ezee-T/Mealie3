#!/usr/bin/env python
#
# $Header: /usr/local/cvsroot/maximo6/MEA/bin/meagen_item.py,v 1.8 2010/11/11 18:57:17 djw Exp $
#
# Read Item Master details from spreadsheet. Start at a certain cell in 
# the spreadsheet and read subsequent rows, same column, until there is an undefined
# value. This is assumed to be the end of the list of Itemnums
#
# Other format assumptions
# - the Item Description is in the column +1 from Itemnum
# - the Item Order Unit is in the column +4 from Itemnum
# - that's it
#
# Usage: meagen_item -i infile -o outfile -r row -c col 
#
#        infile     Name of file to be processed (.CSV format). Required argument.
#        outfile    Name of script file to be generated. Required argument.
#        row        Row number of first employee number to be processed. Required
#                   argument. First row is row 1.
#        col        Column number of first employee number to be processed (A=1, 
#                   B=2, etc.). Required argument.
#
# $Log: meagen_item.py,v $
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
# Revision 1.2  2010/05/17 10:55:41  djw
# 5939. Add SET ECHO ON
#
# Revision 1.1  2010/05/17 07:42:19  djw
# 5939. Initial revision
#
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

   print "Usage: meagen_item -i infile -o outfile -r row -c col"


def spoolfile_name(fullname):

   #
   # Generate and return the name of the spool file
   #

   fname, fext = fullname.split('.')
   return fname


def process_file(ifname, ofname, start_col, start_row):

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
DELETE FROM maximo.mxitem_iface;

-- Insert new data in item interface table

"""

   script_parte = """
-- Insert in queue table

INSERT INTO maximo.mxin_inter_trans(extsysname, ifacename,transid) (SELECT 'EXTSYS1', 'MXITEMInterface', transid FROM maximo.mxitem_iface)
/

COMMIT
/

SPOOL OFF

"""

   cell = ()
   rec_ct = 0

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

   # Convert from index starting at 1 to index starting at 0
   start_col -= 1
   start_row -= 1

   p = re.compile(r'&')

   # Iterate through rows in input spreasheet until we find the end
   # of the list
   for this_row in range(start_row, wsh.nrows):

      ofh.write("INSERT INTO maximo.mxitem_iface (")
      ofh.write("itemnum, description, itemsetid, orderunit, lottype, transid, transseq) ")

      # MA00044 Sometimes itemnum is a float, if so convert to string
      if (wsh.cell(this_row,9).ctype == xlrd.XL_CELL_NUMBER):
         itemnum = str(int(wsh.cell(this_row,9).value))
      else:
         itemnum = p.sub('&\'||\'', wsh.cell(this_row,9).value)

      description = p.sub('&\'||\'', wsh.cell(this_row,10).value)
      description = description.encode('ascii', 'ignore')
      orderunit = wsh.cell(this_row,13).value

      ofh.write("VALUES('%s', '%s', 'ITEMSET', '%s', 'NOLOT', %d, 1)" %
               (itemnum, description, orderunit,
                this_row - start_row + 1))
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

   col_no = 8
   row_no = 3

   print "'%s' will be processed starting at row %d, column %d" % (in_file, row_no, col_no)
   process_file(in_file, out_file, col_no, row_no)
   print "Generated script %s" % out_file

   sys.exit(0)

if __name__ == "__main__":
   main(sys.argv[1:])    # Pass all args except name of program

