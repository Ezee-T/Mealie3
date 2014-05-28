#
#
# $Header: $
#
#
# $Log: $
#
#

'''
Created on 26 Feb 2014

@author: ACE
'''

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

   print "Usage: meagen_servitem -i infile -o outfile -r row -c col"


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
DELETE FROM maximo.mxservitem_iface;

-- Insert new data in service item interface table

"""

   script_parte = """
-- Insert in queue table

INSERT INTO maximo.mxin_inter_trans(extsysname, ifacename,transid) (SELECT 'EXTSYS1', 'MXSERVITEMInterface', transid FROM maximo.mxservitem_iface)
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

      ofh.write("INSERT INTO maximo.mxservitem_iface (")
      ofh.write("itemnum, description, itemsetid, orderunit, transid, transseq) ")

      # MA00044 Sometimes itemnum is a float, if so convert to string
      if (wsh.cell(this_row,9).ctype == xlrd.XL_CELL_NUMBER):
         itemnum = str(int(wsh.cell(this_row,9).value))
      else:
         itemnum = p.sub('&\'||\'', wsh.cell(this_row,9).value)

      description = p.sub('&\'||\'', wsh.cell(this_row,10).value)
      description = description.encode('ascii', 'ignore')
      orderunit = wsh.cell(this_row,13).value

      ofh.write("VALUES('%s', UPPER('%s'), 'ITEMSET', UPPER('%s'), %d, 1)" %
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

