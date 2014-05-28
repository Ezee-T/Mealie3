'''
Created on 24 Jul 2012

@author: ACE
'''

#### Import Module

import cx_Oracle
import xlrd
import sys
import time

processed_ic = []
2222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222
#### Class Definitions

class InterfaceAttrStructureFromXLS:
    
    def __init__(self, worksheet, i, c, logger):
        
        #self.classid = worksheet.cell(i, 0).value
        #self.classname = worksheet.cell(i, 1).value
        self.attrseq = worksheet.cell(i, 4).value
        self.propident = worksheet.cell(i, 5).value
        self.propname = worksheet.cell(i, 6).value
        self.measurecode = worksheet.cell(i, 7).value
        self.data_type = worksheet.cell(i, 10).value
        #self.useinitemdesc = worksheet.cell(i, 11).value
        
        logger.debug("Getting classname for attribute %s" % (self.propident))
        
        logger.debug("Can attribute %s be used in item description %s"  % (self.propident, self.propname))
        
        #if self.useinitemdesc == "Y":
        #    self.useinitemdesc = 1
        #else:
        #    self.useinitemdesc = 0
        
        #q_classid= """
        #   SELECT c.ESCN
        #   FROM maximo.hul_classif_dict c
        #   WHERE c.ESCI = :classid
        #   AND c.ESPI = :propident"""
       
        #c.execute(q_classid, classid=self.classid, propident=self.propident)
        #this_classid = c.fetchone()
        #logger.debug("Class description for classid %s is %s" % (self.classid, this_classid))
           
    def is_valid(self, c, logger):
        
        
               
        return True



########################################################################################
#
#     Functions
#
########################################################################################

def get_next_txn(c, logger):
    
    q_get_id = """
    SELECT maximo.hul_AttributeStruc_seq.NEXTVAL
    FROM DUAL"""
    
    c.execute(q_get_id)
    txn_id = c.fetchone()
    logger.debug("Obtained MEA Txn Id %s" % (txn_id[0]))

    return txn_id[0]

def write_mea_queue(c, txid, logger):

    # Insert row into MEA queue

    q_write_queue = """
    INSERT INTO maximo.mxin_inter_trans (
                           extsysname,
                           ifacename,
                           action,
                           transid
                           )
    VALUES ('EXTSYS1', 'MXATTRIBUTEInterface', 'AddChange', :trans_id)"""

    logger.debug("Writing MEA queue for transid %d" % txid)
    c.execute(q_write_queue, trans_id=txid)

    return True

def Attribute_handler(er):
    type_handler = {
        "XLS": attribute_handler_xls_ora
        }
    return type_handler.get(er.source_type)(er)

def attr_exists(conn, c, cls, attr, logger):
    
    #Verify the existence of attribute and write into mxattribute_iface if it doesn't exist
    
    q_get_exists = """
    SELECT a.assetattrid
    FROM maximo.assetattribute a
    WHERE a.assetattrid = :assetattrid"""
    
    c.execute(q_get_exists, assetattrid=attr)
    result = c.fetchone()
    if result == None:
        logger.debug("Lookup q_get_exists failed in attrid_exists(): assetattrid %s" % (attr))
        retval = False
    else:
        retval = True
    return retval

def process_classid(ia, c, logger):
    
    q_write_interface = """
    INSERT INTO maximo.mxattribute_iface(
        assetattrid,
        description,
        measureunitid,
        datatype,
        transid,
        transseq)
    VALUES (
        :assetattrid,
        :description,
        :measureunitid,
        :datatype,
        :transid,
        :transseq)"""
        
    if ia.propident not in processed_ic:

        # Obtain new MEA transaction id for next asset 

        logger.debug("Getting new transaction id for this load")
        trans_id = get_next_txn(c, logger)
        

        trans_seq = 1          # Sequence within batch

        # Write the MEA interface table
        logger.debug("Processing interface record for attribute=%s" % (ia.propident))
                
        logger.info("Writing interface: assetattrid=%s description=%s measureunitcode=%s datatype=%s attrdescprefix=%s transid=%s transseq=%s " % (ia.propident, ia.propname, ia.measurecode, ia.data_type, ia.propname, trans_id, trans_seq))

        #logger.debug("commoditygroup=%s" % (ic.commoditygroup))
        logger.debug("classification=%s" % (ia.propident))
        c.execute(q_write_interface, assetattrid=ia.propident,
                                     description=ia.propname,
                                     measureunitid=ia.measurecode,
                                     datatype=ia.data_type,
                                     #attrdescprefix=ia.propname+":",
                                     transid=trans_id, 
                                     transseq=trans_seq)
        logger.debug("Wrote interface")
        processed_ic.append(ia.propident)        # We processed it, remember it

        # Write the MEA queue
        write_mea_queue(c, trans_id, logger)
        logger.debug("Wrote MEA queue for transid %d" % trans_id)
        ct = 1
                
    else:
                
        logger.debug("Attribute %s was already processed, skipping" % (ia.propident))
        ct = 0
        
    return ct

def attribute_handler_xls_ora(er):
    
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
    
    # Insert new attribute
    
    q_write_attr_inter = """
    INSERT INTO maximo.mxattribute_iface (
        assetattrid,
        description,
        datatype,
        measureunitid,
        transid,
        transeq)
    VALUES (
        :assetattrid,
        :description,
        :datatype,
        :measuerunitid,
        :transid,
        :transeq)"""

        


    # Initialise local variables for fetch loop

    #queued_batches = 0 
    total_ct = 0       # Total no. of items

    start_row = 1      # The first row in spreadsheet that contains data (start at 0)

    # Iterate through all rows of data in input spreadsheet


    for this_row in range(start_row, wsh.nrows):
  
        ia = InterfaceAttrStructureFromXLS(wsh, this_row, curs_target_seq, er.logger)
        er.logger.info("Processing ESPI %s" % (ia.propident))
        
        if (ia.is_valid(curs_target_lookup, er.logger)):
            
            #Obtain new Transaction Id for next classification
            total_ct += process_classid(ia, curs_target_mea, er.logger)
                
        else:

            er.logger.error("ERROR: Values for itemnum %s are invalid, aborting" % (ia.itemnum))
            print("ERROR: Values for itemnum %s are invalid, aborting" % (ia.itemnum))
            conn_target.rollback()
            conn_target.close()         # Free the connection
            sys.exit(2)
        #End Classification processing loop
        
        # The transaction will not be committed if we are in 'test mode'
    # Otherwise we commit the whole transaction

    if (not er.test_mode):
        er.logger.info("Committing changes to database target %s" % er.target_db)
        conn_target.commit() 
    else:
        er.logger.info("Test Mode: Rolling back changes to %s" % er.target_db)
        conn_target.rollback() 

    er.logger.info("Processed %s %s entities" % (total_ct, er.entity_type))

    print("Processed %s %s entities" % (total_ct, er.entity_type))
    print "See %s for full details of this run" % er.log_file

    # Close open cursors and disconnect from target database
    curs_target_lookup.close()
    conn_target.close()

    return True
