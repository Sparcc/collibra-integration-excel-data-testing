import sys, os
sys.path.append(os.getcwd())
from CompareExcel import *
from openpyxl import load_workbook
import unittest

class CompareExcelMapping(CompareExcel):
    def buildMap(self):
        self.mapping={}
        self.mapping['schema'] = ['R','G']
        self.mapping['size'] = ['P','H']
        self.mapping['nullable'] = ['M','I']
        self.mapping['col_pos_id'] = ['F','J']
        self.mapping['frac_digs'] = ['Q','K']
        self.mapping['default_value'] = ['I','L']
        self.mapping['desc'] = ['E','M']
        self.mapping['pk'] = ['N','N']
    def verifyMapRow(self,rowNum):
        for k,v in self.mapping.items():
            d1 = self.ws["file1"][v[0]+str(rowNum)].value
            d2 = self.ws["file2"][v[1]+str(rowNum)].value
    def buildConcatColMapAndVerifyRow(self,rowNum):
        #define maps
        concatCol = 'A'#column to split
        f1_schema = 0 #1st part of split column
        f1_table = 1
        f1_column = 2
        
        f2_schema = 'B' #column B
        f2_table = 'C'
        f2_column = 'F'

        #process schema,table,column
        row = {} #stores in here for code readability
        row["file1"] = self.ws["file1"][concatCol+str(rowNum)].value.split('::')
        
        #schema
        row["file2"] = self.ws["file2"][f2_schema+str(rowNum)].value
        self.compareData(row["file1"][f1_schema],row["file2"],rowNum)
        
        #table
        row["file2"] = self.ws["file2"][f2_table+str(rowNum)].value
        self.compareData(row["file1"][f1_table],row["file2"],rowNum)
        
        #column
        row["file2"] = self.ws["file2"][f2_column+str(rowNum)].value
        self.compareData(row["file1"][f1_column],row["file2"],rowNum)
    def processRow(self,rowNum):

        self.buildConcatColMapAndVerifyRow(rowNum)
        self.verifyMapRow(rowNum)            

class testcompareExcelIntegrationData(unittest.TestCase):
    def setUp(self):
        #self.comparer = CompareExcel(file1, file2)
        self.comparer = CompareExcel()
        self.comparer.setDataLength(52)
        
        self.oracleComparer = CompareExcelMapping()
        self.oracleComparer.setDataLength(36)
    def testCompareExcel(self):
        file1 = "CompareExcel-test-file1.xlsx"
        file2 = "CompareExcel-test-file2.xlsx"
        self.comparer.setSourceFiles(file1, file2)
        file1 = ['A','C','D']
        file2 = ['A','C','F']
        self.comparer.setColumns(file1, file2)
        self.comparer.compareFiles()
    def testBIDW_Excel_calendar_table(self):
        file1 = "oracle-collibra-export-calendar.xlsx"
        file2 = "oracle-export-calendar.xlsx"
        self.oracleComparer.setSourceFiles(file1, file2)
        self.oracleComparer.buildMap()
        self.oracleComparer.compareFiles(startingRange = 2)
        self.assertEqual(True, self.oracleComparer.testPassed)
    #def tearDown(self):
if __name__ == '__main__':
    unittest.main()