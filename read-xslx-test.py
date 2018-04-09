from openpyxl import load_workbook
import unittest
import logging
import csv
class compareExcelFiles(unittest.TestCase):
    def setUp(self):
        logging.warning('setUp not tested')
    def test_compareImportWithExport(self):
        debugging = False
        wb = load_workbook(filename = 'collibra-tute-test-data.xlsx')
        ws=wb.active
        wb2 = load_workbook(filename = 'Default.xlsx')
        ws2=wb2.active
        i=2 #ignoring headings
        rowString=""
        #ofile  = open('custom-testdata.csv', "w", newline='')
        #writer = csv.writer(ofile)
        firstCol = ["",""]
        secondCol = ["",""]
        while i<51:
            if (debugging):
                rowString=""
                rowString+=ws['A'+str(i)].value
                rowString+=", "
                rowString+=ws['C'+str(i)].value
                rowString+="--Compared To Export From Collibra-->"
                rowString+=ws2['A'+str(i)].value
                rowString+=", "
                rowString+=ws2['C'+str(i)].value
                print(rowString)
            firstCol[0] = ws['A'+str(i)].value
            firstCol[1] = ws2['A'+str(i)].value
            secondCol[0] = ws['C'+str(i)].value
            secondCol[1] = ws2['C'+str(i)].value
            self.assertEqual(firstCol[0],firstCol[1])
            self.assertEqual(secondCol[0],secondCol[1])
            i+= 1
    def tearDown(self):
        logging.warning('tearDown not tested')
if __name__ == '__main__':
    unittest.main()