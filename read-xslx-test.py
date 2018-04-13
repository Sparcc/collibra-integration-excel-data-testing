from openpyxl import load_workbook
import unittest
import logging
import csv

class compareExcel:
    file1Cols = []
    file2Cols = []
    self.wb = {}
    self.ws = {}
    def __init__(self, file1, file2, debugging = False):
        self.debugging = debugging
        self.setSourceFiles(file1,file2)
    def setSourceFiles(self, file1, file2):
        self.wb["file1"] = load_workbook(filename = file1)
        self.ws["file1"] = self.wb["file1"].active
        self.wb["file2"] = load_workbook(filename = file2)
        self.ws["file2"] = self.wb["file2"].active
    def setDataLength(self, upperRange):
        #TODO: Add code to calculate upperRange of both files and see if they are different
        self.upperRange = upperRange
    def compareFiles(self):
        firstCol = {}
        secondCol = {}
        startingRange = 2 #0 is null, 1 is heading
        i=startingRange
        while i<self.upperRange:
            firstCol["file1"] = self.ws["file1"][self.file1Cols[0]+str(i)].value
            secondCol["file1"] = self.ws["file1"][self.file1Cols[1]+str(i)].value
            firstCol["file2"] = self.ws["file2"][self.file2Cols[0]+str(i)].value
            secondCol["file2"] = self.ws["file2"][self.file2Cols[1]+str(i)].value
            if (self.debugging):
                print(firstCol["file1"]+", ")
                print(firstCol["file2"]+"\n")
                print("------------------")
                print(secondCol["file1"]+", ")
                print(secondCol["file2"]+"\n")
            assert(firstCol["file1"] == firstCol["file2"])
            assert(secondCol["file1"] == secondCol["file2"])
            i+=1
        
    def setColumns(self, file1Cols, file2Cols):
        for x in file1Cols:
            self.file1Cols.append(x)
            
        for x in file2Cols:
            self.file2Cols.append(x)

class testcompareExcelFiles(unittest.TestCase):
    def setUp(self):
        self.file1 = {"name":"collibra-tute-test-data.xlsx"}
        self.file2 = {"name":"Default.xlsx"}
        self.comparer = compareExcel(self.file1["name"], self.file2["name"])
        self.comparer.setDataLength(52)
    def test_compareImportWithExport1(self):
        file1 = ['A','C']
        file2 = ['A','C']
        self.comparer.setColumns(file1, file2)
        self.comparer.compareFiles()
    #def tearDown(self):
if __name__ == '__main__':
    unittest.main()