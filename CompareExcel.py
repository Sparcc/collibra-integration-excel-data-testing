from openpyxl import load_workbook
import unittest
import logging
import csv

class CompareExcel:
    file1Cols = []
    file2Cols = []
    numCols = 0
    wb = {}
    ws = {}
    def __init__(self, file1 = "", file2 = "", debugging = False):
        self.debugging = debugging
        if file1 == "" or file2 == "":
            logging.warning('files not set!')
        else:
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
        row = {}
        startingRange = 2 #0 is null, 1 is heading
        i=startingRange
        while i<self.upperRange:
            #compare file1 row and file2 row for each col
            i2 = 0
            while i2 < self.numCols:
                row["file1"] = self.ws["file1"][self.file1Cols[i2]+str(i)].value
                row["file2"] = self.ws["file2"][self.file2Cols[i2]+str(i)].value
                assert(row["file1"] == row["file2"])
                i2+=1
            i+=1
        
    def setColumns(self, file1Cols, file2Cols):
        for x in file1Cols:
            self.file1Cols.append(x)
            
        for x in file2Cols:
            self.file2Cols.append(x)
        self.numCols = len(file1Cols)