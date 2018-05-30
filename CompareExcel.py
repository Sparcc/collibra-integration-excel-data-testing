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
    testPassed = True
    def __init__(self, file1 = "", file2 = "", debugging = False):
        self.debugging = debugging
        if debugging:
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
    def compareFiles(self, startingRange = 2): #0 is null, 1 is heading
        rowNum=startingRange
        while rowNum<self.upperRange:
            #compare file1 row and file2 row for each col
            self.processRow(rowNum)
            rowNum+=1
    def processRow(self,rowNum):
        #go through each column in row
        row = {}
        colNum = 0
        while colNum < self.numCols:
            row["file1"] = self.ws["file1"][self.file1Cols[colNum]+str(rowNum)].value
            row["file2"] = self.ws["file2"][self.file2Cols[colNum]+str(rowNum)].value
            self.compareData(row["file1"],row["file2"],rowNum)
            colNum+=1
    def compareData(self,d1,d2,rowNum):
        if d1 != d2:
            self.testPassed = False
            print('MISMATCH! [{v1}] does not match [{v2}] on rowNum={rn}'.format(v1=d1,v2=d2,rn=rowNum))
    def setColumns(self, file1Cols, file2Cols):
        for x in file1Cols:
            self.file1Cols.append(x)
            
        for x in file2Cols:
            self.file2Cols.append(x)
        self.numCols = len(file1Cols)