from openpyxl import load_workbook
import unittest
import logging
import csv

class CompareExcel:
    destColumns = []
    srcColumns = []
    numCols = 0
    wb = {}
    ws = {}
    testPassed = True
    def __init__(self, dest = "", src = "", debugging = False):
        self.debugging = debugging
        if debugging:
            if dest == "" or src == "":
                logging.warning('files not set!')
            else:
                self.setSourceFiles(dest,src)
    def setSourceFiles(self, dest, src):
        self.wb["dest"] = load_workbook(filename = dest)
        self.ws["dest"] = self.wb["dest"].active
        self.wb["src"] = load_workbook(filename = src)
        self.ws["src"] = self.wb["src"].active
    def setDataLength(self, upperRange):
        #TODO: Add code to calculate upperRange of both files and see if they are different
        self.upperRange = upperRange
    def compareFiles(self, startingRange = 2): #0 is null, 1 is heading
        rowNum=startingRange
        while rowNum<self.upperRange:
            #compare dest row and src row for each col
            self.processRow(rowNum)
            rowNum+=1
    def processRow(self,rowNum):
        #go through each column in row
        row = {}
        colNum = 0
        while colNum < self.numCols:
            row["dest"] = self.ws["dest"][self.destColumns[colNum]+str(rowNum)].value
            row["src"] = self.ws["src"][self.srcColumns[colNum]+str(rowNum)].value
            self.compareData(row["dest"],row["src"],rowNum)
            colNum+=1
    def compareData(self,d1,d2,rowNum):
        if d1 != d2:
            self.testPassed = False
            print('MISMATCH! [{v1}] does not match [{v2}] on rowNum={rn}'.format(v1=d1,v2=d2,rn=rowNum))
    def setColumns(self, destColumns, srcColumns):
        for x in destColumns:
            self.destColumns.append(x)
            
        for x in srcColumns:
            self.srcColumns.append(x)
        self.numCols = len(destColumns)