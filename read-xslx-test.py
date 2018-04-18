import sys, os
sys.path.append(os.getcwd())
from CompareExcel import *

class testcompareExcelIntegrationData(unittest.TestCase):
    def setUp(self):
        #self.comparer = CompareExcel(file1, file2)
        self.comparer = CompareExcel()
        self.comparer.setDataLength(52)
    def test_CompareExcel(self):
        file1 = "CompareExcel-test-file1.xlsx"
        file2 = "CompareExcel-test-file2.xlsx"
        self.comparer.setSourceFiles(file1, file2)
        file1 = ['A','C','D']
        file2 = ['A','C','F']
        self.comparer.setColumns(file1, file2)
        self.comparer.compareFiles()
    #def tearDown(self):
if __name__ == '__main__':
    unittest.main()