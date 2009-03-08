from System.IO import Directory, Path

this_dir = Directory.GetCurrentDirectory()

# import this to load the correct standard library
# As a side effect it changes directory
from runworkbook import RunWorkbook

# Needed for unittest - not shipped with Resolver One
import sys
sys.path.append(r'C:\Python25\Lib')
import unittest

from System.Drawing import Color

from Library.CellRange import CellRange
from Library.Worksheet import Worksheet

Directory.SetCurrentDirectory(this_dir)
spreadsheet_path = Path.Combine(this_dir, 'TestableSheet.rsl')


class TestableSheetTest(unittest.TestCase):
    
    def assertEquals(self, actual, expected, message=''):
        message += ':\nactual = "%s"\nexpected = "%s"\n' % (actual, expected)
        unittest.TestCase.assertEquals(self, actual, expected, message)
        
    def assertTrue(self, arg, message=''):
        unittest.TestCase.assertEquals(self, arg, True, message)
        
    def assertFalse(self, arg, message=''):
        unittest.TestCase.assertEquals(self, arg, False, message)
        
    ##

    def testWorksheets(self):
        workbook = RunWorkbook(spreadsheet_path)
        
        worksheets = [sheet.Name for sheet in workbook]
        self.assertEquals(worksheets, ['Data'],
                          "Incorrect worksheets in spreadsheet")
    
    
    def testBasicSpreadsheetConfiguration(self):
        workbook = RunWorkbook(spreadsheet_path)
        self.assertTrue(workbook['Data'].ShowBounds, "Data.ShowBounds is not set to True")
        self.assertFalse(workbook['Data'].ShowGrid, "Data.ShowGrid is not set to False")
        
    
    def testTotalIncome(self):
        temp = Worksheet("Temporary Sheet")
        cellrange = CellRange(temp.Cells.A1, temp.Cells.A9)
        for i, val in enumerate((1, 2, 3, 4)):
            cellrange[1, i+1] = val
        
        workbook = RunWorkbook(spreadsheet_path, income=cellrange)
        self.assertEquals(workbook.totalincome, 10, "Total Income calculated incorrectly")

    
    def testTotalOutgoings(self):
        temp = Worksheet("Temporary Sheet")
        cellrange = CellRange(temp.Cells.A1, temp.Cells.A17)
        for i, val in enumerate((1, 2, 3, 4)):
            cellrange[1, i+1] = val
            
        workbook = RunWorkbook(spreadsheet_path, outgoings=cellrange)
        self.assertEquals(workbook.totaloutgoings, 10, "Total Outgoings calculated incorrectly")

    
    def DONTtestBalance(self):
        Constants = {
            'Sheet1': {
                (3, 4): 200,
                (7, 4): 100
            }
        }
        workbook = Spreadsheet()
        workbook.recalculate(Constants=Constants)
        self.assertEquals(workbook['Sheet1'].C24, 100, "Balance calculated incorrectly")

    
    def DONTtestBalanceBackColor(self):
        PositiveBalance = {
            'Sheet1': {
                (3, 4): 100,
                (7, 4): 99
            }
        }
        workbook = Spreadsheet()
        workbook.recalculate(Constants=PositiveBalance)
        self.assertNotEquals(workbook['Sheet1'].Cells.C24.BackColor, Color.Red, "Incorrect BackColor for positive balance")
        
        NegativeBalance = {
            'Sheet1': {
                (3, 4): 99,
                (7, 4): 100
            }
        }
        workbook.recalculate(Constants=NegativeBalance)
        self.assertEquals(workbook['Sheet1'].Cells.C24.BackColor, Color.Red, "Incorrect BackColor for negative balance")
        


if __name__ == '__main__':
    suite = unittest.TestSuite()
    suite.addTests(unittest.makeSuite(TestableSheetTest))

    unittest.TextTestRunner(verbosity=2).run(suite)
