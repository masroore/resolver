import unittest

from runworkbook import RunWorkbook

from System.Drawing import Color


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
        workbook = Spreadsheet()
        workbook.recalculate()
        
        worksheets = [sheet.Name for sheet in workbook]
        self.assertEquals(worksheets, ['Sheet1', 'Sheet2', 'Sheet3'],
                          "Incorrect worksheets in spreadsheet")
    
    
    def testBasicSpreadsheetConfiguration(self):
        workbook = Spreadsheet()
        workbook.recalculate()
        self.assertTrue(workbook['Sheet1'].ShowBounds, "Sheet1.ShowBounds is not set to True")
        self.assertFalse(workbook['Sheet1'].ShowGrid, "Sheet1.ShowGrid is not set to False")
        
    
    def testTotalIncome(self):
        Constants = {
            'Sheet1': {
                (3, 4): 1,
                (3, 5): 2,
                (3, 6): 3,
                (3, 7): 4,
            }
        }
        workbook = Spreadsheet()
        workbook.recalculate(Constants=Constants)
        self.assertEquals(workbook['Sheet1'].C22, 10, "Total Income calculated incorrectly")

    
    def testTotalOutgoings(self):
        Constants = {
            'Sheet1': {
                (7, 4): 1,
                (7, 5): 2,
                (7, 6): 3,
                (7, 7): 4,
            }
        }
        workbook = Spreadsheet()
        workbook.recalculate(Constants=Constants)
        self.assertEquals(workbook['Sheet1'].C23, 10, "Total Outgoings calculated incorrectly")

    
    def testBalance(self):
        Constants = {
            'Sheet1': {
                (3, 4): 200,
                (7, 4): 100
            }
        }
        workbook = Spreadsheet()
        workbook.recalculate(Constants=Constants)
        self.assertEquals(workbook['Sheet1'].C24, 100, "Balance calculated incorrectly")

    
    def testBalanceBackColor(self):
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
