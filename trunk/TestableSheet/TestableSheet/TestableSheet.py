import LoadRequiredAssemblies

# Import statements
from System.Drawing import Color, SystemColors
from Library.Button import Button
from Library.CellRange import CellRange
from Library.Date import Date
from Library.Empty import Empty
from Library.ExternalLinks import GetExternalCell
from Library.Errors import (CycleError, DatabaseError, FetchingDataWarning, FormulaError,
                            LogicError, NotFoundError)
from Library.Mappers import MapWorksheets
from Library.Functions import (Aggregate, AND, AVERAGE, CONCATENATE, CopyRange, COUNTIF,
                               IF, ISERROR, MAX, MEDIAN, MIN, OR, PI, SQRT,
                               Stringify, SUM, SUMIF, VALUE, VLOOKUP)
from Library.BloombergFunctions import BLP, BLPH
from Library.BloombergAPI import BloombergError
from Library.Workbook import Workbook

# Constants and Formatting
Constants = {
    'Sheet1': {
        (5, 5): 'Present for Mrs',
        (7, 4): 98.5,
        (6, 4): Date('15/05/2007'),
        (5, 4): 'Chocolate',
        (7, 2): 'Amount',
        (6, 2): 'Date',
        (5, 2): 'Reason',
        (5, 1): 'Outgoings',
        (3, 10): 450.0,
        (3, 9): 225.51,
        (3, 8): 851.42,
        (3, 6): 1800.27,
        (3, 5): 480.0,
        (3, 4): 1240.06,
        (3, 2): 'Amount',
        (2, 10): Date('07/06/2007'),
        (2, 9): Date('06/06/2007'),
        (2, 8): Date('04/06/2007'),
        (2, 6): Date('01/06/2007'),
        (2, 5): Date('29/05/2007'),
        (2, 4): Date('27/05/2007'),
        (2, 2): 'Date',
        (1, 10): 'Business May',
        (1, 9): '????',
        (1, 8): "Wife's Wages May",
        (1, 6): 'My Wages May',
        (1, 5): 'Jerry Paid Back Loan',
        (1, 4): 'Winnings from Horses',
        (1, 2): 'Reason',
        (1, 1): 'Income',
        (1, 7): 'Cashback on Phone',
        (2, 7): Date('03/06/2007'),
        (3, 7): 45.0,
        (6, 5): Date('15/05/2007'),
        (7, 5): 41.23,
        (5, 6): '????',
        (6, 6): Date('18/05/2007'),
        (7, 6): 112.45,
        (5, 7): 'Cash ',
        (6, 7): Date('25/05/2007'),
        (7, 7): 120.0,
        (5, 8): 'Rent',
        (6, 8): Date('01/06/2007'),
        (7, 8): 450.0,
        (5, 9): 'Tescos',
        (6, 9): Date('02/06/2007'),
        (7, 9): 27.32,
        (5, 10): '????',
        (6, 10): Date('02/06/2007'),
        (7, 10): 12.45,
        (5, 11): 'Cash ',
        (6, 11): Date('04/06/2007'),
        (7, 11): 40.0,
        (5, 12): 'Council Tax',
        (6, 12): Date('04/06/2007'),
        (7, 12): 85.0,
        (5, 13): 'Savings',
        (6, 13): Date('05/06/2007'),
        (7, 13): 400.0,
        (5, 14): 'Phone Bill',
        (6, 14): Date('07/06/2007'),
        (7, 14): 27.51,
        (6, 15): Date('09/06/2007'),
        (5, 18): 'Tescos',
        (5, 17): 'Cash',
        (5, 16): '????',
        (5, 15): 'Power Bill',
        (3, 12): 412.09,
        (3, 11): 35.0,
        (2, 12): Date('15/06/2007'),
        (2, 11): Date('08/06/2007'),
        (1, 12): 'Royalties',
        (1, 11): 'Sale of Camera',
        (6, 16): Date('10/06/2007'),
        (6, 17): Date('10/06/2007'),
        (6, 18): Date('10/06/2007'),
        (7, 15): 81.0,
        (7, 16): 24.32,
        (7, 17): 30.0,
        (7, 18): 49.51,
        (1, 22): 'Total Income:',
        (5, 19): 'House Repairs',
        (6, 19): Date('11/06/2007'),
        (7, 19): 2940.06,
        (5, 20): 'Car Repairs',
        (6, 20): Date('12/06/2007'),
        (7, 20): 1001.37,
        (1, 23): 'Total Outgoings:',
        (1, 24): 'Balance:',
    },
}

Formatting = {
    'Sheet1': {
        'Bold': (
            ((7, 2), True),
            ((6, 2), True),
            ((5, 2), True),
            ((5, 1), True),
            ((3, 2), True),
            ((2, 2), True),
            ((1, 2), True),
            ((1, 1), True),
            ((1, 22), True),
            ((1, 23), True),
            ((1, 24), True),
            ((7, 1), True),
            ((6, 1), True),
            ((4, 2), True),
            ((4, 1), True),
            ((3, 1), True),
            ((2, 1), True),
        ),
        'BorderBottom': (
            ((7, 2), True),
            ((6, 2), True),
            ((5, 2), True),
            ((5, 1), True),
            ((3, 2), True),
            ((2, 2), True),
            ((1, 2), True),
            ((1, 1), True),
            ((7, 1), True),
            ((6, 1), True),
            ((3, 1), True),
            ((2, 1), True),
        ),
        'BorderLeft': (
            ((5, 5), True),
            ((6, 4), True),
            ((5, 4), True),
            ((7, 2), True),
            ((6, 2), True),
            ((5, 2), True),
            ((5, 1), True),
            ((2, 10), True),
            ((2, 9), True),
            ((2, 8), True),
            ((2, 6), True),
            ((2, 5), True),
            ((2, 4), True),
            ((2, 2), True),
            ((1, 10), True),
            ((1, 9), True),
            ((1, 8), True),
            ((1, 6), True),
            ((1, 5), True),
            ((1, 4), True),
            ((1, 2), True),
            ((1, 1), True),
            ((1, 7), True),
            ((2, 7), True),
            ((6, 5), True),
            ((5, 6), True),
            ((6, 6), True),
            ((5, 7), True),
            ((6, 7), True),
            ((5, 8), True),
            ((6, 8), True),
            ((5, 9), True),
            ((6, 9), True),
            ((5, 10), True),
            ((6, 10), True),
            ((5, 11), True),
            ((6, 11), True),
            ((5, 12), True),
            ((6, 12), True),
            ((5, 13), True),
            ((6, 13), True),
            ((5, 14), True),
            ((6, 14), True),
            ((6, 15), True),
            ((5, 18), True),
            ((5, 17), True),
            ((5, 16), True),
            ((5, 15), True),
            ((2, 12), True),
            ((2, 11), True),
            ((1, 12), True),
            ((1, 11), True),
            ((6, 16), True),
            ((6, 17), True),
            ((6, 18), True),
            ((5, 19), True),
            ((6, 19), True),
            ((5, 20), True),
            ((6, 20), True),
            ((1, 3), True),
            ((2, 3), True),
            ((5, 3), True),
            ((6, 3), True),
        ),
        'BorderRight': (
            ((7, 4), True),
            ((6, 4), True),
            ((7, 2), True),
            ((3, 10), True),
            ((3, 9), True),
            ((3, 8), True),
            ((3, 6), True),
            ((3, 5), True),
            ((3, 4), True),
            ((3, 2), True),
            ((2, 10), True),
            ((2, 9), True),
            ((2, 8), True),
            ((2, 6), True),
            ((2, 5), True),
            ((2, 4), True),
            ((2, 2), True),
            ((2, 7), True),
            ((3, 7), True),
            ((6, 5), True),
            ((7, 5), True),
            ((6, 6), True),
            ((7, 6), True),
            ((6, 7), True),
            ((7, 7), True),
            ((6, 8), True),
            ((7, 8), True),
            ((6, 9), True),
            ((7, 9), True),
            ((6, 10), True),
            ((7, 10), True),
            ((6, 11), True),
            ((7, 11), True),
            ((6, 12), True),
            ((7, 12), True),
            ((6, 13), True),
            ((7, 13), True),
            ((6, 14), True),
            ((7, 14), True),
            ((6, 15), True),
            ((3, 12), True),
            ((3, 11), True),
            ((2, 12), True),
            ((2, 11), True),
            ((6, 16), True),
            ((6, 17), True),
            ((6, 18), True),
            ((7, 15), True),
            ((7, 16), True),
            ((7, 17), True),
            ((7, 18), True),
            ((6, 19), True),
            ((7, 19), True),
            ((6, 20), True),
            ((7, 20), True),
            ((7, 1), True),
            ((3, 1), True),
            ((2, 3), True),
            ((6, 3), True),
            ((3, 3), True),
            ((7, 3), True),
        ),
        'BorderTop': (
            ((5, 1), True),
            ((1, 1), True),
            ((7, 1), True),
            ((6, 1), True),
            ((3, 1), True),
            ((2, 1), True),
        ),
        'DecimalPlaces': (
            ((3, 24), 2),
        ),
        'StripZeros': (
            ((3, 24), False),
        ),
    },
}


class Spreadsheet(Workbook):
    def __reset(self):
        self._worksheets = []
        self._dataHooks = set()

    def recalculate(self, Constants=Constants, Formatting=Formatting):
        self.__reset()
        workbook = self

        # Worksheet creation
        workbook.AddWorksheet("Sheet1")
        workbook["Sheet1"].Bounds = (1, 1, 7, 24)
        workbook["Sheet1"].Cols["A"].Width = 155
        workbook["Sheet1"].Cols["E"].Width = 173
        workbook["Sheet1"].Cols["I"].Width = 74
        workbook["Sheet1"].ShowBounds = True
        workbook["Sheet1"].ShowGrid = False
        
        workbook.AddWorksheet("Sheet2")
        
        workbook.AddWorksheet("Sheet3")
        

        # Pre-constants user code
        

        workbook.Populate(Constants, Formatting)

        # Pre-formulae user code
        
        sheet = workbook['Sheet1']
        
        # Next let's put nice borders down all the entries automatically
        for row in range(4, sheet.MaxRow+1):
            # First the Income table
            if sheet['A', row].Value is Empty:
                break
            sheet['A', row].BorderLeft = True
            sheet['B', row].BorderLeft = True
            sheet['B', row].BorderRight = True
            sheet['C', row].BorderRight = True
        
        for row in range(4, sheet.MaxRow+1):
            # First the Income table
            if sheet['E', row].Value is Empty:
                break
            # Next the outgoings column
            sheet['E', row].BorderLeft = True
            sheet['F', row].BorderLeft = True
            sheet['F', row].BorderRight = True
            sheet['G', row].BorderRight = True
        
        

        # Formula code
        
        workbook["Sheet1"].C22.Value = SUM(CellRange(workbook["Sheet1"].C1,workbook["Sheet1"].C21))
        
        workbook["Sheet1"].C23.Value = SUM(CellRange(workbook["Sheet1"].G2,workbook["Sheet1"].G21))
        
        workbook["Sheet1"].C24.Value = workbook["Sheet1"].C22.Value-workbook["Sheet1"].C23.Value
        

        # Post-formulae user code
        
        if sheet.C24.Value < 0:
            sheet.C24.BackColor = Color.Red
