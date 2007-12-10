import LoadRequiredAssemblies

# Import statements
from System.Drawing import Color, SystemColors
from Library.Button import Button
from Library.CellRange import CellRange
from Library.Date import Date
from Library.Empty import Empty
from Library.ExternalLinks import GetExternalCell
from Library.Errors import (ArrayError, CycleError, DatabaseError, FetchingDataWarning, 
                            FormulaError, LogicError, NotFoundError)
from Library.Mappers import MapWorksheets
from Library.Functions import (Aggregate, AND, AVERAGE, CONCATENATE, CopyRange, COUNTIF,
                               DDE, DEFAULT_PARAMETER, IF, ImplicitConvert, ISERROR, MAX,
                               MEDIAN, MIN, OR, PI, SQRT, Stringify, SUM, SUMIF, VALUE,
                               VLOOKUP)
from Library.BloombergFunctions import BLP, BLPH
from Library.BloombergAPI import BloombergError
from Library.Workbook import Workbook

# Constants and Formatting
Constants = {
    'Sheet1': {
        (1, 1): 'Income',
        (1, 2): 'Reason',
        (1, 4): 'Winnings from Horses',
        (1, 5): 'Jerry Paid Back Loan',
        (1, 6): 'My Wages May',
        (1, 7): 'Cashback on Phone',
        (1, 8): "Wife's Wages May",
        (1, 9): '????',
        (1, 10): 'Business May',
        (1, 11): 'Sale of Camera',
        (1, 12): 'Royalties',
        (1, 22): 'Total Income:',
        (1, 23): 'Total Outgoings:',
        (1, 24): 'Balance:',
        (2, 2): 'Date',
        (2, 4): Date('27/05/2007'),
        (2, 5): Date('29/05/2007'),
        (2, 6): Date('01/06/2007'),
        (2, 7): Date('03/06/2007'),
        (2, 8): Date('04/06/2007'),
        (2, 9): Date('06/06/2007'),
        (2, 10): Date('07/06/2007'),
        (2, 11): Date('08/06/2007'),
        (2, 12): Date('15/06/2007'),
        (3, 2): 'Amount',
        (3, 4): 1240.06,
        (3, 5): 480.0,
        (3, 6): 1800.27,
        (3, 7): 45.0,
        (3, 8): 851.42,
        (3, 9): 225.51,
        (3, 10): 450.0,
        (3, 11): 35.0,
        (3, 12): 412.09,
        (5, 1): 'Outgoings',
        (5, 2): 'Reason',
        (5, 4): 'Chocolate',
        (5, 5): 'Present for Mrs',
        (5, 6): '????',
        (5, 7): 'Cash ',
        (5, 8): 'Rent',
        (5, 9): 'Tescos',
        (5, 10): '????',
        (5, 11): 'Cash ',
        (5, 12): 'Council Tax',
        (5, 13): 'Savings',
        (5, 14): 'Phone Bill',
        (5, 15): 'Power Bill',
        (5, 16): '????',
        (5, 17): 'Cash',
        (5, 18): 'Tescos',
        (5, 19): 'House Repairs',
        (5, 20): 'Car Repairs',
        (6, 2): 'Date',
        (6, 4): Date('15/05/2007'),
        (6, 5): Date('15/05/2007'),
        (6, 6): Date('18/05/2007'),
        (6, 7): Date('25/05/2007'),
        (6, 8): Date('01/06/2007'),
        (6, 9): Date('02/06/2007'),
        (6, 10): Date('02/06/2007'),
        (6, 11): Date('04/06/2007'),
        (6, 12): Date('04/06/2007'),
        (6, 13): Date('05/06/2007'),
        (6, 14): Date('07/06/2007'),
        (6, 15): Date('09/06/2007'),
        (6, 16): Date('10/06/2007'),
        (6, 17): Date('10/06/2007'),
        (6, 18): Date('10/06/2007'),
        (6, 19): Date('11/06/2007'),
        (6, 20): Date('12/06/2007'),
        (7, 2): 'Amount',
        (7, 4): 98.5,
        (7, 5): 41.23,
        (7, 6): 112.45,
        (7, 7): 120.0,
        (7, 8): 450.0,
        (7, 9): 27.32,
        (7, 10): 12.45,
        (7, 11): 40.0,
        (7, 12): 85.0,
        (7, 13): 400.0,
        (7, 14): 27.51,
        (7, 15): 81.0,
        (7, 16): 24.32,
        (7, 17): 30.0,
        (7, 18): 49.51,
        (7, 19): 2940.06,
        (7, 20): 1001.37,
    },
}

Formatting = {
    'Sheet1': {
        'ColWidth': (
            ((1, 0), 155),
            ((5, 0), 173),
            ((9, 0), 74),
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
            if sheet['A', row] is Empty:
                break
            sheet.Cells['A', row].BorderLeft = True
            sheet.Cells['B', row].BorderLeft = True
            sheet.Cells['B', row].BorderRight = True
            sheet.Cells['C', row].BorderRight = True
        
        for row in range(4, sheet.MaxRow+1):
            # First the Income table
            if sheet['E', row] is Empty:
                break
            # Next the outgoings column
            sheet.Cells['E', row].BorderLeft = True
            sheet.Cells['F', row].BorderLeft = True
            sheet.Cells['F', row].BorderRight = True
            sheet.Cells['G', row].BorderRight = True
        
        

        # Formula code
        
        workbook["Sheet1"].C22 = SUM(CellRange(workbook["Sheet1"].Cells.C1,workbook["Sheet1"].Cells.C21))
        
        workbook["Sheet1"].C23 = SUM(CellRange(workbook["Sheet1"].Cells.G2,workbook["Sheet1"].Cells.G21))
        
        workbook["Sheet1"].C24 = workbook["Sheet1"].C22-workbook["Sheet1"].C23
        

        # Post-formulae user code
        
        if sheet.C24 < 0:
            sheet.Cells.C24.BackColor = Color.Red
