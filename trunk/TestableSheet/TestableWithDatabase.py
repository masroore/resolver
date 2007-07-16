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
                               DDE, IF, ISERROR, MAX, MEDIAN, MIN, OR, PI, SQRT,
                               Stringify, SUM, SUMIF, VALUE, VLOOKUP)
from Library.BloombergFunctions import BLP, BLPH
from Library.BloombergAPI import BloombergError
from Library.Workbook import Workbook

# Constants and Formatting
Constants = {
    'Totals': {
        (1, 1): 'Symbol',
        (3, 6): -1326.48,
        (3, 1): 'Total',
        (2, 1): 'Transactions',
        (1, 2): 'MSFT',
        (3, 2): 173482.31,
        (2, 2): 3.0,
        (1, 3): 'GOOG',
        (3, 3): 153750.2,
        (2, 3): 2.0,
        (1, 4): 'BSY',
        (3, 4): 23865.75,
        (2, 4): 2.0,
        (1, 5): 'ETI',
        (3, 5): -85947.76,
        (2, 5): 1.0,
        (1, 6): 'HNS',
        (2, 6): 2.0,
        (5, 1): 'Total:',
        (5, 2): 'Transactions:',
    },
    'Transactions': {
        (1, 1): '0001',
        (1, 2): '0002',
        (1, 3): '0003',
        (1, 4): '0004',
        (1, 5): '0005',
        (1, 6): '0006',
        (1, 7): '0007',
        (1, 8): '0008',
        (1, 9): '0009',
        (1, 10): '0010',
        (2, 1): 'A37',
        (2, 2): 'G45',
        (2, 3): 'D98',
        (2, 4): 'A01',
        (2, 5): 'C12',
        (2, 6): 'X14',
        (2, 7): 'O39',
        (2, 8): 'A37',
        (2, 9): 'C12',
        (2, 10): 'O39',
        (3, 1): 'MSFT',
        (3, 2): 'GOOG',
        (3, 3): 'BSY',
        (3, 4): 'ETI',
        (3, 5): 'HNS',
        (3, 6): 'GOOG',
        (3, 7): 'HNS',
        (3, 8): 'BSY',
        (3, 9): 'MSFT',
        (3, 10): 'MSFT',
        (4, 1): 237.0,
        (4, 2): 1250.0,
        (4, 3): 658.0,
        (4, 4): 998.0,
        (4, 5): 117.0,
        (4, 6): 4.0,
        (4, 7): 1740.0,
        (4, 8): 345.0,
        (4, 9): 2501.0,
        (4, 10): 1799.0,
        (5, 1): 48.1,
        (5, 2): 123.5,
        (5, 3): 79.5,
        (5, 4): 86.12,
        (5, 5): 0.56,
        (5, 6): 156.2,
        (5, 7): 0.8,
        (5, 8): 82.45,
        (5, 9): 43.0,
        (5, 10): 42.99,
        (6, 1): 'Buy',
        (6, 2): 'Sell',
        (6, 3): 'Sell',
        (6, 4): 'Buy',
        (6, 5): 'Sell',
        (6, 6): 'Buy',
        (6, 7): 'Buy',
        (6, 8): 'Buy',
        (6, 9): 'Sell',
        (6, 10): 'Sell',
    },
}

Formatting = {
    'Totals': {
        'BackColor': (
            ((1, 1), Color.PeachPuff),
            ((2, 1), Color.PeachPuff),
            ((3, 1), Color.PeachPuff),
        ),
        'Bold': (
            ((1, 1), True),
            ((2, 1), True),
            ((3, 1), True),
        ),
        'BorderBottom': (
            ((1, 1), True),
            ((2, 1), True),
            ((3, 1), True),
        ),
    },
    'Transactions': {
        'BackColor': (
            ((1, 1), Color.CornflowerBlue),
            ((1, 3), Color.CornflowerBlue),
            ((1, 5), Color.CornflowerBlue),
            ((1, 7), Color.CornflowerBlue),
            ((1, 9), Color.CornflowerBlue),
            ((2, 1), Color.CornflowerBlue),
            ((2, 3), Color.CornflowerBlue),
            ((2, 5), Color.CornflowerBlue),
            ((2, 7), Color.CornflowerBlue),
            ((2, 9), Color.CornflowerBlue),
            ((3, 1), Color.CornflowerBlue),
            ((3, 3), Color.CornflowerBlue),
            ((3, 5), Color.CornflowerBlue),
            ((3, 7), Color.CornflowerBlue),
            ((3, 9), Color.CornflowerBlue),
            ((4, 1), Color.CornflowerBlue),
            ((4, 3), Color.CornflowerBlue),
            ((4, 5), Color.CornflowerBlue),
            ((4, 7), Color.CornflowerBlue),
            ((4, 9), Color.CornflowerBlue),
            ((5, 1), Color.CornflowerBlue),
            ((5, 3), Color.CornflowerBlue),
            ((5, 5), Color.CornflowerBlue),
            ((5, 7), Color.CornflowerBlue),
            ((5, 9), Color.CornflowerBlue),
            ((6, 1), Color.CornflowerBlue),
            ((6, 3), Color.CornflowerBlue),
            ((6, 5), Color.CornflowerBlue),
            ((6, 7), Color.CornflowerBlue),
            ((6, 9), Color.CornflowerBlue),
        ),
        'Bold': (
            ((1, 1), True),
            ((1, 2), True),
            ((1, 3), True),
            ((1, 4), True),
            ((1, 5), True),
            ((1, 6), True),
            ((1, 7), True),
            ((1, 8), True),
            ((1, 9), True),
            ((1, 10), True),
        ),
        'BorderBottom': (
            ((1, 10), True),
            ((2, 10), True),
            ((3, 10), True),
            ((4, 10), True),
            ((5, 10), True),
            ((6, 10), True),
        ),
        'BorderLeft': (
            ((1, 1), True),
            ((1, 2), True),
            ((1, 3), True),
            ((1, 4), True),
            ((1, 5), True),
            ((1, 6), True),
            ((1, 7), True),
            ((1, 8), True),
            ((1, 9), True),
            ((1, 10), True),
        ),
        'BorderRight': (
            ((1, 1), True),
            ((1, 2), True),
            ((1, 3), True),
            ((1, 4), True),
            ((1, 5), True),
            ((1, 6), True),
            ((1, 7), True),
            ((1, 8), True),
            ((1, 9), True),
            ((1, 10), True),
            ((2, 1), True),
            ((2, 2), True),
            ((2, 3), True),
            ((2, 4), True),
            ((2, 5), True),
            ((2, 6), True),
            ((2, 7), True),
            ((2, 8), True),
            ((2, 9), True),
            ((2, 10), True),
            ((3, 1), True),
            ((3, 2), True),
            ((3, 3), True),
            ((3, 4), True),
            ((3, 5), True),
            ((3, 6), True),
            ((3, 7), True),
            ((3, 8), True),
            ((3, 9), True),
            ((3, 10), True),
            ((4, 1), True),
            ((4, 2), True),
            ((4, 3), True),
            ((4, 4), True),
            ((4, 5), True),
            ((4, 6), True),
            ((4, 7), True),
            ((4, 8), True),
            ((4, 9), True),
            ((4, 10), True),
            ((5, 1), True),
            ((5, 2), True),
            ((5, 3), True),
            ((5, 4), True),
            ((5, 5), True),
            ((5, 6), True),
            ((5, 7), True),
            ((5, 8), True),
            ((5, 9), True),
            ((5, 10), True),
            ((6, 1), True),
            ((6, 2), True),
            ((6, 3), True),
            ((6, 4), True),
            ((6, 5), True),
            ((6, 6), True),
            ((6, 7), True),
            ((6, 8), True),
            ((6, 9), True),
            ((6, 10), True),
        ),
        'BorderTop': (
            ((1, 1), True),
            ((1, 2), True),
            ((1, 3), True),
            ((1, 4), True),
            ((1, 5), True),
            ((1, 6), True),
            ((1, 7), True),
            ((1, 8), True),
            ((1, 9), True),
            ((1, 10), True),
            ((2, 1), True),
            ((2, 2), True),
            ((2, 3), True),
            ((2, 4), True),
            ((2, 5), True),
            ((2, 6), True),
            ((2, 7), True),
            ((2, 8), True),
            ((2, 9), True),
            ((2, 10), True),
            ((3, 1), True),
            ((3, 2), True),
            ((3, 3), True),
            ((3, 4), True),
            ((3, 5), True),
            ((3, 6), True),
            ((3, 7), True),
            ((3, 8), True),
            ((3, 9), True),
            ((3, 10), True),
            ((4, 1), True),
            ((4, 2), True),
            ((4, 3), True),
            ((4, 4), True),
            ((4, 5), True),
            ((4, 6), True),
            ((4, 7), True),
            ((4, 8), True),
            ((4, 9), True),
            ((4, 10), True),
            ((5, 1), True),
            ((5, 2), True),
            ((5, 3), True),
            ((5, 4), True),
            ((5, 5), True),
            ((5, 6), True),
            ((5, 7), True),
            ((5, 8), True),
            ((5, 9), True),
            ((5, 10), True),
            ((6, 1), True),
            ((6, 2), True),
            ((6, 3), True),
            ((6, 4), True),
            ((6, 5), True),
            ((6, 6), True),
            ((6, 7), True),
            ((6, 8), True),
            ((6, 9), True),
            ((6, 10), True),
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
        workbook.AddWorksheet("Totals")
        workbook["Totals"].Bounds = (1, 1, 6, 6)
        workbook["Totals"].Cols["B"].Width = 91
        workbook["Totals"].Cols["E"].Width = 105
        workbook["Totals"].ShowBounds = True
        workbook["Totals"].ShowGrid = False
        
        workbook.AddWorksheet("Transactions")
        workbook["Transactions"].Bounds = (1, 1, 6, 10)
        

        # Pre-constants user code
        
        totals = workbook['Totals']
        cellRange = CellRange(totals.A1, totals.C6)
        cellRange.HeaderRow = cellRange.Rows[1]
        cellRange.HeaderRow.Bold = True
        cellRange.HeaderRow.BorderBottom = True
        cellRange.HeaderRow.BackColor = Color.PeachPuff
        
        cellRange.HeaderRow[1].Value = 'Symbol'
        cellRange.HeaderRow[2].Value = 'Transactions'
        cellRange.HeaderRow[3].Value = 'Total'
        
        connectString = (
            "DRIVER={MySQL ODBC 3.51 Driver};"
            "SERVER=financial_data;"
            "PORT=3306;"
            "DATABASE=transactions;"
            "USER=username;"
            "PASSWORD=password;"
            "OPTION=3;"
        )
        
        query = "SELECT * FROM shares"
        
        transactions = workbook.AddWorksheetFromDB("Transactions", connectString, query)
        

        workbook.Populate(Constants, Formatting)

        # Pre-formulae user code
        
        values = {}
        for row in transactions.Rows:
            modifier = 1
            if row['F'].Value == 'Buy':
                modifier = -1
        
            symbol = row['C'].Value
            currentTotal, numberTransactions = values.setdefault(symbol, (0, 0))
        
            thisTransaction = row['D'].Value * row['E'].Value * modifier
            values[symbol] = currentTotal + thisTransaction, numberTransactions + 1
        
        for index, (symbol, (total, number)) in enumerate(values.items()):
            row = cellRange.Rows[index + 2]
            row[1].Value = symbol
            row[2].Value = number
            row[3].Value = total    
        

        # Formula code
        
        workbook["Totals"].F1.Value = SUM(CellRange(workbook["Totals"].C1,workbook["Totals"].C20))
        
        workbook["Totals"].F2.Value = SUM(CellRange(workbook["Totals"].B1,workbook["Totals"].B20))
        

        # Post-formulae user code
        
        
        
            
        
        
