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
workbook = Workbook()

# Worksheet creation
workbook.AddWorksheet("Colored Range")
workbook["Colored Range"].Bounds = (3, 5, 8, 15)
workbook["Colored Range"].ShowGrid = False

workbook.AddWorksheet("Silly Colors - No Grid")
workbook["Silly Colors - No Grid"].ShowGrid = False

workbook.AddWorksheet("The Silly Colors")

# Pre-constants user code
from System.Drawing import Image
from System.IO import Path
directory = Path.GetDirectoryName(__file__)
imagePath = Path.Combine(directory, '3d-space.jpg')
image = Image.FromFile(imagePath)
workbook.AddImageWorksheet("Gnuplot Image", image)

# Constants and Formatting
Constants = {
    'Colored Range': {
        (3, 5): 'Reference',
        (3, 6): '0001',
        (3, 7): '0002',
        (3, 8): '0003',
        (3, 9): '0004',
        (3, 10): '0005',
        (3, 11): '0006',
        (3, 12): '0007',
        (3, 13): '0008',
        (3, 14): '0009',
        (3, 15): '0010',
        (4, 5): 'Customer',
        (4, 6): 'A37',
        (4, 7): 'G45',
        (4, 8): 'D98',
        (4, 9): 'A01',
        (4, 10): 'C12',
        (4, 11): 'X14',
        (4, 12): 'O39',
        (4, 13): 'A37',
        (4, 14): 'C12',
        (4, 15): 'O39',
        (5, 5): 'Symbol',
        (5, 6): 'MSFT',
        (5, 7): 'GOOG',
        (5, 8): 'BSY',
        (5, 9): 'ETI',
        (5, 10): 'HNS',
        (5, 11): 'GOOG',
        (5, 12): 'HNS',
        (5, 13): 'BSY',
        (5, 14): 'MSFT',
        (5, 15): 'MSFT',
        (6, 5): 'Number',
        (6, 6): 237.0,
        (6, 7): 1250.0,
        (6, 8): 658.0,
        (6, 9): 998.0,
        (6, 10): 117.0,
        (6, 11): 4.0,
        (6, 12): 1740.0,
        (6, 13): 345.0,
        (6, 14): 2501.0,
        (6, 15): 1799.0,
        (7, 5): 'Price',
        (7, 6): 48.1,
        (7, 7): 123.5,
        (7, 8): 79.5,
        (7, 9): 86.12,
        (7, 10): 0.56,
        (7, 11): 156.2,
        (7, 12): 0.8,
        (7, 13): 82.45,
        (7, 14): 43.0,
        (7, 15): 42.99,
        (8, 5): 'Direction',
        (8, 6): 'Buy',
        (8, 7): 'Sell',
        (8, 8): 'Sell',
        (8, 9): 'Buy',
        (8, 10): 'Sell',
        (8, 11): 'Buy',
        (8, 12): 'Buy',
        (8, 13): 'Buy',
        (8, 14): 'Sell',
        (8, 15): 'Sell',
    },
}

Formatting = {
    'Colored Range': {
        'BackColor': (
            ((3, 6), Color.CornflowerBlue),
            ((3, 8), Color.CornflowerBlue),
            ((3, 10), Color.CornflowerBlue),
            ((3, 12), Color.CornflowerBlue),
            ((3, 14), Color.CornflowerBlue),
            ((4, 6), Color.CornflowerBlue),
            ((4, 8), Color.CornflowerBlue),
            ((4, 10), Color.CornflowerBlue),
            ((4, 12), Color.CornflowerBlue),
            ((4, 14), Color.CornflowerBlue),
            ((5, 6), Color.CornflowerBlue),
            ((5, 8), Color.CornflowerBlue),
            ((5, 10), Color.CornflowerBlue),
            ((5, 12), Color.CornflowerBlue),
            ((5, 14), Color.CornflowerBlue),
            ((6, 6), Color.CornflowerBlue),
            ((6, 8), Color.CornflowerBlue),
            ((6, 10), Color.CornflowerBlue),
            ((6, 12), Color.CornflowerBlue),
            ((6, 14), Color.CornflowerBlue),
            ((7, 6), Color.CornflowerBlue),
            ((7, 8), Color.CornflowerBlue),
            ((7, 10), Color.CornflowerBlue),
            ((7, 12), Color.CornflowerBlue),
            ((7, 14), Color.CornflowerBlue),
            ((8, 6), Color.CornflowerBlue),
            ((8, 8), Color.CornflowerBlue),
            ((8, 10), Color.CornflowerBlue),
            ((8, 12), Color.CornflowerBlue),
            ((8, 14), Color.CornflowerBlue),
        ),
        'Bold': (
            ((3, 5), True),
            ((3, 6), True),
            ((3, 7), True),
            ((3, 8), True),
            ((3, 9), True),
            ((3, 10), True),
            ((3, 11), True),
            ((3, 12), True),
            ((3, 13), True),
            ((3, 14), True),
            ((3, 15), True),
            ((4, 5), True),
            ((5, 5), True),
            ((6, 5), True),
            ((7, 5), True),
            ((8, 5), True),
        ),
        'BorderBottom': (
            ((3, 15), True),
            ((4, 15), True),
            ((5, 15), True),
            ((6, 15), True),
            ((7, 15), True),
            ((8, 15), True),
        ),
        'BorderLeft': (
            ((3, 5), True),
            ((3, 6), True),
            ((3, 7), True),
            ((3, 8), True),
            ((3, 9), True),
            ((3, 10), True),
            ((3, 11), True),
            ((3, 12), True),
            ((3, 13), True),
            ((3, 14), True),
            ((3, 15), True),
        ),
        'BorderRight': (
            ((3, 5), True),
            ((3, 6), True),
            ((3, 7), True),
            ((3, 8), True),
            ((3, 9), True),
            ((3, 10), True),
            ((3, 11), True),
            ((3, 12), True),
            ((3, 13), True),
            ((3, 14), True),
            ((3, 15), True),
            ((4, 5), True),
            ((4, 6), True),
            ((4, 7), True),
            ((4, 8), True),
            ((4, 9), True),
            ((4, 10), True),
            ((4, 11), True),
            ((4, 12), True),
            ((4, 13), True),
            ((4, 14), True),
            ((4, 15), True),
            ((5, 5), True),
            ((5, 6), True),
            ((5, 7), True),
            ((5, 8), True),
            ((5, 9), True),
            ((5, 10), True),
            ((5, 11), True),
            ((5, 12), True),
            ((5, 13), True),
            ((5, 14), True),
            ((5, 15), True),
            ((6, 5), True),
            ((6, 6), True),
            ((6, 7), True),
            ((6, 8), True),
            ((6, 9), True),
            ((6, 10), True),
            ((6, 11), True),
            ((6, 12), True),
            ((6, 13), True),
            ((6, 14), True),
            ((6, 15), True),
            ((7, 5), True),
            ((7, 6), True),
            ((7, 7), True),
            ((7, 8), True),
            ((7, 9), True),
            ((7, 10), True),
            ((7, 11), True),
            ((7, 12), True),
            ((7, 13), True),
            ((7, 14), True),
            ((7, 15), True),
            ((8, 5), True),
            ((8, 6), True),
            ((8, 7), True),
            ((8, 8), True),
            ((8, 9), True),
            ((8, 10), True),
            ((8, 11), True),
            ((8, 12), True),
            ((8, 13), True),
            ((8, 14), True),
            ((8, 15), True),
        ),
        'BorderTop': (
            ((3, 5), True),
            ((3, 6), True),
            ((3, 7), True),
            ((3, 8), True),
            ((3, 9), True),
            ((3, 10), True),
            ((3, 11), True),
            ((3, 12), True),
            ((3, 13), True),
            ((3, 14), True),
            ((3, 15), True),
            ((4, 5), True),
            ((4, 6), True),
            ((4, 7), True),
            ((4, 8), True),
            ((4, 9), True),
            ((4, 10), True),
            ((4, 11), True),
            ((4, 12), True),
            ((4, 13), True),
            ((4, 14), True),
            ((4, 15), True),
            ((5, 5), True),
            ((5, 6), True),
            ((5, 7), True),
            ((5, 8), True),
            ((5, 9), True),
            ((5, 10), True),
            ((5, 11), True),
            ((5, 12), True),
            ((5, 13), True),
            ((5, 14), True),
            ((5, 15), True),
            ((6, 5), True),
            ((6, 6), True),
            ((6, 7), True),
            ((6, 8), True),
            ((6, 9), True),
            ((6, 10), True),
            ((6, 11), True),
            ((6, 12), True),
            ((6, 13), True),
            ((6, 14), True),
            ((6, 15), True),
            ((7, 5), True),
            ((7, 6), True),
            ((7, 7), True),
            ((7, 8), True),
            ((7, 9), True),
            ((7, 10), True),
            ((7, 11), True),
            ((7, 12), True),
            ((7, 13), True),
            ((7, 14), True),
            ((7, 15), True),
            ((8, 5), True),
            ((8, 6), True),
            ((8, 7), True),
            ((8, 8), True),
            ((8, 9), True),
            ((8, 10), True),
            ((8, 11), True),
            ((8, 12), True),
            ((8, 13), True),
            ((8, 14), True),
            ((8, 15), True),
        ),
    },
}

workbook.Populate(Constants, Formatting)

# Pre-formulae user code

# Formula code

# Post-formulae user code
from random import random
def RandomColor():
    return Color.FromArgb(random()*256, random()*256, random()*256)

def FillWithSillyColors(sheet):
    for x in range(1, 10):
        for y in range(1, 6):
            cellLoc = (x, y)
            cell = sheet.Cells[cellLoc]
            cell.BackColor = RandomColor()
            cell.Value = cellLoc
            cell.Bold = True
            print cellLoc,
    print

sheet1 = workbook['The Silly Colors']
sheet2 = workbook['Silly Colors - No Grid']
FillWithSillyColors(sheet1)
FillWithSillyColors(sheet2)
