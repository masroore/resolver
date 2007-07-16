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
workbook = Workbook()

# Worksheet creation
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
Constants = {}

Formatting = {}

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
            sheet[cellLoc].BackColor = RandomColor()
            sheet[cellLoc].Value = cellLoc
            sheet[cellLoc].Bold = True
            print cellLoc,
    print

sheet1 = workbook['The Silly Colors']
sheet2 = workbook['Silly Colors - No Grid']
FillWithSillyColors(sheet1)
FillWithSillyColors(sheet2)
