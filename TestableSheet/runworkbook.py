import clr
import sys

from System.IO import Directory, Path

x86path = r'C:\Program Files\Resolver One\bin'
x64path = r'C:\Program Files (x86)\Resolver One\bin'

if Directory.Exists(x86path):
    topLevel = Path.GetDirectoryName(x86path)
    bin = x86path
else:
    topLevel = Path.GetDirectoryName(x64path)
    bin = x64path

Directory.SetCurrentDirectory(topLevel)
sys.path.extend((topLevel, bin))

# This adds references to all .NET assemblies
# that Resolver One spreadsheets use
clr.AddReference('TopLevel')
import LoadAssemblies
LoadAssemblies.binDir = bin

import LoadRequiredAssemblies

from MarketData import Bloomberg, Thomson
from Library.Workbook import Workbook
from Library.ContextDependent import MakeContextDependentFunctions

_workbook = Workbook()

ContextDependentFunctions = MakeContextDependentFunctions(__file__, _workbook, sys)
RunWorkbook = ContextDependentFunctions.RunWorkbook

