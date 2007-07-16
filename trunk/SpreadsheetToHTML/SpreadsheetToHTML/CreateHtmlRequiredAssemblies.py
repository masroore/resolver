# Copyright (c) 2005-2007 Resolver Systems Ltd.
# All Rights Reserved
#

import clr

SYSTEM_REFERENCES = [
    "mscorlib, Version=2.0.0.0, Culture=neutral, PublicKeyToken=b77a5c561934e089",
    "System, Version=2.0.0.0, Culture=neutral, PublicKeyToken=b77a5c561934e089",
    "System.Data, Version=2.0.0.0, Culture=neutral, PublicKeyToken=b77a5c561934e089",
    "System.Drawing, Version=2.0.0.0, Culture=neutral, PublicKeyToken=b03f5f7f11d50a3a",
    "System.Security, Version=2.0.0.0, Culture=neutral, PublicKeyToken=b03f5f7f11d50a3a",
    "System.Windows.Forms, Version=2.0.0.0, Culture=neutral, PublicKeyToken=b77a5c561934e089",
    "System.Xml, Version=2.0.0.0, Culture=neutral, PublicKeyToken=b77a5c561934e089",
]

CUSTOM_REFERENCES = [
    "Syncfusion.Xlsio.Base",
]

for reference in SYSTEM_REFERENCES:
    clr.AddReferenceByName(reference)

from System.Diagnostics import Process
from System.IO import Path
from System.Reflection import Assembly
module = Process.GetCurrentProcess().MainModule
binDir = Path.GetDirectoryName(module.FileName)

for reference in CUSTOM_REFERENCES:
    assemblyPath = "%s\\%s.dll" % (binDir, reference)
    assembly = Assembly.LoadFile(assemblyPath)
    clr.AddReference(assembly)
