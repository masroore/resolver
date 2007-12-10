# Copyright (c) 2005-2007 Resolver Systems Ltd.
# All Rights Reserved
#

from System import Uri
from System.Drawing import Color
from System.Drawing.Imaging import ImageFormat
from System.IO import Path
from System.Security import SecurityElement
from Utils.CellNameUtils import ColumnIndexToName

from SpreadsheetToHTML.html import (
    currentWorksheetMenuLink, imageWorksheet, 
    page, worksheetMenuTable, 
    worksheetMenuLink
)

import re

DATE_RE = re.compile(r"\d{1,2}/\d{1,2}/\d{1,4}")


def AddAlignment(value):
    align = ''
    try:
        if isinstance(value, (int, float)) or DATE_RE.match(str(value)):
            align = ' align="right" '
    except UnicodeDecodeError:
        pass
    return align


def HtmlFromWorkBook(workbook, name, outputDirectory):
    title = 'Resolver Spreadsheet: %s' % name    
    worksheets = []
    for sheet in workbook._worksheets:
        menu = HtmlWorksheetMenu(workbook, sheet.name)
        if hasattr(sheet, 'Image'):
            worksheetHtml = HtmlFromImageSheet(sheet, outputDirectory)
        else:
            worksheetHtml = HtmlFromSheet(sheet)
               
        sheetHtml = page % (title, title, menu, worksheetHtml)
        
        worksheets.append((sheet.name, sheetHtml))
    return worksheets


def HtmlFromImageSheet(sheet, outputDirectory):
    image = sheet.Image
    imageName = Uri.EscapeUriString(sheet.name) + '.jpg'
    fullPath = Path.Combine(outputDirectory, sheet.name + '.jpg')
    image.Save(fullPath, ImageFormat.Jpeg)
    return imageWorksheet % (imageName, imageName, image.Width, image.Height)


def HtmlWorksheetMenu(workbook, current):
    links = []
    for sheet in workbook._worksheets:
        if sheet.name == current:
            links.append(currentWorksheetMenuLink % (Uri.EscapeUriString(sheet.name), SecurityElement.Escape(sheet.name)))
        else:
            links.append(worksheetMenuLink % (Uri.EscapeUriString(sheet.name), SecurityElement.Escape(sheet.name)))
    return worksheetMenuTable % (len(links), ''.join(links))


def HtmlFromSheet(sheet):
    code = []
    if sheet.Error:
        code.append("<p>Error: %s</p>" % SecurityElement.Escape(str(sheet.Error)))

    if not sheet.HasMinMaxRowCol:
        code.append("<p>No data</p>")
    else:
        if sheet.ShowGrid:
            border = ' cellspacing="0" border-color="gray" border="1" '
        else:
            border = ' cellspacing="0" '
        code.append('<table id="spreadsheet" %s>' % border)
        for row in range(sheet.MaxRow + 1):
            if not sheet.ShowGrid and not row:
                continue
            code.append("<tr>")
            for col in range(sheet.MaxCol + 1):
                if not sheet.ShowGrid and not col:
                    continue
                
                code.append(TableSnippetFromLocation(sheet, col, row))
            code.append("</tr>")
        code.append("</table>")

    return "\r\n".join(code)


def TableSnippetFromLocation(sheet, col, row):
    if 0 in (col, row):
        return ThFromHeader(col, row)
    return TdFromCell(sheet, col, row)


def ThFromHeader(col, row):
    if col:
        return "<th scope='col'>%s</th>" % SecurityElement.Escape(ColumnIndexToName(col))
    if row:
        return "<th scope='row'>%s</th>" % SecurityElement.Escape(str(row))
    return "<td>&nbsp;</td>"


def TdFromCell(sheet, col, row):
    h = ""
    w = ""
    if col == 1:
        height = sheet.Rows[row].Height
        if height == -1:
            height = 22
        h = ' height="%s" ' % height
    if row == 1:
        width = sheet.Cols[col].Width
        if width == -1:
            width = 70
        w = ' width="%s" ' % width
    
    cell = sheet.Cells[col, row]
    style = h + w
    if cell.BackColor != Color.White:
        style += ' bgcolor="#%02x%02x%02x" ' % (cell.BackColor.R, cell.BackColor.G, cell.BackColor.B) 
    bold = "%s"
    if cell.Bold:
        bold = "<strong>%s</strong>"
    if cell.Value is None or cell.FormattedValue == "":
        contents = bold % "&nbsp"
    elif hasattr(cell.Value, '__class__') and cell.Value.__class__.__name__ == 'Button':
        contents = bold % "&nbsp"
    elif hasattr(cell.Value, 'inputField'):
        name = '%s-%s' % (col, row)
        contents = bold % (inputField % (name, cell.Value.value)) 
    else:
        contents = bold % SecurityElement.Escape(cell.FormattedValue)
    
    classes = []
    if cell.BorderTop:
        classes.append('b_top')
    if cell.BorderBottom:
        classes.append('b_bottom')
    if cell.BorderLeft:
        classes.append('b_left')
    if cell.BorderRight:
        classes.append('b_right')
    if classes:
        style += ' class="%s" ' % ' '.join(classes)

    style = AddAlignment(cell.Value) + style
    error = ""
    if cell.Error:
        error = "<br />Error: %s" % SecurityElement.Escape(str(cell.Error))
    return "<td%s>%s%s</td>" % (style, contents, error)
