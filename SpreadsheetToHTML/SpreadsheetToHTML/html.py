page = """
<html>
    <head>
        <title>%s</title>
        <style type="text/css">
            body { background-color: #fffff0; font-size: 10; }
            body, td, th { font-family: "Tahome", "Arial"; }
            .b_left   { border-left: solid 2px #000000; }
            .b_right  { border-right: solid 2px #000000; }
            .b_top    { border-top: solid 2px #000000; }
            .b_bottom { border-bottom: solid 2px #000000; }
            #spreadsheet { border-collapse: collapse; }
            td, th { padding:3px; }
            img.align-center {
                display: block;
                margin: auto; 
                text-align: center;
            }
        </style>
    </head>
    <body>
        <h1 align="center">%s</h1>
        <div align="center" id="sheetlink">%s</div>
        <div align="center" id="worksheet">%s</div>
    </body>
</html>
"""

worksheetMenuTable = """

<table border="1" cellpadding="5" width="300">
    <tr>
        <th colspan="%s"><strong>Worksheets</strong></th>
    </tr>
    <tr>%s</tr>
</table>
<br /><br />
"""

imageWorksheet = '''
<div align="center" class="align-center">
    <img alt="%s" class="align-center" src="%s" width="%s" height="%s" />
</div>
'''

worksheetMenuLink = '<td align="center"><a href="%s.html">%s</a></td>'

currentWorksheetMenuLink = '<td bgcolor="#ffffff" align="center"><strong><big><a href="%s.html">%s</a></big></strong></td>'

