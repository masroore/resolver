# This adds references to all the assemblies we may use
import SpreadsheetToHTML.CreateHtmlRequiredAssemblies

import sys

from System.IO import Directory, Path
from System.Drawing import Point
from System.Windows.Forms import (
    Application, Button, Label,
    DialogResult, FolderBrowserDialog , 
    Form, FormBorderStyle,MessageBox, 
    MessageBoxButtons, MessageBoxIcon, 
    OpenFileDialog
)
# Do this import now to speed up loading spreadsheets
from Library.Workbook import Workbook

from SpreadsheetToHTML.Htmliser import HtmlFromWorkBook

import traceback

class NoWorkbook(Exception):
    pass

def LoadSpreadsheet(path):
    """
    Load and execute a spreadsheet exported as Python.
     
    This function takes a path to the Python file and
    returns a workbook instance.
    """
    import sys
    spreadsheetDirectory = Path.GetDirectoryName(path)
    sys.path.append(spreadsheetDirectory)
    
    try:
        context = {'__name__': '__main__', '__file__': path, '__builtins__': __builtins__}
        h = open(path)
        code = h.read() + '\n'
        h.close()
        exec code in context
        Spreadsheet = context.get('Spreadsheet')
        if Spreadsheet is not None:
            # Spreadsheet was exported as a class
            workbook = Spreadsheet()
            workbook.recalculate()
        else:
            # Spreadsheet exported as non-class
            workbook = context.get('workbook')
            if workbook is None:
                # Not a spreadsheet at all
                raise NoWorkbook("No workbook defined in this code. Are you sure it is a Resolver spreadsheet?")
        
        return workbook
    finally:
        sys.path.remove(spreadsheetDirectory)


# Make sure the correct directory is in the path - or our imports will fail
mainDirectory = Directory.GetParent(Path.GetDirectoryName(Application.ExecutablePath))
if not mainDirectory in sys.path:
    sys.path.insert(0, mainDirectory)


class CreateHTMLForm(Form):
    """
    A form to create HTML files (web pages) from Resolver spreadsheets exported as code.
    
    You first choose the file to load and then the directory to put the output files in.
    """
    def __init__(self):
        self.Text = "Spreadsheet to HTML"
        self.Width = 220
        self.Height = 90
        self.FormBorderStyle = FormBorderStyle.Fixed3D
        
        self.openDialog = OpenFileDialog()
        self.openDialog.Title = "Choose a Python Spreadsheet File"
        self.openDialog.Filter = 'Python files (*.py)|*.py|All files (*.*)|*.*'
        
        self.folderDialog = FolderBrowserDialog()
        self.folderDialog.ShowNewFolderButton = True
        self.folderDialog.Description = "Choose a directory to save the HTML files in."
        
        l = Label()
        l.AutoSize = True
        l.Location = Point(50, 5)
        l.Text = "Choose a Spreadsheet"
        
        b = Button(Text="Choose")
        b.Click += self.convert
        b.Location = Point(70, 30)

        self.Controls.Add(l)
        self.Controls.Add(b)
        
        
    def convert(self, sender, event):
        # Present the file dialog to choose a file
        if self.openDialog.ShowDialog() != DialogResult.OK:
            return
        
        filePath = self.openDialog.FileName
        try:
            # Load the spreadsheet
            print 'Loading spreadsheet'
            workbook = LoadSpreadsheet(filePath)
        except Exception, e:
            MessageBox.Show("An error has occurred:\r\n%s: %s" % (e.__class__.__name__, e),
                            "An Error has Occurred", 
                            MessageBoxButtons.OK, MessageBoxIcon.Warning)
            
            # Print the full traceback to the console for debugging
            traceback.print_exc()
            return
        
        # Choose a folder to place output files
        if self.folderDialog.ShowDialog() != DialogResult.OK:
            return
            
        outputDirectory = self.folderDialog.SelectedPath
        spreadsheetName = Path.GetFileNameWithoutExtension(filePath)
        
        # Create html as strings in a dictionary keyed by worksheet name
        # Will save images from ImageWorksheets
        print 'Creating HTML'
        worksheetsAsHtml = HtmlFromWorkBook(workbook, spreadsheetName, outputDirectory)
        
        # Save the output html files
        for name, text in worksheetsAsHtml:
            print 'Writing %s' % (name + '.html')
            try:
                path = Path.Combine(outputDirectory, name + '.html')
                h = open(path, 'w')
                h.write(text)
                h.close()
            except Exception, e:
                MessageBox.Show("An error has occurred:\r\n%s: %s" % (e.__class__.__name__, e),
                                "An Error has Occurred", 
                                MessageBoxButtons.OK, MessageBoxIcon.Warning)
                return

        MessageBox.Show("Done!", "Done", MessageBoxButtons.OK, MessageBoxIcon.Information)


form = CreateHTMLForm()
Application.Run(form)

