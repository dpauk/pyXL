#---File header
#-------------------------------------------------------------------------------
# Name:        pyXL.py (Donationcoder Assignment 8)
# Purpose:     A spreadsheet application written in python and wxpython.
#
# Author:      David Albone (mnemonic)
#
# Created:     27/12/2008 - 06/01/2009
# Icons:       http://tango.freedesktop.org/Tango_Desktop_Project
#-------------------------------------------------------------------------------
# Change List
# -----------
# v1 (27th December 2008):
#   Completed new, load and save functions.
#   Completed update of content bar from grid (and vice-versa).
# v2 (28th December 2008):
#   Added csv import and export.
#   Added about box.
# v3 (3rd January 2009):
#   Split data into a table (DataTable class)
#   Added current field box.
#   Added formula functionality.
# v4 (3rd January 2009):
#   Added printing support.
# v5 (4th January 2009):
#   Added support for tab-separated file import.
#   Added support for space-separated file import.
#   Added support for semicolon-separated file import.
#   Added help text link.
# v6 (6th January 2009):
#   Re-wrote save function to allow "save as..." to be added.
# v7 (6th January 2009):
#   Added support for errored loads.
#   Removed debugging code.
#   Updated csv export to use csv.writer.
#-------------------------------------------------------------------------------

#!/usr/bin/env python

import wx
import wx.grid
import os
import sys
import csv
import re
import sqlite3 as sqlite
from wx.html import HtmlEasyPrinting

NUMBER_GRID_ROWS = 256
NUMBER_GRID_COLS = 256

#---Model objects

class SpreadsheetDatabase(object):
    """Database containing the data from a spreadsheet"""
    def __init__(self, databaseName = ''):
        self.databaseName = databaseName

    def createDatabase(self):
        """Creates a new databsse i.e. save a new spreadsheet"""
        self.__openDatabase()
        self.__createTables()
        self.__closeDatabase()
        
    def __openDatabase(self):
        """Opens a database and creates a cursor"""
        self.con = sqlite.connect(self.databaseName)
        self.cursor = self.con.cursor()

    def __createTables(self):
        """Creates the database tables required for the spreadsheet"""
        # spreadsheet_data (the data held in all of the spreadsheet fields):
        #   row_id = field row
        #   column_id = field column
        #   field_type = reference to spreadsheet_field_type.type
        #   value = value stored in field
        self.cursor.execute("CREATE TABLE spreadsheet_data (row_id INTEGER, column_id INTEGER, value VARCHAR(256))")

    def __databaseCommit(self):
        """Commits inserted / updated data into database"""
        self.con.commit()

    def __closeDatabase(self):
        """Closes the databse"""
        self.con.close()

    def loadDatabase(self):
        """Loads a database i.e. load a spreadsheet"""
        self.__openDatabase()
        try:
            self.cursor.execute("SELECT * from spreadsheet_data")
        except:
            # This executes when the database isn't in the correct format or isn't even a databases
            raise Exception("Load error.")
            self.__closeDatabase()
        allData = self.cursor.fetchall()
        self.__closeDatabase()
        return allData

    def saveDatabase(self, dataList):
        """Saves an existing database i.e. save existing spreadsheet"""
        self.__openDatabase()
        for nextRow in dataList:
            self.__insertRow(nextRow[0], nextRow[1], nextRow[2], nextRow[3])
        self.__databaseCommit()
        self.__closeDatabase()

    def __insertRow(self, row, col, type, value):
        """Inserts a row into spreadsheet_data"""
        self.cursor.execute("INSERT INTO spreadsheet_data VALUES (?, ?, ?)", (row, col, value))

#---wxPython objects (view)

class DataTable(wx.grid.PyGridTableBase):
    """Holds the data displayed in the grid"""
    def __init__(self):
        wx.grid.PyGridTableBase.__init__(self)
        self.data = {}
        self.dataType = wx.grid.GRID_VALUE_STRING
        self.formulas = {}
        self.loadedFile = ''
    
    def IsEmptyCell(self, row, col):
        """Returns a cells state"""
        return self.data.get((row,col)) is not None
    
    def GetNumberRows(self):
        """Returns the number of rows in the table"""
        return NUMBER_GRID_ROWS
    
    def GetNumberCols(self):
        """Returns the number of cols in the table"""
        return NUMBER_GRID_COLS
    
    def GetValue(self, row, col):
        """Gets the value held in a specified cell"""
        value = self.data.get((row, col))
        if value is not None:
            return value
        else:
            return ''
        
    def SetValue(self, row, col, value):
        """Sets the value held in a specified cell"""
        # See if value is a formula
        if len(value) == 0: # i.e. cell has been deleted
            del self.data[(row, col)]
            try:
                del self.formulas[(row,col)]
            except KeyError:
                return
            return
        if (value[0]) == "=":
            splitFormula = self.__breakdownFormula(value)
            if (self.__isFormulaValid(splitFormula)):
                self.data[(row, col)]  = self.__calculateFormula(row, col, splitFormula)
                self.formulas[(row,col)] = value
            else:
                self.data[(row, col)] = "!ERR %s" % value
        else:
            self.data[(row, col)] = value
    
    def reInitialise(self):
        """Re-initialises the grid"""
        self.data = {}
        self.formulas = {}
        self.loadedFile = {}
    
    def getFormula(self, row, col):
        """Returns the value of a formula if available"""
        try:
            return self.formulas[(row, col)] or True
        except (KeyError):
            return None
       
    def GetTypeName(self, row, col):
        """Returns the datatype of the cell"""
        return self.dataType
    
    def isFloat(self, row, col):
        """Returns the result of whether a cell contains a float value"""
        try:
            return float(self.data.get((row, col))) or True
        except (ValueError, TypeError), e:
            return False
        
    def isInt(self, row, col):
        """Returns the result of whether a cell contains an integer value"""
        try:
            return int(self.data.get((row, col))) or True
        except:
            return False
    
    def __breakdownFormula(self, formula):
        """Breaks down a passed formula into two operands and an operator"""
        return re.findall(r"[\w']+|[+-/\*]", formula)
    
    def __isFormulaValid(self, splitFormula):
        """Checks to see if the entered formula is valid"""
        operands, operators = self.__splitIntoOperandsAndOperators(splitFormula)
        if not(self.__checkNumberOfSplits):
            return False
        if not(self.__checkOperands(operands)):
            return False
        if not(self.__checkOperators(operators)):
            return False
        if not(self.__checkOperandGridValuesValid(operands)):
            return False
        return True
    
    def __splitIntoOperandsAndOperators(self, splitFormula):
        """Splits a list into operands and operators"""
        operands = []
        operators = []
        listPosition = 0
        for value in splitFormula:
            if listPosition % 2 == 0:
                operands.append(value)
            else:
                operators.append(value)
            listPosition += 1
        return operands, operators
    
    def __checkNumberOfSplits(self, operands, operators):
        """Ensures that there are one less operators than operands e.g. 2+2 = 2 operands and 1 operator"""
        if (len(operands) - 1) == len(operators):
            return True
        else:
            return False
    
    def __checkOperands(self, operands):
        """Checks the formula operands"""
        for value in operands:
            # Split operand into letters and numbers
            splitValue = re.findall(r"[A-Z]+|[0-9]+", value)
            if len(splitValue) == 2:
                return True
            else:
                return False
    
    def __checkOperators(self, operators):
        """Checks the formula operators"""
        for value in operators:
            splitValue = re.findall(r"[+-/\*]", value)
            if len(splitValue) == 1:
                return True
            else:
                return False
    
    def __checkOperandGridValuesValid(self, operands):
        """Checks that the values contained in the cellReference are valid"""
        for operand in operands:
            row, col = self.__convertCellReferenceIntoRowAndCol(operand)
            if (self.isInt(row, col) or self.isFloat(row, col)):
                return True
            else:
                return False
    
    def refreshFormulas(self):
        """Refreshes all the formulas in the data"""
        for cell, value in self.formulas.iteritems():
            splitFormula = self.__breakdownFormula(value)
            if (self.__isFormulaValid(splitFormula)):
                self.data[(cell[0], cell[1])] = self.__calculateFormula(cell[0], cell[1], splitFormula)
            else:
                self.data[cell] = "!ERR %s" % value

    def __calculateFormula(self, row, col, splitFormula):
        """Calculates the result of a formula and updates the self.data attribute"""
        operatorPosition = 0
        operandPosition = 0
        runningTotal = 0
        operands, operators = self.__splitIntoOperandsAndOperators(splitFormula)
        row, col = self.__convertCellReferenceIntoRowAndCol(operands[operandPosition])
        operandPosition += 1
        operandValue = self.data.get((row, col))
        runningTotal = self.__numberType(operandValue)
        for allOperands in range(len(operands) - 1):
            row, col = self.__convertCellReferenceIntoRowAndCol(operands[operandPosition])
            operandValue = self.GetValue(row, col)
            operandTwo = operands[(operandPosition)]
            if operators[operatorPosition] == "+":
                runningTotal += self.__numberType(operandValue)
            elif operators[operatorPosition] == "-":
                runningTotal -= self.__numberType(operandValue)
            elif operators[operatorPosition] == "*":
                runningTotal *= self.__numberType(operandValue)
            elif operators[operatorPosition] == "/":
                runningTotal /= self.__numberType(operandValue)
            operatorPosition += 1
            operandPosition += 1
        return runningTotal
    
    # Pre-condition: stringNumber must be a number in string format
    def __numberType(self, stringNumber):
        """Converts a string in to either a float or an int"""
        if self.isStringInt(stringNumber):
            return int(stringNumber)
        elif self.isStringFloat(stringNumber):
            return float(stringNumber)
        else:
            print "Error in DataTable.__numberType - input parameter not a number in string form"
            print stringNumber
            print type(stringNumber)
            sys.exit()
    
    def isStringFloat(self, stringNumber):
        """Returns the result of whether a string contains a float value"""
        try:
            return float(stringNumber) or True
        except (ValueError, TypeError), e:
            return False
        
    def isStringInt(self, stringNumber):
        """Returns the result of whether a cell contains an integer value"""
        try:
            return int(stringNumber) or True
        except (ValueError, TypeError), e:
            return False
    
    def __convertCellReferenceIntoRowAndCol(self, cellReference):
        """Converts a cell reference (e.g. A1) into row and col"""
        splitReference = re.findall(r"[A-Z]+|[0-9]+", cellReference)
        col = self.__convertLetterToCol(splitReference[0])
        row = int(splitReference[1]) - 1
        return row, col

    def __convertLetterToCol(self, letter):
        """Converts a column letter to a col number"""
        splitLetter = re.findall(r"[A-Z]", letter)
        if len(splitLetter) == 1:
            numberLetter = ord(letter)
            return numberLetter - 65
        else:
            total = 0
            multiplier = ord(splitLetter[0]) - 65
            total = multiplier * 26
            secondLetterValue = ord(splitLetter[1])
            total += total + (secondLetterValue - 1)
            return total

class SpreadsheetPrinter(HtmlEasyPrinting):
    def __init__(self):
        HtmlEasyPrinting.__init__(self)

    def Print(self, text, doc_name):
        self.SetHeader(doc_name)
        self.PrintText(text,doc_name)

    def PreviewText(self, text, doc_name):
        self.SetHeader(doc_name)
        HtmlEasyPrinting.PreviewText(self, text)

class MainFrame(wx.Frame):
    """Main frame of the spreadsheet"""
    def __init__(self, parent, id, title):
        self.loadedDatabase = ''
        wx.Frame.__init__(self, parent, id, title, size=(934,619), style = wx.DEFAULT_FRAME_STYLE | wx.TAB_TRAVERSAL)

        # Setup frame
        self.__createMenu()
        self.__createToolbar()
        self.__startLayout()
        self.__createContentBar()
        self.__createGrid()
        self.__completeLayout()
        self.__createStatusBar()
        self.__setupEventHandlers()
        self.__setupDataModel()

    def __createMenu(self):
        """Creates the main page menu"""
        # Setup layout, menubar and toolbar
        self.__createFileMenu()
        self.__createHelpMenu()
        self.__completeMenuBarSetup()
        
    def __createFileMenu(self):
        """Creates the main page file menu"""
        self.mainFileMenu = wx.Menu()
        self.newSheet = self.mainFileMenu.Append(-1, "&New", "Clear the current sheet")
        self.openSheet = self.mainFileMenu.Append(-1, "&Open", "Open an existing sheet")
        self.mainFileMenu.AppendSeparator()
        self.saveSheet = self.mainFileMenu.Append(-1, "&Save", "Save the current sheet")
        self.saveAsSheet = self.mainFileMenu.Append(-1, "&Save as...", "Save as the current sheet")
        self.mainFileMenu.AppendSeparator()
        self.importMenu = wx.Menu()
        self.importCsv = self.importMenu.Append(-1, "&CSV", "Import from a CSV file")
        self.importSpace = self.importMenu.Append(-1, "&Space-separated", "Import from a space-separated file")
        self.importTab = self.importMenu.Append(-1, "&Tab-separated", "Import from a tab-separated file")
        self.importSemicolon = self.importMenu.Append(-1, "&Semicolon-separated", "Import from a semicolon-separated file")
        self.mainFileMenu.AppendMenu(-1, "&Import file", self.importMenu)
        self.exportCsv = self.mainFileMenu.Append(-1, "&Export to CSV", "Export to a CSV file")
        self.mainFileMenu.AppendSeparator()
        self.printMenu = self.mainFileMenu.Append(-1, "&Print", "Prints sheet")
        self.printPreviewMenu = self.mainFileMenu.Append(-1, "P&rint preview", "Print preview")
        self.mainFileMenu.AppendSeparator()
        self.exitProg = self.mainFileMenu.Append(-1, "E&xit", "Exit")
        
    def __createHelpMenu(self):
        """Creates the main page help menu"""
        self.mainHelpMenu = wx.Menu()
        self.helpApp = self.mainHelpMenu.Append(-1, "&Help", "Help on pyXL")
        self.mainHelpMenu.AppendSeparator()
        self.aboutApp = self.mainHelpMenu.Append(-1, "&About", "Information about pyXL")
        
    def __completeMenuBarSetup(self):
        """Completes the setup of the main menu bar"""
        self.mainMenuBar = wx.MenuBar(0)
        self.mainMenuBar.Append(self.mainFileMenu, "&File")
        self.mainMenuBar.Append(self.mainHelpMenu, "&Help")
        self.SetMenuBar(self.mainMenuBar)

    def __createToolbar(self):
        """Creates the main page toolbar"""
        self.mainToolBar = self.CreateToolBar(wx.TB_HORIZONTAL, wx.ID_ANY)
        self.__addToolbarButton()
        self.mainToolBar.Realize()

    def __toolbarButtonData(self):
        """The data for the toolbar buttons"""
        return ((wx.ID_NEW, "icons/document-new.png"),
                ("separator", ""),
                (wx.ID_OPEN, "icons/document-open.png"),
                (wx.ID_SAVE, "icons/document-save.png"),
                (wx.ID_SAVEAS, "icons/document-save-as.png"),
                ("separator", ""),
                (wx.ID_PRINT, "icons/document-print.png"))

    def __addToolbarButton(self):
        """Adds buttons to the toolbar"""
        for eachID, eachPath in self.__toolbarButtonData():
            if (eachID == "separator"):
                self.mainToolBar.AddSeparator()
            else:
                self.mainToolBar.AddLabelTool(eachID, "", wx.Bitmap(eachPath))
    
    def __startLayout(self):
        """Starts the layout of the main page"""
        self.mainSizer = wx.BoxSizer(wx.VERTICAL)
        self.mainPanel = wx.Panel(self, wx.ID_ANY, wx.DefaultPosition, wx.DefaultSize, wx.TAB_TRAVERSAL)
        self.displaySizer = wx.BoxSizer(wx.VERTICAL)

    def __createContentBar(self):
        """Creates a bar in the main page that displays the content of the current cell"""
        self.fieldContent = wx.BoxSizer(wx.HORIZONTAL)
        self.currentFieldLabel = wx.StaticText(self.mainPanel, wx.ID_ANY, "Current cell:", wx.DefaultPosition, wx.DefaultSize, 0)
        self.currentFieldLabel.Wrap(-1)
        self.fieldContent.Add(self.currentFieldLabel, 0, wx.ALL, 5);
        self.currentFieldText = wx.TextCtrl(self.mainPanel, wx.ID_ANY, wx.EmptyString, wx.DefaultPosition, wx.DefaultSize, wx.TE_PROCESS_ENTER)
        self.fieldContent.Add(self.currentFieldText, 0, wx.ALL | wx.EXPAND, 0)
        self.fieldContentLabel = wx.StaticText(self.mainPanel, wx.ID_ANY, "Cell content:", wx.DefaultPosition, wx.DefaultSize, 0)
        self.fieldContentLabel.Wrap(-1)
        self.fieldContent.Add(self.fieldContentLabel, 0, wx.ALL, 5);
        self.fieldContentText = wx.TextCtrl(self.mainPanel, wx.ID_ANY, wx.EmptyString, wx.DefaultPosition, wx.DefaultSize, wx.TE_PROCESS_ENTER)
        self.fieldContent.Add(self.fieldContentText, 10, wx.ALL | wx.EXPAND, 0)
        self.displaySizer.Add(self.fieldContent, 1, wx.EXPAND, 5)

    def __createGrid(self):
        """Creates the main spreadsheet grid"""
        self.mainGrid = wx.grid.Grid(self.mainPanel, wx.ID_ANY, wx.DefaultPosition, wx.DefaultSize, 0)

        # Grid
        self.mainGrid.CreateGrid(NUMBER_GRID_ROWS, NUMBER_GRID_COLS)
        self.mainGrid.EnableEditing(True)
        self.mainGrid.EnableGridLines(True)
        self.mainGrid.EnableDragGridSize(False)
        self.mainGrid.SetMargins(0, 0)

        # Columns
        self.mainGrid.EnableDragColMove(False)
        self.mainGrid.EnableDragColSize(True)
        self.mainGrid.SetColLabelSize(30)
        self.mainGrid.SetColLabelAlignment(wx.ALIGN_CENTRE, wx.ALIGN_CENTRE)

        # Rows
        self.mainGrid.EnableDragRowSize(True)
        self.mainGrid.SetRowLabelSize(80)
        self.mainGrid.SetRowLabelAlignment(wx.ALIGN_CENTRE, wx.ALIGN_CENTRE)
        
        # Cell Defaults
        self.mainGrid.SetDefaultCellAlignment(wx.ALIGN_LEFT, wx.ALIGN_TOP)
        self.displaySizer.Add(self.mainGrid, 20, wx.ALL | wx.EXPAND, 0)
        
    def __completeLayout(self):
        """Completes the layout of the main page"""
        self.mainPanel.SetSizer(self.displaySizer)
        self.mainPanel.Layout()
        self.displaySizer.Fit(self.mainPanel)
        self.mainSizer.Add(self.mainPanel, 1, wx.EXPAND | wx.ALL, 0)
        self.SetSizer(self.mainSizer)
        self.Layout()

    def __createStatusBar(self):
        """Creates the status bar on the main page"""
        self.mainStatusBar = self.CreateStatusBar(1, wx.ST_SIZEGRIP, wx.ID_ANY)

    def __setupEventHandlers(self):
        """Setups the event handlers for the main page"""
        self.__setupMenuEvents()
        self.__setupToolbarEvents()
        self.__setupGridEvents()
        self.__setupContentBarEvents()

    def __setupMenuEvents(self):
        """Sets up the menu event handlers"""
        self.Bind(wx.EVT_MENU, self.__OnNew, self.newSheet)
        self.Bind(wx.EVT_MENU, self.__OnOpen, self.openSheet)
        self.Bind(wx.EVT_MENU, self.__OnSave, self.saveSheet)
        self.Bind(wx.EVT_MENU, self.__onSaveAs, self.saveAsSheet)
        self.Bind(wx.EVT_MENU, self.__importCsv, self.importCsv)
        self.Bind(wx.EVT_MENU, self.__importSpace, self.importSpace)
        self.Bind(wx.EVT_MENU, self.__importTab, self.importTab)
        self.Bind(wx.EVT_MENU, self.__importSemicolon, self.importSemicolon)
        self.Bind(wx.EVT_MENU, self.__exportCsv, self.exportCsv)
        self.Bind(wx.EVT_MENU, self.__onPrint, self.printMenu)
        self.Bind(wx.EVT_MENU, self.__onPrintPreview, self.printPreviewMenu)
        self.Bind(wx.EVT_MENU, self.__OnExit, self.exitProg)
        self.Bind(wx.EVT_MENU, self.__onHelp, self.helpApp)
        self.Bind(wx.EVT_MENU, self.__onAbout, self.aboutApp)

    def __setupToolbarEvents(self):
        """Sets up the toolbar event handlers"""
        self.Bind(wx.EVT_TOOL, self.__OnNew, id=wx.ID_NEW)
        self.Bind(wx.EVT_TOOL, self.__OnOpen, id=wx.ID_OPEN)
        self.Bind(wx.EVT_TOOL, self.__OnSave, id=wx.ID_SAVE)
        self.Bind(wx.EVT_TOOL, self.__onSaveAs, id=wx.ID_SAVEAS)
        self.Bind(wx.EVT_TOOL, self.__onPrint, id=wx.ID_PRINT)

    def __setupGridEvents(self):
        """Sets up the menu event handlers"""
        # Grid events
        self.Bind(wx.grid.EVT_GRID_CELL_LEFT_CLICK, self.__updateContentBarWithCellValue)
        self.Bind(wx.grid.EVT_GRID_SELECT_CELL, self.__updateContentBarWithCellValue)
        self.Bind(wx.grid.EVT_GRID_CELL_CHANGE, self.__updateContentBarWithCellValue)
        
    def __setupContentBarEvents(self):
        """Sets up the grid event handlers"""
        self.fieldContentText.Bind(wx.EVT_TEXT_ENTER, self.__enterContentBar)

    def __OnSave(self, event):
        """Deals with the user saving"""
        if self.spreadsheetData.loadedFile == '':
            self.__onSaveAs()
        else:
            try:
                os.remove(self.spreadsheetData.loadedFile)
            except:
                pass
            self.__createSaveFile(self.spreadsheetData.loadedFile)
            dialogText = "New file %s created." % (self.spreadsheetData.loadedFile)
            confirmNewDialog = wx.MessageDialog(None, dialogText, 'New file created', wx.OK)
            confirmNewDialog.ShowModal()
        
    def __saveFile(self):
        """Prompts ths user for a save name and saves the file"""
        saveFilters = 'pyXL files (*.pyx)|*.pyx'
        saveDialog = wx.FileDialog(None, message = "Save spreadsheet file", wildcard = saveFilters, style = wx.SAVE)
        if (saveDialog.ShowModal() == wx.ID_OK):
            if self.__checkIfFileOverwrite(saveDialog.GetPath()):
                try:
                    os.remove(saveDialog.GetPath())
                except:
                    pass
                self.__createSaveFile(saveDialog.GetPath())
    
    def __createSaveFile(self, filePath):
        """Creates the save file"""
        self.saveFile = SpreadsheetDatabase(filePath)
        self.saveFile.createDatabase()
        self.saveFile.saveDatabase(self.__getPopulatedCells())
    
    def __checkIfFileOverwrite(self, saveFilePath):
        """Check to see if the user is trying to overwrite the file and prompt them if they are"""
        if os.path.exists(saveFilePath):
            dialogText = "Are you sure to overwrite %s?" % (saveFilePath)
            overwriteDialog = wx.MessageDialog(None, dialogText, 'Overwrite file?', wx.YES_NO | wx.NO_DEFAULT | wx.ICON_QUESTION)
            if overwriteDialog.ShowModal() == 5103:
                return True
            else:
                return False
        else:
            return True

    def __onSaveAs(self, event=''):
        """Deals with the user saving as..."""
        self.__saveFile()
    
    def __getPopulatedCells(self):
        """Puts all the populated cells into a list of lists [[row, col, type, value], [row, col, type, value], ...]"""
        populatedCellList = []
        for col in range(NUMBER_GRID_COLS):
            for row in range(NUMBER_GRID_ROWS):
                if (self.mainGrid.GetCellValue(row, col) != ''):
                    currentCellList = [row, col, 1, self.mainGrid.GetCellValue(row, col)]
                    populatedCellList.append(currentCellList)
        return populatedCellList

    def __OnOpen(self, event):
        """Deals with the user loading"""
        openDialogResult = self.__promptForLoadFile()
        if (openDialogResult.ShowModal() == wx.ID_OK):
            self.__openFile(openDialogResult.GetPath())
            self.spreadsheetData.loadedFile = openDialogResult.GetPath()

    def __promptForLoadFile(self):
        """Prompts the user for a file to load"""
        openFilters = 'pyXL files (*.pyx)|*.pyx'
        openDialog = wx.FileDialog(None, message = "Open spreadsheet file", wildcard = openFilters, style = wx.OPEN)
        return openDialog

    def __openFile(self, filePath):
        """Loads a file"""
        openFile = SpreadsheetDatabase(filePath)
        try:
            openCellList = openFile.loadDatabase()
        except:
            errorDialog = wx.MessageDialog(None, 'Bad file - loading not completed', 'ERROR', wx.ICON_ERROR | wx.OK)
            errorDialog.ShowModal()
            return
        self.__populateLoadedDataIntoCells(openCellList)

    def __populateLoadedDataIntoCells(self, cellList):
        """Adds all loaded data into the spreadsheet"""
        for cellData in cellList:
            self.mainGrid.SetCellValue(cellData[0], cellData[1], cellData[2])

    def __OnNew(self, event):
        """Clears the spreadsheet"""
        newDialogResult = self.__promptIsUserSure()
        if (newDialogResult == 5103):
            self.mainGrid.ClearGrid()
            self.spreadsheetData.reInitialise()
            self.fieldContentText.Clear()

    def __promptIsUserSure(self):
        """Sees if the user really wants to start a new spreadsheet"""
        newDialog = wx.MessageDialog(None, 'Are you sure?', 'New file', wx.YES_NO | wx.NO_DEFAULT | wx.ICON_QUESTION)
        newDialogResult = newDialog.ShowModal()
        return newDialogResult

    def __importCsv(self, event):
        """Imports a CSV file"""
        self.__importFile(",", "csv files (*.csv)|*.csv", "Import CSV file")
    
    def __importSpace(self, event):
        """Imports a space-separated file"""
        self.__importFile(" ", "txt files (*.txt)|*.txt", "Import space-separated file")
    
    def __importTab(self, event):
        """Imports a tab-separated file"""
        self.__importFile("\t", "txt files (*.txt)|*.txt", "Import tab-separated file")
            
    def __importSemicolon(self, event):
        """Imports a semi-colon separated file"""
        self.__importFile(";", "txt files (*.txt)|*.txt", "Import semicolon-separated file")
    
    def __importFile(self, separator, openFilters, dialogMessage):
        """Imports a file"""
        openDialogResult = self.__promptForImportFile(openFilters, dialogMessage)
        if (openDialogResult.ShowModal() == wx.ID_OK):
            self.mainGrid.ClearGrid()
            self.__openSeparatedFile(openDialogResult.GetPath(), separator)
    
    def __promptForImportFile(self, openFilters, dialogMessage):
        """Prompts the user for a Csv file to load"""
        openDialog = wx.FileDialog(None, message = dialogMessage, wildcard = openFilters, style = wx.OPEN)
        return openDialog

    def __openSeparatedFile(self, filePath, separator):
        """Loads a separated file"""
        csvFile = csv.reader(open(filePath), delimiter=separator)
        rowNum = 0
        for line in csvFile:
            colNum = 0
            for value in line:
                self.mainGrid.SetCellValue(rowNum, colNum, value.strip(" "))
                colNum += 1         
            rowNum += 1

    def __exportCsv(self, event):
        """Exports a CSV file"""
        exportDialogResult = self.__promptForExportCsvFile()
        if (exportDialogResult.ShowModal() == wx.ID_OK):
            self.__exportCsvFile(exportDialogResult.GetPath())

    def __promptForExportCsvFile(self):
        """Prompts the user for a Csv file to load"""
        exportFilters = 'csv files (*.csv)|*.csv'
        exportDialog = wx.FileDialog(None, message = "Export csv file", wildcard = exportFilters, style = wx.SAVE)
        return exportDialog

    def __exportCsvFile(self, filePath):
        """Exports a csv file"""
        finalPopulatedRow = self.__findFinalPopulatedRow()
        csvWriter = open(filePath, 'w')
        for row in range(finalPopulatedRow + 1):
            finalPopulatedColumn = self.__findFinalPopulatedColumnForRow(row)
            outString = ''
            for col in range(finalPopulatedColumn + 1):
                if col == 0:
                    outString += "%s" % self.mainGrid.GetCellValue(row, col)   
                else:
                    outString += ", %s" % self.mainGrid.GetCellValue(row, col)     
            outString += "\n"
            csvWriter.write(outString)
        csvWriter.close()

    def __findFinalPopulatedRow(self):
        """Finds the final populated column in a given row"""
        finalRowNum = 0
        for row in range(NUMBER_GRID_ROWS):
            for col in range(NUMBER_GRID_ROWS):
                if (self.mainGrid.GetCellValue(row, col) != ''):
                    finalRowNum = row
                    break
        return finalRowNum

    def __findFinalPopulatedColumnForRow(self, row):
        """Finds the final populated column in a given row"""
        finalColNum = 0
        for col in range(NUMBER_GRID_COLS):
            if (self.mainGrid.GetCellValue(row, col) != ''):
                finalColNum = col
        return finalColNum
    
    def __onHelp(self, event):
        """Launch help text"""
        os.startfile("pyXL_help.txt")
    
    def __onAbout(self, event):
        """About box"""
        wx.MessageBox("pyXL (Donationcoder assignment 8)\nby David Albone (mnemonic)\nJanuary 2009", "About")  
    
    def __onPrintPreview(self, event):
        """Displays a print preview of the current spreadsheet"""
        printText = self.__formatGridForPrinting(self.mainGrid)
        self.spreadsheetPrint = SpreadsheetPrinter()
        self.spreadsheetPrint.PreviewText(printText, "pyXL sheet")
    
    def __onPrint(self, event):
        """Prints current spreadsheet"""
        printText = self.__formatGridForPrinting(self.mainGrid)
        self.spreadsheetPrint = SpreadsheetPrinter()
        self.spreadsheetPrint.Print(printText, "pyXL sheet")
    
    def __formatGridForPrinting(self, grid):
        """Formats a wxpython grid into a text format ready for printing"""
        finalPopulatedRow = self.__findFinalPopulatedRow()
        finalPopulatedCol = self.__findFinalPopulatedCol()
        printText = "<table>"
        for rowNum in range(finalPopulatedRow+1):
            printText += "<tr>"
            for colNum in range(finalPopulatedCol+1):
                printText += "<td>%s</td>" % self.spreadsheetData.GetValue(rowNum, colNum)
            printText += "</tr>"
        printText += "</table>"
        return printText
    
    def __findFinalPopulatedCol(self):
        """Finds the final populated column"""
        finalColNum = 0
        for col in range(NUMBER_GRID_COLS):
            for row in range(NUMBER_GRID_ROWS):
                if (self.mainGrid.GetCellValue(row, col) != ''):
                    finalColNum = col
        return finalColNum
    
    def __setupDataModel(self):
        """Sets up an instance of class DataModel, used to store the data inside the table"""
        self.spreadsheetData = DataTable()
        self.mainGrid.SetTable(self.spreadsheetData, True)
    
    def __updateContentBarWithCellValue(self, event):
        """Updates the main page content bar when user clicks on a cell"""
        self.mainGrid.GetColLabelValue(event.GetCol()) + self.mainGrid.GetRowLabelValue(event.GetRow())
        self.currentFieldText.SetValue(self.mainGrid.GetColLabelValue(event.GetCol()) + self.mainGrid.GetRowLabelValue(event.GetRow()))
        displayString = self.spreadsheetData.getFormula(event.GetRow(), event.GetCol())
        if (displayString):
            self.fieldContentText.SetValue(displayString)
        else:
            self.fieldContentText.SetValue(self.mainGrid.GetCellValue(event.GetRow(), event.GetCol()))
        self.spreadsheetData.refreshFormulas()
        self.mainGrid.ForceRefresh()
        event.Skip()

    def __enterContentBar(self, event):
        """Updates the grid when the user presses enter in the content bar"""
        self.spreadsheetData.SetValue(self.mainGrid.GetGridCursorRow(), self.mainGrid.GetGridCursorCol(), self.fieldContentText.GetValue())
        self.spreadsheetData.refreshFormulas()
        self.mainGrid.ForceRefresh()
        
    def OnCloseWindow(self, event):
        self.Destroy() 
    
    def __OnExit(self, event):
        """Deals with the user exiting the app"""
        self.Destroy()

#---Main section

def main():
     # Start GUI
    app = wx.App(redirect=False)
    frame = MainFrame(None, -1, "pyXL")
    frame.Show(True)
    app.MainLoop()
    return 0

if __name__ == '__main__':
    main()