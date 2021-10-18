Private Declare PtrSafe Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As LongPtr, ByVal lpOperation As String, _
ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As LongPtr
Option Explicit

Dim cmdDict As Scripting.Dictionary, creationDict As Scripting.Dictionary
Sub sayHi(control As IRibbonControl)
    'Placeholder macro for buttons with unassigned macros
    MsgBox "Hi :)"
End Sub
Public Sub openWorkbook(control As IRibbonControl)
    'Open workbook path passed through the tag provided to this button in the XML file. Opens in this session
    Application.DisplayAlerts = False
    Workbooks.Open control.Tag, False
    Application.DisplayAlerts = True
End Sub

Sub openDatedFile(control As IRibbonControl)
    Dim fso As Object, FindDate As Date
    
    Dim Path As String
    Path = control.Tag
    
    'strip out date format from xml tag
    Dim datefmt As String
    datefmt = Left(Right(Path, Len(Path) - InStr(1, Path, "#")), InStr(1, Right(Path, Len(Path) - InStr(1, Path, "#")), "#") - 1)
    Path = Replace(Path, "#", "")
    
    'begin searching 7 days from now
    Set fso = CreateObject("Scripting.FileSystemObject")
    FindDate = Date + 7

    'from 7 days from now until two months, check each day for a file with that date in it's name. once found, open the file
    Do While Not fso.FileExists(Replace(Path, datefmt, Format(FindDate, datefmt)))
        FindDate = FindDate - 1
        If FindDate < Date - 60 Then
            MsgBox "No file found in last 60 days. Git gud"
            Exit Sub
            End If
    Loop
    Application.DisplayAlerts = False
    Workbooks.Open Replace(Path, datefmt, Format(FindDate, datefmt))
    Application.DisplayAlerts = True
End Sub
Sub openNonExcel(control As IRibbonControl)
    'used to open non excel files (pdfs, word docs, etc.) using the shell
    ShellExecute Application.hwnd, "Open", control.Tag, 0&, 0&, 0&
End Sub
Sub openFolder(control As IRibbonControl)
    'Opens folder specified in XML tags through file explorer in shell
    Dim FolderName As String
    FolderName = control.Tag
    Shell "explorer.exe " & FolderName, 1
End Sub
Sub copySelection(control As IRibbonControl)
    'copies selected visible cells
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    Selection.SpecialCells(xlCellTypeVisible).Select
    Selection.Copy
    Workbooks.Add
    'pasted selection values, format, and column width into new workbook
    Selection.PasteSpecial Paste:=xlPasteColumnWidths, Operation:=xlNone, _
        SkipBlanks:=False, Transpose:=False
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Selection.PasteSpecial Paste:=xlPasteFormats, Operation:=xlNone, _
        SkipBlanks:=False, Transpose:=False
    Application.DisplayAlerts = True
    Application.ScreenUpdating = True
End Sub
