Attribute VB_Name = "Utilities"
Option Explicit
Option Base 0

Function User() As String
    'Zwraca nazwê u¿ytkownika komputera
    User = Application.UserName
End Function

Function ExcelDir()
    'Zwraca œcie¿kê w której jest zainstalowany Excel
    ExcelDir = Application.Path
End Function

Function SheetCount()
    'Zwraca liczbê arkuszy w skoroszycie
    SheetCount = Application.Caller.Parent.Parent.Sheets.Count
End Function

Function SheetName()
    'Zwraca nazwê skoroszytu
    SheetName = Application.Caller.Parent.Name
End Function

Function GetPositionOfSheet() As Integer
    'Funkcja s³u¿¹ca do okreœlenia pozycji arkusza w skoroszycie
    Dim currentSheetName As String
    Dim i As Integer
    
    currentSheetName = SheetName()
    
    GetPositionOfSheet = 1
    For i = 1 To SheetCount
        If ThisWorkbook.Sheets(i).Name = currentSheetName Then
            GetPositionOfSheet = i
            Exit Function
        End If
    Next i
    GetPositionOfSheet = 0
End Function

Function AutoNumber(Optional maxPage As Integer) As String
    'Funkcja s³u¿¹ca do automatycznego numerowania stron
    If maxPage = 0 Then
        maxPage = SheetCount
    End If
    AutoNumber = GetPositionOfSheet & "/" & maxPage
End Function

Function GetCaller()
    Dim v As String
    Select Case TypeName(Application.Caller)
        Case "Range"
            v = Application.Caller.Address
        Case "String"
            v = Application.Caller
        Case "Error"
            v = "Error"
        Case Else
            v = "unknown"
    End Select
    
    GetCaller = v
End Function

Sub SortSheets()
    ' Sorts the sheets of the active workbook
    
    'Display message asking the user to confirm the action
    If MsgBox("Sort the sheets in the active workbook?", vbQuestion + vbYesNo) <> vbYes Then Exit Sub
    
    'Prevent user from breaking the macro
    Application.EnableCancelKey = xlDisabled
    'Exit Sub if there is no workbook active - prevents further errors
    If ActiveWorkbook Is Nothing Then Exit Sub
    
    'Check if workbook is protected and display message box
    If ActiveWorkbook.ProtectStructure Then
        MsgBox ActiveWorkbook.Name & " is protected.", _
        vbCritical, "Cannot sort sheets."
        Exit Sub
    End If
        
    'Turn off screen updating while the sheets are being moved
    Application.ScreenUpdating = False
    
    'Declaring some variables
    Dim SheetNames() As String
    Dim i As Long
    Dim SheetCount As Long
    Dim OldActive As Worksheet
    
    Set OldActive = ActiveSheet
    
    
    'Determine the number of sheets & ReDim Array
    SheetCount = ActiveWorkbook.Sheets.Count
    ReDim SheetNames(1 To SheetCount)
    
    'Fill array with sheet names
    For i = 1 To SheetCount
        SheetNames(i) = ActiveWorkbook.Sheets(i).Name
    Next i
    'Sort the array in ascending order
    Call BubbleSort(SheetNames)
    
    'Move the sheets
    For i = 1 To SheetCount
        ActiveWorkbook.Sheets(SheetNames(i)).Move _
        Before:=ActiveWorkbook.Sheets(i)
    Next i
    
    OldActive.Activate
    
End Sub

Sub BubbleSort(List() As String)
    Dim First As Long, Last As Long
    Dim i As Long, j As Long
    Dim temp As String
    First = LBound(List)
    Last = UBound(List)
    For i = First To Last - 1
        For j = i + 1 To Last
            If UCase(List(i)) > UCase(List(j)) Then
                temp = List(j)
                List(j) = List(i)
                List(i) = temp
            End If
        Next j
    Next i
End Sub

Sub DescribeFunction()
    Dim FuncName As String
    Dim FuncDesc As String
    Dim FuncCat As Long
    Dim Arg1Desc As String
    
    FuncName = "VowelCount"
    FuncDesc = "Zlicza samog³oski"
    FuncCat = 7
    Arg1Desc = "Tekst, do zliczenia samog³osek"
    
    PtrSafe
    
    Application.MacroOptions _
        Macro:=FuncName, _
        Description:=FuncDesc, _
        Category:=FuncCat, _
        ArgumentDescriptions:=Array(Arg1Desc)
End Sub

Public Function ContainsMergedCells(rng As Range)
    'Sprawdza czy w danym zakresie s¹ po³¹czone komórki
    Dim cell As Range
    ContainsMergedCells = False
    For Each cell In rng
        If cell.MergeCells Then
            ContainsMergedCells = True
            Exit Function
        End If
    Next cell
End Function

Sub AutoNumberPages()
    Dim sheet As Worksheet
    Dim i As Integer, lastPage As Integer
    i = 0
    
    lastPage = ActiveWindow.SelectedSheets.Count
    
    For Each sheet In ActiveWindow.SelectedSheets
        i = i + 1
        With sheet.Cells(1, 6)
            .NumberFormat = "@"
            .Value = i & "/" & lastPage
            .HorizontalAlignment = xlCenter
            .Font.Bold = True
        End With
        sheet.Cells(1, 6).NumberFormat = "@"
        sheet.Cells(1, 6).Value = CStr(i) & "/" & CStr(lastPage)
    Next sheet
End Sub

Sub DeleteEmptyColumns(Optional skipPrompt As Boolean = True)
    Dim col As Range
    Dim colsToDelete As Range
    Dim rng As Range
    Dim activeSht As Worksheet
    
    Set activeSht = GetActiveSheet
    On Error Resume Next
    
    Set rng = activeSht.UsedRange
    
    For Each col In rng.Columns
        If WorksheetFunction.CountA(col) = 0 Then
            If colsToDelete Is Nothing Then
                Set colsToDelete = col
            Else
                Set colsToDelete = Union(colsToDelete, col)
            End If
        End If
    Next col
    
    If skipPrompt = False Then
        Dim ans As Integer
        ans = MsgBox(Prompt:="Czy na pewno chcesz usun¹æ puste kolumny?", _
                     Buttons:=vbYesNo, Title:="Potwierdzenie")
        If ans = vbYes Then
            colsToDelete.Delete
        End If
    Else
        colsToDelete.Delete
    End If
    On Error GoTo 0
End Sub

Sub DeleteEmptyRows(Optional skipPrompt As Boolean = True)
    Dim row As Range
    Dim rowsToDelete As Range
    Dim rng As Range
    Dim activeSht As Worksheet
    
    Set activeSht = GetActiveSheet
    On Error Resume Next
    
    Set rng = activeSht.UsedRange
    
    For Each row In rng.Rows
        If WorksheetFunction.CountA(row) = 0 Then
            If rowsToDelete Is Nothing Then
                Set rowsToDelete = row
            Else
                Set rowsToDelete = Union(rowsToDelete, row)
            End If
        End If
    Next row
    
    If skipPrompt = False Then
        Dim ans As Integer
        ans = MsgBox(Prompt:="Czy na pewno chcesz usun¹æ puste wiersze?", _
                     Buttons:=vbYesNo, Title:="Potwierdzenie")
        If ans = vbYes Then
            rowsToDelete.Delete
        End If
    Else
        rowsToDelete.Delete
    End If
    On Error GoTo 0
End Sub

Sub CleanRange()
    Dim rng As Range
    Dim vals As Variant
    If Not TypeOf Selection Is Range Then
        MsgBox "B³êdne zaznaczenie!"
        Exit Sub
    End If
    
    Dim activeSht As Worksheet
    Set activeSht = GetActiveSheet
    
    Dim selectedCells As Range
    Set selectedCells = Intersect(activeSht.UsedRange, Selection)
    
    vals = selectedCells.Value
    Dim val As Variant
    Dim i As Long, j As Long
    For i = LBound(vals, 1) To UBound(vals, 1)
        For j = LBound(vals, 2) To UBound(vals, 2)
            vals(i, j) = WorksheetFunction.Clean(vals(i, j))
            vals(i, j) = Trim(vals(i, j))
        Next j
    Next i
    selectedCells.Value = vals
End Sub

Sub CleanAndDeleteShit()
    Call CleanRange
    Call DeleteEmptyColumns
    Call DeleteEmptyRows
End Sub


Private Function GetActiveSheet()
    GetActiveSheet = Null
    If TypeOf ActiveSheet Is Worksheet Then
        Set GetActiveSheet = ActiveSheet
    End If
End Function

Sub SynchSheets()
    If TypeName(ActiveSheet) <> "Worksheet" Then Exit Sub
    Dim UserSheet As Worksheet, sht As Worksheet
    Dim TopRow As Long, LeftCol As Integer
    Dim UserSel As String
    
    Application.ScreenUpdating = False
    Set UserSheet = ActiveSheet
    
    TopRow = ActiveWindow.ScrollRow
    LeftCol = ActiveWindow.ScrollColumn
    UserSel = ActiveWindow.RangeSelection.Address
    
    For Each sht In ActiveWorkbook.Worksheets
        If sht.Visible Then
            sht.Activate
            Range(UserSel).Select
            ActiveWindow.ScrollRow = TopRow
            ActiveWindow.ScrollColumn = LeftCol
        End If
    Next sht
    
    UserSheet.Activate
    Application.ScreenUpdating = True
End Sub

Public Sub CreateNewWbook()
    Dim NewWbook As Workbook
    Dim NewWbookPath As Variant
    
    On Error Resume Next
    
    Set NewWbook = Workbooks.Item("Obliczenia silnika")
    If Err.Number <> 0 Then
        Set NewWbook = Workbooks.Add()
        NewWbookPath = Application.GetSaveAsFilename(FileFilter:="Excel VBA files (*.xlsm), *.xlsm", _
                        Title:="Zapisz nowy skoroszyt", InitialFileName:="Nowy")
        If NewWbookPath <> False Then
            NewWbook.SaveAs Filename:=NewWbookPath, FileFormat:=xlWorkbookNormal
            Exit Sub
        Else
            MsgBox "Nie podano nazwy" & vbCrLf & "Zamykam skoroszyt!", vbCritical, "Pojebongo?"
            NewWbook.Close SaveChanges:=False
        End If
        Err.Clear
    End If
    On Error GoTo 0
End Sub

Sub SwapCols()
    If TypeName(Selection) <> "Range" Then
        Exit Sub
    End If
    
    If Selection.Areas.Count <> 2 Then
        MsgBox "Proszê wybraæ dwie kolumny do zamiany"
        Exit Sub
    End If
    
    Dim areaFirst As Range, areaSecond As Range
    Set areaFirst = Selection.Areas(1)
    Set areaSecond = Selection.Areas(2)
    
    If areaFirst.Columns.Count > 1 Or areaSecond.Columns.Count > 1 Then
        MsgBox "Proszê wybraæ po jednej kolumnie z ka¿dego zakresu"
        Exit Sub
    End If
    
    If areaFirst.Rows.Count <> areaSecond.Rows.Count Then
        MsgBox "Wybrane zakresy maj¹ ró¿n¹ liczbê wierszy!", vbExclamation
        Exit Sub
    End If
    
    Dim usedRng As Range
    Set usedRng = ActiveSheet.UsedRange
    
    Dim lastRowInSheet As Long
    lastRowInSheet = usedRng.SpecialCells(xlCellTypeLastCell).row
    
    Debug.Print "Last row in sheet: ", lastRowInSheet
    Set areaFirst = Intersect(areaFirst, Range(Cells(1, areaFirst.Column), _
                             Cells(lastRowInSheet, areaFirst.Column)))
    Set areaSecond = Intersect(areaSecond, Range(Cells(1, areaSecond.Column), _
                             Cells(lastRowInSheet, areaSecond.Column)))
    
    Debug.Print "Area First After Intersect: ", areaFirst.Address
    Debug.Print "Area Second After Intersect: ", areaSecond.Address
    
    On Error Resume Next
    
    Dim lastRow_AreaFirst As Long, lastRow_AreaSecond As Long
    
    lastRow_AreaFirst = areaFirst.Rows.Count
    lastRow_AreaSecond = areaSecond.Rows.Count
    
    Debug.Print "Area First last row: ", lastRow_AreaFirst
    Debug.Print "Area second last row: ", lastRow_AreaFirst
    
    Dim areaFirst_Vals As Variant, areaSecond_Vals As Variant
    areaFirst_Vals = areaFirst.Value
    areaSecond_Vals = areaSecond.Value
    
    Dim rowsToInclude As Long
    rowsToInclude = WorksheetFunction.Max(lastRow_AreaFirst, lastRow_AreaSecond)
    
    Debug.Print "Rows to include: ", rowsToInclude
    
    Dim lBndAreaFirst As Long, lBndAreaSecond As Long
    lBndAreaFirst = LBound(areaFirst_Vals)
    lBndAreaSecond = LBound(areaSecond_Vals)
    
    ReDim Preserve areaFirst_Vals(lBndAreaFirst To rowsToInclude)
    ReDim Preserve areaSecond_Vals(lBndAreaSecond To rowsToInclude)
    
    Set areaFirst = areaFirst.Resize(rowsToInclude, areaFirst.Columns.Count)
    Set areaSecond = areaSecond.Resize(rowsToInclude, areaSecond.Columns.Count)
    
    areaFirst.ClearContents
    areaSecond.ClearContents
    
    areaFirst.Value = areaSecond_Vals
    areaSecond.Value = areaFirst_Vals
End Sub
