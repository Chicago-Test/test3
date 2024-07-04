Attribute VB_Name = "Module1"
Option Explicit

Sub checkCellLock()
    On Error GoTo ErrHandler:
    Dim str As String
    Dim sh As Worksheet
    
    Dim totalCount As Long                       ' How many cells checked
    Dim totalCountofFormula As Long              ' How many formula cells checked
    Dim maxrow, maxcol As Long
    Dim isFormula    As Boolean
    Dim isLocked As Boolean
    Dim isSheetProtected As Boolean
    
    Dim isLockedData() As Byte
    Dim prev_hash As String
    Dim xlwb As Excel.Workbook
    Set xlwb = ActiveWorkbook
    For Each sh In xlwb.Worksheets
        isSheetProtected = sh.ProtectContents
    
        'check every cell
        Dim addr As String
        Dim cl As Range
        
        'Get last cell positon excluding empty cells with format style
        'When sheet is empty this cause a problem...
        Dim maxUsedRow, maxUsedCol As Long
        maxUsedRow = FindLastRow(sh)
        maxUsedCol = FindLastColumn(sh)
        
        ReDim isLockedData(maxUsedRow, maxUsedCol)
        
        If maxUsedRow < 1 Or maxUsedCol < 1 Then 'Maybe this is empty sheet so skip
            GoTo continue2
        End If
        
        If (False = sh.Visible And True = xlwb.ProtectStructure) Or (sh.Visible = xlSheetVeryHidden) Then
        End If
        
        'If the sheet is protected and if Worksheet.EnableSelection property is 'xlNoSelection' user cannot access even the cell is unlocked.
        If (True = isSheetProtected) And (sh.EnableSelection < 0) Then
        End If
        
        
        maxrow = maxUsedRow: maxcol = maxUsedCol
        
        'SpecialCells(xlCellTypeFormulas) does not work in protected sheet
        For Each cl In sh.Range(sh.Cells(1, 1), sh.Cells(maxrow, maxcol))
            isFormula = cl.HasFormula
            isLocked = cl.Locked
            If cl.Locked = False Then isLockedData(cl.Row, cl.Column) = 1
            
            totalCount = totalCount + 1
            If isFormula = True Then totalCountofFormula = totalCountofFormula + 1
            'End If
            
            'Formula cell check
            If (isSheetProtected = False And isFormula = True) Or (isSheetProtected = True And isFormula = True And isLocked = False) Then
            End If
            
        Next cl
        str = Application.Run("yoji", isLockedData, sh.Index, prev_hash)
        prev_hash = str
        
continue2:

    Next sh
    
    Exit Sub
ErrHandler:
End Sub

Sub test()
    Dim ret As Object
    Dim str As String
    Dim x(2, 2) As Byte
    x(1, 1) = 1
    x(1, 2) = 0
    'Set ret = Application.Run("yoji", ActiveSheet.Range("a1"))

    'str = Application.Worksheets("Sheet1").Range("H7")

    'str = Application.Run("yoji", 55)
    'Set ret = Application.Run("yoji", x)
    '  Call Application.Run("ReadFormulasMacroType", x)

End Sub

Public Function FindLastRow(sh As Worksheet) As Long
    ' --------------------------------------------------------------------------------
    ' Find the last used Row on a Worksheet
    ' --------------------------------------------------------------------------------
    FindLastRow = 0
    Dim n1 As Long: n1 = 0
    Dim n2 As Long: n2 = 0
    
    On Error Resume Next
    n1 = sh.UsedRange.Find("*", , xlFormulas, , xlByRows, xlPrevious).Row
    n2 = sh.UsedRange.Rows.Count                 'When all the formula cells are hidden, 'sh.UsedRange.Find' may fail
    On Error GoTo 0                              'reset normal error handling
    On Error GoTo ErrHandler:
    
    If n1 <= 0 Then
        FindLastRow = n2
    Else
        FindLastRow = n1
    End If

    'If WorksheetFunction.CountA(sh.Cells) > 0 Then
    ' Search for any entry, by searching backwards by Rows.
    'FindLastRow = ws.Cells.Find(What:="*", After:=ws.Range("a1"), SearchOrder:=xlByRows, SearchDirection:=xlPrevious).Row
    'End If
    Exit Function
ErrHandler:
    FindLastRow = 0
End Function

Public Function FindLastColumn(sh As Worksheet) As Long
    ' --------------------------------------------------------------------------------
    ' Find the last used Column on a Worksheet
    ' --------------------------------------------------------------------------------
    FindLastColumn = 0
    Dim n1 As Long: n1 = 0
    Dim n2 As Long: n2 = 0
    
    On Error Resume Next
    n1 = sh.UsedRange.Find("*", , xlFormulas, , xlByColumns, xlPrevious).Column
    n2 = sh.UsedRange.Columns.Count
    On Error GoTo 0
    On Error GoTo ErrHandler:
    
    If n1 <= 0 Then
        FindLastColumn = n2
    Else
        FindLastColumn = n1
    End If
    
    Exit Function
ErrHandler:
    FindLastColumn = 0
End Function

