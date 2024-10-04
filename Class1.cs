Private Sub CommandButton1_Click()
    Call ExcelRangeToWord
End Sub

Sub ExcelRangeToWord()
 
    'https://www.thespreadsheetguru.com/copy-paste-an-excel-table-into-microsoft-word-with-vba/

    'PURPOSE: Copy/Paste An Excel Table Into a New Word Document
    'NOTE: Must have Word Object Library Active in Order to Run _
    (VBE > Tools > References > Microsoft Word 12.0 Object Library)
    'SOURCE: www.TheSpreadsheetGuru.com
 
    Dim tbl As Excel.range
    Dim WordApp As Word.Application
    Dim myDoc As Word.Document
    Dim WordTable As Word.Table
 
    'Optimize Code
    Application.ScreenUpdating = False
    Application.EnableEvents = False
 
    'Create an Instance of MS Word
    On Error Resume Next
    
    'Is MS Word already opened?
    Set WordApp = GetObject(class:="Word.Application")
    
    'Clear the error between errors
    Err.Clear
 
    'If MS Word is not already open then open MS Word
    If WordApp Is Nothing Then Set WordApp = CreateObject(class:="Word.Application")
    
    'Handle if the Word Application is not found
    If Err.Number = 429 Then
        MsgBox "Microsoft Word could not be found, aborting."
        GoTo EndRoutine
    End If
 
    On Error GoTo 0
  
    'Make MS Word Visible and Active
    WordApp.Visible = True
    WordApp.Activate
    
    'Create a New Document
    Set myDoc = WordApp.Documents.Add
   
    With myDoc.PageSetup
        .TopMargin = WordApp.InchesToPoints(1)
        .BottomMargin = WordApp.InchesToPoints(1)
        .LeftMargin = WordApp.InchesToPoints(1)
        .RightMargin = WordApp.InchesToPoints(1.2)
    End With
  
  
  
    'Copy from Excel
    Set tbl = ThisWorkbook.Worksheets(Sheet1.Name).range("B5:F12")
    tbl.Copy
    'Paste Table into MS Word
    Call myDoc.Paragraphs(1).range.PasteExcelTable(LinkedToExcel:=False, WordFormatting:=False, RTF:=False)
    'Autofit Table so it fits inside Word Document
    Set WordTable = myDoc.Tables(1)
    WordTable.AutoFitBehavior (wdAutoFitWindow)
   
    'Copy from Excel
    Set tbl = ThisWorkbook.Worksheets(Sheet1.Name).range("B5:F12")
    tbl.Copy
    'Paste Table into MS Word
    With myDoc.Content
        .InsertParagraphAfter
        .Paragraphs.Last.range.PasteExcelTable LinkedToExcel:=False, WordFormatting:=False, RTF:=False
    End With
    Set WordTable = myDoc.Tables(2)
    WordTable.AutoFitBehavior (wdAutoFitWindow)
  
    'New page???
    myDoc.Sections.Add
    myDoc.Sections(2).PageSetup.Orientation = wdOrientLandscape
    '  myDoc.Content.InsertParagraphAfter
    'myDoc.Content.InsertAfter ("ab")
    'myDoc.Content.InsertBreak Word.WdBreakType.wdSectionBreakNextPage
   
  
    Set tbl = ThisWorkbook.Worksheets(Sheet1.Name).range("B5:F12")
    tbl.Copy
    With myDoc.Content
        .InsertParagraphAfter
        .Paragraphs.Last.range.PasteExcelTable LinkedToExcel:=False, WordFormatting:=False, RTF:=False
    End With
    Set WordTable = myDoc.Tables(3)
    WordTable.AutoFitBehavior (wdAutoFitWindow)
    'WordApp.Selection.PageSetup.Orientation = wdOrientLandscape
   
    'WordTable.PreferredWidthType = wdPreferredWidthPoints ' may not be reflected
    'WordTable.PreferredWidth = WordApp.InchesToPoints(9)
    'WordTable.Cell(2, 1).Width = WordApp.InchesToPoints(1)
    WordTable.PreferredWidthType = wdPreferredWidthPercent
    WordTable.PreferredWidth = 120               '%
   
    Dim c1 As Word.Cells
   
    Dim myRange As Word.range
    Set myRange = myDoc.range(Start:=myDoc.Tables(3).Cell(2, 2).range.Start, End:=myDoc.Tables(3).Cell(2, 5).range.End)
    myRange.Select
    Word.Selection.Cells.Width = 40              '   Word.Selection.Cells.DistributeWidth
    Word.Selection.ParagraphFormat.Alignment = wdAlignParagraphCenter

    
    Dim rng1 As Variant
    Set rng1 = myDoc.range(Start:=myDoc.Tables(3).Cell(2, 1).range.Start, End:=myDoc.Tables(3).Cell(2, 5).range.End).Cells
  
    Dim cc As Word.Cell
    For Each cc In rng1
        cc.Width = WordApp.InchesToPoints(1)
        cc.range.ParagraphFormat.Alignment = wdAlignParagraphCenter
    Next cc
   
   
   
    'Selection.Cells.DistributeWidth
    'Call myDoc.Paragraphs(2).Range.PasteExcelTable(LinkedToExcel:=False, WordFormatting:=False, RTF:=False)
   
    'Call myDoc.Paragraphs(myDoc.Paragraphs.Count).Range.PasteExcelTable(LinkedToExcel:=False, WordFormatting:=False, RTF:=False)
   
   
   
   
EndRoutine:
    'Optimize Code
    Application.ScreenUpdating = True
    Application.EnableEvents = True
 
    'Clear The Clipboard
    Application.CutCopyMode = False
 
End Sub


