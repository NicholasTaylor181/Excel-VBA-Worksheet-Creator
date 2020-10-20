Attribute VB_Name = "Module1"
Sub ORDER_SHEET()
Attribute ORDER_SHEET.VB_ProcData.VB_Invoke_Func = "A\n14"
'
' ORDER_SHEET Macro
'
' Keyboard Shortcut: Ctrl+Shift+A
'
    Sheets.Add After:=ActiveSheet
    Sheets(1).Activate
    Range("C2").Select
    Selection.Copy
    Sheets("Sheet1").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Selection.Font.Bold = True
    Columns("A:A").EntireColumn.AutoFit
    Range("A2").Select
    ActiveCell.FormulaR1C1 = "PART"
    Range("B2").Select
    ActiveCell.FormulaR1C1 = "ORDER"
    Range("C2").Select
    ActiveCell.FormulaR1C1 = "PULL"
    Range("D2").Select
    ActiveCell.FormulaR1C1 = "INV"
    Range("E2").Select
    ActiveCell.FormulaR1C1 = "SITE"
    Range("F2").Select
    ActiveCell.FormulaR1C1 = "SIZE"
    Range("G2").Select
    ActiveCell.FormulaR1C1 = "ROTATE"

    Call Copy_Parts
    
    
    Range("A2:G2").Select
    Range("A2:G2").Select
    Selection.Font.Bold = True
    Range("A2:G239").Select
    Selection.Borders(xlDiagonalDown).LineStyle = xlNone
    Selection.Borders(xlDiagonalUp).LineStyle = xlNone
    With Selection.Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlInsideVertical)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlInsideHorizontal)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    ActiveWindow.SmallScroll Down:=228
    Range("B240").Select
    Application.CutCopyMode = False
    ActiveCell.FormulaR1C1 = "=SUM(R[-237]C:R[-1]C)"
    Range("B241").Select
    ActiveWindow.SmallScroll Down:=-243
    Range("D3").Select
    ActiveCell.FormulaR1C1 = "=TODAY()"
    Range("E3").Select
    ActiveCell.FormulaR1C1 = "=TODAY()"
    Range("F3").Select
    ActiveCell.FormulaR1C1 = "=TODAY()"
    Range("G3").Select
    ActiveCell.FormulaR1C1 = "=TODAY()"
    Range("D3:G3").Select
    Selection.AutoFill Destination:=Range("D3:G239")
    Range("D3:G239").Select
    Range("A2:G2").Select
    Selection.AutoFilter
    Range("H1").Select
    ActiveWorkbook.Worksheets("Sheet1").AutoFilter.Sort.SortFields.Clear
    ActiveWorkbook.Worksheets("Sheet1").AutoFilter.Sort.SortFields.Add2 Key:= _
        Range("G2"), SortOn:=xlSortOnValues, Order:=xlDescending, DataOption:= _
        xlSortNormal
    With ActiveWorkbook.Worksheets("Sheet1").AutoFilter.Sort
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    Range("B1").Select
    ActiveCell.FormulaR1C1 = "ORDER"
    Range("B1:G1").Select
    With Selection
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlBottom
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
    Selection.Merge
    Range("D3:G239").Select
    With Selection
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlBottom
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
    ActiveWindow.SmallScroll Down:=-12
    Range("B3:B239").Select
    With Selection
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlBottom
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
    ActiveWindow.SmallScroll Down:=-12
    Sheets("Sheet1").Select
    Sheets("Sheet1").Copy After:=Sheets(2)
    Sheets("Sheet1").Select
    Sheets("Sheet1").Copy After:=Sheets(3)
    Sheets("Sheet1 (2)").Select
    Sheets("Sheet1 (2)").Name = "BB"
    Sheets("Sheet1 (3)").Select
    Sheets("Sheet1 (3)").Name = "BBS"
    
'    Call Delete_Rows
'    delete rows only works when sorted by rotation
End Sub
Sub Macro2()
Attribute Macro2.VB_ProcData.VB_Invoke_Func = "S\n14"
'
' Macro2 Macro
'
' Keyboard Shortcut: Ctrl+Shift+S
'
    Windows("850_OReillyAuto_6248280A03GF00.csv").Activate
    Sheets("Sheet1").Select
    Sheets("Sheet1").Copy After:=Sheets(2)
    Sheets("Sheet1").Select
    Sheets("Sheet1").Copy After:=Sheets(3)
    Sheets("Sheet1 (2)").Select
    Sheets("Sheet1 (2)").Name = "BB"
    Sheets("Sheet1 (3)").Select
    Sheets("Sheet1 (3)").Name = "BBS"
    Sheets("BB").Select
    ActiveWindow.SmallScroll Down:=132
    Rows("139:239").Select
    Selection.Delete Shift:=xlUp
    ActiveWindow.SmallScroll Down:=-165
    Sheets("BBS").Select
    Rows("3:138").Select
    Selection.Delete Shift:=xlUp
End Sub

Sub Delete_Rows()
'
' Macro2 Macro
'
' Keyboard Shortcut: Ctrl+Shift+S
'

    Sheets("BB").Select
    Dim bbRangeStart As String
    bbRangeStart = WorksheetFunction.Match("*S*", Range("A:A"), 0)
    Dim bbRangeEnd As String
    bbRangeEnd = WorksheetFunction.Match("*D*", Range("A:A"), 0)
    
    Rows(bbRangeStart & ":" & bbRangeEnd - 1).Select
    Selection.Delete Shift:=xlUp

    
    Sheets("BBS").Select
    Dim bbsRangeStart As String
    bbsRangeStart = WorksheetFunction.Match("*D*", Range("A:A"), 0)
    Dim lRow As String
    lRow = Cells(Rows.Count, 1).End(xlUp).Row
    Rows(bbsRangeStart & ":" & lRow).Select
    Selection.Delete Shift:=xlUp
End Sub
Sub Filter_test()
Attribute Filter_test.VB_ProcData.VB_Invoke_Func = " \n14"
'
' Filter_test Macro
'

'
    Windows("850_OReillyAuto_6241276A27MZ00.xlsx").Activate
    ActiveWindow.SmallScroll Down:=-51
    Windows("850_OReillyAuto_6248280A03GF00.csv").Activate
    ActiveWorkbook.Worksheets("BB").AutoFilter.Sort.SortFields.Clear
    ActiveWorkbook.Worksheets("BB").AutoFilter.Sort.SortFields.Add2 Key:=Range( _
        "A2"), SortOn:=xlSortOnValues, Order:=xlDescending, DataOption:= _
        xlSortNormal
    With ActiveWorkbook.Worksheets("BB").AutoFilter.Sort
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    ActiveWorkbook.Worksheets("BB").AutoFilter.Sort.SortFields.Clear
    ActiveWorkbook.Worksheets("BB").AutoFilter.Sort.SortFields.Add2 Key:=Range( _
        "A2"), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:= _
        xlSortNormal
    With ActiveWorkbook.Worksheets("BB").AutoFilter.Sort
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    ActiveWorkbook.Worksheets("BB").AutoFilter.Sort.SortFields.Clear
    ActiveWorkbook.Worksheets("BB").AutoFilter.Sort.SortFields.Add2 Key:=Range( _
        "A2"), SortOn:=xlSortOnValues, Order:=xlDescending, DataOption:= _
        xlSortNormal
    With ActiveWorkbook.Worksheets("BB").AutoFilter.Sort
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    Sheets("BBS").Select
    ActiveWindow.SmallScroll Down:=-120
    ActiveWorkbook.Worksheets("BBS").AutoFilter.Sort.SortFields.Clear
    ActiveWorkbook.Worksheets("BBS").AutoFilter.Sort.SortFields.Add2 Key:=Range( _
        "A2"), SortOn:=xlSortOnValues, Order:=xlDescending, DataOption:= _
        xlSortNormal
    With ActiveWorkbook.Worksheets("BBS").AutoFilter.Sort
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
End Sub
Sub Copy_Parts()
'
' Copy_Parts Macro
'

'
    Sheets(1).Activate
    
    Dim lRow As String
    lRow = Cells(Rows.Count, 1).End(xlUp).Row
    Range(Cells(4, 7), Cells(lRow, 7)).Select

    Selection.Copy
    Sheets("Sheet1").Select
    Range("A3").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Sheets(1).Activate
    Range(Cells(4, 3), Cells(lRow, 3)).Select
    Application.CutCopyMode = False
    Selection.Copy
    Sheets("Sheet1").Select
    Range("B3").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
End Sub

Sub Select_Sheet()
'
' Select_Sheet Macro
'

'
    Sheets(1).Activate
    Range("A3").Select
    
    
End Sub


