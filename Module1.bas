Attribute VB_Name = "Module1"
Sub ORDER_SHEET()
Attribute ORDER_SHEET.VB_ProcData.VB_Invoke_Func = "A\n14"
'
' ORDER_SHEET Macro
'
' Keyboard Shortcut: Ctrl+Shift+A
'

    Dim macroWorkbook As Workbook
    Set macroWorkbook = ThisWorkbook
    
    Dim isAutozone As Boolean
    isAutozone = Not Mid(Sheets(1).Range("C2").Value, 8, 1) = "A"
    

    Sheets.Add after:=ActiveSheet
    Sheets(1).Activate
    If isAutozone Then
    Range("B2").Copy
    Else
    Range("C2").Copy
    End If
    Sheets("Sheet1").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Selection.Font.Bold = True
    Columns("A:A").EntireColumn.AutoFit
    Range("A2") = "PART"
    Range("B2") = "ORDER"
    Range("C2") = "PULL"
    Range("D2") = "INV"
    Range("E2") = "SITE"
    Range("F2") = "SIZE"
    Range("G2") = "ROTATE"
    If isAutozone Then
    Range("H2") = "NEW"
    Range("H2").Font.Bold = True
    End If

    Call Copy_Parts
    
    
    Range("A2:G2").Select
    Range("A2:G2").Select
    Selection.Font.Bold = True
    Dim lastRow As String
    lastRow = Cells(Rows.Count, 1).End(xlUp).Row
    If isAutozone Then
    Range("A2:H" & lastRow).Select
    Else
    Range("A2:G" & lastRow).Select
    End If
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
    Range("B" & lastRow + 1).Select
    Application.CutCopyMode = False
    ActiveCell.FormulaR1C1 = "=SUM(R[-" & lastRow - 2 & "]C:R[-1]C)"
    If isAutozone Then
    Range("D3").Formula = macroWorkbook.Sheets(1).Range("D6").Formula
    Range("E3").Formula = macroWorkbook.Sheets(1).Range("E6").Formula
    Range("F3").Formula = macroWorkbook.Sheets(1).Range("F6").Formula
    Range("G3").Formula = macroWorkbook.Sheets(1).Range("G6").Formula
    Range("H3").Formula = macroWorkbook.Sheets(1).Range("H6").Formula
    Range("H3:H3").AutoFill Destination:=Range("H3:H" & lastRow)
    Else
    Range("D3").Formula = macroWorkbook.Sheets(1).Range("D3").Formula
    Range("E3").Formula = macroWorkbook.Sheets(1).Range("E3").Formula
    Range("F3").Formula = macroWorkbook.Sheets(1).Range("F3").Formula
    Range("G3").Formula = macroWorkbook.Sheets(1).Range("G3").Formula
    End If
    Range("D3:G3").AutoFill Destination:=Range("D3:G" & lastRow)
    Range("D3:H" & lastRow).Select
    Range("A2:H2").AutoFilter
    Range("H1").Select
    Columns("G:G").EntireColumn.AutoFit
    
    Dim errorCheck As Integer
    errorCheck = 1
    For errorCheck = 3 To lastRow
        If Range("G" & errorCheck).Text = "#N/A" Then
        Range("G" & errorCheck) = "-1"
        End If
    
    Next errorCheck
    
    
    ActiveWorkbook.Worksheets("Sheet1").AutoFilter.Sort.SortFields.Clear
    ActiveWorkbook.Worksheets("Sheet1").AutoFilter.Sort.SortFields.Add2 Key:= _
        Range("G2"), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:= _
        xlSortNormal
        
        
        
    With ActiveWorkbook.Worksheets("Sheet1").AutoFilter.Sort
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    
    
    
    Range("B1").Select
    Call Set_Date
    Range("B1").Font.Bold = True
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
    If isAutozone Then
    Range("D3:H" & lastRow).Select
    Else
    Range("D3:G" & lastRow).Select
    End If
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
    Range("B3:B" & lastRow).Select
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
    
    If isAutozone Then
    Dim start As Integer
    start = 1
    For start = 3 To lastRow
        If Not Range("H" & start).Text = "NEW" Then
        Range("H" & start) = ""
        End If
    
    Next start
    End If
    
    Sheets("Sheet1").Copy after:=Sheets(2)
    Sheets("Sheet1").Copy after:=Sheets(3)
    If isAutozone Then
    Dim isGold As Range
    Set isGold = Columns(1).Find("*G*")
        If Not isGold Is Nothing Then
        Sheets("Sheet1").Copy after:=Sheets(4)
        Sheets("Sheet1 (4)").Name = "GOLD"
        End If
    Sheets("Sheet1 (2)").Name = "PADS"
    Sheets("Sheet1 (3)").Name = "SHOES"
    Else
    Sheets("Sheet1 (2)").Name = "BB"
    Sheets("Sheet1 (3)").Name = "BBS"
    End If
    Call Delete_Rows
    Call Edge_Code
'    delete rows only works when sorted by rotation
End Sub
Sub Set_Date()
Attribute Set_Date.VB_ProcData.VB_Invoke_Func = "S\n14"
'
' Set_Date Macro
'
'
'

    Dim isAutozone As Boolean
    isAutozone = Not Mid(Sheets(1).Range("C2").Value, 8, 1) = "A"
    
    Dim orderDate As String
    Dim shipDate As String
    
    If isAutozone Then
    orderDate = Sheets(1).Range("C2").Value
    shipDate = Sheets(1).Range("D2").Value
    Else
    orderDate = Sheets(1).Range("D2").Value + 1
    shipDate = Sheets(1).Range("G2").Value - 1
    End If
    ActiveCell = "ORDER: " & Mid(orderDate, 5, 2) & "/" & Right(orderDate, 2) & "/" & Left(orderDate, 4) & "          SHIP: " & Mid(shipDate, 5, 2) & "/" & Right(shipDate, 2) & "/" & Left(shipDate, 4)
    
End Sub
Sub Edge_Code()
'
' Edge_Code Macro
'
'
'
    Dim WS_Count As Integer
    Dim I As Integer
    WS_Count = ActiveWorkbook.Worksheets.Count
    
    Dim orderNum As String
    If Mid(Sheets(1).Range("C2").Value, 8, 1) = "A" Then
    orderNum = Mid(Sheets(1).Range("C2").Value, 9, 2)
    
    Else: orderNum = Left(Sheets(1).Range("B2").Value, 2)
    
    End If
    
    For I = 2 To WS_Count
    ActiveWorkbook.Sheets(I).Activate
    
    If orderNum = 23 Or orderNum = 25 Or orderNum = 26 Or orderNum = 10 Then
        
    Rows("1:1").Select
    Selection.Insert Shift:=xlDown
    Selection.Insert Shift:=xlDown
    Range("A1:G2").Select
    Range("A2").Activate
    With Selection
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlBottom
        .WrapText = True
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
    Selection.Merge
    Range("A1") = "PULL ALL PARTS AFTER 8035. IF THERE ARE NONE, SEPERATE, TAKE TO REPACK AREA AND LABEL PALLET WITH FULL PO #"
    Range("A1").Font.Bold = True
    End If
    
    If orderNum = 10 Then
    Range("A1") = "PULL ALL PARTS AFTER 2/5/2018. IF THERE ARE NONE, SEPARATE, TAKE TO REPACK AREA AND LABEL PALLET"
    Range("A1").Font.Bold = True
    
    End If
    
    Next I
End Sub

Sub Delete_Rows()
'
' Macro2 Macro
'
' Keyboard Shortcut: Ctrl+Shift+S
'

    Dim isAutozone As Boolean
    isAutozone = Not Mid(Sheets(1).Range("C2").Value, 8, 1) = "A"
    
    If isAutozone Then
    Dim d As String
    d = "*D*"
    Dim isGold As Range
    Set isGold = Columns(1).Find("*G*")
        If Not isGold Is Nothing Then
        Dim goldStart As String
        goldStart = WorksheetFunction.Match("*G*", Range("A:A"), 0)
        
        Dim goldEnd As Range
        Set goldEnd = Range("A:A").Find(what:=d, after:=Range("A1"), searchorder:=xlByColumns, searchdirection:=xlPrevious)
        Sheets("PADS").Select
        Rows(goldStart & ":" & WorksheetFunction.Match(goldEnd.Value, Range("A:A"), 0)).Delete Shift:=xlUp
        
        Sheets("SHOES").Select
        Rows(goldStart & ":" & WorksheetFunction.Match(goldEnd.Value, Range("A:A"), 0)).Delete Shift:=xlUp
        
        Sheets("GOLD").Select
        Rows("3:" & goldStart - 1).Delete Shift:=xlUp
        End If
    Sheets("PADS").Select
'    Dim d As String
    
    Dim padRangeEnd As Range
'    d = "*D*"
    Set padRangeEnd = Range("A:A").Find(what:=d, after:=Range("A1"), searchorder:=xlByColumns, searchdirection:=xlPrevious)
    Dim lastRow As String
    lastRow = Cells(Rows.Count, 1).End(xlUp).Row
    Rows(WorksheetFunction.Match(padRangeEnd.Value, Range("A:A"), 0) + 1 & ":" & lastRow).Delete Shift:=xlUp
    
    Sheets("SHOES").Select
    Rows("3:" & WorksheetFunction.Match(padRangeEnd.Value, Range("A:A"), 0)).Delete Shift:=xlUp
        
    Else
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
    End If
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
        "A2"), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:= _
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
        "A2"), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:= _
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

Sub test()
Attribute test.VB_ProcData.VB_Invoke_Func = "n\n14"
'
' Test Macro
'

'

'    Range("B14").FormulaR1C1 = "=""'[VDP PO ""&TEXT(TODAY(),""yyyymmdd"")&"".xlsx]BASE'!"""
'    Range("B15").Formula = "=VLOOKUP(B2,INDIRECT(B14&""C:R""),16,0)"
'    Dim macroWorkbook As Workbook
'    Set macroWorkbook = ThisWorkbook
    
    Dim isAutozone As Boolean
    isAutozone = Not Mid(Sheets(1).Range("C2").Value, 8, 1) = "A"
    
    
    
'    Range("G3").Select
'    If isAutozone Then
'    Dim test
'    Set test = WorksheetFunction.Match("*G*", Range("A:A"), 0)


'    Dim job As String
'    Dim searchTerm As Range
'    job = "*D*"
'    Set searchTerm = Range("A:A").Find(what:=job, after:=Range("A1"), searchorder:=xlByColumns, searchdirection:=xlPrevious)
'    ActiveCell = WorksheetFunction.Match(searchTerm.Value, Range("A:A"), 0)
   
'    Range("L6") = Range("H6").Value
'    If Range("H5").Text = "NEW" Then
'    Range("L6") = Range("H6").Value
'    Else
'    Range("L7") = "NO"
'    End If

    Dim lastRow As String
    lastRow = Cells(Rows.Count, 1).End(xlUp).Row
    
    Dim errorCheck As Integer
    errorCheck = 1
    For errorCheck = 3 To lastRow
        If Range("G" & errorCheck).Text = "#N/A" Then
        Range("G" & errorCheck) = "-1"
        End If
    
    Next errorCheck
    
 '   Dim start As Integer
 '   start = 1
 '   For start = 3 To lastRow
 '   Range("K" & start) = start
 '       If Range("H" & start).Value = "NEW" Then
 '       Range("H" & start).Clear
 '       End If
 '   Next start
    
    


    
'    Dim isGold As Range
'    Set isGold = Columns(1).Find("*G*")
'    If Not isGold Is Nothing Then
'    ActiveCell = "No"
'    Else
'    ActiveCell = "Yes"
'    End If
   


'    Dim lastRow As String
 '   lastRow = Cells(Rows.Count, 1).End(xlUp).Row
    
  '  Range("B" & lastRow + 1).Select
  '  ActiveCell.FormulaR1C1 = "=SUM(R[-" & lastRow - 2 & "]C:R[-1]C)"
  
  
  'FOR EDGECODE
  ' USE SHIP STATE, CA OR WA
  
  
  
  'FOR AZ OR OR
  ' IF H3 = OR_SKU
  
  'FOR AZ ADD NEW COLUMN TO H
  'SAME PROCEDURE OTHERWISE
  'MAYBE FOR INV ROTATE SIZE ETC INSTEAD OF COPYING ALL INV INTO TEMP, HAVE JUST THE FORMULA LINKING TO INV, THEN COPY THAT TO SHEETS
  
  
    
    
End Sub


