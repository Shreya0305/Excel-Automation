aSub collectdata()
'
' collectdata Macro
'
' Keyboard Shortcut: Ctrl+Shift+D
'
    Range(Selection, Selection.End(xlDown)).Select
    Range(Selection, Selection.End(xlToRight)).Select
    Selection.Clear
    Range("A1").Select
    
    Sheets("East").Select
    Range("A1").Select
    Range(Selection, Selection.End(xlDown)).Select
    Range(Selection, Selection.End(xlToRight)).Select
    Selection.Copy
    Range("A1").Select
    Sheets("Data").Select
    Range("A2").Select
    ActiveSheet.Paste
    Range("A2").Select
     
    Sheets("West").Select
    Range("A2").Select
    Range(Selection, Selection.End(xlDown)).Select
    Range(Selection, Selection.End(xlToRight)).Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("A2").Select
    Sheets("Data").Select
    Range("A2").Select
    Selection.End(xlDown).Select
    ActiveCell.Offset(1, 0).Range("A1").Select
    ActiveSheet.Paste
    Range("A2").Select
    
     Sheets("South").Select
    Range("A2").Select
    Range(Selection, Selection.End(xlDown)).Select
    Range(Selection, Selection.End(xlToRight)).Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("A2").Select
    Sheets("Data").Select
    Range("A2").Select
    Selection.End(xlDown).Select
    ActiveCell.Offset(1, 0).Range("A1").Select
    ActiveSheet.Paste
    Range("A2").Select
    
    Sheets("North").Select
    Range("A2").Select
    Range(Selection, Selection.End(xlDown)).Select
    Range(Selection, Selection.End(xlToRight)).Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("A2").Select
    Sheets("Data").Select
    Range("A2").Select
    Selection.End(xlDown).Select
    ActiveCell.Offset(1, 0).Range("A1").Select
    ActiveSheet.Paste
    Range("A2").Select
    
    
    Range("A3:A63").Select
    Selection.DataSeries Rowcol:=xlColumns, Type:=xlLinear, Date:=xlDay, _
        Step:=1, Trend:=False
    Range("A11").Select
    Selection.End(xlDown).Select
    
    With Range("A2")
    ActiveSheet.ListObjects.Add(SourceType:=xlSrcRange, Source:=Range(.End(xlDown), .End(xlToRight))).Name = "MyData"
    End With
    
    Range("A2").Select
    
End Sub

Sub copyRPT()
'
' copyRPT Macro
'

'
    Sheets("RPT").Select
    Sheets("RPT").Copy Before:=Sheets(2)
    Sheets("RPT (2)").Select
    Sheets("RPT (2)").Copy Before:=Sheets(2)
    Sheets("RPT (3)").Select
    Sheets("RPT (3)").Copy Before:=Sheets(2)
End Sub

Sub Slicer()
'
' Slicer Macro
'

Sheets("RPT").Select
    ActiveSheet.ChartObjects("MyChart 1").Activate
    ActiveWorkbook.SlicerCaches.Add2(ActiveSheet.PivotTables("MYPT1"), "Date"). _
        Slicers.Add ActiveSheet, , "Date", "Date", 137, 424, 144, 191.249527559055
    ActiveWorkbook.SlicerCaches.Add2(ActiveSheet.PivotTables("MYPT1"), "Product"). _
        Slicers.Add ActiveSheet, , "Product", "Product", 175, 462, 144, _
        191.249527559055
    ActiveWorkbook.SlicerCaches.Add2(ActiveSheet.PivotTables("MYPT1"), "Name"). _
        Slicers.Add ActiveSheet, , "Name", "Name", 213, 500, 144, 191.249527559055
    ActiveWorkbook.SlicerCaches.Add2(ActiveSheet.PivotTables("MYPT1"), "Area"). _
        Slicers.Add ActiveSheet, , "Area", "Area", 251, 538, 144, 191.249527559055
    ActiveSheet.Shapes.Range(Array("Area")).Select
    ActiveSheet.Shapes("Area").IncrementLeft 509
    ActiveSheet.Shapes("Area").IncrementTop -251
    ActiveWorkbook.SlicerCaches("Slicer_Area").Slicers("Area").NumberOfColumns = 4
    ActiveSheet.Shapes("Area").ScaleWidth 1.75, msoFalse, msoScaleFromTopLeft
    ActiveSheet.Shapes("Area").ScaleHeight 0.2509810122, msoFalse, _
        msoScaleFromTopLeft
    ActiveSheet.Shapes.Range(Array("Name")).Select
    ActiveWorkbook.SlicerCaches("Slicer_Name").Slicers("Name").NumberOfColumns = 2
    ActiveSheet.Shapes("Name").IncrementLeft 547
    ActiveSheet.Shapes("Name").IncrementTop -166
    ActiveSheet.Shapes("Name").ScaleHeight 0.4496743134, msoFalse, _
        msoScaleFromTopLeft
    ActiveSheet.Shapes("Name").Width = 252
    ActiveSheet.Shapes.Range(Array("Product")).Select
    ActiveWorkbook.SlicerCaches("Slicer_Product").Slicers("Product"). _
        NumberOfColumns = 2
    ActiveSheet.Shapes("Product").IncrementLeft 585
    ActiveSheet.Shapes("Product").IncrementTop -42
    ActiveSheet.Shapes("Product").Width = 252
    ActiveSheet.Shapes("Product").ScaleHeight 0.4496743134, msoFalse, _
        msoScaleFromTopLeft
    ActiveSheet.Shapes.Range(Array("Date")).Select
    ActiveWorkbook.SlicerCaches("Slicer_Date").Slicers("Date").NumberOfColumns = 2
    ActiveSheet.Shapes("Date").Width = 252
    ActiveSheet.Shapes("Date").ScaleHeight 1.0052287711, msoFalse, _
        msoScaleFromBottomRight
    ActiveSheet.Shapes("Date").IncrementLeft 623
    ActiveSheet.Shapes("Date").IncrementTop 84
    ActiveSheet.Shapes("Date").ScaleHeight 0.6501966563, msoFalse, _
        msoScaleFromTopLeft
    ActiveWorkbook.SlicerCaches("Slicer_Date").Slicers("Date").Style = _
        "SlicerStyleDark6"
    ActiveSheet.Shapes.Range(Array("Product")).Select
    ActiveWorkbook.SlicerCaches("Slicer_Product").Slicers("Product").Style = _
        "SlicerStyleDark5"
    ActiveSheet.Shapes.Range(Array("Name")).Select
    ActiveWorkbook.SlicerCaches("Slicer_Name").Slicers("Name").Style = _
        "SlicerStyleDark4"
    ActiveSheet.Shapes.Range(Array("Area")).Select
    ActiveWorkbook.SlicerCaches("Slicer_Area").Slicers("Area").Style = _
        "SlicerStyleDark2"
    ActiveSheet.Shapes.Range(Array("Area", "Name")).Select
    ActiveSheet.Shapes.Range(Array("Area", "Name", "Product")).Select
    ActiveSheet.Shapes.Range(Array("Area", "Name", "Product", "Date")).Select
    Selection.Cut
    Sheets("Dashboard").Select
    Range("Q1").Select
    ActiveSheet.Paste
End Sub

Sub Delete()
'
' Delete Macro
'

'
    Sheets("Dashboard").Select
    ActiveSheet.ChartObjects("MyChart 1").Activate
    ActiveSheet.Shapes.SelectAll
    Selection.Delete
    Sheets(Array("RPT (4)", "RPT (3)", "RPT (2)", "RPT")).Select
    Sheets("RPT").Activate
    ActiveWindow.SelectedSheets.Delete
    Range("MyData[[#Headers],[Sr. No.]]").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Clear
    Range("A2").Select
End Sub

