Sub PetroHunt()
'
' PetroHunt Macro
' wsh FOR Loop to edit all tabs
 Dim wsh As Worksheet
    Application.ScreenUpdating = False
    For Each wsh In Worksheets
        wsh.Cells.MergeCells = False
        wsh.Cells.WrapText = False
        wsh.Range("A:B").EntireColumn.Insert
        wsh.Range("C3:D4").Copy Destination:=wsh.Range("A5:B6")
        wsh.Range("B6").EntireRow.Insert
        wsh.Range("A7:B60").FillDown
    Next wsh
    Application.ScreenUpdating = True
    
    'Create the Data sheet, adding a header, freeze pane
    Sheets.Add.Name = "Data"
    Sheets("Sheet1").Range("A5:J5").Copy Destination:=Sheets("Data").Range("A1")
    Range("B2").Select
    ActiveWindow.FreezePanes = True
    
    For Each wsh In ActiveWorkbook.Sheets
        If wsh.Name <> "Data" Then
           wsh.Activate
           wsh.Range("A7").CurrentRegion.Select
           Selection.Copy Destination:=Sheets("Data").Cells(Sheets("Data").Rows.Count, 1).End(xlUp)(2)
        End If
    Next wsh
End Sub
