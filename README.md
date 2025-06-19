# ExcelModule
```
Sub HighlightMissingFields()
    Dim ws As Worksheet
    Dim lastRow As Long, lastCol As Long
    Dim cell As Range
    Dim outputCol As Long
    Dim missingList As String
    Dim i As Long, j As Long
    Dim header As String
    Dim isComplete As Boolean

    Set ws = ActiveSheet
    lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
    lastCol = ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column
    outputCol = lastCol + 1 ' Output column for result (Missing Fields)

    ' Set output header
    ws.Cells(1, outputCol).Value = "Missing Fields"

    ' Loop through each row of data
    For i = 2 To lastRow
        missingList = ""
        isComplete = True ' Assume complete unless found missing

        ' Check each cell in the row
        For j = 1 To lastCol
            Set cell = ws.Cells(i, j)
            header = ws.Cells(1, j).Value

            ' Clear previous highlights
            cell.Interior.ColorIndex = xlNone

            ' Check for blank/space-only cells
            If Trim(cell.Value) = "" Then
                cell.Interior.Color = RGB(255, 0, 0) ' Light red
                missingList = missingList & header & ", "
                isComplete = False
            End If
        Next j

        ' Output and highlight
        With ws.Cells(i, outputCol)
            .Interior.ColorIndex = xlNone ' Clear any previous color
            If isComplete Then
                .Value = "All fields are filled."
                .Interior.Color = RGB(17, 255, 0) ' Light green
            Else
                .Value = "Missing: " & Left(missingList, Len(missingList) - 2)
            End If
        End With
    Next i
End Sub
```
