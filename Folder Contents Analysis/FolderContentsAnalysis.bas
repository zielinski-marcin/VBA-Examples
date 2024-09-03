Attribute VB_Name = "FolderContentsAnalysis"
Sub FolderContentsAnalysis()

'   Be mindful that by default this clears your currently active sheet
'   Made in Excel 2016 (xlTreemap was introduced in 2016 version, so most likely it's not backwards compatible)

Dim path As String
Dim filename As String

Dim rowList As Long     ' used for the list of files
Dim rowSum As Long      ' used for the summary section
Dim rowEnd As Long      ' used for the summary section (total number of rows)

'   Select folder to analyze:

With Application.FileDialog(msoFileDialogFolderPicker)
 
        .InitialFileName = Application.DefaultFilePath
        .Title = "Select the folder to analyze..."
        .Show
        If .SelectedItems.Count = 0 Then Exit Sub
 
        path = .SelectedItems(1) & "\"
 
End With

With ThisWorkbook.ActiveSheet

        '   clear the worksheet, including any charts:
        
        .Cells.Clear
        If .ChartObjects.Count > 0 Then
            .ChartObjects.Delete
        End If

        '   summary headers:
        
        .Cells(1, 1).Value = "id"
        .Cells(1, 2).Value = "catalog"
        .Cells(1, 3).Value = "name"
        .Cells(1, 4).Value = "format"
        .Cells(1, 5).Value = "size_MB"
        .Cells(1, 6).Value = "last_modified"
    
        '   formatting:
        
        .Columns("A:Q").ColumnWidth = 3.5
        .Rows(1).Font.FontStyle = "Bold"
        With .Cells.Interior
                .Pattern = xlSolid
                .PatternColorIndex = xlAutomatic
                .ThemeColor = xlThemeColorDark1
                .TintAndShade = 0
                .PatternTintAndShade = 0
        End With

        '   list files in the folder:

        filename = Dir(path, vbReadOnly + vbHidden + vbSystem)
        
        rowList = 2
        
        Do Until filename = ""
                file = path & filename
                Cells(rowList, 1) = rowList - 1
                Cells(rowList, 2) = path
                Cells(rowList, 3) = filename
                Cells(rowList, 4) = Right(filename, Len(filename) - InStrRev(filename, "."))
                Cells(rowList, 5) = FileLen(path & filename) / 1024 ^ 2   'size converted to MB
                Cells(rowList, 6) = FileDateTime(path & filename)
                filename = Dir
                rowList = rowList + 1
        Loop
        
        answer = MsgBox("Would you like to create a summary?", vbQuestion + vbYesNo + vbDefaultButton2, "Statistics?")
        
        If answer = vbNo Then
            .Range("A1:F" & .Cells(Rows.Count, 1).End(xlUp).Row).Borders.LineStyle = xlContinuous
            .Columns("A:F").AutoFit
            Exit Sub
        End If
        
        .Cells(1, 10).Value = "total_size_MB"
        .Cells(1, 11).Value = "files_found"
        .Cells(1, 14).Value = "oldest_file"
        .Cells(1, 15).Value = "newest_file"
        .Cells(1, 16).Value = "total_size_MB"
        .Cells(1, 17).Value = "total_qty"

        Range("D1:D" & rowList).AdvancedFilter Action:=xlFilterCopy, CriteriaRange:=ActiveCell.Range("D1:D" & rowList), CopyToRange:=Range("I1"), Unique:=True 'list distinct file types
        rowEnd = .Cells(Rows.Count, 9).End(xlUp).Row + 1

        '   calculate total size and quantity for each file type:
        
        rowSum = 2
        Do Until rowSum = rowEnd
            .Cells(rowSum, 10).Value = Application.WorksheetFunction.SumIf(Range("d:d"), .Cells(rowSum, 9).Value, Range("E:E"))
            .Cells(rowSum, 11).Value = Application.WorksheetFunction.CountIf(Range("d:d"), .Cells(rowSum, 9).Value)
            rowSum = rowSum + 1
        Loop
        
        '   calculate total size, number of files, oldest and newest file:
        
        .Cells(2, 14).Value = Application.WorksheetFunction.Min(Range("F:F"))   'oldest file
        .Cells(2, 15).Value = Application.WorksheetFunction.Max(Range("F:F"))   'newest file
        .Cells(2, 16).Value = Application.WorksheetFunction.Sum(Range("J:J"))   'total_size
        .Cells(2, 17).Value = rowEnd - 1                                        'total number of files

        '   formatting:
        
        .Range("N2:O2").NumberFormat = "yyyy-mm-dd hh:mm"
        .Columns("E").NumberFormat = "0.0000"
        .Columns("J").NumberFormat = "0.0000"
        .Columns("P").NumberFormat = "0.0000"
        .Range("A1:F" & .Cells(Rows.Count, 1).End(xlUp).Row).Borders.LineStyle = xlContinuous
        .Range("I1:K" & .Cells(Rows.Count, 9).End(xlUp).Row).Borders.LineStyle = xlContinuous
        .Range("N1:Q2").Borders.LineStyle = xlContinuous
        .Rows(1).Font.FontStyle = "Bold"
        With .Cells.Interior
                .Pattern = xlSolid
                .PatternColorIndex = xlAutomatic
                .ThemeColor = xlThemeColorDark1
                .TintAndShade = 0
                .PatternTintAndShade = 0
        End With

        '   set columns widths:
        
        .Columns("A:Q").ColumnWidth = 3.5
        .Columns("A:Q").AutoFit
        If .Columns(2).ColumnWidth > 70 Then
            .Columns(2).ColumnWidth = 70
        End If
        If .Columns(3).ColumnWidth > 70 Then
        .Columns(3).ColumnWidth = 70
        End If

        '   create Treemap chart of file sizes:
        
        .Range("I1:J" & rowEnd - 1).Select
        With ActiveSheet.Shapes.AddChart2(410, xlTreemap)
            .Chart.ChartTitle.Text = "Size per file extension"
        End With
 
End With

End Sub
