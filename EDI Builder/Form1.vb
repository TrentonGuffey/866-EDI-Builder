Imports Excel = Microsoft.Office.Interop.Excel
Public Class Form1
    Private Sub BtnGo_Click(sender As Object, e As EventArgs) Handles btnGo.Click
        Dim theFiles As New List(Of String)
        Dim files() As String = IO.Directory.GetFiles(Application.StartupPath)

        For Each file As String In files

            'MsgBox(Mid(file, file.Length - 3, 4))

            If Mid(file, file.Length - 3, 4) = ".csv" Then

                ParseCSV(file)

            End If


        Next

        MsgBox("Finished")

    End Sub
    Private Sub SetPageBreaks(ByVal sheet As Excel.Worksheet, ByVal interval As Integer)
        ' Set page breaks on the specified sheet at the specified interval
        Dim rowCount As Integer = sheet.UsedRange.Rows.Count

        ' Clear existing horizontal page breaks
        sheet.ResetAllPageBreaks()
        sheet.PageSetup.PrintArea = ""

        ' Set print area and add horizontal page breaks at the specified interval
        If interval > 0 Then
            For i As Integer = interval + 2 To rowCount Step interval
                sheet.HPageBreaks.Add(sheet.Cells(i, 1))
            Next
        End If

        ' Set the print area to cover the entire used range
        sheet.PageSetup.PrintArea = sheet.UsedRange.Address
    End Sub
    Private Sub ParseCSV(ByVal theFileName As String)

        Dim outFileName As String = ""
        outFileName = theFileName
        outFileName = Mid(theFileName, 1, theFileName.Length - 4)
#Disable Warning IDE0054 ' Use compound assignment
        outFileName = outFileName & ".csv"
#Enable Warning IDE0054 ' Use compound assignment

    Dim excelApp As New Excel.Application
        excelApp.DisplayAlerts = False
    Dim excelBook As Excel.Workbook = excelApp.Workbooks.Add(System.Reflection.Missing.Value)
    Dim theSheet As Excel.Worksheet = Nothing
    'try to open the workbook and a worksheet
    Try
            'MsgBox(outFileName)
            excelBook = excelApp.Workbooks.Open(outFileName)
            theSheet = CType(excelBook.Sheets(1), Microsoft.Office.Interop.Excel.Worksheet)

    Dim beginningLineSet As String = txtBeginningLineSet.Text.Trim()
    Dim thruDate As String = txtThruDate.Text.Trim()

    Dim ngvRackSizeInput As String = txtNGVRackSize.Text.Trim()
    Dim frontRackSizeInput As String = txtFrontRackSize.Text.Trim()
    Dim rearRackSizeInput As String = txtRearRackSize.Text.Trim()

    Dim ngvRackSize As Integer = If(String.IsNullOrEmpty(ngvRackSizeInput), 25, Integer.Parse(ngvRackSizeInput))
    Dim frontRackSize As Integer = If(String.IsNullOrEmpty(frontRackSizeInput), 15, Integer.Parse(frontRackSizeInput))
    Dim rearRackSize As Integer = If(String.IsNullOrEmpty(rearRackSizeInput), 15, Integer.Parse(rearRackSizeInput))

    ' Check If the user provided a value
    If Not String.IsNullOrEmpty(beginningLineSet) Then
                ' Delete rows before the specified "Begining LineSet" value
                DeleteRowsAfterValue(theSheet, beginningLineSet, "B")
    End If

    ' Check if the user provided a value
    If Not String.IsNullOrEmpty(thruDate) Then
                ' Delete rows after the specified "Thru Date" value
                DeleteRowsBeforeDate(theSheet, thruDate, "G")
    End If

            theSheet.Range("A1:I1").EntireColumn.AutoFit()
            theSheet.Range("A:K").HorizontalAlignment = Excel.Constants.xlCenter
            theSheet.PageSetup.PrintTitleRows = "$1:$1"

    Dim sheet2 As Excel.Worksheet
            sheet2 = CType(excelBook.Worksheets.Add(After:=theSheet), Excel.Worksheet)
            sheet2.Name = "P-JOBS"

    With sheet2
                .Range("A1").Value = "Status"
                .Range("B1").Value = "Ship Date"
                .Range("C1").Value = "ASN"
                .Range("D1").Value = "Job Sequence"
                .Range("E1").Value = "Set Number"
                .Range("F1").Value = "Part Number"
                .Range("G1").Value = "Job Number"
                .Range("H1").Value = "Product Code"
                .Range("I1").Value = "Product Name"
                .Range("J1").Value = "Date"
                .Range("K1").Value = "ISAControl"
            End With


    Dim range1 As Excel.Range
            range1 = theSheet.UsedRange

    With range1
    .AutoFilter(Field:=2, Criteria1:="=*P*")
                range1.Offset(1, 0).Copy()
                sheet2.Range("D2").PasteSpecial()

    If theSheet.AutoFilterMode = True Then
                    theSheet.AutoFilterMode = False
    End If
    End With

            sheet2.Range("A:K").HorizontalAlignment = Excel.Constants.xlCenter
            sheet2.Range("A1:K1").WrapText = True
            sheet2.Range("A:K").EntireColumn.AutoFit()
            sheet2.PageSetup.PrintTitleRows = "$1:$1"

    Dim sheet3 As Excel.Worksheet
            sheet3 = CType(excelBook.Worksheets.Add(After:=sheet2), Excel.Worksheet)
            sheet3.Name = "NGV"

    With sheet3
        .Range("A1").Value = "Number"
        .Range("B1").Value = "JobSequence"
        .Range("C1").Value = "SetNumber"
        .Range("D1").Value = "PartNumber"
        .Range("E1").Value = "JobNumber"
        .Range("F1").Value = "ProductCode"
        .Range("G1").Value = "ProductName"
        .Range("H1").Value = "Date"
        .Range("I1").Value = "ISAControl"
        .Range("J1").Value = "Rack"
    End With
    With range1
    .AutoFilter(Field:=6, Criteria1:="NGV")
                range1.Offset(1, 0).Copy()
                sheet3.Range("B2").PasteSpecial()
    If theSheet.AutoFilterMode = True Then
                    theSheet.AutoFilterMode = False
    End If
    End With

            sheet3.Range("A1:J1").EntireColumn.AutoFit()
            sheet3.Range("A:J").HorizontalAlignment = Excel.Constants.xlCenter
            sheet3.PageSetup.PrintTitleRows = "$1:$1"
            SetPageBreaks(sheet3, ngvRackSize)


    Dim sheet4 As Excel.Worksheet
            sheet4 = CType(excelBook.Worksheets.Add(After:=sheet3), Excel.Worksheet)
            sheet4.Name = "LT_FRONT"
            sheet3.Range("A1:J1").Copy()
            sheet4.Paste()

    With range1
    .AutoFilter(Field:=6, Criteria1:="LT FRONT")
                range1.Offset(1, 0).Copy()
                sheet4.Range("B2").PasteSpecial()
    If theSheet.AutoFilterMode = True Then
                    theSheet.AutoFilterMode = False
    End If
    End With

            sheet4.Range("A1:J1").EntireColumn.AutoFit()
            sheet4.Range("A:J").HorizontalAlignment = Excel.Constants.xlCenter
            sheet4.PageSetup.PrintTitleRows = "$1:$1"
            SetPageBreaks(sheet4, frontRackSize)


    Dim sheet5 As Excel.Worksheet
            sheet5 = CType(excelBook.Worksheets.Add(After:=sheet4), Excel.Worksheet)
            sheet5.Name = "LT_REAR"
            sheet3.Range("A1:J1").Copy()
            sheet5.Paste()

    With range1
    .AutoFilter(Field:=6, Criteria1:="LT REAR")
                range1.Offset(1, 0).Copy()
                sheet5.Range("B2").PasteSpecial()
    If theSheet.AutoFilterMode = True Then
                    theSheet.AutoFilterMode = False
    End If
    End With


            sheet5.Range("A1:J1").EntireColumn.AutoFit()
            sheet5.Range("A:J").HorizontalAlignment = Excel.Constants.xlCenter
            sheet5.PageSetup.PrintTitleRows = "$1:$1"
            SetPageBreaks(sheet5, rearRackSize)

    Dim sheet6 As Excel.Worksheet
            sheet6 = CType(excelBook.Worksheets.Add(After:=sheet5), Excel.Worksheet)

            sheet6.Name = "TOTALS"

    Dim theValue As String = "value"
    Dim incR As Integer = 1
    Do Until theValue = ""
                theValue = theSheet.Range("A" & incR).Value
                incR += 1
    Loop
#Disable Warning IDE0054 ' Use compound assignment
            incR = incR - 2
#Enable Warning IDE0054 ' Use compound assignment

    Dim lastRow As Long = theSheet.Cells.Rows.Count
    Dim lastCol As Long = theSheet.Cells.Columns.Count

    Dim dataRange As Excel.Range = theSheet.Range("A1:H" & incR)
    Dim ptName As String = "MyPivotTable"
            sheet6.PivotTableWizard(Excel.XlPivotTableSourceType.xlDatabase, dataRange, sheet6.Cells(3, 1), ptName)
            sheet6.Select()
    Dim pt As Excel.PivotTable = sheet6.PivotTables(1)
    With pt
    .TableStyle = Excel.XlPivotTableVersionList.xlPivotTableVersion14
    .TableStyle2 = "PivotStyleLight16"
    .InGridDropZones = False
    End With
    With pt.PivotFields("PartNumber")
    .Orientation = Excel.XlPivotFieldOrientation.xlRowField
    .Position = 1
    End With
    With pt.PivotFields("PartNumber")
    .Orientation = Excel.XlPivotFieldOrientation.xlDataField
    .Position = 1
    End With

    'Formatting the first Sheets header and margins
    With theSheet
        .PageSetup.LeftHeader = "NAVISTAR"
        .PageSetup.CenterHeader = "SHIP DATE:"
        .PageSetup.RightHeader = "WJ FILE:"
        .PageSetup.LeftMargin = "18"
        .PageSetup.RightMargin = "18"
        .PageSetup.TopMargin = "72"
        .PageSetup.BottomMargin = "54"
        .PageSetup.HeaderMargin = "21.6"
        .PageSetup.FooterMargin = "21.6"
    End With

    With sheet2
        .PageSetup.LeftHeader = "NAVISTAR"
        .PageSetup.CenterHeader = "SHIP DATE:"
        .PageSetup.RightHeader = "WJ FILE:"
        .PageSetup.LeftMargin = "18"
        .PageSetup.RightMargin = "18"
        .PageSetup.TopMargin = "72"
        .PageSetup.BottomMargin = "54"
        .PageSetup.HeaderMargin = "21.6"
        .PageSetup.FooterMargin = "21.6"
        .PageSetup.FitToPagesWide = 1
        .PageSetup.FitToPagesTall = False
        .PageSetup.Zoom = False
    End With

    With sheet3
        .PageSetup.LeftHeader = "NAVISTAR-" & Chr(13) & "NGV" & Chr(13) & "RACK &P OF &N"
        .PageSetup.CenterHeader = "SHIP DATE:"
        .PageSetup.RightHeader = "WJ FILE:"
        .PageSetup.LeftMargin = "18"
        .PageSetup.RightMargin = "18"
        .PageSetup.TopMargin = "72"
        .PageSetup.BottomMargin = "54"
        .PageSetup.HeaderMargin = "21.6"
        .PageSetup.FooterMargin = "21.6"
        .PageSetup.FitToPagesWide = 1
        .PageSetup.FitToPagesTall = False
        .PageSetup.Zoom = False
    End With

    With sheet4
        .PageSetup.LeftHeader = "NAVISTAR-" & Chr(13) & "LT FRONT" & Chr(13) & "RACK &P OF &N"
        .PageSetup.CenterHeader = "SHIP DATE:"
        .PageSetup.RightHeader = "WJ FILE:"
        .PageSetup.LeftMargin = "18"
        .PageSetup.RightMargin = "18"
        .PageSetup.TopMargin = "72"
        .PageSetup.BottomMargin = "54"
        .PageSetup.HeaderMargin = "21.6"
        .PageSetup.FooterMargin = "21.6"
        .PageSetup.FitToPagesWide = 1
        .PageSetup.FitToPagesTall = False
        .PageSetup.Zoom = False
    End With

    With sheet5
        .PageSetup.LeftHeader = "NAVISTAR-" & Chr(13) & "LT REAR" & Chr(13) & "RACK &P OF &N"
        .PageSetup.CenterHeader = "SHIP DATE:"
        .PageSetup.RightHeader = "WJ FILE:"
        .PageSetup.LeftMargin = "18"
        .PageSetup.RightMargin = "18"
        .PageSetup.TopMargin = "72"
        .PageSetup.BottomMargin = "54"
        .PageSetup.HeaderMargin = "21.6"
        .PageSetup.FooterMargin = "21.6"
        .PageSetup.FitToPagesWide = 1
        .PageSetup.FitToPagesTall = False
        .PageSetup.Zoom = False
    End With

    With sheet6
        .PageSetup.LeftHeader = "NAVISTAR"
        .PageSetup.CenterHeader = "SHIP DATE:"
        .PageSetup.RightHeader = "WJ FILE:"
        .PageSetup.LeftMargin = "18"
        .PageSetup.RightMargin = "18"
        .PageSetup.TopMargin = "54"
        .PageSetup.BottomMargin = "54"
        .PageSetup.HeaderMargin = "21.6"
        .PageSetup.FooterMargin = "21.6"
        .PageSetup.FitToPagesWide = 1
        .PageSetup.FitToPagesTall = False
        .PageSetup.Zoom = False
    End With

    Dim range2 As Excel.Range
        range2 = sheet6.UsedRange
        range2.BorderAround(Excel.XlLineStyle.xlContinuous, Excel.XlBorderWeight.xlThin, Excel.XlColorIndex.xlColorIndexAutomatic, Excel.XlColorIndex.xlColorIndexAutomatic)
        range2.Borders(Excel.XlBordersIndex.xlInsideHorizontal).Color = 2
        range2.Borders(Excel.XlBordersIndex.xlInsideVertical).Color = 2

        sheet6.Range("A3").Value = "PN"
        sheet6.Range("B3").Value = "COUNT"

    Dim lastRowSheet3 As Long = sheet3.UsedRange.Rows.Count
    With sheet3
        .Range("A2").Value = "1"
        ' Apply the formula from A3 down to the last row
        For rowNum As Integer = 3 To lastRowSheet3
            .Range("A" & rowNum).Formula = "=IF(A" & rowNum - 1 & "=30,1,A" & rowNum - 1 & "+1)"
        Next
        .Range("J2").Value = "1"
        ' Apply the formula from J3 down to the last row
        For rowNum As Integer = 3 To lastRowSheet3
            .Range("J" & rowNum).Formula = "=IF(A" & rowNum & "=1,J" & rowNum - 1 & "+1,J" & rowNum - 1 & ")"
        Next
    End With

    Dim lastRowSheet4 As Long = sheet4.UsedRange.Rows.Count
    With sheet4
        .Range("A2").Value = "1"
        ' Apply the formula from A3 down to the last row
        For rowNum As Integer = 3 To lastRowSheet4
            .Range("A" & rowNum).Formula = "=IF(A" & rowNum - 1 & "=30,1,A" & rowNum - 1 & "+1)"
        Next
        .Range("J2").Value = "1"
        ' Apply the formula from J3 down to the last row
        For rowNum As Integer = 3 To lastRowSheet4
            .Range("J" & rowNum).Formula = "=IF(A" & rowNum & "=1,J" & rowNum - 1 & "+1,J" & rowNum - 1 & ")"
        Next
    End With
    Dim lastRowSheet5 As Long = sheet5.UsedRange.Rows.Count
    With sheet3
        .Range("A2").Value = "1"
        ' Apply the formula from A3 down to the last row
        For rowNum As Integer = 3 To lastRowSheet5
            .Range("A" & rowNum).Formula = "=IF(A" & rowNum - 1 & "=30,1,A" & rowNum - 1 & "+1)"
        Next
        .Range("J2").Value = "1"
        ' Apply the formula from J3 down to the last row
        For rowNum As Integer = 3 To lastRowSheet5
            .Range("J" & rowNum).Formula = "=IF(A" & rowNum & "=1,J" & rowNum - 1 & "+1,J" & rowNum - 1 & ")"
        Next
    End With

    Dim lrow6 As Integer
            lrow6 = sheet6.UsedRange.Rows.Count

    With sheet6
        .Range("A" & lrow6 + 2).Value = "TOTALS"
        .Range("A" & lrow6 + 5).Value = "NGV RACKS"
        .Range("A" & lrow6 + 6).Value = "LT FRONT RACKS"
        .Range("A" & lrow6 + 7).Value = "LT REAR RACKS"
        .Range("A" & lrow6 + 9).Value = "LINE SET:"
        .Range("A" & lrow6 + 5).Font.Bold = True
        .Range("A" & lrow6 + 6).Font.Bold = True
        .Range("A" & lrow6 + 7).Font.Bold = True
        .Range("A" & lrow6 + 9).Font.Bold = True
        .Range("A1:B1").EntireColumn.AutoFit()
    End With

    Catch ex As Exception
            ' Log the exception message and stack trace to the console
            Console.WriteLine("Exception: " & ex.Message)
            Console.WriteLine("Stack Trace: " & ex.StackTrace)

            ' Display a message box with the exception details
            MsgBox("An error occurred. Check the console for details.")
    Finally
            excelBook.SaveAs(outFileName.Replace("csv", "xlsx"), FileFormat:=Excel.XlFileFormat.xlOpenXMLWorkbook)
            'MAKE SURE TO KILL ALL INSTANCES BEFORE QUITING! if you fail to do this
    'The service (excel.exe) will continue to run
            NAR(theSheet)
            excelBook.Close(False)
            NAR(excelBook)
            excelApp.Workbooks.Close()
            NAR(excelApp.Workbooks)
            'quit and dispose app
            excelApp.Quit()
            NAR(excelApp)
            'VERY IMPORTANT
            GC.Collect()
            My.Computer.FileSystem.DeleteFile(outFileName)
    End Try
    End Sub

    Private Sub DeleteRowsAfterValue(ByVal sheet As Excel.Worksheet, ByVal targetValue As String, ByVal column As String)
        Dim usedRange As Excel.Range = sheet.UsedRange
        Dim lastRow As Long = usedRange.Rows.Count

        ' Start from the second row (row 2) since we want to keep the first row
        For i As Long = 2 To lastRow
            Dim cellValue As String = sheet.Range(column & i).Value

            ' Compare the cell value with the target value
            If String.Compare(cellValue, targetValue, StringComparison.OrdinalIgnoreCase) = 0 Then
                ' Exit the loop when the target value is found
                Exit For
            Else
                ' Delete the current row if the target value is not found yet
                sheet.Rows(i).Delete()
                ' Adjust the loop counter since the row has been deleted
                i = i - 1
                ' Adjust the last row count since a row has been deleted
                lastRow = lastRow - 1
            End If
        Next
    End Sub

    Private Sub DeleteRowsBeforeDate(ByVal sheet As Excel.Worksheet, ByVal targetValue As String, ByVal column As String)
        Dim usedRange As Excel.Range = sheet.UsedRange
        Dim lastRow As Long = usedRange.Rows.Count
        For i As Long = lastRow To 1 Step -1
            Dim cellValue As String = sheet.Range(column & i).Value

            ' Compare the cell value with the target value
            If String.Compare(cellValue, targetValue, StringComparison.OrdinalIgnoreCase) = 0 Then
                ' Exit the loop when the target value is found
                Exit For
            Else
                ' Delete the current row if the target value is not found yet
                sheet.Rows(i).Delete()
            End If
        Next
    End Sub
    Private Function Cells(lngRow As Long, lngCol As Long) As Object
        Throw New NotImplementedException()
    End Function

    Private Sub NAR(ByVal o As Object)
        Try
            System.Runtime.InteropServices.Marshal.ReleaseComObject(o)
        Catch ex As Exception
        Finally
#Disable Warning IDE0059 ' Unnecessary assignment of a value
            o = Nothing
#Enable Warning IDE0059 ' Unnecessary assignment of a value
        End Try

    End Sub

End Class
