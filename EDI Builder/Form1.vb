﻿Imports Excel = Microsoft.Office.Interop.Excel
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

            With sheet3
                .Range("A2").Value = "1"
                .Range("A3").Formula = "=IF(A2=30,1,A2+1)"
                .Range("J2").Value = "1"
                .Range("J3").Formula = "=IF(A3=1,J2+1,J2)"
            End With

            With sheet4
                .Range("A2").Value = "1"
                .Range("J2").Value = "1"
                .Range("J3").Formula = "=IF(A3=1,J2+1,J2)"
                .Range("A3").Formula = "=IF(A2=25,1,A2+1)"
            End With

            With sheet5
                .Range("A2").Value = "1"
                .Range("J2").Value = "1"
                .Range("J3").Formula = "=IF(A3=1,J2+1,J2)"
                .Range("A3").Formula = "=IF(A2=15,1,A2+1)"
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
            MsgBox(ex.ToString)
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
