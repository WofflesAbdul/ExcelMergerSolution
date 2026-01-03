Imports Microsoft.Office.Interop

Public Class WorkbookNamedRangeCollector

    Public Function Collect(filePath As String) As WorkbookNamedData
        Dim excelApp As Excel.Application = Nothing
        Dim wb As Excel.Workbook = Nothing

        Dim reportWb As New WorkbookNamedData()

        Try
            excelApp = New Excel.Application()
            wb = excelApp.Workbooks.Open(filePath)

            For Each ws As Excel.Worksheet In wb.Sheets

                If ws.Name.Equals("Cover Page", StringComparison.OrdinalIgnoreCase) Then
                    reportWb.CoverPage = ReadCoverPageSheet(ws)
                Else
                    Dim info = ReadTestSheet(ws)
                    If info IsNot Nothing Then
                        reportWb.TestSheets.Add(info)
                    End If
                End If

            Next

            Return reportWb

        Finally
            If wb IsNot Nothing Then wb.Close(SaveChanges:=False)
            If excelApp IsNot Nothing Then excelApp.Quit()
            ReleaseComObject(wb)
            ReleaseComObject(excelApp)
        End Try
    End Function

    Public Function CollectFromOpenWorkbook(wb As Excel.Workbook) As WorkbookNamedData
        Dim reportWb As New WorkbookNamedData()

        For Each ws As Excel.Worksheet In wb.Sheets
            If ws.Name.Equals("Cover Page", StringComparison.OrdinalIgnoreCase) Then
                reportWb.CoverPage = ReadCoverPageSheet(ws)
            Else
                Dim info = ReadTestSheet(ws)
                If info IsNot Nothing Then
                    reportWb.TestSheets.Add(info)
                End If
            End If
        Next

        Return reportWb
    End Function

    Private Function ReadCoverPageSheet(ws As Excel.Worksheet) As CoverPageSheetInfo
        Return New CoverPageSheetInfo With {
            .PowerSupplyModel = ReadNamed(ws, "PowerSupplyModel"),
            .PowerSupplySerialNumber = ReadNamed(ws, "PowerSupplySerialNumber"),
            .PowerSupplyFirmwareVersion = ReadNamed(ws, "PowerSupplyFirmwareVersion"),
            .TesterName = ReadNamed(ws, "Tester"),
            .Description = ReadNamed(ws, "Description")
        }
    End Function

    Private Function ReadTestSheet(ws As Excel.Worksheet) As TestReportSheetInfo
        ' Only create if at least one named range exists
        Dim devPhase = ReadNamed(ws, "DevelopmentPhase")
        Dim fw = ReadNamed(ws, "FirmwareVersion")
        Dim model = ReadNamed(ws, "ModelNumber")
        Dim serial = ReadNamed(ws, "SerialNumber")
        Dim testedBy = ReadNamed(ws, "TestedBy")

        If String.IsNullOrWhiteSpace(devPhase) AndAlso
           String.IsNullOrWhiteSpace(fw) AndAlso
           String.IsNullOrWhiteSpace(model) AndAlso
           String.IsNullOrWhiteSpace(serial) AndAlso
           String.IsNullOrWhiteSpace(testedBy) Then
            Return Nothing
        End If

        Return New TestReportSheetInfo With {
            .WorksheetName = ws.Name,
            .DevelopmentPhase = devPhase,
            .FirmwareVersion = fw,
            .ModelNumber = model,
            .SerialNumber = serial,
            .TestedBy = testedBy
        }
    End Function

    Private Function ReadNamed(ws As Excel.Worksheet, name As String) As String
        Try
            Dim nm = ws.Names.Item(name)
            Dim rng = DirectCast(nm.RefersToRange, Excel.Range)
            Return If(rng.Value2 IsNot Nothing, rng.Value2.ToString(), Nothing)
        Catch
            Return Nothing
        End Try
    End Function

    Private Sub ReleaseComObject(obj As Object)
        Try
            If obj IsNot Nothing Then
                System.Runtime.InteropServices.Marshal.ReleaseComObject(obj)
            End If
        Finally
            obj = Nothing
        End Try
    End Sub
End Class
