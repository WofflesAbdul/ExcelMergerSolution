Imports System.Runtime.InteropServices
Imports System.Windows.Forms
Imports Microsoft.Office.Interop

Public Class TitleSheetUpdater

    Public Sub UpdateTitleSheet(filePath As String, values As ResolvedTestMetadata)
        Dim excelApp As Excel.Application = Nothing
        Dim wb As Excel.Workbook = Nothing

        Try
            excelApp = New Excel.Application()
            wb = excelApp.Workbooks.Open(filePath)

            UpdateTitleSheetFromOpenWorkbook(wb, values)

        Finally
            If wb IsNot Nothing Then wb.Close(SaveChanges:=True)
            If excelApp IsNot Nothing Then excelApp.Quit()
            ReleaseComObject(wb)
            ReleaseComObject(excelApp)
        End Try
    End Sub

    Public Sub UpdateTitleSheetFromOpenWorkbook(wb As Excel.Workbook, values As ResolvedTestMetadata)
        Dim titleSheet As Excel.Worksheet = Nothing

        ' ---- Locate Title sheet ----
        For Each ws As Excel.Worksheet In wb.Sheets
            If ws.Name.Equals("Title", StringComparison.OrdinalIgnoreCase) Then
                titleSheet = ws
                Exit For
            End If
        Next

        If titleSheet Is Nothing Then Return

        ' ---- Named fields ----
        WriteNamed(titleSheet, "PowerSupplyModel", values.ModelNumber)
        WriteNamed(titleSheet, "PowerSupplySerialNumber", values.SerialNumber)
        WriteNamed(titleSheet, "PowerSupplyFirmwareVersion", values.FirmwareVersion)

        ' ---- Defaults for dialog ----
        Dim suggestedDescription = values.DevelopmentPhase
        Dim suggestedEngineer = values.TestedBy

        ' ---- Read latest Rev from table (if any) ----
        Dim totalRows As Integer
        Dim revList As List(Of String)
        ReadTableColumnValues(titleSheet, "DvtReportOverviewTable", "Rev", totalRows, revList)

        ' Safely get latest revision
        Dim latestRev As String = Nothing
        If revList IsNot Nothing AndAlso revList.Count > 0 Then
            latestRev = revList.Last()
        End If

        ' ---- Auto-increment revision ----
        Dim suggestedRev As String = IncrementRevision(latestRev)

        ' ---- Prompt user (blocks safely) ----
        Using dlg As New TitleTablePromptDialog(suggestedDescription, suggestedEngineer, suggestedRev)
            If dlg.ShowDialog() <> DialogResult.OK Then Return

            ' ---- Build table row values ----
            Dim tableValues As New Dictionary(Of String, Object) From {
            {"Rev", dlg.Revision},
            {"Engineer", dlg.Engineer},
            {"Description", dlg.Description},
            {"Date Prepared", Date.Today}
        }

            WriteToTable(titleSheet, "DvtReportOverviewTable", tableValues)
        End Using

        wb.Save()
    End Sub

    ' ===============================
    ' Helpers
    ' ===============================

    Private Sub WriteNamed(ws As Excel.Worksheet, name As String, value As String)
        If String.IsNullOrWhiteSpace(value) Then Return

        Try
            Dim nm = ws.Names.Item(name)
            Dim rng = DirectCast(nm.RefersToRange, Excel.Range)
            rng.Value2 = value
        Catch
            ' Named range missing → ignore
        End Try
    End Sub

    Private Sub WriteToTable(ws As Excel.Worksheet, tableName As String, values As IDictionary(Of String, Object))
        If values Is Nothing OrElse values.Count = 0 Then Return

        Try
            Dim tbl As Excel.ListObject = ws.ListObjects(tableName)
            Dim isEmpty As Boolean = (tbl.ListRows.Count = 0)
            Dim targetRow As Excel.ListRow = tbl.ListRows.Add() ' always append

            ' ---- Remove phantom row if table was empty ----
            If isEmpty Then
                ' The first added row is now targetRow
                ' The phantom row is automatically shifted down to row 2
                If tbl.ListRows.Count > 1 Then
                    tbl.ListRows(2).Delete()
                End If
            End If

            ' ---- Write column values ----
            For Each kvp In values
                If kvp.Value Is Nothing Then Continue For
                If TypeOf kvp.Value Is String AndAlso String.IsNullOrWhiteSpace(kvp.Value.ToString()) Then Continue For

                Try
                    Dim colIndex = tbl.ListColumns(kvp.Key).Index
                    targetRow.Range.Cells(1, colIndex).Value2 = kvp.Value
                Catch
                    ' Column missing → ignore
                End Try
            Next

        Catch
            ' Table missing → ignore
        End Try
    End Sub

    Private Sub ReadTableColumnValues(ws As Excel.Worksheet, tableName As String, columnName As String,
                                  ByRef totalRows As Integer, ByRef columnValues As List(Of String))
        totalRows = 0
        columnValues = New List(Of String)()

        Try
            Dim tbl As Excel.ListObject = ws.ListObjects(tableName)
            totalRows = tbl.ListRows.Count

            If totalRows = 0 Then Return

            Dim colIndex As Integer = tbl.ListColumns(columnName).Index

            For i As Integer = 1 To totalRows
                Dim val = tbl.ListRows(i).Range.Cells(1, colIndex).Value2
                columnValues.Add(If(val?.ToString(), String.Empty))
            Next

        Catch ex As Exception
            ' Optional: log or handle error
            totalRows = 0
            columnValues.Clear()
        End Try
    End Sub

    Private Function IncrementRevision(latestRev As String) As String
        If String.IsNullOrWhiteSpace(latestRev) Then Return "A"

        ' Single letter
        If latestRev.Length = 1 AndAlso Char.IsLetter(latestRev(0)) Then
            Dim nextChar As Char = Chr(Asc(latestRev(0)) + 1)
            If nextChar > "Z"c Then nextChar = "A"c ' wrap around if needed
            Return nextChar.ToString()
        End If

        ' Try integer
        Dim intVal As Integer
        If Integer.TryParse(latestRev, intVal) Then
            Return (intVal + 1).ToString()
        End If

        ' Try dot-separated numeric revision (e.g., "1.2.3")
        Dim parts() As String = latestRev.Split("."c)
        Dim allNumbers As Boolean = True

        For Each p In parts
            If Not Integer.TryParse(p, 0) Then
                allNumbers = False
                Exit For
            End If
        Next

        If allNumbers Then
            ' Increment the rightmost number
            parts(parts.Length - 1) = (CInt(parts(parts.Length - 1)) + 1).ToString()
            Return String.Join(".", parts)
        End If

        ' Fallback: append "1"
        Return latestRev & ".1"
    End Function

    Private Sub ReleaseComObject(obj As Object)
        Try
            If obj IsNot Nothing Then Marshal.ReleaseComObject(obj)
        Finally
            obj = Nothing
        End Try
    End Sub
End Class