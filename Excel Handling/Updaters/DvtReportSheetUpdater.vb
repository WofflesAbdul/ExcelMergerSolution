Imports System.Runtime.InteropServices
Imports System.Windows.Forms
Imports Microsoft.Office.Interop

' ===============================
' Core partial class – helpers
' ===============================
Partial Public Class DvtReportSheetUpdater

    ' ===============================
    ' Helpers (shared across sheets)
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

    Private Sub ReleaseComObject(obj As Object)
        Try
            If obj IsNot Nothing Then Marshal.ReleaseComObject(obj)
        Finally
            obj = Nothing
        End Try
    End Sub

End Class