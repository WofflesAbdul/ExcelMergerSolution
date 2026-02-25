' ===============================
' Summary Page partial class
' ===============================
Imports Microsoft.Office.Interop

Partial Public Class DvtReportSheetUpdater

    Public Sub UpdateSummarySheet(wb As Excel.Workbook)

        Dim summaryWorkSheet As Excel.Worksheet = Nothing

        For Each ws As Excel.Worksheet In wb.Sheets
            If ws.Name.Equals("Summary", StringComparison.OrdinalIgnoreCase) Then
                summaryWorkSheet = ws
                Exit For
            End If
        Next

        If summaryWorkSheet Is Nothing Then Return

        Dim lookupTable = EnsureLookupTable(summaryWorkSheet)

        PopulateLookupTable(lookupTable)

        ApplyDataValidation(summaryWorkSheet)

        ApplyLookupFormulas(summaryWorkSheet)

    End Sub

    Private Function EnsureLookupTable(ws As Excel.Worksheet) As Excel.ListObject

        Dim lo As Excel.ListObject = Nothing

        'Try find existing table
        For Each tbl As Excel.ListObject In ws.ListObjects
            If tbl.Name.Equals("TestLookup", StringComparison.OrdinalIgnoreCase) Then
                lo = tbl
                Exit For
            End If
        Next

        If lo IsNot Nothing Then
            Return lo
        End If

        ' ---- Create new table ----

        Dim startRow As Integer = 1

        'If DvtReportSummaryTable exists, place below it
        For Each tbl As Excel.ListObject In ws.ListObjects
            If tbl.Name.Equals("DvtReportSummaryTable", StringComparison.OrdinalIgnoreCase) Then
                startRow = tbl.Range.Row + tbl.Range.Rows.Count + 3
                Exit For
            End If
        Next

        Dim headerRange As Excel.Range = ws.Range("A" & startRow & ":C" & startRow)

        headerRange.Cells(1, 1).Value = "DVT"
        headerRange.Cells(1, 2).Value = "OMS"
        headerRange.Cells(1, 3).Value = "ITEMS"

        lo = ws.ListObjects.Add(
            SourceType:=Excel.XlListObjectSourceType.xlSrcRange,
            Source:=headerRange,
            XlListObjectHasHeaders:=Excel.XlYesNoGuess.xlYes
        )

        lo.Name = "TestLookup"

        Return lo

    End Function

    Private Sub PopulateLookupTable(lo As Excel.ListObject)

        Dim coll = TestItemRegistry.GetAll()
        Dim requiredCount As Integer = coll.Count

        'Clear existing rows
        If lo.ListRows.Count > 0 Then
            lo.DataBodyRange.Delete()
        End If

        Dim ti As TestItem

        For Each ti In coll
            Dim newRow As Excel.ListRow = lo.ListRows.Add()
            newRow.Range.Cells(1, 1).Value = ti.DVT
            newRow.Range.Cells(1, 2).Value = ti.OMS
            newRow.Range.Cells(1, 3).Value = ti.ItemName
        Next

        'Hide rows
        lo.Range.EntireRow.Hidden = True

    End Sub

    Private Sub ApplyDataValidation(ws As Excel.Worksheet)

        Dim summaryTable As Excel.ListObject = Nothing

        For Each tbl As Excel.ListObject In ws.ListObjects
            If tbl.Name.Equals("DvtReportSummaryTable", StringComparison.OrdinalIgnoreCase) Then
                summaryTable = tbl
                Exit For
            End If
        Next

        If summaryTable Is Nothing Then Return
        If summaryTable.DataBodyRange Is Nothing Then Return

        Dim itemsColumn As Excel.ListColumn = summaryTable.ListColumns("ITEMS")

        With itemsColumn.DataBodyRange.Validation
            .Delete()
            .Add(Type:=Excel.XlDVType.xlValidateList,
                 AlertStyle:=Excel.XlDVAlertStyle.xlValidAlertStop,
                 Formula1:="=TestLookup[ITEMS]")
            .IgnoreBlank = True
            .InCellDropdown = True
        End With

    End Sub

    Private Sub ApplyLookupFormulas(ws As Excel.Worksheet)

        Dim summaryTable As Excel.ListObject = Nothing

        For Each tbl As Excel.ListObject In ws.ListObjects
            If tbl.Name.Equals("DvtReportSummaryTable", StringComparison.OrdinalIgnoreCase) Then
                summaryTable = tbl
                Exit For
            End If
        Next

        If summaryTable Is Nothing Then Return
        If summaryTable.DataBodyRange Is Nothing Then Return

        summaryTable.ListColumns("DVT").DataBodyRange.Formula =
            "=XLOOKUP([@ITEMS], TestLookup[ITEMS], TestLookup[DVT], """")"

        summaryTable.ListColumns("OMS").DataBodyRange.Formula =
            "=XLOOKUP([@ITEMS], TestLookup[ITEMS], TestLookup[OMS], """")"

    End Sub
End Class
