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

        PrepareSummaryTable(summaryWorkSheet)

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

        If lo IsNot Nothing Then Return lo

        ' ---- Create new table ----

        Dim startRow As Integer = 1

        ' If DvtReportSummaryTable exists, place below it
        For Each tbl As Excel.ListObject In ws.ListObjects
            If tbl.Name.Equals("DvtReportSummaryTable", StringComparison.OrdinalIgnoreCase) Then
                startRow = tbl.Range.Row + tbl.Range.Rows.Count + 200
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

        Dim items = TestItemRegistry.GetAll().ToList()

        ' Clear table completely
        While lo.ListRows.Count > 0
            lo.ListRows(1).Delete()
        End While

        If items.Count = 0 Then Exit Sub

        ' Resize table to correct size
        Dim startCell As Excel.Range = lo.HeaderRowRange.Cells(1, 1)
        Dim newRange As Excel.Range =
            startCell.Resize(items.Count + 1, 3)

        lo.Resize(newRange)

        ' Build 2D array for Excel
        Dim data(items.Count - 1, 2) As Object

        For i As Integer = 0 To items.Count - 1
            data(i, 0) = items(i).DVT
            data(i, 1) = items(i).OMS
            data(i, 2) = items(i).ItemName
        Next

        lo.DataBodyRange.Value = data

        ' Hide lookup table
        lo.Range.EntireRow.Hidden = True

    End Sub

    ' ------------------------------
    ' Prepare summary table states and apply validation/formulas once
    ' ------------------------------
    Private Sub PrepareSummaryTable(ws As Excel.Worksheet)

        Dim summaryTable As Excel.ListObject = Nothing

        For Each tbl As Excel.ListObject In ws.ListObjects
            If tbl.Name.Equals("DvtReportSummaryTable", StringComparison.OrdinalIgnoreCase) Then
                summaryTable = tbl
                Exit For
            End If
        Next

        If summaryTable Is Nothing Then Return

        ' ---------------------------
        ' Step 1: Remove phantom/extra empty rows
        ' ---------------------------
        For i As Integer = summaryTable.ListRows.Count To 1 Step -1
            Dim rowEmpty As Boolean = True
            For Each cell As Excel.Range In summaryTable.ListRows(i).Range.Cells
                If Not String.IsNullOrWhiteSpace(cell.Text) Then
                    rowEmpty = False
                    Exit For
                End If
            Next
            If rowEmpty Then
                summaryTable.ListRows(i).Delete()
            End If
        Next

        ' ---------------------------
        ' Step 2: Add a single empty row if table is empty
        ' ---------------------------
        If summaryTable.ListRows.Count = 0 Then
            summaryTable.ListRows.Add()
        End If

        ' ---------------------------
        ' Step 3: Apply data validation and formulas only if not already applied
        ' ---------------------------
        Dim applyFormulas As Boolean = False
        Try
            Dim dvtCol = summaryTable.ListColumns("DVT").DataBodyRange
            Dim omsCol = summaryTable.ListColumns("OMS").DataBodyRange

            ' Apply formulas only if first cell has no formula
            If Not dvtCol.Cells(1, 1).HasFormula AndAlso Not omsCol.Cells(1, 1).HasFormula Then
                applyFormulas = True
            End If
        Catch
            ' Columns missing? Apply formulas anyway
            applyFormulas = True
        End Try

        If applyFormulas Then
            ApplyDataValidation(ws)
            ApplyLookupFormulas(ws)
        End If

    End Sub

    Private Sub ApplyDataValidation(ws As Excel.Worksheet)

        Dim summaryTable As Excel.ListObject = Nothing
        Dim lookupTable As Excel.ListObject = Nothing

        ' Locate tables
        For Each tbl As Excel.ListObject In ws.ListObjects

            If tbl.Name.Equals("DvtReportSummaryTable",
                               StringComparison.OrdinalIgnoreCase) Then
                summaryTable = tbl
            End If

            If tbl.Name.Equals("TestLookup",
                               StringComparison.OrdinalIgnoreCase) Then
                lookupTable = tbl
            End If

        Next

        If summaryTable Is Nothing Then Return
        If lookupTable Is Nothing Then Return

        ' Ensure at least 1 row in summary table
        If summaryTable.ListRows.Count = 0 Then
            summaryTable.ListRows.Add()
        End If

        If lookupTable.ListRows.Count = 0 Then Return

        Dim itemsColumn As Excel.ListColumn = Nothing
        Dim lookupItemsColumn As Excel.ListColumn = Nothing

        Try
            itemsColumn = summaryTable.ListColumns("ITEMS")
            lookupItemsColumn = lookupTable.ListColumns("ITEMS")
        Catch
            Return
        End Try

        If itemsColumn Is Nothing Then Return
        If lookupItemsColumn Is Nothing Then Return
        If itemsColumn.DataBodyRange Is Nothing Then Return
        If lookupItemsColumn.DataBodyRange Is Nothing Then Return

        ' Get absolute address of lookup range
        Dim lookupRangeAddress As String =
            lookupItemsColumn.DataBodyRange.Address(
                ReferenceStyle:=Excel.XlReferenceStyle.xlA1,
                RowAbsolute:=True,
                ColumnAbsolute:=True,
                External:=True)

        With itemsColumn.DataBodyRange.Validation
            .Delete()
            .Add(Type:=Excel.XlDVType.xlValidateList,
                 AlertStyle:=Excel.XlDVAlertStyle.xlValidAlertStop,
                 Formula1:="=" & lookupRangeAddress)
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

        summaryTable.ListColumns("DVT").DataBodyRange.Formula = ""
        summaryTable.ListColumns("OMS").DataBodyRange.Formula = ""

        summaryTable.ListColumns("DVT").DataBodyRange.Formula =
            "=XLOOKUP([@ITEMS], TestLookup[ITEMS], TestLookup[DVT], """")"

        summaryTable.ListColumns("OMS").DataBodyRange.Formula =
            "=XLOOKUP([@ITEMS], TestLookup[ITEMS], TestLookup[OMS], """")"

    End Sub
End Class