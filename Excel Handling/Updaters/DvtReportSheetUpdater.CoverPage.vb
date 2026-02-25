' ===============================
' Cover Page partial class
' ===============================
Imports System.Windows.Forms
Imports Microsoft.Office.Interop

Partial Public Class DvtReportSheetUpdater

    Public Sub UpdateCoverPageSheetFromOpenWorkbook(wb As Excel.Workbook, values As ResolvedTestMetadata)
        Dim coverPageWorkSheet As Excel.Worksheet = Nothing

        ' ---- Locate Cover Page worksheet ----
        For Each ws As Excel.Worksheet In wb.Sheets
            If ws.Name.Equals("Cover Page", StringComparison.OrdinalIgnoreCase) Then
                coverPageWorkSheet = ws
                Exit For
            End If
        Next

        If coverPageWorkSheet Is Nothing Then Return

        ' ---- Named fields ----
        WriteNamed(coverPageWorkSheet, "PowerSupplyModel", values.ModelNumber)
        WriteNamed(coverPageWorkSheet, "PowerSupplySerialNumber", values.SerialNumber)
        WriteNamed(coverPageWorkSheet, "PowerSupplyFirmwareVersion", values.FirmwareVersion)

        ' ---- Defaults for dialog ----
        Dim suggestedDescription = values.DevelopmentPhase
        Dim suggestedEngineer = values.TestedBy

        ' ---- Read latest Rev from table (if any) ----
        Dim totalRows As Integer
        Dim revList As List(Of String)
        ReadTableColumnValues(coverPageWorkSheet, "DvtReportOverviewTable", "Rev", totalRows, revList)

        ' Safely get latest revision
        Dim latestRev As String = Nothing
        If revList IsNot Nothing AndAlso revList.Count > 0 Then
            latestRev = revList.Last()
        End If

        ' ---- Auto-increment revision ----
        Dim suggestedRev As String = IncrementRevision(latestRev)

        ' ---- Prompt user (blocks safely) ----
        Using dlg As New CoverPageRevisionEntryTablePromptDialog(suggestedDescription, suggestedEngineer, suggestedRev)
            If dlg.ShowDialog() <> DialogResult.OK Then Return

            ' ---- Build table row values ----
            Dim tableValues As New Dictionary(Of String, Object) From {
            {"Rev", dlg.Revision},
            {"Engineer", dlg.Engineer},
            {"Description", dlg.Description},
            {"Date Prepared", Date.Today}
        }

            WriteToTable(coverPageWorkSheet, "DvtReportOverviewTable", tableValues)
        End Using

        wb.Save()
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

End Class