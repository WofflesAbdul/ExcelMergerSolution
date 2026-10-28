Imports Excel = Microsoft.Office.Interop.Excel

''' <summary>
''' Provides shared worksheet lookup, duplicate detection, and copy operations.
''' </summary>
Public NotInheritable Class ExcelWorksheetHelper

    ''' <summary>
    ''' Finds a worksheet by name using a case-insensitive comparison.
    ''' The caller owns the returned COM reference and must release it.
    ''' </summary>
    Public Shared Function GetWorksheetIfExists(sheets As Excel.Sheets, sheetName As String) As Excel.Worksheet
        If sheets Is Nothing Then Throw New ArgumentNullException(NameOf(sheets))
        If String.IsNullOrWhiteSpace(sheetName) Then Throw New ArgumentException("A worksheet name is required.", NameOf(sheetName))

        For i As Integer = 1 To sheets.Count
            Dim worksheet As Excel.Worksheet = Nothing

            Try
                worksheet = TryCast(sheets.Item(i), Excel.Worksheet)

                If worksheet IsNot Nothing AndAlso String.Equals(worksheet.Name, sheetName, StringComparison.OrdinalIgnoreCase) Then
                    Dim result As Excel.Worksheet = worksheet
                    worksheet = Nothing
                    Return result
                End If
            Finally
                ExcelInteropHelper.ReleaseComObject(worksheet)
            End Try
        Next

        Return Nothing
    End Function

    ''' <summary>
    ''' Determines whether a worksheet with the specified name exists.
    ''' </summary>
    Public Shared Function WorksheetExists(sheets As Excel.Sheets, sheetName As String) As Boolean
        Dim worksheet As Excel.Worksheet = Nothing

        Try
            worksheet = GetWorksheetIfExists(sheets, sheetName)
            Return worksheet IsNot Nothing
        Finally
            ExcelInteropHelper.ReleaseComObject(worksheet)
        End Try
    End Function

    ''' <summary>
    ''' Copies a worksheet only if the destination does not already contain a worksheet with the same name.
    ''' Header worksheets are inserted using header priority; ordinary worksheets are appended.
    ''' Returns True if the worksheet was copied; otherwise, returns False.
    ''' </summary>
    Public Shared Function CopyWorksheetIfMissing(sourceSheet As Excel.Worksheet, targetSheets As Excel.Sheets, sheetName As String, Optional headerInserter As ExcelWorksheetHeaderInserter = Nothing) As Boolean
        If sourceSheet Is Nothing Then Throw New ArgumentNullException(NameOf(sourceSheet))
        If targetSheets Is Nothing Then Throw New ArgumentNullException(NameOf(targetSheets))
        If String.IsNullOrWhiteSpace(sheetName) Then Throw New ArgumentException("A worksheet name is required.", NameOf(sheetName))

        If WorksheetExists(targetSheets, sheetName) Then Return False

        If HeaderSequence.GetHeaderPriority(sheetName) <> Integer.MaxValue Then
            If headerInserter Is Nothing Then headerInserter = New ExcelWorksheetHeaderInserter()
            headerInserter.CopyHeaderWorksheet(sourceSheet, targetSheets, sheetName)
        Else
            If targetSheets.Count = 0 Then Throw New InvalidOperationException("The destination workbook has no worksheets.")

            Dim lastSheet As Excel.Worksheet = Nothing

            Try
                lastSheet = CType(targetSheets.Item(targetSheets.Count), Excel.Worksheet)
                sourceSheet.Copy(After:=lastSheet)
            Finally
                ExcelInteropHelper.ReleaseComObject(lastSheet)
            End Try
        End If

        Return True
    End Function

End Class