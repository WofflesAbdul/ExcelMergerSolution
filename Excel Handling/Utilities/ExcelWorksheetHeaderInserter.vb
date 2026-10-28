Imports Microsoft.Office.Interop.Excel

''' <summary>
''' Inserts header worksheets into a workbook according to header priority.
''' </summary>
Public Class ExcelWorksheetHeaderInserter

    ''' <summary>
    ''' Copies a header worksheet into the destination workbook in header-priority order.
    ''' </summary>
    Public Sub CopyHeaderWorksheet(sourceSheet As Worksheet, targetSheets As Sheets, sheetName As String)
        If sourceSheet Is Nothing Then Throw New ArgumentNullException(NameOf(sourceSheet))
        If targetSheets Is Nothing Then Throw New ArgumentNullException(NameOf(targetSheets))
        If targetSheets.Count = 0 Then Throw New InvalidOperationException("The destination workbook has no worksheets.")

        Dim insertionIndex As Integer = 1
        Dim insertionSheet As Worksheet = Nothing

        Try
            Dim newPriority As Integer = HeaderSequence.GetHeaderPriority(sheetName)

            For i As Integer = 1 To targetSheets.Count
                Dim worksheet As Worksheet = Nothing

                Try
                    worksheet = CType(targetSheets.Item(i), Worksheet)

                    Dim existingPriority As Integer = HeaderSequence.GetHeaderPriority(worksheet.Name)

                    ' Header sheets are expected to be grouped at the beginning.
                    If existingPriority = Integer.MaxValue Then Exit For

                    ' Insert after headers of equal or higher priority.
                    If existingPriority > newPriority Then
                        insertionIndex = i
                        Exit For
                    End If

                    insertionIndex = i + 1
                Finally
                    ExcelInteropHelper.ReleaseComObject(worksheet)
                End Try
            Next

            If insertionIndex > targetSheets.Count Then
                insertionSheet = CType(targetSheets.Item(targetSheets.Count), Worksheet)
                sourceSheet.Copy(After:=insertionSheet)
            Else
                insertionSheet = CType(targetSheets.Item(insertionIndex), Worksheet)
                sourceSheet.Copy(Before:=insertionSheet)
            End If
        Finally
            ExcelInteropHelper.ReleaseComObject(insertionSheet)
        End Try
    End Sub

End Class