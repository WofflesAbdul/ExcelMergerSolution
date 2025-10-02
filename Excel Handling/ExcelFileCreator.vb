Imports Microsoft.Office.Interop
Imports System.IO

Public NotInheritable Class ExcelFileCreator

    ' Prevent instantiation
    Private Sub New()
    End Sub

    ' Predefined placeholder sheet names
    Private Shared ReadOnly PlaceholderNames As String() = {"Placeholder", "Sheet1", "Sheet2", "Sheet3"}

    ''' <summary>
    ''' Creates a new Excel workbook at the specified full path.
    ''' Optionally adds a "Placeholder" sheet to mark blank workbooks.
    ''' </summary>
    ''' <param name="fullPath">Full file path including filename.xlsx</param>
    ''' <param name="addPlaceholder">True to create with Placeholder sheet, False for default blank workbook</param>
    ''' <returns>True if file was successfully created and exists, otherwise False</returns>
    Public Shared Function CreateNewExcel(fullPath As String, Optional addPlaceholder As Boolean = False) As Boolean
        If String.IsNullOrWhiteSpace(fullPath) Then
            Throw New ArgumentException("File path cannot be null or empty.", NameOf(fullPath))
        End If

        ' Ensure folder exists
        Dim folder As String = Path.GetDirectoryName(fullPath)
        If Not Directory.Exists(folder) Then
            Directory.CreateDirectory(folder)
        End If

        Dim excelApp As Excel.Application = Nothing
        Dim wb As Excel.Workbook = Nothing

        Try
            excelApp = New Excel.Application()
            excelApp.DisplayAlerts = False

            ' Create new workbook
            wb = excelApp.Workbooks.Add()

            If addPlaceholder Then
                ' Rename the first sheet to Placeholder
                If wb.Sheets.Count > 0 Then
                    wb.Sheets(1).Name = "Placeholder"
                End If
            End If

            wb.SaveAs(fullPath)
            wb.Close(SaveChanges:=True)

            ' Verify file exists
            Return File.Exists(fullPath)

        Catch
            Return False

        Finally
            If wb IsNot Nothing Then MarshalReleaseComObject(wb)
            If excelApp IsNot Nothing Then
                excelApp.Quit()
                MarshalReleaseComObject(excelApp)
            End If
        End Try
    End Function

    ''' <summary>
    ''' Removes placeholder sheets if they exist in the workbook.
    ''' </summary>
    ''' <param name="fullPath">Full file path including filename.xlsx</param>
    Public Shared Sub RemovePlaceholderSheets(fullPath As String)
        If Not File.Exists(fullPath) Then Return

        Dim excelApp As Excel.Application = Nothing
        Dim wb As Excel.Workbook = Nothing

        Try
            excelApp = New Excel.Application()
            excelApp.DisplayAlerts = False
            wb = excelApp.Workbooks.Open(fullPath)

            ' Iterate backwards to safely delete
            For i As Integer = wb.Sheets.Count To 1 Step -1
                Dim ws As Excel.Worksheet = CType(wb.Sheets(i), Excel.Worksheet)
                If PlaceholderNames.Contains(ws.Name) Then
                    ws.Delete()
                End If
                MarshalReleaseComObject(ws)
            Next

            wb.Save()

        Catch
            ' Optionally log error

        Finally
            If wb IsNot Nothing Then MarshalReleaseComObject(wb)
            If excelApp IsNot Nothing Then
                excelApp.Quit()
                MarshalReleaseComObject(excelApp)
            End If
        End Try
    End Sub

    ''' <summary>
    ''' Helper to release COM object
    ''' </summary>
    Private Shared Sub MarshalReleaseComObject(obj As Object)
        Try
            If obj IsNot Nothing Then
                System.Runtime.InteropServices.Marshal.ReleaseComObject(obj)
            End If
        Finally
            obj = Nothing
        End Try
    End Sub

End Class
