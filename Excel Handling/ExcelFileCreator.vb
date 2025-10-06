Imports Microsoft.Office.Interop
Imports System.IO

Public NotInheritable Class ExcelFileCreator

    ' Prevent instantiation
    Private Sub New()
    End Sub

    ' Predefined placeholder sheet names
    Private Shared ReadOnly PlaceholderNames As String() = {"Placeholder", "Sheet1", "Sheet2", "Sheet3"}

    ''' <summary>
    ''' Creates a new Excel workbook in the specified directory with the specified filename.
    ''' Returns the full path of the created workbook.
    ''' </summary>
    ''' <param name="directoryPath">Directory to create the workbook in.</param>
    ''' <param name="fileName">Filename including extension, e.g., "MyWorkbook.xlsx"</param>
    ''' <returns>Full file path of the created workbook.</returns>
    ''' <exception cref="ArgumentException">If directoryPath or fileName is null/empty.</exception>
    ''' <exception cref="ApplicationException">If workbook creation fails.</exception>
    Public Shared Function CreateNewExcel(directoryPath As String, fileName As String) As String
        If String.IsNullOrWhiteSpace(directoryPath) Then Throw New ArgumentException("Directory path cannot be empty.", NameOf(directoryPath))
        If String.IsNullOrWhiteSpace(fileName) Then Throw New ArgumentException("Filename cannot be empty.", NameOf(fileName))

        ' Ensure filename has .xlsx extension
        If Path.GetExtension(fileName).ToLower() <> ".xlsx" Then
            fileName &= ".xlsx"
        End If

        Dim fullPath As String = Path.Combine(directoryPath, fileName)

        ' Ensure folder exists
        If Not Directory.Exists(directoryPath) Then
            Directory.CreateDirectory(directoryPath)
        End If

        Dim excelApp As Excel.Application = Nothing
        Dim wb As Excel.Workbook = Nothing

        Try
            excelApp = New Excel.Application()
            excelApp.DisplayAlerts = False

            wb = excelApp.Workbooks.Add()
            wb.SaveAs(fullPath)
            wb.Close(SaveChanges:=True)

            ' Verify file exists
            If Not File.Exists(fullPath) Then
                Throw New ApplicationException($"Failed to create Excel file at {fullPath}")
            End If

            Return fullPath

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
