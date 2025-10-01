Imports Microsoft.Office.Interop
Imports System.IO

Public Class ExcelFileCreator

    ''' <summary>
    ''' True if the file was newly created as a blank workbook with a placeholder sheet.
    ''' False if created from a template or existing file.
    ''' </summary>
    Public Property HasPlaceholderSheet As Boolean = False

    ''' <summary>
    ''' Creates a new Excel workbook at the specified full path.
    ''' Adds a placeholder sheet "Placeholder" to identify blank workbooks.
    ''' </summary>
    ''' <param name="fullPath">Full file path including filename.xlsx</param>
    ''' <returns>True if file was successfully created and exists, otherwise False</returns>
    Public Function CreateNewExcel(fullPath As String) As Boolean
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

            ' Create new workbook with one sheet
            wb = excelApp.Workbooks.Add()

            ' Rename the default sheet to Placeholder
            If wb.Sheets.Count > 0 Then
                wb.Sheets(1).Name = "Placeholder"
            End If

            wb.SaveAs(fullPath)
            wb.Close(SaveChanges:=True)

            ' Set flag
            HasPlaceholderSheet = True

            ' Verify file exists
            Return File.Exists(fullPath)

        Catch ex As Exception
            ' Could log ex.Message if desired
            HasPlaceholderSheet = False
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
    ''' Removes the placeholder sheet from a workbook, if it exists.
    ''' Only call this for blank-workbook creations.
    ''' </summary>
    ''' <param name="fullPath">Full file path including filename.xlsx</param>
    ''' <param name="placeholderName">Name of the placeholder sheet, default "Placeholder"</param>
    Public Sub RemovePlaceholderSheet(fullPath As String, Optional placeholderName As String = "Placeholder")
        If Not HasPlaceholderSheet Then Return
        If Not File.Exists(fullPath) Then Return

        Dim excelApp As Excel.Application = Nothing
        Dim wb As Excel.Workbook = Nothing

        Try
            excelApp = New Excel.Application()
            excelApp.DisplayAlerts = False

            wb = excelApp.Workbooks.Open(fullPath)

            Dim sheet As Excel.Worksheet = Nothing
            For Each ws As Excel.Worksheet In wb.Sheets
                If ws.Name = placeholderName Then
                    sheet = ws
                    Exit For
                End If
            Next

            If sheet IsNot Nothing Then
                wb.Sheets(sheet.Name).Delete()
                wb.Save()
            End If

        Catch ex As Exception
            ' Could log ex.Message if desired

        Finally
            If wb IsNot Nothing Then MarshalReleaseComObject(wb)
            If excelApp IsNot Nothing Then
                excelApp.Quit()
                MarshalReleaseComObject(excelApp)
            End If

            HasPlaceholderSheet = False ' reset
        End Try
    End Sub

    ''' <summary>
    ''' Helper to release COM object
    ''' </summary>
    Private Sub MarshalReleaseComObject(obj As Object)
        Try
            If obj IsNot Nothing Then
                System.Runtime.InteropServices.Marshal.ReleaseComObject(obj)
            End If
        Finally
            obj = Nothing
        End Try
    End Sub

End Class
