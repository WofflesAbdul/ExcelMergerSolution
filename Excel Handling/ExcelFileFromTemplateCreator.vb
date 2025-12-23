Imports Microsoft.Office.Interop

Public Class ExcelFileFromTemplateCreator

    ' Template path relative to application startup directory
    Private Shared ReadOnly TemplateRelativePath As String =
        IO.Path.Combine("Resources", "Templates", "BaseTemplate.xlt")

    Public Shared Function CreateFromTemplate(
        directoryPath As String,
        fileName As String) As String

        Dim excelApp As Excel.Application = Nothing
        Dim templateWB As Excel.Workbook = Nothing

        Try
            Dim basePath As String = AppDomain.CurrentDomain.BaseDirectory
            Dim templateFullPath As String = IO.Path.Combine(basePath, TemplateRelativePath)


            If Not IO.File.Exists(templateFullPath) Then
                Throw New IO.FileNotFoundException(
                    $"Template file not found at: {templateFullPath}")
            End If

            excelApp = New Excel.Application()

            ' Open template safely
            templateWB = excelApp.Workbooks.Open(
                templateFullPath,
                ReadOnly:=True)

            Dim outputPath As String =
                IO.Path.Combine(directoryPath, fileName & ".xlsx")

            templateWB.SaveAs(outputPath)

            Return outputPath

        Catch ex As Exception
            Throw New ApplicationException(
                $"Failed to create Excel file from template: {ex.Message}", ex)

        Finally
            If templateWB IsNot Nothing Then templateWB.Close(SaveChanges:=False)
            If excelApp IsNot Nothing Then excelApp.Quit()

            ReleaseComObject(templateWB)
            ReleaseComObject(excelApp)
        End Try
    End Function

    Private Shared Sub ReleaseComObject(ByVal obj As Object)
        Try
            If obj IsNot Nothing Then
                Runtime.InteropServices.Marshal.ReleaseComObject(obj)
            End If
        Finally
            obj = Nothing
        End Try
    End Sub

End Class
