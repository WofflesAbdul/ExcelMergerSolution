Imports System.IO
Imports Excel = Microsoft.Office.Interop.Excel

''' <summary>
''' Manages one Excel application and a destination/source workbook pair.
''' </summary>
Public NotInheritable Class ExcelWorkbookSession
    Implements IDisposable

    Private excel As Excel.Application
    Private disposed As Boolean

    ''' <summary>
    ''' Gets the Excel application owned by this session.
    ''' </summary>
    Public ReadOnly Property Application As Excel.Application
        Get
            ThrowIfDisposed()
            Return excel
        End Get
    End Property

    ''' <summary>
    ''' Gets the destination workbook.
    ''' </summary>
    Public ReadOnly Property DestinationWorkbook As Excel.Workbook

    ''' <summary>
    ''' Gets the source workbook, opened read-only.
    ''' </summary>
    Public ReadOnly Property SourceWorkbook As Excel.Workbook

    ''' <summary>
    ''' Opens the destination and source workbooks in one Excel application.
    ''' </summary>
    Public Sub New(destinationPath As String, sourcePath As String)
        ValidateFile(destinationPath, "destination")
        ValidateFile(sourcePath, "source")

        If String.Equals(Path.GetFullPath(destinationPath), Path.GetFullPath(sourcePath), StringComparison.OrdinalIgnoreCase) Then
            Throw New InvalidOperationException("The source and destination workbooks must be different files.")
        End If

        If Threading.Thread.CurrentThread.GetApartmentState() <> Threading.ApartmentState.STA Then
            Throw New InvalidOperationException("Excel operations must run on an STA thread.")
        End If

        Try
            excel = New Excel.Application With {
                .Visible = False,
                .DisplayAlerts = False
            }

            DestinationWorkbook = excel.Workbooks.Open(Path.GetFullPath(destinationPath))
            SourceWorkbook = excel.Workbooks.Open(Path.GetFullPath(sourcePath), ReadOnly:=True)
        Catch
            Dispose()
            Throw
        End Try
    End Sub

    ''' <summary>
    ''' Saves the destination workbook.
    ''' </summary>
    Public Sub SaveDestination()
        ThrowIfDisposed()
        DestinationWorkbook.Save()
    End Sub

    Private Shared Sub ValidateFile(filePath As String, fileRole As String)
        If String.IsNullOrWhiteSpace(filePath) OrElse Not File.Exists(filePath) Then
            Throw New FileNotFoundException("The " & fileRole & " workbook could not be found.", filePath)
        End If
    End Sub

    Private Sub ThrowIfDisposed()
        If disposed Then Throw New ObjectDisposedException(NameOf(ExcelWorkbookSession))
    End Sub

    ''' <summary>
    ''' Closes the workbooks and quits the Excel application owned by this session.
    ''' </summary>
    Public Sub Dispose() Implements IDisposable.Dispose
        If disposed Then Return

        disposed = True

        If SourceWorkbook IsNot Nothing Then
            Try
                SourceWorkbook.Close(SaveChanges:=False)
            Catch
                ' Continue cleanup.
            Finally
                ExcelInteropHelper.ReleaseComObject(SourceWorkbook)
            End Try
        End If

        If DestinationWorkbook IsNot Nothing Then
            Try
                DestinationWorkbook.Close(SaveChanges:=False)
            Catch
                ' Continue cleanup.
            Finally
                ExcelInteropHelper.ReleaseComObject(DestinationWorkbook)
            End Try
        End If

        If excel IsNot Nothing Then
            Try
                excel.Quit()
            Catch
                ' Continue cleanup.
            Finally
                ExcelInteropHelper.ReleaseComObject(excel)
                excel = Nothing
            End Try
        End If
    End Sub

End Class