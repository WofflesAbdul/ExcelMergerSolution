Public Class WorkbookNamedData
    Public Property CoverPage As CoverPageSheetInfo
    Public Property TestSheets As List(Of TestReportSheetInfo)

    Public Sub New()
        TestSheets = New List(Of TestReportSheetInfo)
    End Sub
End Class

