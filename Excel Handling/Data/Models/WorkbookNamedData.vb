Public Class WorkbookNamedData
    Public Property Title As TitleSheetInfo
    Public Property TestSheets As List(Of TestReportSheetInfo)

    Public Sub New()
        TestSheets = New List(Of TestReportSheetInfo)
    End Sub
End Class

