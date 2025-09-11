Imports Microsoft.Office.Interop
''' <summary>
''' Holds sheet info for sorting.
''' </summary>
Public Class SheetInfo
    Public Property Worksheet As Excel.Worksheet
    Public Property OriginalName As String
    Public Property BaseIndex As Integer
    Public Property SuffixIndex As Integer
End Class

