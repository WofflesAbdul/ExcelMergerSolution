Public Class TestItem

    Public ReadOnly Property ItemName As String
    Public ReadOnly Property DVT As String
    Public ReadOnly Property OMS As String

    Public Sub New(itemName As String, dvt As String, oms As String)
        Me.ItemName = itemName
        Me.DVT = dvt
        Me.OMS = oms
    End Sub

End Class