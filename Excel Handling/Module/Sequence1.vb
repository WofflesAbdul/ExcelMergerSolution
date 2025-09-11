Public Module Sequence1
    ''' <summary>
    ''' Defines the base names and their order for sorting.
    ''' </summary>
    Public ReadOnly Property BaseNames As List(Of String)
        Get
            Return New List(Of String) From {
                "Ripple and Noise",
                "Start Up Time",
                "Start Up Rise Time",
                "Holdup Time",
                "Turn Off Fall Time"
            }
        End Get
    End Property
End Module
