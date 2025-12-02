Public Module HeaderSequence
    Private ReadOnly HeaderNames As List(Of String) = New List(Of String) From {
        "Title",
        "Summary",
        "Notes",
        "Observation Notes",
        "Observations"
    }

    ''' <summary>
    ''' Returns a priority number for a header sheet.
    ''' Smaller number = higher priority.
    ''' Headers with any suffix (e.g., "Summary 25C") are recognized.
    ''' </summary>
    Public Function GetHeaderPriority(name As String) As Integer
        If String.IsNullOrWhiteSpace(name) Then Return Integer.MaxValue

        Dim normalized = name.ToLowerInvariant().Trim()

        For i As Integer = 0 To HeaderNames.Count - 1
            Dim header = HeaderNames(i).ToLowerInvariant()
            ' Match if sheet starts with header name
            If normalized.StartsWith(header) Then
                Return i ' 0 = Title, 1 = Summary, 2 = Notes
            End If
        Next

        ' Not a header
        Return Integer.MaxValue
    End Function
End Module
