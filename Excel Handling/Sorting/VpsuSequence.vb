Imports System.Text.RegularExpressions

Public Module VpsuSequence
    ''' <summary>
    ''' Returns a sortable index for a Vpsu suffix.
    ''' Higher percentages are prioritized.
    ''' </summary>
    Public Function GetVpsuIndex(suffix As String) As Integer
        If String.IsNullOrWhiteSpace(suffix) Then Return Integer.MaxValue

        Dim match As Match = Regex.Match(suffix, "Vpsu\s*=\s*(\d+)%", RegexOptions.IgnoreCase)
        If match.Success Then
            ' Descending order: 110% -> 0, 80% -> 30 etc.
            Return 110 - Integer.Parse(match.Groups(1).Value)
        End If

        ' If not a Vpsu suffix, push to the end
        Return Integer.MaxValue
    End Function
End Module

