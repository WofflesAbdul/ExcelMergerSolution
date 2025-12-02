Imports System.Text.RegularExpressions

''' <summary>
''' Converts worksheet suffixes containing temperatures into sortable indices.
''' </summary>
Public Module TemperatureSequence

    ''' <summary>
    ''' Returns a sortable index for a given suffix.
    ''' </summary>
    ''' <param name="suffix">The part of the worksheet name after the base test name.</param>
    ''' <returns>An integer index for sorting; smaller = higher priority.</returns>
    Public Function GetTemperatureIndex(suffix As String) As Integer
        If String.IsNullOrWhiteSpace(suffix) Then
            Return Integer.MaxValue
        End If

        Dim normalized = suffix.ToLowerInvariant()

        ' >>> NEW: Prioritize "Ambient" before numeric temps <<<
        If normalized.Contains("ambient") Then
            Return -1000 ' Ensures ambient sorts first
        End If

        ' Match temperature patterns like "25C", "30C", "-20C"
        Dim tempMatch As Match = Regex.Match(suffix, "(-?\d+)\s*C", RegexOptions.IgnoreCase)
        If tempMatch.Success Then
            Dim tempValue As Integer = Integer.Parse(tempMatch.Groups(1).Value)

            ' Ascending part: 25C to 70C
            If tempValue >= 25 AndAlso tempValue <= 70 Then
                Return tempValue - 25 ' 25C -> 0, 30C -> 5, etc.
            End If

            ' Descending part: 20C down to -40C
            If tempValue < 25 AndAlso tempValue >= -40 Then
                ' Offset by a large number to come after ascending range
                Return 1000 + (25 - tempValue) ' 25C < tempValue -> 1000, 20C -> 1005, ...
            End If

            ' Any other temperature values
            Return Integer.MaxValue - 1
        End If

        ' If no numeric temperature found, push to end
        Return Integer.MaxValue
    End Function

End Module
