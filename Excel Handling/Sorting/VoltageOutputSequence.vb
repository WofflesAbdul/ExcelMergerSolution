Imports System.Text.RegularExpressions

Public Module VoltageOutputSequence
    Public Function GetVoltageOutputIndex(suffix As String) As Integer
        If String.IsNullOrWhiteSpace(suffix) Then Return Integer.MaxValue

        Dim match As Match = Regex.Match(suffix, "\bV(\d+)\b", RegexOptions.IgnoreCase)
        If match.Success Then
            Dim num As Integer = Integer.Parse(match.Groups(1).Value)
            ' V0 should be last, so we assign it the highest sort order
            If num = 0 Then Return Integer.MaxValue - 1
            Return num
        End If

        Return Integer.MaxValue
    End Function
End Module
