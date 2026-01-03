Public Class TestMetadataProcessor

    Public Shared Function ResolveDominant(sheets As List(Of TestReportSheetInfo)) As ResolvedTestMetadata
        Return New ResolvedTestMetadata With {
            .DevelopmentPhase = MostCommon(sheets.Select(Function(x) x.DevelopmentPhase)),
            .FirmwareVersion = MostCommon(sheets.Select(Function(x) x.FirmwareVersion)),
            .ModelNumber = MostCommon(sheets.Select(Function(x) x.ModelNumber)),
            .SerialNumber = MostCommon(sheets.Select(Function(x) x.SerialNumber)),
            .TestedBy = MostCommon(sheets.Select(Function(x) x.TestedBy))
        }
    End Function

    Private Shared Function MostCommon(values As IEnumerable(Of String)) As String
        Return values _
            .Where(Function(v) Not String.IsNullOrWhiteSpace(v)) _
            .GroupBy(Function(v) v) _
            .OrderByDescending(Function(g) g.Count()) _
            .Select(Function(g) g.Key) _
            .FirstOrDefault()
    End Function
End Class


