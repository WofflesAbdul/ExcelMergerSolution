
Imports System.Runtime.InteropServices
Imports System.Threading
Imports Microsoft.Office.Interop

Public Class FunctionalTestSorter

    <ComImport>
    <Guid("00000016-0000-0000-C000-000000000046")>
    <InterfaceType(ComInterfaceType.InterfaceIsIUnknown)>
    Private Interface IMessageFilter
        <PreserveSig>
        Function HandleInComingCall(
            callType As Integer,
            callerTask As IntPtr,
            tickCount As Integer,
            interfaceInfo As IntPtr) As Integer

        <PreserveSig>
        Function RetryRejectedCall(
            calleeTask As IntPtr,
            tickCount As Integer,
            rejectType As Integer) As Integer

        <PreserveSig>
        Function MessagePending(
            calleeTask As IntPtr,
            tickCount As Integer,
            pendingType As Integer) As Integer
    End Interface

    Private Class ExcelMessageFilter
        Implements IMessageFilter

        Public Function HandleInComingCall(
            callType As Integer,
            callerTask As IntPtr,
            tickCount As Integer,
            interfaceInfo As IntPtr) As Integer _
            Implements IMessageFilter.HandleInComingCall

            Return 0
        End Function

        Public Function RetryRejectedCall(
            calleeTask As IntPtr,
            tickCount As Integer,
            rejectType As Integer) As Integer _
            Implements IMessageFilter.RetryRejectedCall

            Const SERVERCALL_RETRYLATER As Integer = 2

            If rejectType = SERVERCALL_RETRYLATER AndAlso tickCount < 10000 Then
                Return 250
            End If

            Return -1
        End Function

        Public Function MessagePending(
            calleeTask As IntPtr,
            tickCount As Integer,
            pendingType As Integer) As Integer _
            Implements IMessageFilter.MessagePending

            Return 2
        End Function
    End Class

    <DllImport("ole32.dll")>
    Private Shared Function CoRegisterMessageFilter(
        newFilter As IMessageFilter,
        ByRef oldFilter As IMessageFilter) As Integer
    End Function

    Public Sub SortSheets(filePath As String)
        If Thread.CurrentThread.GetApartmentState() <> ApartmentState.STA Then
            Throw New InvalidOperationException(
                "Excel sorting must run on an STA thread.")
        End If

        Dim excelApp As Excel.Application = Nothing
        Dim wb As Excel.Workbook = Nothing
        Dim previousFilter As IMessageFilter = Nothing
        Dim filterRegistered As Boolean = False

        Try
            Dim filter As IMessageFilter = New ExcelMessageFilter()
            Dim hr As Integer = CoRegisterMessageFilter(filter, previousFilter)

            If hr < 0 Then
                Marshal.ThrowExceptionForHR(hr)
            End If

            filterRegistered = True

            excelApp = New Excel.Application()
            wb = excelApp.Workbooks.Open(filePath)

            SortSheets(wb)

            wb.Save()

        Catch ex As Exception
            Throw New ApplicationException($"Sort failed: {ex.Message}", ex)

        Finally
            If wb IsNot Nothing Then
                Try
                    wb.Close(SaveChanges:=False)
                Catch
                End Try

                ReleaseComObject(wb)
            End If

            If excelApp IsNot Nothing Then
                Try
                    excelApp.Quit()
                Catch
                End Try

                ReleaseComObject(excelApp)
            End If

            If filterRegistered Then
                Dim ignoredFilter As IMessageFilter = Nothing
                CoRegisterMessageFilter(previousFilter, ignoredFilter)
            End If
        End Try
    End Sub

    Public Sub SortSheets(wb As Excel.Workbook)
        If wb Is Nothing Then
            Throw New ArgumentNullException(NameOf(wb))
        End If

        Dim sheetInfoList As New List(Of SheetNameParsingInfo)
        Dim sheets As Excel.Sheets = Nothing

        Try
            sheets = wb.Worksheets

            For i As Integer = 1 To sheets.Count
                Dim ws As Excel.Worksheet = Nothing

                Try
                    ws = CType(sheets.Item(i), Excel.Worksheet)

                    Dim fullName As String = ws.Name

                    ' Split into base name and suffix.
                    Dim allParts() As String = fullName.Split(" "c)
                    Dim parts() As String

                    If allParts.Length > 1 Then
                        parts = {
                            allParts(0),
                            String.Join(" ", allParts, 1, allParts.Length - 1)
                        }
                    Else
                        parts = allParts
                    End If

                    Dim suffix As String =
                        If(parts.Length > 1, parts(1), String.Empty)

                    ' Determine sorting indices.
                    Dim baseIndex As Integer =
                        FunctionalTestSequence.TestNames.FindIndex(
                            Function(b) fullName.IndexOf(
                                b,
                                StringComparison.OrdinalIgnoreCase) >= 0)

                    If baseIndex = -1 Then
                        baseIndex = Integer.MaxValue
                    End If

                    Dim voltageOutputIndex As Integer =
                        VoltageOutputSequence.GetVoltageOutputIndex(suffix)

                    Dim vpsuIndex As Integer =
                        VpsuSequence.GetVpsuIndex(suffix)

                    Dim tempIndex As Integer =
                        TemperatureSequence.GetTemperatureIndex(suffix)

                    Dim headerPriority As Integer =
                        HeaderSequence.GetHeaderPriority(fullName)

                    sheetInfoList.Add(New SheetNameParsingInfo With {
                        .Worksheet = ws,
                        .OriginalName = fullName,
                        .baseIndex = baseIndex,
                        .voltageOutputIndex = voltageOutputIndex,
                        .vpsuIndex = vpsuIndex,
                        .tempIndex = tempIndex,
                        .headerPriority = headerPriority
                    })

                    ' The list now owns the worksheet reference.
                    ws = Nothing

                Finally
                    ReleaseComObject(ws)
                End Try
            Next

            ' Sort by Header -> Temperature -> BaseName -> VoltageOutput -> Vpsu.
            Dim sortedSheets = sheetInfoList _
                .OrderBy(Function(s) s.headerPriority) _
                .ThenBy(Function(s) s.tempIndex) _
                .ThenBy(Function(s) s.baseIndex) _
                .ThenBy(Function(s) s.voltageOutputIndex) _
                .ThenBy(Function(s) s.vpsuIndex) _
                .ToList()

            ' Move sheets into the required order.
            If sortedSheets.Count > 0 Then
                Dim firstSheet As Excel.Worksheet = Nothing

                Try
                    firstSheet = CType(sheets.Item(1), Excel.Worksheet)

                    If Not String.Equals(
                        firstSheet.Name,
                        sortedSheets(0).OriginalName,
                        StringComparison.Ordinal) Then

                        sortedSheets(0).Worksheet.Move(Before:=firstSheet)
                    End If

                Finally
                    ReleaseComObject(firstSheet)
                End Try

                For i As Integer = 1 To sortedSheets.Count - 1
                    sortedSheets(i).Worksheet.Move(
                        After:=sortedSheets(i - 1).Worksheet)
                Next
            End If

        Catch ex As Exception
            Throw New ApplicationException($"Sort failed: {ex.Message}", ex)

        Finally
            For Each sheetInfo In sheetInfoList
                ReleaseComObject(sheetInfo.Worksheet)
            Next

            ReleaseComObject(sheets)
        End Try
    End Sub

    Private Sub ReleaseComObject(obj As Object)
        Try
            If obj IsNot Nothing AndAlso Marshal.IsComObject(obj) Then
                Marshal.ReleaseComObject(obj)
            End If
        Catch
        End Try
    End Sub

End Class
