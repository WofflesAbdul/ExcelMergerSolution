Imports System.Runtime.InteropServices
Imports System.Threading
Imports Microsoft.Office.Interop

Public NotInheritable Class ExcelHelper

    <ComImport>
    <Guid("00000016-0000-0000-C000-000000000046")>
    <InterfaceType(ComInterfaceType.InterfaceIsIUnknown)>
    Private Interface IMessageFilter
        <PreserveSig>
        Function HandleInComingCall(callType As Integer, callerTask As IntPtr, tickCount As Integer, interfaceInfo As IntPtr) As Integer

        <PreserveSig>
        Function RetryRejectedCall(calleeTask As IntPtr, tickCount As Integer, rejectType As Integer) As Integer

        <PreserveSig>
        Function MessagePending(calleeTask As IntPtr, tickCount As Integer, pendingType As Integer) As Integer
    End Interface

    Private Class ExcelMessageFilter
        Implements IMessageFilter

        Public Function HandleInComingCall(callType As Integer, callerTask As IntPtr, tickCount As Integer, interfaceInfo As IntPtr) As Integer _
            Implements IMessageFilter.HandleInComingCall

            Return 0
        End Function

        Public Function RetryRejectedCall(calleeTask As IntPtr, tickCount As Integer, rejectType As Integer) As Integer _
            Implements IMessageFilter.RetryRejectedCall

            Const SERVERCALL_RETRYLATER As Integer = 2

            If rejectType = SERVERCALL_RETRYLATER AndAlso tickCount < 10000 Then
                Return 250
            End If

            Return -1
        End Function

        Public Function MessagePending(calleeTask As IntPtr, tickCount As Integer, pendingType As Integer) As Integer _
            Implements IMessageFilter.MessagePending

            Return 2
        End Function
    End Class

    <DllImport("ole32.dll")>
    Private Shared Function CoRegisterMessageFilter(newFilter As IMessageFilter, ByRef oldFilter As IMessageFilter) As Integer
    End Function

    Public NotInheritable Class ExcelMessageFilterScope
        Implements IDisposable

        Private ReadOnly previousFilter As IMessageFilter
        Private disposed As Boolean

        Public Sub New()
            If Thread.CurrentThread.GetApartmentState() <> ApartmentState.STA Then
                Throw New InvalidOperationException("Excel operations must run on an STA thread.")
            End If

            Dim filter As IMessageFilter = New ExcelMessageFilter()
            Dim hr As Integer = CoRegisterMessageFilter(filter, previousFilter)

            If hr < 0 Then
                Marshal.ThrowExceptionForHR(hr)
            End If
        End Sub

        Public Sub Dispose() Implements IDisposable.Dispose
            If disposed Then Return

            disposed = True

            Dim ignoredFilter As IMessageFilter = Nothing
            Dim hr As Integer = CoRegisterMessageFilter(previousFilter, ignoredFilter)

            If hr < 0 Then
                Marshal.ThrowExceptionForHR(hr)
            End If
        End Sub
    End Class

    Public Shared Sub ResetWorksheetCursors(workbook As Excel.Workbook, Optional reportProgress As Action(Of Integer) = Nothing, Optional progressStart As Integer = 0, Optional progressEnd As Integer = 100)

        If workbook Is Nothing Then Throw New ArgumentNullException(NameOf(workbook))

        Dim firstSheet As Excel.Worksheet = Nothing
        Dim sheets As Excel.Sheets = Nothing

        Try
            sheets = workbook.Worksheets
            Dim totalSheets As Integer = sheets.Count

            If totalSheets = 0 Then
                reportProgress?.Invoke(progressEnd)
                Return
            End If

            For i As Integer = 1 To totalSheets
                Dim sheet As Excel.Worksheet = Nothing
                Dim cell As Excel.Range = Nothing

                Try
                    sheet = CType(sheets.Item(i), Excel.Worksheet)

                    If i = 1 Then
                        firstSheet = sheet
                        sheet = Nothing
                    End If

                    Dim currentSheet As Excel.Worksheet =
                    If(i = 1, firstSheet, sheet)

                    Try
                        currentSheet.Activate()
                        cell = currentSheet.Range("A1")
                        cell.Select()
                    Catch
                        ' Continue if Excel cannot select A1 on this worksheet.
                    End Try

                Finally
                    ReleaseComObject(cell)
                    ReleaseComObject(sheet)
                End Try

                Dim progress As Integer =
                progressStart +
                CInt(Math.Floor(
                    CDbl(i) / totalSheets *
                    (progressEnd - progressStart)))

                reportProgress?.Invoke(Math.Min(progress, progressEnd))
            Next

            If firstSheet IsNot Nothing Then
                Try
                    firstSheet.Activate()
                Catch
                    ' Do not prevent saving if Excel cannot activate the first worksheet.
                End Try
            End If

        Finally
            ReleaseComObject(firstSheet)
            ReleaseComObject(sheets)
        End Try
    End Sub

    Public Shared Sub ReleaseComObject(obj As Object)
        Try
            If obj IsNot Nothing AndAlso Marshal.IsComObject(obj) Then
                Marshal.ReleaseComObject(obj)
            End If
        Catch
        End Try
    End Sub

End Class