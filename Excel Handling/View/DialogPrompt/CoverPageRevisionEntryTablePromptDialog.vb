Imports System.Windows.Forms
Imports System.Runtime.InteropServices

Public Class CoverPageRevisionEntryTablePromptDialog

    ' --- Windows API constants for forcing topmost ---
    Private Const SWP_NOSIZE As UInteger = &H1
    Private Const SWP_NOMOVE As UInteger = &H2
    Private Const SWP_SHOWWINDOW As UInteger = &H40
    Private Shared ReadOnly HWND_TOPMOST As IntPtr = New IntPtr(-1)

    <DllImport("user32.dll")>
    Private Shared Function SetWindowPos(hWnd As IntPtr, hWndInsertAfter As IntPtr,
                                         X As Integer, Y As Integer, cx As Integer, cy As Integer,
                                         uFlags As UInteger) As Boolean
    End Function

    Public Sub New(suggestedDescription As String, suggestedEngineer As String, suggestedRevision As Object)
        ' Required by designer
        InitializeComponent()

        ' Populate suggested values with type checks and fallback
        txtRevision.Text = If(TypeOf suggestedRevision Is String, suggestedRevision.ToString(),
                           If(TypeOf suggestedRevision Is Char, suggestedRevision.ToString(),
                           If(TypeOf suggestedRevision Is Single, suggestedRevision.ToString(),
                           "A"))) ' fallback
        txtDescription.Text = suggestedDescription
        txtEngineer.Text = suggestedEngineer

        ' UX defaults
        Me.AcceptButton = btnConfirm
        Me.CancelButton = btnCancel
    End Sub

#Region "Public Result Properties"
    Public ReadOnly Property Revision As String
        Get
            Return txtRevision.Text.Trim()
        End Get
    End Property

    Public ReadOnly Property Description As String
        Get
            Return txtDescription.Text.Trim()
        End Get
    End Property

    Public ReadOnly Property Engineer As String
        Get
            Return txtEngineer.Text.Trim()
        End Get
    End Property
#End Region

#Region "Event Handlers"
    Private Sub btnConfirm_Click(sender As Object, e As EventArgs) Handles btnConfirm.Click
        Me.DialogResult = DialogResult.OK
        Me.Close()
    End Sub

    Private Sub btnCancel_Click(sender As Object, e As EventArgs) Handles btnCancel.Click
        Me.DialogResult = DialogResult.Cancel
        Me.Close()
    End Sub
#End Region

#Region "Always On Top"
    ' Add this handler to your form (can do in designer or manually)
    Private Sub CoverPageRevisionEntryTablePromptDialog_Shown(sender As Object, e As EventArgs) Handles Me.Shown
        ' Basic TopMost + BringToFront
        Me.TopMost = True
        Me.BringToFront()
        Me.Activate()

        ' Force topmost even if called from backend thread or async
        SetWindowPos(Me.Handle, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE Or SWP_SHOWWINDOW)
    End Sub
#End Region

End Class
