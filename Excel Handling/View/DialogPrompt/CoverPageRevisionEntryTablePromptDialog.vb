Imports System.Windows.Forms

Public Class CoverPageRevisionEntryTablePromptDialog

    Public Sub New(suggestedDescription As String, suggestedEngineer As String, suggestedRevision As Object)
        ' Required by designer
        InitializeComponent()

        txtRevision.Text = If(TypeOf suggestedRevision Is String, suggestedRevision.ToString(),
                           If(TypeOf suggestedRevision Is Char, suggestedRevision.ToString(),
                           If(TypeOf suggestedRevision Is Single, suggestedRevision.ToString(),
                           "A"))) ' fallback
        ' Populate suggested values
        txtRevision.Text = suggestedRevision
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

End Class
