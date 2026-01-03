<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class CoverPageRevisionEntryTablePromptDialog
    Inherits System.Windows.Forms.Form

    'Form overrides dispose to clean up the component list.
    <System.Diagnostics.DebuggerNonUserCode()> _
    Protected Overrides Sub Dispose(ByVal disposing As Boolean)
        Try
            If disposing AndAlso components IsNot Nothing Then
                components.Dispose()
            End If
        Finally
            MyBase.Dispose(disposing)
        End Try
    End Sub

    'Required by the Windows Form Designer
    Private components As System.ComponentModel.IContainer

    'NOTE: The following procedure is required by the Windows Form Designer
    'It can be modified using the Windows Form Designer.  
    'Do not modify it using the code editor.
    <System.Diagnostics.DebuggerStepThrough()> _
    Private Sub InitializeComponent()
        Me.txtRevision = New System.Windows.Forms.TextBox()
        Me.txtDescription = New System.Windows.Forms.TextBox()
        Me.txtEngineer = New System.Windows.Forms.TextBox()
        Me.lblRevision = New System.Windows.Forms.Label()
        Me.lblDescription = New System.Windows.Forms.Label()
        Me.lblEngineer = New System.Windows.Forms.Label()
        Me.btnCancel = New System.Windows.Forms.Button()
        Me.btnConfirm = New System.Windows.Forms.Button()
        Me.layoutRoot = New System.Windows.Forms.TableLayoutPanel()
        Me.grpEntryDetails = New System.Windows.Forms.GroupBox()
        Me.layoutRoot.SuspendLayout()
        Me.grpEntryDetails.SuspendLayout()
        Me.SuspendLayout()
        '
        'txtRevision
        '
        Me.txtRevision.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.txtRevision.Location = New System.Drawing.Point(84, 32)
        Me.txtRevision.Margin = New System.Windows.Forms.Padding(4, 3, 4, 3)
        Me.txtRevision.Name = "txtRevision"
        Me.txtRevision.Size = New System.Drawing.Size(308, 23)
        Me.txtRevision.TabIndex = 4
        '
        'txtDescription
        '
        Me.txtDescription.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.txtDescription.Location = New System.Drawing.Point(84, 62)
        Me.txtDescription.Margin = New System.Windows.Forms.Padding(4, 3, 4, 3)
        Me.txtDescription.Name = "txtDescription"
        Me.txtDescription.Size = New System.Drawing.Size(308, 23)
        Me.txtDescription.TabIndex = 5
        '
        'txtEngineer
        '
        Me.txtEngineer.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.txtEngineer.Location = New System.Drawing.Point(84, 92)
        Me.txtEngineer.Margin = New System.Windows.Forms.Padding(4, 3, 4, 3)
        Me.txtEngineer.Name = "txtEngineer"
        Me.txtEngineer.Size = New System.Drawing.Size(308, 23)
        Me.txtEngineer.TabIndex = 6
        '
        'lblRevision
        '
        Me.lblRevision.Location = New System.Drawing.Point(7, 36)
        Me.lblRevision.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.lblRevision.Name = "lblRevision"
        Me.lblRevision.RightToLeft = System.Windows.Forms.RightToLeft.Yes
        Me.lblRevision.Size = New System.Drawing.Size(69, 15)
        Me.lblRevision.TabIndex = 7
        Me.lblRevision.Text = "Revision"
        '
        'lblDescription
        '
        Me.lblDescription.Location = New System.Drawing.Point(7, 66)
        Me.lblDescription.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.lblDescription.Name = "lblDescription"
        Me.lblDescription.RightToLeft = System.Windows.Forms.RightToLeft.Yes
        Me.lblDescription.Size = New System.Drawing.Size(69, 15)
        Me.lblDescription.TabIndex = 8
        Me.lblDescription.Text = "Description"
        '
        'lblEngineer
        '
        Me.lblEngineer.Location = New System.Drawing.Point(7, 96)
        Me.lblEngineer.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.lblEngineer.Name = "lblEngineer"
        Me.lblEngineer.RightToLeft = System.Windows.Forms.RightToLeft.Yes
        Me.lblEngineer.Size = New System.Drawing.Size(69, 15)
        Me.lblEngineer.TabIndex = 9
        Me.lblEngineer.Text = "Tested By"
        '
        'btnCancel
        '
        Me.btnCancel.AutoSize = True
        Me.layoutRoot.SetColumnSpan(Me.btnCancel, 3)
        Me.btnCancel.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.btnCancel.Dock = System.Windows.Forms.DockStyle.Right
        Me.btnCancel.Font = New System.Drawing.Font("Calibri", 12.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnCancel.Location = New System.Drawing.Point(153, 158)
        Me.btnCancel.Margin = New System.Windows.Forms.Padding(4, 3, 4, 3)
        Me.btnCancel.Name = "btnCancel"
        Me.btnCancel.Size = New System.Drawing.Size(88, 29)
        Me.btnCancel.TabIndex = 10
        Me.btnCancel.Text = "Cancel"
        Me.btnCancel.UseVisualStyleBackColor = True
        '
        'btnConfirm
        '
        Me.btnConfirm.AutoSize = True
        Me.layoutRoot.SetColumnSpan(Me.btnConfirm, 3)
        Me.btnConfirm.Dock = System.Windows.Forms.DockStyle.Left
        Me.btnConfirm.Font = New System.Drawing.Font("Calibri", 12.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnConfirm.Location = New System.Drawing.Point(249, 158)
        Me.btnConfirm.Margin = New System.Windows.Forms.Padding(4, 3, 4, 3)
        Me.btnConfirm.Name = "btnConfirm"
        Me.btnConfirm.Size = New System.Drawing.Size(88, 29)
        Me.btnConfirm.TabIndex = 11
        Me.btnConfirm.Text = "Confirm"
        Me.btnConfirm.UseVisualStyleBackColor = True
        '
        'layoutRoot
        '
        Me.layoutRoot.ColumnCount = 6
        Me.layoutRoot.ColumnStyles.Add(New System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Absolute, 41.0!))
        Me.layoutRoot.ColumnStyles.Add(New System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 25.0!))
        Me.layoutRoot.ColumnStyles.Add(New System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 25.0!))
        Me.layoutRoot.ColumnStyles.Add(New System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 25.0!))
        Me.layoutRoot.ColumnStyles.Add(New System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 25.0!))
        Me.layoutRoot.ColumnStyles.Add(New System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Absolute, 44.0!))
        Me.layoutRoot.Controls.Add(Me.grpEntryDetails, 1, 1)
        Me.layoutRoot.Controls.Add(Me.btnCancel, 0, 3)
        Me.layoutRoot.Controls.Add(Me.btnConfirm, 3, 3)
        Me.layoutRoot.Dock = System.Windows.Forms.DockStyle.Fill
        Me.layoutRoot.Location = New System.Drawing.Point(0, 0)
        Me.layoutRoot.Margin = New System.Windows.Forms.Padding(4, 3, 4, 3)
        Me.layoutRoot.Name = "layoutRoot"
        Me.layoutRoot.RowCount = 5
        Me.layoutRoot.RowStyles.Add(New System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 33.33333!))
        Me.layoutRoot.RowStyles.Add(New System.Windows.Forms.RowStyle())
        Me.layoutRoot.RowStyles.Add(New System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 33.33333!))
        Me.layoutRoot.RowStyles.Add(New System.Windows.Forms.RowStyle())
        Me.layoutRoot.RowStyles.Add(New System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 33.33333!))
        Me.layoutRoot.RowStyles.Add(New System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Absolute, 20.0!))
        Me.layoutRoot.RowStyles.Add(New System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Absolute, 20.0!))
        Me.layoutRoot.Size = New System.Drawing.Size(493, 203)
        Me.layoutRoot.TabIndex = 12
        '
        'grpEntryDetails
        '
        Me.layoutRoot.SetColumnSpan(Me.grpEntryDetails, 4)
        Me.grpEntryDetails.Controls.Add(Me.txtEngineer)
        Me.grpEntryDetails.Controls.Add(Me.txtRevision)
        Me.grpEntryDetails.Controls.Add(Me.txtDescription)
        Me.grpEntryDetails.Controls.Add(Me.lblRevision)
        Me.grpEntryDetails.Controls.Add(Me.lblEngineer)
        Me.grpEntryDetails.Controls.Add(Me.lblDescription)
        Me.grpEntryDetails.Dock = System.Windows.Forms.DockStyle.Fill
        Me.grpEntryDetails.Location = New System.Drawing.Point(45, 14)
        Me.grpEntryDetails.Margin = New System.Windows.Forms.Padding(4, 3, 4, 3)
        Me.grpEntryDetails.Name = "grpEntryDetails"
        Me.grpEntryDetails.Padding = New System.Windows.Forms.Padding(4, 3, 4, 3)
        Me.grpEntryDetails.Size = New System.Drawing.Size(400, 127)
        Me.grpEntryDetails.TabIndex = 0
        Me.grpEntryDetails.TabStop = False
        Me.grpEntryDetails.Text = "Revision Entry Details"
        '
        'CoverPageRevisionEntryTablePromptDialog
        '
        Me.AcceptButton = Me.btnConfirm
        Me.AutoScaleDimensions = New System.Drawing.SizeF(7.0!, 15.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.CancelButton = Me.btnCancel
        Me.ClientSize = New System.Drawing.Size(493, 203)
        Me.ControlBox = False
        Me.Controls.Add(Me.layoutRoot)
        Me.Font = New System.Drawing.Font("Calibri", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle
        Me.Margin = New System.Windows.Forms.Padding(4, 3, 4, 3)
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.Name = "CoverPageRevisionEntryTablePromptDialog"
        Me.SizeGripStyle = System.Windows.Forms.SizeGripStyle.Hide
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterParent
        Me.Text = "Update Auto-DVT Report Overview Page"
        Me.TopMost = True
        Me.layoutRoot.ResumeLayout(False)
        Me.layoutRoot.PerformLayout()
        Me.grpEntryDetails.ResumeLayout(False)
        Me.grpEntryDetails.PerformLayout()
        Me.ResumeLayout(False)

    End Sub
    Friend WithEvents txtRevision As Windows.Forms.TextBox
    Friend WithEvents txtDescription As Windows.Forms.TextBox
    Friend WithEvents txtEngineer As Windows.Forms.TextBox
    Friend WithEvents lblRevision As Windows.Forms.Label
    Friend WithEvents lblDescription As Windows.Forms.Label
    Friend WithEvents lblEngineer As Windows.Forms.Label
    Friend WithEvents btnCancel As Windows.Forms.Button
    Friend WithEvents btnConfirm As Windows.Forms.Button
    Friend WithEvents layoutRoot As Windows.Forms.TableLayoutPanel
    Friend WithEvents grpEntryDetails As Windows.Forms.GroupBox
End Class
