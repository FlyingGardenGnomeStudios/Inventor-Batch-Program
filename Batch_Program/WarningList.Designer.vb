<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class WarningList
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
        Dim DataGridViewCellStyle1 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle()
        Me.btnOK = New System.Windows.Forms.Button()
        Me.btnCancel = New System.Windows.Forms.Button()
        Me.dgvWarning = New System.Windows.Forms.DataGridView()
        Me.Notification = New System.Windows.Forms.DataGridViewComboBoxColumn()
        Me.DataGridViewTextBoxColumn1 = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.DataGridViewTextBoxColumn2 = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.Warning = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.Context = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.DataGridViewTextBoxColumn3 = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.DataGridViewTextBoxColumn4 = New System.Windows.Forms.DataGridViewTextBoxColumn()
        CType(Me.dgvWarning, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'btnOK
        '
        Me.btnOK.Location = New System.Drawing.Point(598, 259)
        Me.btnOK.Name = "btnOK"
        Me.btnOK.Size = New System.Drawing.Size(75, 23)
        Me.btnOK.TabIndex = 1
        Me.btnOK.Text = "OK"
        Me.btnOK.UseVisualStyleBackColor = True
        '
        'btnCancel
        '
        Me.btnCancel.Location = New System.Drawing.Point(517, 259)
        Me.btnCancel.Name = "btnCancel"
        Me.btnCancel.Size = New System.Drawing.Size(75, 23)
        Me.btnCancel.TabIndex = 2
        Me.btnCancel.Text = "Cancel"
        Me.btnCancel.UseVisualStyleBackColor = True
        '
        'dgvWarning
        '
        Me.dgvWarning.AllowUserToAddRows = False
        Me.dgvWarning.AutoSizeColumnsMode = System.Windows.Forms.DataGridViewAutoSizeColumnsMode.Fill
        Me.dgvWarning.AutoSizeRowsMode = System.Windows.Forms.DataGridViewAutoSizeRowsMode.AllCells
        Me.dgvWarning.Columns.AddRange(New System.Windows.Forms.DataGridViewColumn() {Me.Warning, Me.Context, Me.Notification})
        Me.dgvWarning.EditMode = System.Windows.Forms.DataGridViewEditMode.EditOnEnter
        Me.dgvWarning.Location = New System.Drawing.Point(13, 12)
        Me.dgvWarning.Name = "dgvWarning"
        Me.dgvWarning.RowHeadersVisible = False
        DataGridViewCellStyle1.WrapMode = System.Windows.Forms.DataGridViewTriState.[True]
        Me.dgvWarning.RowsDefaultCellStyle = DataGridViewCellStyle1
        Me.dgvWarning.Size = New System.Drawing.Size(660, 224)
        Me.dgvWarning.TabIndex = 3
        '
        'Notification
        '
        Me.Notification.FillWeight = 30.0!
        Me.Notification.HeaderText = "Notification"
        Me.Notification.Items.AddRange(New Object() {"True", "False"})
        Me.Notification.Name = "Notification"
        '
        'DataGridViewTextBoxColumn1
        '
        Me.DataGridViewTextBoxColumn1.FillWeight = 30.0!
        Me.DataGridViewTextBoxColumn1.HeaderText = "Warning"
        Me.DataGridViewTextBoxColumn1.Name = "DataGridViewTextBoxColumn1"
        Me.DataGridViewTextBoxColumn1.ReadOnly = True
        Me.DataGridViewTextBoxColumn1.Width = 123
        '
        'DataGridViewTextBoxColumn2
        '
        Me.DataGridViewTextBoxColumn2.HeaderText = "Context"
        Me.DataGridViewTextBoxColumn2.Name = "DataGridViewTextBoxColumn2"
        Me.DataGridViewTextBoxColumn2.ReadOnly = True
        Me.DataGridViewTextBoxColumn2.Width = 412
        '
        'Warning
        '
        Me.Warning.FillWeight = 30.0!
        Me.Warning.HeaderText = "Warning"
        Me.Warning.Name = "Warning"
        Me.Warning.ReadOnly = True
        '
        'Context
        '
        Me.Context.HeaderText = "Context"
        Me.Context.Name = "Context"
        Me.Context.ReadOnly = True
        '
        'DataGridViewTextBoxColumn3
        '
        Me.DataGridViewTextBoxColumn3.FillWeight = 30.0!
        Me.DataGridViewTextBoxColumn3.HeaderText = "Column1"
        Me.DataGridViewTextBoxColumn3.Name = "DataGridViewTextBoxColumn3"
        Me.DataGridViewTextBoxColumn3.Width = 206
        '
        'DataGridViewTextBoxColumn4
        '
        Me.DataGridViewTextBoxColumn4.HeaderText = "Column2"
        Me.DataGridViewTextBoxColumn4.Name = "DataGridViewTextBoxColumn4"
        Me.DataGridViewTextBoxColumn4.Width = 205
        '
        'WarningList
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(685, 294)
        Me.Controls.Add(Me.dgvWarning)
        Me.Controls.Add(Me.btnCancel)
        Me.Controls.Add(Me.btnOK)
        Me.Name = "WarningList"
        Me.Text = "WarningList"
        CType(Me.dgvWarning, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)

    End Sub
    Friend WithEvents btnOK As Windows.Forms.Button
    Friend WithEvents btnCancel As Windows.Forms.Button
    Friend WithEvents DataGridViewTextBoxColumn1 As Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents DataGridViewTextBoxColumn2 As Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents dgvWarning As Windows.Forms.DataGridView
    Friend WithEvents DataGridViewTextBoxColumn3 As Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents DataGridViewTextBoxColumn4 As Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents Warning As Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents Context As Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents Notification As Windows.Forms.DataGridViewComboBoxColumn
End Class
