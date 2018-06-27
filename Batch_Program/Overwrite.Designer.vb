<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class Overwrite
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
        Me.Label1 = New System.Windows.Forms.Label()
        Me.dgvOverwrite = New System.Windows.Forms.DataGridView()
        Me.chkOverwrite = New System.Windows.Forms.DataGridViewCheckBoxColumn()
        Me.Total = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.Counter = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.Destin = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.DrawingName = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.DrawSource = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.Type = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.Rev = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.btnOK = New System.Windows.Forms.Button()
        Me.btnCancel = New System.Windows.Forms.Button()
        Me.DataGridViewTextBoxColumn1 = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.DataGridViewTextBoxColumn2 = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.DataGridViewTextBoxColumn3 = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.DataGridViewTextBoxColumn4 = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.DataGridViewTextBoxColumn5 = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.DataGridViewTextBoxColumn6 = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.DataGridViewTextBoxColumn7 = New System.Windows.Forms.DataGridViewTextBoxColumn()
        CType(Me.dgvOverwrite, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(12, 9)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(181, 13)
        Me.Label1.TabIndex = 0
        Me.Label1.Text = "The following files will be overwritten:"
        '
        'dgvOverwrite
        '
        Me.dgvOverwrite.AllowUserToAddRows = False
        Me.dgvOverwrite.AllowUserToDeleteRows = False
        Me.dgvOverwrite.AllowUserToResizeRows = False
        Me.dgvOverwrite.AutoSizeColumnsMode = System.Windows.Forms.DataGridViewAutoSizeColumnsMode.Fill
        Me.dgvOverwrite.BackgroundColor = System.Drawing.SystemColors.Window
        Me.dgvOverwrite.CellBorderStyle = System.Windows.Forms.DataGridViewCellBorderStyle.None
        Me.dgvOverwrite.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.dgvOverwrite.Columns.AddRange(New System.Windows.Forms.DataGridViewColumn() {Me.chkOverwrite, Me.Total, Me.Counter, Me.Destin, Me.DrawingName, Me.DrawSource, Me.Type, Me.Rev})
        Me.dgvOverwrite.Location = New System.Drawing.Point(12, 25)
        Me.dgvOverwrite.Name = "dgvOverwrite"
        Me.dgvOverwrite.RowHeadersVisible = False
        Me.dgvOverwrite.Size = New System.Drawing.Size(444, 150)
        Me.dgvOverwrite.TabIndex = 1
        '
        'chkOverwrite
        '
        Me.chkOverwrite.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.ColumnHeader
        Me.chkOverwrite.FalseValue = "False"
        Me.chkOverwrite.FillWeight = 47.09408!
        Me.chkOverwrite.HeaderText = ""
        Me.chkOverwrite.Name = "chkOverwrite"
        Me.chkOverwrite.TrueValue = "True"
        Me.chkOverwrite.Width = 5
        '
        'Total
        '
        Me.Total.HeaderText = "Total"
        Me.Total.Name = "Total"
        Me.Total.Visible = False
        '
        'Counter
        '
        Me.Counter.HeaderText = "Counter"
        Me.Counter.Name = "Counter"
        Me.Counter.Visible = False
        '
        'Destin
        '
        Me.Destin.HeaderText = "Destin"
        Me.Destin.Name = "Destin"
        Me.Destin.ReadOnly = True
        Me.Destin.Visible = False
        '
        'DrawingName
        '
        Me.DrawingName.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.Fill
        Me.DrawingName.FillWeight = 131.0785!
        Me.DrawingName.HeaderText = "DrawingName"
        Me.DrawingName.Name = "DrawingName"
        '
        'DrawSource
        '
        Me.DrawSource.HeaderText = "DrawSource"
        Me.DrawSource.Name = "DrawSource"
        Me.DrawSource.ReadOnly = True
        Me.DrawSource.Visible = False
        '
        'Type
        '
        Me.Type.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.ColumnHeader
        Me.Type.FillWeight = 121.8274!
        Me.Type.HeaderText = "Export Type"
        Me.Type.Name = "Type"
        Me.Type.ReadOnly = True
        Me.Type.Width = 89
        '
        'Rev
        '
        Me.Rev.HeaderText = "Rev"
        Me.Rev.Name = "Rev"
        Me.Rev.ReadOnly = True
        Me.Rev.Visible = False
        '
        'btnOK
        '
        Me.btnOK.Location = New System.Drawing.Point(381, 181)
        Me.btnOK.Name = "btnOK"
        Me.btnOK.Size = New System.Drawing.Size(75, 23)
        Me.btnOK.TabIndex = 2
        Me.btnOK.Text = "OK"
        Me.btnOK.UseVisualStyleBackColor = True
        '
        'btnCancel
        '
        Me.btnCancel.Location = New System.Drawing.Point(300, 181)
        Me.btnCancel.Name = "btnCancel"
        Me.btnCancel.Size = New System.Drawing.Size(75, 23)
        Me.btnCancel.TabIndex = 3
        Me.btnCancel.Text = "Cancel"
        Me.btnCancel.UseVisualStyleBackColor = True
        '
        'DataGridViewTextBoxColumn1
        '
        Me.DataGridViewTextBoxColumn1.HeaderText = "Destin"
        Me.DataGridViewTextBoxColumn1.Name = "DataGridViewTextBoxColumn1"
        Me.DataGridViewTextBoxColumn1.ReadOnly = True
        Me.DataGridViewTextBoxColumn1.Width = 5
        '
        'DataGridViewTextBoxColumn2
        '
        Me.DataGridViewTextBoxColumn2.HeaderText = "DrawSource"
        Me.DataGridViewTextBoxColumn2.Name = "DataGridViewTextBoxColumn2"
        Me.DataGridViewTextBoxColumn2.ReadOnly = True
        Me.DataGridViewTextBoxColumn2.Width = 5
        '
        'DataGridViewTextBoxColumn3
        '
        Me.DataGridViewTextBoxColumn3.HeaderText = "Total"
        Me.DataGridViewTextBoxColumn3.Name = "DataGridViewTextBoxColumn3"
        Me.DataGridViewTextBoxColumn3.ReadOnly = True
        Me.DataGridViewTextBoxColumn3.Width = 5
        '
        'DataGridViewTextBoxColumn4
        '
        Me.DataGridViewTextBoxColumn4.HeaderText = "Counter"
        Me.DataGridViewTextBoxColumn4.Name = "DataGridViewTextBoxColumn4"
        Me.DataGridViewTextBoxColumn4.ReadOnly = True
        Me.DataGridViewTextBoxColumn4.Width = 5
        '
        'DataGridViewTextBoxColumn5
        '
        Me.DataGridViewTextBoxColumn5.HeaderText = "Type"
        Me.DataGridViewTextBoxColumn5.Name = "DataGridViewTextBoxColumn5"
        Me.DataGridViewTextBoxColumn5.ReadOnly = True
        Me.DataGridViewTextBoxColumn5.Width = 5
        '
        'DataGridViewTextBoxColumn6
        '
        Me.DataGridViewTextBoxColumn6.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.Fill
        Me.DataGridViewTextBoxColumn6.HeaderText = "Name"
        Me.DataGridViewTextBoxColumn6.Name = "DataGridViewTextBoxColumn6"
        Me.DataGridViewTextBoxColumn6.ReadOnly = True
        '
        'DataGridViewTextBoxColumn7
        '
        Me.DataGridViewTextBoxColumn7.HeaderText = "Location"
        Me.DataGridViewTextBoxColumn7.Name = "DataGridViewTextBoxColumn7"
        Me.DataGridViewTextBoxColumn7.ReadOnly = True
        Me.DataGridViewTextBoxColumn7.Width = 5
        '
        'Overwrite
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(468, 214)
        Me.Controls.Add(Me.btnCancel)
        Me.Controls.Add(Me.btnOK)
        Me.Controls.Add(Me.dgvOverwrite)
        Me.Controls.Add(Me.Label1)
        Me.Name = "Overwrite"
        Me.Text = "Overwrite"
        CType(Me.dgvOverwrite, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub

    Friend WithEvents Label1 As Windows.Forms.Label
    Friend WithEvents dgvOverwrite As Windows.Forms.DataGridView
    Friend WithEvents btnOK As Windows.Forms.Button
    Friend WithEvents btnCancel As Windows.Forms.Button
    Friend WithEvents DataGridViewTextBoxColumn1 As Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents DataGridViewTextBoxColumn2 As Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents DataGridViewTextBoxColumn3 As Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents DataGridViewTextBoxColumn4 As Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents DataGridViewTextBoxColumn5 As Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents DataGridViewTextBoxColumn6 As Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents DataGridViewTextBoxColumn7 As Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents chkOverwrite As Windows.Forms.DataGridViewCheckBoxColumn
    Friend WithEvents Total As Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents Counter As Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents Destin As Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents DrawingName As Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents DrawSource As Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents Type As Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents Rev As Windows.Forms.DataGridViewTextBoxColumn
End Class
