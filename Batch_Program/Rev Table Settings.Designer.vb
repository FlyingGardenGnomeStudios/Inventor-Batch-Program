<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class Rev_Table_Settings
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
        Me.components = New System.ComponentModel.Container()
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(Rev_Table_Settings))
        Me.btnOK = New System.Windows.Forms.Button()
        Me.btnCancel = New System.Windows.Forms.Button()
        Me.txtNumRev = New System.Windows.Forms.TextBox()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.txtAlphaRev = New System.Windows.Forms.TextBox()
        Me.GroupBox2 = New System.Windows.Forms.GroupBox()
        Me.nudStartVal = New System.Windows.Forms.NumericUpDown()
        Me.Label3 = New System.Windows.Forms.Label()
        Me.dgvRevTableLayout = New System.Windows.Forms.DataGridView()
        Me.chk = New System.Windows.Forms.DataGridViewCheckBoxColumn()
        Me.DataType = New System.Windows.Forms.DataGridViewComboBoxColumn()
        Me.ContextMenuStrip1 = New System.Windows.Forms.ContextMenuStrip(Me.components)
        Me.DeleteRowToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.GroupBox1 = New System.Windows.Forms.GroupBox()
        Me.rdbBR = New System.Windows.Forms.RadioButton()
        Me.rdbBL = New System.Windows.Forms.RadioButton()
        Me.rdbTR = New System.Windows.Forms.RadioButton()
        Me.PictureBox1 = New System.Windows.Forms.PictureBox()
        Me.rdbTL = New System.Windows.Forms.RadioButton()
        Me.Label4 = New System.Windows.Forms.Label()
        Me.DataGridViewTextBoxColumn1 = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.DataGridViewTextBoxColumn2 = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.Item = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.ColumnName = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.GroupBox2.SuspendLayout()
        CType(Me.nudStartVal, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.dgvRevTableLayout, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.ContextMenuStrip1.SuspendLayout()
        Me.GroupBox1.SuspendLayout()
        CType(Me.PictureBox1, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'btnOK
        '
        Me.btnOK.Location = New System.Drawing.Point(333, 385)
        Me.btnOK.Name = "btnOK"
        Me.btnOK.Size = New System.Drawing.Size(75, 23)
        Me.btnOK.TabIndex = 4
        Me.btnOK.Text = "OK"
        Me.btnOK.UseVisualStyleBackColor = True
        '
        'btnCancel
        '
        Me.btnCancel.Location = New System.Drawing.Point(252, 385)
        Me.btnCancel.Name = "btnCancel"
        Me.btnCancel.Size = New System.Drawing.Size(75, 23)
        Me.btnCancel.TabIndex = 5
        Me.btnCancel.Text = "Cancel"
        Me.btnCancel.UseVisualStyleBackColor = True
        '
        'txtNumRev
        '
        Me.txtNumRev.Location = New System.Drawing.Point(9, 32)
        Me.txtNumRev.Name = "txtNumRev"
        Me.txtNumRev.Size = New System.Drawing.Size(220, 20)
        Me.txtNumRev.TabIndex = 6
        Me.txtNumRev.Text = "APPROVED FOR MANUFACTURING"
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(6, 16)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(147, 13)
        Me.Label1.TabIndex = 7
        Me.Label1.Text = "Numerical revision description"
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Location = New System.Drawing.Point(6, 76)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(158, 13)
        Me.Label2.TabIndex = 9
        Me.Label2.Text = "Alphabetical revision description"
        '
        'txtAlphaRev
        '
        Me.txtAlphaRev.Location = New System.Drawing.Point(9, 92)
        Me.txtAlphaRev.Name = "txtAlphaRev"
        Me.txtAlphaRev.Size = New System.Drawing.Size(220, 20)
        Me.txtAlphaRev.TabIndex = 8
        Me.txtAlphaRev.Text = "ISSUED FOR REVIEW"
        '
        'GroupBox2
        '
        Me.GroupBox2.Controls.Add(Me.Label4)
        Me.GroupBox2.Controls.Add(Me.rdbBR)
        Me.GroupBox2.Controls.Add(Me.nudStartVal)
        Me.GroupBox2.Controls.Add(Me.rdbBL)
        Me.GroupBox2.Controls.Add(Me.Label3)
        Me.GroupBox2.Controls.Add(Me.rdbTR)
        Me.GroupBox2.Controls.Add(Me.Label1)
        Me.GroupBox2.Controls.Add(Me.PictureBox1)
        Me.GroupBox2.Controls.Add(Me.rdbTL)
        Me.GroupBox2.Controls.Add(Me.txtNumRev)
        Me.GroupBox2.Controls.Add(Me.Label2)
        Me.GroupBox2.Controls.Add(Me.txtAlphaRev)
        Me.GroupBox2.Location = New System.Drawing.Point(12, 254)
        Me.GroupBox2.Name = "GroupBox2"
        Me.GroupBox2.Size = New System.Drawing.Size(396, 125)
        Me.GroupBox2.TabIndex = 11
        Me.GroupBox2.TabStop = False
        Me.GroupBox2.Text = "Change Revision"
        '
        'nudStartVal
        '
        Me.nudStartVal.Location = New System.Drawing.Point(103, 53)
        Me.nudStartVal.Name = "nudStartVal"
        Me.nudStartVal.Size = New System.Drawing.Size(37, 20)
        Me.nudStartVal.TabIndex = 11
        '
        'Label3
        '
        Me.Label3.AutoSize = True
        Me.Label3.Location = New System.Drawing.Point(23, 55)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(73, 13)
        Me.Label3.TabIndex = 10
        Me.Label3.Text = "Starting Value"
        '
        'dgvRevTableLayout
        '
        Me.dgvRevTableLayout.AllowUserToResizeColumns = False
        Me.dgvRevTableLayout.AllowUserToResizeRows = False
        Me.dgvRevTableLayout.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.dgvRevTableLayout.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.dgvRevTableLayout.Columns.AddRange(New System.Windows.Forms.DataGridViewColumn() {Me.chk, Me.Item, Me.ColumnName, Me.DataType})
        Me.dgvRevTableLayout.ContextMenuStrip = Me.ContextMenuStrip1
        Me.dgvRevTableLayout.EditMode = System.Windows.Forms.DataGridViewEditMode.EditOnEnter
        Me.dgvRevTableLayout.Location = New System.Drawing.Point(6, 19)
        Me.dgvRevTableLayout.MultiSelect = False
        Me.dgvRevTableLayout.Name = "dgvRevTableLayout"
        Me.dgvRevTableLayout.RowHeadersVisible = False
        Me.dgvRevTableLayout.SelectionMode = System.Windows.Forms.DataGridViewSelectionMode.CellSelect
        Me.dgvRevTableLayout.Size = New System.Drawing.Size(384, 211)
        Me.dgvRevTableLayout.TabIndex = 13
        '
        'chk
        '
        Me.chk.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.AllCells
        Me.chk.HeaderText = ""
        Me.chk.Name = "chk"
        Me.chk.Width = 21
        '
        'DataType
        '
        Me.DataType.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.DisplayedCells
        Me.DataType.HeaderText = "DataType"
        Me.DataType.Items.AddRange(New Object() {"Date", "Number", "Text"})
        Me.DataType.MaxDropDownItems = 3
        Me.DataType.Name = "DataType"
        Me.DataType.Sorted = True
        Me.DataType.Width = 60
        '
        'ContextMenuStrip1
        '
        Me.ContextMenuStrip1.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.DeleteRowToolStripMenuItem})
        Me.ContextMenuStrip1.Name = "ContextMenuStrip1"
        Me.ContextMenuStrip1.Size = New System.Drawing.Size(134, 26)
        '
        'DeleteRowToolStripMenuItem
        '
        Me.DeleteRowToolStripMenuItem.Name = "DeleteRowToolStripMenuItem"
        Me.DeleteRowToolStripMenuItem.Size = New System.Drawing.Size(133, 22)
        Me.DeleteRowToolStripMenuItem.Text = "Delete Row"
        '
        'GroupBox1
        '
        Me.GroupBox1.Controls.Add(Me.dgvRevTableLayout)
        Me.GroupBox1.Location = New System.Drawing.Point(12, 12)
        Me.GroupBox1.Name = "GroupBox1"
        Me.GroupBox1.Size = New System.Drawing.Size(396, 236)
        Me.GroupBox1.TabIndex = 14
        Me.GroupBox1.TabStop = False
        Me.GroupBox1.Text = "Check Items"
        '
        'rdbBR
        '
        Me.rdbBR.AutoSize = True
        Me.rdbBR.Location = New System.Drawing.Point(334, 92)
        Me.rdbBR.Name = "rdbBR"
        Me.rdbBR.Size = New System.Drawing.Size(14, 13)
        Me.rdbBR.TabIndex = 63
        Me.rdbBR.UseVisualStyleBackColor = True
        '
        'rdbBL
        '
        Me.rdbBL.AutoSize = True
        Me.rdbBL.Location = New System.Drawing.Point(284, 92)
        Me.rdbBL.Name = "rdbBL"
        Me.rdbBL.Size = New System.Drawing.Size(14, 13)
        Me.rdbBL.TabIndex = 62
        Me.rdbBL.UseVisualStyleBackColor = True
        '
        'rdbTR
        '
        Me.rdbTR.AutoSize = True
        Me.rdbTR.Location = New System.Drawing.Point(334, 38)
        Me.rdbTR.Name = "rdbTR"
        Me.rdbTR.Size = New System.Drawing.Size(14, 13)
        Me.rdbTR.TabIndex = 61
        Me.rdbTR.UseVisualStyleBackColor = True
        '
        'PictureBox1
        '
        Me.PictureBox1.ErrorImage = Nothing
        Me.PictureBox1.Image = CType(resources.GetObject("PictureBox1.Image"), System.Drawing.Image)
        Me.PictureBox1.InitialImage = Nothing
        Me.PictureBox1.Location = New System.Drawing.Point(297, 55)
        Me.PictureBox1.Name = "PictureBox1"
        Me.PictureBox1.Size = New System.Drawing.Size(36, 34)
        Me.PictureBox1.TabIndex = 60
        Me.PictureBox1.TabStop = False
        Me.PictureBox1.WaitOnLoad = True
        '
        'rdbTL
        '
        Me.rdbTL.AutoSize = True
        Me.rdbTL.Location = New System.Drawing.Point(284, 38)
        Me.rdbTL.Name = "rdbTL"
        Me.rdbTL.Size = New System.Drawing.Size(14, 13)
        Me.rdbTL.TabIndex = 59
        Me.rdbTL.UseVisualStyleBackColor = True
        '
        'Label4
        '
        Me.Label4.AutoSize = True
        Me.Label4.Location = New System.Drawing.Point(242, 16)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(138, 13)
        Me.Label4.TabIndex = 64
        Me.Label4.Text = "Default Rev Table Location"
        '
        'DataGridViewTextBoxColumn1
        '
        Me.DataGridViewTextBoxColumn1.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.Fill
        Me.DataGridViewTextBoxColumn1.FillWeight = 34.04429!
        Me.DataGridViewTextBoxColumn1.HeaderText = "Revision Item"
        Me.DataGridViewTextBoxColumn1.Name = "DataGridViewTextBoxColumn1"
        '
        'DataGridViewTextBoxColumn2
        '
        Me.DataGridViewTextBoxColumn2.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.Fill
        Me.DataGridViewTextBoxColumn2.FillWeight = 53.71433!
        Me.DataGridViewTextBoxColumn2.HeaderText = "Rev Table Column Name"
        Me.DataGridViewTextBoxColumn2.Name = "DataGridViewTextBoxColumn2"
        '
        'Item
        '
        Me.Item.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.Fill
        Me.Item.FillWeight = 34.04429!
        Me.Item.HeaderText = "Revision Item"
        Me.Item.Name = "Item"
        '
        'ColumnName
        '
        Me.ColumnName.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.Fill
        Me.ColumnName.FillWeight = 53.71433!
        Me.ColumnName.HeaderText = "Rev Table Column Name"
        Me.ColumnName.Name = "ColumnName"
        '
        'Rev_Table_Settings
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(420, 413)
        Me.Controls.Add(Me.GroupBox1)
        Me.Controls.Add(Me.GroupBox2)
        Me.Controls.Add(Me.btnCancel)
        Me.Controls.Add(Me.btnOK)
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.Name = "Rev_Table_Settings"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterParent
        Me.Text = "Rev_Table_Settings"
        Me.GroupBox2.ResumeLayout(False)
        Me.GroupBox2.PerformLayout()
        CType(Me.nudStartVal, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.dgvRevTableLayout, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ContextMenuStrip1.ResumeLayout(False)
        Me.GroupBox1.ResumeLayout(False)
        CType(Me.PictureBox1, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)

    End Sub
    Friend WithEvents btnOK As Windows.Forms.Button
    Friend WithEvents btnCancel As Windows.Forms.Button
    Friend WithEvents txtNumRev As Windows.Forms.TextBox
    Friend WithEvents Label1 As Windows.Forms.Label
    Friend WithEvents Label2 As Windows.Forms.Label
    Friend WithEvents txtAlphaRev As Windows.Forms.TextBox
    Friend WithEvents GroupBox2 As Windows.Forms.GroupBox
    Friend WithEvents dgvRevTableLayout As Windows.Forms.DataGridView
    Friend WithEvents ContextMenuStrip1 As Windows.Forms.ContextMenuStrip
    Friend WithEvents DeleteRowToolStripMenuItem As Windows.Forms.ToolStripMenuItem
    Friend WithEvents GroupBox1 As Windows.Forms.GroupBox
    Friend WithEvents chk As Windows.Forms.DataGridViewCheckBoxColumn
    Friend WithEvents Item As Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColumnName As Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents DataType As Windows.Forms.DataGridViewComboBoxColumn
    Friend WithEvents DataGridViewTextBoxColumn1 As Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents DataGridViewTextBoxColumn2 As Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents nudStartVal As Windows.Forms.NumericUpDown
    Friend WithEvents Label3 As Windows.Forms.Label
    Friend WithEvents Label4 As Windows.Forms.Label
    Friend WithEvents rdbBR As Windows.Forms.RadioButton
    Friend WithEvents rdbBL As Windows.Forms.RadioButton
    Friend WithEvents rdbTR As Windows.Forms.RadioButton
    Friend WithEvents PictureBox1 As Windows.Forms.PictureBox
    Friend WithEvents rdbTL As Windows.Forms.RadioButton
End Class
