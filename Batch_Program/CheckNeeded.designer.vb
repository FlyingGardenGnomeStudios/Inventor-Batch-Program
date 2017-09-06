<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class CheckNeeded
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
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(CheckNeeded))
        Me.DateTimePicker1 = New System.Windows.Forms.DateTimePicker()
        Me.btnCancel = New System.Windows.Forms.Button()
        Me.btnIgnore = New System.Windows.Forms.Button()
        Me.btnOK = New System.Windows.Forms.Button()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.lstCheckNeeded = New System.Windows.Forms.ListView()
        Me.DrawingNumber = CType(New System.Windows.Forms.ColumnHeader(), System.Windows.Forms.ColumnHeader)
        Me.CheckedBy = CType(New System.Windows.Forms.ColumnHeader(), System.Windows.Forms.ColumnHeader)
        Me.DateChecked = CType(New System.Windows.Forms.ColumnHeader(), System.Windows.Forms.ColumnHeader)
        Me.Initials = CType(New System.Windows.Forms.ColumnHeader(), System.Windows.Forms.ColumnHeader)
        Me.RevCheck = CType(New System.Windows.Forms.ColumnHeader(), System.Windows.Forms.ColumnHeader)
        Me.RevDate = CType(New System.Windows.Forms.ColumnHeader(), System.Windows.Forms.ColumnHeader)
        Me.More = CType(New System.Windows.Forms.ColumnHeader(), System.Windows.Forms.ColumnHeader)
        Me.TextBox1 = New System.Windows.Forms.TextBox()
        Me.ToolTip1 = New System.Windows.Forms.ToolTip(Me.components)
        Me.btnApplytoall = New System.Windows.Forms.Button()
        Me.ProgressBar1 = New System.Windows.Forms.ProgressBar()
        Me.SuspendLayout()
        '
        'DateTimePicker1
        '
        Me.DateTimePicker1.AllowDrop = True
        Me.DateTimePicker1.Checked = False
        Me.DateTimePicker1.Location = New System.Drawing.Point(390, 9)
        Me.DateTimePicker1.Name = "DateTimePicker1"
        Me.DateTimePicker1.ShowCheckBox = True
        Me.DateTimePicker1.Size = New System.Drawing.Size(170, 20)
        Me.DateTimePicker1.TabIndex = 17
        Me.DateTimePicker1.Visible = False
        '
        'btnCancel
        '
        Me.btnCancel.Location = New System.Drawing.Point(93, 183)
        Me.btnCancel.Name = "btnCancel"
        Me.btnCancel.Size = New System.Drawing.Size(75, 23)
        Me.btnCancel.TabIndex = 14
        Me.btnCancel.Text = "Cancel"
        Me.ToolTip1.SetToolTip(Me.btnCancel, "Exit without making any changes")
        Me.btnCancel.UseVisualStyleBackColor = True
        '
        'btnIgnore
        '
        Me.btnIgnore.Location = New System.Drawing.Point(174, 183)
        Me.btnIgnore.Name = "btnIgnore"
        Me.btnIgnore.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.btnIgnore.Size = New System.Drawing.Size(75, 23)
        Me.btnIgnore.TabIndex = 13
        Me.btnIgnore.Text = "Ignore"
        Me.ToolTip1.SetToolTip(Me.btnIgnore, "Disregard changes and proceed with operation")
        Me.btnIgnore.UseVisualStyleBackColor = True
        '
        'btnOK
        '
        Me.btnOK.Location = New System.Drawing.Point(12, 183)
        Me.btnOK.Name = "btnOK"
        Me.btnOK.Size = New System.Drawing.Size(75, 23)
        Me.btnOK.TabIndex = 12
        Me.btnOK.Text = "OK"
        Me.ToolTip1.SetToolTip(Me.btnOK, "Proceed with changes")
        Me.btnOK.UseVisualStyleBackColor = True
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Location = New System.Drawing.Point(12, 21)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(251, 13)
        Me.Label2.TabIndex = 11
        Me.Label2.Text = "Click to edit. Double-click Drawing Number to open."
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(12, 9)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(211, 13)
        Me.Label1.TabIndex = 10
        Me.Label1.Text = "The following files have not been checked:"
        '
        'lstCheckNeeded
        '
        Me.lstCheckNeeded.Columns.AddRange(New System.Windows.Forms.ColumnHeader() {Me.DrawingNumber, Me.CheckedBy, Me.DateChecked, Me.Initials, Me.RevCheck, Me.RevDate, Me.More})
        Me.lstCheckNeeded.FullRowSelect = True
        Me.lstCheckNeeded.GridLines = True
        Me.lstCheckNeeded.HeaderStyle = System.Windows.Forms.ColumnHeaderStyle.Nonclickable
        Me.lstCheckNeeded.LabelWrap = False
        Me.lstCheckNeeded.Location = New System.Drawing.Point(12, 42)
        Me.lstCheckNeeded.MaximumSize = New System.Drawing.Size(600, 135)
        Me.lstCheckNeeded.MultiSelect = False
        Me.lstCheckNeeded.Name = "lstCheckNeeded"
        Me.lstCheckNeeded.Size = New System.Drawing.Size(600, 135)
        Me.lstCheckNeeded.Sorting = System.Windows.Forms.SortOrder.Ascending
        Me.lstCheckNeeded.TabIndex = 18
        Me.lstCheckNeeded.UseCompatibleStateImageBehavior = False
        Me.lstCheckNeeded.View = System.Windows.Forms.View.Details
        '
        'DrawingNumber
        '
        Me.DrawingNumber.Text = "Drawing Number:"
        Me.DrawingNumber.Width = 100
        '
        'CheckedBy
        '
        Me.CheckedBy.Text = "Checked By:"
        Me.CheckedBy.Width = 85
        '
        'DateChecked
        '
        Me.DateChecked.Text = "Date Checked:"
        Me.DateChecked.Width = 90
        '
        'Initials
        '
        Me.Initials.Text = "Rev Initials:"
        Me.Initials.Width = 75
        '
        'RevCheck
        '
        Me.RevCheck.Text = "Rev Check By:"
        Me.RevCheck.Width = 90
        '
        'RevDate
        '
        Me.RevDate.Text = "Rev Date:"
        Me.RevDate.Width = 90
        '
        'More
        '
        Me.More.Text = "More:"
        Me.More.Width = 40
        '
        'TextBox1
        '
        Me.TextBox1.Location = New System.Drawing.Point(284, 9)
        Me.TextBox1.Name = "TextBox1"
        Me.TextBox1.Size = New System.Drawing.Size(100, 20)
        Me.TextBox1.TabIndex = 19
        Me.TextBox1.Visible = False
        '
        'btnApplytoall
        '
        Me.btnApplytoall.Location = New System.Drawing.Point(255, 183)
        Me.btnApplytoall.Name = "btnApplytoall"
        Me.btnApplytoall.Size = New System.Drawing.Size(75, 23)
        Me.btnApplytoall.TabIndex = 20
        Me.btnApplytoall.Text = "Apply To All"
        Me.ToolTip1.SetToolTip(Me.btnApplytoall, "Apply current information to all items.")
        Me.btnApplytoall.UseVisualStyleBackColor = True
        Me.btnApplytoall.Visible = False
        '
        'ProgressBar1
        '
        Me.ProgressBar1.Location = New System.Drawing.Point(255, 183)
        Me.ProgressBar1.Name = "ProgressBar1"
        Me.ProgressBar1.Size = New System.Drawing.Size(357, 23)
        Me.ProgressBar1.TabIndex = 21
        Me.ProgressBar1.Visible = False
        '
        'CheckNeeded
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(620, 218)
        Me.Controls.Add(Me.ProgressBar1)
        Me.Controls.Add(Me.btnApplytoall)
        Me.Controls.Add(Me.TextBox1)
        Me.Controls.Add(Me.DateTimePicker1)
        Me.Controls.Add(Me.btnCancel)
        Me.Controls.Add(Me.btnIgnore)
        Me.Controls.Add(Me.btnOK)
        Me.Controls.Add(Me.Label2)
        Me.Controls.Add(Me.Label1)
        Me.Controls.Add(Me.lstCheckNeeded)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.Fixed3D
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.Name = "CheckNeeded"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterParent
        Me.Text = "CheckNeeded"
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents DateTimePicker1 As System.Windows.Forms.DateTimePicker
    Friend WithEvents btnCancel As System.Windows.Forms.Button
    Friend WithEvents btnIgnore As System.Windows.Forms.Button
    Friend WithEvents btnOK As System.Windows.Forms.Button
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents lstCheckNeeded As System.Windows.Forms.ListView
    Friend WithEvents CheckedBy As System.Windows.Forms.ColumnHeader
    Friend WithEvents DateChecked As System.Windows.Forms.ColumnHeader
    Friend WithEvents Initials As System.Windows.Forms.ColumnHeader
    Friend WithEvents RevCheck As System.Windows.Forms.ColumnHeader
    Friend WithEvents RevDate As System.Windows.Forms.ColumnHeader
    Friend WithEvents More As System.Windows.Forms.ColumnHeader
    Friend WithEvents TextBox1 As System.Windows.Forms.TextBox
    Protected WithEvents DrawingNumber As System.Windows.Forms.ColumnHeader
    Friend WithEvents ToolTip1 As System.Windows.Forms.ToolTip
    Friend WithEvents btnApplytoall As System.Windows.Forms.Button
    Friend WithEvents ProgressBar1 As System.Windows.Forms.ProgressBar
End Class
