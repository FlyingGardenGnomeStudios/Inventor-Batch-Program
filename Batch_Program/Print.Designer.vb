<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class Print
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
        Me.rdoAllPages = New System.Windows.Forms.RadioButton()
        Me.rdoCurrentPage = New System.Windows.Forms.RadioButton()
        Me.rdoFirstPage = New System.Windows.Forms.RadioButton()
        Me.GroupBox1 = New System.Windows.Forms.GroupBox()
        Me.rdoFull = New System.Windows.Forms.RadioButton()
        Me.rdoScale = New System.Windows.Forms.RadioButton()
        Me.GroupBox2 = New System.Windows.Forms.GroupBox()
        Me.btnOK = New System.Windows.Forms.Button()
        Me.btnCancel = New System.Windows.Forms.Button()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.txtCopies = New System.Windows.Forms.NumericUpDown()
        Me.chkReverse = New System.Windows.Forms.CheckBox()
        Me.GroupBox3 = New System.Windows.Forms.GroupBox()
        Me.rdoColour = New System.Windows.Forms.RadioButton()
        Me.rdoBW = New System.Windows.Forms.RadioButton()
        Me.GroupBox1.SuspendLayout()
        Me.GroupBox2.SuspendLayout()
        CType(Me.txtCopies, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.GroupBox3.SuspendLayout()
        Me.SuspendLayout()
        '
        'rdoAllPages
        '
        Me.rdoAllPages.AutoSize = True
        Me.rdoAllPages.Checked = True
        Me.rdoAllPages.Location = New System.Drawing.Point(6, 49)
        Me.rdoAllPages.Name = "rdoAllPages"
        Me.rdoAllPages.Size = New System.Drawing.Size(69, 17)
        Me.rdoAllPages.TabIndex = 0
        Me.rdoAllPages.TabStop = True
        Me.rdoAllPages.Text = "All Pages"
        Me.rdoAllPages.UseVisualStyleBackColor = True
        '
        'rdoCurrentPage
        '
        Me.rdoCurrentPage.AutoSize = True
        Me.rdoCurrentPage.Location = New System.Drawing.Point(6, 19)
        Me.rdoCurrentPage.Name = "rdoCurrentPage"
        Me.rdoCurrentPage.Size = New System.Drawing.Size(87, 17)
        Me.rdoCurrentPage.TabIndex = 1
        Me.rdoCurrentPage.Text = "Current Page"
        Me.rdoCurrentPage.UseVisualStyleBackColor = True
        '
        'rdoFirstPage
        '
        Me.rdoFirstPage.AutoSize = True
        Me.rdoFirstPage.Location = New System.Drawing.Point(6, 34)
        Me.rdoFirstPage.Name = "rdoFirstPage"
        Me.rdoFirstPage.Size = New System.Drawing.Size(72, 17)
        Me.rdoFirstPage.TabIndex = 2
        Me.rdoFirstPage.Text = "First Page"
        Me.rdoFirstPage.UseVisualStyleBackColor = True
        '
        'GroupBox1
        '
        Me.GroupBox1.Controls.Add(Me.rdoFirstPage)
        Me.GroupBox1.Controls.Add(Me.rdoCurrentPage)
        Me.GroupBox1.Controls.Add(Me.rdoAllPages)
        Me.GroupBox1.Location = New System.Drawing.Point(12, 12)
        Me.GroupBox1.Name = "GroupBox1"
        Me.GroupBox1.Size = New System.Drawing.Size(104, 73)
        Me.GroupBox1.TabIndex = 3
        Me.GroupBox1.TabStop = False
        Me.GroupBox1.Text = "Print Range"
        '
        'rdoFull
        '
        Me.rdoFull.AutoSize = True
        Me.rdoFull.Checked = True
        Me.rdoFull.Location = New System.Drawing.Point(6, 17)
        Me.rdoFull.Name = "rdoFull"
        Me.rdoFull.Size = New System.Drawing.Size(64, 17)
        Me.rdoFull.TabIndex = 4
        Me.rdoFull.TabStop = True
        Me.rdoFull.Text = "Full Size"
        Me.rdoFull.UseVisualStyleBackColor = True
        '
        'rdoScale
        '
        Me.rdoScale.AutoSize = True
        Me.rdoScale.Location = New System.Drawing.Point(6, 32)
        Me.rdoScale.Name = "rdoScale"
        Me.rdoScale.Size = New System.Drawing.Size(83, 17)
        Me.rdoScale.TabIndex = 5
        Me.rdoScale.Text = "Scale Down"
        Me.rdoScale.UseVisualStyleBackColor = True
        '
        'GroupBox2
        '
        Me.GroupBox2.Controls.Add(Me.rdoScale)
        Me.GroupBox2.Controls.Add(Me.rdoFull)
        Me.GroupBox2.Location = New System.Drawing.Point(12, 91)
        Me.GroupBox2.Name = "GroupBox2"
        Me.GroupBox2.Size = New System.Drawing.Size(104, 51)
        Me.GroupBox2.TabIndex = 6
        Me.GroupBox2.TabStop = False
        Me.GroupBox2.Text = "Pages to Print"
        '
        'btnOK
        '
        Me.btnOK.Location = New System.Drawing.Point(203, 117)
        Me.btnOK.Name = "btnOK"
        Me.btnOK.Size = New System.Drawing.Size(75, 23)
        Me.btnOK.TabIndex = 9
        Me.btnOK.Text = "OK"
        Me.btnOK.UseVisualStyleBackColor = True
        '
        'btnCancel
        '
        Me.btnCancel.Location = New System.Drawing.Point(122, 117)
        Me.btnCancel.Name = "btnCancel"
        Me.btnCancel.Size = New System.Drawing.Size(75, 23)
        Me.btnCancel.TabIndex = 10
        Me.btnCancel.Text = "Cancel"
        Me.btnCancel.UseVisualStyleBackColor = True
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(125, 72)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(39, 13)
        Me.Label1.TabIndex = 11
        Me.Label1.Text = "Copies"
        '
        'txtCopies
        '
        Me.txtCopies.Location = New System.Drawing.Point(170, 70)
        Me.txtCopies.Name = "txtCopies"
        Me.txtCopies.Size = New System.Drawing.Size(35, 20)
        Me.txtCopies.TabIndex = 49
        Me.txtCopies.Value = New Decimal(New Integer() {1, 0, 0, 0})
        '
        'chkReverse
        '
        Me.chkReverse.AutoSize = True
        Me.chkReverse.Location = New System.Drawing.Point(128, 91)
        Me.chkReverse.Name = "chkReverse"
        Me.chkReverse.Size = New System.Drawing.Size(95, 17)
        Me.chkReverse.TabIndex = 50
        Me.chkReverse.Text = "Reverse Order"
        Me.chkReverse.UseVisualStyleBackColor = True
        '
        'GroupBox3
        '
        Me.GroupBox3.Controls.Add(Me.rdoColour)
        Me.GroupBox3.Controls.Add(Me.rdoBW)
        Me.GroupBox3.Location = New System.Drawing.Point(122, 12)
        Me.GroupBox3.Name = "GroupBox3"
        Me.GroupBox3.Size = New System.Drawing.Size(119, 51)
        Me.GroupBox3.TabIndex = 51
        Me.GroupBox3.TabStop = False
        Me.GroupBox3.Text = "Colour Settings"
        '
        'rdoColour
        '
        Me.rdoColour.AutoSize = True
        Me.rdoColour.Checked = True
        Me.rdoColour.Location = New System.Drawing.Point(6, 32)
        Me.rdoColour.Name = "rdoColour"
        Me.rdoColour.Size = New System.Drawing.Size(55, 17)
        Me.rdoColour.TabIndex = 5
        Me.rdoColour.TabStop = True
        Me.rdoColour.Text = "Colour"
        Me.rdoColour.UseVisualStyleBackColor = True
        '
        'rdoBW
        '
        Me.rdoBW.AutoSize = True
        Me.rdoBW.Location = New System.Drawing.Point(6, 17)
        Me.rdoBW.Name = "rdoBW"
        Me.rdoBW.Size = New System.Drawing.Size(104, 17)
        Me.rdoBW.TabIndex = 4
        Me.rdoBW.Text = "Black and White"
        Me.rdoBW.UseVisualStyleBackColor = True
        '
        'Print
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(288, 152)
        Me.Controls.Add(Me.GroupBox3)
        Me.Controls.Add(Me.chkReverse)
        Me.Controls.Add(Me.txtCopies)
        Me.Controls.Add(Me.Label1)
        Me.Controls.Add(Me.btnCancel)
        Me.Controls.Add(Me.btnOK)
        Me.Controls.Add(Me.GroupBox2)
        Me.Controls.Add(Me.GroupBox1)
        Me.Name = "Print"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterParent
        Me.Text = "Print"
        Me.GroupBox1.ResumeLayout(False)
        Me.GroupBox1.PerformLayout()
        Me.GroupBox2.ResumeLayout(False)
        Me.GroupBox2.PerformLayout()
        CType(Me.txtCopies, System.ComponentModel.ISupportInitialize).EndInit()
        Me.GroupBox3.ResumeLayout(False)
        Me.GroupBox3.PerformLayout()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub

    Friend WithEvents rdoAllPages As Windows.Forms.RadioButton
    Friend WithEvents rdoCurrentPage As Windows.Forms.RadioButton
    Friend WithEvents rdoFirstPage As Windows.Forms.RadioButton
    Friend WithEvents GroupBox1 As Windows.Forms.GroupBox
    Friend WithEvents rdoFull As Windows.Forms.RadioButton
    Friend WithEvents rdoScale As Windows.Forms.RadioButton
    Friend WithEvents GroupBox2 As Windows.Forms.GroupBox
    Friend WithEvents btnOK As Windows.Forms.Button
    Friend WithEvents btnCancel As Windows.Forms.Button
    Friend WithEvents Label1 As Windows.Forms.Label
    Friend WithEvents txtCopies As Windows.Forms.NumericUpDown
    Friend WithEvents chkReverse As Windows.Forms.CheckBox
    Friend WithEvents GroupBox3 As Windows.Forms.GroupBox
    Friend WithEvents rdoColour As Windows.Forms.RadioButton
    Friend WithEvents rdoBW As Windows.Forms.RadioButton
End Class
