<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class Export
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
        Me.btnCancel = New System.Windows.Forms.Button()
        Me.btnOK = New System.Windows.Forms.Button()
        Me.GroupBox1 = New System.Windows.Forms.GroupBox()
        Me.rdoFirstPage = New System.Windows.Forms.RadioButton()
        Me.rdoCurrentPage = New System.Windows.Forms.RadioButton()
        Me.rdoAllPages = New System.Windows.Forms.RadioButton()
        Me.chkPDF = New System.Windows.Forms.CheckBox()
        Me.chkDWG = New System.Windows.Forms.CheckBox()
        Me.chkDXF = New System.Windows.Forms.CheckBox()
        Me.GroupBox2 = New System.Windows.Forms.GroupBox()
        Me.GroupBox1.SuspendLayout()
        Me.GroupBox2.SuspendLayout()
        Me.SuspendLayout()
        '
        'btnCancel
        '
        Me.btnCancel.Location = New System.Drawing.Point(53, 91)
        Me.btnCancel.Name = "btnCancel"
        Me.btnCancel.Size = New System.Drawing.Size(75, 23)
        Me.btnCancel.TabIndex = 55
        Me.btnCancel.Text = "Cancel"
        Me.btnCancel.UseVisualStyleBackColor = True
        '
        'btnOK
        '
        Me.btnOK.Location = New System.Drawing.Point(134, 91)
        Me.btnOK.Name = "btnOK"
        Me.btnOK.Size = New System.Drawing.Size(75, 23)
        Me.btnOK.TabIndex = 54
        Me.btnOK.Text = "OK"
        Me.btnOK.UseVisualStyleBackColor = True
        '
        'GroupBox1
        '
        Me.GroupBox1.Controls.Add(Me.rdoFirstPage)
        Me.GroupBox1.Controls.Add(Me.rdoCurrentPage)
        Me.GroupBox1.Controls.Add(Me.rdoAllPages)
        Me.GroupBox1.Location = New System.Drawing.Point(12, 12)
        Me.GroupBox1.Name = "GroupBox1"
        Me.GroupBox1.Size = New System.Drawing.Size(104, 73)
        Me.GroupBox1.TabIndex = 52
        Me.GroupBox1.TabStop = False
        Me.GroupBox1.Text = "Print Range"
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
        'chkPDF
        '
        Me.chkPDF.AutoSize = True
        Me.chkPDF.Location = New System.Drawing.Point(6, 20)
        Me.chkPDF.Name = "chkPDF"
        Me.chkPDF.Size = New System.Drawing.Size(47, 17)
        Me.chkPDF.TabIndex = 56
        Me.chkPDF.Text = "PDF"
        Me.chkPDF.UseVisualStyleBackColor = True
        '
        'chkDWG
        '
        Me.chkDWG.AutoSize = True
        Me.chkDWG.Location = New System.Drawing.Point(6, 35)
        Me.chkDWG.Name = "chkDWG"
        Me.chkDWG.Size = New System.Drawing.Size(53, 17)
        Me.chkDWG.TabIndex = 57
        Me.chkDWG.Text = "DWG"
        Me.chkDWG.UseVisualStyleBackColor = True
        '
        'chkDXF
        '
        Me.chkDXF.AutoSize = True
        Me.chkDXF.Location = New System.Drawing.Point(6, 50)
        Me.chkDXF.Name = "chkDXF"
        Me.chkDXF.Size = New System.Drawing.Size(47, 17)
        Me.chkDXF.TabIndex = 58
        Me.chkDXF.Text = "DXF"
        Me.chkDXF.UseVisualStyleBackColor = True
        '
        'GroupBox2
        '
        Me.GroupBox2.Controls.Add(Me.chkPDF)
        Me.GroupBox2.Controls.Add(Me.chkDXF)
        Me.GroupBox2.Controls.Add(Me.chkDWG)
        Me.GroupBox2.Location = New System.Drawing.Point(122, 12)
        Me.GroupBox2.Name = "GroupBox2"
        Me.GroupBox2.Size = New System.Drawing.Size(87, 73)
        Me.GroupBox2.TabIndex = 59
        Me.GroupBox2.TabStop = False
        Me.GroupBox2.Text = "Export Type"
        '
        'Export
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(226, 123)
        Me.Controls.Add(Me.GroupBox2)
        Me.Controls.Add(Me.btnCancel)
        Me.Controls.Add(Me.btnOK)
        Me.Controls.Add(Me.GroupBox1)
        Me.Name = "Export"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterParent
        Me.Text = "Export"
        Me.GroupBox1.ResumeLayout(False)
        Me.GroupBox1.PerformLayout()
        Me.GroupBox2.ResumeLayout(False)
        Me.GroupBox2.PerformLayout()
        Me.ResumeLayout(False)

    End Sub
    Friend WithEvents btnCancel As Windows.Forms.Button
    Friend WithEvents btnOK As Windows.Forms.Button
    Friend WithEvents GroupBox1 As Windows.Forms.GroupBox
    Friend WithEvents rdoFirstPage As Windows.Forms.RadioButton
    Friend WithEvents rdoCurrentPage As Windows.Forms.RadioButton
    Friend WithEvents rdoAllPages As Windows.Forms.RadioButton
    Friend WithEvents chkPDF As Windows.Forms.CheckBox
    Friend WithEvents chkDWG As Windows.Forms.CheckBox
    Friend WithEvents chkDXF As Windows.Forms.CheckBox
    Friend WithEvents GroupBox2 As Windows.Forms.GroupBox
End Class
