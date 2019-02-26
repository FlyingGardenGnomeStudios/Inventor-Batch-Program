<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class Print_Size
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
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(Print_Size))
        Me.lblDWGSize = New System.Windows.Forms.Label()
        Me.rdoDontAsk = New System.Windows.Forms.RadioButton()
        Me.rdoAsk = New System.Windows.Forms.RadioButton()
        Me.btn11x17 = New System.Windows.Forms.Button()
        Me.chkScale = New System.Windows.Forms.CheckBox()
        Me.chkQuestion = New System.Windows.Forms.CheckBox()
        Me.btnCancel = New System.Windows.Forms.Button()
        Me.cmbScaleSize = New System.Windows.Forms.ComboBox()
        Me.btnOK = New System.Windows.Forms.Button()
        Me.GroupBox1 = New System.Windows.Forms.GroupBox()
        Me.GroupBox1.SuspendLayout()
        Me.SuspendLayout()
        '
        'lblDWGSize
        '
        Me.lblDWGSize.Location = New System.Drawing.Point(12, 9)
        Me.lblDWGSize.Name = "lblDWGSize"
        Me.lblDWGSize.Size = New System.Drawing.Size(194, 53)
        Me.lblDWGSize.TabIndex = 0
        Me.lblDWGSize.Text = "The current drawing is a D size sheet. " & Global.Microsoft.VisualBasic.ChrW(13) & Global.Microsoft.VisualBasic.ChrW(10) & "Select the size you would like to print"
        Me.lblDWGSize.TextAlign = System.Drawing.ContentAlignment.TopCenter
        '
        'rdoDontAsk
        '
        Me.rdoDontAsk.AutoSize = True
        Me.rdoDontAsk.Location = New System.Drawing.Point(6, 19)
        Me.rdoDontAsk.Name = "rdoDontAsk"
        Me.rdoDontAsk.Size = New System.Drawing.Size(99, 17)
        Me.rdoDontAsk.TabIndex = 1
        Me.rdoDontAsk.Text = "Don't ask again"
        Me.rdoDontAsk.UseVisualStyleBackColor = True
        '
        'rdoAsk
        '
        Me.rdoAsk.AutoSize = True
        Me.rdoAsk.Checked = True
        Me.rdoAsk.Location = New System.Drawing.Point(6, 36)
        Me.rdoAsk.Name = "rdoAsk"
        Me.rdoAsk.Size = New System.Drawing.Size(88, 17)
        Me.rdoAsk.TabIndex = 2
        Me.rdoAsk.TabStop = True
        Me.rdoAsk.Text = "Ask next time"
        Me.rdoAsk.UseVisualStyleBackColor = True
        '
        'btn11x17
        '
        Me.btn11x17.Location = New System.Drawing.Point(212, 9)
        Me.btn11x17.Name = "btn11x17"
        Me.btn11x17.Size = New System.Drawing.Size(98, 23)
        Me.btn11x17.TabIndex = 4
        Me.btn11x17.Text = "Open Print Dialog"
        Me.btn11x17.UseVisualStyleBackColor = True
        '
        'chkScale
        '
        Me.chkScale.AutoSize = True
        Me.chkScale.Location = New System.Drawing.Point(111, 36)
        Me.chkScale.Name = "chkScale"
        Me.chkScale.Size = New System.Drawing.Size(53, 17)
        Me.chkScale.TabIndex = 5
        Me.chkScale.Text = "Scale"
        Me.chkScale.UseVisualStyleBackColor = True
        Me.chkScale.Visible = False
        '
        'chkQuestion
        '
        Me.chkQuestion.AutoSize = True
        Me.chkQuestion.Location = New System.Drawing.Point(111, 14)
        Me.chkQuestion.Name = "chkQuestion"
        Me.chkQuestion.Size = New System.Drawing.Size(68, 17)
        Me.chkQuestion.TabIndex = 6
        Me.chkQuestion.Text = "Question"
        Me.chkQuestion.UseVisualStyleBackColor = True
        Me.chkQuestion.Visible = False
        '
        'btnCancel
        '
        Me.btnCancel.Location = New System.Drawing.Point(212, 65)
        Me.btnCancel.Name = "btnCancel"
        Me.btnCancel.Size = New System.Drawing.Size(98, 23)
        Me.btnCancel.TabIndex = 7
        Me.btnCancel.Text = "Cancel"
        Me.btnCancel.UseVisualStyleBackColor = True
        '
        'cmbScaleSize
        '
        Me.cmbScaleSize.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cmbScaleSize.FormattingEnabled = True
        Me.cmbScaleSize.Location = New System.Drawing.Point(212, 38)
        Me.cmbScaleSize.Name = "cmbScaleSize"
        Me.cmbScaleSize.Size = New System.Drawing.Size(98, 21)
        Me.cmbScaleSize.TabIndex = 8
        '
        'btnOK
        '
        Me.btnOK.Location = New System.Drawing.Point(212, 94)
        Me.btnOK.Name = "btnOK"
        Me.btnOK.Size = New System.Drawing.Size(98, 23)
        Me.btnOK.TabIndex = 9
        Me.btnOK.Text = "OK"
        Me.btnOK.UseVisualStyleBackColor = True
        '
        'GroupBox1
        '
        Me.GroupBox1.Controls.Add(Me.rdoDontAsk)
        Me.GroupBox1.Controls.Add(Me.rdoAsk)
        Me.GroupBox1.Controls.Add(Me.chkScale)
        Me.GroupBox1.Controls.Add(Me.chkQuestion)
        Me.GroupBox1.Location = New System.Drawing.Point(12, 57)
        Me.GroupBox1.Name = "GroupBox1"
        Me.GroupBox1.Size = New System.Drawing.Size(194, 60)
        Me.GroupBox1.TabIndex = 10
        Me.GroupBox1.TabStop = False
        Me.GroupBox1.Text = "Repeat Options"
        '
        'Print_Size
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(323, 126)
        Me.Controls.Add(Me.GroupBox1)
        Me.Controls.Add(Me.btnOK)
        Me.Controls.Add(Me.cmbScaleSize)
        Me.Controls.Add(Me.btnCancel)
        Me.Controls.Add(Me.btn11x17)
        Me.Controls.Add(Me.lblDWGSize)
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.Name = "Print_Size"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterParent
        Me.Text = "Large Paper Size"
        Me.GroupBox1.ResumeLayout(False)
        Me.GroupBox1.PerformLayout()
        Me.ResumeLayout(False)

    End Sub

    Friend WithEvents lblDWGSize As Windows.Forms.Label
    Friend WithEvents rdoDontAsk As Windows.Forms.RadioButton
    Friend WithEvents rdoAsk As Windows.Forms.RadioButton
    Friend WithEvents btn11x17 As Windows.Forms.Button
    Friend WithEvents chkScale As Windows.Forms.CheckBox
    Friend WithEvents chkQuestion As Windows.Forms.CheckBox
    Friend WithEvents btnCancel As Windows.Forms.Button
    Friend WithEvents cmbScaleSize As Windows.Forms.ComboBox
    Friend WithEvents btnOK As Windows.Forms.Button
    Friend WithEvents GroupBox1 As Windows.Forms.GroupBox
End Class
