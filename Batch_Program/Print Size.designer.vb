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
        Me.btnFullSize = New System.Windows.Forms.Button()
        Me.btn11x17 = New System.Windows.Forms.Button()
        Me.chkScale = New System.Windows.Forms.CheckBox()
        Me.chkQuesiton = New System.Windows.Forms.CheckBox()
        Me.btnCancel = New System.Windows.Forms.Button()
        Me.SuspendLayout()
        '
        'lblDWGSize
        '
        Me.lblDWGSize.AutoSize = True
        Me.lblDWGSize.Location = New System.Drawing.Point(12, 9)
        Me.lblDWGSize.Name = "lblDWGSize"
        Me.lblDWGSize.Size = New System.Drawing.Size(33, 13)
        Me.lblDWGSize.TabIndex = 0
        Me.lblDWGSize.Text = "Label"
        Me.lblDWGSize.TextAlign = System.Drawing.ContentAlignment.TopCenter
        '
        'rdoDontAsk
        '
        Me.rdoDontAsk.AutoSize = True
        Me.rdoDontAsk.Location = New System.Drawing.Point(16, 56)
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
        Me.rdoAsk.Location = New System.Drawing.Point(16, 79)
        Me.rdoAsk.Name = "rdoAsk"
        Me.rdoAsk.Size = New System.Drawing.Size(88, 17)
        Me.rdoAsk.TabIndex = 2
        Me.rdoAsk.TabStop = True
        Me.rdoAsk.Text = "Ask next time"
        Me.rdoAsk.UseVisualStyleBackColor = True
        '
        'btnFullSize
        '
        Me.btnFullSize.Location = New System.Drawing.Point(211, 50)
        Me.btnFullSize.Name = "btnFullSize"
        Me.btnFullSize.Size = New System.Drawing.Size(86, 23)
        Me.btnFullSize.TabIndex = 3
        Me.btnFullSize.Text = "Print Full Size"
        Me.btnFullSize.UseVisualStyleBackColor = True
        '
        'btn11x17
        '
        Me.btn11x17.Location = New System.Drawing.Point(211, 73)
        Me.btn11x17.Name = "btn11x17"
        Me.btn11x17.Size = New System.Drawing.Size(86, 23)
        Me.btn11x17.TabIndex = 4
        Me.btn11x17.Text = "Scale Down"
        Me.btn11x17.UseVisualStyleBackColor = True
        '
        'chkScale
        '
        Me.chkScale.AutoSize = True
        Me.chkScale.Location = New System.Drawing.Point(16, 103)
        Me.chkScale.Name = "chkScale"
        Me.chkScale.Size = New System.Drawing.Size(53, 17)
        Me.chkScale.TabIndex = 5
        Me.chkScale.Text = "Scale"
        Me.chkScale.UseVisualStyleBackColor = True
        Me.chkScale.Visible = False
        '
        'chkQuesiton
        '
        Me.chkQuesiton.AutoSize = True
        Me.chkQuesiton.Location = New System.Drawing.Point(75, 103)
        Me.chkQuesiton.Name = "chkQuesiton"
        Me.chkQuesiton.Size = New System.Drawing.Size(68, 17)
        Me.chkQuesiton.TabIndex = 6
        Me.chkQuesiton.Text = "Question"
        Me.chkQuesiton.UseVisualStyleBackColor = True
        Me.chkQuesiton.Visible = False
        '
        'btnCancel
        '
        Me.btnCancel.Location = New System.Drawing.Point(211, 96)
        Me.btnCancel.Name = "btnCancel"
        Me.btnCancel.Size = New System.Drawing.Size(86, 23)
        Me.btnCancel.TabIndex = 7
        Me.btnCancel.Text = "Cancel"
        Me.btnCancel.UseVisualStyleBackColor = True
        '
        'Print_Size
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(309, 130)
        Me.Controls.Add(Me.btnCancel)
        Me.Controls.Add(Me.chkQuesiton)
        Me.Controls.Add(Me.chkScale)
        Me.Controls.Add(Me.btn11x17)
        Me.Controls.Add(Me.btnFullSize)
        Me.Controls.Add(Me.rdoAsk)
        Me.Controls.Add(Me.rdoDontAsk)
        Me.Controls.Add(Me.lblDWGSize)
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.Name = "Print_Size"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterParent
        Me.Text = "Large Paper Size"
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub

    Friend WithEvents lblDWGSize As Windows.Forms.Label
    Friend WithEvents rdoDontAsk As Windows.Forms.RadioButton
    Friend WithEvents rdoAsk As Windows.Forms.RadioButton
    Friend WithEvents btnFullSize As Windows.Forms.Button
    Friend WithEvents btn11x17 As Windows.Forms.Button
    Friend WithEvents chkScale As Windows.Forms.CheckBox
    Friend WithEvents chkQuesiton As Windows.Forms.CheckBox
    Friend WithEvents btnCancel As Windows.Forms.Button
End Class
