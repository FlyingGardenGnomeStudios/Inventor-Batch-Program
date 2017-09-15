<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class About
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
        Me.PicDonate = New System.Windows.Forms.PictureBox()
        Me.LblDonate = New System.Windows.Forms.Label()
        Me.lblName = New System.Windows.Forms.Label()
        Me.lblPublish = New System.Windows.Forms.Label()
        Me.lblCopyright = New System.Windows.Forms.Label()
        CType(Me.PicDonate, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'PicDonate
        '
        Me.PicDonate.Image = Global.My.Resources.Resources.btn_donate_LG
        Me.PicDonate.Location = New System.Drawing.Point(213, 389)
        Me.PicDonate.Name = "PicDonate"
        Me.PicDonate.Size = New System.Drawing.Size(96, 31)
        Me.PicDonate.TabIndex = 0
        Me.PicDonate.TabStop = False
        Me.PicDonate.Visible = False
        '
        'LblDonate
        '
        Me.LblDonate.AutoSize = True
        Me.LblDonate.Location = New System.Drawing.Point(12, 389)
        Me.LblDonate.Name = "LblDonate"
        Me.LblDonate.Size = New System.Drawing.Size(166, 26)
        Me.LblDonate.TabIndex = 1
        Me.LblDonate.Text = "Has this program saved you time?" & Global.Microsoft.VisualBasic.ChrW(13) & Global.Microsoft.VisualBasic.ChrW(10) & "Consider saying thanks."
        Me.LblDonate.Visible = False
        '
        'lblName
        '
        Me.lblName.AutoSize = True
        Me.lblName.Location = New System.Drawing.Point(16, 328)
        Me.lblName.Name = "lblName"
        Me.lblName.Size = New System.Drawing.Size(83, 13)
        Me.lblName.TabIndex = 2
        Me.lblName.Text = "Batch Program: "
        '
        'lblPublish
        '
        Me.lblPublish.AutoSize = True
        Me.lblPublish.Location = New System.Drawing.Point(16, 354)
        Me.lblPublish.Name = "lblPublish"
        Me.lblPublish.Size = New System.Drawing.Size(147, 13)
        Me.lblPublish.TabIndex = 3
        Me.lblPublish.Text = "Flying Garden Gnome Studios"
        '
        'lblCopyright
        '
        Me.lblCopyright.AutoSize = True
        Me.lblCopyright.Location = New System.Drawing.Point(16, 341)
        Me.lblCopyright.Name = "lblCopyright"
        Me.lblCopyright.Size = New System.Drawing.Size(51, 13)
        Me.lblCopyright.TabIndex = 4
        Me.lblCopyright.Text = "Copyright"
        '
        'About
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.BackgroundImage = Global.My.Resources.Resources.About
        Me.BackgroundImageLayout = System.Windows.Forms.ImageLayout.None
        Me.ClientSize = New System.Drawing.Size(310, 422)
        Me.Controls.Add(Me.lblCopyright)
        Me.Controls.Add(Me.lblPublish)
        Me.Controls.Add(Me.lblName)
        Me.Controls.Add(Me.LblDonate)
        Me.Controls.Add(Me.PicDonate)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedToolWindow
        Me.Name = "About"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterParent
        Me.Text = "About"
        CType(Me.PicDonate, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub

    Friend WithEvents PicDonate As Windows.Forms.PictureBox
    Friend WithEvents LblDonate As Windows.Forms.Label
    Friend WithEvents lblName As Windows.Forms.Label
    Friend WithEvents lblPublish As Windows.Forms.Label
    Friend WithEvents lblCopyright As Windows.Forms.Label
End Class
