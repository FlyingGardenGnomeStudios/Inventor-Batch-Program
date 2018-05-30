<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class Warning
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
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(Warning))
        Me.chkDontShow = New System.Windows.Forms.CheckBox()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.btnOK = New System.Windows.Forms.Button()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.PicDonate = New System.Windows.Forms.PictureBox()
        CType(Me.PicDonate, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'chkDontShow
        '
        Me.chkDontShow.AutoSize = True
        Me.chkDontShow.Location = New System.Drawing.Point(18, 43)
        Me.chkDontShow.Name = "chkDontShow"
        Me.chkDontShow.Size = New System.Drawing.Size(125, 17)
        Me.chkDontShow.TabIndex = 0
        Me.chkDontShow.Text = "Don't show me again"
        Me.chkDontShow.UseVisualStyleBackColor = True
        '
        'Label1
        '
        Me.Label1.Anchor = System.Windows.Forms.AnchorStyles.None
        Me.Label1.Location = New System.Drawing.Point(2, 6)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(433, 30)
        Me.Label1.TabIndex = 1
        Me.Label1.Text = "The renaming function does not currently support Top-Down assembly configurations" &
    "." & Global.Microsoft.VisualBasic.ChrW(13) & Global.Microsoft.VisualBasic.ChrW(10) & "Referenced function parameters will not be re-linked to the new part name."
        Me.Label1.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'btnOK
        '
        Me.btnOK.Location = New System.Drawing.Point(349, 39)
        Me.btnOK.Name = "btnOK"
        Me.btnOK.Size = New System.Drawing.Size(75, 23)
        Me.btnOK.TabIndex = 2
        Me.btnOK.Text = "OK"
        Me.btnOK.UseVisualStyleBackColor = True
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Location = New System.Drawing.Point(161, 48)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(47, 13)
        Me.Label2.TabIndex = 3
        Me.Label2.Text = "Rename"
        Me.Label2.Visible = False
        '
        'PicDonate
        '
        Me.PicDonate.Image = Global.My.Resources.Resources.btn_donate_LG
        Me.PicDonate.Location = New System.Drawing.Point(131, 39)
        Me.PicDonate.Name = "PicDonate"
        Me.PicDonate.Size = New System.Drawing.Size(96, 31)
        Me.PicDonate.TabIndex = 4
        Me.PicDonate.TabStop = False
        Me.PicDonate.Visible = False
        '
        'Warning
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(436, 95)
        Me.ControlBox = False
        Me.Controls.Add(Me.PicDonate)
        Me.Controls.Add(Me.Label2)
        Me.Controls.Add(Me.btnOK)
        Me.Controls.Add(Me.Label1)
        Me.Controls.Add(Me.chkDontShow)
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.MaximumSize = New System.Drawing.Size(452, 150)
        Me.MinimumSize = New System.Drawing.Size(452, 111)
        Me.Name = "Warning"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterParent
        Me.Text = "Warning"
        Me.TopMost = True
        CType(Me.PicDonate, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub

    Friend WithEvents chkDontShow As Windows.Forms.CheckBox
    Friend WithEvents Label1 As Windows.Forms.Label
    Friend WithEvents btnOK As Windows.Forms.Button
    Friend WithEvents Label2 As Windows.Forms.Label
    Friend WithEvents PicDonate As Windows.Forms.PictureBox
End Class
