<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class Rev_Switch
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
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(Rev_Switch))
        Me.btnNumeric = New System.Windows.Forms.Button()
        Me.btnAlpha = New System.Windows.Forms.Button()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.chkResetRev = New System.Windows.Forms.CheckBox()
        Me.SuspendLayout()
        '
        'btnNumeric
        '
        Me.btnNumeric.Location = New System.Drawing.Point(16, 29)
        Me.btnNumeric.Name = "btnNumeric"
        Me.btnNumeric.Size = New System.Drawing.Size(75, 23)
        Me.btnNumeric.TabIndex = 0
        Me.btnNumeric.Text = "Numeric"
        Me.btnNumeric.UseVisualStyleBackColor = True
        '
        'btnAlpha
        '
        Me.btnAlpha.Location = New System.Drawing.Point(98, 29)
        Me.btnAlpha.Name = "btnAlpha"
        Me.btnAlpha.Size = New System.Drawing.Size(75, 23)
        Me.btnAlpha.TabIndex = 1
        Me.btnAlpha.Text = "Alpha"
        Me.btnAlpha.UseVisualStyleBackColor = True
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(13, 13)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(165, 13)
        Me.Label1.TabIndex = 2
        Me.Label1.Text = "Select what format you would like"
        '
        'chkResetRev
        '
        Me.chkResetRev.AutoSize = True
        Me.chkResetRev.Location = New System.Drawing.Point(16, 59)
        Me.chkResetRev.Name = "chkResetRev"
        Me.chkResetRev.Size = New System.Drawing.Size(123, 17)
        Me.chkResetRev.TabIndex = 3
        Me.chkResetRev.Text = "Keep revision history"
        Me.chkResetRev.UseVisualStyleBackColor = True
        '
        'Rev_Switch
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(195, 82)
        Me.Controls.Add(Me.chkResetRev)
        Me.Controls.Add(Me.Label1)
        Me.Controls.Add(Me.btnAlpha)
        Me.Controls.Add(Me.btnNumeric)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.Fixed3D
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.Name = "Rev_Switch"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterParent
        Me.Text = "Rev_Switch"
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents btnNumeric As System.Windows.Forms.Button
    Friend WithEvents btnAlpha As System.Windows.Forms.Button
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents chkResetRev As Windows.Forms.CheckBox
End Class
