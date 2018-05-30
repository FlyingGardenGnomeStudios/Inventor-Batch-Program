<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class ReVerifyNow
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
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(ReVerifyNow))
        Me.lblDescr = New System.Windows.Forms.Label()
        Me.btnExit = New System.Windows.Forms.Button()
        Me.btnReverify = New System.Windows.Forms.Button()
        Me.SuspendLayout()
        '
        'lblDescr
        '
        Me.lblDescr.AutoSize = True
        Me.lblDescr.Location = New System.Drawing.Point(30, 12)
        Me.lblDescr.Name = "lblDescr"
        Me.lblDescr.Size = New System.Drawing.Size(269, 13)
        Me.lblDescr.TabIndex = 11
        Me.lblDescr.Text = "You have X days to re-verify with the activation servers."
        '
        'btnExit
        '
        Me.btnExit.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.btnExit.FlatStyle = System.Windows.Forms.FlatStyle.System
        Me.btnExit.Location = New System.Drawing.Point(212, 49)
        Me.btnExit.Name = "btnExit"
        Me.btnExit.Size = New System.Drawing.Size(130, 24)
        Me.btnExit.TabIndex = 10
        Me.btnExit.Text = "Exit application"
        Me.btnExit.UseVisualStyleBackColor = True
        '
        'btnReverify
        '
        Me.btnReverify.FlatStyle = System.Windows.Forms.FlatStyle.System
        Me.btnReverify.Location = New System.Drawing.Point(32, 49)
        Me.btnReverify.Name = "btnReverify"
        Me.btnReverify.Size = New System.Drawing.Size(92, 24)
        Me.btnReverify.TabIndex = 9
        Me.btnReverify.Text = "Re-verify now"
        Me.btnReverify.UseVisualStyleBackColor = True
        '
        'ReVerifyNow
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(373, 84)
        Me.Controls.Add(Me.lblDescr)
        Me.Controls.Add(Me.btnExit)
        Me.Controls.Add(Me.btnReverify)
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.Name = "ReVerifyNow"
        Me.Text = "ReVerifyNow"
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub

    Private WithEvents lblDescr As Windows.Forms.Label
    Private WithEvents btnExit As Windows.Forms.Button
    Public WithEvents btnReverify As Windows.Forms.Button
End Class
