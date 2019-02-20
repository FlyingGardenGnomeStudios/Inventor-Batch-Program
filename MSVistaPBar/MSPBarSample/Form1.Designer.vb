<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class Form1
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
        Me.MsVistaProgressBar3 = New MSVistaPBar.MSVistaProgressBar()
        Me.MsVistaProgressBar2 = New MSVistaPBar.MSVistaProgressBar()
        Me.MsVistaProgressBar1 = New MSVistaPBar.MSVistaProgressBar()
        Me.PropertyGrid1 = New System.Windows.Forms.PropertyGrid()
        Me.SuspendLayout()
        '
        'MsVistaProgressBar3
        '
        Me.MsVistaProgressBar3.BackColor = System.Drawing.Color.Transparent
        Me.MsVistaProgressBar3.BarColorMethod = MSVistaPBar.MSVistaProgressBar.BarColorsWhen.None
        Me.MsVistaProgressBar3.BlockSize = 10
        Me.MsVistaProgressBar3.BlockSpacing = 5
        Me.MsVistaProgressBar3.DisplayText = "%P%"
        Me.MsVistaProgressBar3.DisplayTextColor = System.Drawing.SystemColors.ControlText
        Me.MsVistaProgressBar3.DisplayTextFont = New System.Drawing.Font("Arial", 8.0!)
        Me.MsVistaProgressBar3.GradiantStyle = MSVistaPBar.MSVistaProgressBar.BackGradiant.None
        Me.MsVistaProgressBar3.Location = New System.Drawing.Point(13, 75)
        Me.MsVistaProgressBar3.Name = "MsVistaProgressBar3"
        Me.MsVistaProgressBar3.ProgressBarStyle = MSVistaPBar.MSVistaProgressBar.BarStyle.Blocks
        Me.MsVistaProgressBar3.ShowText = True
        Me.MsVistaProgressBar3.Size = New System.Drawing.Size(264, 32)
        Me.MsVistaProgressBar3.StartColor = System.Drawing.Color.FromArgb(CType(CType(0, Byte), Integer), CType(CType(211, Byte), Integer), CType(CType(40, Byte), Integer))
        Me.MsVistaProgressBar3.TabIndex = 4
        Me.MsVistaProgressBar3.Value = 50
        '
        'MsVistaProgressBar2
        '
        Me.MsVistaProgressBar2.BackColor = System.Drawing.Color.Transparent
        Me.MsVistaProgressBar2.BackgroundColor = System.Drawing.Color.White
        Me.MsVistaProgressBar2.BackgroundColor2 = System.Drawing.Color.Silver
        Me.MsVistaProgressBar2.BarColorMethod = MSVistaPBar.MSVistaProgressBar.BarColorsWhen.None
        Me.MsVistaProgressBar2.BlockSize = 10
        Me.MsVistaProgressBar2.BlockSpacing = 10
        Me.MsVistaProgressBar2.DisplayText = "%P%"
        Me.MsVistaProgressBar2.DisplayTextColor = System.Drawing.SystemColors.ControlText
        Me.MsVistaProgressBar2.DisplayTextFont = New System.Drawing.Font("Arial", 8.0!)
        Me.MsVistaProgressBar2.GradiantStyle = MSVistaPBar.MSVistaProgressBar.BackGradiant.Horizontal
        Me.MsVistaProgressBar2.Location = New System.Drawing.Point(13, 51)
        Me.MsVistaProgressBar2.MarqueeSpeed = 40
        Me.MsVistaProgressBar2.Name = "MsVistaProgressBar2"
        Me.MsVistaProgressBar2.OuterStrokeColor = System.Drawing.Color.FromArgb(CType(CType(95, Byte), Integer), CType(CType(169, Byte), Integer), CType(CType(207, Byte), Integer))
        Me.MsVistaProgressBar2.ShowText = False
        Me.MsVistaProgressBar2.Size = New System.Drawing.Size(264, 17)
        Me.MsVistaProgressBar2.StartColor = System.Drawing.Color.FromArgb(CType(CType(152, Byte), Integer), CType(CType(205, Byte), Integer), CType(CType(216, Byte), Integer))
        Me.MsVistaProgressBar2.TabIndex = 3
        '
        'MsVistaProgressBar1
        '
        Me.MsVistaProgressBar1.BackColor = System.Drawing.Color.Transparent
        Me.MsVistaProgressBar1.BackgroundColor2 = System.Drawing.Color.White
        Me.MsVistaProgressBar1.BlockSize = 10
        Me.MsVistaProgressBar1.BlockSpacing = 10
        Me.MsVistaProgressBar1.DisplayText = "%P%"
        Me.MsVistaProgressBar1.DisplayTextColor = System.Drawing.SystemColors.ControlText
        Me.MsVistaProgressBar1.DisplayTextFont = New System.Drawing.Font("Arial", 8.0!)
        Me.MsVistaProgressBar1.GradiantStyle = MSVistaPBar.MSVistaProgressBar.BackGradiant.None
        Me.MsVistaProgressBar1.Location = New System.Drawing.Point(12, 12)
        Me.MsVistaProgressBar1.MarqueeSpeed = 10
        Me.MsVistaProgressBar1.Name = "MsVistaProgressBar1"
        Me.MsVistaProgressBar1.ShowText = True
        Me.MsVistaProgressBar1.Size = New System.Drawing.Size(264, 32)
        Me.MsVistaProgressBar1.TabIndex = 2
        Me.MsVistaProgressBar1.Value = 50
        '
        'PropertyGrid1
        '
        Me.PropertyGrid1.Dock = System.Windows.Forms.DockStyle.Right
        Me.PropertyGrid1.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.PropertyGrid1.Location = New System.Drawing.Point(283, 0)
        Me.PropertyGrid1.Name = "PropertyGrid1"
        Me.PropertyGrid1.SelectedObject = Me.MsVistaProgressBar1
        Me.PropertyGrid1.Size = New System.Drawing.Size(236, 334)
        Me.PropertyGrid1.TabIndex = 1
        '
        'Form1
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(519, 334)
        Me.Controls.Add(Me.MsVistaProgressBar3)
        Me.Controls.Add(Me.MsVistaProgressBar2)
        Me.Controls.Add(Me.MsVistaProgressBar1)
        Me.Controls.Add(Me.PropertyGrid1)
        Me.Name = "Form1"
        Me.Text = "Form1"
        Me.ResumeLayout(False)

    End Sub
    Friend WithEvents PropertyGrid1 As System.Windows.Forms.PropertyGrid
    Friend WithEvents MsVistaProgressBar1 As MSVistaPBar.MSVistaProgressBar
    Friend WithEvents MsVistaProgressBar2 As MSVistaPBar.MSVistaProgressBar
    Friend WithEvents MsVistaProgressBar3 As MSVistaPBar.MSVistaProgressBar

End Class
