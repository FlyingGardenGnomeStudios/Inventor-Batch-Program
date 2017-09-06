<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class Rename
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
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(Rename))
        Me.DGVRename = New System.Windows.Forms.DataGridView()
        Me.btnThumbs = New System.Windows.Forms.Button()
        Me.btnExcel = New System.Windows.Forms.Button()
        Me.btnRename = New System.Windows.Forms.Button()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.btnImport = New System.Windows.Forms.Button()
        Me.txtParent = New System.Windows.Forms.Label()
        Me.Remaining = New System.Windows.Forms.Label()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.chkCCParts = New System.Windows.Forms.CheckBox()
        Me.chkPParts = New System.Windows.Forms.CheckBox()
        Me.chkDParts = New System.Windows.Forms.CheckBox()
        Me.RenameProgress = New MSVistaProgressBar()
        Me.txtParentSource = New System.Windows.Forms.Label()
        Me.chkStructure = New System.Windows.Forms.CheckBox()
        Me.FileLocation = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.Drawing = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.Part = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.NewName = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.Thumbnail = New System.Windows.Forms.DataGridViewImageColumn()
        Me.ID = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.Reuse = New System.Windows.Forms.DataGridViewCheckBoxColumn()
        CType(Me.DGVRename, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'DGVRename
        '
        Me.DGVRename.AllowDrop = True
        Me.DGVRename.AllowUserToAddRows = False
        Me.DGVRename.AllowUserToResizeRows = False
        Me.DGVRename.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.DGVRename.Columns.AddRange(New System.Windows.Forms.DataGridViewColumn() {Me.FileLocation, Me.Drawing, Me.Part, Me.NewName, Me.Thumbnail, Me.ID, Me.Reuse})
        Me.DGVRename.Location = New System.Drawing.Point(27, 25)
        Me.DGVRename.Name = "DGVRename"
        Me.DGVRename.RowHeadersVisible = False
        Me.DGVRename.Size = New System.Drawing.Size(1169, 387)
        Me.DGVRename.TabIndex = 0
        '
        'btnThumbs
        '
        Me.btnThumbs.Location = New System.Drawing.Point(1214, 25)
        Me.btnThumbs.Name = "btnThumbs"
        Me.btnThumbs.Size = New System.Drawing.Size(96, 23)
        Me.btnThumbs.TabIndex = 1
        Me.btnThumbs.Text = "Large Thumbs"
        Me.btnThumbs.UseVisualStyleBackColor = True
        '
        'btnExcel
        '
        Me.btnExcel.Location = New System.Drawing.Point(1214, 55)
        Me.btnExcel.Name = "btnExcel"
        Me.btnExcel.Size = New System.Drawing.Size(96, 23)
        Me.btnExcel.TabIndex = 2
        Me.btnExcel.Text = "Export to Excel"
        Me.btnExcel.UseVisualStyleBackColor = True
        '
        'btnRename
        '
        Me.btnRename.Location = New System.Drawing.Point(1214, 388)
        Me.btnRename.Name = "btnRename"
        Me.btnRename.Size = New System.Drawing.Size(96, 23)
        Me.btnRename.TabIndex = 3
        Me.btnRename.Text = "Rename"
        Me.btnRename.UseVisualStyleBackColor = True
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(24, 9)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(549, 13)
        Me.Label1.TabIndex = 4
        Me.Label1.Text = "Parts can be rearranged by click-and-drag. if you need a more robust method of or" &
    "ganization export this list to excel."
        '
        'btnImport
        '
        Me.btnImport.Location = New System.Drawing.Point(1214, 84)
        Me.btnImport.Name = "btnImport"
        Me.btnImport.Size = New System.Drawing.Size(96, 23)
        Me.btnImport.TabIndex = 5
        Me.btnImport.Text = "Import from Excel"
        Me.btnImport.UseVisualStyleBackColor = True
        '
        'txtParent
        '
        Me.txtParent.AutoSize = True
        Me.txtParent.Location = New System.Drawing.Point(603, 9)
        Me.txtParent.Name = "txtParent"
        Me.txtParent.Size = New System.Drawing.Size(38, 13)
        Me.txtParent.TabIndex = 6
        Me.txtParent.Text = "Parent"
        Me.txtParent.Visible = False
        '
        'Remaining
        '
        Me.Remaining.AutoSize = True
        Me.Remaining.Location = New System.Drawing.Point(1202, 427)
        Me.Remaining.Name = "Remaining"
        Me.Remaining.Size = New System.Drawing.Size(81, 13)
        Me.Remaining.TabIndex = 8
        Me.Remaining.Text = "Go get a coffee"
        Me.Remaining.Visible = False
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Location = New System.Drawing.Point(1202, 414)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(75, 13)
        Me.Label2.TabIndex = 9
        Me.Label2.Text = "Est. Time Left:"
        Me.Label2.Visible = False
        '
        'chkCCParts
        '
        Me.chkCCParts.AutoSize = True
        Me.chkCCParts.Checked = True
        Me.chkCCParts.CheckState = System.Windows.Forms.CheckState.Checked
        Me.chkCCParts.Location = New System.Drawing.Point(1214, 113)
        Me.chkCCParts.Name = "chkCCParts"
        Me.chkCCParts.Size = New System.Drawing.Size(101, 30)
        Me.chkCCParts.TabIndex = 11
        Me.chkCCParts.Text = "Include Content" & Global.Microsoft.VisualBasic.ChrW(13) & Global.Microsoft.VisualBasic.ChrW(10) & "Center Parts"
        Me.chkCCParts.UseVisualStyleBackColor = True
        '
        'chkPParts
        '
        Me.chkPParts.AutoSize = True
        Me.chkPParts.Checked = True
        Me.chkPParts.CheckState = System.Windows.Forms.CheckState.Checked
        Me.chkPParts.Location = New System.Drawing.Point(1214, 141)
        Me.chkPParts.Name = "chkPParts"
        Me.chkPParts.Size = New System.Drawing.Size(104, 30)
        Me.chkPParts.TabIndex = 12
        Me.chkPParts.Text = "Include" & Global.Microsoft.VisualBasic.ChrW(13) & Global.Microsoft.VisualBasic.ChrW(10) & "Purchased Parts"
        Me.chkPParts.UseVisualStyleBackColor = True
        '
        'chkDParts
        '
        Me.chkDParts.AutoSize = True
        Me.chkDParts.Checked = True
        Me.chkDParts.CheckState = System.Windows.Forms.CheckState.Checked
        Me.chkDParts.Location = New System.Drawing.Point(1214, 168)
        Me.chkDParts.Name = "chkDParts"
        Me.chkDParts.Size = New System.Drawing.Size(90, 30)
        Me.chkDParts.TabIndex = 13
        Me.chkDParts.Text = "Include " & Global.Microsoft.VisualBasic.ChrW(13) & Global.Microsoft.VisualBasic.ChrW(10) & "Derived Parts"
        Me.chkDParts.UseVisualStyleBackColor = True
        '
        'RenameProgress
        '
        Me.RenameProgress.BackColor = System.Drawing.Color.Transparent
        Me.RenameProgress.BlockSize = 10
        Me.RenameProgress.BlockSpacing = 10
        Me.RenameProgress.DisplayText = "%P%"
        Me.RenameProgress.DisplayTextColor = System.Drawing.SystemColors.ControlText
        Me.RenameProgress.DisplayTextFont = New System.Drawing.Font("Arial", 8.0!)
        Me.RenameProgress.GradiantStyle = MSVistaProgressBar.BackGradiant.None
        Me.RenameProgress.Location = New System.Drawing.Point(27, 414)
        Me.RenameProgress.Name = "RenameProgress"
        Me.RenameProgress.ShowText = True
        Me.RenameProgress.Size = New System.Drawing.Size(1169, 22)
        Me.RenameProgress.TabIndex = 10
        Me.RenameProgress.Visible = False
        '
        'txtParentSource
        '
        Me.txtParentSource.AutoSize = True
        Me.txtParentSource.Location = New System.Drawing.Point(659, 9)
        Me.txtParentSource.Name = "txtParentSource"
        Me.txtParentSource.Size = New System.Drawing.Size(72, 13)
        Me.txtParentSource.TabIndex = 14
        Me.txtParentSource.Text = "ParentSource"
        Me.txtParentSource.Visible = False
        '
        'chkStructure
        '
        Me.chkStructure.AutoSize = True
        Me.chkStructure.Checked = True
        Me.chkStructure.CheckState = System.Windows.Forms.CheckState.Checked
        Me.chkStructure.Location = New System.Drawing.Point(1214, 352)
        Me.chkStructure.Name = "chkStructure"
        Me.chkStructure.Size = New System.Drawing.Size(83, 30)
        Me.chkStructure.TabIndex = 15
        Me.chkStructure.Text = "Keep Folder" & Global.Microsoft.VisualBasic.ChrW(13) & Global.Microsoft.VisualBasic.ChrW(10) & "Structure"
        Me.chkStructure.UseVisualStyleBackColor = True
        '
        'FileLocation
        '
        Me.FileLocation.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.Fill
        Me.FileLocation.HeaderText = "File Location:"
        Me.FileLocation.MinimumWidth = 200
        Me.FileLocation.Name = "FileLocation"
        Me.FileLocation.ReadOnly = True
        Me.FileLocation.Resizable = System.Windows.Forms.DataGridViewTriState.[False]
        '
        'Drawing
        '
        Me.Drawing.HeaderText = "Drawing Number:"
        Me.Drawing.MinimumWidth = 50
        Me.Drawing.Name = "Drawing"
        Me.Drawing.ReadOnly = True
        Me.Drawing.Width = 125
        '
        'Part
        '
        Me.Part.HeaderText = "Part Number:"
        Me.Part.MinimumWidth = 50
        Me.Part.Name = "Part"
        Me.Part.ReadOnly = True
        '
        'NewName
        '
        Me.NewName.HeaderText = "New Name:"
        Me.NewName.MinimumWidth = 50
        Me.NewName.Name = "NewName"
        '
        'Thumbnail
        '
        Me.Thumbnail.FillWeight = 125.0!
        Me.Thumbnail.HeaderText = "Thumbnail"
        Me.Thumbnail.ImageLayout = System.Windows.Forms.DataGridViewImageCellLayout.Zoom
        Me.Thumbnail.MinimumWidth = 50
        Me.Thumbnail.Name = "Thumbnail"
        Me.Thumbnail.ReadOnly = True
        Me.Thumbnail.Resizable = System.Windows.Forms.DataGridViewTriState.[True]
        Me.Thumbnail.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.Automatic
        '
        'ID
        '
        Me.ID.HeaderText = "ID"
        Me.ID.MinimumWidth = 50
        Me.ID.Name = "ID"
        Me.ID.ReadOnly = True
        Me.ID.Visible = False
        '
        'Reuse
        '
        Me.Reuse.FalseValue = "False"
        Me.Reuse.HeaderText = "Reuse"
        Me.Reuse.MinimumWidth = 50
        Me.Reuse.Name = "Reuse"
        Me.Reuse.TrueValue = "True"
        Me.Reuse.Width = 50
        '
        'Rename
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(1334, 447)
        Me.Controls.Add(Me.chkStructure)
        Me.Controls.Add(Me.txtParentSource)
        Me.Controls.Add(Me.chkDParts)
        Me.Controls.Add(Me.chkPParts)
        Me.Controls.Add(Me.chkCCParts)
        Me.Controls.Add(Me.RenameProgress)
        Me.Controls.Add(Me.Label2)
        Me.Controls.Add(Me.Remaining)
        Me.Controls.Add(Me.txtParent)
        Me.Controls.Add(Me.btnImport)
        Me.Controls.Add(Me.Label1)
        Me.Controls.Add(Me.btnRename)
        Me.Controls.Add(Me.btnExcel)
        Me.Controls.Add(Me.btnThumbs)
        Me.Controls.Add(Me.DGVRename)
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.MinimumSize = New System.Drawing.Size(900, 350)
        Me.Name = "Rename"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Rename"
        CType(Me.DGVRename, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub

    Friend WithEvents DGVRename As Windows.Forms.DataGridView
    Friend WithEvents btnThumbs As Windows.Forms.Button
    Friend WithEvents btnExcel As Windows.Forms.Button
    Friend WithEvents btnRename As Windows.Forms.Button
    Friend WithEvents Label1 As Windows.Forms.Label
    Friend WithEvents btnImport As Windows.Forms.Button
    Friend WithEvents txtParent As Windows.Forms.Label
    Friend WithEvents Remaining As Windows.Forms.Label
    Friend WithEvents Label2 As Windows.Forms.Label
    Friend WithEvents RenameProgress As MSVistaProgressBar
    Friend WithEvents chkCCParts As Windows.Forms.CheckBox
    Friend WithEvents chkPParts As Windows.Forms.CheckBox
    Friend WithEvents chkDParts As Windows.Forms.CheckBox
    Friend WithEvents txtParentSource As Windows.Forms.Label
    Friend WithEvents chkStructure As Windows.Forms.CheckBox
    Friend WithEvents FileLocation As Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents Drawing As Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents Part As Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents NewName As Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents Thumbnail As Windows.Forms.DataGridViewImageColumn
    Friend WithEvents ID As Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents Reuse As Windows.Forms.DataGridViewCheckBoxColumn
End Class
