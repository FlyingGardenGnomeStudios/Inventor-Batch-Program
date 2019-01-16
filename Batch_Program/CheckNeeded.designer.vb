<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()>
Partial Class CheckNeeded
    Inherits System.Windows.Forms.Form

    'Form overrides dispose to clean up the component list.
    <System.Diagnostics.DebuggerNonUserCode()>
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
    <System.Diagnostics.DebuggerStepThrough()>
    Private Sub InitializeComponent()
        Me.components = New System.ComponentModel.Container()
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(CheckNeeded))
        Me.btnCancel = New System.Windows.Forms.Button()
        Me.btnHide = New System.Windows.Forms.Button()
        Me.btnOK = New System.Windows.Forms.Button()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.ToolTip1 = New System.Windows.Forms.ToolTip(Me.components)
        Me.ProgressBar1 = New System.Windows.Forms.ProgressBar()
        Me.cmsApplyValues = New System.Windows.Forms.ContextMenuStrip(Me.components)
        Me.ApplyRowValuesToAllRowsToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.VisibleRowsToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.EmptyValuesToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.EmptyCellsOnlyToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.ApplyCellValueToEntireColumnToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.CellsInColumnToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.VisibleCellsInColumnToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.EmptyCellsInColumnToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.ClearCellToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.RemoveRevisionToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.AddRevisionToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.tgvCheckNeeded = New AdvancedDataGridView.TreeGridView()
        Me.Splitter1 = New System.Windows.Forms.Splitter()
        Me.DataGridViewTextBoxColumn1 = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.DataGridViewTextBoxColumn2 = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.GridDateControl1 = New GridDateControl()
        Me.GridDateControl2 = New GridDateControl()
        Me.DataGridViewTextBoxColumn3 = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.DataGridViewTextBoxColumn4 = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.DataGridViewTextBoxColumn5 = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.DataGridViewTextBoxColumn6 = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.DrawingName = New AdvancedDataGridView.TreeGridColumn()
        Me.IsDirty = New System.Windows.Forms.DataGridViewCheckBoxColumn()
        Me.cmsApplyValues.SuspendLayout()
        CType(Me.tgvCheckNeeded, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'btnCancel
        '
        Me.btnCancel.Location = New System.Drawing.Point(93, 183)
        Me.btnCancel.Name = "btnCancel"
        Me.btnCancel.Size = New System.Drawing.Size(75, 23)
        Me.btnCancel.TabIndex = 14
        Me.btnCancel.Text = "Cancel"
        Me.ToolTip1.SetToolTip(Me.btnCancel, "Exit without making any changes")
        Me.btnCancel.UseVisualStyleBackColor = True
        '
        'btnHide
        '
        Me.btnHide.Location = New System.Drawing.Point(174, 183)
        Me.btnHide.Name = "btnHide"
        Me.btnHide.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.btnHide.Size = New System.Drawing.Size(105, 23)
        Me.btnHide.TabIndex = 13
        Me.btnHide.Text = "Hide Completed"
        Me.ToolTip1.SetToolTip(Me.btnHide, "Disregard changes and proceed with operation")
        Me.btnHide.UseVisualStyleBackColor = True
        '
        'btnOK
        '
        Me.btnOK.Location = New System.Drawing.Point(12, 183)
        Me.btnOK.Name = "btnOK"
        Me.btnOK.Size = New System.Drawing.Size(75, 23)
        Me.btnOK.TabIndex = 12
        Me.btnOK.Text = "OK"
        Me.ToolTip1.SetToolTip(Me.btnOK, "Proceed with changes")
        Me.btnOK.UseVisualStyleBackColor = True
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Location = New System.Drawing.Point(12, 9)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(221, 13)
        Me.Label2.TabIndex = 11
        Me.Label2.Text = "Click to edit. Right-click for additional options."
        '
        'ProgressBar1
        '
        Me.ProgressBar1.Location = New System.Drawing.Point(285, 183)
        Me.ProgressBar1.Name = "ProgressBar1"
        Me.ProgressBar1.Size = New System.Drawing.Size(323, 23)
        Me.ProgressBar1.TabIndex = 21
        Me.ProgressBar1.Visible = False
        '
        'cmsApplyValues
        '
        Me.cmsApplyValues.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.ApplyRowValuesToAllRowsToolStripMenuItem, Me.ApplyCellValueToEntireColumnToolStripMenuItem, Me.ClearCellToolStripMenuItem, Me.RemoveRevisionToolStripMenuItem, Me.AddRevisionToolStripMenuItem})
        Me.cmsApplyValues.Name = "ContextMenuStrip1"
        Me.cmsApplyValues.Size = New System.Drawing.Size(243, 114)
        '
        'ApplyRowValuesToAllRowsToolStripMenuItem
        '
        Me.ApplyRowValuesToAllRowsToolStripMenuItem.DropDownItems.AddRange(New System.Windows.Forms.ToolStripItem() {Me.VisibleRowsToolStripMenuItem, Me.EmptyValuesToolStripMenuItem, Me.EmptyCellsOnlyToolStripMenuItem})
        Me.ApplyRowValuesToAllRowsToolStripMenuItem.Name = "ApplyRowValuesToAllRowsToolStripMenuItem"
        Me.ApplyRowValuesToAllRowsToolStripMenuItem.Size = New System.Drawing.Size(242, 22)
        Me.ApplyRowValuesToAllRowsToolStripMenuItem.Text = "Apply selected row values to all:"
        '
        'VisibleRowsToolStripMenuItem
        '
        Me.VisibleRowsToolStripMenuItem.Name = "VisibleRowsToolStripMenuItem"
        Me.VisibleRowsToolStripMenuItem.Size = New System.Drawing.Size(139, 22)
        Me.VisibleRowsToolStripMenuItem.Text = "Rows"
        '
        'EmptyValuesToolStripMenuItem
        '
        Me.EmptyValuesToolStripMenuItem.Name = "EmptyValuesToolStripMenuItem"
        Me.EmptyValuesToolStripMenuItem.Size = New System.Drawing.Size(139, 22)
        Me.EmptyValuesToolStripMenuItem.Text = "Visible Rows"
        '
        'EmptyCellsOnlyToolStripMenuItem
        '
        Me.EmptyCellsOnlyToolStripMenuItem.Name = "EmptyCellsOnlyToolStripMenuItem"
        Me.EmptyCellsOnlyToolStripMenuItem.Size = New System.Drawing.Size(139, 22)
        Me.EmptyCellsOnlyToolStripMenuItem.Text = "Empty Cells"
        '
        'ApplyCellValueToEntireColumnToolStripMenuItem
        '
        Me.ApplyCellValueToEntireColumnToolStripMenuItem.DropDownItems.AddRange(New System.Windows.Forms.ToolStripItem() {Me.CellsInColumnToolStripMenuItem, Me.VisibleCellsInColumnToolStripMenuItem, Me.EmptyCellsInColumnToolStripMenuItem})
        Me.ApplyCellValueToEntireColumnToolStripMenuItem.Name = "ApplyCellValueToEntireColumnToolStripMenuItem"
        Me.ApplyCellValueToEntireColumnToolStripMenuItem.Size = New System.Drawing.Size(242, 22)
        Me.ApplyCellValueToEntireColumnToolStripMenuItem.Text = "Apply selected cell value to all:"
        '
        'CellsInColumnToolStripMenuItem
        '
        Me.CellsInColumnToolStripMenuItem.Name = "CellsInColumnToolStripMenuItem"
        Me.CellsInColumnToolStripMenuItem.Size = New System.Drawing.Size(195, 22)
        Me.CellsInColumnToolStripMenuItem.Text = "Cells in Column"
        '
        'VisibleCellsInColumnToolStripMenuItem
        '
        Me.VisibleCellsInColumnToolStripMenuItem.Name = "VisibleCellsInColumnToolStripMenuItem"
        Me.VisibleCellsInColumnToolStripMenuItem.Size = New System.Drawing.Size(195, 22)
        Me.VisibleCellsInColumnToolStripMenuItem.Text = "Visible Cells in Column"
        '
        'EmptyCellsInColumnToolStripMenuItem
        '
        Me.EmptyCellsInColumnToolStripMenuItem.Name = "EmptyCellsInColumnToolStripMenuItem"
        Me.EmptyCellsInColumnToolStripMenuItem.Size = New System.Drawing.Size(195, 22)
        Me.EmptyCellsInColumnToolStripMenuItem.Text = "Empty Cells in Column"
        '
        'ClearCellToolStripMenuItem
        '
        Me.ClearCellToolStripMenuItem.Name = "ClearCellToolStripMenuItem"
        Me.ClearCellToolStripMenuItem.Size = New System.Drawing.Size(242, 22)
        Me.ClearCellToolStripMenuItem.Text = "Clear Cell"
        '
        'RemoveRevisionToolStripMenuItem
        '
        Me.RemoveRevisionToolStripMenuItem.Name = "RemoveRevisionToolStripMenuItem"
        Me.RemoveRevisionToolStripMenuItem.Size = New System.Drawing.Size(242, 22)
        Me.RemoveRevisionToolStripMenuItem.Text = "Remove Revision"
        '
        'AddRevisionToolStripMenuItem
        '
        Me.AddRevisionToolStripMenuItem.Name = "AddRevisionToolStripMenuItem"
        Me.AddRevisionToolStripMenuItem.Size = New System.Drawing.Size(242, 22)
        Me.AddRevisionToolStripMenuItem.Text = "Add Revision"
        '
        'tgvCheckNeeded
        '
        Me.tgvCheckNeeded.AllowUserToAddRows = False
        Me.tgvCheckNeeded.AllowUserToDeleteRows = False
        Me.tgvCheckNeeded.AllowUserToResizeRows = False
        Me.tgvCheckNeeded.AutoSizeColumnsMode = System.Windows.Forms.DataGridViewAutoSizeColumnsMode.ColumnHeader
        Me.tgvCheckNeeded.Columns.AddRange(New System.Windows.Forms.DataGridViewColumn() {Me.DrawingName, Me.IsDirty})
        Me.tgvCheckNeeded.EditMode = System.Windows.Forms.DataGridViewEditMode.EditOnEnter
        Me.tgvCheckNeeded.ImageList = Nothing
        Me.tgvCheckNeeded.Location = New System.Drawing.Point(12, 25)
        Me.tgvCheckNeeded.MultiSelect = False
        Me.tgvCheckNeeded.Name = "tgvCheckNeeded"
        Me.tgvCheckNeeded.RowHeadersVisible = False
        Me.tgvCheckNeeded.Size = New System.Drawing.Size(596, 152)
        Me.tgvCheckNeeded.TabIndex = 23
        '
        'Splitter1
        '
        Me.Splitter1.Location = New System.Drawing.Point(0, 0)
        Me.Splitter1.Name = "Splitter1"
        Me.Splitter1.Size = New System.Drawing.Size(3, 218)
        Me.Splitter1.TabIndex = 24
        Me.Splitter1.TabStop = False
        '
        'DataGridViewTextBoxColumn1
        '
        Me.DataGridViewTextBoxColumn1.HeaderText = "Drawing Number"
        Me.DataGridViewTextBoxColumn1.Name = "DataGridViewTextBoxColumn1"
        Me.DataGridViewTextBoxColumn1.Width = 69
        '
        'DataGridViewTextBoxColumn2
        '
        Me.DataGridViewTextBoxColumn2.HeaderText = "Revision"
        Me.DataGridViewTextBoxColumn2.Name = "DataGridViewTextBoxColumn2"
        Me.DataGridViewTextBoxColumn2.Visible = False
        Me.DataGridViewTextBoxColumn2.Width = 69
        '
        'GridDateControl1
        '
        Me.GridDateControl1.HeaderText = "Check Date"
        Me.GridDateControl1.Name = "GridDateControl1"
        Me.GridDateControl1.Resizable = System.Windows.Forms.DataGridViewTriState.[True]
        Me.GridDateControl1.Width = 84
        '
        'GridDateControl2
        '
        Me.GridDateControl2.HeaderText = "Rev Date"
        Me.GridDateControl2.Name = "GridDateControl2"
        Me.GridDateControl2.Resizable = System.Windows.Forms.DataGridViewTriState.[True]
        Me.GridDateControl2.Width = 85
        '
        'DataGridViewTextBoxColumn3
        '
        Me.DataGridViewTextBoxColumn3.HeaderText = "Description"
        Me.DataGridViewTextBoxColumn3.Name = "DataGridViewTextBoxColumn3"
        Me.DataGridViewTextBoxColumn3.Width = 69
        '
        'DataGridViewTextBoxColumn4
        '
        Me.DataGridViewTextBoxColumn4.HeaderText = "RevisionBy"
        Me.DataGridViewTextBoxColumn4.Name = "DataGridViewTextBoxColumn4"
        Me.DataGridViewTextBoxColumn4.Width = 69
        '
        'DataGridViewTextBoxColumn5
        '
        Me.DataGridViewTextBoxColumn5.HeaderText = "Approved By"
        Me.DataGridViewTextBoxColumn5.Name = "DataGridViewTextBoxColumn5"
        Me.DataGridViewTextBoxColumn5.Width = 69
        '
        'DataGridViewTextBoxColumn6
        '
        Me.DataGridViewTextBoxColumn6.HeaderText = "MoreRevs"
        Me.DataGridViewTextBoxColumn6.Name = "DataGridViewTextBoxColumn6"
        Me.DataGridViewTextBoxColumn6.Width = 69
        '
        'DrawingName
        '
        Me.DrawingName.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.AllCells
        Me.DrawingName.DataPropertyName = "Drawing Name"
        Me.DrawingName.DefaultNodeImage = Nothing
        Me.DrawingName.HeaderText = "Drawing Name"
        Me.DrawingName.Name = "DrawingName"
        Me.DrawingName.ReadOnly = True
        Me.DrawingName.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.DrawingName.Width = 83
        '
        'IsDirty
        '
        Me.IsDirty.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.AllCells
        Me.IsDirty.HeaderText = "IsDirty"
        Me.IsDirty.Name = "IsDirty"
        Me.IsDirty.ReadOnly = True
        Me.IsDirty.Visible = False
        Me.IsDirty.Width = 42
        '
        'CheckNeeded
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(624, 218)
        Me.Controls.Add(Me.Splitter1)
        Me.Controls.Add(Me.tgvCheckNeeded)
        Me.Controls.Add(Me.ProgressBar1)
        Me.Controls.Add(Me.btnCancel)
        Me.Controls.Add(Me.btnHide)
        Me.Controls.Add(Me.btnOK)
        Me.Controls.Add(Me.Label2)
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.MinimumSize = New System.Drawing.Size(640, 256)
        Me.Name = "CheckNeeded"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterParent
        Me.Text = "CheckNeeded"
        Me.cmsApplyValues.ResumeLayout(False)
        CType(Me.tgvCheckNeeded, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents btnCancel As System.Windows.Forms.Button
    Friend WithEvents btnHide As System.Windows.Forms.Button
    Friend WithEvents btnOK As System.Windows.Forms.Button
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents ToolTip1 As System.Windows.Forms.ToolTip
    Friend WithEvents ProgressBar1 As System.Windows.Forms.ProgressBar
    Friend WithEvents cmsApplyValues As Windows.Forms.ContextMenuStrip
    Friend WithEvents ApplyRowValuesToAllRowsToolStripMenuItem As Windows.Forms.ToolStripMenuItem
    Friend WithEvents ApplyCellValueToEntireColumnToolStripMenuItem As Windows.Forms.ToolStripMenuItem
    Friend WithEvents tgvCheckNeeded As AdvancedDataGridView.TreeGridView
    Friend WithEvents DataGridViewTextBoxColumn1 As Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents DataGridViewTextBoxColumn2 As Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents DataGridViewTextBoxColumn3 As Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents DataGridViewTextBoxColumn4 As Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents DataGridViewTextBoxColumn5 As Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents DataGridViewTextBoxColumn6 As Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents Splitter1 As Windows.Forms.Splitter
    Friend WithEvents GridDateControl1 As GridDateControl
    Friend WithEvents GridDateControl2 As GridDateControl
    Friend WithEvents ClearCellToolStripMenuItem As Windows.Forms.ToolStripMenuItem
    Friend WithEvents RemoveRevisionToolStripMenuItem As Windows.Forms.ToolStripMenuItem
    Friend WithEvents ToolStripMenuItem1 As Windows.Forms.ToolStripMenuItem
    Friend WithEvents AddRevisionToolStripMenuItem As Windows.Forms.ToolStripMenuItem
    Friend WithEvents VisibleRowsToolStripMenuItem As Windows.Forms.ToolStripMenuItem
    Friend WithEvents EmptyValuesToolStripMenuItem As Windows.Forms.ToolStripMenuItem
    Friend WithEvents EmptyCellsOnlyToolStripMenuItem As Windows.Forms.ToolStripMenuItem
    Friend WithEvents CellsInColumnToolStripMenuItem As Windows.Forms.ToolStripMenuItem
    Friend WithEvents VisibleCellsInColumnToolStripMenuItem As Windows.Forms.ToolStripMenuItem
    Friend WithEvents EmptyCellsInColumnToolStripMenuItem As Windows.Forms.ToolStripMenuItem
    Friend WithEvents DrawingName As AdvancedDataGridView.TreeGridColumn
    Friend WithEvents IsDirty As Windows.Forms.DataGridViewCheckBoxColumn
End Class
