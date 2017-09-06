<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class iPropertySettings
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
        Me.iPropertiesTab = New System.Windows.Forms.TabControl()
        Me.SummaryTab = New System.Windows.Forms.TabPage()
        Me.DGVSummary = New System.Windows.Forms.DataGridView()
        Me.iProperty = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.Reference = New System.Windows.Forms.DataGridViewComboBoxColumn()
        Me.ProjectTab = New System.Windows.Forms.TabPage()
        Me.DGVProject = New System.Windows.Forms.DataGridView()
        Me.DataGridViewTextBoxColumn1 = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.PReference = New System.Windows.Forms.DataGridViewComboBoxColumn()
        Me.StatusTab = New System.Windows.Forms.TabPage()
        Me.DGVStatus = New System.Windows.Forms.DataGridView()
        Me.SiProperty = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.SReference = New System.Windows.Forms.DataGridViewComboBoxColumn()
        Me.cmbDefault = New System.Windows.Forms.ComboBox()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.btnOK = New System.Windows.Forms.Button()
        Me.btnCancel = New System.Windows.Forms.Button()
        Me.btnApply = New System.Windows.Forms.Button()
        Me.btnDefaults = New System.Windows.Forms.Button()
        Me.CustomTab = New System.Windows.Forms.TabPage()
        Me.DGVCustom = New System.Windows.Forms.DataGridView()
        Me.DataGridViewTextBoxColumn2 = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.CReference = New System.Windows.Forms.DataGridViewComboBoxColumn()
        Me.iPropertiesTab.SuspendLayout()
        Me.SummaryTab.SuspendLayout()
        CType(Me.DGVSummary, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.ProjectTab.SuspendLayout()
        CType(Me.DGVProject, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.StatusTab.SuspendLayout()
        CType(Me.DGVStatus, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.CustomTab.SuspendLayout()
        CType(Me.DGVCustom, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'iPropertiesTab
        '
        Me.iPropertiesTab.Controls.Add(Me.SummaryTab)
        Me.iPropertiesTab.Controls.Add(Me.ProjectTab)
        Me.iPropertiesTab.Controls.Add(Me.StatusTab)
        Me.iPropertiesTab.Controls.Add(Me.CustomTab)
        Me.iPropertiesTab.Location = New System.Drawing.Point(12, 12)
        Me.iPropertiesTab.Name = "iPropertiesTab"
        Me.iPropertiesTab.SelectedIndex = 0
        Me.iPropertiesTab.Size = New System.Drawing.Size(379, 309)
        Me.iPropertiesTab.TabIndex = 0
        '
        'SummaryTab
        '
        Me.SummaryTab.Controls.Add(Me.DGVSummary)
        Me.SummaryTab.Location = New System.Drawing.Point(4, 22)
        Me.SummaryTab.Name = "SummaryTab"
        Me.SummaryTab.Padding = New System.Windows.Forms.Padding(3)
        Me.SummaryTab.Size = New System.Drawing.Size(371, 283)
        Me.SummaryTab.TabIndex = 0
        Me.SummaryTab.Text = "Summary"
        Me.SummaryTab.UseVisualStyleBackColor = True
        '
        'DGVSummary
        '
        Me.DGVSummary.AllowUserToAddRows = False
        Me.DGVSummary.AllowUserToDeleteRows = False
        Me.DGVSummary.AllowUserToOrderColumns = True
        Me.DGVSummary.AllowUserToResizeColumns = False
        Me.DGVSummary.AllowUserToResizeRows = False
        Me.DGVSummary.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.DGVSummary.Columns.AddRange(New System.Windows.Forms.DataGridViewColumn() {Me.iProperty, Me.Reference})
        Me.DGVSummary.EditMode = System.Windows.Forms.DataGridViewEditMode.EditOnEnter
        Me.DGVSummary.Location = New System.Drawing.Point(6, 6)
        Me.DGVSummary.MultiSelect = False
        Me.DGVSummary.Name = "DGVSummary"
        Me.DGVSummary.RowHeadersVisible = False
        Me.DGVSummary.SelectionMode = System.Windows.Forms.DataGridViewSelectionMode.CellSelect
        Me.DGVSummary.Size = New System.Drawing.Size(359, 271)
        Me.DGVSummary.TabIndex = 0
        '
        'iProperty
        '
        Me.iProperty.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.Fill
        Me.iProperty.HeaderText = "iProperty"
        Me.iProperty.Name = "iProperty"
        Me.iProperty.ReadOnly = True
        '
        'Reference
        '
        Me.Reference.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.Fill
        Me.Reference.HeaderText = "Reference"
        Me.Reference.Items.AddRange(New Object() {"Drawing", "Model"})
        Me.Reference.Name = "Reference"
        '
        'ProjectTab
        '
        Me.ProjectTab.Controls.Add(Me.DGVProject)
        Me.ProjectTab.Location = New System.Drawing.Point(4, 22)
        Me.ProjectTab.Name = "ProjectTab"
        Me.ProjectTab.Padding = New System.Windows.Forms.Padding(3)
        Me.ProjectTab.Size = New System.Drawing.Size(371, 283)
        Me.ProjectTab.TabIndex = 1
        Me.ProjectTab.Text = "Project"
        Me.ProjectTab.UseVisualStyleBackColor = True
        '
        'DGVProject
        '
        Me.DGVProject.AllowUserToAddRows = False
        Me.DGVProject.AllowUserToDeleteRows = False
        Me.DGVProject.AllowUserToOrderColumns = True
        Me.DGVProject.AllowUserToResizeColumns = False
        Me.DGVProject.AllowUserToResizeRows = False
        Me.DGVProject.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.DGVProject.Columns.AddRange(New System.Windows.Forms.DataGridViewColumn() {Me.DataGridViewTextBoxColumn1, Me.PReference})
        Me.DGVProject.EditMode = System.Windows.Forms.DataGridViewEditMode.EditOnEnter
        Me.DGVProject.Location = New System.Drawing.Point(6, 6)
        Me.DGVProject.MultiSelect = False
        Me.DGVProject.Name = "DGVProject"
        Me.DGVProject.RowHeadersVisible = False
        Me.DGVProject.Size = New System.Drawing.Size(359, 271)
        Me.DGVProject.TabIndex = 1
        '
        'DataGridViewTextBoxColumn1
        '
        Me.DataGridViewTextBoxColumn1.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.Fill
        Me.DataGridViewTextBoxColumn1.HeaderText = "iProperty"
        Me.DataGridViewTextBoxColumn1.Name = "DataGridViewTextBoxColumn1"
        Me.DataGridViewTextBoxColumn1.ReadOnly = True
        '
        'PReference
        '
        Me.PReference.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.Fill
        Me.PReference.HeaderText = "Reference"
        Me.PReference.Items.AddRange(New Object() {"Drawing", "Model"})
        Me.PReference.Name = "PReference"
        '
        'StatusTab
        '
        Me.StatusTab.Controls.Add(Me.DGVStatus)
        Me.StatusTab.Location = New System.Drawing.Point(4, 22)
        Me.StatusTab.Name = "StatusTab"
        Me.StatusTab.Padding = New System.Windows.Forms.Padding(3)
        Me.StatusTab.Size = New System.Drawing.Size(371, 283)
        Me.StatusTab.TabIndex = 2
        Me.StatusTab.Text = "Status"
        Me.StatusTab.UseVisualStyleBackColor = True
        '
        'DGVStatus
        '
        Me.DGVStatus.AllowUserToAddRows = False
        Me.DGVStatus.AllowUserToDeleteRows = False
        Me.DGVStatus.AllowUserToOrderColumns = True
        Me.DGVStatus.AllowUserToResizeColumns = False
        Me.DGVStatus.AllowUserToResizeRows = False
        Me.DGVStatus.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.DGVStatus.Columns.AddRange(New System.Windows.Forms.DataGridViewColumn() {Me.SiProperty, Me.SReference})
        Me.DGVStatus.EditMode = System.Windows.Forms.DataGridViewEditMode.EditOnEnter
        Me.DGVStatus.Location = New System.Drawing.Point(6, 6)
        Me.DGVStatus.MultiSelect = False
        Me.DGVStatus.Name = "DGVStatus"
        Me.DGVStatus.RowHeadersVisible = False
        Me.DGVStatus.Size = New System.Drawing.Size(359, 271)
        Me.DGVStatus.TabIndex = 1
        '
        'SiProperty
        '
        Me.SiProperty.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.Fill
        Me.SiProperty.HeaderText = "iProperty"
        Me.SiProperty.Name = "SiProperty"
        Me.SiProperty.ReadOnly = True
        '
        'SReference
        '
        Me.SReference.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.Fill
        Me.SReference.HeaderText = "Reference"
        Me.SReference.Items.AddRange(New Object() {"Drawing", "Model"})
        Me.SReference.Name = "SReference"
        '
        'cmbDefault
        '
        Me.cmbDefault.FormattingEnabled = True
        Me.cmbDefault.Items.AddRange(New Object() {"Assembly", "Model"})
        Me.cmbDefault.Location = New System.Drawing.Point(206, 323)
        Me.cmbDefault.Name = "cmbDefault"
        Me.cmbDefault.Size = New System.Drawing.Size(79, 21)
        Me.cmbDefault.TabIndex = 1
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(13, 328)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(187, 13)
        Me.Label1.TabIndex = 2
        Me.Label1.Text = "On identical drawing names default to:"
        '
        'btnOK
        '
        Me.btnOK.Location = New System.Drawing.Point(316, 350)
        Me.btnOK.Name = "btnOK"
        Me.btnOK.Size = New System.Drawing.Size(75, 23)
        Me.btnOK.TabIndex = 3
        Me.btnOK.Text = "OK"
        Me.btnOK.UseVisualStyleBackColor = True
        '
        'btnCancel
        '
        Me.btnCancel.Location = New System.Drawing.Point(235, 350)
        Me.btnCancel.Name = "btnCancel"
        Me.btnCancel.Size = New System.Drawing.Size(75, 23)
        Me.btnCancel.TabIndex = 4
        Me.btnCancel.Text = "Cancel"
        Me.btnCancel.UseVisualStyleBackColor = True
        '
        'btnApply
        '
        Me.btnApply.Location = New System.Drawing.Point(154, 350)
        Me.btnApply.Name = "btnApply"
        Me.btnApply.Size = New System.Drawing.Size(75, 23)
        Me.btnApply.TabIndex = 5
        Me.btnApply.Text = "Apply"
        Me.btnApply.UseVisualStyleBackColor = True
        '
        'btnDefaults
        '
        Me.btnDefaults.Location = New System.Drawing.Point(291, 323)
        Me.btnDefaults.Name = "btnDefaults"
        Me.btnDefaults.Size = New System.Drawing.Size(100, 23)
        Me.btnDefaults.TabIndex = 6
        Me.btnDefaults.Text = "Reset Defaults"
        Me.btnDefaults.UseVisualStyleBackColor = True
        '
        'CustomTab
        '
        Me.CustomTab.Controls.Add(Me.DGVCustom)
        Me.CustomTab.Location = New System.Drawing.Point(4, 22)
        Me.CustomTab.Name = "CustomTab"
        Me.CustomTab.Padding = New System.Windows.Forms.Padding(3)
        Me.CustomTab.Size = New System.Drawing.Size(371, 283)
        Me.CustomTab.TabIndex = 3
        Me.CustomTab.Text = "Custom"
        Me.CustomTab.UseVisualStyleBackColor = True
        '
        'DGVCustom
        '
        Me.DGVCustom.AllowUserToAddRows = False
        Me.DGVCustom.AllowUserToDeleteRows = False
        Me.DGVCustom.AllowUserToOrderColumns = True
        Me.DGVCustom.AllowUserToResizeColumns = False
        Me.DGVCustom.AllowUserToResizeRows = False
        Me.DGVCustom.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.DGVCustom.Columns.AddRange(New System.Windows.Forms.DataGridViewColumn() {Me.DataGridViewTextBoxColumn2, Me.CReference})
        Me.DGVCustom.EditMode = System.Windows.Forms.DataGridViewEditMode.EditOnEnter
        Me.DGVCustom.Location = New System.Drawing.Point(6, 6)
        Me.DGVCustom.MultiSelect = False
        Me.DGVCustom.Name = "DGVCustom"
        Me.DGVCustom.RowHeadersVisible = False
        Me.DGVCustom.Size = New System.Drawing.Size(359, 271)
        Me.DGVCustom.TabIndex = 2
        '
        'DataGridViewTextBoxColumn2
        '
        Me.DataGridViewTextBoxColumn2.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.Fill
        Me.DataGridViewTextBoxColumn2.HeaderText = "iProperty"
        Me.DataGridViewTextBoxColumn2.Name = "DataGridViewTextBoxColumn2"
        Me.DataGridViewTextBoxColumn2.ReadOnly = True
        '
        'CReference
        '
        Me.CReference.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.Fill
        Me.CReference.HeaderText = "Reference"
        Me.CReference.Items.AddRange(New Object() {"Drawing", "Model"})
        Me.CReference.Name = "CReference"
        '
        'iPropertySettings
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(403, 393)
        Me.Controls.Add(Me.btnDefaults)
        Me.Controls.Add(Me.btnApply)
        Me.Controls.Add(Me.btnCancel)
        Me.Controls.Add(Me.btnOK)
        Me.Controls.Add(Me.Label1)
        Me.Controls.Add(Me.cmbDefault)
        Me.Controls.Add(Me.iPropertiesTab)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedToolWindow
        Me.Name = "iPropertySettings"
        Me.Text = "iPropertySettings"
        Me.iPropertiesTab.ResumeLayout(False)
        Me.SummaryTab.ResumeLayout(False)
        CType(Me.DGVSummary, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ProjectTab.ResumeLayout(False)
        CType(Me.DGVProject, System.ComponentModel.ISupportInitialize).EndInit()
        Me.StatusTab.ResumeLayout(False)
        CType(Me.DGVStatus, System.ComponentModel.ISupportInitialize).EndInit()
        Me.CustomTab.ResumeLayout(False)
        CType(Me.DGVCustom, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub

    Friend WithEvents iPropertiesTab As Windows.Forms.TabControl
    Friend WithEvents SummaryTab As Windows.Forms.TabPage
    Friend WithEvents ProjectTab As Windows.Forms.TabPage
    Friend WithEvents StatusTab As Windows.Forms.TabPage
    Friend WithEvents cmbDefault As Windows.Forms.ComboBox
    Friend WithEvents Label1 As Windows.Forms.Label
    Friend WithEvents btnOK As Windows.Forms.Button
    Friend WithEvents btnCancel As Windows.Forms.Button
    Friend WithEvents btnApply As Windows.Forms.Button
    Friend WithEvents DGVSummary As Windows.Forms.DataGridView
    Friend WithEvents iProperty As Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents Reference As Windows.Forms.DataGridViewComboBoxColumn
    Friend WithEvents DGVProject As Windows.Forms.DataGridView
    Friend WithEvents DGVStatus As Windows.Forms.DataGridView
    Friend WithEvents btnDefaults As Windows.Forms.Button
    Friend WithEvents DataGridViewTextBoxColumn1 As Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents PReference As Windows.Forms.DataGridViewComboBoxColumn
    Friend WithEvents SiProperty As Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents SReference As Windows.Forms.DataGridViewComboBoxColumn
    Friend WithEvents CustomTab As Windows.Forms.TabPage
    Friend WithEvents DGVCustom As Windows.Forms.DataGridView
    Friend WithEvents DataGridViewTextBoxColumn2 As Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents CReference As Windows.Forms.DataGridViewComboBoxColumn
End Class
