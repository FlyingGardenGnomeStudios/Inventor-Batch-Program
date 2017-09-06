<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()>
Partial Class Main
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
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(Main))
        Me.chkPres = New System.Windows.Forms.CheckBox()
        Me.chkDerived = New System.Windows.Forms.CheckBox()
        Me.Label3 = New System.Windows.Forms.Label()
        Me.btnExit = New System.Windows.Forms.Button()
        Me.btnOK = New System.Windows.Forms.Button()
        Me.chkClose = New System.Windows.Forms.CheckBox()
        Me.chkExport = New System.Windows.Forms.CheckBox()
        Me.chkDXF = New System.Windows.Forms.CheckBox()
        Me.chkPrint = New System.Windows.Forms.CheckBox()
        Me.chkCheck = New System.Windows.Forms.CheckBox()
        Me.chkOpen = New System.Windows.Forms.CheckBox()
        Me.chkiProp = New System.Windows.Forms.CheckBox()
        Me.chkAssy = New System.Windows.Forms.CheckBox()
        Me.chkParts = New System.Windows.Forms.CheckBox()
        Me.chkDrawings = New System.Windows.Forms.CheckBox()
        Me.GroupBox1 = New System.Windows.Forms.GroupBox()
        Me.GroupBox2 = New System.Windows.Forms.GroupBox()
        Me.PictureBox1 = New System.Windows.Forms.PictureBox()
        Me.chkPartSelect = New System.Windows.Forms.CheckBox()
        Me.lstOpenfiles = New System.Windows.Forms.CheckedListBox()
        Me.GroupBox4 = New System.Windows.Forms.GroupBox()
        Me.chkUseDrawings = New System.Windows.Forms.CheckBox()
        Me.chkSkipAssy = New System.Windows.Forms.CheckBox()
        Me.chkRRev = New System.Windows.Forms.CheckBox()
        Me.ChkRevType = New System.Windows.Forms.CheckBox()
        Me.chkDWG = New System.Windows.Forms.CheckBox()
        Me.chkPDF = New System.Windows.Forms.CheckBox()
        Me.btnRename = New System.Windows.Forms.Button()
        Me.btnFlatPattern = New System.Windows.Forms.Button()
        Me.btnSpreadsheet = New System.Windows.Forms.Button()
        Me.GroupBox3 = New System.Windows.Forms.GroupBox()
        Me.LVSubFiles = New System.Windows.Forms.ListView()
        Me.ColumnHeader1 = CType(New System.Windows.Forms.ColumnHeader(), System.Windows.Forms.ColumnHeader)
        Me.txtSearch = New System.Windows.Forms.TextBox()
        Me.chkDWGSelect = New System.Windows.Forms.CheckBox()
        Me.tmr = New System.Windows.Forms.Timer(Me.components)
        Me.ToolTip1 = New System.Windows.Forms.ToolTip(Me.components)
        Me.btnRef = New System.Windows.Forms.Button()
        Me.PictureBox2 = New System.Windows.Forms.PictureBox()
        Me.GroupBox5 = New System.Windows.Forms.GroupBox()
        Me.CMSSubFiles = New System.Windows.Forms.ContextMenuStrip(Me.components)
        Me.SortToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.CMSAlphabetical = New System.Windows.Forms.ToolStripMenuItem()
        Me.CMSHeirarchical = New System.Windows.Forms.ToolStripMenuItem()
        Me.ShowHideToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.CMSShow = New System.Windows.Forms.ToolStripMenuItem()
        Me.CMSHide = New System.Windows.Forms.ToolStripMenuItem()
        Me.ExportToToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.CMSSubSpreadsheet = New System.Windows.Forms.ToolStripMenuItem()
        Me.CMSSubText = New System.Windows.Forms.ToolStripMenuItem()
        Me.CMSOpenFiles = New System.Windows.Forms.ContextMenuStrip(Me.components)
        Me.ExportToToolStripMenuItem1 = New System.Windows.Forms.ToolStripMenuItem()
        Me.CMSSpreadsheet = New System.Windows.Forms.ToolStripMenuItem()
        Me.CMSTextFile = New System.Windows.Forms.ToolStripMenuItem()
        Me.MenuStrip1 = New System.Windows.Forms.MenuStrip()
        Me.FileToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.DefaultSettingsToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.IPropertySettingsToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.AboutToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.AboutBatchProgramToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.HowToToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.ContextMenuStrip1 = New System.Windows.Forms.ContextMenuStrip(Me.components)
        Me.MsVistaProgressBar1 = New MSVistaProgressBar()
        Me.GroupBox1.SuspendLayout()
        Me.GroupBox2.SuspendLayout()
        CType(Me.PictureBox1, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.GroupBox4.SuspendLayout()
        Me.GroupBox3.SuspendLayout()
        CType(Me.PictureBox2, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.GroupBox5.SuspendLayout()
        Me.CMSSubFiles.SuspendLayout()
        Me.CMSOpenFiles.SuspendLayout()
        Me.MenuStrip1.SuspendLayout()
        Me.SuspendLayout()
        '
        'chkPres
        '
        Me.chkPres.AutoSize = True
        Me.chkPres.Location = New System.Drawing.Point(6, 60)
        Me.chkPres.Name = "chkPres"
        Me.chkPres.Size = New System.Drawing.Size(90, 17)
        Me.chkPres.TabIndex = 49
        Me.chkPres.Text = "Presentations"
        Me.ToolTip1.SetToolTip(Me.chkPres, "Show the presentations that are currently open")
        Me.chkPres.UseVisualStyleBackColor = True
        '
        'chkDerived
        '
        Me.chkDerived.AutoSize = True
        Me.chkDerived.Location = New System.Drawing.Point(17, 44)
        Me.chkDerived.Name = "chkDerived"
        Me.chkDerived.Size = New System.Drawing.Size(101, 17)
        Me.chkDerived.TabIndex = 47
        Me.chkDerived.Text = "Multi-body Parts"
        Me.ToolTip1.SetToolTip(Me.chkDerived, "Display parts with derived components in the Referenced Drawing window")
        Me.chkDerived.UseVisualStyleBackColor = True
        Me.chkDerived.Visible = False
        '
        'Label3
        '
        Me.Label3.AutoSize = True
        Me.Label3.Enabled = False
        Me.Label3.FlatStyle = System.Windows.Forms.FlatStyle.System
        Me.Label3.Location = New System.Drawing.Point(57, 79)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(63, 13)
        Me.Label3.TabIndex = 46
        Me.Label3.Text = "Print Copies"
        '
        'btnExit
        '
        Me.btnExit.FlatStyle = System.Windows.Forms.FlatStyle.System
        Me.btnExit.Location = New System.Drawing.Point(566, 240)
        Me.btnExit.Name = "btnExit"
        Me.btnExit.Size = New System.Drawing.Size(75, 23)
        Me.btnExit.TabIndex = 45
        Me.btnExit.Text = "Exit"
        Me.btnExit.UseVisualStyleBackColor = True
        '
        'btnOK
        '
        Me.btnOK.FlatStyle = System.Windows.Forms.FlatStyle.System
        Me.btnOK.Location = New System.Drawing.Point(647, 240)
        Me.btnOK.Name = "btnOK"
        Me.btnOK.Size = New System.Drawing.Size(75, 23)
        Me.btnOK.TabIndex = 44
        Me.btnOK.Text = "OK"
        Me.btnOK.UseVisualStyleBackColor = True
        '
        'chkClose
        '
        Me.chkClose.AutoSize = True
        Me.chkClose.Location = New System.Drawing.Point(10, 121)
        Me.chkClose.Name = "chkClose"
        Me.chkClose.Size = New System.Drawing.Size(138, 17)
        Me.chkClose.TabIndex = 43
        Me.chkClose.Text = "Close Open Documents"
        Me.ToolTip1.SetToolTip(Me.chkClose, "Close all the drawings selected")
        Me.chkClose.UseVisualStyleBackColor = True
        '
        'chkExport
        '
        Me.chkExport.AutoSize = True
        Me.chkExport.Location = New System.Drawing.Point(10, 105)
        Me.chkExport.Name = "chkExport"
        Me.chkExport.Size = New System.Drawing.Size(98, 17)
        Me.chkExport.TabIndex = 42
        Me.chkExport.Text = "Export Drawing"
        Me.ToolTip1.SetToolTip(Me.chkExport, "Create PDF's of the drawings selected")
        Me.chkExport.UseVisualStyleBackColor = True
        '
        'chkDXF
        '
        Me.chkDXF.AutoSize = True
        Me.chkDXF.Location = New System.Drawing.Point(113, 121)
        Me.chkDXF.Name = "chkDXF"
        Me.chkDXF.Size = New System.Drawing.Size(47, 17)
        Me.chkDXF.TabIndex = 41
        Me.chkDXF.Text = "DXF"
        Me.ToolTip1.SetToolTip(Me.chkDXF, "Create DXF's of the drawings selected")
        Me.chkDXF.UseVisualStyleBackColor = True
        Me.chkDXF.Visible = False
        '
        'chkPrint
        '
        Me.chkPrint.AutoSize = True
        Me.chkPrint.Location = New System.Drawing.Point(10, 60)
        Me.chkPrint.Name = "chkPrint"
        Me.chkPrint.Size = New System.Drawing.Size(94, 17)
        Me.chkPrint.TabIndex = 40
        Me.chkPrint.Text = "Print Drawings"
        Me.ToolTip1.SetToolTip(Me.chkPrint, "Print the drawings selected")
        Me.chkPrint.UseVisualStyleBackColor = True
        '
        'chkCheck
        '
        Me.chkCheck.AutoSize = True
        Me.chkCheck.Location = New System.Drawing.Point(10, 45)
        Me.chkCheck.Name = "chkCheck"
        Me.chkCheck.Size = New System.Drawing.Size(104, 17)
        Me.chkCheck.TabIndex = 39
        Me.chkCheck.Text = "Check Drawings"
        Me.ToolTip1.SetToolTip(Me.chkCheck, "Check the revision tables of the drawings selected")
        Me.chkCheck.UseVisualStyleBackColor = True
        '
        'chkOpen
        '
        Me.chkOpen.AutoSize = True
        Me.chkOpen.Location = New System.Drawing.Point(10, 30)
        Me.chkOpen.Name = "chkOpen"
        Me.chkOpen.Size = New System.Drawing.Size(76, 17)
        Me.chkOpen.TabIndex = 38
        Me.chkOpen.Text = "Open Only"
        Me.ToolTip1.SetToolTip(Me.chkOpen, "Open all the drawings selected")
        Me.chkOpen.UseVisualStyleBackColor = True
        '
        'chkiProp
        '
        Me.chkiProp.AutoSize = True
        Me.chkiProp.Location = New System.Drawing.Point(10, 15)
        Me.chkiProp.Name = "chkiProp"
        Me.chkiProp.Size = New System.Drawing.Size(115, 17)
        Me.chkiProp.TabIndex = 37
        Me.chkiProp.Text = "Change iProperties"
        Me.ToolTip1.SetToolTip(Me.chkiProp, "Modify the iProperties on all the drawings selected")
        Me.chkiProp.UseVisualStyleBackColor = True
        '
        'chkAssy
        '
        Me.chkAssy.AutoSize = True
        Me.chkAssy.Location = New System.Drawing.Point(6, 44)
        Me.chkAssy.Name = "chkAssy"
        Me.chkAssy.Size = New System.Drawing.Size(78, 17)
        Me.chkAssy.TabIndex = 36
        Me.chkAssy.Text = "Assemblies"
        Me.ToolTip1.SetToolTip(Me.chkAssy, "Show the assemblies that are currently open")
        Me.chkAssy.UseVisualStyleBackColor = True
        '
        'chkParts
        '
        Me.chkParts.AutoSize = True
        Me.chkParts.Location = New System.Drawing.Point(6, 29)
        Me.chkParts.Name = "chkParts"
        Me.chkParts.Size = New System.Drawing.Size(50, 17)
        Me.chkParts.TabIndex = 35
        Me.chkParts.Text = "Parts"
        Me.ToolTip1.SetToolTip(Me.chkParts, "Show the parts that are currently open")
        Me.chkParts.UseVisualStyleBackColor = True
        '
        'chkDrawings
        '
        Me.chkDrawings.AutoSize = True
        Me.chkDrawings.Location = New System.Drawing.Point(6, 13)
        Me.chkDrawings.Name = "chkDrawings"
        Me.chkDrawings.Size = New System.Drawing.Size(70, 17)
        Me.chkDrawings.TabIndex = 34
        Me.chkDrawings.Text = "Drawings"
        Me.ToolTip1.SetToolTip(Me.chkDrawings, "Show the drawings that are currently open")
        Me.chkDrawings.UseVisualStyleBackColor = True
        '
        'GroupBox1
        '
        Me.GroupBox1.Controls.Add(Me.chkAssy)
        Me.GroupBox1.Controls.Add(Me.chkDerived)
        Me.GroupBox1.Controls.Add(Me.chkDrawings)
        Me.GroupBox1.Controls.Add(Me.chkParts)
        Me.GroupBox1.Controls.Add(Me.chkPres)
        Me.GroupBox1.Location = New System.Drawing.Point(12, 27)
        Me.GroupBox1.Name = "GroupBox1"
        Me.GroupBox1.Size = New System.Drawing.Size(142, 79)
        Me.GroupBox1.TabIndex = 50
        Me.GroupBox1.TabStop = False
        Me.GroupBox1.Text = "Selection"
        '
        'GroupBox2
        '
        Me.GroupBox2.Controls.Add(Me.PictureBox1)
        Me.GroupBox2.Controls.Add(Me.chkPartSelect)
        Me.GroupBox2.Controls.Add(Me.lstOpenfiles)
        Me.GroupBox2.Location = New System.Drawing.Point(160, 27)
        Me.GroupBox2.Name = "GroupBox2"
        Me.GroupBox2.Size = New System.Drawing.Size(191, 214)
        Me.GroupBox2.TabIndex = 51
        Me.GroupBox2.TabStop = False
        Me.GroupBox2.Text = "           Open Parts"
        '
        'PictureBox1
        '
        Me.PictureBox1.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Center
        Me.PictureBox1.Image = Global.My.Resources.Resources.inverse1
        Me.PictureBox1.InitialImage = Nothing
        Me.PictureBox1.Location = New System.Drawing.Point(25, -1)
        Me.PictureBox1.Name = "PictureBox1"
        Me.PictureBox1.Size = New System.Drawing.Size(15, 15)
        Me.PictureBox1.TabIndex = 57
        Me.PictureBox1.TabStop = False
        '
        'chkPartSelect
        '
        Me.chkPartSelect.AutoSize = True
        Me.chkPartSelect.Location = New System.Drawing.Point(10, 0)
        Me.chkPartSelect.Name = "chkPartSelect"
        Me.chkPartSelect.Size = New System.Drawing.Size(15, 14)
        Me.chkPartSelect.TabIndex = 1
        Me.chkPartSelect.UseVisualStyleBackColor = True
        '
        'lstOpenfiles
        '
        Me.lstOpenfiles.BackColor = System.Drawing.Color.White
        Me.lstOpenfiles.CheckOnClick = True
        Me.lstOpenfiles.FormattingEnabled = True
        Me.lstOpenfiles.IntegralHeight = False
        Me.lstOpenfiles.Location = New System.Drawing.Point(7, 18)
        Me.lstOpenfiles.Name = "lstOpenfiles"
        Me.lstOpenfiles.Size = New System.Drawing.Size(180, 191)
        Me.lstOpenfiles.TabIndex = 0
        Me.lstOpenfiles.ThreeDCheckBoxes = True
        Me.ToolTip1.SetToolTip(Me.lstOpenfiles, "Documents that are currently open")
        '
        'GroupBox4
        '
        Me.GroupBox4.Controls.Add(Me.chkClose)
        Me.GroupBox4.Controls.Add(Me.chkUseDrawings)
        Me.GroupBox4.Controls.Add(Me.chkSkipAssy)
        Me.GroupBox4.Controls.Add(Me.chkExport)
        Me.GroupBox4.Controls.Add(Me.chkDXF)
        Me.GroupBox4.Controls.Add(Me.chkRRev)
        Me.GroupBox4.Controls.Add(Me.ChkRevType)
        Me.GroupBox4.Controls.Add(Me.chkDWG)
        Me.GroupBox4.Controls.Add(Me.chkPDF)
        Me.GroupBox4.Controls.Add(Me.Label3)
        Me.GroupBox4.Controls.Add(Me.chkPrint)
        Me.GroupBox4.Controls.Add(Me.chkCheck)
        Me.GroupBox4.Controls.Add(Me.chkOpen)
        Me.GroupBox4.Controls.Add(Me.chkiProp)
        Me.GroupBox4.Location = New System.Drawing.Point(554, 27)
        Me.GroupBox4.Name = "GroupBox4"
        Me.GroupBox4.Size = New System.Drawing.Size(176, 144)
        Me.GroupBox4.TabIndex = 53
        Me.GroupBox4.TabStop = False
        Me.GroupBox4.Text = "Drawing Action"
        '
        'chkUseDrawings
        '
        Me.chkUseDrawings.AutoSize = True
        Me.chkUseDrawings.Location = New System.Drawing.Point(25, 151)
        Me.chkUseDrawings.Name = "chkUseDrawings"
        Me.chkUseDrawings.Size = New System.Drawing.Size(92, 17)
        Me.chkUseDrawings.TabIndex = 62
        Me.chkUseDrawings.Text = "Use Drawings"
        Me.ToolTip1.SetToolTip(Me.chkUseDrawings, "Create DXF's from the drawing.")
        Me.chkUseDrawings.UseVisualStyleBackColor = True
        Me.chkUseDrawings.Visible = False
        '
        'chkSkipAssy
        '
        Me.chkSkipAssy.AutoSize = True
        Me.chkSkipAssy.Location = New System.Drawing.Point(25, 136)
        Me.chkSkipAssy.Name = "chkSkipAssy"
        Me.chkSkipAssy.Size = New System.Drawing.Size(102, 17)
        Me.chkSkipAssy.TabIndex = 36
        Me.chkSkipAssy.Text = "Skip Assemblies"
        Me.ToolTip1.SetToolTip(Me.chkSkipAssy, "Do not make a DXF of the assemby")
        Me.chkSkipAssy.UseVisualStyleBackColor = True
        Me.chkSkipAssy.Visible = False
        '
        'chkRRev
        '
        Me.chkRRev.AutoSize = True
        Me.chkRRev.Location = New System.Drawing.Point(10, 90)
        Me.chkRRev.Name = "chkRRev"
        Me.chkRRev.Size = New System.Drawing.Size(110, 17)
        Me.chkRRev.TabIndex = 60
        Me.chkRRev.Text = "Remove Revision"
        Me.chkRRev.UseVisualStyleBackColor = True
        '
        'ChkRevType
        '
        Me.ChkRevType.AutoSize = True
        Me.ChkRevType.Location = New System.Drawing.Point(10, 75)
        Me.ChkRevType.Name = "ChkRevType"
        Me.ChkRevType.Size = New System.Drawing.Size(113, 17)
        Me.ChkRevType.TabIndex = 59
        Me.ChkRevType.Text = "Change Rev Type"
        Me.ChkRevType.UseVisualStyleBackColor = True
        '
        'chkDWG
        '
        Me.chkDWG.AutoSize = True
        Me.chkDWG.Location = New System.Drawing.Point(62, 121)
        Me.chkDWG.Name = "chkDWG"
        Me.chkDWG.Size = New System.Drawing.Size(53, 17)
        Me.chkDWG.TabIndex = 58
        Me.chkDWG.Text = "DWG"
        Me.chkDWG.UseVisualStyleBackColor = True
        Me.chkDWG.Visible = False
        '
        'chkPDF
        '
        Me.chkPDF.AutoSize = True
        Me.chkPDF.Location = New System.Drawing.Point(17, 121)
        Me.chkPDF.Name = "chkPDF"
        Me.chkPDF.Size = New System.Drawing.Size(47, 17)
        Me.chkPDF.TabIndex = 57
        Me.chkPDF.Text = "PDF"
        Me.chkPDF.UseVisualStyleBackColor = True
        Me.chkPDF.Visible = False
        '
        'btnRename
        '
        Me.btnRename.Location = New System.Drawing.Point(6, 65)
        Me.btnRename.Name = "btnRename"
        Me.btnRename.Size = New System.Drawing.Size(130, 23)
        Me.btnRename.TabIndex = 2
        Me.btnRename.Text = "Rename Parts"
        Me.ToolTip1.SetToolTip(Me.btnRename, "Rename the components of a selected assembly")
        Me.btnRename.UseVisualStyleBackColor = True
        '
        'btnFlatPattern
        '
        Me.btnFlatPattern.Location = New System.Drawing.Point(6, 90)
        Me.btnFlatPattern.Name = "btnFlatPattern"
        Me.btnFlatPattern.Size = New System.Drawing.Size(130, 23)
        Me.btnFlatPattern.TabIndex = 1
        Me.btnFlatPattern.Text = "Export Flat Pattern"
        Me.ToolTip1.SetToolTip(Me.btnFlatPattern, "Add the parent drawing name to the Reference table in each drawing")
        Me.btnFlatPattern.UseVisualStyleBackColor = True
        '
        'btnSpreadsheet
        '
        Me.btnSpreadsheet.Location = New System.Drawing.Point(6, 15)
        Me.btnSpreadsheet.Name = "btnSpreadsheet"
        Me.btnSpreadsheet.Size = New System.Drawing.Size(130, 23)
        Me.btnSpreadsheet.TabIndex = 0
        Me.btnSpreadsheet.Text = "Create Spreadsheet"
        Me.ToolTip1.SetToolTip(Me.btnSpreadsheet, "Create a spreadsheet of all the components of a selected assembly")
        Me.btnSpreadsheet.UseVisualStyleBackColor = True
        '
        'GroupBox3
        '
        Me.GroupBox3.Controls.Add(Me.LVSubFiles)
        Me.GroupBox3.Controls.Add(Me.txtSearch)
        Me.GroupBox3.Controls.Add(Me.chkDWGSelect)
        Me.GroupBox3.Cursor = System.Windows.Forms.Cursors.Default
        Me.GroupBox3.Location = New System.Drawing.Point(357, 27)
        Me.GroupBox3.Name = "GroupBox3"
        Me.GroupBox3.Size = New System.Drawing.Size(191, 214)
        Me.GroupBox3.TabIndex = 52
        Me.GroupBox3.TabStop = False
        Me.GroupBox3.Text = "          Referenced Drawings"
        '
        'LVSubFiles
        '
        Me.LVSubFiles.Activation = System.Windows.Forms.ItemActivation.OneClick
        Me.LVSubFiles.AutoArrange = False
        Me.LVSubFiles.CheckBoxes = True
        Me.LVSubFiles.Columns.AddRange(New System.Windows.Forms.ColumnHeader() {Me.ColumnHeader1})
        Me.LVSubFiles.HeaderStyle = System.Windows.Forms.ColumnHeaderStyle.None
        Me.LVSubFiles.Location = New System.Drawing.Point(6, 18)
        Me.LVSubFiles.Name = "LVSubFiles"
        Me.LVSubFiles.Size = New System.Drawing.Size(178, 190)
        Me.LVSubFiles.TabIndex = 62
        Me.LVSubFiles.UseCompatibleStateImageBehavior = False
        Me.LVSubFiles.View = System.Windows.Forms.View.Details
        '
        'ColumnHeader1
        '
        Me.ColumnHeader1.Width = 160
        '
        'txtSearch
        '
        Me.txtSearch.Location = New System.Drawing.Point(6, 188)
        Me.txtSearch.Name = "txtSearch"
        Me.txtSearch.Size = New System.Drawing.Size(178, 20)
        Me.txtSearch.TabIndex = 1
        Me.txtSearch.Visible = False
        '
        'chkDWGSelect
        '
        Me.chkDWGSelect.AutoSize = True
        Me.chkDWGSelect.Checked = True
        Me.chkDWGSelect.CheckState = System.Windows.Forms.CheckState.Checked
        Me.chkDWGSelect.Location = New System.Drawing.Point(6, 0)
        Me.chkDWGSelect.Name = "chkDWGSelect"
        Me.chkDWGSelect.Size = New System.Drawing.Size(15, 14)
        Me.chkDWGSelect.TabIndex = 58
        Me.chkDWGSelect.UseVisualStyleBackColor = True
        '
        'tmr
        '
        '
        'ToolTip1
        '
        '
        'btnRef
        '
        Me.btnRef.Location = New System.Drawing.Point(6, 40)
        Me.btnRef.Name = "btnRef"
        Me.btnRef.Size = New System.Drawing.Size(130, 23)
        Me.btnRef.TabIndex = 3
        Me.btnRef.Text = "Update Ref Table"
        Me.ToolTip1.SetToolTip(Me.btnRef, "Add the parent drawing name to the Reference table in each drawing")
        Me.btnRef.UseVisualStyleBackColor = True
        '
        'PictureBox2
        '
        Me.PictureBox2.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Center
        Me.PictureBox2.Image = Global.My.Resources.Resources.inverse1
        Me.PictureBox2.InitialImage = Nothing
        Me.PictureBox2.Location = New System.Drawing.Point(378, 27)
        Me.PictureBox2.Name = "PictureBox2"
        Me.PictureBox2.Size = New System.Drawing.Size(15, 15)
        Me.PictureBox2.TabIndex = 59
        Me.PictureBox2.TabStop = False
        '
        'GroupBox5
        '
        Me.GroupBox5.Controls.Add(Me.btnRef)
        Me.GroupBox5.Controls.Add(Me.btnRename)
        Me.GroupBox5.Controls.Add(Me.btnSpreadsheet)
        Me.GroupBox5.Controls.Add(Me.btnFlatPattern)
        Me.GroupBox5.Cursor = System.Windows.Forms.Cursors.Default
        Me.GroupBox5.Location = New System.Drawing.Point(12, 123)
        Me.GroupBox5.Name = "GroupBox5"
        Me.GroupBox5.Size = New System.Drawing.Size(142, 118)
        Me.GroupBox5.TabIndex = 60
        Me.GroupBox5.TabStop = False
        Me.GroupBox5.Text = "Part Actions"
        '
        'CMSSubFiles
        '
        Me.CMSSubFiles.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.SortToolStripMenuItem, Me.ShowHideToolStripMenuItem, Me.ExportToToolStripMenuItem})
        Me.CMSSubFiles.Name = "ContextMenuStrip1"
        Me.CMSSubFiles.Size = New System.Drawing.Size(134, 70)
        '
        'SortToolStripMenuItem
        '
        Me.SortToolStripMenuItem.DropDownItems.AddRange(New System.Windows.Forms.ToolStripItem() {Me.CMSAlphabetical, Me.CMSHeirarchical})
        Me.SortToolStripMenuItem.Name = "SortToolStripMenuItem"
        Me.SortToolStripMenuItem.Size = New System.Drawing.Size(133, 22)
        Me.SortToolStripMenuItem.Text = "Sort"
        '
        'CMSAlphabetical
        '
        Me.CMSAlphabetical.Name = "CMSAlphabetical"
        Me.CMSAlphabetical.Size = New System.Drawing.Size(140, 22)
        Me.CMSAlphabetical.Text = "Alphabetical"
        '
        'CMSHeirarchical
        '
        Me.CMSHeirarchical.Checked = True
        Me.CMSHeirarchical.CheckState = System.Windows.Forms.CheckState.Checked
        Me.CMSHeirarchical.Name = "CMSHeirarchical"
        Me.CMSHeirarchical.Size = New System.Drawing.Size(140, 22)
        Me.CMSHeirarchical.Text = "Hierarchical"
        '
        'ShowHideToolStripMenuItem
        '
        Me.ShowHideToolStripMenuItem.DropDownItems.AddRange(New System.Windows.Forms.ToolStripItem() {Me.CMSShow, Me.CMSHide})
        Me.ShowHideToolStripMenuItem.Name = "ShowHideToolStripMenuItem"
        Me.ShowHideToolStripMenuItem.Size = New System.Drawing.Size(133, 22)
        Me.ShowHideToolStripMenuItem.Text = "Show/Hide"
        '
        'CMSShow
        '
        Me.CMSShow.Checked = True
        Me.CMSShow.CheckState = System.Windows.Forms.CheckState.Checked
        Me.CMSShow.Name = "CMSShow"
        Me.CMSShow.Size = New System.Drawing.Size(199, 22)
        Me.CMSShow.Text = "Show Missing Drawings"
        '
        'CMSHide
        '
        Me.CMSHide.Name = "CMSHide"
        Me.CMSHide.Size = New System.Drawing.Size(199, 22)
        Me.CMSHide.Text = "Hide Missing Drawings"
        '
        'ExportToToolStripMenuItem
        '
        Me.ExportToToolStripMenuItem.DropDownItems.AddRange(New System.Windows.Forms.ToolStripItem() {Me.CMSSubSpreadsheet, Me.CMSSubText})
        Me.ExportToToolStripMenuItem.Name = "ExportToToolStripMenuItem"
        Me.ExportToToolStripMenuItem.Size = New System.Drawing.Size(133, 22)
        Me.ExportToToolStripMenuItem.Text = "Export To"
        '
        'CMSSubSpreadsheet
        '
        Me.CMSSubSpreadsheet.Name = "CMSSubSpreadsheet"
        Me.CMSSubSpreadsheet.Size = New System.Drawing.Size(138, 22)
        Me.CMSSubSpreadsheet.Text = "Spreadsheet"
        '
        'CMSSubText
        '
        Me.CMSSubText.Name = "CMSSubText"
        Me.CMSSubText.Size = New System.Drawing.Size(138, 22)
        Me.CMSSubText.Text = "Text File"
        '
        'CMSOpenFiles
        '
        Me.CMSOpenFiles.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.ExportToToolStripMenuItem1})
        Me.CMSOpenFiles.Name = "CMSOpenFiles"
        Me.CMSOpenFiles.Size = New System.Drawing.Size(122, 26)
        '
        'ExportToToolStripMenuItem1
        '
        Me.ExportToToolStripMenuItem1.DropDownItems.AddRange(New System.Windows.Forms.ToolStripItem() {Me.CMSSpreadsheet, Me.CMSTextFile})
        Me.ExportToToolStripMenuItem1.Name = "ExportToToolStripMenuItem1"
        Me.ExportToToolStripMenuItem1.Size = New System.Drawing.Size(121, 22)
        Me.ExportToToolStripMenuItem1.Text = "Export to"
        '
        'CMSSpreadsheet
        '
        Me.CMSSpreadsheet.Name = "CMSSpreadsheet"
        Me.CMSSpreadsheet.Size = New System.Drawing.Size(138, 22)
        Me.CMSSpreadsheet.Text = "Spreadsheet"
        '
        'CMSTextFile
        '
        Me.CMSTextFile.Name = "CMSTextFile"
        Me.CMSTextFile.Size = New System.Drawing.Size(138, 22)
        Me.CMSTextFile.Text = "Text File"
        '
        'MenuStrip1
        '
        Me.MenuStrip1.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.FileToolStripMenuItem, Me.AboutToolStripMenuItem})
        Me.MenuStrip1.Location = New System.Drawing.Point(0, 0)
        Me.MenuStrip1.Name = "MenuStrip1"
        Me.MenuStrip1.Size = New System.Drawing.Size(742, 24)
        Me.MenuStrip1.TabIndex = 61
        Me.MenuStrip1.Text = "MenuStrip1"
        '
        'FileToolStripMenuItem
        '
        Me.FileToolStripMenuItem.DropDownItems.AddRange(New System.Windows.Forms.ToolStripItem() {Me.DefaultSettingsToolStripMenuItem, Me.IPropertySettingsToolStripMenuItem})
        Me.FileToolStripMenuItem.Name = "FileToolStripMenuItem"
        Me.FileToolStripMenuItem.Size = New System.Drawing.Size(37, 20)
        Me.FileToolStripMenuItem.Text = "File"
        '
        'DefaultSettingsToolStripMenuItem
        '
        Me.DefaultSettingsToolStripMenuItem.Name = "DefaultSettingsToolStripMenuItem"
        Me.DefaultSettingsToolStripMenuItem.Size = New System.Drawing.Size(167, 22)
        Me.DefaultSettingsToolStripMenuItem.Text = "Export Settings"
        '
        'IPropertySettingsToolStripMenuItem
        '
        Me.IPropertySettingsToolStripMenuItem.Name = "IPropertySettingsToolStripMenuItem"
        Me.IPropertySettingsToolStripMenuItem.Size = New System.Drawing.Size(167, 22)
        Me.IPropertySettingsToolStripMenuItem.Text = "iProperty Settings"
        '
        'AboutToolStripMenuItem
        '
        Me.AboutToolStripMenuItem.DropDownItems.AddRange(New System.Windows.Forms.ToolStripItem() {Me.AboutBatchProgramToolStripMenuItem, Me.HowToToolStripMenuItem})
        Me.AboutToolStripMenuItem.Name = "AboutToolStripMenuItem"
        Me.AboutToolStripMenuItem.Size = New System.Drawing.Size(52, 20)
        Me.AboutToolStripMenuItem.Text = "About"
        '
        'AboutBatchProgramToolStripMenuItem
        '
        Me.AboutBatchProgramToolStripMenuItem.Name = "AboutBatchProgramToolStripMenuItem"
        Me.AboutBatchProgramToolStripMenuItem.Size = New System.Drawing.Size(189, 22)
        Me.AboutBatchProgramToolStripMenuItem.Text = "About Batch Program"
        '
        'HowToToolStripMenuItem
        '
        Me.HowToToolStripMenuItem.Name = "HowToToolStripMenuItem"
        Me.HowToToolStripMenuItem.Size = New System.Drawing.Size(189, 22)
        Me.HowToToolStripMenuItem.Text = "How to..."
        '
        'ContextMenuStrip1
        '
        Me.ContextMenuStrip1.Name = "ContextMenuStrip1"
        Me.ContextMenuStrip1.Size = New System.Drawing.Size(61, 4)
        '
        'MsVistaProgressBar1
        '
        Me.MsVistaProgressBar1.BackColor = System.Drawing.Color.Transparent
        Me.MsVistaProgressBar1.BlockSize = 10
        Me.MsVistaProgressBar1.BlockSpacing = 10
        Me.MsVistaProgressBar1.Cursor = System.Windows.Forms.Cursors.AppStarting
        Me.MsVistaProgressBar1.DisplayText = "%P%"
        Me.MsVistaProgressBar1.DisplayTextColor = System.Drawing.SystemColors.ControlText
        Me.MsVistaProgressBar1.DisplayTextFont = New System.Drawing.Font("Arial", 8.0!)
        Me.MsVistaProgressBar1.GradiantStyle = MSVistaProgressBar.BackGradiant.None
        Me.MsVistaProgressBar1.Location = New System.Drawing.Point(12, 243)
        Me.MsVistaProgressBar1.Name = "MsVistaProgressBar1"
        Me.MsVistaProgressBar1.ShowText = True
        Me.MsVistaProgressBar1.Size = New System.Drawing.Size(536, 20)
        Me.MsVistaProgressBar1.TabIndex = 56
        Me.MsVistaProgressBar1.Visible = False
        '
        'Main
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.BackColor = System.Drawing.SystemColors.Control
        Me.ClientSize = New System.Drawing.Size(742, 268)
        Me.Controls.Add(Me.MenuStrip1)
        Me.Controls.Add(Me.PictureBox2)
        Me.Controls.Add(Me.GroupBox4)
        Me.Controls.Add(Me.MsVistaProgressBar1)
        Me.Controls.Add(Me.btnExit)
        Me.Controls.Add(Me.btnOK)
        Me.Controls.Add(Me.GroupBox3)
        Me.Controls.Add(Me.GroupBox2)
        Me.Controls.Add(Me.GroupBox1)
        Me.Controls.Add(Me.GroupBox5)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.Fixed3D
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.MainMenuStrip = Me.MenuStrip1
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.Name = "Main"
        Me.Text = "Batch Program"
        Me.TransparencyKey = System.Drawing.Color.Maroon
        Me.GroupBox1.ResumeLayout(False)
        Me.GroupBox1.PerformLayout()
        Me.GroupBox2.ResumeLayout(False)
        Me.GroupBox2.PerformLayout()
        CType(Me.PictureBox1, System.ComponentModel.ISupportInitialize).EndInit()
        Me.GroupBox4.ResumeLayout(False)
        Me.GroupBox4.PerformLayout()
        Me.GroupBox3.ResumeLayout(False)
        Me.GroupBox3.PerformLayout()
        CType(Me.PictureBox2, System.ComponentModel.ISupportInitialize).EndInit()
        Me.GroupBox5.ResumeLayout(False)
        Me.CMSSubFiles.ResumeLayout(False)
        Me.CMSOpenFiles.ResumeLayout(False)
        Me.MenuStrip1.ResumeLayout(False)
        Me.MenuStrip1.PerformLayout()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents chkPres As System.Windows.Forms.CheckBox
    Friend WithEvents chkDerived As System.Windows.Forms.CheckBox
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents btnExit As System.Windows.Forms.Button
    Friend WithEvents btnOK As System.Windows.Forms.Button
    Friend WithEvents chkClose As System.Windows.Forms.CheckBox
    Friend WithEvents chkExport As System.Windows.Forms.CheckBox
    Friend WithEvents chkDXF As System.Windows.Forms.CheckBox
    Friend WithEvents chkPrint As System.Windows.Forms.CheckBox
    Friend WithEvents chkCheck As System.Windows.Forms.CheckBox
    Friend WithEvents chkOpen As System.Windows.Forms.CheckBox
    Friend WithEvents chkiProp As System.Windows.Forms.CheckBox
    Friend WithEvents chkAssy As System.Windows.Forms.CheckBox
    Friend WithEvents chkParts As System.Windows.Forms.CheckBox
    Friend WithEvents chkDrawings As System.Windows.Forms.CheckBox
    Friend WithEvents GroupBox1 As System.Windows.Forms.GroupBox
    Friend WithEvents GroupBox2 As System.Windows.Forms.GroupBox
    Friend WithEvents lstOpenfiles As System.Windows.Forms.CheckedListBox
    Friend WithEvents GroupBox4 As System.Windows.Forms.GroupBox
    Friend WithEvents chkSkipAssy As System.Windows.Forms.CheckBox
    Friend WithEvents btnRename As System.Windows.Forms.Button
    Friend WithEvents btnFlatPattern As System.Windows.Forms.Button
    Friend WithEvents btnSpreadsheet As System.Windows.Forms.Button
    Friend WithEvents GroupBox3 As System.Windows.Forms.GroupBox
    Friend WithEvents tmr As System.Windows.Forms.Timer
    Friend WithEvents ToolTip1 As System.Windows.Forms.ToolTip
    Friend WithEvents MsVistaProgressBar1 As MSVistaProgressBar
    Friend WithEvents txtSearch As System.Windows.Forms.TextBox
    Friend WithEvents chkDWG As Windows.Forms.CheckBox
    Friend WithEvents chkPDF As Windows.Forms.CheckBox
    Friend WithEvents PictureBox1 As Windows.Forms.PictureBox
    Friend WithEvents chkPartSelect As Windows.Forms.CheckBox
    Friend WithEvents PictureBox2 As Windows.Forms.PictureBox
    Friend WithEvents chkDWGSelect As Windows.Forms.CheckBox
    Friend WithEvents chkRRev As Windows.Forms.CheckBox
    Friend WithEvents ChkRevType As Windows.Forms.CheckBox
    Friend WithEvents GroupBox5 As Windows.Forms.GroupBox
    Friend WithEvents chkUseDrawings As Windows.Forms.CheckBox
    Friend WithEvents btnRef As Windows.Forms.Button
    Friend WithEvents CMSSubFiles As Windows.Forms.ContextMenuStrip
    Friend WithEvents SortAlphabeticallyToolStripMenuItem As Windows.Forms.ToolStripMenuItem
    Friend WithEvents SortHierarchicallyToolStripMenuItem As Windows.Forms.ToolStripMenuItem
    Friend WithEvents AlphabeticalToolStripMenuItem As Windows.Forms.ToolStripMenuItem
    Friend WithEvents HierarchicalToolStripMenuItem As Windows.Forms.ToolStripMenuItem
    Friend WithEvents SpreadsheetToolStripMenuItem As Windows.Forms.ToolStripMenuItem
    Friend WithEvents TextFileToolStripMenuItem As Windows.Forms.ToolStripMenuItem
    Friend WithEvents SortToolStripMenuItem As Windows.Forms.ToolStripMenuItem
    Friend WithEvents AlphabeticalToolStripMenuItem1 As Windows.Forms.ToolStripMenuItem
    Friend WithEvents HeirarchicalToolStripMenuItem As Windows.Forms.ToolStripMenuItem
    Friend WithEvents SpreadsheetToolStripMenuItem1 As Windows.Forms.ToolStripMenuItem
    Friend WithEvents TextFileToolStripMenuItem1 As Windows.Forms.ToolStripMenuItem
    Friend WithEvents CMSAlphabetical As Windows.Forms.ToolStripMenuItem
    Friend WithEvents CMSHeirarchical As Windows.Forms.ToolStripMenuItem
    Friend WithEvents CMSOpenFiles As Windows.Forms.ContextMenuStrip
    Friend WithEvents ExportToToolStripMenuItem1 As Windows.Forms.ToolStripMenuItem
    Friend WithEvents CMSSpreadsheet As Windows.Forms.ToolStripMenuItem
    Friend WithEvents CMSTextFile As Windows.Forms.ToolStripMenuItem
    Friend WithEvents ShowHideToolStripMenuItem As Windows.Forms.ToolStripMenuItem
    Friend WithEvents CMSShow As Windows.Forms.ToolStripMenuItem
    Friend WithEvents CMSHide As Windows.Forms.ToolStripMenuItem
    Friend WithEvents MenuStrip1 As Windows.Forms.MenuStrip
    Friend WithEvents FileToolStripMenuItem As Windows.Forms.ToolStripMenuItem
    Friend WithEvents DefaultSettingsToolStripMenuItem As Windows.Forms.ToolStripMenuItem
    Friend WithEvents AboutToolStripMenuItem As Windows.Forms.ToolStripMenuItem
    Friend WithEvents AboutBatchProgramToolStripMenuItem As Windows.Forms.ToolStripMenuItem
    Friend WithEvents ContextMenuStrip1 As Windows.Forms.ContextMenuStrip
    Friend WithEvents HowToToolStripMenuItem As Windows.Forms.ToolStripMenuItem
    Friend WithEvents LVSubFiles As Windows.Forms.ListView
    Friend WithEvents ColumnHeader1 As Windows.Forms.ColumnHeader
    Friend WithEvents ExportToToolStripMenuItem As Windows.Forms.ToolStripMenuItem
    Friend WithEvents CMSSubSpreadsheet As Windows.Forms.ToolStripMenuItem
    Friend WithEvents CMSSubText As Windows.Forms.ToolStripMenuItem
    Friend WithEvents IPropertySettingsToolStripMenuItem As Windows.Forms.ToolStripMenuItem
End Class
