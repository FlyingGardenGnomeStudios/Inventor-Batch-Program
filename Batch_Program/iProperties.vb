Imports System.Windows.Forms
Imports Inventor
Imports System.Runtime.InteropServices
Imports System.Globalization
Imports System.Drawing

Public Class iProperties
#Region "Form Setup"
    Dim _invApp As Inventor.Application
    Dim SumDicRows As New Dictionary(Of String, DataGridViewRow)
    Dim ProjDicRows As New Dictionary(Of String, DataGridViewRow)
    Dim StatusDicRows As New Dictionary(Of String, DataGridViewRow)
    Dim RevDicRows As New Dictionary(Of String, DataGridViewRow)
    Dim CustomModDicRows As New Dictionary(Of String, DataGridViewRow)
    Dim CustomDrawDicRows As New Dictionary(Of String, DataGridViewRow)
    Dim SumDic As New Dictionary(Of String, Items)
    Dim ProjDic As New Dictionary(Of String, Items)
    Dim StatusDic As New Dictionary(Of String, Items)
    Dim CustomModDic As New Dictionary(Of String, Items)
    Dim CustomDrawDic As New Dictionary(Of String, Items)
    Dim RevDic As New Dictionary(Of String, Items)
    Dim oDateTimePicker As New DateTimePicker
    Dim cmbYesNo As DataGridViewComboBoxCell
    Dim InvStringDic As New Dictionary(Of String, String)
    Dim InvDateDic As New Dictionary(Of String, Date)
    Dim materiallist As New List(Of String)
    Dim dpControl As String = ""
    Dim dgvRevControl As DataGridView
    Dim dgvcmbcell As DataGridViewComboBoxCell
    Dim InvTitle, InvSubject, InvManager, InvCategory, InvKeywords, InvAuthor, InvCompany, Invcomments, InvDescription, InvPartNumber, InvStockNumber _
                 , InvProject, InvDesigner, InvVendor, InvModDate, InvDrawDate, InvDrawBy As Inventor.Property
    Dim InvDesignState As String
    Dim CustomPropSet As PropertySet
    Dim InvEngineer, InvStatus, InvRevision, InvDesignerState, InvCheckedBy, InvCheckDate
    Dim InvRef As New Dictionary(Of String, String)
    Dim InvRefProp(0 To 9) As Inventor.Property
    Dim MaterialCell As DataGridViewComboBoxCell = New DataGridViewComboBoxCell
#End Region
    Dim Main As Main
#Region "Setup iProperties"
    Private Sub dgvProject_CellEnter(sender As Object, e As DataGridViewCellEventArgs) Handles dgvProject.CellEnter
        Dim DateRow As Boolean = False

        If e.RowIndex = -1 Or e.ColumnIndex <= 0 Then Exit Sub
        If InStr(dgvProject(dgvProject.Columns("ProjItem").Index, (e.RowIndex)).Value, "Date") <> 0 Then
            DateRow = True
            dpControl = "Project"
            DateTimePicker(e, DateRow, dgvProject)
        Else
            oDateTimePicker.Visible = False
        End If
    End Sub
    Private Sub dgvCustomDrawing_CellEnter(sender As Object, e As DataGridViewCellEventArgs) Handles dgvCustomDrawing.CellEnter
        Dim DateRow As Boolean = False
        If e.RowIndex = -1 Or e.ColumnIndex <= 0 Then Exit Sub
        If dgvCustomDrawing(dgvCustomDrawing.Columns("DCusType").Index, e.RowIndex).Value = "Date" AndAlso
            dgvCustomDrawing.Columns("DCusValue").Index = e.ColumnIndex Then
            DateRow = True
            dpControl = "CustomDrawing"
            DateTimePicker(e, DateRow, dgvCustomDrawing)
        Else
            oDateTimePicker.Visible = False
        End If
    End Sub
    Private Sub dgvCustomModel_CellEnter(sender As Object, e As DataGridViewCellEventArgs) Handles dgvCustomModel.CellEnter
        Dim DateRow As Boolean = False
        Dim YesNoRow As Boolean = False
        If e.RowIndex = -1 Or e.ColumnIndex <= 0 Then Exit Sub
        If dgvCustomModel(dgvCustomModel.Columns("PCusType").Index, e.RowIndex).Value = "Date" AndAlso
           dgvCustomModel.Columns("PCusValue").Index = e.ColumnIndex Then
            DateRow = True
            dpControl = "CustomModel"
            DateTimePicker(e, DateRow, dgvCustomModel)
            'ElseIf dgvCustomModel(dgvCustomModel.Columns("PCusType").Index, e.RowIndex).Value = "Yes or No" AndAlso
            '   dgvCustomModel.Columns("PCusValue").Index = e.ColumnIndex Then
            '    YesNoRow = True
            '    cmbYesNo(e, YesNoRow, dgvCustomModel)
            'Else
            oDateTimePicker.Visible = False
        End If
    End Sub
    Private Sub DateTimePicker(ByVal e As DataGridViewCellEventArgs, ByRef DateRow As Boolean, dgvcontrol As DataGridView)
        If DateRow = True Then
            'Adding DateTimePicker control into DataGridView 
            dgvcontrol.Controls.Add(oDateTimePicker)
            ' Setting the format (i.e. 2014-10-10)
            oDateTimePicker.Format = DateTimePickerFormat.Short
            oDateTimePicker.ShowCheckBox = True
            If dgvcontrol.CurrentCell.Value = "" Then
                oDateTimePicker.Checked = False
            Else
                oDateTimePicker.Checked = True
            End If
            ' It returns the retangular area that represents the Display area for a cell
            Dim oRectangle As Rectangle = dgvcontrol.GetCellDisplayRectangle(e.ColumnIndex, e.RowIndex, True)
            'Setting area for DateTimePicker Control
            oDateTimePicker.Size = New Size(oRectangle.Width, oRectangle.Height)
            ' Setting Location
            oDateTimePicker.Location = New Drawing.Point(oRectangle.X, oRectangle.Y)
            ' An event attached to dateTimePicker Control which is fired when DateTimeControl is closed
            AddHandler oDateTimePicker.CloseUp, AddressOf iPropDateTimePicker_CloseUp
            ' An event attached to dateTimePicker Control which is fired when any date is selected
            AddHandler oDateTimePicker.TextChanged, AddressOf iPropdateTimePicker_OnTextChange
            ' Now make it visible
            oDateTimePicker.Visible = True
        Else
            oDateTimePicker.Visible = False
        End If
    End Sub
    Private Sub dgvStatus_CellEnter(sender As Object, e As DataGridViewCellEventArgs) Handles dgvStatus.CellEnter
        Dim DateRow As Boolean = False
        If e.RowIndex = -1 Or e.ColumnIndex <= 0 Then Exit Sub
        If InStr(dgvStatus(dgvStatus.Columns("StatusItem").Index, (e.RowIndex)).Value, "Date") <> 0 Then
            DateRow = True
            dpControl = "Status"
            DateTimePicker(e, DateRow, dgvStatus)
        Else
            oDateTimePicker.Visible = False
        End If
    End Sub
    Private Sub dgvRev1_CellEnter(sender As Object, e As DataGridViewCellEventArgs) Handles dgvRev1.CellEnter
        Dim DateRow As Boolean = False
        If e.RowIndex = -1 Or e.ColumnIndex <= 0 Then Exit Sub
        If InStr(dgvRev1(dgvRev1.Columns("Rev1Item").Index, (e.RowIndex)).Value, "Date") <> 0 Then
            DateRow = True
            dgvRevControl = dgvRev1
            DateTimePicker(e, DateRow, dgvRev1)
        Else
            oDateTimePicker.Visible = False
        End If
    End Sub
    Private Sub RevDatePicker(Sender As Object, e As DataGridViewCellEventArgs)
        Dim DateRow As Boolean = False
        If e.RowIndex = -1 Or e.ColumnIndex <= 0 Then Exit Sub
        For Each control In RevisionTabs.SelectedTab.Controls
            If Strings.InStr(control.name, "Rev") <> 0 Then
                dgvRevControl = control
                Exit For
            End If
        Next
        If InStr(dgvRevControl(dgvRevControl.Columns("Rev" & RevisionTabs.SelectedTab.TabIndex + 1 & "Item").Index, (e.RowIndex)).Value, "Date") <> 0 Then
            DateRow = True
            DateTimePicker(e, DateRow, dgvRevControl)
        Else
            oDateTimePicker.Visible = False
        End If
    End Sub
    Private Sub iPropdateTimePicker_OnTextChange(ByVal sender As Object, ByVal e As EventArgs)
        ' Saving the 'Selected Date on Calendar' into DataGridView current cell
        If dpControl = "Status" Then
            If oDateTimePicker.Checked = True Then
                dgvStatus.CurrentCell.Value = oDateTimePicker.Value.ToString(System.Globalization.CultureInfo.CurrentCulture.DateTimeFormat.ShortDatePattern())
            Else
                dgvStatus.CurrentCell.Value = ""
            End If
        ElseIf dpControl = "Project" Then
            If oDateTimePicker.Checked = True Then
                dgvProject.CurrentCell.Value = oDateTimePicker.Value.ToString(System.Globalization.CultureInfo.CurrentCulture.DateTimeFormat.ShortDatePattern())
            Else
                dgvProject.CurrentCell.Value = ""
            End If
        ElseIf dpControl = "CustomDrawing" Then
            If oDateTimePicker.Checked = True Then
                dgvCustomDrawing.CurrentCell.Value = oDateTimePicker.Value.ToString(System.Globalization.CultureInfo.CurrentCulture.DateTimeFormat.ShortDatePattern())
            Else
                dgvCustomDrawing.CurrentCell.Value = ""
            End If
        ElseIf dpControl = "CustomModel" Then
            If oDateTimePicker.Checked = True Then
                dgvCustomModel.CurrentCell.Value = oDateTimePicker.Value.ToString(System.Globalization.CultureInfo.CurrentCulture.DateTimeFormat.ShortDatePattern())
            Else
                dgvCustomModel.CurrentCell.Value = ""
            End If
        Else
            If oDateTimePicker.Checked = True Then
                dgvRevControl.CurrentCell.Value = oDateTimePicker.Value.ToString(System.Globalization.CultureInfo.CurrentCulture.DateTimeFormat.ShortDatePattern())
            Else
                dgvRevControl.CurrentCell.Value = ""
            End If
        End If
        'dpControl = ""
    End Sub
    Private Sub iPropDateTimePicker_CloseUp(ByVal sender As Object, ByVal e As EventArgs)
        ' Hiding the control after use 
        oDateTimePicker.Visible = False
    End Sub
    Dim RevTable As New RevTable
    Public Sub New()

        ' This call is required by the designer.
        InitializeComponent()
        _invApp = Marshal.GetActiveObject("Inventor.Application")
        ' Add any initialization after the InitializeComponent() call.
        dgvSummary.Rows.Add("Title:")
        dgvSummary.Rows.Add("Subject:")
        dgvSummary.Rows.Add("Author:")
        dgvSummary.Rows.Add("Manager:")
        dgvSummary.Rows.Add("Company:")
        dgvSummary.Rows.Add("Category:")
        dgvSummary.Rows.Add("Keywords:")
        dgvSummary.Rows.Add("Comments:")
        dgvSummary.Rows.Add("Material:")
        dgvSummary(dgvSummary.Columns("SumModel").Index, dgvSummary.RowCount - 1) = New DataGridViewComboBoxCell
        dgvcmbcell = dgvSummary(dgvSummary.Columns("SumModel").Index, dgvSummary.RowCount - 1)
        SumDicRows.Add("Title:", dgvSummary.Rows(0))
        SumDicRows.Add("Subject:", dgvSummary.Rows(1))
        SumDicRows.Add("Author:", dgvSummary.Rows(2))
        SumDicRows.Add("Manager:", dgvSummary.Rows(3))
        SumDicRows.Add("Company:", dgvSummary.Rows(4))
        SumDicRows.Add("Category:", dgvSummary.Rows(5))
        SumDicRows.Add("Keywords:", dgvSummary.Rows(6))
        SumDicRows.Add("Comments:", dgvSummary.Rows(7))
        SumDicRows.Add("Material:", dgvSummary.Rows(8))
        Dim cmbwidth As Integer = Nothing
        For x = 1 To _invApp.ActiveMaterialLibrary.MaterialAssets.Count - 1
            materiallist.Add(_invApp.ActiveMaterialLibrary.MaterialAssets.Item(x).DisplayName)
        Next

        materiallist.Sort()
        dgvcmbcell.Items.AddRange(materiallist.ToArray)

        dgvProject.Rows.Add("Location:")
        dgvProject.Rows.Item(dgvProject.RowCount - 1).ReadOnly = True
        dgvProject.Rows.Item(dgvProject.RowCount - 1).DefaultCellStyle.BackColor = Drawing.Color.LightGray
        dgvProject.Rows.Add("File Subtype:")
        dgvProject.Rows.Item(dgvProject.RowCount - 1).ReadOnly = True
        dgvProject.Rows.Item(dgvProject.RowCount - 1).DefaultCellStyle.BackColor = Drawing.Color.LightGray
        dgvProject.Rows.Add("Part Number:")
        dgvProject.Rows.Add("Stock Number:")
        dgvProject.Rows.Add("Description:")
        dgvProject.Rows.Add("Revision Number:")
        dgvProject.Rows.Add("Project:")
        dgvProject.Rows.Add("Designer:")
        dgvProject.Rows.Add("Engineer:")
        dgvProject.Rows.Add("Authority:")
        dgvProject.Rows.Add("Cost Center:")
        dgvProject.Rows.Add("Estimated Cost:")
        dgvProject.Rows.Add("Creation Date:")
        dgvProject.Rows.Add("Vendor:")
        dgvProject.Rows.Add("WEB Link:")
        ProjDicRows.Add("Location:", dgvProject.Rows(0))
        ProjDicRows.Add("File Subtype:", dgvProject.Rows(1))
        ProjDicRows.Add("Part Number:", dgvProject.Rows(2))
        ProjDicRows.Add("Stock Number:", dgvProject.Rows(3))
        ProjDicRows.Add("Description:", dgvProject.Rows(4))
        ProjDicRows.Add("Revision Number:", dgvProject.Rows(5))
        ProjDicRows.Add("Project:", dgvProject.Rows(6))
        ProjDicRows.Add("Designer:", dgvProject.Rows(7))
        ProjDicRows.Add("Engineer:", dgvProject.Rows(8))
        ProjDicRows.Add("Authority:", dgvProject.Rows(9))
        ProjDicRows.Add("Cost Center:", dgvProject.Rows(10))
        ProjDicRows.Add("Estimated Cost:", dgvProject.Rows(11))
        ProjDicRows.Add("Creation Date:", dgvProject.Rows(12))

        ProjDicRows.Add("Vendor:", dgvProject.Rows(13))
        ProjDicRows.Add("WEB Link:", dgvProject.Rows(14))

        dgvStatus.Rows.Add("Part Number:")
        dgvStatus.Rows.Item(dgvStatus.RowCount - 1).ReadOnly = True
        dgvStatus.Rows.Item(dgvStatus.RowCount - 1).DefaultCellStyle.BackColor = Drawing.Color.LightGray
        dgvStatus.Rows.Add("Stock Number:")
        dgvStatus.Rows.Add("Status:")
        dgvStatus.Rows.Add("Design State:")
        dgvStatus(dgvStatus.Columns("StatusDrawing").Index, dgvStatus.RowCount - 1) = New DataGridViewComboBoxCell
        dgvcmbcell = dgvStatus(dgvStatus.Columns("StatusDrawing").Index, dgvStatus.RowCount - 1)
        dgvcmbcell.Items.AddRange("Work In Progress", "Pending", "Released")
        dgvStatus(dgvStatus.Columns("StatusModel").Index, dgvStatus.RowCount - 1) = New DataGridViewComboBoxCell
        dgvcmbcell = dgvStatus(dgvStatus.Columns("StatusModel").Index, dgvStatus.RowCount - 1)
        dgvcmbcell.Items.AddRange("Work In Progress", "Pending", "Released")
        dgvStatus.Rows.Add("Checked By:")
        dgvStatus.Rows.Add("Checked Date:")
        dgvStatus.Rows.Add("Eng. Approved By:")
        dgvStatus.Rows.Add("Eng. Approved Date:")
        dgvStatus.Rows.Add("Mfg. Approved By:")
        dgvStatus.Rows.Add("Mfg. Approved Date:")
        StatusDicRows.Add("Part Number:", dgvStatus.Rows(0))
        StatusDicRows.Add("Stock Number:", dgvStatus.Rows(1))
        StatusDicRows.Add("Status:", dgvStatus.Rows(2))
        StatusDicRows.Add("Design State:", dgvStatus.Rows(3))
        StatusDicRows.Add("Checked By:", dgvStatus.Rows(4))
        StatusDicRows.Add("Checked Date:", dgvStatus.Rows(5))
        StatusDicRows.Add("Eng. Approved By:", dgvStatus.Rows(6))
        StatusDicRows.Add("Eng. Approved Date:", dgvStatus.Rows(7))
        StatusDicRows.Add("Mfg. Approved By:", dgvStatus.Rows(8))
        StatusDicRows.Add("Mfg. Approved Date:", dgvStatus.Rows(9))

        If My.Settings.RTSCheckedBy = True Then
            dgvRev1.Rows.Add("Checked By")
        End If
        If My.Settings.RTSCheckedDate = True Then
            dgvRev1.Rows.Add("Check Date")
        End If
        If My.Settings.RTSDate = True Then
            dgvRev1.Rows.Add(My.Settings.RTSDateCol)
        End If
        If My.Settings.RTSDesc = True Then
            dgvRev1.Rows.Add(My.Settings.RTSDescCol)
        End If
        If My.Settings.RTSName = True Then
            dgvRev1.Rows.Add(My.Settings.RTSNameCol)
        End If

        If My.Settings.RTSApproved = True Then
            dgvRev1.Rows.Add(My.Settings.RTSApprovedCol)
        End If

        If My.Settings.RTS1 = True Then
            dgvRev1.Rows.Add(My.Settings.RTS1Item)
        End If
        If My.Settings.RTS2 = True Then
            dgvRev1.Rows.Add(My.Settings.RTS2Item)
        End If
        If My.Settings.RTS3 = True Then
            dgvRev1.Rows.Add(My.Settings.RTS3Item)
        End If
        If My.Settings.RTS4 = True Then
            dgvRev1.Rows.Add(My.Settings.RTS4Item)
        End If
        If My.Settings.RTS5 = True Then
            dgvRev1.Rows.Add(My.Settings.RTS5Item)
        End If
    End Sub
#Region "Adjust Combobox"
    Private Function dgvSummary_EditingContrtolShowing(ByVal sender As System.Object, ByVal e As System.Windows.Forms.DataGridViewEditingControlShowingEventArgs) Handles dgvSummary.EditingControlShowing
        Try
            If dgvSummary.CurrentCellAddress = New Drawing.Point(1, 8) Then
                Dim combo As ComboBox = CType(e.Control, ComboBox)
                AdjustComboBoxWidth(combo)
            End If
        Catch
        End Try
    End Function
    Private Function dgvStatus_EditingContrtolShowing(ByVal sender As System.Object, ByVal e As System.Windows.Forms.DataGridViewEditingControlShowingEventArgs) Handles dgvStatus.EditingControlShowing
        Try
            If dgvStatus.CurrentCell.ColumnIndex <> 0 AndAlso dgvStatus.CurrentCell.RowIndex = 3 Then
                Dim combo As ComboBox = CType(e.Control, ComboBox)
                AdjustComboBoxWidth(combo)
            End If
        Catch
        End Try
    End Function
    Private Function dgvCustomModel_EditingContrtolShowing(ByVal sender As System.Object, ByVal e As System.Windows.Forms.DataGridViewEditingControlShowingEventArgs) Handles dgvCustomModel.EditingControlShowing
        Try
            If dgvStatus.CurrentCell.ColumnIndex = 2 Then
                Dim combo As ComboBox = CType(e.Control, ComboBox)
                AdjustComboBoxWidth(combo)
            End If
        Catch
        End Try
    End Function
    Private Function dgvCustomDrawing_EditingContrtolShowing(ByVal sender As System.Object, ByVal e As System.Windows.Forms.DataGridViewEditingControlShowingEventArgs) Handles dgvCustomDrawing.EditingControlShowing
        Try
            If dgvStatus.CurrentCell.ColumnIndex = 2 Then
                Dim combo As ComboBox = CType(e.Control, ComboBox)
                AdjustComboBoxWidth(combo)
            End If
        Catch
        End Try
    End Function
    Public Shared Function AdjustComboBoxWidth(ByVal sender As Object)
        Dim senderComboBox = DirectCast(sender, ComboBox)
        Dim width As Integer = senderComboBox.DropDownWidth
        Dim g As Graphics = senderComboBox.CreateGraphics()
        Dim font As Font = senderComboBox.Font

        Dim vertScrollBarWidth As Integer = If((senderComboBox.Items.Count > senderComboBox.MaxDropDownItems), SystemInformation.VerticalScrollBarWidth, 0)

        Dim newWidth As Integer
        For Each s As String In DirectCast(sender, ComboBox).Items
            newWidth = CInt(g.MeasureString(s, font).Width) + vertScrollBarWidth
            If width < newWidth Then
                width = newWidth
            End If
        Next

        senderComboBox.DropDownWidth = width
        Return False
    End Function
#End Region
    Public Function PopMain(CalledFunction As Main)
        Main = CalledFunction
        Return Nothing
    End Function
    Public Function PopRevTable(CalledFunction As RevTable)
        RevTable = CalledFunction
        Return Nothing
    End Function
    Public Sub PopulateiProps(Path As Documents, ByRef oDoc As Document, ByRef Archive As String _
                             , ByRef DrawingName As String, ByRef DrawSource As String, OpenDocs As ArrayList _
                             , ByRef Read As Boolean)
        Dim Y, Total As Integer
        For Y = 0 To Main.dgvSubFiles.RowCount - 1
            If Main.dgvSubFiles(Main.dgvSubFiles.Columns("chkSubFiles").Index, Y).Value = True Then
                Total += 1
            End If
        Next
        'Go through iProperties of each selected item and retrieve the values
        iPropIdentify(Path, oDoc, Archive, DrawingName, DrawSource, OpenDocs, 0, Total, Read)
    End Sub
    Private Sub btnAddRev_Click(sender As Object, e As EventArgs) Handles btnAddRev.Click
        AddRev()
    End Sub
    Private Sub AddRev()
        'Rev1
        Dim NewRev As New TabPage
        Dim newDGV As New DataGridView
        newDGV.Parent = NewRev
        newDGV.Location = New Drawing.Point(6, 19)
        Dim NewItemColumn As New DataGridViewTextBoxColumn With {
            .Name = "Rev" & RevisionTabs.TabCount + 1 & "Item",
            .HeaderText = "Item"}
        Dim NewDrawingColumn As New DataGridViewTextBoxColumn With {
            .Name = "Rev" & RevisionTabs.TabCount + 1 & "Drawing",
            .HeaderText = "Value"}
        Dim NewDirtyColumn As New DataGridViewCheckBoxColumn With {
            .Name = "Rev" & RevisionTabs.TabCount + 1 & "IsDirty",
            .HeaderText = "Is Dirty",
            .Visible = False}

        RevisionTabs.TabPages.Add(NewRev)
        Dim RevLabel As New Label With {
            .AutoSize = True,
            .Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte)),
        .Location = New System.Drawing.Point(3, 0),
        .Name = "lblRev" & RevisionTabs.TabCount,
        .Size = New System.Drawing.Size(307, 13),
        .TabIndex = 0,
        .Text = "Changes here overwrite the revision table, proceed with caution" & Global.Microsoft.VisualBasic.ChrW(13) & Global.Microsoft.VisualBasic.ChrW(10)}

        NewRev.Controls.Add(newDGV)
        NewRev.Controls.Add(RevLabel)
        NewRev.Location = New System.Drawing.Point(4, 22)
        NewRev.Name = "Rev" & RevisionTabs.TabCount
        NewRev.Padding = New System.Windows.Forms.Padding(3)
        NewRev.Size = New System.Drawing.Size(339, 366)
        NewRev.TabIndex = RevisionTabs.TabCount - 1
        NewRev.Text = "Rev" & RevisionTabs.TabCount
        NewRev.UseVisualStyleBackColor = True
        btnAddRev.Parent = NewRev
        NewRev.Name = "Rev" & RevisionTabs.TabCount
        'dgvRev1
        '
        newDGV.AllowUserToAddRows = False
        newDGV.AllowUserToDeleteRows = False
        newDGV.AutoSizeColumnsMode = System.Windows.Forms.DataGridViewAutoSizeColumnsMode.Fill
        newDGV.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        newDGV.Columns.AddRange(New System.Windows.Forms.DataGridViewColumn() {NewItemColumn, NewDrawingColumn, NewDirtyColumn})
        newDGV.EditMode = System.Windows.Forms.DataGridViewEditMode.EditOnEnter
        newDGV.Location = New System.Drawing.Point(6, 19)
        newDGV.Name = "dgvRev" & RevisionTabs.TabCount
        newDGV.RowHeadersVisible = False
        newDGV.Size = New System.Drawing.Size(326, 312)
        newDGV.TabIndex = 17
        If My.Settings.RTSCheckedBy = True Then
            newDGV.Rows.Add("Checked By")
        End If
        If My.Settings.RTSCheckedDate = True Then
            newDGV.Rows.Add("Check Date")
        End If
        If My.Settings.RTSDate = True Then
            newDGV.Rows.Add(My.Settings.RTSDateCol)
        End If
        If My.Settings.RTSDesc = True Then
            newDGV.Rows.Add(My.Settings.RTSDescCol)
        End If
        If My.Settings.RTSName = True Then
            newDGV.Rows.Add(My.Settings.RTSNameCol)
        End If

        If My.Settings.RTSApproved = True Then
            newDGV.Rows.Add(My.Settings.RTSApprovedCol)
        End If

        If My.Settings.RTS1 = True Then
            newDGV.Rows.Add(My.Settings.RTS1Item)
        End If
        If My.Settings.RTS2 = True Then
            newDGV.Rows.Add(My.Settings.RTS2Item)
        End If
        If My.Settings.RTS3 = True Then
            newDGV.Rows.Add(My.Settings.RTS3Item)
        End If
        If My.Settings.RTS4 = True Then
            newDGV.Rows.Add(My.Settings.RTS4Item)
        End If
        If My.Settings.RTS5 = True Then
            newDGV.Rows.Add(My.Settings.RTS5Item)
        End If

        AddHandler newDGV.CellValueChanged, AddressOf Me.RevCellChange
        AddHandler newDGV.CellEnter, AddressOf Me.RevDatePicker
        newDGV.Name = "dgvRev" & RevisionTabs.TabCount
    End Sub
    Public Sub iPropIdentify(Path As Documents, ByRef oDoc As Document, ByRef Archive As String _
                             , ByRef DrawingName As String, ByRef DrawSource As String, OpenDocs As ArrayList _
                             , ByRef RevTotal As Integer, ByRef Total As Integer, ByRef Read As Boolean)
        Dim X, Y As Integer
        X = 0
        Dim Write As Boolean = False
        Dim List As String = ""
        Dim dDoc As Document = Nothing
        'Dim DrawSource As String = Nothing
        Dim Modelsource As String = Nothing
        RevTable.PopiProperties(Me)
        Main.PopiProperties(Me)
        RevTable.dgvRevTable.Rows.Clear()
        _invApp.SilentOperation = True
        If Read = False Then Call RevTable.iPropRevTable(oDoc)
        'Go through drawings to see which ones are selected
        Dim Warning As Boolean = False
        For Y = 0 To Main.dgvSubFiles.RowCount - 1

            'Look through all sub files in open documents to get the part sourcefile
            If Main.dgvSubFiles(Main.dgvSubFiles.Columns("chkSubFiles").Index, Y).Value = True Then
                Modelsource = Main.dgvSubFiles(Main.dgvSubFiles.Columns("DrawingSource").Index, Y).Value
                DrawSource = Main.dgvSubFiles(Main.dgvSubFiles.Columns("DrawingLocation").Index, Y).Value
                If Modelsource <> "" Then oDoc = _invApp.Documents.Open(Modelsource, False)
                If DrawSource <> "" Then dDoc = _invApp.Documents.Open(DrawSource, False)
                If Read = True Then
                    InvDateDic.Clear()
                    InvStringDic.Clear()
                    InvRef.Clear()
                    If Modelsource <> "" AndAlso DrawSource <> "" Then
                        Call SetBothProps(oDoc, dDoc)
                        'Call ModifyRevTable(dDoc, True)
                    Else
                        Call SetModelProps(oDoc)
                    End If
                    Call InitializeiProperties()
                Else
                    Call WriteProps(oDoc, "Model")
                    oDoc.Update()
                    Call WriteProps(dDoc, "Drawing")
                    ' Call ModifyRevTable(dDoc, False)
                    dDoc.Update()
                End If
                Main.CloseLater(DrawingName, oDoc)
                Main.MatchDrawing(DrawSource, DrawingName, Y)
                ''oDoc = _invApp.Documents.Open(DrawSource, False)
                'If Read = True Then
                '    'Call SetDrawingProps(oDoc, Warning)
                '    'Call RevTableCheck(oDoc, RevTotal)

                '    '  Main.ProgressBar(Main.LVSubFiles.CheckedItems.Count, Y + 1, "Reading: ", DrawingName)
                'Else
                '    ProgressBar2.Visible = True
                '    ' ProgressBar2.Value = ((X + 1) / Main.LVSubFiles.CheckedItems.Count) * 100
                '    ProgressBar2.PerformStep()
                '    Call WriteDrawingProps(oDoc)

                '    _invApp.ActiveDocument.Update()
                'End If
                Try
                    Main.CloseLater(oDoc.DisplayName, oDoc)
                Catch
                    Main.writeDebug("Couldn't close " & DrawingName & ". Could already be closed")
                End Try
                X += 1
                ProgressBar2.Value = (X / Total) * 100
                ' Main.bgwRun.ReportProgress((X / Total) * 100, "Getting iProperties from: " & oDoc.DisplayName)
                ' End If
            End If
            _invApp.SilentOperation = False
        Next
        'If Read = True Then
        '    For X = RevTotal To RevisionTabs.TabPages.Count - 2
        '        RevisionTabs.TabPages.Remove(RevisionTabs.TabPages(RevisionTabs.TabPages.Count - 1))
        '    Next
        '    Try
        '        With RevisionTabs.TabPages.Item(RevisionTabs.TabPages.Count - 1)
        '            .Text = "Add Rev"
        '            .Controls.Item("lblRev" & RevisionTabs.TabPages.Count.ToString).Visible = False
        '            .Controls.Item("chkRev" & RevisionTabs.TabPages.Count.ToString & "AddRev").Visible = True
        '            .Controls.Item("txtRev" & RevisionTabs.TabPages.Count.ToString & "Rev").Text = "*Varies*"
        '            Call ChangeAddRev(False)
        '        End With
        '    Catch
        '        'MsgBox("There was a problem retrieving the revision history")
        '        'TabControl1.Enabled = False
        '    End Try

        'End If
        ProgressBar2.Visible = False
        RevTable.Close()
        If List <> "" Then
            MsgBox("The following drawings are missing revision tables:" & vbNewLine & List & vbNewLine _
                   & "You may add the tables when the operation completes.", , "Missing Revision Table")
        End If
    End Sub
    Private Sub SetBothProps(invModelDoc As Document, invDrawDoc As Document)
        'Set the identifiers for the model properties
        SumDic.Clear()
        SumDic.Add("Title:", New Items With {.ModelValue = invModelDoc.PropertySets.Item("{F29F85E0-4FF9-1068-AB91-08002B27B3D9}").ItemByPropId("2").Value,
                   .DrawingValue = invDrawDoc.PropertySets.Item("{F29F85E0-4FF9-1068-AB91-08002B27B3D9}").ItemByPropId("2").Value})
        SumDic.Add("Subject:", New Items With {.ModelValue = invModelDoc.PropertySets.Item("{F29F85E0-4FF9-1068-AB91-08002B27B3D9}").ItemByPropId("3").Value,
                   .DrawingValue = invDrawDoc.PropertySets.Item("{F29F85E0-4FF9-1068-AB91-08002B27B3D9}").ItemByPropId("3").Value})
        SumDic.Add("Manager:", New Items With {.ModelValue = invModelDoc.PropertySets.Item("{D5CDD502-2E9C-101B-9397-08002B2CF9AE}").ItemByPropId("14").Value,
                   .DrawingValue = invDrawDoc.PropertySets.Item("{D5CDD502-2E9C-101B-9397-08002B2CF9AE}").ItemByPropId("14").Value})
        SumDic.Add("Category:", New Items With {.ModelValue = invModelDoc.PropertySets.Item("{D5CDD502-2E9C-101B-9397-08002B2CF9AE}").ItemByPropId("2").Value,
                   .DrawingValue = invDrawDoc.PropertySets.Item("{D5CDD502-2E9C-101B-9397-08002B2CF9AE}").ItemByPropId("2").Value})
        SumDic.Add("Keywords:", New Items With {.ModelValue = invModelDoc.PropertySets.Item("{F29F85E0-4FF9-1068-AB91-08002B27B3D9}").ItemByPropId("5").Value,
                   .DrawingValue = invDrawDoc.PropertySets.Item("{F29F85E0-4FF9-1068-AB91-08002B27B3D9}").ItemByPropId("5").Value})
        SumDic.Add("Author:", New Items With {.ModelValue = invModelDoc.PropertySets.Item("{F29F85E0-4FF9-1068-AB91-08002B27B3D9}").ItemByPropId("4").Value,
                   .DrawingValue = invDrawDoc.PropertySets.Item("{F29F85E0-4FF9-1068-AB91-08002B27B3D9}").ItemByPropId("4").Value})
        SumDic.Add("Company:", New Items With {.ModelValue = invModelDoc.PropertySets.Item("{D5CDD502-2E9C-101B-9397-08002B2CF9AE}").ItemByPropId("15").Value,
                   .DrawingValue = invDrawDoc.PropertySets.Item("{D5CDD502-2E9C-101B-9397-08002B2CF9AE}").ItemByPropId("15").Value})
        SumDic.Add("Comments:", New Items With {.ModelValue = invModelDoc.PropertySets.Item("{F29F85E0-4FF9-1068-AB91-08002B27B3D9}").ItemByPropId("6").Value,
                   .DrawingValue = invDrawDoc.PropertySets.Item("{F29F85E0-4FF9-1068-AB91-08002B27B3D9}").ItemByPropId("6").Value})
        Try
            SumDic.Add("Material:", New Items With {.ModelValue = invModelDoc.ComponentDefinition.Material.Name,
                   .DrawingValue = Nothing})
        Catch
        End Try
        ProjDic.Clear()
        ProjDic.Add("Location:", New Items With {.ModelValue = invModelDoc.FullFileName,
                    .DrawingValue = invDrawDoc.FullFileName})
        ProjDic.Add("File Subtype:", New Items With {.ModelValue = invModelDoc.PropertySets.Item("{32853F0F-3444-11D1-9E93-0060B03C1CA6}").ItemByPropId("32").Value,
                    .DrawingValue = invDrawDoc.PropertySets.Item("{32853F0F-3444-11D1-9E93-0060B03C1CA6}").ItemByPropId("32").Value})
        ProjDic.Add("Part Number:", New Items With {.ModelValue = invModelDoc.PropertySets.Item("{32853F0F-3444-11D1-9E93-0060B03C1CA6}").ItemByPropId("5").Value,
                    .DrawingValue = invDrawDoc.PropertySets.Item("{32853F0F-3444-11D1-9E93-0060B03C1CA6}").ItemByPropId("5").Value})
        ProjDic.Add("Stock Number:", New Items With {.ModelValue = invModelDoc.PropertySets.Item("{32853F0F-3444-11D1-9E93-0060B03C1CA6}").ItemByPropId("55").Value,
                    .DrawingValue = invDrawDoc.PropertySets.Item("{32853F0F-3444-11D1-9E93-0060B03C1CA6}").ItemByPropId("55").Value})
        ProjDic.Add("Description:", New Items With {.ModelValue = invModelDoc.PropertySets.Item("{32853F0F-3444-11D1-9E93-0060B03C1CA6}").ItemByPropId("29").Value,
                    .DrawingValue = invDrawDoc.PropertySets.Item("{32853F0F-3444-11D1-9E93-0060B03C1CA6}").ItemByPropId("29").Value})
        ProjDic.Add("Revision Number:", New Items With {.ModelValue = invModelDoc.PropertySets.Item("{F29F85E0-4FF9-1068-AB91-08002B27B3D9}").ItemByPropId("9").Value,
                    .DrawingValue = invDrawDoc.PropertySets.Item("{F29F85E0-4FF9-1068-AB91-08002B27B3D9}").ItemByPropId("9").Value})
        ProjDic.Add("Project:", New Items With {.ModelValue = invModelDoc.PropertySets.Item("{32853F0F-3444-11D1-9E93-0060B03C1CA6}").ItemByPropId("7").Value,
                  .DrawingValue = invDrawDoc.PropertySets.Item("{32853F0F-3444-11D1-9E93-0060B03C1CA6}").ItemByPropId("7").Value})
        ProjDic.Add("Designer:", New Items With {.ModelValue = invModelDoc.PropertySets.Item("{32853F0F-3444-11D1-9E93-0060B03C1CA6}").ItemByPropId("41").Value,
                    .DrawingValue = invDrawDoc.PropertySets.Item("{32853F0F-3444-11D1-9E93-0060B03C1CA6}").ItemByPropId("41").Value})
        ProjDic.Add("Engineer:", New Items With {.ModelValue = invModelDoc.PropertySets.Item("{32853F0F-3444-11D1-9E93-0060B03C1CA6}").ItemByPropId("42").Value,
                    .DrawingValue = invDrawDoc.PropertySets.Item("{32853F0F-3444-11D1-9E93-0060B03C1CA6}").ItemByPropId("42").Value})
        ProjDic.Add("Authority:", New Items With {.ModelValue = invModelDoc.PropertySets.Item("{32853F0F-3444-11D1-9E93-0060B03C1CA6}").ItemByPropId("43").Value,
                    .DrawingValue = invDrawDoc.PropertySets.Item("{32853F0F-3444-11D1-9E93-0060B03C1CA6}").ItemByPropId("42").Value})
        ProjDic.Add("Cost Center:", New Items With {.ModelValue = invModelDoc.PropertySets.Item("{32853F0F-3444-11D1-9E93-0060B03C1CA6}").ItemByPropId("9").Value,
                    .DrawingValue = invDrawDoc.PropertySets.Item("{32853F0F-3444-11D1-9E93-0060B03C1CA6}").ItemByPropId("9").Value})
        ProjDic.Add("Estimated Cost:", New Items With {.ModelValue = invModelDoc.PropertySets.Item("{32853F0F-3444-11D1-9E93-0060B03C1CA6}").ItemByPropId("36").Value,
                    .DrawingValue = invDrawDoc.PropertySets.Item("{32853F0F-3444-11D1-9E93-0060B03C1CA6}").ItemByPropId("36").Value})
        ProjDic.Add("Creation Date:", New Items With {.ModelValue = invModelDoc.PropertySets.Item("{32853F0F-3444-11D1-9E93-0060B03C1CA6}").ItemByPropId("4").Value,
                    .DrawingValue = invDrawDoc.PropertySets.Item("{32853F0F-3444-11D1-9E93-0060B03C1CA6}").ItemByPropId("4").Value})
        ProjDic.Add("Vendor:", New Items With {.ModelValue = invModelDoc.PropertySets.Item("{32853F0F-3444-11D1-9E93-0060B03C1CA6}").ItemByPropId("30").Value,
                    .DrawingValue = invDrawDoc.PropertySets.Item("{32853F0F-3444-11D1-9E93-0060B03C1CA6}").ItemByPropId("30").Value})
        ProjDic.Add("WEB Link:", New Items With {.ModelValue = invModelDoc.PropertySets.Item("{32853F0F-3444-11D1-9E93-0060B03C1CA6}").ItemByPropId("23").Value,
                    .DrawingValue = invDrawDoc.PropertySets.Item("{32853F0F-3444-11D1-9E93-0060B03C1CA6}").ItemByPropId("23").Value})
        StatusDic.Clear()
        StatusDic.Add("Part Number:", New Items With {.ModelValue = invModelDoc.PropertySets.Item("{32853F0F-3444-11D1-9E93-0060B03C1CA6}").ItemByPropId("5").Value,
                    .DrawingValue = invDrawDoc.PropertySets.Item("{32853F0F-3444-11D1-9E93-0060B03C1CA6}").ItemByPropId("5").Value})
        StatusDic.Add("Stock Number:", New Items With {.ModelValue = invModelDoc.PropertySets.Item("{32853F0F-3444-11D1-9E93-0060B03C1CA6}").ItemByPropId("55").Value,
                    .DrawingValue = invDrawDoc.PropertySets.Item("{32853F0F-3444-11D1-9E93-0060B03C1CA6}").ItemByPropId("55").Value})
        StatusDic.Add("Status:", New Items With {.ModelValue = invModelDoc.PropertySets.Item("{32853F0F-3444-11D1-9E93-0060B03C1CA6}").ItemByPropId("17").Value,
                    .DrawingValue = invDrawDoc.PropertySets.Item("{32853F0F-3444-11D1-9E93-0060B03C1CA6}").ItemByPropId("17").Value})
        StatusDic.Add("Design State:", New Items With {.ModelValue = invModelDoc.PropertySets.Item("{32853F0F-3444-11D1-9E93-0060B03C1CA6}").ItemByPropId("40").Value,
                    .DrawingValue = invDrawDoc.PropertySets.Item("{32853F0F-3444-11D1-9E93-0060B03C1CA6}").ItemByPropId("40").Value})
        StatusDic.Add("Checked By:", New Items With {.ModelValue = invModelDoc.PropertySets.Item("{32853F0F-3444-11D1-9E93-0060B03C1CA6}").ItemByPropId("10").Value,
                    .DrawingValue = invDrawDoc.PropertySets.Item("{32853F0F-3444-11D1-9E93-0060B03C1CA6}").ItemByPropId("10").Value})
        StatusDic.Add("Checked Date:", New Items With {.ModelValue = invModelDoc.PropertySets.Item("{32853F0F-3444-11D1-9E93-0060B03C1CA6}").ItemByPropId("11").Value,
                    .DrawingValue = invDrawDoc.PropertySets.Item("{32853F0F-3444-11D1-9E93-0060B03C1CA6}").ItemByPropId("11").Value})
        StatusDic.Add("Eng. Approved By:", New Items With {.ModelValue = invModelDoc.PropertySets.Item("{32853F0F-3444-11D1-9E93-0060B03C1CA6}").ItemByPropId("12").Value,
                    .DrawingValue = invDrawDoc.PropertySets.Item("{32853F0F-3444-11D1-9E93-0060B03C1CA6}").ItemByPropId("12").Value})
        StatusDic.Add("Eng. Approved Date:", New Items With {.ModelValue = invModelDoc.PropertySets.Item("{32853F0F-3444-11D1-9E93-0060B03C1CA6}").ItemByPropId("13").Value,
                    .DrawingValue = invDrawDoc.PropertySets.Item("{32853F0F-3444-11D1-9E93-0060B03C1CA6}").ItemByPropId("13").Value})
        StatusDic.Add("Mfg. Approved By:", New Items With {.ModelValue = invModelDoc.PropertySets.Item("{32853F0F-3444-11D1-9E93-0060B03C1CA6}").ItemByPropId("34").Value,
                    .DrawingValue = invDrawDoc.PropertySets.Item("{32853F0F-3444-11D1-9E93-0060B03C1CA6}").ItemByPropId("34").Value})
        StatusDic.Add("Mfg. Approved Date:", New Items With {.ModelValue = invModelDoc.PropertySets.Item("{32853F0F-3444-11D1-9E93-0060B03C1CA6}").ItemByPropId("35").Value,
                    .DrawingValue = invDrawDoc.PropertySets.Item("{32853F0F-3444-11D1-9E93-0060B03C1CA6}").ItemByPropId("35").Value})
        CustomModDic.Clear()
        CustomDrawDic.Clear()

        For Each iProp As Inventor.Property In invModelDoc.PropertySets.Item("Inventor User Defined Properties")
            CustomModDic.Add(iProp.Name, New Items With {.ModelValue = invModelDoc.PropertySets.Item("Inventor User Defined Properties").Item(iProp.Name).Value,
                          .Type = iProp.Value.GetType})
        Next
        For Each iProp As Inventor.Property In invDrawDoc.PropertySets.Item("Inventor User Defined Properties")
            CustomDrawDic.Add(iProp.Name, New Items With {.DrawingValue = invDrawDoc.PropertySets.Item("Inventor User Defined Properties").Item(iProp.Name).Value,
                          .Type = iProp.Value.GetType.Name})
        Next
    End Sub
    Private Sub SetModelProps(invModelDoc As Document)
        'Set the identifiers for the model properties
        SumDic.Clear()
        SumDic.Add("Title:", New Items With {.ModelValue = invModelDoc.PropertySets.Item("{F29F85E0-4FF9-1068-AB91-08002B27B3D9}").ItemByPropId("2").Value})
        SumDic.Add("Subject:", New Items With {.ModelValue = invModelDoc.PropertySets.Item("{F29F85E0-4FF9-1068-AB91-08002B27B3D9}").ItemByPropId("3").Value})
        SumDic.Add("Manager:", New Items With {.ModelValue = invModelDoc.PropertySets.Item("{D5CDD502-2E9C-101B-9397-08002B2CF9AE}").ItemByPropId("14").Value})
        SumDic.Add("Category:", New Items With {.ModelValue = invModelDoc.PropertySets.Item("{D5CDD502-2E9C-101B-9397-08002B2CF9AE}").ItemByPropId("2").Value})
        SumDic.Add("Keywords:", New Items With {.ModelValue = invModelDoc.PropertySets.Item("{F29F85E0-4FF9-1068-AB91-08002B27B3D9}").ItemByPropId("5").Value})
        SumDic.Add("Author:", New Items With {.ModelValue = invModelDoc.PropertySets.Item("{F29F85E0-4FF9-1068-AB91-08002B27B3D9}").ItemByPropId("4").Value})
        SumDic.Add("Company:", New Items With {.ModelValue = invModelDoc.PropertySets.Item("{D5CDD502-2E9C-101B-9397-08002B2CF9AE}").ItemByPropId("15").Value})
        SumDic.Add("Comments:", New Items With {.ModelValue = invModelDoc.PropertySets.Item("{F29F85E0-4FF9-1068-AB91-08002B27B3D9}").ItemByPropId("6").Value})
        SumDic.Add("Material:", New Items With {.ModelValue = invModelDoc.PropertySets.Item("{32853F0F-3444-11D1-9E93-0060B03C1CA6}").ItemByPropId("20").Value})
        ProjDic.Clear()
        ProjDic.Add("Location:", New Items With {.ModelValue = invModelDoc.FullFileName})
        ProjDic.Add("File Subtype:", New Items With {.ModelValue = invModelDoc.PropertySets.Item("{32853F0F-3444-11D1-9E93-0060B03C1CA6}").ItemByPropId("32").Value})
        ProjDic.Add("Part Number:", New Items With {.ModelValue = invModelDoc.PropertySets.Item("{32853F0F-3444-11D1-9E93-0060B03C1CA6}").ItemByPropId("5").Value})
        ProjDic.Add("Stock Number:", New Items With {.ModelValue = invModelDoc.PropertySets.Item("{32853F0F-3444-11D1-9E93-0060B03C1CA6}").ItemByPropId("55").Value})
        ProjDic.Add("Description:", New Items With {.ModelValue = invModelDoc.PropertySets.Item("{32853F0F-3444-11D1-9E93-0060B03C1CA6}").ItemByPropId("29").Value})
        ProjDic.Add("Revision Number:", New Items With {.ModelValue = invModelDoc.PropertySets.Item("{F29F85E0-4FF9-1068-AB91-08002B27B3D9}").ItemByPropId("9").Value})
        ProjDic.Add("Project:", New Items With {.ModelValue = invModelDoc.PropertySets.Item("{32853F0F-3444-11D1-9E93-0060B03C1CA6}").ItemByPropId("7").Value})
        ProjDic.Add("Designer:", New Items With {.ModelValue = invModelDoc.PropertySets.Item("{32853F0F-3444-11D1-9E93-0060B03C1CA6}").ItemByPropId("41").Value})
        ProjDic.Add("Engineer:", New Items With {.ModelValue = invModelDoc.PropertySets.Item("{32853F0F-3444-11D1-9E93-0060B03C1CA6}").ItemByPropId("42").Value})
        ProjDic.Add("Authority:", New Items With {.ModelValue = invModelDoc.PropertySets.Item("{32853F0F-3444-11D1-9E93-0060B03C1CA6}").ItemByPropId("43").Value})
        ProjDic.Add("Cost Center:", New Items With {.ModelValue = invModelDoc.PropertySets.Item("{32853F0F-3444-11D1-9E93-0060B03C1CA6}").ItemByPropId("9").Value})
        ProjDic.Add("Estimated Cost:", New Items With {.ModelValue = invModelDoc.PropertySets.Item("{32853F0F-3444-11D1-9E93-0060B03C1CA6}").ItemByPropId("36").Value})
        ProjDic.Add("Creation Date:", New Items With {.ModelValue = invModelDoc.PropertySets.Item("{32853F0F-3444-11D1-9E93-0060B03C1CA6}").ItemByPropId("4").Value})
        ProjDic.Add("Vendor:", New Items With {.ModelValue = invModelDoc.PropertySets.Item("{32853F0F-3444-11D1-9E93-0060B03C1CA6}").ItemByPropId("30").Value})
        ProjDic.Add("WEB Link:", New Items With {.ModelValue = invModelDoc.PropertySets.Item("{32853F0F-3444-11D1-9E93-0060B03C1CA6}").ItemByPropId("23").Value})
        StatusDic.Clear()
        StatusDic.Add("Part Number:", New Items With {.ModelValue = invModelDoc.PropertySets.Item("{32853F0F-3444-11D1-9E93-0060B03C1CA6}").ItemByPropId("5").Value})
        StatusDic.Add("Stock Number:", New Items With {.ModelValue = invModelDoc.PropertySets.Item("{32853F0F-3444-11D1-9E93-0060B03C1CA6}").ItemByPropId("55").Value})
        StatusDic.Add("Status:", New Items With {.ModelValue = invModelDoc.PropertySets.Item("{32853F0F-3444-11D1-9E93-0060B03C1CA6}").ItemByPropId("17").Value})
        StatusDic.Add("Design State:", New Items With {.ModelValue = invModelDoc.PropertySets.Item("{32853F0F-3444-11D1-9E93-0060B03C1CA6}").ItemByPropId("40").Value})
        StatusDic.Add("Checked By:", New Items With {.ModelValue = invModelDoc.PropertySets.Item("{32853F0F-3444-11D1-9E93-0060B03C1CA6}").ItemByPropId("10").Value})
        StatusDic.Add("Checked Date:", New Items With {.ModelValue = invModelDoc.PropertySets.Item("{32853F0F-3444-11D1-9E93-0060B03C1CA6}").ItemByPropId("11").Value})
        StatusDic.Add("Eng. Approved By:", New Items With {.ModelValue = invModelDoc.PropertySets.Item("{32853F0F-3444-11D1-9E93-0060B03C1CA6}").ItemByPropId("12").Value})
        StatusDic.Add("Eng. Approved Date:", New Items With {.ModelValue = invModelDoc.PropertySets.Item("{32853F0F-3444-11D1-9E93-0060B03C1CA6}").ItemByPropId("13").Value})
        StatusDic.Add("Mfg. Approved By:", New Items With {.ModelValue = invModelDoc.PropertySets.Item("{32853F0F-3444-11D1-9E93-0060B03C1CA6}").ItemByPropId("34").Value})
        StatusDic.Add("Mfg. Approved Date:", New Items With {.ModelValue = invModelDoc.PropertySets.Item("{32853F0F-3444-11D1-9E93-0060B03C1CA6}").ItemByPropId("35").Value})
        CustomModDic.Clear()
        CustomDrawDic.Clear()

        For Each iProp As Inventor.Property In invModelDoc.PropertySets.Item("Inventor User Defined Properties")
            CustomModDic.Add(iProp.Name, New Items With {.ModelValue = invModelDoc.PropertySets.Item("Inventor User Defined Properties").Item(iProp.Name).Value,
                          .Type = iProp.Value.GetType})
        Next

    End Sub
    Private Sub SetDrawingProps(ByRef invDrawingDoc As Document, ByRef Warning As Boolean)


        'For A = 1 To CustomPropSet.Count
        '    If InStr(CustomPropSet.Item(A).Name, "Reference") = 0 Then
        '        CustomPropSet.Item(A).Delete()
        '    End If
        'Next
        ' For P = 0 To 9
        If invDrawingDoc.PropertySets.Item("Inventor User Defined Properties").Count = 0 Then Exit Sub
        For p = 0 To invDrawingDoc.PropertySets.Item("Inventor User Defined Properties").Count
            Try
                Debug.WriteLine(p.ToString)
                InvRef.Add("InvRef" & p.ToString, CustomPropSet.Item("Reference" & p).Value)
            Catch
                SetiProp(InvRefProp(p), "Reference" & p)
                InvRef.Add("InvRef" & p.ToString, CustomPropSet.Item("Reference" & p).Value)
                'If Warning = False Then
                '    MsgBox("Some drawings are missing reference keys, the references will be updated." & vbNewLine &
                '        "The data will be saved, but a new Titleblock is required to display the values." & vbNewLine &
                '        "To fix this update the titleblock through the Drawing Transfer Resource Wizard")
                '    Warning = True
                'End If
            End Try
        Next


    End Sub
    Private Sub SetiProp(InvRef As Inventor.Property, Ref As String)
        InvRef = CustomPropSet.Add("", Ref)
    End Sub
    Public Sub ModifyRevTable(ByRef oDoc As Document, ByRef read As Boolean)
        Dim CheckedBy As String
        Dim FirstRun As Boolean = True
        Dim RevCheckNameColNum, RevCheckDateColNum, RevCheckRevColNum, RevCheckDescColNum, RevCheckApproveColNum As Integer
        Dim RevCheck1ColNum, RevCheck2ColNum, RevCheck3ColNum, RevCheck4ColNum, RevCheck5ColNum As Integer
        Dim i, j As Integer
        Dim k As Integer = 0
        Dim DateChecked As Date
        CheckedBy = oDoc.PropertySets.Item("{32853F0F-3444-11D1-9E93-0060B03C1CA6}").ItemByPropId("10").Value
        DateChecked = CDate(oDoc.PropertySets.Item("{32853F0F-3444-11D1-9E93-0060B03C1CA6}").ItemByPropId("11").Value)
        Dim oPoint As Point2d
        Dim oRevTable As RevisionTable
        'RevTable.PopMain(Me)
        'Set the active sheet
        Dim Sheet As Sheet = oDoc.Activesheet
        Dim oBorder As Border = Sheet.Border
        If Not oBorder Is Nothing Then
            oPoint = _invApp.TransientGeometry.CreatePoint2d(0.965193999999989, 2.7178)
        Else
            oPoint = _invApp.TransientGeometry.CreatePoint2d(Sheet.Width, Sheet.Height)
        End If
        'Define the revision table from the active sheet
        oRevTable = Sheet.RevisionTables(1)
        If Err.Number = 5 Then
            oRevTable = oDoc.ActiveSheet.Revisiontables.Add2(oPoint, False, True, False, 0)
            'clear the error for future code.
            Err.Clear()
            Exit Sub
        End If
        'retrieve the data from the revision table
        Dim RevNo As String = oDoc.PropertySets.Item("{F29F85E0-4FF9-1068-AB91-08002B27B3D9}").ItemByPropId("9").Value
        Dim tg As TransientGeometry = _invApp.TransientGeometry
        Dim dd As DrawingDocument = _invApp.Documents.ItemByName(oDoc.FullDocumentName) '.ActiveDocument
        Dim s As Sheet = dd.ActiveSheet
        ' Get revision table
        Dim rt As RevisionTable = s.RevisionTables(1)
        ' Get dimensions
        ' Dim c As Integer = rt.RevisionTableColumns.Count
        Dim r As Integer = rt.RevisionTableRows.Count
        ' Counter
        ' Get headers and column widths
        Dim rtc As RevisionTableColumn
        i = 0
        If RevisionTabs.TabCount < oRevTable.RevisionTableRows.Count Then
            AddRev()
        End If
        For Each rtc In rt.RevisionTableColumns
            Dim h As New DataGridViewTextBoxColumn 'ColumnHeader
            h.Name = CultureInfo.CurrentCulture.TextInfo.ToTitleCase(LCase(rtc.Title))
            If FirstRun = True Then
                Select Case UCase(h.Name)
                    Case UCase(My.Settings.RTSRevCol)
                        If My.Settings.RTSRev = True Then
                            RevCheckRevColNum = i
                        End If
                    Case UCase(My.Settings.RTSDateCol)
                        If My.Settings.RTSDate = True Then
                            RevCheckDateColNum = i
                        End If
                    Case UCase(My.Settings.RTSDescCol)
                        If My.Settings.RTSDesc = True Then
                            RevCheckDescColNum = i
                        End If
                    Case UCase(My.Settings.RTSNameCol)
                        If My.Settings.RTSName = True Then
                            RevCheckNameColNum = i
                        End If
                    Case UCase(My.Settings.RTSApprovedCol)
                        If My.Settings.RTSApproved = True Then
                            RevCheckApproveColNum = i
                        End If
                    Case UCase(My.Settings.RTS1Col)
                        If My.Settings.RTS1 = True Then
                            RevCheck1ColNum = i
                        End If
                    Case UCase(My.Settings.RTS2Col)
                        If My.Settings.RTS2 = True Then
                            RevCheck2ColNum = i
                        End If
                    Case UCase(My.Settings.RTS3Col)
                        If My.Settings.RTS3 = True Then
                            RevCheck3ColNum = i
                        End If
                    Case UCase(My.Settings.RTS4Col)
                        If My.Settings.RTS4 = True Then
                            RevCheck4ColNum = i
                        End If
                    Case UCase(My.Settings.RTS5Col)
                        If My.Settings.RTS4 = True Then
                            RevCheck5ColNum = i
                        End If
                End Select
                i = i + 1
            End If

        Next
        FirstRun = False
        '' Get contents and row heights

        Dim c As Integer = rt.RevisionTableColumns.Count
        Dim Input(c) As String
        Dim heights(r) As Double
        Dim rtr As RevisionTableRow
        Dim DGVRows As Integer
        i = 0 : j = 1
        For Each rtr In rt.RevisionTableRows
            Dim rtcell As RevisionTableCell
            Dim contents(c * (r + 1)) As String

            Dim x As Integer = 0
            For Each rtcell In rtr
                contents(i) = rtcell.Text
                If k = 0 Then Input(x) = contents(i)
                i += 1
                x += 1
            Next
            For Each control In RevisionTabs.TabPages(j - 1).Controls
                If Strings.InStr(control.name, "Rev" & j) <> 0 Then
                    dgvRevControl = control
                    Exit For
                End If
            Next
            For z = 1 To dgvRevControl.RowCount - 1

                Select Case UCase(dgvRevControl(0, z).Value)

                    Case UCase("Checked By")
                        If read = True Then
                            If dgvRevControl(1, z).Value = Nothing OrElse dgvRevControl(1, z).Value = CheckedBy Then
                                dgvRevControl(1, z).Value = CheckedBy
                            Else
                                dgvRevControl(1, z).Value = "*Varies*"
                            End If
                        Else
                            If dgvRevControl(2, z).Value = True Then
                                oDoc.PropertySets.Item("{32853F0F-3444-11D1-9E93-0060B03C1CA6}").ItemByPropId("10").Value = dgvRevControl(1, z).Value
                            End If
                        End If
                    Case UCase("Check Date")
                        If read = True Then
                            If k = 0 Then
                                If DateChecked = #1/1/1601# Then
                                    dgvRevControl(1, z).Value = ""
                                Else
                                    If dgvRevControl(1, z).Value = Nothing OrElse dgvRevControl(1, z).Value = CStr(DateChecked.ToShortDateString) Then
                                        dgvRevControl(1, z).Value = CStr(DateChecked.ToShortDateString)
                                    Else
                                        dgvRevControl(1, z).Value = "*Varies*"
                                    End If
                                End If
                            End If
                        Else
                            If dgvRevControl(2, z).Value = True Then
                                oDoc.PropertySets.Item("{32853F0F-3444-11D1-9E93-0060B03C1CA6}").ItemByPropId("11").Value = CDate(dgvRevControl(1, z).Value)
                            End If
                        End If

                    Case UCase(My.Settings.RTSRevCol)
                        If read = True Then
                            If dgvRevControl(1, z).Value = Nothing OrElse dgvRevControl(1, z).Value = contents((k * rt.RevisionTableColumns.Count) + RevCheckRevColNum) Then
                                dgvRevControl(1, z).Value = contents((k * rt.RevisionTableColumns.Count) + RevCheckRevColNum)
                            Else
                                dgvRevControl(1, z).Value = "*Varies*"
                            End If
                        Else
                            If dgvRevControl(2, z).Value = True Then
                                rt.RevisionTableColumns.Item(RevCheckRevColNum + 1).value = dgvRevControl(1, z).Value
                            End If
                        End If
                    Case UCase(My.Settings.RTSDescCol)
                        If read = True Then
                            If dgvRevControl(1, z).Value = Nothing OrElse dgvRevControl(1, z).Value = contents((k * rt.RevisionTableColumns.Count) + RevCheckDescColNum) Then
                                dgvRevControl(1, z).Value = contents((k * rt.RevisionTableColumns.Count) + RevCheckDescColNum)
                            Else
                                dgvRevControl(1, z).Value = "*Varies*"
                            End If
                        Else
                            If dgvRevControl(2, z).Value = True Then
                                rt.RevisionTableRows(j).Item(RevCheckDescColNum + 1).Text = dgvRevControl(1, z).Value
                            End If
                        End If
                    Case UCase(My.Settings.RTSDateCol)
                        If read = True Then
                            If contents((k * rt.RevisionTableColumns.Count) + RevCheckDateColNum) = #1/1/1601# Then
                                dgvRevControl(1, z).Value = ""
                            Else
                                If dgvRevControl(1, z).Value = Nothing OrElse dgvRevControl(1, z).Value = CStr(DateTime.Parse(contents((k * rt.RevisionTableColumns.Count) + RevCheckDateColNum))) Then
                                    dgvRevControl(1, z).Value = CStr(DateTime.Parse(contents((k * rt.RevisionTableColumns.Count) + RevCheckDateColNum)))
                                Else
                                    dgvRevControl(1, z).Value = "*Varies*"
                                End If
                            End If
                        Else
                            If dgvRevControl(2, z).Value = True Then
                                rt.RevisionTableRows(j).Item(RevCheckDateColNum + 1).Text = dgvRevControl(1, z).Value
                            End If
                        End If
                    Case UCase(My.Settings.RTSNameCol)
                        If read = True Then
                            If dgvRevControl(1, z).Value = Nothing OrElse dgvRevControl(1, z).Value = contents((k * rt.RevisionTableColumns.Count) + RevCheckNameColNum) Then
                                dgvRevControl(1, z).Value = contents((k * rt.RevisionTableColumns.Count) + RevCheckNameColNum)
                            Else
                                dgvRevControl(1, z).Value = "*Varies*"
                            End If
                        Else
                            If dgvRevControl(2, z).Value = True Then
                                rt.RevisionTableRows(j).Item(RevCheckNameColNum + 1).Text = dgvRevControl(1, z).Value
                            End If
                        End If
                    Case UCase(My.Settings.RTSApprovedCol)
                        If read = True Then
                            If dgvRevControl(1, z).Value = Nothing OrElse dgvRevControl(1, z).Value = contents((k * rt.RevisionTableColumns.Count) + RevCheckApproveColNum) Then
                                dgvRevControl(1, z).Value = contents((k * rt.RevisionTableColumns.Count) + RevCheckApproveColNum)
                            Else
                                dgvRevControl(1, z).Value = "*Varies*"
                            End If
                        Else
                            If dgvRevControl(2, z).Value = True Then
                                rt.RevisionTableRows(j).Item(RevCheckApproveColNum + 1).Text = dgvRevControl(1, z).Value
                            End If
                        End If
                    Case UCase(My.Settings.RTS1Item)
                        If read = True Then
                            If dgvRevControl(1, z).Value = Nothing OrElse dgvRevControl(1, z).Value = contents((k * rt.RevisionTableColumns.Count) + RevCheck1ColNum) Then
                                If RevCheck1ColNum <> 0 Then dgvRevControl(1, z).Value = contents((k * rt.RevisionTableColumns.Count) + RevCheck1ColNum)
                            Else
                                dgvRevControl(1, z).Value = "*Varies*"
                            End If
                        Else
                            If dgvRevControl(2, z).Value = True Then
                                rt.RevisionTableRows(j).Item(RevCheck1ColNum + 1).Text = dgvRevControl(1, z).Value
                            End If
                        End If
                    Case UCase(My.Settings.RTS2Item)
                        If read = True Then
                            If dgvRevControl(1, z).Value = Nothing OrElse dgvRevControl(1, z).Value = contents((k * rt.RevisionTableColumns.Count) + RevCheck2ColNum) Then
                                If RevCheck2ColNum <> 0 Then dgvRevControl(1, z).Value = contents((k * rt.RevisionTableColumns.Count) + RevCheck2ColNum)
                            Else
                                dgvRevControl(1, z).Value = "*Varies*"
                            End If
                        Else
                            If dgvRevControl(2, z).Value = True Then
                                rt.RevisionTableRows(j).Item(RevCheck2ColNum + 1).Text = dgvRevControl(1, z).Value
                            End If
                        End If
                    Case UCase(My.Settings.RTS3Item)
                        If read = True Then
                            If dgvRevControl(1, z).Value = Nothing OrElse dgvRevControl(1, z).Value = contents((k * rt.RevisionTableColumns.Count) + RevCheck3ColNum) Then
                                If RevCheck3ColNum <> 0 Then dgvRevControl(1, z).Value = contents((k * rt.RevisionTableColumns.Count) + RevCheck3ColNum)
                            Else
                                dgvRevControl(1, z).Value = "*Varies*"
                            End If
                        Else
                            If dgvRevControl(2, z).Value = True Then
                                rt.RevisionTableRows(j).Item(RevCheck3ColNum + 1).Text = dgvRevControl(1, z).Value
                            End If
                        End If
                    Case UCase(My.Settings.RTS4Item)
                        If read = True Then
                            If dgvRevControl(1, z).Value = Nothing OrElse dgvRevControl(1, z).Value = contents((k * rt.RevisionTableColumns.Count) + RevCheck4ColNum) Then
                                If RevCheck4ColNum <> 0 Then dgvRevControl(1, z).Value = contents((k * rt.RevisionTableColumns.Count) + RevCheck4ColNum)
                            Else
                                dgvRevControl(1, z).Value = "*Varies*"
                            End If
                        Else
                            If dgvRevControl(2, z).Value = True Then
                                rt.RevisionTableRows(j).Item(RevCheck4ColNum + 1).Text = dgvRevControl(1, z).Value
                            End If
                        End If
                    Case UCase(My.Settings.RTS5Item)
                        If read = True Then
                            If dgvRevControl(1, z).Value = Nothing OrElse dgvRevControl(1, z).Value = contents((k * rt.RevisionTableColumns.Count) + RevCheck4ColNum) Then
                                If RevCheck5ColNum <> 0 Then dgvRevControl(1, z).Value = contents((k * rt.RevisionTableColumns.Count) + RevCheck5ColNum)
                            Else
                                dgvRevControl(1, z).Value = "*Varies*"
                            End If
                        Else
                            If dgvRevControl(2, z).Value = True Then
                                rt.RevisionTableRows(j).Item(RevCheck5ColNum + 1).Text = dgvRevControl(1, z).Value
                            End If
                        End If
                End Select
            Next
            k += 1
            j += 1
        Next
    End Sub
    Private Sub InitializeiProperties()
        'Go through each iProperty in both the part and drawing and assign them to their respective output strings
        'When the values between documents are different, display *Varies* as the output.
        Dim Type, Status, DateString As String
        Dim cmbMaterial As DataGridViewComboBoxCell = dgvSummary(1, dgvSummary.RowCount - 1)
        For Row = 0 To dgvSummary.RowCount - 1
            dgvSummary(dgvSummary.Columns("SummaryIsDirty").Index, Row).Value = "False"
            If dgvSummary("SumItem", Row).Value = "Material:" Then
                If cmbMaterial.Value <> "*Varies*" AndAlso cmbMaterial.Value = "" AndAlso cmbMaterial.Value Is Nothing Then
                    If cmbMaterial.Items.Contains(SumDic.Item(dgvSummary("SumItem", Row).Value).ModelValue) Then
                        cmbMaterial.Value = SumDic.Item(dgvSummary("SumItem", Row).Value).ModelValue
                    ElseIf Not cmbMaterial.Items.Contains(SumDic.Item(dgvSummary("SumItem", Row).Value).ModelValue) Then
                        cmbMaterial.Items.Add(SumDic.Item(dgvSummary("SumItem", Row).Value).ModelValue)
                        cmbMaterial.Value = SumDic.Item(dgvSummary("SumItem", Row).Value).ModelValue
                    Else
                        cmbMaterial.Items.Add("*Varies*")
                        cmbMaterial.Value = "*Varies*"
                    End If
                End If
            ElseIf SumDicRows(dgvSummary("SumItem", Row).Value).Cells("SumModel").Value = "" Then
                SumDicRows(dgvSummary("SumItem", Row).Value).Cells("SumModel").Value = SumDic.Item(dgvSummary("SumItem", Row).Value).ModelValue
            ElseIf SumDicRows(dgvSummary("SumItem", Row).Value).Cells("SumModel").Value <> SumDic.Item(dgvSummary("SumItem", Row).Value).ModelValue Then
                SumDicRows(dgvSummary("SumItem", Row).Value).Cells("SumModel").Value = "*Varies*"
            End If
            Try
                If SumDicRows(dgvSummary("SumItem", Row).Value).Cells("SumDrawing").Value = "" Then
                    SumDicRows(dgvSummary("SumItem", Row).Value).Cells("SumDrawing").Value = SumDic.Item(dgvSummary("SumItem", Row).Value).DrawingValue
                ElseIf SumDicRows(dgvSummary("SumItem", Row).Value).Cells("SumDrawing").Value <> SumDic.Item(dgvSummary("SumItem", Row).Value).DrawingValue Then
                    SumDicRows(dgvSummary("SumItem", Row).Value).Cells("SumDrawing").Value = "*Varies*"
                End If
            Catch
            End Try
        Next
        For Each Row As DataGridViewRow In dgvProject.Rows
            dgvProject(dgvProject.Columns("ProjectIsDirty").Index, Row.Index).Value = "False"
            If InStr(Strings.LCase(dgvProject("ProjItem", Row.Index).Value), "date") <> 0 Then
                If Date.Parse(ProjDic.Item(dgvProject("ProjItem", Row.Index).Value).ModelValue).ToString(CultureInfo.CurrentCulture.DateTimeFormat.ShortDatePattern) = "01/01/01" Then
                    DateString = ""
                ElseIf Date.Parse(ProjDic.Item(dgvProject("ProjItem", Row.Index).Value).ModelValue).ToString(CultureInfo.CurrentCulture.DateTimeFormat.ShortDatePattern) <> ProjDicRows(dgvProject("ProjItem", Row.Index).Value).Cells("ProjModel").Value And
                    ProjDicRows(dgvProject("ProjItem", Row.Index).Value).Cells("ProjModel").Value <> Nothing Then
                    DateString = "*Varies*"
                Else
                    DateString = Date.Parse(ProjDic.Item(dgvProject("ProjItem", Row.Index).Value).ModelValue).ToString(CultureInfo.CurrentCulture.DateTimeFormat.ShortDatePattern)
                End If
                ProjDicRows(dgvProject("ProjItem", Row.Index).Value).Cells("ProjModel").Value = DateString
                Try
                    If Date.Parse(ProjDic.Item(dgvProject("ProjItem", Row.Index).Value).DrawingValue).ToString(CultureInfo.CurrentCulture.DateTimeFormat.ShortDatePattern) = "01/01/01" Then
                        DateString = ""
                    ElseIf Date.Parse(ProjDic.Item(dgvProject("ProjItem", Row.Index).Value).DrawingValue).ToString(CultureInfo.CurrentCulture.DateTimeFormat.ShortDatePattern) <> ProjDicRows(dgvProject("ProjItem", Row.Index).Value).Cells("ProjDrawing").Value And
                    ProjDicRows(dgvProject("ProjItem", Row.Index).Value).Cells("ProjDrawing").Value <> Nothing Then
                        DateString = "*Varies*"
                    Else
                        DateString = Date.Parse(ProjDic.Item(dgvProject("ProjItem", Row.Index).Value).DrawingValue).ToString(CultureInfo.CurrentCulture.DateTimeFormat.ShortDatePattern)
                    End If
                Catch
                End Try
                ProjDicRows(dgvProject("ProjItem", Row.Index).Value).Cells("ProjDrawing").Value = DateString
            Else
                If ProjDicRows(dgvProject("ProjItem", Row.Index).Value).Cells("ProjModel").Value = "" Then
                    ProjDicRows(dgvProject("ProjItem", Row.Index).Value).Cells("ProjModel").Value = ProjDic.Item(dgvProject("ProjItem", Row.Index).Value).ModelValue
                ElseIf ProjDicRows(dgvProject("ProjItem", Row.Index).Value).Cells("ProjModel").Value <> ProjDic.Item(dgvProject("ProjItem", Row.Index).Value).ModelValue Then
                    ProjDicRows(dgvProject("ProjItem", Row.Index).Value).Cells("ProjModel").Value = "*Varies*"
                End If
                Try
                    If ProjDicRows(dgvProject("ProjItem", Row.Index).Value).Cells("ProjDrawing").Value = "" Then
                        ProjDicRows(dgvProject("ProjItem", Row.Index).Value).Cells("ProjDrawing").Value = ProjDic.Item(dgvProject("ProjItem", Row.Index).Value).DrawingValue
                    ElseIf ProjDicRows(dgvProject("ProjItem", Row.Index).Value).Cells("ProjDrawing").Value <> ProjDic.Item(dgvProject("ProjItem", Row.Index).Value).DrawingValue Then
                        ProjDicRows(dgvProject("ProjItem", Row.Index).Value).Cells("ProjDrawing").Value = "*Varies*"
                    End If
                Catch
                End Try
            End If
        Next
        For Each Row As DataGridViewRow In dgvStatus.Rows
            dgvStatus(dgvStatus.Columns("StatusIsDirty").Index, Row.Index).Value = "False"
            Try
                If dgvStatus("StatusItem", Row.Index).Value = "Design State:" Then
                    Status = ConvertDesignState(StatusDic.Item(dgvStatus("StatusItem", Row.Index).Value).ModelValue)
                    StatusDicRows(dgvStatus("StatusItem", Row.Index).Value).Cells("StatusModel").Value = Status
                    Status = ConvertDesignState(StatusDic.Item(dgvStatus("StatusItem", Row.Index).Value).DrawingValue)
                    StatusDicRows(dgvStatus("StatusItem", Row.Index).Value).Cells("StatusDrawing").Value = Status
                ElseIf InStr(Strings.LCase(dgvStatus("StatusItem", Row.Index).Value), "date") <> 0 Then
                    If Date.Parse(StatusDic.Item(dgvStatus("StatusItem", Row.Index).Value).ModelValue).ToString(CultureInfo.CurrentCulture.DateTimeFormat.ShortDatePattern) = "01/01/01" Then
                        DateString = ""
                    ElseIf Date.Parse(StatusDic.Item(dgvStatus("StatusItem", Row.Index).Value).ModelValue).ToString(CultureInfo.CurrentCulture.DateTimeFormat.ShortDatePattern) <>
                        StatusDicRows(dgvStatus("StatusItem", Row.Index).Value).Cells("StatusModel").Value AndAlso
                        StatusDicRows(dgvStatus("StatusItem", Row.Index).Value).Cells("StatusModel").Value <> Nothing Then
                        DateString = "*Varies*"
                    Else
                        DateString = Date.Parse(StatusDic.Item(dgvStatus("StatusItem", Row.Index).Value).ModelValue).ToString(CultureInfo.CurrentCulture.DateTimeFormat.ShortDatePattern)
                    End If
                    StatusDicRows(dgvStatus("StatusItem", Row.Index).Value).Cells("StatusModel").Value = DateString
                    If Date.Parse(StatusDic.Item(dgvStatus("StatusItem", Row.Index).Value).DrawingValue).ToString(CultureInfo.CurrentCulture.DateTimeFormat.ShortDatePattern) = "01/01/01" Then
                        DateString = ""
                    ElseIf Date.Parse(StatusDic.Item(dgvStatus("StatusItem", Row.Index).Value).DrawingValue).ToString(CultureInfo.CurrentCulture.DateTimeFormat.ShortDatePattern) <>
                        StatusDicRows(dgvStatus("StatusItem", Row.Index).Value).Cells("StatusDrawing").Value AndAlso
                         StatusDicRows(dgvStatus("StatusItem", Row.Index).Value).Cells("StatusDrawing").Value <> Nothing Then
                        DateString = "*Varies*"
                    Else
                        DateString = Date.Parse(StatusDic.Item(dgvStatus("StatusItem", Row.Index).Value).DrawingValue).ToString(CultureInfo.CurrentCulture.DateTimeFormat.ShortDatePattern)
                    End If
                    StatusDicRows(dgvStatus("StatusItem", Row.Index).Value).Cells("StatusDrawing").Value = DateString
                Else
                    If StatusDicRows(dgvStatus("StatusItem", Row.Index).Value).Cells("StatusModel").Value = "" Then
                        StatusDicRows(dgvStatus("StatusItem", Row.Index).Value).Cells("StatusModel").Value = StatusDic.Item(dgvStatus("StatusItem", Row.Index).Value).ModelValue
                    ElseIf StatusDicRows(dgvStatus("StatusItem", Row.Index).Value).Cells("StatusModel").Value <> StatusDic.Item(dgvStatus("StatusItem", Row.Index).Value).ModelValue Then
                        StatusDicRows(dgvStatus("StatusItem", Row.Index).Value).Cells("StatusModel").Value = "*Varies*"
                    End If
                    If StatusDicRows(dgvStatus("StatusItem", Row.Index).Value).Cells("StatusDrawing").Value = "" Then
                        StatusDicRows(dgvStatus("StatusItem", Row.Index).Value).Cells("StatusDrawing").Value = StatusDic.Item(dgvStatus("StatusItem", Row.Index).Value).DrawingValue
                    ElseIf StatusDicRows(dgvStatus("StatusItem", Row.Index).Value).Cells("StatusDrawing").Value <> StatusDic.Item(dgvStatus("StatusItem", Row.Index).Value).DrawingValue Then
                        StatusDicRows(dgvStatus("StatusItem", Row.Index).Value).Cells("StatusDrawing").Value = "*Varies*"
                    End If
                End If
            Catch
            End Try
        Next

        For Each Item In CustomModDic

            Dim skip As Boolean = False
            For Each row In dgvCustomModel.Rows
                dgvCustomModel(dgvCustomModel.Columns("ModelIsDirty").Index, row.Index).Value = "False"
                If dgvCustomModel(dgvCustomModel.Columns("PCusName").Index, row.index).Value = Item.Key Then
                    skip = True
                    Exit For
                End If
            Next
            Type = ConvertiPropType(True, Item.Value.Type)
            If skip = False Then
                dgvCustomModel.Rows.Add(Item.Key, Item.Value.ModelValue, Type.ToString)
            End If

        Next
        For Each Item In CustomDrawDic
            Dim skip As Boolean = False
            For Each row In dgvCustomDrawing.Rows
                dgvCustomDrawing(dgvCustomDrawing.Columns("DrawingIsDirty").Index, row.Index).Value = "False"
                If dgvCustomDrawing(dgvCustomDrawing.Columns("DCusName").Index, row.index).Value = Item.Key Then
                    skip = True
                    Exit For
                End If
            Next
            Type = ConvertiPropType(False, Item.Value.Type)
            If skip = False Then
                dgvCustomDrawing.Rows.Add(Item.Key, Item.Value.ModelValue, Type.ToString)
            End If
        Next
    End Sub
    Private Function ConvertiPropType(Model As Boolean, iPropType As Object) As String
        If Model = True Then
            If iPropType.name = "String" Then Return "Text"
            If iPropType.name = "DateTime" Then Return "Date"
            If iPropType.name = "Int32" Then Return "Number"
            If iPropType.name = "Boolean" Then Return "Yes or No"
        Else
            If iPropType = "String" Then Return "Text"
            If iPropType = "DateTime" Then Return "Date"
            If iPropType = "Int32" Then Return "Number"
            If iPropType = "Boolean" Then Return "Yes or No"
        End If
        Return Nothing
    End Function
    Private Function ConvertDesignState(State As Integer) As String
        If State = 1 Then Return "Work In Progress"
        If State = 2 Then Return "Pending"
        If State = 3 Then Return "Released"
        Return "Work In Progress"
    End Function
#End Region
    Private Sub btnCancel_Click(sender As System.Object, e As System.EventArgs) Handles btnCancel.Click
        Me.Hide()
    End Sub
    Private Sub btnOK_Click(sender As System.Object, e As System.EventArgs) Handles btnOK.Click
        Dim Path As Documents
        Dim oDoc As Document = Nothing
        Dim OpenDocs As New ArrayList

        Path = _invApp.Documents
        Dim Archive As String = ""
        Dim DrawSource As String = ""
        Dim DrawingName As String = ""
        Main.CreateOpenDocs(OpenDocs)
        Call PopulateiProps(Path, oDoc, Archive, DrawingName, DrawSource, OpenDocs, Read:=False)
        'Call Main.OpenSelected(Path, oDoc, Archive, DrawingName, DrawSource)
        Me.Hide()
        Me.Focus()
    End Sub
#Region "Write iProperties"
    Private Sub WriteProps(ByRef oDoc As Document, ByRef Document As String)
        For Each row In dgvSummary.Rows
            If dgvSummary(dgvSummary.Columns("SummaryIsDirty").Index, row.index).Value = "True" Then
                ' If dgvSummary(dgvSummary.Columns(Document).Index, row.index).Value <> "*Varies*" Then
                If dgvSummary(dgvSummary.Columns("Sum" & Document).Index, row.index).Value Is Nothing Then dgvSummary(dgvSummary.Columns("Sum" & Document).Index, row.index).Value = ""
                Select Case dgvSummary(0, row.index).Value

                    Case "Title:"
                        oDoc.PropertySets.Item("{F29F85E0-4FF9-1068-AB91-08002B27B3D9}").ItemByPropId("2").Value = dgvSummary(dgvSummary.Columns("Sum" & Document).Index, row.index).Value
                    Case "Subject:"
                        oDoc.PropertySets.Item("{F29F85E0-4FF9-1068-AB91-08002B27B3D9}").ItemByPropId("3").Value = dgvSummary(dgvSummary.Columns("Sum" & Document).Index, row.index).Value
                    Case "Author:"
                        oDoc.PropertySets.Item("{F29F85E0-4FF9-1068-AB91-08002B27B3D9}").ItemByPropId("4").Value = dgvSummary(dgvSummary.Columns("Sum" & Document).Index, row.index).Value
                    Case "Manager:"
                        oDoc.PropertySets.Item("{D5CDD502-2E9C-101B-9397-08002B2CF9AE}").ItemByPropId("14").Value = dgvSummary(dgvSummary.Columns("Sum" & Document).Index, row.index).Value
                    Case "Company:"
                        oDoc.PropertySets.Item("{D5CDD502-2E9C-101B-9397-08002B2CF9AE}").ItemByPropId("15").Value = dgvSummary(dgvSummary.Columns("Sum" & Document).Index, row.index).Value
                    Case "Category:"
                        oDoc.PropertySets.Item("{D5CDD502-2E9C-101B-9397-08002B2CF9AE}").ItemByPropId("2").Value = dgvSummary(dgvSummary.Columns("Sum" & Document).Index, row.index).Value
                    Case "Keywords:"
                        oDoc.PropertySets.Item("{F29F85E0-4FF9-1068-AB91-08002B27B3D9}").ItemByPropId("5").Value = dgvSummary(dgvSummary.Columns("Sum" & Document).Index, row.index).Value
                    Case "Comments:"
                        oDoc.PropertySets.Item("{F29F85E0-4FF9-1068-AB91-08002B27B3D9}").ItemByPropId("6").Value = dgvSummary(dgvSummary.Columns("Sum" & Document).Index, row.index).Value
                    Case "Material:"
                        If dgvSummary(dgvSummary.Columns("Sum" & Document).Index, row.index).Value <> "*Varies*" AndAlso
                            dgvSummary.Columns("Sum" & Document).Index = 1 AndAlso
                            dgvSummary(dgvSummary.Columns("Sum" & Document).Index, row.index).Value <> "" Then
                            Try
                                oDoc.componentdefinition.material.name = dgvSummary(dgvSummary.Columns("Sum" & Document).Index, row.index).Value
                            Catch
                            End Try
                        End If
                End Select
            End If
        Next
        For Each row In dgvProject.Rows
            If dgvProject(dgvProject.Columns("ProjectIsDirty").Index, row.index).Value = "True" Then
                If dgvProject(dgvProject.Columns("Proj" & Document).Index, row.index).Value Is Nothing Then dgvProject(dgvProject.Columns("Proj" & Document).Index, row.index).Value = ""
                'If dgvProject(1, row.index).Value <> "*Varies*" Then
                Select Case dgvProject(0, row.index).Value
                    Case "Part Number:"
                        oDoc.PropertySets.Item("{32853F0F-3444-11D1-9E93-0060B03C1CA6}").ItemByPropId("5").Value = dgvProject(dgvProject.Columns("Proj" & Document).Index, row.index).Value
                    Case "Stock Number:"
                        oDoc.PropertySets.Item("{32853F0F-3444-11D1-9E93-0060B03C1CA6}").ItemByPropId("55").Value = dgvProject(dgvProject.Columns("Proj" & Document).Index, row.index).Value
                    Case "Description:"
                        oDoc.PropertySets.Item("{32853F0F-3444-11D1-9E93-0060B03C1CA6}").ItemByPropId("29").Value = dgvProject(dgvProject.Columns("Proj" & Document).Index, row.index).Value
                    Case "Revision Number:"
                        oDoc.PropertySets.Item("{F29F85E0-4FF9-1068-AB91-08002B27B3D9}").ItemByPropId("9").Value = dgvProject(dgvProject.Columns("Proj" & Document).Index, row.index).Value
                    Case "Project:"
                        oDoc.PropertySets.Item("{32853F0F-3444-11D1-9E93-0060B03C1CA6}").ItemByPropId("7").Value = dgvProject(dgvProject.Columns("Proj" & Document).Index, row.index).Value
                    Case "Designer:"
                        oDoc.PropertySets.Item("{32853F0F-3444-11D1-9E93-0060B03C1CA6}").ItemByPropId("41").Value = dgvProject(dgvProject.Columns("Proj" & Document).Index, row.index).Value
                    Case "Engineer:"
                        oDoc.PropertySets.Item("{32853F0F-3444-11D1-9E93-0060B03C1CA6}").ItemByPropId("42").Value = dgvProject(dgvProject.Columns("Proj" & Document).Index, row.index).Value
                    Case "Authority:"
                        oDoc.PropertySets.Item("{32853F0F-3444-11D1-9E93-0060B03C1CA6}").ItemByPropId("43").Value = dgvProject(dgvProject.Columns("Proj" & Document).Index, row.index).Value
                    Case "Cost Center:"
                        oDoc.PropertySets.Item("{32853F0F-3444-11D1-9E93-0060B03C1CA6}").ItemByPropId("9").Value = dgvProject(dgvProject.Columns("Proj" & Document).Index, row.index).Value
                    Case "Estimated Cost:"
                        oDoc.PropertySets.Item("{32853F0F-3444-11D1-9E93-0060B03C1CA6}").ItemByPropId("36").Value = dgvProject(dgvProject.Columns("Proj" & Document).Index, row.index).Value
                    Case "Creation Date:"
                        If dgvProject(dgvProject.Columns("Proj" & Document).Index, row.index).Value = "" Then
                            oDoc.PropertySets.Item("{32853F0F-3444-11D1-9E93-0060B03C1CA6}").ItemByPropId("4").Value = #1/1/1601#
                        Else
                            oDoc.PropertySets.Item("{32853F0F-3444-11D1-9E93-0060B03C1CA6}").ItemByPropId("4").Value = dgvProject(dgvProject.Columns("Proj" & Document).Index, row.index).Value

                        End If
                    Case "Vendor:"
                        oDoc.PropertySets.Item("{32853F0F-3444-11D1-9E93-0060B03C1CA6}").ItemByPropId("4").Value = dgvProject(dgvProject.Columns("Proj" & Document).Index, row.index).Value
                    Case "Vendor:"
                        oDoc.PropertySets.Item("{32853F0F-3444-11D1-9E93-0060B03C1CA6}").ItemByPropId("30").Value = dgvProject(dgvProject.Columns("Proj" & Document).Index, row.index).Value
                    Case "WEB Link:"
                        oDoc.PropertySets.Item("{32853F0F-3444-11D1-9E93-0060B03C1CA6}").ItemByPropId("23").Value = dgvProject(dgvProject.Columns("Proj" & Document).Index, row.index).Value
                End Select
            End If

        Next
        For Each row In dgvStatus.Rows
            If dgvStatus(dgvStatus.Columns("StatusIsDirty").Index, row.index).Value = "True" Then
                If dgvStatus(dgvStatus.Columns("Status" & Document).Index, row.index).Value Is Nothing Then dgvStatus(dgvStatus.Columns("Status" & Document).Index, row.index).Value = ""
                ' If dgvStatus(1, row.index).Value <> "*Varies*" Then
                Select Case dgvStatus(0, row.index).Value
                    Case "Part Number:"
                        oDoc.PropertySets.Item("{32853F0F-3444-11D1-9E93-0060B03C1CA6}").ItemByPropId("5").Value = dgvStatus(dgvStatus.Columns("Status" & Document).Index, row.index).Value
                    Case "Stock Number:"
                        oDoc.PropertySets.Item("{32853F0F-3444-11D1-9E93-0060B03C1CA6}").ItemByPropId("55").Value = dgvStatus(dgvStatus.Columns("Status" & Document).Index, row.index).Value
                    Case "Status:"
                        oDoc.PropertySets.Item("{32853F0F-3444-11D1-9E93-0060B03C1CA6}").ItemByPropId("17").Value = dgvStatus(dgvStatus.Columns("Status" & Document).Index, row.index).Value
                    Case "Design State:"
                        Select Case dgvStatus(dgvStatus.Columns("Status" & Document).Index, row.index).Value
                            Case "Work In Progress"
                                oDoc.PropertySets.Item("{32853F0F-3444-11D1-9E93-0060B03C1CA6}").ItemByPropId("40").Value = 0
                            Case "Pending"
                                oDoc.PropertySets.Item("{32853F0F-3444-11D1-9E93-0060B03C1CA6}").ItemByPropId("40").Value = 2
                            Case "Released"
                                oDoc.PropertySets.Item("{32853F0F-3444-11D1-9E93-0060B03C1CA6}").ItemByPropId("40").Value = 3
                        End Select
                    Case "Checked By:"
                        oDoc.PropertySets.Item("{32853F0F-3444-11D1-9E93-0060B03C1CA6}").ItemByPropId("10").Value = dgvStatus(dgvStatus.Columns("Status" & Document).Index, row.index).Value
                    Case "Checked Date:"
                        If dgvStatus(dgvStatus.Columns("Status" & Document).Index, row.index).Value = "" Then
                            oDoc.PropertySets.Item("{32853F0F-3444-11D1-9E93-0060B03C1CA6}").ItemByPropId("11").Value = #1/1/1601#
                        Else
                            oDoc.PropertySets.Item("{32853F0F-3444-11D1-9E93-0060B03C1CA6}").ItemByPropId("11").Value = dgvStatus(dgvStatus.Columns("Status" & Document).Index, row.index).Value
                        End If
                    Case "Eng. Approved By:"
                        oDoc.PropertySets.Item("{32853F0F-3444-11D1-9E93-0060B03C1CA6}").ItemByPropId("12").Value = dgvStatus(dgvStatus.Columns("Status" & Document).Index, row.index).Value
                    Case "Eng. Approved Date:"
                        If dgvStatus(dgvStatus.Columns("Status" & Document).Index, row.index).Value = "" Then
                            oDoc.PropertySets.Item("{32853F0F-3444-11D1-9E93-0060B03C1CA6}").ItemByPropId("13").Value = #1/1/1601#
                        Else
                            oDoc.PropertySets.Item("{32853F0F-3444-11D1-9E93-0060B03C1CA6}").ItemByPropId("13").Value = dgvStatus(dgvStatus.Columns("Status" & Document).Index, row.index).Value
                        End If
                    Case "Mfg. Approved By:"
                        oDoc.PropertySets.Item("{32853F0F-3444-11D1-9E93-0060B03C1CA6}").ItemByPropId("34").Value = dgvStatus(dgvStatus.Columns("Status" & Document).Index, row.index).Value
                    Case "Mfg. Approved Date:"
                        If dgvStatus(dgvStatus.Columns("Status" & Document).Index, row.index).Value = "" Then
                            oDoc.PropertySets.Item("{32853F0F-3444-11D1-9E93-0060B03C1CA6}").ItemByPropId("35").Value = #1/1/1601#
                        Else
                            oDoc.PropertySets.Item("{32853F0F-3444-11D1-9E93-0060B03C1CA6}").ItemByPropId("35").Value = dgvStatus(dgvStatus.Columns("Status" & Document).Index, row.index).Value
                        End If
                End Select
            End If
            oDoc.Update()
        Next
    End Sub
    Private Sub WriteDrawingProps(ByRef oDoc As Document)
        'If txtTitle.Text <> "*Varies*" And My.Settings.Title = "Drawing" Then
        '    oDoc.PropertySets.Item("{F29F85E0-4FF9-1068-AB91-08002B27B3D9}").ItemByPropId("2").Value = txtTitle.Text
        'End If
        'If txtSubject.Text <> "*Varies*" And My.Settings.Subject = "Drawing" Then
        '    oDoc.PropertySets.Item("{F29F85E0-4FF9-1068-AB91-08002B27B3D9}").ItemByPropId("3").Value = txtSubject.Text
        'End If
        'If txtManager.Text <> "*Varies*" And My.Settings.Manager = "Drawing" Then
        '    oDoc.PropertySets.Item("{D5CDD502-2E9C-101B-9397-08002B2CF9AE}").ItemByPropId("14").Value = txtManager.Text
        'End If
        'If txtCategory.Text <> "*Varies*" And My.Settings.Category = "Drawing" Then
        '    oDoc.PropertySets.Item("{D5CDD502-2E9C-101B-9397-08002B2CF9AE}").ItemByPropId("2").Value = txtCategory.Text
        'End If
        'If txtKeywords.Text <> "*Varies*" And My.Settings.Keywords = "Drawing" Then
        '    oDoc.PropertySets.Item("{F29F85E0-4FF9-1068-AB91-08002B27B3D9}").ItemByPropId("5").Value = txtKeywords.Text
        'End If
        'If txtAuthor.Text <> "*Varies*" And My.Settings.Author = "Drawing" Then
        '    oDoc.PropertySets.Item("{F29F85E0-4FF9-1068-AB91-08002B27B3D9}").ItemByPropId("4").Value = txtAuthor.Text
        'End If
        'If txtCompany.Text <> "*Varies*" And My.Settings.Company = "Drawing" Then
        '    oDoc.PropertySets.Item("{D5CDD502-2E9C-101B-9397-08002B2CF9AE}").ItemByPropId("15").Value = txtCompany.Text
        'End If
        'If txtComments.Text <> "*Varies*" And My.Settings.Comments = "Drawing" Then
        '    oDoc.PropertySets.Item("{F29F85E0-4FF9-1068-AB91-08002B27B3D9}").ItemByPropId("6").Value = txtComments.Text
        'End If
        'If txtDescription.Text <> "*Varies*" And My.Settings.Description = "Drawing" Then
        '    oDoc.PropertySets.Item("{32853F0F-3444-11D1-9E93-0060B03C1CA6}").ItemByPropId("29").Value = txtDescription.Text
        'End If
        'If txtPartNumber.Text <> "*Varies*" And My.Settings.PPartNumber = "Drawing" Then
        '    oDoc.PropertySets.Item("{32853F0F-3444-11D1-9E93-0060B03C1CA6}").ItemByPropId("5").Value = txtPartNumber.Text
        'End If
        'If txtStockNumber.Text <> "*Varies*" And My.Settings.PStockNumber = "Drawing" Then
        '    oDoc.PropertySets.Item("{32853F0F-3444-11D1-9E93-0060B03C1CA6}").ItemByPropId("55").Value = txtStockNumber.Text
        'End If
        'If txtProject.Text <> "*Varies*" And My.Settings.Project = "Drawing" Then
        '    oDoc.PropertySets.Item("{32853F0F-3444-11D1-9E93-0060B03C1CA6}").ItemByPropId("7").Value = txtProject.Text
        'End If
        'If txtDesigner.Text <> "*Varies*" And My.Settings.Designer = "Drawing" Then
        '    oDoc.PropertySets.Item("{32853F0F-3444-11D1-9E93-0060B03C1CA6}").ItemByPropId("41").Value = txtDesigner.Text
        'End If
        'If txtVendor.Text <> "*Varies*" And My.Settings.Vendor = "Drawing" Then
        '    oDoc.PropertySets.Item("{32853F0F-3444-11D1-9E93-0060B03C1CA6}").ItemByPropId("30").Value = txtVendor.Text
        'End If
        'If dtModelDate.Checked = True And My.Settings.ModelDate = "Drawing" Then
        '    oDoc.PropertySets.Item("{32853F0F-3444-11D1-9E93-0060B03C1CA6}").ItemByPropId("4").Value = dtModelDate.Value
        'End If
        'If txtEngineer.Text <> "*Varies*" And My.Settings.Engineer = "Drawing" Then
        '    oDoc.PropertySets.Item("{32853F0F-3444-11D1-9E93-0060B03C1CA6}").ItemByPropId("42").Value = txtEngineer.Text
        'End If
        'If txtStatus.Text <> "*Varies*" And My.Settings.Status = "Drawing" Then
        '    oDoc.PropertySets.Item("{32853F0F-3444-11D1-9E93-0060B03C1CA6}").ItemByPropId("17").Value = txtStatus.Text
        'End If
        'If txtRevision.Text <> "*Varies*" And My.Settings.Revision = "Drawing" Then
        '    oDoc.PropertySets.Item("{F29F85E0-4FF9-1068-AB91-08002B27B3D9}").ItemByPropId("9").Value = txtRevision.Text
        'End If
        'If txtDrawingBy.Text <> "*Varies*" And My.Settings.DrawingBy = "Drawing" Then
        '    oDoc.PropertySets.Item("{32853F0F-3444-11D1-9E93-0060B03C1CA6}").ItemByPropId("41").Value = txtDrawingBy.Text
        'End If
        'If cmbDesignState.Text = "Released" And My.Settings.DesignState = "Drawing" Then
        '    oDoc.PropertySets.Item("{32853F0F-3444-11D1-9E93-0060B03C1CA6}").ItemByPropId("40").Value = 3
        'ElseIf cmbDesignState.Text = "Pending" And My.Settings.DesignState = "Drawing" Then
        '    oDoc.PropertySets.Item("{32853F0F-3444-11D1-9E93-0060B03C1CA6}").ItemByPropId("40").Value = 2
        'ElseIf cmbDesignState.Text = "Work In Progress" And My.Settings.DesignState = "Drawing" Then
        '    oDoc.PropertySets.Item("{32853F0F-3444-11D1-9E93-0060B03C1CA6}").ItemByPropId("40").Value = 1
        'End If
        'If txtCheckedBy.Text <> "*Varies*" And My.Settings.CheckedBy = "Drawing" Then
        '    oDoc.PropertySets.Item("{32853F0F-3444-11D1-9E93-0060B03C1CA6}").ItemByPropId("10").Value = txtCheckedBy.Text
        'End If
        'If dtCheckDate.Checked = True And My.Settings.CheckedDate = "Drawing" Then
        '    oDoc.PropertySets.Item("{32853F0F-3444-11D1-9E93-0060B03C1CA6}").ItemByPropId("11").Value = dtCheckDate.Value
        'Else
        '    ' oDoc.PropertySets.Item("{32853F0F-3444-11D1-9E93-0060B03C1CA6}").ItemByPropId("11").Expression = "1601, 1, 1"
        'End If
        'If dtDrawingDate.Checked = True And My.Settings.DrawingDate = "Drawing" Then
        '    oDoc.PropertySets.Item("{32853F0F-3444-11D1-9E93-0060B03C1CA6}").ItemByPropId("4").Value = dtDrawingDate.Value
        'Else
        '    ' oDoc.PropertySets.Item("{32853F0F-3444-11D1-9E93-0060B03C1CA6}").ItemByPropId("4").Expression = "1601, 1, 1"
        'End If
        'Try
        '    If txtRef0.Text <> "*Varies*" And My.Settings.Custom = "Drawing" Then
        '        oDoc.PropertySets.Item("Inventor User Defined Properties").Item("Reference0").Value = txtRef0.Text
        '    End If
        '    If txtRef1.Text <> "*Varies*" And My.Settings.Custom = "Drawing" Then
        '        oDoc.PropertySets.Item("Inventor User Defined Properties").Item("Reference1").Value = txtRef1.Text
        '    End If
        '    If txtRef2.Text <> "*Varies*" And My.Settings.Custom = "Drawing" Then
        '        oDoc.PropertySets.Item("Inventor User Defined Properties").Item("Reference2").Value = txtRef2.Text
        '    End If
        '    If txtRef3.Text <> "*Varies*" And My.Settings.Custom = "Drawing" Then
        '        oDoc.PropertySets.Item("Inventor User Defined Properties").Item("Reference3").Value = txtRef3.Text
        '    End If
        '    If txtRef4.Text <> "*Varies*" And My.Settings.Custom = "Drawing" Then
        '        oDoc.PropertySets.Item("Inventor User Defined Properties").Item("Reference4").Value = txtRef4.Text
        '    End If
        '    If txtRef5.Text <> "*Varies*" And My.Settings.Custom = "Drawing" Then
        '        oDoc.PropertySets.Item("Inventor User Defined Properties").Item("Reference5").Value = txtRef5.Text
        '    End If
        '    If txtRef6.Text <> "*Varies*" And My.Settings.Custom = "Drawing" Then
        '        oDoc.PropertySets.Item("Inventor User Defined Properties").Item("Reference6").Value = txtRef6.Text
        '    End If
        '    If txtRef7.Text <> "*Varies*" And My.Settings.Custom = "Drawing" Then
        '        oDoc.PropertySets.Item("Inventor User Defined Properties").Item("Reference7").Value = txtRef7.Text
        '    End If
        '    If txtRef8.Text <> "*Varies*" And My.Settings.Custom = "Drawing" Then
        '        oDoc.PropertySets.Item("Inventor User Defined Properties").Item("Reference8").Value = txtRef8.Text
        '    End If
        '    If txtRef9.Text <> "*Varies*" And My.Settings.Custom = "Drawing" Then
        '        oDoc.PropertySets.Item("Inventor User Defined Properties").Item("Reference9").Value = txtRef9.Text
        '    End If
        'Catch
        'End Try
    End Sub
    Private Sub WriteRevTable(ByRef oDoc As Document, Opendocs As ArrayList)
        Dim shtRevTable As RevisionTable
        'RevTable.PopiProperties(Me)
        'Dim Path As Documents
        'Path = _invApp.Documents
        Dim strPath, strFile, Rev As String
        Dim Sheet As Sheet
        Dim Col, Row, I, J As Integer
        Dim oPoint As Point2d = _invApp.TransientGeometry.CreatePoint2d(0.965193999999989, 2.7178)
        'For X = 0 To Main.lstSubfiles.Items.Count - 1
        'If Main.lstSubfiles.GetItemCheckState(X) = CheckState.Checked Then
        'For J = 1 To Path.Count
        'oDoc = Path.Item(J)
        _invApp.SilentOperation = True
        strPath = Strings.Left(oDoc.FullDocumentName, Len(oDoc.FullDocumentName) - 3) & "idw"
        strFile = Strings.Right(strPath, Strings.Len(strPath) - InStrRev(strPath, "\"))
        'If Trim(Main.lstSubfiles.Items.Item(X)) = strFile Then
        oDoc = _invApp.Documents.Open(strPath, True)
        Sheet = oDoc.ActiveSheet
        On Error Resume Next
        shtRevTable = Sheet.RevisionTables(1)
        Dim oRevTable As RevisionTable
        If Err.Number = 5 Then
            oRevTable = oDoc.ActiveSheet.Revisiontables.Add2(oPoint, False, True, False, 0)
            Err.Clear()
        End If
        Col = shtRevTable.RevisionTableColumns.Count
        Row = shtRevTable.RevisionTableRows.Count
        Rev = oDoc.PropertySets.Item("{F29F85E0-4FF9-1068-AB91-08002B27B3D9}").ItemByPropId("9").Value
        'Iterate through rev table and populate from the userform
        Dim Contents(0 To Col * Row) As String
        I = 0 : J = 0
        Dim Initials(0 To Row), RevCheckBy(0 To Row), RevDate(0 To Row) As String
        'For Each RevRow As RevisionTableRow In shtRevTable.RevisionTableRows
        '    Dim RtCell As RevisionTableCell
        '    For Each RtCell In RevRow
        '        'skip over items not used and proceed to next row
        '        If J Mod 5 = 0 And J > 0 Then
        '            J = 0
        '            I += 1
        '        End If
        '        If I <shtRevTable.RevisionTableRows.Count Then
        '            If J > 0 Then
        '                If RevTable.lstCheckFiles.Items(I).SubItems(J).Text <> "*Varies*" Then
        '                    RtCell.Text = RevTable.lstCheckFiles.Items(I).SubItems(J).Text
        '                End If
        '                J += 1
        '            Else
        '                If Strings.Len(RevTable.lstCheckFiles.Items(I).SubItems(J).Text) = 1 Then
        '                    J += 1
        '                End If
        '            End If
        '        End If
        '    Next
        'Next
        ' Adding Revs
        Dim RevTab As Integer = RevisionTabs.TabCount
        'If RevisionTabs.TabPages.Item(RevisionTabs.TabPages.Count - 1).Controls.Item("txtrev" & RevTab.ToString & "Init").Enabled = True Then
        '    shtRevTable.RevisionTableRows.Add()
        '    I = 0 : J = 0
        '    For Each revrow As RevisionTableRow In shtRevTable.RevisionTableRows
        '        Dim RtCell As RevisionTableCell
        '        For Each RtCell In revrow
        '            ' I = shtRevTable.RevisionTableColumns.Count * shtRevTable.RevisionTableRows.Count - 3 _
        '            'To shtRevTable.RevisionTableColumns.Count * shtRevTable.RevisionTableRows.Count
        '            If I = shtRevTable.RevisionTableRows.Count - 1 Then
        '                Select Case J
        '                    Case = shtRevTable.RevisionTableColumns.Count * shtRevTable.RevisionTableRows.Count - 3
        '                        RtCell.Text = RevisionTabs.TabPages.Item(RevisionTabs.TabPages.Count - 1).Controls.Item("txtRev" & RevTab.ToString & "Des").Text
        '                    Case = shtRevTable.RevisionTableColumns.Count * shtRevTable.RevisionTableRows.Count - 2
        '                        RtCell.Text = RevisionTabs.TabPages.Item(RevisionTabs.TabPages.Count - 1).Controls.Item("txtRev" & RevTab.ToString & "Init").Text
        '                    Case = shtRevTable.RevisionTableColumns.Count * shtRevTable.RevisionTableRows.Count - 1
        '                        RtCell.Text = RevisionTabs.TabPages.Item(RevisionTabs.TabPages.Count - 1).Controls.Item("txtRev" & RevTab.ToString & "Approved").Text
        '                End Select
        '            End If
        '            'Next
        '            J += 1
        '        Next
        '        I += 1
        '    Next
        'End If
        oDoc.Update()
        oDoc.Save()
        Main.CloseLater(strFile, oDoc)
        _invApp.SilentOperation = False
        '   Exit For
        'End If
        'Next
        'End If
        'Next
    End Sub
    Private Sub dgvCheckNeeded_CellClick(sender As Object, e As DataGridViewCellEventArgs) Handles dgvSummary.CellClick

        Dim DateCol As Boolean = False
        If e.RowIndex = -1 Or e.ColumnIndex = -1 Then Exit Sub
        'If dgvSummary.CurrentCellAddress.X = 1 And dgvSummary.CurrentCellAddress.Y = 8 Then
        '    dgvSummary(e.ColumnIndex, e.RowIndex) = MaterialCell
        '    '    Dim oRectangle As Rectangle = dgvSummary.GetCellDisplayRectangle(e.ColumnIndex, e.RowIndex, True)
        '    '    'Setting area for DateTimePicker Control
        '    '    'MaterialCell.Size = New Size(oRectangle.Width, oRectangle.Height)
        '    '    ' Setting Location
        '    '    'cmbSummMat.Location = New Drawing.Point(oRectangle.X, oRectangle.Y)
        'End If

        'If dgvSummary.CurrentCellAddress = dgvSummary(dgvSummary.Columns.Item("Item"), dgvSummary.Rows.Item("Material")) Then

        'End If

        'If DateCol = True Then

        '    'Adding DateTimePicker control into DataGridView 
        '    dgvCheckNeeded.Controls.Add(oDateTimePicker)
        '    ' Setting the format (i.e. 2014-10-10)
        '    oDateTimePicker.Format = DateTimePickerFormat.Short
        '    oDateTimePicker.ShowCheckBox = True
        '    ' It returns the retangular area that represents the Display area for a cell
        '    Dim oRectangle As Rectangle = dgvCheckNeeded.GetCellDisplayRectangle(e.ColumnIndex, e.RowIndex, True)
        '    'Setting area for DateTimePicker Control
        '    oDateTimePicker.Size = New Size(oRectangle.Width, oRectangle.Height)
        '    ' Setting Location
        '    oDateTimePicker.Location = New Drawing.Point(oRectangle.X, oRectangle.Y)
        '    ' An event attached to dateTimePicker Control which is fired when DateTimeControl is closed
        '    AddHandler oDateTimePicker.CloseUp, AddressOf Me.oDateTimePicker_CloseUp
        '    ' An event attached to dateTimePicker Control which is fired when any date is selected
        '    AddHandler oDateTimePicker.TextChanged, AddressOf Me.dateTimePicker_OnTextChange
        '    ' Now make it visible
        '    oDateTimePicker.Visible = True
        'Else
        '    oDateTimePicker.Visible = False
        'End If
    End Sub

    Private Sub dgvSummary_CellValueChanged(sender As Object, e As DataGridViewCellEventArgs) Handles dgvSummary.CellValueChanged
        If Me.Created Then dgvSummary(dgvSummary.Columns("SummaryIsDirty").Index, e.RowIndex).Value = "True"
    End Sub

    Private Sub dgvStatus_CellValueChanged(sender As Object, e As DataGridViewCellEventArgs) Handles dgvStatus.CellValueChanged
        If Me.Created Then dgvStatus(dgvStatus.Columns("StatusIsDirty").Index, e.RowIndex).Value = "True"
    End Sub

    Private Sub dgvProject_CellValueChanged(sender As Object, e As DataGridViewCellEventArgs) Handles dgvProject.CellValueChanged
        If Me.Created Then dgvProject(dgvProject.Columns("ProjectIsDirty").Index, e.RowIndex).Value = "True"
    End Sub
    Private Sub dgvRev1_CellValueChanged(sender As Object, e As DataGridViewCellEventArgs) Handles dgvRev1.CellValueChanged
        If Me.Created Then dgvRev1(dgvRev1.Columns("Rev1IsDirty").Index, e.RowIndex).Value = "True"
    End Sub
    Private Sub RevCellChange(Sender As Object, e As DataGridViewCellEventArgs)
        If Me.Created Then
            For Each control In RevisionTabs.SelectedTab.Controls
                If Strings.InStr(control.name, "Rev") <> 0 Then
                    dgvRevControl = control
                    Exit For
                End If
            Next
            dgvRevControl(dgvRevControl.Columns("Rev" & RevisionTabs.SelectedTab.TabIndex + 1 & "IsDirty").Index, (e.RowIndex)).Value = True
            Exit Sub
        End If
    End Sub
    Private Sub dgvCustomModel_CellValueChanged(sender As Object, e As DataGridViewCellEventArgs) Handles dgvCustomModel.CellValueChanged
        If Me.Created Then dgvCustomModel(dgvCustomModel.Columns("ModelIsDirty").Index, e.RowIndex).Value = "True"
        If Me.Created Then
            If dgvCustomModel(dgvCustomModel.Columns("PCusType").Index, e.RowIndex).Value = "Yes or No" AndAlso
                dgvCustomModel(dgvCustomModel.Columns("PCusValue").Index, e.RowIndex).GetType.ToString <> "System.Windows.Forms.DataGridViewComboBoxCell" Then
                dgvCustomModel(dgvCustomModel.Columns("PCusValue").Index, e.RowIndex) = New DataGridViewComboBoxCell
                cmbYesNo = dgvCustomModel(dgvCustomModel.Columns("PCusValue").Index, e.RowIndex)
                cmbYesNo.Items.AddRange("Yes", "No")
                Exit Sub
                cmbYesNo.Value = "Yes"
            ElseIf dgvCustomModel(dgvCustomModel.Columns("PCusType").Index, e.RowIndex).Value = "Text" Then
                Dim TextCell As DataGridViewTextBoxCell
                dgvCustomModel(dgvCustomModel.Columns("PCusValue").Index, e.RowIndex) = New DataGridViewTextBoxCell
                TextCell = dgvCustomModel(dgvCustomModel.Columns("PCusValue").Index, e.RowIndex)
            ElseIf dgvCustomModel(dgvCustomModel.Columns("PCusType").Index, e.RowIndex).Value = "Number" Then
                Dim strSource As String = Numeric(dgvCustomModel(dgvCustomModel.Columns("PCusValue").Index, e.RowIndex).Value)
                If dgvCustomModel(dgvCustomModel.Columns("PCusValue").Index, e.RowIndex).Value <> strSource Then dgvCustomModel(dgvCustomModel.Columns("PCusValue").Index, e.RowIndex).Value = strSource
            End If
        End If
    End Sub
    Public Function Numeric(ByVal StrSource As String) As String
        Dim strResult As String = ""
        Dim Dec As Boolean = False
        For i As Integer = 1 To Strings.Len(StrSource)
            If 45 < Asc(Mid(StrSource, i, 1)) AndAlso Asc(Mid(StrSource, i, 1)) < 58 Then
                If Asc(Mid(StrSource, i, 1)) = 46 And Dec = False Then
                    Dec = True
                    strResult = strResult & Mid(StrSource, i, 1)
                ElseIf Asc(Mid(StrSource, i, 1)) <> 46 Then
                    strResult = strResult & Mid(StrSource, i, 1)
                End If
            End If
        Next
        If strResult <> "" Then Numeric = strResult
    End Function
    Private Sub dgvCustomDrawing_CellValueChanged(sender As Object, e As DataGridViewCellEventArgs)
        If Me.Created Then dgvCustomDrawing(dgvCustomDrawing.Columns("DrawingIsDirty").Index, e.RowIndex).Value = "True"
    End Sub
#End Region

#Region "Combobox 1-Click"
    'Private Sub dgvStatus_CellMouseClick(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellMouseEventArgs) Handles dgvStatus.CellMouseClick
    '    'Make sure the type is set to something, a work around if the grid contains checkboxes.
    '    If e.RowIndex > -1 And e.ColumnIndex > -1 Then
    '        If dgvStatus(e.ColumnIndex, e.RowIndex).EditType IsNot Nothing Then
    '            'Check if the control is a combo box if so, edit on enter
    '            If dgvStatus(e.ColumnIndex, e.RowIndex).EditType.ToString() = "System.Windows.Forms.DataGridViewComboBoxEditingControl" Then
    '                SendKeys.Send("{F4}")
    '            End If
    '        End If
    '    End If
    'End Sub
    'Private Sub dgvSummary_CellMouseClick(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellMouseEventArgs) Handles dgvSummary.CellMouseClick
    '    'Make sure the type is set to something, a work around if the grid contains checkboxes.
    '    If e.RowIndex > -1 And e.ColumnIndex > -1 Then
    '        'Debug.WriteLine(dgvSummary(e.ColumnIndex, e.RowIndex))
    '        If dgvSummary(e.ColumnIndex, e.RowIndex).EditType IsNot Nothing Then
    '            'Check if the control is a combo box if so, edit on enter
    '            If dgvSummary(e.ColumnIndex, e.RowIndex).EditType.ToString() = "System.Windows.Forms.DataGridViewComboBoxEditingControl" Then
    '                SendKeys.Send("{F4}")
    '            End If
    '        End If
    '    End If
    'End Sub
    'Private Sub dgvCustomModel_CellMouseClick(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellMouseEventArgs) Handles dgvCustomModel.CellMouseClick
    '    'Make sure the type is set to something, a work around if the grid contains checkboxes.
    '    If e.RowIndex > -1 And e.ColumnIndex > -1 Then
    '        If dgvCustomModel(e.ColumnIndex, e.RowIndex).EditType IsNot Nothing Then
    '            'Check if the control is a combo box if so, edit on enter
    '            If dgvCustomModel(e.ColumnIndex, e.RowIndex).EditType.ToString() = "System.Windows.Forms.DataGridViewComboBoxEditingControl" Then
    '                SendKeys.Send("{F4}")
    '            End If
    '        End If
    '    End If
    'End Sub
    'Private Sub dgvCustomDrawing_CellMouseClick(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellMouseEventArgs) Handles dgvCustomDrawing.CellMouseClick
    '    'Make sure the type is set to something, a work around if the grid contains checkboxes.
    '    If e.RowIndex > -1 And e.ColumnIndex > -1 Then
    '        If dgvCustomDrawing(e.ColumnIndex, e.RowIndex).EditType IsNot Nothing Then
    '            'Check if the control is a combo box if so, edit on enter
    '            If dgvCustomDrawing(e.ColumnIndex, e.RowIndex).EditType.ToString() = "System.Windows.Forms.DataGridViewComboBoxEditingControl" Then
    '                SendKeys.Send("{F4}")
    '            End If
    '        End If
    '    End If
    'End Sub
#End Region
End Class
Public Class Items
    Public Property ModelValue As String
    Public Property DrawingValue As String
    Public Property Type As Object
End Class