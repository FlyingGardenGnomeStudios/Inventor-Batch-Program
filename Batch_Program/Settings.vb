Imports System.Windows.Forms
Imports System.Text.RegularExpressions
#Region "Export Settings"
Public Class Settings
    Private RowIndex As Integer = 0
    Public Sub New()
        InitializeComponent()
        LoadFormValues()
        LoadGeneralsettings()
        LoadRevTableSettings()
    End Sub
    Private Sub LoadFormValues()
        If My.Settings.PDFSaveNewLoc = True Then
            txtPDFSaveLoc.Text = My.Settings.PDFSaveLoc
            txtPDFSaveLoc.Enabled = True
            rdoPDFSaveLoc.Checked = True
        ElseIf My.Settings.PDFSaveTag = True Then
            txtPDFTag.Enabled = True
            rdoPDFTag.Checked = True
            txtPDFTag.Text = My.Settings.PDFTag
        Else
            rdoPDFChoose.Checked = True
        End If
        If My.Settings.DXFSaveNewLoc = True Then
            txtDXFSaveLoc.Text = My.Settings.DXFSaveLoc
            txtDXFSaveLoc.Enabled = True
            rdoDXFSaveLoc.Checked = True
        ElseIf My.Settings.DXFSaveTag = True Then
            txtDXFTag.Enabled = True
            rdoDXFTag.Checked = True
            txtDXFTag.Text = My.Settings.DXFTag
        Else
            rdoDXFChoose.Checked = True
        End If
        If My.Settings.DWGSaveNewLoc = True Then
            txtDWGSaveLoc.Text = My.Settings.DWGSaveLoc
            txtDWGSaveLoc.Enabled = True
            rdoDWGSaveLoc.Checked = True
        ElseIf My.Settings.DWGSaveTag = True Then
            txtDWGTag.Enabled = True
            rdoDWGTag.Checked = True
            txtDWGTag.Text = My.Settings.DWGTag
        Else
            rdoDWGChoose.Checked = True
        End If
        If My.Settings.PDFRev = False Then
            chkPDFRev.Checked = False
        Else
            chkPDFRev.Checked = True
        End If
        If My.Settings.DXFRev = False Then
            chkDXFRev.Checked = False
        Else
            chkDXFRev.Checked = True
        End If
        If My.Settings.DWGRev = False Then
            chkDWGRev.Checked = False
        Else
            chkDWGRev.Checked = True
        End If
        If My.Settings.ArchiveExport = False Then
            chkArchive.Checked = False
        Else
            chkArchive.Checked = True
        End If
        If My.Settings.DWGini = True Then
            chkCustDWGini.Checked = True
            txtCustDWGini.Text = My.Settings.DWGiniLoc
        End If
        If My.Settings.DXFini = True Then
            chkCustDXFini.Checked = True
            txtCustDXFini.Text = My.Settings.DXFiniLoc
        End If
        If My.Settings.PDFLineWeights = True Then
            chkLineWeights.Checked = True
        Else
            chkLineWeights.Checked = False
        End If
        If My.Settings.PDFColoursBlack = True Then
            chkPDFBW.Checked = True
        Else
            chkPDFBW.Checked = False
        End If
        cmbSheets.SelectedIndex = My.Settings.PDFRange
        numRes.Value = My.Settings.PDFRes
    End Sub
    Private Sub rdoPDFSaveLoc_CheckedChanged(sender As Object, e As EventArgs) Handles rdoPDFSaveLoc.CheckedChanged
        If rdoPDFSaveLoc.Checked = True Then
            rdoPDFChoose.Checked = False
            rdoPDFTag.Checked = False
            txtPDFSaveLoc.Enabled = True
            btnPDFLocBrowse.Enabled = True
            txtPDFTag.Enabled = False
            If txtPDFTag.Text = "" Then txtPDFTag.Text = "ex: \Drawings\PDF"
        End If
    End Sub
    Private Sub rdoPDFChoose_CheckedChanged(sender As Object, e As EventArgs) Handles rdoPDFChoose.CheckedChanged
        If rdoPDFChoose.Checked = True Then
            rdoPDFTag.Checked = False
            rdoPDFSaveLoc.Checked = False
            txtPDFSaveLoc.Enabled = False
            btnPDFLocBrowse.Enabled = False
            txtPDFTag.Enabled = False
            If txtPDFTag.Text = "" Then txtPDFTag.Text = "ex: \Drawings\PDF"
        End If
    End Sub
    Private Sub rdoPDFSub_CheckedChanged(sender As Object, e As EventArgs) Handles rdoPDFTag.CheckedChanged
        If rdoPDFTag.Checked = True Then
            rdoPDFChoose.Checked = False
            rdoPDFSaveLoc.Checked = False
            txtPDFSaveLoc.Enabled = False
            txtPDFTag.Enabled = True
            If txtPDFTag.Text.Contains("ex:") Then txtPDFTag.Text = ""
        End If
    End Sub
    Private Sub rdoDXFChoose_CheckedChanged(sender As Object, e As EventArgs) Handles rdoDXFChoose.CheckedChanged
        If rdoDXFChoose.Checked = True Then
            rdoDXFSaveLoc.Checked = False
            rdoDXFTag.Checked = False
            txtDXFTag.Enabled = False
            txtDXFSaveLoc.Enabled = False
            If txtDXFTag.Text = "" Then txtDXFTag.Text = "ex: \Drawings\DXF"
        End If
    End Sub
    Private Sub rdoDXFSaveLoc_CheckedChanged(sender As Object, e As EventArgs) Handles rdoDXFSaveLoc.CheckedChanged
        If rdoDXFSaveLoc.Checked = True Then
            rdoDXFChoose.Checked = False
            rdoDXFTag.Checked = False
            txtDXFSaveLoc.Enabled = True
            txtDXFTag.Enabled = False
            If txtDXFTag.Text = "" Then txtDXFTag.Text = "ex: \Drawings\DXF"
        End If
    End Sub
    Private Sub rdoDXFSub_CheckedChanged(sender As Object, e As EventArgs) Handles rdoDXFTag.CheckedChanged
        If rdoDXFTag.Checked = True Then
            rdoDXFChoose.Checked = False
            rdoDXFSaveLoc.Checked = False
            txtDXFSaveLoc.Enabled = False
            txtDXFTag.Enabled = True
            If txtDXFTag.Text.Contains("ex:") Then txtDXFTag.Text = ""
        End If
    End Sub
    Private Sub rdoDWGChoose_CheckedChanged(sender As Object, e As EventArgs) Handles rdoDWGChoose.CheckedChanged
        If rdoDWGChoose.Checked = True Then
            rdoDWGSaveLoc.Checked = False
            rdoDWGTag.Checked = False
            txtDWGTag.Enabled = False
            txtDWGSaveLoc.Enabled = False
            If txtDWGTag.Text = "" Then txtDWGTag.Text = "ex: \Drawings\DWG"
        End If
    End Sub
    Private Sub rdoDWGSaveLoc_CheckedChanged(sender As Object, e As EventArgs) Handles rdoDWGSaveLoc.CheckedChanged
        If rdoDWGSaveLoc.Checked = True Then
            rdoDWGChoose.Checked = False
            rdoDWGTag.Checked = False
            txtDWGSaveLoc.Enabled = True
            txtDWGTag.Enabled = False
            If txtDWGTag.Text = "" Then txtDWGTag.Text = "ex: \Drawings\DWG"
        End If
    End Sub
    Private Sub rdoDWGSub_CheckedChanged(sender As Object, e As EventArgs) Handles rdoDWGTag.CheckedChanged
        If rdoDWGTag.Checked = True Then
            rdoDWGChoose.Checked = False
            rdoDWGSaveLoc.Checked = False
            txtDWGSaveLoc.Enabled = False
            txtDWGTag.Enabled = True
            If txtDWGTag.Text.Contains("ex:") Then txtDWGTag.Text = ""
        End If
    End Sub
    Private Sub btnCancel_Click(sender As Object, e As EventArgs) Handles btnCancel.Click
        LoadFormValues()
        Me.Close()
    End Sub
    Private Sub btnOK_Click(sender As Object, e As EventArgs) Handles btnOK.Click


        If rdoPDFSaveLoc.Checked = True Then
            My.Settings.PDFSaveLoc = txtPDFSaveLoc.Text
            My.Settings.PDFSaveNewLoc = True
            My.Settings.PDFSaveTag = False
        ElseIf rdoPDFTag.Checked = True Then
            If Regex.IsMatch(txtPDFTag.Text, ("^[a-zA-Z0-9_]*$")) Then
                MsgBox("The PDF location you have selected Is invalid")
                Exit Sub
            End If
            If Strings.Left(txtPDFTag.Text, 1) <> "\" Then
                txtPDFTag.Text = "\" & txtPDFTag.Text
            End If
            If Strings.Right(txtPDFTag.Text, 1) = "\" Then
                txtPDFTag.Text = Strings.Left(txtPDFTag.Text, Len(txtPDFTag.Text) - 1)
            End If
            My.Settings.PDFTag = txtPDFTag.Text
            My.Settings.PDFSaveNewLoc = False
            My.Settings.PDFSaveTag = True
        Else
            My.Settings.PDFSaveNewLoc = False
            My.Settings.PDFSaveTag = False
        End If
        If chkPDFRev.Checked = True Then
            My.Settings.PDFRev = True
        Else
            My.Settings.PDFRev = False
        End If


        If rdoDXFSaveLoc.Checked = True Then
            My.Settings.DXFSaveLoc = txtDXFSaveLoc.Text
            My.Settings.DXFSaveNewLoc = True
            My.Settings.DXFSaveTag = False
        ElseIf rdoDXFTag.Checked = True Then
            If Regex.IsMatch(txtDXFTag.Text, ("^[a-zA-Z0-9_]*$")) Then
                MsgBox("The DXF location you have selected Is invalid")
                Exit Sub
            End If
            If Strings.Left(txtDXFTag.Text, 1) <> "\" Then
                txtDXFTag.Text = "\" & txtDXFTag.Text
            End If
            If Strings.Right(txtDXFTag.Text, 1) = "\" Then
                txtDXFTag.Text = Strings.Left(txtDXFTag.Text, Len(txtDXFTag.Text) - 1)
            End If
            My.Settings.DXFTag = txtDXFTag.Text
            My.Settings.DXFSaveNewLoc = False
            My.Settings.DXFSaveTag = True
        Else
            My.Settings.DXFSaveNewLoc = False
            My.Settings.DXFSaveTag = False
        End If
        If chkDXFRev.Checked = True Then
            My.Settings.DXFRev = True
        Else
            My.Settings.DXFRev = False
        End If

        If rdoDWGSaveLoc.Checked = True Then
            My.Settings.DWGSaveLoc = txtDWGSaveLoc.Text
            My.Settings.DWGSaveNewLoc = True
            My.Settings.DWGSaveTag = False
        ElseIf rdoDWGTag.Checked = True Then
            If Regex.IsMatch(txtDWGTag.Text, ("^[a-zA-Z0-9_]*$")) Then
                MsgBox("The DWG location you have selected Is invalid")
                Exit Sub
            End If
            If Strings.Left(txtDWGTag.Text, 1) <> "\" Then
                txtDWGTag.Text = "\" & txtDWGTag.Text
            End If
            If Strings.Right(txtDWGTag.Text, 1) = "\" Then
                txtDWGTag.Text = Strings.Left(txtDWGTag.Text, Len(txtDWGTag.Text) - 1)
            End If
            My.Settings.DWGTag = txtDWGTag.Text
            My.Settings.DWGSaveNewLoc = False
            My.Settings.DWGSaveTag = True
        Else
            My.Settings.DWGSaveNewLoc = False
            My.Settings.DWGSaveTag = False
        End If
        If chkDWGRev.Checked = True Then
            My.Settings.DWGRev = True
        Else
            My.Settings.DWGRev = False
        End If
        If chkArchive.Checked = True Then
            My.Settings.ArchiveExport = True
        Else
            My.Settings.ArchiveExport = False
        End If
        If chkCustDWGini.Checked = True Then
            My.Settings.DWGini = True
            My.Settings.DWGiniLoc = txtCustDWGini.Text
        Else
            My.Settings.DWGini = False
        End If
        If chkCustDXFini.Checked = True Then
            My.Settings.DXFini = True
            My.Settings.DXFiniLoc = txtCustDXFini.Text
        Else
            My.Settings.DXFini = False
        End If
        If chkLineWeights.Checked = True Then
            My.Settings.PDFLineWeights = True
        Else
            My.Settings.PDFLineWeights = False
        End If
        If chkPDFBW.Checked Then
            My.Settings.PDFColoursBlack = True
        Else
            My.Settings.PDFColoursBlack = False
        End If
        My.Settings.PDFRange = cmbSheets.SelectedIndex
        My.Settings.PDFRes = numRes.Value
        My.Settings.Save()
        Me.Close()
        If rdoStrict.Checked = True Then
            My.Settings.StrictSearch = True
        Else
            My.Settings.StrictSearch = False
        End If
        My.Settings.Save()
        Me.Close()
    End Sub
    Private Sub btnPDFLocBrowse_Click(sender As Object, e As EventArgs) Handles btnPDFLocBrowse.Click
        Dim Folder As FolderBrowserDialog = New FolderBrowserDialog
        Folder.Description = "Choose the location you wish to save To"
        Folder.RootFolder = System.Environment.SpecialFolder.Desktop
        Try
            If Folder.ShowDialog() = Windows.Forms.DialogResult.OK Then
                txtPDFSaveLoc.Text = Folder.SelectedPath
            Else
                Exit Sub
            End If
        Catch ex As Exception
            MessageBox.Show(ex.Message, "Exception Details", MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
        End Try
    End Sub
    Private Sub btnDXFLocBrowse_Click(sender As Object, e As EventArgs) Handles btnDXFLocBrowse.Click
        Dim Folder As FolderBrowserDialog = New FolderBrowserDialog
        Folder.Description = "Choose the location you wish to save To"
        Folder.RootFolder = System.Environment.SpecialFolder.Desktop
        Try
            If Folder.ShowDialog() = Windows.Forms.DialogResult.OK Then
                txtDXFSaveLoc.Text = Folder.SelectedPath
            Else
                Exit Sub
            End If
        Catch ex As Exception
            MessageBox.Show(ex.Message, "Exception Details", MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
        End Try
    End Sub
    Private Sub btnDWGBrowse_Click(sender As Object, e As EventArgs) Handles btnDWGBrowse.Click
        Dim Folder As FolderBrowserDialog = New FolderBrowserDialog
        Folder.Description = "Choose the location you wish To save To"
        Folder.RootFolder = System.Environment.SpecialFolder.Desktop
        Try
            If Folder.ShowDialog() = Windows.Forms.DialogResult.OK Then
                txtDWGSaveLoc.Text = Folder.SelectedPath
            Else
                Exit Sub
            End If
        Catch ex As Exception
            MessageBox.Show(ex.Message, "Exception Details", MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
        End Try
    End Sub
    Private Function IsAlphaNum(ByVal strInputText As String) As Boolean
        Dim IsAlpha As Boolean = False
        If System.Text.RegularExpressions.Regex.IsMatch(strInputText, "^[a-zA-Z0-9-\\]+$") Then
            IsAlpha = True
        Else
            IsAlpha = False
        End If
        Return IsAlpha
    End Function
    Private Sub chkCustDWGini_CheckedChanged(sender As Object, e As EventArgs) Handles chkCustDWGini.CheckedChanged
        If chkCustDWGini.Checked = True Then
            txtCustDWGini.Enabled = True
            btnDWGiniBrowse.Enabled = True
        Else
            txtCustDWGini.Enabled = False
            btnDWGiniBrowse.Enabled = False
        End If
    End Sub
    Private Sub btnDWGiniBrowse_Click(sender As Object, e As EventArgs) Handles btnDWGiniBrowse.Click
        Dim Folder As OpenFileDialog = New OpenFileDialog
        Folder.Title = "Choose the location of the .ini file"
        Folder.InitialDirectory = System.Environment.SpecialFolder.Desktop
        Folder.Filter = ".ini Files (*.ini*)|*.ini*"
        Folder.RestoreDirectory = True
        Try
            If Folder.ShowDialog() = Windows.Forms.DialogResult.OK Then
                txtCustDWGini.Text = Folder.FileName
            Else
                Exit Sub
            End If
        Catch ex As Exception
            MessageBox.Show(ex.Message, "Exception Details", MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
        End Try
    End Sub
    Private Sub chkDXFCustini_CheckedChanged(sender As Object, e As EventArgs) Handles chkCustDXFini.CheckedChanged
        If chkCustDXFini.Checked = True Then
            txtCustDXFini.Enabled = True
            btnDXFiniBrowse.Enabled = True
        Else
            txtCustDXFini.Enabled = False
            btnDXFiniBrowse.Enabled = False
        End If
    End Sub
    Private Sub btnDXFiniBrowse_Click(sender As Object, e As EventArgs) Handles btnDXFiniBrowse.Click
        Dim Folder As OpenFileDialog = New OpenFileDialog
        Folder.Title = "Choose the location of the .ini file"
        Folder.InitialDirectory = System.Environment.SpecialFolder.Desktop
        Folder.Filter = ".ini Files (*.ini*)|*.ini*"
        Folder.RestoreDirectory = True
        Try
            If Folder.ShowDialog() = Windows.Forms.DialogResult.OK Then
                txtCustDXFini.Text = Folder.FileName
            Else
                Exit Sub
            End If
        Catch ex As Exception
            MessageBox.Show(ex.Message, "Exception Details", MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
        End Try
    End Sub
    Private Sub RefColour_ItemActivate(sender As Object, e As EventArgs) Handles RefColour.ItemActivate
        Dim cDialog As New ColorDialog()
        Dim Index As Integer = RefColour.FocusedItem.Index
        cDialog.Color = RefColour.Items.Item(Index).ForeColor ' initial selection is current color.

        If (cDialog.ShowDialog() = DialogResult.OK) Then
            RefColour.Items.Item(Index).ForeColor = cDialog.Color ' update with user selected color.
        End If
        Select Case Index
            Case 0
                My.Settings.REF = cDialog.Color
            Case 1
                My.Settings.DNE = cDialog.Color
            Case 2
                My.Settings.PPM = cDialog.Color
        End Select

        RefColour.SelectedItems.Clear()
    End Sub
    Private Sub LoadGeneralsettings()
        RefColour.Items.Item(0).ForeColor = My.Settings.REF
        RefColour.Items.Item(1).ForeColor = My.Settings.DNE
        RefColour.Items.Item(2).ForeColor = My.Settings.PPM
        If My.Settings.StrictSearch = False Then
            rdoLoose.Checked = True
            rdoStrict.Checked = False
        Else
            rdoStrict.Checked = True
            rdoLoose.Checked = False
        End If
    End Sub
    Public Sub LoadRevTableSettings()

        dgvRevTableLayout.Rows.Add(My.Settings.RTSCheckedBy, "Checked By", "Sheet iProperty", "Text")
        dgvRevTableLayout.Rows.Add(My.Settings.RTSCheckedDate, "Check Date", "Sheet iProperty", "Date")
        dgvRevTableLayout.Rows.Add(My.Settings.RTSRev, "Revision Number", My.Settings.RTSRevCol, "Number")
        dgvRevTableLayout.Rows.Add(My.Settings.RTSDate, "Revision Date", My.Settings.RTSDateCol, "Date")
        dgvRevTableLayout.Rows.Add(My.Settings.RTSDesc, "Revision Description", My.Settings.RTSDescCol, "Text")
        dgvRevTableLayout.Rows.Add(My.Settings.RTSName, "Revision By", My.Settings.RTSNameCol, "Text")
        dgvRevTableLayout.Rows.Add(My.Settings.RTSApproved, "Revision Approved", My.Settings.RTSApprovedCol, "Text")
        For column = 0 To dgvRevTableLayout.Columns.Count - 1
            For Row = 0 To 1
                If column <> 0 Then
                    dgvRevTableLayout(column, Row).ReadOnly = True
                    dgvRevTableLayout(column, Row).Style.BackColor = Drawing.Color.LightGray
                End If
            Next
            For row = 2 To dgvRevTableLayout.Rows.Count - 2
                If column <> 0 AndAlso column <> 2 Then
                    dgvRevTableLayout(column, row).ReadOnly = True
                    dgvRevTableLayout(column, row).Style.BackColor = Drawing.Color.LightGray
                End If
            Next
        Next
        If My.Settings.RTS1Item <> "" Then dgvRevTableLayout.Rows.Add(My.Settings.RTS1, My.Settings.RTS1Item, My.Settings.RTS1Col, My.Settings.RTS1Value)
        If My.Settings.RTS2Item <> "" Then dgvRevTableLayout.Rows.Add(My.Settings.RTS2, My.Settings.RTS2Item, My.Settings.RTS2Col, My.Settings.RTS2Value)
        If My.Settings.RTS3Item <> "" Then dgvRevTableLayout.Rows.Add(My.Settings.RTS3, My.Settings.RTS3Item, My.Settings.RTS3Col, My.Settings.RTS3Value)
        If My.Settings.RTS4Item <> "" Then dgvRevTableLayout.Rows.Add(My.Settings.RTS4, My.Settings.RTS4Item, My.Settings.RTS4Col, My.Settings.RTS4Value)
        If My.Settings.RTS5Item <> "" Then dgvRevTableLayout.Rows.Add(My.Settings.RTS5, My.Settings.RTS5Item, My.Settings.RTS5Col, My.Settings.RTS5Value)
        txtNumRev.Text = My.Settings.NumRev
        txtAlphaRev.Text = My.Settings.AlphaRev
        nudStartVal.Value = My.Settings.StartVal
        Select Case My.Settings.DefRevLoc
            Case "BL"
                rdbBL.Checked = True
            Case "BR"
                rdbBR.Checked = True
            Case "TL"
                rdbTL.Checked = True
            Case "TR"
                rdbTR.Checked = True
        End Select
    End Sub
    Public Sub RevTableOK()
        For Row = 2 To 10
            Select Case Row - 2
                Case 0
                    My.Settings.RTSRev = dgvRevTableLayout(0, Row).Value
                    My.Settings.RTSRevCol = dgvRevTableLayout(2, Row).Value
                Case 1
                    My.Settings.RTSDate = dgvRevTableLayout(0, Row).Value
                    My.Settings.RTSDateCol = dgvRevTableLayout(2, Row).Value
                Case 2
                    My.Settings.RTSDesc = dgvRevTableLayout(0, Row).Value
                    My.Settings.RTSDescCol = dgvRevTableLayout(2, Row).Value
                Case 3
                    My.Settings.RTSName = dgvRevTableLayout(0, Row).Value
                    My.Settings.RTSNameCol = dgvRevTableLayout(2, Row).Value
                Case 4
                    My.Settings.RTSApproved = dgvRevTableLayout(0, Row).Value
                    My.Settings.RTSApprovedCol = dgvRevTableLayout(2, Row).Value
                Case 5
                    Try
                        My.Settings.RTS1 = dgvRevTableLayout(0, Row).Value
                        My.Settings.RTS1Item = dgvRevTableLayout(1, Row).Value
                        My.Settings.RTS1Col = dgvRevTableLayout(2, Row).Value
                        My.Settings.RTS1Value = dgvRevTableLayout(3, Row).Value
                    Catch
                        My.Settings.RTS1 = False
                        My.Settings.RTS1Item = Nothing
                        My.Settings.RTS1Col = Nothing
                        My.Settings.RTS1Value = Nothing
                    End Try
                Case 6
                    Try
                        My.Settings.RTS2 = dgvRevTableLayout(0, Row).Value
                        My.Settings.RTS2Item = dgvRevTableLayout(1, Row).Value
                        My.Settings.RTS2Col = dgvRevTableLayout(2, Row).Value
                        My.Settings.RTS2Value = dgvRevTableLayout(3, Row).Value
                    Catch
                        My.Settings.RTS2 = False
                        My.Settings.RTS2Item = Nothing
                        My.Settings.RTS2Col = Nothing
                        My.Settings.RTS2Value = Nothing
                    End Try
                Case 7
                    Try
                        My.Settings.RTS3 = dgvRevTableLayout(0, Row).Value
                        My.Settings.RTS3Item = dgvRevTableLayout(1, Row).Value
                        My.Settings.RTS3Col = dgvRevTableLayout(2, Row).Value
                        My.Settings.RTS3Value = dgvRevTableLayout(3, Row).Value
                    Catch
                        My.Settings.RTS3 = False
                        My.Settings.RTS3Item = Nothing
                        My.Settings.RTS3Col = Nothing
                        My.Settings.RTS3Value = Nothing
                    End Try
                Case 8
                    Try
                        My.Settings.RTS4 = dgvRevTableLayout(0, Row).Value
                        My.Settings.RTS4Item = dgvRevTableLayout(1, Row).Value
                        My.Settings.RTS4Col = dgvRevTableLayout(2, Row).Value
                        My.Settings.RTS4Value = dgvRevTableLayout(3, Row).Value
                    Catch
                        My.Settings.RTS4 = False
                        My.Settings.RTS4Item = Nothing
                        My.Settings.RTS4Col = Nothing
                        My.Settings.RTS4Value = Nothing
                    End Try
                Case 9
                    Try
                        My.Settings.RTS5 = dgvRevTableLayout(0, Row).Value
                        My.Settings.RTS5Item = dgvRevTableLayout(1, Row).Value
                        My.Settings.RTS5Col = dgvRevTableLayout(2, Row).Value
                        My.Settings.RTS5Value = dgvRevTableLayout(3, Row).Value
                    Catch
                        My.Settings.RTS5 = False
                        My.Settings.RTS5Item = Nothing
                        My.Settings.RTS5Col = Nothing
                        My.Settings.RTS5Value = Nothing
                    End Try
                Case 10
                    MsgBox("Batch Program currently only supports 5 custom fields")
            End Select
        Next
        If txtAlphaRev.Text <> "" Then
            My.Settings.AlphaRev = txtAlphaRev.Text
        Else
            MsgBox("Please enter a default description for alphabetical revisions")
            Exit Sub
        End If
        If txtNumRev.Text <> "" Then
            My.Settings.NumRev = txtNumRev.Text
        Else
            MsgBox("Please enter a default description for numerical revisions")
            Exit Sub
        End If
        My.Settings.StartVal = nudStartVal.Value
        My.Settings.Save()
        Me.Close()
        If rdbBL.Checked = True Then
            My.Settings.DefRevLoc = "BL"
        ElseIf rdbBR.Checked = True Then
            My.Settings.DefRevLoc = "BR"
        ElseIf rdbTL.Checked = True Then
            My.Settings.DefRevLoc = "TL"
        ElseIf rdbTR.Checked = True Then
            My.Settings.DefRevLoc = "TR"
        End If
    End Sub
    'Private Sub dgvRevTableLayout_CellMouseClick(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellMouseEventArgs) Handles dgvRevTableLayout.CellMouseClick
    '    'Make sure the type is set to something, a work around if the grid contains checkboxes.
    '    If e.RowIndex > -1 And e.ColumnIndex > -1 Then
    '        If dgvRevTableLayout(e.ColumnIndex, e.RowIndex).EditType IsNot Nothing Then
    '            'Check if the control is a combo box if so, edit on enter
    '            If dgvRevTableLayout(e.ColumnIndex, e.RowIndex).EditType.ToString() = "System.Windows.Forms.DataGridViewComboBoxEditingControl" Then
    '                'SendKeys.Send("{F4}")
    '            End If
    '        End If
    '    End If
    'End Sub
    Private Sub ContextMenuStrip1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ContextMenuStrip1.Click
        Try
            If dgvRevTableLayout.CurrentCell.RowIndex > 6 Then
                dgvRevTableLayout.Rows.RemoveAt(dgvRevTableLayout.CurrentCell.RowIndex)
            Else
                MsgBox("Cannot remove default item." & vbNewLine & "Uncheck this item if it does not apply")
            End If
        Catch
        End Try
    End Sub
    Private Sub btnApply_Click(sender As Object, e As EventArgs) Handles btnApply.Click
        My.Settings.Save()
    End Sub
End Class
#End Region