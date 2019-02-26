Imports System.Windows.Forms
Imports System.Text.RegularExpressions
Imports System.IO
Imports System.Data
Public Class Settings
    'Private RowIndex As Integer = 0
    Dim NameFormatDT As New DataTable
    Public Sub New()
        InitializeComponent()
        PrintLoad()
        LoadGeneralSettings()
        LoadRevTableSettings()
        ExportLoad()
    End Sub
#Region "General"
    Private Sub RefColour_ItemActivate(sender As Object, e As EventArgs) Handles lvRefColour.ItemActivate
        Dim cDialog As New ColorDialog()
        Dim Index As Integer = lvRefColour.FocusedItem.Index
        cDialog.Color = lvRefColour.Items.Item(Index).ForeColor ' initial selection is current color.

        If (cDialog.ShowDialog() = DialogResult.OK) Then
            lvRefColour.Items.Item(Index).ForeColor = cDialog.Color ' update with user selected color.
        End If
        Select Case Index
            Case 0
                My.Settings.REF = cDialog.Color
            Case 1
                My.Settings.DNE = cDialog.Color
            Case 2
                My.Settings.PPM = cDialog.Color
        End Select

        lvRefColour.SelectedItems.Clear()
    End Sub
    Private Sub txtNameFormat_Click(sender As Object, e As EventArgs) Handles txtNameFormat.Click
        txtNameFormat.ForeColor = Drawing.Color.Black
        If txtNameFormat.Text.Contains("ex: ") Then
            txtNameFormat.Text = ""
        End If
    End Sub
    Private Sub txtNameFormat_LostFocus(sender As Object, e As EventArgs) Handles txtNameFormat.LostFocus
        If txtNameFormat.Text = "" Then
            txtNameFormat.Text = "ex: 1234-12.ABC"
            txtNameFormat.ForeColor = Drawing.Color.Gray
        End If
    End Sub
    Private Sub btnAdd_Click(sender As Object, e As EventArgs) Handles btnAdd.Click
        Dim NameFormat As String = RemoveInvalidFileNameChars(txtNameFormat.Text)
        Dim Example As String = Nothing
        Dim Alpha As String = "@"
        Dim Num As Integer = "0"
        If NameFormat <> txtNameFormat.Text Then
            MsgBox("Invalid characters have been removed from the submitted filename")
            txtNameFormat.Text = NameFormat
        End If
        NameFormat = Nothing
        For Letter = 0 To Len(txtNameFormat.Text) - 1
            If Num > 9 Then Num = Num - 10
            If Chr(Asc(Alpha)) > Chr(Asc("A") + 25) Then Alpha = Chr(Asc(Alpha) - 25)
            If Char.IsNumber(txtNameFormat.Text.Chars(Letter)) Then ' IsNumeric(Strings.Left(txtNameFormat.Text, Letter)) = True Then
                NameFormat = NameFormat & "#" ' Replace(Letter, Strings.Left(NameFormat, Letter), "#")
                Example = Example & Num + 1
                Num += 1
            ElseIf Char.IsLetter(txtNameFormat.Text.Chars(Letter)) Then ' Regex.IsMatch(Strings.Left(UCase(txtNameFormat.Text), Letter), "^[A-Z]{1}$") Then
                NameFormat = NameFormat & "?" 'Replace(Letter, Strings.Left(NameFormat, Letter), "?")
                Example = Example & Chr(Asc(Alpha) + 1)
                Alpha = Chr(Asc(Alpha) + 1)
            Else
                NameFormat = NameFormat & txtNameFormat.Text.Chars(Letter)
                Example = Example & txtNameFormat.Text.Chars(Letter)
                Alpha = "@"
                Num = 0
            End If
        Next
        Dim Dup As Boolean = False
        For Each Row In dgvNameFormat.Rows
            If NameFormat = dgvNameFormat(dgvNameFormat.Columns("Format").Index, Row.index).Value Then
                Dup = True
                Exit For
            End If
        Next
        If Dup = False Then
            dgvNameFormat.Rows.Add({NameFormat, Example})
        End If
    End Sub
    Private Sub btnRemove_Click(sender As Object, e As EventArgs) Handles btnRemove.Click
        If dgvNameFormat.SelectedRows Is Nothing Then
            Exit Sub
        Else
            For Each Row In dgvNameFormat.SelectedRows
                dgvNameFormat.Rows.Remove(dgvNameFormat.Rows(Row.index))
            Next
        End If

    End Sub
    Private Sub LoadGeneralSettings()
        lvRefColour.Items.Item(0).ForeColor = My.Settings.REF
        lvRefColour.Items.Item(1).ForeColor = My.Settings.DNE
        lvRefColour.Items.Item(2).ForeColor = My.Settings.PPM
        If My.Settings.StrictSearch = True Then
            rdoLoose.Checked = False
            rdoStrict.Checked = True
        Else
            rdoStrict.Checked = False
            rdoLoose.Checked = True
        End If

        chkExperimental.Checked = My.Settings.Experimental
        chkColourCode.Checked = My.Settings.ColourCode
        If dgvNameFormat.RowCount = 0 Then
            If Not My.Settings.NameFormat Is Nothing Then
                For Each Item As String In My.Settings.NameFormat
                    Dim Items As String() = Strings.Split(Item, "*")
                    dgvNameFormat.Rows.Add({Items(0), Items(1)})
                Next
            End If
        End If
        If chkExperimental.Checked = True Then
            GroupBox5.Height = 238
        Else
            GroupBox5.Height = 39
        End If
        nudMaxRef.Value = My.Settings.MaxRefNum
        If chkColourCode.Checked = False Then
            lvRefColour.Enabled = False
        Else
            lvRefColour.Enabled = True
        End If
    End Sub
    Private Sub chkColourCode_CheckedChanged(sender As Object, e As EventArgs) Handles chkColourCode.CheckedChanged
        If chkColourCode.Checked = False Then
            lvRefColour.Enabled = False
        Else
            lvRefColour.Enabled = True
        End If

    End Sub
    Private Sub chkExperimental_CheckedChanged(sender As Object, e As EventArgs) Handles chkExperimental.CheckedChanged
        If chkExperimental.Checked = True Then
            GroupBox5.Height = 238
        Else
            GroupBox5.Height = 39
        End If
    End Sub
    Private Sub GeneralOK()
        My.Settings.StrictSearch = rdoStrict.Checked
        My.Settings.ColourCode = chkColourCode.Checked
        My.Settings.Experimental = chkExperimental.Checked
        Dim NameFormatDT As DataTable = New DataTable
        For Each col As DataGridViewColumn In dgvNameFormat.Columns
            NameFormatDT.Columns.Add(col.Name)
        Next
        For Each row As DataGridViewRow In dgvNameFormat.Rows
            Dim dRow As DataRow = NameFormatDT.NewRow
            For Each cell As DataGridViewCell In row.Cells
                dRow(cell.ColumnIndex) = cell.Value
            Next
            NameFormatDT.Rows.Add(dRow)
        Next
        dgvNameFormat.SelectAll()
        If Not My.Settings.NameFormat Is Nothing Then My.Settings.NameFormat.Clear()
        If My.Settings.NameFormat Is Nothing Then My.Settings.NameFormat = New Specialized.StringCollection
        If dgvNameFormat.SelectedRows.Count > 0 Then
            Dim DRC As DataGridViewSelectedRowCollection = dgvNameFormat.SelectedRows
            Dim IDS As New List(Of String)
            For i As Integer = 0 To DRC.Count - 1
                Dim id As String = DRC(i).Cells(0).Value & "*" & DRC(i).Cells(1).Value
                My.Settings.NameFormat.Add(id)
            Next
        End If
        My.Settings.MaxRefNum = nudMaxRef.Value
    End Sub
#End Region
#Region "Print"
    Private Sub PrintLoad()
        chkDWGLocation.Checked = My.Settings.PrintDwgLoc
        If My.Settings.PrintDwgLoc = True Then
            Select Case My.Settings.DWGLoc
                Case "TR"
                    rdoTR.Checked = True
                Case "TL"
                    rdoTL.Checked = True
                Case "BR"
                    rdoBR.Checked = True
                Case "BL"
                    rdoBL.Checked = True
            End Select
        End If
        Select Case My.Settings.PrintRange
            Case 0
                rdoAllPages.Checked = True
            Case 1
                rdoFirstPage.Checked = True
            Case 2
                rdoCurrentPage.Checked = True
        End Select
        Select Case My.Settings.PrintColour
            Case True
                rdoColour.Checked = True
            Case False
                rdoBW.Checked = True
        End Select
        If My.Settings.Scale = False Then
            rdoScale.Checked = False
            rdoFull.Checked = True
        Else
            rdoScale.Checked = True
            rdoFull.Checked = False
        End If
        txtCopies.Value = My.Settings.PrintCopies
        chkReverse.Checked = My.Settings.PrintReverse
        cmbScaleB.SelectedIndex = My.Settings.BScale
        cmbScaleC.SelectedIndex = My.Settings.CScale
        cmbScaleD.SelectedIndex = My.Settings.DScale
        cmbScaleE.SelectedIndex = My.Settings.EScale
        cmbScaleF.SelectedIndex = My.Settings.FScale
        chkA.Checked = My.Settings.ASize
        chkB.Checked = My.Settings.BSize
        chkC.Checked = My.Settings.CSize
        chkD.Checked = My.Settings.DSize
        chkE.Checked = My.Settings.ESize
        chkF.Checked = My.Settings.FSize
    End Sub
    Private Sub rdoScale_CheckedChanged(sender As Object, e As EventArgs) Handles rdoScale.CheckedChanged
        gbxScale.Visible = rdoScale.Checked
        My.Settings.Scale = rdoScale.Checked
    End Sub
    Private Sub rdoBR_CheckedChanged(sender As Object, e As EventArgs) Handles rdoBR.CheckedChanged
        If rdoBR.Checked = True Then
            rdoBL.Checked = False
            rdoTL.Checked = False
            rdoTR.Checked = False
            My.Settings.DWGLoc = "BR"
        End If
    End Sub
    Private Sub rdoTL_CheckedChanged(sender As Object, e As EventArgs) Handles rdoTL.CheckedChanged
        If rdoTL.Checked = True Then
            rdoBL.Checked = False
            rdoTR.Checked = False
            rdoBR.Checked = False
            My.Settings.DWGLoc = "TL"
        End If
    End Sub
    Private Sub rdoBL_CheckedChanged(sender As Object, e As EventArgs) Handles rdoBL.CheckedChanged
        If rdoBL.Checked = True Then
            rdoTL.Checked = False
            rdoTR.Checked = False
            rdoBR.Checked = False
            My.Settings.DWGLoc = "BL"
        End If
    End Sub
    Private Sub rdoTR_CheckedChanged(sender As Object, e As EventArgs) Handles rdoTR.CheckedChanged
        If rdoTR.Checked = True Then
            rdoBL.Checked = False
            rdoTL.Checked = False
            rdoBR.Checked = False
            My.Settings.DWGLoc = "TR"
        End If
    End Sub
    Private Sub chkDWGLocation_CheckedChanged(sender As Object, e As EventArgs) Handles chkDWGLocation.CheckedChanged
        If chkDWGLocation.Checked = True Then
            Select Case My.Settings.DWGLoc
                Case "TR"
                    rdoTR.Checked = True
                Case "TL"
                    rdoTL.Checked = True
                Case "BR"
                    rdoBR.Checked = True
                Case "BL"
                    rdoBL.Checked = True
            End Select
        Else
            rdoTR.Checked = False
            rdoTL.Checked = False
            rdoBR.Checked = False
            rdoBL.Checked = False
        End If
    End Sub
    Private Sub PrintOK()
        My.Settings.Scale = rdoScale.Checked
        My.Settings.BScale = cmbScaleB.SelectedIndex
        My.Settings.CScale = cmbScaleC.SelectedIndex
        My.Settings.DScale = cmbScaleD.SelectedIndex
        My.Settings.EScale = cmbScaleE.SelectedIndex
        My.Settings.FScale = cmbScaleF.SelectedIndex
        My.Settings.ASize = chkA.Checked
        My.Settings.BSize = chkb.Checked
        My.Settings.CSize = chkc.checked
        My.Settings.DSize = chkD.Checked
        My.Settings.ESize = chkE.Checked
        My.Settings.FSize = chkF.Checked
        If rdoCurrentPage.Checked = True Then
            My.Settings.PrintRange = 2
        ElseIf rdoFirstPage.Checked = True Then
            My.Settings.PrintRange = 1
        Else
            My.Settings.PrintRange = 0
        End If
        If rdoFull.Checked = True Then
            My.Settings.PrintSize = 0
        Else
            My.Settings.PrintSize = 1
        End If
        If rdoColour.Checked = True Then
            My.Settings.PrintColour = True
        Else
            My.Settings.PrintColour = False
        End If
        My.Settings.PrintDwgLoc = chkDWGLocation.Checked
        My.Settings.PrintReverse = chkReverse.Checked
        txtCopies.Value = My.Settings.PrintCopies
    End Sub
#End Region
#Region "Export"
    Private Sub ExportLoad()
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
        Select Case My.Settings.PrintSize
            Case 0
                rdoFull.Checked = True
            Case 1
                rdoScale.Checked = True
        End Select
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
    Private Sub ExportOK()
        If rdoPDFSaveLoc.Checked = True Then
            My.Settings.PDFSaveLoc = txtPDFSaveLoc.Text
            My.Settings.PDFSaveNewLoc = True
            My.Settings.PDFSaveTag = False
        ElseIf rdoPDFTag.Checked = True Then
            My.Settings.PDFTag = txtPDFTag.Text
            My.Settings.PDFSaveNewLoc = False
            My.Settings.PDFSaveTag = True
        Else
            My.Settings.PDFSaveNewLoc = False
            My.Settings.PDFSaveTag = False
        End If
        My.Settings.PDFRev = chkPDFRev.Checked
        If rdoDXFSaveLoc.Checked = True Then
            My.Settings.DXFSaveLoc = txtDXFSaveLoc.Text
            My.Settings.DXFSaveNewLoc = True
            My.Settings.DXFSaveTag = False
        ElseIf rdoDXFTag.Checked = True Then

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

            My.Settings.DWGTag = txtDWGTag.Text
            My.Settings.DWGSaveNewLoc = False
            My.Settings.DWGSaveTag = True
        Else
            My.Settings.DWGSaveNewLoc = False
            My.Settings.DWGSaveTag = False
        End If
        My.Settings.DWGRev = chkDWGRev.Checked
        My.Settings.ArchiveExport = chkArchive.Checked
        My.Settings.DWGini = chkCustDWGini.Checked
        My.Settings.DWGiniLoc = txtCustDWGini.Text
        My.Settings.DXFini = chkCustDXFini.Checked
        My.Settings.DXFiniLoc = txtCustDXFini.Text
        My.Settings.PDFLineWeights = chkLineWeights.Checked
        My.Settings.PDFColoursBlack = chkPDFBW.Checked
        My.Settings.PDFRange = cmbSheets.SelectedIndex
        My.Settings.PDFRes = numRes.Value
        rdoStrict.Checked = My.Settings.StrictSearch
        My.Settings.Experimental = chkExperimental.Checked
    End Sub
#End Region
#Region "Rev Table"
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
        My.Settings.StartVal = nudStartVal.Value
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
#End Region
#Region "Functions"
    Private Function RemoveInvalidFileNameChars(UserInput As String) As String
        For Each invalidChar In IO.Path.GetInvalidFileNameChars
            UserInput = UserInput.Replace(invalidChar, "")
        Next
        Return UserInput
    End Function
    Public Function ConvertDatatableToXML(ByVal dt As DataTable) As String
        Dim str As New MemoryStream()
        dt.WriteXml(str, True)
        str.Seek(0, SeekOrigin.Begin)
        Dim sr As New StreamReader(str)
        Dim xmlstr As String
        xmlstr = sr.ReadToEnd()
        Return (xmlstr)
    End Function
    Private Function IsAlphaNum(ByVal strInputText As String) As Boolean
        Dim IsAlpha As Boolean = False
        If System.Text.RegularExpressions.Regex.IsMatch(strInputText, "^[a-zA-Z0-9-\\]+$") Then
            IsAlpha = True
        Else
            IsAlpha = False
        End If
        Return IsAlpha
    End Function
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
    Public Shared Function FilenameIsOK(ByVal fileName As String) As Boolean
        If fileName = Nothing Then Return False
        Dim file As String = Path.GetFileName(fileName)
        Dim directory As String = Path.GetDirectoryName(fileName)

        Return Not (file.Intersect(Path.GetInvalidFileNameChars()).Any() _
                OrElse
                directory.Intersect(Path.GetInvalidPathChars()).Any())
    End Function
    Function TestValues() As Boolean
        If txtPDFSaveLoc.Text = Nothing AndAlso rdoPDFSaveLoc.Checked = True Then
            MsgBox("A valid location is needed for the PDF save location")
            tcrlSettings.SelectedIndex = 2
            tabSaveLoc.SelectedIndex = 0
            Return False
        End If

        If txtDXFSaveLoc.Text = Nothing AndAlso rdoDXFSaveLoc.Checked = True Then
            MsgBox("A valid location is needed for the DXF save location")
            tcrlSettings.SelectedIndex = 2
            tabSaveLoc.SelectedIndex = 1
            Return False
        End If

        If txtDWGSaveLoc.Text = Nothing AndAlso rdoDWGSaveLoc.Checked = True Then
            MsgBox("A valid location is needed for the DWG save location")
            tcrlSettings.SelectedIndex = 2
            tabSaveLoc.SelectedIndex = 2
            Return False
        End If
        If txtPDFTag.Text = Nothing AndAlso rdoPDFTag.Checked = True Then
            If FilenameIsOK(txtPDFTag.Text) = False Then
                'If Regex.IsMatch(txtPDFTag.Text, ("^[a-zA-Z0-9_]*$")) Then
                MsgBox("The PDF export location you have selected Is invalid")
                tcrlSettings.SelectedIndex = 2
                tabSaveLoc.SelectedIndex = 0
                Return False
            End If
            If Strings.Left(txtPDFTag.Text, 1) <> "\" Then
                txtPDFTag.Text = "\" & txtPDFTag.Text
            End If
            If Strings.Right(txtPDFTag.Text, 1) = "\" Then
                txtPDFTag.Text = Strings.Left(txtPDFTag.Text, Len(txtPDFTag.Text) - 1)
            End If
        End If
        If txtDXFTag.Text = Nothing AndAlso rdoDXFTag.Checked = True Then
            If FilenameIsOK(txtDXFTag.Text) = False Then
                'If Regex.IsMatch(txtDXFTag.Text, ("^[a-zA-Z0-9_]*$")) Then
                MsgBox("The DXF export location you have selected Is invalid")
                tcrlSettings.SelectedIndex = 2
                tabSaveLoc.SelectedIndex = 1
                Return False
            End If
            If Strings.Left(txtDXFTag.Text, 1) <> "\" Then
                txtDXFTag.Text = "\" & txtDXFTag.Text
            End If
            If Strings.Right(txtDXFTag.Text, 1) = "\" Then
                txtDXFTag.Text = Strings.Left(txtDXFTag.Text, Len(txtDXFTag.Text) - 1)
            End If
        End If
        If txtDWGTag.Text = Nothing AndAlso rdoDWGTag.Checked = True Then
            If FilenameIsOK(txtDWGTag.Text) = False Then
                'If Regex.IsMatch(txtDWGTag.Text, ("^[a-zA-Z0-9_]*$")) Then
                MsgBox("The DWG export location you have selected Is invalid")
                tcrlSettings.SelectedIndex = 2
                tabSaveLoc.SelectedIndex = 2
                Return False
            End If
            If Strings.Left(txtDWGTag.Text, 1) <> "\" Then
                txtDWGTag.Text = "\" & txtDWGTag.Text
            End If
            If Strings.Right(txtDWGTag.Text, 1) = "\" Then
                txtDWGTag.Text = Strings.Left(txtDWGTag.Text, Len(txtDWGTag.Text) - 1)
            End If
        End If
        If txtAlphaRev.Text = Nothing Then
            MsgBox("Please enter a default description for alphabetical revisions")
            tcrlSettings.SelectedIndex = 3
            Return False
        End If
        If txtNumRev.Text = Nothing Then
            MsgBox("Please enter a default description for numerical revisions")
            tcrlSettings.SelectedIndex = 3
            Return False
        End If
        Return True
    End Function
#End Region
    Private Sub btnCancel_Click(sender As Object, e As EventArgs) Handles btnCancel.Click
        Me.Close()
    End Sub
    Private Sub btnOK_Click(sender As Object, e As EventArgs) Handles btnOK.Click
        ExportOK()
        GeneralOK()
        RevTableOK()
        PrintOK()
        If TestValues() = True Then
            My.Settings.Save()
            Me.Close()
        End If
    End Sub
    Private Sub btnApply_Click(sender As Object, e As EventArgs) Handles btnApply.Click
        If TestValues() = True Then My.Settings.Save()
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
End Class