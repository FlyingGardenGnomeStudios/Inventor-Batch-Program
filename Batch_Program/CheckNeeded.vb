Imports System.Windows.Forms
Imports Inventor
Imports System.Runtime.InteropServices
Imports System.Collections.Generic
Imports System.Globalization
Imports System.Drawing
Imports AdvancedDataGridView
Imports System.Windows.Forms.DataGridView
Imports System.ComponentModel

Public Class CheckNeeded
    Public Declare Function mouse_event Lib "user32" Alias "mouse_event" (ByVal dwFlags As Long, ByVal dx As Long, ByVal dy As Long, ByVal cButtons As Long, ByVal dwExtraInfo As Long)
    'Create global references for RevTable form
    Dim bCancelEdit As Boolean
    Dim _invApp As Inventor.Application
    Dim CurrentItem As Windows.Forms.ListViewItem
    Dim CurrentSB As Windows.Forms.ListViewItem.ListViewSubItem
    Dim Main As Main
    Dim RevTable As New RevTable
    Dim oDateTimePicker As New DateTimePicker
    Dim LastCell As DataGridViewCell
    Dim ChangeFlag As Boolean = False
    Dim SelectedCell As DataGridViewCell
    Public Sub New()
        ' This call is required by the designer.
        InitializeComponent()
        _invApp = Marshal.GetActiveObject("Inventor.Application")
        If My.Settings.RTSCheckedBy = True Then
            Dim RTSCheckedBy As New DataGridViewTextBoxColumn With
                {.Name = "CheckedBy",
                .HeaderText = "Checked By",
                .SortMode = DataGridViewColumnSortMode.Automatic,
                .ReadOnly = False}
            tgvCheckNeeded.Columns.Add(RTSCheckedBy)
        End If
        If My.Settings.RTSCheckedDate = True Then
            Dim RTSCheckDate As New GridDateControl With
              {.Name = "CheckDate",
              .HeaderText = "Check Date",
              .SortMode = DataGridViewColumnSortMode.Automatic,
              .ReadOnly = False}
            tgvCheckNeeded.Columns.Add(RTSCheckDate)
        End If
        If My.Settings.RTSRev = True Then
            Dim RTSRev As New DataGridViewTextBoxColumn With
             {.Name = My.Settings.RTSRevCol,
        .HeaderText = My.Settings.RTSRevCol,
        .SortMode = DataGridViewColumnSortMode.Automatic,
        .Tag = "Number",
        .ReadOnly = True}
            tgvCheckNeeded.Columns.Add(RTSRev)
        End If
        If My.Settings.RTSDate = True Then
            'Dim RTSRevDate As New GridDateControl With
            Dim RTSRevDate As New DataGridViewTextBoxColumn With
               {.Name = My.Settings.RTSDateCol,
              .HeaderText = My.Settings.RTSDateCol,
              .SortMode = DataGridViewColumnSortMode.Automatic,
              .ReadOnly = False}
            tgvCheckNeeded.Columns.Add(RTSRevDate)
        End If
        If My.Settings.RTSDesc = True Then
            Dim RTSDescription As New DataGridViewTextBoxColumn With
                {.Name = My.Settings.RTSDescCol,
                .HeaderText = My.Settings.RTSDescCol,
                .SortMode = DataGridViewColumnSortMode.Automatic,
                .ReadOnly = False}
            tgvCheckNeeded.Columns.Add(RTSDescription)
        End If
        If My.Settings.RTSName = True Then
            Dim RTSName As New DataGridViewTextBoxColumn With
                {.Name = My.Settings.RTSNameCol,
                .HeaderText = My.Settings.RTSNameCol,
                .SortMode = DataGridViewColumnSortMode.Automatic,
                .ReadOnly = False}
            tgvCheckNeeded.Columns.Add(RTSName)
        End If
        If My.Settings.RTSApproved = True Then
            Dim RTSApproved As New DataGridViewTextBoxColumn With
    {.Name = My.Settings.RTSApprovedCol,
    .HeaderText = My.Settings.RTSApprovedCol,
    .SortMode = DataGridViewColumnSortMode.Automatic,
    .ReadOnly = False}
            tgvCheckNeeded.Columns.Add(RTSApproved)
        End If

        If My.Settings.RTS1 = True Then
            Select Case My.Settings.RTS1Value
                Case "Text"
                    Dim RTS1Value As New DataGridViewTextBoxColumn With
                        {.Name = My.Settings.RTS1Item,
                        .HeaderText = My.Settings.RTS1Col,
                        .SortMode = DataGridViewColumnSortMode.Automatic,
                        .ReadOnly = False}
                    tgvCheckNeeded.Columns.Add(RTS1Value)
                Case "Number"
                    Dim RTS1Value As New DataGridViewTextBoxColumn With
                        {.Name = My.Settings.RTS1Item,
                        .HeaderText = My.Settings.RTS1Col,
                        .SortMode = DataGridViewColumnSortMode.Automatic,
                        .Tag = "Number",
                        .ReadOnly = False}
                    tgvCheckNeeded.Columns.Add(RTS1Value)
                Case "Date"
                    'Dim RTS1Value As New GridDateControl With
                    Dim RTS1Value As New DataGridViewTextBoxColumn With
                        {.Name = My.Settings.RTS1Item,
                        .HeaderText = My.Settings.RTS1Col,
                        .SortMode = DataGridViewColumnSortMode.Automatic,
                        .ReadOnly = False}
                    tgvCheckNeeded.Columns.Add(RTS1Value)
            End Select

        End If
        If My.Settings.RTS2 = True Then
            Select Case My.Settings.RTS2Value
                Case "Text"
                    Dim RTS2Value As New DataGridViewTextBoxColumn With
                        {.Name = My.Settings.RTS2Item,
                        .HeaderText = My.Settings.RTS2Col,
                        .SortMode = DataGridViewColumnSortMode.Automatic,
                        .ReadOnly = False}
                    tgvCheckNeeded.Columns.Add(RTS2Value)
                Case "Number"
                    Dim RTS2Value As New DataGridViewTextBoxColumn With
                        {.Name = My.Settings.RTS2Item,
                        .HeaderText = My.Settings.RTS2Col,
                        .SortMode = DataGridViewColumnSortMode.Automatic,
                        .Tag = "Number",
                        .ReadOnly = False}
                    tgvCheckNeeded.Columns.Add(RTS2Value)
                Case "Date"
                    Dim RTS2Value As New GridDateControl With
                        {.Name = My.Settings.RTS2Item,
                        .HeaderText = My.Settings.RTS2Col,
                        .SortMode = DataGridViewColumnSortMode.Automatic,
                        .ReadOnly = False}
                    tgvCheckNeeded.Columns.Add(RTS2Value)
            End Select
        End If
        If My.Settings.RTS3 = True Then
            Select Case My.Settings.RTS3Value
                Case "Text"
                    Dim RTS3Value As New DataGridViewTextBoxColumn With
                        {.Name = My.Settings.RTS3Item,
                        .HeaderText = My.Settings.RTS3Col,
                        .SortMode = DataGridViewColumnSortMode.Automatic,
                        .ReadOnly = False}
                    tgvCheckNeeded.Columns.Add(RTS3Value)
                Case "Number"
                    Dim RTS3Value As New DataGridViewTextBoxColumn With
                        {.Name = My.Settings.RTS3Item,
                        .HeaderText = My.Settings.RTS3Col,
                        .SortMode = DataGridViewColumnSortMode.Automatic,
                        .Tag = "Number",
                        .ReadOnly = False}
                    tgvCheckNeeded.Columns.Add(RTS3Value)
                Case "Date"
                    Dim RTS3Value As New GridDateControl With
                        {.Name = My.Settings.RTS3Item,
                        .HeaderText = My.Settings.RTS3Col,
                        .SortMode = DataGridViewColumnSortMode.Automatic,
                        .ReadOnly = False}
                    tgvCheckNeeded.Columns.Add(RTS3Value)
            End Select
        End If
        If My.Settings.RTS4 = True Then
            Select Case My.Settings.RTS4Value
                Case "Text"
                    Dim RTS4Value As New DataGridViewTextBoxColumn With
                        {.Name = My.Settings.RTS4Item,
                        .HeaderText = My.Settings.RTS4Col,
                        .SortMode = DataGridViewColumnSortMode.Automatic,
                        .ReadOnly = False}
                    tgvCheckNeeded.Columns.Add(RTS4Value)
                Case "Number"
                    Dim RTS4Value As New DataGridViewTextBoxColumn With
                        {.Name = My.Settings.RTS4Item,
                        .HeaderText = My.Settings.RTS4Col,
                        .SortMode = DataGridViewColumnSortMode.Automatic,
                        .Tag = "Number",
                        .ReadOnly = False}
                    tgvCheckNeeded.Columns.Add(RTS4Value)
                Case "Date"
                    Dim RTS4Value As New GridDateControl With
                        {.Name = My.Settings.RTS4Item,
                        .HeaderText = My.Settings.RTS4Col,
                        .SortMode = DataGridViewColumnSortMode.Automatic,
                        .ReadOnly = False}
                    tgvCheckNeeded.Columns.Add(RTS4Value)
            End Select
        End If
        If My.Settings.RTS5 = True Then
            Select Case My.Settings.RTS5Value
                Case "Text"
                    Dim RTS5Value As New DataGridViewTextBoxColumn With
                        {.Name = My.Settings.RTS5Item,
                        .HeaderText = My.Settings.RTS5Col,
                        .SortMode = DataGridViewColumnSortMode.Automatic,
                        .ReadOnly = False}
                    tgvCheckNeeded.Columns.Add(RTS5Value)
                Case "Number"
                    Dim RTS5Value As New DataGridViewTextBoxColumn With
                        {.Name = My.Settings.RTS5Item,
                        .HeaderText = My.Settings.RTS5Col,
                        .SortMode = DataGridViewColumnSortMode.Automatic,
                        .Tag = "Number",
                        .ReadOnly = False}
                    tgvCheckNeeded.Columns.Add(RTS5Value)
                Case "Date"
                    Dim RTS5Value As New GridDateControl With
                        {.Name = My.Settings.RTS5Item,
                        .HeaderText = My.Settings.RTS5Col,
                        .SortMode = DataGridViewColumnSortMode.Automatic,
                        .ReadOnly = False}
                    tgvCheckNeeded.Columns.Add(RTS5Value)
            End Select
        End If
    End Sub
    Public Function PopMain(CalledFunction As Main)
        Main = CalledFunction
        Return Nothing
    End Function
    Public Function PopRevTable(CalledFunction As RevTable)
        RevTable = CalledFunction
        Return Nothing
    End Function
    Private Sub btnCancel_Click(sender As System.Object, e As System.EventArgs) Handles btnCancel.Click
        Main.chkCheck.Checked = False
        Main.chkDXF.Checked = False
        Main.chkPDF.Checked = False
        Me.Close()
    End Sub
    Private Sub btnHide_Click(sender As System.Object, e As System.EventArgs) Handles btnHide.Click
        Dim Remove As Boolean
        If btnHide.Text = "Hide Completed" Then

            For Each ParentNode As TreeGridNode In tgvCheckNeeded.Nodes
                Remove = True
                For Each Cell In ParentNode.Cells
                    If Cell.Value = "" Then
                        Remove = False
                        Exit For
                    End If
                Next
                If Remove = True AndAlso ParentNode.HasChildren Then
                    For Each ChildNode In ParentNode.Nodes
                        For Each Childcell In ChildNode.Cells
                            If Childcell.value = Nothing Then
                                If Childcell.readonly = False Then
                                    '     If tgvCheckNeeded.Columns(Childcell.columnindex).headertext <> "Checked By" AndAlso
                                    '      tgvCheckNeeded.Columns(Childcell.columnindex).headertext <> "Check Date" Then
                                    Remove = False
                                        Exit For
                                    End If
                                End If

                        Next
                    Next
                End If
                If Remove = True Then
                    btnHide.Text = "Show Completed"
                    For Each Childnode In ParentNode.Nodes
                        Childnode.Visible = False
                    Next
                    ParentNode.Visible = False
                End If
            Next
        Else
            btnHide.Text = "Hide Completed"
            For Each Node As TreeGridNode In tgvCheckNeeded.Rows
                Node.Visible = True
                If Node.HasChildren = True Then
                    For Each Childnode In Node.Nodes
                        Childnode.Visible = True
                    Next
                End If
            Next
        End If
    End Sub
    Private Sub btnOK_Click(sender As System.Object, e As System.EventArgs) Handles btnOK.Click
        Dim DrawingName, DrawSource, Rev As String
        Dim Sheet As Sheet
        Dim Col, Row, Ans As Integer
        Dim RevTable As RevisionTable
        Dim Path As Documents = _invApp.Documents
        Dim oDoc As Document = Nothing
        Dim Opendocs As New ArrayList
        Main.CreateOpenDocs()
        ProgressBar1.Visible = True
        Dim RevNode As TreeGridNode
        For Each node As TreeGridNode In tgvCheckNeeded.Rows
            node.Expand()
        Next
        'For each drawing name, find the drawing source related to it
        'If Me.btnIgnore.Visible = True Then
        '    For Y = 0 To lstCheckNeeded.Items.Count - 1
        '        For Z = 0 To lstCheckNeeded.Items(Y).SubItems.Count - 2
        '            If lstCheckNeeded.Items(Y).SubItems(Z).Text = "" And Me.btnIgnore.Visible = True And Ans <> vbOK Then
        '                Ans = MsgBox("Some categories are missing information" & vbNewLine &
        '                       "(Make sure you check the 'More' category)" & vbNewLine &
        '                       "Do you wish to continue?", MsgBoxStyle.OkCancel)
        '                If Ans = vbCancel Then
        '                    'Me.Hide()
        '                    Exit Sub
        '                End If
        '            End If
        '        Next
        '    Next
        'End If
        'Iterate through checkneeded table to get drawing name
        Dim Flag As Boolean = False
        DrawSource = Nothing
        For Each node In tgvCheckNeeded.Nodes
            Dim TotRevs As Integer = 1
            If node.HasChildren Then
                For Each ChildNode In node.Nodes
                    TotRevs += 1
                Next
            End If
            'For Each row In dgvCheckNeeded.Rows
            'If tgvCheckNeeded.IsCurrentRowDirty Then
            DrawingName = tgvCheckNeeded(tgvCheckNeeded.Columns("DrawingName").Index, node.RowIndex).Value
            For Each X In Main.dgvSubFiles.Rows
                If DrawingName = Main.dgvSubFiles(Main.dgvSubFiles.Columns("DrawingName").Index, X.index).Value Then
                    DrawSource = Main.dgvSubFiles(Main.dgvSubFiles.Columns("DrawingLocation").Index, X.index).Value
                    Exit For
                End If
            Next
            If DrawSource = Nothing Then MsgBox("The drawing could not be found." & vbNewLine & "If the drawing is saved in a different location" & vbNewLine &
                                   "than the model, this can cause an error.")
            'Open the related drawing in the background
            Try
                oDoc = _invApp.Documents.Open(DrawSource, False)

                'MsgBox("An error occurred in opening " & DrawSource & vbNewLine & "The values for this drawing will not be updated")
                For Each Column In tgvCheckNeeded.Columns
                    Select Case Column.name
                        Case "CheckedBy"
                            If tgvCheckNeeded(Column.index, node.RowIndex).Value = "" Then
                                oDoc.PropertySets.Item("{32853F0F-3444-11D1-9E93-0060B03C1CA6}").ItemByPropId("10").Value = ""
                            Else
                                oDoc.PropertySets.Item("{32853F0F-3444-11D1-9E93-0060B03C1CA6}").ItemByPropId("10").Value = tgvCheckNeeded(Column.index, node.RowIndex).Value
                            End If
                        Case "CheckDate"
                            If tgvCheckNeeded(Column.index, node.RowIndex).Value = "" Then
                                oDoc.PropertySets.Item("{32853F0F-3444-11D1-9E93-0060B03C1CA6}").ItemByPropId("11").Value = #1/1/1601#
                            Else
                                oDoc.PropertySets.Item("{32853F0F-3444-11D1-9E93-0060B03C1CA6}").ItemByPropId("11").Value = tgvCheckNeeded(Column.index, node.RowIndex).Value
                            End If
                    End Select
                Next
                Sheet = oDoc.ActiveSheet
                RevTable = Sheet.RevisionTables(1)
                If RevTable.RevisionTableRows.Count < TotRevs Then
                    Do Until RevTable.RevisionTableRows.Count = TotRevs
                        RevTable.RevisionTableRows.Add()
                    Loop
                ElseIf RevTable.RevisionTableRows.Count > TotRevs Then
                    RevTable.RevisionTableRows.Add()
                    Do Until RevTable.RevisionTableRows.Count = TotRevs
                        RevTable.RevisionTableRows(RevTable.RevisionTableRows.Count - 1).Delete()
                    Loop
                End If

                Col = RevTable.RevisionTableColumns.Count
                Row = RevTable.RevisionTableRows.Count
                Rev = oDoc.PropertySets.Item("{F29F85E0-4FF9-1068-AB91-08002B27B3D9}").ItemByPropId("9").Value
                'Iterate through rev table and populate from the userform
                Dim Contents(0 To Col * Row) As String
                Dim J As Integer = 0
                Dim Initials(0 To Row), RevCheckBy(0 To Row), RevDate(0 To Row) As String

                For RevRow = 1 To RevTable.RevisionTableRows.Count
                    Dim i As Integer = 1
                    If RevRow > 1 Then
                        RevNode = node.Nodes.Item(RevRow - 2)
                    Else
                        RevNode = node
                    End If
                    For Each rtc In RevTable.RevisionTableColumns
                        Dim h As New DataGridViewTextBoxColumn
                        Dim rtcell As RevisionTableCell = Nothing
                        h.Name = CultureInfo.CurrentCulture.TextInfo.ToTitleCase(LCase(rtc.Title))

                        WriteRevCase(RevNode, h.Name, RevRow, i, RevTable, oDoc)
                        i = i + 1
                    Next
                Next
                oDoc.Update()
                _invApp.SilentOperation = True
                Try
                    oDoc.Save()
                Catch ex As Exception
                    Ans = MsgBox("Drawing " & DrawingName & " could not be saved." & vbNewLine _
                                 & "Make sure the drawing is not read-only", MsgBoxStyle.OkCancel, "Error During Save")
                    If Ans = vbCancel Then
                        ProgressBar1.Hide()
                        Exit Sub
                    End If
                End Try
                Main.CloseLater(DrawingName, oDoc)
                _invApp.SilentOperation = False
                ProgressBar1.Value = (Row / tgvCheckNeeded.RowCount) * 100
                ProgressBar1.PerformStep()
                'Exit For
            Catch ex As Exception
            End Try
            'End If
            If node.HasChildren Then
                For Each Childnode In node.Nodes

                Next
            End If
        Next
        'Main.chkCheck.CheckState = CheckState.Indeterminate
        ProgressBar1.Visible = False
        Me.Close()
        'If Main.chkCheck.CheckState = CheckState.Indeterminate Then Main.ExportCheck(Path, odoc:=Nothing, Archive:="", Main.dgvSubFiles, DrawingName:="", DrawSource:="", ExportType:="dxf")
    End Sub
    Private Sub WriteRevCase(node As TreeGridNode, Title As String, RevRow As Integer, i As Integer, RevTable As RevisionTable, oDoc As Document)
        Select Case UCase(Title)
            Case UCase(My.Settings.RTSRevCol)
                If My.Settings.RTSDate = True Then
                    If node.Cells(tgvCheckNeeded.Columns(My.Settings.RTSRevCol).Index).Value Is Nothing Then
                        RevTable.RevisionTableRows(RevRow).Item(i).Text = ""
                    End If
                    RevTable.RevisionTableRows(RevRow).Item(i).Text = node.Cells(tgvCheckNeeded.Columns(My.Settings.RTSRevCol).Index).Value
                End If
            Case UCase(My.Settings.RTSDateCol)
                If My.Settings.RTSDate = True Then
                    If node.Cells(tgvCheckNeeded.Columns(My.Settings.RTSDateCol).Index).Value Is Nothing Then
                        RevTable.RevisionTableRows(RevRow).Item(i).Text = ""
                    End If
                    RevTable.RevisionTableRows(RevRow).Item(i).Text = node.Cells(tgvCheckNeeded.Columns(My.Settings.RTSDateCol).Index).Value
                End If
            Case UCase(My.Settings.RTSDescCol)
                If My.Settings.RTSDesc = True Then
                    If node.Cells(tgvCheckNeeded.Columns(My.Settings.RTSDescCol).Index).Value Is Nothing Then
                        RevTable.RevisionTableRows.Item(RevRow).Item(i).Text = ""
                    End If
                    RevTable.RevisionTableRows.Item(RevRow).Item(i).Text = node.Cells(tgvCheckNeeded.Columns(My.Settings.RTSDescCol).Index).Value
                End If
            Case UCase(My.Settings.RTSNameCol)
                If My.Settings.RTSName = True Then
                    If node.Cells(tgvCheckNeeded.Columns(My.Settings.RTSNameCol).Index).Value Is Nothing Then
                        RevTable.RevisionTableRows.Item(RevRow).Item(i).Text = ""
                    End If
                    RevTable.RevisionTableRows.Item(RevRow).Item(i).Text = node.Cells(tgvCheckNeeded.Columns(My.Settings.RTSNameCol).Index).Value
                End If
            Case UCase(My.Settings.RTSApprovedCol)
                If My.Settings.RTSApproved = True Then
                    If node.Cells(tgvCheckNeeded.Columns(My.Settings.RTSApprovedCol).Index).Value Is Nothing Then
                        RevTable.RevisionTableRows.Item(RevRow).Item(i).Text = ""
                    End If
                    RevTable.RevisionTableRows.Item(RevRow).Item(i).Text = node.Cells(tgvCheckNeeded.Columns(My.Settings.RTSApprovedCol).Index).Value
                End If
            Case UCase(My.Settings.RTS1Col)
                If My.Settings.RTS1 = True Then
                    If node.Cells(tgvCheckNeeded.Columns(My.Settings.RTS1Col).Index).Value Is Nothing Then
                        RevTable.RevisionTableRows.Item(RevRow).Item(i).Text = ""
                    End If
                    RevTable.RevisionTableRows.Item(RevRow).Item(i).Text = node.Cells(tgvCheckNeeded.Columns(My.Settings.RTS1Col).Index).Value
                End If
            Case UCase(My.Settings.RTS2Col)
                If My.Settings.RTS2 = True Then
                    If node.Cells(tgvCheckNeeded.Columns(My.Settings.RTS2Col).Index).Value Is Nothing Then
                        RevTable.RevisionTableRows.Item(RevRow).Item(i).Text = ""
                    End If
                    RevTable.RevisionTableRows.Item(RevRow).Item(i).Text = node.Cells(tgvCheckNeeded.Columns(My.Settings.RTS2Col).Index).Value
                End If
            Case UCase(My.Settings.RTS3Col)
                If My.Settings.RTS3 = True Then
                    If node.Cells(tgvCheckNeeded.Columns(My.Settings.RTS4Col).Index).Value Is Nothing Then
                        RevTable.RevisionTableRows.Item(RevRow).Item(i).Text = ""
                    End If
                    RevTable.RevisionTableRows.Item(RevRow).Item(i).Text = node.Cells(tgvCheckNeeded.Columns(My.Settings.RTS3Col).Index).Value
                End If
            Case UCase(My.Settings.RTS4Col)
                If My.Settings.RTS4 = True Then
                    If node.Cells(tgvCheckNeeded.Columns(My.Settings.RTS4Col).Index).Value Is Nothing Then
                        RevTable.RevisionTableRows.Item(RevRow).Item(i).Text = ""
                    End If
                    RevTable.RevisionTableRows.Item(RevRow).Item(i).Text = node.Cells(tgvCheckNeeded.Columns(My.Settings.RTS4Col).Index).Value
                End If
            Case UCase(My.Settings.RTS5Col)
                If My.Settings.RTS5 = True Then
                    If node.Cells(tgvCheckNeeded.Columns(My.Settings.RTS5Col).Index).Value Is Nothing Then
                        RevTable.RevisionTableRows.Item(RevRow).Item(i).Text = ""
                    End If
                    RevTable.RevisionTableRows.Item(RevRow).Item(i).Text = node.Cells(tgvCheckNeeded.Columns(My.Settings.RTS5Col).Index).Value
                End If
        End Select
    End Sub
    Public Sub PopulateCheckNeeded(Path As Documents, ByRef odoc As Document, ByRef Archive As String _
                                 , ByRef DrawingName As String, ByRef DrawSource As String, OpenDocs As ArrayList)

        Dim RevNo, CheckedBy, RevCheckBy, Initials As String
        Initials = ""
        RevCheckBy = ""
        Dim FirstRun As Boolean = True
        Dim DateChecked As Date
        Dim Sheet As Sheet
        Dim RevisionTable As RevisionTable
        Dim i, j As Integer
        Dim RevCheckNameColNum, RevCheckDateColNum, RevCheckRevColNum, RevCheckDescColNum, RevCheckApproveColNum As Integer
        Dim RevCheck1ColNum, RevCheck2ColNum, RevCheck3ColNum, RevCheck4ColNum, RevCheck5ColNum As Integer
        Dim More As Boolean = False
        Main.pgbMain.ProgressBarStyle = MSVistaProgressBar.BarStyle.Continuous
        'Iterate through all files in the subfiles window
        For Y = 0 To Main.dgvSubFiles.RowCount - 1
            'get the checkbox state of each item
            Dim k As Integer = 0
            If Main.dgvSubFiles(Main.dgvSubFiles.Columns("chkSubFiles").Index, Y).Value = True Then
                'retrieve name of selected item
                Main.MatchDrawing(DrawSource, DrawingName, Y)
                'open the document in the background
                odoc = _invApp.Documents.Open(DrawSource, False)
                'set active sheet and retrieve the values from the revision table
                CheckedBy = odoc.PropertySets.Item("{32853F0F-3444-11D1-9E93-0060B03C1CA6}").ItemByPropId("10").Value
                DateChecked = CDate(odoc.PropertySets.Item("{32853F0F-3444-11D1-9E93-0060B03C1CA6}").ItemByPropId("11").Value)
                Dim oPoint As Point2d
                Dim oRevTable As RevisionTable
                'RevTable.PopMain(Me)
                'Set the active sheet
                Sheet = odoc.Activesheet
                Dim oBorder As Border = Sheet.Border
                If Not oBorder Is Nothing Then
                    oPoint = _invApp.TransientGeometry.CreatePoint2d(0.965193999999989, 2.7178)
                Else
                    oPoint = _invApp.TransientGeometry.CreatePoint2d(Sheet.Width, Sheet.Height)
                End If
                'Define the revision table from the active sheet
                RevisionTable = Sheet.RevisionTables(1)
                If Err.Number = 5 Then
                    oRevTable = odoc.ActiveSheet.Revisiontables.Add2(oPoint, False, True, False, 0)
                    'clear the error for future code.
                    Err.Clear()
                    Exit Sub
                End If
                'retrieve the data from the revision table
                RevNo = odoc.PropertySets.Item("{F29F85E0-4FF9-1068-AB91-08002B27B3D9}").ItemByPropId("9").Value
                Dim tg As TransientGeometry = _invApp.TransientGeometry
                Dim dd As DrawingDocument = _invApp.Documents.ItemByName(DrawSource) '.ActiveDocument
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
                Dim node As TreeGridNode = Nothing
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
                    Dim Rev(tgvCheckNeeded.Columns.Count - 1) As String
                    DGVRows = tgvCheckNeeded.Rows.Count - 1
                    Rev(0) = DrawingName
                    For z = 1 To tgvCheckNeeded.ColumnCount - 1
                        Select Case UCase(tgvCheckNeeded.Columns(z).Name)
                            Case UCase("CheckedBy")
                                If k = 0 Then Rev(z) = CheckedBy
                            Case UCase("CheckDate")
                                If k = 0 Then
                                    If DateChecked = #1/1/1601# Then
                                        Rev(z) = ""
                                    Else
                                        Rev(z) = CStr(DateChecked.ToShortDateString)
                                    End If
                                End If
                            Case UCase(My.Settings.RTSRevCol)
                                Rev(z) = contents((k * rt.RevisionTableColumns.Count) + RevCheckRevColNum)
                            Case UCase(My.Settings.RTSDescCol)
                                Rev(z) = contents((k * rt.RevisionTableColumns.Count) + RevCheckDescColNum)
                            Case UCase(My.Settings.RTSDateCol)
                                If contents((k * rt.RevisionTableColumns.Count) + RevCheckDateColNum) = "" Then
                                    Rev(z) = ""
                                ElseIf contents((k * rt.RevisionTableColumns.Count) + RevCheckDateColNum) = #1/1/1601# Then
                                    Rev(z) = ""
                                Else
                                    Rev(z) = CStr(DateTime.Parse(contents((k * rt.RevisionTableColumns.Count) + RevCheckDateColNum)))
                                End If
                            Case UCase(My.Settings.RTSNameCol)
                                Rev(z) = contents((k * rt.RevisionTableColumns.Count) + RevCheckNameColNum)
                            Case UCase(My.Settings.RTSApprovedCol)
                                Rev(z) = contents((k * rt.RevisionTableColumns.Count) + RevCheckApproveColNum)
                            Case UCase(My.Settings.RTS1Item)
                                If RevCheck1ColNum <> 0 Then Rev(z) = contents((k * rt.RevisionTableColumns.Count) + RevCheck1ColNum)
                            Case UCase(My.Settings.RTS2Item)
                                If RevCheck2ColNum <> 0 Then Rev(z) = contents((k * rt.RevisionTableColumns.Count) + RevCheck2ColNum)
                            Case UCase(My.Settings.RTS3Item)
                                If RevCheck3ColNum <> 0 Then Rev(z) = contents((k * rt.RevisionTableColumns.Count) + RevCheck3ColNum)
                            Case UCase(My.Settings.RTS4Item)
                                If RevCheck4ColNum <> 0 Then Rev(z) = contents((k * rt.RevisionTableColumns.Count) + RevCheck4ColNum)
                            Case UCase(My.Settings.RTS5Item)
                                If RevCheck5ColNum <> 0 Then Rev(z) = contents((k * rt.RevisionTableColumns.Count) + RevCheck5ColNum)
                        End Select
                    Next
                    Dim add As Boolean = True
                    'For check = 0 To tgvCheckNeeded.Columns.Count - 1
                    '    If Rev(check) = "" Then
                    '        add = True
                    '        Exit For
                    '    End If
                    'Next
                    Dim result As DataGridViewRow
                    result = New DataGridViewRow
                    If k = 0 And add = True Then
                        node = tgvCheckNeeded.Nodes.Add(Rev)
                    ElseIf k = 1 Then
                        node = node.Nodes.Add(Rev)
                        For Column = 1 To tgvCheckNeeded.Columns.Count - 1
                            If tgvCheckNeeded.Columns(Column).HeaderText = "Checked By" Or
                                    tgvCheckNeeded.Columns(Column).HeaderText = "Check Date" Then
                                node.Cells(Column).ReadOnly = True
                                node.Cells(Column).Style.BackColor = Drawing.Color.LightGray

                            End If
                        Next
                    Else
                        node = node.Parent.Nodes.Add(Rev)
                        For Column = 1 To tgvCheckNeeded.Columns.Count - 1
                            If tgvCheckNeeded.Columns(Column).HeaderText = "Checked By" Or
                                    tgvCheckNeeded.Columns(Column).HeaderText = "Check Date" Then
                                node.Cells(Column).ReadOnly = True
                                node.Cells(Column).Style.BackColor = Drawing.Color.LightGray
                            End If
                        Next
                    End If

                    k += 1
                    j = j + 1
                Next
            End If
            ' Main.ProgressBar(Main.LVSubFiles.CheckedItems.Count, Y + 1, "Checking Revs: ", DrawingName)
        Next
        Main.pgbMain.Visible = False
        Dim tgvColWidth As Double = Nothing
        For Each column In tgvCheckNeeded.Columns
            tgvColWidth = column.width + tgvColWidth
        Next
        If tgvColWidth < tgvCheckNeeded.Width Then

            '    tgvCheckNeeded.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill
            '    tgvCheckNeeded.AutoResizeColumns()
            'Else
            tgvCheckNeeded.AutoSizeColumnsMode = DataGridViewAutoSizeColumnMode.DisplayedCells
            tgvCheckNeeded.AutoResizeColumns()
        End If
    End Sub
    Private Sub tgvCheckNeeded_MouseUp(sender As Object, e As MouseEventArgs) Handles tgvCheckNeeded.MouseUp
        Me.TopMost = True
        Dim hit As DataGridView.HitTestInfo = tgvCheckNeeded.HitTest(e.X, e.Y)
        If e.Button <> MouseButtons.Right Then
            Return
        End If

        If e.X < 1 Or e.Y < 1 Then Exit Sub
        If (hit.RowIndex) < 0 Or (hit.ColumnIndex) < 0 Then Exit Sub
        SelectedCell = tgvCheckNeeded.Rows(hit.RowIndex).Cells(hit.ColumnIndex)
        Dim VisNodes As Integer = 0
        For Each Node As TreeGridNode In tgvCheckNeeded.Rows
            If Node.Visible = True Then
                VisNodes += 1
            End If
        Next
        If VisNodes = 1 Then
            cmsApplyValues.Items.Item(0).Visible = False
            cmsApplyValues.Items.Item(1).Visible = False
        Else
            cmsApplyValues.Items.Item(0).Visible = True
            cmsApplyValues.Items.Item(1).Visible = True
        End If
        If SelectedCell.RowIndex > 0 AndAlso tgvCheckNeeded.Rows(hit.RowIndex - 1).Cells(0).Value = tgvCheckNeeded.Rows(hit.RowIndex).Cells(0).Value Then
            cmsApplyValues.Items.Item(3).Visible = True
        Else
            cmsApplyValues.Items.Item(3).Visible = False
        End If
        cmsApplyValues.Show(Cursor.Position)
        Me.TopMost = False
    End Sub
    Private Sub DateTimePicker(ByVal e As DataGridViewCellEventArgs, ByRef DateRow As Boolean, dgvcontrol As DataGridView)
        If DateRow = True Then
            'Adding DateTimePicker control into DataGridView 
            dgvcontrol.Controls.Add(oDateTimePicker)
            ' Setting the format (i.e. 2014-10-10)
            oDateTimePicker.Format = DateTimePickerFormat.Short
            oDateTimePicker.ShowCheckBox = True
            If dgvcontrol.CurrentCell.Value = Nothing Or dgvcontrol Is Nothing Then
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
    Private Sub iPropdateTimePicker_OnTextChange(ByVal sender As Object, ByVal e As EventArgs)
        ' Saving the 'Selected Date on Calendar' into DataGridView current cell
        If oDateTimePicker.Checked = True Then
            tgvCheckNeeded.CurrentCell.Value = oDateTimePicker.Value.ToString(System.Globalization.CultureInfo.CurrentCulture.DateTimeFormat.ShortDatePattern())
        Else
            tgvCheckNeeded.CurrentCell.Value = ""
        End If
    End Sub

    Private Sub iPropDateTimePicker_CloseUp(ByVal sender As Object, ByVal e As EventArgs)
        ' Hiding the control after use 
        oDateTimePicker.Visible = False
    End Sub
    Private Sub ApplyRowValuesToAllRowsToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles ApplyRowValuesToAllRowsToolStripMenuItem.Click
        tgvCheckNeeded.CurrentCell = Nothing
        tgvCheckNeeded.ClearSelection()

        For Column = 0 To tgvCheckNeeded.ColumnCount - 1
            If tgvCheckNeeded.Columns(Column).ReadOnly = False AndAlso
                tgvCheckNeeded.Columns(Column).Name <> "CheckedBy" AndAlso
                tgvCheckNeeded.Columns(Column).Name <> "CheckDate" Then
                For Each Node As TreeGridNode In tgvCheckNeeded.Rows
                    If Node.Visible = True Then Node.Cells(Column).Value = tgvCheckNeeded(Column, SelectedCell.RowIndex).Value
                Next
            End If

        Next

    End Sub
    Private Sub ApplyCellValueToEntireColumnToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles ApplyCellValueToEntireColumnToolStripMenuItem.Click
        If tgvCheckNeeded.Columns(SelectedCell.ColumnIndex).ReadOnly = True Then
            MessageBox.Show("This property cannot be modified")
            Exit Sub
        End If
        Dim Value As String = SelectedCell.Value
        tgvCheckNeeded.CurrentCell = Nothing


        For Each Node As TreeGridNode In tgvCheckNeeded.Rows
            Node.Cells(SelectedCell.ColumnIndex).Value = Value
        Next
    End Sub
    Private Sub tgvCheckNeeded_CellEndEdit(sender As Object, e As DataGridViewCellEventArgs) Handles tgvCheckNeeded.CellEndEdit
        If tgvCheckNeeded.Columns(e.ColumnIndex).Tag = "Number" And ChangeFlag = False Then

            If Not IsNumeric(tgvCheckNeeded(e.ColumnIndex, e.RowIndex).Value) AndAlso tgvCheckNeeded(e.ColumnIndex, e.RowIndex).Value <> Nothing Then
                MessageBox.Show("Column """ & tgvCheckNeeded.Columns(e.ColumnIndex).HeaderText & """ requires a numeric value")
                ChangeFlag = True
                tgvCheckNeeded(e.ColumnIndex, e.RowIndex).Value = Nothing
                ChangeFlag = False

            End If

        End If

    End Sub
    Private Sub CheckNeeded_Resize(sender As Object, e As EventArgs) Handles Me.Resize
        tgvCheckNeeded.Width = Me.Width - 44
        tgvCheckNeeded.Height = Me.Height - 104
        btnOK.Location = New Drawing.Point(btnOK.Location.X, Me.Height - 73)
        btnCancel.Location = New Drawing.Point(btnCancel.Location.X, Me.Height - 73)
        btnHide.Location = New Drawing.Point(btnHide.Location.X, Me.Height - 73)

    End Sub
    Private Sub ClearCellToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles ClearCellToolStripMenuItem.Click
        If tgvCheckNeeded.Columns(SelectedCell.ColumnIndex).ReadOnly = True Then
            MessageBox.Show("This property cannot be modified")
            Exit Sub
        End If
        SelectedCell.Value = Nothing
        tgvCheckNeeded.ClearSelection()
    End Sub
    Private Sub RemoveRevisionToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles RemoveRevisionToolStripMenuItem.Click
        Dim Rev As String = Nothing
        For Each node As TreeGridNode In tgvCheckNeeded.Nodes
            If tgvCheckNeeded(0, SelectedCell.RowIndex).Value = tgvCheckNeeded(0, node.RowIndex).Value Then
                Rev = tgvCheckNeeded(tgvCheckNeeded.Columns(My.Settings.RTSRevCol).Index, node.RowIndex).Value
                For Each childnode As TreeGridNode In node.Nodes
                    If childnode.RowIndex = SelectedCell.RowIndex Then
                        node.Nodes.Remove(childnode)
                        Exit For
                    End If
                Next

                For Column = 1 To tgvCheckNeeded.Columns.Count - 1
                    If tgvCheckNeeded.Columns(Column).HeaderText = My.Settings.RTSRevCol Then
                        For Each childnode As TreeGridNode In node.Nodes
                            If IsNumeric(Rev) Then
                                Rev += 1
                            Else
                                Rev = GetAlphaString(Asc(Rev) - 63)
                            End If
                            childnode.Cells(Column).Value = Rev
                        Next
                    End If

                Next
                Exit For
            End If
        Next
    End Sub
    Private Sub AddRevisionToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles AddRevisionToolStripMenuItem.Click
        For Each node As TreeGridNode In tgvCheckNeeded.Nodes
            Dim NewRev As String
            Dim TotRevs As Integer
            Dim Rev(tgvCheckNeeded.Columns.Count - 1) As String
            If tgvCheckNeeded(0, SelectedCell.RowIndex).Value = tgvCheckNeeded(0, node.RowIndex).Value Then
                Rev(0) = tgvCheckNeeded(0, SelectedCell.RowIndex).Value
                For Column = 1 To tgvCheckNeeded.Columns.Count - 1
                    Select Case tgvCheckNeeded.Columns(Column).HeaderText
                        Case My.Settings.RTSRevCol
                            If My.Settings.RTSRev = True Then
                                For Each objNode In node.Nodes
                                    TotRevs += 1
                                Next
                                If IsNumeric(tgvCheckNeeded(Column, SelectedCell.RowIndex).Value) Then
                                    If TotRevs = 0 Then
                                        NewRev = tgvCheckNeeded(Column, SelectedCell.RowIndex).Value + 1
                                    Else
                                        NewRev = node.Nodes(TotRevs - 1).Cells(Column).Value + 1
                                    End If
                                Else
                                    NewRev = GetAlphaString(TotRevs + 2)
                                End If
                                Rev(Column) = NewRev
                            End If
                        Case My.Settings.RTSNameCol
                            Rev(Column) = _invApp.GeneralOptions.UserName
                        Case My.Settings.RTSDateCol
                            Rev(Column) = Today.ToShortDateString
                    End Select

                Next
                node = node.Nodes.Add(Rev)
                ' node.Expand()
                For Column = 1 To tgvCheckNeeded.Columns.Count - 1
                    If tgvCheckNeeded.Columns(Column).HeaderText = "Checked By" Or
                            tgvCheckNeeded.Columns(Column).HeaderText = "Check Date" Then
                        node.Cells(Column).ReadOnly = True
                        node.Cells(Column).Style.BackColor = Drawing.Color.LightGray
                    End If
                Next
                Exit Sub
            End If
        Next
    End Sub
    Public Function GetAlphaString(AlphaNum As Integer)
        Dim alphaCD1 As String
        Dim alphaCD2 As String
        Dim DV() As String
        Dim TV() As String
        alphaCD1 = Math.Round(Val((AlphaNum - 1) / 26), 3)
        alphaCD2 = Math.Round(Val((AlphaNum - 1) / 676), 3)
        DV = Split(alphaCD1, ".")
        TV = Split(alphaCD2, ".")

        Select Case AlphaNum
            Case 1 To 27 'A to Z
                GetAlphaString = Chr(64 + AlphaNum)
            Case 27 To 702 'AA to ZZ
                GetAlphaString = Chr(64 + DV(0)) & Chr(64 + (AlphaNum - (26 * DV(0))))
            Case 702 To 18278 'AAA to ZZZ
                If TV(0) < 27 Then 'If does not exceeds the limit of 26 (total no. of english alphabets).
                    GetAlphaString = Chr(64 + TV(0)) & Chr(64 + (DV(0) - (26 * TV(0)))) & Chr(64 + (AlphaNum - (26 * DV(0))))
                Else 'If exceeds the limit of 26 (total no. of english alphabets).
                    GetAlphaString = "Z" & "Z" & Chr(64 + (AlphaNum - (26 * DV(0))))
                End If
            Case Else
                Return Nothing
                Exit Function
        End Select
    End Function

    Private Sub tgvCheckNeeded_CellEnter(sender As Object, e As DataGridViewCellEventArgs) Handles tgvCheckNeeded.CellEnter
        Dim DateRow As Boolean = False
        If e.RowIndex = -1 Or e.ColumnIndex <= 0 Then Exit Sub
        If (LCase(CStr(tgvCheckNeeded.Columns(e.ColumnIndex).HeaderText))).Contains("date") Then
            ' If InStr(tgvCheckNeeded(tgvCheckNeeded.Columns("StatusItem").Index, (e.RowIndex)).Value, "Date") <> 0 Then
            DateRow = True
            DateTimePicker(e, DateRow, tgvCheckNeeded)
        Else
            oDateTimePicker.Visible = False
        End If
    End Sub
End Class