Imports System.Windows.Forms
Imports Inventor
Imports System.Runtime.InteropServices
Imports System.Globalization

Public Class RevTable
    'Create global references for RevTable form
    Dim bCancelEdit As Boolean
    Dim _invApp As Inventor.Application
    Dim CurrentItem As Windows.Forms.ListViewItem
    Dim CurrentSB As Windows.Forms.ListViewItem.ListViewSubItem
    Dim iProperties As iProperties
    Dim Main As Main
    Dim CheckNeeded As CheckNeeded
    Dim RevTable As RevTable
    Public Sub New()
        ' This call is required by the designer.
        InitializeComponent()
        _invApp = Marshal.GetActiveObject("Inventor.Application")

        ' Add any initialization after the InitializeComponent() call.
    End Sub
    Public Function PopMain(CalledFunction As Main)
        Main = CalledFunction
        Return Nothing
    End Function
    Public Function PopiProperties(CalledFunction As iProperties)
        iProperties = CalledFunction
        Return Nothing
    End Function
    Public Function PopCheckNeeded(CalledFunction As CheckNeeded)
        CheckNeeded = CalledFunction
        Return Nothing
    End Function
    Public Sub CreateOpenDocs(OpenDocs As ArrayList)
        Dim Archive, DocSource, DocName As String
        For Each oDoc In _invApp.Documents.VisibleDocuments
            Archive = oDoc.FullDocumentName
            'Use the Partsource file to create the drawingsource file
            DocSource = Strings.Left(Archive, Strings.Len(Archive))
            DocName = Strings.Right(DocSource, Strings.Len(DocSource) - Strings.InStrRev(DocSource, "\"))
            OpenDocs.Add(DocName)
        Next
    End Sub
    Private Sub btnOK_Click(sender As System.Object, e As System.EventArgs) Handles btnOK.Click
        CheckNeeded.PopRevTable(Me)
        Dim oDoc As Document = _invApp.ActiveDocument
        Dim Archive As String = ""
        Dim DrawingName As String = ""
        Dim DrawSource As String = ""
        Dim Adjust As String = ""
        Dim OpenDocs As New ArrayList
        CreateOpenDocs(OpenDocs)
        Dim oPointY(0 To oDoc.sheets.count) As String
        Dim oPointX(0 To oDoc.sheets.count) As String
        Dim Point As Point2d = Nothing
        Dim oRevtable As RevisionTable = Nothing
        Dim Rev As String = oDoc.PropertySets.Item("{F29F85E0-4FF9-1068-AB91-08002B27B3D9}").ItemByPropId("9").Value
        Dim X As Integer = 0
        'Dim Y As Integer = lstCheckFiles.Items.Count
        'Dim Col, Row, I, J As Integer
        'Dim Text As String = ""
        'If btnIgnore.Visible = False Then
        '    'Create Loop, first Loop deletes the revtable, Second loop repopulates the rev table
        '    'Using the data from the userform
        '    Sheets = oDoc.Sheets
        '    Sheets.Item(1).Activate()
        '    Dim S As Integer = 0
        '    'Delete each rev table in each sheet
        '    Sheets.Item(1).Activate()
        '    For Each Sheet In Sheets
        '        Sheet.Activate()

        '        oRevtable = oDoc.activesheet.RevisionTables(1)
        '        oPointX(S) = oRevtable.Position.X
        '        oPointY(S) = oRevtable.Position.Y
        '        S += 1
        '    Next
        '    S = 0
        '    For Each Sheet In Sheets
        '        If S = 0 Then

        '            Sheet.Activate()
        '            oRevtable = oDoc.activesheet.RevisionTables(1)
        '            oRevtable.RevisionTableRows.Add()
        '            '
        '            'THIS SECTION IS FOR FUTURE DEVELOPMENT
        '            'TRYING TO REMOVE REVISIONS WITHOUT DELETING REVISION TABLE
        '            '
        '            'For R = 1 To oRevtable.RevisionTableRows.Count - 1
        '            '    For Y = 1 To lstCheckFiles.Items.Count - 1
        '            '        Text = oRevtable.RevisionTableRows.Item(Y).Item(R).Text
        '            '        Text = oRevtable.RevisionTableRows.Item(Y).Item(R + 1).Text
        '            '        If oRevtable.RevisionTableRows.Item(Y).Item(1).Text = lstCheckFiles.Items(Y).SubItems(0).Text _
        '            '            And oRevtable.RevisionTableRows.Item(Y).Item(3).Text = lstCheckFiles.Items(Y).SubItems(2).Text Then
        '            '        Else
        '            '            oRevtable.RevisionTableRows.Item(Y).Delete()
        '            '        End If
        '            '    Next
        '            For Each oRow As RevisionTableRow In oRevtable.RevisionTableRows
        '                If oRow.IsActiveRow Then
        '                Else
        '                    oRow.Delete()
        '                End If
        '            Next
        '            oTitleBlock = oDoc.activesheet.TitleBlock
        '            If S = 0 Then
        '                Adjust = oPointY(S) - oRevtable.Position.Y
        '            End If

        '            oRevtable.Delete()
        '            S += 1
        '        ElseIf S > 0 Then
        '            Sheet.Activate()
        '            oRevtable = oDoc.activesheet.RevisionTables(1)
        '            oRevtable.Delete()
        '        End If
        '    Next Sheet
        '    S = 0
        '    Sheets.Item(1).Activate()
        '    For Each Sheet In Sheets
        '        Sheet.Activate()
        '        'Set counter to numeric or alpha based on previous RevTable
        '        Point = _invApp.TransientGeometry.CreatePoint2d(oPointX(S), oPointY(S) - Adjust)

        '        If IsNumeric(Rev) Then
        '            oRevtable = oDoc.ActiveSheet.RevisionTables.Add2(Point, False, True, False, 0)
        '        Else
        '            oRevtable = oDoc.ActiveSheet.RevisionTables.Add2(Point, False, True, True, "A")
        '        End If
        '        'Add items corresponding to which table form is being populated
        '        Y = lstCheckFiles.Items.Count
        '        S += 1
        '    Next Sheet
        '    If oRevtable.RevisionTableRows.Count < Y Then
        '        Do Until oRevtable.RevisionTableRows.Count = Y
        '            oRevtable.RevisionTableRows.Add()
        '        Loop
        '    End If
        'End If


        'oDoc = _invApp.ActiveDocument

        'oDoc.Sheets.Item(1).activate()
        'Dim Flag As Boolean = False
        'Sheet = oDoc.ActiveSheet
        'oRevtable = Sheet.RevisionTables(1)
        'Col = oRevtable.RevisionTableColumns.Count
        'Row = oRevtable.RevisionTableRows.Count
        'Rev = oDoc.PropertySets.Item("{F29F85E0-4FF9-1068-AB91-08002B27B3D9}").ItemByPropId("9").Value
        ''Iterate through rev table and populate from the userform
        'Dim Contents(0 To Col * Row) As String
        'I = 0 : J = 0
        'Dim Initials(0 To Row), RevCheckBy(0 To Row), RevDate(0 To Row) As String
        'For Each RevRow As RevisionTableRow In oRevtable.RevisionTableRows
        '    Dim RtCell As RevisionTableCell
        '    For Each RtCell In RevRow
        '        'skip over items not used and proceed to next row
        '        If J Mod 5 = 0 And J > 0 Then
        '            J = 0
        '            I += 1
        '        End If
        '        If J > 0 Then
        '            RtCell.Text = lstCheckFiles.Items(I).SubItems(J).Text
        '            If Flag = False And RtCell.Text = "" Then Flag = True
        '            J += 1
        '        Else
        '            If Strings.Len(lstCheckFiles.Items(I).SubItems(J).Text) = 1 Then
        '                J += 1
        '            End If
        '        End If
        '    Next
        'Next
        'Exit For
        'End If
        'End If
        ' Next
        'If Flag = False Then
        '    For X = 0 To CheckNeeded.lstCheckNeeded.Items.Count
        '        If CheckNeeded.lstCheckNeeded.Items(X).SubItems(0).Text = Strings.Right(oDoc.FullDocumentName, Len(oDoc.FullDocumentName) - InStrRev(oDoc.FullDocumentName, "\")) Then
        '            CheckNeeded.lstCheckNeeded.Items(X).Remove()
        '            Exit For
        '        End If
        '    Next
        'End If
        ' lstCheckFiles.Items.Clear()
        Me.Hide()
        'CheckNeeded.lstCheckNeeded.Items.Clear()

        'CheckNeeded.PopulateCheckNeeded(Path, oDoc, Archive, DrawingName, DrawSource, OpenDocs)
        'CheckNeeded.Hide()
        'If CheckNeeded.lstCheckNeeded.Items.Count > 0 Then
        '    ' CheckNeeded.btnIgnore.Visible = True
        '    CheckNeeded.btnOK.Visible = False
        '    CheckNeeded.btnCancel.Location = CheckNeeded.btnOK.Location
        '    CheckNeeded.btnCancel.Text = "Finished"
        '    CheckNeeded.Refresh()
        '    CheckNeeded.Show()
        'Else
        '    CheckNeeded.Hide()
        '    MsgBox("All drawings have been checked.")
        'End If
    End Sub

    Public Sub PopulateRevTable(oDoc As Document, ByRef RevNo As String, iProp As Boolean)
        If iProp = True Then dgvRevTable.Rows.Clear()
        Dim RevisionTable As RevisionTable
        Dim Sheet As Sheet
        Dim oPoint As Point2d
        Dim oRevTable As RevisionTable
        'RevTable.PopMain(Me)
        'Set the active sheet
        Sheet = oDoc.ActiveSheet
        Dim oBorder As Border = Sheet.Border
        If Not oBorder Is Nothing Then
            oPoint = _invApp.TransientGeometry.CreatePoint2d(0.965193999999989, 2.7178)
        Else
            oPoint = _invApp.TransientGeometry.CreatePoint2d(Sheet.Width, Sheet.Height)
        End If
        On Error Resume Next
        'Define the revision table from the active sheet
        RevisionTable = Sheet.RevisionTables(1)
        If Err.Number = 5 Then
            oRevTable = oDoc.ActiveSheet.Revisiontables.Add2(oPoint, False, True, False, 0)
            Main.RevTable_Location(Sheet, oRevTable, oPoint)
            oRevTable.Position = oPoint
            'clear the error for future code.
            Err.Clear()
            Exit Sub
        End If
        'retrieve the data from the revision table
        RevNo = oDoc.PropertySets.Item("{F29F85E0-4FF9-1068-AB91-08002B27B3D9}").ItemByPropId("9").Value
        Dim tg As TransientGeometry
        tg = _invApp.TransientGeometry
        Dim dd As DrawingDocument
        dd = _invApp.ActiveDocument
        Dim s As Sheet
        s = dd.ActiveSheet
        ' Get revision table
        Dim rt As RevisionTable
        rt = s.RevisionTables(1)
        ' Get dimensions
        Dim c As Integer
        Dim r As Integer
        c = rt.RevisionTableColumns.Count
        r = rt.RevisionTableRows.Count
        ' Counter
        Dim i As Integer, j As Integer
        ' Get headers and column widths
        Dim rtc As RevisionTableColumn = Nothing
        For Each rtc In rt.RevisionTableColumns
            Dim h As New DataGridViewTextBoxColumn With
            {.Name = CultureInfo.CurrentCulture.TextInfo.ToTitleCase(LCase(rtc.Title)),
            .HeaderText = CultureInfo.CurrentCulture.TextInfo.ToTitleCase(LCase(rtc.Title)),
            .SortMode = DataGridViewColumnSortMode.NotSortable,
            .ReadOnly = True}
            h.Name = CultureInfo.CurrentCulture.TextInfo.ToTitleCase(LCase(rtc.Title))
            dgvRevTable.Columns.Add(h)
        Next
        Dim Vis As New DataGridViewTextBoxColumn With
            {.Name = "Visible",
            .HeaderText = "Visible",
            .SortMode = DataGridViewColumnSortMode.NotSortable,
            .ReadOnly = True}
        Vis.Name = CultureInfo.CurrentCulture.TextInfo.ToTitleCase(LCase(rtc.Title))
        dgvRevTable.Columns.Add(Vis)
        ' Get contents and row heights
        Dim contents(c * (r + 1)) As String
        Dim heights(r) As Double
        Dim rtr As RevisionTableRow
        i = 1 : j = 1
        For Each rtr In rt.RevisionTableRows
            Dim rtcell As RevisionTableCell
            Dim Input(c + 1) As String
            Dim x As Integer = 0
            For Each rtcell In rtr
                contents(i) = rtcell.Text
                Input(x) = contents(i)
                i += 1
                x += 1
            Next
            'Write the saved row to the next row of the Userform table
            If rtr.Visible = True Then
                Input(x) = "Visible"
            Else
                Input(x) = "Hidden"
            End If
            Dim result As DataGridViewRow
            result = New DataGridViewRow
            dgvRevTable.Rows.Add(Input)

            heights(j) = rtr.Height
            'Save each row of data in a listview

            'Go to the next row
            j = j + 1
        Next

        Dim dgvColWidth As Double = Nothing
        For Each column In dgvRevTable.Columns
            dgvColWidth = column.width + dgvColWidth
        Next
        If dgvColWidth < dgvRevTable.Width Then

            dgvRevTable.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill
            dgvRevTable.AutoResizeColumns()
        Else
            dgvRevTable.AutoSizeColumnsMode = DataGridViewAutoSizeColumnMode.DisplayedCells
            dgvRevTable.AutoResizeColumns()
        End If
    End Sub
    Public Sub iPropRevTable(ByRef oDoc As Document)
        ''iProperties.PopRevTable(Me)
        'Dim input(4) As String
        'Dim Result As DataGridViewRow
        'For X As Integer = 0 To iProperties.TabControl1.TabCount - 1
        '    Select Case X
        '        Case 0
        '            input(0) = iProperties.txtRev1Rev.Text
        '            input(1) = iProperties.dtRev1Date.Value.ToString("MM/dd/yy")
        '            input(2) = iProperties.txtRev1Des.Text
        '            input(3) = iProperties.txtRev1Init.Text
        '            input(4) = iProperties.txtRev1Approved.Text
        '        Case 1
        '            input(0) = iProperties.txtRev2Rev.Text
        '            input(1) = iProperties.dtRev2Date.Value.ToString("MM/dd/yy")
        '            input(2) = iProperties.txtRev2Des.Text
        '            input(3) = iProperties.txtRev2Init.Text
        '            input(4) = iProperties.txtRev2Approved.Text

        '        Case 2
        '            input(0) = iProperties.txtRev3Rev.Text
        '            input(1) = iProperties.dtRev3Date.Value.ToString("MM/dd/yy")
        '            input(2) = iProperties.txtRev3Des.Text
        '            input(3) = iProperties.txtRev3Init.Text
        '            input(4) = iProperties.txtRev3Approved.Text
        '        Case (3)
        '            input(0) = iProperties.txtRev4Rev.Text
        '            input(1) = iProperties.dtRev4Date.Value.ToString("MM/dd/yy")
        '            input(2) = iProperties.txtRev4Des.Text
        '            input(3) = iProperties.txtRev4Init.Text
        '            input(4) = iProperties.txtRev4Approved.Text
        '        Case 4
        '            input(0) = iProperties.txtRev5Rev.Text
        '            input(1) = iProperties.dtRev5Date.Value.ToString("MM/dd/yy")
        '            input(2) = iProperties.txtRev5Des.Text
        '            input(3) = iProperties.txtRev5Init.Text
        '            input(4) = iProperties.txtRev5Approved.Text
        '        Case 5
        '            input(0) = iProperties.txtRev6Rev.Text
        '            input(1) = iProperties.dtRev6Date.Value.ToString("MM/dd/yy")
        '            input(2) = iProperties.txtRev6Des.Text
        '            input(3) = iProperties.txtRev6Init.Text
        '            input(4) = iProperties.txtRev6Approved.Text
        '        Case 6
        '            input(0) = iProperties.txtRev7Rev.Text
        '            input(1) = iProperties.dtRev7Date.Value.ToString("MM/dd/yy")
        '            input(2) = iProperties.txtRev7Des.Text
        '            input(3) = iProperties.txtRev7Init.Text
        '            input(4) = iProperties.txtRev7Approved.Text

        '        Case 7
        '            input(0) = iProperties.txtRev8Rev.Text
        '            input(1) = iProperties.dtRev8Date.Value.ToString("MM/dd/yy")
        '            input(2) = iProperties.txtRev8Des.Text
        '            input(3) = iProperties.txtRev8Init.Text
        '            input(4) = iProperties.txtRev8Approved.Text
        '    End Select
        '    Result = New DataGridViewRow
        '    dgvRevTable.Rows.Add(input)
        'Next
    End Sub
    'Private Sub btnIgnore_Click(sender As System.Object, e As System.EventArgs) Handles btnIgnore.Click
    '    lstCheckFiles.Items.Clear()
    '    Me.Hide()
    '    Exit Sub
    'End Sub
    'Private Sub lstCheckFiles_MouseClick(sender As Object, e As System.Windows.Forms.MouseEventArgs)
    '    RevDtPicker.Visible = False
    '    Dim Ans, Rev As String
    '    Dim Index As Integer

    '    If btnIgnore.Visible = False Then
    '        'Keep the base revision locked
    '        If lstCheckFiles.SelectedIndices(0) = 0 Then
    '            MsgBox("You cannot remove the base revision")
    '        Else
    '            'Record the row selected
    '            Index = lstCheckFiles.SelectedIndices(0)
    '            'confirm delete
    '            Ans = MsgBox("Do you wish to remove this revision?" & vbNewLine & "All revision tags will be removed as well.", vbYesNo, "Remove Revision")
    '            If Ans = vbYes Then
    '                'remove the selected row
    '                For X As Integer = lstCheckFiles.SelectedItems.Count - 1 To 0
    '                    lstCheckFiles.SelectedItems(X).Remove()
    '                Next
    '                'Test if the rev is numeric or alpha
    '                Rev = lstCheckFiles.Items(0).SubItems(0).Text
    '                If IsNumeric(Rev) Then
    '                    'for numeric, renumber the revs because one was removed
    '                    For X = 0 To lstCheckFiles.Items.Count - 1
    '                        lstCheckFiles.Items(X).SubItems(0).Text = X
    '                    Next
    '                Else
    '                    'for alpha, redo the revs alphabetically because one was removed
    '                    For X = 0 To lstCheckFiles.Items.Count - 1
    '                        lstCheckFiles.Items(X).SubItems(0).Text = Chr(65 + X)
    '                    Next
    '                End If

    '            End If
    '        End If
    '    Else
    '        RevDtPicker.Visible = False
    '        ' See which column has been clicked
    '        CurrentItem = lstCheckFiles.GetItemAt(e.X, e.Y)     ' which listviewitem was clicked
    '        If CurrentItem Is Nothing Then Exit Sub
    '        CurrentSB = CurrentItem.GetSubItemAt(e.X, e.Y)  ' which subitem was clicked
    '        ' NOTE: This portion is important. Here you can define your own
    '        '       rules as to which column can be edited and which cannot.
    '        Dim iSubIndex As Integer = CurrentItem.SubItems.IndexOf(CurrentSB)
    '        If CurrentItem.Text = "A" Or CurrentItem.Text = "0" Then
    '            If iSubIndex = 1 Or iSubIndex = 2 Then Exit Sub
    '        End If
    '        Select Case iSubIndex
    '            Case 2, 3, 4
    '                RevDtPicker.Visible = False
    '                'Create textbox for text editable items       
    '                Dim lLeft = CurrentSB.Bounds.Left + 2
    '                Dim lWidth As Integer = CurrentSB.Bounds.Width
    '                With RevTxtBox
    '                    .SetBounds(lLeft + lstCheckFiles.Left, CurrentSB.Bounds.Top +
    '                               lstCheckFiles.Top, lWidth, CurrentSB.Bounds.Height)
    '                    .Text = CurrentSB.Text
    '                    .Show()
    '                    .Focus()
    '                End With
    '            Case 1
    '                'If CurrentSB = 1 Then Exit Sub
    '                'Create Date Picker for Date column
    '                Dim lLeft = CurrentSB.Bounds.Left + 2
    '                Dim lWidth As Integer = CurrentSB.Bounds.Width
    '                With RevDtPicker

    '                    .SetBounds(lLeft + lstCheckFiles.Left, CurrentSB.Bounds.Top + lstCheckFiles.Top, lWidth, CurrentSB.Bounds.Height)
    '                    .Show()
    '                    .Focus()
    '                    .Visible = True
    '                End With
    '                'Expand datepicker
    '                'Windows.Forms.SendKeys.Send("{F4}")
    '                RevDtPicker.Checked = False
    '            Case 0, 5
    '                Exit Sub
    '            Case Else
    '                ' Exit for un-editable items
    '                Exit Sub
    '        End Select
    '    End If
    'End Sub

    'Private Sub dgvCheckNeeded_CellClick(sender As Object, e As DataGridViewCellEventArgs) Handles dgvRevTable.CellClick
    '    Dim DateCol As Boolean = False
    '    If e.RowIndex = -1 Or e.ColumnIndex = -1 Then Exit Sub
    '    Select Case dgvRevTable.Columns(e.ColumnIndex).Name
    '        Case My.Settings.RTSDateCol
    '            DateCol = True
    '        Case "Check Date"
    '            DateCol = True
    '        Case My.Settings.RTS1
    '            If My.Settings.RTS2 = "Date" Then DateCol = True
    '        Case My.Settings.RTS2
    '            If My.Settings.RTS2 = "Date" Then DateCol = True
    '        Case My.Settings.RTS3
    '            If My.Settings.RTS2 = "Date" Then DateCol = True
    '        Case My.Settings.RTS4
    '            If My.Settings.RTS2 = "Date" Then DateCol = True
    '    End Select

    '    If DateCol = True Then

    '        'Adding DateTimePicker control into DataGridView 
    '        dgvRevTable.Controls.Add(oDateTimePicker)
    '        ' Setting the format (i.e. 2014-10-10)
    '        oDateTimePicker.Format = DateTimePickerFormat.Short
    '        oDateTimePicker.ShowCheckBox = True
    '        ' It returns the retangular area that represents the Display area for a cell
    '        Dim oRectangle As Rectangle = dgvRevTable.GetCellDisplayRectangle(e.ColumnIndex, e.RowIndex, True)
    '        'Setting area for DateTimePicker Control
    '        oDateTimePicker.Size = New Size(oRectangle.Width, oRectangle.Height)
    '        ' Setting Location
    '        oDateTimePicker.Location = New Drawing.Point(oRectangle.X, oRectangle.Y)
    '        ' An event attached to dateTimePicker Control which is fired when DateTimeControl is closed
    '        AddHandler oDateTimePicker.CloseUp, AddressOf Me.oDateTimePicker_CloseUp
    '        ' An event attached to dateTimePicker Control which is fired when any date is selected
    '        AddHandler oDateTimePicker.TextChanged, AddressOf Me.dateTimePicker_OnTextChange
    '        ' Now make it visible
    '        oDateTimePicker.Visible = True
    '    Else
    '        oDateTimePicker.Visible = False
    '    End If
    'End Sub

    'Private Sub dateTimePicker_OnTextChange(ByVal sender As Object, ByVal e As EventArgs)
    '    ' Saving the 'Selected Date on Calendar' into DataGridView current cell
    '    If oDateTimePicker.Checked = True Then
    '        dgvRevTable.CurrentCell.Value = oDateTimePicker.Value.ToString(System.Globalization.CultureInfo.CurrentCulture.DateTimeFormat.ShortDatePattern())
    '    Else
    '        dgvRevTable.CurrentCell.Value = ""
    '    End If
    'End Sub

    'Private Sub oDateTimePicker_CloseUp(ByVal sender As Object, ByVal e As EventArgs)
    '    ' Hiding the control after use 
    '    oDateTimePicker.Visible = False
    'End Sub

    'Private Sub dgvCheckNeeded_MouseUp(sender As Object, e As MouseEventArgs) Handles dgvRevTable.MouseUp
    '    If dgvRevTable(dgvRevTable.Columns("MoreRevs").Index, dgvRevTable.CurrentRow.Index).Value = False Then
    '        cmsApplyValues.Items.Item(2).Enabled = False
    '    Else
    '        cmsApplyValues.Items.Item(2).Enabled = True
    '    End If
    '    If e.Button = MouseButtons.Right And dgvRevTable.RowCount > 1 Then
    '        cmsApplyValues.Show(Cursor.Position)
    '    ElseIf e.Button = MouseButtons.Right And dgvRevTable.RowCount = 1 And dgvRevTable(dgvRevTable.Columns("MoreRevs").Index, dgvRevTable.CurrentRow.Index).Value = True Then
    '        cmsApplyValues.Items.Item(0).Enabled = False
    '        cmsApplyValues.Items.Item(1).Enabled = False
    '        cmsApplyValues.Show(Cursor.Position)
    '    End If
    'End Sub
End Class