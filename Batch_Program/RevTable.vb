Imports System.Windows.Forms
Imports Inventor
Imports System.Runtime.InteropServices

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
    Private Sub lstCheckFiles_Click(sender As Object, e As System.EventArgs)
        RevDtPicker.Visible = False
    End Sub
    Private Sub btnCancel_Click(sender As System.Object, e As System.EventArgs) Handles btnCancel.Click
        lstCheckFiles.Items.Clear()
        Me.Hide()
    End Sub
    Protected Overrides Function ProcessCmdKey(ByRef msg As System.Windows.Forms.Message, ByVal keyData As System.Windows.Forms.Keys) As Boolean
        ' This subroutine closes the text box
        Select Case msg.WParam.ToInt32()
            ' Enter key is pressed
            Case CInt(Windows.Forms.Keys.Enter)
                ' editing completed
                bCancelEdit = False
                RevTxtBox.Hide()
                ' Escape key is pressed
            Case CInt(Windows.Forms.Keys.Escape)
                ' editing was cancelled
                bCancelEdit = True
                RevTxtBox.Hide()
        End Select
        Return Nothing
    End Function
    Private Sub TextBox1_LostFocus(sender As Object, e As System.EventArgs) Handles RevTxtBox.LostFocus
        RevTxtBox.Hide()
        RevDtPicker.Hide()
        If bCancelEdit = False Then
            ' update listviewitem
            CurrentSB.Text = RevTxtBox.Text
        Else
            ' Editing was cancelled by user
            bCancelEdit = False
        End If
        lstCheckFiles.Focus()
    End Sub
    Private Sub DateTimePicker1_ValueChanged(sender As Object, e As System.EventArgs) Handles RevDtPicker.ValueChanged
        If RevDtPicker.Visible = True Then
            'Input the date selected into the cell
            CurrentSB.Text = RevDtPicker.Value.ToString("MM/dd/yy")
            RevDtPicker.Visible = False
        End If
    End Sub
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
        Dim Sheet As Sheet
        Dim Sheets As Sheets
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
        Dim oTitleBlock As TitleBlock
        Dim Rev As String = oDoc.PropertySets.Item("{F29F85E0-4FF9-1068-AB91-08002B27B3D9}").ItemByPropId("9").Value
        Dim X As Integer = 0
        Dim Y As Integer = lstCheckFiles.Items.Count
        Dim Col, Row, I, J As Integer
        Dim Text As String = ""
        If btnIgnore.Visible = False Then
            'Create Loop, first Loop deletes the revtable, Second loop repopulates the rev table
            'Using the data from the userform
            Sheets = oDoc.Sheets
            Sheets.Item(1).Activate()
            Dim S As Integer = 0
            'Delete each rev table in each sheet
            Sheets.Item(1).Activate()
            For Each Sheet In Sheets
                Sheet.Activate()

                oRevtable = oDoc.activesheet.RevisionTables(1)
                oPointX(S) = oRevtable.Position.X
                oPointY(S) = oRevtable.Position.Y
                S += 1
            Next
            S = 0
            For Each Sheet In Sheets
                If S = 0 Then

                    Sheet.Activate()
                    oRevtable = oDoc.activesheet.RevisionTables(1)
                    oRevtable.RevisionTableRows.Add()
                    '
                    'THIS SECTION IS FOR FUTURE DEVELOPMENT
                    'TRYING TO REMOVE REVISIONS WITHOUT DELETING REVISION TABLE
                    '
                    'For R = 1 To oRevtable.RevisionTableRows.Count - 1
                    '    For Y = 1 To lstCheckFiles.Items.Count - 1
                    '        Text = oRevtable.RevisionTableRows.Item(Y).Item(R).Text
                    '        Text = oRevtable.RevisionTableRows.Item(Y).Item(R + 1).Text
                    '        If oRevtable.RevisionTableRows.Item(Y).Item(1).Text = lstCheckFiles.Items(Y).SubItems(0).Text _
                    '            And oRevtable.RevisionTableRows.Item(Y).Item(3).Text = lstCheckFiles.Items(Y).SubItems(2).Text Then
                    '        Else
                    '            oRevtable.RevisionTableRows.Item(Y).Delete()
                    '        End If
                    '    Next
                    For Each oRow As RevisionTableRow In oRevtable.RevisionTableRows
                        If oRow.IsActiveRow Then
                        Else
                            oRow.Delete()
                        End If
                    Next
                    oTitleBlock = oDoc.activesheet.TitleBlock
                    If S = 0 Then
                        Adjust = oPointY(S) - oRevtable.Position.Y
                    End If

                    oRevtable.Delete()
                    S += 1
                ElseIf S > 0 Then
                    Sheet.Activate()
                    oRevtable = oDoc.activesheet.RevisionTables(1)
                    oRevtable.Delete()
                End If
            Next Sheet
            S = 0
            Sheets.Item(1).Activate()
            For Each Sheet In Sheets
                Sheet.Activate()
                'Set counter to numeric or alpha based on previous RevTable
                Point = _invApp.TransientGeometry.CreatePoint2d(oPointX(S), oPointY(S) - Adjust)

                If IsNumeric(Rev) Then
                    oRevtable = oDoc.ActiveSheet.RevisionTables.Add2(Point, False, True, False, 0)
                Else
                    oRevtable = oDoc.ActiveSheet.RevisionTables.Add2(Point, False, True, True, "A")
                End If
                'Add items corresponding to which table form is being populated
                Y = lstCheckFiles.Items.Count
                S += 1
            Next sheet
            If oRevtable.RevisionTableRows.Count < Y Then
                    Do Until oRevtable.RevisionTableRows.Count = Y
                        oRevtable.RevisionTableRows.Add()
                    Loop
                End If
        End If


        oDoc = _invApp.ActiveDocument
        'Path = _invApp.Documents
        'For J = 1 To Path.Count
        'oDoc = Path.Item(J)
        'strPath = oDoc.FullDocumentName
        'strFile = Strings.Right(strPath, Strings.Len(strPath) - InStrRev(strPath, "\"))
        'If strPath IsNot Nothing Then
        'If My.Computer.FileSystem.FileExists(Strings.Left(strPath, (Strings.Len(strPath) - 3)) & "idw") Then
        'If _invApp.ActiveStrings.Right(Document.FullDocumentName, len(Document.FullDocumentName)-InStrRev(Document.FullDocumentName, "\")) = strFile Then
        'Main.MatchDrawing(DrawSource, DrawingName, X)
        'oDoc = _invApp.Documents.Open(Strings.Left(strPath, (Strings.Len(strPath) - 3)) & "idw", True)
        'oDoc = _invApp.Documents.Open(DrawSource, True)
        oDoc.Sheets.Item(1).activate()
        Dim Flag As Boolean = False
        Sheet = oDoc.ActiveSheet
        oRevtable = Sheet.RevisionTables(1)
        Col = oRevtable.RevisionTableColumns.Count
        Row = oRevtable.RevisionTableRows.Count
        Rev = oDoc.PropertySets.Item("{F29F85E0-4FF9-1068-AB91-08002B27B3D9}").ItemByPropId("9").Value
        'Iterate through rev table and populate from the userform
        Dim Contents(0 To Col * Row) As String
        I = 0 : J = 0
        Dim Initials(0 To Row), RevCheckBy(0 To Row), RevDate(0 To Row) As String
        For Each RevRow As RevisionTableRow In oRevtable.RevisionTableRows
            Dim RtCell As RevisionTableCell
            For Each RtCell In RevRow
                'skip over items not used and proceed to next row
                If J Mod 5 = 0 And J > 0 Then
                    J = 0
                    I += 1
                End If
                If J > 0 Then
                    RtCell.Text = lstCheckFiles.Items(I).SubItems(J).Text
                    If Flag = False And RtCell.Text = "" Then Flag = True
                    J += 1
                Else
                    If Strings.Len(lstCheckFiles.Items(I).SubItems(J).Text) = 1 Then
                        J += 1
                    End If
                End If
            Next
        Next
        'Exit For
        'End If
        'End If
        ' Next
        If Flag = False Then
            For X = 0 To CheckNeeded.lstCheckNeeded.Items.Count
                If CheckNeeded.lstCheckNeeded.Items(X).SubItems(0).Text = Strings.Right(oDoc.FullDocumentName, Len(oDoc.FullDocumentName) - InStrRev(oDoc.FullDocumentName, "\")) Then
                    CheckNeeded.lstCheckNeeded.Items(X).Remove()
                    Exit For
                End If
            Next
        End If
        lstCheckFiles.Items.Clear()
        Me.Hide()
        'CheckNeeded.lstCheckNeeded.Items.Clear()

        'CheckNeeded.PopulateCheckNeeded(Path, oDoc, Archive, DrawingName, DrawSource, OpenDocs)
        'CheckNeeded.Hide()
        If CheckNeeded.lstCheckNeeded.Items.Count > 0 Then
            ' CheckNeeded.btnIgnore.Visible = True
            CheckNeeded.btnOK.Visible = False
            CheckNeeded.btnCancel.Location = CheckNeeded.btnOK.Location
            CheckNeeded.btnCancel.Text = "Finished"
            CheckNeeded.Refresh()
            CheckNeeded.Show()
        Else
            CheckNeeded.Hide()
            MsgBox("All drawings have been checked.")
        End If
    End Sub
    Public Sub PopulateRevTable(oDoc As Document, ByRef RevNo As String, ByRef J As Integer, iProp As Boolean)
        If iProp = True Then lstCheckFiles.Clear()
        Dim RevisionTable As RevisionTable
        Dim Col, Row, i As Integer
        Dim Sheet As Sheet
        Dim oPoint As Point2d = _invApp.TransientGeometry.CreatePoint2d(0.965193999999989, 2.7178)
        Dim oRevTable As RevisionTable
        'RevTable.PopMain(Me)
        'Set the active sheet
        Sheet = oDoc.ActiveSheet
        On Error Resume Next
        'Define the revision table from the active sheet
        RevisionTable = Sheet.RevisionTables(1)
        If Err.Number = 5 Then
            oRevTable = oDoc.ActiveSheet.Revisiontables.Add2(oPoint, False, True, False, 0)
            'clear the error for future code.
            Err.Clear()
            Exit Sub
        End If
        'retrieve the data from the revision table
        Col = RevisionTable.RevisionTableColumns.Count
        Row = RevisionTable.RevisionTableRows.Count
        RevNo = oDoc.PropertySets.Item("{F29F85E0-4FF9-1068-AB91-08002B27B3D9}").ItemByPropId("9").Value

        'set parameters for going through the table. Set i for each column
        'set J for each row

        Dim Contents(0 To Col * Row) As String
        i = 0 : J = 1

        'Set variables for initials, Rows and Date for each row of the table
        Dim initials(0 To Row), RevCheckBy(0 To Row), RevDate(0 To Row) As String
        'Iterate through rev table to retrieve data
        For Each RevRow As RevisionTableRow In RevisionTable.RevisionTableRows
            Dim RTCell As RevisionTableCell
            'Iterate through the rev table cell by cell
            For Each RTCell In RevRow
                'save the contents of the cell
                Contents(i) = RTCell.Text
                i += 1
            Next
            'Record whether the revision row is visible or not
            Dim Input(5) As String
            If RevRow.Visible = True Then
                Input(5) = "Visible"
            Else
                Input(5) = "Hidden"
            End If
            'Save each row of data in a listview
            Dim result As ListViewItem
            Input(0) = Contents(((J - 1) * 5) + 0)
            Input(1) = Contents(((J - 1) * 5) + 1)
            Input(2) = Contents(((J - 1) * 5) + 2)
            Input(3) = Contents(((J - 1) * 5) + 3)
            Input(4) = Contents(((J - 1) * 5) + 4)
            'Write the saved row to the next row of the Userform table
            result = New ListViewItem(Input)
            lstCheckFiles.Items.Add(result)
            'Go to the next row
            J += 1
        Next
    End Sub
    Public Sub iPropRevTable(ByRef oDoc As Document)
        'iProperties.PopRevTable(Me)
        Dim input(4) As String
        Dim Result As ListViewItem
        For X As Integer = 0 To iProperties.TabControl1.TabCount - 1
            Select Case X
                Case 0
                    input(0) = iProperties.txtRev1Rev.Text
                    input(1) = iProperties.dtRev1Date.Value.ToString("MM/dd/yy")
                    input(2) = iProperties.txtRev1Des.Text
                    input(3) = iProperties.txtRev1Init.Text
                    input(4) = iProperties.txtRev1Approved.Text
                    Result = New ListViewItem(input)
                    lstCheckFiles.Items.Add(Result)
                Case 1
                    input(0) = iProperties.txtRev2Rev.Text
                    input(1) = iProperties.dtRev2Date.Value.ToString("MM/dd/yy")
                    input(2) = iProperties.txtRev2Des.Text
                    input(3) = iProperties.txtRev2Init.Text
                    input(4) = iProperties.txtRev2Approved.Text
                    Result = New ListViewItem(input)
                    lstCheckFiles.Items.Add(Result)
                Case 2
                    input(0) = iProperties.txtRev3Rev.Text
                    input(1) = iProperties.dtRev3Date.Value.ToString("MM/dd/yy")
                    input(2) = iProperties.txtRev3Des.Text
                    input(3) = iProperties.txtRev3Init.Text
                    input(4) = iProperties.txtRev3Approved.Text
                    Result = New ListViewItem(input)
                    lstCheckFiles.Items.Add(Result)
                Case (3)
                    input(0) = iProperties.txtRev4Rev.Text
                    input(1) = iProperties.dtRev4Date.Value.ToString("MM/dd/yy")
                    input(2) = iProperties.txtRev4Des.Text
                    input(3) = iProperties.txtRev4Init.Text
                    input(4) = iProperties.txtRev4Approved.Text
                    Result = New ListViewItem(input)
                    lstCheckFiles.Items.Add(Result)
                Case 4
                    input(0) = iProperties.txtRev5Rev.Text
                    input(1) = iProperties.dtRev5Date.Value.ToString("MM/dd/yy")
                    input(2) = iProperties.txtRev5Des.Text
                    input(3) = iProperties.txtRev5Init.Text
                    input(4) = iProperties.txtRev5Approved.Text
                    Result = New ListViewItem(input)
                    lstCheckFiles.Items.Add(Result)
                Case 5
                    input(0) = iProperties.txtRev6Rev.Text
                    input(1) = iProperties.dtRev6Date.Value.ToString("MM/dd/yy")
                    input(2) = iProperties.txtRev6Des.Text
                    input(3) = iProperties.txtRev6Init.Text
                    input(4) = iProperties.txtRev6Approved.Text
                    Result = New ListViewItem(input)
                    lstCheckFiles.Items.Add(Result)
                Case 6
                    input(0) = iProperties.txtRev7Rev.Text
                    input(1) = iProperties.dtRev7Date.Value.ToString("MM/dd/yy")
                    input(2) = iProperties.txtRev7Des.Text
                    input(3) = iProperties.txtRev7Init.Text
                    input(4) = iProperties.txtRev7Approved.Text
                    Result = New ListViewItem(input)
                    lstCheckFiles.Items.Add(Result)
                Case 7
                    input(0) = iProperties.txtRev8Rev.Text
                    input(1) = iProperties.dtRev8Date.Value.ToString("MM/dd/yy")
                    input(2) = iProperties.txtRev8Des.Text
                    input(3) = iProperties.txtRev8Init.Text
                    input(4) = iProperties.txtRev8Approved.Text
                    Result = New ListViewItem(input)
                    lstCheckFiles.Items.Add(Result)
            End Select
        Next
    End Sub
    Private Sub btnIgnore_Click(sender As System.Object, e As System.EventArgs) Handles btnIgnore.Click
        lstCheckFiles.Items.Clear()
        Me.Hide()
        Exit Sub
    End Sub
    Private Sub lstCheckFiles_MouseClick(sender As Object, e As System.Windows.Forms.MouseEventArgs) Handles lstCheckFiles.MouseClick
        RevDtPicker.Visible = False
        Dim Ans, Rev As String
        Dim Index As Integer

        If btnIgnore.Visible = False Then
            'Keep the base revision locked
            If lstCheckFiles.SelectedIndices(0) = 0 Then
                MsgBox("You cannot remove the base revision")
            Else
                'Record the row selected
                Index = lstCheckFiles.SelectedIndices(0)
                'confirm delete
                Ans = MsgBox("Do you wish to remove this revision?" & vbNewLine & "All revision tags will be removed as well.", vbYesNo, "Remove Revision")
                If Ans = vbYes Then
                    'remove the selected row
                    For X As Integer = lstCheckFiles.SelectedItems.Count - 1 To 0
                        lstCheckFiles.SelectedItems(X).Remove()
                    Next
                    'Test if the rev is numeric or alpha
                    Rev = lstCheckFiles.Items(0).SubItems(0).Text
                    If IsNumeric(Rev) Then
                        'for numeric, renumber the revs because one was removed
                        For X = 0 To lstCheckFiles.Items.Count - 1
                            lstCheckFiles.Items(X).SubItems(0).Text = X
                        Next
                    Else
                        'for alpha, redo the revs alphabetically because one was removed
                        For X = 0 To lstCheckFiles.Items.Count - 1
                            lstCheckFiles.Items(X).SubItems(0).Text = Chr(65 + X)
                        Next
                    End If

                End If
            End If
        Else
            RevDtPicker.Visible = False
            ' See which column has been clicked
            CurrentItem = lstCheckFiles.GetItemAt(e.X, e.Y)     ' which listviewitem was clicked
            If CurrentItem Is Nothing Then Exit Sub
            CurrentSB = CurrentItem.GetSubItemAt(e.X, e.Y)  ' which subitem was clicked
            ' NOTE: This portion is important. Here you can define your own
            '       rules as to which column can be edited and which cannot.
            Dim iSubIndex As Integer = CurrentItem.SubItems.IndexOf(CurrentSB)
            If CurrentItem.Text = "A" Or CurrentItem.Text = "0" Then
                If iSubIndex = 1 Or iSubIndex = 2 Then Exit Sub
            End If
            Select Case iSubIndex
                Case 2, 3, 4
                    RevDtPicker.Visible = False
                    'Create textbox for text editable items       
                    Dim lLeft = CurrentSB.Bounds.Left + 2
                    Dim lWidth As Integer = CurrentSB.Bounds.Width
                    With RevTxtBox
                        .SetBounds(lLeft + lstCheckFiles.Left, CurrentSB.Bounds.Top +
                                   lstCheckFiles.Top, lWidth, CurrentSB.Bounds.Height)
                        .Text = CurrentSB.Text
                        .Show()
                        .Focus()
                    End With
                Case 1
                    'If CurrentSB = 1 Then Exit Sub
                    'Create Date Picker for Date column
                    Dim lLeft = CurrentSB.Bounds.Left + 2
                    Dim lWidth As Integer = CurrentSB.Bounds.Width
                    With RevDtPicker

                        .SetBounds(lLeft + lstCheckFiles.Left, CurrentSB.Bounds.Top + lstCheckFiles.Top, lWidth, CurrentSB.Bounds.Height)
                        .Show()
                        .Focus()
                        .Visible = True
                    End With
                    'Expand datepicker
                    'Windows.Forms.SendKeys.Send("{F4}")
                    RevDtPicker.Checked = False
                Case 0, 5
                    Exit Sub
                Case Else
                    ' Exit for un-editable items
                    Exit Sub
            End Select
            End If
    End Sub
End Class