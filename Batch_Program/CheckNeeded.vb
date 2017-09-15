Imports System.Windows.Forms
Imports Inventor
Imports System.Runtime.InteropServices
Imports System.Collections.Generic
Imports System.Globalization
Public Class CheckNeeded
    Public Declare Function mouse_event Lib "user32" Alias "mouse_event" (ByVal dwFlags As Long, ByVal dx As Long, ByVal dy As Long, ByVal cButtons As Long, ByVal dwExtraInfo As Long)
    'Create global references for RevTable form
    Dim bCancelEdit As Boolean
    Dim _invApp As Inventor.Application
    Dim CurrentItem As Windows.Forms.ListViewItem
    Dim CurrentSB As Windows.Forms.ListViewItem.ListViewSubItem
    Dim Main As Main
    Dim RevTable As New RevTable
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
    Public Function PopRevTable(CalledFunction As RevTable)
        RevTable = CalledFunction
        Return Nothing
    End Function
    Private Sub btnCancel_Click(sender As System.Object, e As System.EventArgs) Handles btnCancel.Click
        Main.chkCheck.CheckState = CheckState.Unchecked
        Main.chkDXF.CheckState = CheckState.Unchecked
        Main.chkPDF.CheckState = CheckState.Unchecked
        Me.Hide()
    End Sub
    Private Sub lstCheckNeeded_MouseClick(sender As Object, e As System.Windows.Forms.MouseEventArgs) Handles lstCheckNeeded.MouseClick
        DateTimePicker1.Visible = False
        TextBox1.Visible = False
        ' See which column has been clicked
        CurrentItem = lstCheckNeeded.GetItemAt(e.X, e.Y)     ' which listviewitem was clicked
        If CurrentItem Is Nothing Or btnCancel.Text = "Finished" Then Exit Sub
        CurrentSB = CurrentItem.GetSubItemAt(e.X, e.Y)  ' which subitem was clicked

        ' NOTE: This portion is important. Here you can define your own
        '       rules as to which column can be edited and which cannot.
        Dim iSubIndex As Integer = CurrentItem.SubItems.IndexOf(CurrentSB)
        Select Case iSubIndex
            Case 1, 3, 4
                btnApplytoall.Visible = False
                DateTimePicker1.Visible = False
                'Create textbox for text editable items       
                Dim lLeft = CurrentSB.Bounds.Left + 2
                Dim lWidth As Integer = CurrentSB.Bounds.Width
                With TextBox1
                    .SetBounds(lLeft + lstCheckNeeded.Left, CurrentSB.Bounds.Top +
                               lstCheckNeeded.Top, lWidth, CurrentSB.Bounds.Height)
                    .Text = CurrentSB.Text
                    .Show()
                    .Focus()
                End With
            Case 2, 5
                btnApplytoall.Visible = False
                'Create Date Picker for Date column
                Dim lLeft = CurrentSB.Bounds.Left + 2
                Dim lWidth As Integer = CurrentSB.Bounds.Width
                With DateTimePicker1

                    .SetBounds(lLeft + lstCheckNeeded.Left, CurrentSB.Bounds.Top + lstCheckNeeded.Top, lWidth, CurrentSB.Bounds.Height)
                    .Show()
                    .Focus()
                    .Visible = True
                End With
                'Expand datepicker
                'Windows.Forms.SendKeys.Send("{F4}")
                DateTimePicker1.Checked = False
            Case 0, 6
                If lstCheckNeeded.Items.Count > 1 Then btnApplytoall.Visible = True
                If btnIgnore.Visible = False Then btnApplytoall.Location = btnIgnore.Location
                Exit Sub
            Case Else
                ' Exit for un-editable items
                Exit Sub
        End Select
    End Sub
    Private Sub lstCheckNeeded_MouseDoubleClick(sender As Object, e As System.Windows.Forms.MouseEventArgs) Handles lstCheckNeeded.MouseDoubleClick
        Dim List As String = ""
        Dim oDoc As Document
        Dim Path As Documents = _invApp.Documents
        Dim TestSB As Windows.Forms.ListViewItem.ListViewSubItem
        Dim DrawingName As String
        Dim RevNo As String = ""
        Dim DrawSource As String = ""
        RevTable.PopCheckNeeded(Me)
        ' which listviewitem was clicked
        CurrentItem = lstCheckNeeded.GetItemAt(e.X, e.Y)
        If CurrentItem Is Nothing Then Exit Sub
        TestSB = CurrentItem.GetSubItemAt("548", e.Y)
        If btnCancel.Text = "Finished" Then
            RevTable.Label1.Text = "The drawing includes the following revisions:"
            RevTable.Label2.Text = "Click to remove a revision"
        End If
        If TestSB.Text = "False" Then
            MsgBox("There are no other revisions for this drawing.")
            Exit Sub
        Else
            RevTable.lstCheckFiles.Items.Clear()
            CurrentItem = lstCheckNeeded.GetItemAt("40", e.Y)
            DrawingName = CurrentItem.GetSubItemAt("40", e.Y).Text
            'Look for selected item
            For X = 0 To Main.LVSubFiles.CheckedItems.Count - 1
                If Main.LVSubFiles.Items(X).Checked = True And DrawingName = Trim(Main.LVSubFiles.Items(X).Text) Then
                    Main.MatchDrawing(DrawSource, DrawingName, X)
                    oDoc = _invApp.Documents.Open(DrawSource, True)
                    If btnIgnore.Visible = False Then
                        RevTable.btnIgnore.Visible = False
                    Else
                        RevTable.btnIgnore.Visible = True
                    End If
                    RevTable.Refresh()
                    'Dim RevTable As New RevTable
                    Call RevTable.PopCheckNeeded(Me)
                    Call RevTable.PopulateRevTable(oDoc, RevNo, 1, False)
                    _invApp.SilentOperation = True
                    Try
                        oDoc.Save()
                    Catch
                    End Try
                    Main.CloseLater(DrawingName, oDoc)
                    _invApp.SilentOperation = False
                    Exit For
                End If
            Next
            'For J As Integer = 1 To _invApp.Documents.Count
            'oDoc = Path.Item(J)
            'Archive = oDoc.FullDocumentName
            'Use the Partsource file to create the drawingsource file
            'If Archive IsNot Nothing Then
            ' DrawSource = Strings.Left(Archive, Strings.Len(Archive))
            'DrawSource = Strings.Left(DrawSource, Len(DrawSource) - 3) & "idw"
            'If DrawingName = Strings.Right(DrawSource, Len(DrawSource) - InStrRev(DrawSource, "\")) Then
            'If My.Computer.FileSystem.FileExists(DrawSource) = True Then
            '
            'Exit For
            'End If
            'End If
            'End If
            'Next
        End If
        'RevTable.btnIgnore.Visible = True
        btnApplytoall.Visible = False
        RevTable.ShowDialog()
    End Sub
    Protected Overrides Function ProcessCmdKey(ByRef msg As System.Windows.Forms.Message, ByVal keyData As System.Windows.Forms.Keys) As Boolean
        ' This subroutine closes the text box
        Select Case msg.WParam.ToInt32()
            ' Enter key is pressed
            Case CInt(Windows.Forms.Keys.Enter)
                ' editing completed
                bCancelEdit = False
                TextBox1.Hide()
                ' Escape key is pressed
            Case CInt(Windows.Forms.Keys.Escape)
                ' editing was cancelled
                bCancelEdit = True
                TextBox1.Hide()
        End Select
        Return Nothing
    End Function
    Private Sub TextBox1_LostFocus(sender As Object, e As System.EventArgs) Handles TextBox1.LostFocus
        TextBox1.Hide()
        DateTimePicker1.Hide()
        If bCancelEdit = False Then
            ' update listviewitem
            CurrentSB.Text = TextBox1.Text
        Else
            ' Editing was cancelled by user
            bCancelEdit = False
        End If
        lstCheckNeeded.Focus()
    End Sub
    Private Sub DateTimePicker1_ValueChanged(sender As Object, e As System.EventArgs) Handles DateTimePicker1.ValueChanged
        If DateTimePicker1.Visible = True Then
            'Input the date selected into the cell
            If DateTimePicker1.Checked = True Then
                CurrentSB.Text = DateTimePicker1.Value.ToString("MM/dd/yy")
            Else
                CurrentSB.Text = ""
                'DateTimePicker1.Visible = False
            End If
        End If
    End Sub
    Private Sub btnIgnore_Click(sender As System.Object, e As System.EventArgs) Handles btnIgnore.Click
        Dim Path As Documents = _invApp.Documents
        Main.chkCheck.CheckState = CheckState.Unchecked
        Me.Hide()
        'Main.ExportCheck(Path, odoc:=Nothing, Archive:="", DrawingName:="", DrawSource:="", ExportType:="DXF")
    End Sub
    Private Sub btnOK_Click(sender As System.Object, e As System.EventArgs) Handles btnOK.Click
        Dim DrawingName, Archive, DrawSource, Rev As String
        Dim Sheet As Sheet
        Dim Col, Row, Ans As Integer
        Dim RevTable As RevisionTable
        Dim Path As Documents = _invApp.Documents
        Dim oDoc As Document = Nothing
        Dim Opendocs As New ArrayList
        Main.CreateOpenDocs(Opendocs)
        ProgressBar1.Visible = True
        'For each drawing name, find the drawing source related to it
        btnApplytoall.Visible = False
        If Me.btnIgnore.Visible = True Then
            For Y = 0 To lstCheckNeeded.Items.Count - 1
                For Z = 0 To lstCheckNeeded.Items(Y).SubItems.Count - 2
                    If lstCheckNeeded.Items(Y).SubItems(Z).Text = "" And Me.btnIgnore.Visible = True And Ans <> vbOK Then
                        Ans = MsgBox("Some categories are missing information" & vbNewLine &
                               "(Make sure you check the 'More' category)" & vbNewLine &
                               "Do you wish to continue?", MsgBoxStyle.OkCancel)
                        If Ans = vbCancel Then
                            'Me.Hide()
                            Exit Sub
                        End If
                    End If
                Next
            Next
        End If
        'Iterate through checkneeded table to get drawing name
        Dim Flag As Boolean = False
        For X As Integer = 0 To lstCheckNeeded.Items.Count - 1
            DrawingName = lstCheckNeeded.Items(X).Text
            For J As Integer = 1 To _invApp.Documents.Count
                oDoc = Path.Item(J)
                Archive = oDoc.FullDocumentName
                DrawSource = Strings.Left(Archive, Strings.Len(Archive) - 3) & "idw"
                If InStr(DrawSource, DrawingName) <> 0 Then
                    'Open the related drawing in the background
                    Try
                        oDoc = _invApp.Documents.Open(DrawSource, False)
                    Catch ex As Exception
                        MsgBox("The drawing could not be found." & vbNewLine & "If the drawing is saved in a different location" & vbNewLine &
                               "than the model, this can cause an error.")
                        Exit Sub
                    End Try


                    oDoc.PropertySets.Item("{32853F0F-3444-11D1-9E93-0060B03C1CA6}").ItemByPropId("10").Value = lstCheckNeeded.Items(X).SubItems(1).Text
                    If lstCheckNeeded.Items(X).SubItems(2).Text = "" Then
                        oDoc.PropertySets.Item("{32853F0F-3444-11D1-9E93-0060B03C1CA6}").ItemByPropId("11").Value = #1/1/1601#
                    Else
                        oDoc.PropertySets.Item("{32853F0F-3444-11D1-9E93-0060B03C1CA6}").ItemByPropId("11").Value = lstCheckNeeded.Items(X).SubItems(2).Text
                    End If
                    Sheet = oDoc.ActiveSheet
                    RevTable = Sheet.RevisionTables(1)
                    Col = RevTable.RevisionTableColumns.Count
                    Row = RevTable.RevisionTableRows.Count
                    Rev = oDoc.PropertySets.Item("{F29F85E0-4FF9-1068-AB91-08002B27B3D9}").ItemByPropId("9").Value
                    'Iterate through rev table and populate from the userform
                    Dim Contents(0 To Col * Row) As String
                    J = 0
                    Dim Initials(0 To Row), RevCheckBy(0 To Row), RevDate(0 To Row) As String
                    For Each RevRow As RevisionTableRow In RevTable.RevisionTableRows
                        Dim RtCell As RevisionTableCell
                        For Each RtCell In RevRow
                            Select Case J
                                Case 1
                                    RtCell.Text = lstCheckNeeded.Items(X).SubItems(5).Text
                                Case 3
                                    RtCell.Text = lstCheckNeeded.Items(X).SubItems(3).Text
                                Case 4
                                    RtCell.Text = lstCheckNeeded.Items(X).SubItems(4).Text
                            End Select
                            J += 1
                        Next
                    Next
                    oDoc.Update()
                    _invApp.SilentOperation = True
                    oDoc.Save()
                    Main.CloseLater(DrawingName, oDoc)
                    _invApp.SilentOperation = False
                    ProgressBar1.Value = ((X + 1) / lstCheckNeeded.Items.Count) * 100
                    ProgressBar1.PerformStep()
                    Exit For
                End If
            Next

        Next
        'Main.chkCheck.CheckState = CheckState.Indeterminate
        ProgressBar1.Visible = False
        Me.Hide()
        If Main.chkCheck.CheckState = CheckState.Indeterminate Then Main.ExportCheck(Path, odoc:=Nothing, Archive:="", DrawingName:="", DrawSource:="", ExportType:="DXF")
    End Sub
    Private Sub lstCheckNeeded_SelectedIndexChanged(sender As System.Object, e As System.EventArgs) Handles lstCheckNeeded.SelectedIndexChanged
        DateTimePicker1.Visible = False
        btnApplytoall.Visible = False
    End Sub
    Private Sub lstCheckNeeded_Click(sender As Object, e As System.EventArgs) Handles lstCheckNeeded.Click
        DateTimePicker1.Visible = False
    End Sub
    Public Sub PopulateCheckNeeded(Path As Documents, ByRef odoc As Document, ByRef Archive As String _
                                 , ByRef DrawingName As String, ByRef DrawSource As String, OpenDocs As ArrayList)
        Dim RevNo, CheckedBy, RevCheckBy, Initials As String
        Initials = ""
        RevCheckBy = ""
        Dim DateChecked As Date
        Dim Input(8) As String
        Dim Sheet As Sheet
        Dim RevisionTable As RevisionTable
        Dim Col, Row, i, k As Integer
        Dim More As Boolean = False
        Main.MsVistaProgressBar1.ProgressBarStyle = MSVistaProgressBar.BarStyle.Continuous
        'Iterate through all files in the subfiles window
        For X = 0 To Main.LVSubFiles.Items.Count - 1
            'get the checkbox state of each item
            If Main.LVSubFiles.Items(X).Checked = True Then
                'retrieve name of selected item

                'iterate through documents in inventor
                'For j = 1 To Path.Count
                'odoc = Path.Item(j)
                'If odoc.FullDocumentName = Nothing Then GoTo Skip
                'Set drawing source for current inventor document
                'DrawSource = Strings.Left(odoc.FullDocumentName, Strings.Len(odoc.FullDocumentName) - 3) & "idw"
                Main.MatchDrawing(DrawSource, DrawingName, X)
                DrawingName = Trim(Main.LVSubFiles.Items.Item(X).Text)
                'strPath = Strings.Right(DrawSource, Strings.Len(DrawSource) - InStrRev(DrawSource, "\"))
                'test to see if the document has the same name as the listed file
                'If InStr(DrawSource, DrawingName) <> 0 Then
                'open the document in the background
                odoc = _invApp.Documents.Open(DrawSource, False)
                'set active sheet and retrieve the values from the revision table
                Sheet = odoc.ActiveSheet
                RevisionTable = Sheet.RevisionTables(1)
                Col = RevisionTable.RevisionTableColumns.Count
                Row = RevisionTable.RevisionTableRows.Count
                CheckedBy = odoc.PropertySets.Item("{32853F0F-3444-11D1-9E93-0060B03C1CA6}").ItemByPropId("10").Value.ToString
                RevNo = odoc.PropertySets.Item("{F29F85E0-4FF9-1068-AB91-08002B27B3D9}").ItemByPropId("9").Value.ToString
                DateChecked = CDate(odoc.PropertySets.Item("{32853F0F-3444-11D1-9E93-0060B03C1CA6}").ItemByPropId("11").Value)

                Dim Contents(0 To Col * Row) As String
                i = 0 : k = 0

                For Each RevRow In RevisionTable.RevisionTableRows
                    For Each RTCell In RevRow
                        Contents(i) = RTCell.Text
                        i += 1
                        'Record the contents of each cell in each row
                    Next
                    If k = 0 Then
                        Input(0) = DrawingName
                        Input(1) = CheckedBy
                        If DateChecked = #1/1/1601# Then
                            Input(2) = ""
                        Else
                            Input(2) = CStr(DateChecked)
                        End If
                        Input(3) = Contents((k * 5) + 3)
                        Input(4) = Contents((k * 5) + 4)
                        Input(5) = Contents((k * 5) + 1)
                        Input(6) = CStr(More)
                        Input(7) = CStr(More)
                    Else
                        If Contents((k * 5) + 3) = "" Or Contents((k * 5) + 4) = "" _
                        Or Contents((k * 5) + 1) = "" Or Contents((k * 5) + 2) = "" Then
                            More = True
                            Input(6) = CStr(More)
                            Input(7) = "Extra"
                            More = False
                            Exit For
                        Else
                            Input(6) = "True"
                            Input(7) = "Extra"
                        End If
                    End If
                    k += 1
                Next
                Dim result As ListViewItem
                result = New ListViewItem(Input)
                If Input(0) = "" Or Input(1) = "" Or Input(2) = "" Or Input(3) = "" Or Input(4) = "" Or Input(5) = "" Or Input(6) = "True" Then
                    lstCheckNeeded.Items.Add(result)
                ElseIf Input(7) = "Extra" And Strings.InStr(Label1.Text, "checked") = 0 Then
                    lstCheckNeeded.Items.Add(result)
                End If
                Main.CloseLater(DrawingName, odoc)
                'Exit For
Skip:
                ' End If
                'Next
            End If
            Main.ProgressBar(Main.LVSubFiles.CheckedItems.Count, X + 1, "Checking Revs: ", DrawingName)
        Next
        Main.MsVistaProgressBar1.Visible = False
    End Sub
    Private Sub btnApplytoall_Click(sender As System.Object, e As System.EventArgs) Handles btnApplytoall.Click
        Dim CheckedBy, RevInit, RevCheckBy, DateChecked, RevDate As String
        CheckedBy = CurrentItem.SubItems(1).Text
        DateChecked = CurrentItem.SubItems(2).Text
        RevInit = CurrentItem.SubItems(3).Text
        RevDate = CurrentItem.SubItems(4).Text
        RevCheckBy = CurrentItem.SubItems(5).Text
        For X = 0 To lstCheckNeeded.Items.Count - 1
            lstCheckNeeded.Items(X).SubItems(1).Text = CheckedBy
            lstCheckNeeded.Items(X).SubItems(2).Text = DateChecked
            lstCheckNeeded.Items(X).SubItems(3).Text = RevInit
            lstCheckNeeded.Items(X).SubItems(4).Text = RevDate
            lstCheckNeeded.Items(X).SubItems(5).Text = RevCheckBy
        Next
    End Sub
    Private Sub TextBox1_PreviewKeyDown(sender As Object, e As PreviewKeyDownEventArgs) Handles TextBox1.PreviewKeyDown
        If e.KeyCode = Keys.Tab Then
            Windows.Forms.Cursor.Position = New System.Drawing.Point(TextBox1.Location.X + 100 + Me.Location.X, TextBox1.Location.Y + Me.Location.Y + 35)
            Dim C As New MouseEventArgs(1048576, 1, TextBox1.Location.X + 100 + Me.Location.X, TextBox1.Location.Y + Me.Location.Y + 35, 0)
            'e = {X = 315 Y = 33 Button = Left {1048576}}
            Call lstCheckNeeded_MouseClick(lstCheckNeeded, C)
        End If
    End Sub
    Private Sub DateTimePicker1_PreviewKeyDown(sender As Object, e As PreviewKeyDownEventArgs) Handles DateTimePicker1.PreviewKeyDown
        If e.KeyCode = Keys.Tab Then
            Dim Loc As Integer = DateTimePicker1.Location.X

        End If
    End Sub
End Class