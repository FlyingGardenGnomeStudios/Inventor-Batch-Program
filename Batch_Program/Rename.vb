Imports System.Drawing
Imports System.Windows.Forms
Imports Inventor
Imports System
Imports System.IO
Imports Microsoft.Office.Interop
Imports System.Runtime.InteropServices
Imports System.IO.Packaging

Public Class Rename
    Private fromIndex As Integer
    Private dragIndex As Integer
    Private dragRect As Drawing.Rectangle
    'Dim _ExcelApp As New Excel.Application
    Dim _invapp As Inventor.Application
    Dim Time As System.DateTime = Now
    Public Sub New()
        ' This call is required by the designer.
        InitializeComponent()
        _invapp = Marshal.GetActiveObject("Inventor.Application")
        ' Add any initialization after the InitializeComponent() call.
    End Sub
    Public Function PopMain(CalledFunction As Main)
        Main = CalledFunction
        Return Nothing
    End Function
    Private Sub DGVRename_DragDrop(sender As Object, e As DragEventArgs) Handles DGVRename.DragDrop
        Dim p As Drawing.Point = DGVRename.PointToClient(New Drawing.Point(e.X, e.Y))
        dragIndex = DGVRename.HitTest(p.X, p.Y).RowIndex
        If (e.Effect = DragDropEffects.Move) Then
            Dim dragRow As DataGridViewRow = e.Data.GetData(GetType(DataGridViewRow))
            DGVRename.Rows.RemoveAt(fromIndex)
            DGVRename.Rows.Insert(dragIndex, dragRow)
        End If
    End Sub
    Private Sub DGVRename_DragOver(ByVal sender As Object,
                                       ByVal e As DragEventArgs) _
                                       Handles DGVRename.DragOver
        e.Effect = DragDropEffects.Move
    End Sub
    Private Sub DGVRename_MouseDown(ByVal sender As Object,
                                    ByVal e As MouseEventArgs) _
                                    Handles DGVRename.MouseDown
        fromIndex = DGVRename.HitTest(e.X, e.Y).RowIndex
        If fromIndex > -1 Then
            Dim dragSize As Drawing.Size = SystemInformation.DragSize
            dragRect = New Drawing.Rectangle(New Drawing.Point(e.X - (dragSize.Width / 2),
                                               e.Y - (dragSize.Height / 2)),
                                     dragSize)
        Else
            dragRect = System.Drawing.Rectangle.Empty
        End If
    End Sub

    Private Sub DGVRename_MouseMove(ByVal sender As Object,
                                        ByVal e As MouseEventArgs) _
                                        Handles DGVRename.MouseMove



        If (e.Button And MouseButtons.Left) = MouseButtons.Left Then
            If (dragRect <> System.Drawing.Rectangle.Empty _
            AndAlso Not dragRect.Contains(e.X, e.Y)) Then
                DGVRename.DoDragDrop(DGVRename.Rows(fromIndex),
                                         DragDropEffects.Move)
            End If
        End If
        'Debug.Print(MousePosition.X - Me.Bounds.X)
        If MousePosition.X - Me.Bounds.X < 990 And MousePosition.X - Me.Bounds.X > 1100 Then
            Cursor.Current = Cursors.Cross
        Else
            Cursor.Current = Cursors.Default
        End If
    End Sub

    Private Sub btnThumbs_Click(sender As Object, e As EventArgs) Handles btnThumbs.Click
        If btnThumbs.Text = "Small Thumbs" Then
            For X = 0 To DGVRename.RowCount - 1
                DGVRename.Rows(X).Height = 22
            Next
            DGVRename.Width = DGVRename.Width - 0
            btnThumbs.Text = "Large Thumbs"
        Else
            For X = 0 To DGVRename.RowCount - 1
                DGVRename.Rows(X).Height = 100
            Next
            DGVRename.Width = DGVRename.Width + 0
            btnThumbs.Text = "Small Thumbs"
        End If
    End Sub
    Private Sub Rename_Resize(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Resize
        Me.DGVRename.Width = Me.Width - 155
        Me.DGVRename.Height = Me.Height - 100
        Me.RenameProgress.Location = New Drawing.Point(DGVRename.Location.X, DGVRename.Location.Y + DGVRename.Height + 5)
        Me.RenameProgress.Width = DGVRename.Width
        Me.Label2.Location = New Drawing.Point(Me.RenameProgress.Location.X + Me.RenameProgress.Width + 5, Me.RenameProgress.Location.Y)
        Me.Remaining.Location = New Drawing.Point(Me.RenameProgress.Location.X + Me.RenameProgress.Width + 5, Me.RenameProgress.Location.Y + 15)
        btnThumbs.Location = New Drawing.Point(Me.Remaining.Location.X + 5, Me.DGVRename.Location.Y)
        btnExcel.Location = New Drawing.Point(Me.btnThumbs.Location.X, Me.btnThumbs.Location.Y + 25)
        btnImport.Location = New Drawing.Point(Me.btnThumbs.Location.X, Me.btnThumbs.Location.Y + 50)
        btnRename.Location = New Drawing.Point(Me.btnThumbs.Location.X, Me.DGVRename.Height)
        chkCCParts.Location = New Drawing.Point(Me.btnThumbs.Location.X, Me.btnThumbs.Location.Y + 75)
        chkPParts.Location = New Drawing.Point(Me.btnThumbs.Location.X, Me.btnThumbs.Location.Y + 105)
        chkDParts.Location = New Drawing.Point(Me.btnThumbs.Location.X, Me.btnThumbs.Location.Y + 135)
        chkStructure.Location = New Drawing.Point(Me.btnThumbs.Location.X, Me.btnRename.Location.Y - 35)
        Dim X As Integer = -25
        For Each column In DGVRename.Columns
            If DGVRename.Columns(column.index).visible = True Then
                If DGVRename.Columns(column.index).headertext = "Reuse" Then Exit For
                X += DGVRename.Columns(column.index).width
                Debug.WriteLine(X)
            Else
            End If
        Next
        chkReuse.Location = New Drawing.Point(X, Me.chkReuse.Location.Y)
    End Sub
    Public Sub btnExcel_Click(sender As Object, e As EventArgs) Handles btnExcel.Click
        RenameProgress.Visible = True
        Label2.Visible = True
        Remaining.Visible = True
        Dim _ExcelApp As Excel.Application = New Excel.Application
        Dim ExcelDocs As Excel.Workbooks = _ExcelApp.Workbooks
        Dim Start As Date = Now()
        Main.writeDebug("Accessing excel for rename function")
        If My.Computer.FileSystem.FileExists(IO.Path.Combine(IO.Path.GetTempPath, "Rename.xlsm")) Then
            Kill(IO.Path.Combine(IO.Path.GetTempPath, "Rename.xlsm"))
        End If
        IO.File.WriteAllBytes(IO.Path.Combine(IO.Path.GetTempPath, "Rename.xlsm"), My.Resources.Rename)
        Dim xlPath = IO.Path.Combine(IO.Path.GetTempPath, "Rename.xlsm") 'IO.Path.Combine(exeDir.DirectoryName, "Rename.xlsm")
        Dim ExcelDoc As Excel.Workbook = ExcelDocs.Open(xlPath)
        ExcelDoc.Application.EnableEvents = False
        Try
            ExcelDoc.ActiveSheet.name = "Rename" 'DGVRename.Rows(0).Cells("Part").Value
            Main.writeDebug("Success")
            Dim Lastrow As Integer = ExcelDoc.ActiveSheet.Cells(ExcelDoc.ActiveSheet.Rows.Count, 1).End(Excel.XlDirection.xlUp).Row
            If Lastrow > 4 Then ExcelDoc.ActiveSheet.range("A4:E" & Lastrow).clear
            Dim Y As Integer = 4
            For X = 0 To DGVRename.RowCount - 1
                ProgressBar(X + 1, DGVRename.RowCount, "Exporting: ", DGVRename.Rows(X).Cells("Part").Value, Start)
                If DGVRename.Rows(X).ReadOnly = False And DGVRename.Rows(X).Cells("Reuse").Value = False Then
                    With ExcelDoc.ActiveSheet
                        .range("A" & Y).value = DGVRename.Rows(X).Cells("FileLocation").Value
                        .Range("B" & Y).value = DGVRename.Rows(X).Cells("Drawing").Value
                        .Range("C" & Y).value = DGVRename.Rows(X).Cells("Part").Value
                        .Range("D" & Y).value = DGVRename.Rows(X).Cells("NewName").Value
                        .range("F" & Y).value = DGVRename.Rows(X).Cells("ID").Value
                    End With
                    Dim Img As Image = DGVRename.Rows(X).Cells("Thumbnail").Value
                    If Img Is Nothing Then
                    Else
                        Img.Save(My.Computer.FileSystem.SpecialDirectories.Temp & "\Thumbnail.jpg")
                    End If
                    If IO.File.Exists(My.Computer.FileSystem.SpecialDirectories.Temp & "\Thumbnail.jpg") Then
                        With ExcelDoc.ActiveSheet.Range("E" & Y)
                            .value = ExcelDoc.ActiveSheet.range("C" & Y).value
                            ExcelDoc.ActiveSheet.range("E" & Y).font.color = -16777024
                            .addcomment
                            .comment.visible = False
                            .Comment.Shape.TextFrame.Characters.Text = ""
                            .comment.shape.fill.userpicture(My.Computer.FileSystem.SpecialDirectories.Temp & "\Thumbnail.jpg")
                            .comment.shape.height = 50
                            .comment.shape.width = 50
                            .comment.shape.left = ExcelDoc.ActiveSheet.range("F" & Y).left
                            .comment.shape.top = ExcelDoc.ActiveSheet.range("F" & Y).top
                        End With
                        ExcelDoc.ActiveSheet.rows(Y).rowheight = 50
                        Kill(My.Computer.FileSystem.SpecialDirectories.Temp & "\Thumbnail.jpg")
                    Else
                        ExcelDoc.ActiveSheet.Range("E" & Y).value = "No Thumbnail"
                    End If
                    Main.writeDebug("Exported:  " & DGVRename.Rows(X).Cells("Part").Value)
                    Y += 1
                End If
            Next
            Main.writeDebug(DGVRename.RowCount - 1 & " Items exported to Excel successfully")
            ExcelDoc.ActiveSheet.rows("4:4").Select
            ExcelDoc.Application.ActiveWindow.FreezePanes = True
            Lastrow = ExcelDoc.ActiveSheet.Cells(ExcelDoc.ActiveSheet.Rows.Count, 1).End(Excel.XlDirection.xlUp).Row
            ExcelDoc.ActiveSheet.Rows("4:" & Lastrow).EntireRow.autofit
            ExcelDoc.Application.EnableEvents = True
            If My.Computer.FileSystem.FileExists(My.Computer.FileSystem.SpecialDirectories.Temp & " \Rename.xlsm") Then
                ExcelDoc.Save()
            Else
                ExcelDoc.SaveAs(My.Computer.FileSystem.SpecialDirectories.Temp & "\Rename.xlsm", Excel.XlFileFormat.xlOpenXMLWorkbookMacroEnabled)
            End If
            _ExcelApp.DisplayAlerts = True
            _ExcelApp.Visible = True
            btnImport.Enabled = True
            RenameProgress.Visible = True
        Catch ex As Exception
            Main.writeDebug("Error occurred in performing excel function" & vbNewLine & ex.Message)
            MessageBox.Show(ex.Message, "Exception Details", MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
            If Strings.InStr(ex.Message, "is not trusted") <> 0 Then
                MsgBox("1. Open Excel and access the Options section." & vbNewLine _
                & "2. Select the Trust Center section" & vbNewLine _
                & "3. Click the Trust Center Settings button" & vbNewLine _
                & "4. Select the Macro Settings section" & vbNewLine _
                & "5. Tick the box labelled 'Trust access to the VBA project object model'")
            End If
            If _ExcelApp.Application.Workbooks.Count < 1 Then
                _ExcelApp.ThisWorkbook.Close()
            Else
                _ExcelApp.Application.Quit()
            End If
            Marshal.ReleaseComObject(ExcelDocs)
            Marshal.ReleaseComObject(_ExcelApp)
            Main.KillAllExcels(Time)
        Finally
            Main.writeDebug("Closing Excel")
            ExcelDoc.Application.EnableEvents = True
            'KillAllExcels(Time)
        End Try
        RenameProgress.Visible = False
        Label2.Visible = False
        Remaining.Visible = False
    End Sub

    Private Sub btnImport_Click(sender As Object, e As EventArgs) Handles btnImport.Click
        Dim _ExcelApp As Excel.Application
        _ExcelApp = New Excel.Application
        Dim Start As Date = Now()
        'DGVRename.Rows.Clear()
        Dim ExcelDoc As Excel.Workbook
        Dim Img As New DataGridViewImageCell
        Dim Thumbnail As Image = Nothing
        Try
            _ExcelApp.Visible = False
            _ExcelApp.DisplayAlerts = False
            _ExcelApp.Workbooks.Open(My.Computer.FileSystem.SpecialDirectories.Temp & "\Rename.xlsm")
            ExcelDoc = _ExcelApp.ActiveWorkbook
            Main.writeDebug("Accessing Excel for import")
        Catch ex As Exception
            Main.writeDebug("Error occurred while importing From excel" & vbNewLine & ex.Message)
            MessageBox.Show(ex.Message & vbNewLine & "Couldn't locate file" & vbNewLine & "Ensure you click 'Finished' when complete.",
                            "Exception Details", MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
            If _ExcelApp.Workbooks.Count < 1 Then
                _ExcelApp.Quit()
            Else
                _ExcelApp.ThisWorkbook.Close()
            End If
            Main.KillAllExcels(Time)
            Exit Sub
        End Try
        Dim Lastrow As Integer = ExcelDoc.ActiveSheet.Cells(ExcelDoc.ActiveSheet.Rows.Count, 1).End(Excel.XlDirection.xlUp).Row
        For X = 0 To Lastrow - 4
            RenameProgress.Visible = True
            Label2.Visible = True
            Remaining.Visible = True
            ProgressBar(X, Lastrow - 4, "Importing: ", ExcelDoc.ActiveSheet.range("D" & X + 4).value, Start)
            'DGVRename.Rows.Add()
            'DGVRename.Rows(X).Cells("FileLocation).Value = ExcelDoc.ActiveSheet.Range("A" & X + 4).value
            'If ExcelDoc.ActiveSheet.Range("B" & X + 4).value Is Nothing Then
            'Else
            '    DGVRename.Rows(X).Cells("Drawing").Value = ExcelDoc.ActiveSheet.Range("B" & X + 4).value
            'End If
            For row = 0 To DGVRename.Rows.Count - 1
                If DGVRename.Rows(row).Cells("Part").Value = ExcelDoc.ActiveSheet.Range("C" & X + 4).value And
                    DGVRename.Rows(row).Cells("FileLocation").Value = ExcelDoc.ActiveSheet.range("A" & X + 4).value And
                    DGVRename.Rows(row).Cells("Drawing").Value = ExcelDoc.ActiveSheet.range("B" & X + 4).value Then
                    DGVRename.Rows(row).Cells("NewName").Value = ExcelDoc.ActiveSheet.range("D" & X + 4).value
                    Main.writeDebug("Imported: " & DGVRename.Rows(row).Cells("Part").Value = ExcelDoc.ActiveSheet.Range("C" & X + 4).value)
                End If
            Next
            Main.writeDebug(DGVRename.Rows.Count - 1 & " items successfully imported from Excel")
            'If ExcelDoc.ActiveSheet.range("E" & X + 4).value <> "No Thumbnail" Then
            '    Main.ExtractThumb(DGVRename.Rows(X).Cells("Part").Value, Thumbnail)
            '    DGVRename.Rows(X).Cells("Thumbnail").Value = Thumbnail
            'End If
            Main.writeDebug("Imported :" & DGVRename.Rows(X).Cells("Part").Value)
        Next
        RenameProgress.Visible = False
        Label2.Visible = False
        Remaining.Visible = False
        'If _ExcelApp.Application.Workbooks.Count < 1 Then
        '    _ExcelApp.ThisWorkbook.Close()
        'Else
        _ExcelApp.Workbooks.Close()
        _ExcelApp.Application.Quit()
        'End If
        Marshal.ReleaseComObject(ExcelDoc)
        Marshal.ReleaseComObject(_ExcelApp)
        Main.KillAllExcels(Time)

        If My.Computer.FileSystem.FileExists(My.Computer.FileSystem.SpecialDirectories.Temp & "\Rename.xlsm") Then
            Try
                ' Kill(My.Computer.FileSystem.SpecialDirectories.Temp & "\Rename.xlsm")
            Catch
            End Try
        End If
        _ExcelApp = Nothing
        ExcelDoc = Nothing
        btnImport.Enabled = False
    End Sub
    Private Sub btnRename_Click(sender As Object, e As EventArgs) Handles btnRename.Click
        Dim SaveLoc As String = ""
        Dim TempLoc As String = ""
        Dim Folder As FolderBrowserDialog = New FolderBrowserDialog
        Folder.Description = "Choose the location you wish to save to"
        Folder.RootFolder = System.Environment.SpecialFolder.Desktop
        Folder.SelectedPath = DGVRename.Rows(0).Cells("FileLocation").Value
        Try
            If Folder.ShowDialog() = Windows.Forms.DialogResult.OK Then
                SaveLoc = Folder.SelectedPath & "\"
            Else
                Exit Sub
            End If
            For X = 0 To DGVRename.RowCount - 1
                If DGVRename.Rows(X).Cells("Part").Value = txtParent.Text Then
                    TempLoc = DGVRename.Rows(X).Cells("FileLocation").Value & "Temp\"
                    Main.writeDebug("Temp location set to :" & TempLoc)
                    If Not My.Computer.FileSystem.DirectoryExists(TempLoc) Then
                        MkDir(TempLoc)
                        Main.writeDebug("Temp folder created")
                    End If
                End If
            Next
        Catch ex As Exception
            Main.writeDebug("Error occurred in setting save location" & vbNewLine & ex.Message)
            MessageBox.Show(ex.Message, "Exception Details", MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
        End Try
        CopyFiles(TempLoc, SaveLoc)
        If txtParentSource.Text = SaveLoc Then
            Main.writeDebug("Updating references")
            MsgBox(txtParent.Text & " needs to be closed in order to update part references")
            ' _invapp.Documents.ItemByName(IO.Path.Combine(txtParentSource.Text, txtParent.Text)).Save()
            ' _invapp.Documents.ItemByName(IO.Path.Combine(txtParentSource.Text, txtParent.Text)).Close()
            For Each Doc As Document In _invapp.Documents
                Main.CloseLater(Doc.FullFileName, Doc)
            Next
            For Each Row In DGVRename.Rows
                If My.Computer.FileSystem.FileExists(IO.Path.Combine(DGVRename(DGVRename.Columns("FileLocation").Index, Row.index).Value,
                                                                     DGVRename(DGVRename.Columns("NewName").Index, Row.index).Value)) Then
                    Main.writeDebug("Deleting file: " & IO.Path.Combine(DGVRename(DGVRename.Columns("FileLocation").Index, Row.index).Value,
                                         DGVRename(DGVRename.Columns("Part").Index, Row.index).Value))
                    Kill(IO.Path.Combine(DGVRename(DGVRename.Columns("FileLocation").Index, Row.index).Value,
                                         DGVRename(DGVRename.Columns("Part").Index, Row.index).Value))
                    Try
                        If DGVRename(DGVRename.Columns("Drawing").Index, Row.index).Value <> "" Then
                            Main.writeDebug("Deleting file: " & IO.Path.Combine(DGVRename(DGVRename.Columns("FileLocation").Index, Row.index).Value,
                                                     Strings.Replace(DGVRename(DGVRename.Columns("NewName").Index, Row.index).Value,
                                                 IO.Path.GetExtension(DGVRename(DGVRename.Columns("NewName").Index, Row.index).Value), ".idw")))
                            Kill(IO.Path.Combine(DGVRename(DGVRename.Columns("FileLocation").Index, Row.index).Value,
                                                     Strings.Replace(DGVRename(DGVRename.Columns("NewName").Index, Row.index).Value,
                                                 IO.Path.GetExtension(DGVRename(DGVRename.Columns("NewName").Index, Row.index).Value), ".idw")))
                        End If
                    Catch
                    End Try
                End If
            Next
        End If
        Try
            Main.writeDebug("Moving directory " & TempLoc & " to " & SaveLoc)
            My.Computer.FileSystem.MoveDirectory(TempLoc, SaveLoc, True)
        Catch
            If My.Computer.FileSystem.GetFiles(TempLoc).Count > 1 Then
                MsgBox("Not all files could be moved. In order to proceed please open " & vbNewLine &
                    TempLoc & " and move the renamed files to your appropriate directory.")
            End If
        End Try
        Main.writeDebug("Rename operation completed")
        MsgBox("The operation has completed." & vbNewLine & "The new files are stored in: " & vbNewLine _
       & SaveLoc)
        Dim Document As Document

        For Each Document In _invapp.Documents
            Main.CloseLater(Strings.Right(Document.FullDocumentName, Len(Document.FullDocumentName) - InStrRev(Document.FullDocumentName, "\")), Document)
        Next
        Try
            If My.Computer.FileSystem.DirectoryExists(TempLoc) Then
                My.Computer.FileSystem.DeleteDirectory(TempLoc, FileIO.DeleteDirectoryOption.DeleteAllContents)
            End If
        Catch
            If My.Computer.FileSystem.GetFiles(TempLoc).Count > 0 Then
                MsgBox("Some  files are currently in use and cannot be deleted" & vbNewLine &
                    "Please manually delete " & TempLoc & " after you close your assembly.")
            End If
        End Try

        RenameProgress.Visible = False
        Label2.Visible = False
        Remaining.Visible = False
    End Sub
    Private Sub CopyFiles(TempLoc As String, SaveLoc As String)
        Dim Backup As Boolean = False
        RenameProgress.Visible = True
        Label2.Visible = True
        Remaining.Visible = True
        Dim Start As Date = Now()
        Dim ans As String
        Dim Overwrite As Boolean = Nothing
        Dim Source As String = ""
        Dim Dest As String = ""
        Dim TagLoc As String = ""
        Dim oDoc As Document = Nothing
        Dim dDoc As DrawingDocument = Nothing
        Dim oFileDesc As FileDescriptor
        Dim Y As Integer = 1
        Dim Z As Integer = 0
        Dim DWGError As String = "The following drawing were copied but with errors:" & vbNewLine & vbNewLine
        _invapp.SilentOperation = True
        If txtParentSource.Text <> SaveLoc Then
            ans = MsgBox("Do you wish to make a backup of the original assembly?", vbYesNoCancel)
            Main.writeDebug("Backing up original documents")
        Else
            ans = vbYes
        End If
        If ans = vbYes Then
            If My.Computer.FileSystem.FileExists(IO.Path.Combine(IO.Path.GetDirectoryName(txtParentSource.Text), IO.Path.GetFileNameWithoutExtension(txtParent.Text)) & "-Backup.zip") Then
                Kill(IO.Path.Combine(IO.Path.GetDirectoryName(txtParentSource.Text), IO.Path.GetFileNameWithoutExtension(txtParent.Text)) & "-Backup.zip")
            End If
            Dim ZipPath As String = IO.Path.Combine(IO.Path.GetDirectoryName(txtParentSource.Text), IO.Path.GetFileNameWithoutExtension(txtParent.Text)) & "-Backup.zip"
            AddtoZip(ZipPath)
            Main.writeDebug("Backup saved in " & ZipPath)
            Backup = True
        ElseIf ans = vbNo Then
        Else
            Exit Sub
        End If
        Try
            For X = 0 To DGVRename.RowCount - 1
                If DGVRename.Rows(X).Cells("Reuse").Value = False Then
                    Main.writeDebug("Selected file to be copied: " & Source)
                    Source = DGVRename.Rows.Item(X).Cells("FileLocation").Value & DGVRename.Rows(X).Cells("Part").Value
                    Main.writeDebug("Selected file to be copied: " & Source)
                    If chkStructure.Checked = True Then
                        If InStr(Source, txtParentSource.Text) <> 0 Then
                            TagLoc = Strings.Replace(Strings.Left(Source, (InStrRev(Source, "\"))), txtParentSource.Text, "")
                        Else
                            TagLoc = "Transferred Files\"
                        End If
                    End If
                    Main.writeDebug("Opening file: " & Source)
                    oDoc = _invapp.Documents.Open(Source, False)
                    If DGVRename.Rows(X).Cells("NewName").Value = "" Then
                        DGVRename.Rows(X).Cells("NewName").Value = DGVRename.Rows(X).Cells("Part").Value
                        Main.writeDebug("New part name: " & DGVRename.Rows(X).Cells("Part").Value)
                    ElseIf Strings.LCase(IO.Path.GetExtension(DGVRename.Rows(X).Cells("NewName").Value)) <> "ipt" AndAlso
                        Strings.LCase(IO.Path.GetExtension(DGVRename.Rows(X).Cells("NewName").Value)) <> "iam" Then
                        DGVRename.Rows(X).Cells("NewName").Value = DGVRename.Rows(X).Cells("NewName").Value & Strings.LCase(IO.Path.GetExtension(DGVRename.Rows(X).Cells("Part").Value))
                        Main.writeDebug("New part name " & DGVRename.Rows(X).Cells("NewName").Value)
                    End If
                    ProgressBar(X + 1, DGVRename.RowCount, "Copying: ", DGVRename.Rows(X).Cells("NewName").Value, Start)

                    If My.Computer.FileSystem.FileExists(SaveLoc & TagLoc & DGVRename.Rows(X).Cells("NewName").Value) And Overwrite = Nothing Then
                        ans = MsgBox("Any files of the same name will be overwritten" & vbNewLine & "Continue?", vbYesNo)
                        If ans = vbYes Then
                            Main.writeDebug("Overwrite files authorized")
                            Overwrite = True
                        Else
                            Main.writeDebug("Overwrite files denied - exiting operation")
                            Exit Sub
                        End If
                    End If
                    If Not My.Computer.FileSystem.DirectoryExists(TempLoc & "Transferred Files\") Then
                        Main.writeDebug("Create temporary location for transferred files")
                        My.Computer.FileSystem.CreateDirectory(TempLoc & "Transferred Files\")
                    End If
                    If Overwrite = True And My.Computer.FileSystem.FileExists(IO.Path.Combine(TempLoc, TagLoc, DGVRename.Rows(X).Cells("NewName").Value)) Then
                        Main.writeDebug(IO.Path.Combine(TempLoc, TagLoc, DGVRename.Rows(X).Cells("NewName").Value) & " already exists" & vbNewLine &
                            "Deleting file")
                        My.Computer.FileSystem.DeleteFile(IO.Path.Combine(TempLoc, TagLoc, DGVRename.Rows(X).Cells("NewName").Value))
                    End If
                    Try
                        Debug.Print(IO.Path.Combine(TempLoc, TagLoc, DGVRename.Rows(X).Cells("NewName").Value))
                        If isReadOnly(IO.Path.Combine(TempLoc, TagLoc, DGVRename.Rows(X).Cells("NewName").Value)) = True Then
                            Main.writeDebug("Destination file: " & IO.Path.Combine(TempLoc, TagLoc, DGVRename.Rows(X).Cells("NewName").Value) & " is read-only")
                            ans = MsgBox("The file " & IO.Path.Combine(TempLoc, TagLoc, DGVRename.Rows(X).Cells("NewName").Value) & " is read-only or currently open in another context" & vbNewLine &
                                "The file needs to be closed or checked out of vault before performing the rename operation" & vbNewLine &
                                "Do you wish to continue with the current process?", vbYesNo)
                            If ans = vbNo Then
                                Main.writeDebug("User exited rename operation")
                                Exit Sub
                            Else
                                Main.writeDebug("Skipping file " & IO.Path.Combine(TempLoc, TagLoc, DGVRename.Rows(X).Cells("NewName").Value))
                                Exit Try
                            End If
                        End If
                        oDoc.SaveAs(TempLoc & TagLoc & DGVRename.Rows(X).Cells("NewName").Value, True)
                    Catch ex As Exception
                        If ex.Message <> "Unspecified error (Exception from HRESULT: 0x80004005 (E_FAIL))" Then
                            Main.writeDebug("Error occurred while saving " & DGVRename.Rows(X).Cells("OldName").Value & " to " &
                                            vbNewLine & IO.Path.Combine(TempLoc, TagLoc, DGVRename.Rows(X).Cells("NewName").Value))
                            MessageBox.Show(ex.Message, "Exception Details", MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
                        End If
                    End Try
                    Main.CloseLater(DGVRename.Rows(X).Cells("Part").Value, oDoc)
                    If DGVRename.Rows(X).Cells("Drawing").Value <> "" Then

                        Try
                            Source = IO.Path.Combine(DGVRename.Rows.Item(X).Cells("FileLocation").Value, DGVRename.Rows(X).Cells("Drawing").Value)
                            dDoc = _invapp.Documents.Open(Source, False)
                            If isReadOnly(IO.Path.Combine(TempLoc, TagLoc, IO.Path.GetFileNameWithoutExtension(DGVRename.Rows(X).Cells("NewName").Value)) & ".idw") = True Then
                                Main.writeDebug("Destination file: " & IO.Path.Combine(TempLoc, TagLoc, IO.Path.GetFileNameWithoutExtension(DGVRename.Rows(X).Cells("NewName").Value)) & ".idw" & " is read-only")
                                ans = MsgBox("The file " & IO.Path.Combine(TempLoc, TagLoc, IO.Path.GetFileNameWithoutExtension(DGVRename.Rows(X).Cells("NewName").Value)) & ".idw" & " is read-only or currently open in another context" & vbNewLine &
                                "The file needs to be closed or checked out of vault before performing the rename operation" & vbNewLine &
                                "Do you wish to continue with the current process?", vbYesNo)
                                If ans = vbNo Then
                                    Main.writeDebug("User exited rename operation")
                                    Exit Sub
                                Else
                                    Main.writeDebug("Skipping file " & IO.Path.Combine(TempLoc, TagLoc, IO.Path.GetFileNameWithoutExtension(DGVRename.Rows(X).Cells("NewName").Value)) & ".idw")
                                    Exit Try
                                End If
                            End If
                            dDoc.SaveAs(IO.Path.Combine(TempLoc, TagLoc, IO.Path.GetFileNameWithoutExtension(DGVRename.Rows(X).Cells("NewName").Value)) & ".idw", True)
                            Main.writeDebug("Saved document: " & IO.Path.Combine(TempLoc, TagLoc, IO.Path.GetFileNameWithoutExtension(DGVRename.Rows(X).Cells("NewName").Value)) & ".idw")
                            Main.CloseLater(IO.Path.Combine(TempLoc, TagLoc, IO.Path.GetFileNameWithoutExtension(DGVRename.Rows(X).Cells("NewName").Value)) & ".idw", dDoc)
                            dDoc = _invapp.Documents.Open((IO.Path.Combine(TempLoc, TagLoc, IO.Path.GetFileNameWithoutExtension(DGVRename.Rows(X).Cells("NewName").Value)) & ".idw"), False)
                            For Y = 1 To dDoc.File.ReferencedFileDescriptors.Count
                                For Z = 0 To DGVRename.RowCount - 1
                                    If dDoc.File.ReferencedFileDescriptors.Item(Y).ReferencedFile.FullFileName = DGVRename.Rows(Z).Cells("FileLocation").Value & DGVRename.Rows(Z).Cells("Part").Value Then
                                        Try
                                            oFileDesc = dDoc.File.ReferencedFileDescriptors.Item(Y)
                                            Main.writeDebug("Replacing reference: " & oFileDesc.FullFileName & " with " & IO.Path.Combine(TempLoc, TagLoc, (DGVRename.Rows(X).Cells("NewName").Value)))
                                            Call oFileDesc.ReplaceReference(IO.Path.Combine(TempLoc, TagLoc, (DGVRename.Rows(X).Cells("NewName").Value)))
                                            dDoc.Save()
                                        Catch ex As Exception
                                            Main.writeDebug("An Error occurred While replacing reference To " & TempLoc & DGVRename.Rows(Z).Cells("NewName").Value & "." & vbNewLine &
                                                ex.Message)
                                            MsgBox("An error occurred while replacing reference to " & TempLoc & DGVRename.Rows(Z).Cells("NewName").Value & "." & vbNewLine &
                                                ex.Message & vbNewLine & "This replacement will be skipped.")

                                        End Try
                                        Exit For
                                    End If
                                Next
                            Next
                            Main.CloseLater(Strings.Right(dDoc.FullDocumentName, Len(dDoc.FullDocumentName) - InStrRev(dDoc.FullDocumentName, "\")), dDoc)
                        Catch
                            If Err.Description.Length > 0 And X < DGVRename.RowCount - 1 Then
                                DWGError = DWGError & Strings.Left(DGVRename.Rows(X).Cells("NewName").Value, Len(DGVRename.Rows(X).Cells("NewName").Value) - 3) & "idw" & vbNewLine
                                Err.Clear()
                            End If
                        End Try
                    End If
                Else
                    Main.writeDebug("Reused file: " & Source)
                    Source = DGVRename.Rows.Item(X).Cells("FileLocation").Value & DGVRename.Rows(X).Cells("Part").Value
                    oDoc = _invapp.Documents.Open(Source, False)
                End If
            Next
        Catch ex As Exception
            Main.writeDebug("An unknown error occurred during the rename operation " & ex.Message)
            MessageBox.Show(ex.Message, "Exception Details", MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
        End Try
        For Each Row In DGVRename.Rows
            If DGVRename(DGVRename.Columns("Part").Index, Row.index).Value = txtParent.Text AndAlso DGVRename(DGVRename.Columns("NewName").Index, Row.index).Value <> "" Then
                txtParent.Text = DGVRename(DGVRename.Columns("NewName").Index, Row.index).Value
                txtParentSource.Text = TempLoc
                Exit For
            End If
        Next
        ReplaceReferences(oDoc, TempLoc, IO.Path.Combine(txtParentSource.Text, txtParent.Text))

        Try
            Dim atts As System.IO.FileAttributes = System.IO.File.GetAttributes(oDoc.FullFileName)
            If atts <> 1 And atts <> 33 Then
                oDoc.Save()
            End If
        Catch ex As Exception
            MessageBox.Show(ex.Message, "Exception Details", MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
        End Try
        _invapp.SilentOperation = False
        If DWGError.Length > 54 Then
            MsgBox(DWGError)
        End If
    End Sub
    Private Function isReadOnly(sFile) As Boolean
        If GetAttr(sFile) And vbReadOnly Then
            isReadOnly = True
        Else
            isReadOnly = False
        End If
    End Function
    Private Sub AddtoZip(ByVal ZipPath As String)
        Dim Zip As Package = ZipPackage.Open(ZipPath, IO.FileMode.OpenOrCreate, IO.FileAccess.ReadWrite)
        Dim FileToAdd As String
        For Each row In DGVRename.Rows
            If DGVRename(DGVRename.Columns("Drawing").Index, row.index).Value = "" Then
                FileToAdd = IO.Path.Combine(DGVRename(DGVRename.Columns("FileLocation").Index, row.index).Value,
                DGVRename(DGVRename.Columns("Part").Index, row.index).Value)
                AddToArchive(Zip, FileToAdd)
            Else
                FileToAdd = IO.Path.Combine(DGVRename(DGVRename.Columns("FileLocation").Index, row.index).Value,
                                DGVRename(DGVRename.Columns("Part").Index, row.index).Value)
                AddToArchive(Zip, FileToAdd)
                FileToAdd = IO.Path.Combine(DGVRename(DGVRename.Columns("FileLocation").Index, row.index).Value,
                                DGVRename(DGVRename.Columns("Drawing").Index, row.index).Value)
                AddToArchive(Zip, FileToAdd)
            End If
            Main.writeDebug(FileToAdd & " added to backup file")
        Next

        Zip.Close()
    End Sub
    Private Sub AddToArchive(ByVal zip As Package,
                     ByVal FileToAdd As String)

        'Replace spaces with an underscore (_)
        Dim uriFileName As String = FileToAdd.Replace(" ", "_")

        'A Uri always starts with a forward slash "/"
        Dim zipUri As String = String.Concat("/",
               IO.Path.GetFileName(uriFileName))

        Dim partUri As New Uri(zipUri, UriKind.Relative)
        Dim contentType As String =
               Net.Mime.MediaTypeNames.Application.Zip

        'The PackagePart contains the information:
        ' Where to extract the file when it's extracted (partUri)
        ' The type of content stream (MIME type):  (contentType)
        ' The type of compression:  (CompressionOption.Normal)
        Dim pkgPart As PackagePart = zip.CreatePart(partUri,
               contentType, CompressionOption.Normal)

        'Read all of the bytes from the file to add to the zip file
        Dim bites As Byte() = IO.File.ReadAllBytes(FileToAdd)

        'Compress and write the bytes to the zip file
        pkgPart.GetStream().Write(bites, 0, bites.Length)

    End Sub
    Private Sub ReplaceReferences(ByVal oDoc As Document, ByVal SaveLoc As String, ByVal Source As String)
        Dim Start As Date = Now()
        'For Row As Integer = 0 To DGVRename.Rows.Count - 1
        '    If DGVRename.Rows(Row).Cells("Part").Value = txtParent.Text Then
        '        If DGVRename.Rows(Row).Cells("Reuse").Value = False Then
        '            Source = SaveLoc & DGVRename.Rows(Row).Cells("NewName").Value
        '            Exit For
        '        Else
        '            Source = DGVRename.Rows(Row).Cells("FileLocation").Value & DGVRename.Rows(Row).Cells("Part").Value
        '            Exit For
        '        End If
        '    End If
        'Next
        Main.writeDebug("Replacing part references")
        Dim oAssDoc As AssemblyDocument = _invapp.Documents.Open(Source, False)
        Dim oAssDef As AssemblyComponentDefinition = oAssDoc.ComponentDefinition
        Dim oCompOccs As ComponentOccurrences = oAssDef.Occurrences
        TraverseAssembly(oAssDoc, oCompOccs, SaveLoc, Start, 0)
        Try
            _invapp.SilentOperation = True
            oAssDoc.Save()

        Catch
        Finally
            _invapp.SilentOperation = False
        End Try
        Main.CloseLater(Source, oAssDoc)
        ProgressBar(1, 1, "Cleaning Up", "", Start)
    End Sub
    Private Sub TraverseAssembly(ByVal oAssDoc As Document, ByVal oCompOccs As ComponentOccurrences, ByVal SaveLoc As String, ByVal Start As Date,
                                 ByRef Y As Integer)
        Dim Source As String = ""
        Dim oCompOcc As ComponentOccurrence
        Dim TagLoc As String = ""
        For Each oCompOcc In oCompOccs
            For X = 0 To DGVRename.RowCount - 1
                If oCompOcc.Definition.Document.fullfilename = DGVRename.Rows(X).Cells("FileLocation").Value & DGVRename.Rows(X).Cells("Part").Value And
                    DGVRename.Rows(X).Cells("Reuse").Value = Nothing Then
                    Source = DGVRename.Rows.Item(X).Cells("FileLocation").Value & DGVRename.Rows(X).Cells("NewName").Value
                    If chkStructure.Checked = True Then
                        If InStr(Source, txtParentSource.Text) <> 0 Then
                            TagLoc = Strings.Replace(Strings.Left(Source, (InStrRev(Source, "\"))), txtParentSource.Text, "")
                        Else
                            TagLoc = "Transferred Files\"
                        End If
                    End If
                    ' TagLoc = Replace(Strings.Left(DGVRename.Rows(X).Cells("FileLocation").Value, Strings.InStrRev(DGVRename.Rows(X).Cells("FileLocation").Value, "\")), txtParentSource.Text, "")
                    Try
                        Main.writeDebug("Replacing " & oCompOcc.Name & " with " & SaveLoc & TagLoc & DGVRename.Rows(X).Cells("NewName").Value)
                        Call oCompOcc.Replace(SaveLoc & TagLoc & DGVRename.Rows(X).Cells("NewName").Value, True)
                    Catch
                        Main.writeDebug("Unknown error replacing reference to " & SaveLoc & TagLoc & DGVRename.Rows(X).Cells("NewName").Value)
                        MsgBox("Unknown error replacing reference to " & SaveLoc & TagLoc & DGVRename.Rows(X).Cells("NewName").Value & "." & vbNewLine &
                               "This replacement will be skipped.")

                    End Try
                    Exit For
                End If
            Next
            If oCompOcc.DefinitionDocumentType = DocumentTypeEnum.kAssemblyDocumentObject Then
                ProgressBar(Y, oCompOccs.Count, "Fixing Reference: ", oCompOcc.Definition.Document.displayName, Start)
                TraverseAssembly(oAssDoc, oCompOcc.SubOccurrences, SaveLoc, Start, Y + 1)
            End If
            Y += 1
            ProgressBar(Y, DGVRename.RowCount - 1, "Fixing Reference: ", oCompOcc.Definition.Document.displayName, Start)
        Next
    End Sub
    Private Sub ProgressBar(X As Integer, Total As Integer, Pretext As String, PartName As String, Start As Date)
        Dim Percent As Integer = (X / Total) * 100
        RenameProgress.Value = Percent
        RenameProgress.DisplayText = Pretext & PartName & "    " & Percent & "%"
        Dim Time As TimeSpan = Now - Start
        If RenameProgress.Value = 0 Then Exit Sub
        Dim Remaining As String = ((Time.TotalSeconds / RenameProgress.Value) * (100 - Percent)).ToString
        'Dim Remaining As TimeSpan = TimeSpan.FromSeconds(Time)
        ' Dim Timespan As Integer = (Time.TotalSeconds / RenameProgress.Value) * (100 - Percent)
        If Remaining > 0 Then SecondsToTime(Remaining)
        Me.Remaining.Text = Remaining
    End Sub
    Private Sub SecondsToTime(ByRef Remaining As String)
        Dim hours, minutes, seconds

        ' calculates whole hours (like a div operator)
        hours = Remaining \ 3600

        ' calculates the remaining number of seconds
        Remaining = Remaining Mod 3600

        ' calculates the whole number of minutes in the remaining number of seconds
        minutes = Remaining \ 60
        If minutes < 10 Then minutes = "0" & minutes
        ' calculates the remaining number of seconds after taking the number of minutes
        seconds = Remaining Mod 60


        seconds = Strings.Left(seconds, Strings.InStr(seconds, ".") - 1)
        If seconds < 10 Then seconds = "0" & seconds
        ' returns as a string

        Remaining = hours & ":" & minutes & ":" & seconds
    End Sub

    Private Sub chkDParts_CheckedChanged(sender As Object, e As EventArgs) Handles chkDParts.CheckedChanged
        For X = 0 To DGVRename.RowCount - 1
            If DGVRename.Rows(X).Cells("ID").Value = "DP" Then
                If chkDParts.Checked = False Then
                    DGVRename.Rows(X).ReadOnly = True
                    DGVRename.Rows(X).DefaultCellStyle.BackColor = System.Drawing.Color.LightGray
                    DGVRename.Rows(X).Cells("Reuse").Value = True
                Else
                    DGVRename.Rows(X).ReadOnly = False
                    DGVRename.Rows(X).DefaultCellStyle.BackColor = System.Drawing.Color.White
                    DGVRename.Rows(X).Cells("Reuse").Value = False
                End If
            End If
        Next
    End Sub

    Private Sub chkPParts_CheckedChanged(sender As Object, e As EventArgs) Handles chkPParts.CheckedChanged
        For X = 0 To DGVRename.RowCount - 1
            If DGVRename.Rows(X).Cells("ID").Value = "PP" Then
                If chkPParts.Checked = False Then
                    DGVRename.Rows(X).ReadOnly = True
                    DGVRename.Rows(X).DefaultCellStyle.BackColor = System.Drawing.Color.LightGray
                    DGVRename.Rows(X).Cells("Reuse").Value = True

                Else
                    DGVRename.Rows(X).ReadOnly = False
                    DGVRename.Rows(X).DefaultCellStyle.BackColor = System.Drawing.Color.White
                    DGVRename.Rows(X).Cells("Reuse").Value = False
                End If
            End If
        Next
    End Sub

    Private Sub chkCCParts_CheckedChanged(sender As Object, e As EventArgs) Handles chkCCParts.CheckedChanged
        For X = 0 To DGVRename.RowCount - 1
            If InStr(DGVRename.Rows(X).Cells("FileLocation").Value, "Content Center") <> 0 Then
                If chkCCParts.Checked = False Then
                    DGVRename.Rows(X).ReadOnly = True
                    DGVRename.Rows(X).DefaultCellStyle.BackColor = System.Drawing.Color.LightGray
                    DGVRename.Rows(X).Cells("Reuse").Value = True
                Else
                    DGVRename.Rows(X).ReadOnly = False
                    DGVRename.Rows(X).DefaultCellStyle.BackColor = System.Drawing.Color.White
                    DGVRename.Rows(X).Cells("Reuse").Value = False
                End If
            End If
        Next
    End Sub

    Private Sub Rename_Shown(sender As Object, e As EventArgs) Handles Me.Shown
        chkReuse.Location = New Drawing.Point(1150, chkReuse.Location.Y)
    End Sub

    Private Sub chkReuse_CheckedChanged(sender As Object, e As EventArgs) Handles chkReuse.CheckedChanged
        If chkReuse.Checked = True Then
            For Each row In DGVRename.Rows
                DGVRename(DGVRename.Columns("Reuse").Index, row.index).Value = True
            Next
        Else
            For Each row In DGVRename.Rows
                DGVRename(DGVRename.Columns("Reuse").Index, row.index).Value = False
            Next
        End If
    End Sub
End Class