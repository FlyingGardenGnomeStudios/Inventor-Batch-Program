Imports System.Drawing
Imports System.Windows.Forms
Imports Inventor
Imports System
Imports System.IO
Imports Microsoft.Office.Interop
Imports System.Runtime.InteropServices

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
    End Sub
    Public Sub btnExcel_Click(sender As Object, e As EventArgs) Handles btnExcel.Click
        RenameProgress.Visible = True
        Label2.Visible = True
        Remaining.Visible = True
        Dim _ExcelApp As Excel.Application = New Excel.Application
        Dim ExcelDocs As Excel.Workbooks = _ExcelApp.Workbooks
        Dim Start As Date = Now()
         
        If My.Computer.FileSystem.FileExists(IO.Path.Combine(IO.Path.GetTempPath, "Rename.xlsm")) Then
            Kill(IO.Path.Combine(IO.Path.GetTempPath, "Rename.xlsm"))
        End If
        IO.File.WriteAllBytes(IO.Path.Combine(IO.Path.GetTempPath, "Rename.xlsm"), My.Resources.Rename)
        Dim xlPath = IO.Path.Combine(IO.Path.GetTempPath, "Rename.xlsm") 'IO.Path.Combine(exeDir.DirectoryName, "Rename.xlsm")
        Dim ExcelDoc As Excel.Workbook = ExcelDocs.Open(xlPath)
        ExcelDoc.Application.EnableEvents = False
        Try
            ExcelDoc.ActiveSheet.name = "Rename" 'DGVRename.Rows(0).Cells(2).Value
            Dim Lastrow As Integer = ExcelDoc.ActiveSheet.Cells(ExcelDoc.ActiveSheet.Rows.Count, 1).End(Excel.XlDirection.xlUp).Row
            If Lastrow > 4 Then ExcelDoc.ActiveSheet.range("A4:E" & Lastrow).clear
            Dim Y As Integer = 4
            For X = 0 To DGVRename.RowCount - 1
                ProgressBar(X + 1, DGVRename.RowCount, "Exporting: ", DGVRename.Rows(X).Cells(2).Value, Start)
                If DGVRename.Rows(X).ReadOnly = False Then
                    With ExcelDoc.ActiveSheet
                        .range("A" & Y).value = DGVRename.Rows(X).Cells(0).Value
                        .Range("B" & Y).value = DGVRename.Rows(X).Cells(1).Value
                        .Range("C" & Y).value = DGVRename.Rows(X).Cells(2).Value
                        .Range("D" & Y).value = DGVRename.Rows(X).Cells(3).Value
                        .range("F" & Y).value = DGVRename.Rows(X).Cells(5).Value
                    End With
                    Dim Img As Image = DGVRename.Rows(Y - 4).Cells(4).Value
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
                    Main.writeDebug("Exported: " & DGVRename.Rows(Y - 4).Cells(2).Value)
                    Y += 1
                End If
            Next
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
            KillAllExcels(Time)
        Finally
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
        Catch ex As Exception
            MessageBox.Show(ex.Message & vbNewLine & "Couldn't locate file" & vbNewLine & "Ensure you click 'Finished' when complete.",
                            "Exception Details", MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
            If _ExcelApp.Workbooks.Count < 1 Then
                _ExcelApp.Quit()
            Else
                _ExcelApp.ThisWorkbook.Close()
            End If
            KillAllExcels(Time)
            Exit Sub
        End Try
        Dim Lastrow As Integer = ExcelDoc.ActiveSheet.Cells(ExcelDoc.ActiveSheet.Rows.Count, 1).End(Excel.XlDirection.xlUp).Row
        For X = 0 To Lastrow - 4
            RenameProgress.Visible = True
            Label2.Visible = True
            Remaining.Visible = True
            ProgressBar(X, Lastrow - 4, "Importing: ", ExcelDoc.ActiveSheet.range("D" & X + 4).value, Start)
            'DGVRename.Rows.Add()
            'DGVRename.Rows(X).Cells(0).Value = ExcelDoc.ActiveSheet.Range("A" & X + 4).value
            'If ExcelDoc.ActiveSheet.Range("B" & X + 4).value Is Nothing Then
            'Else
            '    DGVRename.Rows(X).Cells(1).Value = ExcelDoc.ActiveSheet.Range("B" & X + 4).value
            'End If
            For row = 0 To DGVRename.Rows.Count - 1
                If DGVRename.Rows(row).Cells("Part").Value = ExcelDoc.ActiveSheet.Range("C" & X + 4).value And
                    DGVRename.Rows(row).Cells("FileLocation").Value = ExcelDoc.ActiveSheet.range("A" & X + 4).value And
                    DGVRename.Rows(row).Cells("Drawing").Value = ExcelDoc.ActiveSheet.range("B" & X + 4).value Then
                    DGVRename.Rows(row).Cells("NewName").Value = ExcelDoc.ActiveSheet.range("D" & X + 4).value
                End If
            Next
            'If ExcelDoc.ActiveSheet.range("E" & X + 4).value <> "No Thumbnail" Then
            '    Main.ExtractThumb(DGVRename.Rows(X).Cells(2).Value, Thumbnail)
            '    DGVRename.Rows(X).Cells(4).Value = Thumbnail
            'End If
            Main.writeDebug("Imported :" & DGVRename.Rows(X).Cells(2).Value)
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
        KillAllExcels(Time)

        If My.Computer.FileSystem.FileExists(My.Computer.FileSystem.SpecialDirectories.Temp & "\Rename.xlsm") Then
            Try
                Kill(My.Computer.FileSystem.SpecialDirectories.Temp & "\Rename.xlsm")
            Catch
            End Try
        End If
        _ExcelApp = Nothing
        ExcelDoc = Nothing
        btnImport.Enabled = False
    End Sub
    Sub KillAllExcels(Time As System.DateTime)
        Dim proc As System.Diagnostics.Process
        For Each proc In System.Diagnostics.Process.GetProcessesByName("EXCEL")
            If proc.StartTime > Time Then proc.Kill()
            Main.writeDebug("Killed Excel Process")
        Next
    End Sub
    Private Sub btnRename_Click(sender As Object, e As EventArgs) Handles btnRename.Click
        Dim SaveLoc As String = ""
        Dim TempLoc As String = ""
        Dim Folder As FolderBrowserDialog = New FolderBrowserDialog
        Folder.Description = "Choose the location you wish to save to"
        Folder.RootFolder = System.Environment.SpecialFolder.Desktop
        Folder.SelectedPath = DGVRename.Rows(0).Cells(0).Value
        Try
            If Folder.ShowDialog() = Windows.Forms.DialogResult.OK Then
                SaveLoc = Folder.SelectedPath & "\"
            Else
                Exit Sub
            End If
            For X = 0 To DGVRename.RowCount - 1
                If DGVRename.Rows(X).Cells(2).Value = txtParent.Text Then
                    TempLoc = DGVRename.Rows(X).Cells(0).Value & "Temp\"
                    If Not My.Computer.FileSystem.DirectoryExists(TempLoc) Then
                        MkDir(TempLoc)
                    End If
                End If
            Next
        Catch ex As Exception
            MessageBox.Show(ex.Message, "Exception Details", MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
        End Try
        CopyFiles(TempLoc, SaveLoc)
        Try
            My.Computer.FileSystem.MoveDirectory(TempLoc, SaveLoc, True)
        Catch
            If My.Computer.FileSystem.GetFiles(TempLoc).Count > 1 Then
                MsgBox("Not all files could be moved. In order to proceed please open " & vbNewLine &
                    TempLoc & " and move the renamed files to your appropriate directory.")
            End If
        End Try
        MsgBox("The operation has completed." & vbNewLine & "The new files are stored in: " & vbNewLine _
       & SaveLoc)
        Dim Document As Document

        For Each Document In _invapp.Documents
            Main.CloseLater(Strings.Right(Document.FullDocumentName, len(Document.FullDocumentName) - InStrRev(Document.FullDocumentName, "\")), Document)
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
        Try
            For X = 0 To DGVRename.RowCount - 1
                If DGVRename.Rows(X).Cells("Reuse").Value = False Then
                    Source = DGVRename.Rows.Item(X).Cells(0).Value & DGVRename.Rows(X).Cells(2).Value
                    If chkStructure.Checked = True Then
                        If InStr(Source, txtParentSource.Text) <> 0 Then
                            TagLoc = Strings.Replace(Strings.Left(Source, (InStrRev(Source, "\"))), txtParentSource.Text, "")
                        Else
                            TagLoc = "Transferred Files\"
                        End If
                    End If
                    ' TagLoc = Replace(Strings.Left(Source, Strings.InStrRev(Source, "\")), txtParentSource.Text, "")
                    oDoc = _invapp.Documents.Open(Source, False)
                    If DGVRename.Rows(X).Cells(3).Value = "" Then
                        DGVRename.Rows(X).Cells(3).Value = DGVRename.Rows(X).Cells(2).Value
                    ElseIf Strings.LCase(Strings.Right(DGVRename.Rows(X).Cells(3).Value, 4)) <> ".ipt" And
                    Strings.LCase(Strings.Right(DGVRename.Rows(X).Cells(3).Value, 4)) <> ".iam" Then
                        DGVRename.Rows(X).Cells(3).Value = DGVRename.Rows(X).Cells(3).Value & Strings.LCase(Strings.Right(DGVRename.Rows(X).Cells(2).Value, 4))
                    End If
                    ProgressBar(X + 1, DGVRename.RowCount, "Copying: ", DGVRename.Rows(X).Cells(3).Value, Start)

                    If My.Computer.FileSystem.FileExists(SaveLoc & TagLoc & DGVRename.Rows(X).Cells(3).Value) And Overwrite = Nothing Then
                        ans = MsgBox("Any files of the same name will be overwritten" & vbNewLine & "Continue?", vbYesNo)
                        If ans = vbYes Then
                            Overwrite = True
                        Else
                            Exit Sub
                        End If
                    End If
                    If Not My.Computer.FileSystem.DirectoryExists(TempLoc & "Transferred Files\") Then
                        My.Computer.FileSystem.CreateDirectory(TempLoc & "Transferred Files\")
                    End If
                    If Overwrite = True And My.Computer.FileSystem.FileExists(TempLoc & DGVRename.Rows(X).Cells(3).Value) Then
                        My.Computer.FileSystem.DeleteFile(TempLoc & DGVRename.Rows(X).Cells(3).Value)
                    End If
                    Try
                        oDoc.SaveAs(TempLoc & TagLoc & DGVRename.Rows(X).Cells(3).Value, True)
                    Catch ex As Exception
                        If ex.Message <> "Unspecified error (Exception from HRESULT: 0x80004005 (E_FAIL))" Then
                            MessageBox.Show(ex.Message, "Exception Details", MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
                        End If
                    End Try
                    Main.CloseLater(DGVRename.Rows(X).Cells(2).Value, oDoc)
                    If DGVRename.Rows(X).Cells(1).Value <> "" Then
                        Source = DGVRename.Rows.Item(X).Cells(0).Value & DGVRename.Rows(X).Cells(1).Value
                        dDoc = _invapp.Documents.Open(Source, False)
                        dDoc.SaveAs(TempLoc & TagLoc & Strings.Left(DGVRename.Rows(X).Cells(3).Value, Len(DGVRename.Rows(X).Cells(3).Value) - 3) & "idw", True)
                        Main.CloseLater(Strings.Left(DGVRename.Rows(X).Cells(3).Value, Len(DGVRename.Rows(X).Cells(3).Value) - 3) & "idw", dDoc)
                        Try
                            dDoc = _invapp.Documents.Open(TempLoc & Strings.Left(DGVRename.Rows(X).Cells(3).Value, Len(DGVRename.Rows(X).Cells(3).Value) - 3) & "idw", False)
                            For Y = 1 To dDoc.File.ReferencedFileDescriptors.Count
                                For Z = 0 To DGVRename.RowCount - 1
                                    If dDoc.File.ReferencedFileDescriptors.Item(Y).ReferencedFile.FullFileName = DGVRename.Rows(Z).Cells(0).Value & DGVRename.Rows(Z).Cells(2).Value Then
                                        Try
                                            oFileDesc = dDoc.File.ReferencedFileDescriptors.Item(Y)
                                            Call oFileDesc.ReplaceReference(TempLoc & DGVRename.Rows(Z).Cells(3).Value)
                                            dDoc.Save()
                                        Catch
                                            MsgBox("Unknown error replacing reference to " & TempLoc & DGVRename.Rows(Z).Cells(3).Value & "." & vbNewLine &
                                                   "This replacement will be skipped.")

                                        End Try
                                        Exit For
                                    End If
                                Next
                            Next
                            Main.CloseLater(Strings.Right(dDoc.FullDocumentName, Len(dDoc.FullDocumentName) - InStrRev(dDoc.FullDocumentName, "\")), dDoc)
                        Catch
                            If Err.Description.Length > 0 And X < DGVRename.RowCount - 1 Then
                                DWGError = DWGError & Strings.Left(DGVRename.Rows(X).Cells(3).Value, Len(DGVRename.Rows(X).Cells(3).Value) - 3) & "idw" & vbNewLine
                                Err.Clear()
                            End If
                        End Try
                    End If
                End If
            Next
        Catch ex As Exception
            MessageBox.Show(ex.Message, "Exception Details", MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
        End Try
        ReplaceReferences(oDoc, TempLoc, Source)

        Try
            Dim atts As System.IO.FileAttributes = System.IO.File.GetAttributes(oDoc.FullFileName)
            If atts <> 1 and atts <> 33 Then
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
    Private Sub ReplaceReferences(ByVal oDoc As Document, ByVal SaveLoc As String, ByVal Source As String)
        Dim Start As Date = Now()
        For Row As Integer = 0 To DGVRename.Rows.Count - 1
            If DGVRename.Rows(Row).Cells(2).Value = txtParent.Text Then
                Source = SaveLoc & DGVRename.Rows(Row).Cells(3).Value
                Exit For
            End If
        Next
        Dim oAssDoc As AssemblyDocument = _invapp.Documents.Open(Source, False)
        Dim oAssDef As AssemblyComponentDefinition = oAssDoc.ComponentDefinition
        Dim oCompOccs As ComponentOccurrences = oAssDef.Occurrences
        TraverseAssembly(oAssDoc, oCompOccs, SaveLoc, Start, 0)
        oAssDoc.Save()
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
                If oCompOcc.Definition.Document.fullfilename = DGVRename.Rows(X).Cells(0).Value & DGVRename.Rows(X).Cells(2).Value And
                    DGVRename.Rows(X).Cells("Reuse").Value = False Then
                    Source = DGVRename.Rows.Item(X).Cells(0).Value & DGVRename.Rows(X).Cells(2).Value
                    If chkStructure.Checked = True Then
                        If InStr(Source, txtParentSource.Text) <> 0 Then
                            TagLoc = Strings.Replace(Strings.Left(Source, (InStrRev(Source, "\"))), txtParentSource.Text, "")
                        Else
                            TagLoc = "Transferred Files\"
                        End If
                    End If
                    ' TagLoc = Replace(Strings.Left(DGVRename.Rows(X).Cells(0).Value, Strings.InStrRev(DGVRename.Rows(X).Cells(0).Value, "\")), txtParentSource.Text, "")
                    Try
                        Call oCompOcc.Replace(SaveLoc & TagLoc & DGVRename.Rows(X).Cells(3).Value, True)
                    Catch
                        MsgBox("Unknown error replacing reference to " & SaveLoc & TagLoc & DGVRename.Rows(X).Cells(3).Value & "." & vbNewLine &
                               "This replacement will be skipped.")

                    End Try
                    Exit For
                End If
            Next
            If oCompOcc.DefinitionDocumentType = DocumentTypeEnum.kAssemblyDocumentObject Then
                ProgressBar(Y, DGVRename.RowCount - 1, "Fixing Reference: ", oCompOcc.Definition.Document.displayName, Start)
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
            If DGVRename.Rows(X).Cells(5).Value = "DP" Then
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
            If DGVRename.Rows(X).Cells(5).Value = "PP" Then
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
            If InStr(DGVRename.Rows(X).Cells(0).Value, "Content Center") <> 0 Then
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
End Class