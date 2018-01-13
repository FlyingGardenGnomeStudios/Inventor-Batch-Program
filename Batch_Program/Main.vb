Imports System.Type
Imports System.Activator
Imports System.Runtime.InteropServices
Imports Inventor
Imports System.Windows.Forms
Imports Microsoft.Office.Interop
Imports System.Drawing
Imports System.Text.RegularExpressions


Public Class Main
    Dim _invApp As Inventor.Application
    Dim _started As Boolean = False
    Public iProperties As iProperties
    Public Shared RevTable As RevTable
    Public CheckNeeded As CheckNeeded
    Public Settings As New Settings
    Dim RenameTable As New List(Of List(Of String))
    Dim ThumbList As New List(Of Image)
    Dim OpenFiles As New List(Of KeyValuePair(Of String, String))
    Dim SubFiles As New List(Of KeyValuePair(Of String, String))
    Dim AlphaSub As SortedList(Of String, String) = New SortedList(Of String, String)
    Dim OpenDocs As New ArrayList
    Dim VBAFlag As String = "False"
    Dim InvRef(0 To 9) As Inventor.Property
    Dim CustomPropSet As PropertySet
    Dim Time As System.DateTime = Now
    Dim Loading As Boolean = True
    Dim strFileToEncrypt As String
    Dim strFileToDecrypt As String
    Dim strOutputEncrypt As String
    Dim strOutputDecrypt As String
    Dim fsInput As System.IO.FileStream
    Dim fsOutput As System.IO.FileStream
    Dim ReVerifyNow As ReVerifyNow
    Private ta As TurboActivate
    Private isGenuine As Boolean
    Dim myProcess As Process
    Dim InventorExited As Boolean = False
    Private Delegate Sub IsActivatedDelegate()
    ' Set the trial flags you want to use. Here we've selected that the
    ' trial data should be stored system-wide (TA_SYSTEM) and that we should
    ' use un-resetable verified trials (TA_VERIFIED_TRIAL).
    Private trialFlags As TA_Flags = TA_Flags.TA_SYSTEM Or TA_Flags.TA_VERIFIED_TRIAL

    ' Don't use 0 for either of these values.
    ' We recommend 90, 14. But if you want to lower the values
    ' we don't recommend going below 7 days for each value.
    ' Anything lower and you're just punishing legit users.
    Private Const DaysBetweenChecks As UInteger = 90
    Private Const GracePeriodLength As UInteger = 14

#Region "Setup"
    Public Sub writeDebug(ByVal x As String)
        Dim path As String = My.Computer.FileSystem.SpecialDirectories.Temp
        Dim FILE_NAME As String = path & "\Debug.txt"
        If System.IO.File.Exists(FILE_NAME) = False Then
            System.IO.File.Create(FILE_NAME).Dispose()
        End If
        Dim objWriter As New System.IO.StreamWriter(FILE_NAME, True)
        objWriter.WriteLine(DateTime.Now.ToLongTimeString & " " & x & vbNewLine)
        objWriter.Close()

    End Sub
    Public Sub ProcessStarted()
        myProcess = Process.GetProcessesByName("Inventor").First
        myProcess.EnableRaisingEvents = True
        AddHandler myProcess.Exited, AddressOf ProcessExited
    End Sub
    Public Sub ProcessExited(ByVal sender As Object, ByVal e As System.EventArgs)
        InventorExited = True
    End Sub
    Public Sub New()
        ' This call Is required by the designer.
        'My.Settings.Donated = False
        'My.Settings.DonateShowMe = True
        'My.Settings.FirstRun = True
        'My.Settings.DonateCount = 0
        InitializeComponent()
        Dim Warning As New Warning
        If My.Settings.FirstRun = True Then
            Warning.FirstRun()
            Warning.ShowDialog()
            Warning.btnOK.Location = New Drawing.Point(Warning.btnOK.Location.X, Warning.btnOK.Location.Y - 40)
        End If

        If My.Settings.DonateShowMe = True Then
            If My.Settings.DonateCount = 0 Then
                Warning.Donate()
            ElseIf My.Settings.DonateCount < 4 Then
                My.Settings.DonateCount = My.Settings.DonateCount + 1
            ElseIf My.Settings.DonateCount = 4 Then
                My.Settings.DonateCount = 0
            End If
        End If

        'Add any initialization after the InitializeComponent() call.
        Try
            _invApp = Marshal.GetActiveObject("Inventor.Application")
            ProcessStarted()
        Catch ex As Exception
            Try
                Dim Inv As New ProcessStartInfo("Inventor.exe")
                Dim p As New Process
                p.StartInfo = Inv
                p.Start()
                'p.WaitForInputIdle()
                While p.MainWindowTitle = Nothing
                    System.Threading.Thread.Sleep(1000)
                    p.Refresh()
                End While
                _invApp = Marshal.GetActiveObject("Inventor.Application")
                _started = True
                ProcessStarted()
            Catch ex2 As Exception
                MsgBox(ex2.ToString())
                MsgBox("Unable to get or start Inventor")
            End Try
        End Try
        If My.Computer.FileSystem.FileExists(My.Computer.FileSystem.SpecialDirectories.Temp & "\Debug.txt") Then
            Kill(My.Computer.FileSystem.SpecialDirectories.Temp & "\Debug.txt")
        End If
        writeDebug("Inventor Accessed")

        LVSubFiles.Columns(0).Width = LVSubFiles.Width - 10
        Try
            'TODO: goto the version page at LimeLM and paste this GUID here
            ta = New TurboActivate("3d2a7b7e59bfcc74c5df44.47834669")

            ' Check if we're activated, and every 90 days verify it with the activation servers
            ' In this example we won't show an error if the activation was done offline
            ' (see the 3rd parameter of the IsGenuine() function)
            ' https://wyday.com/limelm/help/offline-activation/
            Dim gr As IsGenuineResult = ta.IsGenuine(DaysBetweenChecks, GracePeriodLength, True)

            isGenuine = (gr = IsGenuineResult.Genuine _
                         OrElse gr = IsGenuineResult.GenuineFeaturesChanged _
                         OrElse gr = IsGenuineResult.InternetError)
            ' an internet error means the user is activated but
            ' TurboActivate failed to contact the LimeLM servers



            ' If IsGenuineEx() is telling us we're not activated
            ' but the IsActivated() function is telling us that the activation
            ' data on the computer is valid (i.e. the crypto-signed-fingerprint matches the computer)
            ' then that means that the customer has passed the grace period and they must re-verify
            ' with the servers to continue to use your app.

            'Note: DO NOT allow the customer to just continue to use your app indefinitely with absolutely
            '      no reverification with the servers. If you want to do that then don't use IsGenuine() or
            '      IsGenuineEx() at all -- just use IsActivated().
            If Not isGenuine AndAlso ta.IsActivated() Then

                ' We're treating the customer as is if they aren't activated, so they can't use your app.

                ' However, we show them a dialog where they can reverify with the servers immediately.

                Dim frmReverify As ReVerifyNow = New ReVerifyNow(ta, DaysBetweenChecks, GracePeriodLength)

                If frmReverify.ShowDialog(Me) = DialogResult.OK Then
                    isGenuine = True
                ElseIf Not frmReverify.noLongerActivated Then ' the user clicked cancel and the user is still activated

                    ' Just bail out of your app
                    Close()
                    Return
                End If
            End If

        Catch ex As TurboActivateException
            ' failed to check if activated, meaning the customer screwed
            ' something up so kill the app immediately
            MessageBox.Show("Failed to check if activated:  " + ex.Message)
            Close()
            Return
        End Try

        'Show a trial if we're not genuine
        'See step 9, below.
        ShowTrial(Not isGenuine)
        CreateOpenDocs(OpenDocs)
    End Sub
#End Region
#Region "Form Updating"
    Public Sub UpdateForm()
        'Clear the Open Documents listbox
        OpenFiles.Clear()
        lstOpenfiles.Items.Clear()
        Dim oDoc As Document
        'Iterate through each document open in Inventor and retrieve the display name
        Try
            For Each oDoc In _invApp.Documents.VisibleDocuments
                If oDoc.FullFileName <> Nothing Then
                    'Compare file type to the files chosen to display and only display the selected documents.
                    'Add the document name to key & location to value for faster recall
                    If chkAssy.Checked = True And oDoc.DocumentType = DocumentTypeEnum.kAssemblyDocumentObject Then
                        OpenFiles.Add(New KeyValuePair(Of String, String)(Strings.Right(oDoc.FullFileName, Len(oDoc.FullFileName) - InStrRev(oDoc.FullFileName, "\")), oDoc.FullFileName))
                        'lstOpenfiles.Items.Add(Strings.Right(oDoc.FullFileName, len(oDoc.FullFileName)-InStrRev(oDoc.FullFileName, "\")))
                    ElseIf chkParts.Checked = True And oDoc.DocumentType = DocumentTypeEnum.kPartDocumentObject Then
                        OpenFiles.Add(New KeyValuePair(Of String, String)(Strings.Right(oDoc.FullFileName, Len(oDoc.FullFileName) - InStrRev(oDoc.FullFileName, "\")), oDoc.FullFileName))
                        'lstOpenfiles.Items.Add(Strings.Right(oDoc.FullFileName, len(oDoc.FullFileName)-InStrRev(oDoc.FullFileName, "\")))
                    ElseIf chkDrawings.Checked = True And oDoc.DocumentType = DocumentTypeEnum.kDrawingDocumentObject Then
                        OpenFiles.Add(New KeyValuePair(Of String, String)(Strings.Right(oDoc.FullFileName, Len(oDoc.FullFileName) - InStrRev(oDoc.FullFileName, "\")), oDoc.FullFileName))
                        'lstOpenfiles.Items.Add(Strings.Right(oDoc.FullFileName, len(oDoc.FullFileName)-InStrRev(oDoc.FullFileName, "\")))
                    ElseIf chkPres.Checked = True And oDoc.DocumentType = DocumentTypeEnum.kPresentationDocumentObject Then
                        OpenFiles.Add(New KeyValuePair(Of String, String)(Strings.Right(oDoc.FullFileName, Len(oDoc.FullFileName) - InStrRev(oDoc.FullFileName, "\")), oDoc.FullFileName))
                        'lstOpenfiles.Items.Add(Strings.Right(oDoc.FullFileName, len(oDoc.FullFileName)-InStrRev(oDoc.FullFileName, "\")))
                    End If
                End If
            Next
            writeDebug("Open document list refreshed")
        Catch
            If Strings.Len(Err.ToString) > 0 Then
                writeDebug("***ERROR***" & vbNewLine & Err.Description & vbNewLine & "***ERROR***" & vbNewLine)
                Err.Clear()
            End If

        End Try
        If OpenFiles.Count = 0 Then
            OpenFiles.Clear()
        Else
            'Add an entry for each item in the keyvalue pair list
            For Each Pair As KeyValuePair(Of String, String) In OpenFiles
                lstOpenfiles.Items.Add(Pair.Key)
            Next
        End If

        If LVSubFiles.Items.Count = 0 Then
            'change display style to simulate an unusable entity
            LVSubFiles.Items.Add("No Drawings Found")
            'LVSubFiles.ThreeDCheckBoxes = False
            LVSubFiles.ForeColor = Drawing.Color.Gray
            For x = 0 To LVSubFiles.Items.Count - 1
                LVSubFiles.Items(x).Checked = False
            Next
        End If
        If lstOpenfiles.Items.Count > 0 Then
            If InStr(lstOpenfiles.Items(0), "ipt") > 0 Or
                InStr(lstOpenfiles.Items(0), "idw") > 0 Or
                InStr(lstOpenfiles.Items(0), "iam") > 0 Or
                InStr(lstOpenfiles.Items(0), "ipn") > 0 Then
            Else
                MsgBox("Some parts are not listed as they either have an unknown extension" & vbNewLine &
                       "or the file extensions are invisible." & vbNewLine &
                   "(You can enable file name extension visibility through the file explorer)")
                End
            End If
        End If
        LVSubFiles.AutoResizeColumns(ColumnHeaderAutoResizeStyle.ColumnContent)
    End Sub
    Private Sub CheckboxReorder()
        If chkAssy.Checked = True And chkParts.Checked = False Then
            chkDerived.Visible = True
            chkAssy.Top = chkParts.Location.Y + 15
            chkDerived.Top = chkAssy.Location.Y + 15
            chkPres.Top = chkDerived.Location.Y + 15
            GroupBox1.Height = 95
        ElseIf chkAssy.Checked = False And chkParts.Checked = True Then
            chkDerived.Visible = True
            chkDerived.Top = chkParts.Location.Y + 15
            chkAssy.Top = chkDerived.Location.Y + 15
            chkPres.Top = chkAssy.Location.Y + 15
            GroupBox1.Height = 95
        ElseIf chkAssy.Checked = True And chkParts.Checked = True Then
            chkDerived.Visible = True
            chkAssy.Top = chkParts.Location.Y + 15
            chkDerived.Top = chkAssy.Location.Y + 15
            chkPres.Top = chkDerived.Location.Y + 15
            GroupBox1.Height = 95
        Else
            chkAssy.Checked = False And chkParts.Checked = False
            chkDerived.Visible = False
            chkDerived.Checked = False
            chkDerived.Top = chkParts.Location.Y + 15
            chkAssy.Top = chkParts.Location.Y + 15
            chkPres.Top = chkAssy.Location.Y + 15
            GroupBox1.Height = 80
        End If
    End Sub
    Private Sub tmr_Tick(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles tmr.Tick
        MsVistaProgressBar.Visible = True
        LVSubFiles.Enabled = False
        'MsVistaProgressBar.ProgressBarStyle = MSVistaProgressBar.BarStyle.Marquee
        tmr.Enabled = False
        'set the colour and style of the subfiles window
        SubFiles.Clear()
        AlphaSub.Clear()
        LVSubFiles.Items.Clear()
        RenameTable.Clear()
        ThumbList.Clear()
        'LVSubFiles.ThreeDCheckBoxes = True
        LVSubFiles.ForeColor = Drawing.Color.Black
        Dim watch As Stopwatch = Stopwatch.StartNew
        Dim oDoc As Document
        Dim Path As Documents
        Dim PartSource As String
        Dim Level As Integer = 0
        Dim Total As Integer = 1
        Dim Counter As Integer = 1
        Dim Title As String = "Found: "
        ProgressBar(Total, Counter, Title, "")
        CreateOpenDocs(OpenDocs)
        Dim Elog As String = ""
        'iterate through the open files and display the available drawings in the Subfiles window

        Dim X As Integer
        For X = 0 To lstOpenfiles.CheckedItems.Count - 1
            'for drawing documents add name and location from openfiles list
            If Strings.InStr(lstOpenfiles.CheckedItems(X), "idw") <> 0 Then
                SubFiles.Add(New KeyValuePair(Of String, String)(lstOpenfiles.CheckedItems(X), OpenFiles.Item(lstOpenfiles.Items.IndexOf(lstOpenfiles.CheckedItems(X))).Value))
                AlphaSub.Add(lstOpenfiles.CheckedItems(X), OpenFiles.Item(lstOpenfiles.Items.IndexOf(lstOpenfiles.CheckedItems(X))).Value)
                Elog = SubFiles.Item(SubFiles.Count - 1).Value & ": Added to Open File list" & vbNewLine
            Else
                'iterate through open documents looking for the item corresponding to the checked item, if it exists, display it in the subfiles window.
                For J = 1 To _invApp.Documents.Count
                    Path = _invApp.Documents
                    oDoc = Path.Item(J)
                    'get the source path of the targeted file
                    PartSource = oDoc.FullFileName

                    'If the targeted document is the checked document, identify its type
                    If lstOpenfiles.CheckedItems.Item(X) = Strings.Right(PartSource, Strings.Len(PartSource) - InStrRev(PartSource, "\")) Then
                        'Drawing documents can be directly sent to the subfiles list
                        If oDoc.DocumentType = DocumentTypeEnum.kDrawingDocumentObject Then
                            TestForDrawing(PartSource, 0, Total, Counter, OpenDocs, Elog, False)
                            'Part/Presentation documents require a search for the associated drawings
                        ElseIf oDoc.DocumentType = DocumentTypeEnum.kPartDocumentObject Or oDoc.DocumentType = DocumentTypeEnum.kPresentationDocumentObject Then

                            'Test to see if assumed drawing name exists
                            TestForDrawing(PartSource, 0, Total, Counter, OpenDocs, Elog, False)
                            CheckForDev(PartSource, Total, Counter, OpenDocs, Elog, Level + 1)
                            'Assembly documents require a search for subfiles
                        ElseIf oDoc.DocumentType = DocumentTypeEnum.kAssemblyDocumentObject Then
                            'Total = Total + oDoc.ReferencedDocuments.Count 
                            TraverseAssemblyLoad(PartSource, 0, Total, Counter, OpenDocs, Elog, Dev:=False)
                        End If
                    End If
                Next
            End If
        Next
        'update the listbox when complete
        writeDebug(Elog)
        RunCompleted()
    End Sub

#End Region
#Region "Left side buttons"
    Private Sub chkDrawings_CheckedChanged(sender As System.Object, e As System.EventArgs) Handles chkDrawings.CheckedChanged
        UpdateForm()
        writeDebug("Drawings Checked = " & chkDrawings.CheckState)
    End Sub
    Private Sub chkParts_CheckedChanged(sender As System.Object, e As System.EventArgs) Handles chkParts.CheckedChanged
        UpdateForm()
        writeDebug("Parts Checked = " & chkParts.CheckState)
        CheckboxReorder()
    End Sub
    Private Sub chkAssy_CheckedChanged(sender As System.Object, e As System.EventArgs) Handles chkAssy.CheckedChanged
        UpdateForm()
        writeDebug("Assembly Checked = " & chkAssy.CheckState)
        CheckboxReorder()
    End Sub
    Private Sub chkPres_CheckedChanged(sender As System.Object, e As System.EventArgs) Handles chkPres.CheckedChanged
        UpdateForm()
        writeDebug("Presentation Checked = " & chkPres.CheckState)
    End Sub
    Private Sub WriteChildren(Occurrences As ComponentOccurrences, Layer As Integer,
                          ByRef R As Integer, ByRef Col As Integer, ByRef Maxlayer As Integer,
                          Parent As List(Of List(Of String)), ParentName As String, ByRef Total As Integer, ByRef Counter As Integer)
        Dim Occ As ComponentOccurrence
        Dim Title As String = "Getting References: "
        'Dim oDoc As Document
        'Dim ColVal As Integer
        'Dim Dup, Fresh As Boolean
        Dim Match As Boolean
        Dim PartName As String
        'Parent(Layer, Col) = Occ.Name
        ParentName = ParentName
        'go through each sub-file
        For Each Occ In Occurrences
            If Occ.Name <> "" Then
                'Parse out part name from the occurance name
                If InStr(Occ.Name, ":") <> 0 Then
                    PartName = Strings.Left(Occ.Name, InStrRev(Occ.Name, ":") - 1)
                Else
                    PartName = Occ.Name
                End If
                'Match the case of the name to IMM Standard
                Match = (PartName Like "?###[-.]#####") Or (PartName Like "?##[-.]#####") Or (PartName Like "?###[-.]#####[-.]##") Or (PartName Like "?##[-.]#####[-.]##")
                'If Match = True Then
                Col += 1
                ProgressBar(Total, Counter, Title, Col)
                'For Parents with more children, add another list to the current list
                If Layer > Maxlayer Then
                    Maxlayer = Layer
                    Parent.Add(New List(Of String))
                End If
                'search for repetitions in the current list
                If Parent(Layer).Contains(ParentName & "," & PartName) Then
                Else
                    'add the new part to the current list



                    Parent(Layer).Add(ParentName & "," & PartName)
                End If
                'End If
                If Occ.DefinitionDocumentType = DocumentTypeEnum.kAssemblyDocumentObject Then
                    'Go through for each assembly found
                    Counter += 1
                    WriteChildren(Occ.SubOccurrences, Layer + 1, R, Col, Maxlayer, Parent, PartName, Total, Counter)
                End If
            End If
        Next
    End Sub
    Private Sub btnSpreadsheet_Click(sender As System.Object, e As System.EventArgs) Handles btnSpreadsheet.Click
        VBAFlag = "NA"
        Dim C, Total, Ans As Integer
        Dim Counter As Integer = 1
        Total = 0
        Dim AssyName, SaveName, PreSave As String
        SaveName = "" : PreSave = ""
        Dim AsmDoc As AssemblyDocument
        Dim oDoc As Document = Nothing
        Dim AsmDef As AssemblyComponentDefinition
        Dim PropSets As PropertySets
        Dim ExcelDoc As Excel.Workbook = Nothing
        Dim OpenDocs As New ArrayList
        CreateOpenDocs(OpenDocs)
        Dim Elog As String = ""
        Dim StartTime As Date = Now
        Dim _ExcelApp As New Excel.Application
        For X = 0 To lstOpenfiles.Items.Count - 1
            If lstOpenfiles.GetItemCheckState(X) = CheckState.Checked Then
                If Strings.Right(lstOpenfiles.Items.Item(X).ToString, 3) = "iam" Then
                    C += 1
                End If
            End If
        Next
        If C = 0 Then
            MsgBox("Please select an assembly")
            Exit Sub
        End If
        If C > 1 Then
            MsgBox("Please only select one assembly")
            Exit Sub
        End If
        For X = 0 To lstOpenfiles.Items.Count - 1
            If lstOpenfiles.GetItemCheckState(X) = CheckState.Checked Then
                AssyName = Trim(lstOpenfiles.Items.Item(X).ToString)

                For Each oDoc In _invApp.Documents
                    If InStr(oDoc.FullFileName, AssyName) <> 0 Then
                        oDoc.Activate()
                        If oDoc.DocumentType = DocumentTypeEnum.kAssemblyDocumentObject Then
                            AsmDoc = _invApp.ActiveDocument
                        Else
                            MsgBox("The assembly must be active.")
                            Exit Sub
                        End If
                        Try
                            MsVistaProgressBar.Visible = True
                            AsmDef = AsmDoc.ComponentDefinition
                            CountParts(Total)
                            If My.Computer.FileSystem.FileExists(IO.Path.Combine(IO.Path.GetTempPath, "List-Blank.xlsm")) Then
                                Kill(IO.Path.Combine(IO.Path.GetTempPath, "List-Blank.xlsm"))
                            End If
                            If My.Computer.FileSystem.FileExists(IO.Path.Combine(IO.Path.GetTempPath, "Quote-Blank.xlsm")) Then
                                Kill(IO.Path.Combine(IO.Path.GetTempPath, "Quote-Blank.xlsm"))
                            End If
                            IO.File.WriteAllBytes(IO.Path.Combine(IO.Path.GetTempPath, "List-Blank.xlsm"), My.Resources.List_Blank)
                            IO.File.WriteAllBytes(IO.Path.Combine(IO.Path.GetTempPath, "Quote-Blank.xlsm"), My.Resources.Quote_Blank)
                            Dim xlPath = IO.Path.Combine(IO.Path.GetTempPath, "List-Blank.xlsm")
                            _ExcelApp.Workbooks.Open(xlPath)
                            _ExcelApp.Visible = False
                            ExcelDoc = _ExcelApp.ActiveWorkbook
                            'ShowParameters(oDoc)
                            GetProperties(oDoc, AsmDef.Occurrences, 0, 0, ExcelDoc, Total)
                            MsVistaProgressBar.Visible = False
                            ExcelDoc.Worksheets("Saw Cut Lengths").visible = False
                            'UpdateCustomiProperty(oDoc, "", "")
                            PropSets = AsmDoc.PropertySets
                            'SaveName = Strings.Left(AsmDoc.FullFileName, InStrRev(AsmDoc.FullFileName, "\") - 1)
                            SaveName = Strings.Left(AsmDoc.FullFileName, InStrRev(AsmDoc.FullFileName, "\"))
                            PreSave = Strings.Left(Strings.Right(AsmDoc.FullFileName, Len(AsmDoc.FullFileName) - InStrRev(AsmDoc.FullFileName, "\")),
                                                   Strings.Len(Strings.Right(AsmDoc.FullFileName, Len(AsmDoc.FullFileName) - InStrRev(AsmDoc.FullFileName, "\"))) - 4) & "-List.xlsm"
                            If My.Computer.FileSystem.DirectoryExists(SaveName & "Vendor Quotes\") Then
                            Else
                                My.Computer.FileSystem.CreateDirectory(SaveName & "Vendor Quotes\")
                            End If
                            _ExcelApp.ActiveWorkbook.SaveCopyAs(SaveName & "Vendor Quotes\" & PreSave)
                            _ExcelApp.DisplayAlerts = False
                            ExcelDoc.Close()
                            _ExcelApp.Quit()
                            KillAllExcels(StartTime)

                        Catch ex As Exception
                            MessageBox.Show(ex.Message, "Exception Details", MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
                        Finally
                            Ans = MsgBox("A spreadsheet of all part properties has been created" & vbNewLine &
                                         "Do you wish to open the file?", vbYesNo, "Finished")
                            If Ans = vbYes Then
                                System.Diagnostics.Process.Start(SaveName & "Vendor Quotes\" & PreSave)
                                '_ExcelApp = Marshal.GetActiveObject("Excel.application")
                                '_ExcelApp.Workbooks.Open(SaveName & "Vendor Quotes\" & PreSave)
                                '_ExcelApp.Visible = True
                            Else
                                '_ExcelApp.Visible = True

                                '_ExcelApp.Quit()
                                '_ExcelApp.DisplayAlerts = True
                                '_ExcelApp = Nothing
                                'KillAllExcels(StartTime)
                            End If
                        End Try
                        Exit Sub
                    End If
                Next
            End If
        Next
    End Sub
    Private Sub btnRename_Click(sender As System.Object, e As System.EventArgs) Handles btnRename.Click

        If lstOpenfiles.CheckedItems.Count = 0 Then
            MsgBox("Please select an item to rename")
            Exit Sub
        ElseIf lstOpenfiles.CheckedItems.Count > 1 Then
            MsgBox("Only select only one item to be renamed")
            Exit Sub
        End If
        If InStr(lstOpenfiles.CheckedItems.Item(0).ToString, ".ipt") <> 0 Then
            MsgBox("Rename currently only supports assemblies")
            Exit Sub
        End If
        Dim Infoform As New Warning
        If My.Settings.RenameShowMe = True Then
            Infoform.ShowDialog()
        End If
        CreateOpenDocs(OpenDocs)
        Dim ErrList As String = ""
        Dim Elog As String = ""
        If VBAFlag = "False" Then CreateVBA()
        Dim Rename As New Rename

        Dim X As Integer = 0
        Rename.PopMain(Me)
        Try
            For Each Entry As List(Of String) In RenameTable
                If X = 0 Then
                    Rename.txtParent.Text = Entry.Item(2).ToString
                    Rename.txtParentSource.Text = Entry.Item(0).ToString
                End If
                Rename.DGVRename.Rows.Add(Entry.Item(0).ToString, Entry.Item(1).ToString, Entry.Item(2).ToString, "", ThumbList(X), Entry.Item(3).ToString)
                Elog = Elog & Entry.Item(2).ToString & ": Added to Rename Table" & vbNewLine
                X += 1
            Next
        Catch
            ErrList = Err.Description & vbNewLine
            Err.Clear()
        Finally
            If ErrList.Length > 0 Then
                MsgBox("The following occurred during the creation of the rename table" & vbNewLine & ErrList)
            End If
            writeDebug(Elog)
        End Try
        Rename.chkCCParts.Checked = False

        Rename.Show()
        'RenameTable.Clear()
    End Sub

#End Region
#Region "Right side buttons"
    Public Function PopiProperties(CalledFunction As iProperties)
        iProperties = CalledFunction
        Return Nothing
    End Function
    Private Sub chkPartSelect_CheckedChanged(sender As Object, e As EventArgs) Handles chkPartSelect.CheckedChanged
        If chkPartSelect.CheckState = CheckState.Checked Then
            For X = 0 To lstOpenfiles.Items.Count - 1
                lstOpenfiles.SetItemChecked(X, True)
            Next
        Else
            For X = 0 To lstOpenfiles.Items.Count - 1
                lstOpenfiles.SetItemChecked(X, False)
            Next
        End If
    End Sub
    Private Sub chkDWGSelect_CheckedChanged(sender As Object, e As EventArgs) Handles chkDWGSelect.CheckedChanged
        If chkDWGSelect.CheckState = CheckState.Checked Then
            For X = 0 To LVSubFiles.Items.Count - 1
                LVSubFiles.Items(X).Checked = True
            Next

        Else
            For X = 0 To LVSubFiles.Items.Count - 1
                LVSubFiles.Items(X).Checked = False
            Next
        End If
    End Sub
    Private Sub PictureBox2_Click(sender As Object, e As EventArgs) Handles PictureBox2.Click
        For X = 0 To LVSubFiles.Items.Count - 1
            If LVSubFiles.Items(X).Checked = True Then
                LVSubFiles.Items(X).Checked = False
            Else
                LVSubFiles.Items(X).Checked = True
            End If
        Next
    End Sub
    Public Sub OpenSelected(Path As Documents, ByRef odoc As Document, ByRef Archive As String _
                             , ByRef DrawingName As String, ByRef DrawSource As String)
        'Go through drawings to see which ones are selected
        For X = 0 To LVSubFiles.Items.Count - 1
            'Look through all sub files in open documents to get the part sourcefile
            If LVSubFiles.Items(X).Checked = True Then
                MatchDrawing(DrawSource, DrawingName, X)
                'DrawSource = Strings.Left(SubFiles.Item(X).Value, Len(SubFiles.Item(X).Value) - 3) & "idw"
                ProgressBar(LVSubFiles.CheckedItems.Count, X, "Opening:   ", DrawingName)
                Try
                    odoc = _invApp.Documents.Open(DrawSource, True)
                    writeDebug("Opened Drawing: " & DrawingName)
                Catch
                    MsgBox("Could not open drawing. Ensure the drawing exists")
                    writeDebug("Could not open drawing: " & DrawingName)
                End Try
            End If
        Next
    End Sub
    Public Sub ExportCheck(Path As Documents, ByRef odoc As Document, ByRef Archive As String _
                             , ByRef DrawingName As String, ByRef DrawSource As String, ExportType As String)
        Dim OpenDocs As New ArrayList
        CreateOpenDocs(OpenDocs)
        Dim X, Ans, Total, Counter, OWCounter As Integer
        Counter = 1
        Dim Overwrite, Title, PDFSource, DXFSource, DWGSource, RevNo As String
        Overwrite = ""
        PDFSource = ""
        DXFSource = ""
        DWGSource = ""
        RevNo = ""
        'Get total amount of documents for the process bar
        Total = LVSubFiles.CheckedItems.Count
        'Go through drawings to see which ones are selected
        Dim Destin As String = ""

        For X = 0 To LVSubFiles.Items.Count - 1
            'Look through all sub files in open documents to get the part sourcefile
            If LVSubFiles.Items(X).Checked = True Then
                MatchPart(DrawSource, DrawingName, X)
                If chkSkipAssy.Checked = True And DrawSource.Contains(".iam") Then
                Else
                    'iterate through opend documents to find the selected file
                    MatchDrawing(DrawSource, DrawingName, X)
                    If My.Settings(ExportType & "SaveNewLoc") = False And My.Settings(ExportType & "SaveTag") = False Then
                        Dim Folder As FolderBrowserDialog = New FolderBrowserDialog
                        Folder.Description = "Choose the location you wish to save the " & UCase(ExportType)
                        Folder.RootFolder = System.Environment.SpecialFolder.Desktop
                        Folder.SelectedPath = Strings.Left(DrawSource, InStrRev(DrawSource, "\"))
                        If Destin = "" Then
                            Try
                                If Folder.ShowDialog() = Windows.Forms.DialogResult.OK Then
                                    Destin = Folder.SelectedPath
                                Else
                                    Exit Sub
                                End If
                            Catch ex As Exception
                                MessageBox.Show(ex.Message, "Exception Details", MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
                            End Try
                        End If
                    ElseIf My.Settings(ExportType & "SaveNewLoc") = True Then
                        Destin = My.Settings(ExportType & "SaveLoc")
                    ElseIf My.Settings(ExportType & "Tag") <> "" Then
                        Destin = Strings.Left(DrawSource, InStrRev(DrawSource, "\") - 1) & My.Settings(ExportType & "Tag")
                    End If
                    'open drawing
                    odoc = _invApp.Documents.Open(DrawSource, False)
                    'Get revision number
                    If My.Settings(ExportType & "Rev") = True Then
                        RevNo = "-R" & odoc.PropertySets.Item("{F29F85E0-4FF9-1068-AB91-08002B27B3D9}").ItemByPropId("9").Value
                    Else
                        RevNo = ""
                    End If
                    'Check for files that will be overwritten
                    'Set the filename to be saved along with the revision

                    'Set the PDF/DXF name and search to see if it exists
                    DXFSource = Destin & "\" & DrawingName
                    PDFSource = Destin & "\" & DrawingName
                    DXFSource = DXFSource.Insert(DXFSource.LastIndexOf("."), RevNo)
                    PDFSource = PDFSource.Insert(PDFSource.LastIndexOf("."), RevNo)
                    DXFSource = Replace(DXFSource, "idw", "dxf")
                    PDFSource = Replace(PDFSource, "idw", "pdf")
                    DWGSource = Replace(DXFSource, "dxf", "dwg")
                    'If the file exists, save it for display later
                    If ExportType = "PDF" Then
                        If chkPDF.CheckState = CheckState.Checked Then
                            If My.Computer.FileSystem.FileExists(PDFSource) Then
                                Overwrite = Overwrite & Strings.Left(DrawingName, Strings.Len(DrawingName) - 4) & RevNo & ".pdf" & vbNewLine
                                OWCounter += 1
                            End If
                        End If
                    ElseIf ExportType = "dwg" Then
                        If chkDWG.CheckState = CheckState.Checked Then
                            If My.Computer.FileSystem.FileExists(DWGSource) Or
                                My.Computer.FileSystem.FileExists(DWGSource.Insert(DWGSource.LastIndexOf("."), "_Sheet_1")) Then
                                Overwrite = Overwrite & Strings.Left(DrawingName, Strings.Len(DrawingName) - 4) & RevNo & ".dwg" & vbNewLine
                                OWCounter += 1
                            End If
                        End If
                    ElseIf ExportType = "dxf" Then
                        If My.Computer.FileSystem.FileExists(DXFSource) Or
                        My.Computer.FileSystem.FileExists(DXFSource.Insert(DXFSource.LastIndexOf("."), "_Sheet_1")) Then
                            Overwrite = Overwrite & Strings.Left(DrawingName, Strings.Len(DrawingName) - 4) & RevNo & ".dxf" & vbNewLine
                            OWCounter += 1
                        End If
                    End If

                    CloseLater(DrawingName, odoc)
                    Counter += 1
                    Title = "Checking"
                    ProgressBar(Total, Counter, Title, DrawingName)
                    'Display any files that will be overwritten
                End If
            End If
        Next
        Counter = 1
        If Overwrite = "" And ExportType = "PDF" Then
            PDFCreator(Path, odoc, PDFSource, Archive, DrawingName, DrawSource, Destin, OpenDocs, Total, Counter)
        ElseIf Overwrite = "" And ExportType = "dxf" Then
            DXFDWGCreator(Path, odoc, Destin, Archive, DrawingName, DrawSource, OpenDocs, Total, Counter, DXFSource, "dxf")
        ElseIf Overwrite = "" And ExportType = "dwg" Then
            DXFDWGCreator(Path, odoc, Destin, Archive, DrawingName, DrawSource, OpenDocs, Total, Counter, DWGSource, "dwg")
        ElseIf OWCounter > 10 Then
            Ans = MsgBox("More than 10 files" & vbNewLine & "will be overwritten." _
                       & vbNewLine & " Do you wish to continue?", vbOKCancel, "Overwrite")
            If Ans = vbOK And ExportType = "PDF" Then
                PDFCreator(Path, odoc, PDFSource, Archive, DrawingName, DrawSource, Destin, OpenDocs, Total, Counter)
            ElseIf Ans = vbOK And ExportType = "dxf" Then
                DXFDWGCreator(Path, odoc, Destin, Archive, DrawingName, DrawSource, OpenDocs, Total, Counter, DXFSource, "dxf")
            Else
                chkCheck.CheckState = CheckState.Unchecked
                MsVistaProgressBar.Visible = False
                Exit Sub
            End If
        Else
            Ans = MsgBox("The following files will be overwritten:" _
                         & vbNewLine & vbNewLine & Overwrite, vbOKCancel, "Overwrite")
            writeDebug("Notify user of files to be overwritten: " & vbNewLine & Overwrite)
            If Ans = vbOK And ExportType = "PDF" Then
                PDFCreator(Path, odoc, PDFSource, Archive, DrawingName, DrawSource, Destin, OpenDocs, Total, Counter)
            ElseIf Ans = vbOK And ExportType = "dxf" Then
                DXFDWGCreator(Path, odoc, Destin, Archive, DrawingName, DrawSource, OpenDocs, Total, Counter, DXFSource, "dxf")
            ElseIf Ans = vbOK And ExportType = "dwg" Then
                DXFDWGCreator(Path, odoc, Destin, Archive, DrawingName, DrawSource, OpenDocs, Total, Counter, DWGSource, "dwg")
            Else
                chkCheck.CheckState = CheckState.Unchecked
                Exit Sub
            End If
        End If
    End Sub
    Private Sub PDFCreator(Path As Documents, ByRef oDoc As Document, ByRef PDFSource As String, ByRef Archive As String _
                             , ByRef DrawingName As String, ByRef DrawSource As String, ByRef Destin As String, OpenDocs As ArrayList _
                             , Total As Integer, Counter As Integer)
        ' Get the PDF translator Add-In.
        Dim ErrorReport As String = ""
        Dim Title, RevNo As String
        Dim oPDFTrans As Inventor.ApplicationAddIn
        oPDFTrans = _invApp.ApplicationAddIns.ItemById("{0AC6FD96-2F4D-42CE-8BE0-8AEA580399E4}")
        If chkPDF.Checked = True And oPDFTrans Is Nothing Then
            MsgBox("Could Not access PDF translator.")
            Exit Sub
        End If
        ' Create some objects that are used to pass information to the translator Add-In.   
        Dim oContext As TranslationContext
        oContext = _invApp.TransientObjects.CreateTranslationContext
        Dim oOptions As NameValueMap
        oOptions = _invApp.TransientObjects.CreateNameValueMap
        'Go through drawings to see which ones are selected
        _invApp.SilentOperation = True
        For X = 0 To LVSubFiles.Items.Count - 1
            'Look through all sub files in open documents to get the part sourcefile
            If LVSubFiles.Items(X).Checked = True Then
                'iterate through open documents to find the selected file
                MatchDrawing(DrawSource, DrawingName, X)

                'open drawing
                oDoc = _invApp.Documents.Open(DrawSource, True)

                If oPDFTrans.HasSaveCopyAsOptions(_invApp.ActiveDocument, oContext, oOptions) Then
                    Try
                        For P = 1 To oDoc.sheets.count
                            oDoc.sheets.item(P).activate()
                        Next
                        Dim Mass As Double = oDoc.ReferencedFiles.Item(1).componentdefinition.massproperties.mass
                        'Next
                        Call oDoc.Update()
                        oOptions.Value("Sheet_Range") = Inventor.PrintRangeEnum.kPrintAllSheets
                        oOptions.Value("All_Color_AS_Black") = False
                        oOptions.Value("Remove_Line_Weights") = True
                        oOptions.Value("Vector_Resolution") = 5000
                        ' Define various settings and input to provide the translator.  
                        oContext.Type = IOMechanismEnum.kFileBrowseIOMechanism
                        Dim oData As DataMedium
                        oData = _invApp.TransientObjects.CreateDataMedium
                        'create save locations
                        If My.Settings.PDFRev = True Then
                            RevNo = "-R" & oDoc.PropertySets.Item("{F29F85E0-4FF9-1068-AB91-08002B27B3D9}").ItemByPropId("9").Value
                        Else
                            RevNo = ""
                        End If
                        PDFSource = Destin & "\" & DrawingName
                        PDFSource = PDFSource.Insert(PDFSource.LastIndexOf("."), RevNo)
                        'PDFSource = Destin & "\" & DrawingName & "-R" & RevNo
                        PDFSource = Replace(PDFSource, ".idw", ".pdf")
                        'check if the file exists. If the directory is missing, create a new one
                        If chkPDF.CheckState = CheckState.Checked Then
                            If My.Computer.FileSystem.DirectoryExists(Strings.Left(PDFSource, InStrRev(PDFSource, "\"))) = False Then
                                MkDir(Strings.Left(PDFSource, InStrRev(PDFSource, "\")))
                            End If
                            If My.Computer.FileSystem.DirectoryExists(Strings.Left(PDFSource, InStrRev(PDFSource, "\")) & "Archived\") = False Then
                                MkDir(Strings.Left(PDFSource, InStrRev(PDFSource, "\")) & "Archived\")
                            End If
                            'if an older revision exists move it into the archive folder
                            'if the archive folder doesn't exist create a new one
                            Dim Filetype As String = ".pdf"
                            Search_For_Duplicates(PDFSource, DrawingName, Filetype)
                            Dim oDocument As Document
                            oDocument = _invApp.ActiveDocument
                            ' Call the translator.  
                            oData.FileName = PDFSource
                            Call oPDFTrans.SaveCopyAs(oDocument, oContext, oOptions, oData)
                            writeDebug("Created PDF: " & PDFSource)
                        End If
                        'If the drawingcreation causes an error, save drawing name and inform user when process is complete
                        CloseLater(DrawingName, oDoc)
                    Catch
                        ErrorReport = DrawingName & vbNewLine
                    End Try
                End If

                Title = "Saving: "
                ProgressBar(Total, Counter, Title, Replace(DrawingName, ".idw", ".pdf"))
                Counter += 1
            End If
        Next
        _invApp.SilentOperation = False
        MsVistaProgressBar.Visible = False
        'notify user of drawings that caused problems
        If Len(ErrorReport) > 0 Then
            MsgBox("The following files were Not created: " & vbNewLine & vbNewLine & ErrorReport & vbNewLine &
                   "Check to make sure the drawing isn't open or write-protected.")
            writeDebug("The following files were Not created:  " & vbNewLine & vbNewLine & ErrorReport & vbNewLine &
                   "Check to make sure the drawing isn't open or write-protected.")
        End If
    End Sub
    Public Sub DXFDWGCreator(Path As Documents, ByRef oDoc As Document, ByRef Destin As String, ByRef Archive As String _
                             , ByRef DrawingName As String, ByRef DrawSource As String, OpenDocs As ArrayList _
                             , Total As Integer, Counter As Integer, Source As String, ExportType As String)
        Dim sReadableType As String = ""
        Dim Title, RevNo As String
        Dim NotMade As String = ""
        _invApp.SilentOperation = True
        'Go through drawings to see which ones are selected
        For X = 0 To LVSubFiles.Items.Count - 1
            'Look through all sub files in open documents to get the part sourcefile
            MatchPart(DrawSource, DrawingName, X)
            If chkSkipAssy.Checked = True And DrawSource.Contains(".iam") Then
            Else
                If LVSubFiles.Items(X).Checked = True Then
                    'iterate through opend documents to find the selected file
                    MatchDrawing(DrawSource, DrawingName, X)
                    'open drawing
                    oDoc = _invApp.Documents.Open(DrawSource, False)
                    If My.Settings(ExportType & "Rev") = True Then
                        RevNo = "-R" & oDoc.PropertySets.Item("{F29F85E0-4FF9-1068-AB91-08002B27B3D9}").ItemByPropId("9").Value
                    Else
                        RevNo = ""
                    End If
                    Source = Destin & "\" & DrawingName
                    Source = Source.Insert(Source.LastIndexOf("."), RevNo)
                    'PDFSource = Destin & "\" & DrawingName & "-R" & RevNo
                    Source = Replace(Source, "idw", ExportType)
                    'DXFSource = DrawSource.Insert(DrawSource.LastIndexOf("\"), "\DWG_DXF")
                    'DXFSource = DXFSource.Insert(DXFSource.LastIndexOf("."), "-R" & RevNo)
                    'DXFSource = Replace(DXFSource, "idw", "dxf")
                    If My.Computer.FileSystem.DirectoryExists(Strings.Left(Source, InStrRev(Source, "\"))) = False Then
                        MkDir(Strings.Left(Source, InStrRev(Source, "\")))
                    End If
                    If My.Computer.FileSystem.DirectoryExists(Strings.Left(Source, InStrRev(Source, "\")) & "Archived\") = False Then
                        MkDir(Strings.Left(Source, InStrRev(Source, "\")) & "Archived\")
                    End If

                    'if an older revision exists move it into the archive folder
                    'if the archive folder doesn't exist create a new one
                    Dim Filetype As String = "." & ExportType
                    Search_For_Duplicates(Source, DrawingName, Filetype)
                    If My.Computer.FileSystem.FileExists(Strings.Replace(DrawSource, "idw", "ipt")) = True Then
                        Archive = Strings.Replace(DrawSource, "idw", "ipt")
                        oDoc = _invApp.Documents.Open(Archive, True)
                        SheetMetalTest(Archive, oDoc, sReadableType)
                        If sReadableType = "P" And ExportType = "dxf" Then
                            Call ExportPart(DrawSource, Archive, False, Destin, DrawingName, OpenDocs, ExportType, RevNo)
                        ElseIf sReadableType = "S" And ExportType = "dxf" And chkUseDrawings.Checked = False Then
                            Call SMDXF(oDoc, Source, True, Replace(DrawingName, ".idw", ".dxf"))
                            CloseLater(Strings.Left(DrawingName, Len(DrawingName) - 3) & "ipt", oDoc)
                        ElseIf sReadableType = "S" And ExportType = "dxf" And chkUseDrawings.Checked = True Then
                            Call ExportPart(DrawSource, Archive, False, Destin, DrawingName, OpenDocs, "dxf", RevNo)
                        ElseIf sReadableType <> "" And ExportType = "dwg" Then
                            Call ExportPart(DrawSource, Archive, False, Destin, DrawingName, OpenDocs, "dwg", RevNo)
                        ElseIf sReadableType = "" Then
                            CloseLater(Strings.Right(oDoc.FullFileName, Len(oDoc.FullFileName) - InStrRev(oDoc.FullFileName, "\")), oDoc)
                        End If
                        If chkCheck.CheckState = CheckState.Indeterminate Then
                            chkCheck.CheckState = CheckState.Unchecked
                        End If
                    ElseIf My.Computer.FileSystem.FileExists(Strings.Replace(DrawSource, "idw", "iam")) = False Then
                        NotMade = DrawingName & vbNewLine
                    ElseIf My.Computer.FileSystem.FileExists(Strings.Replace(DrawSource, "idw", "iam")) = True Then
                        Call ExportPart(DrawSource, Archive, False, Destin, DrawingName, OpenDocs, ExportType, RevNo)
                    End If
                    Title = "Saving:  "
                    ProgressBar(Total, Counter, Title, Replace(DrawingName, ".idw", "." & ExportType))
                    Counter += 1
                End If
            End If
        Next
        If NotMade <> "" Then
            MsgBox("There was a problem generating the following files:" & vbNewLine & vbNewLine & NotMade)
        End If
        _invApp.SilentOperation = False
    End Sub
    Private Sub ChangeRev()
        Dim Ans As String
        Ans = MsgBox("This will reset all current revisions" _
                     & vbNewLine & "Continue?", MsgBoxStyle.OkCancel, "Change Revision")
        If Ans = vbCancel Then Exit Sub
        Dim Rev_Switch As New Rev_Switch
        Dim RevType As Integer = 0
        Rev_Switch.PopMain(Me)
        Rev_Switch.First(RevType)
    End Sub
    Public Sub ChangeRev(ByRef RevType As Integer)
        Dim oDoc As Document = Nothing
        Dim dDoc As DrawingDocument = Nothing
        Dim Path As Documents = _invApp.Documents
        Dim DrawingName As String = Nothing
        Dim Archive As String = Nothing
        Dim DrawSource As String = Nothing
        Dim OpenDocs As New ArrayList
        Dim Y, Q As Integer
        Q = 0
        CreateOpenDocs(OpenDocs)
        _invApp.SilentOperation = True
        For Y = 0 To LVSubFiles.Items.Count - 1
            If LVSubFiles.Items(Y).Checked = True Then
                Q += 1
                MatchDrawing(DrawSource, DrawingName, Y)
                'open drawing
                oDoc = _invApp.Documents.Open(DrawSource, True)
                Dim Sheets As Sheets
                Dim Sheet As Sheet
                Dim Rev, RevNo As String
                Dim oPoint(0 To oDoc.sheets.count * 2) As String
                Dim Tx(0 To oDoc.sheets.count) As Decimal
                Dim Ty(0 To oDoc.sheets.count) As Decimal
                Dim Adjust As String = ""
                Dim Point As Point2d
                Dim oRevTable As RevisionTable
                Dim oTitleblock As TitleBlock
                Dim C, R, i As Integer
                MsVistaProgressBar.Visible = True
                'Check the titleblock to see if it is the IMM TitleBlock, inform user if it is not
                Dim oTitleBlockDef As TitleBlockDefinition
                oTitleBlockDef = oDoc.TitleBlockDefinitions.Item(1)
                'If Strings.InStr(oTitleBlockDef.Name, "IMM") = 0 Then
                '    MsgBox("This operation cannot function on" & vbNewLine &
                '           "non-IMM title blocks")
                '    Exit Sub
                'End If
                'Set the point at which the rev table should be inserted
                ProgressBar(LVSubFiles.CheckedItems.Count, Q, "Changing Rev: ", Strings.Right(oDoc.FullFileName, Len(oDoc.FullFileName) - InStrRev(oDoc.FullFileName, "\")))
                RevNo = oDoc.PropertySets.Item("{F29F85E0-4FF9-1068-AB91-08002B27B3D9}").ItemByPropId("9").Value
                oRevTable = oDoc.activesheet.RevisionTables(1)
                'oRevTable.RevisionTableRows.Add()
                If IsNumeric(RevNo) Then
                    Rev = RevNo
                Else
                    Rev = Asc(RevNo) - Asc("A")
                End If
                Dim X As Integer = 0
                'Iterate through each sheet of the open drawing
                Sheets = oDoc.Sheets
                Sheets.Item(1).Activate()
                Dim S As Integer = 0
                'Delete each rev table in each sheet
                Sheets.Item(1).Activate()
                For Each Sheet In Sheets
                    Sheet.Activate()

                    oRevTable = oDoc.activesheet.RevisionTables(1)
                    Tx(S) = oRevTable.Position.X
                    Ty(S) = oRevTable.Position.Y
                    S += 1
                Next
                S = 0
                For Each Sheet In Sheets
                    If S = 0 Then

                        Sheet.Activate()
                        oRevTable = oDoc.activesheet.RevisionTables(1)
                        oRevTable.RevisionTableRows.Add()

                        For Each oRow As RevisionTableRow In oRevTable.RevisionTableRows
                            If oRow.IsActiveRow Then
                            Else
                                oRow.Delete()
                            End If
                        Next
                        oTitleblock = oDoc.activesheet.TitleBlock
                        If S = 0 Then
                            Adjust = Ty(S) - oRevTable.Position.Y
                        End If

                        oRevTable.Delete()
                        S += 1
                    ElseIf S > 0 Then
                        Sheet.Activate()
                        oRevTable = oDoc.activesheet.RevisionTables(1)
                        oRevTable.Delete()
                    End If
                Next Sheet
                S = 0
                Sheets.Item(1).Activate()
                For Each Sheet In Sheets
                    Sheet.Activate()
                    'Set counter to numeric or alpha based on previous RevTable
                    Point = _invApp.TransientGeometry.CreatePoint2d(Tx(S), Ty(S) - Adjust)

                    If RevType = 1 Then
                        oRevTable = oDoc.ActiveSheet.RevisionTables.add2(Point, False, True, False, 0)
                    Else
                        oRevTable = oDoc.ActiveSheet.RevisionTables.Add2(Point, False, True, True, "A")
                    End If
                    'Add items corresponding to which table form is being populated
                    S += 1
                Next Sheet
                Sheet = oDoc.activesheet
                Sheet.Activate()
                oRevTable = Sheet.RevisionTables(1)
                'Iterate through the new revision table and insert the standard information for the respective table
                C = oRevTable.RevisionTableColumns.Count
                R = oRevTable.RevisionTableRows.Count
                Dim Contents(0 To C * R) As String
                Dim Rtr As RevisionTableRow
                Dim RTCell As RevisionTableCell
                i = 1
                For Each Rtr In oRevTable.RevisionTableRows
                    For Each RTCell In Rtr
                        If i = 3 And RevType = 1 Then
                            RTCell.Text = "APPROVED FOR MANUFACTURING"
                            Exit For
                        ElseIf i = 3 And RevType = 0 Then
                            RTCell.Text = "ISSUED FOR REVIEW"
                            Exit For
                        End If
                        i += 1
                    Next
                Next
                oDoc.sheets.item(1).activate()
                _invApp.SilentOperation = True
                Try
                    oDoc.Save()
                Catch ex As Exception
                    MessageBox.Show(ex.Message, "Exception Details", MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
                End Try
                CloseLater(DrawingName, oDoc)
                _invApp.SilentOperation = False
            End If
        Next
        _invApp.SilentOperation = False
        MsVistaProgressBar.Visible = False
    End Sub
    Private Sub RRev()
        Dim OpenDocs As New ArrayList
        CreateOpenDocs(OpenDocs)
        RevTable.PopMain(Me)
        Dim oDoc As Document
        Dim Path As Documents
        'Dim oRevTable As RevisionTable
        'Dim Point As Point2d
        Path = _invApp.Documents
        'Dim Loc As Drawing.Point
        'Dim Sheets As Sheets
        'Dim Sheet As Sheet
        Dim X As Integer = 0
        'Dim oTitleblock As TitleBlock
        Dim DwgName, strPath, strFile, RevNo As String
        RevNo = ""
        DwgName = ""
        RevTable.btnIgnore.Visible = False
        chkiProp.CheckState = CheckState.Unchecked
        'iterate through the files in subfiles
        For X = 0 To LVSubFiles.Items.Count - 1
            'select the files that have been checked
            If LVSubFiles.Items(X).Checked = True Then
                'save the name of the checked file
                DwgName = LVSubFiles.Items.Item(X).Text
                Exit For
            End If
        Next
        'Iterate through all the open documents in inventor
        For j = 1 To Path.Count
            oDoc = Path.Item(j)

            'Parse the path from the document
            strPath = Strings.Left(oDoc.FullFileName, Strings.Len(oDoc.FullFileName) - 3) & "idw"
            'Parse out the filename from the document
            strFile = Strings.Right(strPath, Strings.Len(strPath) - InStrRev(strPath, "\"))
            'Compare the selected file to the current document, if they are the same proceed
            If Trim(DwgName) = strFile Then
                'Modify the revtable layout to give instruction for removing revisions
                'Open the selected document in the background
                oDoc = _invApp.Documents.Open(strPath, False)
                'Go through each sheet to check for revisiontable
                'Do Until X > 1
                'Dim S As Integer = 0
                'Sheets = oDoc.Sheets
                'Dim oPoint(0 To oDoc.sheets.count * 2) As String
                'For Each Sheet In Sheets
                ' 'For each sheet, on the first cycle, delete the table, and create a new table on the second round
                ' If X <> 1 Then
                ' Sheet.Activate()
                ' On Error Resume Next
                ' oRevTable = oDoc.activesheet.RevisionTables(1)
                ' oTitleblock = oDoc.activesheet.TitleBlock
                ' oPoint(S) = oRevTable.Position.X
                ' oPoint(S + 1) = oRevTable.RangeBox.MinPoint.Y + 1.14359999263287
                ' oRevTable.Delete()
                ' S += 2
                'Else
                '   Sheet.Activate()
                '  Point = _invApp.TransientGeometry.CreatePoint2d(oPoint(S), oPoint(S + 1))
                ' oRevTable = oDoc.ActiveSheet.Revisiontables.Add2(Point, False, True, False, 0)
                'S += 2
                'End If
                'Next
                'X += 1
                'Loop
                Call RevTable.PopulateRevTable(oDoc, RevNo, 1, False)
                Exit For
            End If
        Next
        'Show the userform
        RevTable.Show()
    End Sub
    Private Sub btnRef_Click(sender As System.Object, e As System.EventArgs) Handles btnRef.Click
        Dim Written As Boolean = False
        Dim Warning As Boolean = False
        Dim Warning2 As Boolean = False
        Dim PropExists As Boolean = True
        Dim Fix As String = "The following files have more references" & vbNewLine &
                            "than the titleblock permits: " & vbNewLine & vbNewLine
        Dim C, X, Y, R, P, Maxlayer, Layer, Col As Integer
        Col = R = P = 0
        Layer = 0
        'Dim Parent As New ArrayList
        Dim Parent As New List(Of List(Of String))
        Dim Total As Integer = LVSubFiles.CheckedItems.Count
        Dim Counter As Integer = 1
        Parent.Clear()
        Dim oDoc As Document
        Dim OpenDocs As New ArrayList
        Dim AsmDoc As AssemblyDocument
        Dim DocType As DocumentTypeEnum
        Dim Ans, strFile, PartName, Archive, DrawSource As String
        Dim DrawingName As String = ""
        Dim test, ChildFile, ParentFile As String
        ChildFile = ""
        Dim Path As Documents
        Path = _invApp.Documents
        'Go through drawings to see which ones are selected
        For X = 0 To lstOpenfiles.Items.Count - 1
            'Look through all sub files in open documents to get the part sourcefile
            If lstOpenfiles.GetItemCheckState(X) = CheckState.Checked Then
                'iterate through opend documents to find the selected file
                If Strings.Right(lstOpenfiles.Items.Item(X), 3) = "iam" Then
                    C += 1
                End If
            End If
        Next
        If C = 0 Then
            MsgBox("Please select an assembly")
            Exit Sub
        End If
        If C > 1 Then
            MsgBox("Please only select one assembly")
            Exit Sub
        End If
        'Go through drawings to see which ones are selected
        For X = 0 To lstOpenfiles.Items.Count - 1
            'Look through all sub files in open documents to get the part sourcefile
            If lstOpenfiles.GetItemCheckState(X) = CheckState.Checked Then
                'iterate through opend documents to find the selected file
                For Y = 1 To _invApp.Documents.Count
                    'If selected document = Found document
                    If InStr(_invApp.Documents.Item(Y).FullFileName, lstOpenfiles.Items.Item(X)) <> 0 Then
                        oDoc = _invApp.Documents.Open(_invApp.Documents.Item(Y).FullFileName)
                        Ans = MsgBox("This action will overwrite" & vbNewLine & "all previous references." _
                                     & vbNewLine & "Do you wish to continue?", vbYesNo, "Overwrite References")
                        If Ans = vbNo Then
                            Exit Sub
                        Else
                            ' activate matching document
                            AsmDoc = _invApp.ActiveDocument
                            DocType = AsmDoc.DocumentType
                            strFile = Strings.Left(AsmDoc.FullFileName, Strings.Len(AsmDoc.FullFileName) - 4)
                            strFile = Strings.Right(strFile, Strings.Len(strFile) - InStrRev(strFile, "\"))
                            'Add list to the array for new Parent
                            Parent.Add(New List(Of String))
                            Parent(Layer).Add(strFile)
                            'call recursive sub to go through all sub files
                            WriteChildren(AsmDoc.ComponentDefinition.Occurrences, Layer + 1, R, Col, Maxlayer, Parent, strFile, Total, Counter)
                            Exit For
                        End If
                    End If
                Next
            End If
        Next
        Counter = 1
        For X = 0 To LVSubFiles.Items.Count - 1
            Counter += 1
            For Y = 1 To _invApp.Documents.Count - 1
                oDoc = Path.Item(Y)
                Archive = oDoc.FullFileName
                'Use the Partsource file to create the drawingsource file
                DrawSource = Strings.Left(Archive, Strings.Len(Archive) - 3) & "idw"
                If My.Computer.FileSystem.FileExists(DrawSource) Then
                    DrawingName = Strings.Right(DrawSource, Strings.Len(DrawSource) - Strings.InStrRev(DrawSource, "\"))
                    'If the drawing file is listed, open the drawing in Inventor
                    If Trim(LVSubFiles.Items.Item(X).Text) = DrawingName Then
                        oDoc = _invApp.Documents.Open(DrawSource, False)
                        CustomPropSet = oDoc.PropertySets.Item("{D5CDD505-2E9C-101B-9397-08002B2CF9AE}")
                        PartName = Strings.Right(oDoc.FullFileName, Len(oDoc.FullFileName) - InStrRev(oDoc.FullFileName, "\"))


                        For P = 0 To 9
                            Try
                                InvRef(P) = CustomPropSet.Item("Reference" & P)
                            Catch
                                SetiProp(InvRef(P), "Reference" & P)
                                InvRef(P) = CustomPropSet.Item("Reference" & P)
                            End Try
                        Next
                        For P = 0 To 9
                            InvRef(P).Value = ""
                        Next
                        P = 0
                        Dim Fullstring As New List(Of String)
                        For Each l As List(Of String) In Parent
                            For M = 0 To l.Count - 1
                                test = l(M)
                                If InStr(test, ",") = 0 Then
                                    ParentFile = test
                                    ChildFile = ""
                                Else
                                    ChildFile = Strings.Right(test, Len(test) - (InStrRev(test, ",")))
                                    ParentFile = Strings.Left(test, (InStr(test, ",") - 1))
                                End If
                                If (Strings.Left(PartName, Len(PartName) - 4) = ChildFile) And ChildFile <> "" Then
                                    If P = 7 Then
                                        Warning2 = True
                                        If InStr(Fix, ChildFile) = 0 Then
                                            Fix = Fix & ChildFile & vbNewLine
                                        End If
                                    ElseIf P = 6 Then
                                        MsgBox("The current titleblock holds 6 references." & vbNewLine &
                                            "All following references will be added, but the Titleblock needs to be updated before they can be shown.")
                                    ElseIf P = 10 Then
                                        MsgBox("Currently the program can only keep 10 references." & vbNewLine &
                                            "All following references will not be added.")
                                    End If
                                    Try
                                        InvRef(P) = CustomPropSet.Item("Reference" & P)
                                        InvRef(P).Value = ParentFile
                                    Catch
                                        InvRef(P) = CustomPropSet.Add("", "Reference" & P)
                                        InvRef(P).Value = ParentFile
                                    End Try
                                    P += 1
                                    Written = True
                                End If
                            Next
                        Next
                        _invApp.SilentOperation = True
                        oDoc.Update()
                        Try
                            oDoc.Save2()
                        Catch ex As Exception
                            MessageBox.Show(ex.Message, "Exception Details", MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
                        End Try
                        CloseLater(DrawingName, oDoc)
                        _invApp.SilentOperation = False
                    End If
                End If
                'MsVistaProgressBar.ProgressBarStyle = MSVistaProgressBar.BarStyle.Marquee
                ProgressBar(Total, Counter, "Writing Reference: ", DrawingName)
                If Written = True Then
                    Written = False
                    Exit For
                End If
            Next
        Next
        MsVistaProgressBar.Visible = False
        If Warning2 = True Then
            MsgBox(Fix & vbNewLine & "These drawings need to be updated manually.")
        End If

        MsgBox("All the references have been updated")
    End Sub
    Private Sub SetiProp(InvRef As Inventor.Property, Ref As String)
        InvRef = CustomPropSet.Add("", Ref)
    End Sub
    Private Sub chkDXF_CheckedChanged(sender As System.Object, e As System.EventArgs) Handles chkDXF.CheckedChanged
        If chkDWG.Checked = True Then Exit Sub
        If chkDXF.CheckState = CheckState.Checked Then
            chkUseDrawings.CheckState = CheckState.Unchecked
            chkUseDrawings.Visible = True
            chkSkipAssy.CheckState = CheckState.Checked
            chkSkipAssy.Visible = True
            chkClose.Top = chkClose.Top + 30
            GroupBox4.Height = GroupBox4.Size.Height.ToString + 30
        Else
            chkUseDrawings.CheckState = CheckState.Checked
            chkUseDrawings.Visible = False
            chkSkipAssy.Visible = False
            chkClose.Top = chkClose.Top - 30
            GroupBox4.Height = GroupBox4.Size.Height.ToString - 30
        End If
    End Sub
    Private Sub chkExport_CheckedChanged(sender As Object, e As EventArgs) Handles chkExport.CheckedChanged
        If chkExport.CheckState = CheckState.Checked Then
            chkDWG.Visible = True
            chkPDF.Visible = True
            chkDXF.Visible = True
            chkClose.Top = chkClose.Top + 15
            GroupBox4.Height = GroupBox4.Height + 15
            chkPDF.CheckState = CheckState.Unchecked
            chkDWG.CheckState = CheckState.Unchecked
            chkDXF.CheckState = CheckState.Unchecked
        Else
            chkDWG.Visible = False
            chkPDF.Visible = False
            chkDXF.Visible = False
            chkDXF.CheckState = CheckState.Unchecked
            chkDWG.CheckState = CheckState.Unchecked
            chkPDF.CheckState = CheckState.Unchecked
            chkClose.Top = chkClose.Top - 15
            GroupBox4.Height = GroupBox4.Height - 15
        End If
    End Sub
    Private Sub chkDWG_CheckedChanged(sender As Object, e As EventArgs) Handles chkDWG.CheckedChanged
        If chkDXF.CheckState = CheckState.Checked Then Exit Sub
        If chkDWG.CheckState = CheckState.Checked Then
            chkUseDrawings.CheckState = CheckState.Unchecked
            chkUseDrawings.Visible = True
            chkSkipAssy.CheckState = CheckState.Checked
            chkSkipAssy.Visible = True
            chkClose.Top = chkClose.Top + 30
            GroupBox4.Height = GroupBox4.Size.Height.ToString + 30
        Else
            chkUseDrawings.CheckState = CheckState.Checked
            chkUseDrawings.Visible = False
            chkSkipAssy.Visible = False
            chkClose.Top = chkClose.Top - 30
            GroupBox4.Height = GroupBox4.Size.Height.ToString - 30
        End If
    End Sub

#End Region
#Region "Background Calls"
    Private Sub lstOpenfiles_ItemCheck(sender As Object, e As System.Windows.Forms.ItemCheckEventArgs) Handles lstOpenfiles.ItemCheck
        'Run checkbox click routine through a timer so that the most recent item will show up in the selection
        tmr.Enabled = False
        tmr.Enabled = True
    End Sub
    Private Sub TestForDrawing(Partsource As String, ByVal Level As Integer, ByRef Total As Integer,
                               ByRef Counter As Integer, OpenDocs As ArrayList, ByRef Elog As String, Ref As Boolean)
        Dim Dup As Boolean = False
        Dim strFile, DrawingName As String
        'Extract the file name from the full file path
        strFile = Strings.Right(Partsource, Strings.Len(Partsource) - InStrRev(Partsource, "\"))
        'Assume drawing is stored in same location as part/assembly (same inherent procedure as Inventor)
        'Create drawing name for search purposes
        DrawingName = Strings.Left(strFile, (Strings.Len(strFile) - 3)) & "idw"
        'Test to see if the file exists in the expected
        ProgressBar(Total, Counter, "Found: ", DrawingName)
        If Not My.Computer.FileSystem.FileExists(Partsource) Then
            DrawingName = DrawingName & "(PPM)"
        ElseIf Not My.Computer.FileSystem.FileExists(Strings.Left(Partsource, (Strings.Len(Partsource) - 3)) & "idw") Then
            DrawingName = DrawingName & "(DNE)"
        ElseIf Ref = True Then
            DrawingName = DrawingName & "(REF)"
        End If


        Elog = Elog & DrawingName & ": "
        'Check to see if the part has already been added to the list
        If SubFiles.Count = 0 Then
            SubFiles.Add(New KeyValuePair(Of String, String)((Space(Level * 3) & DrawingName).ToString, Partsource))
            AlphaSub.Add((DrawingName).ToString, Partsource)
            Elog = Elog & "Added to Open File list" & vbNewLine
            AddtoRenameTable(Partsource, DrawingName, strFile, True, "")
            Counter += 1
        Else
            For x = 0 To SubFiles.Count - 1
                If Trim(SubFiles.Item(x).Key).Contains(DrawingName) Then
                    Elog = Elog & "Skipped in Open File list" & vbNewLine
                    Exit Sub
                End If
            Next
            'Tab over part to show it is a sub-component
            SubFiles.Add(New KeyValuePair(Of String, String)((Space(Level * 3) & DrawingName).ToString, Partsource))
            AlphaSub.Add((DrawingName).ToString, Partsource)
            Elog = Elog & "Added to Open File list" & vbNewLine
            Counter += 1
        End If

        'End If

        'For Each pair As KeyValuePair(Of String, String) In AlphaSub

        'Next
        ' AlphaSub = SubFiles.ToList
        ' AlphaSub.Sort(Function(firstpair As KeyValuePair(Of String, String), nextpair As KeyValuePair(Of String, String)) firstpair.Key.CompareTo(nextpair.Key))
        'AlphaSub = AlphaSub.Sort()
    End Sub
    Private Sub CheckForDev(PartSource As String, ByRef Total As Integer, ByRef Counter As Integer, OpenDocs As ArrayList, Elog As String, ByVal Level As Integer)
        If chkDerived.Checked = False Then Exit Sub
        Dim oAsmDoc As AssemblyDocument = Nothing
        Dim oDoc, oDerived As Document
        Dim DerDrawing As String
        oDoc = _invApp.Documents.Open(PartSource, False)
        If oDoc.DocumentType = DocumentTypeEnum.kPartDocumentObject And oDoc.ReferencedDocuments.Count > 0 Then
            For X = 1 To oDoc.ReferencedDocuments.Count
                oDerived = oDoc.ReferencedDocuments.Item(X)
                DerDrawing = oDerived.PropertySets.Item("{32853F0F-3444-11D1-9E93-0060B03C1CA6}").ItemByPropId("5").Value.ToString & ".idw"
                PartSource = oDerived.FullFileName
                If oDerived.DocumentType = DocumentTypeEnum.kAssemblyDocumentObject Then
                    Total = Total + oDoc.ReferencedDocuments.Count - 1
                    If _invApp.ActiveDocument.FullFileName = PartSource Then
                        oAsmDoc = _invApp.ActiveDocument
                    Else
                        Try
                            oAsmDoc = _invApp.Documents.Open(PartSource, False)
                        Catch ex As Exception
                            MsgBox("Something went wrong accessing " & PartSource & vbNewLine &
                           "Please verify the file exists And/Or check" & vbNewLine &
                           "if there are invisible documents open")
                            Exit Sub
                        End Try
                    End If

                    TraverseAssembly(oAsmDoc.ComponentDefinition.Occurrences, PartSource, Level, Total, Counter, OpenDocs, Elog, True)
                Else
                    TestForDrawing(PartSource, Level, Total, Counter, OpenDocs, Elog, False)
                End If
            Next
        End If
        'CloseLater(Strings.Right(PartSource, (Strings.Len(PartSource) - InStrRev(PartSource, "\"))), oDoc)
    End Sub
    Private Sub TraverseAssemblyLoad(PartSource As String, Level As Integer, ByRef Total As Integer,
                                     ByRef Counter As Integer, Opendocs As ArrayList, ByRef Elog As String, ByRef Dev As Boolean)
        Dim oAsmDoc As AssemblyDocument = Nothing
        Dim strFile, DrawingName, ID As String
        Dim Add As String = "True"
        Dim RenameList As New List(Of String)
        'Open assembly in background
        If _invApp.ActiveDocument.FullFileName = PartSource Then
            oAsmDoc = _invApp.ActiveDocument
        Else
            Try
                oAsmDoc = _invApp.Documents.Open(PartSource, False)
            Catch ex As Exception
                MsgBox("Something went wrong accessing " & PartSource & vbNewLine &
                           "Please verify the file exists And/Or check" & vbNewLine &
                           "if there are invisible documents open")
                Exit Sub
            End Try
        End If
        PartSource = oAsmDoc.FullFileName
        Rename.PopMain(Me)
        'Create a drawing name assuming the name/location is the same as the part
        strFile = Strings.Right(PartSource, Strings.Len(PartSource) - InStrRev(PartSource, "\"))
        'DrawingName = Strings.Left(strFile, (Strings.Len(strFile) - 3)) & "idw"
        'If Strings.InStr(PartSource, "Content Center") = 0 Then
        If Dev = False Then
            If Strings.InStr(oAsmDoc.File.FullFileName, "Purchased Parts") > 0 Then
                ID = "PP"
            ElseIf Strings.InStr(oAsmDoc.File.FullFileName, "Content Center") > 0 Then
                ID = "CC"
            Else
                ID = oAsmDoc.PropertySets.Item("{F29F85E0-4FF9-1068-AB91-08002B27B3D9}").ItemByPropId("5").Value
            End If
        Else
            ID = "DP"
        End If

        AddtoRenameTable(PartSource, DrawingName, strFile, Add, ID)
        'End If
        'Check to see if the drawing exists and add to the list
        TestForDrawing(PartSource, Level, Total, Counter, Opendocs, Elog, False)
        'If the part is an assembly, go through recursively and retrieve the sub-components
        ' Get all of the referenced documents. 
        Dim oRefDocs As DocumentsEnumerator
        oRefDocs = oAsmDoc.AllReferencedDocuments

        ' Iterate through the list of documents. 
        Dim oRefDoc As Document
        For Each oRefDoc In oRefDocs
            Total = Total + 1
        Next
        TraverseAssembly(oAsmDoc.ComponentDefinition.Occurrences, PartSource, Level + 1, Total, Counter, Opendocs, Elog, Dev)
        CloseLater(strFile, oAsmDoc)
    End Sub
    Private Sub TraverseAssembly(Occurrences As ComponentOccurrences, PartSource As String, Level As Integer, ByRef Total As Integer, ByRef Counter As Integer,
                                 OpenDocs As ArrayList, ByRef Elog As String, ByVal Dev As Boolean)
        'Count the level of the sub-component
        Dim Add As Boolean = True
        Dim strFile, DrawingName, ID As String
        Dim oOcc As ComponentOccurrence
        Dim Ref As Boolean
        'iterate through each occurrence in the assembly
        'Total = Total + Occurrences.Count
        For Each oOcc In Occurrences
            Try
                If oOcc.BOMStructure = BOMStructureEnum.kReferenceBOMStructure Then
                    Ref = True
                Else
                    Ref = False
                End If

                PartSource = oOcc.Definition.Document.fullfilename
                'setting the mass to something updates the mass property in the drawing
                'create drawing name
                strFile = Strings.Right(PartSource, Strings.Len(PartSource) - InStrRev(PartSource, "\"))
                DrawingName = Strings.Left(strFile, (Strings.Len(strFile) - 3)) & "idw"
                ProgressBar(Total, Counter, "Found:  ", DrawingName)
                'FindWorker.ReportProgress(CInt((Counter / Total) * 100))
                'Add to list if the drawing exists
                TestForDrawing(PartSource, Level, Total, Counter, OpenDocs, Elog, Ref)
                'iterate through again for each sub-assembly found 
                'Counter += 1
                'Create a list of sub parts for use in the renaming section. this saves time having to recreate it again later.
                'If VBAFlag <> "NA" And Strings.InStr(PartSource, "Content Center") = 0 Then 
                If Dev = False Then
                    If Strings.InStr(PartSource, "Purchased Parts") > 0 Then
                        ID = "PP"
                    ElseIf Strings.InStr(PartSource, "Content Center") > 0 Then
                        ID = "CC"
                    Else
                        ID = oOcc.Definition.Document.propertysets.Item("{F29F85E0-4FF9-1068-AB91-08002B27B3D9}").ItemByPropId("5").Value
                    End If
                Else
                    ID = "DP"
                End If
                AddtoRenameTable(PartSource, DrawingName, strFile, Add, ID)
                'If the sub assembly is an assembly itself, redo the iteration until all parts have been found.
                If oOcc.DefinitionDocumentType = DocumentTypeEnum.kAssemblyDocumentObject Then
                    'Level += 1
                    TraverseAssembly(oOcc.SubOccurrences, PartSource, Level + 1, Total, Counter, OpenDocs, Elog, Dev)
                ElseIf oOcc.DefinitionDocumentType = DocumentTypeEnum.kPartDocumentObject Then
                    CheckForDev(PartSource, Total, Counter, OpenDocs, Elog, Level + 1)
                End If
            Catch

            End Try

        Next
    End Sub
    Private Sub AddtoRenameTable(ByVal PartSource As String, ByVal DrawingName As String, ByVal strFile As String,
                                 ByVal add As Boolean, ByVal ID As String)
        Dim oDoc As Document = Nothing
        Dim exists As Boolean
        Dim RenameList As New List(Of String)
        RenameList.Add(Strings.Left(PartSource, InStrRev(PartSource, "\")))
        If My.Computer.FileSystem.FileExists(Strings.Left(PartSource, InStrRev(PartSource, "\")) & DrawingName) Then
            RenameList.Add(DrawingName)
        Else
            RenameList.Add("")
        End If
        RenameList.Add(strFile)
        For X = 0 To RenameTable.Count - 1
            If Trim(RenameTable(X).Item(2).ToString) = Trim(strFile) Or
                My.Computer.FileSystem.FileExists(PartSource) = False Then
                add = False
                Exit For
            Else
                add = True
            End If
        Next
        'pull the thumbnail from Inventor. This must be done through inventor's 32-bit api. 
        If add = True Then
            RenameTable.Add(RenameList)
            Dim oP As Inventor.InventorVBAProject
            For Each oP In _invApp.VBAProjects
                If oP.Name = "ApplicationProject" Then
                    Dim oC As Inventor.InventorVBAComponent
                    For Each oC In oP.InventorVBAComponents
                        If oC.Name = "Get64BitPicture" Then
                            exists = True
                            Dim oM As Inventor.InventorVBAMember
                            For Each oM In oC.InventorVBAMembers
                                If oM.Name = "Thumbnail" Then
                                    If My.Computer.FileSystem.FileExists(My.Computer.FileSystem.SpecialDirectories.Temp & "\Thumbnail.jpg") Then
                                        Kill(My.Computer.FileSystem.SpecialDirectories.Temp & "\Thumbnail.jpg")
                                    End If

                                    IO.File.WriteAllText(My.Computer.FileSystem.SpecialDirectories.Temp & "\PartSource.txt", PartSource)
                                    oM.Execute()
                                    Kill(My.Computer.FileSystem.SpecialDirectories.Temp & "\PartSource.txt")
                                    Exit For
                                End If
                            Next oM
                        End If
                    Next oC
                End If
            Next

            If IO.File.Exists(My.Computer.FileSystem.SpecialDirectories.Temp & "\Thumbnail.jpg") Then
                Dim Thumbnail As Image = Image.FromFile(My.Computer.FileSystem.SpecialDirectories.Temp & "\Thumbnail.jpg", False)
                ThumbList.Add(Thumbnail)
                Kill(My.Computer.FileSystem.SpecialDirectories.Temp & "\Thumbnail.jpg")
            Else
                ThumbList.Add(Nothing)
            End If
        End If
        RenameList.Add(ID)
        If exists = True And add = True Then
            VBAFlag = "True"
        End If
    End Sub
    Private Sub CreateVBA()
        If Not IO.File.Exists(My.Computer.FileSystem.SpecialDirectories.Temp & "\Get64BitPicture.bas") Then
            IO.File.WriteAllText(My.Computer.FileSystem.SpecialDirectories.Temp & "\Get64BitPicture.txt", My.Resources.Get64BitPicture)
            FileSystem.Rename(My.Computer.FileSystem.SpecialDirectories.Temp & "\Get64BitPicture.txt", My.Computer.FileSystem.SpecialDirectories.Temp & "\Get64BitPicture.bas")
        End If
        writeDebug("VBA Code missing. Generated VBA code in temp directory")
        MsgBox("The thumbnail script is missing from Inventor" & vbNewLine & vbNewLine &
               "If you no longer wish to receive this error" & vbNewLine &
               "Open VBA Editor in Inventor and import" & vbNewLine &
               My.Computer.FileSystem.SpecialDirectories.Temp & "\Get64BitPicture.bas into the ""modules list""")
    End Sub
    Private Sub RunCompleted()
        If SubFiles.Count > 10 Then
            txtSearch.Location = New Drawing.Point(LVSubFiles.Location.X, Me.Height - 122)
            txtSearch.Visible = True
            txtSearch.Text = "Search"
            txtSearch.ForeColor = Drawing.Color.Gray
        Else
            LVSubFiles.Height = lstOpenfiles.Height
            txtSearch.Visible = False
        End If

        LVSubFiles.Enabled = True
        'if no drawings are found, notify the user
        If SubFiles.Count = 0 Then
            'change display style to simulate an unusable entity
            LVSubFiles.Items.Add("No Drawings Found")
            'LVSubFiles.ThreeDCheckBoxes = False
            LVSubFiles.ForeColor = Drawing.Color.Gray

        ElseIf SubFiles.Count <> 0 And LVSubFiles.Items.Count = 0 And CMSHeirarchical.Checked = True Then
            For Each pair As KeyValuePair(Of String, String) In SubFiles
                If InStr(pair.Key, "(DNE)") <> 0 Then
                    LVSubFiles.Items.Add(Strings.Replace(pair.Key, "(DNE)", ""))
                    LVSubFiles.Items(LVSubFiles.Items.Count - 1).ForeColor = Drawing.Color.Gray
                ElseIf InStr(pair.Key, "(REF)") <> 0 Then
                    LVSubFiles.Items.Add(Strings.Replace(pair.Key, "(REF)", ""))
                    LVSubFiles.Items(LVSubFiles.Items.Count - 1).ForeColor = Drawing.Color.Blue
                    LVSubFiles.Items(LVSubFiles.Items.Count - 1).Checked = True
                ElseIf InStr(pair.Key, "(PPM)") <> 0 Then
                    LVSubFiles.Items.Add(Strings.Replace(pair.Key, "(PPM)", ""))
                    LVSubFiles.Items(LVSubFiles.Items.Count - 1).ForeColor = Drawing.Color.Red
                Else
                    LVSubFiles.Items.Add(pair.Key)
                    LVSubFiles.Items(LVSubFiles.Items.Count - 1).Checked = True
                End If
            Next
        ElseIf SubFiles.Count <> 0 And LVSubFiles.Items.Count = 0 And CMSAlphabetical.Checked = True Then
            For Each pair As KeyValuePair(Of String, String) In AlphaSub
                If InStr(pair.Key, "(DNE)") <> 0 Then
                    LVSubFiles.Items.Add(Strings.Replace(pair.Key, "(DNE)", ""))
                    LVSubFiles.Items(LVSubFiles.Items.Count - 1).ForeColor = Drawing.Color.Gray
                ElseIf InStr(pair.Key, "(REF)") <> 0 Then
                    LVSubFiles.Items.Add(Strings.Replace(pair.Key, "(REF)", ""))
                    LVSubFiles.Items(LVSubFiles.Items.Count - 1).ForeColor = Drawing.Color.Blue
                    LVSubFiles.Items(LVSubFiles.Items.Count - 1).Checked = True
                ElseIf InStr(pair.Key, "(PPM)") <> 0 Then
                    LVSubFiles.Items.Add(Strings.Replace(pair.Key, "(PPM)", ""))
                    LVSubFiles.Items(LVSubFiles.Items.Count - 1).ForeColor = Drawing.Color.Red
                Else
                    LVSubFiles.Items.Add(pair.Key)
                    LVSubFiles.Items(LVSubFiles.Items.Count - 1).Checked = True
                End If
            Next
        End If
        writeDebug("Finished updating OpenFiles list")
        LVSubFiles.AutoResizeColumns(ColumnHeaderAutoResizeStyle.ColumnContent)
        MsVistaProgressBar.Visible = False
        Me.Update()

    End Sub
    Public Sub CreateOpenDocs(OpenDocs As ArrayList)
        OpenDocs.Clear()
        Dim Archive, DocSource, DocName As String
        For Each oDoc In _invApp.Documents.VisibleDocuments
            Archive = oDoc.FullFileName
            'Use the Partsource file to create the drawingsource file
            DocSource = Strings.Left(Archive, Strings.Len(Archive))
            DocName = Strings.Right(DocSource, Strings.Len(DocSource) - Strings.InStrRev(DocSource, "\"))
            OpenDocs.Add(DocName)
            writeDebug("Created Open Document List :" & OpenDocs.Count & " documents")
        Next
    End Sub
    Public Sub MatchDrawing(ByRef DrawSource As String, ByRef DrawingName As String, Y As Integer)
        If CMSHeirarchical.Checked = True Then
            For Each item In SubFiles
                If Strings.Replace(item.Key, "(REF)", "") = LVSubFiles.Items(Y).Text Then
                    DrawSource = Strings.Left(item.Value, Len(item.Value) - 3) & "idw"
                    DrawingName = Trim(Strings.Replace(item.Key, "(REF)", ""))
                    Exit For
                End If
            Next
        Else
            For Each item In AlphaSub
                If item.Key = LVSubFiles.Items(Y).Text Then
                    DrawSource = Strings.Left(item.Value, Len(item.Value) - 3) & "idw"
                    DrawingName = Trim(item.Key)
                    Exit For
                End If
            Next
        End If
        writeDebug("Matched Drawing: " & DrawingName & " to " & DrawSource)
    End Sub
    Public Sub MatchPart(ByRef DrawSource As String, ByRef DrawingName As String, Y As Integer)

        If CMSHeirarchical.Checked = True Then

            'If SubFiles.Count > LVSubFiles.Items.Count Then
            For Each item In SubFiles
                If item.Key.Contains(LVSubFiles.Items(Y).Text) Then
                    DrawingName = item.Key.ToString
                    DrawSource = item.Value.ToString
                    If System.IO.File.Exists(Strings.Left(DrawSource, Len(DrawSource) - 3) & "ipt") And
                            System.IO.File.Exists(Strings.Left(DrawSource, Len(DrawSource) - 3) & "iam") Then
                        If My.Settings.DupName = "Model" Then
                            DrawSource = Strings.Left(DrawSource, Len(DrawSource) - 3) & "ipt"
                            DrawingName = Strings.Left(DrawingName, Len(DrawingName) - 3) & "ipt"
                        Else
                            DrawSource = Strings.Left(DrawSource, Len(DrawSource) - 3) & "iam"
                            DrawingName = Strings.Left(DrawingName, Len(DrawingName) - 3) & "iam"
                        End If
                        Exit For
                    ElseIf System.IO.File.Exists(Strings.Left(DrawSource, Len(DrawSource) - 3) & "ipt") And
                            Not System.IO.File.Exists(Strings.Left(DrawSource, Len(DrawSource) - 3) & "iam") Then
                        DrawSource = Strings.Left(DrawSource, Len(DrawSource) - 3) & "ipt"
                        DrawingName = Strings.Left(DrawingName, Len(DrawingName) - 3) & "ipt"
                        Exit For
                    ElseIf System.IO.File.Exists(Strings.Left(DrawSource, Len(DrawSource) - 3) & "iam") And
                            Not System.IO.File.Exists(Strings.Left(DrawSource, Len(DrawSource) - 3) & "ipt") Then
                        DrawSource = Strings.Left(DrawSource, Len(DrawSource) - 3) & "iam"
                        DrawingName = Strings.Left(DrawingName, Len(DrawingName) - 3) & "iam"
                        Exit For
                    Else
                        MsgBox("There is a discrepancy between the drawing name and model name." & vbNewLine &
                               "Model data could not be located for drawing: " & DrawingName)
                        Exit Sub
                    End If
                End If
            Next
        Else
            For Each item In AlphaSub
                If item.Key.Contains(LVSubFiles.Items(Y).Text) Then
                    DrawingName = item.Key(Y).ToString
                    DrawSource = item.Value(Y).ToString
                    If System.IO.File.Exists(Strings.Left(DrawSource, Len(DrawSource) - 3) & "ipt") And
                            System.IO.File.Exists(Strings.Left(DrawSource, Len(DrawSource) - 3) & "iam") Then
                        If My.Settings.DupName = "Model" Then
                            DrawSource = Strings.Left(DrawSource, Len(DrawSource) - 3) & "ipt"
                            DrawingName = Strings.Left(DrawingName, Len(DrawingName) - 3) & "ipt"
                        Else
                            DrawSource = Strings.Left(DrawSource, Len(DrawSource) - 3) & "iam"
                            DrawingName = Strings.Left(DrawingName, Len(DrawingName) - 3) & "iam"
                        End If
                        Exit For
                    ElseIf System.IO.File.Exists(Strings.Left(DrawSource, Len(DrawSource) - 3) & "ipt") And
                            Not System.IO.File.Exists(Strings.Left(DrawSource, Len(DrawSource) - 3) & "iam") Then
                        DrawSource = Strings.Left(DrawSource, Len(DrawSource) - 3) & "ipt"
                        DrawingName = Strings.Left(DrawingName, Len(DrawingName) - 3) & "ipt"
                        Exit For
                    ElseIf System.IO.File.Exists(Strings.Left(DrawSource, Len(DrawSource) - 3) & "iam") And
                            Not System.IO.File.Exists(Strings.Left(DrawSource, Len(DrawSource) - 3) & "ipt") Then
                        DrawSource = Strings.Left(DrawSource, Len(DrawSource) - 3) & "iam"
                        DrawingName = Strings.Left(DrawingName, Len(DrawingName) - 3) & "iam"
                        Exit For
                    Else
                        MsgBox("Couldn't locate model data for " & DrawingName)
                        Exit Sub
                    End If
                End If
            Next
        End If
        writeDebug("Matched Part: " & DrawingName & " to " & DrawSource)
        'End If
        'Look for selected item
        'Dim CheckedFile As String = Trim(lstSubfiles.Items.Item(X))
        '        For J = 1 To _invApp.Documents.Count
        '            oDoc = Path.Item(J)
        '            Archive = oDoc.FullFileName
        '            If Archive = Nothing Then GoTo skip
        '            If Strings.Right(Archive, 1) <> ">" Then
        '                If Strings.Right(Archive, 3) <> "idw" Then
        '                    Dim mass As Double = oDoc.ComponentDefinition.MassProperties.Mass
        '                End If
        '                'Use the Partsource file to create the drawingsource file
        '                If System.IO.File.Exists(Strings.Left(Archive, Strings.Len(Archive) - 3) & "ipt") Then
        '                    PartSource = Strings.Left(Archive, Strings.Len(Archive) - 3) & "ipt"
        '                ElseIf System.IO.File.Exists(Strings.Left(Archive, Strings.Len(Archive) - 3) & "iam") Then
        '                    PartSource = Strings.Left(Archive, Strings.Len(Archive) - 3) & "iam"
        '                End If
        '                If PartSource = "" Then
        '                    MsgBox("The drawing " & CheckedFile & " has a model with a different name" & vbNewLine _
        '                                               & "or has been saved in a different project location." & vbNewLine _
        '                                               & "This model's properties could not be located.")
        '                    GoTo skip
        '                End If

        '                PartName = Strings.Right(PartSource, Strings.Len(PartSource) - Strings.InStrRev(PartSource, "\"))
        '                'If the drawing file is checked, open the drawing in Inventor
        '                If Strings.Left(CheckedFile, Len(CheckedFile) - 4) = Strings.Left(PartName, Len(PartName) - 4) Then
        '                    Archive = PartSource
        '                    Exit Sub
        '                End If
        '            End If
        'skip:
        '        Next
    End Sub
    Private Sub Search_For_Duplicates(ByVal Filename As String, DrawingName As String, ByVal Filetype As String)
        For Each file As IO.FileInfo In Get_Files(Strings.Left(Filename, InStrRev(Filename, "\")),
                                                                  IO.SearchOption.TopDirectoryOnly, Filetype,
                                                                  "\" & Strings.Left(DrawingName, InStrRev(DrawingName, ".")))
            If file.FullName <> Filename Then
                If My.Computer.FileSystem.FileExists(file.FullName.Insert(file.FullName.LastIndexOf("\"), "\Archived")) Then
                    Kill(file.FullName)
                Else
                    System.IO.File.Move(file.FullName, file.FullName.Insert(file.FullName.LastIndexOf("\"), "\Archived"))
                End If
            End If
        Next
        For Each file As IO.FileInfo In Get_Files(Strings.Left(Filename, InStrRev(Filename, "\")),
                                                                  IO.SearchOption.TopDirectoryOnly, Filetype,
                                                                  "\" & Strings.Left(DrawingName, InStrRev(DrawingName, ".") - 1) & "-R")
            If file.FullName <> Filename Then
                If My.Computer.FileSystem.FileExists(file.FullName.Insert(file.FullName.LastIndexOf("\"), "\Archived")) Then
                    Kill(file.FullName)
                Else
                    System.IO.File.Move(file.FullName, file.FullName.Insert(file.FullName.LastIndexOf("\"), "\Archived"))
                End If
            End If
        Next
    End Sub
    Private Function Get_Files(ByVal directory As String,
                           ByVal recursive As IO.SearchOption,
                           ByVal ext As String,
                           ByVal filename As String) As List(Of IO.FileInfo)

        'Check the directory for older revisions that have the same base name and select them to be moved to the archive folder
        Return IO.Directory.GetFiles(directory, "*" & If(ext.StartsWith("*"), ext.Substring(1), ext), recursive) _
                           .Where(Function(o) o.ToLower.Contains(filename.ToLower)) _
                           .Select(Function(p) New IO.FileInfo(p)).ToList


    End Function
    Private Sub SheetMetalTest(ByRef Archive As String, oDoc As Document, ByRef sReadableType As String)
        Dim sDocumentSubType As String = oDoc.SubType
        'Get document type
        Dim eDocumentType As DocumentTypeEnum = oDoc.DocumentType
        'Check document type for sheet metal
        If chkSkipAssy.CheckState = CheckState.Checked Then
            If sDocumentSubType = "{28EC8354-9024-440F-A8A2-0E0E55D635B0}" Or sDocumentSubType = "{E60F81E1-49B3-11D0-93C3-7E0706000000}" Then
                sReadableType = ""
                Exit Sub
            End If
        ElseIf chkSkipAssy.CheckState = CheckState.Unchecked And sDocumentSubType = "{E60F81E1-49B3-11D0-93C3-7E0706000000}" Then
            sReadableType = "P"
            writeDebug("Sheetmetal test for " & oDoc.DisplayName & ": Part")
        End If
        If chkUseDrawings.CheckState = CheckState.Checked And chkSkipAssy.CheckState = CheckState.Unchecked Then
            sReadableType = "P"
            writeDebug("Sheetmetal test for " & oDoc.DisplayName & ": Part")
            Exit Sub
        End If
        Select Case sDocumentSubType
            Case "{4D29B490-49B2-11D0-93C3-7E0706000000}"
                'Part
                sReadableType = "P"
                writeDebug("Sheetmetal test for " & oDoc.DisplayName & ": Part")
            Case "{28EC8354-9024-440F-A8A2-0E0E55D635B0}"
                sReadableType = "P"
            Case "{9C464203-9BAE-11D3-8BAD-0060B0CE6BB4}"
                sReadableType = "S"
                writeDebug("Sheetmetal test for " & oDoc.DisplayName & ": Sheet")
                'Sheet Metal
        End Select
    End Sub
    Public Sub CloseLater(Name As String, oDoc As Document)
        Dim CloseLater As Boolean = True
        'Go through string of originally open documents to see if document has been opened by the program
        For Each Str As String In OpenDocs
            If Str.Contains(Name) Then
                'If found, don't close
                CloseLater = False
                Exit For
            End If
        Next
        'If the drawing wasn't found close it. This keeps the system from having too many open documents in memory.
        If CloseLater = True Then
            Try
                _invApp.Documents.ItemByName(oDoc.FullDocumentName).Close(True)
                writeDebug("Drawing " & Name & " closed")
            Catch
                writeDebug("Failed to close " & Name & ". Perhaps not yet saved, or already closed.")
            End Try

            'oDoc.Close(True)
        End If
    End Sub
    Public Sub ProgressBar(Total As Integer, Counter As Integer, Title As String, FileName As String)
        Dim Percent As Integer = (Counter / Total) * 100
        MsVistaProgressBar.Visible = True
        If Percent > 100 Then Percent = 100
        Me.MsVistaProgressBar.Value = Percent
        If Title = "Found: " Or Title = "Getting References: " Then
            Me.MsVistaProgressBar.DisplayText = Title & " " & FileName
        Else
            Me.MsVistaProgressBar.DisplayText = Title & " " & FileName & " " & Percent & "%"
        End If
    End Sub
    Private Sub CountParts(ByRef Total As Integer)
        ' Get the active assembly. 
        Dim oAsmDoc As AssemblyDocument
        oAsmDoc = _invApp.ActiveDocument

        ' Get the assembly component definition. 
        Dim oAsmDef As AssemblyComponentDefinition
        oAsmDef = oAsmDoc.ComponentDefinition

        ' Get all of the leaf occurrences of the assembly. 
        Dim oLeafOccs As ComponentOccurrencesEnumerator
        oLeafOccs = oAsmDef.Occurrences.AllLeafOccurrences

        ' Iterate through the occurrences and print the name. 
        Dim oOcc As ComponentOccurrence
        For Each oOcc In oLeafOccs
            Total = Total + 1
        Next
    End Sub
    Public Sub KillAllExcels(Time As System.DateTime)
        Dim proc As System.Diagnostics.Process
        For Each proc In System.Diagnostics.Process.GetProcessesByName("EXCEL")
            If proc.StartTime > Time Then proc.Kill()
            writeDebug("Killed Excel Process")
        Next
    End Sub
    Private Sub GetProperties(ByRef oDoc As Document, ByVal Occurrences As ComponentOccurrences, ByRef Properties As Double, ByRef Counter As Integer, ExcelDoc As Excel.Workbook, Total As Integer)
        Dim Occ As ComponentOccurrence
        Dim oPartDef As ComponentDefinition
        Dim PropSets, customPropSet, UserDefProps As Inventor.PropertySets
        Dim Process, StockNo, Material, PartNo, Description, Vendor, Length, Width, Thk, Area, Title As String
        Dim Offset As Integer = 1
        Title = "Extracting"
        For Each Occ In Occurrences
            Counter += 1
            'Check if the document is a part or assembly
            If Occ.DefinitionDocumentType = DocumentTypeEnum.kPartDocumentObject Or
                DocumentTypeEnum.kAssemblyDocumentObject And Occ.BOMStructure = BOMStructureEnum.kPurchasedBOMStructure Then
                'Get iProperties for this document
                PropSets = Occ.Definition.Document.PropertySets()
                Process = PropSets.Item("{D5CDD502-2E9C-101B-9397-08002B2CF9AE}").ItemByPropId("2").Value.ToString
                If Occ.BOMStructure = BOMStructureEnum.kPurchasedBOMStructure Then
                    Process = "PP"
                ElseIf Occ.BOMStructure = BOMStructureEnum.kReferenceBOMStructure Then
                    Process = "Ref"
                ElseIf Occ.BOMStructure = BOMStructureEnum.kPhantomBOMStructure Then
                    Process = "PH"
                End If
                StockNo = PropSets.Item("{32853F0F-3444-11D1-9E93-0060B03C1CA6}").ItemByPropId("55").Value.ToString
                Material = PropSets.Item("{32853F0F-3444-11D1-9E93-0060B03C1CA6}").ItemByPropId("20").Value.ToString
                PartNo = PropSets.Item("{32853F0F-3444-11D1-9E93-0060B03C1CA6}").ItemByPropId("5").Value.ToString
                Description = PropSets.Item("{32853F0F-3444-11D1-9E93-0060B03C1CA6}").ItemByPropId("29").Value.ToString
                Vendor = PropSets.Item("{32853F0F-3444-11D1-9E93-0060B03C1CA6}").ItemByPropId("30").Value.ToString

                oPartDef = Occ.Definition
                On Error Resume Next
                'Dim ExcelDoc As Excel.Workbook
                Offset += 1

                Length = oPartDef.Parameters.Item("Length").Value / 30.48
                Width = oPartDef.Parameters.Item("Width").Value / 30.48
                Thk = oPartDef.Parameters.Item("Thickness").Value / 2.54
                oDoc = _invApp.Documents.Open(Occ.Definition.Document.FullFileName, False)
                Dim sReadablefiletype As String = ""
                SheetMetalTest(oDoc.FullFileName, oDoc, sReadablefiletype)
                If sReadablefiletype = "S" Then
                    CreateFlatPattern(oDoc.ComponentDefinition, oDoc.DisplayName, oDoc, False)
                End If
                Area = AreaCalculate(oDoc, Occ)
                If Process = "SC" Or Process = "EX" Then
                    ExcelDoc.Worksheets("Saw Cut").activate()
                    Do Until ExcelDoc.ActiveSheet.Range("F" & Offset).Value = ""
                        Offset = Offset + 1
                        If CStr(ExcelDoc.ActiveSheet.Range("F" & Offset).Value) = CStr(StockNo) And
                            CStr(ExcelDoc.ActiveSheet.Range("E" & Offset).Value) = CStr(Material) Then
                            ExcelDoc.ActiveSheet.Range("A" & Offset).Value = ExcelDoc.ActiveSheet.Range("A" & Offset).Value + (Length)
                            If InStrRev(ExcelDoc.ActiveSheet.Range("G" & Offset).Value, PartNo) = "" And
                                ExcelDoc.ActiveSheet.Range("G" & Offset).Value <> "" And
                                InStr(ExcelDoc.ActiveSheet.range("G" & Offset).value, PartNo) = 0 Then
                                ExcelDoc.ActiveSheet.Range("G" & Offset).Value = ExcelDoc.ActiveSheet.Range("G" & Offset).Value & ", " & PartNo
                            End If
                            Exit Do
                        ElseIf ExcelDoc.ActiveSheet.Range("F" & Offset).Value = Nothing Then
                            ExcelDoc.ActiveSheet.Range("A" & Offset).Value = ExcelDoc.ActiveSheet.Range("A" & Offset).Value + (Length)
                            ExcelDoc.ActiveSheet.Range("B" & Offset).Value = "Ft."
                            ExcelDoc.ActiveSheet.Range("F" & Offset).Value = StockNo
                            ExcelDoc.ActiveSheet.Range("E" & Offset).Value = Material
                            ExcelDoc.ActiveSheet.Range("G" & Offset).Value = PartNo
                            Exit Do
                        Else
                            'MsgBox "Jump to next line"
                        End If
                    Loop
                    Offset = 1
                    ExcelDoc.Worksheets("Saw Cut Lengths").Activate()
                    Do Until ExcelDoc.ActiveSheet.Range("C" & Offset).Value = ""
                        Offset = Offset + 1

                        If ExcelDoc.ActiveSheet.Range("C" & Offset).Value = Nothing Then
                            ExcelDoc.ActiveSheet.Range("A" & Offset).Value = Length
                            ExcelDoc.ActiveSheet.Range("B" & Offset).Value = "Ft."
                            ExcelDoc.ActiveSheet.Range("D" & Offset).Value = StockNo
                            ExcelDoc.ActiveSheet.Range("C" & Offset).Value = Material
                            ExcelDoc.ActiveSheet.Range("E" & Offset).Value = PartNo
                            Exit Do
                        End If
                    Loop
                ElseIf Process = "LC" Then
                    ExcelDoc.Worksheets("Profile Cut").Activate()
                    Do Until ExcelDoc.ActiveSheet.Range("D" & Offset).Value = ""
                        Offset = Offset + 1
                        If CStr(ExcelDoc.ActiveSheet.Range("D" & Offset).Value) = CStr(StockNo) Then
                            ExcelDoc.ActiveSheet.Range("A" & Offset).Value = ExcelDoc.ActiveSheet.Range("A" & Offset).Value + Area
                            If InStrRev(ExcelDoc.ActiveSheet.Range("E" & Offset).Value, PartNo) = "" And
                                ExcelDoc.ActiveSheet.Range("E" & Offset).Value <> "" Then
                                ExcelDoc.ActiveSheet.Range("E" & Offset).Value = ExcelDoc.ActiveSheet.Range("E" & Offset).Value & ", " & PartNo
                            End If
                            Exit Do
                        ElseIf ExcelDoc.ActiveSheet.Range("C" & Offset).Value = Nothing Then

                            ExcelDoc.ActiveSheet.Range("A" & Offset).Value = ExcelDoc.ActiveSheet.Range("A" & Offset).Value + Area
                            ExcelDoc.ActiveSheet.Range("B" & Offset).Value = "Sq Ft."
                            ExcelDoc.ActiveSheet.Range("D" & Offset).Value = StockNo
                            ExcelDoc.ActiveSheet.Range("C" & Offset).Value = Material
                            ExcelDoc.ActiveSheet.Range("E" & Offset).Value = PartNo
                            Exit Do
                        Else
                            'MsgBox "Jump to next line"
                        End If
                    Loop
                ElseIf Process = "HC" Then
                    ExcelDoc.Worksheets("Torch Cut").Activate()
                    Do Until ExcelDoc.ActiveSheet.Range("D" & Offset).Value = ""
                        Offset = Offset + 1
                        If CStr(ExcelDoc.ActiveSheet.Range("D" & Offset).Value) = CStr(StockNo) Then
                            ExcelDoc.ActiveSheet.Range("A" & Offset).Value = ExcelDoc.ActiveSheet.Range("A" & Offset).Value + (Length * Width)
                            If InStrRev(ExcelDoc.ActiveSheet.Range("E" & Offset).Value, PartNo) = "" Then
                                ExcelDoc.ActiveSheet.Range("E" & Offset).Value = ExcelDoc.ActiveSheet.Range("E" & Offset).Value & ", " & PartNo
                            End If
                            Exit Do
                        ElseIf ExcelDoc.ActiveSheet.Range("D" & Offset).Value = Nothing Then
                            ExcelDoc.ActiveSheet.Range("A" & Offset).Value = ExcelDoc.ActiveSheet.Range("A" & Offset).Value + (Length * Width)
                            ExcelDoc.ActiveSheet.Range("B" & Offset).Value = "Sq Ft."
                            ExcelDoc.ActiveSheet.Range("D" & Offset).Value = StockNo
                            ExcelDoc.ActiveSheet.Range("C" & Offset).Value = Material
                            ExcelDoc.ActiveSheet.Range("E" & Offset).Value = PartNo
                            Exit Do
                        Else
                            'MsgBox "Jump to next line"
                        End If
                    Loop
                ElseIf Process = "PP" Then
                    ExcelDoc.Worksheets("Purchased Parts").Activate()
                    Do Until ExcelDoc.ActiveSheet.Range("B" & Offset).Value = ""
                        Offset = Offset + 1
                        If CStr(ExcelDoc.ActiveSheet.Range("B" & Offset).Value) = CStr(Description) And
                            CStr(ExcelDoc.ActiveSheet.Range("D" & Offset).Value) = CStr(PartNo) Then
                            ExcelDoc.ActiveSheet.Range("A" & Offset).Value = ExcelDoc.ActiveSheet.Range("A" & Offset).Value + 1
                            Exit Do
                        ElseIf ExcelDoc.ActiveSheet.Range("B" & Offset).Value = Nothing Then
                            ExcelDoc.ActiveSheet.Range("A" & Offset).Value = 1
                            ExcelDoc.ActiveSheet.Range("D" & Offset).Value = PartNo
                            ExcelDoc.ActiveSheet.Range("B" & Offset).Value = Description
                            ExcelDoc.ActiveSheet.Range("C" & Offset).Value = Vendor
                            Exit Do
                        Else
                            'MsgBox "Jump to next line"
                        End If
                    Loop
                ElseIf Process = "" And Occ.DefinitionDocumentType = DocumentTypeEnum.kPartDocumentObject Then
                    ExcelDoc.Worksheets("Other Parts").Activate()
                    Do Until ExcelDoc.ActiveSheet.Range("B" & Offset).Value = ""
                        Offset = Offset + 1
                        If CStr(ExcelDoc.ActiveSheet.Range("E" & Offset).Value) = CStr(Description) Then
                            ExcelDoc.ActiveSheet.Range("A" & Offset).Value = ExcelDoc.ActiveSheet.Range("A" & Offset).Value + 1
                            ExcelDoc.ActiveSheet.Range("B" & Offset).Value = ExcelDoc.ActiveSheet.Range("B" & Offset).Value + (Length)
                            'ExcelDoc.ActiveSheet.Range("D" & Offset).Value = ExcelDoc.ActiveSheet.Range("D" & Offset).Value + (Length * Width)
                            If InStrRev(ExcelDoc.ActiveSheet.Range("G" & Offset).Value, PartNo) = "" And
                                ExcelDoc.ActiveSheet.Range("G" & Offset).Value <> "" Then
                                ExcelDoc.ActiveSheet.Range("G" & Offset).Value = ExcelDoc.ActiveSheet.Range("G" & Offset).Value & ", " & PartNo
                            End If
                            Exit Do
                        ElseIf ExcelDoc.ActiveSheet.Range("B" & Offset).Value = Nothing Then
                            ExcelDoc.ActiveSheet.Range("A" & Offset).Value = 1
                            'ExcelDoc.ActiveSheet.Range("B" & Offset).Value = ExcelDoc.ActiveSheet.Range("B" & Offset).Value + (Length)
                            ExcelDoc.ActiveSheet.Range("B" & Offset).Value = "Ea."
                            'ExcelDoc.ActiveSheet.Range("D" & Offset).Value = ExcelDoc.ActiveSheet.Range("D" & Offset).Value + (Length * Width)
                            'ExcelDoc.ActiveSheet.Range("E" & Offset).Value = "Sq Ft."
                            ExcelDoc.ActiveSheet.Range("E" & Offset).Value = Description
                            ExcelDoc.ActiveSheet.Range("C" & Offset).Value = Material
                            ExcelDoc.ActiveSheet.Range("D" & Offset).Value = StockNo
                            ExcelDoc.ActiveSheet.Range("G" & Offset).Value = PartNo
                            ExcelDoc.ActiveSheet.Range("F" & Offset).Value = Vendor
                            Exit Do
                        Else
                            'MsgBox "Jump to next line"
                        End If
                    Loop
                End If
                Offset = 1
                ProgressBar(Total, Counter, Title, Occ.Name)
                If Err.Number = 0 Then
                    Properties = Properties + Length
                End If
                On Error GoTo 0
            Else
                GetProperties(oDoc, Occ.SubOccurrences, Properties, Counter, ExcelDoc, Total)
            End If
        Next
        ExcelDoc.Worksheets("Saw Cut").activate()
    End Sub
    Function AreaCalculate(ByRef oDoc As Document, ByRef Occ As ComponentOccurrence) As Decimal

        Dim oDef As SheetMetalComponentDefinition
        oDef = oDoc.ComponentDefinition
        Dim oFlatPattern As FlatPattern
        oFlatPattern = oDef.FlatPattern
        Dim oTransaction As Transaction
        oTransaction = _invApp.TransactionManager.StartTransaction(oDoc, "Find area")
        Dim oSketch As PlanarSketch
        oSketch = oFlatPattern.Sketches.Add(oFlatPattern.TopFace)
        Dim oEdgeLoop As EdgeLoop = Nothing
        For Each oEdgeLoop In oFlatPattern.TopFace.EdgeLoops
            If oEdgeLoop.IsOuterEdgeLoop Then
                Exit For
            End If
        Next
        Dim oEdge As Edge
        For Each oEdge In oEdgeLoop.Edges
            Call oSketch.AddByProjectingEntity(oEdge)
        Next
        Dim oProfile As Profile
        oProfile = oSketch.Profiles.AddForSolid
        Dim dArea As Double
        dArea = oProfile.RegionProperties.Area
        oTransaction.Abort()
        AreaCalculate = dArea / (12 * 2.54) ^ 2
    End Function
    Public Sub ExtractThumb(ByRef PartName As String, ByRef Thumbnail As Image)
        Dim X As Integer = 0
        For Each entry As List(Of String) In RenameTable
            If entry.Item(2).ToString = PartName Then
                Thumbnail = ThumbList(X)
                Exit For
            End If
            X += 1
        Next
    End Sub
    Private Sub CloseDrawings(Path As Documents, ByRef odoc As Document, ByRef Archive As String _
                                 , ByRef DrawingName As String, ByRef DrawSource As String, OpenDocs As ArrayList)
        Dim Ans As String
        Dim Checkin As Boolean
        Ans = MsgBox("Do you wish to save these drawings?", vbYesNoCancel, "Close Drawings")
        If Ans = vbYes Then
            Checkin = True
        ElseIf Ans = vbNo Then
            Checkin = False
        Else
            Exit Sub
        End If
        'Go through drawings to see which ones are selected

        For x = 0 To LVSubFiles.CheckedItems.Count - 1
            'iterate through all open documents in inventor to find selected document

            'For Each item As Document In _invApp.Documents
            ' If Strings.Right(item.FullFileName, Len(item.FullFileName) - InStrRev(item.FullFileName, "\")) = Trim(LVSubFiles.CheckedItems(x).Text) Then
            'save document if user selected save
            MatchDrawing(DrawSource, DrawingName, LVSubFiles.SelectedItems(x).Index)
            If Checkin = True Then
                Try
                    _invApp.Documents.ItemByName(DrawSource).Save()
                    _invApp.Documents.ItemByName(DrawSource).Close(True)
                    ' otherwise close without save
                Catch
                End Try
            Else
                _invApp.Documents.ItemByName(DrawSource).Close(True)
                writeDebug("Drawing " & DrawingName & " closed without save")
            End If
        Next
        'UpdateForm()
    End Sub
    Private Sub ToolTip1_Popup(sender As System.Object, e As System.Windows.Forms.PopupEventArgs) Handles ToolTip1.Popup
        Dim ctrl As Control = GetChildAtPoint(Me.PointToClient(MousePosition))
        Dim tt As String = ToolTip1.GetToolTip(ctrl)
        If Not tt = Nothing Then
            If tt.Length > 75 Then ToolTip1.SetToolTip(ctrl, SplitToolTip(tt))
        End If
    End Sub
    Friend Function SplitToolTip(ByVal strOrig As String) As String
        Dim strArray As String()
        Dim SPACE As String = " "
        Dim strOneWord As String
        Dim strBuilder As String = ""
        Dim strReturn As String = ""
        strArray = strOrig.Split(SPACE)
        For Each strOneWord In strArray
            strBuilder = strOneWord & SPACE
            If Len(strBuilder) > 70 Then
                strReturn = strReturn & strBuilder & vbNewLine &
                strBuilder = ""
            End If
        Next
        Return strReturn & strBuilder
    End Function
    Private Sub Main_FormClosed(sender As Object, e As FormClosedEventArgs) Handles Me.FormClosed
        writeDebug("Form close clean up started")
        Try
            _invApp = Marshal.GetActiveObject("Inventor.Application")
        Catch ex As Exception
            Exit Sub
        End Try
        If _invApp.Documents.VisibleDocuments.Count = 0 Then
            _invApp.Documents.CloseAll()
        ElseIf _invApp.Documents.Count > _invApp.Documents.VisibleDocuments.Count And OpenDocs.Count <> 0 Then
            For Each odoc As Document In _invApp.Documents
                Try
                    CloseLater(odoc.DisplayName, odoc)
                Catch
                    Try
                        writeDebug("Error closing " & odoc.DisplayName)
                    Catch
                        writeDebug("Error closing unknown document")
                    End Try
                End Try
            Next
        End If
        If _started = True Then
            Dim ans As Boolean = MsgBox("Do you wish to close Inventor as well?", vbYesNo, "Application Closed")
            If ans = vbYes Then
                _invApp.Quit()
                writeDebug("Inventor Closed")
            End If
        End If
        writeDebug("Main Form closed")
    End Sub
    Private Sub Main_MouseEnter(sender As Object, e As EventArgs) Handles Me.MouseEnter
        If InventorExited = True Then
            writeDebug("Connection to Inventor lost")
            If MsgBox("Connection to Inventor was lost" & vbNewLine & "Do you wish to try to restart Inventor?", vbYesNo, "Unexpected Error") = vbYes Then
                Try
                    Dim Inv As New ProcessStartInfo("Inventor.exe")
                    Dim p As New Process
                    p.StartInfo = Inv
                    p.Start()
                    'p.WaitForInputIdle()
                    While p.MainWindowTitle = Nothing
                        System.Threading.Thread.Sleep(1000)
                        p.Refresh()
                    End While
                    _invApp = Marshal.GetActiveObject("Inventor.Application")
                    _started = True
                    ProcessStarted()
                Catch ex2 As Exception
                    MsgBox(ex2.ToString())
                    MsgBox("Unable to get or start Inventor")
                    writeDebug("Unable to start Inventor")
                End Try
            End If
            Me.Close()
        End If
        InventorExited = False
    End Sub
    Private Sub Rename_Resize(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Resize
        GroupBox4.Location = New Drawing.Point(Me.Width - 203, 27)
        btnOK.Location = New Drawing.Point(Me.Width - 115, Me.Height - 70)
        btnExit.Location = New Drawing.Point(Me.Width - 196, btnOK.Location.Y)
        GroupBox2.Height = (Me.Height - 96)
        GroupBox2.Width = (Me.Width - 375) / 2
        GroupBox3.Height = GroupBox2.Height
        GroupBox3.Width = GroupBox2.Width
        GroupBox3.Location = New Drawing.Point(GroupBox2.Location.X + GroupBox2.Width + 6, GroupBox2.Location.Y)
        lstOpenfiles.Height = GroupBox2.Height - 24
        lstOpenfiles.Width = GroupBox2.Width - 22
        LVSubFiles.Width = GroupBox3.Width - 22
        LVSubFiles.AutoResizeColumns(ColumnHeaderAutoResizeStyle.ColumnContent)
        MsVistaProgressBar.Location = New Drawing.Point(12, Me.Height - 67)
        MsVistaProgressBar.Width = Me.Width - 226
        GroupBox5.Location = New Drawing.Point(12, Me.Height - 187)
        If LVSubFiles.Items.Count < LVSubFiles.Height / 18 And txtSearch.ForeColor = Drawing.Color.Gray Or LVSubFiles.Items.Count = 0 Then
            txtSearch.Visible = False
            LVSubFiles.Height = GroupBox2.Height - 24
        Else
            txtSearch.Visible = True
            LVSubFiles.Height = GroupBox2.Height - 49
            txtSearch.Width = LVSubFiles.Width
            txtSearch.Location = New Drawing.Point(LVSubFiles.Location.X, Me.Height - 122)

        End If
        PictureBox2.Location = New Drawing.Point(GroupBox3.Location.X + 21, 27)
    End Sub
    Private Sub btnFlatPattern_Click(sender As Object, e As EventArgs) Handles btnFlatPattern.Click
        Dim Archive As String = ""
        Dim PartName As String = ""
        Dim PartSource As String = ""
        Dim oDoc As Document = Nothing
        Dim sReadableType As String = ""
        Dim Path As Documents = _invApp.Documents
        Dim RevNo As String = ""
        Dim DXFSource As String = ""
        Dim Title As String = "Saving: "
        Dim Total As String = lstOpenfiles.CheckedItems.Count
        Dim Flag As Boolean = False
        For X = 0 To lstOpenfiles.Items.Count - 1
            If InStr(lstOpenfiles.Items(X), "ipt") <> 0 Then
                Flag = True
                Exit For

            End If
        Next
        If Flag = False Then
            MsgBox("No part selected." & vbNewLine & "For drawings use the drawing action list.")
        Else
            For X = 0 To lstOpenfiles.Items.Count - 1
                'Look through all sub files in open documents to get the part sourcefile
                If lstOpenfiles.GetItemCheckState(X) = CheckState.Checked Then
                    PartName = lstOpenfiles.Items(X)
                    PartSource = OpenFiles.Item(X).Value
                    oDoc = _invApp.Documents.Open(PartSource, False)
                    If My.Computer.FileSystem.FileExists(Strings.Replace(PartSource, "ipt", "idw")) Then
                        Dim dDoc As DrawingDocument = _invApp.Documents.Open(Strings.Replace(PartSource, "ipt", "idw"), False)
                        RevNo = dDoc.PropertySets.Item("{F29F85E0-4FF9-1068-AB91-08002B27B3D9}").ItemByPropId("9").Value
                        dDoc.Close()
                    Else
                        RevNo = 0
                    End If

                    DXFSource = PartSource.Insert(PartSource.LastIndexOf("\"), "\DWG_DXF")
                    If My.Settings.DXFRev = True Then
                        DXFSource = DXFSource.Insert(DXFSource.LastIndexOf("."), "-R" & RevNo)
                    End If
                    DXFSource = Replace(DXFSource, "ipt", "dxf")
                    SheetMetalTest(Archive, oDoc, sReadableType)
                    If sReadableType = "S" Then
                        Call SMDXF(oDoc, DXFSource, True, Replace(PartName, ".idw", ".dxf"))
                    End If
                    ProgressBar(Total, X + 1, "Saving: ", PartName)
                End If
            Next
        End If
        MsVistaProgressBar.Visible = False
    End Sub
    Private Sub ExportPart(DrawSource As String, Archive As String, FlatPattern As Boolean, Destin As String, DrawingName As String,
             OpenDocs As ArrayList, Output As String, RevNo As String)

        Dim odoc As Document = _invApp.ActiveDocument
        If FlatPattern = False Then
            odoc = _invApp.Documents.Open(DrawSource, True)
        Else
            odoc = _invApp.Documents.Open(Archive, True)
        End If
        Dim oDWGAddIn As TranslatorAddIn = Nothing
        Dim i As Long
        Dim strIniFile As String = ""
        Dim DwgName, DWGSource, ExportName As String
        Archive = _invApp.ActiveDocument.FullFileName
        DwgName = Strings.Right(Archive, Len(Archive) - InStrRev(Archive, "\"))
        DWGSource = Destin & "\" & DrawingName
        DWGSource = DWGSource.Insert(DWGSource.LastIndexOf("."), RevNo)
        ExportName = Replace(DWGSource, "idw", LCase(Output))

        'If My.Computer.FileSystem.DirectoryExists(Strings.Left(ExportName, InStrRev(ExportName, "\"))) = False Then
        '    MkDir(Strings.Left(ExportName, InStrRev(ExportName, "\")))
        'End If
        'If My.Computer.FileSystem.DirectoryExists(Strings.Left(ExportName, InStrRev(ExportName, "\")) & "Archived\") = False Then
        '    MkDir(Strings.Left(ExportName, InStrRev(ExportName, "\")) & "Archived\")
        'End If
        'Dim Filetype As String = "." & LCase(Output)
        'Search_For_Duplicates(ExportName, DrawingName, Filetype)
        Dim oNameValueMap As NameValueMap
        oNameValueMap = _invApp.TransientObjects.CreateNameValueMap
        ' Define the type of output by
        ' specifying the filename.
        Dim oOutputFile As DataMedium
        oOutputFile = _invApp.TransientObjects.CreateDataMedium

        For i = 1 To _invApp.ApplicationAddIns.Count
            If Output = "dxf" Then
                If _invApp.ApplicationAddIns.Item(i).ClassIdString = "{C24E3AC4-122E-11D5-8E91-0010B541CD80}" Then
                    oDWGAddIn = _invApp.ApplicationAddIns.Item(i)
                    strIniFile = IO.Path.GetDirectoryName(My.Application.Info.DirectoryPath & "\Resources\dxf.ini")
                    Exit For
                End If
            ElseIf Output = "dwg" Then
                If _invApp.ApplicationAddIns.Item(i).ClassIdString = "{C24E3AC2-122E-11D5-8E91-0010B541CD80}" Then
                    oDWGAddIn = _invApp.ApplicationAddIns.Item(i)
                    strIniFile = IO.Path.GetDirectoryName(My.Application.Info.DirectoryPath & "\Resources\dwg.ini")
                    Exit For
                End If
            Else
                MsgBox("Error: No export type selected")
                Exit Sub
            End If
        Next
        Call oNameValueMap.Add("Export_Acad_IniFile", strIniFile)
        oOutputFile.FileName = ExportName

        If oDWGAddIn Is Nothing Then
            MsgBox("DWG add-in not found.")
            Exit Sub
        End If

        ' Check to make sure the add-in
        ' is activated.
        If Not oDWGAddIn.Activated Then
            oDWGAddIn.Activate()
        End If

        ' Create a translation context and define
        ' that we want to output to a file.
        Dim oContext As TranslationContext
        oContext = _invApp.TransientObjects.CreateTranslationContext
        oContext.Type = IOMechanismEnum.kFileBrowseIOMechanism
        ' Call the SaveCopyAs method of the add-in.
        Try
            Call oDWGAddIn.SaveCopyAs(_invApp.ActiveDocument, oContext, oNameValueMap, oOutputFile)
            writeDebug("Created " & Output & " " & ExportName)
        Catch
            MessageBox.Show("An error occurred while creating " & ExportName)
            writeDebug("Unable to create " & Output & " " & ExportName)
        End Try
        CloseLater(Strings.Right(DrawSource, Strings.Len(DrawSource) - Strings.InStrRev(DrawSource, " \ ")), _invApp.Documents.Open(DrawSource, True))
        CloseLater(Strings.Right(Archive, Strings.Len(Archive) - Strings.InStrRev(Archive, " \ ")), _invApp.Documents.Open(Archive, True))
        Me.Focus()
    End Sub
    Private Sub CreateFlatPattern(oCompDef As ComponentDefinition, DrawingName As String, oDoc As Document, DXF As Boolean)

        Dim oDef As ControlDefinition
        If oCompDef.HasFlatPattern = False Then
            Try
                'go to flat pattern or create it if it doesn't exist
                oCompDef.Unfold()
                oDef = _invApp.CommandManager.ControlDefinitions.Item("PartSwitchRepresentationCmd")
                oDef.Execute()
                'Return to folded model
                oDef = _invApp.CommandManager.ControlDefinitions.Item("PartConvertToSheetMetalCmd")
                oDef.Execute()
                'Get active document
            Catch
                If DXF = True Then
                    MsgBox("An Error occurred during the flat pattern creation Of " & Strings.Right(oDoc.FullFileName, Len(oDoc.FullFileName) - InStrRev(oDoc.FullFileName, "\")) & "." & vbNewLine _
                      & "The .dxf was Not created" & vbNewLine & "Some parts require the flat patterns To be created manually." _
                      & vbNewLine & vbNewLine & Err.Description)
                Else
                    MsgBox("An Error occurred during the flat pattern creation Of " & Strings.Right(oDoc.FullFileName, Len(oDoc.FullFileName) - InStrRev(oDoc.FullFileName, "\")) & "." & vbNewLine _
                       & "The area for this part will not be included.")
                End If
                writeDebug("Error accurred during flat pattern creation of " & DrawingName)
            End Try
        End If
    End Sub
    Private Sub SMDXF(oDoc As Document, DXFSource As String, Flatpattern As Boolean, DrawingName As String)
        'Dim oPartDoc As Document = _invApp.ActiveDocument
        _invApp.SilentOperation = True
        CreateFlatPattern(oDoc.ComponentDefinition, DrawingName, oDoc, True)

        If My.Computer.FileSystem.DirectoryExists(Strings.Left(DXFSource, InStrRev(DXFSource, "\"))) = False Then
            MkDir(Strings.Left(DXFSource, InStrRev(DXFSource, "\")))
        End If
        If My.Computer.FileSystem.DirectoryExists(Strings.Left(DXFSource, InStrRev(DXFSource, "\")) & "Archived\") = False Then
            MkDir(Strings.Left(DXFSource, InStrRev(DXFSource, "\")) & "Archived\")
        End If
        Dim Filetype As String = ".dxf"
        Search_For_Duplicates(DXFSource, DrawingName, Filetype)
        Dim oDataIO As DataIO
        oDataIO = oDoc.ComponentDefinition.DataIO
        'Build the string that defines the format of the DXF file
        Dim sOut As String
        '&OuterProfileLayer=IV_INTERIOR_PROFILES"
        sOut = "FLAT PATTERN DXF?AcadVersion=2010" &
        "&OuterProfileLayer=IV_OUTER_PROFILE&OuterProfileLayerLineType=37633&OuterProfileLayerLineWeight=0,0500&OuterProfileLayerColor=0;0;0" &
        "&InteriorProfilesLayer=IV_INTERIOR_PROFILES&InteriorProfilesLayerLineType=37633&InteriorProfilesLayerLineWeight=0,0500&InteriorProfilesLayerColor=0;0;0" &
        "&InvisibleLayers=IV_TANGENT;IV_BEND;IV_Bend_Down;IV_Bend​_Up;IV_ARC_CENTERS"
        '
        '
        Try
            oDataIO.WriteDataToFile(sOut, DXFSource)
        Catch
            MsgBox("An error occurred while saving " & Strings.Right(oDoc.FullFileName, Len(oDoc.FullFileName) - InStrRev(oDoc.FullFileName, "\")) & "." & vbNewLine _
                  & "Ensure the dxf is not open or write protected." _
                  & vbNewLine & vbNewLine & Err.Description)
            writeDebug("An error occurred while saving " & Strings.Right(oDoc.FullFileName, Len(oDoc.FullFileName) - InStrRev(oDoc.FullFileName, "\")) & "." & vbNewLine _
                  & "Ensure the dxf is not open or write protected." _
                  & vbNewLine & vbNewLine & Err.Description)
            Err.Clear()
        End Try
        _invApp.SilentOperation = False
        ' "&TangentLayer=IV_TANGENT&deletelayer=TrueTangentLayerLineType=37633&TangentLayerLineWeight=0,0500&TangentLayerColor=0;0;0" & _
        '"&BendLayer=BEND&BendLayerLineType=37633&BendLayerLineWeight=0,0500&BendLayerColor=0;0;0" & _
        '"&BendDownLayer=BENDD&BendDownLayerLineType=37633&BendDownLayerLineWeight=0,0500&BendDownLayerColor=0;0;0" & _
        ' "&ArcCentersLayer=ARC&ArcCentersLayerLineType=37633&ArcCentersLayerLineWeight=0,0500&ArcCentersLayerColor=0;0;0" & _
        ' "&OuterProfileLayer=IV_OUTER_PROFILE&OuterProfileLayerLineType=37633&OuterProfileLayerLineWeight=0,0500&OuterProfileLayerColor=0;0;0" & _
        ' "&InteriorProfilesLayer=IV_INTERIOR_PROFILES&InteriorProfilesLayerLineType=37633&InteriorProfilesLayerLineWeight=0,0500&InteriorProfilesLayerColor=0;0;0"

        '' ---- Creating DXF from flat pattern through the .INI file doesn't seem to be possible
        '' ---- the code below is kept for reference only
        'Dim oDXFAddin As TranslatorAddIn = Nothing
        'For i As Long = 1 To _invApp.ApplicationAddIns.Count
        ' If _invApp.ApplicationAddIns.Item(i).ClassIdString = "{C24E3AC4-122E-11D5-8E91-0010B541CD80}" Then
        ' oDXFAddin = _invApp.ApplicationAddIns.Item(i)
        ' Exit For
        ' End If
        ' Next
        ' If oDXFAddin Is Nothing Then
        ' MsgBox("The DXF Add-in could not be found")
        ' Exit Sub
        ' End If
        ' If Not oDXFAddin.Activated Then oDXFAddin.Activate()
        ' Dim oContext As TranslationContext = _invApp.TransientObjects.CreateTranslationContext
        ' oContext.Type = IOMechanismEnum.kFileBrowseIOMechanism
        ' Dim oNameValueMap As NameValueMap = _invApp.TransientObjects.CreateNameValueMap
        ' If oDXFAddin.HasSaveCopyAsOptions(oPartDoc, oContext, oNameValueMap) Then
        ' For x = 1 To oNameValueMap.Count
        ' Next
        ' oNameValueMap.Value("DwgVersion") = 25
        ' Dim strIniFile As String = "P:\Inventor System Files\Inventor Add Ins and Programs\IO Profile.ini"
        ' oNameValueMap.Value("Export_Acad_IniFile") = strIniFile
        ' End If

        ' Dim oOutputFile As DataMedium = _invApp.TransientObjects.CreateDataMedium
        ' oOutputFile.FileName = DXFSource

        'Call oDXFAddin.SaveCopyAs(oPartDoc, oContext, oNameValueMap, oOutputFile)

    End Sub
    Private Sub releaseObject(ByVal obj As Object)
        Try
            System.Runtime.InteropServices.Marshal.ReleaseComObject(obj)
            obj = Nothing
        Catch ex As Exception
            obj = Nothing
        Finally
            GC.Collect()
        End Try
    End Sub
    Private Sub ExportList(ByRef oDoc As Inventor.Document, ByVal Occurrences As ComponentOccurrences,
                           ByRef Properties As Double, ByRef Counter As Integer, ExcelDoc As Excel.Workbook, Total As Integer)
        Dim Occ As ComponentOccurrence
        Dim PropSets As Inventor.PropertySets
        Dim StockNo, Material, PartNo, Description, Mass, Title As String
        Dim Offset As Integer = 1
        Title = "Extracting"
        For Each Occ In Occurrences
            Counter += 1
            'Check if the document is a part or assembly
            If Occ.DefinitionDocumentType = DocumentTypeEnum.kPartDocumentObject Or
                DocumentTypeEnum.kAssemblyDocumentObject And Occ.BOMStructure = BOMStructureEnum.kPurchasedBOMStructure Then
                'Get iProperties for this document
                PropSets = Occ.Definition.Document.PropertySets()
                StockNo = PropSets.Item("{32853F0F-3444-11D1-9E93-0060B03C1CA6}").ItemByPropId("55").Value.ToString
                Material = PropSets.Item("{32853F0F-3444-11D1-9E93-0060B03C1CA6}").ItemByPropId("20").Value.ToString
                PartNo = PropSets.Item("{32853F0F-3444-11D1-9E93-0060B03C1CA6}").ItemByPropId("5").Value.ToString
                Description = PropSets.Item("{32853F0F-3444-11D1-9E93-0060B03C1CA6}").ItemByPropId("29").Value.ToString
                Mass = Math.Round(Occ.MassProperties.Mass, 2)
                ExcelDoc.ActiveSheet.Range("A" & Offset).Value = StockNo
                ExcelDoc.ActiveSheet.Range("B" & Offset).Value = Material
                ExcelDoc.ActiveSheet.Range("C" & Offset).Value = PartNo
                ExcelDoc.ActiveSheet.Range("D" & Offset).Value = Description
                ExcelDoc.ActiveSheet.Range("E" & Offset).Value = Mass
                ProgressBar(Total, Counter, Title, Occ.Name)
                Offset += 1
            Else
                ExportList(oDoc, Occ.SubOccurrences, Properties, Counter, ExcelDoc, Total)
            End If
        Next
    End Sub

#End Region
#Region "List Calls"
    Private Sub txtSearch_Click(sender As Object, e As System.EventArgs) Handles txtSearch.Click
        If txtSearch.ForeColor = Drawing.Color.Gray Then
            txtSearch.ForeColor = Drawing.Color.Black
            txtSearch.Text = ""

        End If
    End Sub
    Private Sub txtSearch_LostFocus(sender As Object, e As EventArgs) Handles txtSearch.LostFocus
        If txtSearch.Text = "" Then
            txtSearch.ForeColor = Drawing.Color.Gray
            txtSearch.Text = "Search"
        End If
    End Sub
    Private Sub txtSearch_TextChanged(sender As Object, e As System.EventArgs) Handles txtSearch.TextChanged
        If txtSearch.ForeColor = Drawing.Color.Gray Or txtSearch.Text = "Search" Then Exit Sub
        LVSubFiles.Items.Clear()
        FilterSubFiles()
    End Sub
    Private Sub FilterSubFiles()
        LVSubFiles.Items.Clear()
        Dim search As String = txtSearch.Text
        If txtSearch.Text = "Search" Then
            search = ""
        End If
        If CMSHeirarchical.Checked = True Then
            For Each pair As KeyValuePair(Of String, String) In SubFiles

                If Strings.InStr(pair.Key, search) <> 0 Then

                    If Strings.InStr(pair.Key, "(DNE)") <> 0 Then
                        If CMSMissingDWG.Text = "Hide Missing Drawings" Then
                            LVSubFiles.Items.Add(Strings.Replace(pair.Key, "(DNE)", ""))
                            LVSubFiles.Items(LVSubFiles.Items.Count - 1).Checked = False
                            LVSubFiles.Items(LVSubFiles.Items.Count - 1).ForeColor = Drawing.Color.Gray
                        End If
                    ElseIf Strings.InStr(pair.Key, "(REF)") <> 0 Then
                        If CMSReference.Text = "Hide Reference Drawings" Then
                            LVSubFiles.Items.Add(Strings.Replace(pair.Key, "(REF)", ""))
                            LVSubFiles.Items(LVSubFiles.Items.Count - 1).Checked = False
                            LVSubFiles.Items(LVSubFiles.Items.Count - 1).ForeColor = Drawing.Color.Blue
                        End If
                    ElseIf Strings.InStr(pair.Key, "(PPM)") <> 0 Then
                        If CMSMissingParts.Text = "Hide Missing Parts" Then
                            LVSubFiles.Items.Add(Strings.Replace(pair.Key, "(PPM)", ""))
                            LVSubFiles.Items(LVSubFiles.Items.Count - 1).Checked = False
                            LVSubFiles.Items(LVSubFiles.Items.Count - 1).ForeColor = Drawing.Color.Red
                        End If
                    Else
                        LVSubFiles.Items.Add(pair.Key)
                        LVSubFiles.Items(LVSubFiles.Items.Count - 1).Checked = True
                    End If
                End If

            Next

        Else
            For Each pair As KeyValuePair(Of String, String) In AlphaSub
                If Strings.InStr(pair.Key, search) <> 0 Then
                    If Strings.InStr(pair.Key, "(DNE)") <> 0 Then
                        If CMSMissingDWG.Text = "Hide Missing Drawings" Then
                            LVSubFiles.Items.Add(Strings.Replace(pair.Key, "(DNE)", ""))
                            LVSubFiles.Items(LVSubFiles.Items.Count - 1).Checked = False
                            LVSubFiles.Items(LVSubFiles.Items.Count - 1).ForeColor = Drawing.Color.Gray
                        End If
                    ElseIf Strings.InStr(pair.Key, "(REF)") <> 0 Then
                        If CMSReference.Text = "Hide Reference Drawings" Then
                            LVSubFiles.Items.Add(Strings.Replace(pair.Key, "(REF)", ""))
                            LVSubFiles.Items(LVSubFiles.Items.Count - 1).Checked = False
                            LVSubFiles.Items(LVSubFiles.Items.Count - 1).ForeColor = Drawing.Color.Blue
                        End If
                    ElseIf Strings.InStr(pair.Key, "(PPM)") <> 0 Then
                        If CMSMissingParts.Text = "Hide Missing Parts" Then
                            LVSubFiles.Items.Add(Strings.Replace(pair.Key, "(PPM)", ""))
                            LVSubFiles.Items(LVSubFiles.Items.Count - 1).Checked = False
                            LVSubFiles.Items(LVSubFiles.Items.Count - 1).ForeColor = Drawing.Color.Red
                        End If
                    Else
                        LVSubFiles.Items.Add(pair.Key)
                        LVSubFiles.Items(LVSubFiles.Items.Count - 1).Checked = True
                    End If
                End If
            Next
        End If
    End Sub
    Private Sub LVSubFiles_MouseUp(sender As Object, e As MouseEventArgs) Handles LVSubFiles.MouseUp
        Try
            If e.Button = Windows.Forms.MouseButtons.Right Then
                CMSSubFiles.Show(Cursor.Position)
                If LVSubFiles.Items(LVSubFiles.FocusedItem.Index).Checked = True Then
                    LVSubFiles.Items(LVSubFiles.FocusedItem.Index).Checked = False
                Else
                    LVSubFiles.Items(LVSubFiles.FocusedItem.Index).Checked = True
                End If
            End If
        Catch
        End Try
    End Sub
    Private Sub CMSAlphabetical_Click(sender As Object, e As EventArgs) Handles CMSAlphabetical.Click
        LVSubFiles.Items.Clear()
        For Each pair As KeyValuePair(Of String, String) In AlphaSub
            If Strings.InStr(pair.Key, "(DNE)") <> 0 Then
                If CMSMissingDWG.Text = "Hide Missing Drawings" Then
                    LVSubFiles.Items.Add(Strings.Replace(pair.Key, "(DNE)", ""))
                    LVSubFiles.Items(LVSubFiles.Items.Count - 1).Checked = False
                    LVSubFiles.Items(LVSubFiles.Items.Count - 1).ForeColor = Drawing.Color.Gray
                End If
            ElseIf Strings.InStr(pair.Key, "(REF)") <> 0 Then
                If CMSReference.Text = "Hide Reference Drawings" Then
                    LVSubFiles.Items.Add(Strings.Replace(pair.Key, "(REF)", ""))
                    LVSubFiles.Items(LVSubFiles.Items.Count - 1).Checked = False
                    LVSubFiles.Items(LVSubFiles.Items.Count - 1).ForeColor = Drawing.Color.Blue
                End If
            ElseIf Strings.InStr(pair.Key, "(PPM)") <> 0 Then
                If CMSMissingParts.Text = "Hide Missing Parts" Then
                    LVSubFiles.Items.Add(Strings.Replace(pair.Key, "(PPM)", ""))
                    LVSubFiles.Items(LVSubFiles.Items.Count - 1).Checked = False
                    LVSubFiles.Items(LVSubFiles.Items.Count - 1).ForeColor = Drawing.Color.Red
                End If
            Else
                LVSubFiles.Items.Add(pair.Key)
                LVSubFiles.Items(LVSubFiles.Items.Count - 1).Checked = True
            End If
        Next
        CMSAlphabetical.Checked = True
        CMSHeirarchical.Checked = False
    End Sub
    Private Sub CMSHeirarchical_Click(sender As Object, e As EventArgs) Handles CMSHeirarchical.Click
        LVSubFiles.Items.Clear()
        For Each pair As KeyValuePair(Of String, String) In SubFiles
            If Strings.InStr(pair.Key, "(DNE)") <> 0 Then
                If CMSMissingDWG.Text = "Hide Missing Drawings" Then
                    LVSubFiles.Items.Add(Strings.Replace(pair.Key, "(DNE)", ""))
                    LVSubFiles.Items(LVSubFiles.Items.Count - 1).Checked = False
                    LVSubFiles.Items(LVSubFiles.Items.Count - 1).ForeColor = Drawing.Color.Gray
                End If
            ElseIf Strings.InStr(pair.Key, "(REF)") <> 0 Then
                If CMSReference.Text = "Hide Reference Drawings" Then
                    LVSubFiles.Items.Add(Strings.Replace(pair.Key, "(REF)", ""))
                    LVSubFiles.Items(LVSubFiles.Items.Count - 1).Checked = False
                    LVSubFiles.Items(LVSubFiles.Items.Count - 1).ForeColor = Drawing.Color.Blue
                End If
            ElseIf Strings.InStr(pair.Key, "(PPM)") <> 0 Then
                If CMSMissingParts.Text = "Hide Missing Parts" Then
                    LVSubFiles.Items.Add(Strings.Replace(pair.Key, "(PPM)", ""))
                    LVSubFiles.Items(LVSubFiles.Items.Count - 1).Checked = False
                    LVSubFiles.Items(LVSubFiles.Items.Count - 1).ForeColor = Drawing.Color.Red
                End If
            Else
                LVSubFiles.Items.Add(pair.Key)
                LVSubFiles.Items(LVSubFiles.Items.Count - 1).Checked = True
            End If
        Next
        CMSAlphabetical.Checked = False
        CMSHeirarchical.Checked = True
    End Sub
    Private Sub CMSSpreadsheet_Click(sender As Object, e As EventArgs)
        ExportText(CMSHeirarchical.Checked, CMSMissingDWG.Checked, False, True)
    End Sub
    Private Sub LVSubFiles_Click(sender As Object, e As EventArgs) Handles LVSubFiles.Click
        If LVSubFiles.Items(LVSubFiles.FocusedItem.Index).ForeColor <> Drawing.Color.Gray And
            LVSubFiles.Items(LVSubFiles.FocusedItem.Index).ForeColor <> Drawing.Color.Red Then
            If LVSubFiles.Items(LVSubFiles.FocusedItem.Index).Checked = True Then
                LVSubFiles.Items(LVSubFiles.FocusedItem.Index).Checked = False
            Else
                LVSubFiles.Items(LVSubFiles.FocusedItem.Index).Checked = True
            End If
        End If
    End Sub
    Private Sub LVSubFiles_ItemChecked(sender As Object, e As ItemCheckedEventArgs) Handles LVSubFiles.ItemChecked
        If LVSubFiles.Items(e.Item.Index).ForeColor = Drawing.Color.Gray OrElse
            LVSubFiles.Items(e.Item.Index).ForeColor = Drawing.Color.Red Then
            e.Item.Checked = False
        End If
        LVSubFiles.FocusedItem = Nothing
        If LVSubFiles.Items.Count > 10 Then LVSubFiles.Height = lstOpenfiles.Height - 25
    End Sub
    Private Sub CMSSubSpreadsheet_Click(sender As Object, e As EventArgs) Handles CMSSubSpreadsheet.Click
        ExportText(CMSHeirarchical.Checked, CMSMissingDWG.Checked, False, False)
    End Sub
    Private Sub CMSSubText_Click(sender As Object, e As EventArgs) Handles CMSSubText.Click
        ExportText(CMSHeirarchical.Checked, CMSMissingDWG.Checked, True, False)
    End Sub
    Private Sub ShowMissingDrawingsToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles CMSMissingDWG.Click
        If CMSMissingDWG.Text = "Show Missing Drawings" Then
            CMSMissingDWG.Text = "Hide Missing Drawings"
        Else
            CMSMissingDWG.Text = "Show Missing Drawings"
        End If
        FilterSubFiles()
    End Sub
    Private Sub HideReferenceDrawingsToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles CMSReference.Click
        If CMSReference.Text = "Show Reference Drawings" Then
            CMSReference.Text = "Hide Reference Drawings"
        Else
            CMSReference.Text = "Show Reference Drawings"
        End If
        FilterSubFiles()
    End Sub
    Private Sub HideMissingPartsToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles CMSMissingParts.Click
        If CMSMissingParts.Text = "Show Missing Parts" Then
            CMSMissingParts.Text = "Hide Missing Parts"
        Else
            CMSMissingParts.Text = "Show Missing Parts"
        End If
        FilterSubFiles()
    End Sub
    Private Sub ExportText(Heirarchical As Boolean, Hidden As Boolean, Text As Boolean, OpenFiles As Boolean)
        Dim filetype As String = ""
        If Text = True Then
            filetype = ".txt"
            Dim File_Name As String = My.Computer.FileSystem.SpecialDirectories.Temp & "\Parts List" & filetype
            If System.IO.File.Exists(File_Name) = False Then
                System.IO.File.Create(File_Name).Dispose()
            End If
            If My.Computer.FileSystem.FileExists(File_Name) Then
                Kill(File_Name)
            End If
            If System.IO.File.Exists(File_Name) = False Then
                System.IO.File.Create(File_Name).Dispose()
            End If
            Dim objWriter As New System.IO.StreamWriter(File_Name, True)
            'objWriter.WriteLine(My.Computer.FileSystem.OpenTextFileWriter(File_Name, True))
            Dim Parents As String = ""
            Try
                objWriter.WriteLine("File created at-- " & DateTime.Now)
                For Each item In lstOpenfiles.CheckedItems
                    Parents = vbTab & item.ToString & vbNewLine & Parents
                Next
                objWriter.WriteLine(vbNewLine & "Files associated with " & vbNewLine & vbNewLine & Parents & vbNewLine & vbNewLine &
                                "Filename:" & vbTab & vbTab & vbTab & vbTab & "Location:" & vbNewLine)
                If CMSHeirarchical.Checked = True Then
                    For Each pair As KeyValuePair(Of String, String) In SubFiles
                        If CMSMissingDWG.Text = "Hide Missing Drawings" And InStr(pair.Key, "(DNE)") <> 0 Then
                            objWriter.WriteLine(Strings.Replace(pair.Key, "(DNE)", "") & ":" & vbTab & vbTab & vbTab & pair.Value)
                        ElseIf CMSMissingParts.Text = "Hide Missing Parts" And InStr(pair.Key, "(PPM)") <> 0 Then
                            objWriter.WriteLine(Strings.Replace(pair.Key, "(PPM)", "") & ":" & vbTab & vbTab & vbTab & pair.Value)
                        ElseIf CMSReference.Text = "Hide Reference Parts" And InStr(pair.Key, "(REF)") <> 0 Then
                            objWriter.WriteLine(Strings.Replace(pair.Key, "(REF)", "") & ":" & vbTab & vbTab & vbTab & pair.Value)
                        ElseIf InStr(pair.Key, "(DNE)") = 0 And InStr(pair.Key, "(PPM)") = 0 And InStr(pair.Key, "(REF)") = 0 Then
                            objWriter.WriteLine(pair.Key & ":" & vbTab & vbTab & vbTab & pair.Value)
                        End If
                    Next
                ElseIf CMSHeirarchical.Checked = False Then
                    For Each pair As KeyValuePair(Of String, String) In AlphaSub
                        If CMSMissingDWG.Text = "Hide Missing Drawings" And InStr(pair.Key, "(DNE)") <> 0 Then
                            objWriter.WriteLine(Strings.Replace(pair.Key, "(DNE)", "") & ":" & vbTab & vbTab & vbTab & pair.Value)
                        ElseIf CMSMissingParts.Text = "Hide Missing Parts" And InStr(pair.Key, "(PPM)") <> 0 Then
                            objWriter.WriteLine(Strings.Replace(pair.Key, "(PPM)", "") & ":" & vbTab & vbTab & vbTab & pair.Value)
                        ElseIf CMSReference.Text = "Hide Reference Parts" And InStr(pair.Key, "(REF)") <> 0 Then
                            objWriter.WriteLine(Strings.Replace(pair.Key, "(REF)", "") & ":" & vbTab & vbTab & vbTab & pair.Value)
                        ElseIf InStr(pair.Key, "(DNE)") = 0 And InStr(pair.Key, "(PPM)") = 0 And InStr(pair.Key, "(REF)") = 0 Then
                            objWriter.WriteLine(pair.Key & ":" & vbTab & vbTab & vbTab & pair.Value)
                        End If
                    Next
                End If
            Catch ex As Exception
                MessageBox.Show(ex.Message, "Error Creating Log File", MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
            End Try
            objWriter.Close()
        Else
            filetype = ".xls"
            If My.Computer.FileSystem.FileExists(My.Computer.FileSystem.SpecialDirectories.Temp & "\Parts List" & filetype) Then
                Kill(My.Computer.FileSystem.SpecialDirectories.Temp & "\Parts List" & filetype)
            End If
            Dim xlApp As Excel.Application = New Microsoft.Office.Interop.Excel.Application()
            If xlApp Is Nothing Then
                MessageBox.Show("Excel is not properly installed!!")
                Return
            End If
            Dim xlWorkBook As Excel.Workbook
            Dim xlWorkSheet As Excel.Worksheet
            xlApp.Visible = False
            Dim misValue As Object = System.Reflection.Missing.Value
            xlWorkBook = xlApp.Workbooks.Add(misValue)
            xlWorkSheet = xlWorkBook.Sheets("sheet1")
            xlWorkSheet.Cells(2, 2) = "Files associated with"
            Dim X As Integer
            For X = 0 To lstOpenfiles.CheckedItems.Count - 1
                xlWorkSheet.Cells(3 + X, 2) = lstOpenfiles.Items(X)
            Next
            If CMSHeirarchical.Checked = True Then
                For Each pair As KeyValuePair(Of String, String) In SubFiles
                    If CMSMissingDWG.Text = "Hide Missing Drawings" And InStr(pair.Key, "(DNE)") <> 0 Then
                        xlWorkSheet.Cells(X + 4, 1 + Regex.Match(pair.Key, "[^ ]").Index / 3) = Strings.Replace(Trim(pair.Key), "(DNE)", "")
                        xlWorkSheet.Cells(X + 4, 2 + Regex.Match(pair.Key, "[^ ]").Index / 3) = pair.Value
                    ElseIf CMSMissingParts.Text = "Hide Missing Parts" And InStr(pair.Key, "(PPM)") <> 0 Then
                        xlWorkSheet.Cells(X + 4, 1 + Regex.Match(pair.Key, "[^ ]").Index / 3) = Strings.Replace(Trim(pair.Key), "(PPM)", "")
                        xlWorkSheet.Cells(X + 4, 2 + Regex.Match(pair.Key, "[^ ]").Index / 3) = pair.Value
                    ElseIf CMSReference.Text = "Hide Reference Parts" And InStr(pair.Key, "(REF)") <> 0 Then
                        xlWorkSheet.Cells(X + 4, 1 + Regex.Match(pair.Key, "[^ ]").Index / 3) = Strings.Replace(Trim(pair.Key), "(REF)", "")
                        xlWorkSheet.Cells(X + 4, 2 + Regex.Match(pair.Key, "[^ ]").Index / 3) = pair.Value
                    Else
                        If Not InStr(pair.Key, "(DNE)") <> 0 And
                                Not InStr(pair.Key, "(PPM)") <> 0 And
                           Not InStr(pair.Key, "(REF)") <> 0 Then
                            xlWorkSheet.Cells(X + 4, 1 + Regex.Match(pair.Key, "[^ ]").Index / 3) = Trim(pair.Key)
                            xlWorkSheet.Cells(X + 4, 2 + Regex.Match(pair.Key, "[^ ]").Index / 3) = pair.Value
                        End If
                    End If
                    X += 1
                Next
            Else
                For Each pair As KeyValuePair(Of String, String) In AlphaSub
                    If CMSMissingDWG.Text = "Hide Missing Drawings" And InStr(pair.Key, "(DNE)") <> 0 Then
                        xlWorkSheet.Cells(X + 4, 1 + Regex.Match(pair.Key, "[^ ]").Index / 3) = Strings.Replace(Trim(pair.Key), "(DNE)", "")
                        xlWorkSheet.Cells(X + 4, 2 + Regex.Match(pair.Key, "[^ ]").Index / 3) = pair.Value
                    ElseIf CMSMissingParts.Text = "Hide Missing Parts" And InStr(pair.Key, "(PPM)") <> 0 Then
                        xlWorkSheet.Cells(X + 4, 1 + Regex.Match(pair.Key, "[^ ]").Index / 3) = Strings.Replace(Trim(pair.Key), "(PPM)", "")
                        xlWorkSheet.Cells(X + 4, 2 + Regex.Match(pair.Key, "[^ ]").Index / 3) = pair.Value
                    ElseIf CMSReference.Text = "Hide Reference Parts" And InStr(pair.Key, "(REF)") <> 0 Then
                        xlWorkSheet.Cells(X + 4, 1 + Regex.Match(pair.Key, "[^ ]").Index / 3) = Strings.Replace(Trim(pair.Key), "(REF)", "")
                        xlWorkSheet.Cells(X + 4, 2 + Regex.Match(pair.Key, "[^ ]").Index / 3) = pair.Value
                    Else
                        If Not InStr(pair.Key, "(DNE)") <> 0 And
                                Not InStr(pair.Key, "(PPM)") <> 0 And
                           Not InStr(pair.Key, "(REF)") <> 0 Then
                            xlWorkSheet.Cells(X + 4, 1 + Regex.Match(pair.Key, "[^ ]").Index / 3) = Trim(pair.Key)
                            xlWorkSheet.Cells(X + 4, 2 + Regex.Match(pair.Key, "[^ ]").Index / 3) = pair.Value
                        End If
                    End If
                    X += 1
                Next
            End If
            xlWorkBook.SaveAs(My.Computer.FileSystem.SpecialDirectories.Temp & "\Parts List" & filetype, Excel.XlFileFormat.xlWorkbookNormal, misValue, misValue, misValue, misValue,
            Excel.XlSaveAsAccessMode.xlExclusive, misValue, misValue, misValue, misValue, misValue)
            xlWorkBook.Close(True, misValue, misValue)
            xlApp.Quit()
            releaseObject(xlWorkSheet)
            releaseObject(xlWorkBook)
            releaseObject(xlApp)

        End If
        Process.Start(My.Computer.FileSystem.SpecialDirectories.Temp & "\Parts List" & filetype)
    End Sub

#End Region
#Region "Top Menus"
    Private Sub mnuActDeact_Click(sender As Object, e As EventArgs) Handles mnuActDeact.Click
        If isGenuine Then
            ' deactivate product without deleting the product key
            ' allows the user to easily reactivate
            Try
                ta.Deactivate(False)
            Catch ex As TurboActivateException
                MessageBox.Show("Failed to deactivate: " + ex.Message)
                Return
            End Try

            isGenuine = False
            ShowTrial(True)
        Else
            'Note: you can launch the TurboActivate wizard or you can create you own interface

            'launch TurboActivate.exe to get the product key from the user
            Dim TAProcess As New Process
            If My.Application.IsNetworkDeployed Then
                TAProcess.StartInfo.FileName = IO.Path.Combine(IO.Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().CodeBase), "TurboActivate.exe")
            Else
                TAProcess.StartInfo.FileName = IO.Path.Combine(IO.Path.GetDirectoryName(My.Application.Info.DirectoryPath), "debug\TurboActivate.exe")
            End If
            TAProcess.EnableRaisingEvents = True
            AddHandler TAProcess.Exited, New EventHandler(AddressOf p_Exited)
            TAProcess.Start()
        End If
    End Sub
    ''' This event handler is called when TurboActivate.exe closes.
    Private Sub p_Exited(ByVal sender As Object, ByVal e As EventArgs)

        ' remove the event
        RemoveHandler DirectCast(sender, Process).Exited, New EventHandler(AddressOf p_Exited)

        ' the UI thread is running asynchronous to TurboActivate closing
        ' that's why we can't call TAIsActivated(); directly
        Invoke(New IsActivatedDelegate(AddressOf CheckIfActivated))
    End Sub
    ''' Rechecks if we're activated -- if so enable the app features.
    Private Sub CheckIfActivated()
        ' recheck if activated
        Dim isNowActivated As Boolean = False

        Try
            isNowActivated = ta.IsActivated
        Catch ex As TurboActivateException
            MessageBox.Show("Failed to check if activated: " + ex.Message)
            Exit Sub
        End Try

        If isNowActivated Then
            isGenuine = True
            ReEnableAppFeatures()
            ShowTrial(False)
        End If
    End Sub
    Private Sub ShowTrial(ByVal show As Boolean)

        'lblTrialMessage.Visible = show
        'btnExtendTrial.Visible = show

        mnuActDeact.Text = (If(show, "Activate...", "Deactivate"))

        If show Then
            Dim trialDaysRemaining As UInteger = 0UI

            Try
                ta.UseTrial(trialFlags)

                ' get the number of remaining trial days
                trialDaysRemaining = ta.TrialDaysRemaining(trialFlags)
            Catch ex As TurboActivateException
                MessageBox.Show("Failed to start the trial: " + ex.Message)
            End Try

            ' if no more trial days then disable all app features
            If trialDaysRemaining = 0 Then
                DisableAppFeatures()
            Else
                'lblTrialMessage.Text = "Your trial expires in " & trialDaysRemaining & " days."
            End If
        End If
    End Sub
    Private Sub DisableAppFeatures()
        'TODO: disable all the features of the program
    End Sub
    Private Sub ReEnableAppFeatures()
        'TODO: re-enable all the features of the program
    End Sub
    Private Sub IFoundABugToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles IFoundABugToolStripMenuItem.Click
        MessageBox.Show("Issues with the program can be sent to flyinggardengnomestudios@gmail.com. Please include a description of the problem and how to re-create it. You can attach the error-log which is located in: " & My.Computer.FileSystem.SpecialDirectories.Temp & "\debug.txt. " & "If possible, it is always helpful to include a pack-and-go of the assembly/part that created the error.")

    End Sub
    Private Sub AboutBatchProgramToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles AboutBatchProgramToolStripMenuItem.Click
        About.ShowDialog()
    End Sub
    Private Sub IPropertySettingsToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles IPropertySettingsToolStripMenuItem.Click
        iPropertySettings.ShowDialog()
    End Sub
    Private Sub HowToToolStripMenuItem1_Click(sender As Object, e As EventArgs) Handles HowToToolStripMenuItem1.Click
        System.IO.File.WriteAllText(My.Computer.FileSystem.SpecialDirectories.Temp & "\changelog.txt", My.Resources.Changelog)
        Process.Start(My.Computer.FileSystem.SpecialDirectories.Temp & "\changelog.txt")
    End Sub
    Private Sub DefaultSettingsToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles DefaultSettingsToolStripMenuItem.Click
        Settings.ShowDialog()
    End Sub
    Private Sub TutorialsToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles TutorialsToolStripMenuItem.Click
        Process.Start("https://www.youtube.com/playlist?list=PL5Jsz3IKLQ40-UCC4Fkm6WmQ9hINzAzEw")
    End Sub
#End Region
    Private Sub btnExit_Click(sender As System.Object, e As System.EventArgs) Handles btnExit.Click
        If btnExit.Text = "Exit" Then
            writeDebug("Batch Program closed")
            Me.Close()
            End
        End If
    End Sub
    Public Sub btnOK_Click(sender As System.Object, e As System.EventArgs) Handles btnOK.Click
        MsVistaProgressBar.ProgressBarStyle = MSVistaProgressBar.BarStyle.Continuous
        Dim NoPart, NoDraw As Boolean
        Dim Path As Documents = _invApp.Documents
        Dim oDoc As Document = Nothing
        Dim Archive As String = ""
        Dim DrawSource As String = ""
        Dim DrawingName As String = ""
        Dim DocSource, DocName As String
        Dim CheckNeeded As New CheckNeeded
        CreateOpenDocs(OpenDocs)
        If lstOpenfiles.CheckedItems.Count = 0 Then
            NoPart = True
        End If
        If LVSubFiles.CheckedItems.Count = 0 Then
            NoDraw = True
        End If
        If NoPart = True Then
            MsgBox("Please Select a Part/Assembly")
            Exit Sub
        ElseIf NoDraw = True Then
            MsgBox("Please Select a Drawing")
            Exit Sub
        End If
        If chkiProp.Checked = True Then
            Dim iProperties As New iProperties
            iProperties.PopMain(Me)
            writeDebug("Accessing iProperties")
            iProperties.PopulateiProps(Path, oDoc, Archive, DrawingName, DrawSource, OpenDocs, True)
            If DrawSource = "" And DrawingName = "" Then
                MsgBox("Could not locate drawing" & vbNewLine & "Ensure the drawing and model are saved to the same directory")
                Exit Sub
            End If
            iProperties.ShowDialog()
            chkiProp.CheckState = CheckState.Unchecked
        End If
        If ChkRevType.Checked = True Then
            writeDebug("Accessing RevType")
            Call ChangeRev()
            ChkRevType.Checked = False
        End If

        If chkCheck.Checked = True Then
            writeDebug("Accessing Checked properties")
            CheckNeeded.PopMain(Me)
            CheckNeeded.PopulateCheckNeeded(Path, oDoc, Archive, DrawingName, DrawSource, OpenDocs)
            If CheckNeeded.lstCheckNeeded.Items.Count > 0 Then
                CheckNeeded.ShowDialog()
            ElseIf CheckNeeded.lstCheckNeeded.Items.Count > 0 Then
                CheckNeeded.btnIgnore.Visible = True
                CheckNeeded.ShowDialog()
            Else
                MsgBox("All drawings have been checked.")
            End If
            chkCheck.Checked = False
        End If
        If chkRRev.Checked = True Then
            writeDebug("Accessing remove rev details")
            CheckNeeded.PopMain(Me)
            CheckNeeded.btnIgnore.Visible = False
            CheckNeeded.Label1.Text = "These drawings have revisions that can be removed:"
            CheckNeeded.Label2.Visible = False
            CheckNeeded.Refresh()
            CheckNeeded.btnOK.Visible = False
            CheckNeeded.btnCancel.Location = CheckNeeded.btnOK.Location
            CheckNeeded.btnCancel.Text = "Finished"
            CheckNeeded.PopulateCheckNeeded(Path, oDoc, Archive, DrawingName, DrawSource, OpenDocs)
            'CheckNeeded.lstCheckNeeded
            CheckNeeded.ShowDialog()
            'CheckNeeded.btnIgnore.Visible = True
            'CheckNeeded.Label1.Text = "The following files have not been checked:"
            'CheckNeeded.Label2.Visible = True
            'CheckNeeded.Refresh()
            chkRRev.Checked = False
        End If
        If chkPrint.Checked = True Then
            writeDebug("Accessing Print dialogue")
            Print.PopMain(Me)
            Print.PopulatePrint(Path, oDoc, Archive, DrawingName, DrawSource, OpenDocs, SubFiles, AlphaSub)
            Print.ShowDialog()
            'PrintDrawings(Path, oDoc, Archive, DrawingName, DrawSource, OpenDocs)
            chkPrint.Checked = False
        End If
        If chkDXF.Checked = True Then
            writeDebug("Accessing DXF creation")
            ExportCheck(Path, oDoc, Archive, DrawingName, DrawSource, "dxf")
            chkDXF.Checked = False
        End If
        If chkPDF.Checked = True Then
            writeDebug("Accessing PDF creation")
            ExportCheck(Path, oDoc, Archive, DrawingName, DrawSource, "PDF")
        End If
        If chkDWG.Checked = True Then
            writeDebug("Accessing DWG creation")
            ExportCheck(Path, oDoc, Archive, DrawingName, DrawSource, "dwg")
            chkDWG.Checked = False
        End If
        If chkOpen.Checked = True Then
            writeDebug("Accessing Open drawings command")
            OpenSelected(Path, oDoc, Archive, DrawingName, DrawSource)
            chkOpen.Checked = False
        End If
        If chkClose.Checked = True Then
            writeDebug("Accessing Close open documents")
            CloseDrawings(Path, oDoc, Archive, DrawingName, DrawSource, OpenDocs)
            chkClose.Checked = False
        End If
        RunCompleted()
        writeDebug("Cleaning up documents")
        Try
            For Each oDoc In _invApp.Documents.VisibleDocuments
                Archive = oDoc.FullFileName
                'Use the Partsource file to create the drawingsource file
                DocSource = Strings.Left(Archive, Strings.Len(Archive))
                DocName = Strings.Right(DocSource, Strings.Len(DocSource) - Strings.InStrRev(DocSource, "\"))
                OpenDocs.Add(DocName)
            Next

            If OpenDocs.Count > 0 Then
            Else
                _invApp.SilentOperation = True
                _invApp.Documents.CloseAll()
                _invApp.SilentOperation = False
            End If
        Catch ex As Exception
            MessageBox.Show(ex.Message, "Exception Details", MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
            _invApp.SilentOperation = False
        End Try

    End Sub


End Class
