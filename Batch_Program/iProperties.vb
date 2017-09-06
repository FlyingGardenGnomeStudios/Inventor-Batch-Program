Imports System.Windows.Forms
Imports Inventor
Imports System.Runtime.InteropServices
Imports System.Globalization
Public Class iProperties
    Dim _invApp As Inventor.Application
    Dim InvTitle, InvSubject, InvManager, InvCategory, InvKeywords, InvAuthor, InvCompany, Invcomments, InvDescription, InvPartNumber, InvStockNumber _
                 , InvProject, InvDesigner, InvVendor, InvModDate, InvDrawDate, InvDrawBy As Inventor.Property
    Dim InvDesignState As String
    Dim CustomPropSet As PropertySet
    Dim InvEngineer, InvStatus, InvRevision, InvDesignerState, InvCheckedBy, InvCheckDate
    Dim InvRef(0 To 9) As Inventor.Property

    Private Sub chkRev2AddRev_CheckedChanged(sender As Object, e As EventArgs) Handles chkRev2AddRev.CheckedChanged
        If chkRev2AddRev.CheckState = CheckState.Checked Then
            Call ChangeAddRev(True)
        Else
            Call ChangeAddRev(False)
        End If
    End Sub
    Private Sub chkRev3AddRev_CheckedChanged(sender As Object, e As EventArgs) Handles chkRev3AddRev.CheckedChanged
        If chkRev3AddRev.CheckState = CheckState.Checked Then
            Call ChangeAddRev(True)
        Else
            Call ChangeAddRev(False)
        End If
    End Sub
    Private Sub chkRev4AddRev_CheckedChanged(sender As Object, e As EventArgs) Handles chkRev4AddRev.CheckedChanged
        If chkRev4AddRev.CheckState = CheckState.Checked Then
            Call ChangeAddRev(True)
        Else
            Call ChangeAddRev(False)
        End If
    End Sub
    Private Sub chkRev5AddRev_CheckedChanged(sender As Object, e As EventArgs) Handles chkRev5AddRev.CheckedChanged
        If chkRev5AddRev.CheckState = CheckState.Checked Then
            Call ChangeAddRev(True)
        Else
            Call ChangeAddRev(False)
        End If
    End Sub
    Private Sub chkRev6AddRev_CheckedChanged(sender As Object, e As EventArgs) Handles chkRev6AddRev.CheckedChanged
        If chkRev6AddRev.CheckState = CheckState.Checked Then
            Call ChangeAddRev(True)
        Else
            Call ChangeAddRev(False)
        End If
    End Sub
    Private Sub chkRev7AddRev_CheckedChanged(sender As Object, e As EventArgs) Handles chkRev7AddRev.CheckedChanged
        If chkRev7AddRev.CheckState = CheckState.Checked Then
            Call ChangeAddRev(True)
        Else
            Call ChangeAddRev(False)
        End If
    End Sub
    Private Sub chkRev8AddRev_CheckedChanged(sender As Object, e As EventArgs) Handles chkRev8AddRev.CheckedChanged
        If chkRev8AddRev.CheckState = CheckState.Checked Then
            Call ChangeAddRev(True)
        Else
            Call ChangeAddRev(False)
        End If
    End Sub
    Private Sub ChangeAddRev(Click As Boolean)
        TabControl1.TabPages.Item(TabControl1.TabPages.Count - 1).Controls.Item("dtRev" & TabControl1.TabPages.Count.ToString & "Date").Enabled = Click
        TabControl1.TabPages.Item(TabControl1.TabPages.Count - 1).Controls.Item("txtRev" & TabControl1.TabPages.Count.ToString & "Des").Enabled = Click
        TabControl1.TabPages.Item(TabControl1.TabPages.Count - 1).Controls.Item("txtRev" & TabControl1.TabPages.Count.ToString & "Init").Enabled = Click
        TabControl1.TabPages.Item(TabControl1.TabPages.Count - 1).Controls.Item("txtRev" & TabControl1.TabPages.Count.ToString & "Approved").Enabled = Click
    End Sub

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
    Public Sub PopulateiProps(Path As Documents, ByRef oDoc As Document, ByRef Archive As String _
                             , ByRef DrawingName As String, ByRef DrawSource As String, OpenDocs As ArrayList _
                             , ByRef Read As Boolean)
        Dim Y, Total As Integer
        For Y = 0 To Main.LVSubFiles.Items.Count - 1
            If Main.LVSubFiles.Items(Y).Checked = True Then
                Total += 1
            End If
        Next
        'Go through iProperties of each selected item and retrieve the values
        iPropIdentify(Path, oDoc, Archive, DrawingName, DrawSource, OpenDocs, 0, Read)
    End Sub
    Public Sub iPropIdentify(Path As Documents, ByRef oDoc As Document, ByRef Archive As String _
                             , ByRef DrawingName As String, ByRef DrawSource As String, OpenDocs As ArrayList _
                             , ByRef RevTotal As Integer, ByRef Read As Boolean)
        Dim X, Y As Integer
        X = 0
        Dim Write As Boolean = False
        Dim List As String = ""
        RevTable.PopiProperties(Me)
        Main.PopiProperties(Me)
        RevTable.lstCheckFiles.Clear()
        _invApp.SilentOperation = True
        If Read = False Then Call RevTable.iPropRevTable(oDoc)
        'Go through drawings to see which ones are selected
        Dim Warning As Boolean = False
        For Y = 0 To Main.LVSubFiles.Items.Count - 1
            'Look through all sub files in open documents to get the part sourcefile
            If Main.LVSubFiles.Items(Y).Checked = True Then
                Main.MatchPart(DrawSource, DrawingName, Y)
                If DrawSource = "" And DrawingName = "" Then
                    Exit Sub
                End If
                'Path, oDoc, Archive, DrawingName, Trim(Main.lstSubfiles.Items.Item(X)), DrawSource, Y)
                'DrawSource = Strings.Left(SubFiles.Item(Y).Value, Len(SubFiles.Item(Y).Value) - 3) & "idw"
                oDoc = _invApp.Documents.Open(DrawSource, False)
                If Read = True Then
                    Call SetModelProps(oDoc)
                Else
                    Call WriteModelProps(oDoc)
                    oDoc.Update()
                End If
                Main.CloseLater(DrawingName, oDoc)
                Main.MatchDrawing(DrawSource, DrawingName, Y)
                'Path, oDoc, Archive, DrawingName, DrawSource, Y)
                'DrawSource = SubFiles.Item(Y).Value
                oDoc = _invApp.Documents.Open(DrawSource, False)
                If Read = True Then

                    Call SetDrawingProps(oDoc, Warning)
                    Call RevTableCheck(oDoc, RevTotal)
                    Call InitializeiProperties(X, DrawSource)
                    Main.ProgressBar(Main.LVSubFiles.CheckedItems.Count, Y + 1, "Reading: ", DrawingName)
                Else
                    ProgressBar2.Visible = True
                    ProgressBar2.Value = ((X + 1) / Main.LVSubFiles.CheckedItems.Count) * 100
                    ProgressBar2.PerformStep()
                    Call WriteDrawingProps(oDoc)
                    Call WriteRevTable(oDoc, OpenDocs)
                    _invApp.ActiveDocument.Update()
                End If
                X += 1
            End If

            _invApp.SilentOperation = False
        Next
        If Read = True Then
            For X = RevTotal To TabControl1.TabPages.Count - 2
                TabControl1.TabPages.Remove(TabControl1.TabPages(TabControl1.TabPages.Count - 1))
            Next
            Try
                With TabControl1.TabPages.Item(TabControl1.TabPages.Count - 1)
                    .Text = "Add Rev"
                    .Controls.Item("lblRev" & TabControl1.TabPages.Count.ToString).Visible = False
                    .Controls.Item("chkRev" & TabControl1.TabPages.Count.ToString & "AddRev").Visible = True
                    .Controls.Item("txtRev" & TabControl1.TabPages.Count.ToString & "Rev").Text = "*Varies*"
                    Call ChangeAddRev(False)
                End With
            Catch
                MsgBox("There was a problem retrieving the revision history")
                TabControl1.Enabled = False
            End Try

        End If
            ProgressBar2.Visible = False
        RevTable.Close()
        If List <> "" Then
            MsgBox("The following drawings are missing revision tables:" & vbNewLine & List & vbNewLine _
                   & "You may add the tables when the operation completes.", , "Missing Revision Table")
        End If
    End Sub
    Private Sub SetModelProps(invModelDoc As Document)
        'Set the identifiers for the model properties

        If My.Settings.Title = "Model" Then InvTitle = invModelDoc.PropertySets.Item("{F29F85E0-4FF9-1068-AB91-08002B27B3D9}").ItemByPropId("2")
        If My.Settings.Subject = "Model" Then InvSubject = invModelDoc.PropertySets.Item("{F29F85E0-4FF9-1068-AB91-08002B27B3D9}").ItemByPropId("3")
        If My.Settings.Manager = "Model" Then InvManager = invModelDoc.PropertySets.Item("{D5CDD502-2E9C-101B-9397-08002B2CF9AE}").ItemByPropId("14")
        If My.Settings.Category = "Model" Then InvCategory = invModelDoc.PropertySets.Item("{D5CDD502-2E9C-101B-9397-08002B2CF9AE}").ItemByPropId("2")
        If My.Settings.Keywords = "Model" Then InvKeywords = invModelDoc.PropertySets.Item("{F29F85E0-4FF9-1068-AB91-08002B27B3D9}").ItemByPropId("5")
        If My.Settings.Author = "Model" Then InvAuthor = invModelDoc.PropertySets.Item("{F29F85E0-4FF9-1068-AB91-08002B27B3D9}").ItemByPropId("4")
        If My.Settings.Company = "Model" Then InvCompany = invModelDoc.PropertySets.Item("{D5CDD502-2E9C-101B-9397-08002B2CF9AE}").ItemByPropId("15")
        If My.Settings.Comments = "Model" Then Invcomments = invModelDoc.PropertySets.Item("{F29F85E0-4FF9-1068-AB91-08002B27B3D9}").ItemByPropId("6")
        If My.Settings.Description = "Model" Then InvDescription = invModelDoc.PropertySets.Item("{32853F0F-3444-11D1-9E93-0060B03C1CA6}").ItemByPropId("29")
        If My.Settings.PPartNumber = "Model" Then InvPartNumber = invModelDoc.PropertySets.Item("{32853F0F-3444-11D1-9E93-0060B03C1CA6}").ItemByPropId("5")
        If My.Settings.PStockNumber = "Model" Then InvStockNumber = invModelDoc.PropertySets.Item("{32853F0F-3444-11D1-9E93-0060B03C1CA6}").ItemByPropId("55")
        If My.Settings.Project = "Model" Then InvProject = invModelDoc.PropertySets.Item("{32853F0F-3444-11D1-9E93-0060B03C1CA6}").ItemByPropId("7")
        If My.Settings.Designer = "Model" Then InvDesigner = invModelDoc.PropertySets.Item("{32853F0F-3444-11D1-9E93-0060B03C1CA6}").ItemByPropId("41")
        If My.Settings.Vendor = "Model" Then InvVendor = invModelDoc.PropertySets.Item("{32853F0F-3444-11D1-9E93-0060B03C1CA6}").ItemByPropId("30")
        If My.Settings.ModelDate = "Model" Then InvModDate = invModelDoc.PropertySets.Item("{32853F0F-3444-11D1-9E93-0060B03C1CA6}").ItemByPropId("4")
        If My.Settings.Custom = "Model" Then CustomPropSet = invModelDoc.PropertySets.Item("{D5CDD505-2E9C-101B-9397-08002B2CF9AE}")
        If My.Settings.Engineer = "Model" Then InvEngineer = invModelDoc.PropertySets.Item("{32853F0F-3444-11D1-9E93-0060B03C1CA6}").ItemByPropId("42")
        If My.Settings.Status = "Model" Then InvStatus = invModelDoc.PropertySets.Item("{32853F0F-3444-11D1-9E93-0060B03C1CA6}").ItemByPropId("17")
        If My.Settings.Revision = "Model" Then InvRevision = invModelDoc.PropertySets.Item("{F29F85E0-4FF9-1068-AB91-08002B27B3D9}").ItemByPropId("9")
        If My.Settings.DesignState = "Model" Then InvDesignState = invModelDoc.PropertySets.Item("{32853F0F-3444-11D1-9E93-0060B03C1CA6}").ItemByPropId("40").Value
        If My.Settings.CheckedBy = "Model" Then InvCheckedBy = invModelDoc.PropertySets.Item("{32853F0F-3444-11D1-9E93-0060B03C1CA6}").ItemByPropId("10")
        If My.Settings.CheckedDate = "Model" Then InvCheckDate = invModelDoc.PropertySets.Item("{32853F0F-3444-11D1-9E93-0060B03C1CA6}").ItemByPropId("11")
        If My.Settings.DrawingBy = "Model" Then InvDrawDate = invModelDoc.PropertySets.Item("{32853F0F-3444-11D1-9E93-0060B03C1CA6}").ItemByPropId("4")
        If My.Settings.DrawingDate = "Model" Then InvDrawBy = invModelDoc.PropertySets.Item("{32853F0F-3444-11D1-9E93-0060B03C1CA6}").ItemByPropId("41")
    End Sub
    Private Sub SetDrawingProps(ByRef invDrawingDoc As Document, ByRef Warning As Boolean)

        'Set the identifiers for the drawing properties
        If My.Settings.Title = "Drawing" Then InvTitle = invDrawingDoc.PropertySets.Item("{F29F85E0-4FF9-1068-AB91-08002B27B3D9}").ItemByPropId("2")
        If My.Settings.Subject = "Drawing" Then InvSubject = invDrawingDoc.PropertySets.Item("{F29F85E0-4FF9-1068-AB91-08002B27B3D9}").ItemByPropId("3")
        If My.Settings.Manager = "Drawing" Then InvManager = invDrawingDoc.PropertySets.Item("{D5CDD502-2E9C-101B-9397-08002B2CF9AE}").ItemByPropId("14")
        If My.Settings.Category = "Drawing" Then InvCategory = invDrawingDoc.PropertySets.Item("{D5CDD502-2E9C-101B-9397-08002B2CF9AE}").ItemByPropId("2")
        If My.Settings.Keywords = "Drawing" Then InvKeywords = invDrawingDoc.PropertySets.Item("{F29F85E0-4FF9-1068-AB91-08002B27B3D9}").ItemByPropId("5")
        If My.Settings.Author = "Drawing" Then InvAuthor = invDrawingDoc.PropertySets.Item("{F29F85E0-4FF9-1068-AB91-08002B27B3D9}").ItemByPropId("4")
        If My.Settings.Company = "Drawing" Then InvCompany = invDrawingDoc.PropertySets.Item("{D5CDD502-2E9C-101B-9397-08002B2CF9AE}").ItemByPropId("15")
        If My.Settings.Comments = "Drawing" Then Invcomments = invDrawingDoc.PropertySets.Item("{F29F85E0-4FF9-1068-AB91-08002B27B3D9}").ItemByPropId("6")
        If My.Settings.Description = "Drawing" Then InvDescription = invDrawingDoc.PropertySets.Item("{32853F0F-3444-11D1-9E93-0060B03C1CA6}").ItemByPropId("29")
        If My.Settings.PPartNumber = "Drawing" Then InvPartNumber = invDrawingDoc.PropertySets.Item("{32853F0F-3444-11D1-9E93-0060B03C1CA6}").ItemByPropId("5")
        If My.Settings.PStockNumber = "Drawing" Then InvStockNumber = invDrawingDoc.PropertySets.Item("{32853F0F-3444-11D1-9E93-0060B03C1CA6}").ItemByPropId("55")
        If My.Settings.Project = "Drawing" Then InvProject = invDrawingDoc.PropertySets.Item("{32853F0F-3444-11D1-9E93-0060B03C1CA6}").ItemByPropId("7")
        If My.Settings.Designer = "Drawing" Then InvDesigner = invDrawingDoc.PropertySets.Item("{32853F0F-3444-11D1-9E93-0060B03C1CA6}").ItemByPropId("41")
        If My.Settings.Vendor = "Drawing" Then InvVendor = invDrawingDoc.PropertySets.Item("{32853F0F-3444-11D1-9E93-0060B03C1CA6}").ItemByPropId("30")
        If My.Settings.ModelDate = "Drawing" Then InvModDate = invDrawingDoc.PropertySets.Item("{32853F0F-3444-11D1-9E93-0060B03C1CA6}").ItemByPropId("4")
        If My.Settings.Custom = "Drawing" Then CustomPropSet = invDrawingDoc.PropertySets.Item("{D5CDD505-2E9C-101B-9397-08002B2CF9AE}")
        If My.Settings.Engineer = "Drawing" Then InvEngineer = invDrawingDoc.PropertySets.Item("{32853F0F-3444-11D1-9E93-0060B03C1CA6}").ItemByPropId("42")
        If My.Settings.Status = "Drawing" Then InvStatus = invDrawingDoc.PropertySets.Item("{32853F0F-3444-11D1-9E93-0060B03C1CA6}").ItemByPropId("17")
        If My.Settings.Revision = "Drawing" Then InvRevision = invDrawingDoc.PropertySets.Item("{F29F85E0-4FF9-1068-AB91-08002B27B3D9}").ItemByPropId("9")
        If My.Settings.DesignState = "Drawing" Then InvDesignState = invDrawingDoc.PropertySets.Item("{32853F0F-3444-11D1-9E93-0060B03C1CA6}").ItemByPropId("40").Value
        If My.Settings.CheckedBy = "Drawing" Then InvCheckedBy = invDrawingDoc.PropertySets.Item("{32853F0F-3444-11D1-9E93-0060B03C1CA6}").ItemByPropId("10")
        If My.Settings.CheckedDate = "Drawing" Then InvCheckDate = invDrawingDoc.PropertySets.Item("{32853F0F-3444-11D1-9E93-0060B03C1CA6}").ItemByPropId("11")
        If My.Settings.DrawingBy = "Drawing" Then InvDrawDate = invDrawingDoc.PropertySets.Item("{32853F0F-3444-11D1-9E93-0060B03C1CA6}").ItemByPropId("4")
        If My.Settings.DrawingDate = "Drawing" Then InvDrawBy = invDrawingDoc.PropertySets.Item("{32853F0F-3444-11D1-9E93-0060B03C1CA6}").ItemByPropId("41")
        'For A = 1 To CustomPropSet.Count
        '    If InStr(CustomPropSet.Item(A).Name, "Reference") = 0 Then
        '        CustomPropSet.Item(A).Delete()
        '    End If
        'Next
        For P = 0 To 9
            Try
                InvRef(P) = CustomPropSet.Item("Reference" & P)
            Catch
                SetiProp(InvRef(P), "Reference" & P)
                InvRef(P) = CustomPropSet.Item("Reference" & P)
                'If Warning = False Then
                '    MsgBox("Some drawings are missing reference keys, the references will be updated." & vbNewLine &
                '        "The data will be saved, but a new Titleblock is required to display the values." & vbNewLine &
                '        "To fix this update the titleblock through the Drawing Transfer Resource Wizard")
                '    Warning = True
                'End If
            End Try
        Next
        'Try
        '    InvRef(1) = CustomPropSet.Item("Reference1")
        'Catch
        '    SetiProp(InvRef(1), "Reference1")
        'End Try
        'Try
        '    InvRef(2) = CustomPropSet.Item("Reference2")
        'Catch
        '    SetiProp(InvRef(2), "Reference2")
        'End Try
        'Try
        '    InvRef(3) = CustomPropSet.Item("Reference3")
        'Catch
        '    SetiProp(InvRef(3), "Reference3")
        'End Try
        'Try
        '    InvRef(4) = CustomPropSet.Item("Reference4")
        'Catch
        '    SetiProp(InvRef(4), "Reference4")
        'End Try
        'Try
        '    InvRef(5) = CustomPropSet.Item("Reference5")
        'Catch
        '    SetiProp(InvRef(5), "Reference5")
        'End Try
        'Try
        '    InvRef(6) = CustomPropSet.Item("Reference6")
        'Catch
        '    SetiProp(InvRef(6), "Reference6")
        'End Try
        'Try
        '    InvRef(7) = CustomPropSet.Item("Reference7")
        'Catch
        '    SetiProp(InvRef(7), "Reference7")
        'End Try
        'Try
        '    InvRef(8) = CustomPropSet.Item("Reference8")
        'Catch
        '    SetiProp(InvRef(8), "Reference8")
        'End Try
        'Try
        '    InvRef(9) = CustomPropSet.Item("Referenc.9")
        'Catch
        '    SetiProp(InvRef(9), "Reference9")
        'End Try

    End Sub
    Private Sub SetiProp(InvRef As Inventor.Property, Ref As String)
        InvRef = CustomPropSet.Add("", Ref)
    End Sub
    Public Sub RevTableCheck(ByRef oDoc As Document, ByRef RevTotal As Integer)
        Dim DT As DateTime
        Dim J, Rev As Integer
        Dim iProp As Boolean = False
        If Main.chkiProp.CheckState = CheckState.Checked Then iProp = True
        Call RevTable.PopulateRevTable(oDoc, "", J, iProp)
        If RevTable.lstCheckFiles.Items.Count > RevTotal Then
            RevTotal = RevTable.lstCheckFiles.Items.Count
        End If
        For Rev = 1 To J - 1
            Select Case Rev

                Case Is = 1
                    If txtRev1Rev.Text = "" Or UCase(txtRev1Rev.Text) = UCase(RevTable.lstCheckFiles.Items(0).SubItems(0).Text) Then
                        txtRev1Rev.Text = RevTable.lstCheckFiles.Items(0).SubItems(0).Text
                    Else
                        txtRev1Rev.Text = "*Varies*"
                    End If
                    DT = DateTime.Parse(RevTable.lstCheckFiles.Items(0).SubItems(1).Text)
                    If CStr(dtRev1Date.Value) <> "#1/1/1601#" Or dtRev1Date.Value <> DT Then
                        dtRev1Date.Value = DT
                        dtRev1Date.Checked = False
                    Else
                        dtRev1Date.Value = Now
                    End If
                    If txtRev1Des.Text = "" Or UCase(txtRev1Des.Text) = UCase(RevTable.lstCheckFiles.Items(0).SubItems(2).Text) Then
                        txtRev1Des.Text = RevTable.lstCheckFiles.Items(0).SubItems(2).Text
                    Else
                        txtRev1Des.Text = "*Varies*"
                    End If
                    If txtRev1Init.Text = "" Or UCase(txtRev1Init.Text) = UCase(RevTable.lstCheckFiles.Items(0).SubItems(3).Text) Then
                        txtRev1Init.Text = RevTable.lstCheckFiles.Items(0).SubItems(3).Text
                    Else
                        txtRev1Init.Text = "*Varies*"
                    End If
                    If txtRev1Approved.Text = "" Or UCase(txtRev1Approved.Text) = UCase(RevTable.lstCheckFiles.Items(0).SubItems(4).Text) Then
                        txtRev1Approved.Text = RevTable.lstCheckFiles.Items(0).SubItems(4).Text
                    Else
                        txtRev1Approved.Text = "*Varies*"
                    End If
                Case Is = 2
                    If txtRev2Rev.Text = "" Or UCase(txtRev2Rev.Text) = UCase(RevTable.lstCheckFiles.Items(1).SubItems(0).Text) Then
                        txtRev2Rev.Text = RevTable.lstCheckFiles.Items(1).SubItems(0).Text
                    Else
                        txtRev2Rev.Text = "*Varies*"
                    End If
                    DT = DateTime.Parse(RevTable.lstCheckFiles.Items(1).SubItems(1).Text)
                    If CStr(dtRev2Date.Value) <> "#1/1/1601#" Or dtRev2Date.Value = DT Then
                        dtRev2Date.Value = DT
                        dtRev2Date.Checked = False
                    Else
                        dtRev2Date.Value = Now
                    End If
                    If txtRev2Des.Text = "" Or UCase(txtRev2Des.Text) = UCase(RevTable.lstCheckFiles.Items(1).SubItems(2).Text) Then
                        txtRev2Des.Text = RevTable.lstCheckFiles.Items(1).SubItems(2).Text
                    Else
                        txtRev2Des.Text = "*Varies*"
                    End If
                    If txtRev2Init.Text = "" Or UCase(txtRev2Init.Text) = UCase(RevTable.lstCheckFiles.Items(1).SubItems(3).Text) Then
                        txtRev2Init.Text = RevTable.lstCheckFiles.Items(1).SubItems(3).Text
                    Else
                        txtRev2Init.Text = "*Varies*"
                    End If
                    If txtRev2Approved.Text = "" Or UCase(txtRev2Approved.Text) = UCase(RevTable.lstCheckFiles.Items(1).SubItems(4).Text) Then
                        txtRev2Approved.Text = RevTable.lstCheckFiles.Items(1).SubItems(4).Text
                    Else
                        txtRev2Approved.Text = "*Varies*"
                    End If
                Case Is = 3
                    If txtRev3Rev.Text = "" Or UCase(txtRev3Rev.Text) = UCase(RevTable.lstCheckFiles.Items(2).SubItems(0).Text) Then
                        txtRev3Rev.Text = RevTable.lstCheckFiles.Items(2).SubItems(0).Text
                    Else
                        txtRev3Rev.Text = "*Varies*"
                    End If
                    DT = DateTime.Parse(RevTable.lstCheckFiles.Items(2).SubItems(1).Text)
                    If CStr(dtRev3Date.Value) <> "#1/1/1601#" Or dtRev3Date.Value = DT Then
                        dtRev3Date.Value = DT
                        dtRev3Date.Checked = False
                    Else
                        dtRev3Date.Value = Now
                    End If
                    If txtRev3Des.Text = "" Or UCase(txtRev3Des.Text) = UCase(RevTable.lstCheckFiles.Items(2).SubItems(2).Text) Then
                        txtRev3Des.Text = RevTable.lstCheckFiles.Items(2).SubItems(2).Text
                    Else
                        txtRev3Des.Text = "*Varies*"
                    End If
                    If txtRev3Init.Text = "" Or UCase(txtRev3Init.Text) = UCase(RevTable.lstCheckFiles.Items(2).SubItems(3).Text) Then
                        txtRev3Init.Text = RevTable.lstCheckFiles.Items(2).SubItems(3).Text
                    Else
                        txtRev3Init.Text = "*Varies*"
                    End If
                    If txtRev3Approved.Text = "" Or UCase(txtRev3Approved.Text) = UCase(RevTable.lstCheckFiles.Items(2).SubItems(4).Text) Then
                        txtRev3Approved.Text = RevTable.lstCheckFiles.Items(2).SubItems(4).Text
                    Else
                        txtRev3Approved.Text = "*Varies*"
                    End If
                Case Is = 4
                    If txtRev4Rev.Text = "" Or UCase(txtRev4Rev.Text) = UCase(RevTable.lstCheckFiles.Items(3).SubItems(0).Text) Then
                        txtRev4Rev.Text = RevTable.lstCheckFiles.Items(3).SubItems(0).Text
                    Else
                        txtRev4Rev.Text = "*Varies*"
                    End If
                    DT = DateTime.Parse(RevTable.lstCheckFiles.Items(3).SubItems(1).Text)
                    If CStr(dtRev4Date.Value) <> "#1/1/1601#" Or dtRev4Date.Value = DT Then
                        dtRev4Date.Value = DT
                        dtRev4Date.Checked = False
                    Else
                        dtRev4Date.Value = Now
                    End If
                    If txtRev4Des.Text = "" Or UCase(txtRev4Des.Text) = UCase(RevTable.lstCheckFiles.Items(3).SubItems(2).Text) Then
                        txtRev4Des.Text = RevTable.lstCheckFiles.Items(3).SubItems(2).Text
                    Else
                        txtRev4Des.Text = "*Varies*"
                    End If
                    If txtRev4Init.Text = "" Or UCase(txtRev4Init.Text) = UCase(RevTable.lstCheckFiles.Items(3).SubItems(3).Text) Then
                        txtRev4Init.Text = RevTable.lstCheckFiles.Items(3).SubItems(3).Text
                    Else
                        txtRev4Init.Text = "*Varies*"
                    End If
                    If txtRev4Approved.Text = "" Or UCase(txtRev4Approved.Text) = UCase(RevTable.lstCheckFiles.Items(3).SubItems(4).Text) Then
                        txtRev4Approved.Text = RevTable.lstCheckFiles.Items(3).SubItems(4).Text
                    Else
                        txtRev4Approved.Text = "*Varies*"
                    End If
                Case Is = 5
                    If txtRev5Rev.Text = "" Or UCase(txtRev5Rev.Text) = UCase(RevTable.lstCheckFiles.Items(4).SubItems(0).Text) Then
                        txtRev5Rev.Text = RevTable.lstCheckFiles.Items(4).SubItems(0).Text
                    Else
                        txtRev5Rev.Text = "*Varies*"
                    End If
                    DT = DateTime.Parse(RevTable.lstCheckFiles.Items(4).SubItems(1).Text)
                    If CStr(dtRev5Date.Value) <> "#1/1/1601#" Or dtRev5Date.Value = DT Then
                        dtRev5Date.Value = DT
                        dtRev5Date.Checked = False
                    Else
                        dtRev5Date.Value = Now
                    End If
                    If txtRev5Des.Text = "" Or UCase(txtRev5Des.Text) = UCase(RevTable.lstCheckFiles.Items(4).SubItems(2).Text) Then
                        txtRev5Des.Text = RevTable.lstCheckFiles.Items(4).SubItems(2).Text
                    Else
                        txtRev5Des.Text = "*Varies*"
                    End If
                    If txtRev5Init.Text = "" Or UCase(txtRev5Init.Text) = UCase(RevTable.lstCheckFiles.Items(4).SubItems(3).Text) Then
                        txtRev5Init.Text = RevTable.lstCheckFiles.Items(4).SubItems(3).Text
                    Else
                        txtRev5Init.Text = "*Varies*"
                    End If
                    If txtRev5Approved.Text = "" Or UCase(txtRev5Approved.Text) = UCase(RevTable.lstCheckFiles.Items(4).SubItems(4).Text) Then
                        txtRev5Approved.Text = RevTable.lstCheckFiles.Items(4).SubItems(4).Text
                    Else
                        txtRev5Approved.Text = "*Varies*"
                    End If
                Case Is = 6
                    If txtRev6Rev.Text = "" Or UCase(txtRev6Rev.Text) = UCase(RevTable.lstCheckFiles.Items(5).SubItems(0).Text) Then
                        txtRev6Rev.Text = RevTable.lstCheckFiles.Items(5).SubItems(0).Text
                    Else
                        txtRev6Rev.Text = "*Varies*"
                    End If
                    DT = DateTime.Parse(RevTable.lstCheckFiles.Items(5).SubItems(1).Text)
                    If CStr(dtRev6Date.Value) <> "#1/1/1601#" Or dtRev6Date.Value = DT Then
                        dtRev6Date.Value = DT
                        dtRev6Date.Checked = False
                    Else
                        dtRev6Date.Value = Now
                    End If
                    If txtRev6Des.Text = "" Or UCase(txtRev6Des.Text) = UCase(RevTable.lstCheckFiles.Items(5).SubItems(2).Text) Then
                        txtRev6Des.Text = RevTable.lstCheckFiles.Items(5).SubItems(2).Text
                    Else
                        txtRev6Des.Text = "*Varies*"
                    End If
                    If txtRev6Init.Text = "" Or UCase(txtRev6Init.Text) = UCase(RevTable.lstCheckFiles.Items(5).SubItems(3).Text) Then
                        txtRev6Init.Text = RevTable.lstCheckFiles.Items(5).SubItems(3).Text
                    Else
                        txtRev6Init.Text = "*Varies*"
                    End If
                    If txtRev6Approved.Text = "" Or UCase(txtRev6Approved.Text) = UCase(RevTable.lstCheckFiles.Items(5).SubItems(4).Text) Then
                        txtRev6Approved.Text = RevTable.lstCheckFiles.Items(5).SubItems(4).Text
                    Else
                        txtRev6Approved.Text = "*Varies*"
                    End If
                Case Is = 7
                    If txtRev7Rev.Text = "" Or UCase(txtRev7Rev.Text) = UCase(RevTable.lstCheckFiles.Items(6).SubItems(0).Text) Then
                        txtRev7Rev.Text = RevTable.lstCheckFiles.Items(6).SubItems(0).Text
                    Else
                        txtRev7Rev.Text = "*Varies*"
                    End If
                    DT = DateTime.Parse(RevTable.lstCheckFiles.Items(6).SubItems(1).Text)
                    If CStr(dtRev7Date.Value) <> "#1/1/1601#" Or dtRev7Date.Value = DT Then
                        dtRev7Date.Value = DT
                        dtRev7Date.Checked = False
                    Else
                        dtRev7Date.Value = Now
                    End If
                    If txtRev7Des.Text = "" Or UCase(txtRev7Des.Text) = UCase(RevTable.lstCheckFiles.Items(6).SubItems(2).Text) Then
                        txtRev7Des.Text = RevTable.lstCheckFiles.Items(6).SubItems(2).Text
                    Else
                        txtRev7Des.Text = "*Varies*"
                    End If
                    If txtRev7Init.Text = "" Or UCase(txtRev7Init.Text) = UCase(RevTable.lstCheckFiles.Items(6).SubItems(3).Text) Then
                        txtRev7Init.Text = RevTable.lstCheckFiles.Items(6).SubItems(3).Text
                    Else
                        txtRev7Init.Text = "*Varies*"
                    End If
                    If txtRev7Approved.Text = "" Or UCase(txtRev7Approved.Text) = UCase(RevTable.lstCheckFiles.Items(6).SubItems(4).Text) Then
                        txtRev7Approved.Text = RevTable.lstCheckFiles.Items(6).SubItems(4).Text
                    Else
                        txtRev7Approved.Text = "*Varies*"
                    End If
                Case Is = 8
                    If txtRev8Rev.Text = "" Or UCase(txtRev8Rev.Text) = UCase(RevTable.lstCheckFiles.Items(7).SubItems(0).Text) Then
                        txtRev8Rev.Text = RevTable.lstCheckFiles.Items(7).SubItems(0).Text
                    Else
                        txtRev8Rev.Text = "*Varies*"
                    End If
                    DT = DateTime.Parse(RevTable.lstCheckFiles.Items(7).SubItems(1).Text)
                    If CStr(dtRev8Date.Value) <> "#1/1/1601#" Or dtRev8Date.Value = DT Then
                        dtRev8Date.Value = DT
                        dtRev8Date.Checked = False
                    Else
                        dtRev8Date.Value = Now
                    End If
                    If txtRev8Des.Text = "" Or UCase(txtRev8Des.Text) = UCase(RevTable.lstCheckFiles.Items(7).SubItems(2).Text) Then
                        txtRev8Des.Text = RevTable.lstCheckFiles.Items(7).SubItems(2).Text
                    Else
                        txtRev8Des.Text = "*Varies*"
                    End If
                    If txtRev8Init.Text = "" Or UCase(txtRev8Init.Text) = UCase(RevTable.lstCheckFiles.Items(7).SubItems(3).Text) Then
                        txtRev8Init.Text = RevTable.lstCheckFiles.Items(7).SubItems(3).Text
                    Else
                        txtRev8Init.Text = "*Varies*"
                    End If
                    If txtRev8Approved.Text = "" Or UCase(txtRev8Approved.Text) = UCase(RevTable.lstCheckFiles.Items(7).SubItems(4).Text) Then
                        txtRev8Approved.Text = RevTable.lstCheckFiles.Items(7).SubItems(4).Text
                    Else
                        txtRev8Approved.Text = "*Varies*"
                    End If
            End Select
        Next
    End Sub
    Private Sub InitializeiProperties(ByRef X As Integer, DrawSource As String)
        'Go through each iProperty in both the part and drawing and assign them to their respective output strings
        'When the values between documents are different, display *Varies* as the output.
        Dim TempState As String = ""
        If X = 0 And txtTitle.Text = "" Then
            txtTitle.Text = CStr(InvTitle.Value)
        ElseIf X <> 0 And UCase(txtTitle.Text) <> UCase(CStr(InvTitle.Value)) Then
            txtTitle.Text = "*Varies*"
        End If

        If X = 0 And txtSubject.Text = "" Then
            txtSubject.Text = CStr(InvSubject.Value)
        ElseIf X <> 0 And UCase(txtSubject.Text) <> UCase(CStr(InvSubject.Value)) Then
            txtSubject.Text = "*Varies*"
        End If
        If X = 0 And txtAuthor.Text = "" Then
            txtAuthor.Text = CStr(InvAuthor.Value)
        ElseIf X <> 0 And UCase(txtAuthor.Text) <> UCase(CStr(InvAuthor.Value)) Then
            txtAuthor.Text = "*Varies*"
        End If

        If X = 0 And txtManager.Text = "" Then
            txtManager.Text = CStr(InvManager.Value)
        ElseIf X <> 0 And UCase(txtManager.Text) <> UCase(CStr(InvManager.Value)) Then
            txtManager.Text = "*Varies*"
        End If

        If X = 0 And txtCompany.Text = "" Then
            txtCompany.Text = CStr(InvCompany.Value)
        ElseIf X <> 0 And UCase(txtCompany.Text) <> UCase(CStr(InvCompany.Value)) Then
            txtCompany.Text = "*Varies*"
        End If

        If X = 0 And txtCategory.Text = "" Then
            txtCategory.Text = CStr(InvCategory.Value)
        ElseIf X <> 0 And UCase(txtCategory.Text) <> UCase(CStr(InvCategory.Value)) Then
            txtCategory.Text = "*Varies*"
        End If

        If X = 0 And txtKeywords.Text = "" Then
            txtKeywords.Text = CStr(InvKeywords.Value)
        ElseIf X <> 0 And UCase(txtKeywords.Text) <> UCase(CStr(InvKeywords.Value)) Then
            txtKeywords.Text = "*Varies*"
        End If

        If X = 0 And txtComments.Text = "" Then
            txtComments.Text = CStr(Invcomments.Value)
        ElseIf X <> 0 And UCase(txtComments.Text) <> UCase(CStr(Invcomments.Value)) Then
            txtComments.Text = "*Varies*"
        End If

        If X = 0 And txtLocation.Text = "" Then
            txtLocation.Text = DrawSource
        ElseIf X <> 0 And UCase(txtLocation.Text) <> UCase(DrawSource) Then
            txtLocation.Text = "*Varies*"
        End If

        If X = 0 And txtPNumber.Text = "" Then
            txtPNumber.Text = CStr(InvPartNumber.Value)
        ElseIf X <> 0 And UCase(txtPNumber.Text) <> UCase(CStr(InvPartNumber.Value)) Then
            txtPNumber.Text = "*Varies*"
        End If

        If X = 0 And txtDescription.Text = "" Then
            txtDescription.Text = CStr(InvDescription.Value)
        ElseIf X <> 0 And UCase(txtDescription.Text) <> UCase(CStr(InvDescription.Value)) Then
            txtDescription.Text = "*Varies*"
        End If

        If X = 0 And txtRevision.Text = "" Then
            txtRevision.Text = CStr(InvRevision.Value)
        ElseIf X <> 0 And UCase(txtRevision.Text) <> UCase(CStr(InvRevision.Value)) Then
            txtRevision.Text = "*Varies*"
        End If

        If X = 0 And txtProject.Text = "" Then
            txtProject.Text = CStr(InvProject.Value)
        ElseIf X <> 0 And UCase(txtProject.Text) <> UCase(CStr(InvProject.Value)) Then
            txtProject.Text = "*Varies*"
        End If

        If X = 0 And txtDesigner.Text = "" Then
            txtDesigner.Text = CStr(InvDesigner.Value)
        ElseIf X <> 0 And UCase(txtDesigner.Text) <> UCase(CStr(InvDesigner.Value)) Then
            txtDesigner.Text = "*Varies*"
        End If

        If X = 0 And txtPartNumber.Text = "" Then
            txtPartNumber.Text = CStr(InvPartNumber.Value)
        ElseIf X <> 0 And UCase(txtPartNumber.Text) <> UCase(CStr(InvPartNumber.Value)) Then
            txtPartNumber.Text = "*Varies*"
        End If

        If X = 0 And txtSNumber.Text = "" Then
            txtSNumber.Text = CStr(InvStockNumber.Value)
        ElseIf X <> 0 And UCase(txtSNumber.Text) <> UCase(CStr(InvStockNumber.Value)) Then
            txtSNumber.Text = "*Varies*"
        End If

        If X = 0 And txtEngineer.Text = "" Then
            txtEngineer.Text = CStr(InvEngineer.Value)
        ElseIf X <> 0 And UCase(txtEngineer.Text) <> UCase(CStr(InvEngineer.Value)) Then
            txtEngineer.Text = "*Varies*"
        End If

        If X = 0 And txtVendor.Text = "" Then
            txtVendor.Text = CStr(InvVendor.Value)
        ElseIf X <> 0 And UCase(txtVendor.Text) <> UCase(CStr(InvVendor.Value)) Then
            txtVendor.Text = "*Varies*"
        End If

        If CStr(InvModDate.Value) = DateTime.MinValue Then
            If X = 0 Then
                dtModelDate.Checked = True
                dtModelDate.Value = InvModDate.Value
            ElseIf X <> 0 And dtModelDate.Value <> InvModDate.Value Then
                dtModelDate.Checked = False
            End If
        Else
            dtModelDate.Checked = False
        End If

        If CStr(InvDrawDate.Value) = DateTime.MinValue Then
            If X = 0 Then
                dtDrawingDate.Checked = True
                dtDrawingDate.Value = InvDrawDate.Value
            ElseIf X <> 0 And dtDrawingDate.Value <> InvDrawDate.Value Then
                dtDrawingDate.Checked = False
            End If
        Else
            dtDrawingDate.Checked = False
        End If

        If X = 0 And txtStockNumber.Text = "" Then
            txtStockNumber.Text = CStr(InvStockNumber.Value)
        ElseIf X <> 0 And UCase(txtStockNumber.Text) <> UCase(CStr(InvStockNumber.Value)) Then
            txtStockNumber.Text = "*Varies*"
        End If

        If X = 0 And txtStatus.Text = "" Then
            txtStatus.Text = CStr(InvStatus.Value)
        ElseIf X <> 0 And UCase(txtStatus.Text) <> UCase(CStr(InvStatus.Value)) Then
            txtStatus.Text = "*Varies*"
        End If

        If CStr(InvModDate.Value) = DateTime.MinValue Then
            If X = 0 Then
                dtModelDate.Checked = True
                dtModelDate.Value = InvModDate.Value
            ElseIf X <> 0 And dtModelDate.Value <> InvModDate.Value Then
                dtModelDate.Checked = False
            End If
        Else
            dtModelDate.Checked = False
        End If

        If X = 0 And txtDrawing.Text = "" Then
            txtDrawing.Text = CStr(InvDrawBy.Value)
        ElseIf X <> 0 And UCase(txtDrawing.Text) <> UCase(CStr(InvDrawBy.Value)) Then
            txtDrawing.Text = "*Varies*"
        End If

        If X = 0 Then
            If InvDesignState = 3 Then
                cmbDesignState.Text = "Released"
            ElseIf InvDesignState = 2 Then
                cmbDesignState.Text = "Pending"
            ElseIf InvDesignState = 1 Then
                cmbDesignState.Text = "Work In Progress"
            End If
        Else
            If cmbDesignState.Text = "Released" Then
                TempState = 3
            ElseIf cmbDesignState.Text = "Pending" Then
                TempState = 2
            ElseIf cmbDesignState.Text = "Work In Progress" Then
                TempState = 1
            Else
                TempState = "*Varies*"
            End If
        End If
        If X <> 0 And TempState <> InvDesignState Then
            cmbDesignState.Text = "*Varies*"
        End If

        If X = 0 And txtCheckedBy.Text = "" Then
            txtCheckedBy.Text = CStr(InvCheckedBy.Value)
        ElseIf X <> 0 And UCase(txtCheckedBy.Text) <> UCase(CStr(InvCheckedBy.Value)) Then
            txtCheckedBy.Text = "*Varies*"
        End If
        If CStr(InvCheckDate.Value) = DateTime.MinValue Then
            If X = 0 Then
                dtCheckDate.Value = InvCheckDate.Value
                dtCheckDate.Checked = True
            ElseIf X <> 0 And dtCheckDate.Value <> InvCheckDate.Value Then
                dtCheckDate.Checked = False
            End If
        Else
            dtCheckDate.Checked = False
        End If
        If X = 0 And txtRef0.Text = "" Then
            Me.txtRef0.Text = CStr(InvRef(0).Value)
        ElseIf X <> 0 And UCase(txtRef0.Text) <> UCase(CStr(InvRef(0).Value)) Then
            Me.txtRef0.Text = "*Varies*"
        End If
        If X = 0 And txtRef1.Text = "" Then
            Me.txtRef1.Text = CStr(InvRef(1).Value)
        ElseIf X <> 0 And UCase(txtRef1.Text) <> UCase(CStr(InvRef(1).Value)) Then
            Me.txtRef1.Text = "*Varies*"
        End If
        If X = 0 And txtRef2.Text = "" Then
            Me.txtRef2.Text = CStr(InvRef(2).Value)
        ElseIf X <> 0 And UCase(txtRef2.Text) <> UCase(CStr(InvRef(2).Value)) Then
            Me.txtRef2.Text = "*Varies*"
        End If
        If X = 0 And txtRef3.Text = "" Then
            Me.txtRef2.Text = CStr(InvRef(3).Value)
        ElseIf X <> 0 And UCase(txtRef3.Text) <> UCase(CStr(InvRef(3).Value)) Then
            Me.txtRef3.Text = "*Varies*"
        End If
        If X = 0 And txtRef4.Text = "" Then
            Me.txtRef4.Text = CStr(InvRef(4).Value)
        ElseIf X <> 0 And UCase(txtRef4.Text) <> UCase(CStr(InvRef(4).Value)) Then
            Me.txtRef4.Text = "*Varies*"
        End If
        If X = 0 And txtRef5.Text = "" Then
            Me.txtRef5.Text = CStr(InvRef(5).Value)
        ElseIf X <> 0 And UCase(txtRef5.Text) <> UCase(CStr(InvRef(5).Value)) Then
            Me.txtRef5.Text = "*Varies*"
        End If
        If X = 0 And txtRef6.Text = "" Then
            Me.txtRef6.Text = CStr(InvRef(6).Value)
        ElseIf X <> 0 And UCase(txtRef6.Text) <> UCase(CStr(InvRef(6).Value)) Then
            Me.txtRef6.Text = "*Varies*"
        End If
        If X = 0 And txtRef7.Text = "" Then
            Me.txtRef7.Text = CStr(InvRef(7).Value)
        ElseIf X <> 0 And UCase(txtRef7.Text) <> UCase(CStr(InvRef(7).Value)) Then
            Me.txtRef7.Text = "*Varies*"
        End If
        If X = 0 And txtRef8.Text = "" Then
            Me.txtRef8.Text = CStr(InvRef(8).Value)
        ElseIf X <> 0 And UCase(txtRef8.Text) <> UCase(CStr(InvRef(8).Value)) Then
            Me.txtRef8.Text = "*Varies*"
        End If
        If X = 0 And txtRef9.Text = "" Then
            Me.txtRef9.Text = CStr(InvRef(9).Value)
        ElseIf X <> 0 And UCase(txtRef9.Text) <> UCase(CStr(InvRef(9).Value)) Then
            Me.txtRef9.Text = "*Varies*"
        End If
        'Increase counter to require comparison to the last ipropertiy collected
        ' X = X + 1

    End Sub
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
    Private Sub WriteModelProps(ByRef oDoc As Document)
        If txtTitle.Text <> "*Varies*" And My.Settings.Title = "Model" Then
            oDoc.PropertySets.Item("{F29F85E0-4FF9-1068-AB91-08002B27B3D9}").ItemByPropId("2").Value = txtTitle.Text
        End If
        If txtSubject.Text <> "*Varies*" And My.Settings.Subject = "Model" Then
            oDoc.PropertySets.Item("{F29F85E0-4FF9-1068-AB91-08002B27B3D9}").ItemByPropId("3").Value = txtSubject.Text
        End If
        If txtManager.Text <> "*Varies*" And My.Settings.Manager = "Model" Then
            oDoc.PropertySets.Item("{D5CDD502-2E9C-101B-9397-08002B2CF9AE}").ItemByPropId("14").Value = txtManager.Text
        End If
        If txtCategory.Text <> "*Varies*" And My.Settings.Category = "Model" Then
            oDoc.PropertySets.Item("{D5CDD502-2E9C-101B-9397-08002B2CF9AE}").ItemByPropId("2").Value = txtCategory.Text
        End If
        If txtKeywords.Text <> "*Varies*" And My.Settings.Keywords = "Model" Then
            oDoc.PropertySets.Item("{F29F85E0-4FF9-1068-AB91-08002B27B3D9}").ItemByPropId("5").Value = txtKeywords.Text
        End If
        If txtAuthor.Text <> "*Varies*" And My.Settings.Author = "Model" Then
            oDoc.PropertySets.Item("{F29F85E0-4FF9-1068-AB91-08002B27B3D9}").ItemByPropId("4").Value = txtAuthor.Text
        End If
        If txtCompany.Text <> "*Varies*" And My.Settings.Company = "Model" Then
            oDoc.PropertySets.Item("{D5CDD502-2E9C-101B-9397-08002B2CF9AE}").ItemByPropId("15").Value = txtCompany.Text
        End If
        If txtComments.Text <> "*Varies*" And My.Settings.Comments = "Model" Then
            oDoc.PropertySets.Item("{F29F85E0-4FF9-1068-AB91-08002B27B3D9}").ItemByPropId("6").Value = txtComments.Text
        End If
        If txtDescription.Text <> "*Varies*" And My.Settings.Description = "Model" Then
            oDoc.PropertySets.Item("{32853F0F-3444-11D1-9E93-0060B03C1CA6}").ItemByPropId("29").Value = txtDescription.Text
        End If
        If txtPartNumber.Text <> "*Varies*" And My.Settings.PPartNumber = "Model" Then
            oDoc.PropertySets.Item("{32853F0F-3444-11D1-9E93-0060B03C1CA6}").ItemByPropId("5").Value = txtPartNumber.Text
        End If
        If txtStockNumber.Text <> "*Varies*" And My.Settings.PStockNumber = "Model" Then
            oDoc.PropertySets.Item("{32853F0F-3444-11D1-9E93-0060B03C1CA6}").ItemByPropId("55").Value = txtStockNumber.Text
        End If
        If txtProject.Text <> "*Varies*" And My.Settings.Project = "Model" Then
            oDoc.PropertySets.Item("{32853F0F-3444-11D1-9E93-0060B03C1CA6}").ItemByPropId("7").Value = txtProject.Text
        End If
        If txtDesigner.Text <> "*Varies*" And My.Settings.Designer = "Model" Then
            oDoc.PropertySets.Item("{32853F0F-3444-11D1-9E93-0060B03C1CA6}").ItemByPropId("41").Value = txtDesigner.Text
        End If
        If txtVendor.Text <> "*Varies*" And My.Settings.Vendor = "Model" Then
            oDoc.PropertySets.Item("{32853F0F-3444-11D1-9E93-0060B03C1CA6}").ItemByPropId("30").Value = txtVendor.Text
        End If
        If dtModelDate.Checked = True And My.Settings.ModelDate = "Model" Then
            oDoc.PropertySets.Item("{32853F0F-3444-11D1-9E93-0060B03C1CA6}").ItemByPropId("4").Value = dtModelDate.Value
        End If
        If txtEngineer.Text <> "*Varies*" And My.Settings.Engineer = "Model" Then
            oDoc.PropertySets.Item("{32853F0F-3444-11D1-9E93-0060B03C1CA6}").ItemByPropId("42").Value = txtEngineer.Text
        End If
        If txtStatus.Text <> "*Varies*" And My.Settings.Status = "Model" Then
            oDoc.PropertySets.Item("{32853F0F-3444-11D1-9E93-0060B03C1CA6}").ItemByPropId("17").Value = txtStatus.Text
        End If
        If txtRevision.Text <> "*Varies*" And My.Settings.Revision = "Model" Then
            oDoc.PropertySets.Item("{F29F85E0-4FF9-1068-AB91-08002B27B3D9}").ItemByPropId("9").Value = txtRevision.Text
        End If
        If txtDrawing.Text <> "*Varies*" And My.Settings.DrawingBy = "Model" Then
            oDoc.PropertySets.Item("{32853F0F-3444-11D1-9E93-0060B03C1CA6}").ItemByPropId("41").Value = txtDrawing.Text
        End If
        If cmbDesignState.Text = "Released" And My.Settings.DesignState = "Model" Then
            oDoc.PropertySets.Item("{32853F0F-3444-11D1-9E93-0060B03C1CA6}").ItemByPropId("40").Value = 3
        ElseIf cmbDesignState.Text = "Pending" And My.Settings.DesignState = "Model" Then
            oDoc.PropertySets.Item("{32853F0F-3444-11D1-9E93-0060B03C1CA6}").ItemByPropId("40").Value = 2
        ElseIf cmbDesignState.Text = "Work In Progress" And My.Settings.DesignState = "Model" Then
            oDoc.PropertySets.Item("{32853F0F-3444-11D1-9E93-0060B03C1CA6}").ItemByPropId("40").Value = 1
        End If
        If txtCheckedBy.Text <> "*Varies*" And My.Settings.CheckedBy = "Model" Then
            oDoc.PropertySets.Item("{32853F0F-3444-11D1-9E93-0060B03C1CA6}").ItemByPropId("10").Value = txtCheckedBy.Text
        End If
        If dtCheckDate.Checked = True And My.Settings.CheckedDate = "Model" Then
            oDoc.PropertySets.Item("{32853F0F-3444-11D1-9E93-0060B03C1CA6}").ItemByPropId("11").Value = dtCheckDate.Value
        Else
            ' oDoc.PropertySets.Item("{32853F0F-3444-11D1-9E93-0060B03C1CA6}").ItemByPropId("11").Expression = "1601, 1, 1"
        End If
        If dtDrawingDate.Checked = True And My.Settings.DrawingDate = "Model" Then
            oDoc.PropertySets.Item("{32853F0F-3444-11D1-9E93-0060B03C1CA6}").ItemByPropId("4").Value = dtDrawingDate.Value
        Else
            ' oDoc.PropertySets.Item("{32853F0F-3444-11D1-9E93-0060B03C1CA6}").ItemByPropId("4").Expression = "1601, 1, 1"
        End If
        If txtRef0.Text <> "*Varies*" And My.Settings.Custom = "Model" Then
            oDoc.PropertySets.Item("Inventor User Defined Properties").Item("Reference0").Value = txtRef0.Text
        End If
        If txtRef1.Text <> "*Varies*" And My.Settings.Custom = "Model" Then
            oDoc.PropertySets.Item("Inventor User Defined Properties").Item("Reference1").Value = txtRef1.Text
        End If
        If txtRef2.Text <> "*Varies*" And My.Settings.Custom = "Model" Then
            oDoc.PropertySets.Item("Inventor User Defined Properties").Item("Reference2").Value = txtRef2.Text
        End If
        If txtRef3.Text <> "*Varies*" And My.Settings.Custom = "Model" Then
            oDoc.PropertySets.Item("Inventor User Defined Properties").Item("Reference3").Value = txtRef3.Text
        End If
        If txtRef4.Text <> "*Varies*" And My.Settings.Custom = "Model" Then
            oDoc.PropertySets.Item("Inventor User Defined Properties").Item("Reference4").Value = txtRef4.Text
        End If
        If txtRef5.Text <> "*Varies*" And My.Settings.Custom = "Model" Then
            oDoc.PropertySets.Item("Inventor User Defined Properties").Item("Reference5").Value = txtRef5.Text
        End If
        If txtRef6.Text <> "*Varies*" And My.Settings.Custom = "Model" Then
            oDoc.PropertySets.Item("Inventor User Defined Properties").Item("Reference6").Value = txtRef6.Text
        End If
        If txtRef7.Text <> "*Varies*" And My.Settings.Custom = "Model" Then
            oDoc.PropertySets.Item("Inventor User Defined Properties").Item("Reference7").Value = txtRef7.Text
        End If
        If txtRef8.Text <> "*Varies*" And My.Settings.Custom = "Model" Then
            oDoc.PropertySets.Item("Inventor User Defined Properties").Item("Reference8").Value = txtRef8.Text
        End If
        If txtRef9.Text <> "*Varies*" And My.Settings.Custom = "Model" Then
            oDoc.PropertySets.Item("Inventor User Defined Properties").Item("Reference9").Value = txtRef9.Text
        End If
    End Sub
    Private Sub WriteDrawingProps(ByRef oDoc As Document)
        If txtTitle.Text <> "*Varies*" And My.Settings.Title = "Drawing" Then
            oDoc.PropertySets.Item("{F29F85E0-4FF9-1068-AB91-08002B27B3D9}").ItemByPropId("2").Value = txtTitle.Text
        End If
        If txtSubject.Text <> "*Varies*" And My.Settings.Subject = "Drawing" Then
            oDoc.PropertySets.Item("{F29F85E0-4FF9-1068-AB91-08002B27B3D9}").ItemByPropId("3").Value = txtSubject.Text
        End If
        If txtManager.Text <> "*Varies*" And My.Settings.Manager = "Drawing" Then
            oDoc.PropertySets.Item("{D5CDD502-2E9C-101B-9397-08002B2CF9AE}").ItemByPropId("14").Value = txtManager.Text
        End If
        If txtCategory.Text <> "*Varies*" And My.Settings.Category = "Drawing" Then
            oDoc.PropertySets.Item("{D5CDD502-2E9C-101B-9397-08002B2CF9AE}").ItemByPropId("2").Value = txtCategory.Text
        End If
        If txtKeywords.Text <> "*Varies*" And My.Settings.Keywords = "Drawing" Then
            oDoc.PropertySets.Item("{F29F85E0-4FF9-1068-AB91-08002B27B3D9}").ItemByPropId("5").Value = txtKeywords.Text
        End If
        If txtAuthor.Text <> "*Varies*" And My.Settings.Author = "Drawing" Then
            oDoc.PropertySets.Item("{F29F85E0-4FF9-1068-AB91-08002B27B3D9}").ItemByPropId("4").Value = txtAuthor.Text
        End If
        If txtCompany.Text <> "*Varies*" And My.Settings.Company = "Drawing" Then
            oDoc.PropertySets.Item("{D5CDD502-2E9C-101B-9397-08002B2CF9AE}").ItemByPropId("15").Value = txtCompany.Text
        End If
        If txtComments.Text <> "*Varies*" And My.Settings.Comments = "Drawing" Then
            oDoc.PropertySets.Item("{F29F85E0-4FF9-1068-AB91-08002B27B3D9}").ItemByPropId("6").Value = txtComments.Text
        End If
        If txtDescription.Text <> "*Varies*" And My.Settings.Description = "Drawing" Then
            oDoc.PropertySets.Item("{32853F0F-3444-11D1-9E93-0060B03C1CA6}").ItemByPropId("29").Value = txtDescription.Text
        End If
        If txtPartNumber.Text <> "*Varies*" And My.Settings.PPartNumber = "Drawing" Then
            oDoc.PropertySets.Item("{32853F0F-3444-11D1-9E93-0060B03C1CA6}").ItemByPropId("5").Value = txtPartNumber.Text
        End If
        If txtStockNumber.Text <> "*Varies*" And My.Settings.PStockNumber = "Drawing" Then
            oDoc.PropertySets.Item("{32853F0F-3444-11D1-9E93-0060B03C1CA6}").ItemByPropId("55").Value = txtStockNumber.Text
        End If
        If txtProject.Text <> "*Varies*" And My.Settings.Project = "Drawing" Then
            oDoc.PropertySets.Item("{32853F0F-3444-11D1-9E93-0060B03C1CA6}").ItemByPropId("7").Value = txtProject.Text
        End If
        If txtDesigner.Text <> "*Varies*" And My.Settings.Designer = "Drawing" Then
            oDoc.PropertySets.Item("{32853F0F-3444-11D1-9E93-0060B03C1CA6}").ItemByPropId("41").Value = txtDesigner.Text
        End If
        If txtVendor.Text <> "*Varies*" And My.Settings.Vendor = "Drawing" Then
            oDoc.PropertySets.Item("{32853F0F-3444-11D1-9E93-0060B03C1CA6}").ItemByPropId("30").Value = txtVendor.Text
        End If
        If dtModelDate.Checked = True And My.Settings.ModelDate = "Drawing" Then
            oDoc.PropertySets.Item("{32853F0F-3444-11D1-9E93-0060B03C1CA6}").ItemByPropId("4").Value = dtModelDate.Value
        End If
        If txtEngineer.Text <> "*Varies*" And My.Settings.Engineer = "Drawing" Then
            oDoc.PropertySets.Item("{32853F0F-3444-11D1-9E93-0060B03C1CA6}").ItemByPropId("42").Value = txtEngineer.Text
        End If
        If txtStatus.Text <> "*Varies*" And My.Settings.Status = "Drawing" Then
            oDoc.PropertySets.Item("{32853F0F-3444-11D1-9E93-0060B03C1CA6}").ItemByPropId("17").Value = txtStatus.Text
        End If
        If txtRevision.Text <> "*Varies*" And My.Settings.Revision = "Drawing" Then
            oDoc.PropertySets.Item("{F29F85E0-4FF9-1068-AB91-08002B27B3D9}").ItemByPropId("9").Value = txtRevision.Text
        End If
        If txtDrawing.Text <> "*Varies*" And My.Settings.DrawingBy = "Drawing" Then
            oDoc.PropertySets.Item("{32853F0F-3444-11D1-9E93-0060B03C1CA6}").ItemByPropId("41").Value = txtDrawing.Text
        End If
        If cmbDesignState.Text = "Released" And My.Settings.DesignState = "Drawing" Then
            oDoc.PropertySets.Item("{32853F0F-3444-11D1-9E93-0060B03C1CA6}").ItemByPropId("40").Value = 3
        ElseIf cmbDesignState.Text = "Pending" And My.Settings.DesignState = "Drawing" Then
            oDoc.PropertySets.Item("{32853F0F-3444-11D1-9E93-0060B03C1CA6}").ItemByPropId("40").Value = 2
        ElseIf cmbDesignState.Text = "Work In Progress" And My.Settings.DesignState = "Drawing" Then
            oDoc.PropertySets.Item("{32853F0F-3444-11D1-9E93-0060B03C1CA6}").ItemByPropId("40").Value = 1
        End If
        If txtCheckedBy.Text <> "*Varies*" And My.Settings.CheckedBy = "Drawing" Then
            oDoc.PropertySets.Item("{32853F0F-3444-11D1-9E93-0060B03C1CA6}").ItemByPropId("10").Value = txtCheckedBy.Text
        End If
        If dtCheckDate.Checked = True And My.Settings.CheckedDate = "Drawing" Then
            oDoc.PropertySets.Item("{32853F0F-3444-11D1-9E93-0060B03C1CA6}").ItemByPropId("11").Value = dtCheckDate.Value
        Else
            ' oDoc.PropertySets.Item("{32853F0F-3444-11D1-9E93-0060B03C1CA6}").ItemByPropId("11").Expression = "1601, 1, 1"
        End If
        If dtDrawingDate.Checked = True And My.Settings.DrawingDate = "Drawing" Then
            oDoc.PropertySets.Item("{32853F0F-3444-11D1-9E93-0060B03C1CA6}").ItemByPropId("4").Value = dtDrawingDate.Value
        Else
            ' oDoc.PropertySets.Item("{32853F0F-3444-11D1-9E93-0060B03C1CA6}").ItemByPropId("4").Expression = "1601, 1, 1"
        End If
        If txtRef0.Text <> "*Varies*" And My.Settings.Custom = "Drawing" Then
            oDoc.PropertySets.Item("Inventor User Defined Properties").Item("Reference0").Value = txtRef0.Text
        End If
        If txtRef1.Text <> "*Varies*" And My.Settings.Custom = "Drawing" Then
            oDoc.PropertySets.Item("Inventor User Defined Properties").Item("Reference1").Value = txtRef1.Text
        End If
        If txtRef2.Text <> "*Varies*" And My.Settings.Custom = "Drawing" Then
            oDoc.PropertySets.Item("Inventor User Defined Properties").Item("Reference2").Value = txtRef2.Text
        End If
        If txtRef3.Text <> "*Varies*" And My.Settings.Custom = "Drawing" Then
            oDoc.PropertySets.Item("Inventor User Defined Properties").Item("Reference3").Value = txtRef3.Text
        End If
        If txtRef4.Text <> "*Varies*" And My.Settings.Custom = "Drawing" Then
            oDoc.PropertySets.Item("Inventor User Defined Properties").Item("Reference4").Value = txtRef4.Text
        End If
        If txtRef5.Text <> "*Varies*" And My.Settings.Custom = "Drawing" Then
            oDoc.PropertySets.Item("Inventor User Defined Properties").Item("Reference5").Value = txtRef5.Text
        End If
        If txtRef6.Text <> "*Varies*" And My.Settings.Custom = "Drawing" Then
            oDoc.PropertySets.Item("Inventor User Defined Properties").Item("Reference6").Value = txtRef6.Text
        End If
        If txtRef7.Text <> "*Varies*" And My.Settings.Custom = "Drawing" Then
            oDoc.PropertySets.Item("Inventor User Defined Properties").Item("Reference7").Value = txtRef7.Text
        End If
        If txtRef8.Text <> "*Varies*" And My.Settings.Custom = "Drawing" Then
            oDoc.PropertySets.Item("Inventor User Defined Properties").Item("Reference8").Value = txtRef8.Text
        End If
        If txtRef9.Text <> "*Varies*" And My.Settings.Custom = "Drawing" Then
            oDoc.PropertySets.Item("Inventor User Defined Properties").Item("Reference9").Value = txtRef9.Text
        End If
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
        For Each RevRow As RevisionTableRow In shtRevTable.RevisionTableRows
            Dim RtCell As RevisionTableCell
            For Each RtCell In RevRow
                'skip over items not used and proceed to next row
                If J Mod 5 = 0 And J > 0 Then
                    J = 0
                    I += 1
                End If
                If I < shtRevTable.RevisionTableRows.Count Then
                    If J > 0 Then
                        If RevTable.lstCheckFiles.Items(I).SubItems(J).Text <> "*Varies*" Then
                            RtCell.Text = RevTable.lstCheckFiles.Items(I).SubItems(J).Text
                        End If
                        J += 1
                    Else
                        If Strings.Len(RevTable.lstCheckFiles.Items(I).SubItems(J).Text) = 1 Then
                            J += 1
                        End If
                    End If
                End If
            Next
        Next
        ' Adding Revs
        Dim RevTab As Integer = TabControl1.TabCount
        If TabControl1.TabPages.Item(TabControl1.TabPages.Count - 1).Controls.Item("txtrev" & RevTab.ToString & "Init").Enabled = True Then
            shtRevTable.RevisionTableRows.Add()
            I = 0 : J = 0
            For Each revrow As RevisionTableRow In shtRevTable.RevisionTableRows
                Dim RtCell As RevisionTableCell
                For Each RtCell In revrow
                    ' I = shtRevTable.RevisionTableColumns.Count * shtRevTable.RevisionTableRows.Count - 3 _
                    'To shtRevTable.RevisionTableColumns.Count * shtRevTable.RevisionTableRows.Count
                    If I = shtRevTable.RevisionTableRows.Count - 1 Then
                        Select Case J
                            Case = shtRevTable.RevisionTableColumns.Count * shtRevTable.RevisionTableRows.Count - 3
                                RtCell.Text = TabControl1.TabPages.Item(TabControl1.TabPages.Count - 1).Controls.Item("txtRev" & RevTab.ToString & "Des").Text
                            Case = shtRevTable.RevisionTableColumns.Count * shtRevTable.RevisionTableRows.Count - 2
                                RtCell.Text = TabControl1.TabPages.Item(TabControl1.TabPages.Count - 1).Controls.Item("txtRev" & RevTab.ToString & "Init").Text
                            Case = shtRevTable.RevisionTableColumns.Count * shtRevTable.RevisionTableRows.Count - 1
                                RtCell.Text = TabControl1.TabPages.Item(TabControl1.TabPages.Count - 1).Controls.Item("txtRev" & RevTab.ToString & "Approved").Text
                        End Select
                    End If
                    'Next
                    J += 1
                Next
                I += 1
            Next
        End If
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
End Class