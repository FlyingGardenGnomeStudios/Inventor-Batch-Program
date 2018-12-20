Imports System.Runtime.InteropServices
Imports System.Windows.Forms
Imports Inventor
Public Class Print
    Dim Main As Main
    Dim Print_Size As New Print_Size
    Dim _invApp As Inventor.Application
    Dim PPath, PoDoc As Documents
    Dim PArchive, PDrawingName, PDrawSource As String
    Dim POpenDocs As ArrayList
    Dim PSubfiles As List(Of KeyValuePair(Of String, String))
    Dim ASubfiles As SortedList(Of String, String)
    Public Function PopMain(CalledFunction As Main)
        Main = CalledFunction
        Return Nothing
    End Function
    Public Sub New()
        ' This call is required by the designer.
        InitializeComponent()
        _invApp = Marshal.GetActiveObject("Inventor.Application")
        ' Add any initialization after the InitializeComponent() call.
    End Sub
    Public Sub PopulatePrint(Path As Documents, ByRef odoc As Document, ByRef Archive As String _
                             , ByRef DrawingName As String, ByRef DrawSource As String, OpenDocs As ArrayList _
                             , ByRef SubFiles As List(Of KeyValuePair(Of String, String)), ByRef AlphaSub As SortedList(Of String, String))
        PPath = Path
        PoDoc = odoc
        PArchive = Archive
        PDrawingName = DrawingName
        PDrawSource = DrawSource
        POpenDocs = OpenDocs
        PSubfiles = SubFiles
        ASubfiles = AlphaSub
        Select Case My.Settings.PrintSize
            Case 0
                rdoFull.Checked = True
            Case 1
                rdoScale.Checked = True
        End Select
        If My.Settings.PrintDwgLoc = True Then
            Select Case My.Settings.DWGLoc
                Case "TR"
                    rdbTR.Checked = True
                Case "TL"
                    rdbTL.Checked = True
                Case "BR"
                    rdbBR.Checked = True
                Case "BL"
                    rdbBL.Checked = True
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
            Case 0
                rdoColour.Checked = True
            Case 1
                rdoBW.Checked = True
        End Select
        chkDWGLocation.Checked = My.Settings.PrintDwgLoc
    End Sub

    Private Sub rdoAllPages_CheckedChanged(sender As Object, e As EventArgs) Handles rdoAllPages.CheckedChanged
        If rdoAllPages.Checked = True Then
            rdoCurrentPage.Checked = False
            rdoFirstPage.Checked = False
        End If
    End Sub

    Private Sub rdoFirstPage_CheckedChanged(sender As Object, e As EventArgs) Handles rdoFirstPage.CheckedChanged
        If rdoFirstPage.Checked = True Then
            rdoAllPages.Checked = False
            rdoCurrentPage.Checked = False
        End If
    End Sub

    Private Sub rdoBW_CheckedChanged(sender As Object, e As EventArgs) Handles rdoBW.CheckedChanged
        If rdoBW.Checked = True Then
            rdoColour.Checked = False
        End If
    End Sub

    Private Sub rdoColour_CheckedChanged(sender As Object, e As EventArgs) Handles rdoColour.CheckedChanged
        If rdoColour.Checked = True Then
            rdoBW.Checked = False
        End If
    End Sub

    Private Sub rdoCurrentPage_CheckedChanged(sender As Object, e As EventArgs) Handles rdoCurrentPage.CheckedChanged
        If rdoCurrentPage.Checked = True Then
            rdoAllPages.Checked = False
            rdoFirstPage.Checked = False
        End If
    End Sub

    Private Sub btnExpand_Click(sender As Object, e As EventArgs) Handles btnExpand.Click
        If btnExpand.Text = ">>" Then
            Me.Width = 400
            btnExpand.Text = "<<"
        Else
            Me.Width = 280
            btnExpand.Text = ">>"
        End If
    End Sub

    Private Sub RadioButton2_CheckedChanged(sender As Object, e As EventArgs) Handles rdbBR.CheckedChanged
        If rdbBR.Checked = True Then
            rdbBL.Checked = False
            rdbTL.Checked = False
            rdbTR.Checked = False
            My.Settings.DWGLoc = "BR"
            My.Settings.Save()
        End If
    End Sub

    Private Sub rdbBL_CheckedChanged(sender As Object, e As EventArgs) Handles rdbBL.CheckedChanged
        If rdbBL.Checked = True Then
            rdbTL.Checked = False
            rdbTR.Checked = False
            rdbBR.Checked = False
            My.Settings.DWGLoc = "BL"
            My.Settings.Save()
        End If
    End Sub

    Private Sub rdbTL_CheckedChanged(sender As Object, e As EventArgs) Handles rdbTL.CheckedChanged
        If rdbTL.Checked = True Then
            rdbBL.Checked = False
            rdbTR.Checked = False
            rdbBR.Checked = False
            My.Settings.DWGLoc = "TL"
            My.Settings.Save()
        End If
    End Sub

    Private Sub rdbTR_CheckedChanged(sender As Object, e As EventArgs) Handles rdbTR.CheckedChanged
        If rdbTR.Checked = True Then
            rdbBL.Checked = False
            rdbTL.Checked = False
            rdbBR.Checked = False
            My.Settings.DWGLoc = "True"
            My.Settings.Save()
        End If
    End Sub

    Private Sub dwgLocation_CheckedChanged(sender As Object, e As EventArgs) Handles chkDWGLocation.CheckedChanged
        If chkDWGLocation.Checked = True Then
            Select Case My.Settings.DWGLoc
                Case "TR"
                    rdbTR.Checked = True
                Case "TL"
                    rdbTL.Checked = True
                Case "BR"
                    rdbBR.Checked = True
                Case "BL"
                    rdbBL.Checked = True
            End Select
        Else
            rdbTR.Checked = False
            rdbTL.Checked = False
            rdbBR.Checked = False
            rdbBL.Checked = False
        End If
    End Sub

    Private Sub btnCancel_Click(sender As Object, e As EventArgs) Handles btnCancel.Click
        Me.Close()
    End Sub

    Private Sub btnOK_Click(sender As Object, e As EventArgs) Handles btnOK.Click
        Print_Size.PopPrint(Me)
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
            My.Settings.PrintColour = 0
        Else
            My.Settings.PrintColour = 1
        End If
        My.Settings.PrintDwgLoc = chkDWGLocation.Checked

        My.Settings.PrintReverse = chkReverse.Checked
        My.Settings.Save()
        Me.Hide()
        Dim dDoc As Inventor.DrawingDocument
        Dim Range, Direction As Integer : Range = 0 : Direction = 0
        Dim PEnd, PStart, Z As Integer
        Dim Colour As Boolean
        Dim Mass As Double
        Dim DrawSource As String = ""
        Dim DrawingName As String = ""
        Z = 1
        Dim ScaleSelect As Boolean
        If rdoFull.Checked = True Then
            ScaleSelect = False
        Else
            ScaleSelect = True
        End If
        If rdoAllPages.Checked = True Then
            Range = 3
        ElseIf rdoCurrentPage.Checked = True Then
            Range = 2
        Else
            Range = 1
        End If
        If rdoBW.Checked = True Then
            Colour = False
        Else
            Colour = True
        End If
        For Y = 1 To txtCopies.Value
            If chkReverse.Checked = False Then
                Direction = 1
                PEnd = Main.dgvSubFiles.RowCount - 1
                PStart = 0
            Else
                Direction = -1
                PStart = Main.dgvSubFiles.RowCount - 1
                PEnd = 0
            End If
            For X = PStart To PEnd Step Direction
                If Main.bgwRun.CancellationPending = True Then Exit Sub
                If Main.dgvSubFiles(Main.dgvSubFiles.Columns("chkSubFiles").Index, X).Value = True Then
                    DrawingName = Main.dgvSubFiles(Main.dgvSubFiles.Columns("DrawingName").Index, X).Value
                    dDoc = _invApp.Documents.Open(Main.dgvSubFiles(Main.dgvSubFiles.Columns("DrawingLocation").Index, X).Value, True)
                    'For Each oDoc In dDoc.ReferencedFiles
                    'Mass = oDoc.ComponentDefinition.MassProperties.mass
                    Try
                        Mass = dDoc.ReferencedFiles.Item(1).componentdefinition.massproperties.mass
                        'Next
                        Call dDoc.Update()
                    Catch
                    End Try
                    'PDrawingName = Strings.Right(dDoc.FullDocumentName, Len(dDoc.FullDocumentName) - InStrRev(dDoc.FullDocumentName, "\"))
                    'Main.ProgressBar(PEnd + 1 * Y, Z, "Printing: ", DrawingName)
                    Main.bgwRun.ReportProgress((Z / (PEnd + 1 * Y)) * 100, "Printing: " & DrawingName)
                    Z += 1

                    PrintSheets(DrawingName, ScaleSelect, Range, dDoc, Colour)
                    Main.CloseLater(DrawingName, dDoc)
                End If
            Next
        Next
        Print_Size.rdoAsk.Checked = True
        Print_Size.rdoDontAsk.Checked = False
        Print_Size.chkQuesiton.Checked = False
        Print_Size.chkScale.Checked = False
        Me.Close()
    End Sub
    Public Sub PrintSheets(DrawingName As String, ScaleSelect As Boolean, Range As Integer, dDoc As Document, Colour As Boolean)
        Dim Size As Double
        'Set up print manager according to generic standards
        Dim oDoc As Inventor.DrawingDocument
        oDoc = _invApp.ActiveDocument
        Dim oSheet As Inventor.Sheet
        oSheet = oDoc.ActiveSheet
        Dim oPM As Inventor.PrintManager
        Dim odef As ControlDefinition
        oPM = oDoc.PrintManager
        oPM.scalemode = Inventor.PrintScaleModeEnum.kPrintBestFitScale
        oPM.printrange = Inventor.PrintRangeEnum.kPrintCurrentSheet
        If Range <> 2 Then
            dDoc.Sheets.Item(1).Activate()
        End If

        If Colour = True Then
            oPM.ColorMode = PrintColorModeEnum.kPrintDefaultColorMode
        Else
            oPM.ColorMode = PrintColorModeEnum.kPrintGrayScale
        End If
        'Scale drawings to the same size if chosen by the user
        If Print_Size.lblDWGSize.Text = "Cancel" Then Exit Sub
        For X = 1 To dDoc.sheets.count
            If Range = 3 Then dDoc.sheets.item(X).activate
            Size = dDoc.activesheet.size
            If ScaleSelect = True Then
                If Size = 9987 Or Size = 9988 Then
                    oPM.PaperSize = PaperSizeEnum.kPaperSizeLetter
                Else
                    oPM.PaperSize = PaperSizeEnum.kPaperSize11x17
                End If
            Else
                'Scale the drawings individually if chosen by the user
                If Size = 9987 Then
                    oPM.PaperSize = PaperSizeEnum.kPaperSizeLetter
                ElseIf Size = 9988 Then
                    oPM.PaperSize = PaperSizeEnum.kPaperSize11x17
                ElseIf Size <> 9987 And Size <> 9988 And Print_Size.chkScale.Checked = False And Print_Size.rdoAsk.Checked = True Then

                    Print_Size.lblDWGSize.Text = "Drawing " & DrawingName & " Sht " & X & " is larger than 11x17" & vbNewLine &
                    "Do you still wish to print this drawing at full size?"
                    Print_Size.ShowDialog()
                    'notify if the drawing is larger than 11x17 and open Inventor print dialogue.

                    If Print_Size.chkScale.Checked = True Then
                        odef = _invApp.CommandManager.ControlDefinitions("AppFilePrintCmd")
                        odef.Execute()
                    Else
                        oPM.PaperSize = PaperSizeEnum.kPaperSize11x17
                    End If
                ElseIf Size <> 9987 And Size <> 9988 And Print_Size.chkScale.Checked = True Then
                    odef = _invApp.CommandManager.ControlDefinitions("AppFilePrintCmd")
                    odef.Execute()
                ElseIf Size <> 9987 And Size <> 9988 And Print_Size.rdoDontAsk.Checked = True Then
                    oPM.PaperSize = PaperSizeEnum.kPaperSize11x17
                End If
            End If
            If Print_Size.lblDWGSize.Text = "Cancel" Then Exit Sub
            If oDoc.Sheets.Item(X).ExcludeFromPrinting = False Then
                Dim oSketch As DrawingSketch = Nothing
                If My.Settings.PrintDwgLoc = True Then
                    oSketch = dDoc.activesheet.sketches.add
                    oSketch.Edit()
                    Dim oTG As TransientGeometry = _invApp.TransientGeometry
                    Dim DWGLoc As String = dDoc.FullFileName
                    Dim oBorder As Border = Nothing
                    Try
                        oBorder = oDoc.ActiveSheet.Border
                    Catch ex As Exception
                        MessageBox.Show("Drawing border missing" & vbNewLine & ex.Message)
                    End Try
                    Dim oTextBox As Inventor.TextBox
                    Select Case My.Settings.DWGLoc
                        Case "TR"
                            oTextBox = oSketch.TextBoxes.AddFitted(oTG.CreatePoint2d(oBorder.RangeBox.MaxPoint.X, oBorder.RangeBox.MaxPoint.Y + 0.45), DWGLoc)
                            oTextBox.HorizontalJustification = HorizontalTextAlignmentEnum.kAlignTextRight
                        Case "BR"
                            oTextBox = oSketch.TextBoxes.AddFitted(oTG.CreatePoint2d(oBorder.RangeBox.MaxPoint.X, oBorder.RangeBox.MinPoint.Y - 0.3), DWGLoc)
                            oTextBox.HorizontalJustification = HorizontalTextAlignmentEnum.kAlignTextRight
                        Case "TL"
                            oTextBox = oSketch.TextBoxes.AddFitted(oTG.CreatePoint2d(oBorder.RangeBox.MinPoint.X, oBorder.RangeBox.MaxPoint.Y + 0.45), DWGLoc)
                            oTextBox.HorizontalJustification = HorizontalTextAlignmentEnum.kAlignTextLeft
                        Case "BL"
                            oTextBox = oSketch.TextBoxes.AddFitted(oTG.CreatePoint2d(oBorder.RangeBox.MinPoint.X, oBorder.RangeBox.MinPoint.Y - 0.3), DWGLoc)
                            oTextBox.HorizontalJustification = HorizontalTextAlignmentEnum.kAlignTextLeft
                    End Select
                End If
                oPM.SubmitPrint()
                Try
                    oSketch.ExitEdit()
                    oSketch.Delete()
                Catch
                End Try
            End If

                If Range <> 3 Then Exit For
        Next
        'oPM.PaperSize = PaperSizeEnum.kPaperSizeLetter
    End Sub
End Class