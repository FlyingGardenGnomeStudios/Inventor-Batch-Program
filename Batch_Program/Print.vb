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

    Private Sub btnCancel_Click(sender As Object, e As EventArgs) Handles btnCancel.Click
        Me.Close()
    End Sub

    Private Sub btnOK_Click(sender As Object, e As EventArgs) Handles btnOK.Click
        Print_Size.PopPrint(Me)
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
                PEnd = Main.LVSubFiles.Items.Count - 1
                PStart = 0
            Else
                Direction = -1
                PStart = Main.LVSubFiles.Items.Count - 1
                PEnd = 0
            End If
            For X = PStart To PEnd Step Direction
                If Main.LVSubFiles.Items(X).Checked = True Then
                    Main.MatchDrawing(DrawSource, DrawingName, X)
                    dDoc = _invApp.Documents.Open(DrawSource, True)
                    'For Each oDoc In dDoc.ReferencedFiles
                    'Mass = oDoc.ComponentDefinition.MassProperties.mass
                    Try
                        Mass = dDoc.ReferencedFiles.Item(1).componentdefinition.massproperties.mass
                        'Next
                        Call dDoc.Update()
                    Catch
                    End Try
                    'PDrawingName = Strings.Right(dDoc.FullDocumentName, Len(dDoc.FullDocumentName) - InStrRev(dDoc.FullDocumentName, "\"))
                    Main.ProgressBar(PEnd + 1 * Y, Z, "Printing: ", PDrawingName)
                    Z += 1
                    PrintSheets(PDrawingName, ScaleSelect, Range, dDoc, Colour)
                    Main.CloseLater(PDrawingName, dDoc)
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
            oPM.SubmitPrint()

            If Range <> 3 Then Exit For
        Next
        'oPM.PaperSize = PaperSizeEnum.kPaperSizeLetter
    End Sub
End Class