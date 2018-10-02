Imports System.Runtime.InteropServices
Imports Inventor
Public Class Overwrite
    Dim Main As Main
    Dim _invApp As Inventor.Application

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

    Private Sub btnOK_Click(sender As Object, e As EventArgs) Handles btnOK.Click
        Dim sReadableType As String = ""
        Dim NotMade As String = ""
        Dim Archive As String = ""
        Dim oDoc As Document
        For Each Row In dgvOverwrite.Rows
            Dim Destin As String = dgvOverwrite(dgvOverwrite.Columns("Destin").Index, Row.index).Value
            Dim ExportType As String = dgvOverwrite(dgvOverwrite.Columns("Type").Index, Row.index).Value
            Dim DrawSource As String = dgvOverwrite(dgvOverwrite.Columns("DrawSource").Index, Row.index).Value
            Dim DrawingName As String = dgvOverwrite(dgvOverwrite.Columns("DrawingName").Index, Row.index).Value
            Dim RevNo As String = dgvOverwrite(dgvOverwrite.Columns("Rev").Index, Row.index).Value
            Dim Total As Integer = CInt(dgvOverwrite(dgvOverwrite.Columns("Total").Index, Row.index).Value)
            Dim Counter As Integer = CInt(dgvOverwrite(dgvOverwrite.Columns("Counter").Index, Row.index).Value)

            If dgvOverwrite(dgvOverwrite.Columns("Type").Index, Row.index).Value = "PDF" Then
                Main.PDFCreator(Destin, Total, Counter)
            Else
                Main.Search_For_Duplicates(Destin, DrawingName, "." & ExportType)

                If My.Computer.FileSystem.FileExists(Strings.Replace(DrawSource, "idw", "ipt")) = True Then
                    Archive = Strings.Replace(DrawSource, "idw", "ipt")
                    oDoc = _invApp.Documents.Open(Archive, True)
                    Main.SheetMetalTest(oDoc, sReadableType)
                    If sReadableType = "P" And ExportType <> "PDF" Then
                        Call Main.ExportPart(DrawSource, Archive, False, Destin, DrawingName, ExportType, RevNo)
                    ElseIf sReadableType = "S" And ExportType <> "PDF" Then
                        Call Main.SMDXF(oDoc, DrawSource, True, Replace(DrawingName, ".idw", "." & ExportType))
                        Main.CloseLater(Strings.Left(DrawingName, Len(DrawingName) - 3) & "ipt", oDoc)
                    ElseIf sReadableType = "S" And ExportType <> "PDF" Then
                        Call Main.ExportPart(DrawSource, Archive, False, Destin, DrawingName, ExportType, RevNo)
                    ElseIf sReadableType <> "" And ExportType <> "PDF" Then
                        Call Main.ExportPart(DrawSource, Archive, False, Destin, DrawingName, ExportType, RevNo)
                    ElseIf sReadableType = "" Then
                        Main.CloseLater(IO.Path.GetFileName(oDoc.FullDocumentName), oDoc)
                    End If
                    If Main.chkEditRev.CheckState = Windows.Forms.CheckState.Indeterminate Then
                        Main.chkEditRev.CheckState = Windows.Forms.CheckState.Unchecked
                    End If
                ElseIf My.Computer.FileSystem.FileExists(Strings.Replace(DrawSource, "idw", "iam")) = False Then
                    NotMade = DrawingName & vbNewLine
                ElseIf My.Computer.FileSystem.FileExists(Strings.Replace(DrawSource, "idw", "iam")) = True AndAlso
                    Main.chkSkipAssy.Checked = False Then
                    Call Main.ExportPart(DrawSource, Strings.Replace(Archive, "idw", "iam"), False, Destin, DrawingName, ExportType, RevNo)
                End If
                Main.bgwRun.ReportProgress((Counter / Total) * 100, "Saving: " & Replace(DrawingName, ".idw", "." & ExportType))
            End If
        Next
        Me.Close()
    End Sub

    Private Sub btnCancel_Click(sender As Object, e As EventArgs) Handles btnCancel.Click
        Me.Close()
    End Sub
End Class