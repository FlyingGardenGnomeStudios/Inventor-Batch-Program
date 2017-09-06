Imports System.Runtime.InteropServices
Imports System.Windows.Forms
Imports Inventor
Public Class Export
    Dim Main As Main
    Dim _invApp As Inventor.Application
    Dim PPath, PoDoc As Documents
    Dim PArchive, PDrawingName, PDrawSource As String
    Dim POpenDocs As ArrayList
    Dim PSubfiles As List(Of KeyValuePair(Of String, String))
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
                             , ByRef SubFiles As List(Of KeyValuePair(Of String, String)))
        PPath = Path
        PoDoc = odoc
        PArchive = Archive
        PDrawingName = DrawingName
        PDrawSource = DrawSource
        POpenDocs = OpenDocs
        PSubfiles = SubFiles
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
    Private Sub rdoCurrentPage_CheckedChanged(sender As Object, e As EventArgs) Handles rdoCurrentPage.CheckedChanged
        If rdoCurrentPage.Checked = True Then
            rdoAllPages.Checked = False
            rdoFirstPage.Checked = False
        End If
    End Sub
    Private Sub btnCancel_Click(sender As Object, e As EventArgs) Handles btnCancel.Click
        Me.Close()
    End Sub
End Class