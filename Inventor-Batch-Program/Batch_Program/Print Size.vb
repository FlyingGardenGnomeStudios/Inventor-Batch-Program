Imports System.Runtime.InteropServices

Public Class Print_Size
    Dim _invApp As Inventor.Application
    Dim Main As Main
    Dim Print As Print
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
    Public Function PopPrint(CalledFunction As Print)
        PopPrint = CalledFunction
        Return Nothing
    End Function
    Private Sub rdoAsk_CheckedChanged(sender As Object, e As EventArgs) Handles rdoAsk.CheckedChanged
        If rdoAsk.Checked = True Then
            rdoDontAsk.Checked = False
            chkQuesiton.Checked = True
        End If
    End Sub

    Private Sub rdoDontAsk_CheckedChanged(sender As Object, e As EventArgs) Handles rdoDontAsk.CheckedChanged
        If rdoDontAsk.Checked = True Then
            rdoAsk.Checked = False
            chkQuesiton.Checked = False
        End If
    End Sub

    Private Sub btnFullSize_Click(sender As Object, e As EventArgs) Handles btnFullSize.Click
        chkScale.Checked = True
        Me.Hide()
    End Sub

    Private Sub btn11x17_Click(sender As Object, e As EventArgs) Handles btn11x17.Click
        chkScale.Checked = False
        Me.Hide()
    End Sub

    Private Sub btnCancel_Click(sender As Object, e As EventArgs) Handles btnCancel.Click
        lblDWGSize.Text = "Cancel"
        Me.Close()
    End Sub
End Class