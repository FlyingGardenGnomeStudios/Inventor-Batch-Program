Imports System.Runtime.InteropServices
Imports Inventor

Public Class Print_Size
    Dim _invApp As Inventor.Application
    Dim Main As Main
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
    Private Sub rdoAsk_CheckedChanged(sender As Object, e As EventArgs) Handles rdoAsk.CheckedChanged
        If rdoAsk.Checked = True Then
            rdoDontAsk.Checked = False
            chkQuestion.Checked = True
        End If
    End Sub

    Private Sub rdoDontAsk_CheckedChanged(sender As Object, e As EventArgs) Handles rdoDontAsk.CheckedChanged
        If rdoDontAsk.Checked = True Then
            rdoAsk.Checked = False
            chkQuestion.Checked = False
        End If
    End Sub

    Private Sub btnFullSize_Click(sender As Object, e As EventArgs)

        Me.Hide()
    End Sub

    Private Sub btn11x17_Click(sender As Object, e As EventArgs) Handles btn11x17.Click
        Dim oDef As ControlDefinition = _invApp.CommandManager.ControlDefinitions("AppFilePrintCmd")
        oDef.Execute()
        chkScale.Checked = False
        Me.Close()
    End Sub

    Private Sub btnCancel_Click(sender As Object, e As EventArgs) Handles btnCancel.Click
        lblDWGSize.Text = "Cancel"
        Me.Close()
    End Sub

    Private Sub cmbScaleSize_SelectedIndexChanged(sender As Object, e As EventArgs) Handles cmbScaleSize.SelectedIndexChanged
        chkScale.Checked = True
    End Sub

    Private Sub btnOK_Click(sender As Object, e As EventArgs) Handles btnOK.Click
        If cmbScaleSize.SelectedItem.contains("Scale to") Then
            chkScale.CheckState = Windows.Forms.CheckState.Checked
        Else
            chkScale.CheckState = Windows.Forms.CheckState.Indeterminate
        End If
        Me.Close()
    End Sub

    Private Sub Print_Size_Shown(sender As Object, e As EventArgs) Handles Me.Shown

        cmbScaleSize.SelectedIndex = 0
    End Sub
End Class