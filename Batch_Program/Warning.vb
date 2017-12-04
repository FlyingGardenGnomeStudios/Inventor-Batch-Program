
Imports System.Runtime.InteropServices
Public Class Warning
    Public Sub New()
        ' This call is required by the designer.
        InitializeComponent()
    End Sub
    Public Sub Donate()
        Label1.Text = "Developing takes time and money. " & vbNewLine & "If this program has saved you either, consider saying thanks."
        PicDonate.Visible = True
        If My.Settings.Donated = True Then
            chkDontShow.Visible = True
            chkDontShow.Text = "I already did"
        Else
            chkDontShow.Visible = False
            PicDonate.Left = chkDontShow.Left
        End If

        btnOK.Width = 200
        Label2.Text = "Donate"
        btnOK.Left = btnOK.Left - 120
        btnOK.Text = "I'll just keep using it for free thanks"
        ' Add any initialization after the InitializeComponent() call.
        Me.ShowDialog()
    End Sub
    Public Sub FirstRun()
        Me.Text = "First time information"
        Me.Height = 150
        Label2.Text = "FirstRun"
        btnOK.Location = New Drawing.Point(btnOK.Location.X, btnOK.Location.Y + 40)
        Label1.Text = "Batch Program is a program created to automate many tasks in Inventor. The program is in a constant state of developement so bugs are likely. If you find any errors/bugs let me know by emailing a description of how to replicate the error to flyinggardengnomestudios@gmail.com. Please include the error log which is saved in: " & My.Computer.FileSystem.SpecialDirectories.Temp & "\debug.txt"
    End Sub
    Private Sub btnOK_Click(sender As Object, e As EventArgs) Handles btnOK.Click


        Select Case Label2.Text
            Case "Rename"
                If chkDontShow.Checked = True Then
                    My.Settings.RenameShowMe = False
                End If
            Case "Donate"
                If btnOK.Text = "I'll just keep using it for free thanks" Then
                    My.Settings.DonateCount = My.Settings.DonateCount + 1
                Else
                    My.Settings.DonateShowMe = False
                End If
            Case "FirstRun"
                My.Settings.FirstRun = False
        End Select
        Me.Close()
    End Sub
    Private Sub PictureBox1_Click(sender As Object, e As EventArgs) Handles PicDonate.Click
        Try
            System.Diagnostics.Process.Start("https://www.paypal.com/cgi-bin/webscr?cmd=_s-xclick&hosted_button_id=HXDSA8NWCRM2G")
        Catch
            'Code to handle the error.
        End Try
        Me.Close()
        My.Settings.Donated = True
    End Sub

    Private Sub chkDontShow_CheckedChanged(sender As Object, e As EventArgs) Handles chkDontShow.CheckedChanged
        PicDonate.Visible = False
        btnOK.Text = "My bad, carry on and disregard"
    End Sub
End Class