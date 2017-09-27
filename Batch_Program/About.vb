Imports System.Reflection
Imports System.Windows.Forms

Public Class About
    Public Sub New()
        InitializeComponent()
        Dim fecha As Date = IO.File.GetCreationTime(Assembly.GetExecutingAssembly().Location)
        lblCopyright.Text = My.Application.Info.Copyright
        If My.Application.IsNetworkDeployed Then
            With System.Deployment.Application.ApplicationDeployment.CurrentDeployment.CurrentVersion
                lblName.Text = "Inventor Batch Program " & .Major.ToString() & "." &
                .Minor.ToString() & "." &
                .Build.ToString() & "." &
                .Revision.ToString() & " Build Date: " & fecha.ToShortDateString.ToString

            End With
        Else
            lblPublish.Text = "Test Mode" & " Build Date: " & fecha.ToShortDateString.ToString
        End If

    End Sub
    Private Sub PictureBox1_Click(sender As Object, e As EventArgs) Handles PicDonate.Click
        Try
            System.Diagnostics.Process.Start("https://www.paypal.com/cgi-bin/webscr?cmd=_s-xclick&hosted_button_id=HXDSA8NWCRM2G")
        Catch
            'Code to handle the error.
        End Try
    End Sub

End Class

