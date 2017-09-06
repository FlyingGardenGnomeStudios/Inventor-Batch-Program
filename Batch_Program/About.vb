Public Class About
    Public Sub New()
        InitializeComponent()
        lblCopyright.Text = My.Application.Info.Copyright
        If My.Application.IsNetworkDeployed Then
            With System.Deployment.Application.ApplicationDeployment.CurrentDeployment.CurrentVersion
                lblName.Text = "Inventor Batch Program " & .Major.ToString() & "." &
                .Minor.ToString() & "." &
                .Build.ToString() & "." &
                .Revision.ToString()

            End With
        Else
            lblPublish.Text = "Test Mode"
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

