Imports System.Reflection
Imports System.Windows.Forms

Public Class About
    Public Sub New()
        InitializeComponent()
        Dim fecha As Date = IO.File.GetCreationTime(Assembly.GetExecutingAssembly().Location)
        lblCopyright.Text = My.Application.Info.Copyright
        Dim AssemblyLocation = Assembly.GetExecutingAssembly().Location
        Dim VersionNumber As String = System.Diagnostics.FileVersionInfo.GetVersionInfo(AssemblyLocation).FileVersion

        lblPublish.Text = "Inventor Batch Program " & VersionNumber


    End Sub

    Private Sub PicDonate_Click(sender As Object, e As EventArgs) Handles PicDonate.Click
        Try
            System.Diagnostics.Process.Start("https://www.PayPal.Me/fggs")
        Catch
            'Code to handle the error.
        End Try
    End Sub
End Class

