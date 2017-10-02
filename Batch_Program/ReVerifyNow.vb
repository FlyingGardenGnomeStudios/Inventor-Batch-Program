Imports System.Windows.Forms
Imports wyDay
Imports wyDay.TurboActivate

Public Class ReVerifyNow
    Dim Main As Main
    Private ta As TurboActivate
    Private inGrace As Boolean = False
    Public GenuineDaysLeft As UInteger

    Public noLongerActivated As Boolean = False

    Public Sub New(ta As TurboActivate, DaysBetweenChecks As UInteger, GracePeriodLength As UInteger)
        ' Set the TurboActivate instance from the main form
        Me.ta = ta

        InitializeComponent()

        ' Use the days between checks and grace period from
        ' the main form
        GenuineDaysLeft = ta.GenuineDays(DaysBetweenChecks, GracePeriodLength, inGrace)

        If GenuineDaysLeft = 0UI Then
            lblDescr.Text = "You must re-verify with the activation servers to continue using this app."
        Else
            lblDescr.Text = "You have " + GenuineDaysLeft + " days to re-verify with the activation servers."
        End If
    End Sub
    Public Function PopMain(CalledFunction As Main)
        Main = CalledFunction
        Return Nothing
    End Function

    Private Sub btnReverify_Click(sender As Object, e As EventArgs) Handles btnReverify.Click
        Try
            Select Case ta.IsGenuine()
                Case IsGenuineResult.Genuine, IsGenuineResult.GenuineFeaturesChanged
                    DialogResult = DialogResult.OK
                    Close()

                Case IsGenuineResult.NotGenuine, IsGenuineResult.NotGenuineInVM
                    noLongerActivated = True
                    DialogResult = DialogResult.Cancel
                    Close()

                Case IsGenuineResult.InternetError
                    MessageBox.Show("Failed to connect with the activation servers.")
            End Select
        Catch ex As Exception
            MessageBox.Show("Failed to re-verify with the activation servers. Full error: " + ex.Message)
        End Try
    End Sub

    Private Sub btnExit_Click(sender As Object, e As EventArgs) Handles btnExit.Click
        DialogResult = DialogResult.Cancel
        Close()
    End Sub
End Class