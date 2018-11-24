Public Class BugFound
    Private Sub btnOK_Click(sender As Object, e As EventArgs) Handles btnOK.Click
        Me.Close()
    End Sub

    Private Sub btnLogIssue_Click(sender As Object, e As EventArgs) Handles btnLogIssue.Click
        Try
            Process.Start("https://github.com/FlyingGardenGnomeStudios/Inventor-Batch-Program/issues")
        Catch ex As Exception
            MsgBox("Umm, oops..." & vbNewLine & "I can't even find the website." & vbNewLine & "Unfortunately it looks like you've been abandoned. I'm truly sorry")
        Finally
            Me.Close()
        End Try
    End Sub

    Private Sub PictureBox1_Click(sender As Object, e As EventArgs) Handles PictureBox1.Click
        MsgBox("CONGRATULATIONS!!!!" & vbNewLine &
               "Your random clicking has won you the chance to close this window!!!!")
    End Sub

    Private Sub Label1_Click(sender As Object, e As EventArgs) Handles Label1.Click

    End Sub
End Class