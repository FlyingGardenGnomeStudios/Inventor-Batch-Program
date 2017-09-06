Public Class Rev_Switch
    Dim Main As Main
    Dim RevType As Integer
    Public Function PopMain(CalledFunction As Main)
        Main = CalledFunction
        Return Nothing
    End Function
    Public Sub First(ByRef RevType As Integer)
        Me.ShowDialog()
    End Sub
    Private Sub btnNumeric_Click(sender As System.Object, e As System.EventArgs) Handles btnNumeric.Click
        Main.ChangeRev(1)
        Me.Close()
    End Sub

    Private Sub btnAlpha_Click(sender As System.Object, e As System.EventArgs) Handles btnAlpha.Click
        Main.ChangeRev(0)
        Me.Close()
    End Sub
End Class