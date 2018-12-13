Public Class Select_Template
    Public Sub New()

        ' This call is required by the designer.
        InitializeComponent()

        ' Add any initialization after the InitializeComponent() call.

    End Sub
    Public Function PopMain(CalledFunction As Main)
        Main = CalledFunction
        Return Nothing
    End Function
    Private Sub btnOK_Click(sender As Object, e As EventArgs) Handles btnOK.Click

    End Sub
End Class