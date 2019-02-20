Public Class Form1


    Private Sub MsVistaProgressBar2_MouseClick(ByVal sender As System.Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles MsVistaProgressBar2.MouseClick, MsVistaProgressBar1.MouseClick, MsVistaProgressBar3.MouseClick
        PropertyGrid1.SelectedObject = sender
    End Sub

End Class
