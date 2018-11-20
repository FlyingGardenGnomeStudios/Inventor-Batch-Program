Public Class WarningList
    Private Sub WarningList_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        dgvWarning.Rows.Add("First Run", "Batch Program is a program created to automate many tasks in Inventor. The program is in a constant state of developement so bugs are likely. If you find any errors/bugs let me know by emailing a description of how to replicate the error to flyinggardengnomestudios@gmail.com. Please include the error log which is saved in: " & My.Computer.FileSystem.SpecialDirectories.Temp & "\debug.txt", My.Settings.FirstRun.ToString)
        dgvWarning.Rows.Add("Rename", "The renaming function does not currently support Top-Down assembly configurations.
Referenced function parameters will not be re-linked to the new part name.", My.Settings.RenameShowMe.ToString)
        dgvWarning.Rows.Add("Bad File Type", "Some items were found that have an unsupported extension or the file extensions are invisible." &
                       " These files may Not work correctly With the program. " &
                       "Currently only the native Inventor extensions are supported (.ipt, .idw, .iam, & .ipn)" & vbNewLine &
                       "(You can enable file name extension visibility through the file explorer)", My.Settings.BadFileTypeWarning.ToString)
    End Sub

    Private Sub btnCancel_Click(sender As Object, e As EventArgs) Handles btnCancel.Click
       Me.Close
    End Sub

    Private Sub btnOK_Click(sender As Object, e As EventArgs) Handles btnOK.Click
        My.Settings.FirstRun = dgvWarning.Rows(0).Cells("Notification").Value
        My.Settings.RenameShowMe = dgvWarning.Rows(1).Cells("Notification").Value
        My.Settings.BadFileTypeWarning = dgvWarning.Rows(2).Cells("Notification").Value
    End Sub
End Class