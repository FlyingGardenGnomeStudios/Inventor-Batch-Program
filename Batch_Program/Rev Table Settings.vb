Imports System.Windows.Forms

Public Class Rev_Table_Settings
    Private RowIndex As Integer = 0
    Public Sub New()

        ' This call is required by the designer.
        InitializeComponent()
        dgvRevTableLayout.Rows.Add(My.Settings.RTSCheckedBy, "Checked By", "Sheet iProperty", "Text")
        dgvRevTableLayout.Rows.Add(My.Settings.RTSCheckedDate, "Check Date", "Sheet iProperty", "Date")
        dgvRevTableLayout.Rows.Add(My.Settings.RTSRev, "Revision Number", My.Settings.RTSRevCol, "Number")
        dgvRevTableLayout.Rows.Add(My.Settings.RTSDate, "Revision Date", My.Settings.RTSDateCol, "Date")
        dgvRevTableLayout.Rows.Add(My.Settings.RTSDesc, "Revision Description", My.Settings.RTSDescCol, "Text")
        dgvRevTableLayout.Rows.Add(My.Settings.RTSName, "Revision By", My.Settings.RTSNameCol, "Text")
        dgvRevTableLayout.Rows.Add(My.Settings.RTSApproved, "Revision Approved", My.Settings.RTSApprovedCol, "Text")
        For column = 0 To dgvRevTableLayout.Columns.Count - 1
            For Row = 0 To 1
                If column <> 0 Then
                    dgvRevTableLayout(column, Row).ReadOnly = True
                    dgvRevTableLayout(column, Row).Style.BackColor = Drawing.Color.LightGray
                End If
            Next
            For row = 2 To dgvRevTableLayout.Rows.Count - 2
                If column <> 0 AndAlso column <> 2 Then
                    dgvRevTableLayout(column, row).ReadOnly = True
                    dgvRevTableLayout(column, row).Style.BackColor = Drawing.Color.LightGray
                End If
            Next
        Next
        If My.Settings.RTS1Item <> "" Then dgvRevTableLayout.Rows.Add(My.Settings.RTS1, My.Settings.RTS1Item, My.Settings.RTS1Col, My.Settings.RTS1Value)
        If My.Settings.RTS2Item <> "" Then dgvRevTableLayout.Rows.Add(My.Settings.RTS2, My.Settings.RTS2Item, My.Settings.RTS2Col, My.Settings.RTS2Value)
        If My.Settings.RTS3Item <> "" Then dgvRevTableLayout.Rows.Add(My.Settings.RTS3, My.Settings.RTS3Item, My.Settings.RTS3Col, My.Settings.RTS3Value)
        If My.Settings.RTS4Item <> "" Then dgvRevTableLayout.Rows.Add(My.Settings.RTS4, My.Settings.RTS4Item, My.Settings.RTS4Col, My.Settings.RTS4Value)
        If My.Settings.RTS5Item <> "" Then dgvRevTableLayout.Rows.Add(My.Settings.RTS5, My.Settings.RTS5Item, My.Settings.RTS5Col, My.Settings.RTS5Value)
        txtNumRev.Text = My.Settings.NumRev
        txtAlphaRev.Text = My.Settings.AlphaRev
        nudStartVal.Value = My.Settings.StartVal
        Select Case My.Settings.DefRevLoc
            Case "BL"
                rdbBL.Checked = True
            Case "BR"
                rdbBR.Checked = True
            Case "TL"
                rdbTL.Checked = True
            Case "TR"
                rdbTR.Checked = True
        End Select
    End Sub


    Private Sub btnOK_Click(sender As Object, e As EventArgs) Handles btnOK.Click
        For Row = 2 To 10
            Select Case Row - 2
                Case 0
                    My.Settings.RTSRev = dgvRevTableLayout(0, Row).Value
                    My.Settings.RTSRevCol = dgvRevTableLayout(2, Row).Value
                Case 1
                    My.Settings.RTSDate = dgvRevTableLayout(0, Row).Value
                    My.Settings.RTSDateCol = dgvRevTableLayout(2, Row).Value
                Case 2
                    My.Settings.RTSDesc = dgvRevTableLayout(0, Row).Value
                    My.Settings.RTSDescCol = dgvRevTableLayout(2, Row).Value
                Case 3
                    My.Settings.RTSName = dgvRevTableLayout(0, Row).Value
                    My.Settings.RTSNameCol = dgvRevTableLayout(2, Row).Value
                Case 4
                    My.Settings.RTSApproved = dgvRevTableLayout(0, Row).Value
                    My.Settings.RTSApprovedCol = dgvRevTableLayout(2, Row).Value
                Case 5
                    Try
                        My.Settings.RTS1 = dgvRevTableLayout(0, Row).Value
                        My.Settings.RTS1Item = dgvRevTableLayout(1, Row).Value
                        My.Settings.RTS1Col = dgvRevTableLayout(2, Row).Value
                        My.Settings.RTS1Value = dgvRevTableLayout(3, Row).Value
                    Catch
                        My.Settings.RTS1 = False
                        My.Settings.RTS1Item = Nothing
                        My.Settings.RTS1Col = Nothing
                        My.Settings.RTS1Value = Nothing
                    End Try
                Case 6
                    Try
                        My.Settings.RTS2 = dgvRevTableLayout(0, Row).Value
                        My.Settings.RTS2Item = dgvRevTableLayout(1, Row).Value
                        My.Settings.RTS2Col = dgvRevTableLayout(2, Row).Value
                        My.Settings.RTS2Value = dgvRevTableLayout(3, Row).Value
                    Catch
                        My.Settings.RTS2 = False
                        My.Settings.RTS2Item = Nothing
                        My.Settings.RTS2Col = Nothing
                        My.Settings.RTS2Value = Nothing
                    End Try
                Case 7
                    Try
                        My.Settings.RTS3 = dgvRevTableLayout(0, Row).Value
                        My.Settings.RTS3Item = dgvRevTableLayout(1, Row).Value
                        My.Settings.RTS3Col = dgvRevTableLayout(2, Row).Value
                        My.Settings.RTS3Value = dgvRevTableLayout(3, Row).Value
                    Catch
                        My.Settings.RTS3 = False
                        My.Settings.RTS3Item = Nothing
                        My.Settings.RTS3Col = Nothing
                        My.Settings.RTS3Value = Nothing
                    End Try
                Case 8
                    Try
                        My.Settings.RTS4 = dgvRevTableLayout(0, Row).Value
                        My.Settings.RTS4Item = dgvRevTableLayout(1, Row).Value
                        My.Settings.RTS4Col = dgvRevTableLayout(2, Row).Value
                        My.Settings.RTS4Value = dgvRevTableLayout(3, Row).Value
                    Catch
                        My.Settings.RTS4 = False
                        My.Settings.RTS4Item = Nothing
                        My.Settings.RTS4Col = Nothing
                        My.Settings.RTS4Value = Nothing
                    End Try
                Case 9
                    Try
                        My.Settings.RTS5 = dgvRevTableLayout(0, Row).Value
                        My.Settings.RTS5Item = dgvRevTableLayout(1, Row).Value
                        My.Settings.RTS5Col = dgvRevTableLayout(2, Row).Value
                        My.Settings.RTS5Value = dgvRevTableLayout(3, Row).Value
                    Catch
                        My.Settings.RTS5 = False
                        My.Settings.RTS5Item = Nothing
                        My.Settings.RTS5Col = Nothing
                        My.Settings.RTS5Value = Nothing
                    End Try
                Case 10
                    MsgBox("Batch Program currently only supports 5 custom fields")
            End Select
        Next
        If txtAlphaRev.Text <> "" Then
            My.Settings.AlphaRev = txtAlphaRev.Text
        Else
            MsgBox("Please enter a default description for alphabetical revisions")
            Exit Sub
        End If
        If txtNumRev.Text <> "" Then
            My.Settings.NumRev = txtNumRev.Text
        Else
            MsgBox("Please enter a default description for numerical revisions")
            Exit Sub
        End If
        My.Settings.StartVal = nudStartVal.Value
        My.Settings.Save()
        Me.Close()
        If rdbBL.Checked = True Then
            My.Settings.DefRevLoc = "BL"
        ElseIf rdbBR.Checked = True Then
            My.Settings.DefRevLoc = "BR"
        ElseIf rdbTL.Checked = True Then
            My.Settings.DefRevLoc = "TL"
        ElseIf rdbTR.Checked = True Then
            My.Settings.DefRevLoc = "TR"
        End If
    End Sub

    Private Sub btnCancel_Click(sender As Object, e As EventArgs) Handles btnCancel.Click
        Me.Close()
    End Sub
    Private Sub ContextMenuStrip1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ContextMenuStrip1.Click
        Try
            If dgvRevTableLayout.CurrentCell.RowIndex > 6 Then
                dgvRevTableLayout.Rows.RemoveAt(dgvRevTableLayout.CurrentCell.RowIndex)
            Else
                MsgBox("Cannot remove default item." & vbNewLine & "Uncheck this item if it does not apply")
            End If
        Catch
        End Try
    End Sub
    Private Sub dgvRevTableLayout_CellMouseClick(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellMouseEventArgs) Handles dgvRevTableLayout.CellMouseClick
        'Make sure the type is set to something, a work around if the grid contains checkboxes.
        If e.RowIndex > -1 And e.ColumnIndex > -1 Then
            If dgvRevTableLayout(e.ColumnIndex, e.RowIndex).EditType IsNot Nothing Then
                'Check if the control is a combo box if so, edit on enter
                If dgvRevTableLayout(e.ColumnIndex, e.RowIndex).EditType.ToString() = "System.Windows.Forms.DataGridViewComboBoxEditingControl" Then
                    SendKeys.Send("{F4}")
                End If
            End If
        End If
    End Sub
End Class