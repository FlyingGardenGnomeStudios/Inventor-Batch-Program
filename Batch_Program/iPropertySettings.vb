Imports System.Windows.Forms

Public Class iPropertySettings
    Public Sub New()

        ' This call is required by the designer.
        InitializeComponent()
        DGVSummary.Rows.Add("Title")
        DGVSummary.Rows.Add("Subject")
        DGVSummary.Rows.Add("Author")
        DGVSummary.Rows.Add("Manager")
        DGVSummary.Rows.Add("Company")
        DGVSummary.Rows.Add("Category")
        DGVSummary.Rows.Add("Keywords")
        DGVSummary.Rows.Add("Comments")
        DGVProject.Rows.Add("Location")
        DGVProject.Rows.Add("File Subtype")
        DGVProject.Rows.Add("Part Number")
        DGVProject.Rows.Add("Stock Number")
        DGVProject.Rows.Add("Description")
        DGVProject.Rows.Add("Revision Number")
        DGVProject.Rows.Add("Project")
        DGVProject.Rows.Add("Designer")
        DGVProject.Rows.Add("Engineer")
        DGVProject.Rows.Add("Vendor")
        DGVStatus.Rows.Add("Part Number")
        DGVStatus.Rows.Add("Stock Number")
        DGVStatus.Rows.Add("Status")
        DGVStatus.Rows.Add("Model Date")
        DGVStatus.Rows.Add("Drawing Date")
        DGVStatus.Rows.Add("Drawing By")
        DGVStatus.Rows.Add("Design State")
        DGVStatus.Rows.Add("Checked By")
        DGVStatus.Rows.Add("Checked Date")
        DGVCustom.Rows.Add("Custom")
        DGVSummary.Rows(0).Cells("Reference").Value = My.Settings.Title
        DGVSummary.Rows(1).Cells("Reference").Value = My.Settings.Subject
        DGVSummary.Rows(2).Cells("Reference").Value = My.Settings.Author
        DGVSummary.Rows(3).Cells("Reference").Value = My.Settings.Manager
        DGVSummary.Rows(4).Cells("Reference").Value = My.Settings.Company
        DGVSummary.Rows(5).Cells("Reference").Value = My.Settings.Category
        DGVSummary.Rows(6).Cells("Reference").Value = My.Settings.Keywords
        DGVSummary.Rows(7).Cells("Reference").Value = My.Settings.Comments
        DGVProject.Rows(0).Cells("PReference").Value = My.Settings.Location
        DGVProject.Rows(1).Cells("PReference").Value = My.Settings.Subtype
        DGVProject.Rows(2).Cells("PReference").Value = My.Settings.PPartNumber
        DGVProject.Rows(3).Cells("PReference").Value = My.Settings.PStockNumber
        DGVProject.Rows(4).Cells("PReference").Value = My.Settings.Description
        DGVProject.Rows(5).Cells("PReference").Value = My.Settings.Revision
        DGVProject.Rows(6).Cells("PReference").Value = My.Settings.Project
        DGVProject.Rows(7).Cells("PReference").Value = My.Settings.Designer
        DGVProject.Rows(8).Cells("PReference").Value = My.Settings.Engineer
        DGVProject.Rows(9).Cells("PReference").Value = My.Settings.Vendor
        DGVStatus.Rows(0).Cells("SReference").Value = My.Settings.SPartNumber
        DGVStatus.Rows(1).Cells("SReference").Value = My.Settings.SStockNumber
        DGVStatus.Rows(2).Cells("SReference").Value = My.Settings.Status
        DGVStatus.Rows(3).Cells("SReference").Value = My.Settings.ModelDate
        DGVStatus.Rows(4).Cells("SReference").Value = My.Settings.DrawingDate
        DGVStatus.Rows(5).Cells("SReference").Value = My.Settings.DrawingBy
        DGVStatus.Rows(6).Cells("SReference").Value = My.Settings.DesignState
        DGVStatus.Rows(7).Cells("SReference").Value = My.Settings.CheckedBy
        DGVStatus.Rows(8).Cells("SReference").Value = My.Settings.CheckedDate
        DGVCustom.Rows(0).Cells("CReference").Value = My.Settings.Custom
        cmbDefault.Text = My.Settings.DupName

        ' Add any initialization after the InitializeComponent() call.

    End Sub
    Private Sub btnApply_Click(sender As Object, e As EventArgs) Handles btnApply.Click
        SaveSettings()
    End Sub
    Private Sub SaveSettings()
        My.Settings.Title = DGVSummary.Rows(0).Cells("Reference").Value
        My.Settings.Subject = DGVSummary.Rows(1).Cells("Reference").Value
        My.Settings.Author = DGVSummary.Rows(2).Cells("Reference").Value
        My.Settings.Manager = DGVSummary.Rows(3).Cells("Reference").Value
        My.Settings.Company = DGVSummary.Rows(4).Cells("Reference").Value
        My.Settings.Category = DGVSummary.Rows(5).Cells("Reference").Value
        My.Settings.Keywords = DGVSummary.Rows(6).Cells("Reference").Value
        My.Settings.Comments = DGVSummary.Rows(7).Cells("Reference").Value
        My.Settings.Location = DGVProject.Rows(0).Cells("PReference").Value
        My.Settings.Subtype = DGVProject.Rows(1).Cells("PReference").Value
        My.Settings.PPartNumber = DGVProject.Rows(2).Cells("PReference").Value
        My.Settings.PStockNumber = DGVProject.Rows(3).Cells("PReference").Value
        My.Settings.Description = DGVProject.Rows(4).Cells("PReference").Value
        My.Settings.Revision = DGVProject.Rows(5).Cells("PReference").Value
        My.Settings.Project = DGVProject.Rows(6).Cells("PReference").Value
        My.Settings.Designer = DGVProject.Rows(7).Cells("PReference").Value
        My.Settings.Engineer = DGVProject.Rows(8).Cells("PReference").Value
        My.Settings.Vendor = DGVProject.Rows(9).Cells("PReference").Value
        My.Settings.SPartNumber = DGVStatus.Rows(0).Cells("SReference").Value
        My.Settings.SStockNumber = DGVStatus.Rows(1).Cells("SReference").Value
        My.Settings.Status = DGVStatus.Rows(2).Cells("SReference").Value
        My.Settings.ModelDate = DGVStatus.Rows(3).Cells("SReference").Value
        My.Settings.DrawingDate = DGVStatus.Rows(4).Cells("SReference").Value
        My.Settings.DrawingBy = DGVStatus.Rows(5).Cells("SReference").Value
        My.Settings.DesignState = DGVStatus.Rows(6).Cells("SReference").Value
        My.Settings.CheckedBy = DGVStatus.Rows(7).Cells("SReference").Value
        My.Settings.CheckedDate = DGVStatus.Rows(8).Cells("SReference").Value
        My.Settings.Custom = DGVCustom.Rows(0).Cells("CReference").Value
        My.Settings.DupName = cmbDefault.Text

    End Sub

    Private Sub btnCancel_Click(sender As Object, e As EventArgs) Handles btnCancel.Click
        Me.Close()
    End Sub

    Private Sub btnOK_Click(sender As Object, e As EventArgs) Handles btnOK.Click
        SaveSettings()
        Me.Close()
    End Sub

    Private Sub DGVSummary_CellEnter(sender As Object, e As DataGridViewCellEventArgs) Handles DGVSummary.CellEnter
        SendKeys.Send("{F4}")
    End Sub
    Private Sub DGVProject_CellEnter(sender As Object, e As DataGridViewCellEventArgs) Handles DGVProject.CellEnter
        SendKeys.Send("{F4}")
    End Sub
    Private Sub DGVStatus_CellEnter(sender As Object, e As DataGridViewCellEventArgs) Handles DGVStatus.CellEnter
        SendKeys.Send("{F4}")
    End Sub

    Private Sub btnDefaults_Click(sender As Object, e As EventArgs) Handles btnDefaults.Click
        DGVSummary.Rows(0).Cells("Reference").Value = My.Settings.PropertyValues("Title").Property.DefaultValue
        DGVSummary.Rows(1).Cells("Reference").Value = My.Settings.PropertyValues("Subject").Property.DefaultValue
        DGVSummary.Rows(2).Cells("Reference").Value = My.Settings.PropertyValues("Author").Property.DefaultValue
        DGVSummary.Rows(3).Cells("Reference").Value = My.Settings.PropertyValues("Manager").Property.DefaultValue
        DGVSummary.Rows(4).Cells("Reference").Value = My.Settings.PropertyValues("Company").Property.DefaultValue
        DGVSummary.Rows(5).Cells("Reference").Value = My.Settings.PropertyValues("Category").Property.DefaultValue
        DGVSummary.Rows(6).Cells("Reference").Value = My.Settings.PropertyValues("Keywords").Property.DefaultValue
        DGVSummary.Rows(7).Cells("Reference").Value = My.Settings.PropertyValues("Comments").Property.DefaultValue
        DGVProject.Rows(0).Cells("PReference").Value = My.Settings.PropertyValues("Location").Property.DefaultValue
        DGVProject.Rows(1).Cells("PReference").Value = My.Settings.PropertyValues("Subtype").Property.DefaultValue
        DGVProject.Rows(2).Cells("PReference").Value = My.Settings.PropertyValues("PPartNumber").Property.DefaultValue
        DGVProject.Rows(3).Cells("PReference").Value = My.Settings.PropertyValues("PStockNumber").Property.DefaultValue
        DGVProject.Rows(4).Cells("PReference").Value = My.Settings.PropertyValues("Description").Property.DefaultValue
        DGVProject.Rows(5).Cells("PReference").Value = My.Settings.PropertyValues("Revision").Property.DefaultValue
        DGVProject.Rows(6).Cells("PReference").Value = My.Settings.PropertyValues("Project").Property.DefaultValue
        DGVProject.Rows(7).Cells("PReference").Value = My.Settings.PropertyValues("Designer").Property.DefaultValue
        DGVProject.Rows(8).Cells("PReference").Value = My.Settings.PropertyValues("Engineer").Property.DefaultValue
        DGVProject.Rows(9).Cells("PReference").Value = My.Settings.PropertyValues("Vendor").Property.DefaultValue
        DGVStatus.Rows(0).Cells("SReference").Value = My.Settings.PropertyValues("SPartNumber").Property.DefaultValue
        DGVStatus.Rows(1).Cells("SReference").Value = My.Settings.PropertyValues("SStockNumber").Property.DefaultValue
        DGVStatus.Rows(2).Cells("SReference").Value = My.Settings.PropertyValues("Status").Property.DefaultValue
        DGVStatus.Rows(3).Cells("SReference").Value = My.Settings.PropertyValues("ModelDate").Property.DefaultValue
        DGVStatus.Rows(4).Cells("SReference").Value = My.Settings.PropertyValues("DrawingDate").Property.DefaultValue
        DGVStatus.Rows(5).Cells("SReference").Value = My.Settings.PropertyValues("DrawingBy").Property.DefaultValue
        DGVStatus.Rows(6).Cells("SReference").Value = My.Settings.PropertyValues("DesignState").Property.DefaultValue
        DGVStatus.Rows(7).Cells("SReference").Value = My.Settings.PropertyValues("CheckedBy").Property.DefaultValue
        DGVStatus.Rows(8).Cells("SReference").Value = My.Settings.PropertyValues("CheckedDate").Property.DefaultValue
        DGVCustom.Rows(0).Cells("CReference").Value = My.Settings.PropertyValues("Custom").Property.DefaultValue
        cmbDefault.Text = My.Settings.PropertyValues("DupName").Property.DefaultValue
    End Sub
End Class