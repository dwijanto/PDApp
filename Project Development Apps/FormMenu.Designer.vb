<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class FormMenu
    Inherits System.Windows.Forms.Form

    'Form overrides dispose to clean up the component list.
    <System.Diagnostics.DebuggerNonUserCode()> _
    Protected Overrides Sub Dispose(ByVal disposing As Boolean)
        Try
            If disposing AndAlso components IsNot Nothing Then
                components.Dispose()
            End If
        Finally
            MyBase.Dispose(disposing)
        End Try
    End Sub

    'Required by the Windows Form Designer
    Private components As System.ComponentModel.IContainer

    'NOTE: The following procedure is required by the Windows Form Designer
    'It can be modified using the Windows Form Designer.  
    'Do not modify it using the code editor.
    <System.Diagnostics.DebuggerStepThrough()> _
    Private Sub InitializeComponent()
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(FormMenu))
        Me.MenuStrip1 = New System.Windows.Forms.MenuStrip()
        Me.ImportToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.ProjectStageStatusToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.ProjectMaintenanceToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.ClosingPeriodToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.MasterToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.UserToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.VendorToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.FailedStatusAdjustmentToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.ShowOpenUserToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.GeneratePeriodToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.SaveHistoryToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.ReportToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.RawDataToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.HelpToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.GuidelineToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.MenuStrip1.SuspendLayout()
        Me.SuspendLayout()
        '
        'MenuStrip1
        '
        Me.MenuStrip1.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.ImportToolStripMenuItem, Me.MasterToolStripMenuItem, Me.ReportToolStripMenuItem, Me.HelpToolStripMenuItem})
        Me.MenuStrip1.Location = New System.Drawing.Point(0, 0)
        Me.MenuStrip1.Name = "MenuStrip1"
        Me.MenuStrip1.Size = New System.Drawing.Size(603, 24)
        Me.MenuStrip1.TabIndex = 2
        Me.MenuStrip1.Text = "MenuStrip1"
        '
        'ImportToolStripMenuItem
        '
        Me.ImportToolStripMenuItem.DropDownItems.AddRange(New System.Windows.Forms.ToolStripItem() {Me.ProjectStageStatusToolStripMenuItem, Me.ProjectMaintenanceToolStripMenuItem, Me.ClosingPeriodToolStripMenuItem})
        Me.ImportToolStripMenuItem.Name = "ImportToolStripMenuItem"
        Me.ImportToolStripMenuItem.Size = New System.Drawing.Size(80, 20)
        Me.ImportToolStripMenuItem.Text = "Transaction"
        '
        'ProjectStageStatusToolStripMenuItem
        '
        Me.ProjectStageStatusToolStripMenuItem.Name = "ProjectStageStatusToolStripMenuItem"
        Me.ProjectStageStatusToolStripMenuItem.Size = New System.Drawing.Size(183, 22)
        Me.ProjectStageStatusToolStripMenuItem.Tag = "FormProjectSigned"
        Me.ProjectStageStatusToolStripMenuItem.Text = "Project Stage Status"
        '
        'ProjectMaintenanceToolStripMenuItem
        '
        Me.ProjectMaintenanceToolStripMenuItem.Name = "ProjectMaintenanceToolStripMenuItem"
        Me.ProjectMaintenanceToolStripMenuItem.Size = New System.Drawing.Size(183, 22)
        Me.ProjectMaintenanceToolStripMenuItem.Tag = "FormProjectMaintenance"
        Me.ProjectMaintenanceToolStripMenuItem.Text = "Project Maintenance"
        '
        'ClosingPeriodToolStripMenuItem
        '
        Me.ClosingPeriodToolStripMenuItem.Name = "ClosingPeriodToolStripMenuItem"
        Me.ClosingPeriodToolStripMenuItem.Size = New System.Drawing.Size(183, 22)
        Me.ClosingPeriodToolStripMenuItem.Text = "Closing Period"
        '
        'MasterToolStripMenuItem
        '
        Me.MasterToolStripMenuItem.DropDownItems.AddRange(New System.Windows.Forms.ToolStripItem() {Me.UserToolStripMenuItem, Me.VendorToolStripMenuItem, Me.FailedStatusAdjustmentToolStripMenuItem, Me.ShowOpenUserToolStripMenuItem, Me.GeneratePeriodToolStripMenuItem, Me.SaveHistoryToolStripMenuItem})
        Me.MasterToolStripMenuItem.Name = "MasterToolStripMenuItem"
        Me.MasterToolStripMenuItem.Size = New System.Drawing.Size(55, 20)
        Me.MasterToolStripMenuItem.Text = "Admin"
        '
        'UserToolStripMenuItem
        '
        Me.UserToolStripMenuItem.Name = "UserToolStripMenuItem"
        Me.UserToolStripMenuItem.Size = New System.Drawing.Size(205, 22)
        Me.UserToolStripMenuItem.Tag = "FormUser"
        Me.UserToolStripMenuItem.Text = "User && Security"
        '
        'VendorToolStripMenuItem
        '
        Me.VendorToolStripMenuItem.Name = "VendorToolStripMenuItem"
        Me.VendorToolStripMenuItem.Size = New System.Drawing.Size(205, 22)
        Me.VendorToolStripMenuItem.Tag = "FormVendorV2"
        Me.VendorToolStripMenuItem.Text = "Vendor"
        '
        'FailedStatusAdjustmentToolStripMenuItem
        '
        Me.FailedStatusAdjustmentToolStripMenuItem.Name = "FailedStatusAdjustmentToolStripMenuItem"
        Me.FailedStatusAdjustmentToolStripMenuItem.Size = New System.Drawing.Size(205, 22)
        Me.FailedStatusAdjustmentToolStripMenuItem.Tag = "FormFailedStatusAdjustment"
        Me.FailedStatusAdjustmentToolStripMenuItem.Text = "Failed Status Adjustment"
        '
        'ShowOpenUserToolStripMenuItem
        '
        Me.ShowOpenUserToolStripMenuItem.Name = "ShowOpenUserToolStripMenuItem"
        Me.ShowOpenUserToolStripMenuItem.Size = New System.Drawing.Size(205, 22)
        Me.ShowOpenUserToolStripMenuItem.Tag = "DialogOpenUser"
        Me.ShowOpenUserToolStripMenuItem.Text = "Show Open User"
        '
        'GeneratePeriodToolStripMenuItem
        '
        Me.GeneratePeriodToolStripMenuItem.Name = "GeneratePeriodToolStripMenuItem"
        Me.GeneratePeriodToolStripMenuItem.Size = New System.Drawing.Size(205, 22)
        Me.GeneratePeriodToolStripMenuItem.Tag = "DialogExtendPeriod"
        Me.GeneratePeriodToolStripMenuItem.Text = "Generate Period"
        Me.GeneratePeriodToolStripMenuItem.Visible = False
        '
        'SaveHistoryToolStripMenuItem
        '
        Me.SaveHistoryToolStripMenuItem.Name = "SaveHistoryToolStripMenuItem"
        Me.SaveHistoryToolStripMenuItem.Size = New System.Drawing.Size(205, 22)
        Me.SaveHistoryToolStripMenuItem.Text = "Save History"
        '
        'ReportToolStripMenuItem
        '
        Me.ReportToolStripMenuItem.DropDownItems.AddRange(New System.Windows.Forms.ToolStripItem() {Me.RawDataToolStripMenuItem})
        Me.ReportToolStripMenuItem.Name = "ReportToolStripMenuItem"
        Me.ReportToolStripMenuItem.Size = New System.Drawing.Size(54, 20)
        Me.ReportToolStripMenuItem.Text = "Report"
        '
        'RawDataToolStripMenuItem
        '
        Me.RawDataToolStripMenuItem.Name = "RawDataToolStripMenuItem"
        Me.RawDataToolStripMenuItem.Size = New System.Drawing.Size(91, 22)
        Me.RawDataToolStripMenuItem.Tag = "FormRawData"
        Me.RawDataToolStripMenuItem.Text = "KPI"
        '
        'HelpToolStripMenuItem
        '
        Me.HelpToolStripMenuItem.DropDownItems.AddRange(New System.Windows.Forms.ToolStripItem() {Me.GuidelineToolStripMenuItem})
        Me.HelpToolStripMenuItem.Name = "HelpToolStripMenuItem"
        Me.HelpToolStripMenuItem.Size = New System.Drawing.Size(44, 20)
        Me.HelpToolStripMenuItem.Text = "Help"
        '
        'GuidelineToolStripMenuItem
        '
        Me.GuidelineToolStripMenuItem.Name = "GuidelineToolStripMenuItem"
        Me.GuidelineToolStripMenuItem.Size = New System.Drawing.Size(124, 22)
        Me.GuidelineToolStripMenuItem.Text = "Guideline"
        '
        'FormMenu
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(603, 109)
        Me.Controls.Add(Me.MenuStrip1)
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.Name = "FormMenu"
        Me.Text = "FormMenu"
        Me.MenuStrip1.ResumeLayout(False)
        Me.MenuStrip1.PerformLayout()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents MenuStrip1 As System.Windows.Forms.MenuStrip
    Friend WithEvents ImportToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents MasterToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents ReportToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents RawDataToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents UserToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents VendorToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents ProjectStageStatusToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents ProjectMaintenanceToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents ClosingPeriodToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents FailedStatusAdjustmentToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents ShowOpenUserToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents GeneratePeriodToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents SaveHistoryToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents HelpToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents GuidelineToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
End Class
