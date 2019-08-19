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
        Me.ImportSalesAnalysisReportToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.ImportExportSalesAnalysisReportToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.ImportAccountingDataToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.ImportToolStripMenuItem1 = New System.Windows.Forms.ToolStripMenuItem()
        Me.ExportToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.CashNettingToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.MasteToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.BPToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.ReportToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.ExceptionReportToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.SupplierExceptionToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.SAPExceptionToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.MatchedReportToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.SupplierMatchedToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.SAPMatchedToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.OriginalDataToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.SAPRawDataToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.MenuStrip1.SuspendLayout()
        Me.SuspendLayout()
        '
        'MenuStrip1
        '
        Me.MenuStrip1.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.ImportToolStripMenuItem, Me.ExportToolStripMenuItem, Me.MasteToolStripMenuItem, Me.ReportToolStripMenuItem})
        Me.MenuStrip1.Location = New System.Drawing.Point(0, 0)
        Me.MenuStrip1.Name = "MenuStrip1"
        Me.MenuStrip1.Size = New System.Drawing.Size(620, 24)
        Me.MenuStrip1.TabIndex = 0
        Me.MenuStrip1.Text = "MenuStrip1"
        '
        'ImportToolStripMenuItem
        '
        Me.ImportToolStripMenuItem.DropDownItems.AddRange(New System.Windows.Forms.ToolStripItem() {Me.ImportSalesAnalysisReportToolStripMenuItem, Me.ImportExportSalesAnalysisReportToolStripMenuItem, Me.ImportAccountingDataToolStripMenuItem, Me.ImportToolStripMenuItem1})
        Me.ImportToolStripMenuItem.Name = "ImportToolStripMenuItem"
        Me.ImportToolStripMenuItem.Size = New System.Drawing.Size(81, 20)
        Me.ImportToolStripMenuItem.Text = "Transaction"
        '
        'ImportSalesAnalysisReportToolStripMenuItem
        '
        Me.ImportSalesAnalysisReportToolStripMenuItem.Name = "ImportSalesAnalysisReportToolStripMenuItem"
        Me.ImportSalesAnalysisReportToolStripMenuItem.ShortcutKeyDisplayString = ""
        Me.ImportSalesAnalysisReportToolStripMenuItem.Size = New System.Drawing.Size(259, 22)
        Me.ImportSalesAnalysisReportToolStripMenuItem.Tag = "FormImportSalesAnalysis"
        Me.ImportSalesAnalysisReportToolStripMenuItem.Text = "Import Sales Analysis Report"
        Me.ImportSalesAnalysisReportToolStripMenuItem.TextAlign = System.Drawing.ContentAlignment.BottomLeft
        Me.ImportSalesAnalysisReportToolStripMenuItem.Visible = False
        '
        'ImportExportSalesAnalysisReportToolStripMenuItem
        '
        Me.ImportExportSalesAnalysisReportToolStripMenuItem.Name = "ImportExportSalesAnalysisReportToolStripMenuItem"
        Me.ImportExportSalesAnalysisReportToolStripMenuItem.Size = New System.Drawing.Size(259, 22)
        Me.ImportExportSalesAnalysisReportToolStripMenuItem.Tag = "FormImportExportSalesAnalysis"
        Me.ImportExportSalesAnalysisReportToolStripMenuItem.Text = "Import Export Sales Analysis Report"
        '
        'ImportAccountingDataToolStripMenuItem
        '
        Me.ImportAccountingDataToolStripMenuItem.Name = "ImportAccountingDataToolStripMenuItem"
        Me.ImportAccountingDataToolStripMenuItem.Size = New System.Drawing.Size(259, 22)
        Me.ImportAccountingDataToolStripMenuItem.Text = "Import SAP Invoice (037)"
        '
        'ImportToolStripMenuItem1
        '
        Me.ImportToolStripMenuItem1.Name = "ImportToolStripMenuItem1"
        Me.ImportToolStripMenuItem1.Size = New System.Drawing.Size(259, 22)
        Me.ImportToolStripMenuItem1.Text = "Import Supplier Invoice"
        '
        'ExportToolStripMenuItem
        '
        Me.ExportToolStripMenuItem.DropDownItems.AddRange(New System.Windows.Forms.ToolStripItem() {Me.CashNettingToolStripMenuItem})
        Me.ExportToolStripMenuItem.Name = "ExportToolStripMenuItem"
        Me.ExportToolStripMenuItem.Size = New System.Drawing.Size(52, 20)
        Me.ExportToolStripMenuItem.Text = "Export"
        Me.ExportToolStripMenuItem.Visible = False
        '
        'CashNettingToolStripMenuItem
        '
        Me.CashNettingToolStripMenuItem.Name = "CashNettingToolStripMenuItem"
        Me.CashNettingToolStripMenuItem.Size = New System.Drawing.Size(143, 22)
        Me.CashNettingToolStripMenuItem.Tag = "FormCashNetting"
        Me.CashNettingToolStripMenuItem.Text = "Cash Netting"
        '
        'MasteToolStripMenuItem
        '
        Me.MasteToolStripMenuItem.DropDownItems.AddRange(New System.Windows.Forms.ToolStripItem() {Me.BPToolStripMenuItem})
        Me.MasteToolStripMenuItem.Name = "MasteToolStripMenuItem"
        Me.MasteToolStripMenuItem.Size = New System.Drawing.Size(55, 20)
        Me.MasteToolStripMenuItem.Text = "Master"
        '
        'BPToolStripMenuItem
        '
        Me.BPToolStripMenuItem.Name = "BPToolStripMenuItem"
        Me.BPToolStripMenuItem.Size = New System.Drawing.Size(127, 22)
        Me.BPToolStripMenuItem.Tag = "FormBusinessPartner"
        Me.BPToolStripMenuItem.Text = "Master BP"
        '
        'ReportToolStripMenuItem
        '
        Me.ReportToolStripMenuItem.DropDownItems.AddRange(New System.Windows.Forms.ToolStripItem() {Me.ExceptionReportToolStripMenuItem, Me.MatchedReportToolStripMenuItem, Me.OriginalDataToolStripMenuItem})
        Me.ReportToolStripMenuItem.Name = "ReportToolStripMenuItem"
        Me.ReportToolStripMenuItem.Size = New System.Drawing.Size(54, 20)
        Me.ReportToolStripMenuItem.Text = "Report"
        '
        'ExceptionReportToolStripMenuItem
        '
        Me.ExceptionReportToolStripMenuItem.DropDownItems.AddRange(New System.Windows.Forms.ToolStripItem() {Me.SupplierExceptionToolStripMenuItem, Me.SAPExceptionToolStripMenuItem})
        Me.ExceptionReportToolStripMenuItem.Name = "ExceptionReportToolStripMenuItem"
        Me.ExceptionReportToolStripMenuItem.Size = New System.Drawing.Size(163, 22)
        Me.ExceptionReportToolStripMenuItem.Text = "Exception Report"
        '
        'SupplierExceptionToolStripMenuItem
        '
        Me.SupplierExceptionToolStripMenuItem.Name = "SupplierExceptionToolStripMenuItem"
        Me.SupplierExceptionToolStripMenuItem.Size = New System.Drawing.Size(171, 22)
        Me.SupplierExceptionToolStripMenuItem.Text = "Supplier Exception"
        '
        'SAPExceptionToolStripMenuItem
        '
        Me.SAPExceptionToolStripMenuItem.Name = "SAPExceptionToolStripMenuItem"
        Me.SAPExceptionToolStripMenuItem.Size = New System.Drawing.Size(171, 22)
        Me.SAPExceptionToolStripMenuItem.Text = "SAP Exception"
        '
        'MatchedReportToolStripMenuItem
        '
        Me.MatchedReportToolStripMenuItem.DropDownItems.AddRange(New System.Windows.Forms.ToolStripItem() {Me.SupplierMatchedToolStripMenuItem, Me.SAPMatchedToolStripMenuItem})
        Me.MatchedReportToolStripMenuItem.Name = "MatchedReportToolStripMenuItem"
        Me.MatchedReportToolStripMenuItem.Size = New System.Drawing.Size(163, 22)
        Me.MatchedReportToolStripMenuItem.Text = "Matched Report"
        '
        'SupplierMatchedToolStripMenuItem
        '
        Me.SupplierMatchedToolStripMenuItem.Name = "SupplierMatchedToolStripMenuItem"
        Me.SupplierMatchedToolStripMenuItem.Size = New System.Drawing.Size(167, 22)
        Me.SupplierMatchedToolStripMenuItem.Text = "Supplier Matched"
        '
        'SAPMatchedToolStripMenuItem
        '
        Me.SAPMatchedToolStripMenuItem.Name = "SAPMatchedToolStripMenuItem"
        Me.SAPMatchedToolStripMenuItem.Size = New System.Drawing.Size(167, 22)
        Me.SAPMatchedToolStripMenuItem.Text = "SAP Matched"
        '
        'OriginalDataToolStripMenuItem
        '
        Me.OriginalDataToolStripMenuItem.DropDownItems.AddRange(New System.Windows.Forms.ToolStripItem() {Me.SAPRawDataToolStripMenuItem})
        Me.OriginalDataToolStripMenuItem.Name = "OriginalDataToolStripMenuItem"
        Me.OriginalDataToolStripMenuItem.Size = New System.Drawing.Size(163, 22)
        Me.OriginalDataToolStripMenuItem.Text = "Original Data"
        '
        'SAPRawDataToolStripMenuItem
        '
        Me.SAPRawDataToolStripMenuItem.Name = "SAPRawDataToolStripMenuItem"
        Me.SAPRawDataToolStripMenuItem.Size = New System.Drawing.Size(152, 22)
        Me.SAPRawDataToolStripMenuItem.Text = "SAP Raw Data"
        '
        'FormMenu
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(620, 91)
        Me.Controls.Add(Me.MenuStrip1)
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.Location = New System.Drawing.Point(100, 100)
        Me.MainMenuStrip = Me.MenuStrip1
        Me.Name = "FormMenu"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.Manual
        Me.Text = "FormMenu"
        Me.MenuStrip1.ResumeLayout(False)
        Me.MenuStrip1.PerformLayout()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents MenuStrip1 As System.Windows.Forms.MenuStrip
    Friend WithEvents ImportToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents ImportSalesAnalysisReportToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents ExportToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents MasteToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents CashNettingToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents BPToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents ImportExportSalesAnalysisReportToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents ImportAccountingDataToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents ImportToolStripMenuItem1 As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents ReportToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents ExceptionReportToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents SupplierExceptionToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents SAPExceptionToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents MatchedReportToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents SupplierMatchedToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents SAPMatchedToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents OriginalDataToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents SAPRawDataToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
End Class
