<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class CaseFinder
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
        Me.components = New System.ComponentModel.Container()
        Dim DataGridViewCellStyle1 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle()
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(CaseFinder))
        Me.panelCaseDirectory = New System.Windows.Forms.Panel()
        Me.treeViewFolders = New System.Windows.Forms.TreeView()
        Me.lblNuixCaseDirectory = New System.Windows.Forms.Label()
        Me.btnGetData = New System.Windows.Forms.Button()
        Me.txtReportLocation = New System.Windows.Forms.TextBox()
        Me.lblReportLocation = New System.Windows.Forms.Label()
        Me.btnReportLocation = New System.Windows.Forms.Button()
        Me.grdCaseInfo = New System.Windows.Forms.DataGridView()
        Me.CaseGuid = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.CollectionStatus = New System.Windows.Forms.DataGridViewComboBoxColumn()
        Me.PercentComplete = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.ReportLoadDuration = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.CaseName = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.CurrentCaseVersion = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.UpgradedCaseVersion = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.BatchLoadInfo = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.DataExport = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.CaseLocation = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.BackUpLocation = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.CaseDescription = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.CaseSizeOnDisk = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.CaseFileSize = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.CaseAuditSize = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.OldestTopLevel = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.NewestTopLevel = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.IsCompound = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.CasesContained = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.ContainedInCase = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.Investigator = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.InvestigatorSessions = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.InvalidSessions = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.InvestigatorTimeSummary = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.BrokerMemory = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.WorkerCount = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.WorkerMemory = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.EvidenceName = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.EvidenceLocation = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.EvidenceDescription = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.EvidenceCustomMetadata = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.LanguagesContained = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.MimeTypes = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.ItemTypes = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.IrregularItems = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.CreationDate = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.ModifiedDate = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.LoadStartDate = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.LoadEndDate = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.LoadTime = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.LoadEvents = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.TotalLoadTime = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.ProcessingSpeed = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.TotalCaseItemCount = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.ItemCounts = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.OriginalItems = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.DuplicateItems = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.Custodians = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.CustodianCount = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.SearchTerm = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.SearchSize = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.SearchHitCount = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.CustodianSearchHit = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.HitCountPercent = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.NuixLogLocation = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.cboReportType = New System.Windows.Forms.ComboBox()
        Me.lblReportType = New System.Windows.Forms.Label()
        Me.ImgIcons = New System.Windows.Forms.ImageList(Me.components)
        Me.lblSearchTerm = New System.Windows.Forms.Label()
        Me.txtSearchTerm = New System.Windows.Forms.TextBox()
        Me.btnExportCaseStatistics = New System.Windows.Forms.Button()
        Me.ExportContextMenuStrip = New System.Windows.Forms.ContextMenuStrip(Me.components)
        Me.ExcelToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.CSVToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.JSonToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.XMLToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.lblNuixVersionLocation = New System.Windows.Forms.Label()
        Me.txtNuixConsoleLocation = New System.Windows.Forms.TextBox()
        Me.btnConsoleLocation = New System.Windows.Forms.Button()
        Me.lblNuixConsoleVersion = New System.Windows.Forms.Label()
        Me.cboNuixLicenseType = New System.Windows.Forms.ComboBox()
        Me.lblNuixVersion = New System.Windows.Forms.Label()
        Me.btnLoadPreviousReportingRun = New System.Windows.Forms.Button()
        Me.lblNuixLogDir = New System.Windows.Forms.Label()
        Me.txtNuixLogDir = New System.Windows.Forms.TextBox()
        Me.btnNuixLogSelector = New System.Windows.Forms.Button()
        Me.chkBackUpCase = New System.Windows.Forms.CheckBox()
        Me.txtBackupLocation = New System.Windows.Forms.TextBox()
        Me.lblBackUpLocation = New System.Windows.Forms.Label()
        Me.btnBackupLocationChooser = New System.Windows.Forms.Button()
        Me.grpLicenseType = New System.Windows.Forms.GroupBox()
        Me.numExportWorkers = New System.Windows.Forms.NumericUpDown()
        Me.lblExportWorkers = New System.Windows.Forms.Label()
        Me.lblExportMemory = New System.Windows.Forms.Label()
        Me.numExportWorkerMemory = New System.Windows.Forms.NumericUpDown()
        Me.lblNuixAppMemory = New System.Windows.Forms.Label()
        Me.numNuixAppMemory = New System.Windows.Forms.NumericUpDown()
        Me.lblServerType = New System.Windows.Forms.Label()
        Me.cboLicenseType = New System.Windows.Forms.ComboBox()
        Me.txtRegistryServer = New System.Windows.Forms.TextBox()
        Me.lblRegistryServer = New System.Windows.Forms.Label()
        Me.txtNMSInfo = New System.Windows.Forms.TextBox()
        Me.lblNMSInfo = New System.Windows.Forms.Label()
        Me.txtNMSUserName = New System.Windows.Forms.TextBox()
        Me.txtNMSAddress = New System.Windows.Forms.TextBox()
        Me.lblNMSAddress = New System.Windows.Forms.Label()
        Me.lblCalculateProcessingSpeeds = New System.Windows.Forms.Label()
        Me.cboCalculateProcessingSpeeds = New System.Windows.Forms.ComboBox()
        Me.chkExportSearchResults = New System.Windows.Forms.CheckBox()
        Me.btnGetFileSystemData = New System.Windows.Forms.Button()
        Me.btnSaveConfig = New System.Windows.Forms.Button()
        Me.btnLoadConfig = New System.Windows.Forms.Button()
        Me.cboSizeReporting = New System.Windows.Forms.ComboBox()
        Me.lblShowSizeIn = New System.Windows.Forms.Label()
        Me.grpCaseSelector = New System.Windows.Forms.GroupBox()
        Me.radFile = New System.Windows.Forms.RadioButton()
        Me.radFileSystem = New System.Windows.Forms.RadioButton()
        Me.txtCaseFileLocations = New System.Windows.Forms.TextBox()
        Me.btnFileLocation = New System.Windows.Forms.Button()
        Me.lblFileLocation = New System.Windows.Forms.Label()
        Me.btnStopProcessing = New System.Windows.Forms.Button()
        Me.lblExportLocation = New System.Windows.Forms.Label()
        Me.txtExportLocation = New System.Windows.Forms.TextBox()
        Me.btnExportLocation = New System.Windows.Forms.Button()
        Me.cboExportType = New System.Windows.Forms.ComboBox()
        Me.grpSearchTerm = New System.Windows.Forms.GroupBox()
        Me.radSearchFile = New System.Windows.Forms.RadioButton()
        Me.radSearchTerm = New System.Windows.Forms.RadioButton()
        Me.btnSearchTermFile = New System.Windows.Forms.Button()
        Me.chkExportOnly = New System.Windows.Forms.CheckBox()
        Me.chkRollUpReporting = New System.Windows.Forms.CheckBox()
        Me.cboCopyMoveCases = New System.Windows.Forms.ComboBox()
        Me.cboUpgradeCasees = New System.Windows.Forms.ComboBox()
        Me.lblUpgradeCases = New System.Windows.Forms.Label()
        Me.chkIncludeDiskSize = New System.Windows.Forms.CheckBox()
        Me.panelCaseDirectory.SuspendLayout()
        CType(Me.grdCaseInfo, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.ExportContextMenuStrip.SuspendLayout()
        Me.grpLicenseType.SuspendLayout()
        CType(Me.numExportWorkers, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.numExportWorkerMemory, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.numNuixAppMemory, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.grpCaseSelector.SuspendLayout()
        Me.grpSearchTerm.SuspendLayout()
        Me.SuspendLayout()
        '
        'panelCaseDirectory
        '
        Me.panelCaseDirectory.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
            Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.panelCaseDirectory.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.panelCaseDirectory.Controls.Add(Me.treeViewFolders)
        Me.panelCaseDirectory.Controls.Add(Me.lblNuixCaseDirectory)
        Me.panelCaseDirectory.Location = New System.Drawing.Point(8, 89)
        Me.panelCaseDirectory.Margin = New System.Windows.Forms.Padding(2)
        Me.panelCaseDirectory.Name = "panelCaseDirectory"
        Me.panelCaseDirectory.Size = New System.Drawing.Size(284, 423)
        Me.panelCaseDirectory.TabIndex = 37
        '
        'treeViewFolders
        '
        Me.treeViewFolders.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
            Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.treeViewFolders.BackColor = System.Drawing.SystemColors.Control
        Me.treeViewFolders.CheckBoxes = True
        Me.treeViewFolders.Location = New System.Drawing.Point(3, 15)
        Me.treeViewFolders.Margin = New System.Windows.Forms.Padding(2)
        Me.treeViewFolders.Name = "treeViewFolders"
        Me.treeViewFolders.Size = New System.Drawing.Size(276, 401)
        Me.treeViewFolders.TabIndex = 1
        '
        'lblNuixCaseDirectory
        '
        Me.lblNuixCaseDirectory.AutoSize = True
        Me.lblNuixCaseDirectory.Location = New System.Drawing.Point(2, 0)
        Me.lblNuixCaseDirectory.Margin = New System.Windows.Forms.Padding(2, 0, 2, 0)
        Me.lblNuixCaseDirectory.Name = "lblNuixCaseDirectory"
        Me.lblNuixCaseDirectory.Size = New System.Drawing.Size(100, 13)
        Me.lblNuixCaseDirectory.TabIndex = 5
        Me.lblNuixCaseDirectory.Text = "Nuix Case Directory"
        '
        'btnGetData
        '
        Me.btnGetData.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.btnGetData.Location = New System.Drawing.Point(1023, 788)
        Me.btnGetData.Margin = New System.Windows.Forms.Padding(2)
        Me.btnGetData.Name = "btnGetData"
        Me.btnGetData.Size = New System.Drawing.Size(78, 36)
        Me.btnGetData.TabIndex = 23
        Me.btnGetData.Text = "Get Data"
        Me.btnGetData.UseVisualStyleBackColor = True
        '
        'txtReportLocation
        '
        Me.txtReportLocation.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.txtReportLocation.Location = New System.Drawing.Point(123, 550)
        Me.txtReportLocation.Margin = New System.Windows.Forms.Padding(2)
        Me.txtReportLocation.Name = "txtReportLocation"
        Me.txtReportLocation.Size = New System.Drawing.Size(169, 20)
        Me.txtReportLocation.TabIndex = 9
        '
        'lblReportLocation
        '
        Me.lblReportLocation.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.lblReportLocation.AutoSize = True
        Me.lblReportLocation.Location = New System.Drawing.Point(10, 555)
        Me.lblReportLocation.Margin = New System.Windows.Forms.Padding(2, 0, 2, 0)
        Me.lblReportLocation.Name = "lblReportLocation"
        Me.lblReportLocation.Size = New System.Drawing.Size(86, 13)
        Me.lblReportLocation.TabIndex = 42
        Me.lblReportLocation.Text = "Report Location:"
        '
        'btnReportLocation
        '
        Me.btnReportLocation.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.btnReportLocation.Location = New System.Drawing.Point(296, 550)
        Me.btnReportLocation.Margin = New System.Windows.Forms.Padding(2)
        Me.btnReportLocation.Name = "btnReportLocation"
        Me.btnReportLocation.Size = New System.Drawing.Size(24, 19)
        Me.btnReportLocation.TabIndex = 10
        Me.btnReportLocation.Text = "..."
        Me.btnReportLocation.UseVisualStyleBackColor = True
        '
        'grdCaseInfo
        '
        Me.grdCaseInfo.AllowUserToAddRows = False
        Me.grdCaseInfo.AllowUserToDeleteRows = False
        Me.grdCaseInfo.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
            Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.grdCaseInfo.BackgroundColor = System.Drawing.SystemColors.Control
        Me.grdCaseInfo.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.grdCaseInfo.Columns.AddRange(New System.Windows.Forms.DataGridViewColumn() {Me.CaseGuid, Me.CollectionStatus, Me.PercentComplete, Me.ReportLoadDuration, Me.CaseName, Me.CurrentCaseVersion, Me.UpgradedCaseVersion, Me.BatchLoadInfo, Me.DataExport, Me.CaseLocation, Me.BackUpLocation, Me.CaseDescription, Me.CaseSizeOnDisk, Me.CaseFileSize, Me.CaseAuditSize, Me.OldestTopLevel, Me.NewestTopLevel, Me.IsCompound, Me.CasesContained, Me.ContainedInCase, Me.Investigator, Me.InvestigatorSessions, Me.InvalidSessions, Me.InvestigatorTimeSummary, Me.BrokerMemory, Me.WorkerCount, Me.WorkerMemory, Me.EvidenceName, Me.EvidenceLocation, Me.EvidenceDescription, Me.EvidenceCustomMetadata, Me.LanguagesContained, Me.MimeTypes, Me.ItemTypes, Me.IrregularItems, Me.CreationDate, Me.ModifiedDate, Me.LoadStartDate, Me.LoadEndDate, Me.LoadTime, Me.LoadEvents, Me.TotalLoadTime, Me.ProcessingSpeed, Me.TotalCaseItemCount, Me.ItemCounts, Me.OriginalItems, Me.DuplicateItems, Me.Custodians, Me.CustodianCount, Me.SearchTerm, Me.SearchSize, Me.SearchHitCount, Me.CustodianSearchHit, Me.HitCountPercent, Me.NuixLogLocation})
        Me.grdCaseInfo.Location = New System.Drawing.Point(295, 8)
        Me.grdCaseInfo.Margin = New System.Windows.Forms.Padding(2)
        Me.grdCaseInfo.Name = "grdCaseInfo"
        Me.grdCaseInfo.RowTemplate.Height = 28
        Me.grdCaseInfo.Size = New System.Drawing.Size(984, 499)
        Me.grdCaseInfo.TabIndex = 2
        '
        'CaseGuid
        '
        Me.CaseGuid.Frozen = True
        Me.CaseGuid.HeaderText = "Case Guid"
        Me.CaseGuid.Name = "CaseGuid"
        Me.CaseGuid.Visible = False
        '
        'CollectionStatus
        '
        Me.CollectionStatus.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.AllCells
        DataGridViewCellStyle1.BackColor = System.Drawing.Color.White
        DataGridViewCellStyle1.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        DataGridViewCellStyle1.ForeColor = System.Drawing.Color.Black
        DataGridViewCellStyle1.SelectionBackColor = System.Drawing.Color.Black
        DataGridViewCellStyle1.SelectionForeColor = System.Drawing.Color.White
        Me.CollectionStatus.DefaultCellStyle = DataGridViewCellStyle1
        Me.CollectionStatus.Frozen = True
        Me.CollectionStatus.HeaderText = "Collection Status"
        Me.CollectionStatus.Items.AddRange(New Object() {"File System and Case Data collected", "File System Info Collected", "File System Info Collected - Case Version Mismatch", "File System Info Collected - Case Migrating", "File System Info Collected - Waiting for Case Data", "Getting Case Data from Case...", "Exporting Case Data...", "Collection Stopped...", "Case Data Exported", "Get New Data", "Case Copied or Moved ", "Case No Longer Exists", "Error Checking Case Data (see logs for details)", "Case Locked"})
        Me.CollectionStatus.Name = "CollectionStatus"
        Me.CollectionStatus.Resizable = System.Windows.Forms.DataGridViewTriState.[True]
        Me.CollectionStatus.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.Automatic
        Me.CollectionStatus.Width = 102
        '
        'PercentComplete
        '
        Me.PercentComplete.Frozen = True
        Me.PercentComplete.HeaderText = "Percent Complete"
        Me.PercentComplete.Name = "PercentComplete"
        '
        'ReportLoadDuration
        '
        Me.ReportLoadDuration.Frozen = True
        Me.ReportLoadDuration.HeaderText = "Report Load Duration"
        Me.ReportLoadDuration.Name = "ReportLoadDuration"
        '
        'CaseName
        '
        Me.CaseName.Frozen = True
        Me.CaseName.HeaderText = "Case Name"
        Me.CaseName.Name = "CaseName"
        Me.CaseName.Resizable = System.Windows.Forms.DataGridViewTriState.[True]
        '
        'CurrentCaseVersion
        '
        Me.CurrentCaseVersion.Frozen = True
        Me.CurrentCaseVersion.HeaderText = "Current Case Version"
        Me.CurrentCaseVersion.Name = "CurrentCaseVersion"
        Me.CurrentCaseVersion.Width = 90
        '
        'UpgradedCaseVersion
        '
        Me.UpgradedCaseVersion.Frozen = True
        Me.UpgradedCaseVersion.HeaderText = "Upgraded Case Version"
        Me.UpgradedCaseVersion.Name = "UpgradedCaseVersion"
        Me.UpgradedCaseVersion.Width = 90
        '
        'BatchLoadInfo
        '
        Me.BatchLoadInfo.HeaderText = "Batch Load Info"
        Me.BatchLoadInfo.Name = "BatchLoadInfo"
        '
        'DataExport
        '
        Me.DataExport.HeaderText = "Data Export"
        Me.DataExport.Name = "DataExport"
        '
        'CaseLocation
        '
        Me.CaseLocation.HeaderText = "Case Location"
        Me.CaseLocation.Name = "CaseLocation"
        '
        'BackUpLocation
        '
        Me.BackUpLocation.HeaderText = "Back-up Location"
        Me.BackUpLocation.Name = "BackUpLocation"
        '
        'CaseDescription
        '
        Me.CaseDescription.HeaderText = "Case Description"
        Me.CaseDescription.Name = "CaseDescription"
        '
        'CaseSizeOnDisk
        '
        Me.CaseSizeOnDisk.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.AllCells
        Me.CaseSizeOnDisk.HeaderText = "Case Size On Disk"
        Me.CaseSizeOnDisk.Name = "CaseSizeOnDisk"
        Me.CaseSizeOnDisk.Width = 91
        '
        'CaseFileSize
        '
        Me.CaseFileSize.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.AllCells
        Me.CaseFileSize.HeaderText = "Case File Size"
        Me.CaseFileSize.Name = "CaseFileSize"
        Me.CaseFileSize.Width = 72
        '
        'CaseAuditSize
        '
        Me.CaseAuditSize.HeaderText = "Case Audit Size"
        Me.CaseAuditSize.Name = "CaseAuditSize"
        '
        'OldestTopLevel
        '
        Me.OldestTopLevel.HeaderText = "Oldest Top Level Item"
        Me.OldestTopLevel.Name = "OldestTopLevel"
        '
        'NewestTopLevel
        '
        Me.NewestTopLevel.HeaderText = "Newest Top Level Item"
        Me.NewestTopLevel.Name = "NewestTopLevel"
        '
        'IsCompound
        '
        Me.IsCompound.HeaderText = "Is Compound Case"
        Me.IsCompound.Name = "IsCompound"
        '
        'CasesContained
        '
        Me.CasesContained.HeaderText = "Cases Contained"
        Me.CasesContained.Name = "CasesContained"
        '
        'ContainedInCase
        '
        Me.ContainedInCase.HeaderText = "Contained In Case"
        Me.ContainedInCase.Name = "ContainedInCase"
        '
        'Investigator
        '
        Me.Investigator.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.AllCells
        Me.Investigator.HeaderText = "Investigator"
        Me.Investigator.Name = "Investigator"
        Me.Investigator.Width = 87
        '
        'InvestigatorSessions
        '
        Me.InvestigatorSessions.HeaderText = "Investigator Sessions"
        Me.InvestigatorSessions.Name = "InvestigatorSessions"
        '
        'InvalidSessions
        '
        Me.InvalidSessions.HeaderText = "Invalid Sessions"
        Me.InvalidSessions.Name = "InvalidSessions"
        '
        'InvestigatorTimeSummary
        '
        Me.InvestigatorTimeSummary.HeaderText = "Investigator Time Summary"
        Me.InvestigatorTimeSummary.Name = "InvestigatorTimeSummary"
        '
        'BrokerMemory
        '
        Me.BrokerMemory.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.None
        Me.BrokerMemory.HeaderText = "Broker Memory"
        Me.BrokerMemory.Name = "BrokerMemory"
        Me.BrokerMemory.Width = 75
        '
        'WorkerCount
        '
        Me.WorkerCount.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.None
        Me.WorkerCount.HeaderText = "Worker Count"
        Me.WorkerCount.Name = "WorkerCount"
        Me.WorkerCount.Width = 75
        '
        'WorkerMemory
        '
        Me.WorkerMemory.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.None
        Me.WorkerMemory.HeaderText = "Worker Memory"
        Me.WorkerMemory.Name = "WorkerMemory"
        Me.WorkerMemory.Width = 75
        '
        'EvidenceName
        '
        Me.EvidenceName.HeaderText = "Evidence Processed"
        Me.EvidenceName.Name = "EvidenceName"
        '
        'EvidenceLocation
        '
        Me.EvidenceLocation.HeaderText = "Evidence Location"
        Me.EvidenceLocation.Name = "EvidenceLocation"
        '
        'EvidenceDescription
        '
        Me.EvidenceDescription.HeaderText = "Evidence Description"
        Me.EvidenceDescription.Name = "EvidenceDescription"
        '
        'EvidenceCustomMetadata
        '
        Me.EvidenceCustomMetadata.HeaderText = "Evidence Custom Metadata"
        Me.EvidenceCustomMetadata.Name = "EvidenceCustomMetadata"
        '
        'LanguagesContained
        '
        Me.LanguagesContained.HeaderText = "Languages"
        Me.LanguagesContained.Name = "LanguagesContained"
        '
        'MimeTypes
        '
        Me.MimeTypes.HeaderText = "Mime Types"
        Me.MimeTypes.Name = "MimeTypes"
        '
        'ItemTypes
        '
        Me.ItemTypes.HeaderText = "Item Types"
        Me.ItemTypes.Name = "ItemTypes"
        '
        'IrregularItems
        '
        Me.IrregularItems.HeaderText = "Irregular Items"
        Me.IrregularItems.Name = "IrregularItems"
        '
        'CreationDate
        '
        Me.CreationDate.HeaderText = "Creation Date"
        Me.CreationDate.Name = "CreationDate"
        '
        'ModifiedDate
        '
        Me.ModifiedDate.HeaderText = "Modified Date"
        Me.ModifiedDate.Name = "ModifiedDate"
        '
        'LoadStartDate
        '
        Me.LoadStartDate.HeaderText = "Load Start Date"
        Me.LoadStartDate.Name = "LoadStartDate"
        '
        'LoadEndDate
        '
        Me.LoadEndDate.HeaderText = "Load End Date"
        Me.LoadEndDate.Name = "LoadEndDate"
        '
        'LoadTime
        '
        Me.LoadTime.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.AllCells
        Me.LoadTime.HeaderText = "Load Time"
        Me.LoadTime.Name = "LoadTime"
        Me.LoadTime.Width = 76
        '
        'LoadEvents
        '
        Me.LoadEvents.HeaderText = "Data Load Events"
        Me.LoadEvents.Name = "LoadEvents"
        '
        'TotalLoadTime
        '
        Me.TotalLoadTime.HeaderText = "Total Data Load Time"
        Me.TotalLoadTime.Name = "TotalLoadTime"
        '
        'ProcessingSpeed
        '
        Me.ProcessingSpeed.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.AllCells
        Me.ProcessingSpeed.HeaderText = "Processing Speed"
        Me.ProcessingSpeed.Name = "ProcessingSpeed"
        Me.ProcessingSpeed.Width = 108
        '
        'TotalCaseItemCount
        '
        Me.TotalCaseItemCount.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.AllCells
        Me.TotalCaseItemCount.HeaderText = "Total Case Count"
        Me.TotalCaseItemCount.Name = "TotalCaseItemCount"
        Me.TotalCaseItemCount.Width = 105
        '
        'ItemCounts
        '
        Me.ItemCounts.HeaderText = "Item Counts"
        Me.ItemCounts.Name = "ItemCounts"
        '
        'OriginalItems
        '
        Me.OriginalItems.HeaderText = "Original Items"
        Me.OriginalItems.Name = "OriginalItems"
        '
        'DuplicateItems
        '
        Me.DuplicateItems.HeaderText = "Duplicate Items"
        Me.DuplicateItems.Name = "DuplicateItems"
        '
        'Custodians
        '
        Me.Custodians.HeaderText = "Custodians"
        Me.Custodians.Name = "Custodians"
        '
        'CustodianCount
        '
        Me.CustodianCount.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.AllCells
        Me.CustodianCount.HeaderText = "Custodian Count"
        Me.CustodianCount.Name = "CustodianCount"
        Me.CustodianCount.Width = 101
        '
        'SearchTerm
        '
        Me.SearchTerm.HeaderText = "Search Term"
        Me.SearchTerm.Name = "SearchTerm"
        '
        'SearchSize
        '
        Me.SearchSize.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.AllCells
        Me.SearchSize.HeaderText = "Search Size"
        Me.SearchSize.Name = "SearchSize"
        Me.SearchSize.Width = 82
        '
        'SearchHitCount
        '
        Me.SearchHitCount.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.AllCells
        Me.SearchHitCount.HeaderText = "Search Hit Count"
        Me.SearchHitCount.Name = "SearchHitCount"
        Me.SearchHitCount.Width = 78
        '
        'CustodianSearchHit
        '
        Me.CustodianSearchHit.HeaderText = "Custodian Search Hit"
        Me.CustodianSearchHit.Name = "CustodianSearchHit"
        '
        'HitCountPercent
        '
        Me.HitCountPercent.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.AllCells
        Me.HitCountPercent.HeaderText = "Hit Count Percent"
        Me.HitCountPercent.Name = "HitCountPercent"
        Me.HitCountPercent.Width = 106
        '
        'NuixLogLocation
        '
        Me.NuixLogLocation.HeaderText = "Nuix Log"
        Me.NuixLogLocation.Name = "NuixLogLocation"
        '
        'cboReportType
        '
        Me.cboReportType.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.cboReportType.FormattingEnabled = True
        Me.cboReportType.Items.AddRange(New Object() {"All", "App Memory per case", "Case by Investigator", "Case Evidence", "Case Location", "Case Size", "Compound Case Analysis", "Custodians in Case", "Metadata type", "Processing time", "Processing speed (GB per hour)", "Search Term Hit", "Total Number of Items", "Total Number of workers"})
        Me.cboReportType.Location = New System.Drawing.Point(123, 520)
        Me.cboReportType.Margin = New System.Windows.Forms.Padding(2)
        Me.cboReportType.Name = "cboReportType"
        Me.cboReportType.Size = New System.Drawing.Size(169, 21)
        Me.cboReportType.TabIndex = 3
        '
        'lblReportType
        '
        Me.lblReportType.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.lblReportType.AutoSize = True
        Me.lblReportType.Location = New System.Drawing.Point(10, 525)
        Me.lblReportType.Margin = New System.Windows.Forms.Padding(2, 0, 2, 0)
        Me.lblReportType.Name = "lblReportType"
        Me.lblReportType.Size = New System.Drawing.Size(69, 13)
        Me.lblReportType.TabIndex = 48
        Me.lblReportType.Text = "Report Type:"
        '
        'ImgIcons
        '
        Me.ImgIcons.ImageStream = CType(resources.GetObject("ImgIcons.ImageStream"), System.Windows.Forms.ImageListStreamer)
        Me.ImgIcons.TransparentColor = System.Drawing.Color.Transparent
        Me.ImgIcons.Images.SetKeyName(0, "folder_Closed_32xMD.png")
        Me.ImgIcons.Images.SetKeyName(1, "folder_Open_32xMD.png")
        '
        'lblSearchTerm
        '
        Me.lblSearchTerm.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.lblSearchTerm.AutoSize = True
        Me.lblSearchTerm.Location = New System.Drawing.Point(10, 585)
        Me.lblSearchTerm.Margin = New System.Windows.Forms.Padding(2, 0, 2, 0)
        Me.lblSearchTerm.Name = "lblSearchTerm"
        Me.lblSearchTerm.Size = New System.Drawing.Size(71, 13)
        Me.lblSearchTerm.TabIndex = 49
        Me.lblSearchTerm.Text = "Search Term:"
        '
        'txtSearchTerm
        '
        Me.txtSearchTerm.Anchor = CType(((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.txtSearchTerm.Location = New System.Drawing.Point(298, 580)
        Me.txtSearchTerm.Margin = New System.Windows.Forms.Padding(2)
        Me.txtSearchTerm.Name = "txtSearchTerm"
        Me.txtSearchTerm.Size = New System.Drawing.Size(709, 20)
        Me.txtSearchTerm.TabIndex = 11
        '
        'btnExportCaseStatistics
        '
        Me.btnExportCaseStatistics.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.btnExportCaseStatistics.Location = New System.Drawing.Point(1192, 788)
        Me.btnExportCaseStatistics.Margin = New System.Windows.Forms.Padding(2)
        Me.btnExportCaseStatistics.Name = "btnExportCaseStatistics"
        Me.btnExportCaseStatistics.Size = New System.Drawing.Size(81, 36)
        Me.btnExportCaseStatistics.TabIndex = 25
        Me.btnExportCaseStatistics.Text = "Export Data..."
        Me.btnExportCaseStatistics.UseVisualStyleBackColor = True
        '
        'ExportContextMenuStrip
        '
        Me.ExportContextMenuStrip.ImageScalingSize = New System.Drawing.Size(24, 24)
        Me.ExportContextMenuStrip.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.ExcelToolStripMenuItem, Me.CSVToolStripMenuItem, Me.JSonToolStripMenuItem, Me.XMLToolStripMenuItem})
        Me.ExportContextMenuStrip.Name = "ExportContextMenuStrip"
        Me.ExportContextMenuStrip.Size = New System.Drawing.Size(101, 92)
        '
        'ExcelToolStripMenuItem
        '
        Me.ExcelToolStripMenuItem.Name = "ExcelToolStripMenuItem"
        Me.ExcelToolStripMenuItem.Size = New System.Drawing.Size(100, 22)
        Me.ExcelToolStripMenuItem.Text = "Excel"
        '
        'CSVToolStripMenuItem
        '
        Me.CSVToolStripMenuItem.Name = "CSVToolStripMenuItem"
        Me.CSVToolStripMenuItem.Size = New System.Drawing.Size(100, 22)
        Me.CSVToolStripMenuItem.Text = "CSV"
        '
        'JSonToolStripMenuItem
        '
        Me.JSonToolStripMenuItem.Name = "JSonToolStripMenuItem"
        Me.JSonToolStripMenuItem.Size = New System.Drawing.Size(100, 22)
        Me.JSonToolStripMenuItem.Text = "JSon"
        '
        'XMLToolStripMenuItem
        '
        Me.XMLToolStripMenuItem.Name = "XMLToolStripMenuItem"
        Me.XMLToolStripMenuItem.Size = New System.Drawing.Size(100, 22)
        Me.XMLToolStripMenuItem.Text = "XML"
        '
        'lblNuixVersionLocation
        '
        Me.lblNuixVersionLocation.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.lblNuixVersionLocation.AutoSize = True
        Me.lblNuixVersionLocation.Location = New System.Drawing.Point(10, 645)
        Me.lblNuixVersionLocation.Margin = New System.Windows.Forms.Padding(2, 0, 2, 0)
        Me.lblNuixVersionLocation.Name = "lblNuixVersionLocation"
        Me.lblNuixVersionLocation.Size = New System.Drawing.Size(116, 13)
        Me.lblNuixVersionLocation.TabIndex = 52
        Me.lblNuixVersionLocation.Text = "Nuix Console Location:"
        '
        'txtNuixConsoleLocation
        '
        Me.txtNuixConsoleLocation.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.txtNuixConsoleLocation.Location = New System.Drawing.Point(123, 640)
        Me.txtNuixConsoleLocation.Margin = New System.Windows.Forms.Padding(2)
        Me.txtNuixConsoleLocation.Name = "txtNuixConsoleLocation"
        Me.txtNuixConsoleLocation.Size = New System.Drawing.Size(168, 20)
        Me.txtNuixConsoleLocation.TabIndex = 14
        '
        'btnConsoleLocation
        '
        Me.btnConsoleLocation.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.btnConsoleLocation.Location = New System.Drawing.Point(296, 640)
        Me.btnConsoleLocation.Margin = New System.Windows.Forms.Padding(2)
        Me.btnConsoleLocation.Name = "btnConsoleLocation"
        Me.btnConsoleLocation.Size = New System.Drawing.Size(24, 19)
        Me.btnConsoleLocation.TabIndex = 15
        Me.btnConsoleLocation.Text = "..."
        Me.btnConsoleLocation.UseVisualStyleBackColor = True
        '
        'lblNuixConsoleVersion
        '
        Me.lblNuixConsoleVersion.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.lblNuixConsoleVersion.AutoSize = True
        Me.lblNuixConsoleVersion.Location = New System.Drawing.Point(321, 645)
        Me.lblNuixConsoleVersion.Margin = New System.Windows.Forms.Padding(2, 0, 2, 0)
        Me.lblNuixConsoleVersion.Name = "lblNuixConsoleVersion"
        Me.lblNuixConsoleVersion.Size = New System.Drawing.Size(110, 13)
        Me.lblNuixConsoleVersion.TabIndex = 55
        Me.lblNuixConsoleVersion.Text = "Nuix Console Version:"
        '
        'cboNuixLicenseType
        '
        Me.cboNuixLicenseType.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.cboNuixLicenseType.FormattingEnabled = True
        Me.cboNuixLicenseType.Items.AddRange(New Object() {"Corporate eDiscovery", "eDiscovery Workstation", "eDiscovery Reviewer", "Email Archive Examiner", "Investigation and Response", "Ultimate Workstation"})
        Me.cboNuixLicenseType.Location = New System.Drawing.Point(123, 670)
        Me.cboNuixLicenseType.Margin = New System.Windows.Forms.Padding(2)
        Me.cboNuixLicenseType.Name = "cboNuixLicenseType"
        Me.cboNuixLicenseType.Size = New System.Drawing.Size(168, 21)
        Me.cboNuixLicenseType.TabIndex = 16
        '
        'lblNuixVersion
        '
        Me.lblNuixVersion.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.lblNuixVersion.AutoSize = True
        Me.lblNuixVersion.Location = New System.Drawing.Point(10, 675)
        Me.lblNuixVersion.Margin = New System.Windows.Forms.Padding(2, 0, 2, 0)
        Me.lblNuixVersion.Name = "lblNuixVersion"
        Me.lblNuixVersion.Size = New System.Drawing.Size(71, 13)
        Me.lblNuixVersion.TabIndex = 62
        Me.lblNuixVersion.Text = "Nuix License:"
        '
        'btnLoadPreviousReportingRun
        '
        Me.btnLoadPreviousReportingRun.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.btnLoadPreviousReportingRun.Location = New System.Drawing.Point(1105, 788)
        Me.btnLoadPreviousReportingRun.Margin = New System.Windows.Forms.Padding(2)
        Me.btnLoadPreviousReportingRun.Name = "btnLoadPreviousReportingRun"
        Me.btnLoadPreviousReportingRun.Size = New System.Drawing.Size(84, 36)
        Me.btnLoadPreviousReportingRun.TabIndex = 24
        Me.btnLoadPreviousReportingRun.Text = "Load Previous Reports"
        Me.btnLoadPreviousReportingRun.UseVisualStyleBackColor = True
        '
        'lblNuixLogDir
        '
        Me.lblNuixLogDir.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.lblNuixLogDir.AutoSize = True
        Me.lblNuixLogDir.Location = New System.Drawing.Point(10, 615)
        Me.lblNuixLogDir.Margin = New System.Windows.Forms.Padding(2, 0, 2, 0)
        Me.lblNuixLogDir.Name = "lblNuixLogDir"
        Me.lblNuixLogDir.Size = New System.Drawing.Size(87, 13)
        Me.lblNuixLogDir.TabIndex = 64
        Me.lblNuixLogDir.Text = "Nuix Log File Dir:"
        '
        'txtNuixLogDir
        '
        Me.txtNuixLogDir.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.txtNuixLogDir.Location = New System.Drawing.Point(123, 610)
        Me.txtNuixLogDir.Margin = New System.Windows.Forms.Padding(2)
        Me.txtNuixLogDir.Name = "txtNuixLogDir"
        Me.txtNuixLogDir.Size = New System.Drawing.Size(168, 20)
        Me.txtNuixLogDir.TabIndex = 12
        '
        'btnNuixLogSelector
        '
        Me.btnNuixLogSelector.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.btnNuixLogSelector.Location = New System.Drawing.Point(296, 610)
        Me.btnNuixLogSelector.Margin = New System.Windows.Forms.Padding(2)
        Me.btnNuixLogSelector.Name = "btnNuixLogSelector"
        Me.btnNuixLogSelector.Size = New System.Drawing.Size(24, 19)
        Me.btnNuixLogSelector.TabIndex = 13
        Me.btnNuixLogSelector.Text = "..."
        Me.btnNuixLogSelector.UseVisualStyleBackColor = True
        '
        'chkBackUpCase
        '
        Me.chkBackUpCase.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.chkBackUpCase.AutoSize = True
        Me.chkBackUpCase.Location = New System.Drawing.Point(571, 553)
        Me.chkBackUpCase.Name = "chkBackUpCase"
        Me.chkBackUpCase.Size = New System.Drawing.Size(90, 17)
        Me.chkBackUpCase.TabIndex = 6
        Me.chkBackUpCase.Text = "Backup Case"
        Me.chkBackUpCase.UseVisualStyleBackColor = True
        '
        'txtBackupLocation
        '
        Me.txtBackupLocation.Anchor = CType(((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.txtBackupLocation.Location = New System.Drawing.Point(827, 550)
        Me.txtBackupLocation.Name = "txtBackupLocation"
        Me.txtBackupLocation.Size = New System.Drawing.Size(423, 20)
        Me.txtBackupLocation.TabIndex = 7
        '
        'lblBackUpLocation
        '
        Me.lblBackUpLocation.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.lblBackUpLocation.AutoSize = True
        Me.lblBackUpLocation.Location = New System.Drawing.Point(730, 555)
        Me.lblBackUpLocation.Name = "lblBackUpLocation"
        Me.lblBackUpLocation.Size = New System.Drawing.Size(91, 13)
        Me.lblBackUpLocation.TabIndex = 71
        Me.lblBackUpLocation.Text = "Backup Location:"
        '
        'btnBackupLocationChooser
        '
        Me.btnBackupLocationChooser.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.btnBackupLocationChooser.Location = New System.Drawing.Point(1255, 550)
        Me.btnBackupLocationChooser.Margin = New System.Windows.Forms.Padding(2)
        Me.btnBackupLocationChooser.Name = "btnBackupLocationChooser"
        Me.btnBackupLocationChooser.Size = New System.Drawing.Size(24, 19)
        Me.btnBackupLocationChooser.TabIndex = 8
        Me.btnBackupLocationChooser.Text = "..."
        Me.btnBackupLocationChooser.UseVisualStyleBackColor = True
        '
        'grpLicenseType
        '
        Me.grpLicenseType.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.grpLicenseType.Controls.Add(Me.numExportWorkers)
        Me.grpLicenseType.Controls.Add(Me.lblExportWorkers)
        Me.grpLicenseType.Controls.Add(Me.lblExportMemory)
        Me.grpLicenseType.Controls.Add(Me.numExportWorkerMemory)
        Me.grpLicenseType.Controls.Add(Me.lblNuixAppMemory)
        Me.grpLicenseType.Controls.Add(Me.numNuixAppMemory)
        Me.grpLicenseType.Controls.Add(Me.lblServerType)
        Me.grpLicenseType.Controls.Add(Me.cboLicenseType)
        Me.grpLicenseType.Controls.Add(Me.txtRegistryServer)
        Me.grpLicenseType.Controls.Add(Me.lblRegistryServer)
        Me.grpLicenseType.Controls.Add(Me.txtNMSInfo)
        Me.grpLicenseType.Controls.Add(Me.lblNMSInfo)
        Me.grpLicenseType.Controls.Add(Me.txtNMSUserName)
        Me.grpLicenseType.Controls.Add(Me.txtNMSAddress)
        Me.grpLicenseType.Controls.Add(Me.lblNMSAddress)
        Me.grpLicenseType.Location = New System.Drawing.Point(10, 696)
        Me.grpLicenseType.Name = "grpLicenseType"
        Me.grpLicenseType.Size = New System.Drawing.Size(483, 128)
        Me.grpLicenseType.TabIndex = 73
        Me.grpLicenseType.TabStop = False
        Me.grpLicenseType.Text = "License Info"
        '
        'numExportWorkers
        '
        Me.numExportWorkers.Increment = New Decimal(New Integer() {2, 0, 0, 0})
        Me.numExportWorkers.Location = New System.Drawing.Point(287, 105)
        Me.numExportWorkers.Maximum = New Decimal(New Integer() {16, 0, 0, 0})
        Me.numExportWorkers.Minimum = New Decimal(New Integer() {2, 0, 0, 0})
        Me.numExportWorkers.Name = "numExportWorkers"
        Me.numExportWorkers.Size = New System.Drawing.Size(38, 20)
        Me.numExportWorkers.TabIndex = 81
        Me.numExportWorkers.Value = New Decimal(New Integer() {2, 0, 0, 0})
        '
        'lblExportWorkers
        '
        Me.lblExportWorkers.AutoSize = True
        Me.lblExportWorkers.Location = New System.Drawing.Point(200, 110)
        Me.lblExportWorkers.Name = "lblExportWorkers"
        Me.lblExportWorkers.Size = New System.Drawing.Size(83, 13)
        Me.lblExportWorkers.TabIndex = 80
        Me.lblExportWorkers.Text = "Export Workers:"
        '
        'lblExportMemory
        '
        Me.lblExportMemory.AutoSize = True
        Me.lblExportMemory.Location = New System.Drawing.Point(340, 110)
        Me.lblExportMemory.Name = "lblExportMemory"
        Me.lblExportMemory.Size = New System.Drawing.Size(80, 13)
        Me.lblExportMemory.TabIndex = 79
        Me.lblExportMemory.Text = "Export Memory:"
        '
        'numExportWorkerMemory
        '
        Me.numExportWorkerMemory.Increment = New Decimal(New Integer() {2, 0, 0, 0})
        Me.numExportWorkerMemory.Location = New System.Drawing.Point(429, 105)
        Me.numExportWorkerMemory.Maximum = New Decimal(New Integer() {64, 0, 0, 0})
        Me.numExportWorkerMemory.Minimum = New Decimal(New Integer() {2, 0, 0, 0})
        Me.numExportWorkerMemory.Name = "numExportWorkerMemory"
        Me.numExportWorkerMemory.Size = New System.Drawing.Size(38, 20)
        Me.numExportWorkerMemory.TabIndex = 78
        Me.numExportWorkerMemory.Value = New Decimal(New Integer() {2, 0, 0, 0})
        '
        'lblNuixAppMemory
        '
        Me.lblNuixAppMemory.AutoSize = True
        Me.lblNuixAppMemory.Location = New System.Drawing.Point(10, 110)
        Me.lblNuixAppMemory.Name = "lblNuixAppMemory"
        Me.lblNuixAppMemory.Size = New System.Drawing.Size(93, 13)
        Me.lblNuixAppMemory.TabIndex = 77
        Me.lblNuixAppMemory.Text = "Nuix App Memory:"
        '
        'numNuixAppMemory
        '
        Me.numNuixAppMemory.Increment = New Decimal(New Integer() {4, 0, 0, 0})
        Me.numNuixAppMemory.Location = New System.Drawing.Point(115, 105)
        Me.numNuixAppMemory.Minimum = New Decimal(New Integer() {4, 0, 0, 0})
        Me.numNuixAppMemory.Name = "numNuixAppMemory"
        Me.numNuixAppMemory.Size = New System.Drawing.Size(38, 20)
        Me.numNuixAppMemory.TabIndex = 22
        Me.numNuixAppMemory.Value = New Decimal(New Integer() {4, 0, 0, 0})
        '
        'lblServerType
        '
        Me.lblServerType.Anchor = System.Windows.Forms.AnchorStyles.None
        Me.lblServerType.AutoSize = True
        Me.lblServerType.Location = New System.Drawing.Point(10, 20)
        Me.lblServerType.Name = "lblServerType"
        Me.lblServerType.Size = New System.Drawing.Size(68, 13)
        Me.lblServerType.TabIndex = 76
        Me.lblServerType.Text = "Server Type:"
        '
        'cboLicenseType
        '
        Me.cboLicenseType.Anchor = System.Windows.Forms.AnchorStyles.None
        Me.cboLicenseType.FormattingEnabled = True
        Me.cboLicenseType.Items.AddRange(New Object() {"Desktop", "Desktop (dongleless)", "Server"})
        Me.cboLicenseType.Location = New System.Drawing.Point(115, 15)
        Me.cboLicenseType.Name = "cboLicenseType"
        Me.cboLicenseType.Size = New System.Drawing.Size(168, 21)
        Me.cboLicenseType.TabIndex = 17
        '
        'txtRegistryServer
        '
        Me.txtRegistryServer.Anchor = System.Windows.Forms.AnchorStyles.None
        Me.txtRegistryServer.Location = New System.Drawing.Point(380, 50)
        Me.txtRegistryServer.Margin = New System.Windows.Forms.Padding(2)
        Me.txtRegistryServer.Name = "txtRegistryServer"
        Me.txtRegistryServer.Size = New System.Drawing.Size(87, 20)
        Me.txtRegistryServer.TabIndex = 19
        '
        'lblRegistryServer
        '
        Me.lblRegistryServer.Anchor = System.Windows.Forms.AnchorStyles.None
        Me.lblRegistryServer.AutoSize = True
        Me.lblRegistryServer.Location = New System.Drawing.Point(300, 50)
        Me.lblRegistryServer.Margin = New System.Windows.Forms.Padding(2, 0, 2, 0)
        Me.lblRegistryServer.Name = "lblRegistryServer"
        Me.lblRegistryServer.Size = New System.Drawing.Size(82, 13)
        Me.lblRegistryServer.TabIndex = 74
        Me.lblRegistryServer.Text = "Registry Server:"
        '
        'txtNMSInfo
        '
        Me.txtNMSInfo.Anchor = System.Windows.Forms.AnchorStyles.None
        Me.txtNMSInfo.Location = New System.Drawing.Point(203, 75)
        Me.txtNMSInfo.Margin = New System.Windows.Forms.Padding(2)
        Me.txtNMSInfo.Name = "txtNMSInfo"
        Me.txtNMSInfo.PasswordChar = Global.Microsoft.VisualBasic.ChrW(42)
        Me.txtNMSInfo.Size = New System.Drawing.Size(79, 20)
        Me.txtNMSInfo.TabIndex = 21
        '
        'lblNMSInfo
        '
        Me.lblNMSInfo.Anchor = System.Windows.Forms.AnchorStyles.None
        Me.lblNMSInfo.AutoSize = True
        Me.lblNMSInfo.Location = New System.Drawing.Point(10, 80)
        Me.lblNMSInfo.Margin = New System.Windows.Forms.Padding(2, 0, 2, 0)
        Me.lblNMSInfo.Name = "lblNMSInfo"
        Me.lblNMSInfo.Size = New System.Drawing.Size(89, 13)
        Me.lblNMSInfo.TabIndex = 73
        Me.lblNMSInfo.Text = "NMS Information:"
        '
        'txtNMSUserName
        '
        Me.txtNMSUserName.Anchor = System.Windows.Forms.AnchorStyles.None
        Me.txtNMSUserName.Location = New System.Drawing.Point(115, 75)
        Me.txtNMSUserName.Margin = New System.Windows.Forms.Padding(2)
        Me.txtNMSUserName.Name = "txtNMSUserName"
        Me.txtNMSUserName.Size = New System.Drawing.Size(82, 20)
        Me.txtNMSUserName.TabIndex = 20
        '
        'txtNMSAddress
        '
        Me.txtNMSAddress.Anchor = System.Windows.Forms.AnchorStyles.None
        Me.txtNMSAddress.Location = New System.Drawing.Point(115, 45)
        Me.txtNMSAddress.Margin = New System.Windows.Forms.Padding(2)
        Me.txtNMSAddress.Name = "txtNMSAddress"
        Me.txtNMSAddress.Size = New System.Drawing.Size(168, 20)
        Me.txtNMSAddress.TabIndex = 18
        '
        'lblNMSAddress
        '
        Me.lblNMSAddress.Anchor = System.Windows.Forms.AnchorStyles.None
        Me.lblNMSAddress.AutoSize = True
        Me.lblNMSAddress.Location = New System.Drawing.Point(10, 50)
        Me.lblNMSAddress.Margin = New System.Windows.Forms.Padding(2, 0, 2, 0)
        Me.lblNMSAddress.Name = "lblNMSAddress"
        Me.lblNMSAddress.Size = New System.Drawing.Size(75, 13)
        Me.lblNMSAddress.TabIndex = 72
        Me.lblNMSAddress.Text = "NMS Address:"
        '
        'lblCalculateProcessingSpeeds
        '
        Me.lblCalculateProcessingSpeeds.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.lblCalculateProcessingSpeeds.AutoSize = True
        Me.lblCalculateProcessingSpeeds.Location = New System.Drawing.Point(583, 523)
        Me.lblCalculateProcessingSpeeds.Name = "lblCalculateProcessingSpeeds"
        Me.lblCalculateProcessingSpeeds.Size = New System.Drawing.Size(178, 13)
        Me.lblCalculateProcessingSpeeds.TabIndex = 74
        Me.lblCalculateProcessingSpeeds.Text = "Calculate Processing Speeds Using:"
        '
        'cboCalculateProcessingSpeeds
        '
        Me.cboCalculateProcessingSpeeds.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.cboCalculateProcessingSpeeds.FormattingEnabled = True
        Me.cboCalculateProcessingSpeeds.Items.AddRange(New Object() {"Audit Size", "File Size"})
        Me.cboCalculateProcessingSpeeds.Location = New System.Drawing.Point(765, 518)
        Me.cboCalculateProcessingSpeeds.Name = "cboCalculateProcessingSpeeds"
        Me.cboCalculateProcessingSpeeds.Size = New System.Drawing.Size(76, 21)
        Me.cboCalculateProcessingSpeeds.TabIndex = 4
        '
        'chkExportSearchResults
        '
        Me.chkExportSearchResults.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.chkExportSearchResults.AutoSize = True
        Me.chkExportSearchResults.Location = New System.Drawing.Point(1057, 580)
        Me.chkExportSearchResults.Name = "chkExportSearchResults"
        Me.chkExportSearchResults.Size = New System.Drawing.Size(131, 17)
        Me.chkExportSearchResults.TabIndex = 75
        Me.chkExportSearchResults.Text = "Export Search Results"
        Me.chkExportSearchResults.UseVisualStyleBackColor = True
        '
        'btnGetFileSystemData
        '
        Me.btnGetFileSystemData.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.btnGetFileSystemData.Location = New System.Drawing.Point(943, 788)
        Me.btnGetFileSystemData.Name = "btnGetFileSystemData"
        Me.btnGetFileSystemData.Size = New System.Drawing.Size(78, 36)
        Me.btnGetFileSystemData.TabIndex = 77
        Me.btnGetFileSystemData.Text = "Get File System Data"
        Me.btnGetFileSystemData.UseVisualStyleBackColor = True
        '
        'btnSaveConfig
        '
        Me.btnSaveConfig.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.btnSaveConfig.Location = New System.Drawing.Point(501, 788)
        Me.btnSaveConfig.Name = "btnSaveConfig"
        Me.btnSaveConfig.Size = New System.Drawing.Size(78, 36)
        Me.btnSaveConfig.TabIndex = 78
        Me.btnSaveConfig.Text = "Save Config"
        Me.btnSaveConfig.UseVisualStyleBackColor = True
        '
        'btnLoadConfig
        '
        Me.btnLoadConfig.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.btnLoadConfig.Location = New System.Drawing.Point(583, 788)
        Me.btnLoadConfig.Name = "btnLoadConfig"
        Me.btnLoadConfig.Size = New System.Drawing.Size(78, 36)
        Me.btnLoadConfig.TabIndex = 79
        Me.btnLoadConfig.Text = "Load Config"
        Me.btnLoadConfig.UseVisualStyleBackColor = True
        '
        'cboSizeReporting
        '
        Me.cboSizeReporting.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.cboSizeReporting.FormattingEnabled = True
        Me.cboSizeReporting.Items.AddRange(New Object() {"Bytes", "Megabytes", "Gigabytes"})
        Me.cboSizeReporting.Location = New System.Drawing.Point(483, 518)
        Me.cboSizeReporting.Name = "cboSizeReporting"
        Me.cboSizeReporting.Size = New System.Drawing.Size(79, 21)
        Me.cboSizeReporting.TabIndex = 80
        '
        'lblShowSizeIn
        '
        Me.lblShowSizeIn.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.lblShowSizeIn.AutoSize = True
        Me.lblShowSizeIn.Location = New System.Drawing.Point(409, 523)
        Me.lblShowSizeIn.Name = "lblShowSizeIn"
        Me.lblShowSizeIn.Size = New System.Drawing.Size(71, 13)
        Me.lblShowSizeIn.TabIndex = 81
        Me.lblShowSizeIn.Text = "Show Size in:"
        '
        'grpCaseSelector
        '
        Me.grpCaseSelector.Controls.Add(Me.radFile)
        Me.grpCaseSelector.Controls.Add(Me.radFileSystem)
        Me.grpCaseSelector.Location = New System.Drawing.Point(8, 8)
        Me.grpCaseSelector.Name = "grpCaseSelector"
        Me.grpCaseSelector.Size = New System.Drawing.Size(281, 42)
        Me.grpCaseSelector.TabIndex = 82
        Me.grpCaseSelector.TabStop = False
        Me.grpCaseSelector.Text = "Case Location"
        '
        'radFile
        '
        Me.radFile.AutoSize = True
        Me.radFile.Location = New System.Drawing.Point(100, 19)
        Me.radFile.Name = "radFile"
        Me.radFile.Size = New System.Drawing.Size(41, 17)
        Me.radFile.TabIndex = 1
        Me.radFile.TabStop = True
        Me.radFile.Text = "File"
        Me.radFile.UseVisualStyleBackColor = True
        '
        'radFileSystem
        '
        Me.radFileSystem.AutoSize = True
        Me.radFileSystem.Location = New System.Drawing.Point(16, 19)
        Me.radFileSystem.Name = "radFileSystem"
        Me.radFileSystem.Size = New System.Drawing.Size(78, 17)
        Me.radFileSystem.TabIndex = 0
        Me.radFileSystem.TabStop = True
        Me.radFileSystem.Text = "File System"
        Me.radFileSystem.UseVisualStyleBackColor = True
        '
        'txtCaseFileLocations
        '
        Me.txtCaseFileLocations.Location = New System.Drawing.Point(37, 56)
        Me.txtCaseFileLocations.Name = "txtCaseFileLocations"
        Me.txtCaseFileLocations.Size = New System.Drawing.Size(225, 20)
        Me.txtCaseFileLocations.TabIndex = 83
        '
        'btnFileLocation
        '
        Me.btnFileLocation.Location = New System.Drawing.Point(265, 56)
        Me.btnFileLocation.Name = "btnFileLocation"
        Me.btnFileLocation.Size = New System.Drawing.Size(24, 19)
        Me.btnFileLocation.TabIndex = 84
        Me.btnFileLocation.Text = "..."
        Me.btnFileLocation.UseVisualStyleBackColor = True
        '
        'lblFileLocation
        '
        Me.lblFileLocation.AutoSize = True
        Me.lblFileLocation.Location = New System.Drawing.Point(5, 59)
        Me.lblFileLocation.Name = "lblFileLocation"
        Me.lblFileLocation.Size = New System.Drawing.Size(26, 13)
        Me.lblFileLocation.TabIndex = 85
        Me.lblFileLocation.Text = "File:"
        '
        'btnStopProcessing
        '
        Me.btnStopProcessing.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.btnStopProcessing.Location = New System.Drawing.Point(862, 788)
        Me.btnStopProcessing.Name = "btnStopProcessing"
        Me.btnStopProcessing.Size = New System.Drawing.Size(78, 36)
        Me.btnStopProcessing.TabIndex = 86
        Me.btnStopProcessing.Text = "Stop Processing"
        Me.btnStopProcessing.UseVisualStyleBackColor = True
        '
        'lblExportLocation
        '
        Me.lblExportLocation.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.lblExportLocation.AutoSize = True
        Me.lblExportLocation.Location = New System.Drawing.Point(923, 615)
        Me.lblExportLocation.Name = "lblExportLocation"
        Me.lblExportLocation.Size = New System.Drawing.Size(84, 13)
        Me.lblExportLocation.TabIndex = 87
        Me.lblExportLocation.Text = "Export Location:"
        '
        'txtExportLocation
        '
        Me.txtExportLocation.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.txtExportLocation.Location = New System.Drawing.Point(1013, 610)
        Me.txtExportLocation.Name = "txtExportLocation"
        Me.txtExportLocation.Size = New System.Drawing.Size(237, 20)
        Me.txtExportLocation.TabIndex = 88
        '
        'btnExportLocation
        '
        Me.btnExportLocation.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.btnExportLocation.Location = New System.Drawing.Point(1255, 610)
        Me.btnExportLocation.Margin = New System.Windows.Forms.Padding(2)
        Me.btnExportLocation.Name = "btnExportLocation"
        Me.btnExportLocation.Size = New System.Drawing.Size(24, 19)
        Me.btnExportLocation.TabIndex = 89
        Me.btnExportLocation.Text = "..."
        Me.btnExportLocation.UseVisualStyleBackColor = True
        '
        'cboExportType
        '
        Me.cboExportType.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.cboExportType.FormattingEnabled = True
        Me.cboExportType.Items.AddRange(New Object() {"", "Case Subset", "Native", "NLI", "Mailbox", "PDF"})
        Me.cboExportType.Location = New System.Drawing.Point(1183, 580)
        Me.cboExportType.Name = "cboExportType"
        Me.cboExportType.Size = New System.Drawing.Size(96, 21)
        Me.cboExportType.TabIndex = 76
        '
        'grpSearchTerm
        '
        Me.grpSearchTerm.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.grpSearchTerm.Controls.Add(Me.radSearchFile)
        Me.grpSearchTerm.Controls.Add(Me.radSearchTerm)
        Me.grpSearchTerm.Location = New System.Drawing.Point(123, 575)
        Me.grpSearchTerm.Name = "grpSearchTerm"
        Me.grpSearchTerm.Size = New System.Drawing.Size(170, 30)
        Me.grpSearchTerm.TabIndex = 90
        Me.grpSearchTerm.TabStop = False
        '
        'radSearchFile
        '
        Me.radSearchFile.AutoSize = True
        Me.radSearchFile.Location = New System.Drawing.Point(98, 10)
        Me.radSearchFile.Name = "radSearchFile"
        Me.radSearchFile.Size = New System.Drawing.Size(65, 17)
        Me.radSearchFile.TabIndex = 1
        Me.radSearchFile.TabStop = True
        Me.radSearchFile.Text = "CSV File"
        Me.radSearchFile.UseVisualStyleBackColor = True
        '
        'radSearchTerm
        '
        Me.radSearchTerm.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.radSearchTerm.AutoSize = True
        Me.radSearchTerm.Location = New System.Drawing.Point(6, 10)
        Me.radSearchTerm.Name = "radSearchTerm"
        Me.radSearchTerm.Size = New System.Drawing.Size(86, 17)
        Me.radSearchTerm.TabIndex = 0
        Me.radSearchTerm.TabStop = True
        Me.radSearchTerm.Text = "Search Term"
        Me.radSearchTerm.UseVisualStyleBackColor = True
        '
        'btnSearchTermFile
        '
        Me.btnSearchTermFile.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.btnSearchTermFile.Location = New System.Drawing.Point(1012, 580)
        Me.btnSearchTermFile.Name = "btnSearchTermFile"
        Me.btnSearchTermFile.Size = New System.Drawing.Size(24, 19)
        Me.btnSearchTermFile.TabIndex = 91
        Me.btnSearchTermFile.Text = "..."
        Me.btnSearchTermFile.UseVisualStyleBackColor = True
        '
        'chkExportOnly
        '
        Me.chkExportOnly.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.chkExportOnly.AutoSize = True
        Me.chkExportOnly.Location = New System.Drawing.Point(769, 615)
        Me.chkExportOnly.Name = "chkExportOnly"
        Me.chkExportOnly.Size = New System.Drawing.Size(145, 17)
        Me.chkExportOnly.TabIndex = 92
        Me.chkExportOnly.Text = "Export Only (no reporting)"
        Me.chkExportOnly.UseVisualStyleBackColor = True
        '
        'chkRollUpReporting
        '
        Me.chkRollUpReporting.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.chkRollUpReporting.AutoSize = True
        Me.chkRollUpReporting.Location = New System.Drawing.Point(1192, 766)
        Me.chkRollUpReporting.Name = "chkRollUpReporting"
        Me.chkRollUpReporting.Size = New System.Drawing.Size(94, 17)
        Me.chkRollUpReporting.TabIndex = 93
        Me.chkRollUpReporting.Text = "Roll-up Report"
        Me.chkRollUpReporting.UseVisualStyleBackColor = True
        '
        'cboCopyMoveCases
        '
        Me.cboCopyMoveCases.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.cboCopyMoveCases.FormattingEnabled = True
        Me.cboCopyMoveCases.Items.AddRange(New Object() {"Move", "Copy"})
        Me.cboCopyMoveCases.Location = New System.Drawing.Point(668, 551)
        Me.cboCopyMoveCases.Name = "cboCopyMoveCases"
        Me.cboCopyMoveCases.Size = New System.Drawing.Size(58, 21)
        Me.cboCopyMoveCases.TabIndex = 94
        '
        'cboUpgradeCasees
        '
        Me.cboUpgradeCasees.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.cboUpgradeCasees.FormattingEnabled = True
        Me.cboUpgradeCasees.Items.AddRange(New Object() {"No", "Upgrade Only", "Upgrade and Report"})
        Me.cboUpgradeCasees.Location = New System.Drawing.Point(452, 550)
        Me.cboUpgradeCasees.Name = "cboUpgradeCasees"
        Me.cboUpgradeCasees.Size = New System.Drawing.Size(113, 21)
        Me.cboUpgradeCasees.TabIndex = 96
        '
        'lblUpgradeCases
        '
        Me.lblUpgradeCases.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.lblUpgradeCases.AutoSize = True
        Me.lblUpgradeCases.Location = New System.Drawing.Point(365, 554)
        Me.lblUpgradeCases.Name = "lblUpgradeCases"
        Me.lblUpgradeCases.Size = New System.Drawing.Size(83, 13)
        Me.lblUpgradeCases.TabIndex = 97
        Me.lblUpgradeCases.Text = "Upgrade Cases:"
        '
        'chkIncludeDiskSize
        '
        Me.chkIncludeDiskSize.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.chkIncludeDiskSize.AutoSize = True
        Me.chkIncludeDiskSize.Location = New System.Drawing.Point(298, 522)
        Me.chkIncludeDiskSize.Name = "chkIncludeDiskSize"
        Me.chkIncludeDiskSize.Size = New System.Drawing.Size(108, 17)
        Me.chkIncludeDiskSize.TabIndex = 98
        Me.chkIncludeDiskSize.Text = "Include Disk Size"
        Me.chkIncludeDiskSize.UseVisualStyleBackColor = True
        '
        'CaseFinder
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(96.0!, 96.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Dpi
        Me.AutoScroll = True
        Me.AutoScrollMinSize = New System.Drawing.Size(800, 600)
        Me.ClientSize = New System.Drawing.Size(1290, 831)
        Me.Controls.Add(Me.chkIncludeDiskSize)
        Me.Controls.Add(Me.lblUpgradeCases)
        Me.Controls.Add(Me.cboUpgradeCasees)
        Me.Controls.Add(Me.cboCopyMoveCases)
        Me.Controls.Add(Me.chkRollUpReporting)
        Me.Controls.Add(Me.chkExportOnly)
        Me.Controls.Add(Me.btnSearchTermFile)
        Me.Controls.Add(Me.grpSearchTerm)
        Me.Controls.Add(Me.btnExportLocation)
        Me.Controls.Add(Me.txtExportLocation)
        Me.Controls.Add(Me.lblExportLocation)
        Me.Controls.Add(Me.btnStopProcessing)
        Me.Controls.Add(Me.lblFileLocation)
        Me.Controls.Add(Me.btnFileLocation)
        Me.Controls.Add(Me.txtCaseFileLocations)
        Me.Controls.Add(Me.grpCaseSelector)
        Me.Controls.Add(Me.lblShowSizeIn)
        Me.Controls.Add(Me.cboSizeReporting)
        Me.Controls.Add(Me.btnLoadConfig)
        Me.Controls.Add(Me.btnSaveConfig)
        Me.Controls.Add(Me.btnGetFileSystemData)
        Me.Controls.Add(Me.cboExportType)
        Me.Controls.Add(Me.chkExportSearchResults)
        Me.Controls.Add(Me.cboCalculateProcessingSpeeds)
        Me.Controls.Add(Me.lblCalculateProcessingSpeeds)
        Me.Controls.Add(Me.grpLicenseType)
        Me.Controls.Add(Me.btnBackupLocationChooser)
        Me.Controls.Add(Me.lblBackUpLocation)
        Me.Controls.Add(Me.txtBackupLocation)
        Me.Controls.Add(Me.chkBackUpCase)
        Me.Controls.Add(Me.btnNuixLogSelector)
        Me.Controls.Add(Me.txtNuixLogDir)
        Me.Controls.Add(Me.lblNuixLogDir)
        Me.Controls.Add(Me.btnLoadPreviousReportingRun)
        Me.Controls.Add(Me.lblNuixVersion)
        Me.Controls.Add(Me.cboNuixLicenseType)
        Me.Controls.Add(Me.lblNuixConsoleVersion)
        Me.Controls.Add(Me.btnConsoleLocation)
        Me.Controls.Add(Me.txtNuixConsoleLocation)
        Me.Controls.Add(Me.lblNuixVersionLocation)
        Me.Controls.Add(Me.btnExportCaseStatistics)
        Me.Controls.Add(Me.txtSearchTerm)
        Me.Controls.Add(Me.lblSearchTerm)
        Me.Controls.Add(Me.lblReportType)
        Me.Controls.Add(Me.cboReportType)
        Me.Controls.Add(Me.grdCaseInfo)
        Me.Controls.Add(Me.btnReportLocation)
        Me.Controls.Add(Me.lblReportLocation)
        Me.Controls.Add(Me.txtReportLocation)
        Me.Controls.Add(Me.btnGetData)
        Me.Controls.Add(Me.panelCaseDirectory)
        Me.Margin = New System.Windows.Forms.Padding(2)
        Me.MinimumSize = New System.Drawing.Size(1300, 850)
        Me.Name = "CaseFinder"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterParent
        Me.Text = "Universal Case Reporting tool"
        Me.panelCaseDirectory.ResumeLayout(False)
        Me.panelCaseDirectory.PerformLayout()
        CType(Me.grdCaseInfo, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ExportContextMenuStrip.ResumeLayout(False)
        Me.grpLicenseType.ResumeLayout(False)
        Me.grpLicenseType.PerformLayout()
        CType(Me.numExportWorkers, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.numExportWorkerMemory, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.numNuixAppMemory, System.ComponentModel.ISupportInitialize).EndInit()
        Me.grpCaseSelector.ResumeLayout(False)
        Me.grpCaseSelector.PerformLayout()
        Me.grpSearchTerm.ResumeLayout(False)
        Me.grpSearchTerm.PerformLayout()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents panelCaseDirectory As System.Windows.Forms.Panel
    Friend WithEvents lblNuixCaseDirectory As System.Windows.Forms.Label
    Friend WithEvents btnGetData As System.Windows.Forms.Button
    Friend WithEvents txtReportLocation As System.Windows.Forms.TextBox
    Friend WithEvents lblReportLocation As System.Windows.Forms.Label
    Friend WithEvents btnReportLocation As System.Windows.Forms.Button
    Friend WithEvents grdCaseInfo As System.Windows.Forms.DataGridView
    Friend WithEvents cboReportType As System.Windows.Forms.ComboBox
    Friend WithEvents lblReportType As System.Windows.Forms.Label
    Friend WithEvents treeViewFolders As System.Windows.Forms.TreeView
    Friend WithEvents ImgIcons As System.Windows.Forms.ImageList
    Friend WithEvents lblSearchTerm As System.Windows.Forms.Label
    Friend WithEvents txtSearchTerm As System.Windows.Forms.TextBox
    Friend WithEvents btnExportCaseStatistics As System.Windows.Forms.Button
    Friend WithEvents ExportContextMenuStrip As System.Windows.Forms.ContextMenuStrip
    Friend WithEvents CSVToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents JSonToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents XMLToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents lblNuixVersionLocation As System.Windows.Forms.Label
    Friend WithEvents txtNuixConsoleLocation As System.Windows.Forms.TextBox
    Friend WithEvents btnConsoleLocation As System.Windows.Forms.Button
    Friend WithEvents lblNuixConsoleVersion As System.Windows.Forms.Label
    Friend WithEvents cboNuixLicenseType As System.Windows.Forms.ComboBox
    Friend WithEvents lblNuixVersion As System.Windows.Forms.Label
    Friend WithEvents btnLoadPreviousReportingRun As System.Windows.Forms.Button
    Friend WithEvents lblNuixLogDir As System.Windows.Forms.Label
    Friend WithEvents txtNuixLogDir As System.Windows.Forms.TextBox
    Friend WithEvents btnNuixLogSelector As System.Windows.Forms.Button
    Friend WithEvents chkBackUpCase As System.Windows.Forms.CheckBox
    Friend WithEvents txtBackupLocation As System.Windows.Forms.TextBox
    Friend WithEvents lblBackUpLocation As System.Windows.Forms.Label
    Friend WithEvents btnBackupLocationChooser As System.Windows.Forms.Button
    Friend WithEvents grpLicenseType As System.Windows.Forms.GroupBox
    Friend WithEvents txtRegistryServer As System.Windows.Forms.TextBox
    Friend WithEvents lblRegistryServer As System.Windows.Forms.Label
    Friend WithEvents txtNMSInfo As System.Windows.Forms.TextBox
    Friend WithEvents lblNMSInfo As System.Windows.Forms.Label
    Friend WithEvents txtNMSUserName As System.Windows.Forms.TextBox
    Friend WithEvents txtNMSAddress As System.Windows.Forms.TextBox
    Friend WithEvents lblNMSAddress As System.Windows.Forms.Label
    Friend WithEvents cboLicenseType As System.Windows.Forms.ComboBox
    Friend WithEvents lblServerType As System.Windows.Forms.Label
    Friend WithEvents numNuixAppMemory As System.Windows.Forms.NumericUpDown
    Friend WithEvents lblNuixAppMemory As System.Windows.Forms.Label
    Friend WithEvents lblCalculateProcessingSpeeds As System.Windows.Forms.Label
    Friend WithEvents cboCalculateProcessingSpeeds As System.Windows.Forms.ComboBox
    Friend WithEvents chkExportSearchResults As System.Windows.Forms.CheckBox
    Friend WithEvents btnGetFileSystemData As System.Windows.Forms.Button
    Friend WithEvents ExcelToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents btnSaveConfig As System.Windows.Forms.Button
    Friend WithEvents btnLoadConfig As System.Windows.Forms.Button
    Friend WithEvents cboSizeReporting As System.Windows.Forms.ComboBox
    Friend WithEvents lblShowSizeIn As System.Windows.Forms.Label
    Friend WithEvents grpCaseSelector As System.Windows.Forms.GroupBox
    Friend WithEvents radFile As System.Windows.Forms.RadioButton
    Friend WithEvents radFileSystem As System.Windows.Forms.RadioButton
    Friend WithEvents txtCaseFileLocations As System.Windows.Forms.TextBox
    Friend WithEvents btnFileLocation As System.Windows.Forms.Button
    Friend WithEvents lblFileLocation As System.Windows.Forms.Label
    Friend WithEvents btnStopProcessing As System.Windows.Forms.Button
    Friend WithEvents lblExportLocation As System.Windows.Forms.Label
    Friend WithEvents txtExportLocation As System.Windows.Forms.TextBox
    Friend WithEvents btnExportLocation As System.Windows.Forms.Button
    Friend WithEvents cboExportType As System.Windows.Forms.ComboBox
    Friend WithEvents grpSearchTerm As System.Windows.Forms.GroupBox
    Friend WithEvents radSearchFile As System.Windows.Forms.RadioButton
    Friend WithEvents radSearchTerm As System.Windows.Forms.RadioButton
    Friend WithEvents btnSearchTermFile As System.Windows.Forms.Button
    Friend WithEvents chkExportOnly As System.Windows.Forms.CheckBox
    Friend WithEvents lblExportMemory As System.Windows.Forms.Label
    Friend WithEvents numExportWorkerMemory As System.Windows.Forms.NumericUpDown
    Friend WithEvents lblExportWorkers As System.Windows.Forms.Label
    Friend WithEvents numExportWorkers As System.Windows.Forms.NumericUpDown
    Friend WithEvents chkRollUpReporting As System.Windows.Forms.CheckBox
    Friend WithEvents cboCopyMoveCases As System.Windows.Forms.ComboBox
    Friend WithEvents cboUpgradeCasees As System.Windows.Forms.ComboBox
    Friend WithEvents lblUpgradeCases As System.Windows.Forms.Label
    Friend WithEvents CaseGuid As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents CollectionStatus As System.Windows.Forms.DataGridViewComboBoxColumn
    Friend WithEvents PercentComplete As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ReportLoadDuration As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents CaseName As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents CurrentCaseVersion As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents UpgradedCaseVersion As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents BatchLoadInfo As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents DataExport As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents CaseLocation As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents BackUpLocation As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents CaseDescription As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents CaseSizeOnDisk As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents CaseFileSize As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents CaseAuditSize As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents OldestTopLevel As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents NewestTopLevel As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents IsCompound As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents CasesContained As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ContainedInCase As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents Investigator As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents InvestigatorSessions As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents InvalidSessions As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents InvestigatorTimeSummary As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents BrokerMemory As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents WorkerCount As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents WorkerMemory As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents EvidenceName As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents EvidenceLocation As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents EvidenceDescription As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents EvidenceCustomMetadata As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents LanguagesContained As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents MimeTypes As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ItemTypes As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents IrregularItems As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents CreationDate As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ModifiedDate As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents LoadStartDate As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents LoadEndDate As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents LoadTime As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents LoadEvents As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents TotalLoadTime As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ProcessingSpeed As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents TotalCaseItemCount As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ItemCounts As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents OriginalItems As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents DuplicateItems As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents Custodians As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents CustodianCount As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents SearchTerm As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents SearchSize As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents SearchHitCount As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents CustodianSearchHit As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents HitCountPercent As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents NuixLogLocation As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents chkIncludeDiskSize As System.Windows.Forms.CheckBox

End Class
