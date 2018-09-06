Imports System.IO
Imports System.Text.RegularExpressions
Imports System.Xml
Imports System.Xml.Serialization
Imports System.Xml.Linq
Imports System
Imports System.IO.StreamReader
Imports System.Collections
Imports System.Threading
Imports System.Globalization
Imports System.Data.SQLite
Imports System.Reflection
Imports Microsoft.Office.Interop
Imports System.Configuration


Public Class CaseFinder
    Public plstSoureFolders As List(Of String)
    Public pbNoMoreJobs As Boolean
    Public psReportType As String
    Public psLicenseType As String
    Public pDataLoaded As Boolean
    Public psUCRTLogFile As String
    Public psShowSizeIn As String

    Private ReportFormUpdate As System.Threading.Thread
    Private SQLiteDBReadThread As System.Threading.Thread

    Private Sub btnReportLocation_Click(sender As Object, e As EventArgs) Handles btnReportLocation.Click
        Dim fldrBrowserDialog As New FolderBrowserDialog

        If (fldrBrowserDialog.ShowDialog = System.Windows.Forms.DialogResult.OK) Then

            txtReportLocation.Text = fldrBrowserDialog.SelectedPath

        End If
    End Sub

    Private Sub btnGetData_Click(sender As Object, e As EventArgs) Handles btnGetData.Click
        Dim bStatus As Boolean
        Dim lstMailBoxTotals As New List(Of String)
        Dim lstExporterMetrics As New List(Of String)
        Dim lstCreatedPST As New List(Of String)
        Dim lstCreatedZIP As New List(Of String)
        Dim lstPSTCustodianName As New List(Of String)
        Dim sNuixConsoleVersion As String
        Dim sSearchTerm As String
        Dim sNuixAppMemory As String
        Dim sScriptsDirectory As String
        Dim msgboxReturn As DialogResult
        Dim bReloadSelectedInfo As Boolean
        Dim sReportFilePath As String
        Dim asCaseFolders() As String
        Dim bMigrateCases As Boolean
        Dim bBackUpCases As Boolean
        Dim bIncludeDiskSize As Boolean
        Dim sBackUpLocation As String
        Dim sMachineName As String
        Dim sShowSizeIn As String
        Dim iCounter As Integer
        Dim NuixCaseFileCSV As StreamReader
        Dim bCopyCases As Boolean
        Dim sCaseName As String
        Dim sCaseDirectory As String
        Dim sCollectionStatus As String
        Dim sBackUpCaseLocation As String
        Dim lstNuixCases As List(Of String)
        Dim lstCaseGUIDs As List(Of String)
        Dim value As System.Version = My.Application.Info.Version
        Dim dblCaseSizeOnDisk As Double


        lstNuixCases = New List(Of String)
        lstCaseGUIDs = New List(Of String)

        bIncludeDiskSize = chkIncludeDiskSize.Checked
        grdCaseInfo.Rows.Clear()

        Try
            If pbNoMoreJobs = False Then
                MessageBox.Show("Currently collecting data please wait until processing is completed to get get data.", "Case Data Still Processing.")
                Exit Sub
            End If

            If cboReportType.Text = vbNullString Then
                MessageBox.Show("You must select the report to run.", "Report Type Not selected")
                cboReportType.Focus()
                Exit Sub
            End If
            If txtNuixLogDir.Text = vbNullString Then
                MessageBox.Show("You must select the location of the Nuix Log Files.", "Nuix Log File Directory not selected")
                txtNuixLogDir.Focus()
                Exit Sub
            End If

            If (txtReportLocation.Text = vbNullString) Then
                MsgBox("You Must Enter the location to create the Case report file.")
                txtReportLocation.Focus()
                Exit Sub
            Else
                sReportFilePath = txtReportLocation.Text
                sScriptsDirectory = sReportFilePath & "\" & "Scripts"
                Directory.CreateDirectory(sScriptsDirectory)

                bStatus = blnBuildSQLiteDB(sScriptsDirectory)
                bStatus = blnBuildSQLiteDatabaseScript(sScriptsDirectory)
                bStatus = blnBuildSQLiteRubyScript(sScriptsDirectory)
            End If
            If (txtNuixConsoleLocation.Text = vbNullString) Then
                MessageBox.Show("You must enter the location of the Nuix Console version to use.", "Nuix Console Version Location")
                txtNuixConsoleLocation.Focus()
                Exit Sub
            Else
                sNuixConsoleVersion = lblNuixConsoleVersion.Text.Replace("Nuix Console Version: ", "")
            End If

            If cboLicenseType.Text = "Server" Then
                If txtNMSAddress.Text = vbNullString Then
                    MessageBox.Show("You must enter the appropriate NMS Address.", "NMS Address not entered")
                    txtNMSAddress.Focus()
                    Exit Sub
                End If
                If txtNMSUserName.Text = vbNullString Then
                    MessageBox.Show("You must enter the appropriate NMS Username.", "NMS Username not entered")
                    txtNMSUserName.Focus()
                    Exit Sub
                End If

                If txtNMSInfo.Text = vbNullString Then
                    MessageBox.Show("You must enter the appropriate NMS Info.", "NMS Info not entered")
                    txtNMSInfo.Focus()
                    Exit Sub
                End If
            End If

            If cboNuixLicenseType.Text = vbNullString Then
                MessageBox.Show("You must select the appropriate license to use.", "Select appropriate license")
                cboNuixLicenseType.Focus()
                Exit Sub
            End If

            If radFile.Checked = True Then
                If txtCaseFileLocations.Text = vbNullString Then
                    MessageBox.Show("You must select a CSV file containing paths to Nuix Case files", "Case file CSV required", MessageBoxButtons.OK)
                    txtCaseFileLocations.Focus()
                    Exit Sub
                Else
                    NuixCaseFileCSV = New StreamReader(txtCaseFileLocations.Text)
                    While Not NuixCaseFileCSV.EndOfStream
                        plstSoureFolders.Add(NuixCaseFileCSV.ReadLine)
                    End While
                End If
            End If

            If plstSoureFolders.Count = 0 Then
                MessageBox.Show("You must select at least one folder to search for Nuix Cases in.", "Select Nuix Case Directory")
                Exit Sub
            End If

            If grdCaseInfo.Rows.Count > 1 Then
                msgboxReturn = MessageBox.Show("Would you like to search for new case data on the file system (Yes) or reload selected case data (No)?", "Case data selection", MessageBoxButtons.YesNoCancel)
                If msgboxReturn = vbYes Then
                    grdCaseInfo.Rows.Clear()
                    bReloadSelectedInfo = False
                ElseIf msgboxReturn = vbNo Then
                    bReloadSelectedInfo = True
                Else
                    Exit Sub
                End If
            End If

            If cboUpgradeCasees.Text = "No" Then
                bMigrateCases = False
            ElseIf cboUpgradeCasees.Text = "Upgrade and Report" Then
                bMigrateCases = True

            ElseIf cboUpgradeCasees.Text = "Upgrade Only" Then
                bMigrateCases = True
            End If

            If chkBackUpCase.Checked = True Then
                If cboCopyMoveCases.Text = vbNullString Then
                    MessageBox.Show("You must select whether you want to copy or move the cases to the backup location.", "Copy or Move cases to backup location", MessageBoxButtons.OK)
                    Exit Sub
                Else
                    bBackUpCases = True
                    If txtBackupLocation.Text = vbNullString Then
                        MessageBox.Show("You have not selected a back up location.", "Select backup location", MessageBoxButtons.OK)
                        txtBackupLocation.Focus()
                        Exit Sub
                    Else
                        sBackUpLocation = txtBackupLocation.Text
                    End If
                    If cboCopyMoveCases.Text = "Copy" Then
                        bCopyCases = True
                    Else
                        bCopyCases = False
                    End If
                End If
            Else
                bBackUpCases = False
                sBackUpLocation = ""
            End If

            If cboCalculateProcessingSpeeds.Visible = True Then
                If cboCalculateProcessingSpeeds.Text = vbNullString Then
                    MessageBox.Show("You must select a size value to calculate processing speeds.", "Select Processing Speed Calculation", MessageBoxButtons.OK)
                    cboCalculateProcessingSpeeds.Focus()
                    Exit Sub
                End If
            End If

            'If chkExportSearchResults.Checked = True Then
            '    If cboExportType.Text = vbNullString Then
            '        MessageBox.Show("You must select an export type to export the search results.", "Export Search Results", MessageBoxButtons.OK)
            '        cboExportType.Focus()
            '        Exit Sub
            '    ElseIf cboExportType.Text = "Case Subset" Then
            '        MessageBox.Show("Currently Case Subset is not supported.", "Export Search Results", MessageBoxButtons.OK)
            '        cboExportType.Focus()
            '        Exit Sub
            '    End If

            'End If

            If cboSizeReporting.Text = vbNullString Then
                MessageBox.Show("You must select the size value to report in", "Show Size in", MessageBoxButtons.OK)
                cboSizeReporting.Focus()
                Exit Sub
            Else
                sShowSizeIn = cboSizeReporting.Text
            End If

            If chkExportSearchResults.Checked = True Then
                If cboExportType.Text = vbNullString Then
                    MessageBox.Show("You must select an export type if you want to export the search results.", "Select Export Type", MessageBoxButtons.OK)
                    cboExportType.Focus()
                    Exit Sub

                End If
            End If
            sSearchTerm = Me.txtSearchTerm.Text

            sNuixAppMemory = "-Xmx" & numNuixAppMemory.Value.ToString & "g"


            asCaseFolders = plstSoureFolders.ToArray
            sMachineName = System.Net.Dns.GetHostName

            psUCRTLogFile = sScriptsDirectory & "\UCRT Log - " & sMachineName & "-" & DateTime.Now.ToString("yyyy-MM-dd hh-mm-ss") & ".log"

            Logger(psUCRTLogFile, "Nuix Universal Case Reporting Tool - " & psUCRTLogFile)
            iCounter = 0
            For Each sourceLocation In plstSoureFolders
                Logger(psUCRTLogFile, "Source - " & iCounter & " - " & sourceLocation.ToString)
                iCounter = iCounter + 1
            Next
            Logger(psUCRTLogFile, "Report Type - " & cboReportType.Text)
            Logger(psUCRTLogFile, "Report Location - " & txtReportLocation.Text)
            Logger(psUCRTLogFile, "Search Term - " & txtSearchTerm.Text)
            Logger(psUCRTLogFile, "Nuix Log Directory - " & txtNuixLogDir.Text)
            Logger(psUCRTLogFile, "Nuix Console Location - " & txtNuixConsoleLocation.Text)
            Logger(psUCRTLogFile, "Nuix License - " & cboNuixLicenseType.Text)
            Logger(psUCRTLogFile, "Nuix Console Location - " & txtNuixConsoleLocation.Text)
            Logger(psUCRTLogFile, "Server Type - " & cboLicenseType.Text)
            Logger(psUCRTLogFile, "NMS Address - " & txtNMSAddress.Text)
            Logger(psUCRTLogFile, "NMS Username - " & txtNMSUserName.Text)
            Logger(psUCRTLogFile, "NMS Registry Server- " & txtRegistryServer.Text)
            Logger(psUCRTLogFile, "Nuix Application Memory - " & numNuixAppMemory.Text)
            Logger(psUCRTLogFile, "Upgrade Case Version Mismatch - " & cboUpgradeCasees.Text)
            Logger(psUCRTLogFile, "Backup Cases - " & chkBackUpCase.Checked.ToString)
            Logger(psUCRTLogFile, "Backup Cases Location - " & txtBackupLocation.Text)

            Try
                If (bReloadSelectedInfo = False) Then
                    If Not IsNothing(asCaseFolders) Then
                        bStatus = blnGetAllNuixCaseFiles(sScriptsDirectory, sNuixConsoleVersion, asCaseFolders, bMigrateCases, False, bIncludeDiskSize, lstNuixCases, lstCaseGUIDs)
                        If chkBackUpCase.Checked = True Then
                            Me.Text = "Universal Case Reporting tool - " & value.ToString & " - Copying/Moving Cases - Please Wait"

                            For Each NuixGUID In lstCaseGUIDs
                                bStatus = blnGetUpdatedDBInfo(sScriptsDirectory, NuixGUID.ToString, "CaseName", sCaseName, "TEXT")
                                bStatus = blnGetUpdatedDBInfo(sScriptsDirectory, NuixGUID.ToString, "CaseLocation", sCaseDirectory, "TEXT")
                                bStatus = blnGetUpdatedDBInfo(sScriptsDirectory, NuixGUID.ToString, "CaseSizeOnDisk", dblCaseSizeOnDisk, "INT")
                                bStatus = blnCopyCase(sCaseName, sCaseDirectory, sBackUpLocation, bCopyCases, sBackUpCaseLocation, dblCaseSizeOnDisk)
                                sCollectionStatus = "Case Copied or Moved"
                                bStatus = blnUpdateSQLiteReportingDB(sScriptsDirectory, NuixGUID.ToString, "BackUpLocation", sBackUpCaseLocation)
                            Next

                        End If
                    End If
                    Me.Text = "Universal Case Reporting tool - " & value.ToString

                    bStatus = blnPopulateCaseInfoGrid(grdCaseInfo, sScriptsDirectory, sShowSizeIn, "'File System Info Collected - Case Migrating', 'File System Info Collected - Case Version Mismatch', 'File System Info Collected - Waiting for Case Data', 'Case Locked'")

                    psReportType = cboReportType.Text

                    pDataLoaded = False

                    If cboCopyMoveCases.Text <> "Move" Then
                        ReportFormUpdate = New System.Threading.Thread(AddressOf Me.BuildNuixAllDataReports)
                        ReportFormUpdate.Start()
                    Else
                        MessageBox.Show("You cannot move the Cases and collect the case data.  If necessary rerun the case stats on the " & txtBackupLocation.Text & " location", "Cannot move case data and collect statisics", MessageBoxButtons.OK)
                        Exit Sub
                    End If

                Else
                    ReportFormUpdate = New System.Threading.Thread(AddressOf Me.UpdateNuixAllDataReports)
                    ReportFormUpdate.Start()
                End If

            Catch ex As Exception
                Logger(psUCRTLogFile, ex.ToString)
            End Try
        Catch ex As Exception
            MessageBox.Show("Get Data Click - " & ex.ToString)
            Logger(psUCRTLogFile, ex.ToString)
        End Try

    End Sub

    Private Function blnPopulateCaseInfoGrid(ByVal grdCaseInfo As DataGridView, ByVal sSQLiteDBLocation As String, ByVal sShowSizeIn As String, ByVal sCollectionStatus As String) As Boolean
        blnPopulateCaseInfoGrid = False
        Dim mSQL As String
        Dim dt As DataTable
        Dim ds As DataSet
        Dim dataReader As SQLiteDataReader
        Dim sqlCommand As SQLiteCommand
        Dim sqlConnection As SQLiteConnection

        Dim sCaseGUID As String
        Dim sCaseName As String
        Dim sReportLoadDuration As String
        Dim sBatchLoadInfo As String
        Dim sCaseLocation As String
        Dim sCurrentCaseVersion As String
        Dim sUpgradedCaseVersion As String
        Dim dCaseSizeOnDisk As Double
        Dim dCaseFileSize As Double
        Dim dCaseAuditSize As Double
        Dim sIsCompound As String
        Dim sCasesContained As String
        Dim sContainedInCase As String
        Dim sInvestigator As String
        Dim sInvestigatorSessions As String
        Dim sInvalidSessions As String
        Dim sInvestigatorTimeSummary As String
        Dim iBrokerMemory As Integer
        Dim iWorkerCount As Integer
        Dim iWorkerMemory As Integer
        Dim sEvidenceProcessed As String
        Dim sEvidenceLocation As String
        Dim sEvidenceCustomMetadata As String
        Dim sMimeTypes As String
        Dim sItemTypes As String
        Dim sCreationDate As String
        Dim sModifiedDate As String
        Dim sLoadDataStart As String
        Dim sLoadDataEnd As String
        Dim sLoadEvents As String
        Dim sDataExport As String
        Dim sTotalLoadTime As String
        Dim sProcessingSpeed As String
        Dim sCustodians As String
        Dim sSearchTerm As String
        Dim dSearchSize As Double
        Dim dHitCount As Double
        Dim dHitCountPercentage As Decimal
        Dim sCustodianSearchHit As String
        Dim sReportLoadTime As String
        Dim sNuixLogFile As String
        Dim sOldestTopLevel As String
        Dim sNewestTopLevel As String
        Dim sLanguages As String
        Dim sCaseDescription As String
        Dim sEvidenceDescription As String
        Dim sIrregularItems As String
        Dim iCaseGUIDcol As Integer
        Dim iCaseNamecol As Integer
        Dim iCollectionStatuscol As Integer
        Dim iReportLoadDurationcol As Integer
        Dim iBatchLoadInfocol As Integer
        Dim iDataExportCol As Integer
        Dim iCaseLocationcol As Integer
        Dim iCurrentCaseVersioncol As Integer
        Dim iUpgradedCaseVersioncol As Integer
        Dim iCaseSizeOnDiskcol As Integer
        Dim iCaseFileSizecol As Integer
        Dim iCaseAuditSizecol As Integer
        Dim iIsCompoundcol As Integer
        Dim iCasesContainedcol As Integer
        Dim iContainedInCasecol As Integer
        Dim iInvestigatorcol As Integer
        Dim iBrokerMemorycol As Integer
        Dim iWorkerCountcol As Integer
        Dim iWorkerMemorycol As Integer
        Dim iEvidenceProcessedcol As Integer
        Dim iEvidenceLocationcol As Integer
        Dim iEvidenceCustomMetadataCol As Integer
        Dim iMimeTypescol As Integer
        Dim iItemTypescol As Integer
        Dim iCreationDatecol As Integer
        Dim iModifiedDatecol As Integer
        Dim iLoadDataStartcol As Integer
        Dim iLoadDataEndcol As Integer
        Dim iLoadTimecol As Integer
        Dim iLoadEventscol As Integer
        Dim iTotalLoadTimecol As Integer
        Dim iProcessingSpeedcol As Integer
        Dim iCustodianscol As Integer
        Dim iCustodianCountcol As Integer
        Dim iSearchTermcol As Integer
        Dim iSearchSizecol As Integer
        Dim iHitCountcol As Integer
        Dim iCustodianSearchHitcol As Integer
        Dim iItemCountcol As Integer
        Dim iReportLoadTimecol As Integer
        Dim iOldestTopLevelCol As Integer
        Dim iNewestTopLevelCol As Integer
        Dim iLanguagescol As Integer
        Dim iCaseDescriptioncol As Integer
        Dim iEvidenceDescriptioncol As Integer
        Dim iIrregularItemscol As Integer
        Dim iNuixLogFileCol As Integer
        Dim iModifyDateCompare As Integer
        Dim iRowIndex As Integer
        Dim lSeconds As Long
        Dim lHours As Long
        Dim lMinutes As Long
        Dim sCaseModifiedDate As String
        Dim sHitCountPercentage As String
        Dim sLoadTimeHMS As String
        Dim iPercentCompleteCol As Integer
        Dim iPercentComplete As Integer
        Dim iDuplicateItemCol As Integer
        Dim sDuplicateItems As String
        Dim iOriginalItemCol As Integer
        Dim sOriginaltems As String
        Dim iItemCountsCol As Integer
        Dim sItemCounts As String
        Dim iTotalItemCountCol As Integer
        Dim dTotalItemCount As Integer
        Dim iBackUpLocationCol As Integer
        Dim sBackUpLocation As String

        Dim sLoadDatePart As String
        Dim sLoadTimePart As String
        Dim iLoadTime As Integer
        Dim dLoadDateStart As Date
        Dim dLoadDateEnd As Date
        Dim asDateParts() As String
        Dim iCustodianCount As Integer
        Dim bStatus As Boolean
        Dim CaseInfoRow As DataGridViewRow
        Dim asNuixVersionNumber() As String
        Dim sExportItems As String

        Try

            dt = Nothing
            ds = New DataSet
            sqlConnection = New SQLiteConnection("Data Source=" & sSQLiteDBLocation & "\NuixCaseReports.db3;Version=3;Read Only=True;New=False;Compress=True;")

            If sCollectionStatus = "%" Then
                mSQL = "select CaseGUID, PercentComplete, ReportLoadStart, ReportLoadEnd, CaseName, CollectionStatus, ReportLoadDuration, BatchLoadInfo, CaseLocation, BackUpLocation, CurrentCaseVersion, UpgradedCaseVersion, CaseSizeOnDisk, CaseFileSize, CaseAuditSize, IsCompound, CasesContained, ContainedInCase, Investigator, InvestigatorSessions, InvalidSessions, InvestigatorTimeSummary, BrokerMemory, WorkerCount, WorkerMemory, EvidenceProcessed, EvidenceLocation, EvidenceCustomMetadata, MimeTypes, ItemTypes, CreationDate, ModifiedDate, LoadDataStart, LoadDataEnd, TotalLoadTime, LoadEvents, TotalLoadTime, ProcessingSpeed, Custodians, CustodianCount, SearchTerm, SearchSize, HitCount, CustodianSearchHit,TotalItemCount, ItemCounts, OriginalItems, DuplicateItems, CaseUsers, ReportLoadTime, NuixLogLocation, OldestItem, NewestItem, Languages, CustomMetadata, CaseDescription, EvidenceDescription, IrregularItems from NuixReportingInfo where CollectionStatus Like '" & sCollectionStatus & "'"
            Else
                mSQL = "select CaseGUID, PercentComplete, ReportLoadStart, ReportLoadEnd, CaseName, CollectionStatus, ReportLoadDuration, BatchLoadInfo, CaseLocation, BackUpLocation, CurrentCaseVersion, UpgradedCaseVersion, CaseSizeOnDisk, CaseFileSize, CaseAuditSize, IsCompound, CasesContained, ContainedInCase, Investigator, InvestigatorSessions, InvalidSessions, InvestigatorTimeSummary, BrokerMemory, WorkerCount, WorkerMemory, EvidenceProcessed, EvidenceLocation, EvidenceCustomMetadata, MimeTypes, ItemTypes, CreationDate, ModifiedDate, LoadDataStart, LoadDataEnd, TotalLoadTime, LoadEvents, TotalLoadTime, ProcessingSpeed, Custodians, CustodianCount, SearchTerm, SearchSize, HitCount, CustodianSearchHit,TotalItemCount, ItemCounts, OriginalItems, DuplicateItems, CaseUsers, ReportLoadTime, NuixLogLocation, OldestItem, NewestItem, Languages, CustomMetadata, CaseDescription, EvidenceDescription, IrregularItems from NuixReportingInfo where CollectionStatus in (" & sCollectionStatus & ")"
            End If
            sqlCommand = New SQLiteCommand(mSQL, sqlConnection)
            sqlConnection.Open()

            dataReader = sqlCommand.ExecuteReader

            While dataReader.Read
                iCaseGUIDcol = dataReader.GetOrdinal("CaseGUID")
                If dataReader.IsDBNull(iCaseGUIDcol) Then
                    sCaseGUID = vbNullString
                Else
                    sCaseGUID = dataReader.GetString(iCaseGUIDcol)
                End If

                iCollectionStatuscol = dataReader.GetOrdinal("CollectionStatus")
                If dataReader.IsDBNull(iCollectionStatuscol) Then
                    sCollectionStatus = vbNullString
                Else
                    sCollectionStatus = dataReader.GetString(iCollectionStatuscol)
                End If

                iPercentCompleteCol = dataReader.GetOrdinal("PercentComplete")
                If dataReader.IsDBNull(iCollectionStatuscol) Then
                    iPercentComplete = 0
                Else
                    iPercentComplete = dataReader.GetInt16(iPercentCompleteCol)
                End If

                iCaseNamecol = dataReader.GetOrdinal("CaseName")
                If dataReader.IsDBNull(iCaseNamecol) Then
                    sCaseName = vbNullString
                Else
                    sCaseName = dataReader.GetString(iCaseNamecol)
                End If

                iReportLoadDurationcol = dataReader.GetOrdinal("ReportLoadDuration")
                If dataReader.IsDBNull(iReportLoadDurationcol) Then
                    sReportLoadDuration = vbNullString
                Else
                    sReportLoadDuration = dataReader.GetString(iReportLoadDurationcol)
                End If

                iBatchLoadInfocol = dataReader.GetOrdinal("BatchLoadInfo")
                If dataReader.IsDBNull(iBatchLoadInfocol) Then
                    sBatchLoadInfo = vbNullString
                Else
                    sBatchLoadInfo = dataReader.GetString(iBatchLoadInfocol)
                End If

                iIsCompoundcol = dataReader.GetOrdinal("IsCompound")
                If dataReader.IsDBNull(iIsCompoundcol) Then
                    sIsCompound = vbNullString
                Else
                    sIsCompound = dataReader.GetString(iIsCompoundcol)
                End If

                iCasesContainedcol = dataReader.GetOrdinal("CasesContained")
                If dataReader.IsDBNull(iCasesContainedcol) Then
                    sCasesContained = vbNullString
                Else
                    sCasesContained = dataReader.GetString(iCasesContainedcol)
                End If

                iContainedInCasecol = dataReader.GetOrdinal("ContainedInCase")
                If dataReader.IsDBNull(iContainedInCasecol) Then
                    sContainedInCase = vbNullString
                Else
                    sContainedInCase = dataReader.GetString(iContainedInCasecol)
                End If

                iCaseLocationcol = dataReader.GetOrdinal("CaseLocation")
                If dataReader.IsDBNull(iCaseLocationcol) Then
                    sCaseLocation = vbNullString
                Else
                    sCaseLocation = dataReader.GetString(iCaseLocationcol)
                End If

                iBackUpLocationCol = dataReader.GetOrdinal("BackUpLocation")
                If dataReader.IsDBNull(iBackUpLocationCol) Then
                    sBackUpLocation = vbNullString
                Else
                    sBackUpLocation = dataReader.GetString(iBackUpLocationCol)
                End If

                iCurrentCaseVersioncol = dataReader.GetOrdinal("CurrentCaseVersion")
                If dataReader.IsDBNull(iCurrentCaseVersioncol) Then
                    sCurrentCaseVersion = vbNullString
                Else
                    sCurrentCaseVersion = dataReader.GetString(iCurrentCaseVersioncol)
                End If

                iUpgradedCaseVersioncol = dataReader.GetOrdinal("UpgradedCaseVersion")
                If dataReader.IsDBNull(iUpgradedCaseVersioncol) Then
                    sUpgradedCaseVersion = vbNullString
                Else
                    sUpgradedCaseVersion = dataReader.GetString(iUpgradedCaseVersioncol)
                End If
                iCaseSizeOnDiskcol = dataReader.GetOrdinal("CaseSizeOnDisk")
                If dataReader.IsDBNull(iCaseSizeOnDiskcol) Then
                    dCaseSizeOnDisk = vbNullString
                Else
                    dCaseSizeOnDisk = dataReader.GetInt64(iCaseSizeOnDiskcol)
                End If

                iCaseFileSizecol = dataReader.GetOrdinal("CaseFileSize")
                If dataReader.IsDBNull(iCaseFileSizecol) Then
                    dCaseFileSize = 0.0
                Else
                    dCaseFileSize = dataReader.GetInt64(iCaseFileSizecol)
                End If

                iCaseAuditSizecol = dataReader.GetOrdinal("CaseAuditSize")
                If dataReader.IsDBNull(iCaseAuditSizecol) Then
                    dCaseAuditSize = 0.0
                Else
                    dCaseAuditSize = dataReader.GetInt64(iCaseAuditSizecol)
                End If

                iInvestigatorcol = dataReader.GetOrdinal("Investigator")
                If dataReader.IsDBNull(iInvestigatorcol) Then
                    sInvestigator = vbNullString
                Else
                    sInvestigator = dataReader.GetString(iInvestigatorcol)
                End If

                sInvestigatorSessions = vbNullString
                sInvestigatorTimeSummary = vbNullString
                sExportItems = vbNullString
                bStatus = blnGetInvestigatorSessions(sSQLiteDBLocation, sCaseGUID, sInvestigatorSessions, sInvalidSessions)
                bStatus = blnGetInvestigatorTimeSummary(sSQLiteDBLocation, sCaseGUID, sInvestigatorTimeSummary)
                bStatus = blnGetExportItemsInfo(sSQLiteDBLocation, sCaseGUID, sExportItems)

                iBrokerMemorycol = dataReader.GetOrdinal("BrokerMemory")
                If dataReader.IsDBNull(iBrokerMemorycol) Then
                    iBrokerMemory = 0
                Else
                    iBrokerMemory = dataReader.GetInt16(iBrokerMemorycol)
                End If

                iWorkerCountcol = dataReader.GetOrdinal("WorkerCount")
                If dataReader.IsDBNull(iWorkerCountcol) Then
                    iWorkerCount = 0
                Else
                    iWorkerCount = dataReader.GetInt16(iWorkerCountcol)
                End If

                iWorkerMemorycol = dataReader.GetOrdinal("WorkerMemory")
                If dataReader.IsDBNull(iWorkerMemorycol) Then
                    iWorkerMemory = 0
                Else
                    iWorkerMemory = dataReader.GetInt16(iWorkerMemorycol)
                End If

                iEvidenceProcessedcol = dataReader.GetOrdinal("EvidenceProcessed")
                If dataReader.IsDBNull(iEvidenceProcessedcol) Then
                    sEvidenceProcessed = vbNullString
                Else
                    sEvidenceProcessed = dataReader.GetString(iEvidenceProcessedcol)
                End If

                iEvidenceLocationcol = dataReader.GetOrdinal("EvidenceLocation")
                If dataReader.IsDBNull(iEvidenceLocationcol) Then
                    sEvidenceLocation = vbNullString
                Else
                    sEvidenceLocation = dataReader.GetString(iEvidenceLocationcol)
                End If

                iEvidenceCustomMetadataCol = dataReader.GetOrdinal("EvidenceCustomMetadata")
                If dataReader.IsDBNull(iEvidenceCustomMetadataCol) Then
                    sEvidenceCustomMetadata = vbNullString
                Else
                    sEvidenceCustomMetadata = dataReader.GetString(iEvidenceCustomMetadataCol)
                End If

                iMimeTypescol = dataReader.GetOrdinal("MimeTypes")
                If dataReader.IsDBNull(iMimeTypescol) Then
                    sMimeTypes = vbNullString
                Else
                    sMimeTypes = dataReader.GetString(iMimeTypescol)
                End If

                iTotalItemCountCol = dataReader.GetOrdinal("TotalItemCount")
                If dataReader.IsDBNull(iTotalItemCountCol) Then
                    dTotalItemCount = 0.0
                Else
                    dTotalItemCount = dataReader.GetValue(iTotalItemCountCol)
                End If

                iItemCountsCol = dataReader.GetOrdinal("ItemCounts")
                If dataReader.IsDBNull(iItemCountcol) Then
                    sItemCounts = ""
                Else
                    sItemCounts = dataReader.GetValue(iItemCountsCol)
                End If

                iDuplicateItemCol = dataReader.GetOrdinal("DuplicateItems")
                If dataReader.IsDBNull(iDuplicateItemCol) Then
                    sDuplicateItems = ""
                Else
                    sDuplicateItems = dataReader.GetValue(iDuplicateItemCol)
                End If

                iOriginalItemCol = dataReader.GetOrdinal("OriginalItems")
                If dataReader.IsDBNull(iOriginalItemCol) Then
                    sOriginaltems = ""
                Else
                    sOriginaltems = dataReader.GetValue(iOriginalItemCol)
                End If

                iItemTypescol = dataReader.GetOrdinal("ItemTypes")
                If dataReader.IsDBNull(iItemTypescol) Then
                    sItemTypes = vbNullString
                Else
                    sItemTypes = dataReader.GetString(iItemTypescol)
                End If

                iCreationDatecol = dataReader.GetOrdinal("CreationDate")
                If dataReader.IsDBNull(iCreationDatecol) Then
                    sCreationDate = vbNullString
                Else
                    sCreationDate = dataReader.GetString(iCreationDatecol)
                End If

                iModifiedDatecol = dataReader.GetOrdinal("ModifiedDate")
                If dataReader.IsDBNull(iModifiedDatecol) Then
                    sModifiedDate = vbNullString
                Else
                    sModifiedDate = dataReader.GetString(iModifiedDatecol)
                End If

                iLoadDataStartcol = dataReader.GetOrdinal("LoadDataStart")
                If dataReader.IsDBNull(iLoadDataStartcol) Then
                    sLoadDataStart = vbNullString
                Else
                    sLoadDataStart = dataReader.GetString(iLoadDataStartcol)
                    If sLoadDataStart.Contains("T") Then
                        asDateParts = Split(sLoadDataStart, "T")
                        sLoadDatePart = asDateParts(0)
                        If asDateParts(1).Contains(",") Then
                            sLoadTimePart = asDateParts(1).Substring(0, asDateParts(1).IndexOf(","))
                        End If
                        If (asDateParts(1).Contains(" + ")) Then
                            sLoadTimePart = asDateParts(1).Substring(0, asDateParts(1).IndexOf(" + "))
                        End If
                        sLoadDataStart = sLoadDatePart & " " & sLoadTimePart
                        dLoadDateStart = Date.Parse(sLoadDataStart, System.Globalization.DateTimeFormatInfo.InvariantInfo, Globalization.DateTimeStyles.None)
                        sLoadDataStart = dLoadDateStart.ToString
                        ' bStatus = blnUpdateSQLiteReportingDB(sSQLiteDBLocation, sCaseGUID, sCaseName, "LoadDataStart", sLoadDataStart)
                    End If
                End If

                iLoadDataEndcol = dataReader.GetOrdinal("LoadDataEnd")
                If dataReader.IsDBNull(iLoadDataEndcol) Then
                    sLoadDataEnd = vbNullString
                Else
                    sLoadDataEnd = dataReader.GetString(iLoadDataEndcol)
                    If sLoadDataEnd.Contains("T") Then
                        asDateParts = Split(sLoadDataEnd, "T")
                        sLoadDatePart = asDateParts(0)
                        If asDateParts(1).Contains(",") Then
                            sLoadTimePart = asDateParts(1).Substring(0, asDateParts(1).IndexOf(","))
                        End If
                        If asDateParts(1).Contains(" + ") Then
                            sLoadTimePart = asDateParts(1).Substring(0, asDateParts(1).IndexOf(" + "))
                        End If
                        sLoadDataEnd = sLoadDatePart & " " & sLoadTimePart
                        dLoadDateEnd = Date.Parse(sLoadDataEnd, System.Globalization.DateTimeFormatInfo.InvariantInfo, Globalization.DateTimeStyles.None)
                        sLoadDataEnd = dLoadDateEnd.ToString
                        'bStatus = blnUpdateSQLiteReportingDB(sSQLiteDBLocation, sCaseGUID, sCaseName, "LoadDataEnd", sLoadDataEnd)
                    End If
                End If

                iLoadTimecol = dataReader.GetOrdinal("LoadTime")
                If dataReader.IsDBNull(iLoadTimecol) Then
                    iLoadTime = 0
                Else
                    iLoadTime = dataReader.GetValue(iLoadTimecol)
                End If

                iLoadEventscol = dataReader.GetOrdinal("LoadEvents")
                If dataReader.IsDBNull(iLoadEventscol) Then
                    sLoadEvents = 0
                Else
                    sLoadEvents = dataReader.GetValue(iLoadEventscol)
                End If

                iTotalLoadTimecol = dataReader.GetOrdinal("TotalLoadTime")
                If dataReader.IsDBNull(iTotalLoadTimecol) Then
                    sTotalLoadTime = "0"
                Else
                    sTotalLoadTime = dataReader.GetValue(iTotalLoadTimecol)
                End If

                iProcessingSpeedcol = dataReader.GetOrdinal("ProcessingSpeed")
                If dataReader.IsDBNull(iProcessingSpeedcol) Then
                    sProcessingSpeed = vbNullString
                Else
                    sProcessingSpeed = dataReader.GetString(iProcessingSpeedcol)
                End If

                iCustodianscol = dataReader.GetOrdinal("Custodians")
                If dataReader.IsDBNull(iCustodianscol) Then
                    sCustodians = vbNullString
                Else
                    sCustodians = dataReader.GetString(iCustodianscol)
                End If

                iCustodianCountcol = dataReader.GetOrdinal("CustodianCount")
                If dataReader.IsDBNull(iCustodianCountcol) Then
                    iCustodianCount = vbNullString
                Else
                    iCustodianCount = dataReader.GetInt16(iCustodianCountcol)
                End If

                iSearchTermcol = dataReader.GetOrdinal("SearchTerm")
                If dataReader.IsDBNull(iSearchTermcol) Then
                    sSearchTerm = vbNullString
                Else
                    sSearchTerm = dataReader.GetString(iSearchTermcol)
                End If

                iSearchSizecol = dataReader.GetOrdinal("SearchSize")
                If dataReader.IsDBNull(iSearchSizecol) Then
                    dSearchSize = 0.0
                Else
                    dSearchSize = dataReader.GetValue(iSearchSizecol)
                End If

                iHitCountcol = dataReader.GetOrdinal("HitCount")
                If dataReader.IsDBNull(iHitCountcol) Then
                    dHitCount = 0.0
                Else
                    dHitCount = dataReader.GetValue(iHitCountcol)
                End If

                iCustodianSearchHitcol = dataReader.GetOrdinal("CustodianSearchHit")
                If dataReader.IsDBNull(iCustodianSearchHitcol) Then
                    sCustodianSearchHit = vbNullString
                Else
                    sCustodianSearchHit = dataReader.GetValue(iCustodianSearchHitcol)
                End If

                iCustodianSearchHitcol = dataReader.GetOrdinal("CustodianSearchHit")
                If dataReader.IsDBNull(iCustodianSearchHitcol) Then
                    sCustodianSearchHit = vbNullString
                Else
                    sCustodianSearchHit = dataReader.GetValue(iCustodianSearchHitcol)
                End If

                iCaseDescriptioncol = dataReader.GetOrdinal("CaseDescription")
                If dataReader.IsDBNull(iCaseDescriptioncol) Then
                    sCaseDescription = vbNullString
                Else
                    sCaseDescription = dataReader.GetValue(iCaseDescriptioncol)
                End If

                iOldestTopLevelCol = dataReader.GetOrdinal("OldestItem")
                If dataReader.IsDBNull(iOldestTopLevelCol) Then
                    sOldestTopLevel = vbNullString
                Else
                    sOldestTopLevel = dataReader.GetValue(iOldestTopLevelCol)
                    sOldestTopLevel = Split(sOldestTopLevel, "T").First
                End If

                iNewestTopLevelCol = dataReader.GetOrdinal("NewestItem")
                If dataReader.IsDBNull(iNewestTopLevelCol) Then
                    sNewestTopLevel = vbNullString
                Else
                    sNewestTopLevel = dataReader.GetValue(iNewestTopLevelCol)
                    sNewestTopLevel = Split(sNewestTopLevel, "T").First
                End If

                iLanguagescol = dataReader.GetOrdinal("Languages")
                If dataReader.IsDBNull(iLanguagescol) Then
                    sLanguages = vbNullString
                Else
                    sLanguages = dataReader.GetValue(iLanguagescol)
                End If

                iReportLoadTimecol = dataReader.GetOrdinal("ReportLoadTime")
                If dataReader.IsDBNull(iReportLoadTimecol) Then
                    sReportLoadTime = ""
                Else
                    sReportLoadTime = dataReader.GetValue(iReportLoadTimecol)
                End If

                iNuixLogFileCol = dataReader.GetOrdinal("NuixLogLocation")
                If dataReader.IsDBNull(iNuixLogFileCol) Then
                    sNuixLogFile = ""
                Else
                    sNuixLogFile = dataReader.GetValue(iNuixLogFileCol)
                End If

                iEvidenceDescriptioncol = dataReader.GetOrdinal("EvidenceDescription")
                If dataReader.IsDBNull(iEvidenceDescriptioncol) Then
                    sEvidenceDescription = ""
                Else
                    sEvidenceDescription = dataReader.GetValue(iEvidenceDescriptioncol)
                End If

                iIrregularItemscol = dataReader.GetOrdinal("IrregularItems")
                If dataReader.IsDBNull(iIrregularItemscol) Then
                    sIrregularItems = ""
                Else
                    sIrregularItems = dataReader.GetValue(iIrregularItemscol)
                End If

                If (CInt(dHitCount) > 0) Then
                    dHitCountPercentage = CInt(dHitCount) / CInt(dTotalItemCount)
                Else
                    dHitCountPercentage = CDec("0.0")
                End If
                sHitCountPercentage = FormatPercent(dHitCountPercentage)

                iRowIndex = grdCaseInfo.Rows.Add()

                CaseInfoRow = grdCaseInfo.Rows(iRowIndex)

                With CaseInfoRow
                    CaseInfoRow.Cells("CaseGUID").Value = sCaseGUID
                    If File.Exists(sCaseLocation & "\case.fbi2") Then
                        sCaseModifiedDate = System.IO.File.GetLastWriteTime(sCaseLocation & "\case.fbi2")

                        iModifyDateCompare = DateTime.Compare(DateTime.Parse(sReportLoadTime), DateTime.Parse(sCaseModifiedDate))

                        If (iModifyDateCompare < 0) Then
                            CaseInfoRow.Cells("CollectionStatus").Value = "Get New Data"
                            CaseInfoRow.DefaultCellStyle.ForeColor = Color.Orange
                        Else
                            CaseInfoRow.Cells("CollectionStatus").Value = sCollectionStatus
                        End If
                    Else
                        CaseInfoRow.Cells("CollectionStatus").Value = "Case No Longer Exists"
                    End If

                    lSeconds = CLng(sTotalLoadTime)

                    lHours = Int(lSeconds / 3600)
                    lMinutes = (Int(lSeconds / 60)) - (lHours * 60)
                    lSeconds = Int(lSeconds Mod 60)

                    If lSeconds = 60 Then
                        lMinutes = lMinutes + 1
                        lSeconds = 0
                    End If

                    If lMinutes = 60 Then
                        lMinutes = 0
                        lHours = lHours + 1
                    End If
                    sLoadTimeHMS = lHours.ToString & ":" & lMinutes.ToString & ":" & lSeconds.ToString
                    CaseInfoRow.Cells("CaseGUID").Value = sCaseGUID
                    CaseInfoRow.Cells("PercentComplete").Value = iPercentComplete
                    CaseInfoRow.Cells("CaseName").Value = sCaseName
                    CaseInfoRow.Cells("ReportLoadDuration").Value = sReportLoadDuration
                    asNuixVersionNumber = Split(sCurrentCaseVersion, ".")
                    If (CInt(asNuixVersionNumber(0)) < 7) Then
                        CaseInfoRow.DefaultCellStyle.ForeColor = Color.Red
                    ElseIf (CInt(asNuixVersionNumber(0)) = 7) And (CInt(asNuixVersionNumber(1)) < 3) Then
                        CaseInfoRow.DefaultCellStyle.ForeColor = Color.Orange
                    Else
                        CaseInfoRow.DefaultCellStyle.ForeColor = Color.Green
                    End If
                    CaseInfoRow.Cells("CurrentCaseVersion").Value = sCurrentCaseVersion


                    CaseInfoRow.Cells("UpgradedCaseVersion").Value = sUpgradedCaseVersion
                    CaseInfoRow.Cells("BatchLoadInfo").Value = sBatchLoadInfo
                    CaseInfoRow.Cells("DataExport").Value = sExportItems
                    CaseInfoRow.Cells("CaseLocation").Value = sCaseLocation
                    CaseInfoRow.Cells("BackUpLocation").Value = sBackUpLocation
                    CaseInfoRow.Cells("CaseDescription").Value = sCaseDescription
                    Select Case sShowSizeIn
                        Case "Bytes"
                            CaseInfoRow.Cells("CaseSizeOnDisk").Value = FormatNumber(dCaseSizeOnDisk, 2, , TriState.True)
                            CaseInfoRow.Cells("CaseFileSize").Value = FormatNumber(dCaseFileSize, 2, , TriState.True)
                            CaseInfoRow.Cells("CaseAuditSize").Value = FormatNumber(dCaseAuditSize, 2, , TriState.True)
                            CaseInfoRow.Cells("SearchSize").Value = FormatNumber(dSearchSize, 2, TriState.True)

                        Case "Megabytes"
                            CaseInfoRow.Cells("CaseSizeOnDisk").Value = FormatNumber((dCaseSizeOnDisk / 1024 / 1024), 2, , TriState.True)
                            CaseInfoRow.Cells("CaseFileSize").Value = FormatNumber((dCaseFileSize / 1024 / 1024), 2, , TriState.True)
                            CaseInfoRow.Cells("CaseAuditSize").Value = FormatNumber((dCaseAuditSize / 1024 / 1024), 2, , TriState.True)
                            CaseInfoRow.Cells("SearchSize").Value = FormatNumber((dSearchSize / 1024 / 1024), 2, , TriState.True)

                        Case "Gigabytes"
                            CaseInfoRow.Cells("CaseSizeOnDisk").Value = FormatNumber((dCaseSizeOnDisk / 1024 / 1024 / 1024), 2, , TriState.True)
                            CaseInfoRow.Cells("CaseFileSize").Value = FormatNumber((dCaseFileSize / 1024 / 1024 / 1024), 2, , TriState.True)
                            CaseInfoRow.Cells("CaseAuditSize").Value = FormatNumber((dCaseAuditSize / 1024 / 1024 / 1024), 2, , TriState.True)
                            CaseInfoRow.Cells("SearchSize").Value = FormatNumber((dSearchSize / 1024 / 1024 / 1024), 2, , TriState.True)
                    End Select
                    CaseInfoRow.Cells("OldestTopLevel").Value = sOldestTopLevel
                    CaseInfoRow.Cells("NewestTopLevel").Value = sNewestTopLevel
                    CaseInfoRow.Cells("IsCompound").Value = sIsCompound
                    CaseInfoRow.Cells("CasesContained").Value = sCasesContained
                    CaseInfoRow.Cells("ContainedInCase").Value = sContainedInCase
                    CaseInfoRow.Cells("Investigator").Value = sInvestigator
                    CaseInfoRow.Cells("InvestigatorSessions").Value = sInvestigatorSessions
                    CaseInfoRow.Cells("InvestigatorTimeSummary").Value = sInvestigatorTimeSummary
                    CaseInfoRow.Cells("DataExport").Value = sExportItems
                    CaseInfoRow.Cells("BrokerMemory").Value = iBrokerMemory
                    CaseInfoRow.Cells("WorkerCount").Value = iWorkerCount
                    CaseInfoRow.Cells("WorkerMemory").Value = iWorkerMemory
                    CaseInfoRow.Cells("EvidenceName").Value = sEvidenceProcessed
                    CaseInfoRow.Cells("EvidenceLocation").Value = sEvidenceLocation
                    CaseInfoRow.Cells("EvidenceDescription").Value = sEvidenceDescription
                    CaseInfoRow.Cells("EvidenceCustomMetadata").Value = sEvidenceCustomMetadata
                    CaseInfoRow.Cells("LanguagesContained").Value = sLanguages
                    CaseInfoRow.Cells("MimeTypes").Value = sMimeTypes
                    CaseInfoRow.Cells("ItemTypes").Value = sItemTypes
                    CaseInfoRow.Cells("IrregularItems").Value = sIrregularItems
                    CaseInfoRow.Cells("CreationDate").Value = sCreationDate
                    CaseInfoRow.Cells("ModifiedDate").Value = sModifiedDate
                    CaseInfoRow.Cells("LoadStartDate").Value = sLoadDataStart
                    CaseInfoRow.Cells("LoadEndDate").Value = sLoadDataEnd
                    CaseInfoRow.Cells("LoadTime").Value = iLoadTime
                    CaseInfoRow.Cells("LoadEvents").Value = sLoadEvents
                    CaseInfoRow.Cells("TotalLoadTime").Value = sLoadTimeHMS
                    CaseInfoRow.Cells("ProcessingSpeed").Value = sProcessingSpeed
                    CaseInfoRow.Cells("Custodians").Value = sCustodians
                    CaseInfoRow.Cells("CustodianCount").Value = iCustodianCount
                    CaseInfoRow.Cells("SearchTerm").Value = sSearchTerm

                    CaseInfoRow.Cells("SearchHitCount").Value = FormatNumber(dHitCount, 0, , TriState.True)
                    CaseInfoRow.Cells("CustodianSearchHit").Value = sCustodianSearchHit
                    CaseInfoRow.Cells("TotalCaseItemCount").Value = FormatNumber(dTotalItemCount, 0, , TriState.True)
                    CaseInfoRow.Cells("ItemCounts").Value = sItemCounts
                    CaseInfoRow.Cells("DuplicateItems").Value = sDuplicateItems
                    CaseInfoRow.Cells("OriginalItems").Value = sOriginaltems
                    CaseInfoRow.Cells("HitCountPercent").Value = sHitCountPercentage
                    CaseInfoRow.Cells("NuixLogLocation").Value = sNuixLogFile
                    'CaseInfoRow.Cells("")
                End With
            End While
            sqlConnection.Close()

        Catch ex As Exception
            Logger(psUCRTLogFile, "Error in Populate Case Info Grid - " & ex.ToString)
        End Try
        blnPopulateCaseInfoGrid = True
    End Function

    Public Function blnGetModifiedDate(ByVal sCaseLocation As String, sModifiedDate As String) As Boolean
        blnGetModifiedDate = False


        blnGetModifiedDate = True
    End Function
    Public Sub UpdateNuixAllDataReports()

        Dim sCaseName As String
        Dim sRubyFileName As String
        Dim sScriptsDirectory As String
        Dim sBatchFileName As String
        Dim bStatus As Boolean
        Dim NuixConsoleProcessStartInfo As ProcessStartInfo
        Dim NuixConsoleProcess As Process
        Dim sSearchTerm As String
        Dim sItemCount As String
        Dim sNuixAppMemory As String
        Dim sReportFilePath As String
        Dim sCasePath As String
        Dim sCaseFileSize As String
        Dim sMimeTypes As String
        Dim sItemTypes As String
        Dim sLoadDataStart As String
        Dim sLoadDataEnd As String
        Dim dLoadDataStart As DateTime
        Dim dLoadDataEnd As DateTime
        Dim sCustodians As String
        Dim iCustodianCount As Integer
        Dim sReportType As String
        Dim grdCaseData As DataGridView
        Dim CaseGridRow As DataGridViewRow
        Dim iRowIndex As Integer
        Dim LoadTime As TimeSpan
        Dim dTotalMinutes As Double
        Dim sSQLiteDBLocation As String
        Dim asDateParts() As String
        Dim sLoadDatePart As String
        Dim sLoadTimePart As String
        Dim dLoadDateStart As Date
        Dim dLoadDateEnd As Date
        Dim sUserSearchString As String
        Dim sUserSearchFile As String
        Dim dStartTime As DateTime
        Dim dEndTime As DateTime
        Dim tsReportLoadDuration As TimeSpan
        Dim sReportLoadDuration As String
        Dim sNMSLocation As String
        Dim sNMSUserName As String
        Dim sNMSUserInfo As String
        Dim sLogFileDir As String
        Dim bNoMoreJobs As Boolean
        Dim sCaseLogFileDir As Boolean
        Dim sCaseGUID As String
        Dim sRegistryServer As String
        Dim bMigrateCase As Boolean
        Dim bExportSearchResults As Boolean
        Dim sExportDirectory As String
        Dim iTotalCases As Integer
        Dim bExportOnly As Boolean
        Dim sExportWorkers As String
        Dim sExportWorkerMemory As String
        Dim sExportType As String
        Dim sUpgradeCases As String

        bNoMoreJobs = False
        Try
            sReportFilePath = txtReportLocation.Text
            sScriptsDirectory = sReportFilePath & "\" & "Scripts"
            sSearchTerm = txtSearchTerm.Text
            sNuixAppMemory = "-Xmx4g"
            Invoke(Sub()
                       If chkExportSearchResults.Checked = True Then
                           bExportSearchResults = True
                           sExportDirectory = txtExportLocation.Text
                       End If
                   End Sub)
            Invoke(Sub()
                       sUpgradeCases = cboUpgradeCasees.Text
                       If sUpgradeCases = "No" Then
                           bMigrateCase = False
                       ElseIf sUpgradeCases = "Upgrade Only" Then
                           bMigrateCase = True
                       ElseIf sUpgradeCases = "Upgrade and Report" Then
                           bMigrateCase = True
                       End If
                   End Sub)
            Invoke(Sub()
                       sNMSUserInfo = txtNMSUserName.Text
                   End Sub)
            Invoke(Sub()
                       sNMSUserInfo = txtNMSInfo.Text
                   End Sub)
            Invoke(Sub()
                       sNMSLocation = txtNMSAddress.Text
                   End Sub)
            Invoke(Sub()
                       sLogFileDir = txtNuixLogDir.Text
                   End Sub)

            Invoke(Sub()
                       sReportType = cboReportType.Text
                   End Sub)
            Invoke(Sub()
                       If radSearchFile.Checked = True Then
                           sUserSearchString = ""
                           sUserSearchFile = Me.txtSearchTerm.Text
                       ElseIf radSearchTerm.Checked = True Then
                           sUserSearchString = Me.txtSearchTerm.Text
                           sUserSearchFile = ""
                       End If
                   End Sub)
            Invoke(Sub()
                       grdCaseData = Me.grdCaseInfo
                   End Sub)
            Invoke(Sub()
                       sSQLiteDBLocation = Me.txtReportLocation.Text
                   End Sub)
            Invoke(Sub()
                       sRegistryServer = Me.txtRegistryServer.Text
                   End Sub)
            Invoke(Sub()
                       bExportOnly = chkExportOnly.Checked
                   End Sub)
            Invoke(Sub()
                       sExportWorkers = numExportWorkers.Value
                   End Sub)
            Invoke(Sub()
                       sExportWorkerMemory = numExportWorkerMemory.Value
                   End Sub)
            Invoke(Sub()
                       sExportType = cboExportType.Text
                   End Sub)

            sSQLiteDBLocation = sSQLiteDBLocation & "\Scripts"
            Do While bNoMoreJobs = False
                iRowIndex = 0
                For Each row In grdCaseData.Rows
                    iTotalCases = grdCaseData.RowCount
                    If iRowIndex > iTotalCases Then
                        Invoke(Sub()
                                   CaseGridRow = grdCaseInfo.Rows(iRowIndex)
                               End Sub)
                        If row.cells("CollectionStatus").value = "Get New Data" Then
                            Try
                                sCaseName = CaseGridRow.Cells("CaseName").Value
                                sCaseGUID = CaseGridRow.Cells("CaseGUID").Value
                                If sCaseName <> vbNullString Then
                                    sCaseLogFileDir = sLogFileDir & "\" & sCaseName & "-" & DateTime.Now.ToString("yyyy-MM-dd hh-mm-ss")

                                    dStartTime = DateTime.Now
                                    sCasePath = CaseGridRow.Cells("CaseLocation").Value
                                    CaseGridRow.Cells("CollectionStatus").Value = "Getting Case Data from Case..."
                                    bStatus = blnUpdateSQLiteReportingDB(sSQLiteDBLocation, sCaseGUID, "CollectionStatus", "Getting Case Data from Case...")
                                    'lblCaseFileMessage.Text = "Nuix Case Finder Tool: Getting Case Statistics For: " & sCaseName
                                    sRubyFileName = sScriptsDirectory & "\" & sCaseName & "-" & psReportType & ".rb"
                                    sBatchFileName = sScriptsDirectory & "\" & sCaseName & "-" & psReportType & "_startup.bat"
                                    'bStatus = blnBuildAllCaseDataProcessingRuby(sRubyFileName, sCasePath, sScriptsDirectory, sUserSearchString, bMigrateCase)
                                    bStatus = blnBuildUpdatedAllCaseDataProcessingRuby(sRubyFileName, sCasePath, sScriptsDirectory, sUserSearchString, sUserSearchFile, bMigrateCase, bExportSearchResults, sExportDirectory, bExportOnly, sExportWorkers, sExportWorkerMemory, sExportType, sUpgradeCases)

                                    bStatus = blnBuildBatchFiles(sBatchFileName, sReportType & "-" & sCaseName, txtNuixConsoleLocation.Text, "Server", sNMSLocation, sNMSUserName, sNMSUserInfo, "1", sScriptsDirectory, sRubyFileName, sNuixAppMemory, sCaseLogFileDir, sRegistryServer)
                                    NuixConsoleProcessStartInfo = New ProcessStartInfo(sBatchFileName)
                                    'NuixConsoleProcessStartInfo.WindowStyle = ProcessWindowStyle.Minimized
                                    NuixConsoleProcessStartInfo.WindowStyle = ProcessWindowStyle.Hidden
                                    NuixConsoleProcess = System.Diagnostics.Process.Start(NuixConsoleProcessStartInfo)
                                    Dim bProcessRunning As Boolean
                                    bProcessRunning = True
                                    Do While bProcessRunning = True
                                        bProcessRunning = blnCheckIfProcessIsRunning(NuixConsoleProcess.Id)
                                        Thread.Sleep(2000)
                                    Loop
                                    bStatus = blnGetUpdatedDBInfo(sScriptsDirectory, sCaseGUID, "CaseFileSize", sCaseFileSize, "INT") '
                                    bStatus = blnGetUpdatedDBInfo(sScriptsDirectory, sCaseGUID, "MimeTypes", sMimeTypes, "TEXT")
                                    bStatus = blnGetUpdatedDBInfo(sScriptsDirectory, sCaseGUID, "ItemTypes", sItemTypes, "TEXT")
                                    bStatus = blnGetUpdatedDBInfo(sScriptsDirectory, sCaseGUID, "ItemCounts", sItemCount, "TEXT")
                                    bStatus = blnGetUpdatedDBInfo(sScriptsDirectory, sCaseGUID, "LoadDataStart", sLoadDataStart, "TEXT")
                                    bStatus = blnGetUpdatedDBInfo(sScriptsDirectory, sCaseGUID, "LoadDataEnd", sLoadDataEnd, "TEXT")
                                    bStatus = blnGetUpdatedDBInfo(sScriptsDirectory, sCaseGUID, "Custodians", sCustodians, "TEXT")
                                    bStatus = blnGetUpdatedDBInfo(sScriptsDirectory, sCaseGUID, "CustodianCount", iCustodianCount, "INT")
                                    bStatus = blnGetUpdatedDBInfo(sScriptsDirectory, sCaseGUID, "LoadDataStart", sLoadDataStart, "TEXT")
                                    If sLoadDataStart <> vbNullString Then
                                        asDateParts = Split(sLoadDataStart, "T")
                                        sLoadDatePart = asDateParts(0)

                                        If asDateParts(1).Contains(",") Then
                                            sLoadTimePart = asDateParts(1).Substring(0, asDateParts(1).IndexOf(","))
                                        ElseIf (asDateParts(1).Contains(" + ")) Then
                                            sLoadTimePart = asDateParts(1).Substring(0, asDateParts(1).IndexOf(" + "))
                                        End If

                                        sLoadDataStart = sLoadDatePart & " " & sLoadTimePart
                                        dLoadDateStart = Date.Parse(sLoadDataStart, System.Globalization.DateTimeFormatInfo.InvariantInfo, Globalization.DateTimeStyles.None)
                                        sLoadDataStart = dLoadDateStart.ToString
                                        bStatus = blnUpdateSQLiteReportingDB(sSQLiteDBLocation, sCaseGUID, "LoadDataStart", sLoadDataStart)
                                    End If

                                    bStatus = blnGetUpdatedDBInfo(sScriptsDirectory, sCaseGUID, "LoadDataEnd", sLoadDataEnd, "TEXT")
                                    If sLoadDataEnd <> vbNullString Then
                                        asDateParts = Split(sLoadDataEnd, "T")
                                        sLoadDatePart = asDateParts(0)
                                        If asDateParts(1).Contains(",") Then
                                            sLoadTimePart = asDateParts(1).Substring(0, asDateParts(1).IndexOf(","))
                                        ElseIf (asDateParts(1).Contains(" + ")) Then
                                            sLoadTimePart = asDateParts(1).Substring(0, asDateParts(1).IndexOf(" + "))
                                        End If
                                        sLoadDataEnd = sLoadDatePart & " " & sLoadTimePart
                                        dLoadDateEnd = Date.Parse(sLoadDataEnd, System.Globalization.DateTimeFormatInfo.InvariantInfo, Globalization.DateTimeStyles.None)
                                        sLoadDataEnd = dLoadDateEnd.ToString
                                        bStatus = blnUpdateSQLiteReportingDB(sSQLiteDBLocation, sCaseGUID, "LoadDataEnd", sLoadDataEnd)
                                    End If

                                    CaseGridRow.Cells("TotalCaseItemCount").Value = CDbl(sItemCount).ToString("N0")
                                    CaseGridRow.Cells("CollectionStatus").Value = "File System and Case Data collected"
                                    CaseGridRow.Cells("CaseFileSize").Value = CDbl(sCaseFileSize).ToString("N0")
                                    CaseGridRow.Cells("MimeTypes").Value = sMimeTypes
                                    CaseGridRow.Cells("ItemTypes").Value = sItemTypes
                                    CaseGridRow.Cells("LoadStartDate").Value = sLoadDataStart
                                    CaseGridRow.Cells("LoadEndDate").Value = sLoadDataEnd
                                    If sLoadDataStart <> vbNullString Then
                                        dLoadDataStart = DateTime.Parse(sLoadDataStart)
                                        dLoadDataEnd = DateTime.Parse(sLoadDataEnd)
                                        LoadTime = dLoadDataEnd.Subtract(dLoadDataStart)
                                        dTotalMinutes = LoadTime.TotalMinutes
                                        CaseGridRow.Cells("LoadTime").Value = Math.Round(dTotalMinutes, 2)
                                        bStatus = blnUpdateSQLiteReportingDB(sSQLiteDBLocation, sCaseGUID, "LoadTime", dTotalMinutes)
                                    Else
                                        CaseGridRow.Cells("LoadTime").Value = "0"
                                    End If
                                    CaseGridRow.Cells("Custodians").Value = sCustodians
                                    CaseGridRow.Cells("CustodianCount").Value = iCustodianCount
                                    dEndTime = DateTime.Now
                                    tsReportLoadDuration = (dEndTime - dStartTime)
                                    sReportLoadDuration = tsReportLoadDuration.Minutes & " minutes and " & tsReportLoadDuration.Seconds & " seconds"
                                    CaseGridRow.Cells("ReportLoadDuration").Value = sReportLoadDuration

                                    bStatus = blnUpdateSQLiteReportingDB(sSQLiteDBLocation, sCaseGUID, "ReportLoadDuration", sReportLoadDuration)
                                    bStatus = blnUpdateSQLiteReportingDB(sSQLiteDBLocation, sCaseGUID, "CollectionStatus", "File System and Case Data collected")

                                End If

                            Catch ex As Exception
                                Logger(psUCRTLogFile, "Exception in BuildNuixAllDataReports." & ex.ToString)
                            End Try
                        End If
                        iRowIndex = iRowIndex + 1
                    End If
                Next
                bNoMoreJobs = True
                pbNoMoreJobs = True
            Loop

        Catch ex As Exception
            Logger(psUCRTLogFile, "Error in UpdateNuixAllDataReports - " & ex.ToString)
        End Try
        'btnGetDataThread.Enabled = True
        pDataLoaded = True
        MessageBox.Show("All Reporting Information has completed processing.  Export Data and Run additional report if necessary.", "All Reporting Data Processed", MessageBoxButtons.OK, MessageBoxIcon.None, MessageBoxDefaultButton.Button1, MessageBoxOptions.ServiceNotification, False)

    End Sub

    Public Sub BuildNuixAllDataReports()

        Dim sCaseName As String
        Dim sRubyFileName As String
        Dim sLogFileDir As String
        Dim sScriptsDirectory As String
        Dim sBatchFileName As String
        Dim bStatus As Boolean
        Dim NuixConsoleProcessStartInfo As ProcessStartInfo
        Dim NuixConsoleProcess As Process
        Dim sSearchTerm As String
        Dim sHitCount As String
        Dim sNuixAppMemory As String
        Dim sReportFilePath As String
        Dim sCasePath As String
        Dim dHitCountPercentage As Decimal
        Dim sCaseFileSize As String
        Dim sCaseAuditSize As String
        Dim sSearchSize As String
        Dim sMimeTypes As String
        Dim sItemTypes As String
        Dim sLoadDataStart As String
        Dim sLoadDataEnd As String
        Dim dLoadDataStart As DateTime
        Dim dLoadDataEnd As DateTime
        Dim sCustodians As String
        Dim iCustodianCount As Integer
        Dim sReportType As String
        Dim grdCaseData As DataGridView
        Dim CaseGridRow As DataGridViewRow
        Dim iRowIndex As Integer
        Dim LoadTime As TimeSpan
        Dim dTotalMinutes As Double
        Dim sSQLiteDBLocation As String
        Dim asDateParts() As String
        Dim sLoadDatePart As String
        Dim sLoadTimePart As String
        Dim dLoadDateStart As Date
        Dim dLoadDateEnd As Date
        Dim sUserSearchString As String
        Dim sCustodianSearchHit As String
        Dim dStartTime As DateTime
        Dim dEndTime As DateTime
        Dim tsReportLoadDuration As TimeSpan
        Dim sReportLoadDuration As String
        Dim sBatchLoadInfo As String
        Dim asBatchLoadInfo() As String
        Dim asBatchLoadDetails() As String
        Dim sBatchLoadDate As String
        Dim dBatchLoadDate As String
        Dim sBatchLoadCount As String
        Dim sBatchLoadFileSize As String
        Dim sBatchLoadAuditSize As String
        Dim sAllBatchLoadInfo As String
        Dim sBatchDatePart As String
        Dim sBatchTimePart As String
        Dim sIsCompound As String
        Dim sCasesContained As String
        Dim sNMSLocation As String
        Dim sNMSUserName As String
        Dim sNMSUserInfo As String
        Dim bNoMoreJobs As Boolean
        Dim asCasesContainedAll() As String
        Dim sCaseContainedValue As String
        Dim sCaseContainedInValue As String
        Dim sCaseLogFileDir As String
        Dim sCaseGUID As String
        Dim sLoadEvents As String
        Dim sTotalLoadTime As String
        Dim lHrs As Long
        Dim lMinutes As Long
        Dim lSeconds As Long
        Dim sLoadTimeHMS As String
        Dim sRegistryServer As String
        Dim sInvestigatorSessions As String
        Dim sInvestigatorTimeSummary As String
        Dim bMigrateCase As Boolean
        Dim sUpgradeCases As String
        Dim sUpgradedCaseVersion As String
        Dim sCaseDirectory As String
        Dim sServerType As String
        Dim sNuixLogFile As String
        Dim iNuixAppMemory As Integer
        Dim sInvalidSessions As String
        Dim dProcessingSpeed As Double
        Dim sProcessingCalculation As String
        Dim sReportLoadEnd As String
        Dim bExportSearchResults As Boolean
        Dim sOldestItemDate As String
        Dim sNewestItemDate As String
        Dim sLanguages As String
        Dim sEvidenceCustomMetadata As String
        Dim sCaseDescription As String
        Dim sEvidenceDescription As String
        Dim sIrregularItems As String
        Dim sPercentComplete As String
        Dim sTotalCaseItemCount As String
        Dim sDuplicateItems As String
        Dim sOriginalItems As String
        Dim sItemCounts As String
        Dim sShowSizeIn As String
        Dim dCaseFileSize As Double
        Dim dCaseAuditSize As Double
        Dim dSearchSize As Double
        Dim sUserSearchFile As String
        Dim sExportDirectory As String
        Dim bExportOnly As Boolean
        Dim sExportWorkers As String
        Dim sExportWorkerMemory As String
        Dim sExportItems As String
        Dim sExportType As String
        Dim sErrorDescription As String


        bNoMoreJobs = False
        Try
            Invoke(Sub()
                       If radSearchFile.Checked = True Then
                           sUserSearchFile = txtSearchTerm.Text
                           sUserSearchString = ""
                       ElseIf radSearchTerm.Checked = True Then
                           sUserSearchFile = ""
                           sUserSearchString = txtSearchTerm.Text
                       End If
                   End Sub)
            Invoke(Sub()
                       If chkExportSearchResults.Checked = True Then
                           'sExportType = cboExportType.Text
                           bExportSearchResults = True
                           sExportDirectory = txtExportLocation.Text
                       Else
                           bExportSearchResults = False
                       End If
                   End Sub)
            Invoke(Sub()
                       iNuixAppMemory = numNuixAppMemory.Value
                       sNuixAppMemory = "-Xmx" & iNuixAppMemory & "g"
                   End Sub)
            Invoke(Sub()
                       sShowSizeIn = cboSizeReporting.Text
                   End Sub)
            Invoke(Sub()
                       sUpgradeCases = cboUpgradeCasees.Text
                       If sUpgradeCases = "No" Then
                           bMigrateCase = False
                       ElseIf sUpgradeCases = "Upgrade Only" Then
                           bMigrateCase = True
                       ElseIf sUpgradeCases = "Upgrade and Report Then" Then
                           bMigrateCase = True
                       End If
                   End Sub)
            Invoke(Sub()
                       sReportFilePath = txtReportLocation.Text
                   End Sub)
            sScriptsDirectory = sReportFilePath & "\" & "Scripts"
            Invoke(Sub()
                       sSearchTerm = txtSearchTerm.Text
                   End Sub)
            Invoke(Sub()
                       sServerType = cboLicenseType.Text
                   End Sub)
            Invoke(Sub()
                       sNMSUserName = txtNMSUserName.Text
                   End Sub)
            Invoke(Sub()
                       sNMSUserInfo = txtNMSInfo.Text
                   End Sub)
            Invoke(Sub()
                       sNMSLocation = txtNMSAddress.Text
                   End Sub)
            Invoke(Sub()
                       sRegistryServer = Me.txtRegistryServer.Text
                   End Sub)
            Invoke(Sub()
                       sLogFileDir = txtNuixLogDir.Text
                   End Sub)

            Invoke(Sub()
                       sReportType = cboReportType.Text
                   End Sub)
            Invoke(Sub()
                       grdCaseData = Me.grdCaseInfo
                   End Sub)
            Invoke(Sub()
                       sSQLiteDBLocation = Me.txtReportLocation.Text
                   End Sub)
            Invoke(Sub()
                       'btnGetDataThread = Me.btnGetData
                       'btnGetDataThread = System.Windows.Forms.Button.btnGetdata
                   End Sub)

            Invoke(Sub()
                       sProcessingCalculation = cboCalculateProcessingSpeeds.Text
                   End Sub)

            Invoke(Sub()
                       bExportOnly = chkExportOnly.Checked
                   End Sub)

            Invoke(Sub()
                       sExportWorkers = numExportWorkers.Value.ToString
                   End Sub)

            Invoke(Sub()
                       sExportWorkerMemory = numExportWorkerMemory.Value.ToString
                   End Sub)

            Invoke(Sub()
                       sExportType = cboExportType.Text
                   End Sub)
            sSQLiteDBLocation = sScriptsDirectory

            Do While bNoMoreJobs = False
                iRowIndex = 0
                For Each row In grdCaseData.Rows
                    sAllBatchLoadInfo = ""

                    Invoke(Sub()
                               CaseGridRow = grdCaseInfo.Rows(iRowIndex)
                           End Sub)
                    If (row.cells("CollectionStatus").value = "File System Info Collected - Case Version Mismatch") And (bMigrateCase <> True) Then
                    ElseIf (row.cells("CollectionStatus").value = "Case Locked") Then
                    Else
                        Try
                            sCaseName = CaseGridRow.Cells("CaseName").Value
                            sCaseGUID = CaseGridRow.Cells("CaseGUID").Value
                            sCaseDirectory = CaseGridRow.Cells("CaseLocation").Value
                            If sCaseName <> vbNullString Then
                                sCaseLogFileDir = sLogFileDir & "\" & sCaseName & "-" & DateTime.Now.ToString("yyyy-MM-dd hh-mm-ss")
                                dStartTime = DateTime.Now

                                sCasePath = CaseGridRow.Cells("CaseLocation").Value

                                CaseGridRow.Cells("CollectionStatus").Value = "Getting Case Data from Case..."
                                Logger(psUCRTLogFile, "Updating Collection Status - " & sCaseName & " - Getting Case Data from Case...")
                                bStatus = blnUpdateSQLiteReportingDB(sSQLiteDBLocation, sCaseGUID, "CollectionStatus", "Getting Case Data from Case...")
                                'lblCaseFileMessage.Text = "Nuix Case Finder Tool: Getting Case Statistics For: " & sCaseName
                                sRubyFileName = sScriptsDirectory & "\" & sCaseName & "-" & psReportType & ".rb"
                                sBatchFileName = sScriptsDirectory & "\" & sCaseName & "-" & psReportType & "_startup.bat"
                                Logger(psUCRTLogFile, "Building Case Ruby Scripts - " & sCaseName)
                                'bStatus = blnBuildAllCaseDataProcessingRuby(sRubyFileName, sCasePath, sScriptsDirectory, sUserSearchString, bMigrateCase)
                                bStatus = blnBuildUpdatedAllCaseDataProcessingRuby(sRubyFileName, sCasePath, sScriptsDirectory, sUserSearchString, sUserSearchFile, bMigrateCase, bExportSearchResults, sExportDirectory, bExportOnly, sExportWorkers, sExportWorkerMemory, sExportType, sUpgradeCases)

                                Logger(psUCRTLogFile, "Building Case Batch File - " & sCaseName)
                                bStatus = blnBuildBatchFiles(sBatchFileName, sReportType & "-" & sCaseName, txtNuixConsoleLocation.Text, sServerType, sNMSLocation, sNMSUserName, sNMSUserInfo, "1", sScriptsDirectory, sRubyFileName, sNuixAppMemory, sCaseLogFileDir, sRegistryServer)
                                Logger(psUCRTLogFile, "Launching Nuix for - " & sCaseName)
                                bStatus = blnUpdateSQLiteReportingDB(sSQLiteDBLocation, sCaseGUID, "ReportLoadStart", dStartTime.ToString)
                                NuixConsoleProcessStartInfo = New ProcessStartInfo(sBatchFileName)
                                'NuixConsoleProcessStartInfo.WindowStyle = ProcessWindowStyle.Minimized
                                NuixConsoleProcessStartInfo.WindowStyle = ProcessWindowStyle.Hidden
                                NuixConsoleProcess = System.Diagnostics.Process.Start(NuixConsoleProcessStartInfo)
                                Dim bProcessRunning As Boolean
                                bProcessRunning = True

                                Do While bProcessRunning = True
                                    bProcessRunning = blnCheckIfProcessIsRunning(NuixConsoleProcess.Id)
                                    Thread.Sleep(1000)
                                    bStatus = blnGetUpdatedDBInfo(sScriptsDirectory, sCaseGUID, "PercentComplete", sPercentComplete, "INT")
                                    CaseGridRow.Cells("PercentComplete").Value = sPercentComplete
                                Loop

                                Logger(psUCRTLogFile, "Checking Nuix Logs for errors - " & sCaseName)
                                bStatus = blnCheckNuixLogForErrors(sLogFileDir, sNuixLogFile, sErrorDescription)
                                Logger(psUCRTLogFile, "Checking Nuix Logs for Errors return - " & bStatus.ToString & sErrorDescription)
                                CaseGridRow.Cells("NuixLogLocation").Value = sNuixLogFile
                                If bStatus = False Then
                                    If sUpgradeCases = "Upgrade Only" Then
                                        If bMigrateCase = True Then
                                            bStatus = blnGetCurrentCaseVersion(sCaseDirectory, sUpgradedCaseVersion)
                                            bStatus = blnUpdateSQLiteReportingDB(sSQLiteDBLocation, sCaseGUID, "UpgradedCaseVersion", sUpgradedCaseVersion)
                                            CaseGridRow.Cells("UpgradedCaseVersion").Value = sUpgradedCaseVersion
                                        End If
                                    Else
                                        If bMigrateCase = True Then
                                            bStatus = blnGetCurrentCaseVersion(sCaseDirectory, sUpgradedCaseVersion)
                                            bStatus = blnUpdateSQLiteReportingDB(sSQLiteDBLocation, sCaseGUID, "UpgradedCaseVersion", sUpgradedCaseVersion)
                                            CaseGridRow.Cells("UpgradedCaseVersion").Value = sUpgradedCaseVersion
                                        End If

                                        Try
                                            bStatus = blnUpdateSQLiteReportingDB(sSQLiteDBLocation, sCaseGUID, "NuixLogLocation", sNuixLogFile)
                                            bStatus = blnGetUpdatedDBInfo(sScriptsDirectory, sCaseGUID, "CaseFileSize", sCaseFileSize, "INT") '
                                            bStatus = blnGetUpdatedDBInfo(sScriptsDirectory, sCaseGUID, "CaseAuditSize", sCaseAuditSize, "INT")
                                            bStatus = blnGetUpdatedDBInfo(sScriptsDirectory, sCaseGUID, "OldestItem", sOldestItemDate, "TEXT")
                                            bStatus = blnGetUpdatedDBInfo(sScriptsDirectory, sCaseGUID, "NewestItem", sNewestItemDate, "TEXT")
                                            bStatus = blnGetUpdatedDBInfo(sScriptsDirectory, sCaseGUID, "MimeTypes", sMimeTypes, "TEXT")
                                            bStatus = blnGetUpdatedDBInfo(sScriptsDirectory, sCaseGUID, "ItemTypes", sItemTypes, "TEXT")
                                            bStatus = blnGetUpdatedDBInfo(sScriptsDirectory, sCaseGUID, "LoadDataStart", sLoadDataStart, "TEXT")
                                            bStatus = blnGetUpdatedDBInfo(sScriptsDirectory, sCaseGUID, "LoadDataEnd", sLoadDataEnd, "TEXT")
                                            bStatus = blnGetUpdatedDBInfo(sScriptsDirectory, sCaseGUID, "Custodians", sCustodians, "TEXT")
                                            bStatus = blnGetUpdatedDBInfo(sScriptsDirectory, sCaseGUID, "CustodianCount", iCustodianCount, "INT")
                                            bStatus = blnGetUpdatedDBInfo(sScriptsDirectory, sCaseGUID, "LoadDataStart", sLoadDataStart, "TEXT")
                                            bStatus = blnGetUpdatedDBInfo(sScriptsDirectory, sCaseGUID, "CustodianSearchHit", sCustodianSearchHit, "TEXT")
                                            bStatus = blnGetUpdatedDBInfo(sScriptsDirectory, sCaseGUID, "HitCount", sHitCount, "INT")
                                            bStatus = blnGetUpdatedDBInfo(sScriptsDirectory, sCaseGUID, "CaseFileSize", sCaseFileSize, "INT")
                                            bStatus = blnGetUpdatedDBInfo(sScriptsDirectory, sCaseGUID, "SearchSize", sSearchSize, "INT")
                                            bStatus = blnGetUpdatedDBInfo(sScriptsDirectory, sCaseGUID, "BatchLoadInfo", sBatchLoadInfo, "TEXT")
                                            bStatus = blnGetUpdatedDBInfo(sScriptsDirectory, sCaseGUID, "IsCompound", sIsCompound, "TEXT")
                                            bStatus = blnGetUpdatedDBInfo(sScriptsDirectory, sCaseGUID, "CasesContained", sCasesContained, "TEXT")
                                            bStatus = blnGetUpdatedDBInfo(sScriptsDirectory, sCaseGUID, "LoadEvents", sLoadEvents, "INT")
                                            bStatus = blnGetUpdatedDBInfo(sScriptsDirectory, sCaseGUID, "TotalLoadTime", sTotalLoadTime, "INT")
                                            bStatus = blnGetUpdatedDBInfo(sScriptsDirectory, sCaseGUID, "Languages", sLanguages, "TEXT")
                                            bStatus = blnGetUpdatedDBInfo(sScriptsDirectory, sCaseGUID, "EvidenceCustomMetadata", sEvidenceCustomMetadata, "TEXT")
                                            bStatus = blnGetUpdatedDBInfo(sScriptsDirectory, sCaseGUID, "CaseDescription", sCaseDescription, "TEXT")
                                            bStatus = blnGetUpdatedDBInfo(sScriptsDirectory, sCaseGUID, "EvidenceDescription", sEvidenceDescription, "TEXT")
                                            bStatus = blnGetUpdatedDBInfo(sScriptsDirectory, sCaseGUID, "IrregularItems", sIrregularItems, "TEXT")
                                            bStatus = blnGetUpdatedDBInfo(sScriptsDirectory, sCaseGUID, "DuplicateItems", sDuplicateItems, "TEXT")
                                            bStatus = blnGetUpdatedDBInfo(sScriptsDirectory, sCaseGUID, "OriginalItems", sOriginalItems, "TEXT")
                                            bStatus = blnGetUpdatedDBInfo(sScriptsDirectory, sCaseGUID, "ItemCounts", sItemCounts, "TEXT")
                                            bStatus = blnGetUpdatedDBInfo(sScriptsDirectory, sCaseGUID, "TotalItemCount", sTotalCaseItemCount, "INT")

                                            sInvestigatorSessions = vbNullString
                                            sInvestigatorTimeSummary = vbNullString
                                            sInvalidSessions = vbNullString
                                            sExportItems = vbNullString
                                            bStatus = blnGetInvestigatorSessions(sScriptsDirectory, sCaseGUID, sInvestigatorSessions, sInvalidSessions)
                                            bStatus = blnGetInvestigatorTimeSummary(sScriptsDirectory, sCaseGUID, sInvestigatorTimeSummary)
                                            bStatus = blnGetExportItemsInfo(sScriptsDirectory, sCaseGUID, sExportItems)
                                            bStatus = blnUpdateSQLiteReportingDB(sSQLiteDBLocation, sCaseGUID, "InvestigatorSessions", sInvestigatorSessions)
                                            bStatus = blnUpdateSQLiteReportingDB(sSQLiteDBLocation, sCaseGUID, "InvalidSessions", sInvestigatorSessions)

                                            lSeconds = CLng(sTotalLoadTime)

                                            lHrs = Int(lSeconds / 3600)
                                            lMinutes = (Int(lSeconds / 60)) - (lHrs * 60)
                                            lSeconds = Int(lSeconds Mod 60)

                                            If lSeconds = 60 Then
                                                lMinutes = lMinutes + 1
                                                lSeconds = 0
                                            End If

                                            If lMinutes = 60 Then
                                                lMinutes = 0
                                                lHrs = lHrs + 1
                                            End If

                                            sLoadTimeHMS = lHrs.ToString & ":" & lMinutes.ToString & ":" & lSeconds.ToString

                                            If sBatchLoadInfo <> vbNullString Then
                                                asBatchLoadInfo = Split(sBatchLoadInfo, ";")
                                                If asBatchLoadInfo.Count > 0 Then
                                                    For iCounter = 0 To asBatchLoadInfo.Count - 1
                                                        asBatchLoadDetails = Split(asBatchLoadInfo(iCounter), "::")
                                                        If asBatchLoadDetails(0) <> vbNullString Then
                                                            sBatchLoadDate = asBatchLoadDetails(0)
                                                            sBatchLoadCount = asBatchLoadDetails(1)
                                                            sBatchLoadFileSize = asBatchLoadDetails(2)
                                                            sBatchLoadAuditSize = asBatchLoadDetails(3)
                                                            If sBatchLoadDate <> vbNullString Then
                                                                asDateParts = Split(sBatchLoadDate, "T")
                                                                sBatchDatePart = asDateParts(0)

                                                                If asDateParts(1).Contains(",") Then
                                                                    sBatchTimePart = asDateParts(1).Substring(0, asDateParts(1).IndexOf(","))
                                                                ElseIf (asDateParts(1).Contains(" + ")) Then
                                                                    sBatchTimePart = asDateParts(1).Substring(0, asDateParts(1).IndexOf(" + "))
                                                                Else
                                                                    sBatchTimePart = asDateParts(1)
                                                                End If

                                                                sBatchLoadDate = sBatchDatePart & " " & sBatchTimePart
                                                                dBatchLoadDate = Date.Parse(sBatchLoadDate, System.Globalization.DateTimeFormatInfo.InvariantInfo, Globalization.DateTimeStyles.None)
                                                                sBatchLoadDate = dBatchLoadDate.ToString
                                                                sAllBatchLoadInfo = sAllBatchLoadInfo & sBatchLoadDate & "::" & sBatchLoadCount & "::" & sBatchLoadFileSize & "::" & sBatchLoadAuditSize & ";"
                                                            End If
                                                        End If
                                                    Next
                                                End If
                                                bStatus = blnUpdateSQLiteReportingDB(sSQLiteDBLocation, sCaseGUID, "BatchLoadInfo", sAllBatchLoadInfo)

                                            End If

                                            If sLoadDataStart <> vbNullString Then
                                                asDateParts = Split(sLoadDataStart, "T")
                                                sLoadDatePart = asDateParts(0)

                                                If asDateParts(1).Contains(",") Then
                                                    sLoadTimePart = asDateParts(1).Substring(0, asDateParts(1).IndexOf(","))
                                                ElseIf (asDateParts(1).Contains(" + ")) Then
                                                    sLoadTimePart = asDateParts(1).Substring(0, asDateParts(1).IndexOf(" + "))
                                                End If

                                                sLoadDataStart = sLoadDatePart & " " & sLoadTimePart
                                                dLoadDateStart = Date.Parse(sLoadDataStart, System.Globalization.DateTimeFormatInfo.InvariantInfo, Globalization.DateTimeStyles.None)
                                                sLoadDataStart = dLoadDateStart.ToString
                                                bStatus = blnUpdateSQLiteReportingDB(sSQLiteDBLocation, sCaseGUID, "LoadDataStart", sLoadDataStart)
                                            End If

                                            bStatus = blnGetUpdatedDBInfo(sScriptsDirectory, sCaseGUID, "LoadDataEnd", sLoadDataEnd, "TEXT")
                                            If sLoadDataEnd <> vbNullString Then
                                                asDateParts = Split(sLoadDataEnd, "T")
                                                sLoadDatePart = asDateParts(0)
                                                If asDateParts(1).Contains(",") Then
                                                    sLoadTimePart = asDateParts(1).Substring(0, asDateParts(1).IndexOf(","))
                                                ElseIf (asDateParts(1).Contains(" + ")) Then
                                                    sLoadTimePart = asDateParts(1).Substring(0, asDateParts(1).IndexOf(" + "))
                                                End If
                                                sLoadDataEnd = sLoadDatePart & " " & sLoadTimePart
                                                dLoadDateEnd = Date.Parse(sLoadDataEnd, System.Globalization.DateTimeFormatInfo.InvariantInfo, Globalization.DateTimeStyles.None)
                                                sLoadDataEnd = dLoadDateEnd.ToString
                                                bStatus = blnUpdateSQLiteReportingDB(sSQLiteDBLocation, sCaseGUID, "LoadDataEnd", sLoadDataEnd)
                                            End If

                                            If sOldestItemDate <> vbNullString Then
                                                asDateParts = Split(sOldestItemDate, "T")
                                                sLoadDatePart = asDateParts(0)
                                                If asDateParts(1).Contains(",") Then
                                                    sLoadTimePart = asDateParts(1).Substring(0, asDateParts(1).IndexOf(","))
                                                ElseIf (asDateParts(1).Contains(" + ")) Then
                                                    sLoadTimePart = asDateParts(1).Substring(0, asDateParts(1).IndexOf(" + "))
                                                ElseIf (asDateParts(1).Contains("-")) Then
                                                    sLoadTimePart = asDateParts(1).Substring(0, asDateParts(1).IndexOf("-"))
                                                End If
                                            End If

                                        Catch ex As Exception
                                            Logger(psUCRTLogFile, "Error in Updating Database - " & ex.ToString)
                                        End Try
                                        CaseGridRow.Cells("OldestTopLevel").Value = sLoadDatePart

                                        sLoadDatePart = ""
                                        sLoadTimePart = ""
                                        If sNewestItemDate <> vbNullString Then
                                            asDateParts = Split(sNewestItemDate, "T")
                                            sLoadDatePart = asDateParts(0)
                                            If asDateParts(1).Contains(",") Then
                                                sLoadTimePart = asDateParts(1).Substring(0, asDateParts(1).IndexOf(","))
                                            ElseIf (asDateParts(1).Contains(" + ")) Then
                                                sLoadTimePart = asDateParts(1).Substring(0, asDateParts(1).IndexOf(" + "))
                                            ElseIf (asDateParts(1).Contains("-")) Then
                                                sLoadTimePart = asDateParts(1).Substring(0, asDateParts(1).IndexOf("-"))
                                            End If
                                        End If

                                        CaseGridRow.Cells("NewestTopLevel").Value = sLoadDatePart

                                        CaseGridRow.Cells("BatchLoadInfo").Value = sAllBatchLoadInfo
                                        CaseGridRow.Cells("DataExport").Value = sExportItems
                                        CaseGridRow.Cells("InvestigatorSessions").Value = sInvestigatorSessions
                                        CaseGridRow.Cells("InvalidSessions").Value = sInvalidSessions
                                        CaseGridRow.Cells("InvestigatorTimeSummary").Value = sInvestigatorTimeSummary
                                        CaseGridRow.Cells("TotalCaseItemCount").Value = CDbl(sTotalCaseItemCount).ToString("N0")
                                        CaseGridRow.Cells("DuplicateItems").Value = sDuplicateItems
                                        CaseGridRow.Cells("OriginalItems").Value = sOriginalItems
                                        CaseGridRow.Cells("ItemCounts").Value = sItemCounts
                                        CaseGridRow.Cells("CollectionStatus").Value = "File System and Case Data collected"
                                        dCaseFileSize = CDbl(sCaseFileSize)
                                        dCaseAuditSize = CDbl(sCaseAuditSize)
                                        dSearchSize = CDbl(sSearchSize)

                                        Select Case sShowSizeIn
                                            Case "Bytes"
                                                CaseGridRow.Cells("CaseFileSize").Value = FormatNumber(dCaseFileSize, 2, , TriState.True)
                                                CaseGridRow.Cells("CaseAuditSize").Value = FormatNumber(dCaseAuditSize, 2, , TriState.True)
                                                CaseGridRow.Cells("SearchSize").Value = FormatNumber(dSearchSize, 2, , TriState.True)
                                            Case "Megabytes"
                                                dCaseFileSize = dCaseFileSize / 1024 / 1024
                                                dCaseAuditSize = dCaseAuditSize / 1024 / 1024
                                                dSearchSize = dSearchSize / 1024 / 1024
                                                CaseGridRow.Cells("CaseFileSize").Value = FormatNumber(dCaseFileSize, 2, , TriState.True)
                                                CaseGridRow.Cells("CaseAuditSize").Value = FormatNumber(dCaseAuditSize, 2, , TriState.True)
                                                CaseGridRow.Cells("SearchSize").Value = FormatNumber(dSearchSize, 2, , TriState.True)
                                            Case "Gigabytes"
                                                dCaseFileSize = dCaseFileSize / 1024 / 1024 / 1024
                                                dCaseAuditSize = dCaseAuditSize / 1024 / 1024 / 1024
                                                dSearchSize = dSearchSize / 1024 / 1024 / 1024
                                                CaseGridRow.Cells("CaseFileSize").Value = FormatNumber(dCaseFileSize, 2, , TriState.True)
                                                CaseGridRow.Cells("CaseAuditSize").Value = FormatNumber(dCaseAuditSize, 2, , TriState.True)
                                                CaseGridRow.Cells("SearchSize").Value = FormatNumber(dSearchSize, 2, , TriState.True)

                                        End Select
                                        CaseGridRow.Cells("MimeTypes").Value = sMimeTypes
                                        CaseGridRow.Cells("ItemTypes").Value = sItemTypes
                                        CaseGridRow.Cells("LoadStartDate").Value = sLoadDataStart
                                        CaseGridRow.Cells("LoadEndDate").Value = sLoadDataEnd

                                        If sLoadDataStart <> vbNullString Then
                                            CaseGridRow.Cells("LoadTime").Value = Math.Round(CLng(sTotalLoadTime) / 60, 2)
                                            bStatus = blnUpdateSQLiteReportingDB(sSQLiteDBLocation, sCaseGUID, "LoadTime", Math.Round(CLng(sTotalLoadTime) / 60, 2))
                                        Else
                                            CaseGridRow.Cells("LoadTime").Value = "0"
                                        End If

                                        If (sProcessingCalculation = "File Size") Then
                                            dProcessingSpeed = CDbl(sCaseFileSize) / Math.Round(CLng(sTotalLoadTime) / 60, 2)
                                            dProcessingSpeed = ((dProcessingSpeed / 1024 / 1024 / 1024) * 60)
                                        ElseIf (sProcessingCalculation = "Audit Size") Then
                                            dProcessingSpeed = CDbl(sCaseAuditSize) / Math.Round(CLng(sTotalLoadTime) / 60, 2)
                                            dProcessingSpeed = ((dProcessingSpeed / 1024 / 1024 / 1024) * 60)
                                        Else
                                            dProcessingSpeed = CDbl(sCaseAuditSize) / Math.Round(CLng(sTotalLoadTime) / 60, 2)
                                            dProcessingSpeed = ((dProcessingSpeed / 1024 / 1024 / 1024) * 60)

                                        End If
                                        CaseGridRow.Cells("ProcessingSpeed").Value = Math.Round(dProcessingSpeed, 2).ToString & " GB/hour"
                                        CaseGridRow.Cells("LanguagesContained").Value = sLanguages
                                        CaseGridRow.Cells("EvidenceCustomMetadata").Value = sEvidenceCustomMetadata
                                        CaseGridRow.Cells("EvidenceDescription").Value = sEvidenceDescription
                                        CaseGridRow.Cells("IrregularItems").Value = sIrregularItems
                                        CaseGridRow.Cells("CaseDescription").Value = sCaseDescription
                                        CaseGridRow.Cells("Custodians").Value = sCustodians
                                        CaseGridRow.Cells("CustodianCount").Value = iCustodianCount
                                        CaseGridRow.Cells("CustodianSearchHit").Value = sCustodianSearchHit
                                        CaseGridRow.Cells("SearchHitCount").Value = CInt(sHitCount).ToString("N0")
                                        If (CInt(sHitCount) > 0) Then
                                            dHitCountPercentage = CInt(sHitCount) / CInt(sTotalCaseItemCount)
                                            CaseGridRow.Cells("HitCountPercent").Value = FormatPercent(dHitCountPercentage)
                                        Else
                                            CaseGridRow.Cells("HitCountPercent").Value = "0.0%"
                                        End If
                                        dEndTime = DateTime.Now
                                        tsReportLoadDuration = dEndTime.Subtract(dStartTime)
                                        'sReportLoadDuration = Format(tsReportLoadDuration.Minutes, "#0") & " minutes and " & Format(tsReportLoadDuration.Seconds, "00") & " seconds"
                                        sReportLoadDuration = tsReportLoadDuration.TotalSeconds.ToString("0.0") & " seconds"
                                        CaseGridRow.Cells("ReportLoadDuration").Value = sReportLoadDuration
                                        CaseGridRow.Cells("IsCompound").Value = sIsCompound
                                        CaseGridRow.Cells("CasesContained").Value = sCasesContained
                                        CaseGridRow.Cells("LoadEvents").Value = sLoadEvents
                                        CaseGridRow.Cells("TotalLoadTime").Value = sLoadTimeHMS
                                        bStatus = blnUpdateSQLiteReportingDB(sSQLiteDBLocation, sCaseGUID, "ReportLoadDuration", sReportLoadDuration)
                                        bStatus = blnUpdateSQLiteReportingDB(sSQLiteDBLocation, sCaseGUID, "CollectionStatus", "File System and Case Data collected")
                                        sReportLoadEnd = DateTime.Now.ToString
                                        bStatus = blnUpdateSQLiteReportingDB(sSQLiteDBLocation, sCaseGUID, "ReportLoadEnd", sReportLoadEnd)

                                    End If


                                Else
                                    bStatus = blnUpdateSQLiteReportingDB(sSQLiteDBLocation, sCaseGUID, "NuixLogLocation", sNuixLogFile)
                                    If CaseGridRow.Cells("CaseName").Value <> vbNullString Then
                                        CaseGridRow.Cells("CollectionStatus").Value = "Error Checking Case Data (see logs for details)"
                                        CaseGridRow.DefaultCellStyle.ForeColor = Color.Red
                                        bStatus = blnUpdateSQLiteReportingDB(sSQLiteDBLocation, sCaseGUID, "CollectionStatus", "Error Checking Case Data (see logs for details)")
                                    End If
                                End If
                            Else
                                If CaseGridRow.Cells("CaseName").Value <> vbNullString Then
                                    CaseGridRow.Cells("CollectionStatus").Value = "Error Checking Case Data (see logs for details)"
                                    CaseGridRow.DefaultCellStyle.ForeColor = Color.Red
                                    bStatus = blnUpdateSQLiteReportingDB(sSQLiteDBLocation, sCaseGUID, "CollectionStatus", "Error Checking Case Data (see logs for details)")
                                End If
                            End If

                        Catch ex As Exception
                            CaseGridRow.Cells("CollectionStatus").Value = "Error Checking Case Data (see logs for details)"
                            CaseGridRow.DefaultCellStyle.ForeColor = Color.Red
                            bStatus = blnUpdateSQLiteReportingDB(sSQLiteDBLocation, sCaseGUID, "CollectionStatus", "Error Checking Case Data (see logs for details)")
                            Logger(psUCRTLogFile, " Error in Exception in BuildNuixAllDataReports - " & ex.ToString)
                        End Try
                    End If
                    iRowIndex = iRowIndex + 1
                Next
                bNoMoreJobs = True
                pbNoMoreJobs = True
            Loop

            For Each row In grdCaseInfo.Rows
                If row.cells("CasesContained").value <> vbNullString Then
                    sCasesContained = row.Cells("CasesContained").value
                    asCasesContainedAll = Split(sCasesContained, ";")
                    sCaseContainedInValue = row.Cells("CaseName").value
                    For iCounter = 0 To asCasesContainedAll.Count - 1
                        sCaseContainedValue = asCasesContainedAll(iCounter)
                        For Each caserow In grdCaseInfo.Rows
                            If caserow.cells("CaseName").value <> vbNullString Then
                                If caserow.cells("CaseName").value = sCaseContainedValue Then
                                    caserow.cells("ContainedInCase").Value = sCaseContainedInValue
                                    bStatus = blnUpdateSQLiteReportingDB(sSQLiteDBLocation, caserow.cells("CaseGUID").Value, "ContainedInCase", sCaseContainedInValue)
                                End If
                            End If
                        Next
                    Next
                End If
            Next
            'btnGetDataThread.Enabled = True

        Catch ex As Exception
            Logger(psUCRTLogFile, "Error in BuildNuixAllDataReports - " & ex.ToString)
        End Try
        pDataLoaded = True
        MessageBox.Show("All Reporting Information has completed processing.  Export Data and Run additional report if necessary.", "All Reporting Data Processed", MessageBoxButtons.OK, MessageBoxIcon.None, MessageBoxDefaultButton.Button1, MessageBoxOptions.ServiceNotification, False)

    End Sub

    Public Function blnGetCurrentCaseVersion(ByVal sCurrentCaseDirectory As String, ByRef sUpgradedCaseVersion As String) As Boolean
        blnGetCurrentCaseVersion = False
        Dim NuixCaseFile As Xml.XmlDocument
        Dim ChildNodes As Xml.XmlNodeList
        Dim oMetaDataNodeList As Xml.XmlNodeList

        Dim sNuixName As String

        NuixCaseFile = New Xml.XmlDocument

        NuixCaseFile.Load(sCurrentCaseDirectory & "\case.fbi2")
        oMetaDataNodeList = NuixCaseFile.GetElementsByTagName("metadata")
        For Each MetadataNode In oMetaDataNodeList
            If MetadataNode.haschildnodes Then
                ChildNodes = MetadataNode.childnodes
                For Each Child In ChildNodes
                    If Child.name = "saved-by-product" Then
                        sNuixName = Child.GetAttribute("name")
                        sUpgradedCaseVersion = Child.GetAttribute("version")
                    End If
                Next
            End If
        Next
        blnGetCurrentCaseVersion = True
    End Function
    Public Function blnCheckNuixLogForErrors(ByVal sNuixLogFileDir As String, ByRef sNuixLogFile As String, ByRef sErrorDescription As String) As Boolean
        Dim bStatus As Boolean
        Dim NuixLogStreamReader As StreamReader
        Dim sCurrentRow As String

        blnCheckNuixLogForErrors = False

        bStatus = blnGetNuixLogFiles(sNuixLogFileDir, sNuixLogFile)
        Try
            NuixLogStreamReader = New StreamReader(sNuixLogFile)
            While Not NuixLogStreamReader.EndOfStream
                sCurrentRow = NuixLogStreamReader.ReadLine
                If sCurrentRow.Contains("Error running script:") Then
                    blnCheckNuixLogForErrors = True
                    Exit Function
                ElseIf sCurrentRow.Contains("FATAL com.nuix.investigator.main.b - Couldn't acquire a licence") Then
                    blnCheckNuixLogForErrors = True
                    sErrorDescription = "FATAL com.nuix.investigator.main.b - Couldn't acquire a licence"
                    Exit Function
                ElseIf sCurrentRow.Contains("No licences were found.") Then
                    blnCheckNuixLogForErrors = True
                    sErrorDescription = "No licences were found."
                    Exit Function
                ElseIf sCurrentRow.Contains("Error opening case") Then
                    blnCheckNuixLogForErrors = True
                    sErrorDescription = "Error opening case"
                    Exit Function
                ElseIf sCurrentRow.Contains("Items have an invalid product class") Then
                    blnCheckNuixLogForErrors = True
                    sErrorDescription = "Items have an invalid product class"
                    Exit Function
                ElseIf sCurrentRow.Contains("Couldn't acquire a licence") Then
                    blnCheckNuixLogForErrors = True
                    sErrorDescription = "Couldn't acquire a licence"
                    Exit Function
                End If

            End While

            blnCheckNuixLogForErrors = False

        Catch ex As Exception
            Logger(psUCRTLogFile, "Error in blnCheckNuixLogForErrors - " & ex.ToString)
        End Try
    End Function

    Public Function blnGetNuixLogFiles(ByVal sNuixLogFileDir As String, ByRef sNuixLogFile As String) As Boolean
        Dim CurrentDirectory As DirectoryInfo

        CurrentDirectory = New DirectoryInfo(sNuixLogFileDir)
        If Not CurrentDirectory.Attributes.HasFlag(FileAttributes.ReadOnly) Then
            Try
                For Each Directory In CurrentDirectory.GetDirectories
                    blnGetNuixLogFiles(Directory.FullName, sNuixLogFile)
                Next

                For Each Files In CurrentDirectory.GetFiles
                    If Files.Name = "nuix.log" Then
                        sNuixLogFile = Files.FullName
                    End If
                Next
            Catch ex As Exception
                Logger(psUCRTLogFile, "Error in blnGetNuixLogFiles -" & ex.Message)
            End Try
        End If

        blnGetNuixLogFiles = True

    End Function

    Private Function blnUpdateSQLiteReportingDB(ByVal sSQLiteLocation As String, ByVal sCaseGuid As String, ByVal sColumnName As String, ByVal sColumnValue As String) As Boolean
        blnUpdateSQLiteReportingDB = False

        Dim sUpdateNuixReportingInfo As String
        Dim sParameterValue As String

        Dim SQLiteConnection As SQLiteConnection

        Try
            sParameterValue = "" & "@" & sColumnName & ""
            SQLiteConnection = New SQLiteConnection("Data Source=" & sSQLiteLocation & "\NuixCaseReports.db3;Version=3;New=False;Compress=True;")
            SQLiteConnection.Open()
            Using SQLiteConnection

                sUpdateNuixReportingInfo = "Update NuixReportingInfo set " & sColumnName & " = " & sParameterValue
                sUpdateNuixReportingInfo = sUpdateNuixReportingInfo & " WHERE CaseGUID = @CaseGuid"

                Using oUpdateEWSExtractionDataCommand As New SQLiteCommand()
                    With oUpdateEWSExtractionDataCommand
                        .Connection = SQLiteConnection
                        .CommandText = sUpdateNuixReportingInfo
                        .Parameters.AddWithValue("@CaseGUID", sCaseGuid)
                        .Parameters.AddWithValue(sParameterValue, sColumnValue)
                    End With
                    Try
                        oUpdateEWSExtractionDataCommand.ExecuteNonQuery()
                        SQLiteConnection.Close()
                    Catch ex As Exception
                        Logger(psUCRTLogFile, "Error blnUpdateSQLiteReportingDB - " & ex.Message.ToString())
                    End Try
                End Using
                SQLiteConnection.Close()
                SQLiteConnection.Dispose()
            End Using

        Catch ex As Exception
            Logger(psUCRTLogFile, "Error in blnUpdateSQLiteReportingDB - " & ex.ToString)
        End Try

        blnUpdateSQLiteReportingDB = True
    End Function

    Private Function blnUpdateSessionDuration(ByVal sSQLiteLocation As String, ByVal sCaseGuid As String, ByVal sStartDate As String, ByVal sSessionEvent As String, ByVal sDuration As String) As Boolean
        blnUpdateSessionDuration = False

        Dim sUpdateNuixReportingInfo As String

        Dim SQLiteConnection As SQLiteConnection

        Try

            SQLiteConnection = New SQLiteConnection("Data Source=" & sSQLiteLocation & "\NuixCaseReports.db3;Version=3;New=False;Compress=True;")
            SQLiteConnection.Open()
            Using SQLiteConnection

                sUpdateNuixReportingInfo = "Update UCRTSessionEvents set Duration = '" & sDuration & "'"
                sUpdateNuixReportingInfo = sUpdateNuixReportingInfo & " WHERE CaseGUID = @CaseGuid and StartDate = @StartDate and SessionEvent = @SessionEvent"

                Using oUpdateSessionDurationCommand As New SQLiteCommand()
                    With oUpdateSessionDurationCommand
                        .Connection = SQLiteConnection
                        .CommandText = sUpdateNuixReportingInfo
                        .Parameters.AddWithValue("@CaseGUID", sCaseGuid)
                        .Parameters.AddWithValue("@StartDate", sStartDate)
                        .Parameters.AddWithValue("@SessionEvent", sSessionEvent)
                    End With
                    Try
                        oUpdateSessionDurationCommand.ExecuteNonQuery()
                        SQLiteConnection.Close()
                    Catch ex As Exception
                        Logger(psUCRTLogFile, "Error blnUpdateSessionDuration - " & ex.Message.ToString())
                    End Try
                End Using
                SQLiteConnection.Close()
                SQLiteConnection.Dispose()
            End Using

        Catch ex As Exception
            Logger(psUCRTLogFile, "Error in blnUpdateSessionDuration - " & ex.ToString)
        End Try

        blnUpdateSessionDuration = True
    End Function

    Private Function blnCheckIfProcessIsRunning(ByVal sProcessID As String) As Boolean

        Dim NuixProcess As System.Diagnostics.Process

        blnCheckIfProcessIsRunning = False

        Try
            NuixProcess = Process.GetProcessById(CInt(sProcessID))
            blnCheckIfProcessIsRunning = True
        Catch ex As Exception
            blnCheckIfProcessIsRunning = False
        End Try

    End Function

    Public Function blnGetUpdatedDBInfo(ByRef sSQLiteLocation As String, ByRef sCaseGUID As String, ByRef sFieldName As String, ByRef sFieldValue As String, ByVal sFieldType As String) As Boolean
        blnGetUpdatedDBInfo = False

        Dim Connection As SQLiteConnection
        Dim SQLCommand As SQLiteCommand
        Dim sCustodianQuery As String
        Dim dataReader As SQLiteDataReader
        Dim dFieldValue As Double

        Try
            Connection = New SQLiteConnection("Data Source=" & sSQLiteLocation & "\NuixCaseReports.db3;Version=3;New=False;Compress=True;Read Only=True")
            Connection.Open()

            '            Logger(psUCRTLogFile, "Case GUID = " & sCaseGUID & " Field Name = " & sFieldName & " Field Type = " & sFieldType)

            sCustodianQuery = "SELECT " & sFieldName & " FROM NuixReportingInfo WHERE CaseGUID = '" & sCaseGUID & "'"

            SQLCommand = New SQLiteCommand(sCustodianQuery, Connection)
            dataReader = SQLCommand.ExecuteReader
            If dataReader.HasRows Then
                While dataReader.Read
                    If sFieldType = "TEXT" Then
                        If Not IsDBNull(dataReader(sFieldName)) Then
                            sFieldValue = dataReader.GetValue(0)
                        Else
                            sFieldValue = vbNullString
                        End If
                    ElseIf sFieldType = "INT" Then

                        If Not IsDBNull(dataReader(sFieldName)) Then
                            dFieldValue = dataReader.GetInt64(0)
                            sFieldValue = dFieldValue.ToString
                        Else
                            dFieldValue = 0.0
                            sFieldValue = dFieldValue.ToString
                        End If

                    End If
                End While
            End If
            Connection.Close()

        Catch ex As Exception
            Logger(psUCRTLogFile, "Error in blnGetUpdatedDBInfo" & ex.ToString)
            Connection.Close()
        End Try

        blnGetUpdatedDBInfo = True
    End Function


    Public Function blnGetInvestigatorSessions(ByRef sSQLiteLocation As String, ByRef sCaseGUID As String, ByRef sInvestigatorSession As String, ByRef sInvalidSessions As String) As Boolean
        blnGetInvestigatorSessions = False

        Dim Connection As SQLiteConnection
        Dim SQLCommand As SQLiteCommand
        Dim sCustodianQuery As String
        Dim dataReader As SQLiteDataReader
        Dim sStartDate As String
        Dim asStartDate() As String
        Dim asEndDate() As String
        Dim sEndDate As String
        Dim sInvestigatorName As String
        Dim provider As New CultureInfo("en-US")
        Dim sSessionEvent As String
        Dim bLastEventOpen As Boolean
        Dim dStartDate As DateTime
        Dim dEndDate As DateTime
        Dim iDuration As Integer
        Dim sInvalidSessionDates As String
        Dim iSessionEventMismatch As Integer

        Connection = New SQLiteConnection("Data Source=" & sSQLiteLocation & "\NuixCaseReports.db3;Version=3;New=False;Compress=True;Read Only=True")
        Connection.Open()

        sCustodianQuery = "SELECT SessionEvent, StartDate, EndDate, user, Duration FROM UCRTSessionEvents WHERE CaseGUID = '" & sCaseGUID & "' and SessionEvent in ('SessionOpen', 'SessionClose', 'openSession', 'closeSession') order by StartDate ASC"

        SQLCommand = New SQLiteCommand(sCustodianQuery, Connection)
        dataReader = SQLCommand.ExecuteReader
        If dataReader.HasRows Then
            iSessionEventMismatch = 0
            While dataReader.Read
                sSessionEvent = dataReader(0)
                sInvestigatorName = dataReader(3)

                If ((sSessionEvent = "SessionOpen") Or (sSessionEvent = "openSession")) Then
                    If (bLastEventOpen = True) Then
                        sStartDate = dataReader.GetValue(1)
                        asStartDate = Split(sStartDate, " + ")

                        sStartDate = asStartDate(0)

                        asStartDate = Split(sStartDate, "T")

                        dStartDate = Date.Parse(asStartDate(0) & " " & asStartDate(1))
                        sInvalidSessionDates = sInvalidSessionDates & dStartDate.ToString & ";"

                        iSessionEventMismatch = iSessionEventMismatch + 1
                    Else
                        sStartDate = dataReader.GetValue(1)
                        asStartDate = Split(sStartDate, " + ")

                        sStartDate = asStartDate(0)

                        asStartDate = Split(sStartDate, "T")

                        dStartDate = Date.Parse(asStartDate(0) & " " & asStartDate(1))

                        bLastEventOpen = True
                    End If
                ElseIf ((sSessionEvent = "SessionClose") Or (sSessionEvent = "closeSession")) Then
                    sEndDate = dataReader.GetValue(1)
                    asEndDate = Split(sEndDate, " + ")
                    sEndDate = asEndDate(0)
                    asEndDate = Split(sEndDate, "T")
                    dEndDate = Date.Parse(asEndDate(0) & " " & asEndDate(1))

                    iDuration = DateDiff(DateInterval.Minute, dStartDate, dEndDate)
                    bLastEventOpen = False
                    sInvestigatorSession = sInvestigatorSession & sInvestigatorName & "--Open:" & dStartDate.ToString & "--Close:" & dEndDate.ToString & "(" & iDuration & ");"
                    '                    bStatus = blnUpdateSessionDuration(sSQLiteLocation, sCaseGUID, sStartDate, "SessionOpen", iDuration.ToString)
                End If
            End While
        End If

        If iSessionEventMismatch > 0 Then
            sInvalidSessions = "(" & iSessionEventMismatch & ")-" & sInvalidSessionDates

        Else
            sInvalidSessions = vbNullString
        End If
        Connection.Close()
        blnGetInvestigatorSessions = True
    End Function

    Public Function blnGetExportItemsInfo(ByRef sSQLiteLocation As String, ByRef sCaseGUID As String, ByRef sExportItems As String) As Boolean
        blnGetExportItemsInfo = False

        Dim Connection As SQLiteConnection
        Dim SQLCommand As SQLiteCommand
        Dim sCustodianQuery As String
        Dim dataReader As SQLiteDataReader
        Dim sStartDate As String
        Dim asStartDate() As String
        Dim asEndDate() As String
        Dim sEndDate As String
        Dim sInvestigatorName As String
        Dim provider As New CultureInfo("en-US")
        Dim sSessionEvent As String
        Dim bLastEventOpen As Boolean
        Dim sSuccess As String
        Dim sFailures As String
        Dim dStartDate As DateTime
        Dim dEndDate As DateTime
        Dim iDuration As Integer
        Dim sInvalidSessionDates As String
        Dim iSessionEventMismatch As Integer

        Connection = New SQLiteConnection("Data Source=" & sSQLiteLocation & "\NuixCaseReports.db3;Version=3;New=False;Compress=True;Read Only=True")
        Connection.Open()

        sCustodianQuery = "SELECT SessionEvent, StartDate, EndDate, user, Success, Failures FROM UCRTSessionEvents WHERE CaseGUID = '" & sCaseGUID & "' and SessionEvent in ('export') order by StartDate ASC"

        SQLCommand = New SQLiteCommand(sCustodianQuery, Connection)
        dataReader = SQLCommand.ExecuteReader
        If dataReader.HasRows Then
            iSessionEventMismatch = 0
            While dataReader.Read
                sSessionEvent = dataReader(0)
                sInvestigatorName = dataReader(3)
                sSuccess = dataReader(4)
                sFailures = dataReader(5)

                sStartDate = dataReader.GetValue(1)
                asStartDate = Split(sStartDate, " + ")

                sStartDate = asStartDate(0)

                asStartDate = Split(sStartDate, "T")

                dStartDate = Date.Parse(asStartDate(0) & " " & asStartDate(1))
                sEndDate = dataReader.GetValue(2)
                asEndDate = Split(sEndDate, " + ")

                sEndDate = asEndDate(0)

                asEndDate = Split(sEndDate, "T")

                dEndDate = Date.Parse(asEndDate(0) & " " & asEndDate(1))
                iDuration = DateDiff(DateInterval.Second, dStartDate, dEndDate)
                iSessionEventMismatch = iSessionEventMismatch + 1
                sExportItems = sExportItems & sInvestigatorName & "--Start:" & dStartDate.ToString & "--End:" & dEndDate.ToString & "--Success:" & sSuccess & "--Failures:" & sFailures & "--duration:" & iDuration & ";"
            End While
        End If

        Connection.Close()
        blnGetExportItemsInfo = True
    End Function

    Public Function blnGetInvestigatorTimeSummary(ByRef sSQLiteLocation As String, ByRef sCaseGUID As String, ByRef sInvestigatorTimeSummary As String) As Boolean
        blnGetInvestigatorTimeSummary = False

        Dim Connection As SQLiteConnection
        Dim SQLCommand As SQLiteCommand
        Dim dataReader As SQLiteDataReader
        Dim sInvestigatorQuery As String
        Dim sDurationQuery As String
        Dim lstInvestigators As List(Of String)
        Dim iDuration As Integer
        Dim sStartDate As String
        Dim asStartDate() As String
        Dim sEndDate As String
        Dim asEndDate() As String
        Dim dStartDate As DateTime
        Dim dEndDate As DateTime
        Dim sSessionEvent As String
        Dim bLastEventOpen As Boolean
        Dim iSessionEventMismatch As Integer
        Dim iUserDuration As Integer

        Connection = New SQLiteConnection("Data Source=" & sSQLiteLocation & "\NuixCaseReports.db3;Version=3;New=False;Compress=True;Read Only=True")
        Connection.Open()

        lstInvestigators = New List(Of String)

        sInvestigatorQuery = "Select distinct User from UCRTSessionEvents Where CaseGUID = '" & sCaseGUID & "'"
        SQLCommand = New SQLiteCommand(sInvestigatorQuery, Connection)
        dataReader = SQLCommand.ExecuteReader
        If dataReader.HasRows Then
            While dataReader.Read
                lstInvestigators.Add(dataReader.GetValue(0))
            End While
        End If

        For Each User In lstInvestigators
            iUserDuration = 0
            sDurationQuery = "SELECT SessionEvent, StartDate FROM UCRTSessionEvents WHERE CaseGUID = '" & sCaseGUID & "' and SessionEvent in ('SessionOpen', 'SessionClose', 'openSession', 'closeSession') and User = '" & User.ToString & "' order by StartDate ASC"
            SQLCommand = New SQLiteCommand(sDurationQuery, Connection)
            dataReader = SQLCommand.ExecuteReader
            If dataReader.HasRows Then
                While dataReader.Read
                    sSessionEvent = dataReader.GetValue(0)
                    If ((sSessionEvent = "SessionOpen") Or (sSessionEvent = "openSession")) Then
                        If (bLastEventOpen = True) Then
                            iSessionEventMismatch = iSessionEventMismatch + 1
                        Else
                            sStartDate = dataReader.GetValue(1)
                            asStartDate = Split(sStartDate, " + ")

                            sStartDate = asStartDate(0)

                            asStartDate = Split(sStartDate, "T")

                            dStartDate = Date.Parse(asStartDate(0) & " " & asStartDate(1))
                            bLastEventOpen = True
                        End If
                    ElseIf ((sSessionEvent = "SessionClose") Or (sSessionEvent = "closeSession")) Then
                        sEndDate = dataReader.GetValue(1)
                        asEndDate = Split(sEndDate, " + ")
                        sEndDate = asEndDate(0)
                        asEndDate = Split(sEndDate, "T")
                        dEndDate = Date.Parse(asEndDate(0) & " " & asEndDate(1))

                        iDuration = DateDiff(DateInterval.Minute, dStartDate, dEndDate)
                        iUserDuration = iUserDuration + iDuration
                        bLastEventOpen = False
                    End If
                End While
            End If
            sInvestigatorTimeSummary = sInvestigatorTimeSummary & User.ToString & "(" & iUserDuration & ");"
        Next
        Connection.Close()
        blnGetInvestigatorTimeSummary = True
    End Function
    Private Function blnUpdateSQLiteExtrationDB(ByVal sSQLiteLocation As String, ByVal sCaseName As String, ByVal sColumnName As String, ByVal sColumnValue As String) As Boolean
        blnUpdateSQLiteExtrationDB = False
        Dim sUpdateEWSExtractionDate As String
        Dim sParameterValue As String

        Dim SQLiteConnection As SQLiteConnection

        sParameterValue = "" & "@" & sColumnName & ""
        SQLiteConnection = New SQLiteConnection("Data Source=" & sSQLiteLocation & "\NuixCaseReports.db3;Version=3;New=False;Compress=True;Read Only=False")
        SQLiteConnection.Open()

        Using SQLiteConnection


            sUpdateEWSExtractionDate = "Update NuixReportingInfo set " & sColumnName & " = " & sParameterValue
            sUpdateEWSExtractionDate = sUpdateEWSExtractionDate & " WHERE CaseName = @CaseName"

            Using oUpdateEWSExtractionDataCommand As New SQLiteCommand()
                With oUpdateEWSExtractionDataCommand
                    .Connection = SQLiteConnection
                    .CommandText = sUpdateEWSExtractionDate
                    .Parameters.AddWithValue("@CaseName", sCaseName)
                    .Parameters.AddWithValue(sParameterValue, sColumnValue)
                End With
                Try
                    oUpdateEWSExtractionDataCommand.ExecuteNonQuery()
                    SQLiteConnection.Close()
                Catch ex As Exception
                    Logger(psUCRTLogFile, "Error in blnUpdateSQLLiteExtractionDB - " & ex.ToString)
                End Try
            End Using
            SQLiteConnection.Close()
            SQLiteConnection.Dispose()
        End Using

        blnUpdateSQLiteExtrationDB = True
    End Function

    Private Function blnUpdateSQLiteAllCaseInfo(ByVal sSQLiteLocation As String, ByVal sCaseGUID As String, ByVal sCollectionStatus As String, ByVal iPercentComplete As Integer, ByVal sCaseName As String, ByVal sBatchLoadInfo As String, ByVal sCaseSize As String, ByVal sCaseLocation As String, ByVal sCurrentCaseVersion As String, ByVal sUpgradedCaseVersion As String, ByVal sIsCompound As String, ByVal sCasesContained As String, ByVal sContainedInCase As String, ByVal sInvestigator As String, ByVal sBrokerMemory As String, ByVal sWorkerCount As String, ByVal sWorkerMemory As String, ByVal sEvidenceName As String, ByVal sEvidenceLocation As String, ByVal sEvidenceCustomMetadata As String, ByVal sMimeTypes As String, ByVal sItemTypes As String, ByVal sCreationDate As String, ByVal sModifiedDate As String, ByVal sLoadDataStart As String, ByVal sLoadDataEnd As String, ByVal sLoadTime As String, ByVal sProcessingSpeed As String, ByVal sCustodians As String, ByVal sCustodianCount As Integer, ByVal sHitCount As String, ByVal sSearchTerm As String, ByVal sSearchSize As String, ByVal sCustodianSearchHit As String, ByVal sItemCount As String, ByVal sCaseUsers As String, ByVal dReportLoadTime As DateTime, ByVal sEvidenceDescription As String) As Boolean
        blnUpdateSQLiteAllCaseInfo = False
        Dim sInsertEWSExtractionDate As String
        Dim sUpdateEWSExtractionDate As String
        Dim sQueryReturnedCustodian As String
        Dim SQLiteConnection As SQLiteConnection

        SQLiteConnection = New SQLiteConnection("Data Source=" & sSQLiteLocation & "\NuixCaseReports.db3;Version=3;New=False;Compress=True;Read Only=True")
        SQLiteConnection.Open()

        Using SQLiteConnection
            Using SQLSelectCommand As New SQLiteCommand("SELECT CaseName FROM NuixReportingInfo WHERE CaseGUID='" & sCaseGUID & "'")
                With SQLSelectCommand
                    .Connection = SQLiteConnection
                    Using readerObject As SQLiteDataReader = SQLSelectCommand.ExecuteReader
                        While readerObject.Read
                            sQueryReturnedCustodian = readerObject("CaseName").ToString
                        End While
                    End Using
                End With
            End Using
            SQLiteConnection.Close()
            SQLiteConnection.Dispose()
        End Using

        If sQueryReturnedCustodian = vbNullString Then

            SQLiteConnection = New SQLiteConnection("Data Source=" & sSQLiteLocation & "\NuixCaseReports.db3;Version=3;New=False;Compress=True;Read Only=False")
            SQLiteConnection.Open()
            Using SQLiteConnection

                sInsertEWSExtractionDate = "Insert into NuixReportingInfo (CaseGUID,CaseName,CollectionStatus,PercentComplete,CaseLocation,BatchLoadInfo,CurrentCaseVersion,UpgradedCaseVersion,CaseSizeOnDisk,CaseFileSize,IsCompound,CasesContained,ContainedInCase,Investigator,BrokerMemory,WorkerCount,WorkerMemory,EvidenceProcessed,EvidenceLocation,EvidenceCustomMetadata,MimeTypes,ItemTypes,CreationDate,ModifiedDate,LoadDataStart,LoadDataEnd,LoadTime,ProcessingSpeed,Custodians,CustodianCount,SearchTerm,SearchSize,HitCount,CustodianSearchHit,ItemCounts,CaseUsers,ReportLoadTime,EvidenceDescription) Values "
                sInsertEWSExtractionDate = sInsertEWSExtractionDate & "(@CaseGUID,@CaseName,@CollectionStatus,@PercentComplete,@CaseLocation,@BatchLoadInfo,@CurrentCaseVersion,@UpgradedCaseVersion,@CaseSizeOnDisk,@CaseFileSize,@IsCompound,@CasesContained,@ContainedInCase,@Investigator,@BrokerMemory,@WorkerCount,@WorkerMemory,@EvidenceProcessed,@EvidenceLocation,@EvidenceCustomMetadata,@MimeTypes,@ItemTypes,@CreationDate,@ModifiedDate,@LoadDataStart,@LoadDataEnd,@LoadTime,@ProcessingSpeed,@Custodians,@CustodianCount,@SearchTerm,@SearchSize,@HitCount,@CustodianSearchHit,@ItemCounts,@CaseUsers,@ReportLoadTime,@EvidenceDescription)"
                Using oInsertEWSExtractionDataCommand As New SQLiteCommand()
                    With oInsertEWSExtractionDataCommand
                        .Connection = SQLiteConnection
                        .CommandText = sInsertEWSExtractionDate
                        .Parameters.AddWithValue("@CaseGUID", sCaseGUID)
                        .Parameters.AddWithValue("@CaseName", sCaseName)
                        .Parameters.AddWithValue("@CollectionStatus", sCollectionStatus)
                        .Parameters.AddWithValue("@PercentComplete", iPercentComplete)
                        .Parameters.AddWithValue("@BatchLoadInfo", sBatchLoadInfo)
                        .Parameters.AddWithValue("@CaseLocation", sCaseLocation)
                        .Parameters.AddWithValue("@CurrentCaseVersion", sCurrentCaseVersion)
                        .Parameters.AddWithValue("@UpgradedCaseVersion", sUpgradedCaseVersion)
                        .Parameters.AddWithValue("@CaseSizeOnDisk", sCaseSize)
                        .Parameters.AddWithValue("@CaseFileSize", "0")
                        .Parameters.AddWithValue("@IsCompound", sIsCompound)
                        .Parameters.AddWithValue("@CasesContained", sCasesContained)
                        .Parameters.AddWithValue("@ContainedInCase", sContainedInCase)
                        .Parameters.AddWithValue("@Investigator", sInvestigator)
                        .Parameters.AddWithValue("@BrokerMemory", sBrokerMemory)
                        .Parameters.AddWithValue("@WorkerCount", sWorkerCount)
                        .Parameters.AddWithValue("@WorkerMemory", sWorkerMemory)
                        .Parameters.AddWithValue("@EvidenceProcessed", sEvidenceName)
                        .Parameters.AddWithValue("@EvidenceLocation", sEvidenceLocation)
                        .Parameters.AddWithValue("@EvidenceCustomMetadata", sEvidenceCustomMetadata)
                        .Parameters.AddWithValue("@MimeTypes", sMimeTypes)
                        .Parameters.AddWithValue("@ItemTypes", sItemTypes)
                        .Parameters.AddWithValue("@CreationDate", sCreationDate)
                        .Parameters.AddWithValue("@ModifiedDate", sModifiedDate)
                        .Parameters.AddWithValue("@LoadDataStart", sLoadDataStart)
                        .Parameters.AddWithValue("@LoadDataEnd", sLoadDataEnd)
                        .Parameters.AddWithValue("@LoadTime", sLoadTime)
                        .Parameters.AddWithValue("@ProcessingSpeed", sProcessingSpeed)
                        .Parameters.AddWithValue("@Custodians", sCustodians)
                        .Parameters.AddWithValue("@CustodianCount", sCustodianCount)
                        .Parameters.AddWithValue("@SearchTerm", sSearchTerm)
                        .Parameters.AddWithValue("@SearchSize", sSearchSize)
                        .Parameters.AddWithValue("@HitCount", sHitCount)
                        .Parameters.AddWithValue("@CustodianSearchHit", sCustodianSearchHit)
                        .Parameters.AddWithValue("@ItemCounts", sItemCount)
                        .Parameters.AddWithValue("@CaseUsers", sCaseUsers)
                        .Parameters.AddWithValue("@ReportLoadTime", dReportLoadTime.ToString)
                        .Parameters.AddWithValue("@EvidenceDescription", sEvidenceDescription)
                    End With
                    Try
                        'SQLiteConnection.Open()
                        oInsertEWSExtractionDataCommand.ExecuteNonQuery()
                        SQLiteConnection.Close()
                    Catch ex As Exception
                        Logger(psUCRTLogFile, "Error in blnUpdateSQLiteAllCaseInfo - " & ex.Message.ToString())
                    End Try
                End Using
                SQLiteConnection.Close()
                SQLiteConnection.Dispose()
            End Using
        Else
            SQLiteConnection = New SQLiteConnection("Data Source=" & sSQLiteLocation & "\NuixCaseReports.db3;Version=3;New=False;Compress=True;Read Only=False")
            SQLiteConnection.Open()
            Using SQLiteConnection
                sUpdateEWSExtractionDate = "Update NuixReportingInfo set CaseName = @CaseName, CollectionStatus = @CollectionStatus, PercentComplete = @PercentComplete, BatchLoadInfo = @BatchLoadInfo, CaseLocation = @CaseLocation, CurrentCaseVersion = @CurrentCaseVersion,UpgradedCaseVersion = @UpgradedCaseVersion, CaseSizeOnDisk = @CaseSizeOnDisk, CaseFileSize = @CaseFileSize, IsCompound = @IsCompound, CasesContained = @CasesContained, ContainedInCase = @ContainedInCase, "
                sUpdateEWSExtractionDate = sUpdateEWSExtractionDate & "Investigator = @Investigator, BrokerMemory = @BrokerMemory, WorkerCount = @WorkerCount, WorkerMemory = @WorkerMemory, EvidenceProcessed = @EvidenceProcessed, EvidenceLocation = @EvidenceLocation, EvidenceCustomMetadata = @EvidenceCustomMetadata, "
                sUpdateEWSExtractionDate = sUpdateEWSExtractionDate & "MimeTypes = @MimeTypes, ItemTypes = @ItemTypes ,CreationDate = @CreationDate, ModifiedDate = @ModifiedDate, LoadDataStart = @LoadDataStart, LoadDataEnd = @LoadDataEnd, LoadTime = @LoadTime, "
                sUpdateEWSExtractionDate = sUpdateEWSExtractionDate & "ProcessingSpeed = @ProcessingSpeed, Custodians = @Custodians, CustodianCount = @CustodianCount,SearchTerm = @SearchTerm, SearchSize = @SearchSize, CustodianSearchHit = @CustodianSearchHit, HitCount = @HitCount, ItemCounts = @ItemCounts, CaseUsers = @CaseUsers, ReportLoadTime = @ReportLoadTime, EvidenceDescription = @EvidenceDescription "
                sUpdateEWSExtractionDate = sUpdateEWSExtractionDate & "WHERE CaseGUID = @CaseGUID"

                Using oUpdateEWSExtractionDataCommand As New SQLiteCommand()
                    With oUpdateEWSExtractionDataCommand
                        .Connection = SQLiteConnection
                        .CommandText = sUpdateEWSExtractionDate
                        .Parameters.AddWithValue("@CaseGUID", sCaseGUID)
                        .Parameters.AddWithValue("@CaseName", sCaseName)
                        .Parameters.AddWithValue("@CollectionStatus", sCollectionStatus)
                        .Parameters.AddWithValue("@PercentComplete", iPercentComplete)
                        .Parameters.AddWithValue("@BatchLoadInfo", sBatchLoadInfo)
                        .Parameters.AddWithValue("@CaseLocation", sCaseLocation)
                        .Parameters.AddWithValue("@CurrentCaseVersion", sCurrentCaseVersion)
                        .Parameters.AddWithValue("@UpgradedCaseVersion", sUpgradedCaseVersion)
                        .Parameters.AddWithValue("@CaseSizeOnDisk", sCaseSize)
                        .Parameters.AddWithValue("@CaseFileSize", "0")
                        .Parameters.AddWithValue("@IsCompound", sIsCompound)
                        .Parameters.AddWithValue("@CasesContained", sCasesContained)
                        .Parameters.AddWithValue("@ContainedInCase", sContainedInCase)
                        .Parameters.AddWithValue("@Investigator", sInvestigator)
                        .Parameters.AddWithValue("@BrokerMemory", sBrokerMemory)
                        .Parameters.AddWithValue("@WorkerCount", sWorkerCount)
                        .Parameters.AddWithValue("@WorkerMemory", sWorkerMemory)
                        .Parameters.AddWithValue("@EvidenceProcessed", sEvidenceName)
                        .Parameters.AddWithValue("@EvidenceLocation", sEvidenceLocation)
                        .Parameters.AddWithValue("@EvidenceCustomMetadata", sEvidenceCustomMetadata)
                        .Parameters.AddWithValue("@MimeTypes", sMimeTypes)
                        .Parameters.AddWithValue("@ItemTypes", sItemTypes)
                        .Parameters.AddWithValue("@CreationDate", sCreationDate)
                        .Parameters.AddWithValue("@ModifiedDate", sModifiedDate)
                        .Parameters.AddWithValue("@LoadDataStart", sLoadDataStart)
                        .Parameters.AddWithValue("@LoadDataEnd", sLoadDataEnd)
                        .Parameters.AddWithValue("@LoadTime", sLoadTime)
                        .Parameters.AddWithValue("@ProcessingSpeed", sProcessingSpeed)
                        .Parameters.AddWithValue("@Custodians", sCustodians)
                        .Parameters.AddWithValue("@CustodianCount", sCustodianCount)
                        .Parameters.AddWithValue("@SearchTerm", sSearchTerm)
                        .Parameters.AddWithValue("@SearchSize", sSearchSize)
                        .Parameters.AddWithValue("@HitCount", sHitCount)
                        .Parameters.AddWithValue("@CustodianSearchHit", sCustodianSearchHit)
                        .Parameters.AddWithValue("@ItemCounts", sItemCount)
                        .Parameters.AddWithValue("@CaseUsers", sCaseUsers)
                        .Parameters.AddWithValue("@ReportLoadTime", dReportLoadTime.ToString)
                        .Parameters.AddWithValue("@EvidenceDescription", sEvidenceDescription)
                    End With
                    Try
                        'SQLiteConnection.Open()
                        oUpdateEWSExtractionDataCommand.ExecuteNonQuery()
                        SQLiteConnection.Close()
                    Catch ex As Exception
                        Logger(psUCRTLogFile, "Error blnUpdateSQLiteAllCaseInfo - " & ex.Message.ToString())
                    End Try
                End Using
                SQLiteConnection.Close()
                SQLiteConnection.Dispose()
            End Using
        End If
        blnUpdateSQLiteAllCaseInfo = True
    End Function

    Private Function blnBuildSQLiteDB(ByVal sSQLLiteDBLocation As String) As Boolean
        blnBuildSQLiteDB = False
        Dim Connection As SQLiteConnection

        If Not Directory.Exists(sSQLLiteDBLocation) Then
            Directory.CreateDirectory(sSQLLiteDBLocation)
        End If

        Try
            If Not File.Exists(sSQLLiteDBLocation & "\" & "NuixCaseReports.db3") Then
                SQLiteConnection.CreateFile(sSQLLiteDBLocation & "\" & "NuixCaseReports.db3")

                Connection = New SQLiteConnection("Data Source=" & sSQLLiteDBLocation & "\" & "NuixCaseReports.db3;Version=3;New=False;Read Only=False")
                Connection.Open()

                Using Query As New SQLiteCommand()
                    With Query
                        .Connection = Connection
                        .CommandText = "Create Table NuixReportingInfo(CaseGUID TEXT, ReportLoadStart TEXT, ReportLoadEnd Text, CaseName TEXT, CollectionStatus TEXT, PercentComplete INT, ReportLoadDuration TEXT, BatchLoadInfo TEXT, CaseLocation TEXT, BackUpLocation TEXT, CurrentCaseVersion Text, UpgradedCaseVersion Text, CaseSizeOnDisk INT, CaseFileSize INT, CaseAuditSize INT, IsCompound TEXT, CasesContained TEXT, ContainedInCase TEXT, Investigator TEXT, InvestigatorSessions TEXT, InvalidSessions INT, InvestigatorTimeSummary TEXT, BrokerMemory INT, WorkerCount INT, WorkerMemory INT, EvidenceProcessed TEXT, EvidenceLocation TEXT, EvidenceCustomMetadata TEXT, MimeTypes TEXT, ItemTypes TEXT, CreationDate TEXT, ModifiedDate TEXT,  LoadDataStart TEXT, LoadDataEnd TEXT, LoadTime INT, LoadEvents INT, TotalLoadTime INT, ProcessingSpeed TEXT,Custodians TEXT, CustodianCount INT, SearchTerm TEXT, SearchSize INT, HitCount INT, CustodianSearchHit TEXT, TotalItemCount INT, ItemCounts TEXT, OriginalItems TEXT, DuplicateItems TEXT, CaseUsers TEXT, ReportLoadTime TEXT, NuixLogLocation TEXT, OldestItem TEXT, NewestItem TEXT, Languages TEXT, CustomMetadata TEXT, CaseDescription TEXT, EvidenceDescription TEXT, IrregularItems Text)"
                    End With
                    Try
                        'Connection.Open()
                        Query.ExecuteNonQuery()
                        Connection.Close()
                        Connection.Dispose()
                    Catch ex As Exception
                        Logger(psUCRTLogFile, "Error in blnBuildSQLiteDB - " & ex.Message.ToString())
                    End Try

                End Using

                Connection = New SQLiteConnection("Data Source=" & sSQLLiteDBLocation & "\" & "NuixCaseReports.db3;Version=3;New=False;Read Only=False")
                Connection.Open()

                Using Query As New SQLiteCommand()
                    With Query
                        .Connection = Connection
                        .CommandText = "CREATE TABLE UCRTSessionEvents(CaseGUID TEXT, SessionEvent TEXT, StartDate TEXT, EndDate TEXT, Duration TEXT, User TEXT, Success INT, Failures INT)"
                    End With
                    Try
                        'Connection.Open()
                        Query.ExecuteNonQuery()
                        Connection.Close()
                        Connection.Dispose()
                    Catch ex As Exception
                        Logger(psUCRTLogFile, "Error in Creating UCRTSessionEvents in blnBuildSQLiteDB - " & ex.Message.ToString())
                    End Try

                End Using

                Connection = New SQLiteConnection("Data Source=" & sSQLLiteDBLocation & "\" & "NuixCaseReports.db3;Version=3;New=False;Read Only=False")
                Connection.Open()

                Using Query As New SQLiteCommand()
                    With Query
                        .Connection = Connection
                        .CommandText = "CREATE TABLE UCRTDateRange(CaseGUID TEXT, ItemType TEXT, ItemDate TEXT,  ItemCount INT, Custodian TEXT)"
                    End With
                    Try
                        'Connection.Open()
                        Query.ExecuteNonQuery()
                        Connection.Close()
                        Connection.Dispose()
                    Catch ex As Exception
                        Logger(psUCRTLogFile, "Error in Creating UCRTSessionEvents in blnBuildSQLiteDB - " & ex.Message.ToString())
                    End Try

                End Using

                Connection = New SQLiteConnection("Data Source=" & sSQLLiteDBLocation & "\" & "NuixCaseReports.db3;Version=3;New=False;Read Only=False")
                Connection.Open()

                Using Query As New SQLiteCommand()
                    With Query
                        .Connection = Connection
                        .CommandText = "CREATE TABLE UCRTSearchTermResults(CaseGUID TEXT, SearchTerm TEXT, ItemCount INT, Custodian TEXT, CustodianSearchItemCount, ExportedItems INT)"
                    End With
                    Try
                        'Connection.Open()
                        Query.ExecuteNonQuery()
                        Connection.Close()
                        Connection.Dispose()
                    Catch ex As Exception
                        Logger(psUCRTLogFile, "Error in Creating UCRTSessionEvents in blnBuildSQLiteDB - " & ex.Message.ToString())
                    End Try

                End Using
            End If

        Catch ex As Exception
            MessageBox.Show("Error in blnBuildSQLiteDB - " & ex.ToString)
        End Try


        blnBuildSQLiteDB = True
    End Function

    Private Function blnBuildSQLiteRubyScript(ByVal sSQLiteDBLocation As String) As Boolean
        Dim SQLiteRuby As StreamWriter
        blnBuildSQLiteRubyScript = False
        If File.Exists(sSQLiteDBLocation & "\SQLite.rb_") Then
            Exit Function
        End If

        SQLiteRuby = New StreamWriter(sSQLiteDBLocation & "\SQLite.rb_")

        SQLiteRuby.WriteLine("require 'java'")
        SQLiteRuby.WriteLine("java.sql.DriverManager.registerDriver(org.sqlite.JDBC.new())")
        SQLiteRuby.WriteLine("# This class provides connectivity to SQLite database")
        SQLiteRuby.WriteLine("class SQLite < Database")
        SQLiteRuby.WriteLine("	attr_accessor :file")
        SQLiteRuby.WriteLine("	# file is the db file to open/create")
        SQLiteRuby.WriteLine("	def initialize(file,settings={})")
        SQLiteRuby.WriteLine("		@file = file")
        SQLiteRuby.WriteLine("		@settings = settings")
        SQLiteRuby.WriteLine("	end")
        SQLiteRuby.WriteLine("	# Provides SQLite specific connections")
        SQLiteRuby.WriteLine("	def create_connection()")
        SQLiteRuby.WriteLine("		connection_properties = java.util.Properties.new")
        SQLiteRuby.WriteLine("		@settings.each do |k,v|")
        SQLiteRuby.WriteLine("			connection_properties[k] = v")
        SQLiteRuby.WriteLine("		end")
        SQLiteRuby.WriteLine("		return java.sql.DriverManager.getConnection(""" & "jdbc:sqlite:#{@file}" & """" & ",connection_properties)")
        SQLiteRuby.WriteLine("	end")
        SQLiteRuby.WriteLine("	# Binds data to prepared statement, SQLite has subtle differences")
        SQLiteRuby.WriteLine("	# so must implement its own version from that in class Database")
        SQLiteRuby.WriteLine("	def bind_data(statement,data)")
        SQLiteRuby.WriteLine("		if !data.nil?")
        SQLiteRuby.WriteLine("			data.to_enum.with_index(1) do |value,index|")
        SQLiteRuby.WriteLine("				case value")
        SQLiteRuby.WriteLine("				when Fixnum, Bignum")
        SQLiteRuby.WriteLine("					statement.setLong(index,value)")
        SQLiteRuby.WriteLine("				when Float")
        SQLiteRuby.WriteLine("					#Ruby floats are double precision")
        SQLiteRuby.WriteLine("					statement.setDouble(index,value)")
        SQLiteRuby.WriteLine("				when TrueClass, FalseClass")
        SQLiteRuby.WriteLine("					statement.setBoolean(index,value)")
        SQLiteRuby.WriteLine("				when String")
        SQLiteRuby.WriteLine("                  statement.setString(index,value)")
        SQLiteRuby.WriteLine("				when Java::byte[]")
        SQLiteRuby.WriteLine("					statement.setBytes(index,value)")
        SQLiteRuby.WriteLine("				else")
        SQLiteRuby.WriteLine("					statement.setString(index,value.to_s)")
        SQLiteRuby.WriteLine("				end")
        SQLiteRuby.WriteLine("			end")
        SQLiteRuby.WriteLine("		end")
        SQLiteRuby.WriteLine("	end")
        SQLiteRuby.WriteLine("end")

        SQLiteRuby.Close()
        blnBuildSQLiteRubyScript = True

    End Function
    Private Function blnBuildSQLiteDatabaseScript(ByVal sSQLiteDBLocation As String) As Boolean
        Dim SQLiteDatabaseScript As StreamWriter

        blnBuildSQLiteDatabaseScript = False

        If File.Exists(sSQLiteDBLocation & "\Database.rb_") Then
            blnBuildSQLiteDatabaseScript = True
            Exit Function
        End If
        SQLiteDatabaseScript = New StreamWriter(sSQLiteDBLocation & "\Database.rb_")
        SQLiteDatabaseScript.WriteLine("require 'java'")
        SQLiteDatabaseScript.WriteLine("java_import java.sql.Statement")
        SQLiteDatabaseScript.WriteLine("")
        SQLiteDatabaseScript.WriteLine("# This class wraps a transaction for bulk insertions")
        SQLiteDatabaseScript.WriteLine("class DatabaseInsertBatch")
        SQLiteDatabaseScript.WriteLine("	def initialize(database,sql,max_pending=nil)")
        SQLiteDatabaseScript.WriteLine("		@max_pending = max_pending")
        SQLiteDatabaseScript.WriteLine("		@database = database")
        SQLiteDatabaseScript.WriteLine("		@pending = 0")
        SQLiteDatabaseScript.WriteLine("		@connection = @database.create_connection")
        SQLiteDatabaseScript.WriteLine("		@connection.setAutoCommit(false)")
        SQLiteDatabaseScript.WriteLine("		@statement = @connection.prepareStatement(sql) if !sql.nil?")
        SQLiteDatabaseScript.WriteLine("		@statement_cache = {}")
        SQLiteDatabaseScript.WriteLine("	end")
        SQLiteDatabaseScript.WriteLine("")
        SQLiteDatabaseScript.WriteLine("	# This method will commit whatever is present currently")
        SQLiteDatabaseScript.WriteLine("	# in the transaction")
        SQLiteDatabaseScript.WriteLine("	def commit")
        SQLiteDatabaseScript.WriteLine("		@connection.commit")
        SQLiteDatabaseScript.WriteLine("	end")
        SQLiteDatabaseScript.WriteLine("")
        SQLiteDatabaseScript.WriteLine("	# Inserts some data against the transaction")
        SQLiteDatabaseScript.WriteLine("	# If @max_pending is not nil, will auto commit if")
        SQLiteDatabaseScript.WriteLine("	# the pending insert count is greater than @max_pending")
        SQLiteDatabaseScript.WriteLine("	def insert(data,sql=nil)")
        SQLiteDatabaseScript.WriteLine("		if sql.nil?")
        SQLiteDatabaseScript.WriteLine("			@database.bind_data(@statement,data)")
        SQLiteDatabaseScript.WriteLine("			@statement.executeUpdate")
        SQLiteDatabaseScript.WriteLine("		else")
        SQLiteDatabaseScript.WriteLine("			statement = @statement_cache[sql]")
        SQLiteDatabaseScript.WriteLine("			if statement.nil?")
        SQLiteDatabaseScript.WriteLine("				statement = @statement_cache[sql] = @connection.prepareStatement(sql)")
        SQLiteDatabaseScript.WriteLine("			end")
        SQLiteDatabaseScript.WriteLine("			@database.bind_data(statement,data)")
        SQLiteDatabaseScript.WriteLine("			statement.executeUpdate")
        SQLiteDatabaseScript.WriteLine("		end")
        SQLiteDatabaseScript.WriteLine("		@pending += 1")
        SQLiteDatabaseScript.WriteLine("		if @max_pending && @pending > @max_pending")
        SQLiteDatabaseScript.WriteLine("			commit")
        SQLiteDatabaseScript.WriteLine("			@pending = 0")
        SQLiteDatabaseScript.WriteLine("		end")
        SQLiteDatabaseScript.WriteLine("	end")
        SQLiteDatabaseScript.WriteLine("")
        SQLiteDatabaseScript.WriteLine("	# Ends this transaction and closes connection")
        SQLiteDatabaseScript.WriteLine("	def end")
        SQLiteDatabaseScript.WriteLine("		commit")
        SQLiteDatabaseScript.WriteLine("		@connection.close")
        SQLiteDatabaseScript.WriteLine("	end")
        SQLiteDatabaseScript.WriteLine("end")
        SQLiteDatabaseScript.WriteLine("")
        SQLiteDatabaseScript.WriteLine("# This generically represents databases, provides most common")
        SQLiteDatabaseScript.WriteLine("# functionality, not to be instantiated directly, instead create")
        SQLiteDatabaseScript.WriteLine("# instances of classes that derive from this")
        SQLiteDatabaseScript.WriteLine("class Database")
        SQLiteDatabaseScript.WriteLine("	# Create a connection, must be overidden in derived classes")
        SQLiteDatabaseScript.WriteLine("	def create_connection")
        SQLiteDatabaseScript.WriteLine("		raise " & """" & "Derived class must provide create_connection method" & """")
        SQLiteDatabaseScript.WriteLine("	end")
        SQLiteDatabaseScript.WriteLine("")
        SQLiteDatabaseScript.WriteLine("	# Creates an insert batch class (transaction) against")
        SQLiteDatabaseScript.WriteLine("	# this database object")
        SQLiteDatabaseScript.WriteLine("	def create_insert_batch(sql,max_pending=nil)")
        SQLiteDatabaseScript.WriteLine("		return DatabaseInsertBatch.new(self,sql,max_pending)")
        SQLiteDatabaseScript.WriteLine("	end")
        SQLiteDatabaseScript.WriteLine("")
        SQLiteDatabaseScript.WriteLine("	# Creates and closes an insert batch (transaction) for")
        SQLiteDatabaseScript.WriteLine("	# a provided block")
        SQLiteDatabaseScript.WriteLine("	# max_pending dictates how often to auto commit when provided")
        SQLiteDatabaseScript.WriteLine("	def batch_insert(sql,max_pending=nil,&block)")
        SQLiteDatabaseScript.WriteLine("		batch = create_insert_batch(sql,max_pending)")
        SQLiteDatabaseScript.WriteLine("		yield batch")
        SQLiteDatabaseScript.WriteLine("	ensure")
        SQLiteDatabaseScript.WriteLine("		batch.end")
        SQLiteDatabaseScript.WriteLine("	end")
        SQLiteDatabaseScript.WriteLine("")
        SQLiteDatabaseScript.WriteLine("	# Exceutes a basic update statement")
        SQLiteDatabaseScript.WriteLine("	# data is bound to prepared statement when provided")
        SQLiteDatabaseScript.WriteLine("	def update(sql,data=nil)")
        SQLiteDatabaseScript.WriteLine("		connection = create_connection")
        SQLiteDatabaseScript.WriteLine("		statement = connection.prepareStatement(sql)")
        SQLiteDatabaseScript.WriteLine("		bind_data(statement,data)")
        SQLiteDatabaseScript.WriteLine("		statement.executeUpdate")
        SQLiteDatabaseScript.WriteLine("	rescue => exc")
        SQLiteDatabaseScript.WriteLine("		raise")
        SQLiteDatabaseScript.WriteLine("	ensure")
        SQLiteDatabaseScript.WriteLine("		connection.close if !connection.nil?")
        SQLiteDatabaseScript.WriteLine("	end")
        SQLiteDatabaseScript.WriteLine("")
        SQLiteDatabaseScript.WriteLine("	# Executes an update statement and returns the")
        SQLiteDatabaseScript.WriteLine("	# newly created key, if there was one")
        SQLiteDatabaseScript.WriteLine("	# data is bound to prepared statement when provided")
        SQLiteDatabaseScript.WriteLine("	def insert(sql,data=nil)")
        SQLiteDatabaseScript.WriteLine("		inserted_key = nil")
        SQLiteDatabaseScript.WriteLine("		connection = nil")
        SQLiteDatabaseScript.WriteLine("		connection = create_connection")
        SQLiteDatabaseScript.WriteLine("		statement = connection.prepareStatement(sql,Statement::RETURN_GENERATED_KEYS)")
        SQLiteDatabaseScript.WriteLine("		bind_data(statement,data)")
        SQLiteDatabaseScript.WriteLine("		statement.executeUpdate")
        SQLiteDatabaseScript.WriteLine("		result_set = statement.getGeneratedKeys")
        SQLiteDatabaseScript.WriteLine("		if !result_set.nil? && result_set.next")
        SQLiteDatabaseScript.WriteLine("			insert_key = get_result_column_value(result_set,1)")
        SQLiteDatabaseScript.WriteLine("		end")
        SQLiteDatabaseScript.WriteLine("		return insert_key")
        SQLiteDatabaseScript.WriteLine("	rescue => exc")
        SQLiteDatabaseScript.WriteLine("		raise")
        SQLiteDatabaseScript.WriteLine("	ensure")
        SQLiteDatabaseScript.WriteLine("		connection.close if !connection.nil?")
        SQLiteDatabaseScript.WriteLine("	end")
        SQLiteDatabaseScript.WriteLine("")
        SQLiteDatabaseScript.WriteLine("	# Executes a scalar query (single result)")
        SQLiteDatabaseScript.WriteLine("	# data is bound to prepared statement when provided")
        SQLiteDatabaseScript.WriteLine("	def scalar(sql,data=nil)")
        SQLiteDatabaseScript.WriteLine("		connection = create_connection")
        SQLiteDatabaseScript.WriteLine("		statement = connection.prepareStatement(sql)")
        SQLiteDatabaseScript.WriteLine("		bind_data(statement,data)")
        SQLiteDatabaseScript.WriteLine("		result_set = statement.executeQuery")
        SQLiteDatabaseScript.WriteLine("		if !result_set.nil? && result_set.next")
        SQLiteDatabaseScript.WriteLine("			return get_result_column_value(result_set,1)")
        SQLiteDatabaseScript.WriteLine("		end")
        SQLiteDatabaseScript.WriteLine("	rescue => exc")
        SQLiteDatabaseScript.WriteLine("		raise")
        SQLiteDatabaseScript.WriteLine("	ensure")
        SQLiteDatabaseScript.WriteLine("		connection.close if !connection.nil?")
        SQLiteDatabaseScript.WriteLine("	end")
        SQLiteDatabaseScript.WriteLine("")
        SQLiteDatabaseScript.WriteLine("	# Runs a select query")
        SQLiteDatabaseScript.WriteLine("	# data is bound to prepared statement when provided")
        SQLiteDatabaseScript.WriteLine("	# hashed, true will return records as hash of COLUMN_NAME => value,")
        SQLiteDatabaseScript.WriteLine("	#  otherwise each record will be an array of values")
        SQLiteDatabaseScript.WriteLine("	# when provided a block will yield each record in turn, without a")
        SQLiteDatabaseScript.WriteLine("	# block will return array of results")
        SQLiteDatabaseScript.WriteLine("	def query(sql,data=nil,hashed=true,&block)")
        SQLiteDatabaseScript.WriteLine("		connection = nil")
        SQLiteDatabaseScript.WriteLine("		connection = create_connection")
        SQLiteDatabaseScript.WriteLine("		statement = connection.prepareStatement(sql,Statement::RETURN_GENERATED_KEYS)")
        SQLiteDatabaseScript.WriteLine("		bind_data(statement,data)")
        SQLiteDatabaseScript.WriteLine("		result_set = statement.executeQuery")
        SQLiteDatabaseScript.WriteLine("		column_metadata = result_set.getMetaData")
        SQLiteDatabaseScript.WriteLine("		column_count = column_metadata.getColumnCount")
        SQLiteDatabaseScript.WriteLine("		column_names = (1..column_count).map{|c|column_metadata.getColumnName(c)}")
        SQLiteDatabaseScript.WriteLine("		records = []")
        SQLiteDatabaseScript.WriteLine("		while result_set.next")
        SQLiteDatabaseScript.WriteLine("			values = []")
        SQLiteDatabaseScript.WriteLine("			(1..column_count).each do |c|")
        SQLiteDatabaseScript.WriteLine("				values << get_result_column_value(result_set,c,column_metadata)")
        SQLiteDatabaseScript.WriteLine("			end")
        SQLiteDatabaseScript.WriteLine("			record = nil")
        SQLiteDatabaseScript.WriteLine("			if !hashed")
        SQLiteDatabaseScript.WriteLine("				record = values")
        SQLiteDatabaseScript.WriteLine("			else")
        SQLiteDatabaseScript.WriteLine("				hashed_values = {}")
        SQLiteDatabaseScript.WriteLine("				column_names.each_with_index do |name,column_index|")
        SQLiteDatabaseScript.WriteLine("					hashed_values[name] = values[column_index]")
        SQLiteDatabaseScript.WriteLine("				end")
        SQLiteDatabaseScript.WriteLine("				record = hashed_values")
        SQLiteDatabaseScript.WriteLine("			end")
        SQLiteDatabaseScript.WriteLine("			if block_given?")
        SQLiteDatabaseScript.WriteLine("				yield record")
        SQLiteDatabaseScript.WriteLine("			else")
        SQLiteDatabaseScript.WriteLine("				records << record")
        SQLiteDatabaseScript.WriteLine("			end")
        SQLiteDatabaseScript.WriteLine("		end")
        SQLiteDatabaseScript.WriteLine("		return records if !block_given?")
        SQLiteDatabaseScript.WriteLine("	rescue => exc")
        SQLiteDatabaseScript.WriteLine("		raise")
        SQLiteDatabaseScript.WriteLine("	ensure")
        SQLiteDatabaseScript.WriteLine("		connection.close if !connection.nil?")
        SQLiteDatabaseScript.WriteLine("	end")
        SQLiteDatabaseScript.WriteLine("")
        SQLiteDatabaseScript.WriteLine("	# Helper method for obtaining value from a result set for a given column")
        SQLiteDatabaseScript.WriteLine("	def get_result_column_value(result_set,column_index,column_metadata=nil)")
        SQLiteDatabaseScript.WriteLine("		column_metadata ||= result_set.getMetaData")
        SQLiteDatabaseScript.WriteLine("		value = nil")
        SQLiteDatabaseScript.WriteLine("        puts column_metadata.getColumnTypeName(column_index).downcase")
        SQLiteDatabaseScript.WriteLine("		case column_metadata.getColumnTypeName(column_index).downcase")
        SQLiteDatabaseScript.WriteLine("		when """ & "integer" & """" & ", " & """" & "numeric" & """ ")
        SQLiteDatabaseScript.WriteLine("			value = result_set.getLong(column_index)")
        SQLiteDatabaseScript.WriteLine("		when """ & "float" & """")
        SQLiteDatabaseScript.WriteLine("			value = result_set.getFloat(column_index)")
        SQLiteDatabaseScript.WriteLine("		when """ & "double" & """")
        SQLiteDatabaseScript.WriteLine("			value = result_set.getDouble(column_index)")
        SQLiteDatabaseScript.WriteLine("		when """ & "blob" & """" & ", " & """" & "binary" & """ ")
        SQLiteDatabaseScript.WriteLine("			value = result_set.getBytes(column_index).to_s")
        SQLiteDatabaseScript.WriteLine("		when """ & "varchar" & """")
        SQLiteDatabaseScript.WriteLine("			value = result_set.getString(column_index)")
        SQLiteDatabaseScript.WriteLine("		else")
        SQLiteDatabaseScript.WriteLine("			value = result_set.getString(column_index)")
        SQLiteDatabaseScript.WriteLine("		end")
        SQLiteDatabaseScript.WriteLine("		return value")
        SQLiteDatabaseScript.WriteLine("	end")
        SQLiteDatabaseScript.WriteLine("")
        SQLiteDatabaseScript.WriteLine("	# Binds data to a prepared statement")
        SQLiteDatabaseScript.WriteLine("	def bind_data(statement,data)")
        SQLiteDatabaseScript.WriteLine("		if !data.nil?")
        SQLiteDatabaseScript.WriteLine("			data.to_enum.with_index(1) do |value,index|")
        SQLiteDatabaseScript.WriteLine("				case value")
        SQLiteDatabaseScript.WriteLine("				when Fixnum")
        SQLiteDatabaseScript.WriteLine("					statement.setInt(index,value)")
        SQLiteDatabaseScript.WriteLine("				when Bignum")
        SQLiteDatabaseScript.WriteLine("					statement.setLong(index,value)")
        SQLiteDatabaseScript.WriteLine("				when Float")
        SQLiteDatabaseScript.WriteLine("					statement.setDouble(index,value)")
        SQLiteDatabaseScript.WriteLine("				when TrueClass, FalseClass")
        SQLiteDatabaseScript.WriteLine("					statement.setBoolean(index,value)")
        SQLiteDatabaseScript.WriteLine("				when Time")
        SQLiteDatabaseScript.WriteLine("					statement.setTimestamp(index,java.sql.Timestamp.new(value.to_i*1000))")
        SQLiteDatabaseScript.WriteLine("				when Java::byte[]")
        SQLiteDatabaseScript.WriteLine("					statement.setBytes(index,value)")
        SQLiteDatabaseScript.WriteLine("				when String")
        SQLiteDatabaseScript.WriteLine("					#ASCII-8BIT encoded String is essentially Ruby byte array")
        SQLiteDatabaseScript.WriteLine("					if value.encoding.name == """ & "ASCII-8BIT" & """")
        SQLiteDatabaseScript.WriteLine("						statement.setBytes(index,value.to_java_bytes)")
        SQLiteDatabaseScript.WriteLine("					else")
        SQLiteDatabaseScript.WriteLine("						statement.setString(index,value)")
        SQLiteDatabaseScript.WriteLine("					end")
        SQLiteDatabaseScript.WriteLine("				else")
        SQLiteDatabaseScript.WriteLine("					statement.setString(index,value.to_s)")
        SQLiteDatabaseScript.WriteLine("				end")
        SQLiteDatabaseScript.WriteLine("			end")
        SQLiteDatabaseScript.WriteLine("		end")
        SQLiteDatabaseScript.WriteLine("	end")
        SQLiteDatabaseScript.WriteLine("end")

        SQLiteDatabaseScript.Close()


    End Function

    Public Function blnGetAllNuixCaseFiles(ByVal sSQLiteDBLocation As String, ByVal sNuixConsoleVersion As String, ByVal asCaseFolders() As String, ByVal bMigrateCases As Boolean, ByVal bGetFileSystemDataOnly As Boolean, ByVal bIncludeDiskSize As Boolean, ByRef lstNuixCases As List(Of String), ByRef lstCaseGUIDs As List(Of String)) As Boolean
        Dim CurrentDirectory As DirectoryInfo
        Dim asSubDirectories(0) As String
        Dim bStatus As Boolean
        Dim dblTotalCaseSize As Double
        Dim lstParallelProcessingSettings As List(Of String)
        Dim lstXMLFiles As List(Of String)
        Dim sWorkerTempDir As String
        Dim sWorkerCount As String
        Dim sBrokerMemory As String
        Dim sWorkerMemory As String
        Dim sEvidenceName As String
        Dim sEvidenceLocations As String
        Dim sEvidenceCustodians As String
        Dim sEvidenceDescription As String
        Dim sEvidenceCustomMetadata As String
        Dim NuixCaseFile As Xml.XmlDocument
        Dim ChildNodes As Xml.XmlNodeList
        Dim oMetaDataNodeList As Xml.XmlNodeList
        Dim sGuid As String
        Dim sCaseName As String
        Dim sCreateDate As String
        Dim sInvestigator As String
        Dim sNuixName As String
        Dim sNuixVersion As String
        Dim iEvidenceCustodiansCount As Integer
        Dim sCollectionStatus As String
        Dim sModifiedDate As String
        Dim asDateParts() As String
        Dim sDate As String
        Dim sTime As String
        Dim sCaseLockFile As String
        Dim dCreateDate As DateTime
        Dim dReportLoadTime As DateTime


        lstParallelProcessingSettings = New List(Of String)
        lstXMLFiles = New List(Of String)

        Array.Clear(asSubDirectories, 0, asSubDirectories.Length)
        ReDim asSubDirectories(0)

        For iCounter = 0 To asCaseFolders.Length - 1
            Try
                Array.Clear(asSubDirectories, 0, asSubDirectories.Length)
                ReDim asSubDirectories(0)
                CurrentDirectory = New DirectoryInfo(asCaseFolders(iCounter))
                If Not CurrentDirectory.Attributes.HasFlag(FileAttributes.ReadOnly) Then
                    Try
                        For Each Directory In CurrentDirectory.GetDirectories
                            If asSubDirectories.Length > 0 Then
                                If Not IsNothing(asSubDirectories(0)) Then
                                    ReDim Preserve asSubDirectories(asSubDirectories.Length)
                                    asSubDirectories(asSubDirectories.Length - 1) = Directory.FullName
                                Else
                                    asSubDirectories(0) = Directory.FullName
                                End If
                            Else
                                asSubDirectories(0) = Directory.FullName
                            End If
                        Next

                        If asSubDirectories.Length > 0 Then
                            If Not IsNothing(asSubDirectories(0)) Then
                                bStatus = blnGetAllNuixCaseFiles(sSQLiteDBLocation, sNuixConsoleVersion, asSubDirectories, bMigrateCases, bGetFileSystemDataOnly, bIncludeDiskSize, lstNuixCases, lstCaseGUIDs)
                            End If
                        End If

                        For Each Files In CurrentDirectory.GetFiles
                            If Files.Name = "case.fbi2" Then
                                lstNuixCases.Add(Files.FullName)
                                dblTotalCaseSize = 0.0
                                sModifiedDate = System.IO.File.GetLastWriteTime(Files.FullName)
                                lstXMLFiles.Clear()
                                Dim di As New IO.DirectoryInfo(asCaseFolders(iCounter))
                                If bIncludeDiskSize = True Then
                                    bStatus = blnGetCaseFolderSize(di, dblTotalCaseSize, lstParallelProcessingSettings, lstXMLFiles)
                                End If

                                If lstParallelProcessingSettings.Count > 0 Then
                                    bStatus = blnParseParallelProcessingSettings(lstParallelProcessingSettings(0).ToString, sWorkerTempDir, sWorkerCount, sBrokerMemory, sWorkerMemory)
                                End If
                                If lstXMLFiles.Count > 0 Then
                                    sEvidenceName = ""
                                    sEvidenceCustodians = ""
                                    sEvidenceLocations = ""
                                    sEvidenceDescription = ""
                                    sEvidenceCustomMetadata = ""
                                    iEvidenceCustodiansCount = 0
                                    For Each XMLFile In lstXMLFiles
                                        bStatus = blnParseEvidenceFiles(XMLFile.ToString, sEvidenceName, sEvidenceLocations, sEvidenceCustodians, iEvidenceCustodiansCount, sEvidenceDescription, sEvidenceCustomMetadata)
                                    Next
                                End If

                                NuixCaseFile = New Xml.XmlDocument

                                NuixCaseFile.Load(Files.FullName)
                                oMetaDataNodeList = NuixCaseFile.GetElementsByTagName("metadata")
                                For Each MetadataNode In oMetaDataNodeList
                                    If MetadataNode.haschildnodes Then
                                        ChildNodes = MetadataNode.childnodes
                                        For Each Child In ChildNodes
                                            If Child.name = "guid" Then
                                                sGuid = Child.innertext
                                                lstCaseGUIDs.Add(sGuid)
                                            ElseIf Child.name = "name" Then
                                                sCaseName = Child.innertext
                                                sCaseName = sCaseName.Trim
                                            ElseIf Child.name = "creation-date" Then
                                                sCreateDate = Child.innertext
                                                asDateParts = Split(sCreateDate, "T")
                                                sDate = asDateParts(0)
                                                sTime = asDateParts(1).Substring(0, asDateParts(1).IndexOf(","))
                                                sCreateDate = sDate & " " & sTime
                                                dCreateDate = Date.Parse(sCreateDate, System.Globalization.DateTimeFormatInfo.InvariantInfo, Globalization.DateTimeStyles.None)
                                                sCreateDate = dCreateDate.ToString
                                            ElseIf Child.name = "investigator" Then
                                                sInvestigator = Child.innertext
                                            ElseIf Child.name = "saved-by-product" Then
                                                sNuixName = Child.GetAttribute("name")
                                                sNuixVersion = Child.GetAttribute("version")
                                            End If
                                        Next
                                    End If
                                Next

                                If (sNuixVersion = "7.1.60003") And ((sNuixConsoleVersion = "7.2.1") Or (sNuixConsoleVersion = "7.2.2") Or (sNuixConsoleVersion = "7.2.3") Or (sNuixConsoleVersion = "7.2.4")) Then
                                    sCollectionStatus = "File System Info Collected - Waiting for Case Data"
                                ElseIf (sNuixVersion = "7.3.65878") And (sNuixConsoleVersion = "7.4.0") Then
                                    sCollectionStatus = "File System Info Collected - Waiting for Case Data"
                                ElseIf (sNuixVersion <> sNuixConsoleVersion) Then
                                    If bMigrateCases = True Then
                                        sCollectionStatus = "File System Info Collected - Case Migrating"
                                    Else
                                        If bGetFileSystemDataOnly = True Then
                                            sCollectionStatus = "File System Info Collected"
                                        Else
                                            sCollectionStatus = "File System Info Collected - Case Version Mismatch"
                                        End If
                                    End If
                                Else
                                    If bGetFileSystemDataOnly = True Then
                                        sCollectionStatus = "File System Info Collected"
                                    Else
                                        sCollectionStatus = "File System Info Collected - Waiting for Case Data"
                                    End If
                                End If
                                dReportLoadTime = DateTime.Now

                                sCaseLockFile = System.IO.Path.Combine(CurrentDirectory.FullName, "case.lock")
                                If File.Exists(sCaseLockFile) Then
                                    sCollectionStatus = "Case Locked"
                                End If

                                bStatus = blnUpdateSQLiteAllCaseInfo(sSQLiteDBLocation, sGuid, sCollectionStatus, 0, sCaseName, "", dblTotalCaseSize, Files.DirectoryName, sNuixVersion, "", "", "", "", sInvestigator, sBrokerMemory, sWorkerCount, sWorkerMemory, sEvidenceName, sEvidenceLocations, sEvidenceCustomMetadata, "", "", sCreateDate, sModifiedDate, "", "", "", "", sEvidenceCustodians, iEvidenceCustodiansCount, "", txtSearchTerm.Text, "", "", "", "", dReportLoadTime, sEvidenceDescription)
                            End If
                        Next
                        Array.Clear(asSubDirectories, 0, asSubDirectories.Length)
                        ReDim asSubDirectories(0)
                    Catch ex As Exception
                        MessageBox.Show("blnGetAllNuixCaseFiles Error - " & ex.ToString, "blnGetAllNuixCaseFiles")
                    End Try
                End If

            Catch ex As Exception
                Logger(psUCRTLogFile, ex.ToString)
            End Try
        Next

        blnGetAllNuixCaseFiles = True

    End Function

    Private Function blnCopyCase(ByVal sCurrentCaseName As String, ByVal sCurrentCaseLocation As String, ByVal sBackUpLocation As String, ByVal bCopyCases As Boolean, ByRef sBackUpCaseLocation As String, ByVal dblCaseSizeOnDisk As Double) As Boolean
        blnCopyCase = False
        Dim sNow As String
        Dim dblDiskSizeAvailable As Double
        Dim drvDrive As DriveInfo
        Dim sDriveLetter As String

        sDriveLetter = Path.GetPathRoot(sBackUpLocation)
        drvDrive = New DriveInfo(sDriveLetter)
        dblDiskSizeAvailable = drvDrive.TotalFreeSpace
        If dblDiskSizeAvailable > dblCaseSizeOnDisk Then
            sNow = DateTime.Now.ToString("yyyy-MM-dd hh-mm-ss")
            If Directory.Exists(sCurrentCaseLocation) Then
                If bCopyCases = True Then
                    My.Computer.FileSystem.CreateDirectory(sBackUpLocation & "\" & sCurrentCaseName & "-" & sNow)
                    sBackUpCaseLocation = sBackUpLocation & "\" & sCurrentCaseName & "-" & sNow
                    My.Computer.FileSystem.CopyDirectory(sCurrentCaseLocation, sBackUpLocation & "\" & sCurrentCaseName & "-" & sNow)
                Else
                    sBackUpCaseLocation = sBackUpLocation & "\" & sCurrentCaseName & "-" & sNow
                    My.Computer.FileSystem.MoveDirectory(sCurrentCaseLocation, sBackUpLocation & "\" & sCurrentCaseName)
                End If
            End If
        Else
            MessageBox.Show("You do not have enough disk space to Copy/Move " & sCurrentCaseName & " to " & sBackUpCaseLocation)
            blnCopyCase = False
        End If

        blnCopyCase = True
    End Function
    Private Function blnBuildBatchFiles(ByVal sBatchFileName As String, ByVal sReportType As String, ByVal sNuixAppLocation As String, ByVal sNMSSourceType As String, ByVal sNMSLocation As String, ByVal sNMSUserName As String, ByVal sNMSAdminInfo As String, sNumberOfWorkers As String, ByVal sDirectoryName As String, ByVal sRubyFileName As String, ByVal sNuixAppMemory As String, ByVal sNuixLogDir As String, ByVal sRegistryServer As String) As Boolean

        blnBuildBatchFiles = False

        Dim CustodianBatchFile As StreamWriter
        Dim sLicenceSourceType As String

        If sNMSSourceType = "Desktop" Then
            sLicenceSourceType = "-licencesourcetype dongle"
        ElseIf sNMSSourceType = "Desktop (dongleless)" Then
            sLicenceSourceType = "-Dnuix.licence.handlers=system"
        Else
            sLicenceSourceType = "-licencesourcetype server -licencesourcelocation " & sNMSLocation & " -licencetype " & psLicenseType
        End If
        CustodianBatchFile = New StreamWriter(sBatchFileName)
        CustodianBatchFile.WriteLine("::TITLE is the destination SMTP Address")
        CustodianBatchFile.WriteLine("@TITLE " & sReportType)
        CustodianBatchFile.WriteLine("::Enter NMS Username on Line 4")
        CustodianBatchFile.WriteLine("@SET NUIX_USERNAME=" & sNMSUserName)
        CustodianBatchFile.WriteLine("::Enter NMS Username on Line 6")
        CustodianBatchFile.WriteLine("@SET NUIX_PASSWORD=" & sNMSAdminInfo)

        CustodianBatchFile.Write("""" & sNuixAppLocation & """")
        If sRegistryServer <> vbNullString Then
            CustodianBatchFile.Write(" -Dnuix.registry.servers=" & Trim(sRegistryServer))
        End If
        CustodianBatchFile.Write(" " & Trim(sLicenceSourceType))
        CustodianBatchFile.Write(" -licenceworkers " & sNumberOfWorkers & " " & sNuixAppMemory & " -Dnuix.export.mailbox.maximumFileSizePerMailbox=2GB -Dnuix.logdir=" & """" & sNuixLogDir & """" & " ")

        CustodianBatchFile.WriteLine("""" & sRubyFileName & """")

        'CustodianBatchFile.WriteLine("@pause")
        CustodianBatchFile.WriteLine("@exit")

        CustodianBatchFile.Close()

        blnBuildBatchFiles = True
    End Function

    Private Function blnParseParallelProcessingSettings(ByVal sParallelProcessingSetting As String, ByRef sWorkerTempDir As String, ByRef sWorkerCount As String, ByRef sBrokerMemory As String, ByRef sWorkerMemory As String) As Boolean
        Dim ProcessingSettingsStream As StreamReader
        Dim sCurrentRow As String
        Dim asWorkerTemp() As String
        Dim asWorkerCount() As String
        Dim asBrokerMemory() As String
        Dim asWorkerMemory() As String

        blnParseParallelProcessingSettings = False

        ProcessingSettingsStream = New StreamReader(sParallelProcessingSetting)

        While Not ProcessingSettingsStream.EndOfStream
            sCurrentRow = ProcessingSettingsStream.ReadLine
            If sCurrentRow.Contains("workerTempDirectory") Then
                asWorkerTemp = Split(sCurrentRow, "=")
                sWorkerTempDir = asWorkerTemp(1)
            ElseIf sCurrentRow.Contains("workerCount") Then
                asWorkerCount = Split(sCurrentRow, "=")
                sWorkerCount = asWorkerCount(1)
            ElseIf sCurrentRow.Contains("brokerMemory") Then
                asBrokerMemory = Split(sCurrentRow, "=")
                sBrokerMemory = asBrokerMemory(1)
            ElseIf sCurrentRow.Contains("workerMemory") Then
                asWorkerMemory = Split(sCurrentRow, "=")
                sWorkerMemory = asWorkerMemory(1)
            End If
        End While
        ProcessingSettingsStream.Close()


        blnParseParallelProcessingSettings = True

    End Function

    Private Function blnParseEvidenceFiles(ByVal sEvidenceXML As String, ByRef sEvidenceName As String, ByRef sEvidenceLocations As String, ByRef sEvidenceCustodians As String, ByRef iEvidenceCustodiansCount As Integer, ByRef sEvidenceDescription As String, ByRef sEvidenceCustomMetadata As String) As Boolean
        Dim CaseEvidenceFile As Xml.XmlDocument
        Dim ChildNodes As Xml.XmlNodeList
        Dim oCustodianNodes As Xml.XmlNodeList
        Dim oDataRootNodes As Xml.XmlNodeList
        Dim oFileNodes As Xml.XmlNodeList
        Dim oCustomNode As Xml.XmlNodeList

        Dim oMetaDataNodeList As Xml.XmlNodeList

        blnParseEvidenceFiles = False

        CaseEvidenceFile = New Xml.XmlDocument

        CaseEvidenceFile.Load(sEvidenceXML)
        oMetaDataNodeList = CaseEvidenceFile.GetElementsByTagName("evidence")
        For Each MetadataNode In oMetaDataNodeList
            If MetadataNode.haschildnodes Then
                ChildNodes = MetadataNode.childnodes
                For Each Child In ChildNodes
                    If Child.name = "name" Then
                        sEvidenceName = sEvidenceName & Child.innertext & ";"
                    ElseIf Child.name = "data-roots" Then
                        oDataRootNodes = Child.childnodes
                        For Each oDataRootNode In oDataRootNodes
                            oFileNodes = oDataRootNode.childnodes
                            For Each oFileNode In oFileNodes
                                If oFileNode.name = "file" Then
                                    sEvidenceLocations = sEvidenceLocations & oFileNode.attributes("location").value & ";"
                                End If
                            Next
                        Next
                    ElseIf Child.name = "description" Then
                        sEvidenceDescription = sEvidenceDescription & Child.innertext & ";"
                    ElseIf Child.name = "initialCustodian" Then
                        oCustodianNodes = Child.childnodes
                        For Each custodianNode In oCustodianNodes
                            If custodianNode.name = "name" Then
                                sEvidenceCustodians = sEvidenceCustodians & custodianNode.innertext & ";"
                                iEvidenceCustodiansCount = iEvidenceCustodiansCount + 1
                            End If
                        Next
                    ElseIf Child.name = "metadata" Then
                        oCustomNode = Child.childnodes
                        For Each oValueNode In oCustomNode
                            sEvidenceCustomMetadata = sEvidenceCustomMetadata & oValueNode.attributes("name").value & "::" & oValueNode.attributes("value").value & ";"
                        Next
                    End If
                Next
            End If
        Next

        blnParseEvidenceFiles = True

    End Function


    Private Function blnGetCaseFolderSize(ByVal di As IO.DirectoryInfo, ByRef dblTotalCaseSize As Double, ByRef lstParallelProcessingSettings As List(Of String), ByRef lstXMLFiles As List(Of String)) As Boolean
        Try
            For Each d In di.GetDirectories()
                ProcessData(d, dblTotalCaseSize, lstParallelProcessingSettings, lstXMLFiles)
                blnGetCaseFolderSize(d, dblTotalCaseSize, lstParallelProcessingSettings, lstXMLFiles)
            Next
        Catch
        End Try
    End Function

    Private Sub ProcessData(ByVal di As IO.DirectoryInfo, ByRef dblCaseSize As Double, ByRef lstParallelProcessingSettings As List(Of String), ByRef lstXMLFiles As List(Of String))
        Dim strFileSize As String = ""
        Dim fi As IO.FileInfo

        Try
            di.GetFiles("*.*", SearchOption.AllDirectories)
        Catch
        End Try

        Try
            Dim aryFi As IO.FileInfo() = di.GetFiles("*.*")

            For Each fi In aryFi
                If fi.Exists Then
                    If fi.Name = "parallelprocessingsettings.properties" Then
                        lstParallelProcessingSettings.Add(fi.FullName)
                    ElseIf fi.Name.Contains(".xml") Then
                        lstXMLFiles.Add(fi.FullName)
                    End If
                    dblCaseSize = dblCaseSize + fi.Length
                End If
            Next
        Catch
        End Try
    End Sub

    Sub subCaseFileSearch(ByVal sDir As String, ByRef CaseName As List(Of String), ByRef CasePath As List(Of String), ByRef CaseSize As List(Of Double))
        Dim d As String
        Dim Length As String
        Dim currentdirectory As DirectoryInfo
        Dim extension As String

        Try
            currentdirectory = New DirectoryInfo(sDir)

            For Each File In currentdirectory.GetFiles

                extension = Path.GetExtension(File.ToString)
                If extension = ".fbi2" Then
                    CaseName.Add(File.Name)
                    CasePath.Add(File.DirectoryName)
                    CaseSize.Add(File.Length)
                End If
            Next

            Length = Directory.GetDirectories(sDir).Length
            If Length > 0 Then
                For Each d In Directory.GetDirectories(sDir)
                    subCaseFileSearch(d, CaseName, CasePath, CaseSize)
                Next
            End If
        Catch ex As Exception
            MsgBox("Directory Search" & ex.ToString)
        End Try

    End Sub

    Private Sub CaseFinder_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Dim Tnode As TreeNode
        Dim drives As System.Collections.ObjectModel.ReadOnlyCollection(Of IO.DriveInfo)
        Dim rootDir As String

        Me.Width = 900
        Me.Height = 550
        Me.StartPosition = FormStartPosition.CenterScreen
        cboUpgradeCasees.Text = "No"
        lblSearchTerm.Hide()
        txtSearchTerm.Hide()
        grpSearchTerm.Hide()
        btnSearchTermFile.Hide()
        chkExportSearchResults.Hide()
        lblExportLocation.Hide()
        txtExportLocation.Hide()
        btnExportLocation.Hide()
        chkExportOnly.Hide()
        chkExportSearchResults.Enabled = False
        cboExportType.Enabled = False
        cboSizeReporting.Text = "Bytes"
        psShowSizeIn = "Bytes"
        Dim value As System.Version = My.Application.Info.Version
        radFileSystem.Checked = True

        Me.Text = "Universal Case Reporting tool - " & value.ToString

        cboExportType.Hide()
        txtExportLocation.Enabled = False
        btnExportLocation.Enabled = False
        lblExportLocation.Enabled = False

        treeViewFolders.Nodes.Clear()
        lblBackUpLocation.Enabled = False
        txtBackupLocation.Enabled = False
        btnBackupLocationChooser.Enabled = False
        cboLicenseType.Text = "Server"
        txtNMSAddress.Text = "127.0.0.1:27443"
        txtNMSUserName.Text = "nuixadmin"
        lblCalculateProcessingSpeeds.Hide()
        cboCalculateProcessingSpeeds.Hide()
        cboCopyMoveCases.Enabled = False

        drives = My.Computer.FileSystem.Drives
        For i As Integer = 0 To drives.Count - 1
            If Not drives(i).IsReady Then
                Continue For
            End If
            Tnode = treeViewFolders.Nodes.Add(drives(i).ToString)
            rootDir = drives(i).Name
            AddAllFolders(Tnode, rootDir)
        Next
        plstSoureFolders = New List(Of String)

        pbNoMoreJobs = True
    End Sub

    Private Sub PopulateTreeView(ByVal dir As String, ByVal parentNode As TreeNode)
        Dim folder As String = String.Empty
        Try
            Dim folders() As String = IO.Directory.GetDirectories(dir)
            If folders.Length <> 0 Then
                Dim childNode As TreeNode = Nothing
                For Each folder In folders
                    childNode = New TreeNode(folder)
                    parentNode.Nodes.Add(childNode)
                Next
            End If
        Catch ex As UnauthorizedAccessException
            parentNode.Nodes.Add(folder & ": Access Denied")
        End Try
    End Sub

    Private Sub AddAllFolders(ByVal TNode As TreeNode, ByVal FolderPath As String)
        Dim SubFolderNode As TreeNode
        Dim di As IO.DirectoryInfo

        '  Create a new ImageList
        Dim MyImages As New ImageList()

        'Add the files to treeview
        di = New DirectoryInfo(FolderPath)
        If di.Attributes <> FileAttributes.ReadOnly Then
            Try
                For Each FolderNode As String In IO.Directory.GetDirectories(FolderPath)
                    If FolderNode <> vbNullString Then '
                        SubFolderNode = TNode.Nodes.Add(FolderNode.Substring(FolderNode.LastIndexOf("\"c) + 1))
                        With SubFolderNode
                            .Tag = FolderNode
                            .Nodes.Add("{child}")
                            treeViewFolders.ImageList = ImgIcons
                            .ImageIndex = 0
                            .SelectedImageIndex = 1
                        End With

                    End If
                Next
            Catch e As IOException
                Logger(psUCRTLogFile, "IO Error in AddAllFolders - " & e.ToString)
            Catch ex As Exception
                Logger(psUCRTLogFile, "Error in AddAllFolders - " & ex.ToString)
            End Try

        End If

    End Sub

    Private Sub cboReportType_SelectedIndexChanged(sender As Object, e As EventArgs) Handles cboReportType.SelectedIndexChanged

        Select Case cboReportType.Text
            Case "All"
                grdCaseInfo.Columns("CaseGuid").Visible = False
                grdCaseInfo.Columns("CollectionStatus").Visible = True
                grdCaseInfo.Columns("ReportLoadDuration").Visible = True
                grdCaseInfo.Columns("CaseName").Visible = True
                grdCaseInfo.Columns("CurrentCaseVersion").Visible = True
                grdCaseInfo.Columns("UpgradedCaseVersion").Visible = True
                grdCaseInfo.Columns("BatchLoadInfo").Visible = True
                grdCaseInfo.Columns("DataExport").Visible = True
                grdCaseInfo.Columns("CaseLocation").Visible = True
                grdCaseInfo.Columns("CaseSizeOnDisk").Visible = True
                grdCaseInfo.Columns("CaseFileSize").Visible = True
                grdCaseInfo.Columns("CaseAuditSize").Visible = True
                grdCaseInfo.Columns("OldestTopLevel").Visible = True
                grdCaseInfo.Columns("NewestTopLevel").Visible = True
                grdCaseInfo.Columns("IsCompound").Visible = True
                grdCaseInfo.Columns("CasesContained").Visible = True
                grdCaseInfo.Columns("ContainedInCase").Visible = True
                grdCaseInfo.Columns("Investigator").Visible = True
                grdCaseInfo.Columns("InvestigatorSessions").Visible = True
                grdCaseInfo.Columns("InvalidSessions").Visible = True
                grdCaseInfo.Columns("InvestigatorTimeSummary").Visible = True
                grdCaseInfo.Columns("BrokerMemory").Visible = True
                grdCaseInfo.Columns("WorkerCount").Visible = True
                grdCaseInfo.Columns("WorkerMemory").Visible = True
                grdCaseInfo.Columns("EvidenceName").Visible = True
                grdCaseInfo.Columns("EvidenceLocation").Visible = True
                grdCaseInfo.Columns("MimeTypes").Visible = True
                grdCaseInfo.Columns("ItemTypes").Visible = True
                grdCaseInfo.Columns("IrregularItems").Visible = True
                grdCaseInfo.Columns("CreationDate").Visible = True
                grdCaseInfo.Columns("ModifiedDate").Visible = True
                grdCaseInfo.Columns("LoadStartDate").Visible = False
                grdCaseInfo.Columns("LoadEndDate").Visible = False
                grdCaseInfo.Columns("LoadTime").Visible = False
                grdCaseInfo.Columns("LoadEvents").Visible = True
                grdCaseInfo.Columns("TotalLoadTime").Visible = True
                grdCaseInfo.Columns("ProcessingSpeed").Visible = True
                grdCaseInfo.Columns("Custodians").Visible = True
                grdCaseInfo.Columns("CustodianCount").Visible = True
                grdCaseInfo.Columns("SearchTerm").Visible = True
                grdCaseInfo.Columns("SearchSize").Visible = True
                grdCaseInfo.Columns("SearchHitCount").Visible = True
                grdCaseInfo.Columns("CustodianSearchHit").Visible = True
                grdCaseInfo.Columns("TotalCaseItemCount").Visible = True
                grdCaseInfo.Columns("HitCountPercent").Visible = True
                lblCalculateProcessingSpeeds.Show()
                cboCalculateProcessingSpeeds.Show()
                lblSearchTerm.Show()
                txtSearchTerm.Show()
                grpSearchTerm.Show()
                btnSearchTermFile.Show()

                chkExportSearchResults.Show()
                lblExportLocation.Show()
                txtExportLocation.Show()
                btnExportLocation.Show()
                chkExportSearchResults.Enabled = False
                cboExportType.Enabled = False
                chkExportSearchResults.Checked = False
                cboExportType.Text = vbNullString
                cboExportType.Enabled = False
                'Me.Width = 1550

            Case "App Memory per case"
                grdCaseInfo.Columns("CaseGuid").Visible = False
                grdCaseInfo.Columns("CollectionStatus").Visible = True
                grdCaseInfo.Columns("ReportLoadDuration").Visible = True
                grdCaseInfo.Columns("CaseName").Visible = True
                grdCaseInfo.Columns("CurrentCaseVersion").Visible = True
                grdCaseInfo.Columns("UpgradedCaseVersion").Visible = True
                grdCaseInfo.Columns("BatchLoadInfo").Visible = True
                grdCaseInfo.Columns("DataExport").Visible = True
                grdCaseInfo.Columns("CaseLocation").Visible = True
                grdCaseInfo.Columns("CaseSizeOnDisk").Visible = True
                grdCaseInfo.Columns("CaseFileSize").Visible = True
                grdCaseInfo.Columns("CaseAuditSize").Visible = True
                grdCaseInfo.Columns("OldestTopLevel").Visible = True
                grdCaseInfo.Columns("NewestTopLevel").Visible = True
                grdCaseInfo.Columns("IsCompound").Visible = True
                grdCaseInfo.Columns("CasesContained").Visible = True
                grdCaseInfo.Columns("ContainedInCase").Visible = True
                grdCaseInfo.Columns("Investigator").Visible = True
                grdCaseInfo.Columns("InvestigatorSessions").Visible = True
                grdCaseInfo.Columns("InvalidSessions").Visible = True
                grdCaseInfo.Columns("InvestigatorTimeSummary").Visible = True
                grdCaseInfo.Columns("BrokerMemory").Visible = True
                grdCaseInfo.Columns("WorkerCount").Visible = True
                grdCaseInfo.Columns("WorkerMemory").Visible = True
                grdCaseInfo.Columns("EvidenceName").Visible = False
                grdCaseInfo.Columns("EvidenceLocation").Visible = False
                grdCaseInfo.Columns("MimeTypes").Visible = False
                grdCaseInfo.Columns("ItemTypes").Visible = False
                grdCaseInfo.Columns("IrregularItems").Visible = False
                grdCaseInfo.Columns("CreationDate").Visible = False
                grdCaseInfo.Columns("ModifiedDate").Visible = False
                grdCaseInfo.Columns("LoadStartDate").Visible = False
                grdCaseInfo.Columns("LoadEndDate").Visible = False
                grdCaseInfo.Columns("LoadTime").Visible = False
                grdCaseInfo.Columns("LoadEvents").Visible = False
                grdCaseInfo.Columns("TotalLoadTime").Visible = False
                grdCaseInfo.Columns("ProcessingSpeed").Visible = False
                grdCaseInfo.Columns("Custodians").Visible = False
                grdCaseInfo.Columns("CustodianCount").Visible = False
                grdCaseInfo.Columns("SearchTerm").Visible = False
                grdCaseInfo.Columns("SearchSize").Visible = False
                grdCaseInfo.Columns("SearchHitCount").Visible = False
                grdCaseInfo.Columns("CustodianSearchHit").Visible = False
                grdCaseInfo.Columns("TotalCaseItemCount").Visible = False
                grdCaseInfo.Columns("HitCountPercent").Visible = False
                lblSearchTerm.Hide()
                txtSearchTerm.Hide()
                grpSearchTerm.Hide()
                btnSearchTermFile.Hide()
                chkExportSearchResults.Hide()
                lblExportLocation.Hide()
                txtExportLocation.Hide()
                btnExportLocation.Hide()

                lblCalculateProcessingSpeeds.Hide()
                cboCalculateProcessingSpeeds.Hide()
                chkExportSearchResults.Enabled = False
                chkExportSearchResults.Checked = False
                cboExportType.Text = vbNullString
                cboExportType.Enabled = False

                'Me.Width = 1245
            Case "Case by Investigator"
                grdCaseInfo.Columns("CaseGuid").Visible = False
                grdCaseInfo.Columns("CollectionStatus").Visible = True
                grdCaseInfo.Columns("ReportLoadDuration").Visible = True
                grdCaseInfo.Columns("CaseName").Visible = True
                grdCaseInfo.Columns("CurrentCaseVersion").Visible = True
                grdCaseInfo.Columns("UpgradedCaseVersion").Visible = True
                grdCaseInfo.Columns("BatchLoadInfo").Visible = True
                grdCaseInfo.Columns("DataExport").Visible = True
                grdCaseInfo.Columns("CaseLocation").Visible = True
                grdCaseInfo.Columns("CaseSizeOnDisk").Visible = True
                grdCaseInfo.Columns("CaseFileSize").Visible = True
                grdCaseInfo.Columns("CaseAuditSize").Visible = True
                grdCaseInfo.Columns("OldestTopLevel").Visible = True
                grdCaseInfo.Columns("NewestTopLevel").Visible = True
                grdCaseInfo.Columns("IsCompound").Visible = True
                grdCaseInfo.Columns("CasesContained").Visible = True
                grdCaseInfo.Columns("ContainedInCase").Visible = True
                grdCaseInfo.Columns("Investigator").Visible = True
                grdCaseInfo.Columns("InvestigatorSessions").Visible = True
                grdCaseInfo.Columns("InvalidSessions").Visible = True
                grdCaseInfo.Columns("InvestigatorTimeSummary").Visible = True
                grdCaseInfo.Columns("BrokerMemory").Visible = False
                grdCaseInfo.Columns("WorkerCount").Visible = False
                grdCaseInfo.Columns("WorkerMemory").Visible = False
                grdCaseInfo.Columns("EvidenceName").Visible = False
                grdCaseInfo.Columns("EvidenceLocation").Visible = False
                grdCaseInfo.Columns("MimeTypes").Visible = False
                grdCaseInfo.Columns("ItemTypes").Visible = False
                grdCaseInfo.Columns("IrregularItems").Visible = False
                grdCaseInfo.Columns("CreationDate").Visible = False
                grdCaseInfo.Columns("ModifiedDate").Visible = False
                grdCaseInfo.Columns("LoadStartDate").Visible = False
                grdCaseInfo.Columns("LoadEndDate").Visible = False
                grdCaseInfo.Columns("LoadTime").Visible = False
                grdCaseInfo.Columns("LoadEvents").Visible = False
                grdCaseInfo.Columns("TotalLoadTime").Visible = False
                grdCaseInfo.Columns("ProcessingSpeed").Visible = False
                grdCaseInfo.Columns("Custodians").Visible = False
                grdCaseInfo.Columns("CustodianCount").Visible = False
                grdCaseInfo.Columns("SearchTerm").Visible = False
                grdCaseInfo.Columns("SearchSize").Visible = False
                grdCaseInfo.Columns("SearchHitCount").Visible = False
                grdCaseInfo.Columns("CustodianSearchHit").Visible = False
                grdCaseInfo.Columns("TotalCaseItemCount").Visible = False
                grdCaseInfo.Columns("HitCountPercent").Visible = False
                lblSearchTerm.Hide()
                txtSearchTerm.Hide()
                grpSearchTerm.Hide()
                btnSearchTermFile.Hide()
                chkExportSearchResults.Hide()
                lblExportLocation.Hide()
                txtExportLocation.Hide()
                btnExportLocation.Hide()

                chkExportSearchResults.Enabled = False
                chkExportSearchResults.Checked = False
                cboExportType.Text = vbNullString
                cboExportType.Enabled = False

                'Me.Width = 950

            Case "Case Evidence"
                grdCaseInfo.Columns("CaseGuid").Visible = False
                grdCaseInfo.Columns("CollectionStatus").Visible = True
                grdCaseInfo.Columns("ReportLoadDuration").Visible = True
                grdCaseInfo.Columns("CaseName").Visible = True
                grdCaseInfo.Columns("CurrentCaseVersion").Visible = True
                grdCaseInfo.Columns("UpgradedCaseVersion").Visible = True
                grdCaseInfo.Columns("BatchLoadInfo").Visible = True
                grdCaseInfo.Columns("DataExport").Visible = True
                grdCaseInfo.Columns("CaseLocation").Visible = True
                grdCaseInfo.Columns("CaseSizeOnDisk").Visible = True
                grdCaseInfo.Columns("CaseFileSize").Visible = True
                grdCaseInfo.Columns("CaseAuditSize").Visible = True
                grdCaseInfo.Columns("OldestTopLevel").Visible = True
                grdCaseInfo.Columns("NewestTopLevel").Visible = True
                grdCaseInfo.Columns("IsCompound").Visible = True
                grdCaseInfo.Columns("CasesContained").Visible = True
                grdCaseInfo.Columns("ContainedInCase").Visible = True
                grdCaseInfo.Columns("Investigator").Visible = False
                grdCaseInfo.Columns("InvestigatorSessions").Visible = False
                grdCaseInfo.Columns("InvalidSessions").Visible = False
                grdCaseInfo.Columns("InvestigatorTimeSummary").Visible = False
                grdCaseInfo.Columns("BrokerMemory").Visible = False
                grdCaseInfo.Columns("WorkerCount").Visible = False
                grdCaseInfo.Columns("WorkerMemory").Visible = False
                grdCaseInfo.Columns("EvidenceName").Visible = True
                grdCaseInfo.Columns("EvidenceLocation").Visible = True
                grdCaseInfo.Columns("MimeTypes").Visible = False
                grdCaseInfo.Columns("ItemTypes").Visible = False
                grdCaseInfo.Columns("IrregularItems").Visible = False
                grdCaseInfo.Columns("CreationDate").Visible = False
                grdCaseInfo.Columns("ModifiedDate").Visible = False
                grdCaseInfo.Columns("LoadStartDate").Visible = False
                grdCaseInfo.Columns("LoadEndDate").Visible = False
                grdCaseInfo.Columns("LoadTime").Visible = False
                grdCaseInfo.Columns("LoadEvents").Visible = False
                grdCaseInfo.Columns("TotalLoadTime").Visible = False
                grdCaseInfo.Columns("ProcessingSpeed").Visible = False
                grdCaseInfo.Columns("Custodians").Visible = False
                grdCaseInfo.Columns("CustodianCount").Visible = False
                grdCaseInfo.Columns("SearchTerm").Visible = False
                grdCaseInfo.Columns("SearchSize").Visible = False
                grdCaseInfo.Columns("SearchHitCount").Visible = False
                grdCaseInfo.Columns("CustodianSearchHit").Visible = False
                grdCaseInfo.Columns("TotalCaseItemCount").Visible = False
                grdCaseInfo.Columns("HitCountPercent").Visible = False
                lblSearchTerm.Hide()
                txtSearchTerm.Hide()
                grpSearchTerm.Hide()
                btnSearchTermFile.Hide()
                chkExportSearchResults.Hide()
                lblExportLocation.Hide()
                txtExportLocation.Hide()
                btnExportLocation.Hide()


                lblCalculateProcessingSpeeds.Hide()
                cboCalculateProcessingSpeeds.Hide()
                chkExportSearchResults.Enabled = False
                chkExportSearchResults.Checked = False
                cboExportType.Text = vbNullString
                cboExportType.Enabled = False

            Case "Case Location"
                grdCaseInfo.Columns("CaseGuid").Visible = False
                grdCaseInfo.Columns("CollectionStatus").Visible = True
                grdCaseInfo.Columns("ReportLoadDuration").Visible = True
                grdCaseInfo.Columns("CaseName").Visible = True
                grdCaseInfo.Columns("CurrentCaseVersion").Visible = True
                grdCaseInfo.Columns("UpgradedCaseVersion").Visible = True
                grdCaseInfo.Columns("BatchLoadInfo").Visible = True
                grdCaseInfo.Columns("DataExport").Visible = True
                grdCaseInfo.Columns("CaseLocation").Visible = True
                grdCaseInfo.Columns("CaseSizeOnDisk").Visible = True
                grdCaseInfo.Columns("CaseFileSize").Visible = True
                grdCaseInfo.Columns("CaseAuditSize").Visible = True
                grdCaseInfo.Columns("OldestTopLevel").Visible = True
                grdCaseInfo.Columns("NewestTopLevel").Visible = True
                grdCaseInfo.Columns("IsCompound").Visible = True
                grdCaseInfo.Columns("CasesContained").Visible = True
                grdCaseInfo.Columns("ContainedInCase").Visible = True
                grdCaseInfo.Columns("Investigator").Visible = False
                grdCaseInfo.Columns("InvestigatorSessions").Visible = False
                grdCaseInfo.Columns("InvalidSessions").Visible = False
                grdCaseInfo.Columns("InvestigatorTimeSummary").Visible = False
                grdCaseInfo.Columns("BrokerMemory").Visible = False
                grdCaseInfo.Columns("WorkerCount").Visible = False
                grdCaseInfo.Columns("WorkerMemory").Visible = False
                grdCaseInfo.Columns("EvidenceName").Visible = False
                grdCaseInfo.Columns("EvidenceLocation").Visible = False
                grdCaseInfo.Columns("MimeTypes").Visible = False
                grdCaseInfo.Columns("ItemTypes").Visible = False
                grdCaseInfo.Columns("IrregularItems").Visible = False
                grdCaseInfo.Columns("CreationDate").Visible = False
                grdCaseInfo.Columns("ModifiedDate").Visible = False
                grdCaseInfo.Columns("LoadStartDate").Visible = False
                grdCaseInfo.Columns("LoadEndDate").Visible = False
                grdCaseInfo.Columns("LoadTime").Visible = False
                grdCaseInfo.Columns("LoadEvents").Visible = False
                grdCaseInfo.Columns("TotalLoadTime").Visible = False
                grdCaseInfo.Columns("ProcessingSpeed").Visible = False
                grdCaseInfo.Columns("Custodians").Visible = False
                grdCaseInfo.Columns("CustodianCount").Visible = False
                grdCaseInfo.Columns("SearchTerm").Visible = False
                grdCaseInfo.Columns("SearchSize").Visible = False
                grdCaseInfo.Columns("SearchHitCount").Visible = False
                grdCaseInfo.Columns("CustodianSearchHit").Visible = False
                grdCaseInfo.Columns("TotalCaseItemCount").Visible = False
                grdCaseInfo.Columns("HitCountPercent").Visible = False
                lblSearchTerm.Hide()
                txtSearchTerm.Hide()
                grpSearchTerm.Hide()
                btnSearchTermFile.Hide()
                chkExportSearchResults.Hide()
                lblExportLocation.Hide()
                txtExportLocation.Hide()
                btnExportLocation.Hide()


                lblCalculateProcessingSpeeds.Hide()
                cboCalculateProcessingSpeeds.Hide()
                chkExportSearchResults.Enabled = False
                chkExportSearchResults.Checked = False
                cboExportType.Text = vbNullString
                cboExportType.Enabled = False

                'Me.Width = 850
            Case "Case Size"
                grdCaseInfo.Columns("CaseGuid").Visible = False
                grdCaseInfo.Columns("CollectionStatus").Visible = True
                grdCaseInfo.Columns("ReportLoadDuration").Visible = True
                grdCaseInfo.Columns("CaseName").Visible = True
                grdCaseInfo.Columns("CurrentCaseVersion").Visible = True
                grdCaseInfo.Columns("UpgradedCaseVersion").Visible = True
                grdCaseInfo.Columns("BatchLoadInfo").Visible = True
                grdCaseInfo.Columns("DataExport").Visible = True
                grdCaseInfo.Columns("CaseLocation").Visible = True
                grdCaseInfo.Columns("CaseSizeOnDisk").Visible = True
                grdCaseInfo.Columns("CaseFileSize").Visible = True
                grdCaseInfo.Columns("CaseAuditSize").Visible = True
                grdCaseInfo.Columns("OldestTopLevel").Visible = True
                grdCaseInfo.Columns("NewestTopLevel").Visible = True
                grdCaseInfo.Columns("IsCompound").Visible = True
                grdCaseInfo.Columns("CasesContained").Visible = True
                grdCaseInfo.Columns("ContainedInCase").Visible = True
                grdCaseInfo.Columns("Investigator").Visible = False
                grdCaseInfo.Columns("InvestigatorSessions").Visible = False
                grdCaseInfo.Columns("InvalidSessions").Visible = False
                grdCaseInfo.Columns("InvestigatorTimeSummary").Visible = False
                grdCaseInfo.Columns("BrokerMemory").Visible = False
                grdCaseInfo.Columns("WorkerCount").Visible = False
                grdCaseInfo.Columns("WorkerMemory").Visible = False
                grdCaseInfo.Columns("EvidenceName").Visible = False
                grdCaseInfo.Columns("EvidenceLocation").Visible = False
                grdCaseInfo.Columns("MimeTypes").Visible = False
                grdCaseInfo.Columns("ItemTypes").Visible = False
                grdCaseInfo.Columns("IrregularItems").Visible = False
                grdCaseInfo.Columns("CreationDate").Visible = False
                grdCaseInfo.Columns("ModifiedDate").Visible = False
                grdCaseInfo.Columns("LoadStartDate").Visible = False
                grdCaseInfo.Columns("LoadEndDate").Visible = False
                grdCaseInfo.Columns("LoadTime").Visible = False
                grdCaseInfo.Columns("LoadEvents").Visible = True
                grdCaseInfo.Columns("TotalLoadTime").Visible = True
                grdCaseInfo.Columns("ProcessingSpeed").Visible = True
                grdCaseInfo.Columns("Custodians").Visible = False
                grdCaseInfo.Columns("CustodianCount").Visible = False
                grdCaseInfo.Columns("SearchTerm").Visible = False
                grdCaseInfo.Columns("SearchSize").Visible = False
                grdCaseInfo.Columns("SearchHitCount").Visible = False
                grdCaseInfo.Columns("CustodianSearchHit").Visible = False
                grdCaseInfo.Columns("TotalCaseItemCount").Visible = False
                grdCaseInfo.Columns("HitCountPercent").Visible = False
                lblSearchTerm.Hide()
                txtSearchTerm.Hide()
                grpSearchTerm.Hide()
                btnSearchTermFile.Hide()
                chkExportSearchResults.Hide()
                lblExportLocation.Hide()
                txtExportLocation.Hide()
                btnExportLocation.Hide()


                lblCalculateProcessingSpeeds.Show()
                cboCalculateProcessingSpeeds.Show()
                chkExportSearchResults.Enabled = False
                chkExportSearchResults.Checked = False
                cboExportType.Text = vbNullString
                cboExportType.Enabled = False

                'Me.Width = 1340
            Case "Custodians in Case"
                grdCaseInfo.Columns("CaseGuid").Visible = False
                grdCaseInfo.Columns("CollectionStatus").Visible = True
                grdCaseInfo.Columns("ReportLoadDuration").Visible = True
                grdCaseInfo.Columns("CaseName").Visible = True
                grdCaseInfo.Columns("CurrentCaseVersion").Visible = True
                grdCaseInfo.Columns("UpgradedCaseVersion").Visible = True
                grdCaseInfo.Columns("BatchLoadInfo").Visible = True
                grdCaseInfo.Columns("DataExport").Visible = True
                grdCaseInfo.Columns("CaseLocation").Visible = True
                grdCaseInfo.Columns("CaseSizeOnDisk").Visible = True
                grdCaseInfo.Columns("CaseFileSize").Visible = True
                grdCaseInfo.Columns("CaseAuditSize").Visible = True
                grdCaseInfo.Columns("OldestTopLevel").Visible = True
                grdCaseInfo.Columns("NewestTopLevel").Visible = True
                grdCaseInfo.Columns("IsCompound").Visible = True
                grdCaseInfo.Columns("CasesContained").Visible = True
                grdCaseInfo.Columns("ContainedInCase").Visible = True
                grdCaseInfo.Columns("Investigator").Visible = True
                grdCaseInfo.Columns("InvestigatorSessions").Visible = True
                grdCaseInfo.Columns("InvalidSessions").Visible = True
                grdCaseInfo.Columns("InvestigatorTimeSummary").Visible = True
                grdCaseInfo.Columns("BrokerMemory").Visible = False
                grdCaseInfo.Columns("WorkerCount").Visible = False
                grdCaseInfo.Columns("WorkerMemory").Visible = False
                grdCaseInfo.Columns("EvidenceName").Visible = False
                grdCaseInfo.Columns("EvidenceLocation").Visible = False
                grdCaseInfo.Columns("MimeTypes").Visible = False
                grdCaseInfo.Columns("ItemTypes").Visible = False
                grdCaseInfo.Columns("IrregularItems").Visible = False
                grdCaseInfo.Columns("CreationDate").Visible = False
                grdCaseInfo.Columns("ModifiedDate").Visible = False
                grdCaseInfo.Columns("LoadStartDate").Visible = False
                grdCaseInfo.Columns("LoadEndDate").Visible = False
                grdCaseInfo.Columns("LoadTime").Visible = False
                grdCaseInfo.Columns("LoadEvents").Visible = False
                grdCaseInfo.Columns("TotalLoadTime").Visible = False
                grdCaseInfo.Columns("ProcessingSpeed").Visible = False
                grdCaseInfo.Columns("Custodians").Visible = True
                grdCaseInfo.Columns("CustodianCount").Visible = True
                grdCaseInfo.Columns("SearchTerm").Visible = False
                grdCaseInfo.Columns("SearchSize").Visible = False
                grdCaseInfo.Columns("SearchHitCount").Visible = False
                grdCaseInfo.Columns("CustodianSearchHit").Visible = False
                grdCaseInfo.Columns("TotalCaseItemCount").Visible = False
                grdCaseInfo.Columns("HitCountPercent").Visible = False
                lblSearchTerm.Hide()
                txtSearchTerm.Hide()
                grpSearchTerm.Hide()
                btnSearchTermFile.Hide()
                chkExportSearchResults.Hide()
                lblExportLocation.Hide()
                txtExportLocation.Hide()
                btnExportLocation.Hide()


                lblCalculateProcessingSpeeds.Hide()
                cboCalculateProcessingSpeeds.Hide()
                chkExportSearchResults.Enabled = False
                chkExportSearchResults.Checked = False
                cboExportType.Text = vbNullString
                cboExportType.Enabled = False

                'Me.Width = 1150
            Case "Metadata type"
                grdCaseInfo.Columns("CaseGuid").Visible = False
                grdCaseInfo.Columns("CollectionStatus").Visible = True
                grdCaseInfo.Columns("ReportLoadDuration").Visible = True
                grdCaseInfo.Columns("CaseName").Visible = True
                grdCaseInfo.Columns("CurrentCaseVersion").Visible = True
                grdCaseInfo.Columns("UpgradedCaseVersion").Visible = True
                grdCaseInfo.Columns("BatchLoadInfo").Visible = True
                grdCaseInfo.Columns("DataExport").Visible = True
                grdCaseInfo.Columns("CaseLocation").Visible = True
                grdCaseInfo.Columns("CaseSizeOnDisk").Visible = True
                grdCaseInfo.Columns("CaseFileSize").Visible = True
                grdCaseInfo.Columns("CaseAuditSize").Visible = True
                grdCaseInfo.Columns("OldestTopLevel").Visible = True
                grdCaseInfo.Columns("NewestTopLevel").Visible = True
                grdCaseInfo.Columns("IsCompound").Visible = True
                grdCaseInfo.Columns("CasesContained").Visible = True
                grdCaseInfo.Columns("ContainedInCase").Visible = True
                grdCaseInfo.Columns("Investigator").Visible = False
                grdCaseInfo.Columns("InvestigatorSessions").Visible = False
                grdCaseInfo.Columns("InvalidSessions").Visible = False
                grdCaseInfo.Columns("InvestigatorTimeSummary").Visible = False
                grdCaseInfo.Columns("BrokerMemory").Visible = False
                grdCaseInfo.Columns("WorkerCount").Visible = False
                grdCaseInfo.Columns("WorkerMemory").Visible = False
                grdCaseInfo.Columns("EvidenceName").Visible = False
                grdCaseInfo.Columns("EvidenceLocation").Visible = False
                grdCaseInfo.Columns("MimeTypes").Visible = True
                grdCaseInfo.Columns("ItemTypes").Visible = True
                grdCaseInfo.Columns("IrregularItems").Visible = True
                grdCaseInfo.Columns("CreationDate").Visible = False
                grdCaseInfo.Columns("ModifiedDate").Visible = False
                grdCaseInfo.Columns("LoadStartDate").Visible = False
                grdCaseInfo.Columns("LoadEndDate").Visible = False
                grdCaseInfo.Columns("LoadTime").Visible = False
                grdCaseInfo.Columns("LoadEvents").Visible = False
                grdCaseInfo.Columns("TotalLoadTime").Visible = False
                grdCaseInfo.Columns("ProcessingSpeed").Visible = False
                grdCaseInfo.Columns("Custodians").Visible = False
                grdCaseInfo.Columns("CustodianCount").Visible = False
                grdCaseInfo.Columns("SearchTerm").Visible = False
                grdCaseInfo.Columns("SearchSize").Visible = False
                grdCaseInfo.Columns("SearchHitCount").Visible = False
                grdCaseInfo.Columns("CustodianSearchHit").Visible = False
                grdCaseInfo.Columns("TotalCaseItemCount").Visible = False
                grdCaseInfo.Columns("HitCountPercent").Visible = False
                lblSearchTerm.Hide()
                txtSearchTerm.Hide()
                grpSearchTerm.Hide()
                btnSearchTermFile.Hide()
                chkExportSearchResults.Hide()
                lblExportLocation.Hide()
                txtExportLocation.Hide()
                btnExportLocation.Hide()


                lblCalculateProcessingSpeeds.Hide()
                cboCalculateProcessingSpeeds.Hide()
                chkExportSearchResults.Enabled = False
                chkExportSearchResults.Checked = False
                cboExportType.Text = vbNullString
                cboExportType.Enabled = False

                'Me.Width = 1150
            Case "Processing time"
                grdCaseInfo.Columns("CaseGuid").Visible = False
                grdCaseInfo.Columns("CollectionStatus").Visible = True
                grdCaseInfo.Columns("ReportLoadDuration").Visible = True
                grdCaseInfo.Columns("CaseName").Visible = True
                grdCaseInfo.Columns("CurrentCaseVersion").Visible = True
                grdCaseInfo.Columns("UpgradedCaseVersion").Visible = True
                grdCaseInfo.Columns("BatchLoadInfo").Visible = True
                grdCaseInfo.Columns("DataExport").Visible = True
                grdCaseInfo.Columns("CaseLocation").Visible = True
                grdCaseInfo.Columns("CaseSizeOnDisk").Visible = True
                grdCaseInfo.Columns("CaseFileSize").Visible = True
                grdCaseInfo.Columns("CaseAuditSize").Visible = True
                grdCaseInfo.Columns("OldestTopLevel").Visible = True
                grdCaseInfo.Columns("NewestTopLevel").Visible = True
                grdCaseInfo.Columns("IsCompound").Visible = True
                grdCaseInfo.Columns("CasesContained").Visible = True
                grdCaseInfo.Columns("ContainedInCase").Visible = True
                grdCaseInfo.Columns("Investigator").Visible = False
                grdCaseInfo.Columns("InvestigatorSessions").Visible = False
                grdCaseInfo.Columns("InvalidSessions").Visible = False
                grdCaseInfo.Columns("InvestigatorTimeSummary").Visible = False
                grdCaseInfo.Columns("BrokerMemory").Visible = False
                grdCaseInfo.Columns("WorkerCount").Visible = False
                grdCaseInfo.Columns("WorkerMemory").Visible = False
                grdCaseInfo.Columns("EvidenceName").Visible = False
                grdCaseInfo.Columns("EvidenceLocation").Visible = False
                grdCaseInfo.Columns("MimeTypes").Visible = False
                grdCaseInfo.Columns("ItemTypes").Visible = False
                grdCaseInfo.Columns("IrregularItems").Visible = False
                grdCaseInfo.Columns("CreationDate").Visible = False
                grdCaseInfo.Columns("ModifiedDate").Visible = False
                grdCaseInfo.Columns("LoadStartDate").Visible = False
                grdCaseInfo.Columns("LoadEndDate").Visible = False
                grdCaseInfo.Columns("LoadTime").Visible = False
                grdCaseInfo.Columns("LoadEvents").Visible = True
                grdCaseInfo.Columns("TotalLoadTime").Visible = True
                grdCaseInfo.Columns("ProcessingSpeed").Visible = True
                grdCaseInfo.Columns("Custodians").Visible = False
                grdCaseInfo.Columns("CustodianCount").Visible = False
                grdCaseInfo.Columns("SearchTerm").Visible = False
                grdCaseInfo.Columns("SearchSize").Visible = False
                grdCaseInfo.Columns("SearchHitCount").Visible = False
                grdCaseInfo.Columns("CustodianSearchHit").Visible = False
                grdCaseInfo.Columns("TotalCaseItemCount").Visible = False
                grdCaseInfo.Columns("HitCountPercent").Visible = False
                lblSearchTerm.Hide()
                txtSearchTerm.Hide()
                grpSearchTerm.Hide()
                btnSearchTermFile.Hide()
                chkExportSearchResults.Hide()
                lblExportLocation.Hide()
                txtExportLocation.Hide()
                btnExportLocation.Hide()


                lblCalculateProcessingSpeeds.Show()
                cboCalculateProcessingSpeeds.Show()
                chkExportSearchResults.Enabled = False
                chkExportSearchResults.Checked = False
                cboExportType.Text = vbNullString
                cboExportType.Enabled = False

                'Me.Width = 1350
            Case "Processing speed (GB per hour)"
                grdCaseInfo.Columns("CaseGuid").Visible = False
                grdCaseInfo.Columns("CollectionStatus").Visible = True
                grdCaseInfo.Columns("ReportLoadDuration").Visible = True
                grdCaseInfo.Columns("CaseName").Visible = True
                grdCaseInfo.Columns("CurrentCaseVersion").Visible = True
                grdCaseInfo.Columns("UpgradedCaseVersion").Visible = True
                grdCaseInfo.Columns("BatchLoadInfo").Visible = True
                grdCaseInfo.Columns("DataExport").Visible = True
                grdCaseInfo.Columns("CaseLocation").Visible = True
                grdCaseInfo.Columns("CaseSizeOnDisk").Visible = True
                grdCaseInfo.Columns("CaseFileSize").Visible = True
                grdCaseInfo.Columns("CaseAuditSize").Visible = True
                grdCaseInfo.Columns("OldestTopLevel").Visible = True
                grdCaseInfo.Columns("NewestTopLevel").Visible = True
                grdCaseInfo.Columns("IsCompound").Visible = True
                grdCaseInfo.Columns("CasesContained").Visible = True
                grdCaseInfo.Columns("ContainedInCase").Visible = True
                grdCaseInfo.Columns("Investigator").Visible = False
                grdCaseInfo.Columns("InvestigatorSessions").Visible = False
                grdCaseInfo.Columns("InvalidSessions").Visible = False
                grdCaseInfo.Columns("InvestigatorTimeSummary").Visible = False
                grdCaseInfo.Columns("BrokerMemory").Visible = False
                grdCaseInfo.Columns("WorkerCount").Visible = False
                grdCaseInfo.Columns("WorkerMemory").Visible = False
                grdCaseInfo.Columns("EvidenceName").Visible = False
                grdCaseInfo.Columns("EvidenceLocation").Visible = False
                grdCaseInfo.Columns("MimeTypes").Visible = False
                grdCaseInfo.Columns("ItemTypes").Visible = False
                grdCaseInfo.Columns("IrregularItems").Visible = False
                grdCaseInfo.Columns("CreationDate").Visible = False
                grdCaseInfo.Columns("ModifiedDate").Visible = False
                grdCaseInfo.Columns("LoadStartDate").Visible = False
                grdCaseInfo.Columns("LoadEndDate").Visible = False
                grdCaseInfo.Columns("LoadTime").Visible = False
                grdCaseInfo.Columns("LoadEvents").Visible = True
                grdCaseInfo.Columns("TotalLoadTime").Visible = True
                grdCaseInfo.Columns("ProcessingSpeed").Visible = True
                grdCaseInfo.Columns("Custodians").Visible = False
                grdCaseInfo.Columns("CustodianCount").Visible = False
                grdCaseInfo.Columns("SearchTerm").Visible = False
                grdCaseInfo.Columns("SearchSize").Visible = False
                grdCaseInfo.Columns("SearchHitCount").Visible = False
                grdCaseInfo.Columns("CustodianSearchHit").Visible = False
                grdCaseInfo.Columns("TotalCaseItemCount").Visible = False
                grdCaseInfo.Columns("HitCountPercent").Visible = False
                lblSearchTerm.Hide()
                txtSearchTerm.Hide()
                grpSearchTerm.Hide()
                btnSearchTermFile.Hide()
                chkExportSearchResults.Hide()
                lblExportLocation.Hide()
                txtExportLocation.Hide()
                btnExportLocation.Hide()


                lblCalculateProcessingSpeeds.Show()
                cboCalculateProcessingSpeeds.Show()
                chkExportSearchResults.Enabled = False
                chkExportSearchResults.Checked = False
                cboExportType.Text = vbNullString
                cboExportType.Enabled = False

                'Me.Width = 1350
            Case "Search Term Hit"
                grdCaseInfo.Columns("CaseGuid").Visible = False
                grdCaseInfo.Columns("CollectionStatus").Visible = True
                grdCaseInfo.Columns("ReportLoadDuration").Visible = True
                grdCaseInfo.Columns("CaseName").Visible = True
                grdCaseInfo.Columns("CurrentCaseVersion").Visible = True
                grdCaseInfo.Columns("UpgradedCaseVersion").Visible = True
                grdCaseInfo.Columns("CaseLocation").Visible = True
                grdCaseInfo.Columns("BatchLoadInfo").Visible = True
                grdCaseInfo.Columns("DataExport").Visible = True
                grdCaseInfo.Columns("CaseSizeOnDisk").Visible = True
                grdCaseInfo.Columns("CaseFileSize").Visible = True
                grdCaseInfo.Columns("CaseAuditSize").Visible = True
                grdCaseInfo.Columns("OldestTopLevel").Visible = True
                grdCaseInfo.Columns("NewestTopLevel").Visible = True
                grdCaseInfo.Columns("IsCompound").Visible = True
                grdCaseInfo.Columns("CasesContained").Visible = True
                grdCaseInfo.Columns("ContainedInCase").Visible = True
                grdCaseInfo.Columns("Investigator").Visible = False
                grdCaseInfo.Columns("InvestigatorSessions").Visible = False
                grdCaseInfo.Columns("InvalidSessions").Visible = False
                grdCaseInfo.Columns("InvestigatorTimeSummary").Visible = False
                grdCaseInfo.Columns("BrokerMemory").Visible = False
                grdCaseInfo.Columns("WorkerCount").Visible = False
                grdCaseInfo.Columns("WorkerMemory").Visible = False
                grdCaseInfo.Columns("EvidenceName").Visible = False
                grdCaseInfo.Columns("EvidenceLocation").Visible = False
                grdCaseInfo.Columns("MimeTypes").Visible = False
                grdCaseInfo.Columns("ItemTypes").Visible = False
                grdCaseInfo.Columns("IrregularItems").Visible = False
                grdCaseInfo.Columns("CreationDate").Visible = False
                grdCaseInfo.Columns("ModifiedDate").Visible = False
                grdCaseInfo.Columns("LoadStartDate").Visible = False
                grdCaseInfo.Columns("LoadEndDate").Visible = False
                grdCaseInfo.Columns("LoadTime").Visible = False
                grdCaseInfo.Columns("LoadEvents").Visible = False
                grdCaseInfo.Columns("TotalLoadTime").Visible = False
                grdCaseInfo.Columns("ProcessingSpeed").Visible = False
                grdCaseInfo.Columns("Custodians").Visible = False
                grdCaseInfo.Columns("CustodianCount").Visible = False
                grdCaseInfo.Columns("SearchTerm").Visible = True
                grdCaseInfo.Columns("SearchSize").Visible = True
                grdCaseInfo.Columns("SearchHitCount").Visible = True
                grdCaseInfo.Columns("CustodianSearchHit").Visible = True
                grdCaseInfo.Columns("TotalCaseItemCount").Visible = True
                grdCaseInfo.Columns("HitCountPercent").Visible = True
                lblSearchTerm.Show()
                txtSearchTerm.Show()
                grpSearchTerm.Show()
                btnSearchTermFile.Show()

                chkExportSearchResults.Show()
                lblExportLocation.Show()
                txtExportLocation.Show()
                btnExportLocation.Show()
                chkExportSearchResults.Enabled = False
                cboExportType.Enabled = False
                chkExportSearchResults.Checked = False
                cboExportType.Text = vbNullString
                cboExportType.Enabled = False


                lblCalculateProcessingSpeeds.Hide()
                cboCalculateProcessingSpeeds.Hide()
                chkExportSearchResults.Enabled = True
                chkExportSearchResults.Checked = False
                cboExportType.Text = vbNullString
                cboExportType.Enabled = False

            Case "Total Number of Items"
                grdCaseInfo.Columns("CaseGuid").Visible = False
                grdCaseInfo.Columns("CollectionStatus").Visible = True
                grdCaseInfo.Columns("ReportLoadDuration").Visible = True
                grdCaseInfo.Columns("CaseName").Visible = True
                grdCaseInfo.Columns("CurrentCaseVersion").Visible = True
                grdCaseInfo.Columns("UpgradedCaseVersion").Visible = True
                grdCaseInfo.Columns("BatchLoadInfo").Visible = True
                grdCaseInfo.Columns("DataExport").Visible = True
                grdCaseInfo.Columns("CaseLocation").Visible = True
                grdCaseInfo.Columns("CaseSizeOnDisk").Visible = True
                grdCaseInfo.Columns("CaseFileSize").Visible = True
                grdCaseInfo.Columns("CaseAuditSize").Visible = True
                grdCaseInfo.Columns("OldestTopLevel").Visible = True
                grdCaseInfo.Columns("NewestTopLevel").Visible = True
                grdCaseInfo.Columns("IsCompound").Visible = True
                grdCaseInfo.Columns("CasesContained").Visible = True
                grdCaseInfo.Columns("ContainedInCase").Visible = True
                grdCaseInfo.Columns("Investigator").Visible = False
                grdCaseInfo.Columns("InvestigatorSessions").Visible = False
                grdCaseInfo.Columns("InvalidSessions").Visible = False
                grdCaseInfo.Columns("InvestigatorTimeSummary").Visible = False
                grdCaseInfo.Columns("BrokerMemory").Visible = False
                grdCaseInfo.Columns("WorkerCount").Visible = False
                grdCaseInfo.Columns("WorkerMemory").Visible = False
                grdCaseInfo.Columns("EvidenceName").Visible = False
                grdCaseInfo.Columns("EvidenceLocation").Visible = False
                grdCaseInfo.Columns("MimeTypes").Visible = False
                grdCaseInfo.Columns("ItemTypes").Visible = False
                grdCaseInfo.Columns("IrregularItems").Visible = False
                grdCaseInfo.Columns("CreationDate").Visible = False
                grdCaseInfo.Columns("ModifiedDate").Visible = False
                grdCaseInfo.Columns("LoadStartDate").Visible = False
                grdCaseInfo.Columns("LoadEndDate").Visible = False
                grdCaseInfo.Columns("LoadTime").Visible = False
                grdCaseInfo.Columns("LoadEvents").Visible = False
                grdCaseInfo.Columns("TotalLoadTime").Visible = False
                grdCaseInfo.Columns("ProcessingSpeed").Visible = False
                grdCaseInfo.Columns("Custodians").Visible = False
                grdCaseInfo.Columns("CustodianCount").Visible = False
                grdCaseInfo.Columns("SearchTerm").Visible = False
                grdCaseInfo.Columns("SearchSize").Visible = False
                grdCaseInfo.Columns("SearchHitCount").Visible = False
                grdCaseInfo.Columns("CustodianSearchHit").Visible = False
                grdCaseInfo.Columns("TotalCaseItemCount").Visible = True
                grdCaseInfo.Columns("HitCountPercent").Visible = False
                lblSearchTerm.Hide()
                txtSearchTerm.Hide()
                grpSearchTerm.Hide()
                btnSearchTermFile.Hide()
                chkExportSearchResults.Hide()
                lblExportLocation.Hide()
                txtExportLocation.Hide()
                btnExportLocation.Hide()

                lblCalculateProcessingSpeeds.Hide()
                cboCalculateProcessingSpeeds.Hide()
                chkExportSearchResults.Enabled = False
                chkExportSearchResults.Checked = False
                cboExportType.Text = vbNullString
                cboExportType.Enabled = False

                'Me.Width = 1050
            Case "Total Number of workers"
                grdCaseInfo.Columns("CaseGuid").Visible = False
                grdCaseInfo.Columns("CollectionStatus").Visible = True
                grdCaseInfo.Columns("ReportLoadDuration").Visible = True
                grdCaseInfo.Columns("CaseName").Visible = True
                grdCaseInfo.Columns("CurrentCaseVersion").Visible = True
                grdCaseInfo.Columns("UpgradedCaseVersion").Visible = True
                grdCaseInfo.Columns("BatchLoadInfo").Visible = True
                grdCaseInfo.Columns("DataExport").Visible = True
                grdCaseInfo.Columns("CaseLocation").Visible = True
                grdCaseInfo.Columns("CaseSizeOnDisk").Visible = True
                grdCaseInfo.Columns("CaseFileSize").Visible = True
                grdCaseInfo.Columns("CaseAuditSize").Visible = True
                grdCaseInfo.Columns("OldestTopLevel").Visible = True
                grdCaseInfo.Columns("NewestTopLevel").Visible = True
                grdCaseInfo.Columns("IsCompound").Visible = True
                grdCaseInfo.Columns("CasesContained").Visible = True
                grdCaseInfo.Columns("ContainedInCase").Visible = True
                grdCaseInfo.Columns("Investigator").Visible = True
                grdCaseInfo.Columns("InvestigatorSessions").Visible = True
                grdCaseInfo.Columns("InvalidSessions").Visible = True
                grdCaseInfo.Columns("InvestigatorTimeSummary").Visible = True
                grdCaseInfo.Columns("BrokerMemory").Visible = True
                grdCaseInfo.Columns("WorkerCount").Visible = True
                grdCaseInfo.Columns("WorkerMemory").Visible = True
                grdCaseInfo.Columns("EvidenceName").Visible = False
                grdCaseInfo.Columns("EvidenceLocation").Visible = False
                grdCaseInfo.Columns("MimeTypes").Visible = False
                grdCaseInfo.Columns("ItemTypes").Visible = False
                grdCaseInfo.Columns("IrregularItems").Visible = False
                grdCaseInfo.Columns("CreationDate").Visible = False
                grdCaseInfo.Columns("ModifiedDate").Visible = False
                grdCaseInfo.Columns("LoadStartDate").Visible = False
                grdCaseInfo.Columns("LoadEndDate").Visible = False
                grdCaseInfo.Columns("LoadTime").Visible = False
                grdCaseInfo.Columns("LoadEvents").Visible = True
                grdCaseInfo.Columns("TotalLoadTime").Visible = True
                grdCaseInfo.Columns("ProcessingSpeed").Visible = True
                grdCaseInfo.Columns("Custodians").Visible = False
                grdCaseInfo.Columns("CustodianCount").Visible = False
                grdCaseInfo.Columns("SearchTerm").Visible = False
                grdCaseInfo.Columns("SearchSize").Visible = False
                grdCaseInfo.Columns("SearchHitCount").Visible = False
                grdCaseInfo.Columns("CustodianSearchHit").Visible = False
                grdCaseInfo.Columns("TotalCaseItemCount").Visible = False
                grdCaseInfo.Columns("HitCountPercent").Visible = False
                lblSearchTerm.Hide()
                txtSearchTerm.Hide()
                grpSearchTerm.Hide()
                btnSearchTermFile.Hide()
                chkExportSearchResults.Hide()
                lblExportLocation.Hide()
                txtExportLocation.Hide()
                btnExportLocation.Hide()

                lblCalculateProcessingSpeeds.Show()
                cboCalculateProcessingSpeeds.Show()
                chkExportSearchResults.Enabled = False
                chkExportSearchResults.Checked = False
                cboExportType.Text = vbNullString
                cboExportType.Enabled = False

                'Me.Width = 1100
        End Select
    End Sub

    Public Function blnBuildCaseHistoryRubyScript(ByVal sCaseName As String, ByVal sReportPathLocation As String, ByVal bMigrateCase As Boolean) As Boolean
        blnBuildCaseHistoryRubyScript = False
        Dim sCSVFile As String
        Dim CaseDirectory As DirectoryInfo
        Dim CaseFile As FileInfo
        Dim sCaseDirectory As String
        Dim sCaseDirectoryName As String
        Dim sCaseDirectoryPath As String

        Dim CustodianRuby As StreamWriter

        CaseDirectory = New DirectoryInfo(sCaseName)
        CaseFile = New FileInfo(sCaseName)

        sCaseDirectory = CaseFile.Directory.FullName
        sCaseDirectoryPath = Path.GetDirectoryName(sCaseDirectory).ToString

        sCaseDirectoryName = sCaseDirectory.Replace(sCaseDirectoryPath & "\", "")
        sCSVFile = sReportPathLocation.Replace("\", "\\") & "\\" & sCaseDirectoryName & "-history.csv"

        CustodianRuby = New StreamWriter(sReportPathLocation & sCaseDirectoryName & "-history.rb")

        If bMigrateCase = True Then
            CustodianRuby.WriteLine("$current_case = $utilities.getCaseFactory.open(" & """" & sCaseDirectory & """" & ",{" & """" & "migrate" & """" & "=>true}" & ")")
        Else
            CustodianRuby.WriteLine("$current_case = $utilities.getCaseFactory.open(" & """" & sCaseDirectory & """" & ",{" & """" & "migrate" & """" & "=>false}" & ")")

        End If
        CustodianRuby.WriteLine("csv_file = " & """" & sCSVFile & """")
        CustodianRuby.WriteLine("# Annotation Types")
        CustodianRuby.WriteLine("# openSession - occurs at the start of a session with a case (i.e. when the case is opened.)")
        CustodianRuby.WriteLine("# closeSession - occurs at the end of a session with a case (i.e. when the case is closed.)")
        CustodianRuby.WriteLine("# loadData - occurs when data is loaded into the case.")
        CustodianRuby.WriteLine("# search - occurs when a search is performed.")
        CustodianRuby.WriteLine("# annotation - occurs when items are annotated (e.g. tagged.)")
        CustodianRuby.WriteLine("# export - occurs when data or metadata is exported out of the case.")
        CustodianRuby.WriteLine("# import - occurs when data or metadata is imported into the case. The difference from loadData")
        CustodianRuby.WriteLine("#    is that with import, the data is directly imported without processing.")
        CustodianRuby.WriteLine("# delete - occurs when data in the case is deleted.")
        CustodianRuby.WriteLine("# script - occurs when a script is executed.")
        CustodianRuby.WriteLine("# printPreview - occurs when a print preview action is executed.")
        CustodianRuby.WriteLine("# A value of nil means all types")

        CustodianRuby.WriteLine("# https://download.nuix.com/releases/desktop/stable/docs/en/scripting/api/nuix/Case.html#getHistory--")
        CustodianRuby.WriteLine("settings = {")
        CustodianRuby.WriteLine("""" & "    order" & """" & """" & "=>" & """" & "start_date_descending" & """"",")
        CustodianRuby.WriteLine("""" & "    type" & """" & "=> nil,")
        CustodianRuby.WriteLine("}")

        CustodianRuby.WriteLine("history_events = $current_case.getHistory(settings)")
        CustodianRuby.WriteLine("")
        CustodianRuby.WriteLine("last_progress = Time.now")
        CustodianRuby.WriteLine("require """ & "csv" & """")
        CustodianRuby.WriteLine("CSV.open(csv_file," & """" & "w:utf-8" & """" & ") do |csv|")
        CustodianRuby.WriteLine("	csv << [")
        CustodianRuby.WriteLine("		" & """" & "Type" & """" & ",")
        CustodianRuby.WriteLine("		" & """" & "Failed" & """" & ",")
        CustodianRuby.WriteLine("		" & """" & "Succeeded" & """" & ",")
        CustodianRuby.WriteLine("		" & """" & "Start" & """" & ",")
        CustodianRuby.WriteLine("		" & """" & "End" & """" & ",")
        CustodianRuby.WriteLine("		" & """" & "User" & """" & ",")
        CustodianRuby.WriteLine("		" & """" & "Affected Item Count" & """" & ",")
        CustodianRuby.WriteLine("		" & """" & "Details" & """" & ",")
        CustodianRuby.WriteLine("	]")

        CustodianRuby.WriteLine("   index = 0")
        CustodianRuby.WriteLine("	history_events.each do |event|")
        CustodianRuby.WriteLine("      index += 1")
        CustodianRuby.WriteLine("      detail_blob = []")
        CustodianRuby.WriteLine("       event.getDetails.each do |key,value|")
        CustodianRuby.WriteLine("			detail_blob << " & """" & "#{key}: #{value}" & """")
        CustodianRuby.WriteLine("		end")
        CustodianRuby.WriteLine("		detail_blob = detail_blob.join(" & """" & "; " & """" & ")")
        CustodianRuby.WriteLine("")
        CustodianRuby.WriteLine("		csv << [")
        CustodianRuby.WriteLine("			event.getTypeString,")
        CustodianRuby.WriteLine("			event.getFailed,")
        CustodianRuby.WriteLine("			event.getSucceeded,")
        CustodianRuby.WriteLine("			event.getStartDate.toString,")
        CustodianRuby.WriteLine("			event.getEndDate.toString,")
        CustodianRuby.WriteLine("			" & """" & "#{event.getUser.getShortName} : #{event.getUser.getLongName}" & """" & ",")
        CustodianRuby.WriteLine("			event.getAffectedItems.size,")
        CustodianRuby.WriteLine("			detail_blob,")
        CustodianRuby.WriteLine("		]")
        CustodianRuby.WriteLine("")
        CustodianRuby.WriteLine("		if (Time.now - last_progress) > 2")
        CustodianRuby.WriteLine("          puts " & """" & "#{Time.now}: Processed #{index} events" & """")
        CustodianRuby.WriteLine("       end")
        CustodianRuby.WriteLine("    end")
        CustodianRuby.WriteLine("        puts " & """" & "#{Time.now}: Processed #{index} events" & """")
        CustodianRuby.WriteLine("end")
        CustodianRuby.WriteLine("$current_case.close")
        CustodianRuby.WriteLine("puts " & """" & "Completed reporting to #{csv_file}" & """")

        CustodianRuby.Close()
        blnBuildCaseHistoryRubyScript = True

    End Function

    Private Function blnBuildUpdatedAllCaseDataProcessingRuby(ByVal sRubyFile As String, ByVal sCasePath As String, ByVal sReportPathLocation As String, ByVal sUserEnteredSearch As String, ByVal sUserSearchFile As String, ByVal bMigrateCase As Boolean, ByVal bExportSearchResults As Boolean, ByVal sExportDirectory As String, ByVal bExportOnly As Boolean, ByVal sExportWorkers As String, ByVal sExportWorkerMemory As String, ByVal sExportType As String, ByVal sUpgradeCases As String) As Boolean
        Dim AllCaseDataProcessingRuby As StreamWriter

        blnBuildUpdatedAllCaseDataProcessingRuby = False
        AllCaseDataProcessingRuby = New StreamWriter(sRubyFile)

        AllCaseDataProcessingRuby.WriteLine("require 'thread'")
        AllCaseDataProcessingRuby.WriteLine("require 'json'")
        AllCaseDataProcessingRuby.WriteLine("require 'date'")
        AllCaseDataProcessingRuby.WriteLine("require 'csv'")
        AllCaseDataProcessingRuby.WriteLine("require 'fileutils'")
        AllCaseDataProcessingRuby.WriteLine("")
        AllCaseDataProcessingRuby.WriteLine("class TimeSpanFormatter")
        AllCaseDataProcessingRuby.WriteLine("	SECOND = 1")
        AllCaseDataProcessingRuby.WriteLine("	MINUTE = 60 * SECOND")
        AllCaseDataProcessingRuby.WriteLine("	HOUR = 60 * MINUTE")
        AllCaseDataProcessingRuby.WriteLine("	DAY = 24 * HOUR")
        AllCaseDataProcessingRuby.WriteLine("")
        AllCaseDataProcessingRuby.WriteLine("	def self.format_seconds(seconds)")
        AllCaseDataProcessingRuby.WriteLine("		# Make sure were using whole numbers")
        AllCaseDataProcessingRuby.WriteLine("		seconds = seconds.to_i")
        AllCaseDataProcessingRuby.WriteLine("		days = seconds / DAY")
        AllCaseDataProcessingRuby.WriteLine("		seconds -= days * DAY")
        AllCaseDataProcessingRuby.WriteLine("")
        AllCaseDataProcessingRuby.WriteLine("		hours = seconds / HOUR")
        AllCaseDataProcessingRuby.WriteLine("		seconds -= hours * HOUR")
        AllCaseDataProcessingRuby.WriteLine("")
        AllCaseDataProcessingRuby.WriteLine("		minutes = seconds / MINUTE")
        AllCaseDataProcessingRuby.WriteLine("		seconds -= minutes * MINUTE")
        AllCaseDataProcessingRuby.WriteLine("		days_string = " & """" & """")
        AllCaseDataProcessingRuby.WriteLine("		if days > 0")
        AllCaseDataProcessingRuby.WriteLine("			days_string = " & """" & "#{days} Days" & """")
        AllCaseDataProcessingRuby.WriteLine("		end")
        AllCaseDataProcessingRuby.WriteLine("		hours_string = hours.to_s.rjust(2," & """" & "0" & """" & ")")
        AllCaseDataProcessingRuby.WriteLine("		minutes_string = minutes.to_s.rjust(2," & """" & "0" & """" & ")")
        AllCaseDataProcessingRuby.WriteLine("		seconds_string = seconds.to_s.rjust(2," & """" & "0" & """" & ")")
        AllCaseDataProcessingRuby.WriteLine("")
        AllCaseDataProcessingRuby.WriteLine("		if days > 0")
        AllCaseDataProcessingRuby.WriteLine("			return " & """" & "#{days_string} #{hours_string}:#{minutes_string}:#{seconds_string}" & """")
        AllCaseDataProcessingRuby.WriteLine("		else")
        AllCaseDataProcessingRuby.WriteLine("			return " & """" & "#{hours_string}:#{minutes_string}:#{seconds_string}" & """")
        AllCaseDataProcessingRuby.WriteLine("		end")
        AllCaseDataProcessingRuby.WriteLine("	end")
        AllCaseDataProcessingRuby.WriteLine("end")
        AllCaseDataProcessingRuby.WriteLine("")
        AllCaseDataProcessingRuby.WriteLine("")
        AllCaseDataProcessingRuby.WriteLine("")


        AllCaseDataProcessingRuby.WriteLine("begin")
        AllCaseDataProcessingRuby.WriteLine("   load """ & sReportPathLocation.Replace("\", "\\") & "\\Database.rb_""")
        AllCaseDataProcessingRuby.WriteLine("   load """ & sReportPathLocation.Replace("\", "\\") & "\\SQLite.rb_""")
        AllCaseDataProcessingRuby.WriteLine("   db = SQLite.new(""" & sReportPathLocation.Replace("\", "\\") & "\\NuixCaseReports.db3" & """" & ")")
        If sUpgradeCases = "Upgrade Only" Then
            AllCaseDataProcessingRuby.WriteLine("   $current_case = $utilities.getCaseFactory.open(" & """" & sCasePath.Replace("\", "\\") & """" & ",{" & """" & "migrate" & """" & "=>true}" & ")")
            AllCaseDataProcessingRuby.WriteLine("")
            AllCaseDataProcessingRuby.WriteLine("   case_guid = $current_case.getGuid")
            AllCaseDataProcessingRuby.WriteLine("   case_description = $current_case.getDescription")
            AllCaseDataProcessingRuby.WriteLine("")
            AllCaseDataProcessingRuby.WriteLine("   case_guid = case_guid.tr(" & """" & "-" & """" & "," & """" & """" & ")")
            AllCaseDataProcessingRuby.WriteLine("   case_name = $current_case.getName().strip")
            AllCaseDataProcessingRuby.WriteLine("   total_item_count = $current_case.count('*')")
            AllCaseDataProcessingRuby.WriteLine("   total_file_size = $current_case.getStatistics.getFileSize('*')")
            AllCaseDataProcessingRuby.WriteLine("   total_audit_size = $current_case.getStatistics.getAuditSize('*')")

            AllCaseDataProcessingRuby.WriteLine("   compound_case_contains = ''")
            AllCaseDataProcessingRuby.WriteLine("   is_compound = $current_case.isCompound()")
            AllCaseDataProcessingRuby.WriteLine("   if " & """" & "#{is_compound}" & """" & "== 'true'")
            AllCaseDataProcessingRuby.WriteLine("       child_cases = $current_case.getChildCases")
            AllCaseDataProcessingRuby.WriteLine("       child_cases.each do |cases|")
            AllCaseDataProcessingRuby.WriteLine("           compound_case_contains = compound_case_contains + cases.getName() + ';'")
            AllCaseDataProcessingRuby.WriteLine("       end")
            AllCaseDataProcessingRuby.WriteLine("   end")
            AllCaseDataProcessingRuby.WriteLine("   all_case_users = ''")
            AllCaseDataProcessingRuby.WriteLine("   case_users = $current_case.getAllUsers()")
            AllCaseDataProcessingRuby.WriteLine("      case_users.each do |case_user|")
            AllCaseDataProcessingRuby.WriteLine("	    	all_case_users = all_case_users + " & """" & "#{case_user}" & """" & " + ';'")
            AllCaseDataProcessingRuby.WriteLine("       end")
            AllCaseDataProcessingRuby.WriteLine("   case_users_data = [total_item_count, total_file_size, total_audit_size, all_case_users, case_description, is_compound.to_s, compound_case_contains, 5]")
            AllCaseDataProcessingRuby.WriteLine("   db.update(" & """" & "UPDATE NuixReportingInfo SET TotalItemCount = ?, CaseFileSize = ?, CaseAuditSize = ?, CaseUsers = ?, CaseDescription = ?, IsCompound = ?, CasesContained = ?, PercentComplete = ? WHERE CaseGUID = '#{case_guid}'" & """" & ", case_users_data)")

            AllCaseDataProcessingRuby.WriteLine("   all_batch_load_info = ''")
            AllCaseDataProcessingRuby.WriteLine("   batch_loads = $current_case.getBatchLoads")
            AllCaseDataProcessingRuby.WriteLine("	    batch_loads.each do |load|")
            AllCaseDataProcessingRuby.WriteLine("           batch_guid = " & """" & "#{load.getBatchId}" & """")
            AllCaseDataProcessingRuby.WriteLine("           batch_loaddate = " & """" & "#{load.getLoaded}" & """")
            AllCaseDataProcessingRuby.WriteLine("           batch_load_query = " & """" & "batch-load-guid:#{batch_guid}" & """")
            AllCaseDataProcessingRuby.WriteLine("           batch_load_count = $current_case.count(batch_load_query)")
            AllCaseDataProcessingRuby.WriteLine("           batch_load_file_size = $current_case.getStatistics.getFileSize(batch_load_query)")
            AllCaseDataProcessingRuby.WriteLine("           batch_load_audit_size = $current_case.getStatistics.getAuditSize(batch_load_query)")
            AllCaseDataProcessingRuby.WriteLine("		    all_batch_load_info = all_batch_load_info + " & """" & "#{batch_loaddate}" & """" & " + '::' + " & """" & "#{batch_load_count}" & """" & ".to_s + " & "'::' + " & """" & "#{batch_load_file_size}" & """" & ".to_s + " & "'::' + " & """" & "#{batch_load_audit_size}" & """" & ".to_s + " & """" & ";" & """")
            AllCaseDataProcessingRuby.WriteLine("	    end")
            AllCaseDataProcessingRuby.WriteLine("   case_batchload_data = [all_batch_load_info, 10]")
            AllCaseDataProcessingRuby.WriteLine("   db.update(" & """" & "UPDATE NuixReportingInfo SET BatchLoadInfo = ?, PercentComplete = ? WHERE CaseGUID = '#{case_guid}'" & """" & ", case_batchload_data)")
            AllCaseDataProcessingRuby.WriteLine("rescue")
            AllCaseDataProcessingRuby.WriteLine("   puts 'Error in Block 1'")
            AllCaseDataProcessingRuby.WriteLine("ensure")
            AllCaseDataProcessingRuby.WriteLine("")
            AllCaseDataProcessingRuby.WriteLine("end")
            AllCaseDataProcessingRuby.WriteLine("begin")
            AllCaseDataProcessingRuby.WriteLine("   processing_data = [100]")
            AllCaseDataProcessingRuby.WriteLine("   db.update(" & """" & "UPDATE NuixReportingInfo SET PercentComplete = ? WHERE CaseGUID = '#{case_guid}'" & """" & ", processing_data)")

            AllCaseDataProcessingRuby.WriteLine("   $current_case.close")
            AllCaseDataProcessingRuby.WriteLine("rescue")
            AllCaseDataProcessingRuby.WriteLine("   puts 'Error in Block 10'")
            AllCaseDataProcessingRuby.WriteLine("ensure")
            AllCaseDataProcessingRuby.WriteLine("")
            AllCaseDataProcessingRuby.WriteLine("end")
            AllCaseDataProcessingRuby.WriteLine("")

        Else
            If (bMigrateCase = True) Then
                AllCaseDataProcessingRuby.WriteLine("   $current_case = $utilities.getCaseFactory.open(" & """" & sCasePath.Replace("\", "\\") & """" & ",{" & """" & "migrate" & """" & "=>true}" & ")")
            Else
                AllCaseDataProcessingRuby.WriteLine("   $current_case = $utilities.getCaseFactory.open(" & """" & sCasePath.Replace("\", "\\") & """" & ",{" & """" & "migrate" & """" & "=>false}" & ")")
            End If
            AllCaseDataProcessingRuby.WriteLine("")
            AllCaseDataProcessingRuby.WriteLine("   case_guid = $current_case.getGuid")
            AllCaseDataProcessingRuby.WriteLine("   case_description = $current_case.getDescription")
            AllCaseDataProcessingRuby.WriteLine("")
            AllCaseDataProcessingRuby.WriteLine("   case_guid = case_guid.tr(" & """" & "-" & """" & "," & """" & """" & ")")
            AllCaseDataProcessingRuby.WriteLine("   case_name = $current_case.getName().strip")
            AllCaseDataProcessingRuby.WriteLine("   total_item_count = $current_case.count('*')")
            AllCaseDataProcessingRuby.WriteLine("   total_file_size = $current_case.getStatistics.getFileSize('*')")
            AllCaseDataProcessingRuby.WriteLine("   total_audit_size = $current_case.getStatistics.getAuditSize('*')")

            AllCaseDataProcessingRuby.WriteLine("   compound_case_contains = ''")
            AllCaseDataProcessingRuby.WriteLine("   is_compound = $current_case.isCompound()")
            AllCaseDataProcessingRuby.WriteLine("   if " & """" & "#{is_compound}" & """" & "== 'true'")
            AllCaseDataProcessingRuby.WriteLine("       child_cases = $current_case.getChildCases")
            AllCaseDataProcessingRuby.WriteLine("       child_cases.each do |cases|")
            AllCaseDataProcessingRuby.WriteLine("           compound_case_contains = compound_case_contains + cases.getName() + ';'")
            AllCaseDataProcessingRuby.WriteLine("       end")
            AllCaseDataProcessingRuby.WriteLine("   end")
            AllCaseDataProcessingRuby.WriteLine("   all_case_users = ''")
            AllCaseDataProcessingRuby.WriteLine("   case_users = $current_case.getAllUsers()")
            AllCaseDataProcessingRuby.WriteLine("      case_users.each do |case_user|")
            AllCaseDataProcessingRuby.WriteLine("	    	all_case_users = all_case_users + " & """" & "#{case_user}" & """" & " + ';'")
            AllCaseDataProcessingRuby.WriteLine("       end")
            AllCaseDataProcessingRuby.WriteLine("   case_users_data = [total_item_count, total_file_size, total_audit_size, all_case_users, case_description, is_compound.to_s, compound_case_contains, 5]")
            AllCaseDataProcessingRuby.WriteLine("   db.update(" & """" & "UPDATE NuixReportingInfo SET TotalItemCount = ?, CaseFileSize = ?, CaseAuditSize = ?, CaseUsers = ?, CaseDescription = ?, IsCompound = ?, CasesContained = ?, PercentComplete = ? WHERE CaseGUID = '#{case_guid}'" & """" & ", case_users_data)")

            AllCaseDataProcessingRuby.WriteLine("   all_batch_load_info = ''")
            AllCaseDataProcessingRuby.WriteLine("   batch_loads = $current_case.getBatchLoads")
            AllCaseDataProcessingRuby.WriteLine("	    batch_loads.each do |load|")
            AllCaseDataProcessingRuby.WriteLine("           batch_guid = " & """" & "#{load.getBatchId}" & """")
            AllCaseDataProcessingRuby.WriteLine("           batch_loaddate = " & """" & "#{load.getLoaded}" & """")
            AllCaseDataProcessingRuby.WriteLine("           batch_load_query = " & """" & "batch-load-guid:#{batch_guid}" & """")
            AllCaseDataProcessingRuby.WriteLine("           batch_load_count = $current_case.count(batch_load_query)")
            AllCaseDataProcessingRuby.WriteLine("           batch_load_file_size = $current_case.getStatistics.getFileSize(batch_load_query)")
            AllCaseDataProcessingRuby.WriteLine("           batch_load_audit_size = $current_case.getStatistics.getAuditSize(batch_load_query)")
            AllCaseDataProcessingRuby.WriteLine("		    all_batch_load_info = all_batch_load_info + " & """" & "#{batch_loaddate}" & """" & " + '::' + " & """" & "#{batch_load_count}" & """" & ".to_s + " & "'::' + " & """" & "#{batch_load_file_size}" & """" & ".to_s + " & "'::' + " & """" & "#{batch_load_audit_size}" & """" & ".to_s + " & """" & ";" & """")
            AllCaseDataProcessingRuby.WriteLine("	    end")
            AllCaseDataProcessingRuby.WriteLine("   case_batchload_data = [all_batch_load_info, 10]")
            AllCaseDataProcessingRuby.WriteLine("   db.update(" & """" & "UPDATE NuixReportingInfo SET BatchLoadInfo = ?, PercentComplete = ? WHERE CaseGUID = '#{case_guid}'" & """" & ", case_batchload_data)")
            AllCaseDataProcessingRuby.WriteLine("rescue")
            AllCaseDataProcessingRuby.WriteLine("   puts 'Error in Block 1'")
            AllCaseDataProcessingRuby.WriteLine("ensure")
            AllCaseDataProcessingRuby.WriteLine("")
            AllCaseDataProcessingRuby.WriteLine("end")
            If bExportOnly = False Then
                AllCaseDataProcessingRuby.WriteLine("begin")
                AllCaseDataProcessingRuby.WriteLine("   custodian_count = 0")
                AllCaseDataProcessingRuby.WriteLine("   all_custodians_info = ''")
                AllCaseDataProcessingRuby.WriteLine("   custodian_names = $current_case.getAllCustodians")
                AllCaseDataProcessingRuby.WriteLine("       custodian_names.each do |custodian_name|")
                AllCaseDataProcessingRuby.WriteLine("		    custodian_count += 1")
                AllCaseDataProcessingRuby.WriteLine("		    search_criteria = " & """" & "custodian:\" & """" & "#{custodian_name}\" & """""")
                AllCaseDataProcessingRuby.WriteLine("		    custodian_name_count = $current_case.count(" & """" & "#{search_criteria}" & """" & ")")
                AllCaseDataProcessingRuby.WriteLine("		    all_custodians_info = all_custodians_info + " & """" & "#{custodian_name}" & """" & " + '::' + " & """" & "#{custodian_name_count}" & """" & ".to_s + " & """" & ";" & """")
                AllCaseDataProcessingRuby.WriteLine("       end")
                AllCaseDataProcessingRuby.WriteLine("   case_custodians_data = [all_custodians_info, custodian_count, 15]")
                AllCaseDataProcessingRuby.WriteLine("   db.update(" & """" & "UPDATE NuixReportingInfo SET Custodians = ?, CustodianCount = ?, PercentComplete = ? WHERE CaseGUID = '#{case_guid}'" & """" & ", case_custodians_data)")

                AllCaseDataProcessingRuby.WriteLine("   oldest={" & """" & "defaultFields" & """" & "=>" & """" & "item-date" & """" & "," & """" & "order" & """" & "=>" & """" & "item-date ASC" & """" & "," & """" & "limit" & """" & "=>1}")
                AllCaseDataProcessingRuby.WriteLine("   newest={" & """" & "defaultFields" & """" & "=>" & """" & "item-date" & """" & "," & """" & "order" & """" & "=>" & """" & "item-date DESC" & """" & "," & """" & "limit" & """" & "=>1}")
                '        AllCaseDataProcessingRuby.WriteLine("oldestItemDate = $current_case.search(" & """" & "kind:email" & """" & ", oldest).first().getDate().to_s")
                '        AllCaseDataProcessingRuby.WriteLine("newestItemDate = $current_case.search(" & """" & "kind:email" & """" & ", newest).first().getDate().to_s")
                AllCaseDataProcessingRuby.WriteLine("   oldestItemDate = $current_case.search(" & """" & "flag:top_level item-date:*" & """" & ", oldest).first().getDate().to_s")
                AllCaseDataProcessingRuby.WriteLine("   newestItemDate = $current_case.search(" & """" & "flag:top_level item-date:*" & """" & ", newest).first().getDate().to_s")
                AllCaseDataProcessingRuby.WriteLine("   oldestDate = Date.parse(oldestItemDate.to_s.split(" & """" & "T" & """" & ").first)")
                AllCaseDataProcessingRuby.WriteLine("   newestDate = Date.parse(newestItemDate.to_s.split(" & """" & "T" & """" & ").first)")
                AllCaseDataProcessingRuby.WriteLine("       case_date = [oldestItemDate, newestItemDate, 20]")
                AllCaseDataProcessingRuby.WriteLine("       db.update(" & """" & "UPDATE NuixReportingInfo SET OldestItem = ?, NewestItem = ?, PercentComplete = ? WHERE CaseGUID = '#{case_guid}'" & """" & ", case_date)")
                AllCaseDataProcessingRuby.WriteLine("   lang_detail_count = ''")
                AllCaseDataProcessingRuby.WriteLine("   languages = $current_case.getLanguages")
                AllCaseDataProcessingRuby.WriteLine("   languages.each do |lang|")
                AllCaseDataProcessingRuby.WriteLine("	    lang_count = $current_case.count(" & """" & "lang:#{lang}" & """" & ")")
                AllCaseDataProcessingRuby.WriteLine("	    lang_size = $current_case.getStatistics.getFileSize(" & """" & "lang:#{lang}" & """" & ")")
                AllCaseDataProcessingRuby.WriteLine("       lang_detail_count = lang_detail_count + lang + " & """" & "::" & """" & " + lang_count.to_s + " & """" & "::" & """" & " + lang_size.to_s + " & """" & ";" & """")
                AllCaseDataProcessingRuby.WriteLine("   end")
                AllCaseDataProcessingRuby.WriteLine("   case_lang_details = [lang_detail_count, 17]")
                AllCaseDataProcessingRuby.WriteLine("   db.update(" & """" & "UPDATE NuixReportingInfo SET Languages = ?, PercentComplete = ? WHERE CaseGUID = '#{case_guid}'" & """" & ", case_lang_details)")
                AllCaseDataProcessingRuby.WriteLine("")

                AllCaseDataProcessingRuby.WriteLine("   def sum_filesize(items)")
                AllCaseDataProcessingRuby.WriteLine("       return items.map{|i| i.getFileSize || 0 }.reduce(0,:+)")
                AllCaseDataProcessingRuby.WriteLine("   end")

                AllCaseDataProcessingRuby.WriteLine("   all_items = $current_case.searchUnsorted(" & """" & "*" & """" & ")")
                AllCaseDataProcessingRuby.WriteLine("   all_items_size = $current_case.getStatistics.getFileSize(" & """" & "*" & """" & ")")
                AllCaseDataProcessingRuby.WriteLine("   email_items = $current_case.searchUnsorted(" & """" & "kind:email and flag:top_level" & """" & ")")
                AllCaseDataProcessingRuby.WriteLine("   email_items_size = $current_case.getStatistics.getFileSize(" & """" & "kind:email and flag:top_level" & """" & ")")
                AllCaseDataProcessingRuby.WriteLine("   calendar_items= $current_case.searchUnsorted(" & """" & "kind:calendar" & """" & ")")
                AllCaseDataProcessingRuby.WriteLine("   calendar_items_size = $current_case.getStatistics.getFileSize(" & """" & "kind:calendar" & """" & ")")
                AllCaseDataProcessingRuby.WriteLine("   contact_items = $current_case.searchUnsorted(" & """" & "kind:contact" & """" & ")")
                AllCaseDataProcessingRuby.WriteLine("   contact_items_size = $current_case.getStatistics.getFileSize(" & """" & "kind:contact" & """" & ")")
                AllCaseDataProcessingRuby.WriteLine("   document_items = $current_case.searchUnsorted(" & """" & "kind:document" & """" & ")")
                AllCaseDataProcessingRuby.WriteLine("   document_items_size = $current_case.getStatistics.getFileSize(" & """" & "kind:document" & """" & ")")
                AllCaseDataProcessingRuby.WriteLine("   spreadsheet_items = $current_case.searchUnsorted(" & """" & "kind:spreadsheet" & """" & ")")
                AllCaseDataProcessingRuby.WriteLine("   spreadsheet_items_size = $current_case.getStatistics.getFileSize(" & """" & "kind:spreadsheet" & """" & ")")
                AllCaseDataProcessingRuby.WriteLine("   presentation_items = $current_case.searchUnsorted(" & """" & "kind:presentation" & """" & ")")
                AllCaseDataProcessingRuby.WriteLine("   presentation_items_size = $current_case.getStatistics.getFileSize(" & """" & "kind:presentation" & """" & ")")
                AllCaseDataProcessingRuby.WriteLine("   image_items = $current_case.searchUnsorted(" & """" & "kind:image" & """" & ")")
                AllCaseDataProcessingRuby.WriteLine("   image_items_size = $current_case.getStatistics.getFileSize(" & """" & "kind:image" & """" & ")")
                AllCaseDataProcessingRuby.WriteLine("   drawing_items = $current_case.searchUnsorted(" & """" & "kind:drawing" & """" & ")")
                AllCaseDataProcessingRuby.WriteLine("   drawing_items_size = $current_case.getStatistics.getFileSize(" & """" & "kind:drawing" & """" & ")")
                AllCaseDataProcessingRuby.WriteLine("   other_document_items = $current_case.searchUnsorted(" & """" & "kind:other-document" & """" & ")")
                AllCaseDataProcessingRuby.WriteLine("   other_document_items_size = $current_case.getStatistics.getFileSize(" & """" & "kind:other-document" & """" & ")")
                AllCaseDataProcessingRuby.WriteLine("   multimedia_items = $current_case.searchUnsorted(" & """" & "kind:multimedia" & """" & ")")
                AllCaseDataProcessingRuby.WriteLine("   multimedia_items_size = $current_case.getStatistics.getFileSize(" & """" & "kind:multimedia" & """" & ")")
                AllCaseDataProcessingRuby.WriteLine("   database_items = $current_case.searchUnsorted(" & """" & "kind:database" & """" & ")")
                AllCaseDataProcessingRuby.WriteLine("   database_items_size = $current_case.getStatistics.getFileSize(" & """" & "kind:database" & """" & ")")
                AllCaseDataProcessingRuby.WriteLine("   container_items = $current_case.searchUnsorted(" & """" & "kind:container AND NOT mime-type:application/vnd.nuix-evidence" & """" & ")")
                AllCaseDataProcessingRuby.WriteLine("   container_items_size = $current_case.getStatistics.getFileSize(" & """" & "kind:container AND NOT mime-type:application/vnd.nuix-evidence" & """" & ")")
                AllCaseDataProcessingRuby.WriteLine("   system_items = $current_case.searchUnsorted(" & """" & "kind:system" & """" & ")")
                AllCaseDataProcessingRuby.WriteLine("   system_items_size = $current_case.getStatistics.getFileSize(" & """" & "kind:system" & """" & ")")
                AllCaseDataProcessingRuby.WriteLine("   no_data_items = $current_case.searchUnsorted(" & """" & "kind:no-data" & """" & ")")
                AllCaseDataProcessingRuby.WriteLine("   no_data_items_size = $current_case.getStatistics.getFileSize(" & """" & "kind:no-data" & """" & ")")
                AllCaseDataProcessingRuby.WriteLine("   unrecognised_items = $current_case.searchUnsorted(" & """" & "kind:unrecognised" & """" & ")")
                AllCaseDataProcessingRuby.WriteLine("   unrecognised_items_size = $current_case.getStatistics.getFileSize(" & """" & "kind:unrecognised" & """" & ")")
                AllCaseDataProcessingRuby.WriteLine("   log_items = $current_case.searchUnsorted(" & """" & "kind:log" & """" & ")")
                AllCaseDataProcessingRuby.WriteLine("   log_items_size = $current_case.getStatistics.getFileSize(" & """" & "kind:log" & """" & ")")
                AllCaseDataProcessingRuby.WriteLine("")
                AllCaseDataProcessingRuby.WriteLine("   total_count = all_items.size")
                AllCaseDataProcessingRuby.WriteLine("   email_total_count = email_items.size")
                AllCaseDataProcessingRuby.WriteLine("   calendar_total_count = calendar_items.size")
                AllCaseDataProcessingRuby.WriteLine("   contact_total_count = contact_items.size")
                AllCaseDataProcessingRuby.WriteLine("   document_total_count = document_items.size")
                AllCaseDataProcessingRuby.WriteLine("   spreadsheet_total_count = spreadsheet_items.size")
                AllCaseDataProcessingRuby.WriteLine("   presentation_total_count = presentation_items.size")
                AllCaseDataProcessingRuby.WriteLine("   image_total_count = image_items.size")
                AllCaseDataProcessingRuby.WriteLine("   drawing_total_count = drawing_items.size")
                AllCaseDataProcessingRuby.WriteLine("   other_document_total_count = other_document_items.size")
                AllCaseDataProcessingRuby.WriteLine("   multimedia_total_count = multimedia_items.size")
                AllCaseDataProcessingRuby.WriteLine("   database_total_count = database_items.size")
                AllCaseDataProcessingRuby.WriteLine("   container_total_count = container_items.size")
                AllCaseDataProcessingRuby.WriteLine("   system_total_count = system_items.size")
                AllCaseDataProcessingRuby.WriteLine("   no_data_total_count = no_data_items.size")
                AllCaseDataProcessingRuby.WriteLine("   unrecognised_total_count = unrecognised_items.size")
                AllCaseDataProcessingRuby.WriteLine("   log_total_count = log_items.size")
                AllCaseDataProcessingRuby.WriteLine("")
                AllCaseDataProcessingRuby.WriteLine("   originals = $utilities.getItemUtility.deduplicate(all_items)")
                AllCaseDataProcessingRuby.WriteLine("   duplicates = $utilities.getItemUtility.difference(all_items,originals)")
                AllCaseDataProcessingRuby.WriteLine("   email_originals = $utilities.getItemUtility.deduplicate(email_items)")
                AllCaseDataProcessingRuby.WriteLine("   email_duplicates = $utilities.getItemUtility.difference(email_items,email_originals)")
                AllCaseDataProcessingRuby.WriteLine("   calendar_originals = $utilities.getItemUtility.deduplicate(calendar_items)")
                AllCaseDataProcessingRuby.WriteLine("   calendar_duplicates = $utilities.getItemUtility.difference(calendar_items,calendar_originals)")
                AllCaseDataProcessingRuby.WriteLine("   contact_originals = $utilities.getItemUtility.deduplicate(contact_items)")
                AllCaseDataProcessingRuby.WriteLine("   contact_duplicates = $utilities.getItemUtility.difference(contact_items,contact_originals)")
                AllCaseDataProcessingRuby.WriteLine("   document_originals = $utilities.getItemUtility.deduplicate(document_items)")
                AllCaseDataProcessingRuby.WriteLine("   document_duplicates = $utilities.getItemUtility.difference(document_items,document_originals)")
                AllCaseDataProcessingRuby.WriteLine("   spreadsheet_originals = $utilities.getItemUtility.deduplicate(spreadsheet_items)")
                AllCaseDataProcessingRuby.WriteLine("   spreadsheet_duplicates = $utilities.getItemUtility.difference(spreadsheet_items,spreadsheet_originals)")
                AllCaseDataProcessingRuby.WriteLine("   presentation_originals = $utilities.getItemUtility.deduplicate(presentation_items)")
                AllCaseDataProcessingRuby.WriteLine("   presentation_duplicates = $utilities.getItemUtility.difference(presentation_items,presentation_originals)")
                AllCaseDataProcessingRuby.WriteLine("   image_originals = $utilities.getItemUtility.deduplicate(image_items)")
                AllCaseDataProcessingRuby.WriteLine("   image_duplicates = $utilities.getItemUtility.difference(image_items,image_originals)")
                AllCaseDataProcessingRuby.WriteLine("   drawing_originals = $utilities.getItemUtility.deduplicate(drawing_items)")
                AllCaseDataProcessingRuby.WriteLine("   drawing_duplicates = $utilities.getItemUtility.difference(drawing_items,drawing_originals)")
                AllCaseDataProcessingRuby.WriteLine("   other_document_originals = $utilities.getItemUtility.deduplicate(other_document_items)")
                AllCaseDataProcessingRuby.WriteLine("   other_document_duplicates = $utilities.getItemUtility.difference(other_document_items,other_document_originals)")
                AllCaseDataProcessingRuby.WriteLine("   multimedia_originals = $utilities.getItemUtility.deduplicate(multimedia_items)")
                AllCaseDataProcessingRuby.WriteLine("   multimedia_duplicates = $utilities.getItemUtility.difference(multimedia_items,multimedia_originals)")
                AllCaseDataProcessingRuby.WriteLine("   database_originals = $utilities.getItemUtility.deduplicate(database_items)")
                AllCaseDataProcessingRuby.WriteLine("   database_duplicates = $utilities.getItemUtility.difference(database_items,database_originals)")
                AllCaseDataProcessingRuby.WriteLine("   container_originals = $utilities.getItemUtility.deduplicate(container_items)")
                AllCaseDataProcessingRuby.WriteLine("   container_duplicates = $utilities.getItemUtility.difference(container_items,container_originals)")
                AllCaseDataProcessingRuby.WriteLine("   system_originals = $utilities.getItemUtility.deduplicate(system_items)")
                AllCaseDataProcessingRuby.WriteLine("   system_duplicates = $utilities.getItemUtility.difference(system_items,system_originals)")
                AllCaseDataProcessingRuby.WriteLine("   no_data_originals = $utilities.getItemUtility.deduplicate(no_data_items)")
                AllCaseDataProcessingRuby.WriteLine("   no_data_duplicates = $utilities.getItemUtility.difference(no_data_items,no_data_originals)")
                AllCaseDataProcessingRuby.WriteLine("   unrecognised_originals = $utilities.getItemUtility.deduplicate(unrecognised_items)")
                AllCaseDataProcessingRuby.WriteLine("   unrecognised_duplicates = $utilities.getItemUtility.difference(unrecognised_items,unrecognised_originals)")
                AllCaseDataProcessingRuby.WriteLine("   log_originals = $utilities.getItemUtility.deduplicate(log_items)")
                AllCaseDataProcessingRuby.WriteLine("   log_duplicates = $utilities.getItemUtility.difference(log_items,log_originals)")

                AllCaseDataProcessingRuby.WriteLine("   original_count = originals.size")
                AllCaseDataProcessingRuby.WriteLine("   duplicate_count = duplicates.size")
                AllCaseDataProcessingRuby.WriteLine("   email_original_count = email_originals.size")
                AllCaseDataProcessingRuby.WriteLine("   email_duplicate_count = email_duplicates.size")
                AllCaseDataProcessingRuby.WriteLine("   calendar_original_count = calendar_originals.size")
                AllCaseDataProcessingRuby.WriteLine("   calendar_duplicate_count = calendar_duplicates.size")
                AllCaseDataProcessingRuby.WriteLine("   contact_original_count = contact_originals.size")
                AllCaseDataProcessingRuby.WriteLine("   contact_duplicate_count = contact_duplicates.size")
                AllCaseDataProcessingRuby.WriteLine("   document_original_count = document_originals.size")
                AllCaseDataProcessingRuby.WriteLine("   document_duplicate_count = document_duplicates.size")
                AllCaseDataProcessingRuby.WriteLine("   spreadsheet_original_count = spreadsheet_originals.size")
                AllCaseDataProcessingRuby.WriteLine("   spreadsheet_duplicate_count = spreadsheet_duplicates.size")
                AllCaseDataProcessingRuby.WriteLine("   presentation_original_count = presentation_originals.size")
                AllCaseDataProcessingRuby.WriteLine("   presentation_duplicate_count = presentation_duplicates.size")
                AllCaseDataProcessingRuby.WriteLine("   image_original_count = image_originals.size")
                AllCaseDataProcessingRuby.WriteLine("   image_duplicate_count = image_duplicates.size")
                AllCaseDataProcessingRuby.WriteLine("   drawing_original_count = drawing_originals.size")
                AllCaseDataProcessingRuby.WriteLine("   drawing_duplicate_count = drawing_duplicates.size")
                AllCaseDataProcessingRuby.WriteLine("   other_document_original_count = other_document_originals.size")
                AllCaseDataProcessingRuby.WriteLine("   other_document_duplicate_count = other_document_duplicates.size")
                AllCaseDataProcessingRuby.WriteLine("   multimedia_original_count = multimedia_originals.size")
                AllCaseDataProcessingRuby.WriteLine("   multimedia_duplicate_count = multimedia_duplicates.size")
                AllCaseDataProcessingRuby.WriteLine("   database_original_count = database_originals.size")
                AllCaseDataProcessingRuby.WriteLine("   database_duplicate_count = database_duplicates.size")
                AllCaseDataProcessingRuby.WriteLine("   container_original_count = container_originals.size")
                AllCaseDataProcessingRuby.WriteLine("   container_duplicate_count = container_duplicates.size")
                AllCaseDataProcessingRuby.WriteLine("   system_original_count = system_originals.size")
                AllCaseDataProcessingRuby.WriteLine("   system_duplicate_count = system_duplicates.size")
                AllCaseDataProcessingRuby.WriteLine("   no_data_original_count = no_data_originals.size")
                AllCaseDataProcessingRuby.WriteLine("   no_data_duplicate_count = no_data_duplicates.size")
                AllCaseDataProcessingRuby.WriteLine("   unrecognised_original_count = unrecognised_originals.size")
                AllCaseDataProcessingRuby.WriteLine("   unrecognised_duplicate_count = unrecognised_duplicates.size")
                AllCaseDataProcessingRuby.WriteLine("   log_original_count = log_originals.size")
                AllCaseDataProcessingRuby.WriteLine("   log_duplicate_count = log_duplicates.size")

                AllCaseDataProcessingRuby.WriteLine("")

                AllCaseDataProcessingRuby.WriteLine("   original_filesize = sum_filesize(originals)")
                AllCaseDataProcessingRuby.WriteLine("   duplicate_filesize = sum_filesize(duplicates)")
                AllCaseDataProcessingRuby.WriteLine("   email_original_filesize = sum_filesize(email_originals)")
                AllCaseDataProcessingRuby.WriteLine("   email_duplicate_filesize = sum_filesize(email_duplicates)")
                AllCaseDataProcessingRuby.WriteLine("   calendar_original_filesize = sum_filesize(calendar_originals)")
                AllCaseDataProcessingRuby.WriteLine("   calendar_duplicate_filesize = sum_filesize(calendar_duplicates)")
                AllCaseDataProcessingRuby.WriteLine("   contact_original_filesize = sum_filesize(contact_originals)")
                AllCaseDataProcessingRuby.WriteLine("   contact_duplicate_filesize = sum_filesize(contact_duplicates)")
                AllCaseDataProcessingRuby.WriteLine("   document_original_filesize = sum_filesize(document_originals)")
                AllCaseDataProcessingRuby.WriteLine("   document_duplicate_filesize = sum_filesize(document_duplicates)")
                AllCaseDataProcessingRuby.WriteLine("   spreadsheet_original_filesize = sum_filesize(spreadsheet_originals)")
                AllCaseDataProcessingRuby.WriteLine("   spreadsheet_duplicate_filesize = sum_filesize(spreadsheet_duplicates)")
                AllCaseDataProcessingRuby.WriteLine("   presentation_original_filesize = sum_filesize(presentation_originals)")
                AllCaseDataProcessingRuby.WriteLine("   presentation_duplicate_filesize = sum_filesize(presentation_duplicates)")
                AllCaseDataProcessingRuby.WriteLine("   image_original_filesize = sum_filesize(image_originals)")
                AllCaseDataProcessingRuby.WriteLine("   image_duplicate_filesize = sum_filesize(image_duplicates)")
                AllCaseDataProcessingRuby.WriteLine("   drawing_original_filesize = sum_filesize(drawing_originals)")
                AllCaseDataProcessingRuby.WriteLine("   drawing_duplicate_filesize = sum_filesize(drawing_duplicates)")
                AllCaseDataProcessingRuby.WriteLine("   other_document_original_filesize = sum_filesize(other_document_originals)")
                AllCaseDataProcessingRuby.WriteLine("   other_document_duplicate_filesize = sum_filesize(other_document_duplicates)")
                AllCaseDataProcessingRuby.WriteLine("   multimedia_original_filesize = sum_filesize(multimedia_originals)")
                AllCaseDataProcessingRuby.WriteLine("   multimedia_duplicate_filesize = sum_filesize(multimedia_duplicates)")
                AllCaseDataProcessingRuby.WriteLine("   database_original_filesize = sum_filesize(database_originals)")
                AllCaseDataProcessingRuby.WriteLine("   database_duplicate_filesize = sum_filesize(database_duplicates)")
                AllCaseDataProcessingRuby.WriteLine("   container_original_filesize = sum_filesize(container_originals)")
                AllCaseDataProcessingRuby.WriteLine("   container_duplicate_filesize = sum_filesize(container_duplicates)")
                AllCaseDataProcessingRuby.WriteLine("   system_original_filesize = sum_filesize(system_originals)")
                AllCaseDataProcessingRuby.WriteLine("   system_duplicate_filesize = sum_filesize(system_duplicates)")
                AllCaseDataProcessingRuby.WriteLine("   no_data_original_filesize = sum_filesize(no_data_originals)")
                AllCaseDataProcessingRuby.WriteLine("   no_data_duplicate_filesize = sum_filesize(no_data_duplicates)")
                AllCaseDataProcessingRuby.WriteLine("   unrecognised_original_filesize = sum_filesize(unrecognised_originals)")
                AllCaseDataProcessingRuby.WriteLine("   unrecognised_duplicate_filesize = sum_filesize(unrecognised_duplicates)")
                AllCaseDataProcessingRuby.WriteLine("   log_original_filesize = sum_filesize(log_originals)")
                AllCaseDataProcessingRuby.WriteLine("   log_duplicate_filesize = sum_filesize(log_duplicates)")

                AllCaseDataProcessingRuby.WriteLine("   totals_count = 'Total::' + " & """" & "#{total_count}" & """" & " + " & """" & "::#{all_items_size}" & """" & " + ';Email::' + " & """" & "#{email_total_count}" & """" & " + " & """" & "::#{email_items_size}" & """" & " + ';Calendar::' + " & """" & "#{calendar_total_count}" & """" & " + " & """" & "::#{calendar_items_size}" & """" & " + ';Contact::' + " & """" & "#{contact_total_count}" & """" & " + " & """" & "::#{contact_items_size}" & """" & " + ';Document::' + " & """" & "#{document_total_count}" & """" & " + " & """" & "::#{document_items_size}" & """" & " + ';Spreadsheet::' + " & """" & "#{spreadsheet_total_count}" & """" & " + " & """" & "::#{spreadsheet_items_size}" & """" & " + ';Presentation::' + " & """" & "#{presentation_total_count}" & """" & " + " & """" & "::#{presentation_items_size}" & """" & " + ';Image::' + " & """" & "#{image_total_count}" & """" & " + " & """" & "::#{image_items_size}" & """" & " + ';Drawing::' + " & """" & "#{drawing_total_count}" & """" & " + " & """" & "::#{drawing_items_size}" & """" & " + ';Other-Document::' + " & """" & "#{other_document_total_count}" & """" & " + " & """" & "::#{other_document_items_size}" & """" & " + ';Multimedia::' + " & """" & "#{multimedia_total_count}" & """" & " + " & """" & "::#{multimedia_items_size}" & """" & " + ';Database::' + " & """" & "#{database_total_count}" & """" & " + " & """" & "::#{database_items_size}" & """" & " + ';Container::' + " & """" & "#{container_total_count}" & """" & " + " & """" & "::#{container_items_size}" & """" & " + ';System::' + " & """" & "#{system_total_count}" & """" & " + " & """" & "::#{system_items_size}" & """" & " + ';No-data::' + " & """" & "#{no_data_total_count}" & """" & " + " & """" & "::#{no_data_items_size}" & """" & " + " & "';Unrecognised::' + " & """" & "#{unrecognised_total_count}" & """" & " + " & """" & "::#{unrecognised_items_size}" & """" & " + ';Log::' + " & """" & "#{log_total_count}" & """" & " + " & """" & "::#{log_items_size}" & """")
                AllCaseDataProcessingRuby.WriteLine("   originals_count = 'Total::' + " & """" & "#{original_count}" & """" & " + " & """" & "::#{original_filesize}" & """" & " + ';Email::' + " & """" & "#{email_original_count}" & """" & " + " & """" & "::#{email_original_filesize}" & """" & " + ';Calendar::' + " & """" & "#{calendar_original_count}" & """" & " + " & """" & "::#{calendar_original_filesize}" & """" & " + ';Contact::' + " & """" & "#{contact_original_count}" & """" & " + " & """" & "::#{contact_original_filesize}" & """" & " + ';Document::' + " & """" & "#{document_original_count}" & """" & " + " & """" & "::#{document_original_filesize}" & """" & " + ';Spreadsheet::' + " & """" & "#{spreadsheet_original_count}" & """" & " + " & """" & "::#{spreadsheet_original_filesize}" & """" & " + ';Presentation::' + " & """" & "#{presentation_original_count}" & """" & " + " & """" & "::#{presentation_original_filesize}" & """" & " + ';Image::' + " & """" & "#{image_original_count}" & """" & " + " & """" & "::#{image_original_filesize}" & """" & " + ';Drawing::' + " & """" & "#{drawing_original_count}" & """" & " + " & """" & "::#{drawing_original_filesize}" & """" & " + ';Other-Document::' + " & """" & "#{other_document_original_count}" & """" & " + " & """" & "::#{other_document_original_filesize}" & """" & " + ';Multimedia::' + " & """" & "#{multimedia_original_count}" & """" & " + " & """" & "::#{multimedia_original_filesize}" & """" & " + ';Database::' + " & """" & "#{database_original_count}" & """" & " + " & """" & "::#{database_original_filesize}" & """" & " + ';Container::' + " & """" & "#{container_original_count}" & """" & " + " & """" & "::#{container_original_filesize}" & """" & " + ';System::' + " & """" & "#{system_original_count}" & """" & " + " & """" & "::#{system_original_filesize}" & """" & " + ';No-data::' + " & """" & "#{no_data_original_count}" & """" & " + " & """" & "::#{no_data_original_filesize}" & """" & " + " & "';Unrecognised::' + " & """" & "#{unrecognised_original_count}" & """" & " + " & """" & "::#{unrecognised_original_filesize}" & """" & " + ';Log::' + " & """" & "#{log_original_count}" & """" & " + " & """" & "::#{log_original_filesize}" & """")
                AllCaseDataProcessingRuby.WriteLine("   duplicates_count = 'Total::' + " & """" & "#{duplicate_count}" & """" & " + " & """" & "::#{duplicate_filesize}" & """" & " + ';Email::' + " & """" & "#{email_duplicate_count}" & """" & " + " & """" & "::#{email_duplicate_filesize}" & """" & " + ';Calendar::' + " & """" & "#{calendar_duplicate_count}" & """" & " + " & """" & "::#{calendar_duplicate_filesize}" & """" & " + ';Contact::' + " & """" & "#{contact_duplicate_count}" & """" & " + " & """" & "::#{contact_duplicate_filesize}" & """" & " + ';Document::' + " & """" & "#{document_duplicate_count}" & """" & " + " & """" & "::#{document_duplicate_filesize}" & """" & " + ';Spreadsheet::' + " & """" & "#{spreadsheet_duplicate_count}" & """" & " + " & """" & "::#{spreadsheet_duplicate_filesize}" & """" & " + ';Presentation::' + " & """" & "#{presentation_duplicate_count}" & """" & " + " & """" & "::#{presentation_duplicate_filesize}" & """" & " + ';Image::' + " & """" & "#{image_duplicate_count}" & """" & " + " & """" & "::#{image_duplicate_filesize}" & """" & " + ';Drawing::' + " & """" & "#{drawing_duplicate_count}" & """" & " + " & """" & "::#{drawing_duplicate_filesize}" & """" & " + ';Other-Document::' + " & """" & "#{other_document_duplicate_count}" & """" & " + " & """" & "::#{other_document_duplicate_filesize}" & """" & " + ';Multimedia::' + " & """" & "#{multimedia_duplicate_count}" & """" & " + " & """" & "::#{multimedia_duplicate_filesize}" & """" & " + ';Database::' + " & """" & "#{database_duplicate_count}" & """" & " + " & """" & "::#{database_duplicate_filesize}" & """" & " + ';Container::' + " & """" & "#{container_duplicate_count}" & """" & " + " & """" & "::#{container_duplicate_filesize}" & """" & " + ';System::' + " & """" & "#{system_duplicate_count}" & """" & " + " & """" & "::#{system_duplicate_filesize}" & """" & " + ';No-data::' + " & """" & "#{no_data_duplicate_count}" & """" & " + " & """" & "::#{no_data_duplicate_filesize}" & """" & " + " & "';Unrecognised::' + " & """" & "#{unrecognised_duplicate_count}" & """" & " + " & """" & "::#{unrecognised_duplicate_filesize}" & """" & " + ';Log::' + " & """" & "#{log_duplicate_count}" & """" & " + " & """" & "::#{log_duplicate_filesize}" & """")
                AllCaseDataProcessingRuby.WriteLine("   case_items_data = [totals_count, originals_count, duplicates_count, 22]")
                AllCaseDataProcessingRuby.WriteLine("   db.update(" & """" & "UPDATE NuixReportingInfo SET ItemCounts = ?, OriginalItems = ?, DuplicateItems = ?, PercentComplete = ? WHERE CaseGUID = '#{case_guid}'" & """" & ", case_items_data)")

                AllCaseDataProcessingRuby.WriteLine("rescue")
                AllCaseDataProcessingRuby.WriteLine("   puts 'Error in Block 2'")
                AllCaseDataProcessingRuby.WriteLine("ensure")
                AllCaseDataProcessingRuby.WriteLine("")
                AllCaseDataProcessingRuby.WriteLine("end")
                AllCaseDataProcessingRuby.WriteLine("")
                AllCaseDataProcessingRuby.WriteLine("begin")
                AllCaseDataProcessingRuby.WriteLine("   item_types = $current_case.getItemTypes")
                AllCaseDataProcessingRuby.WriteLine("   mime_type_count = ''")
                AllCaseDataProcessingRuby.WriteLine("   item_types.each do |item_search|")
                AllCaseDataProcessingRuby.WriteLine("       mime_type_count = mime_type_count + " & """" & "#{item_search}::" & """" & " + $current_case.count(" & """" & "mime-type:#{item_search}" & """" & ").to_s + " & """" & "::" & """" & " + $current_case.getStatistics.getFileSize(" & """" & "mime-type:#{item_search}" & """" & ").to_s + " & """" & ";" & """")
                AllCaseDataProcessingRuby.WriteLine("   end")
                AllCaseDataProcessingRuby.WriteLine("   case_mimetype_data = [mime_type_count, 25]")
                AllCaseDataProcessingRuby.WriteLine("   db.update(" & """" & "UPDATE NuixReportingInfo SET MimeTypes = ?, PercentComplete = ? WHERE CaseGUID = '#{case_guid}'" & """" & ", case_mimetype_data)")
                AllCaseDataProcessingRuby.WriteLine("rescue")
                AllCaseDataProcessingRuby.WriteLine("   puts 'Error in Block 3'")
                AllCaseDataProcessingRuby.WriteLine("ensure")
                AllCaseDataProcessingRuby.WriteLine("")
                AllCaseDataProcessingRuby.WriteLine("end")
                AllCaseDataProcessingRuby.WriteLine("")
                AllCaseDataProcessingRuby.WriteLine("begin")
                AllCaseDataProcessingRuby.WriteLine("   date_range_start = nil")
                AllCaseDataProcessingRuby.WriteLine("   date_range_end = nil")
                AllCaseDataProcessingRuby.WriteLine("   user = ''")
                AllCaseDataProcessingRuby.WriteLine("   open_sessions = []")
                AllCaseDataProcessingRuby.WriteLine("   close_sessions = []")
                AllCaseDataProcessingRuby.WriteLine("   loadData_sessions = []")
                AllCaseDataProcessingRuby.WriteLine("   open_sorted = []")
                AllCaseDataProcessingRuby.WriteLine("   close_sorted = []")
                AllCaseDataProcessingRuby.WriteLine("   sorted_array = []")
                AllCaseDataProcessingRuby.WriteLine("   session_data = []")
                AllCaseDataProcessingRuby.WriteLine("   user_sessions = []")
                AllCaseDataProcessingRuby.WriteLine("   session_users = []")
                AllCaseDataProcessingRuby.WriteLine("   loaddata_users = []")
                AllCaseDataProcessingRuby.WriteLine("   time_diff = 0")
                AllCaseDataProcessingRuby.WriteLine("   start_datetime = ''")
                AllCaseDataProcessingRuby.WriteLine("   end_datetime = ''")
                AllCaseDataProcessingRuby.WriteLine("   loadEvents = 0")
                AllCaseDataProcessingRuby.WriteLine("   exportEvents = 0")
                AllCaseDataProcessingRuby.WriteLine("   scriptEvents = 0")
                AllCaseDataProcessingRuby.WriteLine("   scriptdata_users = ''")
                AllCaseDataProcessingRuby.WriteLine("   exportdata_users = ''")
                AllCaseDataProcessingRuby.WriteLine("   total_loadTime = 0")
                AllCaseDataProcessingRuby.WriteLine("   db.update(" & """" & "Delete from UCRTSessionEvents where CaseGUID = '#{case_guid}'" & """" & ")")
                AllCaseDataProcessingRuby.WriteLine("")
                AllCaseDataProcessingRuby.WriteLine("   open_options = {")
                AllCaseDataProcessingRuby.WriteLine("	    " & """" & "startDateAfter" & """" & "=> date_range_start,")
                AllCaseDataProcessingRuby.WriteLine("	    " & """" & "startDateBefore" & """" & "=> date_range_end,")
                AllCaseDataProcessingRuby.WriteLine("	    " & """" & "type" & """" & "=> " & """" & "openSession" & """" & ",")
                AllCaseDataProcessingRuby.WriteLine("   }")
                AllCaseDataProcessingRuby.WriteLine("")
                AllCaseDataProcessingRuby.WriteLine("   close_options = {")
                AllCaseDataProcessingRuby.WriteLine("	    " & """" & "startDateAfter" & """" & "=> date_range_start,")
                AllCaseDataProcessingRuby.WriteLine("	    " & """" & "startDateBefore" & """" & "=> date_range_end,")
                AllCaseDataProcessingRuby.WriteLine("	    " & """" & "type" & """" & "=> " & """" & "closeSession" & """" & ",")
                AllCaseDataProcessingRuby.WriteLine("   }")
                AllCaseDataProcessingRuby.WriteLine("   load_options = {")
                AllCaseDataProcessingRuby.WriteLine("	    " & """" & "startDateAfter" & """" & " => date_range_start,")
                AllCaseDataProcessingRuby.WriteLine("	    " & """" & "startDateBefore" & """" & "=> date_range_end,")
                AllCaseDataProcessingRuby.WriteLine("	    " & """" & "type" & """" & " => " & """" & "loadData" & """" & ",")
                AllCaseDataProcessingRuby.WriteLine("   }")
                AllCaseDataProcessingRuby.WriteLine("   script_options = {")
                AllCaseDataProcessingRuby.WriteLine("	    " & """" & "startDateAfter" & """" & " => date_range_start,")
                AllCaseDataProcessingRuby.WriteLine("	    " & """" & "startDateBefore" & """" & "=> date_range_end,")
                AllCaseDataProcessingRuby.WriteLine("	    " & """" & "type" & """" & " => " & """" & "script" & """" & ",")
                AllCaseDataProcessingRuby.WriteLine("   }")
                AllCaseDataProcessingRuby.WriteLine("   export_options = {")
                AllCaseDataProcessingRuby.WriteLine("	    " & """" & "startDateAfter" & """" & " => date_range_start,")
                AllCaseDataProcessingRuby.WriteLine("	    " & """" & "startDateBefore" & """" & "=> date_range_end,")
                AllCaseDataProcessingRuby.WriteLine("	    " & """" & "type" & """" & " => " & """" & "export" & """" & ",")
                AllCaseDataProcessingRuby.WriteLine("   }")
                AllCaseDataProcessingRuby.WriteLine("#Build history options hash")
                AllCaseDataProcessingRuby.WriteLine("   options = {")
                AllCaseDataProcessingRuby.WriteLine("       " & """" & "startDateAfter" & """" & "=> date_range_start,")
                AllCaseDataProcessingRuby.WriteLine("       " & """" & "startDateBefore" & """" & "=> date_range_end,")
                AllCaseDataProcessingRuby.WriteLine("       " & """" & "user" & """" & "=> user,")
                AllCaseDataProcessingRuby.WriteLine("   }")
                AllCaseDataProcessingRuby.WriteLine("   openhistory = $current_case.getHistory(open_options)")
                AllCaseDataProcessingRuby.WriteLine("	    openhistory.each do |openhist|")
                AllCaseDataProcessingRuby.WriteLine("           username = openhist.getUser")
                AllCaseDataProcessingRuby.WriteLine("           open_datetime = openhist.getStartDate")
                AllCaseDataProcessingRuby.WriteLine("           end_datetime = openhist.getEndDate")
                AllCaseDataProcessingRuby.WriteLine("           event = openhist.getTypeString")
                AllCaseDataProcessingRuby.WriteLine("           session_data = [case_guid, event, open_datetime, end_datetime, '', username]")
                AllCaseDataProcessingRuby.WriteLine("           db.update(" & """" & "INSERT INTO UCRTSessionEvents (CaseGUID, SessionEvent , StartDate , EndDate, Duration, User ) VALUES ( ?, ?, ?, ?, ?, ? )" & """" & ", session_data)")
                AllCaseDataProcessingRuby.WriteLine("		end")
                AllCaseDataProcessingRuby.WriteLine("")
                AllCaseDataProcessingRuby.WriteLine("   closehistory = $current_case.getHistory(close_options)")
                AllCaseDataProcessingRuby.WriteLine("	    closehistory.each do |closehist|")
                AllCaseDataProcessingRuby.WriteLine("           username = closehist.getUser")
                AllCaseDataProcessingRuby.WriteLine("           open_datetime = closehist.getStartDate")
                AllCaseDataProcessingRuby.WriteLine("           end_datetime = closehist.getEndDate")
                AllCaseDataProcessingRuby.WriteLine("           event = closehist.getTypeString")
                AllCaseDataProcessingRuby.WriteLine("           session_data = [case_guid, event, open_datetime, end_datetime, '', username]")
                AllCaseDataProcessingRuby.WriteLine("           db.update(" & """" & "INSERT INTO UCRTSessionEvents (CaseGUID, SessionEvent , StartDate , EndDate, Duration, User ) VALUES ( ?, ?, ?, ?, ?, ? )" & """" & ", session_data)")
                AllCaseDataProcessingRuby.WriteLine("	    end")
                AllCaseDataProcessingRuby.WriteLine("")
                AllCaseDataProcessingRuby.WriteLine("   loadDataHistory = $current_case.getHistory(load_options)")
                AllCaseDataProcessingRuby.WriteLine("	    loadDataHistory.each do |loadhist|")
                AllCaseDataProcessingRuby.WriteLine("           username = loadhist.getUser")
                AllCaseDataProcessingRuby.WriteLine("           start_datetime = loadhist.getStartDate")
                AllCaseDataProcessingRuby.WriteLine("           end_datetime = loadhist.getEndDate")
                AllCaseDataProcessingRuby.WriteLine("           total_time = end_datetime.getMillis - start_datetime.getMillis")
                AllCaseDataProcessingRuby.WriteLine("           total_time = total_time / 1000")

                AllCaseDataProcessingRuby.WriteLine("           event = loadhist.getTypeString")
                AllCaseDataProcessingRuby.WriteLine("           total_loadTime = total_loadTime + total_time")
                AllCaseDataProcessingRuby.WriteLine("           loadEvents = loadEvents + 1")
                AllCaseDataProcessingRuby.WriteLine("           total_time = TimeSpanFormatter.format_seconds(total_time)")

                AllCaseDataProcessingRuby.WriteLine("           session_data = [case_guid, event, start_datetime, end_datetime, total_time, username]")
                AllCaseDataProcessingRuby.WriteLine("           db.update(" & """" & "INSERT INTO UCRTSessionEvents (CaseGUID, SessionEvent , StartDate , EndDate, Duration, User ) VALUES ( ?, ?, ?, ?, ?, ? )" & """" & ", session_data)")
                AllCaseDataProcessingRuby.WriteLine("	    end")
                AllCaseDataProcessingRuby.WriteLine("   total_time = TimeSpanFormatter.format_seconds(total_loadTime)")
                AllCaseDataProcessingRuby.WriteLine("   case_loadtime_data = [total_loadTime, start_datetime, end_datetime, loadEvents, 30]")
                AllCaseDataProcessingRuby.WriteLine("   db.update(" & """" & "UPDATE NuixReportingInfo SET TotalLoadTime = ?, LoadDataStart = ?, LoadDataEnd = ?, LoadEvents = ?, PercentComplete = ? WHERE CaseGUID = '#{case_guid}'" & """" & ", case_loadtime_data)")
                AllCaseDataProcessingRuby.WriteLine("")
                AllCaseDataProcessingRuby.WriteLine("   scriptRunHistory = $current_case.getHistory(script_options)")
                AllCaseDataProcessingRuby.WriteLine("	    scriptRunHistory.each do |scripthist|")
                AllCaseDataProcessingRuby.WriteLine("           username = scripthist.getUser")
                AllCaseDataProcessingRuby.WriteLine("           open_datetime = scripthist.getStartDate")
                AllCaseDataProcessingRuby.WriteLine("           end_datetime = scripthist.getEndDate")
                AllCaseDataProcessingRuby.WriteLine("           event = scripthist.getTypeString")
                AllCaseDataProcessingRuby.WriteLine("           session_data = [case_guid, event, open_datetime, end_datetime, '', username]")
                AllCaseDataProcessingRuby.WriteLine("           db.update(" & """" & "INSERT INTO UCRTSessionEvents (CaseGUID, SessionEvent , StartDate , EndDate, Duration, User ) VALUES ( ?, ?, ?, ?, ?, ? )" & """" & ", session_data)")
                AllCaseDataProcessingRuby.WriteLine("	    end")
                AllCaseDataProcessingRuby.WriteLine("")

                AllCaseDataProcessingRuby.WriteLine("   exportHistory = $current_case.getHistory(export_options)")
                AllCaseDataProcessingRuby.WriteLine("	    exportHistory.each do |exporthist|")
                AllCaseDataProcessingRuby.WriteLine("#	    exporthist.getDetails.each do |key,value|")
                AllCaseDataProcessingRuby.WriteLine("#           detail_blob << " & """" & "#{key}: #{value}" & """")
                AllCaseDataProcessingRuby.WriteLine("#       end")
                AllCaseDataProcessingRuby.WriteLine("#       detail_blob = detail_blob.join(" & """" & "; " & """")
                AllCaseDataProcessingRuby.WriteLine("#       puts detail_blob")
                AllCaseDataProcessingRuby.WriteLine("           exportdetails = exporthist.getDetails")
                AllCaseDataProcessingRuby.WriteLine("           success_items = exportdetails[""" & "processed" & """]")
                AllCaseDataProcessingRuby.WriteLine("           failed_items = exportdetails[""" & "failed" & """]")
                AllCaseDataProcessingRuby.WriteLine("           username = exporthist.getUser")
                AllCaseDataProcessingRuby.WriteLine("           open_datetime = exporthist.getStartDate")
                AllCaseDataProcessingRuby.WriteLine("           end_datetime = exporthist.getEndDate")
                AllCaseDataProcessingRuby.WriteLine("           event = exporthist.getTypeString")
                AllCaseDataProcessingRuby.WriteLine("           session_data = [case_guid, event, open_datetime, end_datetime, '', username, success_items, failed_items]")
                AllCaseDataProcessingRuby.WriteLine("           db.update(" & """" & "INSERT INTO UCRTSessionEvents (CaseGUID, SessionEvent , StartDate , EndDate, Duration, User, Success, Failures ) VALUES ( ?, ?, ?, ?, ?, ?, ?, ? )" & """" & ", session_data)")
                AllCaseDataProcessingRuby.WriteLine("	    end")
                AllCaseDataProcessingRuby.WriteLine("")
                AllCaseDataProcessingRuby.WriteLine("rescue")
                AllCaseDataProcessingRuby.WriteLine("   puts 'Error in Block 4'")
                AllCaseDataProcessingRuby.WriteLine("ensure")
                AllCaseDataProcessingRuby.WriteLine("")
                AllCaseDataProcessingRuby.WriteLine("end")
                AllCaseDataProcessingRuby.WriteLine("")
                AllCaseDataProcessingRuby.WriteLine("begin")

                AllCaseDataProcessingRuby.WriteLine("   db.update(" & """" & "Delete from UCRTDateRange where CaseGUID = '#{case_guid}'" & """" & ")")
                AllCaseDataProcessingRuby.WriteLine("       kinds_count = ''")
                AllCaseDataProcessingRuby.WriteLine("       email_type_size = 0")
                AllCaseDataProcessingRuby.WriteLine("       email_zantaz_type_size = 0")
                AllCaseDataProcessingRuby.WriteLine("       calendar_type_size = 0")
                AllCaseDataProcessingRuby.WriteLine("       contact_type_size = 0")
                AllCaseDataProcessingRuby.WriteLine("       documents_type_size = 0")
                AllCaseDataProcessingRuby.WriteLine("       spreadsheets_type_size = 0")
                AllCaseDataProcessingRuby.WriteLine("       presentations_type_size = 0")
                AllCaseDataProcessingRuby.WriteLine("       image_type_size = 0")
                AllCaseDataProcessingRuby.WriteLine("       drawings_type_size = 0")
                AllCaseDataProcessingRuby.WriteLine("       otherdocuments_type_size = 0")
                AllCaseDataProcessingRuby.WriteLine("       multimedia_type_size = 0")
                AllCaseDataProcessingRuby.WriteLine("       database_type_size = 0")
                AllCaseDataProcessingRuby.WriteLine("       container_type_size = 0")
                AllCaseDataProcessingRuby.WriteLine("       system_type_size = 0")
                AllCaseDataProcessingRuby.WriteLine("       nodata_type_size = 0")
                AllCaseDataProcessingRuby.WriteLine("       unrecognised_type_size = 0")
                AllCaseDataProcessingRuby.WriteLine("       log_type_size = 0")

                AllCaseDataProcessingRuby.WriteLine("       email_search_term = 'kind:email and flag:top_level and has-custodian:0 and not mime-type:application/vnd.zantaz.archive'")
                AllCaseDataProcessingRuby.WriteLine("       email_type_count = $current_case.count(email_search_term)")
                AllCaseDataProcessingRuby.WriteLine("       if email_type_count > 0")
                AllCaseDataProcessingRuby.WriteLine("           email_type_size = email_type_size + $current_case.getStatistics.getFileSize(email_search_term)")
                AllCaseDataProcessingRuby.WriteLine("           oldestItemDate = $current_case.search(" & """" & "#{email_search_term}" & """" & ", oldest).first().getDate().to_s")
                AllCaseDataProcessingRuby.WriteLine("           newestItemDate = $current_case.search(" & """" & "#{email_search_term}" & """" & ", newest).first().getDate().to_s")
                AllCaseDataProcessingRuby.WriteLine("           if oldestItemDate != ''")
                AllCaseDataProcessingRuby.WriteLine("               oldestDate = Date.parse(oldestItemDate.to_s.split(" & """" & "T" & """" & ").first)")
                AllCaseDataProcessingRuby.WriteLine("           end")
                AllCaseDataProcessingRuby.WriteLine("           if newestItemDate != ''")
                AllCaseDataProcessingRuby.WriteLine("               newestDate = Date.parse(newestItemDate.to_s.split(" & """" & "T" & """" & ").first)")
                AllCaseDataProcessingRuby.WriteLine("           end")
                AllCaseDataProcessingRuby.WriteLine("           oldestDate.upto(newestDate) do |dateitem|")
                AllCaseDataProcessingRuby.WriteLine("               dateitemsearch =  dateitem.to_s.delete('-')")
                AllCaseDataProcessingRuby.WriteLine("               email_search_string = " & """" & "#{email_search_term} and item-date:#{dateitemsearch}" & """")
                AllCaseDataProcessingRuby.WriteLine("               email_date_count = $current_case.count(email_search_string)")
                AllCaseDataProcessingRuby.WriteLine("               new_item_date_callback = [case_guid, " & """" & "email" & """" & ", dateitem.to_s, email_date_count, """ & "NA" & """" & "]")
                AllCaseDataProcessingRuby.WriteLine("               db.update(" & """" & "Insert into UCRTDateRange(CaseGUID,ItemType,ItemDate,ItemCount, Custodian) VALUES (?, ?, ?, ?, ?)" & """" & ", new_item_date_callback)")
                AllCaseDataProcessingRuby.WriteLine("           end")
                AllCaseDataProcessingRuby.WriteLine("       end")

                AllCaseDataProcessingRuby.WriteLine("       email_zantaz_search_term = 'kind:email and flag:top_level and has-custodian:0 and mime-type:application/vnd.zantaz.archive'")
                AllCaseDataProcessingRuby.WriteLine("       email_zantaz_type_count = $current_case.count(email_zantaz_search_term)")
                AllCaseDataProcessingRuby.WriteLine("       if email_zantaz_type_count > 0")
                AllCaseDataProcessingRuby.WriteLine("           email_type_size = email_type_size + $current_case.getStatistics.getFileSize(email_search_term)")
                AllCaseDataProcessingRuby.WriteLine("           oldestItemDate = $current_case.search(" & """" & "#{email_zantaz_search_term}" & """" & ", oldest).first().getDate().to_s")
                AllCaseDataProcessingRuby.WriteLine("           newestItemDate = $current_case.search(" & """" & "#{email_zantaz_search_term}" & """" & ", newest).first().getDate().to_s")
                AllCaseDataProcessingRuby.WriteLine("           if oldestItemDate != ''")
                AllCaseDataProcessingRuby.WriteLine("               oldestDate = Date.parse(oldestItemDate.to_s.split(" & """" & "T" & """" & ").first)")
                AllCaseDataProcessingRuby.WriteLine("           end")
                AllCaseDataProcessingRuby.WriteLine("           if newestItemDate != ''")
                AllCaseDataProcessingRuby.WriteLine("               newestDate = Date.parse(newestItemDate.to_s.split(" & """" & "T" & """" & ").first)")
                AllCaseDataProcessingRuby.WriteLine("           end")
                AllCaseDataProcessingRuby.WriteLine("           oldestDate.upto(newestDate) do |dateitem|")
                AllCaseDataProcessingRuby.WriteLine("               dateitemsearch =  dateitem.to_s.delete('-')")
                AllCaseDataProcessingRuby.WriteLine("               email_zantaz_search_string = " & """" & "#{email_zantaz_search_term} and item-date:#{dateitemsearch}" & """")
                AllCaseDataProcessingRuby.WriteLine("               email_zantaz_date_count = $current_case.count(email_zantaz_search_string)")
                AllCaseDataProcessingRuby.WriteLine("               new_item_date_callback = [case_guid, " & """" & "email_zantaz" & """" & ", dateitem.to_s, email_zantaz_date_count, """ & "NA" & """" & "]")
                AllCaseDataProcessingRuby.WriteLine("               db.update(" & """" & "Insert into UCRTDateRange(CaseGUID,ItemType,ItemDate,ItemCount, Custodian) VALUES (?, ?, ?, ?, ?)" & """" & ", new_item_date_callback)")
                AllCaseDataProcessingRuby.WriteLine("           end")
                AllCaseDataProcessingRuby.WriteLine("       end")

                AllCaseDataProcessingRuby.WriteLine("       calendar_search_term = 'kind:calendar and has-custodian:0'")
                AllCaseDataProcessingRuby.WriteLine("       calendar_type_count = $current_case.count(calendar_search_term)")
                AllCaseDataProcessingRuby.WriteLine("       if calendar_type_count > 0 ")
                AllCaseDataProcessingRuby.WriteLine("           calendar_type_size = calendar_type_size + $current_case.getStatistics.getFileSize(calendar_search_term)")
                AllCaseDataProcessingRuby.WriteLine("           oldestItemDate = $current_case.search(" & """" & "#{calendar_search_term}" & """" & ", oldest).first().getDate().to_s")
                AllCaseDataProcessingRuby.WriteLine("           newestItemDate = $current_case.search(" & """" & "#{calendar_search_term}" & """" & ", newest).first().getDate().to_s")
                AllCaseDataProcessingRuby.WriteLine("           if oldestItemDate != ''")
                AllCaseDataProcessingRuby.WriteLine("               oldestDate = Date.parse(oldestItemDate.to_s.split(" & """" & "T" & """" & ").first)")
                AllCaseDataProcessingRuby.WriteLine("           end")
                AllCaseDataProcessingRuby.WriteLine("           if newestItemDate != ''")
                AllCaseDataProcessingRuby.WriteLine("               newestDate = Date.parse(newestItemDate.to_s.split(" & """" & "T" & """" & ").first)")
                AllCaseDataProcessingRuby.WriteLine("           end")
                AllCaseDataProcessingRuby.WriteLine("           oldestDate.upto(newestDate) do |dateitem|")
                AllCaseDataProcessingRuby.WriteLine("               dateitemsearch =  dateitem.to_s.delete('-')")
                AllCaseDataProcessingRuby.WriteLine("               calendar_search_string = " & """" & "#{calendar_search_term} and item-date:#{dateitemsearch}" & """")
                AllCaseDataProcessingRuby.WriteLine("               calendar_date_count = $current_case.count(calendar_search_string)")
                AllCaseDataProcessingRuby.WriteLine("               new_item_date_callback = [case_guid, " & """" & "calendar" & """" & ", dateitem.to_s, calendar_date_count, """ & "NA" & """" & "]")
                AllCaseDataProcessingRuby.WriteLine("               db.update(" & """" & "Insert into UCRTDateRange(CaseGUID,ItemType,ItemDate,ItemCount, Custodian) VALUES (?, ?, ?, ?, ?)" & """" & ", new_item_date_callback)")
                AllCaseDataProcessingRuby.WriteLine("           end")
                AllCaseDataProcessingRuby.WriteLine("       end")

                AllCaseDataProcessingRuby.WriteLine("       contact_search_term = 'kind:contact and has-custodian:0'")
                AllCaseDataProcessingRuby.WriteLine("       contact_type_count = $current_case.count(contact_search_term)")
                AllCaseDataProcessingRuby.WriteLine("       if contact_type_count > 0 ")
                AllCaseDataProcessingRuby.WriteLine("           contact_type_size = contact_type_size + $current_case.getStatistics.getFileSize(contact_search_term)")
                AllCaseDataProcessingRuby.WriteLine("           oldestItemDate = $current_case.search(" & """" & "#{contact_search_term}" & """" & ", oldest).first().getDate().to_s")
                AllCaseDataProcessingRuby.WriteLine("           newestItemDate = $current_case.search(" & """" & "#{contact_search_term}" & """" & ", newest).first().getDate().to_s")
                AllCaseDataProcessingRuby.WriteLine("           if oldestItemDate != ''")
                AllCaseDataProcessingRuby.WriteLine("               oldestDate = Date.parse(oldestItemDate.to_s.split(" & """" & "T" & """" & ").first)")
                AllCaseDataProcessingRuby.WriteLine("           end")
                AllCaseDataProcessingRuby.WriteLine("           if newestItemDate != ''")
                AllCaseDataProcessingRuby.WriteLine("               newestDate = Date.parse(newestItemDate.to_s.split(" & """" & "T" & """" & ").first)")
                AllCaseDataProcessingRuby.WriteLine("           end")
                AllCaseDataProcessingRuby.WriteLine("           oldestDate.upto(newestDate) do |dateitem|")
                AllCaseDataProcessingRuby.WriteLine("               dateitemsearch =  dateitem.to_s.delete('-')")
                AllCaseDataProcessingRuby.WriteLine("               contact_search_string = " & """" & "#{contact_search_term} and item-date:#{dateitemsearch}" & """")
                AllCaseDataProcessingRuby.WriteLine("               contact_date_count = $current_case.count(contact_search_string)")
                AllCaseDataProcessingRuby.WriteLine("               new_item_date_callback = [case_guid, " & """" & "contact" & """" & ", dateitem.to_s, contact_date_count, """ & "NA" & """" & "]")
                AllCaseDataProcessingRuby.WriteLine("               db.update(" & """" & "Insert into UCRTDateRange(CaseGUID,ItemType,ItemDate,ItemCount, Custodian) VALUES (?, ?, ?, ?, ?)" & """" & ", new_item_date_callback)")
                AllCaseDataProcessingRuby.WriteLine("           end")
                AllCaseDataProcessingRuby.WriteLine("       end")

                AllCaseDataProcessingRuby.WriteLine("       document_search_term = 'kind:document and has-custodian:0'")
                AllCaseDataProcessingRuby.WriteLine("       documents_type_count = $current_case.count(document_search_term)")
                AllCaseDataProcessingRuby.WriteLine("       if documents_type_count > 0 ")
                AllCaseDataProcessingRuby.WriteLine("           documents_type_size = documents_type_size + $current_case.getStatistics.getFileSize(document_search_term)")
                AllCaseDataProcessingRuby.WriteLine("           oldestItemDate = $current_case.search(" & """" & "#{document_search_term}" & """" & ", oldest).first().getDate().to_s")
                AllCaseDataProcessingRuby.WriteLine("           newestItemDate = $current_case.search(" & """" & "#{document_search_term}" & """" & ", newest).first().getDate().to_s")
                AllCaseDataProcessingRuby.WriteLine("           if oldestItemDate != ''")
                AllCaseDataProcessingRuby.WriteLine("               oldestDate = Date.parse(oldestItemDate.to_s.split(" & """" & "T" & """" & ").first)")
                AllCaseDataProcessingRuby.WriteLine("           end")
                AllCaseDataProcessingRuby.WriteLine("           if newestItemDate != ''")
                AllCaseDataProcessingRuby.WriteLine("               newestDate = Date.parse(newestItemDate.to_s.split(" & """" & "T" & """" & ").first)")
                AllCaseDataProcessingRuby.WriteLine("           end")
                AllCaseDataProcessingRuby.WriteLine("           oldestDate.upto(newestDate) do |dateitem|")
                AllCaseDataProcessingRuby.WriteLine("               dateitemsearch =  dateitem.to_s.delete('-')")
                AllCaseDataProcessingRuby.WriteLine("               document_search_string = " & """" & "#{document_search_term} and item-date:#{dateitemsearch}" & """")
                AllCaseDataProcessingRuby.WriteLine("               document_date_count = $current_case.count(document_search_string)")
                AllCaseDataProcessingRuby.WriteLine("               new_item_date_callback = [case_guid, " & """" & "document" & """" & ", dateitem.to_s, document_date_count, """ & "NA" & """" & "]")
                AllCaseDataProcessingRuby.WriteLine("               db.update(" & """" & "Insert into UCRTDateRange(CaseGUID,ItemType,ItemDate,ItemCount, Custodian) VALUES (?, ?, ?, ?, ?)" & """" & ", new_item_date_callback)")
                AllCaseDataProcessingRuby.WriteLine("           end")
                AllCaseDataProcessingRuby.WriteLine("       end")

                AllCaseDataProcessingRuby.WriteLine("       spreadsheet_search_term = 'kind:spreadsheet and has-custodian:0'")
                AllCaseDataProcessingRuby.WriteLine("       spreadsheets_type_count = $current_case.count(spreadsheet_search_term)")
                AllCaseDataProcessingRuby.WriteLine("       if spreadsheets_type_count > 0 ")
                AllCaseDataProcessingRuby.WriteLine("           spreadsheets_type_size = spreadsheets_type_size + $current_case.getStatistics.getFileSize(spreadsheet_search_term)")
                AllCaseDataProcessingRuby.WriteLine("           oldestItemDate = $current_case.search(" & """" & "#{spreadsheet_search_term}" & """" & ", oldest).first().getDate().to_s")
                AllCaseDataProcessingRuby.WriteLine("           newestItemDate = $current_case.search(" & """" & "#{spreadsheet_search_term}" & """" & ", newest).first().getDate().to_s")
                AllCaseDataProcessingRuby.WriteLine("           if oldestItemDate != ''")
                AllCaseDataProcessingRuby.WriteLine("               oldestDate = Date.parse(oldestItemDate.to_s.split(" & """" & "T" & """" & ").first)")
                AllCaseDataProcessingRuby.WriteLine("           end")
                AllCaseDataProcessingRuby.WriteLine("           if newestItemDate != ''")
                AllCaseDataProcessingRuby.WriteLine("               newestDate = Date.parse(newestItemDate.to_s.split(" & """" & "T" & """" & ").first)")
                AllCaseDataProcessingRuby.WriteLine("           end")
                AllCaseDataProcessingRuby.WriteLine("           oldestDate.upto(newestDate) do |dateitem|")
                AllCaseDataProcessingRuby.WriteLine("               dateitemsearch =  dateitem.to_s.delete('-')")
                AllCaseDataProcessingRuby.WriteLine("               document_search_string = " & """" & "#{spreadsheet_search_term} and item-date:#{dateitemsearch}" & """")
                AllCaseDataProcessingRuby.WriteLine("               spreadsheets_date_count = $current_case.count(document_search_string)")
                AllCaseDataProcessingRuby.WriteLine("               new_item_date_callback = [case_guid, " & """" & "spreadsheet" & """" & ", dateitem.to_s, spreadsheets_date_count, """ & "NA" & """" & "]")
                AllCaseDataProcessingRuby.WriteLine("               db.update(" & """" & "Insert into UCRTDateRange(CaseGUID,ItemType,ItemDate,ItemCount, Custodian) VALUES (?, ?, ?, ?, ?)" & """" & ", new_item_date_callback)")
                AllCaseDataProcessingRuby.WriteLine("           end")
                AllCaseDataProcessingRuby.WriteLine("       end")

                AllCaseDataProcessingRuby.WriteLine("       presentation_search_term = 'kind:presentation and has-custodian:0'")
                AllCaseDataProcessingRuby.WriteLine("       presentations_type_count = $current_case.count(presentation_search_term)")
                AllCaseDataProcessingRuby.WriteLine("       if presentations_type_count > 0 ")
                AllCaseDataProcessingRuby.WriteLine("           presentations_type_size = presentations_type_size + $current_case.getStatistics.getFileSize(presentation_search_term)")
                AllCaseDataProcessingRuby.WriteLine("           oldestItemDate = $current_case.search(" & """" & "#{presentation_search_term}" & """" & ", oldest).first().getDate().to_s")
                AllCaseDataProcessingRuby.WriteLine("           newestItemDate = $current_case.search(" & """" & "#{presentation_search_term}" & """" & ", newest).first().getDate().to_s")
                AllCaseDataProcessingRuby.WriteLine("           if oldestItemDate != ''")
                AllCaseDataProcessingRuby.WriteLine("               oldestDate = Date.parse(oldestItemDate.to_s.split(" & """" & "T" & """" & ").first)")
                AllCaseDataProcessingRuby.WriteLine("           end")
                AllCaseDataProcessingRuby.WriteLine("           if newestItemDate != ''")
                AllCaseDataProcessingRuby.WriteLine("               newestDate = Date.parse(newestItemDate.to_s.split(" & """" & "T" & """" & ").first)")
                AllCaseDataProcessingRuby.WriteLine("           end")
                AllCaseDataProcessingRuby.WriteLine("           oldestDate.upto(newestDate) do |dateitem|")
                AllCaseDataProcessingRuby.WriteLine("               dateitemsearch =  dateitem.to_s.delete('-')")
                AllCaseDataProcessingRuby.WriteLine("               presentation_search_string = " & """" & "#{presentation_search_term} and item-date:#{dateitemsearch}" & """")
                AllCaseDataProcessingRuby.WriteLine("               presentation_date_count = $current_case.count(presentation_search_string)")
                AllCaseDataProcessingRuby.WriteLine("               new_item_date_callback = [case_guid, " & """" & "presentation" & """" & ", dateitem.to_s, presentation_date_count, """ & "NA" & """" & "]")
                AllCaseDataProcessingRuby.WriteLine("               db.update(" & """" & "Insert into UCRTDateRange(CaseGUID,ItemType,ItemDate,ItemCount, Custodian) VALUES (?, ?, ?, ?, ?)" & """" & ", new_item_date_callback)")
                AllCaseDataProcessingRuby.WriteLine("           end")
                AllCaseDataProcessingRuby.WriteLine("       end")

                AllCaseDataProcessingRuby.WriteLine("       image_search_term = 'kind:image and has-custodian:0'")
                AllCaseDataProcessingRuby.WriteLine("       image_type_count = $current_case.count(image_search_term)")
                AllCaseDataProcessingRuby.WriteLine("       if image_type_count > 0 ")
                AllCaseDataProcessingRuby.WriteLine("           image_type_size = image_type_size + $current_case.getStatistics.getFileSize(image_search_term)")
                AllCaseDataProcessingRuby.WriteLine("           oldestItemDate = $current_case.search(" & """" & "#{image_search_term}" & """" & ", oldest).first().getDate().to_s")
                AllCaseDataProcessingRuby.WriteLine("           newestItemDate = $current_case.search(" & """" & "#{image_search_term}" & """" & ", newest).first().getDate().to_s")
                AllCaseDataProcessingRuby.WriteLine("           if oldestItemDate != ''")
                AllCaseDataProcessingRuby.WriteLine("               oldestDate = Date.parse(oldestItemDate.to_s.split(" & """" & "T" & """" & ").first)")
                AllCaseDataProcessingRuby.WriteLine("           end")
                AllCaseDataProcessingRuby.WriteLine("           if newestItemDate != ''")
                AllCaseDataProcessingRuby.WriteLine("               newestDate = Date.parse(newestItemDate.to_s.split(" & """" & "T" & """" & ").first)")
                AllCaseDataProcessingRuby.WriteLine("           end")
                AllCaseDataProcessingRuby.WriteLine("           oldestDate.upto(newestDate) do |dateitem|")
                AllCaseDataProcessingRuby.WriteLine("               dateitemsearch =  dateitem.to_s.delete('-')")
                AllCaseDataProcessingRuby.WriteLine("               image_search_string = " & """" & "#{image_search_term} and item-date:#{dateitemsearch}" & """")
                AllCaseDataProcessingRuby.WriteLine("               image_date_count = $current_case.count(image_search_string)")
                AllCaseDataProcessingRuby.WriteLine("               new_item_date_callback = [case_guid, " & """" & "image" & """" & ", dateitem.to_s, image_date_count, """ & "NA" & """" & "]")
                AllCaseDataProcessingRuby.WriteLine("               db.update(" & """" & "Insert into UCRTDateRange(CaseGUID,ItemType,ItemDate,ItemCount, Custodian) VALUES (?, ?, ?, ?, ?)" & """" & ", new_item_date_callback)")
                AllCaseDataProcessingRuby.WriteLine("           end")
                AllCaseDataProcessingRuby.WriteLine("       end")

                AllCaseDataProcessingRuby.WriteLine("       drawing_search_term = 'kind:drawing and has-custodian:0'")
                AllCaseDataProcessingRuby.WriteLine("       drawings_type_count = $current_case.count(drawing_search_term)")
                AllCaseDataProcessingRuby.WriteLine("       if drawings_type_count > 0 ")
                AllCaseDataProcessingRuby.WriteLine("           drawings_type_size = drawings_type_size + $current_case.getStatistics.getFileSize(drawing_search_term)")
                AllCaseDataProcessingRuby.WriteLine("           oldestItemDate = $current_case.search(" & """" & "#{drawing_search_term}" & """" & ", oldest).first().getDate().to_s")
                AllCaseDataProcessingRuby.WriteLine("           newestItemDate = $current_case.search(" & """" & "#{drawing_search_term}" & """" & ", newest).first().getDate().to_s")
                AllCaseDataProcessingRuby.WriteLine("           if oldestItemDate != ''")
                AllCaseDataProcessingRuby.WriteLine("               oldestDate = Date.parse(oldestItemDate.to_s.split(" & """" & "T" & """" & ").first)")
                AllCaseDataProcessingRuby.WriteLine("           end")
                AllCaseDataProcessingRuby.WriteLine("           if newestItemDate != ''")
                AllCaseDataProcessingRuby.WriteLine("               newestDate = Date.parse(newestItemDate.to_s.split(" & """" & "T" & """" & ").first)")
                AllCaseDataProcessingRuby.WriteLine("           end")
                AllCaseDataProcessingRuby.WriteLine("           oldestDate.upto(newestDate) do |dateitem|")
                AllCaseDataProcessingRuby.WriteLine("               dateitemsearch =  dateitem.to_s.delete('-')")
                AllCaseDataProcessingRuby.WriteLine("               drawing_search_string = " & """" & "#{drawing_search_term} and item-date:#{dateitemsearch}" & """")
                AllCaseDataProcessingRuby.WriteLine("               drawing_date_count = $current_case.count(drawing_search_string)")
                AllCaseDataProcessingRuby.WriteLine("               new_item_date_callback = [case_guid, " & """" & "drawing" & """" & ", dateitem.to_s, drawing_date_count, """ & "NA" & """" & "]")
                AllCaseDataProcessingRuby.WriteLine("               db.update(" & """" & "Insert into UCRTDateRange(CaseGUID,ItemType,ItemDate,ItemCount, Custodian) VALUES (?, ?, ?, ?, ?)" & """" & ", new_item_date_callback)")
                AllCaseDataProcessingRuby.WriteLine("           end")
                AllCaseDataProcessingRuby.WriteLine("       end")

                AllCaseDataProcessingRuby.WriteLine("       otherdocument_search_term = 'kind:other-document and has-custodian:0'")
                AllCaseDataProcessingRuby.WriteLine("       otherdocuments_type_count = $current_case.count(otherdocument_search_term)")
                AllCaseDataProcessingRuby.WriteLine("       if otherdocuments_type_count > 0 ")
                AllCaseDataProcessingRuby.WriteLine("           otherdocuments_type_size = otherdocuments_type_size + $current_case.getStatistics.getFileSize(otherdocument_search_term)")
                AllCaseDataProcessingRuby.WriteLine("           oldestItemDate = $current_case.search(" & """" & "#{otherdocument_search_term}" & """" & ", oldest).first().getDate().to_s")
                AllCaseDataProcessingRuby.WriteLine("           newestItemDate = $current_case.search(" & """" & "#{otherdocument_search_term}" & """" & ", newest).first().getDate().to_s")
                AllCaseDataProcessingRuby.WriteLine("           if oldestItemDate != ''")
                AllCaseDataProcessingRuby.WriteLine("               oldestDate = Date.parse(oldestItemDate.to_s.split(" & """" & "T" & """" & ").first)")
                AllCaseDataProcessingRuby.WriteLine("           end")
                AllCaseDataProcessingRuby.WriteLine("           if newestItemDate != ''")
                AllCaseDataProcessingRuby.WriteLine("               newestDate = Date.parse(newestItemDate.to_s.split(" & """" & "T" & """" & ").first)")
                AllCaseDataProcessingRuby.WriteLine("           end")
                AllCaseDataProcessingRuby.WriteLine("           oldestDate.upto(newestDate) do |dateitem|")
                AllCaseDataProcessingRuby.WriteLine("               dateitemsearch =  dateitem.to_s.delete('-')")
                AllCaseDataProcessingRuby.WriteLine("               otherdocument_search_string = " & """" & "#{otherdocument_search_term} and item-date:#{dateitemsearch}" & """")
                AllCaseDataProcessingRuby.WriteLine("               otherdocument_date_count = $current_case.count(otherdocument_search_string)")
                AllCaseDataProcessingRuby.WriteLine("               new_item_date_callback = [case_guid, " & """" & "otherdocument" & """" & ", dateitem.to_s, otherdocument_date_count, """ & "NA" & """" & "]")
                AllCaseDataProcessingRuby.WriteLine("               db.update(" & """" & "Insert into UCRTDateRange(CaseGUID,ItemType,ItemDate,ItemCount, Custodian) VALUES (?, ?, ?, ?, ?)" & """" & ", new_item_date_callback)")
                AllCaseDataProcessingRuby.WriteLine("           end")
                AllCaseDataProcessingRuby.WriteLine("       end")

                AllCaseDataProcessingRuby.WriteLine("       multimedia_search_term = 'kind:multimedia and has-custodian:0'")
                AllCaseDataProcessingRuby.WriteLine("       multimedia_type_count = $current_case.count(multimedia_search_term)")
                AllCaseDataProcessingRuby.WriteLine("       if multimedia_type_count > 0 ")
                AllCaseDataProcessingRuby.WriteLine("           multimedia_type_size = multimedia_type_size + $current_case.getStatistics.getFileSize(multimedia_search_term)")
                AllCaseDataProcessingRuby.WriteLine("           oldestItemDate = $current_case.search(" & """" & "#{multimedia_search_term}" & """" & ", oldest).first().getDate().to_s")
                AllCaseDataProcessingRuby.WriteLine("           newestItemDate = $current_case.search(" & """" & "#{multimedia_search_term}" & """" & ", newest).first().getDate().to_s")
                AllCaseDataProcessingRuby.WriteLine("           if oldestItemDate != ''")
                AllCaseDataProcessingRuby.WriteLine("               oldestDate = Date.parse(oldestItemDate.to_s.split(" & """" & "T" & """" & ").first)")
                AllCaseDataProcessingRuby.WriteLine("           end")
                AllCaseDataProcessingRuby.WriteLine("           if newestItemDate != ''")
                AllCaseDataProcessingRuby.WriteLine("               newestDate = Date.parse(newestItemDate.to_s.split(" & """" & "T" & """" & ").first)")
                AllCaseDataProcessingRuby.WriteLine("           end")
                AllCaseDataProcessingRuby.WriteLine("           oldestDate.upto(newestDate) do |dateitem|")
                AllCaseDataProcessingRuby.WriteLine("               dateitemsearch =  dateitem.to_s.delete('-')")
                AllCaseDataProcessingRuby.WriteLine("               multimedia_search_string = " & """" & "#{multimedia_search_term} and item-date:#{dateitemsearch}" & """")
                AllCaseDataProcessingRuby.WriteLine("               multimedia_date_count = $current_case.count(multimedia_search_string)")
                AllCaseDataProcessingRuby.WriteLine("               new_item_date_callback = [case_guid, " & """" & "multimedia" & """" & ", dateitem.to_s, multimedia_date_count, """ & "NA" & """" & "]")
                AllCaseDataProcessingRuby.WriteLine("               db.update(" & """" & "Insert into UCRTDateRange(CaseGUID,ItemType,ItemDate,ItemCount, Custodian) VALUES (?, ?, ?, ?, ?)" & """" & ", new_item_date_callback)")
                AllCaseDataProcessingRuby.WriteLine("           end")
                AllCaseDataProcessingRuby.WriteLine("       end")

                AllCaseDataProcessingRuby.WriteLine("       database_search_term = 'kind:database and has-custodian:0'")
                AllCaseDataProcessingRuby.WriteLine("       database_type_count = $current_case.count(database_search_term)")
                AllCaseDataProcessingRuby.WriteLine("       if database_type_count > 0 ")
                AllCaseDataProcessingRuby.WriteLine("           database_type_size = database_type_size + $current_case.getStatistics.getFileSize(database_search_term)")
                AllCaseDataProcessingRuby.WriteLine("           oldestItemDate = $current_case.search(" & """" & "#{database_search_term}" & """" & ", oldest).first().getDate().to_s")
                AllCaseDataProcessingRuby.WriteLine("           newestItemDate = $current_case.search(" & """" & "#{database_search_term}" & """" & ", newest).first().getDate().to_s")
                AllCaseDataProcessingRuby.WriteLine("           if oldestItemDate != ''")
                AllCaseDataProcessingRuby.WriteLine("               oldestDate = Date.parse(oldestItemDate.to_s.split(" & """" & "T" & """" & ").first)")
                AllCaseDataProcessingRuby.WriteLine("           end")
                AllCaseDataProcessingRuby.WriteLine("           if newestItemDate != ''")
                AllCaseDataProcessingRuby.WriteLine("               newestDate = Date.parse(newestItemDate.to_s.split(" & """" & "T" & """" & ").first)")
                AllCaseDataProcessingRuby.WriteLine("           end")
                AllCaseDataProcessingRuby.WriteLine("           oldestDate.upto(newestDate) do |dateitem|")
                AllCaseDataProcessingRuby.WriteLine("               dateitemsearch =  dateitem.to_s.delete('-')")
                AllCaseDataProcessingRuby.WriteLine("               database_search_string = " & """" & "#{database_search_term} and item-date:#{dateitemsearch}" & """")
                AllCaseDataProcessingRuby.WriteLine("               database_date_count = $current_case.count(database_search_string)")
                AllCaseDataProcessingRuby.WriteLine("               new_item_date_callback = [case_guid, " & """" & "database" & """" & ", dateitem.to_s, database_date_count, """ & "NA" & """" & "]")
                AllCaseDataProcessingRuby.WriteLine("               db.update(" & """" & "Insert into UCRTDateRange(CaseGUID,ItemType,ItemDate,ItemCount, Custodian) VALUES (?, ?, ?, ?, ?)" & """" & ", new_item_date_callback)")
                AllCaseDataProcessingRuby.WriteLine("           end")
                AllCaseDataProcessingRuby.WriteLine("       end")

                AllCaseDataProcessingRuby.WriteLine("       container_search_term = 'kind:container AND NOT mime-type:application/vnd.nuix-evidence and has-custodian:0'")
                AllCaseDataProcessingRuby.WriteLine("       container_type_count = $current_case.count(container_search_term)")
                AllCaseDataProcessingRuby.WriteLine("       if container_type_count > 0 ")
                AllCaseDataProcessingRuby.WriteLine("           container_type_size = container_type_size + $current_case.getStatistics.getFileSize(container_search_term)")
                AllCaseDataProcessingRuby.WriteLine("           oldestItemDate = $current_case.search(" & """" & "#{container_search_term}" & """" & ", oldest).first().getDate().to_s")
                AllCaseDataProcessingRuby.WriteLine("           newestItemDate = $current_case.search(" & """" & "#{container_search_term}" & """" & ", newest).first().getDate().to_s")
                AllCaseDataProcessingRuby.WriteLine("           if oldestItemDate != ''")
                AllCaseDataProcessingRuby.WriteLine("               oldestDate = Date.parse(oldestItemDate.to_s.split(" & """" & "T" & """" & ").first)")
                AllCaseDataProcessingRuby.WriteLine("           end")
                AllCaseDataProcessingRuby.WriteLine("           if newestItemDate != ''")
                AllCaseDataProcessingRuby.WriteLine("               newestDate = Date.parse(newestItemDate.to_s.split(" & """" & "T" & """" & ").first)")
                AllCaseDataProcessingRuby.WriteLine("           end")
                AllCaseDataProcessingRuby.WriteLine("           oldestDate.upto(newestDate) do |dateitem|")
                AllCaseDataProcessingRuby.WriteLine("               dateitemsearch =  dateitem.to_s.delete('-')")
                AllCaseDataProcessingRuby.WriteLine("               container_search_string = " & """" & "container_search_term and item-date:#{dateitemsearch}" & """")
                AllCaseDataProcessingRuby.WriteLine("               container_date_count = $current_case.count(container_search_string)")
                AllCaseDataProcessingRuby.WriteLine("               new_item_date_callback = [case_guid, " & """" & "container" & """" & ", dateitem.to_s, container_date_count, """ & "NA" & """" & "]")
                AllCaseDataProcessingRuby.WriteLine("               db.update(" & """" & "Insert into UCRTDateRange(CaseGUID,ItemType,ItemDate,ItemCount, Custodian) VALUES (?, ?, ?, ?, ?)" & """" & ", new_item_date_callback)")
                AllCaseDataProcessingRuby.WriteLine("           end")
                AllCaseDataProcessingRuby.WriteLine("       end")

                AllCaseDataProcessingRuby.WriteLine("       system_search_term = 'kind:system and has-custodian:0'")
                AllCaseDataProcessingRuby.WriteLine("       system_type_count = $current_case.count(system_search_term)")
                AllCaseDataProcessingRuby.WriteLine("       if system_type_count > 0 ")
                AllCaseDataProcessingRuby.WriteLine("           system_type_size = system_type_size + $current_case.getStatistics.getFileSize(system_search_term)")
                AllCaseDataProcessingRuby.WriteLine("           oldestItemDate = $current_case.search(" & """" & "#{system_search_term}" & """" & ", oldest).first().getDate().to_s")
                AllCaseDataProcessingRuby.WriteLine("           newestItemDate = $current_case.search(" & """" & "#{system_search_term}" & """" & ", newest).first().getDate().to_s")
                AllCaseDataProcessingRuby.WriteLine("           if oldestItemDate != ''")
                AllCaseDataProcessingRuby.WriteLine("               oldestDate = Date.parse(oldestItemDate.to_s.split(" & """" & "T" & """" & ").first)")
                AllCaseDataProcessingRuby.WriteLine("           end")
                AllCaseDataProcessingRuby.WriteLine("           if newestItemDate != ''")
                AllCaseDataProcessingRuby.WriteLine("               newestDate = Date.parse(newestItemDate.to_s.split(" & """" & "T" & """" & ").first)")
                AllCaseDataProcessingRuby.WriteLine("           end")
                AllCaseDataProcessingRuby.WriteLine("           oldestDate.upto(newestDate) do |dateitem|")
                AllCaseDataProcessingRuby.WriteLine("               dateitemsearch =  dateitem.to_s.delete('-')")
                AllCaseDataProcessingRuby.WriteLine("               system_search_string = " & """" & "#{system_search_term} and item-date:#{dateitemsearch}" & """")
                AllCaseDataProcessingRuby.WriteLine("               system_date_count = $current_case.count(system_search_string)")
                AllCaseDataProcessingRuby.WriteLine("               new_item_date_callback = [case_guid, " & """" & "system" & """" & ", dateitem.to_s, system_date_count, """ & "NA" & """" & "]")
                AllCaseDataProcessingRuby.WriteLine("               db.update(" & """" & "Insert into UCRTDateRange(CaseGUID,ItemType,ItemDate,ItemCount, Custodian) VALUES (?, ?, ?, ?, ?)" & """" & ", new_item_date_callback)")
                AllCaseDataProcessingRuby.WriteLine("           end")
                AllCaseDataProcessingRuby.WriteLine("       end")

                AllCaseDataProcessingRuby.WriteLine("       nodata_search_term = 'kind:no-data and has-custodian:0'")
                AllCaseDataProcessingRuby.WriteLine("       nodata_type_count = $current_case.count(nodata_search_term)")
                AllCaseDataProcessingRuby.WriteLine("       if nodata_type_count > 0 ")
                AllCaseDataProcessingRuby.WriteLine("           nodata_type_size = nodata_type_size + $current_case.getStatistics.getFileSize(nodata_search_term)")
                AllCaseDataProcessingRuby.WriteLine("           oldestItemDate = $current_case.search(" & """" & "#{nodata_search_term}" & """" & ", oldest).first().getDate().to_s")
                AllCaseDataProcessingRuby.WriteLine("           newestItemDate = $current_case.search(" & """" & "#{nodata_search_term}" & """" & ", newest).first().getDate().to_s")
                AllCaseDataProcessingRuby.WriteLine("           if oldestItemDate != ''")
                AllCaseDataProcessingRuby.WriteLine("               oldestDate = Date.parse(oldestItemDate.to_s.split(" & """" & "T" & """" & ").first)")
                AllCaseDataProcessingRuby.WriteLine("           end")
                AllCaseDataProcessingRuby.WriteLine("           if newestItemDate != ''")
                AllCaseDataProcessingRuby.WriteLine("               newestDate = Date.parse(newestItemDate.to_s.split(" & """" & "T" & """" & ").first)")
                AllCaseDataProcessingRuby.WriteLine("           end")
                AllCaseDataProcessingRuby.WriteLine("           oldestDate.upto(newestDate) do |dateitem|")
                AllCaseDataProcessingRuby.WriteLine("               dateitemsearch =  dateitem.to_s.delete('-')")
                AllCaseDataProcessingRuby.WriteLine("               nodata_search_string = " & """" & "#{nodata_search_term} and item-date:#{dateitemsearch}" & """")
                AllCaseDataProcessingRuby.WriteLine("               nodata_date_count = $current_case.count(nodata_search_string)")
                AllCaseDataProcessingRuby.WriteLine("               new_item_date_callback = [case_guid, " & """" & "nodata" & """" & ", dateitem.to_s, nodata_date_count, """ & "NA" & """" & "]")
                AllCaseDataProcessingRuby.WriteLine("               db.update(" & """" & "Insert into UCRTDateRange(CaseGUID,ItemType,ItemDate,ItemCount, Custodian) VALUES (?, ?, ?, ?, ?)" & """" & ", new_item_date_callback)")
                AllCaseDataProcessingRuby.WriteLine("           end")
                AllCaseDataProcessingRuby.WriteLine("       end")

                AllCaseDataProcessingRuby.WriteLine("       unrecognised_search_term = 'kind:unrecognised and has-custodian:0'")
                AllCaseDataProcessingRuby.WriteLine("       unrecognised_type_count = $current_case.count(unrecognised_search_term)")
                AllCaseDataProcessingRuby.WriteLine("       if unrecognised_type_count > 0 ")
                AllCaseDataProcessingRuby.WriteLine("           unrecognised_type_size = unrecognised_type_size + $current_case.getStatistics.getFileSize(unrecognised_search_term)")
                AllCaseDataProcessingRuby.WriteLine("           oldestItemDate = $current_case.search(" & """" & "#{unrecognised_search_term}" & """" & ", oldest).first().getDate().to_s")
                AllCaseDataProcessingRuby.WriteLine("           newestItemDate = $current_case.search(" & """" & "#{unrecognised_search_term}" & """" & ", newest).first().getDate().to_s")
                AllCaseDataProcessingRuby.WriteLine("           if oldestItemDate != ''")
                AllCaseDataProcessingRuby.WriteLine("               oldestDate = Date.parse(oldestItemDate.to_s.split(" & """" & "T" & """" & ").first)")
                AllCaseDataProcessingRuby.WriteLine("           end")
                AllCaseDataProcessingRuby.WriteLine("           if newestItemDate != ''")
                AllCaseDataProcessingRuby.WriteLine("               newestDate = Date.parse(newestItemDate.to_s.split(" & """" & "T" & """" & ").first)")
                AllCaseDataProcessingRuby.WriteLine("           end")
                AllCaseDataProcessingRuby.WriteLine("           oldestDate.upto(newestDate) do |dateitem|")
                AllCaseDataProcessingRuby.WriteLine("               dateitemsearch =  dateitem.to_s.delete('-')")
                AllCaseDataProcessingRuby.WriteLine("               unrecognised_search_string = " & """" & "#{unrecognised_search_term} and item-date:#{dateitemsearch}" & """")
                AllCaseDataProcessingRuby.WriteLine("               unrecognised_date_count = $current_case.count(unrecognised_search_string)")
                AllCaseDataProcessingRuby.WriteLine("               new_item_date_callback = [case_guid, " & """" & "unrecognised" & """" & ", dateitem.to_s, unrecognised_date_count, """ & "NA" & """" & "]")
                AllCaseDataProcessingRuby.WriteLine("               db.update(" & """" & "Insert into UCRTDateRange(CaseGUID,ItemType,ItemDate,ItemCount, Custodian) VALUES (?, ?, ?, ?, ?)" & """" & ", new_item_date_callback)")
                AllCaseDataProcessingRuby.WriteLine("           end")
                AllCaseDataProcessingRuby.WriteLine("       end")

                AllCaseDataProcessingRuby.WriteLine("       log_search_term = 'kind:log and has-custodian:0'")
                AllCaseDataProcessingRuby.WriteLine("       log_type_count = $current_case.count(log_search_term)")
                AllCaseDataProcessingRuby.WriteLine("       if log_type_count > 0 ")
                AllCaseDataProcessingRuby.WriteLine("       log_type_size = log_type_size + $current_case.getStatistics.getFileSize(log_search_term)")
                AllCaseDataProcessingRuby.WriteLine("           oldestItemDate = $current_case.search(" & """" & "#{log_search_term}" & """" & ", oldest).first().getDate().to_s")
                AllCaseDataProcessingRuby.WriteLine("           newestItemDate = $current_case.search(" & """" & "#{log_search_term}" & """" & ", newest).first().getDate().to_s")
                AllCaseDataProcessingRuby.WriteLine("           if oldestItemDate != ''")
                AllCaseDataProcessingRuby.WriteLine("               oldestDate = Date.parse(oldestItemDate.to_s.split(" & """" & "T" & """" & ").first)")
                AllCaseDataProcessingRuby.WriteLine("           end")
                AllCaseDataProcessingRuby.WriteLine("           if newestItemDate != ''")
                AllCaseDataProcessingRuby.WriteLine("               newestDate = Date.parse(newestItemDate.to_s.split(" & """" & "T" & """" & ").first)")
                AllCaseDataProcessingRuby.WriteLine("           end")
                AllCaseDataProcessingRuby.WriteLine("           oldestDate.upto(newestDate) do |dateitem|")
                AllCaseDataProcessingRuby.WriteLine("               dateitemsearch =  dateitem.to_s.delete('-')")
                AllCaseDataProcessingRuby.WriteLine("               log_search_string = " & """" & "#{log_search_term} and item-date:#{dateitemsearch}" & """")
                AllCaseDataProcessingRuby.WriteLine("               log_date_count = $current_case.count(log_search_string)")
                AllCaseDataProcessingRuby.WriteLine("               new_item_date_callback = [case_guid, " & """" & "log" & """" & ", dateitem.to_s, log_date_count, """ & "NA" & """" & "]")
                AllCaseDataProcessingRuby.WriteLine("               db.update(" & """" & "Insert into UCRTDateRange(CaseGUID,ItemType,ItemDate,ItemCount, Custodian) VALUES (?, ?, ?, ?, ?)" & """" & ", new_item_date_callback)")
                AllCaseDataProcessingRuby.WriteLine("           end")
                AllCaseDataProcessingRuby.WriteLine("       end")

                AllCaseDataProcessingRuby.WriteLine("   custodian_count = 0")
                AllCaseDataProcessingRuby.WriteLine("   all_custodians_info = ''")
                AllCaseDataProcessingRuby.WriteLine("   custodian_names = $current_case.getAllCustodians")
                AllCaseDataProcessingRuby.WriteLine("       custodian_names.each do |custodian_name|")


                AllCaseDataProcessingRuby.WriteLine("       email_search_term = " & """" & "kind:email and flag:top_level and custodian:'#{custodian_name}' and not mime-type:application/vnd.zantaz.archive" & """")
                AllCaseDataProcessingRuby.WriteLine("       email_type_count_cust = $current_case.count(email_search_term)")
                AllCaseDataProcessingRuby.WriteLine("       if email_type_count_cust > 0")
                AllCaseDataProcessingRuby.WriteLine("       	email_type_count = email_type_count_cust + email_type_count")
                AllCaseDataProcessingRuby.WriteLine("           email_type_size = email_type_size + $current_case.getStatistics.getFileSize(email_search_term)")
                AllCaseDataProcessingRuby.WriteLine("           oldestItemDate = $current_case.search(" & """" & "#{email_search_term}" & """" & ", oldest).first().getDate().to_s")
                AllCaseDataProcessingRuby.WriteLine("           newestItemDate = $current_case.search(" & """" & "#{email_search_term}" & """" & ", newest).first().getDate().to_s")
                AllCaseDataProcessingRuby.WriteLine("           if oldestItemDate != ''")
                AllCaseDataProcessingRuby.WriteLine("               oldestDate = Date.parse(oldestItemDate.to_s.split(" & """" & "T" & """" & ").first)")
                AllCaseDataProcessingRuby.WriteLine("           end")
                AllCaseDataProcessingRuby.WriteLine("           if newestItemDate != ''")
                AllCaseDataProcessingRuby.WriteLine("               newestDate = Date.parse(newestItemDate.to_s.split(" & """" & "T" & """" & ").first)")
                AllCaseDataProcessingRuby.WriteLine("           end")
                AllCaseDataProcessingRuby.WriteLine("           oldestDate.upto(newestDate) do |dateitem|")
                AllCaseDataProcessingRuby.WriteLine("               dateitemsearch =  dateitem.to_s.delete('-')")
                AllCaseDataProcessingRuby.WriteLine("               email_search_string = " & """" & "#{email_search_term} and item-date:#{dateitemsearch}" & """")
                AllCaseDataProcessingRuby.WriteLine("               email_date_count = $current_case.count(email_search_string)")
                AllCaseDataProcessingRuby.WriteLine("               new_item_date_callback = [case_guid, " & """" & "email" & """" & ", dateitem.to_s, email_date_count, custodian_name]")
                AllCaseDataProcessingRuby.WriteLine("               db.update(" & """" & "Insert into UCRTDateRange(CaseGUID,ItemType,ItemDate,ItemCount, Custodian) VALUES (?, ?, ?, ?, ?)" & """" & ", new_item_date_callback)")
                AllCaseDataProcessingRuby.WriteLine("           end")
                AllCaseDataProcessingRuby.WriteLine("       end")

                AllCaseDataProcessingRuby.WriteLine("       email_zantaz_search_term = " & """" & "kind:email and flag:top_level and custodian:'#{custodian_name}' and mime-type:application/vnd.zantaz.archive" & """")
                AllCaseDataProcessingRuby.WriteLine("       email_zantaz_type_count_cust = $current_case.count(email_zantaz_search_term)")
                AllCaseDataProcessingRuby.WriteLine("       if email_zantaz_type_count_cust > 0")
                AllCaseDataProcessingRuby.WriteLine("       	email_zantaz_type_count = email_zantaz_type_count_cust + email_zanaz_type_count")
                AllCaseDataProcessingRuby.WriteLine("           email_zantaz_type_size = email_type_size + $current_case.getStatistics.getFileSize(email_zantaz_search_term)")
                AllCaseDataProcessingRuby.WriteLine("           oldestItemDate = $current_case.search(" & """" & "#{email_zantaz_search_term}" & """" & ", oldest).first().getDate().to_s")
                AllCaseDataProcessingRuby.WriteLine("           newestItemDate = $current_case.search(" & """" & "#{email_zantaz_search_term}" & """" & ", newest).first().getDate().to_s")
                AllCaseDataProcessingRuby.WriteLine("           if oldestItemDate != ''")
                AllCaseDataProcessingRuby.WriteLine("               oldestDate = Date.parse(oldestItemDate.to_s.split(" & """" & "T" & """" & ").first)")
                AllCaseDataProcessingRuby.WriteLine("           end")
                AllCaseDataProcessingRuby.WriteLine("           if newestItemDate != ''")
                AllCaseDataProcessingRuby.WriteLine("               newestDate = Date.parse(newestItemDate.to_s.split(" & """" & "T" & """" & ").first)")
                AllCaseDataProcessingRuby.WriteLine("           end")
                AllCaseDataProcessingRuby.WriteLine("           oldestDate.upto(newestDate) do |dateitem|")
                AllCaseDataProcessingRuby.WriteLine("               dateitemsearch =  dateitem.to_s.delete('-')")
                AllCaseDataProcessingRuby.WriteLine("               email_zantaz_search_string = " & """" & "#{email_zantaz_search_term} and item-date:#{dateitemsearch}" & """")
                AllCaseDataProcessingRuby.WriteLine("               email_zantaz_date_count = $current_case.count(email_zantaz_search_string)")
                AllCaseDataProcessingRuby.WriteLine("               new_item_date_callback = [case_guid, " & """" & "email_santaz" & """" & ", dateitem.to_s, email_zantaz_date_count, custodian_name]")
                AllCaseDataProcessingRuby.WriteLine("               db.update(" & """" & "Insert into UCRTDateRange(CaseGUID,ItemType,ItemDate,ItemCount, Custodian) VALUES (?, ?, ?, ?, ?)" & """" & ", new_item_date_callback)")
                AllCaseDataProcessingRuby.WriteLine("           end")
                AllCaseDataProcessingRuby.WriteLine("       end")

                AllCaseDataProcessingRuby.WriteLine("       calendar_search_term = " & """" & "kind:calendar and custodian:'#{custodian_name}'" & """")
                AllCaseDataProcessingRuby.WriteLine("       calendar_type_count_cust = $current_case.count(calendar_search_term)")
                AllCaseDataProcessingRuby.WriteLine("       if calendar_type_count_cust > 0 ")
                AllCaseDataProcessingRuby.WriteLine("       	calendar_type_count = calendar_type_count_cust + calendar_type_count")
                AllCaseDataProcessingRuby.WriteLine("           calendar_type_size = calendar_type_size + $current_case.getStatistics.getFileSize(calendar_search_term)")
                AllCaseDataProcessingRuby.WriteLine("           oldestItemDate = $current_case.search(" & """" & "#{calendar_search_term}" & """" & ", oldest).first().getDate().to_s")
                AllCaseDataProcessingRuby.WriteLine("           newestItemDate = $current_case.search(" & """" & "#{calendar_search_term}" & """" & ", newest).first().getDate().to_s")
                AllCaseDataProcessingRuby.WriteLine("           if oldestItemDate != ''")
                AllCaseDataProcessingRuby.WriteLine("               oldestDate = Date.parse(oldestItemDate.to_s.split(" & """" & "T" & """" & ").first)")
                AllCaseDataProcessingRuby.WriteLine("           end")
                AllCaseDataProcessingRuby.WriteLine("           if newestItemDate != ''")
                AllCaseDataProcessingRuby.WriteLine("               newestDate = Date.parse(newestItemDate.to_s.split(" & """" & "T" & """" & ").first)")
                AllCaseDataProcessingRuby.WriteLine("           end")
                AllCaseDataProcessingRuby.WriteLine("           oldestDate.upto(newestDate) do |dateitem|")
                AllCaseDataProcessingRuby.WriteLine("               dateitemsearch =  dateitem.to_s.delete('-')")
                AllCaseDataProcessingRuby.WriteLine("               calendar_search_string = " & """" & "#{calendar_search_term} and item-date:#{dateitemsearch}" & """")
                AllCaseDataProcessingRuby.WriteLine("               calendar_date_count = $current_case.count(calendar_search_string)")
                AllCaseDataProcessingRuby.WriteLine("               new_item_date_callback = [case_guid, " & """" & "calendar" & """" & ", dateitem.to_s, calendar_date_count, custodian_name]")
                AllCaseDataProcessingRuby.WriteLine("               db.update(" & """" & "Insert into UCRTDateRange(CaseGUID,ItemType,ItemDate,ItemCount, Custodian) VALUES (?, ?, ?, ?, ?)" & """" & ", new_item_date_callback)")
                AllCaseDataProcessingRuby.WriteLine("           end")
                AllCaseDataProcessingRuby.WriteLine("       end")

                AllCaseDataProcessingRuby.WriteLine("       contact_search_term = " & """" & "kind:contact and custodian:'#{custodian_name}'" & """")
                AllCaseDataProcessingRuby.WriteLine("       contact_type_count_cust = $current_case.count(contact_search_term)")
                AllCaseDataProcessingRuby.WriteLine("       if contact_type_count_cust > 0 ")
                AllCaseDataProcessingRuby.WriteLine("       	contact_type_count = contact_type_count_cust + contact_type_count")
                AllCaseDataProcessingRuby.WriteLine("           contact_type_size = contact_type_size + $current_case.getStatistics.getFileSize(contact_search_term)")
                AllCaseDataProcessingRuby.WriteLine("           oldestItemDate = $current_case.search(" & """" & "#{contact_search_term}" & """" & ", oldest).first().getDate().to_s")
                AllCaseDataProcessingRuby.WriteLine("           newestItemDate = $current_case.search(" & """" & "#{contact_search_term}" & """" & ", newest).first().getDate().to_s")
                AllCaseDataProcessingRuby.WriteLine("           if oldestItemDate != ''")
                AllCaseDataProcessingRuby.WriteLine("               oldestDate = Date.parse(oldestItemDate.to_s.split(" & """" & "T" & """" & ").first)")
                AllCaseDataProcessingRuby.WriteLine("           end")
                AllCaseDataProcessingRuby.WriteLine("           if newestItemDate != ''")
                AllCaseDataProcessingRuby.WriteLine("               newestDate = Date.parse(newestItemDate.to_s.split(" & """" & "T" & """" & ").first)")
                AllCaseDataProcessingRuby.WriteLine("           end")
                AllCaseDataProcessingRuby.WriteLine("           oldestDate.upto(newestDate) do |dateitem|")
                AllCaseDataProcessingRuby.WriteLine("               dateitemsearch =  dateitem.to_s.delete('-')")
                AllCaseDataProcessingRuby.WriteLine("               contact_search_string = " & """" & "#{contact_search_term} and item-date:#{dateitemsearch}" & """")
                AllCaseDataProcessingRuby.WriteLine("               contact_date_count = $current_case.count(contact_search_string)")
                AllCaseDataProcessingRuby.WriteLine("               new_item_date_callback = [case_guid, " & """" & "contact" & """" & ", dateitem.to_s, contact_date_count, custodian_name]")
                AllCaseDataProcessingRuby.WriteLine("               db.update(" & """" & "Insert into UCRTDateRange(CaseGUID,ItemType,ItemDate,ItemCount, Custodian) VALUES (?, ?, ?, ?, ?)" & """" & ", new_item_date_callback)")
                AllCaseDataProcessingRuby.WriteLine("           end")
                AllCaseDataProcessingRuby.WriteLine("       end")

                AllCaseDataProcessingRuby.WriteLine("       document_search_term = " & """" & "kind:document and custodian:'#{custodian_name}'" & """")
                AllCaseDataProcessingRuby.WriteLine("       documents_type_count_cust = $current_case.count(document_search_term)")
                AllCaseDataProcessingRuby.WriteLine("       if documents_type_count_cust > 0 ")
                AllCaseDataProcessingRuby.WriteLine("       	documents_type_count_cust = documents_type_count_cust + documents_type_count")
                AllCaseDataProcessingRuby.WriteLine("           documents_type_size = documents_type_size + $current_case.getStatistics.getFileSize(document_search_term)")
                AllCaseDataProcessingRuby.WriteLine("           oldestItemDate = $current_case.search(" & """" & "#{document_search_term}" & """" & ", oldest).first().getDate().to_s")
                AllCaseDataProcessingRuby.WriteLine("           newestItemDate = $current_case.search(" & """" & "#{document_search_term}" & """" & ", newest).first().getDate().to_s")
                AllCaseDataProcessingRuby.WriteLine("           if oldestItemDate != ''")
                AllCaseDataProcessingRuby.WriteLine("               oldestDate = Date.parse(oldestItemDate.to_s.split(" & """" & "T" & """" & ").first)")
                AllCaseDataProcessingRuby.WriteLine("           end")
                AllCaseDataProcessingRuby.WriteLine("           if newestItemDate != ''")
                AllCaseDataProcessingRuby.WriteLine("               newestDate = Date.parse(newestItemDate.to_s.split(" & """" & "T" & """" & ").first)")
                AllCaseDataProcessingRuby.WriteLine("           end")
                AllCaseDataProcessingRuby.WriteLine("           oldestDate.upto(newestDate) do |dateitem|")
                AllCaseDataProcessingRuby.WriteLine("               dateitemsearch =  dateitem.to_s.delete('-')")
                AllCaseDataProcessingRuby.WriteLine("               document_search_string = " & """" & "#{document_search_term} and item-date:#{dateitemsearch}" & """")
                AllCaseDataProcessingRuby.WriteLine("               document_date_count = $current_case.count(document_search_string)")
                AllCaseDataProcessingRuby.WriteLine("               new_item_date_callback = [case_guid, " & """" & "document" & """" & ", dateitem.to_s, document_date_count, custodian_name]")
                AllCaseDataProcessingRuby.WriteLine("               db.update(" & """" & "Insert into UCRTDateRange(CaseGUID,ItemType,ItemDate,ItemCount, Custodian) VALUES (?, ?, ?, ?, ?)" & """" & ", new_item_date_callback)")
                AllCaseDataProcessingRuby.WriteLine("           end")
                AllCaseDataProcessingRuby.WriteLine("       end")

                AllCaseDataProcessingRuby.WriteLine("       spreadsheet_search_term = " & """" & "kind:spreadsheet and custodian:'#{custodian_name}'" & """")
                AllCaseDataProcessingRuby.WriteLine("       spreadsheets_type_count_cust = $current_case.count(spreadsheet_search_term)")
                AllCaseDataProcessingRuby.WriteLine("       if spreadsheets_type_count_cust > 0 ")
                AllCaseDataProcessingRuby.WriteLine("       	spreadsheets_type_count = spreadsheets_type_count_cust + spreadsheets_type_count")
                AllCaseDataProcessingRuby.WriteLine("           spreadsheets_type_size = spreadsheets_type_size + $current_case.getStatistics.getFileSize(spreadsheet_search_term)")
                AllCaseDataProcessingRuby.WriteLine("           oldestItemDate = $current_case.search(" & """" & "#{spreadsheet_search_term}" & """" & ", oldest).first().getDate().to_s")
                AllCaseDataProcessingRuby.WriteLine("           newestItemDate = $current_case.search(" & """" & "#{spreadsheet_search_term}" & """" & ", newest).first().getDate().to_s")
                AllCaseDataProcessingRuby.WriteLine("           if oldestItemDate != ''")
                AllCaseDataProcessingRuby.WriteLine("               oldestDate = Date.parse(oldestItemDate.to_s.split(" & """" & "T" & """" & ").first)")
                AllCaseDataProcessingRuby.WriteLine("           end")
                AllCaseDataProcessingRuby.WriteLine("           if newestItemDate != ''")
                AllCaseDataProcessingRuby.WriteLine("               newestDate = Date.parse(newestItemDate.to_s.split(" & """" & "T" & """" & ").first)")
                AllCaseDataProcessingRuby.WriteLine("           end")
                AllCaseDataProcessingRuby.WriteLine("           oldestDate.upto(newestDate) do |dateitem|")
                AllCaseDataProcessingRuby.WriteLine("               dateitemsearch =  dateitem.to_s.delete('-')")
                AllCaseDataProcessingRuby.WriteLine("               document_search_string = " & """" & "#{spreadsheet_search_term} and item-date:#{dateitemsearch}" & """")
                AllCaseDataProcessingRuby.WriteLine("               spreadsheets_date_count = $current_case.count(document_search_string)")
                AllCaseDataProcessingRuby.WriteLine("               new_item_date_callback = [case_guid, " & """" & "spreadsheet" & """" & ", dateitem.to_s, spreadsheets_date_count, custodian_name]")
                AllCaseDataProcessingRuby.WriteLine("               db.update(" & """" & "Insert into UCRTDateRange(CaseGUID,ItemType,ItemDate,ItemCount,Custodian) VALUES (?, ?, ?, ?, ?)" & """" & ", new_item_date_callback)")
                AllCaseDataProcessingRuby.WriteLine("           end")
                AllCaseDataProcessingRuby.WriteLine("       end")

                AllCaseDataProcessingRuby.WriteLine("       presentation_search_term = " & """" & "kind:presentation and custodian:'#{custodian_name}'" & """")
                AllCaseDataProcessingRuby.WriteLine("       presentations_type_count_cust = $current_case.count(presentation_search_term)")
                AllCaseDataProcessingRuby.WriteLine("       if presentations_type_count_cust > 0 ")
                AllCaseDataProcessingRuby.WriteLine("       	presentations_type_count = presentations_type_count_cust + presentations_type_count")
                AllCaseDataProcessingRuby.WriteLine("           presentations_type_size = presentations_type_size + $current_case.getStatistics.getFileSize(presentation_search_term)")
                AllCaseDataProcessingRuby.WriteLine("           oldestItemDate = $current_case.search(" & """" & "#{presentation_search_term}" & """" & ", oldest).first().getDate().to_s")
                AllCaseDataProcessingRuby.WriteLine("           newestItemDate = $current_case.search(" & """" & "#{presentation_search_term}" & """" & ", newest).first().getDate().to_s")
                AllCaseDataProcessingRuby.WriteLine("           if oldestItemDate != ''")
                AllCaseDataProcessingRuby.WriteLine("               oldestDate = Date.parse(oldestItemDate.to_s.split(" & """" & "T" & """" & ").first)")
                AllCaseDataProcessingRuby.WriteLine("           end")
                AllCaseDataProcessingRuby.WriteLine("           if newestItemDate != ''")
                AllCaseDataProcessingRuby.WriteLine("               newestDate = Date.parse(newestItemDate.to_s.split(" & """" & "T" & """" & ").first)")
                AllCaseDataProcessingRuby.WriteLine("           end")
                AllCaseDataProcessingRuby.WriteLine("           oldestDate.upto(newestDate) do |dateitem|")
                AllCaseDataProcessingRuby.WriteLine("               dateitemsearch =  dateitem.to_s.delete('-')")
                AllCaseDataProcessingRuby.WriteLine("               presentation_search_string = " & """" & "#{presentation_search_term} and item-date:#{dateitemsearch}" & """")
                AllCaseDataProcessingRuby.WriteLine("               presentation_date_count = $current_case.count(presentation_search_string)")
                AllCaseDataProcessingRuby.WriteLine("               new_item_date_callback = [case_guid, " & """" & "presentation" & """" & ", dateitem.to_s, presentation_date_count, custodian_name]")
                AllCaseDataProcessingRuby.WriteLine("               db.update(" & """" & "Insert into UCRTDateRange(CaseGUID,ItemType,ItemDate,ItemCount, Custodian) VALUES (?, ?, ?, ?, ?)" & """" & ", new_item_date_callback)")
                AllCaseDataProcessingRuby.WriteLine("           end")
                AllCaseDataProcessingRuby.WriteLine("       end")

                AllCaseDataProcessingRuby.WriteLine("       image_search_term = " & """" & "kind:image and custodian:'#{custodian_name}'" & """")
                AllCaseDataProcessingRuby.WriteLine("       image_type_count_cust = $current_case.count(image_search_term)")
                AllCaseDataProcessingRuby.WriteLine("       if image_type_count_cust > 0 ")
                AllCaseDataProcessingRuby.WriteLine("       	image_type_count = image_type_count_cust + image_type_count")
                AllCaseDataProcessingRuby.WriteLine("           image_type_size = image_type_size + $current_case.getStatistics.getFileSize(image_search_term)")
                AllCaseDataProcessingRuby.WriteLine("           oldestItemDate = $current_case.search(" & """" & "#{image_search_term}" & """" & ", oldest).first().getDate().to_s")
                AllCaseDataProcessingRuby.WriteLine("           newestItemDate = $current_case.search(" & """" & "#{image_search_term}" & """" & ", newest).first().getDate().to_s")
                AllCaseDataProcessingRuby.WriteLine("           if oldestItemDate != ''")
                AllCaseDataProcessingRuby.WriteLine("               oldestDate = Date.parse(oldestItemDate.to_s.split(" & """" & "T" & """" & ").first)")
                AllCaseDataProcessingRuby.WriteLine("           end")
                AllCaseDataProcessingRuby.WriteLine("           if newestItemDate != ''")
                AllCaseDataProcessingRuby.WriteLine("               newestDate = Date.parse(newestItemDate.to_s.split(" & """" & "T" & """" & ").first)")
                AllCaseDataProcessingRuby.WriteLine("           end")
                AllCaseDataProcessingRuby.WriteLine("           oldestDate.upto(newestDate) do |dateitem|")
                AllCaseDataProcessingRuby.WriteLine("               dateitemsearch =  dateitem.to_s.delete('-')")
                AllCaseDataProcessingRuby.WriteLine("               image_search_string = " & """" & "#{image_search_term} and item-date:#{dateitemsearch}" & """")
                AllCaseDataProcessingRuby.WriteLine("               image_date_count = $current_case.count(image_search_string)")
                AllCaseDataProcessingRuby.WriteLine("               new_item_date_callback = [case_guid, " & """" & "image" & """" & ", dateitem.to_s, image_date_count, custodian_name]")
                AllCaseDataProcessingRuby.WriteLine("               db.update(" & """" & "Insert into UCRTDateRange(CaseGUID,ItemType,ItemDate,ItemCount, Custodian) VALUES (?, ?, ?, ?, ?)" & """" & ", new_item_date_callback)")
                AllCaseDataProcessingRuby.WriteLine("           end")
                AllCaseDataProcessingRuby.WriteLine("       end")

                AllCaseDataProcessingRuby.WriteLine("       drawing_search_term = " & """" & "kind:drawing and custodian:'#{custodian_name}'" & """")
                AllCaseDataProcessingRuby.WriteLine("       drawings_type_count_cust = $current_case.count(drawing_search_term)")
                AllCaseDataProcessingRuby.WriteLine("       if drawings_type_count_cust > 0 ")
                AllCaseDataProcessingRuby.WriteLine("       	drawings_type_count = drawings_type_count_cust + drawings_type_count")
                AllCaseDataProcessingRuby.WriteLine("           drawings_type_size = drawings_type_size + $current_case.getStatistics.getFileSize(drawing_search_term)")
                AllCaseDataProcessingRuby.WriteLine("           oldestItemDate = $current_case.search(" & """" & "#{drawing_search_term}" & """" & ", oldest).first().getDate().to_s")
                AllCaseDataProcessingRuby.WriteLine("           newestItemDate = $current_case.search(" & """" & "#{drawing_search_term}" & """" & ", newest).first().getDate().to_s")
                AllCaseDataProcessingRuby.WriteLine("           if oldestItemDate != ''")
                AllCaseDataProcessingRuby.WriteLine("               oldestDate = Date.parse(oldestItemDate.to_s.split(" & """" & "T" & """" & ").first)")
                AllCaseDataProcessingRuby.WriteLine("           end")
                AllCaseDataProcessingRuby.WriteLine("           if newestItemDate != ''")
                AllCaseDataProcessingRuby.WriteLine("               newestDate = Date.parse(newestItemDate.to_s.split(" & """" & "T" & """" & ").first)")
                AllCaseDataProcessingRuby.WriteLine("           end")
                AllCaseDataProcessingRuby.WriteLine("           oldestDate.upto(newestDate) do |dateitem|")
                AllCaseDataProcessingRuby.WriteLine("               dateitemsearch =  dateitem.to_s.delete('-')")
                AllCaseDataProcessingRuby.WriteLine("               drawing_search_string = " & """" & "#{drawing_search_term} and item-date:#{dateitemsearch}" & """")
                AllCaseDataProcessingRuby.WriteLine("               drawing_date_count = $current_case.count(drawing_search_string)")
                AllCaseDataProcessingRuby.WriteLine("               new_item_date_callback = [case_guid, " & """" & "drawing" & """" & ", dateitem.to_s, drawing_date_count, custodian_name]")
                AllCaseDataProcessingRuby.WriteLine("               db.update(" & """" & "Insert into UCRTDateRange(CaseGUID,ItemType,ItemDate,ItemCount, Custodian) VALUES (?, ?, ?, ?, ?)" & """" & ", new_item_date_callback)")
                AllCaseDataProcessingRuby.WriteLine("           end")
                AllCaseDataProcessingRuby.WriteLine("       end")

                AllCaseDataProcessingRuby.WriteLine("       otherdocument_search_term = " & """" & "kind:other-document and custodian:'#{custodian_name}'" & """")
                AllCaseDataProcessingRuby.WriteLine("       otherdocuments_type_count_cust = $current_case.count(otherdocument_search_term)")
                AllCaseDataProcessingRuby.WriteLine("       if otherdocuments_type_count_cust > 0 ")
                AllCaseDataProcessingRuby.WriteLine("       	otherdocuments_type_count = otherdocuments_type_count_cust + otherdocuments_type_count")
                AllCaseDataProcessingRuby.WriteLine("           otherdocuments_type_size = otherdocuments_type_size + $current_case.getStatistics.getFileSize(otherdocument_search_term)")
                AllCaseDataProcessingRuby.WriteLine("           oldestItemDate = $current_case.search(" & """" & "#{otherdocument_search_term}" & """" & ", oldest).first().getDate().to_s")
                AllCaseDataProcessingRuby.WriteLine("           newestItemDate = $current_case.search(" & """" & "#{otherdocument_search_term}" & """" & ", newest).first().getDate().to_s")
                AllCaseDataProcessingRuby.WriteLine("           if oldestItemDate != ''")
                AllCaseDataProcessingRuby.WriteLine("               oldestDate = Date.parse(oldestItemDate.to_s.split(" & """" & "T" & """" & ").first)")
                AllCaseDataProcessingRuby.WriteLine("           end")
                AllCaseDataProcessingRuby.WriteLine("           if newestItemDate != ''")
                AllCaseDataProcessingRuby.WriteLine("               newestDate = Date.parse(newestItemDate.to_s.split(" & """" & "T" & """" & ").first)")
                AllCaseDataProcessingRuby.WriteLine("           end")
                AllCaseDataProcessingRuby.WriteLine("           oldestDate.upto(newestDate) do |dateitem|")
                AllCaseDataProcessingRuby.WriteLine("               dateitemsearch =  dateitem.to_s.delete('-')")
                AllCaseDataProcessingRuby.WriteLine("               otherdocument_search_string = " & """" & "#{otherdocument_search_term} and item-date:#{dateitemsearch}" & """")
                AllCaseDataProcessingRuby.WriteLine("               otherdocument_date_count = $current_case.count(otherdocument_search_string)")
                AllCaseDataProcessingRuby.WriteLine("               new_item_date_callback = [case_guid, " & """" & "otherdocument" & """" & ", dateitem.to_s, otherdocument_date_count, custodian_name]")
                AllCaseDataProcessingRuby.WriteLine("               db.update(" & """" & "Insert into UCRTDateRange(CaseGUID,ItemType,ItemDate,ItemCount, Custodian) VALUES (?, ?, ?, ?, ?)" & """" & ", new_item_date_callback)")
                AllCaseDataProcessingRuby.WriteLine("           end")
                AllCaseDataProcessingRuby.WriteLine("       end")

                AllCaseDataProcessingRuby.WriteLine("       multimedia_search_term = " & """" & "kind:multimedia and custodian:'#{custodian_name}'" & """")
                AllCaseDataProcessingRuby.WriteLine("       multimedia_type_count_cust = $current_case.count(multimedia_search_term)")
                AllCaseDataProcessingRuby.WriteLine("       if multimedia_type_count_cust > 0 ")
                AllCaseDataProcessingRuby.WriteLine("       	multimedia_type_count = multimedia_type_count_cust + multimedia_type_count")
                AllCaseDataProcessingRuby.WriteLine("           multimedia_type_size = multimedia_type_size + $current_case.getStatistics.getFileSize(multimedia_search_term)")
                AllCaseDataProcessingRuby.WriteLine("           oldestItemDate = $current_case.search(" & """" & "#{multimedia_search_term}" & """" & ", oldest).first().getDate().to_s")
                AllCaseDataProcessingRuby.WriteLine("           newestItemDate = $current_case.search(" & """" & "#{multimedia_search_term}" & """" & ", newest).first().getDate().to_s")
                AllCaseDataProcessingRuby.WriteLine("           if oldestItemDate != ''")
                AllCaseDataProcessingRuby.WriteLine("               oldestDate = Date.parse(oldestItemDate.to_s.split(" & """" & "T" & """" & ").first)")
                AllCaseDataProcessingRuby.WriteLine("           end")
                AllCaseDataProcessingRuby.WriteLine("           if newestItemDate != ''")
                AllCaseDataProcessingRuby.WriteLine("               newestDate = Date.parse(newestItemDate.to_s.split(" & """" & "T" & """" & ").first)")
                AllCaseDataProcessingRuby.WriteLine("           end")
                AllCaseDataProcessingRuby.WriteLine("           oldestDate.upto(newestDate) do |dateitem|")
                AllCaseDataProcessingRuby.WriteLine("               dateitemsearch =  dateitem.to_s.delete('-')")
                AllCaseDataProcessingRuby.WriteLine("               multimedia_search_string = " & """" & "#{multimedia_search_term} and item-date:#{dateitemsearch}" & """")
                AllCaseDataProcessingRuby.WriteLine("               multimedia_date_count = $current_case.count(multimedia_search_string)")
                AllCaseDataProcessingRuby.WriteLine("               new_item_date_callback = [case_guid, " & """" & "multimedia" & """" & ", dateitem.to_s, multimedia_date_count, custodian_name]")
                AllCaseDataProcessingRuby.WriteLine("               db.update(" & """" & "Insert into UCRTDateRange(CaseGUID,ItemType,ItemDate,ItemCount, Custodian) VALUES (?, ?, ?, ?, ?)" & """" & ", new_item_date_callback)")
                AllCaseDataProcessingRuby.WriteLine("           end")
                AllCaseDataProcessingRuby.WriteLine("       end")

                AllCaseDataProcessingRuby.WriteLine("       database_search_term = " & """" & "kind:database and custodian:'#{custodian_name}'" & """")
                AllCaseDataProcessingRuby.WriteLine("       database_type_count_cust = $current_case.count(database_search_term)")
                AllCaseDataProcessingRuby.WriteLine("       if database_type_count_cust > 0 ")
                AllCaseDataProcessingRuby.WriteLine("       	database_type_count = database_type_count_cust + database_type_count")
                AllCaseDataProcessingRuby.WriteLine("           database_type_size = database_type_size + $current_case.getStatistics.getFileSize(database_search_term)")
                AllCaseDataProcessingRuby.WriteLine("           oldestItemDate = $current_case.search(" & """" & "#{database_search_term}" & """" & ", oldest).first().getDate().to_s")
                AllCaseDataProcessingRuby.WriteLine("           newestItemDate = $current_case.search(" & """" & "#{database_search_term}" & """" & ", newest).first().getDate().to_s")
                AllCaseDataProcessingRuby.WriteLine("           if oldestItemDate != ''")
                AllCaseDataProcessingRuby.WriteLine("               oldestDate = Date.parse(oldestItemDate.to_s.split(" & """" & "T" & """" & ").first)")
                AllCaseDataProcessingRuby.WriteLine("           end")
                AllCaseDataProcessingRuby.WriteLine("           if newestItemDate != ''")
                AllCaseDataProcessingRuby.WriteLine("               newestDate = Date.parse(newestItemDate.to_s.split(" & """" & "T" & """" & ").first)")
                AllCaseDataProcessingRuby.WriteLine("           end")
                AllCaseDataProcessingRuby.WriteLine("           oldestDate.upto(newestDate) do |dateitem|")
                AllCaseDataProcessingRuby.WriteLine("               dateitemsearch =  dateitem.to_s.delete('-')")
                AllCaseDataProcessingRuby.WriteLine("               database_search_string = " & """" & "#{database_search_term} and item-date:#{dateitemsearch}" & """")
                AllCaseDataProcessingRuby.WriteLine("               database_date_count = $current_case.count(database_search_string)")
                AllCaseDataProcessingRuby.WriteLine("               new_item_date_callback = [case_guid, " & """" & "database" & """" & ", dateitem.to_s, database_date_count, custodian_name]")
                AllCaseDataProcessingRuby.WriteLine("               db.update(" & """" & "Insert into UCRTDateRange(CaseGUID,ItemType,ItemDate,ItemCount, Custodian) VALUES (?, ?, ?, ?, ?)" & """" & ", new_item_date_callback)")
                AllCaseDataProcessingRuby.WriteLine("           end")
                AllCaseDataProcessingRuby.WriteLine("       end")

                AllCaseDataProcessingRuby.WriteLine("       container_search_term = " & """" & "kind:container AND NOT mime-type:application/vnd.nuix-evidence and custodian:'#{custodian_name}'" & """")
                AllCaseDataProcessingRuby.WriteLine("       container_type_count_cust = $current_case.count(container_search_term)")
                AllCaseDataProcessingRuby.WriteLine("       if container_type_count_cust > 0 ")
                AllCaseDataProcessingRuby.WriteLine("       	container_type_count = container_type_count_cust + container_type_count")
                AllCaseDataProcessingRuby.WriteLine("           container_type_size = container_type_size + $current_case.getStatistics.getFileSize(container_search_term)")
                AllCaseDataProcessingRuby.WriteLine("           oldestItemDate = $current_case.search(" & """" & "#{container_search_term}" & """" & ", oldest).first().getDate().to_s")
                AllCaseDataProcessingRuby.WriteLine("           newestItemDate = $current_case.search(" & """" & "#{container_search_term}" & """" & ", newest).first().getDate().to_s")
                AllCaseDataProcessingRuby.WriteLine("           if oldestItemDate != ''")
                AllCaseDataProcessingRuby.WriteLine("               oldestDate = Date.parse(oldestItemDate.to_s.split(" & """" & "T" & """" & ").first)")
                AllCaseDataProcessingRuby.WriteLine("           end")
                AllCaseDataProcessingRuby.WriteLine("           if newestItemDate != ''")
                AllCaseDataProcessingRuby.WriteLine("               newestDate = Date.parse(newestItemDate.to_s.split(" & """" & "T" & """" & ").first)")
                AllCaseDataProcessingRuby.WriteLine("           end")
                AllCaseDataProcessingRuby.WriteLine("           oldestDate.upto(newestDate) do |dateitem|")
                AllCaseDataProcessingRuby.WriteLine("               dateitemsearch =  dateitem.to_s.delete('-')")
                AllCaseDataProcessingRuby.WriteLine("               container_search_string = " & """" & "container_search_term and item-date:#{dateitemsearch}" & """")
                AllCaseDataProcessingRuby.WriteLine("               container_date_count = $current_case.count(container_search_string)")
                AllCaseDataProcessingRuby.WriteLine("               new_item_date_callback = [case_guid, " & """" & "container" & """" & ", dateitem.to_s, container_date_count, custodian_name]")
                AllCaseDataProcessingRuby.WriteLine("               db.update(" & """" & "Insert into UCRTDateRange(CaseGUID,ItemType,ItemDate,ItemCount, Custodian) VALUES (?, ?, ?, ?, ?)" & """" & ", new_item_date_callback)")
                AllCaseDataProcessingRuby.WriteLine("           end")
                AllCaseDataProcessingRuby.WriteLine("       end")

                AllCaseDataProcessingRuby.WriteLine("       system_search_term = " & """" & "kind:system and custodian:'#{custodian_name}'" & """")
                AllCaseDataProcessingRuby.WriteLine("       system_type_count_cust = $current_case.count(system_search_term)")
                AllCaseDataProcessingRuby.WriteLine("       if system_type_count_cust > 0 ")
                AllCaseDataProcessingRuby.WriteLine("       	system_type_count = system_type_count_cust + system_type_count")
                AllCaseDataProcessingRuby.WriteLine("           system_type_size = system_type_size + $current_case.getStatistics.getFileSize(system_search_term)")
                AllCaseDataProcessingRuby.WriteLine("           oldestItemDate = $current_case.search(" & """" & "#{system_search_term}" & """" & ", oldest).first().getDate().to_s")
                AllCaseDataProcessingRuby.WriteLine("           newestItemDate = $current_case.search(" & """" & "#{system_search_term}" & """" & ", newest).first().getDate().to_s")
                AllCaseDataProcessingRuby.WriteLine("           if oldestItemDate != ''")
                AllCaseDataProcessingRuby.WriteLine("               oldestDate = Date.parse(oldestItemDate.to_s.split(" & """" & "T" & """" & ").first)")
                AllCaseDataProcessingRuby.WriteLine("           end")
                AllCaseDataProcessingRuby.WriteLine("           if newestItemDate != ''")
                AllCaseDataProcessingRuby.WriteLine("               newestDate = Date.parse(newestItemDate.to_s.split(" & """" & "T" & """" & ").first)")
                AllCaseDataProcessingRuby.WriteLine("           end")
                AllCaseDataProcessingRuby.WriteLine("           oldestDate.upto(newestDate) do |dateitem|")
                AllCaseDataProcessingRuby.WriteLine("               dateitemsearch =  dateitem.to_s.delete('-')")
                AllCaseDataProcessingRuby.WriteLine("               system_search_string = " & """" & "#{system_search_term} and item-date:#{dateitemsearch}" & """")
                AllCaseDataProcessingRuby.WriteLine("               system_date_count = $current_case.count(system_search_string)")
                AllCaseDataProcessingRuby.WriteLine("               new_item_date_callback = [case_guid, " & """" & "system" & """" & ", dateitem.to_s, system_date_count, custodian_name]")
                AllCaseDataProcessingRuby.WriteLine("               db.update(" & """" & "Insert into UCRTDateRange(CaseGUID,ItemType,ItemDate,ItemCount, Custodian) VALUES (?, ?, ?, ?, ?)" & """" & ", new_item_date_callback)")
                AllCaseDataProcessingRuby.WriteLine("           end")
                AllCaseDataProcessingRuby.WriteLine("       end")

                AllCaseDataProcessingRuby.WriteLine("       nodata_search_term = " & """" & "kind:no-data and custodian:'#{custodian_name}'" & """")
                AllCaseDataProcessingRuby.WriteLine("       nodata_type_count_cust = $current_case.count(nodata_search_term)")
                AllCaseDataProcessingRuby.WriteLine("       if nodata_type_count_cust > 0 ")
                AllCaseDataProcessingRuby.WriteLine("       	nondata_type_count = nodata_type_count_cust + nodata_type_count")
                AllCaseDataProcessingRuby.WriteLine("           nodata_type_size = nodata_type_size + $current_case.getStatistics.getFileSize(nodata_search_term)")
                AllCaseDataProcessingRuby.WriteLine("           oldestItemDate = $current_case.search(" & """" & "#{nodata_search_term}" & """" & ", oldest).first().getDate().to_s")
                AllCaseDataProcessingRuby.WriteLine("           newestItemDate = $current_case.search(" & """" & "#{nodata_search_term}" & """" & ", newest).first().getDate().to_s")
                AllCaseDataProcessingRuby.WriteLine("           if oldestItemDate != ''")
                AllCaseDataProcessingRuby.WriteLine("               oldestDate = Date.parse(oldestItemDate.to_s.split(" & """" & "T" & """" & ").first)")
                AllCaseDataProcessingRuby.WriteLine("           end")
                AllCaseDataProcessingRuby.WriteLine("           if newestItemDate != ''")
                AllCaseDataProcessingRuby.WriteLine("               newestDate = Date.parse(newestItemDate.to_s.split(" & """" & "T" & """" & ").first)")
                AllCaseDataProcessingRuby.WriteLine("           end")
                AllCaseDataProcessingRuby.WriteLine("           oldestDate.upto(newestDate) do |dateitem|")
                AllCaseDataProcessingRuby.WriteLine("               dateitemsearch =  dateitem.to_s.delete('-')")
                AllCaseDataProcessingRuby.WriteLine("               nodata_search_string = " & """" & "#{nodata_search_term} and item-date:#{dateitemsearch}" & """")
                AllCaseDataProcessingRuby.WriteLine("               nodata_date_count = $current_case.count(nodata_search_string)")
                AllCaseDataProcessingRuby.WriteLine("               new_item_date_callback = [case_guid, " & """" & "nodata" & """" & ", dateitem.to_s, nodata_date_count, custodian_name]")
                AllCaseDataProcessingRuby.WriteLine("               db.update(" & """" & "Insert into UCRTDateRange(CaseGUID,ItemType,ItemDate,ItemCount, Custodian) VALUES (?, ?, ?, ?, ?)" & """" & ", new_item_date_callback)")
                AllCaseDataProcessingRuby.WriteLine("           end")
                AllCaseDataProcessingRuby.WriteLine("       end")

                AllCaseDataProcessingRuby.WriteLine("       unrecognised_search_term = " & """" & "kind:unrecognised and custodian:'#{custodian_name}'" & """")
                AllCaseDataProcessingRuby.WriteLine("       unrecognised_type_count_cust = $current_case.count(unrecognised_search_term)")
                AllCaseDataProcessingRuby.WriteLine("       if unrecognised_type_count_cust > 0 ")
                AllCaseDataProcessingRuby.WriteLine("       	unrecognised_type_count = unrecognised_type_count_cust + unrecognised_type_count")
                AllCaseDataProcessingRuby.WriteLine("           unrecognised_type_size = unrecognised_type_size + $current_case.getStatistics.getFileSize(unrecognised_search_term)")
                AllCaseDataProcessingRuby.WriteLine("           oldestItemDate = $current_case.search(" & """" & "#{unrecognised_search_term}" & """" & ", oldest).first().getDate().to_s")
                AllCaseDataProcessingRuby.WriteLine("           newestItemDate = $current_case.search(" & """" & "#{unrecognised_search_term}" & """" & ", newest).first().getDate().to_s")
                AllCaseDataProcessingRuby.WriteLine("           if oldestItemDate != ''")
                AllCaseDataProcessingRuby.WriteLine("               oldestDate = Date.parse(oldestItemDate.to_s.split(" & """" & "T" & """" & ").first)")
                AllCaseDataProcessingRuby.WriteLine("           end")
                AllCaseDataProcessingRuby.WriteLine("           if newestItemDate != ''")
                AllCaseDataProcessingRuby.WriteLine("               newestDate = Date.parse(newestItemDate.to_s.split(" & """" & "T" & """" & ").first)")
                AllCaseDataProcessingRuby.WriteLine("           end")
                AllCaseDataProcessingRuby.WriteLine("           oldestDate.upto(newestDate) do |dateitem|")
                AllCaseDataProcessingRuby.WriteLine("               dateitemsearch =  dateitem.to_s.delete('-')")
                AllCaseDataProcessingRuby.WriteLine("               unrecognised_search_string = " & """" & "#{unrecognised_search_term} and item-date:#{dateitemsearch}" & """")
                AllCaseDataProcessingRuby.WriteLine("               unrecognised_date_count = $current_case.count(unrecognised_search_string)")
                AllCaseDataProcessingRuby.WriteLine("               new_item_date_callback = [case_guid, " & """" & "unrecognised" & """" & ", dateitem.to_s, unrecognised_date_count, custodian_name]")
                AllCaseDataProcessingRuby.WriteLine("               db.update(" & """" & "Insert into UCRTDateRange(CaseGUID,ItemType,ItemDate,ItemCount, Custodian) VALUES (?, ?, ?, ?, ?)" & """" & ", new_item_date_callback)")
                AllCaseDataProcessingRuby.WriteLine("           end")
                AllCaseDataProcessingRuby.WriteLine("       end")

                AllCaseDataProcessingRuby.WriteLine("       log_search_term = " & """" & "kind:log and custodian:'#{custodian_name}'" & """")
                AllCaseDataProcessingRuby.WriteLine("       log_type_count_cust = $current_case.count(log_search_term)")
                AllCaseDataProcessingRuby.WriteLine("       if log_type_count_cust > 0 ")
                AllCaseDataProcessingRuby.WriteLine("       	log_type_count = log_type_count_cust + log_type_count")
                AllCaseDataProcessingRuby.WriteLine("       	log_type_size = log_type_size + $current_case.getStatistics.getFileSize(log_search_term)")
                AllCaseDataProcessingRuby.WriteLine("           oldestItemDate = $current_case.search(" & """" & "#{log_search_term}" & """" & ", oldest).first().getDate().to_s")
                AllCaseDataProcessingRuby.WriteLine("           newestItemDate = $current_case.search(" & """" & "#{log_search_term}" & """" & ", newest).first().getDate().to_s")
                AllCaseDataProcessingRuby.WriteLine("           if oldestItemDate != ''")
                AllCaseDataProcessingRuby.WriteLine("               oldestDate = Date.parse(oldestItemDate.to_s.split(" & """" & "T" & """" & ").first)")
                AllCaseDataProcessingRuby.WriteLine("           end")
                AllCaseDataProcessingRuby.WriteLine("           if newestItemDate != ''")
                AllCaseDataProcessingRuby.WriteLine("               newestDate = Date.parse(newestItemDate.to_s.split(" & """" & "T" & """" & ").first)")
                AllCaseDataProcessingRuby.WriteLine("           end")
                AllCaseDataProcessingRuby.WriteLine("           oldestDate.upto(newestDate) do |dateitem|")
                AllCaseDataProcessingRuby.WriteLine("               dateitemsearch =  dateitem.to_s.delete('-')")
                AllCaseDataProcessingRuby.WriteLine("               log_search_string = " & """" & "#{log_search_term} and item-date:#{dateitemsearch}" & """")
                AllCaseDataProcessingRuby.WriteLine("               log_date_count = $current_case.count(log_search_string)")
                AllCaseDataProcessingRuby.WriteLine("               new_item_date_callback = [case_guid, " & """" & "log" & """" & ", dateitem.to_s, log_date_count, custodian_name]")
                AllCaseDataProcessingRuby.WriteLine("               db.update(" & """" & "Insert into UCRTDateRange(CaseGUID,ItemType,ItemDate,ItemCount, Custodian) VALUES (?, ?, ?, ?, ?)" & """" & ", new_item_date_callback)")
                AllCaseDataProcessingRuby.WriteLine("           end")
                AllCaseDataProcessingRuby.WriteLine("       end")
                AllCaseDataProcessingRuby.WriteLine("   end")
                AllCaseDataProcessingRuby.WriteLine("   kinds_count = kinds_count + 'Email::' + " & """" & "#{email_type_count}" & """" & ".to_s + " & """" & "::#{email_type_size}" & """" & " + ' ;'")
                AllCaseDataProcessingRuby.WriteLine("   kinds_count = kinds_count + 'Calendar::' + " & """" & "#{calendar_type_count}" & """" & ".to_s + " & """" & "::#{calendar_type_size}" & """" & " + ';'")
                AllCaseDataProcessingRuby.WriteLine("   kinds_count = kinds_count + 'Contact::' + " & """" & "#{contact_type_count}" & """" & ".to_s + " & """" & "::#{contact_type_size}" & """" & " + ';'")
                AllCaseDataProcessingRuby.WriteLine("   kinds_count = kinds_count + 'Documents::' + " & """" & "#{documents_type_count}" & """" & ".to_s+ " & """" & "::#{documents_type_size}" & """" & " + ';'")
                AllCaseDataProcessingRuby.WriteLine("   kinds_count = kinds_count + 'Spreadsheets::' + " & """" & "#{spreadsheets_type_count}" & """" & ".to_s  + " & """" & "::#{spreadsheets_type_size}" & """" & " + ';'")
                AllCaseDataProcessingRuby.WriteLine("   kinds_count = kinds_count + 'Presentations::' + " & """" & "#{presentations_type_count}" & """" & ".to_s + " & """" & "::#{presentations_type_size}" & """" & " + ';'")
                AllCaseDataProcessingRuby.WriteLine("   kinds_count = kinds_count + 'Image::' + " & """" & "#{image_type_count}" & """" & ".to_s + " & """" & "::#{image_type_size}" & """" & " + ';'")
                AllCaseDataProcessingRuby.WriteLine("   kinds_count = kinds_count + 'Drawings::' + " & """" & "#{drawings_type_count}" & """" & ".to_s + " & """" & "::#{drawings_type_size}" & """" & " + ';'")
                AllCaseDataProcessingRuby.WriteLine("   kinds_count = kinds_count + 'Other Documents::' + " & """" & "#{otherdocuments_type_count}" & """" & ".to_s + " & """" & "::#{otherdocuments_type_size}" & """" & " + ';'")
                AllCaseDataProcessingRuby.WriteLine("   kinds_count = kinds_count + 'Multimedia::' + " & """" & "#{multimedia_type_count}" & """" & ".to_s + " & """" & "::#{multimedia_type_size}" & """" & " + ';'")
                AllCaseDataProcessingRuby.WriteLine("   kinds_count = kinds_count + 'Database::' + " & """" & "#{database_type_count}" & """" & ".to_s + " & """" & "::#{database_type_size}" & """" & " + ';'")
                AllCaseDataProcessingRuby.WriteLine("   kinds_count = kinds_count + 'Container::' + " & """" & "#{container_type_count}" & """" & ".to_s + " & """" & "::#{container_type_size}" & """" & " + ';'")
                AllCaseDataProcessingRuby.WriteLine("   kinds_count = kinds_count + 'System::' + " & """" & "#{system_type_count}" & """" & ".to_s + " & """" & "::#{system_type_size}" & """" & " + ';'")
                AllCaseDataProcessingRuby.WriteLine("   kinds_count = kinds_count + 'No Data::' + " & """" & "#{nodata_type_count}" & """" & ".to_s + " & """" & "::#{nodata_type_size}" & """" & " + ';'")
                AllCaseDataProcessingRuby.WriteLine("   kinds_count = kinds_count + 'Unrecognised::' + " & """" & "#{unrecognised_type_count}" & """" & ".to_s + " & """" & "::#{unrecognised_type_size}" & """" & " + ';'")
                AllCaseDataProcessingRuby.WriteLine("   kinds_count = kinds_count + 'Logs::' + " & """" & "#{log_type_count}" & """" & ".to_s + " & """" & "::#{log_type_size}" & """" & " + ';'")
                AllCaseDataProcessingRuby.WriteLine("   case_kinds_data = [kinds_count, 70]")
                AllCaseDataProcessingRuby.WriteLine("   db.update(" & """" & "UPDATE NuixReportingInfo SET ItemTypes = ?, PercentComplete = ? WHERE CaseGUID = '#{case_guid}'" & """" & ", case_kinds_data)")
                AllCaseDataProcessingRuby.WriteLine("rescue")
                AllCaseDataProcessingRuby.WriteLine("   puts 'Error in Block 5'")
                AllCaseDataProcessingRuby.WriteLine("ensure")
                AllCaseDataProcessingRuby.WriteLine("")
                AllCaseDataProcessingRuby.WriteLine("end")
                AllCaseDataProcessingRuby.WriteLine("")

                AllCaseDataProcessingRuby.WriteLine("begin")
                AllCaseDataProcessingRuby.WriteLine("   corrupted_container_count = $current_case.count('properties:FailureDetail AND NOT flag:encrypted AND has-text:0 AND ( has-embedded-data:1 OR kind:container OR kind:database )')")
                AllCaseDataProcessingRuby.WriteLine("   if corrupted_container_count == 0 ")
                AllCaseDataProcessingRuby.WriteLine("       corrupted_container_size = 0")
                AllCaseDataProcessingRuby.WriteLine("   else")
                AllCaseDataProcessingRuby.WriteLine("       corrupted_container_size = $current_case.getStatistics.getFileSize('properties:FailureDetail AND NOT flag:encrypted AND has-text:0 AND ( has-embedded-data:1 OR kind:container OR kind:database )')")
                AllCaseDataProcessingRuby.WriteLine("   end")

                AllCaseDataProcessingRuby.WriteLine("   unsupported_container_count = $current_case.count('kind:( container OR database ) AND NOT flag:encrypted AND has-embedded-data:0 AND NOT flag:partially_processed AND NOT flag:not_processed AND NOT properties:FailureDetail')")
                AllCaseDataProcessingRuby.WriteLine("   if unsupported_container_count  == 0 ")
                AllCaseDataProcessingRuby.WriteLine("       unsupported_container_size = 0")
                AllCaseDataProcessingRuby.WriteLine("   else")
                AllCaseDataProcessingRuby.WriteLine("       unsupported_container_size = $current_case.getStatistics.getFileSize('kind:( container OR database ) AND NOT flag:encrypted AND has-embedded-data:0 AND NOT flag:partially_processed AND NOT flag:not_processed AND NOT properties:FailureDetail')")
                AllCaseDataProcessingRuby.WriteLine("   end")

                AllCaseDataProcessingRuby.WriteLine("   nonsearchable_pdfs_count = $current_case.count('mime-type:application/pdf AND NOT content:*')")
                AllCaseDataProcessingRuby.WriteLine("   if nonsearchable_pdfs_count == 0 ")
                AllCaseDataProcessingRuby.WriteLine("       nonsearchable_pdfs_size = 0")
                AllCaseDataProcessingRuby.WriteLine("   else")
                AllCaseDataProcessingRuby.WriteLine("       nonsearchable_pdfs_size = $current_case.getStatistics.getFileSize('mime-type:application/pdf AND NOT content:*')")
                AllCaseDataProcessingRuby.WriteLine("   end")

                AllCaseDataProcessingRuby.WriteLine("   text_updated_count = $current_case.count('modifications:text_updated')")
                AllCaseDataProcessingRuby.WriteLine("   if text_updated_count == 0 ")
                AllCaseDataProcessingRuby.WriteLine("       text_updated_size = 0")
                AllCaseDataProcessingRuby.WriteLine("   else")
                AllCaseDataProcessingRuby.WriteLine("       text_updated_size = $current_case.getStatistics.getFileSize('modifications:text_updated')")
                AllCaseDataProcessingRuby.WriteLine("   end")

                AllCaseDataProcessingRuby.WriteLine("   bad_extension_count = $current_case.count('flag:irregular_file_extension')")
                AllCaseDataProcessingRuby.WriteLine("   if bad_extension_count == 0 ")
                AllCaseDataProcessingRuby.WriteLine("       bad_extension_size = 0")
                AllCaseDataProcessingRuby.WriteLine("   else")
                AllCaseDataProcessingRuby.WriteLine("       bad_extension_size = $current_case.getStatistics.getFileSize('flag:irregular_file_extension')")
                AllCaseDataProcessingRuby.WriteLine("   end")

                AllCaseDataProcessingRuby.WriteLine("   unrecognized_count = $current_case.count('kind:unrecognised')")
                AllCaseDataProcessingRuby.WriteLine("   if unrecognized_count == 0 ")
                AllCaseDataProcessingRuby.WriteLine("       unrecognized_size = 0")
                AllCaseDataProcessingRuby.WriteLine("   else")
                AllCaseDataProcessingRuby.WriteLine("       unrecognized_size = $current_case.getStatistics.getFileSize('kind:unrecognised')")
                AllCaseDataProcessingRuby.WriteLine("   end")

                AllCaseDataProcessingRuby.WriteLine("   unsupported_count = $current_case.count('NOT flag:encrypted AND has-embedded-data:0 AND ( ( has-text:0 AND has-image:0 AND NOT flag:not_processed AND NOT kind:multimedia AND NOT mime-type:application/vnd.ms-shortcut AND NOT mime-type:application/x-contact AND NOT kind:system AND NOT mime-type:( application/vnd.apache-error-log-entry OR application/vnd.linux-syslog-entry OR application/vnd.logstash-log-entry OR application/vnd.ms-iis-log-entry OR application/vnd.ms-windows-event-log-record OR application/vnd.ms-windows-event-logx-record OR application/vnd.ms-windows-setup-api-win7-win8-log-boot-entry OR application/vnd.ms-windows-setup-api-win7-win8-log-section-entry OR application/vnd.ms-windows-setup-api-xp-log-entry OR application/vnd.squid-access-log-entry OR application/vnd.tcpdump.record OR application/vnd.tcpdump.tcp.stream OR application/vnd.tcpdump.udp.stream OR application/x-pcapng-entry OR filesystem/x-linux-login-logfile-record OR filesystem/x-ntfs-logfile-record OR server/dropbox-log-event OR text/x-common-log-entry OR text/x-log-entry ) AND NOT kind:log AND NOT mime-type:application/vnd.ms-exchange-stm ) OR mime-type:application/vnd.lotus-notes )')")
                AllCaseDataProcessingRuby.WriteLine("   if unsupported_count == 0 ")
                AllCaseDataProcessingRuby.WriteLine("       unsupported_size = 0")
                AllCaseDataProcessingRuby.WriteLine("   else")
                AllCaseDataProcessingRuby.WriteLine("       unsupported_size = $current_case.getStatistics.getFileSize('NOT flag:encrypted AND has-embedded-data:0 AND ( ( has-text:0 AND has-image:0 AND NOT flag:not_processed AND NOT kind:multimedia AND NOT mime-type:application/vnd.ms-shortcut AND NOT mime-type:application/x-contact AND NOT kind:system AND NOT mime-type:( application/vnd.apache-error-log-entry OR application/vnd.linux-syslog-entry OR application/vnd.logstash-log-entry OR application/vnd.ms-iis-log-entry OR application/vnd.ms-windows-event-log-record OR application/vnd.ms-windows-event-logx-record OR application/vnd.ms-windows-setup-api-win7-win8-log-boot-entry OR application/vnd.ms-windows-setup-api-win7-win8-log-section-entry OR application/vnd.ms-windows-setup-api-xp-log-entry OR application/vnd.squid-access-log-entry OR application/vnd.tcpdump.record OR application/vnd.tcpdump.tcp.stream OR application/vnd.tcpdump.udp.stream OR application/x-pcapng-entry OR filesystem/x-linux-login-logfile-record OR filesystem/x-ntfs-logfile-record OR server/dropbox-log-event OR text/x-common-log-entry OR text/x-log-entry ) AND NOT kind:log AND NOT mime-type:application/vnd.ms-exchange-stm ) OR mime-type:application/vnd.lotus-notes )')")
                AllCaseDataProcessingRuby.WriteLine("   end")

                AllCaseDataProcessingRuby.WriteLine("   empty_count = $current_case.count('mime-type:application/x-empty')")
                AllCaseDataProcessingRuby.WriteLine("   if empty_count == 0 ")
                AllCaseDataProcessingRuby.WriteLine("       empty_size = 0")
                AllCaseDataProcessingRuby.WriteLine("   else")
                AllCaseDataProcessingRuby.WriteLine("       empty_size = $current_case.getStatistics.getFileSize('mime-type:application/x-empty')")
                AllCaseDataProcessingRuby.WriteLine("   end")

                AllCaseDataProcessingRuby.WriteLine("   encrypted_count = $current_case.count('flag:encrypted')")
                AllCaseDataProcessingRuby.WriteLine("   if encrypted_count == 0 ")
                AllCaseDataProcessingRuby.WriteLine("       encrypted_size = 0")
                AllCaseDataProcessingRuby.WriteLine("   else")
                AllCaseDataProcessingRuby.WriteLine("       encrypted_size = $current_case.getStatistics.getFileSize('mime-type:application/x-empty')")
                AllCaseDataProcessingRuby.WriteLine("   end")

                AllCaseDataProcessingRuby.WriteLine("   decrypted_count = $current_case.count('flag:decrypted')")
                AllCaseDataProcessingRuby.WriteLine("   if decrypted_count == 0 ")
                AllCaseDataProcessingRuby.WriteLine("       decrypted_size = 0")
                AllCaseDataProcessingRuby.WriteLine("   else")
                AllCaseDataProcessingRuby.WriteLine("       decrypted_size = $current_case.getStatistics.getFileSize('flag:decrypted')")
                AllCaseDataProcessingRuby.WriteLine("   end")

                AllCaseDataProcessingRuby.WriteLine("   deleted_count = $current_case.count('flag:deleted')")
                AllCaseDataProcessingRuby.WriteLine("   if deleted_count == 0 ")
                AllCaseDataProcessingRuby.WriteLine("       deleted_size = 0")
                AllCaseDataProcessingRuby.WriteLine("   else")
                AllCaseDataProcessingRuby.WriteLine("       deleted_size = $current_case.getStatistics.getFileSize('flag:deleted')")
                AllCaseDataProcessingRuby.WriteLine("   end")

                AllCaseDataProcessingRuby.WriteLine("   corrupted_count = $current_case.count('properties:FailureDetail AND NOT flag:encrypted')")
                AllCaseDataProcessingRuby.WriteLine("   if corrupted_count == 0 ")
                AllCaseDataProcessingRuby.WriteLine("       corrupted_size = 0")
                AllCaseDataProcessingRuby.WriteLine("   else")
                AllCaseDataProcessingRuby.WriteLine("       corrupted_size = $current_case.getStatistics.getFileSize('properties:FailureDetail AND NOT flag:encrypted')")
                AllCaseDataProcessingRuby.WriteLine("   end")

                AllCaseDataProcessingRuby.WriteLine("   digest_mismatch_count = $current_case.count('flag:digest_mismatch')")
                AllCaseDataProcessingRuby.WriteLine("   if digest_mismatch_count  == 0 ")
                AllCaseDataProcessingRuby.WriteLine("       digest_mismatch_size = 0")
                AllCaseDataProcessingRuby.WriteLine("   else")
                AllCaseDataProcessingRuby.WriteLine("       digest_mismatch_size = $current_case.getStatistics.getFileSize('flag:digest_mismatch')")
                AllCaseDataProcessingRuby.WriteLine("   end")

                AllCaseDataProcessingRuby.WriteLine("   text_stripped_count = $current_case.count('flag:text_stripped')")
                AllCaseDataProcessingRuby.WriteLine("   if text_stripped_count == 0 ")
                AllCaseDataProcessingRuby.WriteLine("       text_stripped_size = 0")
                AllCaseDataProcessingRuby.WriteLine("   else")
                AllCaseDataProcessingRuby.WriteLine("       text_stripped_size = $current_case.getStatistics.getFileSize('flag:text_stripped')")
                AllCaseDataProcessingRuby.WriteLine("   end")

                AllCaseDataProcessingRuby.WriteLine("   text_not_indexed_count = $current_case.count('flag:text_not_indexed')")
                AllCaseDataProcessingRuby.WriteLine("   if text_not_indexed_count == 0 ")
                AllCaseDataProcessingRuby.WriteLine("       text_not_indexed_size = 0")
                AllCaseDataProcessingRuby.WriteLine("   else")
                AllCaseDataProcessingRuby.WriteLine("       text_not_indexed_size = $current_case.getStatistics.getFileSize('flag:text_not_indexed')")
                AllCaseDataProcessingRuby.WriteLine("   end")

                AllCaseDataProcessingRuby.WriteLine("   license_restricted_count = $current_case.count('flag:licence_restricted')")
                AllCaseDataProcessingRuby.WriteLine("   if license_restricted_count == 0 ")
                AllCaseDataProcessingRuby.WriteLine("       license_restricted_size = 0")
                AllCaseDataProcessingRuby.WriteLine("   else")
                AllCaseDataProcessingRuby.WriteLine("       license_restricted_size = $current_case.getStatistics.getFileSize('flag:licence_restricted')")
                AllCaseDataProcessingRuby.WriteLine("   end")

                AllCaseDataProcessingRuby.WriteLine("   not_processed_count = $current_case.count('flag:not_processed')")
                AllCaseDataProcessingRuby.WriteLine("   if not_processed_count == 0 ")
                AllCaseDataProcessingRuby.WriteLine("       not_processed_size = 0")
                AllCaseDataProcessingRuby.WriteLine("   else")
                AllCaseDataProcessingRuby.WriteLine("       not_processed_size = $current_case.getStatistics.getFileSize('flag:not_processed')")
                AllCaseDataProcessingRuby.WriteLine("   end")

                AllCaseDataProcessingRuby.WriteLine("   partially_processed_count = $current_case.count('flag:partially_processed')")
                AllCaseDataProcessingRuby.WriteLine("   if partially_processed_count == 0 ")
                AllCaseDataProcessingRuby.WriteLine("       partially_processed_size = 0")
                AllCaseDataProcessingRuby.WriteLine("   else")
                AllCaseDataProcessingRuby.WriteLine("       partially_processed_size = $current_case.getStatistics.getFileSize('flag:partially_processed')")
                AllCaseDataProcessingRuby.WriteLine("   end")

                AllCaseDataProcessingRuby.WriteLine("   text_not_processed_count = $current_case.count('flag:text_not_processed')")
                AllCaseDataProcessingRuby.WriteLine("   if text_not_processed_count == 0 ")
                AllCaseDataProcessingRuby.WriteLine("       text_not_processed_size = 0")
                AllCaseDataProcessingRuby.WriteLine("   else")
                AllCaseDataProcessingRuby.WriteLine("       text_not_processed_size = $current_case.getStatistics.getFileSize('flag:text_not_processed')")
                AllCaseDataProcessingRuby.WriteLine("   end")

                AllCaseDataProcessingRuby.WriteLine("   images_not_processed_count = $current_case.count('flag:images_not_processed')")
                AllCaseDataProcessingRuby.WriteLine("   if images_not_processed_count == 0 ")
                AllCaseDataProcessingRuby.WriteLine("       images_not_processed_size = 0")
                AllCaseDataProcessingRuby.WriteLine("   else")
                AllCaseDataProcessingRuby.WriteLine("       images_not_processed_size = $current_case.getStatistics.getFileSize('flag:images_not_processed')")
                AllCaseDataProcessingRuby.WriteLine("   end")

                AllCaseDataProcessingRuby.WriteLine("   reloaded_count = $current_case.count('flag:reloaded')")
                AllCaseDataProcessingRuby.WriteLine("   if reloaded_count == 0 ")
                AllCaseDataProcessingRuby.WriteLine("       reloaded_size = 0")
                AllCaseDataProcessingRuby.WriteLine("   else")
                AllCaseDataProcessingRuby.WriteLine("       reloaded_size = $current_case.getStatistics.getFileSize('flag:reloaded')")
                AllCaseDataProcessingRuby.WriteLine("   end")

                AllCaseDataProcessingRuby.WriteLine("   poisoned_count = $current_case.count('flag:poison')")
                AllCaseDataProcessingRuby.WriteLine("   if poisoned_count == 0 ")
                AllCaseDataProcessingRuby.WriteLine("       poisoned_size = 0")
                AllCaseDataProcessingRuby.WriteLine("   else")
                AllCaseDataProcessingRuby.WriteLine("       poisoned_size = $current_case.getStatistics.getFileSize('flag:poison')")
                AllCaseDataProcessingRuby.WriteLine("   end")

                AllCaseDataProcessingRuby.WriteLine("   slack_space_count = $current_case.count('flag:slack_space')")
                AllCaseDataProcessingRuby.WriteLine("   if slack_space_count == 0 ")
                AllCaseDataProcessingRuby.WriteLine("       slack_space_size = 0")
                AllCaseDataProcessingRuby.WriteLine("   else")
                AllCaseDataProcessingRuby.WriteLine("       slack_space_size = $current_case.getStatistics.getFileSize('flag:slack_space')")
                AllCaseDataProcessingRuby.WriteLine("   end")

                AllCaseDataProcessingRuby.WriteLine("   unallocated_space_count = $current_case.count('flag:unallocated_space')")
                AllCaseDataProcessingRuby.WriteLine("   if unallocated_space_count == 0 ")
                AllCaseDataProcessingRuby.WriteLine("       unallocated_space_size = 0")
                AllCaseDataProcessingRuby.WriteLine("   else")
                AllCaseDataProcessingRuby.WriteLine("       unallocated_space_size = $current_case.getStatistics.getFileSize('flag:unallocated_space')")
                AllCaseDataProcessingRuby.WriteLine("   end")

                AllCaseDataProcessingRuby.WriteLine("   manually_added_count = $current_case.count('flag:manually_added')")
                AllCaseDataProcessingRuby.WriteLine("   if manually_added_count == 0 ")
                AllCaseDataProcessingRuby.WriteLine("       manually_added_size = 0")
                AllCaseDataProcessingRuby.WriteLine("   else")
                AllCaseDataProcessingRuby.WriteLine("       manually_added_size = $current_case.getStatistics.getFileSize('flag:manually_added')")
                AllCaseDataProcessingRuby.WriteLine("   end")

                AllCaseDataProcessingRuby.WriteLine("   carved_count = $current_case.count('flag:carved')")
                AllCaseDataProcessingRuby.WriteLine("   if carved_count == 0 ")
                AllCaseDataProcessingRuby.WriteLine("       carved_size = 0")
                AllCaseDataProcessingRuby.WriteLine("   else")
                AllCaseDataProcessingRuby.WriteLine("       carved_size = $current_case.getStatistics.getFileSize('flag:carved')")
                AllCaseDataProcessingRuby.WriteLine("   end")

                AllCaseDataProcessingRuby.WriteLine("   fully_recovered_count = $current_case.count('flag:fully_recovered')")
                AllCaseDataProcessingRuby.WriteLine("   if fully_recovered_count == 0 ")
                AllCaseDataProcessingRuby.WriteLine("       fully_recovered_size = 0")
                AllCaseDataProcessingRuby.WriteLine("   else")
                AllCaseDataProcessingRuby.WriteLine("       fully_recovered_size = $current_case.getStatistics.getFileSize('flag:fully_recovered')")
                AllCaseDataProcessingRuby.WriteLine("   end")

                AllCaseDataProcessingRuby.WriteLine("   partially_recovered_count = $current_case.count('flag:partially_recovered')")
                AllCaseDataProcessingRuby.WriteLine("   if partially_recovered_count == 0 ")
                AllCaseDataProcessingRuby.WriteLine("       partially_recovered_size = 0")
                AllCaseDataProcessingRuby.WriteLine("   else")
                AllCaseDataProcessingRuby.WriteLine("       partially_recovered_size = $current_case.getStatistics.getFileSize('flag:partially_recovered')")
                AllCaseDataProcessingRuby.WriteLine("   end")

                AllCaseDataProcessingRuby.WriteLine("   metadata_recovered_count = $current_case.count('flag:metadata_recovered')")
                AllCaseDataProcessingRuby.WriteLine("   if metadata_recovered_count == 0 ")
                AllCaseDataProcessingRuby.WriteLine("       metadata_recovered_size = 0")
                AllCaseDataProcessingRuby.WriteLine("   else")
                AllCaseDataProcessingRuby.WriteLine("       metadata_recovered_size = $current_case.getStatistics.getFileSize('flag:metadata_recovered')")
                AllCaseDataProcessingRuby.WriteLine("   end")

                AllCaseDataProcessingRuby.WriteLine("   hidden_stream_count = $current_case.count('flag:hidden_stream')")
                AllCaseDataProcessingRuby.WriteLine("   if hidden_stream_count == 0 ")
                AllCaseDataProcessingRuby.WriteLine("       hidden_stream_size = 0")
                AllCaseDataProcessingRuby.WriteLine("   else")
                AllCaseDataProcessingRuby.WriteLine("       hidden_stream_size = $current_case.getStatistics.getFileSize('flag:hidden_stream')")
                AllCaseDataProcessingRuby.WriteLine("   end")

                AllCaseDataProcessingRuby.WriteLine("   irregular_items_count = 'Corrupted_container::' + " & """" & "#{corrupted_container_count}" & """" & ".to_s + " & """" & "::#{corrupted_container_size};" & """")
                AllCaseDataProcessingRuby.WriteLine("   irregular_items_count = irregular_items_count + 'Unsupported_Container::' + " & """" & "#{unsupported_container_count}" & """" & ".to_s + " & """" & "::#{unsupported_container_size};" & """")
                AllCaseDataProcessingRuby.WriteLine("   irregular_items_count = irregular_items_count + 'Nonsearchable_PDFs::' + " & """" & "#{nonsearchable_pdfs_count}" & """" & ".to_s + " & """" & "::#{nonsearchable_pdfs_size};" & """")
                AllCaseDataProcessingRuby.WriteLine("   irregular_items_count = irregular_items_count + 'Text_updated::' + " & """" & "#{text_updated_count}" & """" & ".to_s + " & """" & "::#{text_updated_size};" & """")
                AllCaseDataProcessingRuby.WriteLine("   irregular_items_count = irregular_items_count + 'Bad_extension::' + " & """" & "#{bad_extension_count}" & """" & ".to_s + " & """" & "::#{bad_extension_size};" & """")
                AllCaseDataProcessingRuby.WriteLine("   irregular_items_count = irregular_items_count + 'Unrecognized::' + " & """" & "#{unrecognized_count}" & """" & ".to_s + " & """" & "::#{unrecognized_size};" & """")
                AllCaseDataProcessingRuby.WriteLine("   irregular_items_count = irregular_items_count + 'Unsupported::' + " & """" & "#{unsupported_count}" & """" & ".to_s + " & """" & "::#{unsupported_size};" & """")
                AllCaseDataProcessingRuby.WriteLine("   irregular_items_count = irregular_items_count + 'Empty::' + " & """" & "#{empty_count}" & """" & ".to_s + " & """" & "::#{empty_size};" & """")
                AllCaseDataProcessingRuby.WriteLine("   irregular_items_count = irregular_items_count + 'Encrypted::' + " & """" & "#{encrypted_count}" & """" & ".to_s + " & """" & "::#{encrypted_size};" & """")
                AllCaseDataProcessingRuby.WriteLine("   irregular_items_count = irregular_items_count + 'Decrypted::' + " & """" & "#{decrypted_count}" & """" & ".to_s + " & """" & "::#{decrypted_size};" & """")
                AllCaseDataProcessingRuby.WriteLine("   irregular_items_count = irregular_items_count + 'Deleted::' + " & """" & "#{deleted_count}" & """" & ".to_s + " & """" & "::#{deleted_size};" & """")
                AllCaseDataProcessingRuby.WriteLine("   irregular_items_count = irregular_items_count + 'Corrupted::' + " & """" & "#{corrupted_count}" & """" & ".to_s + " & """" & "::#{corrupted_size};" & """")
                AllCaseDataProcessingRuby.WriteLine("   irregular_items_count = irregular_items_count + 'Digest_mismatch::' + " & """" & "#{digest_mismatch_count}" & """" & ".to_s + " & """" & "::#{digest_mismatch_size};" & """")
                AllCaseDataProcessingRuby.WriteLine("   irregular_items_count = irregular_items_count + 'Text_stripped::' + " & """" & "#{text_stripped_count}" & """" & ".to_s + " & """" & "::#{text_stripped_size};" & """")
                AllCaseDataProcessingRuby.WriteLine("   irregular_items_count = irregular_items_count + 'Text_Not_Indexed::' + " & """" & "#{text_not_indexed_count}" & """" & ".to_s + " & """" & "::#{text_not_indexed_size};" & """")
                AllCaseDataProcessingRuby.WriteLine("   irregular_items_count = irregular_items_count + 'License_restricted::' + " & """" & "#{license_restricted_count}" & """" & ".to_s + " & """" & "::#{license_restricted_size};" & """")
                AllCaseDataProcessingRuby.WriteLine("   irregular_items_count = irregular_items_count + 'Not_Processed::' + " & """" & "#{not_processed_count}" & """" & ".to_s + " & """" & "::#{not_processed_size};" & """")
                AllCaseDataProcessingRuby.WriteLine("   irregular_items_count = irregular_items_count + 'Partially_processed::' + " & """" & "#{partially_processed_count}" & """" & ".to_s + " & """" & "::#{partially_processed_size};" & """")
                AllCaseDataProcessingRuby.WriteLine("   irregular_items_count = irregular_items_count + 'Text_Not_Processed::' + " & """" & "#{text_not_processed_count}" & """" & ".to_s + " & """" & "::#{text_not_processed_size};" & """")
                AllCaseDataProcessingRuby.WriteLine("   irregular_items_count = irregular_items_count + 'Images_Not_Processed::' + " & """" & "#{images_not_processed_count}" & """" & ".to_s + " & """" & "::#{images_not_processed_size};" & """")
                AllCaseDataProcessingRuby.WriteLine("   irregular_items_count = irregular_items_count + 'Reload::' + " & """" & "#{reloaded_count}" & """" & ".to_s + " & """" & "::#{reloaded_size};" & """")
                AllCaseDataProcessingRuby.WriteLine("   irregular_items_count = irregular_items_count + 'Poisoned::' + " & """" & "#{poisoned_count}" & """" & ".to_s + " & """" & "::#{poisoned_size};" & """")
                AllCaseDataProcessingRuby.WriteLine("   irregular_items_count = irregular_items_count + 'Slack_space::' + " & """" & "#{slack_space_count}" & """" & ".to_s + " & """" & "::#{slack_space_size};" & """")
                AllCaseDataProcessingRuby.WriteLine("   irregular_items_count = irregular_items_count + 'Unallocated_Space::' + " & """" & "#{unallocated_space_count}" & """" & ".to_s + " & """" & "::#{unallocated_space_size};" & """")
                AllCaseDataProcessingRuby.WriteLine("   irregular_items_count = irregular_items_count + 'Manually_added::' + " & """" & "#{manually_added_count}" & """" & ".to_s + " & """" & "::#{manually_added_size};" & """")
                AllCaseDataProcessingRuby.WriteLine("   irregular_items_count = irregular_items_count + 'Carved::' + " & """" & "#{carved_count}" & """" & ".to_s + " & """" & "::#{carved_size};" & """" & ";")
                AllCaseDataProcessingRuby.WriteLine("   irregular_items_count = irregular_items_count + 'Fully_Recovered::' + " & """" & "#{fully_recovered_count}" & """" & ".to_s + " & """" & "::#{fully_recovered_size};" & """")
                AllCaseDataProcessingRuby.WriteLine("   irregular_items_count = irregular_items_count + 'Partially_Recovered::' + " & """" & "#{partially_recovered_count}" & """" & ".to_s + " & """" & "::#{partially_recovered_size};" & """")
                AllCaseDataProcessingRuby.WriteLine("   irregular_items_count = irregular_items_count + 'Metadata_Recovered::' + " & """" & "#{metadata_recovered_count}" & """" & ".to_s + " & """" & "::#{metadata_recovered_size};" & """")
                AllCaseDataProcessingRuby.WriteLine("   irregular_items_count = irregular_items_count + 'Hidden_Stream::' + " & """" & "#{hidden_stream_count}" & """" & ".to_s + " & """" & "::#{hidden_stream_size};" & """")

                AllCaseDataProcessingRuby.WriteLine("   case_irregularitems_data = [irregular_items_count, 90]")
                AllCaseDataProcessingRuby.WriteLine("   db.update(" & """" & "UPDATE NuixReportingInfo SET IrregularItems = ?, PercentComplete = ? WHERE CaseGUID = '#{case_guid}'" & """" & ", case_irregularitems_data)")
                AllCaseDataProcessingRuby.WriteLine("rescue")
                AllCaseDataProcessingRuby.WriteLine("   puts 'Error in Block 6'")
                AllCaseDataProcessingRuby.WriteLine("ensure")
                AllCaseDataProcessingRuby.WriteLine("")
                AllCaseDataProcessingRuby.WriteLine("end")
                AllCaseDataProcessingRuby.WriteLine("")

                AllCaseDataProcessingRuby.WriteLine("begin")
                AllCaseDataProcessingRuby.WriteLine("   userentered_all_custodians_info=''")
                AllCaseDataProcessingRuby.WriteLine("   userentered_search_hit_count=0")
                AllCaseDataProcessingRuby.WriteLine("   userentered_search_file_size=0")
                If Not sUserEnteredSearch = vbNullString Then
                    AllCaseDataProcessingRuby.WriteLine("   userentered_search_criteria = " & """" & sUserEnteredSearch & """")
                    AllCaseDataProcessingRuby.WriteLine("   userentered_search_hit_count = $current_case.count(userentered_search_criteria)")
                    AllCaseDataProcessingRuby.WriteLine("   userentered_search_file_size = $current_case.getStatistics.getFileSize(userentered_search_criteria)")
                    AllCaseDataProcessingRuby.WriteLine("   userentered_custodian_name_count=0")
                    AllCaseDataProcessingRuby.WriteLine("   custodian_names = $current_case.getAllCustodians")
                    AllCaseDataProcessingRuby.WriteLine("       custodian_names.each do |custodian_name|")
                    AllCaseDataProcessingRuby.WriteLine("		    search_criteria = " & """" & "custodian:\" & """" & "#{custodian_name}\" & """""" & " + " & """" & " and " & """" & " + " & """" & "\" & """" & sUserEnteredSearch & "\" & """""")
                    AllCaseDataProcessingRuby.WriteLine("		    userentered_custodian_name_count = $current_case.count(" & """" & "#{search_criteria}" & """" & ")")
                    AllCaseDataProcessingRuby.WriteLine("		    userentered_all_custodians_info = userentered_all_custodians_info + " & """" & "#{custodian_name}" & """" & " + '::' + " & """" & "#{userentered_custodian_name_count}" & """" & ".to_s + " & """" & ";" & """")
                    AllCaseDataProcessingRuby.WriteLine("       end")
                    AllCaseDataProcessingRuby.WriteLine("   user_search_details = [userentered_search_criteria, 80]")
                    AllCaseDataProcessingRuby.WriteLine("   db.update(" & """" & "UPDATE NuixReportingInfo SET SearchTerm = ?, PercentComplete = ? WHERE CaseGUID = '#{case_guid}'" & """" & ", user_search_details)")
                End If

                If Not sUserSearchFile = vbNullString Then
                    AllCaseDataProcessingRuby.WriteLine("   db.update(" & """" & "Delete from UCRTSearchTermResults where CaseGUID = '#{case_guid}'" & """" & ")")
                    AllCaseDataProcessingRuby.WriteLine("   sSearchTerm = ''")
                    AllCaseDataProcessingRuby.WriteLine("   sExportFolder = ''")
                    AllCaseDataProcessingRuby.WriteLine("   CSV.foreach(" & """" & sUserSearchFile.Replace("\", "\\") & """" & ") do |row|")
                    AllCaseDataProcessingRuby.WriteLine("       sSearchTerm = row[0]")
                    AllCaseDataProcessingRuby.WriteLine("       sExportFolder =  row[1]")
                    AllCaseDataProcessingRuby.WriteLine("       userentered_search_hit_count = $current_case.count(sSearchTerm)")
                    AllCaseDataProcessingRuby.WriteLine("       userentered_search_file_size = $current_case.getStatistics.getFileSize(sSearchTerm)")
                    AllCaseDataProcessingRuby.WriteLine("       user_search_details = [case_guid, sSearchTerm, userentered_search_hit_count]")
                    AllCaseDataProcessingRuby.WriteLine("       db.update(" & """" & "INSERT INTO UCRTSearchTermResults (CaseGUID, SearchTerm, ItemCount) VALUES ( ?, ?, ?)" & """" & ", user_search_details)")
                    AllCaseDataProcessingRuby.WriteLine("       userentered_custodian_name_count=0")
                    AllCaseDataProcessingRuby.WriteLine("       custodian_names = $current_case.getAllCustodians")
                    AllCaseDataProcessingRuby.WriteLine("       custodian_names.each do |custodian_name|")
                    AllCaseDataProcessingRuby.WriteLine("		    search_criteria = " & """" & "custodian:\" & """" & "#{custodian_name}\" & """""" & " + " & """" & " and " & """" & " + " & """" & "\" & """" & "#{sSearchTerm}" & "\" & """""")
                    AllCaseDataProcessingRuby.WriteLine("		    userentered_custodian_name_count = $current_case.count(" & """" & "#{search_criteria}" & """" & ")")
                    AllCaseDataProcessingRuby.WriteLine("		    userentered_all_custodians_info = userentered_all_custodians_info + " & """" & "#{custodian_name}" & """" & " + '::' + " & """" & "#{userentered_custodian_name_count}" & """" & ".to_s + " & """" & ";" & """")
                    AllCaseDataProcessingRuby.WriteLine("           user_search_details = [case_guid, search_criteria, userentered_search_hit_count, custodian_name]")
                    AllCaseDataProcessingRuby.WriteLine("           db.update(" & """" & "INSERT INTO UCRTSearchTermResults (CaseGUID, SearchTerm, Custodian, CustodianSearchTermItemCount) VALUES ( ?, ?, ?, ?)" & """" & ", user_search_details)")
                    AllCaseDataProcessingRuby.WriteLine("       end")
                    AllCaseDataProcessingRuby.WriteLine("   end")

                End If

                '        AllCaseDataProcessingRuby.WriteLine("totals_count = 'Total::' + " & """" & "#{total_count}" & """" & " + ';Email:' + " & """" & "#{email_total_count}" & """" & " + ';Calendar:' + " & """" & "#{calendar_total_count}" & """" & " + ';Contact:' + " & """" & "#{contact_total_count}" & """" & " + ';Document:' + " & """" & "#{document_total_count}" & """" & " + ';Spreadsheet:' + " & """" & "#{spreadsheet_total_count}" & """" & " + ';Presentation:' + " & """" & "#{presentation_total_count}" & """" & " + ';Image:' + " & """" & "#{image_total_count}" & """" & " + ';Drawing:' + " & """" & "#{drawing_total_count}" & """" & " + ';Other-Document:' + " & """" & "#{other_document_total_count}" & """" & " + ';Multimedia:' + " & """" & "#{multimedia_total_count}" & """" & " + ';Database:' + " & """" & "#{database_total_count}" & """" & " + ';Container:' + " & """" & "#{container_total_count}" & """" & " + ';System:' + " & """" & "#{system_total_count}" & """" & " + ';No-data:' + " & """" & "#{no_data_total_count}" & """" & " + ';Unrecognised:' + " & """" & "#{unrecognised_total_count}" & """" & " + ';Log:' + " & """" & "#{log_total_count}" & """")
                AllCaseDataProcessingRuby.WriteLine("   case_items_data = [userentered_all_custodians_info, userentered_search_hit_count, userentered_search_file_size, 85]")
                AllCaseDataProcessingRuby.WriteLine("   db.update(" & """" & "UPDATE NuixReportingInfo SET CustodianSearchHit = ?, HitCount = ?, SearchSize = ?, PercentComplete = ? WHERE CaseGUID = '#{case_guid}'" & """" & ", case_items_data)")
                AllCaseDataProcessingRuby.WriteLine("rescue")
                AllCaseDataProcessingRuby.WriteLine("   puts 'Error in Block 7'")
                AllCaseDataProcessingRuby.WriteLine("ensure")
                AllCaseDataProcessingRuby.WriteLine("")
                AllCaseDataProcessingRuby.WriteLine("end")
                AllCaseDataProcessingRuby.WriteLine("")

            End If

            If bExportSearchResults = True Then
                If sUserSearchFile <> "" Then
                    If sExportType = "Native" Then
                        AllCaseDataProcessingRuby.WriteLine("begin")
                        AllCaseDataProcessingRuby.WriteLine("   CSV.foreach(" & """" & sUserSearchFile.Replace("\", "\\") & """" & ") do |row|")
                        AllCaseDataProcessingRuby.WriteLine("       sSearchTerm = row[0]")
                        AllCaseDataProcessingRuby.WriteLine("       sExportFolder =  row[1]")
                        AllCaseDataProcessingRuby.WriteLine("       items = $current_case.search(sSearchTerm)")
                        AllCaseDataProcessingRuby.WriteLine("       rightnow = DateTime.now.to_s")
                        AllCaseDataProcessingRuby.WriteLine("       rightnow = rightnow.delete(':')")
                        AllCaseDataProcessingRuby.WriteLine("       exporter = $utilities.createBatchExporter('" & sExportDirectory.Replace("\", "\\") & "\\'" & " + sExportFolder " & " + '\\' +  rightnow + '\\' + case_name" & ")")
                        AllCaseDataProcessingRuby.WriteLine("       natives_settings = {")
                        AllCaseDataProcessingRuby.WriteLine("           :naming => " & """" & "guid" & """" & ",")
                        AllCaseDataProcessingRuby.WriteLine("           :path => " & """" & "NATIVE" & """" & ",")
                        AllCaseDataProcessingRuby.WriteLine("           :mailFormat => " & """" & "pst" & """" & ",")
                        AllCaseDataProcessingRuby.WriteLine("           :includeAttachments => true,")
                        AllCaseDataProcessingRuby.WriteLine("       }")
                        AllCaseDataProcessingRuby.WriteLine("       exporter.addProduct(" & """" & "native" & """" & ", natives_settings)")
                        AllCaseDataProcessingRuby.WriteLine("       exporter.setParallelProcessingSettings({")
                        AllCaseDataProcessingRuby.WriteLine("           :workerCount" & " => " & sExportWorkers & ",")
                        AllCaseDataProcessingRuby.WriteLine("           :workerMemory" & " => " & (CInt(sExportWorkerMemory) * 1024) & ",")
                        AllCaseDataProcessingRuby.WriteLine("           :workerTemp" & " => " & """" & "C:/Temp" & """" & ",")
                        AllCaseDataProcessingRuby.WriteLine("           :embedBroker" & " => true,")
                        AllCaseDataProcessingRuby.WriteLine("           :brokerMemory" & " => 768")
                        AllCaseDataProcessingRuby.WriteLine("       })")
                        AllCaseDataProcessingRuby.WriteLine("       exporter.exportItems(items)")
                        AllCaseDataProcessingRuby.WriteLine("   end")
                        AllCaseDataProcessingRuby.WriteLine("rescue")
                        AllCaseDataProcessingRuby.WriteLine("   puts 'Error in Block 8'")
                        AllCaseDataProcessingRuby.WriteLine("ensure")
                        AllCaseDataProcessingRuby.WriteLine("")
                        AllCaseDataProcessingRuby.WriteLine("end")
                    ElseIf sExportType = "PDF" Then
                        AllCaseDataProcessingRuby.WriteLine("begin")
                        AllCaseDataProcessingRuby.WriteLine("   CSV.foreach(" & """" & sUserSearchFile.Replace("\", "\\") & """" & ") do |row|")
                        AllCaseDataProcessingRuby.WriteLine("       sSearchTerm = row[0]")
                        AllCaseDataProcessingRuby.WriteLine("       sExportFolder =  row[1]")
                        AllCaseDataProcessingRuby.WriteLine("       items = $current_case.search(sSearchTerm)")
                        AllCaseDataProcessingRuby.WriteLine("       rightnow = DateTime.now.to_s")
                        AllCaseDataProcessingRuby.WriteLine("       rightnow = rightnow.delete(':')")
                        AllCaseDataProcessingRuby.WriteLine("       exporter = $utilities.createBatchExporter('" & sExportDirectory.Replace("\", "\\") & "\\'" & " + sExportFolder " & " + '\\' +  rightnow + '\\' + case_name" & ")")
                        AllCaseDataProcessingRuby.WriteLine("       pdf_settings = {")
                        AllCaseDataProcessingRuby.WriteLine("           :naming => " & """" & "guid" & """" & ",")
                        AllCaseDataProcessingRuby.WriteLine("           :path => " & """" & "PDF" & """" & ",")
                        AllCaseDataProcessingRuby.WriteLine("           :mailFormat => " & """" & "pst" & """" & ",")
                        AllCaseDataProcessingRuby.WriteLine("           :includeAttachments => true,")
                        AllCaseDataProcessingRuby.WriteLine("       }")
                        AllCaseDataProcessingRuby.WriteLine("       exporter.addProduct(" & """" & "pdf" & """" & ", pdf_settings)")
                        AllCaseDataProcessingRuby.WriteLine("       exporter.setParallelProcessingSettings({")
                        AllCaseDataProcessingRuby.WriteLine("           :workerCount" & " => " & sExportWorkers & ",")
                        AllCaseDataProcessingRuby.WriteLine("           :workerMemory" & " => " & (CInt(sExportWorkerMemory) * 1024) & ",")
                        AllCaseDataProcessingRuby.WriteLine("           :workerTemp" & " => " & """" & "C:/Temp" & """" & ",")
                        AllCaseDataProcessingRuby.WriteLine("           :embedBroker" & " => true,")
                        AllCaseDataProcessingRuby.WriteLine("           :brokerMemory" & " => 768")
                        AllCaseDataProcessingRuby.WriteLine("       })")
                        AllCaseDataProcessingRuby.WriteLine("       exporter.exportItems(items)")
                        AllCaseDataProcessingRuby.WriteLine("   end")
                        AllCaseDataProcessingRuby.WriteLine("rescue")
                        AllCaseDataProcessingRuby.WriteLine("   puts 'Error in Block 8'")
                        AllCaseDataProcessingRuby.WriteLine("ensure")
                        AllCaseDataProcessingRuby.WriteLine("")
                        AllCaseDataProcessingRuby.WriteLine("end")

                    ElseIf sExportType = "NLI" Then
                        AllCaseDataProcessingRuby.WriteLine("begin")
                        AllCaseDataProcessingRuby.WriteLine("   CSV.foreach(" & """" & sUserSearchFile.Replace("\", "\\") & """" & ") do |row|")
                        AllCaseDataProcessingRuby.WriteLine("       sSearchTerm = row[0]")
                        AllCaseDataProcessingRuby.WriteLine("       sExportFolder =  row[1]")
                        AllCaseDataProcessingRuby.WriteLine("       items = $current_case.search(sSearchTerm)")
                        AllCaseDataProcessingRuby.WriteLine("       natives_settings = {")
                        AllCaseDataProcessingRuby.WriteLine("       }")
                        AllCaseDataProcessingRuby.WriteLine("       export_settings = {")
                        AllCaseDataProcessingRuby.WriteLine("       }")
                        AllCaseDataProcessingRuby.WriteLine("       rightnow = DateTime.now.to_s")
                        AllCaseDataProcessingRuby.WriteLine("       rightnow = rightnow.delete(':')")
                        AllCaseDataProcessingRuby.WriteLine("       directory_name = sExportFolder " & " + '\\' +  rightnow + '\\'")
                        AllCaseDataProcessingRuby.WriteLine("       response = FileUtils.mkdir_p(directory_name)")
                        AllCaseDataProcessingRuby.WriteLine("       exporter = $utilities.createLogicalImageExporter(directory_name, case_name, export_settings)")
                        '                    AllCaseDataProcessingRuby.WriteLine("       exporter = $utilities.createLogicalImageExporter('C:\\Users\\CCarlson01<\\Documents\\Nuix\\ISS Governance\\ExportLocation\\NLI\\', '#{case_name}',export_settings)")
                        AllCaseDataProcessingRuby.WriteLine("")
                        AllCaseDataProcessingRuby.WriteLine("       items.each do |item|")
                        AllCaseDataProcessingRuby.WriteLine("           exporter.addItem(item)")
                        AllCaseDataProcessingRuby.WriteLine("       end")
                        AllCaseDataProcessingRuby.WriteLine("   end")
                        AllCaseDataProcessingRuby.WriteLine("   exporter.close")
                        AllCaseDataProcessingRuby.WriteLine("rescue")
                        AllCaseDataProcessingRuby.WriteLine("   puts 'Error in Block 8'")
                        AllCaseDataProcessingRuby.WriteLine("ensure")
                        AllCaseDataProcessingRuby.WriteLine("")
                        AllCaseDataProcessingRuby.WriteLine("end")
                        AllCaseDataProcessingRuby.WriteLine("")
                    ElseIf sExportType = "Mailbox" Then
                        AllCaseDataProcessingRuby.WriteLine("begin")
                        AllCaseDataProcessingRuby.WriteLine("   CSV.foreach(" & """" & sUserSearchFile.Replace("\", "\\") & """" & ") do |row|")
                        AllCaseDataProcessingRuby.WriteLine("       sSearchTerm = row[0]")
                        AllCaseDataProcessingRuby.WriteLine("       sExportFolder =  row[1]")
                        AllCaseDataProcessingRuby.WriteLine("       items = $current_case.search(sSearchTerm)")
                        AllCaseDataProcessingRuby.WriteLine("       rightnow = DateTime.now.to_s")
                        AllCaseDataProcessingRuby.WriteLine("       rightnow = rightnow.delete(':')")
                        AllCaseDataProcessingRuby.WriteLine("       exporter = utilities.getMailboxExporter()")
                        AllCaseDataProcessingRuby.WriteLine("       natives_settings = {")
                        AllCaseDataProcessingRuby.WriteLine("       	:format => " & """" & "pst" & """" & ",")
                        AllCaseDataProcessingRuby.WriteLine("       	:path => nil,")
                        AllCaseDataProcessingRuby.WriteLine("       	:failfast => " & """" & "false" & """")
                        AllCaseDataProcessingRuby.WriteLine("       }")

                        AllCaseDataProcessingRuby.WriteLine("   directory_name = '" & sExportDirectory.Replace("\", "\\") & "\\'" & " + sExportFolder " & " + '\\' +  rightnow + '\\'")
                        AllCaseDataProcessingRuby.WriteLine("   response = FileUtils.mkdir_p(directory_name)")
                        AllCaseDataProcessingRuby.WriteLine("       exporter.exportItems(items, '" & sExportDirectory.Replace("\", "\\") & "\\'" & " + sExportFolder " & " + '\\' +  rightnow + '\\' + case_name + '.pst'" & ", natives_settings)")
                        AllCaseDataProcessingRuby.WriteLine("   end")
                        AllCaseDataProcessingRuby.WriteLine("rescue")
                        AllCaseDataProcessingRuby.WriteLine("   puts 'Error in Block 8'")
                        AllCaseDataProcessingRuby.WriteLine("ensure")
                        AllCaseDataProcessingRuby.WriteLine("")
                        AllCaseDataProcessingRuby.WriteLine("end")

                    ElseIf sExportType = "Case Subset" Then
                        AllCaseDataProcessingRuby.WriteLine("begin")
                        AllCaseDataProcessingRuby.WriteLine("   CSV.foreach(" & """" & sUserSearchFile.Replace("\", "\\") & """" & ") do |row|")
                        AllCaseDataProcessingRuby.WriteLine("       sSearchTerm = row[0]")
                        AllCaseDataProcessingRuby.WriteLine("       sExportFolder =  row[1]")
                        AllCaseDataProcessingRuby.WriteLine("       items = $current_case.search(sSearchTerm)")
                        AllCaseDataProcessingRuby.WriteLine("       rightnow = DateTime.now.to_s")
                        AllCaseDataProcessingRuby.WriteLine("       rightnow = rightnow.delete(':')")
                        AllCaseDataProcessingRuby.WriteLine("       exporter = $utilities.getCaseSubsetExporter()")
                        AllCaseDataProcessingRuby.WriteLine("       natives_settings = {")
                        AllCaseDataProcessingRuby.WriteLine("           :naming => " & """" & "guid" & """" & ",")
                        AllCaseDataProcessingRuby.WriteLine("           :path => " & """" & "NATIVE" & """" & ",")
                        AllCaseDataProcessingRuby.WriteLine("           :mailFormat => " & """" & "pst" & """" & ",")
                        AllCaseDataProcessingRuby.WriteLine("           :includeAttachments => true,")
                        AllCaseDataProcessingRuby.WriteLine("       }")
                        AllCaseDataProcessingRuby.WriteLine("       case_subset_settings = {")
                        AllCaseDataProcessingRuby.WriteLine("           :evidenceStoreCount => 1,")
                        AllCaseDataProcessingRuby.WriteLine("           :includeFamilies => true,")
                        AllCaseDataProcessingRuby.WriteLine("           :copyTags => true,")
                        AllCaseDataProcessingRuby.WriteLine("           :copyComments => true,")
                        AllCaseDataProcessingRuby.WriteLine("           :copyCustodians => true,")
                        AllCaseDataProcessingRuby.WriteLine("           :copyItemSets => true,")
                        AllCaseDataProcessingRuby.WriteLine("           :copyClassifiers => true,")
                        AllCaseDataProcessingRuby.WriteLine("           :copyMarkupSets => true,")
                        AllCaseDataProcessingRuby.WriteLine("           :copyProductionSets => true,")
                        AllCaseDataProcessingRuby.WriteLine("           :copyClusters => true,")
                        AllCaseDataProcessingRuby.WriteLine("           :copyCustomMetadata => true,")
                        AllCaseDataProcessingRuby.WriteLine("           :copyGraphDatabase => false,")
                        AllCaseDataProcessingRuby.WriteLine("           :caseMetadata => nil,")
                        AllCaseDataProcessingRuby.WriteLine("           :processingSettings => nil")
                        AllCaseDataProcessingRuby.WriteLine("       }")
                        AllCaseDataProcessingRuby.WriteLine("   directory_name = '" & sExportDirectory.Replace("\", "\\") & "\\'" & " + sExportFolder " & " + '\\' +  rightnow + '\\'")
                        AllCaseDataProcessingRuby.WriteLine("   response = FileUtils.mkdir_p(directory_name)")
                        AllCaseDataProcessingRuby.WriteLine("   exporter.exportItems(items,'" & sExportDirectory.Replace("\", "\\") & "\\'" & "  + '\\' + sExportFolder  + '\\' +  case_name + '\\' + rightnow  + '\\CaseSubset'',case_subset_settings" & ")")
                        AllCaseDataProcessingRuby.WriteLine("   end")
                        AllCaseDataProcessingRuby.WriteLine("rescue")
                        AllCaseDataProcessingRuby.WriteLine("   puts 'Error in Block 8'")
                        AllCaseDataProcessingRuby.WriteLine("ensure")
                        AllCaseDataProcessingRuby.WriteLine("")
                        AllCaseDataProcessingRuby.WriteLine("end")
                        AllCaseDataProcessingRuby.WriteLine("")
                    End If
                    AllCaseDataProcessingRuby.WriteLine("")
                ElseIf sUserEnteredSearch <> "" Then

                    If sExportType = "Native" Then
                        AllCaseDataProcessingRuby.WriteLine("begin")
                        AllCaseDataProcessingRuby.WriteLine("   sSearchTerm = '" & sUserEnteredSearch & "'")
                        AllCaseDataProcessingRuby.WriteLine("   sExportFolder =  '" & sExportDirectory.Replace("\", "\\") & "'")
                        AllCaseDataProcessingRuby.WriteLine("   items = $current_case.search(sSearchTerm)")
                        AllCaseDataProcessingRuby.WriteLine("   rightnow = DateTime.now.to_s")
                        AllCaseDataProcessingRuby.WriteLine("   rightnow = rightnow.delete(':')")
                        AllCaseDataProcessingRuby.WriteLine("   exporter = $utilities.createBatchExporter(sExportFolder " & " + '\\' +  rightnow + '\\' + case_name" & ")")
                        AllCaseDataProcessingRuby.WriteLine("   natives_settings = {")
                        AllCaseDataProcessingRuby.WriteLine("       :naming => " & """" & "guid" & """" & ",")
                        AllCaseDataProcessingRuby.WriteLine("       :path => " & """" & "NATIVE" & """" & ",")
                        AllCaseDataProcessingRuby.WriteLine("       :mailFormat => " & """" & "pst" & """" & ",")
                        AllCaseDataProcessingRuby.WriteLine("       :includeAttachments => true,")
                        AllCaseDataProcessingRuby.WriteLine("       }")
                        AllCaseDataProcessingRuby.WriteLine("   exporter.addProduct(" & """" & "native" & """" & ", natives_settings)")
                        AllCaseDataProcessingRuby.WriteLine("   exporter.setParallelProcessingSettings({")
                        AllCaseDataProcessingRuby.WriteLine("       :workerCount" & " => " & sExportWorkers & ",")
                        AllCaseDataProcessingRuby.WriteLine("       :workerMemory" & " => " & (CInt(sExportWorkerMemory) * 1024) & ",")
                        AllCaseDataProcessingRuby.WriteLine("       :workerTemp" & " => " & """" & "C:/Temp" & """" & ",")
                        AllCaseDataProcessingRuby.WriteLine("       :embedBroker" & " => true,")
                        AllCaseDataProcessingRuby.WriteLine("       :brokerMemory" & " => 768")
                        AllCaseDataProcessingRuby.WriteLine("       })")
                        AllCaseDataProcessingRuby.WriteLine("   exporter.exportItems(items)")
                        AllCaseDataProcessingRuby.WriteLine("rescue")
                        AllCaseDataProcessingRuby.WriteLine("   puts 'Error in Block 8'")
                        AllCaseDataProcessingRuby.WriteLine("ensure")
                        AllCaseDataProcessingRuby.WriteLine("")
                        AllCaseDataProcessingRuby.WriteLine("end")
                    ElseIf sExportType = "PDF" Then
                        AllCaseDataProcessingRuby.WriteLine("begin")
                        AllCaseDataProcessingRuby.WriteLine("   sSearchTerm = '" & sUserEnteredSearch & "'")
                        AllCaseDataProcessingRuby.WriteLine("   sExportFolder =  '" & sExportDirectory.Replace("\", "\\") & "'")
                        AllCaseDataProcessingRuby.WriteLine("   items = $current_case.search(sSearchTerm)")
                        AllCaseDataProcessingRuby.WriteLine("   rightnow = DateTime.now.to_s")
                        AllCaseDataProcessingRuby.WriteLine("   rightnow = rightnow.delete(':')")
                        AllCaseDataProcessingRuby.WriteLine("   exporter = $utilities.createBatchExporter(sExportFolder " & " + '\\' +  rightnow + '\\' + case_name" & ")")
                        AllCaseDataProcessingRuby.WriteLine("   pdf_settings = {")
                        AllCaseDataProcessingRuby.WriteLine("       :naming => " & """" & "guid" & """" & ",")
                        AllCaseDataProcessingRuby.WriteLine("       :path => " & """" & "PDF" & """" & ",")
                        AllCaseDataProcessingRuby.WriteLine("       :mailFormat => " & """" & "pst" & """" & ",")
                        AllCaseDataProcessingRuby.WriteLine("       :includeAttachments => true,")
                        AllCaseDataProcessingRuby.WriteLine("       }")
                        AllCaseDataProcessingRuby.WriteLine("       exporter.addProduct(" & """" & "pdf" & """" & ", pdf_settings)")
                        AllCaseDataProcessingRuby.WriteLine("       exporter.setParallelProcessingSettings({")
                        AllCaseDataProcessingRuby.WriteLine("           :workerCount" & " => " & sExportWorkers & ",")
                        AllCaseDataProcessingRuby.WriteLine("           :workerMemory" & " => " & (CInt(sExportWorkerMemory) * 1024) & ",")
                        AllCaseDataProcessingRuby.WriteLine("           :workerTemp" & " => " & """" & "C:/Temp" & """" & ",")
                        AllCaseDataProcessingRuby.WriteLine("           :embedBroker" & " => true,")
                        AllCaseDataProcessingRuby.WriteLine("           :brokerMemory" & " => 768")
                        AllCaseDataProcessingRuby.WriteLine("       })")
                        AllCaseDataProcessingRuby.WriteLine("       exporter.exportItems(items)")
                        AllCaseDataProcessingRuby.WriteLine("rescue")
                        AllCaseDataProcessingRuby.WriteLine("   puts 'Error in Block 8'")
                        AllCaseDataProcessingRuby.WriteLine("ensure")
                        AllCaseDataProcessingRuby.WriteLine("")
                        AllCaseDataProcessingRuby.WriteLine("end")
                    ElseIf sExportType = "NLI" Then
                        AllCaseDataProcessingRuby.WriteLine("begin")
                        AllCaseDataProcessingRuby.WriteLine("   sSearchTerm = '" & sUserEnteredSearch & "'")
                        AllCaseDataProcessingRuby.WriteLine("   sExportFolder =  '" & sExportDirectory.Replace("\", "\\") & "'")
                        AllCaseDataProcessingRuby.WriteLine("   items = $current_case.search(sSearchTerm)")
                        AllCaseDataProcessingRuby.WriteLine("   natives_settings = {")
                        AllCaseDataProcessingRuby.WriteLine("   }")
                        AllCaseDataProcessingRuby.WriteLine("   export_settings = {")
                        AllCaseDataProcessingRuby.WriteLine("   }")
                        AllCaseDataProcessingRuby.WriteLine("   rightnow = DateTime.now.to_s")
                        AllCaseDataProcessingRuby.WriteLine("   rightnow = rightnow.delete(':')")
                        AllCaseDataProcessingRuby.WriteLine("   directory_name = sExportFolder " & " + '\\' +  rightnow + '\\'")
                        AllCaseDataProcessingRuby.WriteLine("   response = FileUtils.mkdir_p(directory_name)")
                        AllCaseDataProcessingRuby.WriteLine("   exporter = $utilities.createLogicalImageExporter(directory_name, case_name,export_settings)")
                        AllCaseDataProcessingRuby.WriteLine("")
                        AllCaseDataProcessingRuby.WriteLine("   items.each do |item|")
                        AllCaseDataProcessingRuby.WriteLine("       exporter.addItem(item)")
                        AllCaseDataProcessingRuby.WriteLine("   end")
                        AllCaseDataProcessingRuby.WriteLine("   exporter.close")
                        AllCaseDataProcessingRuby.WriteLine("rescue")
                        AllCaseDataProcessingRuby.WriteLine("   puts 'Error in Block 8'")
                        AllCaseDataProcessingRuby.WriteLine("ensure")
                        AllCaseDataProcessingRuby.WriteLine("")
                        AllCaseDataProcessingRuby.WriteLine("end")
                        AllCaseDataProcessingRuby.WriteLine("")
                    ElseIf sExportType = "Mailbox" Then
                        AllCaseDataProcessingRuby.WriteLine("begin")
                        AllCaseDataProcessingRuby.WriteLine("   sSearchTerm = '" & sUserEnteredSearch & "'")
                        AllCaseDataProcessingRuby.WriteLine("   sExportFolder =  '" & sExportDirectory.Replace("\", "\\") & "'")
                        AllCaseDataProcessingRuby.WriteLine("   items = $current_case.search(sSearchTerm)")
                        AllCaseDataProcessingRuby.WriteLine("   rightnow = DateTime.now.to_s")
                        AllCaseDataProcessingRuby.WriteLine("   rightnow = rightnow.delete(':')")
                        AllCaseDataProcessingRuby.WriteLine("   exporter = utilities.getMailboxExporter()")
                        AllCaseDataProcessingRuby.WriteLine("   natives_settings = {")
                        AllCaseDataProcessingRuby.WriteLine("      :format => " & """" & "pst" & """" & ",")
                        AllCaseDataProcessingRuby.WriteLine("       :path => nil,")
                        AllCaseDataProcessingRuby.WriteLine("       :failfast => " & """" & "false" & """")
                        AllCaseDataProcessingRuby.WriteLine("   }")
                        AllCaseDataProcessingRuby.WriteLine("   directory_name = sExportFolder " & " + '\\' +  rightnow + '\\'")
                        AllCaseDataProcessingRuby.WriteLine("   response = FileUtils.mkdir_p(directory_name)")
                        AllCaseDataProcessingRuby.WriteLine("       exporter.exportItems(items, sExportFolder " & " + '\\' +  rightnow + '\\' + case_name + '.pst'" & ", natives_settings)")
                        AllCaseDataProcessingRuby.WriteLine("rescue")
                        AllCaseDataProcessingRuby.WriteLine("   puts 'Error in Block 8'")
                        AllCaseDataProcessingRuby.WriteLine("ensure")
                        AllCaseDataProcessingRuby.WriteLine("")
                        AllCaseDataProcessingRuby.WriteLine("end")
                    ElseIf sExportType = "Case Subset" Then
                        AllCaseDataProcessingRuby.WriteLine("begin")
                        AllCaseDataProcessingRuby.WriteLine("       sSearchTerm = '" & sUserEnteredSearch & "'")
                        AllCaseDataProcessingRuby.WriteLine("       sExportFolder =  '" & sExportDirectory.Replace("\", "\\") & "'")
                        AllCaseDataProcessingRuby.WriteLine("       items = $current_case.search(sSearchTerm)")
                        AllCaseDataProcessingRuby.WriteLine("       rightnow = DateTime.now.to_s")
                        AllCaseDataProcessingRuby.WriteLine("       rightnow = rightnow.delete(':')")
                        AllCaseDataProcessingRuby.WriteLine("       exporter = $utilities.getCaseSubsetExporter()")
                        AllCaseDataProcessingRuby.WriteLine("       natives_settings = {")
                        AllCaseDataProcessingRuby.WriteLine("           :naming => " & """" & "guid" & """" & ",")
                        AllCaseDataProcessingRuby.WriteLine("           :path => " & """" & "NATIVE" & """" & ",")
                        AllCaseDataProcessingRuby.WriteLine("           :mailFormat => " & """" & "pst" & """" & ",")
                        AllCaseDataProcessingRuby.WriteLine("           :includeAttachments => true,")
                        AllCaseDataProcessingRuby.WriteLine("       }")
                        AllCaseDataProcessingRuby.WriteLine("       case_subset_settings = {")
                        AllCaseDataProcessingRuby.WriteLine("           :evidenceStoreCount => 1,")
                        AllCaseDataProcessingRuby.WriteLine("           :includeFamilies => true,")
                        AllCaseDataProcessingRuby.WriteLine("           :copyTags => true,")
                        AllCaseDataProcessingRuby.WriteLine("           :copyComments => true,")
                        AllCaseDataProcessingRuby.WriteLine("           :copyCustodians => true,")
                        AllCaseDataProcessingRuby.WriteLine("           :copyItemSets => true,")
                        AllCaseDataProcessingRuby.WriteLine("           :copyClassifiers => true,")
                        AllCaseDataProcessingRuby.WriteLine("           :copyMarkupSets => true,")
                        AllCaseDataProcessingRuby.WriteLine("           :copyProductionSets => true,")
                        AllCaseDataProcessingRuby.WriteLine("           :copyClusters => true,")
                        AllCaseDataProcessingRuby.WriteLine("           :copyCustomMetadata => true,")
                        AllCaseDataProcessingRuby.WriteLine("           :copyGraphDatabase => false,")
                        AllCaseDataProcessingRuby.WriteLine("           :caseMetadata => nil,")
                        AllCaseDataProcessingRuby.WriteLine("           :processingSettings => nil")
                        AllCaseDataProcessingRuby.WriteLine("       }")
                        AllCaseDataProcessingRuby.WriteLine("   directory_name = 'sExportFolder " & " + '\\' +  rightnow + '\\'")
                        AllCaseDataProcessingRuby.WriteLine("   response = FileUtils.mkdir_p(directory_name)")
                        AllCaseDataProcessingRuby.WriteLine("   exporter.exportItems(items, sExportFolder  + '\\' +  case_name + '\\' + rightnow  + '\\CaseSubset',case_subset_settings" & ")")
                        AllCaseDataProcessingRuby.WriteLine("rescue")
                        AllCaseDataProcessingRuby.WriteLine("   puts 'Error in Block 8'")
                        AllCaseDataProcessingRuby.WriteLine("ensure")
                        AllCaseDataProcessingRuby.WriteLine("")
                        AllCaseDataProcessingRuby.WriteLine("end")
                        AllCaseDataProcessingRuby.WriteLine("")
                    End If
                End If

            End If
            AllCaseDataProcessingRuby.WriteLine("begin")
            AllCaseDataProcessingRuby.WriteLine("   processing_data = [100]")
            AllCaseDataProcessingRuby.WriteLine("   db.update(" & """" & "UPDATE NuixReportingInfo SET PercentComplete = ? WHERE CaseGUID = '#{case_guid}'" & """" & ", processing_data)")

            AllCaseDataProcessingRuby.WriteLine("   $current_case.close")
            AllCaseDataProcessingRuby.WriteLine("rescue")
            AllCaseDataProcessingRuby.WriteLine("   puts 'Error in Block 10'")
            AllCaseDataProcessingRuby.WriteLine("ensure")
            AllCaseDataProcessingRuby.WriteLine("")
            AllCaseDataProcessingRuby.WriteLine("end")
            AllCaseDataProcessingRuby.WriteLine("")

        End If
        AllCaseDataProcessingRuby.Close()

        blnBuildUpdatedAllCaseDataProcessingRuby = True
    End Function


    Private Sub treeViewFolders_AfterCheck(sender As Object, e As TreeViewEventArgs) Handles treeViewFolders.AfterCheck
        If e.Node.Checked = True Then
            plstSoureFolders.Add(e.Node.Tag)
        Else
            plstSoureFolders.Remove(e.Node.Tag)
        End If
    End Sub

    Private Sub treeViewFolders_AfterExpand(sender As Object, e As TreeViewEventArgs) Handles treeViewFolders.AfterExpand
        If e.Node.Nodes.Count = 1 AndAlso e.Node.Nodes(0).Text = "{child}" Then
            e.Node.Nodes.Clear()
            AddAllFolders(e.Node, CStr(e.Node.Tag))
            e.Node.SelectedImageIndex = 1
        End If
    End Sub


    Private Sub treeViewFolders_BeforeExpand(sender As Object, e As TreeViewCancelEventArgs) Handles treeViewFolders.BeforeExpand
        If e.Node.Nodes.Count = 1 AndAlso e.Node.Nodes(0).Text = "{child}" Then
            e.Node.Nodes.Clear()
            AddAllFolders(e.Node, CStr(e.Node.Tag))
        End If
    End Sub

    Private Sub btnExportCaseStatistics_Click(sender As Object, e As EventArgs) Handles btnExportCaseStatistics.Click
        ExportContextMenuStrip.Show(btnExportCaseStatistics, 0, btnExportCaseStatistics.Height)
    End Sub

    Private Sub CSVToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles CSVToolStripMenuItem.Click

        Dim sReportFilePath As String
        Dim ReportOutputFile As StreamWriter
        Dim sReportType As String
        Dim sOutputFileName As String
        Dim sMachineName As String

        Dim sCaseGuid As String
        Dim sCollectionStatus As String
        Dim sPercentComplete As String
        Dim sReportLoadDuration As String
        Dim sCaseName As String
        Dim sCurrentCaseVersion As String
        Dim sUpgradedCaseVersion As String
        Dim sBatchLoadInfo As String
        Dim sCaseLocation As String
        Dim sCaseDescription As String
        Dim sCaseSizeOnDisk As String
        Dim sCaseFileSize As String
        Dim sCaseAuditSize As String
        Dim sOldestTopLevel As String
        Dim sNewestTopLevel As String
        Dim sIsCompound As String
        Dim sCasesContained As String
        Dim sContainedInCase As String
        Dim sInvestigator As String
        Dim sInvestigatorSessions As String
        Dim sInvalidSessions As String
        Dim sInvestigatorTimeSummary As String
        Dim sBrokerMemory As String
        Dim sWorkerCount As String
        Dim sWorkerMemory As String
        Dim sEvidenceName As String
        Dim sEvidenceLocation As String
        Dim sEvidenceDescription As String
        Dim sEvidenceCustomMetadata As String
        Dim sLanguagesContained As String
        Dim sMimeTypes As String
        Dim sItemTypes As String
        Dim sIrregularItems As String
        Dim sCreationDate As String
        Dim sModifiedDate As String
        Dim sLoadStartDate As String
        Dim sLoadEndDate As String
        Dim sLoadTime As String
        Dim sLoadEvents As String
        Dim sTotalLoadTime As String
        Dim sProcessingSpeed As String
        Dim sTotalCaseItemCount As String
        Dim sDuplicateItems As String
        Dim sCustodians As String
        Dim sCustodianCount As String
        Dim sSearchTerm As String
        Dim sSearchSize As String
        Dim sSearchHitCount As String
        Dim sCustodianSearchHit As String
        Dim sHitCountPercent As String
        Dim sNuixLogLocation As String
        Dim sDataExport As String

        If cboReportType.Text = vbNullString Then
            MessageBox.Show("You must enter a report type to export the data for.", "No Report Type selected")
            cboReportType.Focus()
            Exit Sub
        End If

        sReportFilePath = txtReportLocation.Text
        sReportType = cboReportType.Text
        sMachineName = System.Net.Dns.GetHostName()
        sOutputFileName = sMachineName & "-" & sReportType & "-" & DateTime.Now.ToString("yyyy-MM-dd hh-mm-ss") & ".csv"
        ReportOutputFile = New StreamWriter(sReportFilePath & "\" & sOutputFileName)

        Select Case cboReportType.Text
            Case "All"
                ReportOutputFile.WriteLine("CaseGUID, Collection Status,Percent Complete, Report Load Duration,Case Name,Current Case Version,Upgraded Case Version,Case Location,Batch Load Info,Data Export,Case Description,Case Size On Disk,Case File Size,Case Audit Size,Oldest Item Date,Newest Item Date,Is Compound,Cases Contained, Contained In Case,Investigator,Investigator Sessions, Invalid Sessions, Investigator Time Summary, Broker Memory,Worker Count,Worker Memory,Evidence Name,Evidence Location,Evidence Description,Evidence Custom Metadata,Languages,Mime Types,Item Types,Irregular Items,Creation Date,Modified Date,Load Start,Load End,Load Events,Total Load Time,Processing Speed,Total Item Count, Duplicate Item Count,Custodians,Custodian Count,Search Term,Search Size,Search Hit Count,Custodian Search Hit,Hit Count Percent, Nuix Log Location")
            Case "App Memory per case"
                ReportOutputFile.WriteLine("CaseGUID, Collection Status,Report Load Duration,Case Name,Current Case Version,Upgraded Case Version,Case Location,Batch Load Info,Data Export,Case Size On Disk,Case File Size,Case Audit Size,Oldest Item Date,Newest Item Date,Is Compound,Cases Contained, Contained In Case,Broker Memory,Worker Count,Worker Memory")
            Case "Case by Investigator"
                ReportOutputFile.WriteLine("CaseGUID, Collection Status,Report Load Duration,Case Name,Current Case Version,Upgraded Case Version,Case Location,Batch Load Info,Data Export,Case Size On Disk,Case File Size,Case Audit Size,Oldest Item Date,Newest Item Date,Is Compound,Cases Contained, Contained In Case,Investigator")
            Case "Case Evidence"
                ReportOutputFile.WriteLine("CaseGUID, Collection Status,Report Load Duration,Case Name,Current Case Version,Upgraded Case Version,Case Location,Batch Load Info,Data Export,Case Size On Disk,Case File Size,Case Audit Size,Oldest Item Date,Newest Item Date,Is Compound,Cases Contained, Contained In Case,Evidence Name,Evidence Location")
            Case "Case Location"
                ReportOutputFile.WriteLine("CaseGUID, Collection Status,Report Load Duration,Case Name,Current Case Version,Upgraded Case Version,Case Location,Batch Load Info,Data Export,Case Size On Disk,Case File Size,Case Audit Size,Oldest Item Date,Newest Item Date,Is Compound,Cases Contained")
            Case "Case Size"
                ReportOutputFile.WriteLine("CaseGUID, Collection Status,Report Load Duration,Case Name,Current Case Version,Upgraded Case Version,Case Location,Batch Load Info,Data Export,Case Size On Disk,Case File Size,Case Audit Size,Oldest Item Date,Newest Item Date,Is Compound,Cases Contained, Contained In Case,Load Events,Total Load Time,Processing Speed")
            Case "Custodians in Case"
                ReportOutputFile.WriteLine("CaseGUID, Collection Status,Report Load Duration,Case Name,Current Case Version,Upgraded Case Version,Case Location,Batch Load Info,Data Export,Case Size On Disk,Case File Size,Case Audit Size,Oldest Item Date,Newest Item Date,Is Compound,Cases Contained, Contained In Case,Investigator,Investigator Sessions, Invalid Sessions, Investigator Time Summary,Custodians,Custodian Count")
            Case "Metadata type"
                ReportOutputFile.WriteLine("CaseGUID, Collection Status,Report Load Duration,Case Name,Current Case Version,Upgraded Case Version,Case Location,Batch Load Info,Data Export,Case Size On Disk,Case File Size,Case Audit Size,Oldest Item Date,Newest Item Date,Is Compound,Cases Contained, Contained In Case,Mime Types,Item Types")
            Case "Processing time"
                ReportOutputFile.WriteLine("CaseGUID, Collection Status,Report Load Duration,Case Name,Current Case Version,Upgraded Case Version,Case Location,Batch Load Info,Data Export,Case Size On Disk,Case File Size,Case Audit Size,Oldest Item Date,Newest Item Date,Is Compound,Cases Contained, Contained In Case,Load Events,Total Load Time,Processing Speed")
            Case "Processing speed (GB per hour)"
                ReportOutputFile.WriteLine("CaseGUID, Collection Status,Report Load Duration,Case Name,Current Case Version,Upgraded Case Version,Case Location,Batch Load Info,Data Export,Case Size On Disk,Case File Size,Case Audit Size,Oldest Item Date,Newest Item Date,Is Compound,Cases Contained, Contained In Case,Load Events,TotalLoadTime,ProcessingSpeed")
            Case "Search Term Hit"
                ReportOutputFile.WriteLine("CaseGUID, Collection Status,Report Load Duration,Case Name,Current Case Version,Upgraded Case Version,Case Location,Batch Load Info,Data Export,Case Size On Disk,Case File Size,Case Audit Size,Oldest Item Date,Newest Item Date,Is Compound,Cases Contained, Contained In Case,Search Term,Search Size,Search Hit Count,Custodian Search Hit,Total Case Item Count,Hit Count Percent")
            Case "Total Number of Items"
                ReportOutputFile.WriteLine("CaseGUID, Collection Status,Report Load Duration,Case Name,Current Case Version,Upgraded Case Version,Case Location,Batch Load Info,Data Export,Case Size On Disk,Case File Size,Case Audit Size,Oldest Item Date,Newest Item Date,Is Compound,Cases Contained, Contained In Case,Total Case Item Count")
            Case "Total Number of workers"
                ReportOutputFile.WriteLine("CaseGUID, Collection Status,Report Load Duration,Case Name,Current Case Version,Upgraded Case Version,Case Location,Batch Load Info,Data Export,Case Size On Disk,Case File Size,Case Audit Size,Oldest Item Date,Newest Item Date,Is Compound,Cases Contained, Contained In Case,Investigator,Investigator Sessions, Invalid Sessions, Investigator Time Summary,Broker Memory,Worker Count,Worker Memory,Load Events,Total Load Time,Processing Speed")
        End Select

        For Each row In grdCaseInfo.Rows
            sCaseGuid = row.cells("CaseGUID").value
            If sCaseGuid <> vbNullString Then
                row.cells("CollectionStatus").value = "Exporting Case Data..."
                sCollectionStatus = row.cells("CollectionStatus").value

                sCaseName = row.cells("CaseName").value
                sReportLoadDuration = row.cells("ReportLoadDuration").value
                sCurrentCaseVersion = row.cells("CurrentCaseVersion").value
                sPercentComplete = row.cells("PercentComplete").value
                sUpgradedCaseVersion = row.cells("UpgradedCaseVersion").value
                sBatchLoadInfo = row.cells("BatchLoadInfo").value
                sDataExport = row.cells("DataExport").value
                sCaseLocation = row.cells("CaseLocation").value
                sCaseSizeOnDisk = row.cells("CaseSizeOnDisk").value
                sCaseSizeOnDisk = sCaseSizeOnDisk.Replace(",", "")
                sCaseFileSize = row.cells("CaseFileSize").value
                sCaseFileSize = sCaseFileSize.Replace(",", "")
                sCaseAuditSize = row.cells("CaseAuditSize").value
                sCaseAuditSize = sCaseAuditSize.Replace(",", "")
                sOldestTopLevel = row.cells("OldestTopLevel").Value
                sNewestTopLevel = row.cells("NewestTopLevel").value
                sIsCompound = row.cells("IsCompound").value
                sCasesContained = row.cells("CasesContained").value
                sContainedInCase = row.cells("ContainedInCase").value
                sInvestigator = row.cells("Investigator").value
                sInvestigatorSessions = row.cells("InvestigatorSessions").value
                sInvalidSessions = row.cells("InvalidSessions").value
                sInvestigatorTimeSummary = row.cells("InvestigatorTimeSummary").value
                sBrokerMemory = row.cells("BrokerMemory").value
                sWorkerCount = row.cells("WorkerCount").value
                sWorkerMemory = row.cells("WorkerMemory").value
                sEvidenceName = row.cells("EvidenceName").value
                sEvidenceCustomMetadata = row.cells("EvidenceCustomMetadata").value
                sEvidenceLocation = row.cells("EvidenceLocation").value
                sLanguagesContained = row.cells("LanguagesContained").value
                sMimeTypes = row.cells("MimeTypes").value
                sItemTypes = row.cells("ItemTypes").value
                sIrregularItems = row.cells("IrregularItems").value
                sCreationDate = row.cells("CreationDate").value
                sModifiedDate = row.cells("ModifiedDate").value
                sLoadEvents = row.cells("LoadEvents").value
                sTotalLoadTime = row.cells("TotalLoadTime").value
                sProcessingSpeed = row.cells("ProcessingSpeed").value
                sCustodians = row.cells("Custodians").value
                sCustodianCount = row.cells("CustodianCount").value
                sSearchTerm = row.cells("SearchTerm").value
                sSearchSize = row.cells("SearchSize").value.Replace(",", "")
                sSearchSize = sSearchSize.Replace(",", "")
                sSearchHitCount = row.cells("SearchHitCount").value
                sSearchHitCount = sSearchHitCount.Replace(",", "")
                sCustodianSearchHit = row.cells("CustodianSearchHit").value
                sTotalCaseItemCount = row.cells("TotalCaseItemCount").value
                sTotalCaseItemCount = sTotalCaseItemCount.Replace(",", "")
                sHitCountPercent = row.cells("HitCountPercent").value
                sNuixLogLocation = row.cells("NuixLogLocation").value
                Select Case cboReportType.Text
                    Case "All"
                        ReportOutputFile.WriteLine("""" & sCaseGuid & """" & "," & """" & sCollectionStatus & """" & "," & """" & sPercentComplete & """" & "," & """" & sReportLoadDuration & """" & "," & """" & sCaseName & """" & "," & """" & sCurrentCaseVersion & """" & "," & """" & sUpgradedCaseVersion & """" & "," & """" & sCaseLocation & """" & "," & """" & sBatchLoadInfo & """" & "," & """" & sDataExport & """" & "," & """" & sCaseDescription & """" & "," & """" & sCaseSizeOnDisk & """" & "," & """" & sCaseFileSize & """" & "," & """" & sCaseAuditSize & """" & "," & """" & sOldestTopLevel & """" & "," & """" & sNewestTopLevel & """" & "," & """" & sIsCompound & """" & "," & """" & sCasesContained & """" & "," & """" & sContainedInCase & """" & "," & """" & sInvestigator & """" & "," & """" & sInvestigatorSessions & """" & "," & """" & sInvalidSessions & """" & "," & """" & sInvestigatorTimeSummary & """" & "," & """" & sBrokerMemory & """" & "," & """" & sWorkerCount & """" & "," & """" & sWorkerMemory & """" & "," & """" & sEvidenceName & """" & "," & """" & sEvidenceLocation & """" & "," & """" & sEvidenceDescription & """" & "," & """" & sEvidenceCustomMetadata & """" & "," & """" & sLanguagesContained & """" & "," & """" & sMimeTypes & """" & "," & """" & sItemTypes & """" & "," & """" & sIrregularItems & """" & "," & """" & sCreationDate & """" & "," & """" & sModifiedDate & """" & "," & """" & sLoadStartDate & """" & "," & """" & sLoadEndDate & """" & "," & """" & sLoadEvents & """" & "," & """" & sTotalLoadTime & """" & "," & """" & sProcessingSpeed & """" & "," & """" & sTotalCaseItemCount & """" & "," & """" & sDuplicateItems & """" & "," & """" & sCustodians & """" & "," & """" & sCustodianCount & """" & "," & """" & sSearchTerm & """" & "," & """" & sSearchSize & """" & "," & """" & sSearchHitCount & """" & "," & """" & sCustodianSearchHit & """" & "," & """" & sHitCountPercent & """" & "," & """" & sNuixLogLocation & """")
                    Case "App Memory per case"
                        ReportOutputFile.WriteLine("""" & sCaseGuid & """" & "," & """" & sCollectionStatus & """" & "," & """" & sReportLoadDuration & """" & "," & """" & sCaseName & """" & "," & """" & sCurrentCaseVersion & """" & "," & """" & sUpgradedCaseVersion & """" & "," & """" & sCaseLocation & """" & "," & """" & sBatchLoadInfo & """" & """" & sDataExport & """" & "," & """" & sCaseSizeOnDisk & """" & "," & """" & sCaseFileSize & """" & "," & """" & sCaseAuditSize & """" & sOldestTopLevel & """" & "," & """" & sNewestTopLevel & "," & """" & "," & """" & sIsCompound & """" & "," & """" & sCasesContained & """" & "," & """" & sContainedInCase & """" & "," & """" & sBrokerMemory & """" & "," & """" & sWorkerCount & """" & "," & """" & sWorkerMemory & """")
                    Case "Case by Investigator"
                        ReportOutputFile.WriteLine("""" & sCaseGuid & """" & "," & """" & sCollectionStatus & """" & "," & """" & sReportLoadDuration & """" & "," & """" & sCaseName & """" & "," & """" & sCurrentCaseVersion & """" & "," & """" & sUpgradedCaseVersion & """" & "," & """" & sCaseLocation & """" & "," & """" & sBatchLoadInfo & """" & """" & sDataExport & """" & "," & """" & sCaseSizeOnDisk & """" & "," & """" & sCaseFileSize & """" & "," & """" & sCaseAuditSize & """" & sOldestTopLevel & """" & "," & """" & sNewestTopLevel & "," & """" & "," & """" & sIsCompound & """" & "," & """" & sCasesContained & """" & "," & """" & sContainedInCase & """" & "," & """" & sInvestigator & """" & "," & """" & sInvestigatorSessions & """" & "," & """" & sInvalidSessions & """" & "," & """" & sInvestigatorTimeSummary)
                    Case "Case Evidence"
                        ReportOutputFile.WriteLine("""" & sCaseGuid & """" & "," & """" & sCollectionStatus & """" & "," & """" & sReportLoadDuration & """" & "," & """" & sCaseName & """" & "," & """" & sCurrentCaseVersion & """" & "," & """" & sUpgradedCaseVersion & """" & "," & """" & sCaseLocation & """" & "," & """" & sBatchLoadInfo & """" & """" & sDataExport & """" & "," & """" & sCaseSizeOnDisk & """" & "," & """" & sCaseFileSize & """" & "," & """" & sCaseAuditSize & """" & sOldestTopLevel & """" & "," & """" & sNewestTopLevel & "," & """" & "," & """" & sIsCompound & """" & "," & """" & sCasesContained & """" & "," & """" & sContainedInCase & """" & "," & """" & sEvidenceName & """" & "," & """" & sEvidenceLocation & """")
                    Case "Case Location"
                        ReportOutputFile.WriteLine("""" & sCaseGuid & """" & "," & """" & sCollectionStatus & """" & "," & """" & sReportLoadDuration & """" & "," & """" & sCaseName & """" & "," & """" & sCurrentCaseVersion & """" & "," & """" & sUpgradedCaseVersion & """" & "," & """" & sCaseLocation & """" & "," & """" & sBatchLoadInfo & """" & """" & sDataExport & """" & "," & """" & sCaseSizeOnDisk & """" & "," & """" & sCaseFileSize & """" & "," & """" & sCaseAuditSize & """" & sOldestTopLevel & """" & "," & """" & sNewestTopLevel & "," & """" & "," & """" & sIsCompound & """" & "," & """" & sCasesContained & """")
                    Case "Case Size"
                        ReportOutputFile.WriteLine("""" & sCaseGuid & """" & "," & """" & sCollectionStatus & """" & "," & """" & sReportLoadDuration & """" & "," & """" & sCaseName & """" & "," & """" & sCurrentCaseVersion & """" & "," & """" & sUpgradedCaseVersion & """" & "," & """" & sCaseLocation & """" & "," & """" & sBatchLoadInfo & """" & """" & sDataExport & """" & "," & """" & sCaseSizeOnDisk & """" & "," & """" & sCaseFileSize & """" & "," & """" & sCaseAuditSize & """" & sOldestTopLevel & """" & "," & """" & sNewestTopLevel & "," & """" & "," & """" & sIsCompound & """" & "," & """" & sCasesContained & """" & "," & """" & sContainedInCase & """" & "," & """" & sLoadEvents & """" & "," & """" & sTotalLoadTime & """" & "," & """" & sProcessingSpeed & """")
                    Case "Custodians in Case"
                        ReportOutputFile.WriteLine("""" & sCaseGuid & """" & "," & """" & sCollectionStatus & """" & "," & """" & sReportLoadDuration & """" & "," & """" & sCaseName & """" & "," & """" & sCurrentCaseVersion & """" & "," & """" & sUpgradedCaseVersion & """" & "," & """" & sCaseLocation & """" & "," & """" & sBatchLoadInfo & """" & """" & sDataExport & """" & "," & """" & sCaseSizeOnDisk & """" & "," & """" & sCaseFileSize & """" & "," & """" & sCaseAuditSize & """" & sOldestTopLevel & """" & "," & """" & sNewestTopLevel & "," & """" & "," & """" & sIsCompound & """" & "," & """" & sCasesContained & """" & "," & """" & sContainedInCase & """" & "," & """" & sInvestigator & """" & "," & """" & sInvestigatorSessions & """" & "," & """" & sInvalidSessions & """" & "," & """" & sInvestigatorTimeSummary & """" & "," & """" & sCustodians & """" & "," & """" & sCustodianCount & """")
                    Case "Metadata type"
                        ReportOutputFile.WriteLine("""" & sCaseGuid & """" & "," & """" & sCollectionStatus & """" & "," & """" & sReportLoadDuration & """" & "," & """" & sCaseName & """" & "," & """" & sCurrentCaseVersion & """" & "," & """" & sUpgradedCaseVersion & """" & "," & """" & sCaseLocation & """" & "," & """" & sBatchLoadInfo & """" & """" & sDataExport & """" & "," & """" & sCaseSizeOnDisk & """" & "," & """" & sCaseFileSize & """" & "," & """" & sCaseAuditSize & """" & sOldestTopLevel & """" & "," & """" & sNewestTopLevel & "," & """" & "," & """" & sIsCompound & """" & "," & """" & sCasesContained & """" & "," & """" & sContainedInCase & """" & "," & """" & sMimeTypes & """" & "," & """" & sItemTypes & """")
                    Case "Procesing time"
                        ReportOutputFile.WriteLine("""" & sCaseGuid & """" & "," & """" & sCollectionStatus & """" & "," & """" & sReportLoadDuration & """" & "," & """" & sCaseName & """" & "," & """" & sCurrentCaseVersion & """" & "," & """" & sUpgradedCaseVersion & """" & "," & """" & sCaseLocation & """" & "," & """" & sBatchLoadInfo & """" & """" & sDataExport & """" & "," & """" & sCaseSizeOnDisk & """" & "," & """" & sCaseFileSize & """" & "," & """" & sCaseAuditSize & """" & sOldestTopLevel & """" & "," & """" & sNewestTopLevel & "," & """" & "," & """" & sIsCompound & """" & "," & """" & sCasesContained & """" & "," & """" & sContainedInCase & """" & "," & """" & sLoadEvents & """" & "," & """" & sTotalLoadTime & """" & "," & """" & sProcessingSpeed & """")
                    Case "Procesing speed (GB per hour)"
                        ReportOutputFile.WriteLine("""" & sCaseGuid & """" & "," & """" & sCollectionStatus & """" & "," & """" & sReportLoadDuration & """" & "," & """" & sCaseName & """" & "," & """" & sCurrentCaseVersion & """" & "," & """" & sUpgradedCaseVersion & """" & "," & """" & sCaseLocation & """" & "," & """" & sBatchLoadInfo & """" & """" & sDataExport & """" & "," & """" & sCaseSizeOnDisk & """" & "," & """" & sCaseFileSize & """" & "," & """" & sCaseAuditSize & """" & sOldestTopLevel & """" & "," & """" & sNewestTopLevel & "," & """" & "," & """" & sIsCompound & """" & "," & """" & sCasesContained & """" & "," & """" & sContainedInCase & """" & "," & """" & sLoadEvents & """" & "," & """" & sTotalLoadTime & """" & "," & """" & sProcessingSpeed & """")
                    Case "Search Term Hit"
                        ReportOutputFile.WriteLine("""" & sCaseGuid & """" & "," & """" & sCollectionStatus & """" & "," & """" & sReportLoadDuration & """" & "," & """" & sCaseName & """" & "," & """" & sCurrentCaseVersion & """" & "," & """" & sUpgradedCaseVersion & """" & "," & """" & sCaseLocation & """" & "," & """" & sBatchLoadInfo & """" & """" & sDataExport & """" & "," & """" & sCaseSizeOnDisk & """" & "," & """" & sCaseFileSize & """" & "," & """" & sCaseAuditSize & """" & sOldestTopLevel & """" & "," & """" & sNewestTopLevel & "," & """" & "," & """" & sIsCompound & """" & "," & """" & sCasesContained & """" & "," & """" & sContainedInCase & """" & "," & """" & sSearchTerm & """" & "," & """" & sSearchSize & """" & "," & """" & sSearchHitCount & """" & "," & """" & sCustodianSearchHit & """" & "," & """" & sTotalCaseItemCount & """" & "," & """" & sHitCountPercent & """")
                    Case "Total Number of Items"
                        ReportOutputFile.WriteLine("""" & sCaseGuid & """" & "," & """" & sCollectionStatus & """" & "," & """" & sReportLoadDuration & """" & "," & """" & sCaseName & """" & "," & """" & sCurrentCaseVersion & """" & "," & """" & sUpgradedCaseVersion & """" & "," & """" & sCaseLocation & """" & "," & """" & sBatchLoadInfo & """" & """" & sDataExport & """" & "," & """" & sCaseSizeOnDisk & """" & "," & """" & sCaseFileSize & """" & "," & """" & sCaseAuditSize & """" & sOldestTopLevel & """" & "," & """" & sNewestTopLevel & "," & """" & "," & """" & sIsCompound & """" & "," & """" & sCasesContained & """" & "," & """" & sContainedInCase & """" & "," & """" & sTotalCaseItemCount & """")
                    Case "Total Number of workers"
                        ReportOutputFile.WriteLine("""" & sCaseGuid & """" & "," & """" & sCollectionStatus & """" & "," & """" & sReportLoadDuration & """" & "," & """" & sCaseName & """" & "," & """" & sCurrentCaseVersion & """" & "," & """" & sUpgradedCaseVersion & """" & "," & """" & sCaseLocation & """" & "," & """" & sBatchLoadInfo & """" & """" & sDataExport & """" & "," & """" & sCaseSizeOnDisk & """" & "," & """" & sCaseFileSize & """" & "," & """" & sCaseAuditSize & """" & sOldestTopLevel & """" & "," & """" & sNewestTopLevel & "," & """" & "," & """" & sIsCompound & """" & "," & """" & sCasesContained & """" & "," & """" & sContainedInCase & """" & "," & """" & sInvestigator & """" & "," & """" & sInvestigatorSessions & """" & "," & """" & sInvalidSessions & """" & "," & """" & sInvestigatorTimeSummary & """" & "," & """" & sBrokerMemory & """" & "," & """" & sWorkerCount & """" & "," & """" & sWorkerMemory & """" & "," & """" & sLoadEvents & """" & "," & """" & sTotalLoadTime & """" & "," & """" & sProcessingSpeed & """")
                End Select
            End If
            row.cells("CollectionStatus").value = "Case Data Exported"
        Next

        ReportOutputFile.Close()
        MessageBox.Show("Nuix Case Statistics report finished building.  Report located at: " & sOutputFileName)

    End Sub

    Private Sub JSonToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles JSonToolStripMenuItem.Click
        Dim sReportFilePath As String
        Dim ReportOutputFile As StreamWriter
        Dim sReportType As String
        Dim sOutputFileName As String
        Dim sMachineName As String

        Dim sCaseGUID As String
        Dim sCollectionStatus As String
        Dim sCaseName As String
        Dim sReportLoadDuration As String

        Dim sBatchLoadInfo As String
        Dim sCurrentCaseVersion As String
        Dim sUpgradedCaseVersion As String
        Dim sIsCompound As String
        Dim sCompoundCaseContains As String
        Dim sContainedInCase As String
        Dim sCaseLocation As String
        Dim sCaseSizeOnDisk As String
        Dim sCaseAuditSize As String
        Dim sOldestItem As String
        Dim sNewestItem As String
        Dim sCaseFileSize As String
        Dim sInvestigator As String
        Dim sInvestigatorSessions As String
        Dim sInvalidSessions As String
        Dim sInvestigatorTimeSummary As String
        Dim sBrokerMemory As String
        Dim sWorkerCount As String
        Dim sWorkerMemory As String
        Dim sEvidenceName As String
        Dim sEvidenceLocation As String
        Dim sMimeTypes As String
        Dim sItemTypes As String
        Dim sCreationDate As String
        Dim sModifiedDate As String
        Dim sProcessingSpeed As String
        Dim sCustodians As String
        Dim sCustodianCount As String
        Dim sSearchTerm As String
        Dim sSearchSize As String
        Dim sSearchHitCount As String
        Dim sCustodianSearchHit As String
        Dim sTotalItemCount As String
        Dim sHitCountPercent As String
        Dim sLoadEvents As String
        Dim sTotalLoadTime As String
        Dim iRowCount As Integer
        Dim iCounter As Integer
        Dim sDataExport As String

        If cboReportType.Text = vbNullString Then
            MessageBox.Show("You must enter a report type to export the data for.", "No Report Type selected")
            cboReportType.Focus()
            Exit Sub

        End If
        sReportFilePath = txtReportLocation.Text
        sReportType = cboReportType.Text
        sMachineName = System.Net.Dns.GetHostName()
        sOutputFileName = sMachineName & "-" & sReportType & "-" & DateTime.Now.ToString("yyyy-MM-dd hh-mm-ss") & ".json"
        ReportOutputFile = New StreamWriter(sReportFilePath & "\" & sOutputFileName)

        ReportOutputFile.WriteLine("{")
        ReportOutputFile.WriteLine(" " & """" & "NuixCaseReport" & """" & ": {")
        ReportOutputFile.WriteLine("     " & """" & cboReportType.Text & """" & ": {")

        iRowCount = (grdCaseInfo.RowCount - 1)
        iCounter = 0

        For Each row In grdCaseInfo.Rows
            If row.cells("CaseName").value <> vbNullString Then
                row.cells("CollectionStatus").value = "Exporting Case Data..."
                sCaseGUID = row.cells("CaseGUID").value
                sCollectionStatus = row.cells("CollectionStatus").value
                sCaseName = row.cells("CaseName").value
                sReportLoadDuration = row.cells("ReportLoadDuration").value
                sCurrentCaseVersion = row.cells("CurrentCaseVersion").value
                sUpgradedCaseVersion = row.cells("UpgradedCaseVersion").value
                sCaseLocation = row.cells("CaseLocation").value
                sCaseLocation = sCaseLocation.Replace("\", "\\")
                sBatchLoadInfo = row.cells("BatchLoadInfo").value
                sDataExport = row.cells("DataExport").value
                sCaseSizeOnDisk = row.cells("CaseSizeOnDisk").value
                sCaseAuditSize = row.cells("CaseAuditSize").value
                sCaseFileSize = row.cells("CaseFileSize").value
                sOldestItem = row.cells("OldestTopLevel").value
                sNewestItem = row.cells("NewestTopLevel").value
                sIsCompound = row.cells("IsCompound").value
                sCompoundCaseContains = row.cells("CasesContained").value
                sContainedInCase = row.cells("ContainedInCase").value
                sInvestigator = row.cells("Investigator").value
                sInvestigatorSessions = row.cells("InvestigatorSessions").value
                sInvalidSessions = row.cells("InvalidSessions").value
                sInvestigatorTimeSummary = row.cells("InvestigatorTimeSummary").value
                sBrokerMemory = row.cells("BrokerMemory").value
                sWorkerCount = row.cells("WorkerCount").value
                sWorkerMemory = row.cells("WorkerMemory").value
                sEvidenceName = row.cells("EvidenceName").value
                sEvidenceLocation = row.cells("EvidenceLocation").value
                sEvidenceLocation = sEvidenceLocation.Replace("\", "\\")
                sMimeTypes = row.cells("MimeTypes").value
                sItemTypes = row.cells("ItemTypes").value
                sCreationDate = row.cells("CreationDate").value
                sModifiedDate = row.cells("ModifiedDate").value
                sLoadEvents = row.cells("LoadEvents").value
                sTotalLoadTime = row.cells("TotalLoadTime").value
                sProcessingSpeed = row.cells("ProcessingSpeed").value
                sCustodians = row.cells("Custodians").value
                sCustodianCount = row.cells("CustodianCount").value
                sSearchTerm = row.cells("SearchTerm").value
                sSearchSize = row.cells("SearchSize").value
                sSearchHitCount = row.cells("SearchHitCount").value
                sCustodianSearchHit = row.cells("CustodianSearchHit").value
                sTotalItemCount = row.cells("TotalCaseItemCount").value
                sHitCountPercent = row.cells("HitCountPercent").value
                ReportOutputFile.WriteLine("        " & """" & sCaseName & """" & ": {")

                Select Case cboReportType.Text
                    Case "All"
                        ReportOutputFile.WriteLine("          " & """" & "Case GUID" & """" & ": " & """" & sCaseGUID & """" & ",")
                        ReportOutputFile.WriteLine("          " & """" & "Collection Status" & """" & ": " & """" & sCollectionStatus & """" & ",")
                        ReportOutputFile.WriteLine("          " & """" & "Report Load Duration" & """" & ": " & """" & sReportLoadDuration & """" & ",")
                        ReportOutputFile.WriteLine("          " & """" & "Current Case Version" & """" & ": " & """" & sCurrentCaseVersion & """" & ",")
                        ReportOutputFile.WriteLine("          " & """" & "Upgraded Case Version" & """" & ": " & """" & sUpgradedCaseVersion & """" & ",")
                        ReportOutputFile.WriteLine("          " & """" & "Batch Load Info" & """" & ": " & """" & sBatchLoadInfo & """" & ",")
                        ReportOutputFile.WriteLine("          " & """" & "Data Export" & """" & ": " & """" & sDataExport & """" & ",")
                        ReportOutputFile.WriteLine("          " & """" & "Case Location" & """" & ": " & """" & sCaseLocation & """" & ",")
                        ReportOutputFile.WriteLine("          " & """" & "Case Size On Disk" & """" & ": " & """" & sCaseSizeOnDisk & """" & ",")
                        ReportOutputFile.WriteLine("          " & """" & "Case File Size" & """" & ": " & """" & sCaseFileSize & """" & ",")
                        ReportOutputFile.WriteLine("          " & """" & "Case Audit Size" & """" & ": " & """" & sCaseAuditSize & """" & ",")
                        ReportOutputFile.WriteLine("          " & """" & "Oldest Top Level Item" & """" & ": " & """" & sOldestItem & """" & ",")
                        ReportOutputFile.WriteLine("          " & """" & "Newest Top Level Item" & """" & ": " & """" & sNewestItem & """" & ",")
                        ReportOutputFile.WriteLine("          " & """" & "Is Compound Case" & """" & ": " & """" & sIsCompound & """" & ",")
                        ReportOutputFile.WriteLine("          " & """" & "Cases Contained" & """" & ": " & """" & sCompoundCaseContains & """" & ",")
                        ReportOutputFile.WriteLine("          " & """" & "Contained In Case" & """" & ": " & """" & sContainedInCase & """" & ",")
                        ReportOutputFile.WriteLine("          " & """" & "Investigator" & """" & ": " & """" & sInvestigator & """" & ",")
                        ReportOutputFile.WriteLine("          " & """" & "Investigator Sessions" & """" & ": " & """" & sInvestigatorSessions & """" & ",")
                        ReportOutputFile.WriteLine("          " & """" & "Invalid Sessions" & """" & ": " & """" & sInvalidSessions & """" & ",")
                        ReportOutputFile.WriteLine("          " & """" & "Investigator Time Summary" & """" & ": " & """" & sInvestigatorTimeSummary & """" & ",")
                        ReportOutputFile.WriteLine("          " & """" & "Broker Memory" & """" & ": " & """" & sBrokerMemory & """" & ",")
                        ReportOutputFile.WriteLine("          " & """" & "Worker Count" & """" & ": " & """" & sWorkerCount & """" & ",")
                        ReportOutputFile.WriteLine("          " & """" & "Worker Memory" & """" & ": " & """" & sWorkerMemory & """" & ",")
                        ReportOutputFile.WriteLine("          " & """" & "Evidence Name" & """" & ": " & """" & sEvidenceName & """" & ",")
                        ReportOutputFile.WriteLine("          " & """" & "Evidence Location" & """" & ": " & """" & sEvidenceLocation & """" & ",")
                        ReportOutputFile.WriteLine("          " & """" & "Mime Types" & """" & ": " & """" & sMimeTypes & """" & ",")
                        ReportOutputFile.WriteLine("          " & """" & "Item Types" & """" & ": " & """" & sItemTypes & """" & ",")
                        ReportOutputFile.WriteLine("          " & """" & "sCreationDate" & """" & ": " & """" & sCreationDate & """" & ",")
                        ReportOutputFile.WriteLine("          " & """" & "Modified Date" & """" & ": " & """" & sModifiedDate & """" & ",")
                        ReportOutputFile.WriteLine("          " & """" & "Load Events" & """" & ": " & """" & sLoadEvents & """" & ",")
                        ReportOutputFile.WriteLine("          " & """" & "Total Load Time" & """" & ": " & """" & sTotalLoadTime & """" & ",")
                        ReportOutputFile.WriteLine("          " & """" & "Processing Speed" & """" & ": " & """" & sProcessingSpeed & """" & ",")
                        ReportOutputFile.WriteLine("          " & """" & "Custodians" & """" & ": " & """" & sCustodians & """" & ",")
                        ReportOutputFile.WriteLine("          " & """" & "Custodian Count" & """" & ": " & """" & sCustodianCount & """" & ",")
                        ReportOutputFile.WriteLine("          " & """" & "sSearchTerm" & """" & ": " & """" & sSearchTerm & """" & ",")
                        ReportOutputFile.WriteLine("          " & """" & "Search Size" & """" & ": " & """" & sSearchSize & """" & ",")
                        ReportOutputFile.WriteLine("          " & """" & "Search Hit Count" & """" & ": " & """" & sSearchHitCount & """" & ",")
                        ReportOutputFile.WriteLine("          " & """" & "Custodian Search Hit" & """" & ": " & """" & sCustodianSearchHit & """" & ",")
                        ReportOutputFile.WriteLine("          " & """" & "Total Item Count" & """" & ": " & """" & sTotalItemCount & """" & ",")
                        ReportOutputFile.WriteLine("          " & """" & "Hit Count Percent" & """" & ": " & """" & sHitCountPercent & """")
                        If iCounter < iRowCount - 1 Then
                            ReportOutputFile.WriteLine("        },")
                        Else
                            ReportOutputFile.WriteLine("        }")

                        End If
                    Case "App Memory per case"
                        ReportOutputFile.WriteLine("          " & """" & "Case GUID" & """" & ": " & """" & sCaseGUID & """" & ",")
                        ReportOutputFile.WriteLine("          " & """" & "Collection Status" & """" & ": " & """" & sCollectionStatus & """" & ",")
                        ReportOutputFile.WriteLine("          " & """" & "Report Load Duration" & """" & ": " & """" & sReportLoadDuration & """" & ",")
                        ReportOutputFile.WriteLine("          " & """" & "Current Case Version" & """" & ": " & """" & sCurrentCaseVersion & """" & ",")
                        ReportOutputFile.WriteLine("          " & """" & "Upgraded Case Version" & """" & ": " & """" & sUpgradedCaseVersion & """" & ",")
                        ReportOutputFile.WriteLine("          " & """" & "Batch Load Info" & """" & ": " & """" & sBatchLoadInfo & """" & ",")
                        ReportOutputFile.WriteLine("          " & """" & "Data Export" & """" & ": " & """" & sDataExport & """" & ",")
                        ReportOutputFile.WriteLine("          " & """" & "Is Compound Case" & """" & ": " & """" & sIsCompound & """" & ",")
                        ReportOutputFile.WriteLine("          " & """" & "Cases Contained" & """" & ": " & """" & sCompoundCaseContains & """" & ",")
                        ReportOutputFile.WriteLine("          " & """" & "Contained In Case" & """" & ": " & """" & sContainedInCase & """" & ",")
                        ReportOutputFile.WriteLine("          " & """" & "Case Location" & """" & ": " & """" & sCaseLocation & """" & ",")
                        ReportOutputFile.WriteLine("          " & """" & "Case Size On Disk" & """" & ": " & """" & sCaseSizeOnDisk & """" & ",")
                        ReportOutputFile.WriteLine("          " & """" & "Oldest Top Level Item" & """" & ": " & """" & sOldestItem & """" & ",")
                        ReportOutputFile.WriteLine("          " & """" & "Newest Top Level Item" & """" & ": " & """" & sNewestItem & """" & ",")
                        ReportOutputFile.WriteLine("          " & """" & "Broker Memory" & """" & ": " & """" & sBrokerMemory & """" & ",")
                        ReportOutputFile.WriteLine("          " & """" & "Worker Count" & """" & ": " & """" & sWorkerCount & """" & ",")
                        ReportOutputFile.WriteLine("          " & """" & "Worker Memory" & """" & ": " & """" & sWorkerMemory & """" & ",")
                        ReportOutputFile.WriteLine("          " & """" & "Processing Speed" & """" & ": " & """" & sProcessingSpeed & """")
                        If iCounter < iRowCount - 1 Then
                            ReportOutputFile.WriteLine("        },")
                        Else
                            ReportOutputFile.WriteLine("        }")

                        End If
                    Case "Case by Investigator"
                        ReportOutputFile.WriteLine("          " & """" & "Case GUID" & """" & ": " & """" & sCaseGUID & """" & ",")
                        ReportOutputFile.WriteLine("          " & """" & "Collection Status" & """" & ": " & """" & sCollectionStatus & """" & ",")
                        ReportOutputFile.WriteLine("          " & """" & "Report Load Duration" & """" & ": " & """" & sReportLoadDuration & """" & ",")
                        ReportOutputFile.WriteLine("          " & """" & "Current Case Version" & """" & ": " & """" & sCurrentCaseVersion & """" & ",")
                        ReportOutputFile.WriteLine("          " & """" & "Upgraded Case Version" & """" & ": " & """" & sUpgradedCaseVersion & """" & ",")
                        ReportOutputFile.WriteLine("          " & """" & "Batch Load Info" & """" & ": " & """" & sBatchLoadInfo & """" & ",")
                        ReportOutputFile.WriteLine("          " & """" & "Data Export" & """" & ": " & """" & sDataExport & """" & ",")
                        ReportOutputFile.WriteLine("          " & """" & "Is Compound Case" & """" & ": " & """" & sIsCompound & """" & ",")
                        ReportOutputFile.WriteLine("          " & """" & "Cases Contained" & """" & ": " & """" & sCompoundCaseContains & """" & ",")
                        ReportOutputFile.WriteLine("          " & """" & "Contained In Case" & """" & ": " & """" & sContainedInCase & """" & ",")
                        ReportOutputFile.WriteLine("          " & """" & "Case Location" & """" & ": " & """" & sCaseLocation & """" & ",")
                        ReportOutputFile.WriteLine("          " & """" & "Case Size On Disk" & """" & ": " & """" & sCaseSizeOnDisk & """" & ",")
                        ReportOutputFile.WriteLine("          " & """" & "Oldest Top Level Item" & """" & ": " & """" & sOldestItem & """" & ",")
                        ReportOutputFile.WriteLine("          " & """" & "Newest Top Level Item" & """" & ": " & """" & sNewestItem & """" & ",")
                        ReportOutputFile.WriteLine("          " & """" & "Investigator" & """" & ": " & """" & sInvestigator & """" & ",")
                        ReportOutputFile.WriteLine("          " & """" & "Investigator Sessions" & """" & ": " & """" & sInvestigatorSessions & """" & ",")
                        ReportOutputFile.WriteLine("          " & """" & "Invalid Sessions" & """" & ": " & """" & sInvalidSessions & """" & ",")
                        ReportOutputFile.WriteLine("          " & """" & "Investigator Time Summary" & """" & ": " & """" & sInvestigatorTimeSummary & """")
                        If iCounter < iRowCount - 1 Then
                            ReportOutputFile.WriteLine("        },")
                        Else
                            ReportOutputFile.WriteLine("        }")

                        End If
                    Case "Case Evidence"
                        ReportOutputFile.WriteLine("          " & """" & "Case GUID" & """" & ": " & """" & sCaseGUID & """" & ",")
                        ReportOutputFile.WriteLine("          " & """" & "Collection Status" & """" & ": " & """" & sCollectionStatus & """" & ",")
                        ReportOutputFile.WriteLine("          " & """" & "Report Load Duration" & """" & ": " & """" & sReportLoadDuration & """" & ",")
                        ReportOutputFile.WriteLine("          " & """" & "Current Case Version" & """" & ": " & """" & sCurrentCaseVersion & """" & ",")
                        ReportOutputFile.WriteLine("          " & """" & "Upgraded Case Version" & """" & ": " & """" & sUpgradedCaseVersion & """" & ",")
                        ReportOutputFile.WriteLine("          " & """" & "Batch Load Info" & """" & ": " & """" & sBatchLoadInfo & """" & ",")
                        ReportOutputFile.WriteLine("          " & """" & "Data Export" & """" & ": " & """" & sDataExport & """" & ",")
                        ReportOutputFile.WriteLine("          " & """" & "Is Compound Case" & """" & ": " & """" & sIsCompound & """" & ",")
                        ReportOutputFile.WriteLine("          " & """" & "Cases Contained" & """" & ": " & """" & sCompoundCaseContains & """" & ",")
                        ReportOutputFile.WriteLine("          " & """" & "Contained In Case" & """" & ": " & """" & sContainedInCase & """" & ",")
                        ReportOutputFile.WriteLine("          " & """" & "Case Location" & """" & ": " & """" & sCaseLocation & """" & ",")
                        ReportOutputFile.WriteLine("          " & """" & "Case Size On Disk" & """" & ": " & """" & sCaseSizeOnDisk & """" & ",")
                        ReportOutputFile.WriteLine("          " & """" & "Oldest Top Level Item" & """" & ": " & """" & sOldestItem & """" & ",")
                        ReportOutputFile.WriteLine("          " & """" & "Newest Top Level Item" & """" & ": " & """" & sNewestItem & """" & ",")
                        ReportOutputFile.WriteLine("          " & """" & "Evidence Name" & """" & ": " & """" & sEvidenceName & """" & ",")
                        ReportOutputFile.WriteLine("          " & """" & "Evidence Location" & """" & ": " & """" & sEvidenceLocation & """")
                        If iCounter < iRowCount - 1 Then
                            ReportOutputFile.WriteLine("        },")
                        Else
                            ReportOutputFile.WriteLine("        }")

                        End If
                    Case "Case Location"
                        ReportOutputFile.WriteLine("          " & """" & "Case GUID" & """" & ": " & """" & sCaseGUID & """" & ",")
                        ReportOutputFile.WriteLine("          " & """" & "Collection Status" & """" & ": " & """" & sCollectionStatus & """" & ",")
                        ReportOutputFile.WriteLine("          " & """" & "Report Load Duration" & """" & ": " & """" & sReportLoadDuration & """" & ",")
                        ReportOutputFile.WriteLine("          " & """" & "Current Case Version" & """" & ": " & """" & sCurrentCaseVersion & """" & ",")
                        ReportOutputFile.WriteLine("          " & """" & "Upgraded Case Version" & """" & ": " & """" & sUpgradedCaseVersion & """" & ",")
                        ReportOutputFile.WriteLine("          " & """" & "Is Compound Case" & """" & ": " & """" & sIsCompound & """" & ",")
                        ReportOutputFile.WriteLine("          " & """" & "Cases Contained" & """" & ": " & """" & sCompoundCaseContains & """" & ",")
                        ReportOutputFile.WriteLine("          " & """" & "Contained In Case" & """" & ": " & """" & sContainedInCase & """" & ",")
                        ReportOutputFile.WriteLine("          " & """" & "Batch Load Info" & """" & ": " & """" & sBatchLoadInfo & """" & ",")
                        ReportOutputFile.WriteLine("          " & """" & "Data Export" & """" & ": " & """" & sDataExport & """" & ",")
                        ReportOutputFile.WriteLine("          " & """" & "Case Location" & """" & ": " & """" & sCaseLocation & """" & ",")
                        ReportOutputFile.WriteLine("          " & """" & "Case Size On Disk" & """" & ": " & """" & sCaseSizeOnDisk & """")
                        ReportOutputFile.WriteLine("          " & """" & "Oldest Top Level Item" & """" & ": " & """" & sOldestItem & """" & ",")
                        ReportOutputFile.WriteLine("          " & """" & "Newest Top Level Item" & """" & ": " & """" & sNewestItem & """" & ",")
                        If iCounter < iRowCount - 1 Then
                            ReportOutputFile.WriteLine("        },")
                        Else
                            ReportOutputFile.WriteLine("        }")

                        End If
                    Case "Case Size"
                        ReportOutputFile.WriteLine("          " & """" & "Case GUID" & """" & ": " & """" & sCaseGUID & """" & ",")
                        ReportOutputFile.WriteLine("          " & """" & "Collection Status" & """" & ": " & """" & sCollectionStatus & """" & ",")
                        ReportOutputFile.WriteLine("          " & """" & "Report Load Duration" & """" & ": " & """" & sReportLoadDuration & """" & ",")
                        ReportOutputFile.WriteLine("          " & """" & "Current Case Version" & """" & ": " & """" & sCurrentCaseVersion & """" & ",")
                        ReportOutputFile.WriteLine("          " & """" & "Upgraded Case Version" & """" & ": " & """" & sUpgradedCaseVersion & """" & ",")
                        ReportOutputFile.WriteLine("          " & """" & "Batch Load Info" & """" & ": " & """" & sBatchLoadInfo & """" & ",")
                        ReportOutputFile.WriteLine("          " & """" & "Data Export" & """" & ": " & """" & sDataExport & """" & ",")
                        ReportOutputFile.WriteLine("          " & """" & "Is Compound Case" & """" & ": " & """" & sIsCompound & """" & ",")
                        ReportOutputFile.WriteLine("          " & """" & "Cases Contained" & """" & ": " & """" & sCompoundCaseContains & """" & ",")
                        ReportOutputFile.WriteLine("          " & """" & "Contained In Case" & """" & ": " & """" & sContainedInCase & """" & ",")
                        ReportOutputFile.WriteLine("          " & """" & "Case Location" & """" & ": " & """" & sCaseLocation & """" & ",")
                        ReportOutputFile.WriteLine("          " & """" & "Case Size On Disk" & """" & ": " & """" & sCaseSizeOnDisk & """" & ",")
                        ReportOutputFile.WriteLine("          " & """" & "Case File Size" & """" & ": " & """" & sCaseFileSize & """" & ",")
                        ReportOutputFile.WriteLine("          " & """" & "Case Audit Size" & """" & ": " & """" & sCaseAuditSize & """" & ",")
                        ReportOutputFile.WriteLine("          " & """" & "Oldest Top Level Item" & """" & ": " & """" & sOldestItem & """" & ",")
                        ReportOutputFile.WriteLine("          " & """" & "Newest Top Level Item" & """" & ": " & """" & sNewestItem & """" & ",")
                        ReportOutputFile.WriteLine("          " & """" & "Load Events" & """" & ": " & """" & sLoadEvents & """" & ",")
                        ReportOutputFile.WriteLine("          " & """" & "Total Load Time" & """" & ": " & """" & sTotalLoadTime & """" & ",")
                        ReportOutputFile.WriteLine("          " & """" & "Processing Speed" & """" & ": " & """" & sProcessingSpeed & """")
                        If iCounter < iRowCount - 1 Then
                            ReportOutputFile.WriteLine("        },")
                        Else
                            ReportOutputFile.WriteLine("        }")

                        End If
                    Case "Custodians in Case"
                        ReportOutputFile.WriteLine("          " & """" & "Case GUID" & """" & ": " & """" & sCaseGUID & """" & ",")
                        ReportOutputFile.WriteLine("          " & """" & "Collection Status" & """" & ": " & """" & sCollectionStatus & """" & ",")
                        ReportOutputFile.WriteLine("          " & """" & "Report Load Duration" & """" & ": " & """" & sReportLoadDuration & """" & ",")
                        ReportOutputFile.WriteLine("          " & """" & "Current Case Version" & """" & ": " & """" & sCurrentCaseVersion & """" & ",")
                        ReportOutputFile.WriteLine("          " & """" & "Upgraded Case Version" & """" & ": " & """" & sUpgradedCaseVersion & """" & ",")
                        ReportOutputFile.WriteLine("          " & """" & "Batch Load Info" & """" & ": " & """" & sBatchLoadInfo & """" & ",")
                        ReportOutputFile.WriteLine("          " & """" & "Data Export" & """" & ": " & """" & sDataExport & """" & ",")
                        ReportOutputFile.WriteLine("          " & """" & "Is Compound Case" & """" & ": " & """" & sIsCompound & """" & ",")
                        ReportOutputFile.WriteLine("          " & """" & "Cases Contained" & """" & ": " & """" & sCompoundCaseContains & """" & ",")
                        ReportOutputFile.WriteLine("          " & """" & "Contained In Case" & """" & ": " & """" & sContainedInCase & """" & ",")
                        ReportOutputFile.WriteLine("          " & """" & "Case Location" & """" & ": " & """" & sCaseLocation & """" & ",")
                        ReportOutputFile.WriteLine("          " & """" & "Case Size On Disk" & """" & ": " & """" & sCaseSizeOnDisk & """" & ",")
                        ReportOutputFile.WriteLine("          " & """" & "Oldest Top Level Item" & """" & ": " & """" & sOldestItem & """" & ",")
                        ReportOutputFile.WriteLine("          " & """" & "Newest Top Level Item" & """" & ": " & """" & sNewestItem & """" & ",")
                        ReportOutputFile.WriteLine("          " & """" & "Investigator" & """" & ": " & """" & sInvestigator & """" & ",")
                        ReportOutputFile.WriteLine("          " & """" & "Investigator Sessions" & """" & ": " & """" & sInvestigatorSessions & """" & ",")
                        ReportOutputFile.WriteLine("          " & """" & "Invalid Sessions" & """" & ": " & """" & sInvalidSessions & """" & ",")
                        ReportOutputFile.WriteLine("          " & """" & "Investigator Time Summary" & """" & ": " & """" & sInvestigatorTimeSummary & """" & ",")
                        ReportOutputFile.WriteLine("          " & """" & "Custodians" & """" & ": " & """" & sCustodians & """" & ",")
                        ReportOutputFile.WriteLine("          " & """" & "Custodian Count" & """" & ": " & """" & sCustodianCount & """")
                        If iCounter < iRowCount - 1 Then
                            ReportOutputFile.WriteLine("        },")
                        Else
                            ReportOutputFile.WriteLine("        }")

                        End If
                    Case "Metadata Type"
                        ReportOutputFile.WriteLine("          " & """" & "Case GUID" & """" & ": " & """" & sCaseGUID & """" & ",")
                        ReportOutputFile.WriteLine("          " & """" & "Collection Status" & """" & ": " & """" & sCollectionStatus & """" & ",")
                        ReportOutputFile.WriteLine("          " & """" & "Report Load Duration" & """" & ": " & """" & sReportLoadDuration & """" & ",")
                        ReportOutputFile.WriteLine("          " & """" & "Current Case Version" & """" & ": " & """" & sCurrentCaseVersion & """" & ",")
                        ReportOutputFile.WriteLine("          " & """" & "Upgraded Case Version" & """" & ": " & """" & sUpgradedCaseVersion & """" & ",")
                        ReportOutputFile.WriteLine("          " & """" & "Batch Load Info" & """" & ": " & """" & sBatchLoadInfo & """" & ",")
                        ReportOutputFile.WriteLine("          " & """" & "Data Export" & """" & ": " & """" & sDataExport & """" & ",")
                        ReportOutputFile.WriteLine("          " & """" & "Is Compound Case" & """" & ": " & """" & sIsCompound & """" & ",")
                        ReportOutputFile.WriteLine("          " & """" & "Cases Contained" & """" & ": " & """" & sCompoundCaseContains & """" & ",")
                        ReportOutputFile.WriteLine("          " & """" & "Contained In Case" & """" & ": " & """" & sContainedInCase & """" & ",")
                        ReportOutputFile.WriteLine("          " & """" & "Case Location" & """" & ": " & """" & sCaseLocation & """" & ",")
                        ReportOutputFile.WriteLine("          " & """" & "Case Size On Disk" & """" & ": " & """" & sCaseSizeOnDisk & """" & ",")
                        ReportOutputFile.WriteLine("          " & """" & "Case File Size" & """" & ": " & """" & sCaseFileSize & """" & ",")
                        ReportOutputFile.WriteLine("          " & """" & "Case Audit Size" & """" & ": " & """" & sCaseAuditSize & """" & ",")
                        ReportOutputFile.WriteLine("          " & """" & "Oldest Top Level Item" & """" & ": " & """" & sOldestItem & """" & ",")
                        ReportOutputFile.WriteLine("          " & """" & "Newest Top Level Item" & """" & ": " & """" & sNewestItem & """" & ",")
                        ReportOutputFile.WriteLine("          " & """" & "Mime Types" & """" & ": " & """" & sMimeTypes & """" & ",")
                        ReportOutputFile.WriteLine("          " & """" & "Item Types" & """" & ": " & """" & sItemTypes & """")
                        If iCounter < iRowCount - 1 Then
                            ReportOutputFile.WriteLine("        },")
                        Else
                            ReportOutputFile.WriteLine("        }")

                        End If
                    Case "Processing time"
                        ReportOutputFile.WriteLine("          " & """" & "Case GUID" & """" & ": " & """" & sCaseGUID & """" & ",")
                        ReportOutputFile.WriteLine("          " & """" & "Collection Status" & """" & ": " & """" & sCollectionStatus & """" & ",")
                        ReportOutputFile.WriteLine("          " & """" & "Report Load Duration" & """" & ": " & """" & sReportLoadDuration & """" & ",")
                        ReportOutputFile.WriteLine("          " & """" & "Current Case Version" & """" & ": " & """" & sCurrentCaseVersion & """" & ",")
                        ReportOutputFile.WriteLine("          " & """" & "Upgraded Case Version" & """" & ": " & """" & sUpgradedCaseVersion & """" & ",")
                        ReportOutputFile.WriteLine("          " & """" & "Batch Load Info" & """" & ": " & """" & sBatchLoadInfo & """" & ",")
                        ReportOutputFile.WriteLine("          " & """" & "Data Export" & """" & ": " & """" & sDataExport & """" & ",")
                        ReportOutputFile.WriteLine("          " & """" & "Is Compound Case" & """" & ": " & """" & sIsCompound & """" & ",")
                        ReportOutputFile.WriteLine("          " & """" & "Cases Contained" & """" & ": " & """" & sCompoundCaseContains & """" & ",")
                        ReportOutputFile.WriteLine("          " & """" & "Contained In Case" & """" & ": " & """" & sContainedInCase & """" & ",")
                        ReportOutputFile.WriteLine("          " & """" & "Case Location" & """" & ": " & """" & sCaseLocation & """" & ",")
                        ReportOutputFile.WriteLine("          " & """" & "Case Size On Disk" & """" & ": " & """" & sCaseSizeOnDisk & """" & ",")
                        ReportOutputFile.WriteLine("          " & """" & "Case File Size" & """" & ": " & """" & sCaseFileSize & """" & ",")
                        ReportOutputFile.WriteLine("          " & """" & "Case Audit Size" & """" & ": " & """" & sCaseAuditSize & """" & ",")
                        ReportOutputFile.WriteLine("          " & """" & "Oldest Top Level Item" & """" & ": " & """" & sOldestItem & """" & ",")
                        ReportOutputFile.WriteLine("          " & """" & "Newest Top Level Item" & """" & ": " & """" & sNewestItem & """" & ",")
                        ReportOutputFile.WriteLine("          " & """" & "Load Events" & """" & ": " & """" & sLoadEvents & """" & ",")
                        ReportOutputFile.WriteLine("          " & """" & "Total Load Time" & """" & ": " & """" & sTotalLoadTime & """" & ",")
                        ReportOutputFile.WriteLine("          " & """" & "Processing Speed" & """" & ": " & """" & sProcessingSpeed & """")
                        If iCounter < iRowCount - 1 Then
                            ReportOutputFile.WriteLine("        },")
                        Else
                            ReportOutputFile.WriteLine("        }")

                        End If
                    Case "Processing speed(GB / Hour)"
                        ReportOutputFile.WriteLine("          " & """" & "Case GUID" & """" & ": " & """" & sCaseGUID & """" & ",")
                        ReportOutputFile.WriteLine("          " & """" & "Collection Status" & """" & ": " & """" & sCollectionStatus & """" & ",")
                        ReportOutputFile.WriteLine("          " & """" & "Report Load Duration" & """" & ": " & """" & sReportLoadDuration & """" & ",")
                        ReportOutputFile.WriteLine("          " & """" & "Current Case Version" & """" & ": " & """" & sCurrentCaseVersion & """" & ",")
                        ReportOutputFile.WriteLine("          " & """" & "Upgraded Case Version" & """" & ": " & """" & sUpgradedCaseVersion & """" & ",")
                        ReportOutputFile.WriteLine("          " & """" & "Batch Load Info" & """" & ": " & """" & sBatchLoadInfo & """" & ",")
                        ReportOutputFile.WriteLine("          " & """" & "Data Export" & """" & ": " & """" & sDataExport & """" & ",")
                        ReportOutputFile.WriteLine("          " & """" & "Is Compound Case" & """" & ": " & """" & sIsCompound & """" & ",")
                        ReportOutputFile.WriteLine("          " & """" & "Cases Contained" & """" & ": " & """" & sCompoundCaseContains & """" & ",")
                        ReportOutputFile.WriteLine("          " & """" & "Contained In Case" & """" & ": " & """" & sContainedInCase & """" & ",")
                        ReportOutputFile.WriteLine("          " & """" & "Case Location" & """" & ": " & """" & sCaseLocation & """" & ",")
                        ReportOutputFile.WriteLine("          " & """" & "Case Size On Disk" & """" & ": " & """" & sCaseSizeOnDisk & """" & ",")
                        ReportOutputFile.WriteLine("          " & """" & "Case File Size" & """" & ": " & """" & sCaseFileSize & """" & ",")
                        ReportOutputFile.WriteLine("          " & """" & "Case Audit Size" & """" & ": " & """" & sCaseAuditSize & """" & ",")
                        ReportOutputFile.WriteLine("          " & """" & "Oldest Top Level Item" & """" & ": " & """" & sOldestItem & """" & ",")
                        ReportOutputFile.WriteLine("          " & """" & "Newest Top Level Item" & """" & ": " & """" & sNewestItem & """" & ",")
                        ReportOutputFile.WriteLine("          " & """" & "Load Events" & """" & ": " & """" & sLoadEvents & """" & ",")
                        ReportOutputFile.WriteLine("          " & """" & "Total Load Time" & """" & ": " & """" & sTotalLoadTime & """" & ",")
                        ReportOutputFile.WriteLine("          " & """" & "Processing Speed" & """" & ": " & """" & sProcessingSpeed & """")
                        If iCounter < iRowCount - 1 Then
                            ReportOutputFile.WriteLine("        },")
                        Else
                            ReportOutputFile.WriteLine("        }")

                        End If
                    Case "Search Term Hit"
                        ReportOutputFile.WriteLine("          " & """" & "Case GUID" & """" & ": " & """" & sCaseGUID & """" & ",")
                        ReportOutputFile.WriteLine("          " & """" & "Collection Status" & """" & ": " & """" & sCollectionStatus & """" & ",")
                        ReportOutputFile.WriteLine("          " & """" & "Report Load Duration" & """" & ": " & """" & sReportLoadDuration & """" & ",")
                        ReportOutputFile.WriteLine("          " & """" & "Current Case Version" & """" & ": " & """" & sCurrentCaseVersion & """" & ",")
                        ReportOutputFile.WriteLine("          " & """" & "Upgraded Case Version" & """" & ": " & """" & sUpgradedCaseVersion & """" & ",")
                        ReportOutputFile.WriteLine("          " & """" & "Batch Load Info" & """" & ": " & """" & sBatchLoadInfo & """" & ",")
                        ReportOutputFile.WriteLine("          " & """" & "Data Export" & """" & ": " & """" & sDataExport & """" & ",")
                        ReportOutputFile.WriteLine("          " & """" & "Is Compound Case" & """" & ": " & """" & sIsCompound & """" & ",")
                        ReportOutputFile.WriteLine("          " & """" & "Cases Contained" & """" & ": " & """" & sCompoundCaseContains & """" & ",")
                        ReportOutputFile.WriteLine("          " & """" & "Contained In Case" & """" & ": " & """" & sContainedInCase & """" & ",")
                        ReportOutputFile.WriteLine("          " & """" & "Case Location" & """" & ": " & """" & sCaseLocation & """" & ",")
                        ReportOutputFile.WriteLine("          " & """" & "Case Size On Disk" & """" & ": " & """" & sCaseSizeOnDisk & """" & ",")
                        ReportOutputFile.WriteLine("          " & """" & "Case File Size" & """" & ": " & """" & sCaseFileSize & """" & ",")
                        ReportOutputFile.WriteLine("          " & """" & "Case Audit Size" & """" & ": " & """" & sCaseAuditSize & """" & ",")
                        ReportOutputFile.WriteLine("          " & """" & "Oldest Top Level Item" & """" & ": " & """" & sOldestItem & """" & ",")
                        ReportOutputFile.WriteLine("          " & """" & "Newest Top Level Item" & """" & ": " & """" & sNewestItem & """" & ",")
                        ReportOutputFile.WriteLine("          " & """" & "sSearchTerm" & """" & ": " & """" & sSearchTerm & """" & ",")
                        ReportOutputFile.WriteLine("          " & """" & "Search Size" & """" & ": " & """" & sSearchSize & """" & ",")
                        ReportOutputFile.WriteLine("          " & """" & "Search Hit Count" & """" & ": " & """" & sSearchHitCount & """" & ",")
                        ReportOutputFile.WriteLine("          " & """" & "Custodian Search Hit" & """" & ": " & """" & sCustodianSearchHit & """" & ",")
                        ReportOutputFile.WriteLine("          " & """" & "Total Item Count" & """" & ": " & """" & sTotalItemCount & """" & ",")
                        ReportOutputFile.WriteLine("          " & """" & "Hit Count Percent" & """" & ": " & """" & sHitCountPercent & """")
                        If iCounter < iRowCount - 1 Then
                            ReportOutputFile.WriteLine("        },")
                        Else
                            ReportOutputFile.WriteLine("        }")

                        End If
                    Case "Total Number of Items"
                        ReportOutputFile.WriteLine("          " & """" & "Case GUID" & """" & ": " & """" & sCaseGUID & """" & ",")
                        ReportOutputFile.WriteLine("          " & """" & "Collection Status" & """" & ": " & """" & sCollectionStatus & """" & ",")
                        ReportOutputFile.WriteLine("          " & """" & "Report Load Duration" & """" & ": " & """" & sReportLoadDuration & """" & ",")
                        ReportOutputFile.WriteLine("          " & """" & "Current Case Version" & """" & ": " & """" & sCurrentCaseVersion & """" & ",")
                        ReportOutputFile.WriteLine("          " & """" & "Upgraded Case Version" & """" & ": " & """" & sUpgradedCaseVersion & """" & ",")
                        ReportOutputFile.WriteLine("          " & """" & "Is Compound Case" & """" & ": " & """" & sIsCompound & """" & ",")
                        ReportOutputFile.WriteLine("          " & """" & "Cases Contained" & """" & ": " & """" & sCompoundCaseContains & """" & ",")
                        ReportOutputFile.WriteLine("          " & """" & "Contained In Case" & """" & ": " & """" & sContainedInCase & """" & ",")
                        ReportOutputFile.WriteLine("          " & """" & "Batch Load Info" & """" & ": " & """" & sBatchLoadInfo & """" & ",")
                        ReportOutputFile.WriteLine("          " & """" & "Data Export" & """" & ": " & """" & sDataExport & """" & ",")
                        ReportOutputFile.WriteLine("          " & """" & "Case Location" & """" & ": " & """" & sCaseLocation & """" & ",")
                        ReportOutputFile.WriteLine("          " & """" & "Case Size On Disk" & """" & ": " & """" & sCaseSizeOnDisk & """" & ",")
                        ReportOutputFile.WriteLine("          " & """" & "Case File Size" & """" & ": " & """" & sCaseFileSize & """" & ",")
                        ReportOutputFile.WriteLine("          " & """" & "Case Audit Size" & """" & ": " & """" & sCaseAuditSize & """" & ",")
                        ReportOutputFile.WriteLine("          " & """" & "Oldest Top Level Item" & """" & ": " & """" & sOldestItem & """" & ",")
                        ReportOutputFile.WriteLine("          " & """" & "Newest Top Level Item" & """" & ": " & """" & sNewestItem & """" & ",")
                        ReportOutputFile.WriteLine("          " & """" & "Total Item Count" & """" & ": " & """" & sTotalItemCount & """")
                        If iCounter < iRowCount - 1 Then
                            ReportOutputFile.WriteLine("        },")
                        Else
                            ReportOutputFile.WriteLine("        }")

                        End If
                    Case "Total Number of workers"
                        ReportOutputFile.WriteLine("          " & """" & "Case GUID" & """" & ": " & """" & sCaseGUID & """" & ",")
                        ReportOutputFile.WriteLine("          " & """" & "Collection Status" & """" & ": " & """" & sCollectionStatus & """" & ",")
                        ReportOutputFile.WriteLine("          " & """" & "Report Load Duration" & """" & ": " & """" & sReportLoadDuration & """" & ",")
                        ReportOutputFile.WriteLine("          " & """" & "Current Case Version" & """" & ": " & """" & sCurrentCaseVersion & """" & ",")
                        ReportOutputFile.WriteLine("          " & """" & "Upgraded Case Version" & """" & ": " & """" & sUpgradedCaseVersion & """" & ",")
                        ReportOutputFile.WriteLine("          " & """" & "Batch Load Info" & """" & ": " & """" & sBatchLoadInfo & """" & ",")
                        ReportOutputFile.WriteLine("          " & """" & "Data Export" & """" & ": " & """" & sDataExport & """" & ",")
                        ReportOutputFile.WriteLine("          " & """" & "Is Compound Case" & """" & ": " & """" & sIsCompound & """" & ",")
                        ReportOutputFile.WriteLine("          " & """" & "Cases Contained" & """" & ": " & """" & sCompoundCaseContains & """" & ",")
                        ReportOutputFile.WriteLine("          " & """" & "Contained In Case" & """" & ": " & """" & sContainedInCase & """" & ",")
                        ReportOutputFile.WriteLine("          " & """" & "Case Location" & """" & ": " & """" & sCaseLocation & """" & ",")
                        ReportOutputFile.WriteLine("          " & """" & "Case Size On Disk" & """" & ": " & """" & sCaseSizeOnDisk & """" & ",")
                        ReportOutputFile.WriteLine("          " & """" & "Case File Size" & """" & ": " & """" & sCaseFileSize & """" & ",")
                        ReportOutputFile.WriteLine("          " & """" & "Case Audit Size" & """" & ": " & """" & sCaseAuditSize & """" & ",")
                        ReportOutputFile.WriteLine("          " & """" & "Oldest Top Level Item" & """" & ": " & """" & sOldestItem & """" & ",")
                        ReportOutputFile.WriteLine("          " & """" & "Newest Top Level Item" & """" & ": " & """" & sNewestItem & """" & ",")
                        ReportOutputFile.WriteLine("          " & """" & "Investigator" & """" & ": " & """" & sInvestigator & """" & ",")
                        ReportOutputFile.WriteLine("          " & """" & "Investigator Sessions" & """" & ": " & """" & sInvestigatorSessions & """" & ",")
                        ReportOutputFile.WriteLine("          " & """" & "Invalid Sessions" & """" & ": " & """" & sInvalidSessions & """" & ",")
                        ReportOutputFile.WriteLine("          " & """" & "Investigator Time Summary" & """" & ": " & """" & sInvestigatorTimeSummary & """" & ",")
                        ReportOutputFile.WriteLine("          " & """" & "Broker Memory" & """" & ": " & """" & sBrokerMemory & """" & ",")
                        ReportOutputFile.WriteLine("          " & """" & "Worker Count" & """" & ": " & """" & sWorkerCount & """" & ",")
                        ReportOutputFile.WriteLine("          " & """" & "Worker Memory" & """" & ": " & """" & sWorkerMemory & """" & ",")
                        ReportOutputFile.WriteLine("          " & """" & "Load Events" & """" & ": " & """" & sLoadEvents & """" & ",")
                        ReportOutputFile.WriteLine("          " & """" & "Total Load Time" & """" & ": " & """" & sTotalLoadTime & """" & ",")
                        ReportOutputFile.WriteLine("          " & """" & "Processing Speed" & """" & ": " & """" & sProcessingSpeed & """")
                        If iCounter < iRowCount - 1 Then
                            ReportOutputFile.WriteLine("        },")
                        Else
                            ReportOutputFile.WriteLine("        }")

                        End If
                End Select
                iCounter = iCounter + 1
            End If
            row.cells("CollectionStatus").value = "Case Data Exported"
        Next
        ReportOutputFile.WriteLine("        }")
        ReportOutputFile.WriteLine("    }")
        ReportOutputFile.WriteLine("}")

        ReportOutputFile.Close()
        MessageBox.Show("Nuix Case Statistics report finished building.  Report located at: " & sOutputFileName)
    End Sub

    Private Sub XMLToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles XMLToolStripMenuItem.Click
        Dim sReportFilePath As String
        Dim sReportType As String
        Dim sOutputFileName As String
        Dim sMachineName As String

        Dim sCaseGUID As String
        Dim sCollectionStatus As String
        Dim sReportLoadDuration As String
        Dim sCaseName As String
        Dim sCurrentCaseVersion As String
        Dim sUpgradedCaseVersion As String
        Dim sIsCompound As String
        Dim sCompoundCaseContains As String
        Dim sContainedInCase As String
        Dim sBatchLoadInfo As String
        Dim sDataExport As String
        Dim sCaseLocation As String
        Dim sCaseSizeOnDisk As String
        Dim sCaseFileSize As String
        Dim sCaseAuditSize As String
        Dim sInvestigator As String
        Dim sInvestigatorSessions As String
        Dim sInvalidSessions As String
        Dim sInvestigatorTimeSummary As String
        Dim sBrokerMemory As String
        Dim sWorkerCount As String
        Dim sWorkerMemory As String
        Dim sEvidenceName As String
        Dim sEvidenceLocation As String
        Dim sMimeTypes As String
        Dim sItemTypes As String
        Dim sCreationDate As String
        Dim sModifiedDate As String
        Dim sProcessingSpeed As String
        Dim sCustodians As String
        Dim sCustodianCount As String
        Dim sSearchTerm As String
        Dim sSearchSize As String
        Dim sSearchHitCount As String
        Dim sCustodianSearchHit As String
        Dim sTotalItemCount As String
        Dim sHitCountPercent As String
        Dim sLoadEvents As String
        Dim sTotalLoadTime As String
        Dim sOldestItem As String
        Dim sNewestItem As String
        Dim xmlDeclaration As Xml.XmlDeclaration
        Dim NuixCaseReportXML As Xml.XmlDocument
        Dim CaseReportRoot As Xml.XmlElement
        Dim ReportType As Xml.XmlElement
        Dim CaseGUID As Xml.XmlElement
        Dim CaseName As Xml.XmlElement
        Dim CollectionStatus As Xml.XmlElement
        Dim ReportLoadDuration As Xml.XmlElement
        Dim CurrentCaseVersion As Xml.XmlElement
        Dim UpgradedCaseVersion As Xml.XmlElement
        Dim BatchLoadInfo As Xml.XmlElement
        Dim IsCompoound As Xml.XmlElement
        Dim CasesContained As Xml.XmlElement
        Dim ContainedInCase As Xml.XmlElement
        Dim CaseLocation As Xml.XmlElement
        Dim CaseSizeOnDisk As Xml.XmlElement
        Dim CaseFileSize As Xml.XmlElement
        Dim CaseAuditSize As Xml.XmlElement
        Dim OldestItem As Xml.XmlElement
        Dim NewestItem As Xml.XmlElement
        Dim Investigator As Xml.XmlElement
        Dim InvestigatorSessions As Xml.XmlElement
        Dim InvalidSessions As Xml.XmlElement
        Dim InvestigatorTimeSummary As Xml.XmlElement
        Dim BrokerMemory As Xml.XmlElement
        Dim WorkerCount As Xml.XmlElement
        Dim WorkerMemory As Xml.XmlElement
        Dim EvidenceName As Xml.XmlElement
        Dim EvidenceLocation As Xml.XmlElement
        Dim MimeTypes As Xml.XmlElement
        Dim ItemTypes As Xml.XmlElement
        Dim CreationDate As Xml.XmlElement
        Dim ModifiedDate As Xml.XmlElement
        Dim ProcessingSpeed As Xml.XmlElement
        Dim Custodians As Xml.XmlElement
        Dim CustodianCount As Xml.XmlElement
        Dim SearchTerm As Xml.XmlElement
        Dim SearchSize As Xml.XmlElement
        Dim SearchHitCount As Xml.XmlElement
        Dim CustodianSearchHit As Xml.XmlElement
        Dim TotalItemCount As Xml.XmlElement
        Dim HitCountPercent As Xml.XmlElement
        Dim LoadEvents As Xml.XmlElement
        Dim DataExport As Xml.XmlElement
        Dim TotalLoadTime As Xml.XmlElement

        If cboReportType.Text = vbNullString Then
            MessageBox.Show("You must enter a report type to export the data for.", "No Report Type selected")
            cboReportType.Focus()
            Exit Sub
        End If

        sReportFilePath = txtReportLocation.Text
        sReportType = cboReportType.Text
        sMachineName = System.Net.Dns.GetHostName()
        sOutputFileName = sMachineName & "-" & sReportType & "-" & DateTime.Now.ToString("yyyy-MM-dd hh-mm-ss") & ".xml"

        NuixCaseReportXML = New Xml.XmlDocument
        xmlDeclaration = NuixCaseReportXML.CreateXmlDeclaration("1.0", "UTF-8", "yes")
        NuixCaseReportXML.InsertBefore(xmlDeclaration, CaseReportRoot)
        CaseReportRoot = NuixCaseReportXML.CreateElement("NuixCaseReport")
        NuixCaseReportXML.AppendChild(CaseReportRoot)

        ReportType = NuixCaseReportXML.CreateElement(cboReportType.Text.Replace(" ", ""))
        CaseReportRoot.AppendChild(ReportType)

        For Each row In grdCaseInfo.Rows
            If (row.cells("CaseName").value <> vbNullString) Then
                row.cells("CollectionStatus").value = "Exporting Case Data..."
                sCaseGUID = row.cells("CaseGUID").value
                sCollectionStatus = row.cells("CollectionStatus").value
                sReportLoadDuration = row.cells("ReportLoadDuration").value
                sCaseName = row.cells("CaseName").value
                sBatchLoadInfo = row.cells("BatchLoadInfo").value
                sDataExport = row.cells("DataExport").value
                sCurrentCaseVersion = row.cells("CurrentCaseVersion").value
                sUpgradedCaseVersion = row.cells("UpgradedCaseVersion").value
                sCaseLocation = row.cells("CaseLocation").value
                sCaseSizeOnDisk = row.cells("CaseSizeOnDisk").value
                sCaseFileSize = row.cells("CaseFileSize").value
                sCaseAuditSize = row.cells("CaseAuditSize").value
                sOldestItem = row.cells("OldestTopLevel").value
                sNewestItem = row.cells("NewestTopLevel").value
                sIsCompound = row.cells("IsCompound").value
                sCompoundCaseContains = row.cells("CasesContained").value
                sContainedInCase = row.cells("ContainedInCase").value
                sInvestigator = row.cells("Investigator").value
                sInvestigatorSessions = row.cells("InvestigatorSessions").value
                sInvalidSessions = row.cells("InvalidSessions").value
                sInvestigatorTimeSummary = row.cells("InvestigatorTimeSummary").value
                sBrokerMemory = row.cells("BrokerMemory").value
                sWorkerCount = row.cells("WorkerCount").value
                sWorkerMemory = row.cells("WorkerMemory").value
                sEvidenceName = row.cells("EvidenceName").value
                sEvidenceLocation = row.cells("EvidenceLocation").value
                sMimeTypes = row.cells("MimeTypes").value
                sItemTypes = row.cells("ItemTypes").value
                sCreationDate = row.cells("CreationDate").value
                sModifiedDate = row.cells("ModifiedDate").value
                sLoadEvents = row.cells("LoadEvents").value
                sTotalLoadTime = row.cells("TotalLoadTime").value
                sProcessingSpeed = row.cells("ProcessingSpeed").value
                sCustodians = row.cells("Custodians").value
                sCustodianCount = row.cells("CustodianCount").value
                sSearchTerm = row.cells("SearchTerm").value
                sSearchSize = row.cells("SearchSize").value
                sSearchHitCount = row.cells("SearchHitCount").value
                sCustodianSearchHit = row.cells("CustodianSearchHit").value
                sTotalItemCount = row.cells("TotalCaseItemCount").value
                sHitCountPercent = row.cells("HitCountPercent").value

                Select Case cboReportType.Text
                    Case "All"
                        CaseGUID = NuixCaseReportXML.CreateElement("CaseGUID")
                        ReportType.AppendChild(CaseGUID)
                        CaseGUID.SetAttribute("CaseGUID", sCaseGUID)

                        CaseName = NuixCaseReportXML.CreateElement("CaseName")
                        CaseGUID.AppendChild(CaseName)
                        CaseName.InnerText = sCaseName

                        CollectionStatus = NuixCaseReportXML.CreateElement("CollectionStatus")
                        CaseGUID.AppendChild(CollectionStatus)
                        CollectionStatus.InnerText = sCollectionStatus

                        ReportLoadDuration = NuixCaseReportXML.CreateElement("ReportLoadDuration")
                        CaseGUID.AppendChild(ReportLoadDuration)
                        ReportLoadDuration.InnerText = sReportLoadDuration

                        CurrentCaseVersion = NuixCaseReportXML.CreateElement("CurrentCaseVersion")
                        CaseGUID.AppendChild(CurrentCaseVersion)
                        CurrentCaseVersion.InnerText = sCurrentCaseVersion

                        UpgradedCaseVersion = NuixCaseReportXML.CreateElement("UpgradedCaseVersion")
                        CaseGUID.AppendChild(UpgradedCaseVersion)
                        UpgradedCaseVersion.InnerText = sUpgradedCaseVersion

                        BatchLoadInfo = NuixCaseReportXML.CreateElement("BatchLoadInfo")
                        CaseGUID.AppendChild(BatchLoadInfo)
                        BatchLoadInfo.InnerText = sBatchLoadInfo

                        DataExport = NuixCaseReportXML.CreateElement("DataExport")
                        CaseGUID.AppendChild(DataExport)
                        BatchLoadInfo.InnerText = sDataExport

                        IsCompoound = NuixCaseReportXML.CreateElement("IsCompoundCase")
                        CaseGUID.AppendChild(IsCompoound)
                        IsCompoound.InnerText = sIsCompound

                        CasesContained = NuixCaseReportXML.CreateElement("CasesContained")
                        CaseGUID.AppendChild(CasesContained)
                        CasesContained.InnerText = sCompoundCaseContains

                        ContainedInCase = NuixCaseReportXML.CreateElement("ContainedInCase")
                        CaseGUID.AppendChild(ContainedInCase)
                        ContainedInCase.InnerText = sContainedInCase

                        CaseLocation = NuixCaseReportXML.CreateElement("CaseLocation")
                        CaseGUID.AppendChild(CaseLocation)
                        CaseLocation.InnerText = sCaseLocation

                        CaseSizeOnDisk = NuixCaseReportXML.CreateElement("CaseSizeOnDisk")
                        CaseGUID.AppendChild(CaseSizeOnDisk)
                        CaseSizeOnDisk.InnerText = sCaseSizeOnDisk

                        CaseFileSize = NuixCaseReportXML.CreateElement("CaseFileSize")
                        CaseGUID.AppendChild(CaseFileSize)
                        CaseFileSize.InnerText = sCaseFileSize

                        CaseAuditSize = NuixCaseReportXML.CreateElement("CaseAuditSize")
                        CaseGUID.AppendChild(CaseAuditSize)
                        CaseAuditSize.InnerText = sCaseAuditSize

                        OldestItem = NuixCaseReportXML.CreateElement("OldestTopLevelItem")
                        CaseGUID.AppendChild(OldestItem)
                        OldestItem.InnerText = sOldestItem

                        NewestItem = NuixCaseReportXML.CreateElement("NewesTopLevelItem")
                        CaseGUID.AppendChild(NewestItem)
                        NewestItem.InnerText = sNewestItem

                        Investigator = NuixCaseReportXML.CreateElement("Investigator")
                        CaseGUID.AppendChild(Investigator)
                        Investigator.InnerText = sInvestigator

                        InvestigatorSessions = NuixCaseReportXML.CreateElement("InvestigatorSessions")
                        CaseGUID.AppendChild(InvestigatorSessions)
                        InvestigatorSessions.InnerText = sInvestigatorSessions

                        InvalidSessions = NuixCaseReportXML.CreateElement("InvalidSessions")
                        CaseGUID.AppendChild(InvalidSessions)
                        InvalidSessions.InnerText = sInvalidSessions

                        InvestigatorTimeSummary = NuixCaseReportXML.CreateElement("InvestigatorTimeSummary")
                        CaseGUID.AppendChild(InvestigatorTimeSummary)
                        InvestigatorTimeSummary.InnerText = sInvestigatorTimeSummary

                        BrokerMemory = NuixCaseReportXML.CreateElement("BrokerMemory")
                        CaseGUID.AppendChild(BrokerMemory)
                        BrokerMemory.InnerText = sBrokerMemory

                        WorkerCount = NuixCaseReportXML.CreateElement("WorkerCount")
                        CaseGUID.AppendChild(WorkerCount)
                        WorkerCount.InnerText = sWorkerCount

                        WorkerMemory = NuixCaseReportXML.CreateElement("WorkerCount")
                        CaseGUID.AppendChild(WorkerMemory)
                        WorkerMemory.InnerText = sWorkerMemory

                        EvidenceName = NuixCaseReportXML.CreateElement("EvidenceName")
                        CaseGUID.AppendChild(EvidenceName)
                        EvidenceName.InnerText = sEvidenceName

                        EvidenceLocation = NuixCaseReportXML.CreateElement("EvidenceLocation")
                        CaseGUID.AppendChild(EvidenceLocation)
                        EvidenceLocation.InnerText = sEvidenceLocation

                        MimeTypes = NuixCaseReportXML.CreateElement("MimeTypes")
                        CaseGUID.AppendChild(MimeTypes)
                        MimeTypes.InnerText = sMimeTypes

                        ItemTypes = NuixCaseReportXML.CreateElement("ItemTypes")
                        CaseGUID.AppendChild(ItemTypes)
                        ItemTypes.InnerText = sItemTypes

                        CreationDate = NuixCaseReportXML.CreateElement("CreationDate")
                        CaseGUID.AppendChild(CreationDate)
                        CreationDate.InnerText = sCreationDate

                        ModifiedDate = NuixCaseReportXML.CreateElement("ModifiedDate")
                        CaseGUID.AppendChild(ModifiedDate)
                        ModifiedDate.InnerText = sModifiedDate

                        LoadEvents = NuixCaseReportXML.CreateElement("LoadEvents")
                        CaseGUID.AppendChild(LoadEvents)
                        LoadEvents.InnerText = sLoadEvents

                        TotalLoadTime = NuixCaseReportXML.CreateElement("TotalLoadTime")
                        CaseGUID.AppendChild(TotalLoadTime)
                        TotalLoadTime.InnerText = sTotalLoadTime

                        ProcessingSpeed = NuixCaseReportXML.CreateElement("ProcessingSpeed")
                        CaseGUID.AppendChild(ProcessingSpeed)
                        ProcessingSpeed.InnerText = sProcessingSpeed

                        Custodians = NuixCaseReportXML.CreateElement("Custodians")

                        CaseGUID.AppendChild(Custodians)
                        Custodians.InnerText = sCustodians

                        CustodianCount = NuixCaseReportXML.CreateElement("CustodianCount")
                        CaseGUID.AppendChild(CustodianCount)
                        CustodianCount.InnerText = sCustodianCount

                        SearchTerm = NuixCaseReportXML.CreateElement("SearchTerm")
                        CaseGUID.AppendChild(SearchTerm)
                        SearchTerm.InnerText = sSearchTerm

                        SearchSize = NuixCaseReportXML.CreateElement("SearchSize")
                        CaseGUID.AppendChild(SearchSize)
                        SearchSize.InnerText = sSearchSize

                        SearchHitCount = NuixCaseReportXML.CreateElement("SearchHitCount")
                        CaseGUID.AppendChild(SearchHitCount)
                        SearchHitCount.InnerText = sSearchHitCount

                        CustodianSearchHit = NuixCaseReportXML.CreateElement("CustodianSearchHit")
                        CaseGUID.AppendChild(CustodianSearchHit)
                        CustodianSearchHit.InnerText = sSearchHitCount

                        TotalItemCount = NuixCaseReportXML.CreateElement("TotalItmeCount")
                        CaseGUID.AppendChild(TotalItemCount)
                        TotalItemCount.InnerText = sTotalItemCount

                        HitCountPercent = NuixCaseReportXML.CreateElement("HitCountPercent")
                        CaseGUID.AppendChild(HitCountPercent)
                        HitCountPercent.InnerText = sHitCountPercent
                    Case "App Memory per case"
                        CaseGUID = NuixCaseReportXML.CreateElement("CaseGUID")
                        ReportType.AppendChild(CaseGUID)
                        CaseGUID.SetAttribute("CaseGUID", sCaseGUID)

                        CaseName = NuixCaseReportXML.CreateElement("CaseName")
                        ReportType.AppendChild(CaseName)
                        CaseName.InnerText = sCaseName
                        CollectionStatus = NuixCaseReportXML.CreateElement("CollectionStatus")
                        CaseGUID.AppendChild(CollectionStatus)
                        CollectionStatus.InnerText = sCollectionStatus
                        ReportLoadDuration = NuixCaseReportXML.CreateElement("ReportLoadDuration")
                        CaseGUID.AppendChild(ReportLoadDuration)
                        ReportLoadDuration.InnerText = sReportLoadDuration
                        CurrentCaseVersion = NuixCaseReportXML.CreateElement("CurrentCaseVersion")
                        CaseGUID.AppendChild(CurrentCaseVersion)
                        CurrentCaseVersion.InnerText = sCurrentCaseVersion

                        UpgradedCaseVersion = NuixCaseReportXML.CreateElement("UpgradedCaseVersion")
                        CaseGUID.AppendChild(UpgradedCaseVersion)
                        UpgradedCaseVersion.InnerText = sUpgradedCaseVersion

                        BatchLoadInfo = NuixCaseReportXML.CreateElement("BatchLoadInfo")
                        CaseGUID.AppendChild(BatchLoadInfo)
                        BatchLoadInfo.InnerText = sBatchLoadInfo

                        DataExport = NuixCaseReportXML.CreateElement("DataExport")
                        CaseGUID.AppendChild(DataExport)
                        BatchLoadInfo.InnerText = sDataExport


                        IsCompoound = NuixCaseReportXML.CreateElement("IsCompoundCase")
                        CaseGUID.AppendChild(IsCompoound)
                        IsCompoound.InnerText = sIsCompound
                        CasesContained = NuixCaseReportXML.CreateElement("CasesContained")
                        CaseGUID.AppendChild(CasesContained)
                        CasesContained.InnerText = sCompoundCaseContains
                        ContainedInCase = NuixCaseReportXML.CreateElement("ContainedInCase")
                        CaseGUID.AppendChild(ContainedInCase)
                        ContainedInCase.InnerText = sContainedInCase
                        CaseGUID.AppendChild(BatchLoadInfo)
                        BatchLoadInfo.InnerText = sBatchLoadInfo
                        CaseLocation = NuixCaseReportXML.CreateElement("CaseLocation")
                        CaseGUID.AppendChild(CaseLocation)
                        CaseLocation.InnerText = sCaseLocation
                        CaseSizeOnDisk = NuixCaseReportXML.CreateElement("CaseSizeOnDisk")
                        CaseGUID.AppendChild(CaseSizeOnDisk)
                        CaseSizeOnDisk.InnerText = sCaseSizeOnDisk
                        CaseAuditSize = NuixCaseReportXML.CreateElement("CaseAuditSize")
                        CaseGUID.AppendChild(CaseAuditSize)
                        CaseAuditSize.InnerText = sCaseAuditSize

                        OldestItem = NuixCaseReportXML.CreateElement("OldestTopLevelItem")
                        CaseGUID.AppendChild(OldestItem)
                        OldestItem.InnerText = sOldestItem

                        NewestItem = NuixCaseReportXML.CreateElement("NewesTopLevelItem")
                        CaseGUID.AppendChild(NewestItem)
                        NewestItem.InnerText = sNewestItem
                        BrokerMemory = NuixCaseReportXML.CreateElement("BrokerMemory")
                        CaseGUID.AppendChild(BrokerMemory)
                        BrokerMemory.InnerText = sBrokerMemory
                        WorkerCount = NuixCaseReportXML.CreateElement("WorkerCount")
                        CaseGUID.AppendChild(WorkerCount)
                        WorkerCount.InnerText = sWorkerCount
                        WorkerMemory = NuixCaseReportXML.CreateElement("WorkerMemory")
                        CaseGUID.AppendChild(WorkerMemory)
                        WorkerMemory.InnerText = sWorkerMemory
                        ProcessingSpeed = NuixCaseReportXML.CreateElement("ProcessingSpeed")
                        CaseGUID.AppendChild(ProcessingSpeed)
                        ProcessingSpeed.InnerText = sProcessingSpeed
                    Case "Case by Investigator"
                        CaseGUID = NuixCaseReportXML.CreateElement("CaseGUID")
                        ReportType.AppendChild(CaseGUID)
                        CaseGUID.SetAttribute("CaseGUID", sCaseGUID)

                        CaseName = NuixCaseReportXML.CreateElement("CaseName")
                        ReportType.AppendChild(CaseName)
                        CaseName.InnerText = sCaseName
                        CollectionStatus = NuixCaseReportXML.CreateElement("CollectionStatus")
                        CaseGUID.AppendChild(CollectionStatus)
                        CollectionStatus.InnerText = sCollectionStatus
                        ReportLoadDuration = NuixCaseReportXML.CreateElement("ReportLoadDuration")
                        CaseGUID.AppendChild(ReportLoadDuration)
                        ReportLoadDuration.InnerText = sReportLoadDuration
                        CurrentCaseVersion = NuixCaseReportXML.CreateElement("CurrentCaseVersion")
                        CaseGUID.AppendChild(CurrentCaseVersion)
                        CurrentCaseVersion.InnerText = sCurrentCaseVersion

                        UpgradedCaseVersion = NuixCaseReportXML.CreateElement("UpgradedCaseVersion")
                        CaseGUID.AppendChild(UpgradedCaseVersion)
                        UpgradedCaseVersion.InnerText = sUpgradedCaseVersion
                        BatchLoadInfo = NuixCaseReportXML.CreateElement("BatchLoadInfo")
                        CaseGUID.AppendChild(BatchLoadInfo)
                        BatchLoadInfo.InnerText = sBatchLoadInfo

                        DataExport = NuixCaseReportXML.CreateElement("DataExport")
                        CaseGUID.AppendChild(DataExport)
                        BatchLoadInfo.InnerText = sDataExport

                        CaseAuditSize = NuixCaseReportXML.CreateElement("CaseAuditSize")
                        CaseGUID.AppendChild(CaseAuditSize)
                        CaseAuditSize.InnerText = sCaseAuditSize

                        OldestItem = NuixCaseReportXML.CreateElement("OldestTopLevelItem")
                        CaseGUID.AppendChild(OldestItem)
                        OldestItem.InnerText = sOldestItem

                        NewestItem = NuixCaseReportXML.CreateElement("NewesTopLevelItem")
                        CaseGUID.AppendChild(NewestItem)
                        NewestItem.InnerText = sNewestItem
                        IsCompoound = NuixCaseReportXML.CreateElement("IsCompoundCase")
                        CaseGUID.AppendChild(IsCompoound)
                        IsCompoound.InnerText = sIsCompound
                        CasesContained = NuixCaseReportXML.CreateElement("CasesContained")
                        CaseGUID.AppendChild(CasesContained)
                        CasesContained.InnerText = sCompoundCaseContains
                        ContainedInCase = NuixCaseReportXML.CreateElement("ContainedInCase")
                        CaseGUID.AppendChild(ContainedInCase)
                        ContainedInCase.InnerText = sContainedInCase
                        CaseLocation = NuixCaseReportXML.CreateElement("CaseLocation")
                        CaseGUID.AppendChild(CaseLocation)
                        CaseLocation.InnerText = sCaseLocation
                        CaseSizeOnDisk = NuixCaseReportXML.CreateElement("CaseSizeOnDisk")
                        CaseGUID.AppendChild(CaseSizeOnDisk)
                        CaseSizeOnDisk.InnerText = sCaseSizeOnDisk
                        Investigator = NuixCaseReportXML.CreateElement("Investigator")
                        CaseGUID.AppendChild(Investigator)
                        Investigator.InnerText = sInvestigator

                        InvestigatorSessions = NuixCaseReportXML.CreateElement("InvestigatorSessions")
                        CaseGUID.AppendChild(InvestigatorSessions)
                        InvestigatorSessions.InnerText = sInvestigatorSessions

                        InvalidSessions = NuixCaseReportXML.CreateElement("InvalidSessions")
                        CaseGUID.AppendChild(InvalidSessions)
                        InvalidSessions.InnerText = sInvalidSessions

                        InvestigatorTimeSummary = NuixCaseReportXML.CreateElement("InvestigatorTimeSummary")
                        CaseGUID.AppendChild(InvestigatorTimeSummary)
                        InvestigatorTimeSummary.InnerText = sInvestigatorTimeSummary
                    Case "Case Evidence"
                        CaseGUID = NuixCaseReportXML.CreateElement("CaseGUID")
                        ReportType.AppendChild(CaseGUID)
                        CaseGUID.SetAttribute("CaseGUID", sCaseGUID)

                        CaseName = NuixCaseReportXML.CreateElement("CaseName")
                        ReportType.AppendChild(CaseName)
                        CaseName.InnerText = sCaseName
                        CollectionStatus = NuixCaseReportXML.CreateElement("CollectionStatus")
                        CaseGUID.AppendChild(CollectionStatus)
                        CollectionStatus.InnerText = sCollectionStatus
                        ReportLoadDuration = NuixCaseReportXML.CreateElement("ReportLoadDuration")
                        CaseGUID.AppendChild(ReportLoadDuration)
                        ReportLoadDuration.InnerText = sReportLoadDuration
                        CurrentCaseVersion = NuixCaseReportXML.CreateElement("CurrentCaseVersion")
                        CaseGUID.AppendChild(CurrentCaseVersion)
                        CurrentCaseVersion.InnerText = sCurrentCaseVersion

                        UpgradedCaseVersion = NuixCaseReportXML.CreateElement("UpgradedCaseVersion")
                        CaseGUID.AppendChild(UpgradedCaseVersion)
                        UpgradedCaseVersion.InnerText = sUpgradedCaseVersion
                        BatchLoadInfo = NuixCaseReportXML.CreateElement("BatchLoadInfo")
                        CaseGUID.AppendChild(BatchLoadInfo)
                        BatchLoadInfo.InnerText = sBatchLoadInfo

                        DataExport = NuixCaseReportXML.CreateElement("DataExport")
                        CaseGUID.AppendChild(DataExport)
                        BatchLoadInfo.InnerText = sDataExport

                        CaseAuditSize = NuixCaseReportXML.CreateElement("CaseAuditSize")
                        CaseGUID.AppendChild(CaseAuditSize)
                        CaseAuditSize.InnerText = sCaseAuditSize

                        OldestItem = NuixCaseReportXML.CreateElement("OldestTopLevelItem")
                        CaseGUID.AppendChild(OldestItem)
                        OldestItem.InnerText = sOldestItem

                        NewestItem = NuixCaseReportXML.CreateElement("NewesTopLevelItem")
                        CaseGUID.AppendChild(NewestItem)
                        NewestItem.InnerText = sNewestItem
                        IsCompoound = NuixCaseReportXML.CreateElement("IsCompoundCase")
                        CaseGUID.AppendChild(IsCompoound)
                        IsCompoound.InnerText = sIsCompound
                        CasesContained = NuixCaseReportXML.CreateElement("CasesContained")
                        CaseGUID.AppendChild(CasesContained)
                        CasesContained.InnerText = sCompoundCaseContains
                        ContainedInCase = NuixCaseReportXML.CreateElement("ContainedInCase")
                        CaseGUID.AppendChild(ContainedInCase)
                        ContainedInCase.InnerText = sContainedInCase
                        CaseLocation = NuixCaseReportXML.CreateElement("CaseLocation")
                        CaseGUID.AppendChild(CaseLocation)
                        CaseLocation.InnerText = sCaseLocation
                        CaseSizeOnDisk = NuixCaseReportXML.CreateElement("CaseSizeOnDisk")
                        CaseGUID.AppendChild(CaseSizeOnDisk)
                        CaseSizeOnDisk.InnerText = sCaseSizeOnDisk
                        EvidenceName = NuixCaseReportXML.CreateElement("EvidenceName")
                        CaseGUID.AppendChild(EvidenceName)
                        EvidenceName.InnerText = sEvidenceName
                        EvidenceLocation = NuixCaseReportXML.CreateElement("EvidenceLocation")
                        CaseGUID.AppendChild(EvidenceLocation)
                        EvidenceLocation.InnerText = sEvidenceLocation
                    Case "Case Location"
                        CaseGUID = NuixCaseReportXML.CreateElement("CaseGUID")
                        ReportType.AppendChild(CaseGUID)
                        CaseGUID.SetAttribute("CaseGUID", sCaseGUID)

                        CaseName = NuixCaseReportXML.CreateElement("CaseName")
                        ReportType.AppendChild(CaseName)
                        CaseName.InnerText = sCaseName
                        CollectionStatus = NuixCaseReportXML.CreateElement("CollectionStatus")
                        CaseGUID.AppendChild(CollectionStatus)
                        CollectionStatus.InnerText = sCollectionStatus
                        ReportLoadDuration = NuixCaseReportXML.CreateElement("ReportLoadDuration")
                        CaseGUID.AppendChild(ReportLoadDuration)
                        ReportLoadDuration.InnerText = sReportLoadDuration
                        CurrentCaseVersion = NuixCaseReportXML.CreateElement("CurrentCaseVersion")
                        CaseGUID.AppendChild(CurrentCaseVersion)
                        CurrentCaseVersion.InnerText = sCurrentCaseVersion

                        UpgradedCaseVersion = NuixCaseReportXML.CreateElement("UpgradedCaseVersion")
                        CaseGUID.AppendChild(UpgradedCaseVersion)
                        UpgradedCaseVersion.InnerText = sUpgradedCaseVersion
                        BatchLoadInfo = NuixCaseReportXML.CreateElement("BatchLoadInfo")
                        CaseGUID.AppendChild(BatchLoadInfo)
                        BatchLoadInfo.InnerText = sBatchLoadInfo

                        DataExport = NuixCaseReportXML.CreateElement("DataExport")
                        CaseGUID.AppendChild(DataExport)
                        BatchLoadInfo.InnerText = sDataExport

                        CaseAuditSize = NuixCaseReportXML.CreateElement("CaseAuditSize")
                        CaseGUID.AppendChild(CaseAuditSize)
                        CaseAuditSize.InnerText = sCaseAuditSize

                        OldestItem = NuixCaseReportXML.CreateElement("OldestTopLevelItem")
                        CaseGUID.AppendChild(OldestItem)
                        OldestItem.InnerText = sOldestItem

                        NewestItem = NuixCaseReportXML.CreateElement("NewesTopLevelItem")
                        CaseGUID.AppendChild(NewestItem)
                        NewestItem.InnerText = sNewestItem
                        IsCompoound = NuixCaseReportXML.CreateElement("IsCompoundCase")
                        CaseGUID.AppendChild(IsCompoound)
                        IsCompoound.InnerText = sIsCompound
                        CasesContained = NuixCaseReportXML.CreateElement("CasesContained")
                        CaseGUID.AppendChild(CasesContained)
                        CasesContained.InnerText = sCompoundCaseContains
                        CaseGUID.AppendChild(BatchLoadInfo)
                        ContainedInCase = NuixCaseReportXML.CreateElement("ContainedInCase")
                        CaseGUID.AppendChild(ContainedInCase)
                        ContainedInCase.InnerText = sContainedInCase
                        CaseLocation = NuixCaseReportXML.CreateElement("CaseLocation")
                        CaseGUID.AppendChild(CaseLocation)
                        CaseLocation.InnerText = sCaseLocation
                        CaseSizeOnDisk = NuixCaseReportXML.CreateElement("CaseSizeOnDisk")
                        CaseGUID.AppendChild(CaseSizeOnDisk)
                        CaseSizeOnDisk.InnerText = sCaseSizeOnDisk
                    Case "Case Size"
                        CaseGUID = NuixCaseReportXML.CreateElement("CaseGUID")
                        ReportType.AppendChild(CaseGUID)
                        CaseGUID.SetAttribute("CaseGUID", sCaseGUID)

                        CaseName = NuixCaseReportXML.CreateElement("CaseName")
                        ReportType.AppendChild(CaseName)
                        CaseName.InnerText = sCaseName
                        CollectionStatus = NuixCaseReportXML.CreateElement("CollectionStatus")
                        CaseGUID.AppendChild(CollectionStatus)
                        CollectionStatus.InnerText = sCollectionStatus
                        ReportLoadDuration = NuixCaseReportXML.CreateElement("ReportLoadDuration")
                        CaseGUID.AppendChild(ReportLoadDuration)
                        ReportLoadDuration.InnerText = sReportLoadDuration
                        CurrentCaseVersion = NuixCaseReportXML.CreateElement("CurrentCaseVersion")
                        CaseGUID.AppendChild(CurrentCaseVersion)
                        CurrentCaseVersion.InnerText = sCurrentCaseVersion

                        UpgradedCaseVersion = NuixCaseReportXML.CreateElement("UpgradedCaseVersion")
                        CaseGUID.AppendChild(UpgradedCaseVersion)
                        UpgradedCaseVersion.InnerText = sUpgradedCaseVersion
                        BatchLoadInfo = NuixCaseReportXML.CreateElement("BatchLoadInfo")
                        CaseGUID.AppendChild(BatchLoadInfo)
                        BatchLoadInfo.InnerText = sBatchLoadInfo

                        DataExport = NuixCaseReportXML.CreateElement("DataExport")
                        CaseGUID.AppendChild(DataExport)
                        BatchLoadInfo.InnerText = sDataExport

                        CaseAuditSize = NuixCaseReportXML.CreateElement("CaseAuditSize")
                        CaseGUID.AppendChild(CaseAuditSize)
                        CaseAuditSize.InnerText = sCaseAuditSize

                        OldestItem = NuixCaseReportXML.CreateElement("OldestTopLevelItem")
                        CaseGUID.AppendChild(OldestItem)
                        OldestItem.InnerText = sOldestItem

                        NewestItem = NuixCaseReportXML.CreateElement("NewesTopLevelItem")
                        CaseGUID.AppendChild(NewestItem)
                        NewestItem.InnerText = sNewestItem
                        IsCompoound = NuixCaseReportXML.CreateElement("IsCompoundCase")
                        CaseGUID.AppendChild(IsCompoound)
                        IsCompoound.InnerText = sIsCompound
                        CasesContained = NuixCaseReportXML.CreateElement("CasesContained")
                        CaseGUID.AppendChild(CasesContained)
                        CasesContained.InnerText = sCompoundCaseContains
                        ContainedInCase = NuixCaseReportXML.CreateElement("ContainedInCase")
                        CaseGUID.AppendChild(ContainedInCase)
                        ContainedInCase.InnerText = sContainedInCase
                        CaseLocation = NuixCaseReportXML.CreateElement("CaseLocation")
                        CaseGUID.AppendChild(CaseLocation)
                        CaseLocation.InnerText = sCaseLocation
                        CaseSizeOnDisk = NuixCaseReportXML.CreateElement("CaseSizeOnDisk")
                        CaseGUID.AppendChild(CaseSizeOnDisk)
                        CaseSizeOnDisk.InnerText = sCaseSizeOnDisk
                        LoadEvents = NuixCaseReportXML.CreateElement("LoadEvents")
                        CaseGUID.AppendChild(LoadEvents)
                        LoadEvents.InnerText = sLoadEvents
                        TotalLoadTime = NuixCaseReportXML.CreateElement("TotalLoadTime")
                        CaseGUID.AppendChild(TotalLoadTime)
                        TotalLoadTime.InnerText = sTotalLoadTime
                        ProcessingSpeed = NuixCaseReportXML.CreateElement("ProcessingSpeed")
                        CaseGUID.AppendChild(ProcessingSpeed)
                        ProcessingSpeed.InnerText = sProcessingSpeed
                    Case "Custodians in Case"
                        CaseGUID = NuixCaseReportXML.CreateElement("CaseGUID")
                        ReportType.AppendChild(CaseGUID)
                        CaseGUID.SetAttribute("CaseGUID", sCaseGUID)

                        CaseName = NuixCaseReportXML.CreateElement("CaseName")
                        ReportType.AppendChild(CaseName)
                        CaseName.InnerText = sCaseName
                        CollectionStatus = NuixCaseReportXML.CreateElement("CollectionStatus")
                        CaseGUID.AppendChild(CollectionStatus)
                        CollectionStatus.InnerText = sCollectionStatus
                        ReportLoadDuration = NuixCaseReportXML.CreateElement("ReportLoadDuration")
                        CaseGUID.AppendChild(ReportLoadDuration)
                        ReportLoadDuration.InnerText = sReportLoadDuration
                        CurrentCaseVersion = NuixCaseReportXML.CreateElement("CurrentCaseVersion")
                        CaseGUID.AppendChild(CurrentCaseVersion)
                        CurrentCaseVersion.InnerText = sCurrentCaseVersion

                        UpgradedCaseVersion = NuixCaseReportXML.CreateElement("UpgradedCaseVersion")
                        CaseGUID.AppendChild(UpgradedCaseVersion)
                        UpgradedCaseVersion.InnerText = sUpgradedCaseVersion

                        BatchLoadInfo = NuixCaseReportXML.CreateElement("BatchLoadInfo")
                        CaseGUID.AppendChild(BatchLoadInfo)
                        BatchLoadInfo.InnerText = sBatchLoadInfo

                        DataExport = NuixCaseReportXML.CreateElement("DataExport")
                        CaseGUID.AppendChild(DataExport)
                        BatchLoadInfo.InnerText = sDataExport

                        CaseAuditSize = NuixCaseReportXML.CreateElement("CaseAuditSize")
                        CaseGUID.AppendChild(CaseAuditSize)
                        CaseAuditSize.InnerText = sCaseAuditSize

                        OldestItem = NuixCaseReportXML.CreateElement("OldestTopLevelItem")
                        CaseGUID.AppendChild(OldestItem)
                        OldestItem.InnerText = sOldestItem

                        NewestItem = NuixCaseReportXML.CreateElement("NewesTopLevelItem")
                        CaseGUID.AppendChild(NewestItem)
                        NewestItem.InnerText = sNewestItem
                        IsCompoound = NuixCaseReportXML.CreateElement("IsCompoundCase")
                        CaseGUID.AppendChild(IsCompoound)
                        IsCompoound.InnerText = sIsCompound
                        CasesContained = NuixCaseReportXML.CreateElement("CasesContained")
                        CaseGUID.AppendChild(CasesContained)
                        CasesContained.InnerText = sCompoundCaseContains
                        ContainedInCase = NuixCaseReportXML.CreateElement("ContainedInCase")
                        CaseGUID.AppendChild(ContainedInCase)
                        ContainedInCase.InnerText = sContainedInCase
                        CaseLocation = NuixCaseReportXML.CreateElement("CaseLocation")
                        CaseGUID.AppendChild(CaseLocation)
                        CaseLocation.InnerText = sCaseLocation
                        CaseSizeOnDisk = NuixCaseReportXML.CreateElement("CaseSizeOnDisk")
                        CaseGUID.AppendChild(CaseSizeOnDisk)
                        CaseSizeOnDisk.InnerText = sCaseSizeOnDisk
                        Investigator = NuixCaseReportXML.CreateElement("Investigator")
                        CaseGUID.AppendChild(Investigator)
                        Investigator.InnerText = sInvestigator

                        InvestigatorSessions = NuixCaseReportXML.CreateElement("InvestigatorSessions")
                        CaseGUID.AppendChild(InvestigatorSessions)
                        InvestigatorSessions.InnerText = sInvestigatorSessions

                        InvalidSessions = NuixCaseReportXML.CreateElement("InvalidSessions")
                        CaseGUID.AppendChild(InvalidSessions)
                        InvalidSessions.InnerText = sInvalidSessions

                        InvestigatorTimeSummary = NuixCaseReportXML.CreateElement("InvestigatorTimeSummary")
                        CaseGUID.AppendChild(InvestigatorTimeSummary)
                        InvestigatorTimeSummary.InnerText = sInvestigatorTimeSummary
                        Custodians = NuixCaseReportXML.CreateElement("Custodians")
                        CaseGUID.AppendChild(Custodians)
                        Custodians.InnerText = sCustodians
                        CustodianCount = NuixCaseReportXML.CreateElement("CustodianCount")
                        CaseGUID.AppendChild(CustodianCount)
                        CustodianCount.InnerText = sCustodianCount
                    Case "Metadata Type"
                        CaseGUID = NuixCaseReportXML.CreateElement("CaseGUID")
                        ReportType.AppendChild(CaseGUID)
                        CaseGUID.SetAttribute("CaseGUID", sCaseGUID)

                        CaseName = NuixCaseReportXML.CreateElement("CaseName")
                        ReportType.AppendChild(CaseName)
                        CaseName.InnerText = sCaseName
                        CollectionStatus = NuixCaseReportXML.CreateElement("CollectionStatus")
                        CaseGUID.AppendChild(CollectionStatus)
                        CollectionStatus.InnerText = sCollectionStatus
                        ReportLoadDuration = NuixCaseReportXML.CreateElement("ReportLoadDuration")
                        CaseGUID.AppendChild(ReportLoadDuration)
                        ReportLoadDuration.InnerText = sReportLoadDuration
                        CurrentCaseVersion = NuixCaseReportXML.CreateElement("CurrentCaseVersion")
                        CaseGUID.AppendChild(CurrentCaseVersion)
                        CurrentCaseVersion.InnerText = sCurrentCaseVersion

                        UpgradedCaseVersion = NuixCaseReportXML.CreateElement("UpgradedCaseVersion")
                        CaseGUID.AppendChild(UpgradedCaseVersion)
                        UpgradedCaseVersion.InnerText = sUpgradedCaseVersion
                        BatchLoadInfo = NuixCaseReportXML.CreateElement("BatchLoadInfo")
                        CaseGUID.AppendChild(BatchLoadInfo)
                        BatchLoadInfo.InnerText = sBatchLoadInfo

                        DataExport = NuixCaseReportXML.CreateElement("DataExport")
                        CaseGUID.AppendChild(DataExport)
                        BatchLoadInfo.InnerText = sDataExport

                        CaseAuditSize = NuixCaseReportXML.CreateElement("CaseAuditSize")
                        CaseGUID.AppendChild(CaseAuditSize)
                        CaseAuditSize.InnerText = sCaseAuditSize

                        OldestItem = NuixCaseReportXML.CreateElement("OldestTopLevelItem")
                        CaseGUID.AppendChild(OldestItem)
                        OldestItem.InnerText = sOldestItem

                        NewestItem = NuixCaseReportXML.CreateElement("NewesTopLevelItem")
                        CaseGUID.AppendChild(NewestItem)
                        NewestItem.InnerText = sNewestItem
                        IsCompoound = NuixCaseReportXML.CreateElement("IsCompoundCase")
                        CaseGUID.AppendChild(IsCompoound)
                        IsCompoound.InnerText = sIsCompound
                        CasesContained = NuixCaseReportXML.CreateElement("CasesContained")
                        CaseGUID.AppendChild(CasesContained)
                        CasesContained.InnerText = sCompoundCaseContains
                        ContainedInCase = NuixCaseReportXML.CreateElement("ContainedInCase")
                        CaseGUID.AppendChild(ContainedInCase)
                        ContainedInCase.InnerText = sContainedInCase
                        CaseLocation = NuixCaseReportXML.CreateElement("CaseLocation")
                        CaseGUID.AppendChild(CaseLocation)
                        CaseLocation.InnerText = sCaseLocation
                        CaseSizeOnDisk = NuixCaseReportXML.CreateElement("CaseSizeOnDisk")
                        CaseGUID.AppendChild(CaseSizeOnDisk)
                        CaseSizeOnDisk.InnerText = sCaseSizeOnDisk
                        CaseFileSize = NuixCaseReportXML.CreateElement("CaseFileSize")
                        CaseGUID.AppendChild(CaseFileSize)
                        CaseFileSize.InnerText = sCaseFileSize
                        MimeTypes = NuixCaseReportXML.CreateElement("MimeTypes")
                        CaseGUID.AppendChild(MimeTypes)
                        MimeTypes.InnerText = sMimeTypes
                        ItemTypes = NuixCaseReportXML.CreateElement("ItemTypes")
                        CaseGUID.AppendChild(ItemTypes)
                        ItemTypes.InnerText = sItemTypes
                    Case "Processing time"
                        CaseGUID = NuixCaseReportXML.CreateElement("CaseGUID")
                        ReportType.AppendChild(CaseGUID)
                        CaseGUID.SetAttribute("CaseGUID", sCaseGUID)

                        CaseName = NuixCaseReportXML.CreateElement("CaseName")
                        ReportType.AppendChild(CaseName)
                        CaseName.InnerText = sCaseName
                        CollectionStatus = NuixCaseReportXML.CreateElement("CollectionStatus")
                        CaseGUID.AppendChild(CollectionStatus)
                        CollectionStatus.InnerText = sCollectionStatus
                        ReportLoadDuration = NuixCaseReportXML.CreateElement("ReportLoadDuration")
                        CaseGUID.AppendChild(ReportLoadDuration)
                        ReportLoadDuration.InnerText = sReportLoadDuration
                        CurrentCaseVersion = NuixCaseReportXML.CreateElement("CurrentCaseVersion")
                        CaseGUID.AppendChild(CurrentCaseVersion)
                        CurrentCaseVersion.InnerText = sCurrentCaseVersion

                        UpgradedCaseVersion = NuixCaseReportXML.CreateElement("UpgradedCaseVersion")
                        CaseGUID.AppendChild(UpgradedCaseVersion)
                        UpgradedCaseVersion.InnerText = sUpgradedCaseVersion
                        BatchLoadInfo = NuixCaseReportXML.CreateElement("BatchLoadInfo")
                        CaseGUID.AppendChild(BatchLoadInfo)
                        BatchLoadInfo.InnerText = sBatchLoadInfo

                        DataExport = NuixCaseReportXML.CreateElement("DataExport")
                        CaseGUID.AppendChild(DataExport)
                        BatchLoadInfo.InnerText = sDataExport

                        CaseAuditSize = NuixCaseReportXML.CreateElement("CaseAuditSize")
                        CaseGUID.AppendChild(CaseAuditSize)
                        CaseAuditSize.InnerText = sCaseAuditSize

                        OldestItem = NuixCaseReportXML.CreateElement("OldestTopLevelItem")
                        CaseGUID.AppendChild(OldestItem)
                        OldestItem.InnerText = sOldestItem

                        NewestItem = NuixCaseReportXML.CreateElement("NewesTopLevelItem")
                        CaseGUID.AppendChild(NewestItem)
                        NewestItem.InnerText = sNewestItem
                        IsCompoound = NuixCaseReportXML.CreateElement("IsCompoundCase")
                        CaseGUID.AppendChild(IsCompoound)
                        IsCompoound.InnerText = sIsCompound
                        CasesContained = NuixCaseReportXML.CreateElement("CasesContained")
                        CaseGUID.AppendChild(CasesContained)
                        CasesContained.InnerText = sCompoundCaseContains
                        ContainedInCase = NuixCaseReportXML.CreateElement("ContainedInCase")
                        CaseGUID.AppendChild(ContainedInCase)
                        ContainedInCase.InnerText = sContainedInCase
                        CaseLocation = NuixCaseReportXML.CreateElement("CaseLocation")
                        CaseGUID.AppendChild(CaseLocation)
                        CaseLocation.InnerText = sCaseLocation
                        CaseSizeOnDisk = NuixCaseReportXML.CreateElement("CaseSizeOnDisk")
                        CaseGUID.AppendChild(CaseSizeOnDisk)
                        CaseSizeOnDisk.InnerText = sCaseSizeOnDisk
                        CaseFileSize = NuixCaseReportXML.CreateElement("CaseFileSize")
                        CaseGUID.AppendChild(CaseFileSize)
                        CaseFileSize.InnerText = sCaseFileSize
                        LoadEvents = NuixCaseReportXML.CreateElement("LoadEvents")
                        CaseGUID.AppendChild(LoadEvents)
                        LoadEvents.InnerText = sLoadEvents
                        TotalLoadTime = NuixCaseReportXML.CreateElement("TotalLoadTime")
                        CaseGUID.AppendChild(TotalLoadTime)
                        TotalLoadTime.InnerText = sTotalLoadTime
                        ProcessingSpeed = NuixCaseReportXML.CreateElement("ProcessingSpeed")
                        CaseGUID.AppendChild(ProcessingSpeed)
                        ProcessingSpeed.InnerText = sProcessingSpeed
                    Case "Processing speed(GB / Hour)"
                        CaseGUID = NuixCaseReportXML.CreateElement("CaseGUID")
                        ReportType.AppendChild(CaseGUID)
                        CaseGUID.SetAttribute("CaseGUID", sCaseGUID)

                        CaseName = NuixCaseReportXML.CreateElement("CaseName")
                        ReportType.AppendChild(CaseName)
                        CaseName.InnerText = sCaseName
                        CollectionStatus = NuixCaseReportXML.CreateElement("CollectionStatus")
                        CaseGUID.AppendChild(CollectionStatus)
                        CollectionStatus.InnerText = sCollectionStatus
                        ReportLoadDuration = NuixCaseReportXML.CreateElement("ReportLoadDuration")
                        CaseGUID.AppendChild(ReportLoadDuration)
                        ReportLoadDuration.InnerText = sReportLoadDuration
                        CurrentCaseVersion = NuixCaseReportXML.CreateElement("CurrentCaseVersion")
                        CaseGUID.AppendChild(CurrentCaseVersion)
                        CurrentCaseVersion.InnerText = sCurrentCaseVersion

                        UpgradedCaseVersion = NuixCaseReportXML.CreateElement("UpgradedCaseVersion")
                        CaseGUID.AppendChild(UpgradedCaseVersion)
                        UpgradedCaseVersion.InnerText = sUpgradedCaseVersion

                        BatchLoadInfo = NuixCaseReportXML.CreateElement("BatchLoadInfo")
                        CaseGUID.AppendChild(BatchLoadInfo)
                        BatchLoadInfo.InnerText = sBatchLoadInfo

                        DataExport = NuixCaseReportXML.CreateElement("DataExport")
                        CaseGUID.AppendChild(DataExport)
                        BatchLoadInfo.InnerText = sDataExport

                        IsCompoound = NuixCaseReportXML.CreateElement("IsCompoundCase")
                        CaseGUID.AppendChild(IsCompoound)
                        IsCompoound.InnerText = sIsCompound
                        CasesContained = NuixCaseReportXML.CreateElement("CasesContained")
                        CaseGUID.AppendChild(CasesContained)
                        CasesContained.InnerText = sCompoundCaseContains
                        CaseGUID.AppendChild(BatchLoadInfo)
                        BatchLoadInfo.InnerText = sBatchLoadInfo
                        CaseLocation = NuixCaseReportXML.CreateElement("CaseLocation")
                        CaseGUID.AppendChild(CaseLocation)
                        ContainedInCase = NuixCaseReportXML.CreateElement("ContainedInCase")
                        CaseGUID.AppendChild(ContainedInCase)
                        ContainedInCase.InnerText = sContainedInCase
                        CaseLocation.InnerText = sCaseLocation
                        CaseSizeOnDisk = NuixCaseReportXML.CreateElement("CaseSizeOnDisk")
                        CaseGUID.AppendChild(CaseSizeOnDisk)
                        CaseSizeOnDisk.InnerText = sCaseSizeOnDisk
                        CaseFileSize = NuixCaseReportXML.CreateElement("CaseFileSize")
                        CaseGUID.AppendChild(CaseFileSize)
                        CaseFileSize.InnerText = sCaseFileSize
                        CaseAuditSize = NuixCaseReportXML.CreateElement("CaseAuditSize")
                        CaseGUID.AppendChild(CaseAuditSize)
                        CaseAuditSize.InnerText = sCaseAuditSize

                        OldestItem = NuixCaseReportXML.CreateElement("OldestTopLevelItem")
                        CaseGUID.AppendChild(OldestItem)
                        OldestItem.InnerText = sOldestItem

                        NewestItem = NuixCaseReportXML.CreateElement("NewesTopLevelItem")
                        CaseGUID.AppendChild(NewestItem)
                        NewestItem.InnerText = sNewestItem
                        LoadEvents = NuixCaseReportXML.CreateElement("LoadEvents")
                        CaseGUID.AppendChild(LoadEvents)
                        LoadEvents.InnerText = sLoadEvents
                        TotalLoadTime = NuixCaseReportXML.CreateElement("TotalLoadTime")
                        CaseGUID.AppendChild(TotalLoadTime)
                        TotalLoadTime.InnerText = sTotalLoadTime
                        ProcessingSpeed = NuixCaseReportXML.CreateElement("ProcessingSpeed")
                        CaseGUID.AppendChild(ProcessingSpeed)
                        ProcessingSpeed.InnerText = sProcessingSpeed
                    Case "Search Term Hit"
                        CaseGUID = NuixCaseReportXML.CreateElement("CaseGUID")
                        ReportType.AppendChild(CaseGUID)
                        CaseGUID.SetAttribute("CaseGUID", sCaseGUID)

                        CaseName = NuixCaseReportXML.CreateElement("CaseName")
                        ReportType.AppendChild(CaseName)
                        CaseName.InnerText = sCaseName
                        CollectionStatus = NuixCaseReportXML.CreateElement("CollectionStatus")
                        CaseGUID.AppendChild(CollectionStatus)
                        CollectionStatus.InnerText = sCollectionStatus
                        ReportLoadDuration = NuixCaseReportXML.CreateElement("ReportLoadDuration")
                        CaseGUID.AppendChild(ReportLoadDuration)
                        ReportLoadDuration.InnerText = sReportLoadDuration
                        CurrentCaseVersion = NuixCaseReportXML.CreateElement("CurrentCaseVersion")
                        CaseGUID.AppendChild(CurrentCaseVersion)
                        CurrentCaseVersion.InnerText = sCurrentCaseVersion

                        UpgradedCaseVersion = NuixCaseReportXML.CreateElement("UpgradedCaseVersion")
                        CaseGUID.AppendChild(UpgradedCaseVersion)
                        UpgradedCaseVersion.InnerText = sUpgradedCaseVersion

                        BatchLoadInfo = NuixCaseReportXML.CreateElement("BatchLoadInfo")
                        CaseGUID.AppendChild(BatchLoadInfo)
                        BatchLoadInfo.InnerText = sBatchLoadInfo

                        DataExport = NuixCaseReportXML.CreateElement("DataExport")
                        CaseGUID.AppendChild(DataExport)
                        BatchLoadInfo.InnerText = sDataExport

                        CaseAuditSize = NuixCaseReportXML.CreateElement("CaseAuditSize")
                        CaseGUID.AppendChild(CaseAuditSize)
                        CaseAuditSize.InnerText = sCaseAuditSize

                        OldestItem = NuixCaseReportXML.CreateElement("OldestTopLevelItem")
                        CaseGUID.AppendChild(OldestItem)
                        OldestItem.InnerText = sOldestItem

                        NewestItem = NuixCaseReportXML.CreateElement("NewesTopLevelItem")
                        CaseGUID.AppendChild(NewestItem)
                        NewestItem.InnerText = sNewestItem
                        IsCompoound = NuixCaseReportXML.CreateElement("IsCompoundCase")
                        CaseGUID.AppendChild(IsCompoound)
                        IsCompoound.InnerText = sIsCompound
                        CasesContained = NuixCaseReportXML.CreateElement("CasesContained")
                        CaseGUID.AppendChild(CasesContained)
                        CasesContained.InnerText = sCompoundCaseContains
                        ContainedInCase = NuixCaseReportXML.CreateElement("ContainedInCase")
                        CaseGUID.AppendChild(ContainedInCase)
                        ContainedInCase.InnerText = sContainedInCase
                        CaseLocation = NuixCaseReportXML.CreateElement("CaseLocation")
                        CaseGUID.AppendChild(CaseLocation)
                        CaseLocation.InnerText = sCaseLocation
                        CaseSizeOnDisk = NuixCaseReportXML.CreateElement("CaseSizeOnDisk")
                        CaseGUID.AppendChild(CaseSizeOnDisk)
                        CaseSizeOnDisk.InnerText = sCaseSizeOnDisk
                        CaseFileSize = NuixCaseReportXML.CreateElement("CaseFileSize")
                        CaseGUID.AppendChild(CaseFileSize)
                        CaseFileSize.InnerText = sCaseFileSize
                        SearchTerm = NuixCaseReportXML.CreateElement("SearchTerm")
                        CaseGUID.AppendChild(SearchTerm)
                        SearchTerm.InnerText = sSearchTerm
                        SearchSize = NuixCaseReportXML.CreateElement("SearchSize")
                        CaseGUID.AppendChild(SearchSize)
                        SearchSize.InnerText = sSearchSize
                        SearchHitCount = NuixCaseReportXML.CreateElement("SearchHitCount")
                        CaseGUID.AppendChild(SearchHitCount)
                        SearchHitCount.InnerText = sSearchHitCount
                        CustodianSearchHit = NuixCaseReportXML.CreateElement("CustodianSearchHit")
                        CaseGUID.AppendChild(CustodianSearchHit)
                        CustodianSearchHit.InnerText = sSearchHitCount
                        TotalItemCount = NuixCaseReportXML.CreateElement("TotalItemCount")
                        CaseGUID.AppendChild(TotalItemCount)
                        TotalItemCount.InnerText = sTotalItemCount
                        HitCountPercent = NuixCaseReportXML.CreateElement("sHitCountPercent")
                        CaseGUID.AppendChild(HitCountPercent)
                        HitCountPercent.InnerText = sHitCountPercent
                    Case "Total Number of Items"
                        CaseGUID = NuixCaseReportXML.CreateElement("CaseGUID")
                        ReportType.AppendChild(CaseGUID)
                        CaseGUID.SetAttribute("CaseGUID", sCaseGUID)

                        CaseName = NuixCaseReportXML.CreateElement("CaseName")
                        ReportType.AppendChild(CaseName)
                        CaseName.InnerText = sCaseName
                        CollectionStatus = NuixCaseReportXML.CreateElement("CollectionStatus")
                        CaseGUID.AppendChild(CollectionStatus)
                        CollectionStatus.InnerText = sCollectionStatus
                        ReportLoadDuration = NuixCaseReportXML.CreateElement("ReportLoadDuration")
                        CaseGUID.AppendChild(ReportLoadDuration)
                        ReportLoadDuration.InnerText = sReportLoadDuration

                        CurrentCaseVersion = NuixCaseReportXML.CreateElement("CurrentCaseVersion")
                        CaseGUID.AppendChild(CurrentCaseVersion)
                        CurrentCaseVersion.InnerText = sCurrentCaseVersion

                        UpgradedCaseVersion = NuixCaseReportXML.CreateElement("UpgradedCaseVersion")
                        CaseGUID.AppendChild(UpgradedCaseVersion)
                        UpgradedCaseVersion.InnerText = sUpgradedCaseVersion

                        BatchLoadInfo = NuixCaseReportXML.CreateElement("BatchLoadInfo")
                        CaseGUID.AppendChild(BatchLoadInfo)
                        BatchLoadInfo.InnerText = sBatchLoadInfo

                        DataExport = NuixCaseReportXML.CreateElement("DataExport")
                        CaseGUID.AppendChild(DataExport)
                        BatchLoadInfo.InnerText = sDataExport

                        CaseAuditSize = NuixCaseReportXML.CreateElement("CaseAuditSize")
                        CaseGUID.AppendChild(CaseAuditSize)
                        CaseAuditSize.InnerText = sCaseAuditSize

                        OldestItem = NuixCaseReportXML.CreateElement("OldestTopLevelItem")
                        CaseGUID.AppendChild(OldestItem)
                        OldestItem.InnerText = sOldestItem

                        NewestItem = NuixCaseReportXML.CreateElement("NewesTopLevelItem")
                        CaseGUID.AppendChild(NewestItem)
                        NewestItem.InnerText = sNewestItem
                        IsCompoound = NuixCaseReportXML.CreateElement("IsCompoundCase")
                        CaseGUID.AppendChild(IsCompoound)
                        IsCompoound.InnerText = sIsCompound
                        CasesContained = NuixCaseReportXML.CreateElement("CasesContained")
                        CaseGUID.AppendChild(CasesContained)
                        CasesContained.InnerText = sCompoundCaseContains
                        ContainedInCase = NuixCaseReportXML.CreateElement("ContainedInCase")
                        CaseGUID.AppendChild(ContainedInCase)
                        ContainedInCase.InnerText = sContainedInCase
                        CaseLocation = NuixCaseReportXML.CreateElement("CaseLocation")
                        CaseGUID.AppendChild(CaseLocation)
                        CaseLocation.InnerText = sCaseLocation
                        CaseSizeOnDisk = NuixCaseReportXML.CreateElement("CaseSizeOnDisk")
                        CaseGUID.AppendChild(CaseSizeOnDisk)
                        CaseSizeOnDisk.InnerText = sCaseSizeOnDisk
                        TotalItemCount = NuixCaseReportXML.CreateElement("TotalItemCount")
                        CaseGUID.AppendChild(TotalItemCount)
                        TotalItemCount.InnerText = sTotalItemCount
                    Case "Total Number of workers"
                        CaseGUID = NuixCaseReportXML.CreateElement("CaseGUID")
                        ReportType.AppendChild(CaseGUID)
                        CaseGUID.SetAttribute("CaseGUID", sCaseGUID)

                        CaseName = NuixCaseReportXML.CreateElement("CaseName")
                        ReportType.AppendChild(CaseName)
                        CaseName.InnerText = sCaseName
                        CollectionStatus = NuixCaseReportXML.CreateElement("CollectionStatus")
                        CaseGUID.AppendChild(CollectionStatus)
                        CollectionStatus.InnerText = sCollectionStatus
                        ReportLoadDuration = NuixCaseReportXML.CreateElement("ReportLoadDuration")
                        CaseGUID.AppendChild(ReportLoadDuration)
                        ReportLoadDuration.InnerText = sReportLoadDuration
                        CurrentCaseVersion = NuixCaseReportXML.CreateElement("CurrentCaseVersion")
                        CaseGUID.AppendChild(CurrentCaseVersion)
                        CurrentCaseVersion.InnerText = sCurrentCaseVersion

                        UpgradedCaseVersion = NuixCaseReportXML.CreateElement("UpgradedCaseVersion")
                        CaseGUID.AppendChild(UpgradedCaseVersion)
                        UpgradedCaseVersion.InnerText = sUpgradedCaseVersion

                        BatchLoadInfo = NuixCaseReportXML.CreateElement("BatchLoadInfo")
                        CaseGUID.AppendChild(BatchLoadInfo)
                        BatchLoadInfo.InnerText = sBatchLoadInfo

                        DataExport = NuixCaseReportXML.CreateElement("DataExport")
                        CaseGUID.AppendChild(DataExport)
                        BatchLoadInfo.InnerText = sDataExport

                        CaseAuditSize = NuixCaseReportXML.CreateElement("CaseAuditSize")
                        CaseGUID.AppendChild(CaseAuditSize)
                        CaseAuditSize.InnerText = sCaseAuditSize

                        OldestItem = NuixCaseReportXML.CreateElement("OldestTopLevelItem")
                        CaseGUID.AppendChild(OldestItem)
                        OldestItem.InnerText = sOldestItem

                        NewestItem = NuixCaseReportXML.CreateElement("NewesTopLevelItem")
                        CaseGUID.AppendChild(NewestItem)
                        NewestItem.InnerText = sNewestItem
                        IsCompoound = NuixCaseReportXML.CreateElement("IsCompoundCase")
                        CaseGUID.AppendChild(IsCompoound)
                        IsCompoound.InnerText = sIsCompound
                        CasesContained = NuixCaseReportXML.CreateElement("CasesContained")
                        CaseGUID.AppendChild(CasesContained)
                        CasesContained.InnerText = sCompoundCaseContains
                        ContainedInCase = NuixCaseReportXML.CreateElement("ContainedInCase")
                        CaseGUID.AppendChild(ContainedInCase)
                        ContainedInCase.InnerText = sContainedInCase
                        CaseLocation = NuixCaseReportXML.CreateElement("CaseLocation")
                        CaseGUID.AppendChild(CaseLocation)
                        CaseLocation.InnerText = sCaseLocation
                        CaseSizeOnDisk = NuixCaseReportXML.CreateElement("CaseSizeOnDisk")
                        CaseGUID.AppendChild(CaseSizeOnDisk)
                        CaseSizeOnDisk.InnerText = sCaseSizeOnDisk
                        Investigator = NuixCaseReportXML.CreateElement("Investigator")
                        CaseGUID.AppendChild(Investigator)
                        Investigator.InnerText = sInvestigator

                        InvestigatorSessions = NuixCaseReportXML.CreateElement("InvestigatorSessions")
                        CaseGUID.AppendChild(InvestigatorSessions)
                        InvestigatorSessions.InnerText = sInvestigatorSessions

                        InvalidSessions = NuixCaseReportXML.CreateElement("InvalidSessions")
                        CaseGUID.AppendChild(InvalidSessions)
                        InvalidSessions.InnerText = sInvalidSessions

                        InvestigatorTimeSummary = NuixCaseReportXML.CreateElement("InvestigatorTimeSummary")
                        CaseGUID.AppendChild(InvestigatorTimeSummary)
                        InvestigatorTimeSummary.InnerText = sInvestigatorTimeSummary
                        BrokerMemory = NuixCaseReportXML.CreateElement("BrokerMemory")
                        CaseGUID.AppendChild(BrokerMemory)
                        BrokerMemory.InnerText = sBrokerMemory
                        WorkerCount = NuixCaseReportXML.CreateElement("WorkerCount")
                        CaseGUID.AppendChild(WorkerCount)
                        WorkerCount.InnerText = sWorkerCount
                        WorkerMemory = NuixCaseReportXML.CreateElement("WorkerMemory")
                        CaseGUID.AppendChild(WorkerMemory)
                        WorkerMemory.InnerText = sWorkerMemory
                        LoadEvents = NuixCaseReportXML.CreateElement("LoadEvents")
                        CaseGUID.AppendChild(LoadEvents)
                        LoadEvents.InnerText = sLoadEvents
                        TotalLoadTime = NuixCaseReportXML.CreateElement("TotalLoadTime")
                        CaseGUID.AppendChild(TotalLoadTime)
                        TotalLoadTime.InnerText = sTotalLoadTime
                        ProcessingSpeed = NuixCaseReportXML.CreateElement("ProcessingSpeed")
                        CaseGUID.AppendChild(ProcessingSpeed)
                        ProcessingSpeed.InnerText = sProcessingSpeed
                End Select
            End If
            row.cells("CollectionStatus").value = "Case Data Exported"
        Next

        NuixCaseReportXML.Save(sReportFilePath & "\" & sOutputFileName)

        MessageBox.Show("Nuix Case Statistics report finished building.  Report located at: " & sOutputFileName)

    End Sub

    Private Sub btnConsoleLocation_Click(sender As Object, e As EventArgs) Handles btnConsoleLocation.Click
        Dim OpenFileDialog1 As New OpenFileDialog
        Dim sNuixApp As String
        Dim NuixVersionInfo As FileVersionInfo

        With OpenFileDialog1
            .Filter = "nuix_console.exe|nuix_console.exe"
            .FilterIndex = 1
        End With

        If (OpenFileDialog1.ShowDialog() = System.Windows.Forms.DialogResult.OK) Then
            txtNuixConsoleLocation.Text = OpenFileDialog1.FileName.ToString
            If txtNuixConsoleLocation.Text <> vbNullString Then
                sNuixApp = txtNuixConsoleLocation.Text

                If sNuixApp.Contains("nuix_console.exe") Then
                    NuixVersionInfo = FileVersionInfo.GetVersionInfo(sNuixApp)
                    lblNuixConsoleVersion.Text = "Nuix Console Version: " & NuixVersionInfo.ProductMajorPart & "." & NuixVersionInfo.ProductMinorPart & "." & NuixVersionInfo.ProductBuildPart
                Else
                    MessageBox.Show("You must select the nuix_console.exe application that you will be using to extract case statistics with.", "nuix_console.exe not selected", MessageBoxButtons.OK)
                    Exit Sub
                End If
            End If
        End If
    End Sub

    Private Sub cboNuixLicenseType_SelectedIndexChanged(sender As Object, e As EventArgs) Handles cboNuixLicenseType.SelectedIndexChanged
        Dim sLicenseType As String

        sLicenseType = cboNuixLicenseType.Text
        Select Case sLicenseType
            Case "eDiscovery Workstation"
                psLicenseType = "enterprise-workstation"
            Case "eDiscovery Reviewer"
                psLicenseType = "enterprise-reviewer"
            Case "Corporate eDiscovery"
                psLicenseType = "corporate-ediscovery"
            Case "Investigative Reviewer"
                psLicenseType = "enterprise-workstation"
            Case "Email Archive Examiner"
                psLicenseType = "email-archive-examiner"
            Case "Ultimate Workstation"
                psLicenseType = "ultimate-workstation"
            Case "Investigation and Response"
                psLicenseType = "law-enforcement-desktop"
        End Select
    End Sub

    Private Sub btnLoadPreviousReportingRun_Click(sender As Object, e As EventArgs) Handles btnLoadPreviousReportingRun.Click
        Dim sReportingDBLocation As String
        Dim bStatus As Boolean
        Dim sShowSizeIn As String

        grdCaseInfo.Rows.Clear()

        If txtReportLocation.Text = vbNullString Then
            MessageBox.Show("You must pick the location of the previous reporting run.", "Select Previous Reporting Run Location")
            txtReportLocation.Focus()
            Exit Sub
        End If

        If cboSizeReporting.Text = vbNullString Then
            MessageBox.Show("You must select a value to show the size in", "Show Size", MessageBoxButtons.OK)
            cboSizeReporting.Focus()

            Exit Sub
        Else
            sShowSizeIn = cboSizeReporting.Text
        End If

        If File.Exists(txtReportLocation.Text & "\" & "NuixCaseReports.db3") Then
            sReportingDBLocation = txtReportLocation.Text
        ElseIf File.Exists(txtReportLocation.Text & "\Scripts\NuixCaseReports.db3") Then
            sReportingDBLocation = txtReportLocation.Text & "\Scripts"
        Else
            sReportingDBLocation = vbNullString
        End If

        If sReportingDBLocation = vbNullString Then
            MessageBox.Show("There is no reporting database located in " & txtReportLocation.Text & " or " & txtReportLocation.Text & "\Scripts - please select a different location to load reports from.", "No Reporting DB found.")
            txtReportLocation.Focus()
            Exit Sub
        End If

        bStatus = blnPopulateCaseInfoGrid(grdCaseInfo, sReportingDBLocation, sShowSizeIn, "%")

    End Sub

    Private Sub btnNuixLogSelector_Click(sender As Object, e As EventArgs) Handles btnNuixLogSelector.Click
        Dim fldrBrowserDialog As New FolderBrowserDialog

        If (fldrBrowserDialog.ShowDialog = System.Windows.Forms.DialogResult.OK) Then

            txtNuixLogDir.Text = fldrBrowserDialog.SelectedPath

        End If
    End Sub

    Public Sub Logger(ByVal sLogFileName As String, ByVal sLogMessage As String)
        Dim UCRTLog As StreamWriter

        If (File.Exists(sLogFileName)) Then
            UCRTLog = File.AppendText(sLogFileName)
            UCRTLog.WriteLine(DateTime.Now.ToString("yyyy-MM-dd hh-mm-ss") & " - " & sLogMessage)

            UCRTLog.Close()
        Else
            Try
                UCRTLog = New StreamWriter(sLogFileName)
                UCRTLog.WriteLine(DateTime.Now.ToString("yyyy-MM-dd hh-mm-ss") & " - " & sLogMessage)
                UCRTLog.Close()
            Catch ex As Exception

            End Try
        End If
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles btnBackupLocationChooser.Click
        Dim fldrBrowserDialog As New FolderBrowserDialog

        If (fldrBrowserDialog.ShowDialog = System.Windows.Forms.DialogResult.OK) Then

            txtBackupLocation.Text = fldrBrowserDialog.SelectedPath

        End If
    End Sub

    Private Sub chkBackUpCase_CheckedChanged(sender As Object, e As EventArgs) Handles chkBackUpCase.CheckedChanged
        If chkBackUpCase.Checked = True Then
            lblBackUpLocation.Enabled = True
            txtBackupLocation.Enabled = True
            btnBackupLocationChooser.Enabled = True
            cboCopyMoveCases.Enabled = True
            If cboUpgradeCasees.Text = "Upgrade Only" Then
                cboCopyMoveCases.Text = "Copy"
                cboCopyMoveCases.Enabled = False
            ElseIf cboUpgradeCasees.Text = "Upgrade and Report" Then
                cboCopyMoveCases.Text = "Copy"
                cboCopyMoveCases.Enabled = False
            Else
                cboCopyMoveCases.Text = ""
                cboCopyMoveCases.Enabled = True
            End If
        Else
            lblBackUpLocation.Enabled = False
            txtBackupLocation.Enabled = False
            btnBackupLocationChooser.Enabled = False
            cboCopyMoveCases.Text = ""
        End If
    End Sub

    Private Sub cboLicenseType_SelectedIndexChanged(sender As Object, e As EventArgs) Handles cboLicenseType.SelectedIndexChanged
        If cboLicenseType.Text = "Desktop" Then
            lblNMSAddress.Enabled = False
            txtNMSAddress.Enabled = False
            txtNMSAddress.Text = vbNullString
            lblNMSInfo.Enabled = False
            txtNMSInfo.Enabled = False
            txtNMSInfo.Text = vbNullString
            lblNMSAddress.Enabled = False
            txtNMSUserName.Enabled = False
            txtNMSUserName.Text = vbNullString
            lblRegistryServer.Enabled = False
            txtRegistryServer.Enabled = False
        ElseIf cboLicenseType.Text = "Desktop (dongleless)" Then
            lblNMSAddress.Enabled = False
            txtNMSAddress.Enabled = False
            txtNMSAddress.Text = vbNullString
            lblNMSInfo.Enabled = False
            txtNMSInfo.Enabled = False
            txtNMSInfo.Text = vbNullString
            lblNMSAddress.Enabled = False
            txtNMSUserName.Enabled = False
            txtNMSUserName.Text = vbNullString
            lblRegistryServer.Enabled = False
            txtRegistryServer.Enabled = False
        Else
            lblNMSAddress.Enabled = True
            txtNMSAddress.Enabled = True
            txtNMSAddress.Text = "127.0.0.1:27443"
            lblNMSInfo.Enabled = True
            txtNMSInfo.Enabled = True
            txtNMSInfo.Text = vbNullString
            lblNMSAddress.Enabled = True
            txtNMSUserName.Enabled = True
            txtNMSUserName.Text = "nuixadmin"
            lblRegistryServer.Enabled = True
            txtRegistryServer.Enabled = True
        End If
    End Sub

    Private Sub txtNuixConsoleLocation_LostFocus(sender As Object, e As EventArgs) Handles txtNuixConsoleLocation.LostFocus
        Dim sNuixApp As String
        Dim NuixVersionInfo As FileVersionInfo

        If txtNuixConsoleLocation.Text <> vbNullString Then
            sNuixApp = txtNuixConsoleLocation.Text

            If sNuixApp.Contains("nuix_console.exe") Then
                NuixVersionInfo = FileVersionInfo.GetVersionInfo(sNuixApp)
                lblNuixConsoleVersion.Text = "Nuix Console Version: " & NuixVersionInfo.ProductMajorPart & "." & NuixVersionInfo.ProductMinorPart & "." & NuixVersionInfo.ProductBuildPart
            Else
                If My.Computer.FileSystem.FileExists(sNuixApp & "\nuix_console.exe") Then
                    txtNuixConsoleLocation.Text = sNuixApp & "\nuix_console.exe"
                    sNuixApp = sNuixApp & "\nuix_console.exe"
                    NuixVersionInfo = FileVersionInfo.GetVersionInfo(sNuixApp)
                    lblNuixConsoleVersion.Text = "Nuix Console Version: " & NuixVersionInfo.ProductMajorPart & "." & NuixVersionInfo.ProductMinorPart & "." & NuixVersionInfo.ProductBuildPart
                Else
                    MessageBox.Show("You must select the nuix_console.exe application that you will be using to extract case statistics with.", "nuix_console.exe not selected", MessageBoxButtons.OK)
                    Exit Sub

                End If
            End If
        End If

    End Sub

    Private Sub grdCaseInfo_CellContentDoubleClick(sender As Object, e As DataGridViewCellEventArgs) Handles grdCaseInfo.CellContentDoubleClick
        If grdCaseInfo("NuixLogLocation", e.RowIndex).Value <> vbNullString Then
            System.Diagnostics.Process.Start(grdCaseInfo("NuixLogLocation", e.RowIndex).Value.ToString)
        End If
    End Sub

    Private Sub chkExportSearchResults_CheckedChanged(sender As Object, e As EventArgs) Handles chkExportSearchResults.CheckedChanged
        If (radSearchFile.Checked = False) And (radSearchTerm.Checked = False) Then
            MessageBox.Show("You must Select either a search term or search file as a type of search term.", "Select search type", MessageBoxButtons.OK)
            chkExportSearchResults.Checked = False
            grpSearchTerm.Focus()
            Exit Sub
        End If
        If (txtSearchTerm.Text = vbNullString) And (chkExportSearchResults.Checked = True) Then
            MessageBox.Show("You must enter a search term in order to export the search results for.", "Enter search Term", MessageBoxButtons.OK)
            chkExportSearchResults.Checked = False
            txtSearchTerm.Focus()
            Exit Sub
        End If

        If chkExportSearchResults.Checked = True Then
            cboExportType.Enabled = True
            txtExportLocation.Enabled = True
            btnExportLocation.Enabled = True
            lblExportLocation.Enabled = True
            chkExportOnly.Show()
            cboExportType.Show()
        Else
            cboExportType.Enabled = False
            txtExportLocation.Enabled = False
            btnExportLocation.Enabled = False
            lblExportLocation.Enabled = False
            txtExportLocation.Text = ""
            chkExportOnly.Hide()
            cboExportType.Hide()

        End If
    End Sub

    Private Sub btnGetFileSystemData_Click(sender As Object, e As EventArgs) Handles btnGetFileSystemData.Click
        Dim bStatus As Boolean
        Dim lstMailBoxTotals As New List(Of String)
        Dim lstExporterMetrics As New List(Of String)
        Dim lstCreatedPST As New List(Of String)
        Dim lstCreatedZIP As New List(Of String)
        Dim lstPSTCustodianName As New List(Of String)
        Dim sNuixConsoleVersion As String
        Dim sScriptsDirectory As String
        Dim sReportFilePath As String
        Dim asCaseFolders() As String
        Dim bMigrateCases As Boolean
        Dim bBackUpCases As Boolean
        Dim sBackUpLocation As String
        Dim sMachineName As String
        Dim iCounter As Integer
        Dim sShowSizeIn As String
        Dim bGetFileSystemDataOnly As Boolean
        Dim NuixCaseFileCSV As StreamReader
        Dim bCopyCases As Boolean
        Dim sCaseName As String
        Dim sCaseDirectory As String
        Dim sCollectionStatus As String
        Dim sBackUpCaseLocation As String
        Dim lstNuixCases As List(Of String)
        Dim lstCaseGUIDs As List(Of String)
        Dim dblCaseSizeOnDisk As Double
        Dim bIncludeDiskSize As Boolean

        Dim value As System.Version = My.Application.Info.Version

        bIncludeDiskSize = chkIncludeDiskSize.Checked

        grdCaseInfo.Rows.Clear()
        lstNuixCases = New List(Of String)
        lstCaseGUIDs = New List(Of String)

        Try
            sShowSizeIn = cboSizeReporting.Text
            bGetFileSystemDataOnly = True
            If txtNuixLogDir.Text = vbNullString Then
                MessageBox.Show("You must select the location of the Nuix Log Files.", "Nuix Log File Directory not selected")
                txtNuixLogDir.Focus()
                Exit Sub
            End If

            If (txtReportLocation.Text = vbNullString) Then
                MsgBox("You Must Enter the location to create the Case report file.")
                txtReportLocation.Focus()
                Exit Sub
            Else
                sReportFilePath = txtReportLocation.Text
                sScriptsDirectory = sReportFilePath & "\" & "Scripts"
                Directory.CreateDirectory(sScriptsDirectory)

                bStatus = blnBuildSQLiteDB(sScriptsDirectory)
                bStatus = blnBuildSQLiteDatabaseScript(sScriptsDirectory)
                bStatus = blnBuildSQLiteRubyScript(sScriptsDirectory)
            End If

            If chkBackUpCase.Checked = True Then
                If cboCopyMoveCases.Text = vbNullString Then
                    MessageBox.Show("You must select whether you want to copy or move the cases to the backup location.", "Copy or Move cases to backup location", MessageBoxButtons.OK)
                    Exit Sub
                Else
                    bBackUpCases = True
                    If txtBackupLocation.Text = vbNullString Then
                        MessageBox.Show("You have not selected a back up location.", "Select backup location", MessageBoxButtons.OK)
                        txtBackupLocation.Focus()
                        Exit Sub
                    Else
                        sBackUpLocation = txtBackupLocation.Text
                    End If
                    If cboCopyMoveCases.Text = "Copy" Then
                        bCopyCases = True
                    Else
                        bCopyCases = False
                    End If
                End If
            Else
                bBackUpCases = False
                sBackUpLocation = ""
            End If


            If radFile.Checked = True Then
                If txtCaseFileLocations.Text = vbNullString Then
                    MessageBox.Show("You must select a CSV file containing paths to Nuix Case files", "Case file CSV required", MessageBoxButtons.OK)
                    txtCaseFileLocations.Focus()
                    Exit Sub
                Else
                    NuixCaseFileCSV = New StreamReader(txtCaseFileLocations.Text)
                    While Not NuixCaseFileCSV.EndOfStream
                        plstSoureFolders.Add(NuixCaseFileCSV.ReadLine)
                    End While
                End If
            End If

            If plstSoureFolders.Count = 0 Then
                MessageBox.Show("You must select at least one folder to search for Nuix Cases in.", "Select Nuix Case Directory")
                Exit Sub
            End If

            If (txtNuixConsoleLocation.Text = vbNullString) Then
                MessageBox.Show("You must enter the location of the Nuix Console version to use.", "Nuix Console Version Location")
                txtNuixConsoleLocation.Focus()
                Exit Sub
            Else
                sNuixConsoleVersion = lblNuixConsoleVersion.Text.Replace("Nuix Console Version: ", "")
            End If


            asCaseFolders = plstSoureFolders.ToArray
            sMachineName = System.Net.Dns.GetHostName

            psUCRTLogFile = sScriptsDirectory & "\UCRT Log - " & sMachineName & "-" & DateTime.Now.ToString("yyyy-MM-dd hh-mm-ss") & ".log"

            Logger(psUCRTLogFile, "Nuix Universal Case Reporting Tool - " & psUCRTLogFile)
            iCounter = 0
            For Each sourceLocation In plstSoureFolders
                Logger(psUCRTLogFile, "Source - " & iCounter & " - " & sourceLocation.ToString)
                iCounter = iCounter + 1
            Next
            Try
                If (txtReportLocation.Text = vbNullString) Then
                    MsgBox("You Must Enter the location to create the Case report file.")
                    txtReportLocation.Focus()
                    Exit Sub
                Else
                    sReportFilePath = txtReportLocation.Text
                    sScriptsDirectory = sReportFilePath & "\" & "Scripts"
                    Directory.CreateDirectory(sScriptsDirectory)

                    bStatus = blnBuildSQLiteDB(sScriptsDirectory)
                    bStatus = blnBuildSQLiteDatabaseScript(sScriptsDirectory)
                    bStatus = blnBuildSQLiteRubyScript(sScriptsDirectory)
                End If

                If Not IsNothing(asCaseFolders) Then
                    bStatus = blnGetAllNuixCaseFiles(sScriptsDirectory, sNuixConsoleVersion, asCaseFolders, bMigrateCases, False, bIncludeDiskSize, lstNuixCases, lstCaseGUIDs)
                    If chkBackUpCase.Checked = True Then
                        Me.Text = "Universal Case Reporting tool - " & value.ToString & " - Copying/Moving Cases - Please Wait"

                        For Each NuixGUID In lstCaseGUIDs
                            bStatus = blnGetUpdatedDBInfo(sScriptsDirectory, NuixGUID.ToString, "CaseName", sCaseName, "TEXT")
                            bStatus = blnGetUpdatedDBInfo(sScriptsDirectory, NuixGUID.ToString, "CaseLocation", sCaseDirectory, "TEXT")
                            bStatus = blnGetUpdatedDBInfo(sScriptsDirectory, NuixGUID.ToString, "CaseSizeOnDisk", dblCaseSizeOnDisk, "INT")
                            bStatus = blnCopyCase(sCaseName, sCaseDirectory, sBackUpLocation, bCopyCases, sBackUpCaseLocation, dblCaseSizeOnDisk)
                            sCollectionStatus = "Case Copied or Moved"
                            bStatus = blnUpdateSQLiteReportingDB(sScriptsDirectory, NuixGUID.ToString, "BackUpLocation", sBackUpCaseLocation)
                        Next

                    End If
                End If
                Me.Text = "Universal Case Reporting tool - " & value.ToString

                bStatus = blnPopulateCaseInfoGrid(grdCaseInfo, sScriptsDirectory, sShowSizeIn, "'File System Info Collected - Case Migrating', 'File System Info Collected - Case Version Mismatch', 'File System Info Collected - Waiting for Case Data', 'Case Locked'")

            Catch ex As Exception
                Logger(psUCRTLogFile, ex.ToString)
            End Try
        Catch ex As Exception
            Logger(psUCRTLogFile, ex.ToString)
        End Try
    End Sub


    Private Sub ExcelToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles ExcelToolStripMenuItem.Click
        Dim sCaseGuid As String
        Dim sCollectionStatus As String
        Dim sReportLoadDuration As String
        Dim sCaseName As String
        Dim sCurrentCaseVersion As String
        Dim sUpgradedCaseVersion As String
        Dim sBatchLoadInfo As String
        Dim sPercentComplete As String
        Dim sCaseLocation As String
        Dim sCaseDescription As String
        Dim sCaseSizeOnDisk As String
        Dim sCaseFileSize As String
        Dim sCaseAuditSize As String
        Dim sOldestTopLevel As String
        Dim sNewestTopLevel As String
        Dim sIsCompound As String
        Dim sCasesContained As String
        Dim sContainedInCase As String
        Dim sInvestigator As String
        Dim sInvestigatorSessions As String
        Dim sInvalidSessions As String
        Dim sInvestigatorTimeSummary As String
        Dim sBrokerMemory As String
        Dim sWorkerCount As String
        Dim sWorkerMemory As String
        Dim sEvidenceName As String
        Dim sEvidenceLocation As String
        Dim sEvidenceDescription As String
        Dim sEvidenceCustomMetadata As String
        Dim sLanguagesContained As String
        Dim sMimeTypes As String
        Dim sItemTypes As String
        Dim sIrregularItems As String
        Dim sCreationDate As String
        Dim sModifiedDate As String
        Dim sLoadStartDate As String
        Dim sLoadEndDate As String
        Dim sLoadTime As String
        Dim sLoadEvents As String
        Dim sTotalLoadTime As String
        Dim sProcessingSpeed As String
        Dim sCustodians As String
        Dim sCustodianCount As String
        Dim sSearchTerm As String
        Dim sSearchSize As String
        Dim sSearchHitCount As String
        Dim sCustodianSearchHit As String
        Dim sTotalCaseItemCount As String
        Dim sTopLevelEmailCount As String
        Dim sDuplicateItems As String
        Dim sOriginalItems As String
        Dim sHitCountPercent As String
        Dim asLanguages() As String
        Dim asLanguageDetails() As String
        Dim asMimeTypes() As String
        Dim asMimeTypeDetails() As String
        Dim asItemTypes() As String
        Dim asItemTypeDetails() As String
        Dim asIrregularItems() As String
        Dim asIrregularItemsDetails() As String
        Dim asCustomMetadata() As String
        Dim asCustomMetadataDetails() As String
        Dim asBatchLoadInfo() As String
        Dim asBatchLoadDetails() As String
        Dim sBatchLoadDate As String
        Dim sBatchLoadCount As String
        Dim sBatchLoadFileSize As String
        Dim sBatchLoadAuditSize As String
        Dim iSummarySheetCounter As Integer
        Dim iSessionCount As Integer
        Dim sReportingDBLocation As String
        Dim sShowSizeIN As String
        Dim sNuixLogLocation As String
        Dim asInvestigatorTimeSessions() As String
        Dim asInvestigatorTimeSessionDetails() As String
        Dim iSummarySheetLocation As Integer
        Dim asSessionDuration() As String
        Dim sCloseSession As String
        Dim iIrregularItemsCount As Integer
        Dim iLanguageCount As Integer
        Dim sItemCounts As String
        Dim sDataExport As String

        Dim mSQL As String
        Dim dt As DataTable
        Dim ds As DataSet
        Dim dataReader As SQLiteDataReader
        Dim sqlCommand As SQLiteCommand
        Dim sqlConnection As SQLiteConnection

        Dim lstItemType As List(Of String)
        Dim lstItemTypeCount As List(Of Double)
        Dim lstItemTypeSize As List(Of Double)
        Dim lstMimeType As List(Of String)
        Dim lstMimeTypeCount As List(Of Double)
        Dim lstMimeTypeSize As List(Of Double)
        Dim lstTotalItem As List(Of String)
        Dim lstTotalItemCount As List(Of Double)
        Dim lstTotalItemSize As List(Of Double)
        Dim lstOriginalItem As List(Of String)
        Dim lstOriginalItemCount As List(Of Double)
        Dim lstOriginalItemSize As List(Of Double)
        Dim lstDuplicateItem As List(Of String)
        Dim lstDuplicateItemCount As List(Of Double)
        Dim lstDuplicateItemSize As List(Of Double)

        Dim lstLanguages As List(Of String)
        Dim lstLanguagesCount As List(Of Double)
        Dim lstLanguagesSize As List(Of Double)
        Dim lstIrregularItems As List(Of String)
        Dim lstIrregularItemsCount As List(Of Double)
        Dim lstIrregularItemsSize As List(Of Double)
        Dim iTotalItemIndex As Integer
        Dim dblTotalItemCount As Double
        Dim dblTotalItemSize As Double
        Dim dblOriginalItemCount As Double
        Dim dblOriginalItemSize As Double
        Dim dblDuplicateItemCount As Double
        Dim dblDuplidateItemSize As Double

        Dim iLanguageIndex As Integer
        Dim dblTotalLanguageCount As Double
        Dim dblTotalLanguageSize As Double

        Dim iMimeTypeIndex As Integer
        Dim dblTotalMimeTypeCount As Double
        Dim dblTotalMimeTypeSize As Double

        Dim iItemTypeIndex As Integer
        Dim dblTotalItemItemCount As Double
        Dim dblTotalItemItemSize As Double

        Dim iIrregularItemIndex As Integer
        Dim dblTotalIrregularItemCount As Double
        Dim dblTotalIrregularItemSize As Double

        Dim sItemDate As String
        Dim sMonth As String
        Dim sYear As String
        Dim sWeekday As String
        Dim sMonthNumber As String
        Dim sWeekNumber As String
        Dim dItemDate As Date

        lstItemType = New List(Of String)
        lstItemTypeCount = New List(Of Double)
        lstItemTypeSize = New List(Of Double)
        lstMimeType = New List(Of String)
        lstMimeTypeCount = New List(Of Double)
        lstMimeTypeSize = New List(Of Double)
        lstTotalItem = New List(Of String)
        lstTotalItemCount = New List(Of Double)
        lstTotalItemSize = New List(Of Double)
        lstOriginalItem = New List(Of String)
        lstOriginalItemCount = New List(Of Double)
        lstOriginalItemSize = New List(Of Double)
        lstDuplicateItem = New List(Of String)
        lstDuplicateItemCount = New List(Of Double)
        lstDuplicateItemSize = New List(Of Double)
        lstLanguages = New List(Of String)
        lstLanguagesCount = New List(Of Double)
        lstLanguagesSize = New List(Of Double)
        lstIrregularItems = New List(Of String)
        lstIrregularItemsCount = New List(Of Double)
        lstIrregularItemsSize = New List(Of Double)

        Dim oExcelApp As Excel.Application

        sShowSizeIN = cboSizeReporting.Text
        If sShowSizeIN = vbNullString Then
            MessageBox.Show("You must select the value to show the size in", "Show Size in", MessageBoxButtons.OK)
            cboSizeReporting.Focus()
            Exit Sub
        End If
        If File.Exists(txtReportLocation.Text & "\" & "NuixCaseReports.db3") Then
            sReportingDBLocation = txtReportLocation.Text
        ElseIf File.Exists(txtReportLocation.Text & "\Scripts\NuixCaseReports.db3") Then
            sReportingDBLocation = txtReportLocation.Text & "\Scripts"
        Else
            MessageBox.Show("There is no SQLite DB in the report location selected (or in the scripts location).  Please select a report location where a SQLliteDB exists.", "No SQLiteDB Found", MessageBoxButtons.OK)
            txtReportLocation.Focus()
            Exit Sub
        End If

        For Each row In grdCaseInfo.Rows
            iSummarySheetCounter = 0
            iSummarySheetLocation = 0
            iSummarySheetCounter = 0
            iSessionCount = 0
            If ((row.cells("CaseGuid").value <> vbNullString) And ((row.cells("CollectionStatus").value = "File System and Case Data collected") Or (row.cells("CollectionStatus").value = "Case No Longer Exists"))) Then
                row.cells("CollectionStatus").value = "Exporting Case Data..."
                iIrregularItemsCount = 0

                sCaseGuid = row.cells("CaseGuid").value
                sCollectionStatus = row.cells("CollectionStatus").value
                sPercentComplete = row.cells("PercentComplete").value
                sReportLoadDuration = row.cells("ReportLoadDuration").value
                sCaseName = row.cells("CaseName").value
                sCurrentCaseVersion = row.cells("CurrentCaseVersion").value
                sUpgradedCaseVersion = row.cells("UpgradedCaseVersion").value
                sBatchLoadInfo = row.cells("BatchLoadInfo").value
                sDataExport = row.cells("DataExport").value
                sCaseLocation = row.cells("CaseLocation").value
                sCaseDescription = row.cells("CaseDescription").value
                sCaseSizeOnDisk = row.cells("CaseSizeOnDisk").value
                sCaseFileSize = row.cells("CaseFileSize").value
                sCaseAuditSize = row.cells("CaseAuditSize").value
                sOldestTopLevel = row.cells("OldestTopLevel").value
                sNewestTopLevel = row.cells("NewestTopLevel").value
                sIsCompound = row.cells("IsCompound").value
                sCasesContained = row.cells("CasesContained").value
                sContainedInCase = row.cells("ContainedInCase").value
                sInvestigator = row.cells("Investigator").value
                sInvestigatorSessions = row.cells("InvestigatorSessions").value
                sInvalidSessions = row.cells("InvalidSessions").value
                sInvestigatorTimeSummary = row.cells("InvestigatorTimeSummary").value
                sBrokerMemory = row.cells("BrokerMemory").value
                sWorkerCount = row.cells("WorkerCount").value
                sWorkerMemory = row.cells("WorkerMemory").value
                sEvidenceName = row.cells("EvidenceName").value
                sEvidenceLocation = row.cells("EvidenceLocation").value
                sEvidenceDescription = row.cells("EvidenceDescription").value
                sEvidenceCustomMetadata = row.cells("EvidenceCustomMetadata").value
                sLanguagesContained = row.cells("LanguagesContained").value
                sMimeTypes = row.cells("MimeTypes").value
                sItemTypes = row.cells("ItemTypes").value
                sIrregularItems = row.cells("IrregularItems").value
                sCreationDate = row.cells("CreationDate").value
                sModifiedDate = row.cells("ModifiedDate").value
                sLoadStartDate = row.cells("LoadStartDate").value
                sLoadEndDate = row.cells("LoadEndDate").value
                sLoadTime = row.cells("LoadTime").value
                sLoadEvents = row.cells("LoadEvents").value
                sTotalLoadTime = row.cells("TotalLoadTime").value
                sProcessingSpeed = row.cells("ProcessingSpeed").value
                sTotalCaseItemCount = row.cells("TotalCaseItemCount").value
                sItemCounts = row.cells("ItemCounts").value
                sDuplicateItems = row.cells("DuplicateItems").value
                sOriginalItems = row.cells("OriginalItems").value
                sCustodians = row.cells("Custodians").value
                sCustodianCount = row.cells("CustodianCount").value
                sSearchTerm = row.cells("SearchTerm").value
                sSearchSize = row.cells("SearchSize").value
                sSearchHitCount = row.cells("SearchHitCount").value
                sCustodianSearchHit = row.cells("CustodianSearchHit").value
                sHitCountPercent = row.cells("HitCountPercent").value
                sNuixLogLocation = row.cells("NuixLogLocation").value
                'oXLApp = New Microsoft.Office.Interop.Excel.Application
                Try
                    oExcelApp = New Microsoft.Office.Interop.Excel.Application

                Catch ex As Exception
                    MessageBox.Show("It appears that Microsoft Excel is not installed on this machine. It is required to have Excel installed in order to export to Excel.", "Dependency missing", MessageBoxButtons.OK)
                    Exit Sub
                End Try
                Dim sExcelVersion As String
                sExcelVersion = oExcelApp.Version.ToString
                Logger(psUCRTLogFile, "Excel Version Number = " & sExcelVersion)
                'oExcelWorkbook = New Excel.Workbook
                Dim xlWorkbook As Excel.Workbook = oExcelApp.Workbooks.Add()
                Dim xlIrregularItems As Excel.Worksheet = CType(xlWorkbook.Sheets.Add(), Excel.Worksheet)
                xlIrregularItems.Name = "Irregular Items"
                oExcelApp.Windows.Application.ActiveWindow.DisplayGridlines = False
                Dim xlEncryptedMessagesSheet As Excel.Worksheet = CType(xlWorkbook.Sheets.Add(), Excel.Worksheet)
                xlEncryptedMessagesSheet.Name = "Encrypted Messages"
                oExcelApp.Windows.Application.ActiveWindow.DisplayGridlines = False
                Dim xlCulling As Excel.Worksheet = CType(xlWorkbook.Sheets.Add(), Excel.Worksheet)
                xlCulling.Name = "Culling"
                oExcelApp.Windows.Application.ActiveWindow.DisplayGridlines = False
                Dim xlDateRangeGapPivot As Excel.Worksheet = CType(xlWorkbook.Sheets.Add(), Excel.Worksheet)
                xlDateRangeGapPivot.Name = "DateRangeGap PIVOT"
                oExcelApp.Windows.Application.ActiveWindow.DisplayGridlines = False
                Dim xlDateRangeSheet As Excel.Worksheet = CType(xlWorkbook.Sheets.Add(), Excel.Worksheet)
                xlDateRangeSheet.Name = "DateRangeSummary"
                oExcelApp.Windows.Application.ActiveWindow.DisplayGridlines = False
                Dim xlNoDataDateRangeSheet As Excel.Worksheet = CType(xlWorkbook.Sheets.Add(), Excel.Worksheet)
                xlNoDataDateRangeSheet.Name = "NoItemsDateRange"
                oExcelApp.Windows.Application.ActiveWindow.DisplayGridlines = False
                Dim xlLanguagesSheet As Excel.Worksheet = CType(xlWorkbook.Sheets.Add(), Excel.Worksheet)
                xlLanguagesSheet.Name = "Languages"
                oExcelApp.Windows.Application.ActiveWindow.DisplayGridlines = False
                Dim xlMimeTypeSheet As Excel.Worksheet = CType(xlWorkbook.Sheets.Add(), Excel.Worksheet)
                xlMimeTypeSheet.Name = "MimeTypes"
                oExcelApp.Windows.Application.ActiveWindow.DisplayGridlines = False
                Dim xlItemTypeSheet As Excel.Worksheet = CType(xlWorkbook.Sheets.Add(), Excel.Worksheet)
                xlItemTypeSheet.Name = "ItemTypes"
                oExcelApp.Windows.Application.ActiveWindow.DisplayGridlines = False
                Dim xlSummarySheet As Excel.Worksheet = CType(xlWorkbook.Sheets.Add(), Excel.Worksheet)
                xlSummarySheet.Name = "SummarySheet"
                oExcelApp.Windows.Application.ActiveWindow.DisplayGridlines = False
                Dim xlSummarySheet2 As Excel.Worksheet = CType(xlWorkbook.Sheets.Add(), Excel.Worksheet)
                xlSummarySheet2.Name = "Case Summary"
                oExcelApp.Windows.Application.ActiveWindow.DisplayGridlines = False

                Dim xlPercentRange As Excel.Range
                Dim xlNumberRange As Excel.Range
                Dim xlHighlightRange As Excel.Range
                'xlWorkbook.SaveAs(txtReportLocation.Text & "\" & sCaseName & ".xlsx")

                Dim ImageRange As Excel.Range

                ImageRange = xlSummarySheet.Range("A1:A4")
                ImageRange.MergeCells = True
                Dim sImageFile As String = IO.Path.Combine(Application.StartupPath, "Resources\nuix-logo-updated.jpg")

                xlSummarySheet.Shapes.AddPicture(sImageFile, False, True, 0, 0, 100, 60)

                xlSummarySheet.Cells(5, 1) = "Case Name"
                xlSummarySheet.Cells(5, 2) = sCaseName
                xlSummarySheet.Cells(6, 1) = "Case Guid"
                xlSummarySheet.Cells(6, 2) = sCaseGuid
                xlSummarySheet.Cells(7, 1) = "Case Version"
                xlSummarySheet.Cells(7, 2) = sCurrentCaseVersion
                xlSummarySheet.Cells(8, 1) = "Oldest Case Item"
                xlSummarySheet.Cells(8, 2) = sOldestTopLevel
                xlSummarySheet.Cells(8, 2).NumberFormat = "[$-en-US]mm/dd/yyyy;@"
                xlSummarySheet.Cells(9, 1) = "Newest Case Item"
                xlSummarySheet.Cells(9, 2) = sNewestTopLevel
                xlSummarySheet.Cells(9, 2).NumberFormat = "[$-en-US]mm/dd/yyyy;@"
                xlSummarySheet.Cells(10, 1) = "Case Location"
                xlSummarySheet.Cells(10, 2) = sCaseLocation
                xlSummarySheet.Cells(11, 1) = "Total Case Item Count"
                xlSummarySheet.Cells(11, 2) = sTotalCaseItemCount
                xlNumberRange = xlSummarySheet.Cells(11, 2)
                xlNumberRange.NumberFormat = "#,##0"

                xlSummarySheet.Cells(12, 1) = "Top Level Email Count"
                xlSummarySheet.Cells(12, 2) = ""
                xlSummarySheet.Cells(13, 1) = "Investigator"
                xlSummarySheet.Cells(13, 2) = sInvestigator
                xlSummarySheet.Cells(14, 1) = "Description"
                xlSummarySheet.Cells(14, 2) = sCaseDescription
                xlSummarySheet.Cells(15, 1) = "Ingestion Start Date"
                xlSummarySheet.Cells(15, 2) = sLoadStartDate
                xlSummarySheet.Cells(16, 1) = "Ingestion End Date"
                xlSummarySheet.Cells(16, 2) = sLoadEndDate
                xlSummarySheet.Cells(17, 1) = "Collection Status"
                xlSummarySheet.Cells(17, 2) = sCollectionStatus
                xlSummarySheet.Cells(18, 1) = "Report Load Duration"
                xlSummarySheet.Cells(18, 2) = sReportLoadDuration
                xlSummarySheet.Cells(19, 1) = "Batch Load Date"
                xlSummarySheet.Cells(19, 2) = "Batch Load Number Of Items"
                xlSummarySheet.Cells(19, 3) = "Batch Load File Size"
                xlSummarySheet.Cells(19, 4) = "Batch Load Audit Size"
                iSummarySheetCounter = 19
                asBatchLoadInfo = Split(sBatchLoadInfo, ";")
                For iCounter = 0 To UBound(asBatchLoadInfo) - 1
                    asBatchLoadDetails = Split(asBatchLoadInfo(iCounter), "::")
                    sBatchLoadDate = asBatchLoadDetails(0)
                    sBatchLoadCount = asBatchLoadDetails(1)
                    sBatchLoadFileSize = asBatchLoadDetails(2)
                    sBatchLoadAuditSize = asBatchLoadDetails(3)
                    iSummarySheetCounter = iSummarySheetCounter + 1
                    xlSummarySheet.Cells(iSummarySheetCounter, 1) = sBatchLoadDate
                    xlSummarySheet.Cells(iSummarySheetCounter, 2) = sBatchLoadCount
                    xlSummarySheet.Cells(iSummarySheetCounter, 3) = sBatchLoadFileSize
                    xlNumberRange = xlSummarySheet.Cells(iSummarySheetCounter, 2)
                    xlNumberRange.NumberFormat = "#,##0"
                    xlNumberRange = xlSummarySheet.Cells(iSummarySheetCounter, 3)
                    xlNumberRange.NumberFormat = "#,##0"
                    xlSummarySheet.Cells(iSummarySheetCounter, 4) = sBatchLoadAuditSize
                    xlNumberRange = xlSummarySheet.Cells(iSummarySheetCounter, 4)
                    xlNumberRange.NumberFormat = "#,##0"
                    xlSummarySheet.Cells(iSummarySheetCounter, 5) = "Bytes"

                Next
                If sDataExport <> vbNullString Then
                    xlSummarySheet.Cells(iSummarySheetCounter + 1, 1) = "Data Export Information"
                    xlSummarySheet.Cells(iSummarySheetCounter + 1, 2) = sDataExport
                End If
                xlSummarySheet.Cells(iSummarySheetCounter + 2, 1) = "Case Size On Disk"
                xlSummarySheet.Cells(iSummarySheetCounter + 2, 2) = sCaseSizeOnDisk
                xlNumberRange = xlSummarySheet.Cells(iSummarySheetCounter + 2, 2)
                xlNumberRange.NumberFormat = "#,##0"
                xlSummarySheet.Cells(iSummarySheetCounter + 3, 1) = "Case File Size"
                xlSummarySheet.Cells(iSummarySheetCounter + 3, 2) = sCaseFileSize
                xlNumberRange = xlSummarySheet.Cells(iSummarySheetCounter + 3, 2)
                xlNumberRange.NumberFormat = "#,##0"
                xlSummarySheet.Cells(iSummarySheetCounter + 4, 1) = "Case Audit Size"
                xlSummarySheet.Cells(iSummarySheetCounter + 4, 2) = sCaseFileSize
                xlNumberRange = xlSummarySheet.Cells(iSummarySheetCounter + 4, 2)
                xlNumberRange.NumberFormat = "#,##0"
                xlSummarySheet.Cells(iSummarySheetCounter + 5, 1) = "Is Compound Case"
                xlSummarySheet.Cells(iSummarySheetCounter + 5, 2) = sIsCompound
                xlSummarySheet.Cells(iSummarySheetCounter + 6, 1) = "Cases Contained"
                xlSummarySheet.Cells(iSummarySheetCounter + 6, 2) = sCasesContained
                xlSummarySheet.Cells(iSummarySheetCounter + 7, 1) = "Contained in Case"
                xlSummarySheet.Cells(iSummarySheetCounter + 7, 2) = sContainedInCase
                xlSummarySheet.Cells(iSummarySheetCounter + 8, 1) = "Broker Memory"
                xlSummarySheet.Cells(iSummarySheetCounter + 8, 2) = sBrokerMemory
                xlSummarySheet.Cells(iSummarySheetCounter + 9, 1) = "Worker Count"
                xlSummarySheet.Cells(iSummarySheetCounter + 9, 2) = sWorkerCount
                xlSummarySheet.Cells(iSummarySheetCounter + 10, 1) = "Worker Memory"
                xlSummarySheet.Cells(iSummarySheetCounter + 10, 2) = sWorkerMemory
                xlSummarySheet.Cells(iSummarySheetCounter + 11, 1) = "Evidence Name"
                xlSummarySheet.Cells(iSummarySheetCounter + 11, 2) = sEvidenceName
                xlSummarySheet.Cells(iSummarySheetCounter + 12, 1) = "Evidence Location"
                xlSummarySheet.Cells(iSummarySheetCounter + 12, 2) = sEvidenceLocation
                xlSummarySheet.Cells(iSummarySheetCounter + 13, 1) = "Evidence Description"
                xlSummarySheet.Cells(iSummarySheetCounter + 13, 2) = sEvidenceDescription
                xlSummarySheet.Cells(iSummarySheetCounter + 14, 1) = "Creation Date"
                xlSummarySheet.Cells(iSummarySheetCounter + 14, 2) = sCreationDate
                xlSummarySheet.Cells(iSummarySheetCounter + 15, 1) = "Modified Date"
                xlSummarySheet.Cells(iSummarySheetCounter + 15, 2) = sModifiedDate
                xlSummarySheet.Cells(iSummarySheetCounter + 16, 1) = "Load Time"
                xlSummarySheet.Cells(iSummarySheetCounter + 16, 2) = sLoadTime
                xlSummarySheet.Cells(iSummarySheetCounter + 17, 1) = "Load Events"
                xlSummarySheet.Cells(iSummarySheetCounter + 17, 2) = sLoadEvents
                xlSummarySheet.Cells(iSummarySheetCounter + 18, 1) = "Total Load Time"
                xlSummarySheet.Cells(iSummarySheetCounter + 18, 2) = sTotalLoadTime
                xlSummarySheet.Cells(iSummarySheetCounter + 19, 1) = "Processing Speed"
                xlSummarySheet.Cells(iSummarySheetCounter + 19, 2) = sProcessingSpeed
                xlSummarySheet.Cells(iSummarySheetCounter + 20, 1) = "Custodians"
                xlSummarySheet.Cells(iSummarySheetCounter + 20, 2) = sCustodians
                xlSummarySheet.Cells(iSummarySheetCounter + 21, 1) = "Custodian Count"
                xlSummarySheet.Cells(iSummarySheetCounter + 21, 2) = sCustodianCount
                xlSummarySheet.Cells(iSummarySheetCounter + 22, 1) = "Search Term"
                xlSummarySheet.Cells(iSummarySheetCounter + 22, 2) = sSearchTerm
                xlSummarySheet.Cells(iSummarySheetCounter + 23, 1) = "Search Size"
                xlSummarySheet.Cells(iSummarySheetCounter + 23, 2) = sSearchSize
                xlSummarySheet.Cells(iSummarySheetCounter + 24, 1) = "Custodian Search Hit"
                xlSummarySheet.Cells(iSummarySheetCounter + 24, 2) = sCustodianSearchHit
                xlSummarySheet.Cells(iSummarySheetCounter + 25, 1) = "Hit Count Percent"
                xlSummarySheet.Cells(iSummarySheetCounter + 25, 2) = sHitCountPercent
                xlSummarySheet.Cells(iSummarySheetCounter + 26, 1) = "Nuix Log Location"
                xlSummarySheet.Cells(iSummarySheetCounter + 26, 2) = sNuixLogLocation
                xlSummarySheet.Cells(iSummarySheetCounter + 27, 1) = "Invalid Sessions"
                xlSummarySheet.Cells(iSummarySheetCounter + 27, 2) = sInvalidSessions
                xlSummarySheet.Cells(iSummarySheetCounter + 28, 1) = "Investigator Time Summary"
                xlSummarySheet.Cells(iSummarySheetCounter + 28, 2) = sInvestigatorTimeSummary
                xlSummarySheet.Cells(iSummarySheetCounter + 28, 1) = "Investigator Sessions"
                xlSummarySheet.Cells(iSummarySheetCounter + 28, 2) = "Investigator Name"
                xlSummarySheet.Cells(iSummarySheetCounter + 28, 3) = "Open"
                xlSummarySheet.Cells(iSummarySheetCounter + 28, 4) = "Close"
                xlSummarySheet.Cells(iSummarySheetCounter + 28, 5) = "Duration"

                Dim SummarySheetRange As Excel.Range
                SummarySheetRange = xlSummarySheet.Range("A5:A" & iSummarySheetCounter + 28)
                With SummarySheetRange.Borders(Excel.XlBordersIndex.xlEdgeLeft)
                    .LineStyle = Excel.XlLineStyle.xlContinuous
                    .ColorIndex = 0
                    .TintAndShade = 0
                    .Weight = Excel.XlBorderWeight.xlThick
                End With
                With SummarySheetRange.Borders(Excel.XlBordersIndex.xlEdgeTop)
                    .LineStyle = Excel.XlLineStyle.xlContinuous
                    .ColorIndex = 0
                    .TintAndShade = 0
                    .Weight = Excel.XlBorderWeight.xlThick
                End With
                With SummarySheetRange.Borders(Excel.XlBordersIndex.xlEdgeBottom)
                    .LineStyle = Excel.XlLineStyle.xlContinuous
                    .ColorIndex = 0
                    .TintAndShade = 0
                    .Weight = Excel.XlBorderWeight.xlThick
                End With
                With SummarySheetRange.Borders(Excel.XlBordersIndex.xlEdgeRight)
                    .LineStyle = Excel.XlLineStyle.xlContinuous
                    .ColorIndex = 0
                    .TintAndShade = 0
                    .Weight = Excel.XlBorderWeight.xlThick
                End With


                SummarySheetRange = xlSummarySheet.Range("B5:B" & iSummarySheetCounter + 28)
                With SummarySheetRange.Borders(Excel.XlBordersIndex.xlEdgeLeft)
                    .LineStyle = Excel.XlLineStyle.xlContinuous
                    .ColorIndex = 0
                    .TintAndShade = 0
                    .Weight = Excel.XlBorderWeight.xlThick
                End With
                With SummarySheetRange.Borders(Excel.XlBordersIndex.xlEdgeTop)
                    .LineStyle = Excel.XlLineStyle.xlContinuous
                    .ColorIndex = 0
                    .TintAndShade = 0
                    .Weight = Excel.XlBorderWeight.xlThick
                End With
                With SummarySheetRange.Borders(Excel.XlBordersIndex.xlEdgeBottom)
                    .LineStyle = Excel.XlLineStyle.xlContinuous
                    .ColorIndex = 0
                    .TintAndShade = 0
                    .Weight = Excel.XlBorderWeight.xlThick
                End With
                With SummarySheetRange.Borders(Excel.XlBordersIndex.xlEdgeRight)
                    .LineStyle = Excel.XlLineStyle.xlContinuous
                    .ColorIndex = 0
                    .TintAndShade = 0
                    .Weight = Excel.XlBorderWeight.xlThick
                End With

                SummarySheetRange = xlSummarySheet.Range("C5:C" & iSummarySheetCounter + 28)
                With SummarySheetRange.Borders(Excel.XlBordersIndex.xlEdgeLeft)
                    .LineStyle = Excel.XlLineStyle.xlContinuous
                    .ColorIndex = 0
                    .TintAndShade = 0
                    .Weight = Excel.XlBorderWeight.xlThick
                End With
                With SummarySheetRange.Borders(Excel.XlBordersIndex.xlEdgeTop)
                    .LineStyle = Excel.XlLineStyle.xlContinuous
                    .ColorIndex = 0
                    .TintAndShade = 0
                    .Weight = Excel.XlBorderWeight.xlThick
                End With
                With SummarySheetRange.Borders(Excel.XlBordersIndex.xlEdgeBottom)
                    .LineStyle = Excel.XlLineStyle.xlContinuous
                    .ColorIndex = 0
                    .TintAndShade = 0
                    .Weight = Excel.XlBorderWeight.xlThick
                End With
                With SummarySheetRange.Borders(Excel.XlBordersIndex.xlEdgeRight)
                    .LineStyle = Excel.XlLineStyle.xlContinuous
                    .ColorIndex = 0
                    .TintAndShade = 0
                    .Weight = Excel.XlBorderWeight.xlThick
                End With

                SummarySheetRange = xlSummarySheet.Range("D5:D" & iSummarySheetCounter + 28)
                With SummarySheetRange.Borders(Excel.XlBordersIndex.xlEdgeLeft)
                    .LineStyle = Excel.XlLineStyle.xlContinuous
                    .ColorIndex = 0
                    .TintAndShade = 0
                    .Weight = Excel.XlBorderWeight.xlThick
                End With
                With SummarySheetRange.Borders(Excel.XlBordersIndex.xlEdgeTop)
                    .LineStyle = Excel.XlLineStyle.xlContinuous
                    .ColorIndex = 0
                    .TintAndShade = 0
                    .Weight = Excel.XlBorderWeight.xlThick
                End With
                With SummarySheetRange.Borders(Excel.XlBordersIndex.xlEdgeBottom)
                    .LineStyle = Excel.XlLineStyle.xlContinuous
                    .ColorIndex = 0
                    .TintAndShade = 0
                    .Weight = Excel.XlBorderWeight.xlThick
                End With
                With SummarySheetRange.Borders(Excel.XlBordersIndex.xlEdgeRight)
                    .LineStyle = Excel.XlLineStyle.xlContinuous
                    .ColorIndex = 0
                    .TintAndShade = 0
                    .Weight = Excel.XlBorderWeight.xlThick
                End With

                iSummarySheetLocation = iSummarySheetCounter + 28

                If sInvestigatorSessions <> vbNullString Then
                    asInvestigatorTimeSessions = Split(sInvestigatorSessions, ";")
                    For iCounter = 0 To UBound(asInvestigatorTimeSessions) - 1
                        asInvestigatorTimeSessionDetails = Split(asInvestigatorTimeSessions(iCounter), "--")
                        xlSummarySheet.Cells(iSummarySheetLocation + iCounter, 2) = asInvestigatorTimeSessionDetails(0)
                        xlSummarySheet.Cells(iSummarySheetLocation + iCounter, 3) = asInvestigatorTimeSessionDetails(1).Replace("Open:", "")
                        sCloseSession = asInvestigatorTimeSessionDetails(2).Replace("Close:", "")
                        If sCloseSession.Contains("(") Then
                            sCloseSession = sCloseSession.Substring(0, sCloseSession.IndexOf("("))
                            xlSummarySheet.Cells(iSummarySheetLocation + iCounter, 4) = sCloseSession
                            asSessionDuration = Split(asInvestigatorTimeSessionDetails(2), "(")
                            xlSummarySheet.Cells(iSummarySheetLocation + iCounter, 5) = asSessionDuration(1).Replace(")", "")
                        End If
                        iSessionCount = iSessionCount + 1
                    Next
                End If
                iSummarySheetCounter = iSummarySheetLocation

                SummarySheetRange = xlSummarySheet.Range("A" & iSummarySheetCounter)
                With SummarySheetRange.Borders(Excel.XlBordersIndex.xlEdgeLeft)
                    .LineStyle = Excel.XlLineStyle.xlContinuous
                    .ColorIndex = 0
                    .TintAndShade = 0
                    .Weight = Excel.XlBorderWeight.xlThick
                End With
                With SummarySheetRange.Borders(Excel.XlBordersIndex.xlEdgeTop)
                    .LineStyle = Excel.XlLineStyle.xlContinuous
                    .ColorIndex = 0
                    .TintAndShade = 0
                    .Weight = Excel.XlBorderWeight.xlThick
                End With
                With SummarySheetRange.Borders(Excel.XlBordersIndex.xlEdgeBottom)
                    .LineStyle = Excel.XlLineStyle.xlContinuous
                    .ColorIndex = 0
                    .TintAndShade = 0
                    .Weight = Excel.XlBorderWeight.xlThick
                End With
                With SummarySheetRange.Borders(Excel.XlBordersIndex.xlEdgeRight)
                    .LineStyle = Excel.XlLineStyle.xlContinuous
                    .ColorIndex = 0
                    .TintAndShade = 0
                    .Weight = Excel.XlBorderWeight.xlThick
                End With
                SummarySheetRange = xlSummarySheet.Range("B" & iSummarySheetCounter)
                With SummarySheetRange.Borders(Excel.XlBordersIndex.xlEdgeLeft)
                    .LineStyle = Excel.XlLineStyle.xlContinuous
                    .ColorIndex = 0
                    .TintAndShade = 0
                    .Weight = Excel.XlBorderWeight.xlThick
                End With
                With SummarySheetRange.Borders(Excel.XlBordersIndex.xlEdgeTop)
                    .LineStyle = Excel.XlLineStyle.xlContinuous
                    .ColorIndex = 0
                    .TintAndShade = 0
                    .Weight = Excel.XlBorderWeight.xlThick
                End With
                With SummarySheetRange.Borders(Excel.XlBordersIndex.xlEdgeBottom)
                    .LineStyle = Excel.XlLineStyle.xlContinuous
                    .ColorIndex = 0
                    .TintAndShade = 0
                    .Weight = Excel.XlBorderWeight.xlThick
                End With
                With SummarySheetRange.Borders(Excel.XlBordersIndex.xlEdgeRight)
                    .LineStyle = Excel.XlLineStyle.xlContinuous
                    .ColorIndex = 0
                    .TintAndShade = 0
                    .Weight = Excel.XlBorderWeight.xlThick
                End With
                SummarySheetRange = xlSummarySheet.Range("C" & iSummarySheetCounter)
                With SummarySheetRange.Borders(Excel.XlBordersIndex.xlEdgeLeft)
                    .LineStyle = Excel.XlLineStyle.xlContinuous
                    .ColorIndex = 0
                    .TintAndShade = 0
                    .Weight = Excel.XlBorderWeight.xlThick
                End With
                With SummarySheetRange.Borders(Excel.XlBordersIndex.xlEdgeTop)
                    .LineStyle = Excel.XlLineStyle.xlContinuous
                    .ColorIndex = 0
                    .TintAndShade = 0
                    .Weight = Excel.XlBorderWeight.xlThick
                End With
                With SummarySheetRange.Borders(Excel.XlBordersIndex.xlEdgeBottom)
                    .LineStyle = Excel.XlLineStyle.xlContinuous
                    .ColorIndex = 0
                    .TintAndShade = 0
                    .Weight = Excel.XlBorderWeight.xlThick
                End With
                With SummarySheetRange.Borders(Excel.XlBordersIndex.xlEdgeRight)
                    .LineStyle = Excel.XlLineStyle.xlContinuous
                    .ColorIndex = 0
                    .TintAndShade = 0
                    .Weight = Excel.XlBorderWeight.xlThick
                End With
                SummarySheetRange = xlSummarySheet.Range("D" & iSummarySheetCounter)
                With SummarySheetRange.Borders(Excel.XlBordersIndex.xlEdgeLeft)
                    .LineStyle = Excel.XlLineStyle.xlContinuous
                    .ColorIndex = 0
                    .TintAndShade = 0
                    .Weight = Excel.XlBorderWeight.xlThick
                End With
                With SummarySheetRange.Borders(Excel.XlBordersIndex.xlEdgeTop)
                    .LineStyle = Excel.XlLineStyle.xlContinuous
                    .ColorIndex = 0
                    .TintAndShade = 0
                    .Weight = Excel.XlBorderWeight.xlThick
                End With
                With SummarySheetRange.Borders(Excel.XlBordersIndex.xlEdgeBottom)
                    .LineStyle = Excel.XlLineStyle.xlContinuous
                    .ColorIndex = 0
                    .TintAndShade = 0
                    .Weight = Excel.XlBorderWeight.xlThick
                End With
                With SummarySheetRange.Borders(Excel.XlBordersIndex.xlEdgeRight)
                    .LineStyle = Excel.XlLineStyle.xlContinuous
                    .ColorIndex = 0
                    .TintAndShade = 0
                    .Weight = Excel.XlBorderWeight.xlThick
                End With
                SummarySheetRange = xlSummarySheet.Range("E" & iSummarySheetCounter)
                With SummarySheetRange.Borders(Excel.XlBordersIndex.xlEdgeLeft)
                    .LineStyle = Excel.XlLineStyle.xlContinuous
                    .ColorIndex = 0
                    .TintAndShade = 0
                    .Weight = Excel.XlBorderWeight.xlThick
                End With
                With SummarySheetRange.Borders(Excel.XlBordersIndex.xlEdgeTop)
                    .LineStyle = Excel.XlLineStyle.xlContinuous
                    .ColorIndex = 0
                    .TintAndShade = 0
                    .Weight = Excel.XlBorderWeight.xlThick
                End With
                With SummarySheetRange.Borders(Excel.XlBordersIndex.xlEdgeBottom)
                    .LineStyle = Excel.XlLineStyle.xlContinuous
                    .ColorIndex = 0
                    .TintAndShade = 0
                    .Weight = Excel.XlBorderWeight.xlThick
                End With
                With SummarySheetRange.Borders(Excel.XlBordersIndex.xlEdgeRight)
                    .LineStyle = Excel.XlLineStyle.xlContinuous
                    .ColorIndex = 0
                    .TintAndShade = 0
                    .Weight = Excel.XlBorderWeight.xlThick
                End With
                iSummarySheetCounter = iSummarySheetLocation
                SummarySheetRange = xlSummarySheet.Range("A" & iSummarySheetCounter & ":E" & iSummarySheetCounter + iSessionCount)
                With SummarySheetRange.Borders(Excel.XlBordersIndex.xlEdgeLeft)
                    .LineStyle = Excel.XlLineStyle.xlContinuous
                    .ColorIndex = 0
                    .TintAndShade = 0
                    .Weight = Excel.XlBorderWeight.xlThick
                End With
                With SummarySheetRange.Borders(Excel.XlBordersIndex.xlEdgeTop)
                    .LineStyle = Excel.XlLineStyle.xlContinuous
                    .ColorIndex = 0
                    .TintAndShade = 0
                    .Weight = Excel.XlBorderWeight.xlThick
                End With
                With SummarySheetRange.Borders(Excel.XlBordersIndex.xlEdgeBottom)
                    .LineStyle = Excel.XlLineStyle.xlContinuous
                    .ColorIndex = 0
                    .TintAndShade = 0
                    .Weight = Excel.XlBorderWeight.xlThick
                End With
                With SummarySheetRange.Borders(Excel.XlBordersIndex.xlEdgeRight)
                    .LineStyle = Excel.XlLineStyle.xlContinuous
                    .ColorIndex = 0
                    .TintAndShade = 0
                    .Weight = Excel.XlBorderWeight.xlThick
                End With

                iSummarySheetCounter = iSummarySheetLocation + iSessionCount

                Dim iTotalCustomMetadata As Integer
                Dim iCustomMetadata As Integer

                xlSummarySheet.Cells(iSummarySheetCounter + 1, 1) = "Custom Metadata Key"
                xlSummarySheet.Cells(iSummarySheetCounter + 1, 2) = "Custom Metadata Value"
                iSummarySheetCounter = iSummarySheetCounter + 1

                asCustomMetadata = Split(sEvidenceCustomMetadata, ";")
                iTotalCustomMetadata = UBound(asCustomMetadata)
                For iCounter = 0 To UBound(asCustomMetadata)
                    iCustomMetadata = iCustomMetadata + 1
                    asCustomMetadataDetails = Split(asCustomMetadata(iCounter), "::")
                    If asCustomMetadataDetails(0) <> vbNullString Then
                        xlSummarySheet.Cells(iSummarySheetCounter + 1, 1) = asCustomMetadataDetails(0)
                        xlSummarySheet.Cells(iSummarySheetCounter + 1, 2) = asCustomMetadataDetails(1)
                        iSummarySheetCounter = iSummarySheetCounter + 1
                    End If
                Next

                xlSummarySheet.Columns("A").Autofit()
                xlSummarySheet.Columns("B").Autofit()
                xlSummarySheet.Columns("C").Autofit()
                xlSummarySheet.Columns("D").Autofit()
                xlSummarySheet.Columns("E").Autofit()
                'xlSummarySheet.Range("A1").Select()

                Dim asTotalsCount() As String
                Dim asTotalsCountDetails() As String
                Dim asOriginalsCount() As String
                Dim asOriginalsCountDetails() As String
                Dim asDuplicatesCount() As String
                Dim asDuplicatesCountDetails() As String
                Dim iCullingSheetStartRow As Integer
                Dim iTotalDuplicates As Integer

                'xlCulling.Shapes.AddPicture(sImageFile, False, True, 0, 0, 100, 60)
                xlCulling.Cells(1, 1) = "Type"
                xlCulling.Cells(1, 2) = "Total Item Count"
                xlCulling.Cells(1, 3) = "Total Item Size"
                xlCulling.Cells(1, 4) = "Original Item Count"
                xlCulling.Cells(1, 5) = "Original Item Size"
                xlCulling.Cells(1, 6) = "Duplicate Item Count"
                xlCulling.Cells(1, 7) = "Duplicate Item Size"
                xlCulling.Cells(1, 8) = "Percentage"
                xlCulling.Range("A1:H1").Font.Name = "Arial Black"

                If sItemCounts <> vbNullString Then
                    asDuplicatesCount = Split(sDuplicateItems, ";")
                    asOriginalsCount = Split(sOriginalItems, ";")
                    asTotalsCount = Split(sItemCounts, ";")

                    iCullingSheetStartRow = 2
                    iTotalDuplicates = UBound(asDuplicatesCount)
                    For iDuplicatesCount = 0 To UBound(asDuplicatesCount)

                        asDuplicatesCountDetails = Split(asDuplicatesCount(iDuplicatesCount), "::")
                        asTotalsCountDetails = Split(asTotalsCount(iDuplicatesCount), "::")
                        asTotalsCountDetails = Split(asTotalsCount(iDuplicatesCount), "::")
                        asOriginalsCountDetails = Split(asOriginalsCount(iDuplicatesCount), "::")
                        If lstTotalItem.Contains(asTotalsCountDetails(0)) Then
                            iTotalItemIndex = lstTotalItem.IndexOf(asTotalsCountDetails(0))
                            dblTotalItemCount = lstTotalItemCount(iTotalItemIndex)
                            dblOriginalItemCount = lstOriginalItemCount(iTotalItemIndex)
                            dblDuplicateItemCount = lstDuplicateItemCount(iTotalItemIndex)
                            dblTotalItemSize = lstTotalItemSize(iTotalItemIndex)
                            dblOriginalItemSize = lstOriginalItemSize(iTotalItemIndex)
                            dblDuplidateItemSize = lstDuplicateItemSize(iTotalItemIndex)

                            lstTotalItemCount(iTotalItemIndex) = dblTotalItemCount + asTotalsCountDetails(1)
                            lstTotalItemSize(iTotalItemIndex) = dblTotalItemSize + asTotalsCountDetails(2)

                            lstOriginalItemCount(iTotalItemIndex) = dblOriginalItemCount + asOriginalsCountDetails(1)
                            lstOriginalItemSize(iTotalItemIndex) = dblOriginalItemSize + asOriginalsCountDetails(2)
                            lstDuplicateItemCount(iTotalItemIndex) = dblDuplicateItemCount + asDuplicatesCountDetails(1)
                            lstDuplicateItemSize(iTotalItemIndex) = dblDuplidateItemSize + asDuplicatesCountDetails(2)
                        Else
                            lstTotalItem.Add(asTotalsCountDetails(0))
                            lstOriginalItem.Add(asOriginalsCountDetails(0))
                            lstDuplicateItem.Add(asDuplicatesCountDetails(0))
                            lstTotalItemCount.Add(asTotalsCountDetails(1))
                            lstOriginalItemCount.Add(asOriginalsCountDetails(1))
                            lstDuplicateItemCount.Add(asDuplicatesCountDetails(1))

                            lstTotalItemSize.Add(asTotalsCountDetails(2))
                            lstOriginalItemSize.Add(asOriginalsCountDetails(2))
                            lstDuplicateItemSize.Add(asDuplicatesCountDetails(2))
                        End If

                        Select Case sShowSizeIN
                            Case "Bytes"
                                xlCulling.Cells(iDuplicatesCount + iCullingSheetStartRow, 3) = asTotalsCountDetails(2)
                                xlCulling.Cells(iDuplicatesCount + iCullingSheetStartRow, 5) = asOriginalsCountDetails(2)
                                xlCulling.Cells(iDuplicatesCount + iCullingSheetStartRow, 7) = asDuplicatesCountDetails(2)
                            Case "Megabytes"
                                xlCulling.Cells(iDuplicatesCount + iCullingSheetStartRow, 3) = Math.Round((CDbl(asTotalsCountDetails(2)) / 1024 / 1024), 2)
                                xlCulling.Cells(iDuplicatesCount + iCullingSheetStartRow, 5) = Math.Round((CDbl(asOriginalsCountDetails(2)) / 1024 / 1024), 2)
                                xlCulling.Cells(iDuplicatesCount + iCullingSheetStartRow, 7) = Math.Round((CDbl(asDuplicatesCountDetails(2)) / 1024 / 1024), 2)
                            Case "Gigabytes"
                                xlCulling.Cells(iDuplicatesCount + iCullingSheetStartRow, 3) = Math.Round((CDbl(asTotalsCountDetails(2)) / 1024 / 1024 / 1024), 2)
                                xlCulling.Cells(iDuplicatesCount + iCullingSheetStartRow, 5) = Math.Round((CDbl(asOriginalsCountDetails(2)) / 1024 / 1024 / 1024), 2)
                                xlCulling.Cells(iDuplicatesCount + iCullingSheetStartRow, 7) = Math.Round((CDbl(asDuplicatesCountDetails(2)) / 1024 / 1024 / 1024), 2)
                        End Select
                        xlNumberRange = xlCulling.Cells(iDuplicatesCount + iCullingSheetStartRow, 3)
                        xlNumberRange.NumberFormat = "#,##0"
                        xlNumberRange = xlCulling.Cells(iDuplicatesCount + iCullingSheetStartRow, 5)
                        xlNumberRange.NumberFormat = "#,##0"
                        xlNumberRange = xlCulling.Cells(iDuplicatesCount + iCullingSheetStartRow, 7)
                        xlNumberRange.NumberFormat = "#,##0"

                        xlCulling.Cells(iDuplicatesCount + iCullingSheetStartRow, 1) = asTotalsCountDetails(0)
                        xlNumberRange = xlCulling.Cells(iDuplicatesCount + iCullingSheetStartRow, 1)
                        xlNumberRange.NumberFormat = "#,##0"

                        xlCulling.Cells(iDuplicatesCount + iCullingSheetStartRow, 2) = asTotalsCountDetails(1)
                        xlNumberRange = xlCulling.Cells(iDuplicatesCount + iCullingSheetStartRow, 2)
                        xlNumberRange.NumberFormat = "#,##0"

                        xlCulling.Cells(iDuplicatesCount + iCullingSheetStartRow, 4) = asOriginalsCountDetails(1)
                        xlNumberRange = xlCulling.Cells(iDuplicatesCount + iCullingSheetStartRow, 4)
                        xlNumberRange.NumberFormat = "#,##0"

                        xlCulling.Cells(iDuplicatesCount + iCullingSheetStartRow, 6) = asDuplicatesCountDetails(1)
                        xlNumberRange = xlCulling.Cells(iDuplicatesCount + iCullingSheetStartRow, 6)
                        xlNumberRange.NumberFormat = "#,##0"

                        If asDuplicatesCountDetails(1) > 0 Then
                            xlCulling.Cells(iDuplicatesCount + iCullingSheetStartRow, 8) = "=" & asDuplicatesCountDetails(1) & "/" & asTotalsCountDetails(1)

                        Else
                            xlCulling.Cells(iDuplicatesCount + iCullingSheetStartRow, 8) = "0.0"
                        End If
                        xlPercentRange = xlCulling.Cells(iDuplicatesCount + iCullingSheetStartRow, 8)
                        xlPercentRange.NumberFormat = "0.00%"
                    Next
                End If

                xlCulling.Columns("A").autofit()
                xlCulling.Columns("B").autofit()
                xlCulling.Columns("C").autofit()
                xlCulling.Columns("D").autofit()
                xlCulling.Columns("E").autofit()
                xlCulling.Columns("F").autofit()
                xlCulling.Columns("G").autofit()
                xlCulling.Columns("H").autofit()
                '                xlCulling.Range("A1").Cells.Select()

                xlDateRangeSheet.Cells(1, 1) = "Item Type"
                xlDateRangeSheet.Cells(1, 2) = "ItemDate"
                xlDateRangeSheet.Cells(1, 3) = "ItemCount"
                xlDateRangeSheet.Cells(1, 4) = "Custodian"
                xlDateRangeSheet.Cells(1, 5) = "Weekday"
                xlDateRangeSheet.Cells(1, 6) = "Year"
                xlDateRangeSheet.Cells(1, 7) = "Month"
                xlDateRangeSheet.Cells(1, 8) = "Month Name"
                xlDateRangeSheet.Cells(1, 9) = "Week"

                dt = Nothing
                ds = New DataSet
                sqlConnection = New SQLiteConnection("Data Source=" & sReportingDBLocation & "\NuixCaseReports.db3;Version=3;Read Only=True;New=False;Compress=True;")

                mSQL = "select CaseGUID, ItemType, ItemDate, ItemCount, Custodian from UCRTDateRange where CaseGUID = '" & sCaseGuid & "' and ItemCount > 0 order by ItemDate ASC"
                sqlCommand = New SQLiteCommand(mSQL, sqlConnection)
                sqlConnection.Open()

                '                Dim AllDataDataSet As New DataSet
                'Dim AllDataDataAdapter As SQLiteDataAdapter

                'AllDataDataAdapter = New SQLiteDataAdapter(mSQL, sqlConnection)
                'AllDataDataAdapter.Fill(AllDataDataSet)
                '
                'xlDateRangeSheet.Range("B3:F3").cop()
                dataReader = sqlCommand.ExecuteReader
                Dim iDateRangeCounter As Integer
                iDateRangeCounter = 0
                While dataReader.Read
                    iDateRangeCounter = iDateRangeCounter + 1
                    xlDateRangeSheet.Cells(iDateRangeCounter + 1, 1) = dataReader.GetString(1)
                    sItemDate = dataReader.GetString(2)
                    dItemDate = Convert.ToDateTime(sItemDate)
                    'dItemDate = Date.ParseExact(sItemDate, "YYYY-MM-DD", CultureInfo.InvariantCulture)
                    'dItemDate = Format(sItemDate, "YYYYMMDD")
                    sMonth = MonthName(Month(dItemDate))
                    sYear = Year(dItemDate)
                    sWeekday = WeekdayName(Weekday(dItemDate))
                    sWeekNumber = DatePart("ww", dItemDate)
                    xlDateRangeSheet.Cells(iDateRangeCounter + 1, 2) = dItemDate
                    xlDateRangeSheet.Cells(iDateRangeCounter + 1, 3) = dataReader.GetInt32(3)
                    xlDateRangeSheet.Cells(iDateRangeCounter + 1, 4) = dataReader.GetString(4)
                    xlDateRangeSheet.Cells(iDateRangeCounter + 1, 5) = sWeekday
                    xlDateRangeSheet.Cells(iDateRangeCounter + 1, 6) = sYear
                    xlDateRangeSheet.Cells(iDateRangeCounter + 1, 7) = Month(dItemDate)
                    xlDateRangeSheet.Cells(iDateRangeCounter + 1, 8) = sMonth
                    xlDateRangeSheet.Cells(iDateRangeCounter + 1, 9) = sWeekNumber

                End While
                If iDateRangeCounter > 1 Then
                    Dim xlDateRangeRange As Excel.Range
                    xlDateRangeRange = xlDateRangeSheet.Range("A1:C" & iDateRangeCounter + 1)

                    xlDateRangeSheet.Columns("A").autofit()
                    xlDateRangeSheet.Columns("B").autofit()
                    xlDateRangeSheet.Columns("C").autofit()

                    sqlConnection.Close()

                    Dim xlPivotTableRange As Excel.Range
                    Dim sPivotTableName As String
                    sPivotTableName = "DateRangePivotTable"
                    xlPivotTableRange = xlDateRangeGapPivot.Range("A1")
                    xlDateRangeGapPivot.PivotTableWizard(Excel.XlPivotTableSourceType.xlDatabase, xlDateRangeRange, xlPivotTableRange, xlPivotTableRange)

                    'Dim xlPivotTablePosition As Excel.Range
                    'xlPivotTableRange = xlDateRangeGapPivot.Range("A1")
                    Dim xlDateRangePivot As Excel.PivotTable
                    xlDateRangePivot = xlDateRangeGapPivot.PivotTables(1)
                    With xlDateRangePivot.PivotFields("ItemDate")
                        .orientation = Excel.XlPivotFieldOrientation.xlRowField
                        .position = 1
                    End With
                    With xlDateRangePivot.PivotFields("Item Type")
                        .orientation = Excel.XlPivotFieldOrientation.xlColumnField
                        .position = 1
                    End With
                    xlDateRangePivot.AddDataField(xlDateRangePivot.PivotFields("ItemCount"), "Item Count", Excel.XlConsolidationFunction.xlCount)
                    With xlDateRangePivot.PivotFields("Item Count")
                        .caption = "Sum of ItemCount"
                        .function = Excel.XlConsolidationFunction.xlSum
                    End With

                    If CInt(sExcelVersion) >= 15 Then
                        Dim xlDateRangePivotChart As Object
                        Dim xlDateRangePivotChartType As Excel.XlChartType
                        xlDateRangePivotChartType = Excel.XlChartType.xlColumnClustered
                        Dim xlDateRangePivotRange As Excel.Range
                        xlDateRangePivotRange = xlDateRangePivot.TableRange2
                        xlDateRangePivotChart = xlDateRangeGapPivot.Shapes.AddChart2(201, xlDateRangePivotChartType)

                        xlDateRangePivotChart.chart.ChartType = xlDateRangePivotChartType
                        xlDateRangePivotChart.Chart.SetSourceData(xlDateRangePivotRange)
                    End If

                End If
                xlDateRangeSheet.Range("A1:I1").Font.Name = "Arial Black"
                xlDateRangeSheet.Columns("A").Autofit()
                xlDateRangeSheet.Columns("B").Autofit()
                xlDateRangeSheet.Columns("C").Autofit()
                xlDateRangeSheet.Columns("D").Autofit()
                xlDateRangeSheet.Columns("E").Autofit()
                xlDateRangeSheet.Columns("F").Autofit()
                xlDateRangeSheet.Columns("G").Autofit()
                xlDateRangeSheet.Columns("H").Autofit()
                xlDateRangeSheet.Columns("I").Autofit()
                'xlDateRangeSheet.Range("A1").Select()

                xlNoDataDateRangeSheet.Cells(1, 1) = "Item Type"
                xlNoDataDateRangeSheet.Cells(1, 2) = "ItemDate"
                xlNoDataDateRangeSheet.Cells(1, 3) = "ItemCount"
                xlNoDataDateRangeSheet.Cells(1, 4) = "Weekday"
                xlNoDataDateRangeSheet.Cells(1, 5) = "Year"
                xlNoDataDateRangeSheet.Cells(1, 6) = "Month"
                xlNoDataDateRangeSheet.Cells(1, 7) = "Month(MMM)"
                xlNoDataDateRangeSheet.Cells(1, 8) = "Week"

                dt = Nothing
                ds = New DataSet
                sqlConnection = New SQLiteConnection("Data Source=" & sReportingDBLocation & "\NuixCaseReports.db3;Version=3;Read Only=True;New=False;Compress=True;")

                mSQL = "select CaseGUID, ItemType, ItemDate, ItemCount from UCRTDateRange where CaseGUID = '" & sCaseGuid & "' and ItemCount = 0 order by ItemDate ASC"
                sqlCommand = New SQLiteCommand(mSQL, sqlConnection)
                sqlConnection.Open()

                dataReader = sqlCommand.ExecuteReader
                Dim iNoDataDateRangeCounter As Integer
                Dim NoDataDataSet As New DataSet
                Dim NoDataDataAdapter As SQLiteDataAdapter

                NoDataDataAdapter = New SQLiteDataAdapter(mSQL, sqlConnection)
                NoDataDataAdapter.Fill(NoDataDataSet)

                iNoDataDateRangeCounter = 0
                While dataReader.Read
                    iNoDataDateRangeCounter = iNoDataDateRangeCounter + 1
                    xlNoDataDateRangeSheet.Cells(iNoDataDateRangeCounter + 1, 1) = dataReader.GetString(1)
                    sItemDate = dataReader.GetString(2)
                    dItemDate = Convert.ToDateTime(sItemDate)
                    'dItemDate = Date.ParseExact(sItemDate, "YYYY-MM-DD", CultureInfo.InvariantCulture)
                    'dItemDate = Format(sItemDate, "YYYYMMDD")
                    sMonth = MonthName(Month(dItemDate))
                    sYear = Year(dItemDate)
                    sWeekday = WeekdayName(Weekday(dItemDate))
                    sWeekNumber = DatePart("ww", dItemDate)
                    xlNoDataDateRangeSheet.Cells(iNoDataDateRangeCounter + 1, 2) = dItemDate
                    xlNoDataDateRangeSheet.Cells(iNoDataDateRangeCounter + 1, 3) = dataReader.GetInt32(3)
                    xlNoDataDateRangeSheet.Cells(iNoDataDateRangeCounter + 1, 4) = sWeekday
                    xlNoDataDateRangeSheet.Cells(iNoDataDateRangeCounter + 1, 5) = sYear
                    xlNoDataDateRangeSheet.Cells(iNoDataDateRangeCounter + 1, 6) = Month(dItemDate)
                    xlNoDataDateRangeSheet.Cells(iNoDataDateRangeCounter + 1, 7) = sMonth
                    xlNoDataDateRangeSheet.Cells(iNoDataDateRangeCounter + 1, 8) = sWeekNumber

                End While
                'If iNoDataDateRangeCounter > 1 Then
                '    Dim xlDateRangeRange As Excel.Range
                '    xlDateRangeRange = xlDateRangeSheet.Range("A1:C" & iNoDataDateRangeCounter + 1)

                '    xlDateRangeSheet.Columns("A").autofit()
                '    xlDateRangeSheet.Columns("B").autofit()
                '    xlDateRangeSheet.Columns("C").autofit()

                '    sqlConnection.Close()

                '    Dim xlPivotTableRange As Excel.Range
                '    Dim sPivotTableName As String
                '    sPivotTableName = "DateRangePivotTable"
                '    xlPivotTableRange = xlDateRangeGapPivot.Range("A1")
                '    xlDateRangeGapPivot.PivotTableWizard(Excel.XlPivotTableSourceType.xlDatabase, xlDateRangeRange, xlPivotTableRange, xlPivotTableRange)

                '    'Dim xlPivotTablePosition As Excel.Range
                '    'xlPivotTableRange = xlDateRangeGapPivot.Range("A1")
                '    Dim xlDateRangePivot As Excel.PivotTable
                '    xlDateRangePivot = xlDateRangeGapPivot.PivotTables(1)
                '    With xlDateRangePivot.PivotFields("ItemDate")
                '        .orientation = Excel.XlPivotFieldOrientation.xlRowField
                '        .position = 1
                '    End With
                '    With xlDateRangePivot.PivotFields("Item Type")
                '        .orientation = Excel.XlPivotFieldOrientation.xlColumnField
                '        .position = 1
                '    End With
                '    xlDateRangePivot.AddDataField(xlDateRangePivot.PivotFields("ItemCount"), "Item Count", Excel.XlConsolidationFunction.xlCount)
                '    With xlDateRangePivot.PivotFields("Item Count")
                '        .caption = "Sum of ItemCount"
                '        .function = Excel.XlConsolidationFunction.xlSum
                '    End With

                '    If CInt(sExcelVersion) >= 15 Then
                '        Dim xlDateRangePivotChart As Object
                '        Dim xlDateRangePivotChartType As Excel.XlChartType
                '        xlDateRangePivotChartType = Excel.XlChartType.xlColumnClustered
                '        Dim xlDateRangePivotRange As Excel.Range
                '        xlDateRangePivotRange = xlDateRangePivot.TableRange2
                '        xlDateRangePivotChart = xlDateRangeGapPivot.Shapes.AddChart2(201, xlDateRangePivotChartType)

                '        xlDateRangePivotChart.chart.ChartType = xlDateRangePivotChartType
                '        xlDateRangePivotChart.Chart.SetSourceData(xlDateRangePivotRange)
                '    End If

                'End If
                xlNoDataDateRangeSheet.Range("A1:H1").Font.Name = "Arial Black"
                xlNoDataDateRangeSheet.Columns("A").Autofit()
                xlNoDataDateRangeSheet.Columns("B").Autofit()
                xlNoDataDateRangeSheet.Columns("C").Autofit()
                xlNoDataDateRangeSheet.Columns("D").Autofit()
                xlNoDataDateRangeSheet.Columns("E").Autofit()
                xlNoDataDateRangeSheet.Columns("F").Autofit()
                xlNoDataDateRangeSheet.Columns("G").Autofit()
                xlNoDataDateRangeSheet.Columns("H").Autofit()
                xlNoDataDateRangeSheet.Columns("I").Autofit()
                'xlDateRangeSheet.Range("A1").Select()


                Dim iLanguageItemCount As Integer
                iLanguageItemCount = 0
                Dim dblLanguageItemSize As Double
                dblLanguageItemSize = 0
                Dim iTotalLanguages As Integer
                Dim sLanguagueAbbreviation As String
                Dim sLanguageName As String

                xlLanguagesSheet.Cells(1, 1) = "Language"
                xlLanguagesSheet.Cells(1, 2) = "Count"
                xlLanguagesSheet.Cells(1, 3) = "Size"
                xlLanguagesSheet.Cells(1, 4) = "Percent"
                If sLanguagesContained <> vbNullString Then
                    asLanguages = Split(sLanguagesContained, ";")
                    iLanguageCount = 1
                    iTotalLanguages = UBound(asLanguages)
                    For iCounter = 0 To UBound(asLanguages)
                        iLanguageCount = iLanguageCount + 1
                        asLanguageDetails = Split(asLanguages(iCounter), "::")
                        If asLanguageDetails(0) <> vbNullString Then
                            sLanguagueAbbreviation = asLanguageDetails(0)
                            Select Case sLanguagueAbbreviation
                                Case "afr"
                                    sLanguageName = "Afrikaans"
                                Case "ara"
                                    sLanguageName = "Arabic"
                                Case "bul"
                                    sLanguageName = "Bulgarian"
                                Case "ben"
                                    sLanguageName = "Bengali"
                                Case "ces"
                                    sLanguageName = "Czech"
                                Case "dan"
                                    sLanguageName = "Danish"
                                Case "deu"
                                    sLanguageName = "German"
                                Case "ell"
                                    sLanguageName = "Greek"
                                Case "eng"
                                    sLanguageName = "English"
                                Case "spa"
                                    sLanguageName = "Spanish"
                                Case "est"
                                    sLanguageName = "Estonian"
                                Case "fas"
                                    sLanguageName = "Persian"
                                Case "fin"
                                    sLanguageName = "Finnish"
                                Case "fra"
                                    sLanguageName = "French"
                                Case "guj"
                                    sLanguageName = "Gujarati"
                                Case "heb"
                                    sLanguageName = "Hebrew"
                                Case "hin"
                                    sLanguageName = "Hindi"
                                Case "hrv"
                                    sLanguageName = "Croatian"
                                Case "hun"
                                    sLanguageName = "Hungarian"
                                Case "ind"
                                    sLanguageName = "Indonesian"
                                Case "ita"
                                    sLanguageName = "Italian"
                                Case "jpn"
                                    sLanguageName = "Japanese"
                                Case "kan"
                                    sLanguageName = "Kannada"
                                Case "kor"
                                    sLanguageName = "Korean"
                                Case "lit"
                                    sLanguageName = "Lithuanian"
                                Case "lav"
                                    sLanguageName = "Latvian"
                                Case "mkd"
                                    sLanguageName = "Macedonian"
                                Case "mal"
                                    sLanguageName = "Malayalam"
                                Case "mar"
                                    sLanguageName = "Marathi"
                                Case "nep"
                                    sLanguageName = "Nepali"
                                Case "nld"
                                    sLanguageName = "Dutch"
                                Case "nor"
                                    sLanguageName = "Norwegian"
                                Case "pan"
                                    sLanguageName = "Punjabi"
                                Case "pol"
                                    sLanguageName = "Polish"
                                Case "por"
                                    sLanguageName = "Portuguese"
                                Case "ron"
                                    sLanguageName = "Romanian"
                                Case "rus"
                                    sLanguageName = "Russian"
                                Case "slk"
                                    sLanguageName = "Slovak"
                                Case "slv"
                                    sLanguageName = "Slovene"
                                Case "som"
                                    sLanguageName = "Somali"
                                Case "sqi"
                                    sLanguageName = "Albanian"
                                Case "swe"
                                    sLanguageName = "Swedish"
                                Case "swa"
                                    sLanguageName = "Swahili"
                                Case "tam"
                                    sLanguageName = "Tamil"
                                Case "tel"
                                    sLanguageName = "Telugu"
                                Case "tha"
                                    sLanguageName = "Thai"
                                Case "tgl"
                                    sLanguageName = "Tagalog"
                                Case "tur"
                                    sLanguageName = "Turkish"
                                Case "ukr"
                                    sLanguageName = "Ukrainian"
                                Case "urd"
                                    sLanguageName = "Urdu"
                                Case "vie"
                                    sLanguageName = "Vietnamese"
                                Case "zho"
                                    sLanguageName = "Chinese"
                                Case Else
                                    sLanguageName = sLanguagueAbbreviation
                            End Select
                            xlLanguagesSheet.Range("A1:D1").Font.Name = "Arial Black"
                            xlLanguagesSheet.Cells(iLanguageCount, 1) = sLanguageName
                            xlLanguagesSheet.Cells(iLanguageCount, 2) = asLanguageDetails(1)
                            xlNumberRange = xlLanguagesSheet.Cells(iLanguageCount, 2)
                            xlNumberRange.NumberFormat = "#,##0"

                            iLanguageItemCount = iLanguageItemCount + CInt(asLanguageDetails(1))

                            If lstLanguages.Contains(sLanguageName) Then
                                iLanguageIndex = lstLanguages.IndexOf(sLanguageName)
                                dblTotalLanguageCount = lstLanguagesCount(iLanguageIndex)
                                dblTotalLanguageSize = lstLanguagesSize(iLanguageIndex)

                                lstLanguagesCount(iLanguageIndex) = dblTotalLanguageCount + asLanguageDetails(1)
                                lstLanguagesSize(iLanguageIndex) = dblTotalLanguageSize + asLanguageDetails(2)
                            Else
                                lstLanguages.Add(sLanguageName)
                                lstLanguagesCount.Add(asLanguageDetails(1))
                                lstLanguagesSize.Add(asLanguageDetails(2))
                            End If

                            Select Case sShowSizeIN
                                Case "Bytes"
                                    xlLanguagesSheet.Cells(iLanguageCount, 3) = Math.Round(CDbl(asLanguageDetails(2)), 2)
                                    dblLanguageItemSize = dblLanguageItemSize + Math.Round(CDbl(asLanguageDetails(2)), 2)
                                Case "Megabytes"
                                    xlLanguagesSheet.Cells(iLanguageCount, 3) = Math.Round((CDbl(asLanguageDetails(2)) / 1024 / 1024), 2)
                                    dblLanguageItemSize = dblLanguageItemSize + Math.Round((CDbl(asLanguageDetails(2)) / 1024 / 1024), 2)
                                Case "Gigabytes"
                                    xlLanguagesSheet.Cells(iLanguageCount, 3) = Math.Round((CDbl(asLanguageDetails(2)) / 1024 / 1024 / 1024), 2)
                                    dblLanguageItemSize = dblLanguageItemSize + Math.Round((CDbl(asLanguageDetails(2)) / 1024 / 1024 / 1024), 2)
                            End Select
                            xlNumberRange = xlLanguagesSheet.Cells(iLanguageCount, 3)
                            xlNumberRange.NumberFormat = "#,##0"

                            xlLanguagesSheet.Cells(iLanguageCount, 4) = "=B" & iLanguageCount & "/" & "B" & iTotalLanguages + 2

                            xlPercentRange = xlLanguagesSheet.Cells(iLanguageCount, 4)
                            xlPercentRange.NumberFormat = "0.00%"
                        End If
                    Next

                End If
                xlLanguagesSheet.Sort.SortFields.Add(xlLanguagesSheet.Range("B2:B" & iLanguageCount), Excel.XlSortOn.xlSortOnValues, Excel.XlSortOrder.xlDescending)
                With xlLanguagesSheet.Sort
                    .SetRange(xlLanguagesSheet.Range("A1:C" & iLanguageCount))
                    .Header = Excel.XlYesNoGuess.xlYes
                    .MatchCase = False
                    .Apply()
                End With
                xlLanguagesSheet.Cells(iLanguageCount, 1) = "Totals"
                xlLanguagesSheet.Cells(iLanguageCount, 2) = iLanguageItemCount
                xlNumberRange = xlLanguagesSheet.Cells(iLanguageCount, 2)
                xlNumberRange.NumberFormat = "#,##0"
                xlLanguagesSheet.Cells(iLanguageCount, 3) = dblLanguageItemSize
                xlNumberRange = xlLanguagesSheet.Cells(iLanguageCount, 3)
                xlNumberRange.NumberFormat = "#,##0"

                xlLanguagesSheet.Columns("A").autofit()
                xlLanguagesSheet.Columns("B").autofit()
                xlLanguagesSheet.Columns("C").autofit()
                'xlLanguagesSheet.Range("A1").Select()

                If CInt(sExcelVersion) >= 15 Then
                    Dim oLanguageChart As Object
                    Dim xlLanguageChartType As Excel.XlChartType = Excel.XlChartType.xl3DPie
                    'xlLanguageShape = xlLanguagesSheet.Shapes.AddChart2(286, xlLanguageChartType)
                    Dim xlLanguageRange As Excel.Range
                    If iLanguageCount > 0 Then
                        xlLanguageRange = xlLanguagesSheet.Range("A1:B" & iLanguageCount - 1)
                    Else
                        xlLanguageRange = xlLanguagesSheet.Range("A1:B1")

                    End If

                    oLanguageChart = xlLanguagesSheet.Shapes.AddChart2(286, xlLanguageChartType)

                    oLanguageChart.Chart.ChartType = xlLanguageChartType
                    oLanguageChart.Chart.ChartTitle.Text = "Languages"
                    oLanguageChart.Chart.SetSourceData(xlLanguageRange)
                End If

                xlMimeTypeSheet.Cells(1, 1) = "Mime Type"
                xlMimeTypeSheet.Cells(1, 2) = "Count"
                xlMimeTypeSheet.Cells(1, 3) = "Size"
                xlMimeTypeSheet.Cells(1, 4) = "Percent"
                xlMimeTypeSheet.Range("A1:D1").Font.Name = "Arial Black"

                Dim iMimeTypeCounter As Integer
                Dim iTotalMimeTypes As Integer
                iMimeTypeCounter = 1
                Dim iMimeTypeItemCount As Integer
                Dim dblMimeTypeSizeCount As Double

                iMimeTypeItemCount = 0
                dblMimeTypeSizeCount = 0
                iMimeTypeCounter = 1
                If sMimeTypes <> vbNullString Then
                    asMimeTypes = Split(sMimeTypes, ";")
                    iTotalMimeTypes = UBound(asMimeTypes)
                    For iCounter = 0 To UBound(asMimeTypes)
                        iMimeTypeCounter = iMimeTypeCounter + 1
                        asMimeTypeDetails = Split(asMimeTypes(iCounter), "::")
                        If asMimeTypeDetails(0) <> vbNullString Then
                            xlMimeTypeSheet.Cells(iMimeTypeCounter, 1) = asMimeTypeDetails(0)
                            xlMimeTypeSheet.Cells(iMimeTypeCounter, 2) = asMimeTypeDetails(1)
                            xlMimeTypeSheet.Cells(iMimeTypeCounter, 3) = asMimeTypeDetails(2)
                            iMimeTypeItemCount = iMimeTypeItemCount + CInt(asMimeTypeDetails(1))

                            If lstMimeType.Contains(asMimeTypeDetails(0)) Then
                                iMimeTypeIndex = lstMimeType.IndexOf(asMimeTypeDetails(0))
                                dblTotalMimeTypeCount = lstMimeTypeCount(iMimeTypeIndex)
                                dblTotalMimeTypeSize = lstMimeTypeSize(iMimeTypeIndex)

                                lstMimeTypeCount(iMimeTypeIndex) = dblTotalMimeTypeCount + asMimeTypeDetails(1)
                                lstMimeTypeSize(iMimeTypeIndex) = dblTotalMimeTypeSize + asMimeTypeDetails(2)
                            Else
                                lstMimeType.Add(asMimeTypeDetails(0))
                                lstMimeTypeCount.Add(asMimeTypeDetails(1))
                                lstMimeTypeSize.Add(asMimeTypeDetails(2))
                            End If
                            Select Case sShowSizeIN
                                Case "Bytes"
                                    xlMimeTypeSheet.Cells(iMimeTypeCounter, 3) = Math.Round(CDbl(asMimeTypeDetails(2)), 2)
                                    dblMimeTypeSizeCount = dblMimeTypeSizeCount + Math.Round(CDbl(asMimeTypeDetails(2)), 2)
                                Case "Megabytes"
                                    xlMimeTypeSheet.Cells(iMimeTypeCounter, 3) = Math.Round((CDbl(asMimeTypeDetails(2)) / 1024 / 1024), 2)
                                    dblMimeTypeSizeCount = dblMimeTypeSizeCount + Math.Round((CDbl(asMimeTypeDetails(2)) / 1024 / 1024), 2)
                                Case "Gigabytes"
                                    xlMimeTypeSheet.Cells(iMimeTypeCounter, 3) = Math.Round((CDbl(asMimeTypeDetails(2)) / 1024 / 1024 / 1024), 2)
                                    dblMimeTypeSizeCount = dblMimeTypeSizeCount + Math.Round((CDbl(asMimeTypeDetails(2)) / 1024 / 1024 / 1024), 2)
                            End Select
                            xlMimeTypeSheet.Cells(iMimeTypeCounter, 4) = "=B" & iMimeTypeCounter & "/" & "B" & iTotalMimeTypes + 2
                            xlPercentRange = xlMimeTypeSheet.Cells(iMimeTypeCounter, 4)
                            xlPercentRange.NumberFormat = "0.00%"
                        End If
                    Next
                End If
                xlMimeTypeSheet.Sort.SortFields.Add(xlMimeTypeSheet.Range("B2:B" & iMimeTypeCounter), Excel.XlSortOn.xlSortOnValues, Excel.XlSortOrder.xlDescending)
                With xlMimeTypeSheet.Sort
                    .SetRange(xlMimeTypeSheet.Range("A1:C" & iMimeTypeCounter))
                    .Header = Excel.XlYesNoGuess.xlYes
                    .MatchCase = False
                    .Apply()
                End With
                xlMimeTypeSheet.Cells(iMimeTypeCounter, 1) = "Totals"
                xlMimeTypeSheet.Cells(iMimeTypeCounter, 2) = iMimeTypeItemCount
                xlMimeTypeSheet.Cells(iMimeTypeCounter, 3) = dblMimeTypeSizeCount
                xlNumberRange = xlMimeTypeSheet.Range("B:C")
                xlNumberRange.NumberFormat = "#,##0"


                If CInt(sExcelVersion >= 15) Then
                    Dim oMimeTypeChart As Object
                    Dim xlMimeTypeChartType As Excel.XlChartType = Excel.XlChartType.xl3DPie
                    'xlMimeTypeShape = xlMimeTypeSheet.Shapes.AddChart2(286, xlMimeTypeChartType)
                    Dim xlMimeTypeRange As Excel.Range
                    If iMimeTypeCounter > 0 Then
                        xlMimeTypeRange = xlMimeTypeSheet.Range("A1:B" & iMimeTypeCounter - 1)
                    Else
                        xlMimeTypeRange = xlMimeTypeSheet.Range("A1:B1")

                    End If

                    oMimeTypeChart = xlMimeTypeSheet.Shapes.AddChart2(286, xlMimeTypeChartType)

                    oMimeTypeChart.Chart.ChartType = xlMimeTypeChartType
                    oMimeTypeChart.Chart.ChartTitle.Text = "Mime Types"
                    oMimeTypeChart.Chart.SetSourceData(xlMimeTypeRange)
                End If

                xlMimeTypeSheet.Columns("A").autofit()
                xlMimeTypeSheet.Columns("B").autofit()
                xlMimeTypeSheet.Columns("C").autofit()
                'xlMimeTypeSheet.Range("A1").Select()


                xlItemTypeSheet.Cells(1, 1) = "Item Type"
                xlItemTypeSheet.Cells(1, 2) = "Count"
                xlItemTypeSheet.Cells(1, 3) = "Size"
                xlItemTypeSheet.Cells(1, 4) = "Percent"
                xlItemTypeSheet.Range("A1:D1").Font.Name = "Arial Black"

                Dim iItemTypeCount As Integer
                Dim iTotalItemTypes As Integer
                Dim iItemTypeItemCount As Integer
                Dim dblItemTypeSizeCount As Double
                iItemTypeItemCount = 0
                dblItemTypeSizeCount = 0
                iItemTypeCount = 1
                If sItemTypes <> vbNullString Then
                    asItemTypes = Split(sItemTypes, ";")
                    iTotalItemTypes = UBound(asItemTypes)
                    For iCounter = 0 To UBound(asItemTypes)
                        iItemTypeCount = iItemTypeCount + 1
                        asItemTypeDetails = Split(asItemTypes(iCounter), "::")
                        If asItemTypeDetails(0) <> vbNullString Then
                            If asItemTypeDetails(0) = "Email" Then
                                sTopLevelEmailCount = asItemTypeDetails(1)
                                xlSummarySheet.Cells(12, 2) = asItemTypeDetails(1)
                                xlNumberRange = xlSummarySheet.Cells(12, 2)
                                xlNumberRange.NumberFormat = "#,##0"
                            End If
                            xlItemTypeSheet.Cells(iItemTypeCount, 1) = asItemTypeDetails(0)
                            xlNumberRange = xlItemTypeSheet.Cells(iItemTypeCount, 1)
                            xlNumberRange.NumberFormat = "#,##0"

                            xlItemTypeSheet.Cells(iItemTypeCount, 2) = asItemTypeDetails(1)
                            xlNumberRange = xlItemTypeSheet.Cells(iItemTypeCount, 2)
                            xlNumberRange.NumberFormat = "#,##0"

                            iItemTypeItemCount = iItemTypeItemCount + CInt(asItemTypeDetails(1))
                            If lstItemType.Contains(asItemTypeDetails(0)) Then
                                iItemTypeIndex = lstItemType.IndexOf(asItemTypeDetails(0))
                                dblTotalItemItemCount = lstItemTypeCount(iItemTypeIndex)
                                dblTotalItemItemSize = lstItemTypeSize(iItemTypeIndex)

                                lstItemTypeCount(iItemTypeIndex) = dblTotalItemItemCount + asItemTypeDetails(1)
                                lstItemTypeSize(iItemTypeIndex) = dblTotalItemItemSize + asItemTypeDetails(2)
                            Else
                                lstItemType.Add(asItemTypeDetails(0))
                                lstItemTypeCount.Add(asItemTypeDetails(1))
                                lstItemTypeSize.Add(asItemTypeDetails(2))
                            End If
                            Select Case sShowSizeIN
                                Case "Bytes"
                                    xlItemTypeSheet.Cells(iItemTypeCount, 3) = Math.Round(CDbl(asItemTypeDetails(2)), 2)
                                    dblItemTypeSizeCount = dblItemTypeSizeCount + Math.Round(CDbl(asItemTypeDetails(2)), 2)
                                Case "Megabytes"
                                    xlItemTypeSheet.Cells(iItemTypeCount, 3) = Math.Round((CDbl(asItemTypeDetails(2)) / 1024 / 1024), 2)
                                    dblItemTypeSizeCount = dblItemTypeSizeCount + Math.Round((CDbl(asItemTypeDetails(2)) / 1024 / 1024), 2)
                                Case "Gigabytes"
                                    xlItemTypeSheet.Cells(iItemTypeCount, 3) = Math.Round((CDbl(asItemTypeDetails(2)) / 1024 / 1024 / 1024), 2)
                                    dblItemTypeSizeCount = dblItemTypeSizeCount + Math.Round((CDbl(asItemTypeDetails(2)) / 1024 / 1024 / 1024), 2)
                            End Select
                            xlItemTypeSheet.Cells(iItemTypeCount, 4) = "=B" & iItemTypeCount & "/" & "B" & iTotalItemTypes + 2
                            xlPercentRange = xlItemTypeSheet.Cells(iItemTypeCount, 4)
                            xlPercentRange.NumberFormat = "0.00%"
                        End If
                    Next
                End If
                xlItemTypeSheet.Sort.SortFields.Add(xlItemTypeSheet.Range("B2:B" & iItemTypeCount), Excel.XlSortOn.xlSortOnValues, Excel.XlSortOrder.xlDescending)
                With xlItemTypeSheet.Sort
                    .SetRange(xlItemTypeSheet.Range("A1:C" & iItemTypeCount))
                    .Header = Excel.XlYesNoGuess.xlYes
                    .MatchCase = False
                    .Apply()
                End With

                xlItemTypeSheet.Cells(iItemTypeCount, 1) = "Totals"
                xlItemTypeSheet.Cells(iItemTypeCount, 2) = iItemTypeItemCount
                xlNumberRange = xlItemTypeSheet.Cells(iItemTypeCount, 2)
                xlNumberRange.NumberFormat = "#,##0"
                xlItemTypeSheet.Cells(iItemTypeCount, 3) = dblItemTypeSizeCount
                xlNumberRange = xlItemTypeSheet.Cells(iItemTypeCount, 3)
                xlNumberRange.NumberFormat = "#,##0"
                xlNumberRange = xlItemTypeSheet.Range("B:C")
                xlNumberRange.NumberFormat = "#,##0"

                If CInt(sExcelVersion) >= 15 Then
                    Dim oItemTypeChart As Object
                    Dim xlItemTypeChartType As Excel.XlChartType = Excel.XlChartType.xl3DPie
                    'xlItemTypeShape = xlItemTypeSheet.Shapes.AddChart2(286, xlItemTypeChartType)
                    Dim xlItemTypeRange As Excel.Range = xlItemTypeSheet.Range("A1:B" & iItemTypeCount)
                    If iItemTypeCount > 0 Then
                        xlItemTypeRange = xlItemTypeSheet.Range("A1:B" & iItemTypeCount - 1)
                    Else
                        xlItemTypeRange = xlItemTypeSheet.Range("A1:B1")

                    End If

                    oItemTypeChart = xlItemTypeSheet.Shapes.AddChart2(286, xlItemTypeChartType)

                    oItemTypeChart.chart.ChartType = xlItemTypeChartType
                    oItemTypeChart.Chart.ChartTitle.Text = "Item Types"
                    oItemTypeChart.Chart.SetSourceData(xlItemTypeRange)
                End If

                xlItemTypeSheet.Columns("A").autofit()
                xlItemTypeSheet.Columns("B").autofit()
                xlItemTypeSheet.Columns("C").autofit()
                'xlItemTypeSheet.Range("A1").Select()

                xlIrregularItems.Cells(1, 1) = "Type"
                xlIrregularItems.Cells(1, 2) = "Count"
                xlIrregularItems.Cells(1, 3) = "Size"
                xlIrregularItems.Cells(1, 4) = "Percent"
                xlEncryptedMessagesSheet.Cells(1, 1) = "Type"
                xlEncryptedMessagesSheet.Cells(1, 2) = "Count"
                xlEncryptedMessagesSheet.Cells(1, 3) = "Size"
                xlEncryptedMessagesSheet.Cells(1, 4) = "Percent"
                xlIrregularItems.Range("A1:D1").Font.Name = "Arial Black"
                xlEncryptedMessagesSheet.Range("A1:D1").Font.Name = "Arial Black"

                Dim iEncryptedItems As Integer
                iEncryptedItems = 1
                Dim iIrregularItems As Integer
                iIrregularItems = 1
                Dim iEncryptedItemCount As Integer
                Dim dblEncrypedItemSize As Double
                Dim iIrregularItemCount As Integer
                Dim dblIrregularItemSize As Double
                iEncryptedItemCount = 0
                dblEncrypedItemSize = 0
                iIrregularItemCount = 0
                dblIrregularItemSize = 0
                If sIrregularItems <> vbNullString Then
                    asIrregularItems = Split(sIrregularItems, ";")
                    Dim iTotalIrregularItems As Integer
                    iTotalIrregularItems = UBound(asIrregularItems)
                    For iCounter = 0 To UBound(asIrregularItems)

                        asIrregularItemsDetails = Split(asIrregularItems(iCounter), "::")
                        If asIrregularItemsDetails(0) <> vbNullString Then

                            If lstIrregularItems.Contains(asIrregularItemsDetails(0)) Then
                                iIrregularItemIndex = lstIrregularItems.IndexOf(asIrregularItemsDetails(0))
                                dblTotalIrregularItemCount = lstIrregularItemsCount(iIrregularItemIndex)
                                dblTotalIrregularItemSize = lstIrregularItemsSize(iIrregularItemIndex)

                                lstIrregularItemsCount(iIrregularItemIndex) = dblTotalIrregularItemCount + asIrregularItemsDetails(1)
                                lstIrregularItemsSize(iIrregularItemIndex) = dblTotalIrregularItemSize + asIrregularItemsDetails(2)
                            Else
                                lstIrregularItems.Add(asIrregularItemsDetails(0))
                                lstIrregularItemsCount.Add(asIrregularItemsDetails(1))
                                lstIrregularItemsSize.Add(asIrregularItemsDetails(2))
                            End If
                            If ((asIrregularItemsDetails(0) = "Encrypted") Or (asIrregularItems(0) = "Decrypted")) Then
                                iEncryptedItems = iEncryptedItems + 1
                                xlEncryptedMessagesSheet.Cells(iEncryptedItems, 1) = asIrregularItemsDetails(0)
                                xlEncryptedMessagesSheet.Cells(iEncryptedItems, 2) = asIrregularItemsDetails(1)
                                xlEncryptedMessagesSheet.Cells(iEncryptedItems, 3) = asIrregularItemsDetails(2)
                                iIrregularItemsCount = iIrregularItemsCount - 1
                                If asIrregularItemsDetails(1) > 0 Then
                                    xlEncryptedMessagesSheet.Cells(iEncryptedItems, 4) = "=B" & iEncryptedItems & "/" & "B" & 4
                                Else
                                    xlEncryptedMessagesSheet.Cells(iEncryptedItems, 4) = "0"
                                End If
                                iEncryptedItemCount = iEncryptedItemCount + CInt(asIrregularItemsDetails(1))
                                dblEncrypedItemSize = dblEncrypedItemSize + CDbl(asIrregularItemsDetails(2))

                                xlPercentRange = xlEncryptedMessagesSheet.Cells(iEncryptedItems + 1, 4)
                                xlPercentRange.NumberFormat = "0.00%"
                            ElseIf (asIrregularItemsDetails(0) = "Decrypted") Then
                                iEncryptedItems = iEncryptedItems + 1
                                xlEncryptedMessagesSheet.Cells(iEncryptedItems, 1) = asIrregularItemsDetails(0)
                                xlEncryptedMessagesSheet.Cells(iEncryptedItems, 2) = asIrregularItemsDetails(1)
                                Select Case sShowSizeIN
                                    Case "Bytes"
                                        xlEncryptedMessagesSheet.Cells(iEncryptedItems, 3) = Math.Round(CDbl(asIrregularItemsDetails(2)), 2)
                                        dblEncrypedItemSize = dblEncrypedItemSize + Math.Round(CDbl(asIrregularItemsDetails(2)), 2)
                                    Case "Megabytes"
                                        xlEncryptedMessagesSheet.Cells(iEncryptedItems, 3) = Math.Round((CDbl(asIrregularItemsDetails(2)) / 1024 / 1024), 2)
                                        dblEncrypedItemSize = dblEncrypedItemSize + Math.Round((CDbl(asIrregularItemsDetails(2)) / 1024 / 1024), 2)
                                    Case "Gigabytes"
                                        xlEncryptedMessagesSheet.Cells(iEncryptedItems, 3) = Math.Round((CDbl(asIrregularItemsDetails(2)) / 1024 / 1024 / 1024), 2)
                                        dblEncrypedItemSize = dblEncrypedItemSize + Math.Round((CDbl(asIrregularItemsDetails(2)) / 1024 / 1024 / 1024), 2)

                                End Select

                                iIrregularItemsCount = iIrregularItemsCount - 1
                                iEncryptedItemCount = iEncryptedItemCount + CInt(asIrregularItemsDetails(1))
                                If asIrregularItemsDetails(1) > 0 Then
                                    xlEncryptedMessagesSheet.Cells(iEncryptedItems, 4) = "=B" & iEncryptedItems & "/" & "B" & 4
                                Else
                                    xlEncryptedMessagesSheet.Cells(iEncryptedItems, 4) = "0"
                                End If

                                xlPercentRange = xlEncryptedMessagesSheet.Cells((iEncryptedItems), 4)
                                xlPercentRange.NumberFormat = "0.00%"
                            Else
                                iIrregularItems = iIrregularItems + 1
                                xlIrregularItems.Cells(iIrregularItems, 1) = asIrregularItemsDetails(0)
                                xlIrregularItems.Cells(iIrregularItems, 2) = asIrregularItemsDetails(1)
                                Select Case sShowSizeIN
                                    Case "Bytes"
                                        xlIrregularItems.Cells(iIrregularItems, 3) = Math.Round(CDbl(asIrregularItemsDetails(2)), 2)
                                        dblIrregularItemSize = dblIrregularItemSize + Math.Round(CDbl(asIrregularItemsDetails(2)), 2)
                                    Case "Megabytes"
                                        xlIrregularItems.Cells(iIrregularItems, 3) = Math.Round((CDbl(asIrregularItemsDetails(2)) / 1024 / 1024), 2)
                                        dblIrregularItemSize = dblIrregularItemSize + Math.Round((CDbl(asIrregularItemsDetails(2)) / 1024 / 1024), 2)
                                    Case "Gigabytes"
                                        xlIrregularItems.Cells(iIrregularItems, 3) = Math.Round((CDbl(asIrregularItemsDetails(2)) / 1024 / 1024 / 1024), 2)
                                        dblIrregularItemSize = dblIrregularItemSize + Math.Round((CDbl(asIrregularItemsDetails(2)) / 1024 / 1024 / 1024), 2)

                                End Select

                                iIrregularItemCount = iIrregularItemCount + CInt(asIrregularItemsDetails(1))
                                xlIrregularItems.Cells(iIrregularItems, 4) = "=B" & iIrregularItems & "/" & "B" & (iTotalIrregularItems - 1)
                                xlPercentRange = xlIrregularItems.Cells(iIrregularItems, 4)
                                xlPercentRange.NumberFormat = "0.00%"
                            End If
                        End If
                    Next

                End If
                xlNumberRange = xlIrregularItems.Range("B:C")
                xlNumberRange.NumberFormat = "#,##0"

                If iEncryptedItemCount > 0 Then
                    xlEncryptedMessagesSheet.Sort.SortFields.Add(xlEncryptedMessagesSheet.Range("B2:B" & iEncryptedItemCount), Excel.XlSortOn.xlSortOnValues, Excel.XlSortOrder.xlDescending)
                    With xlEncryptedMessagesSheet.Sort
                        .SetRange(xlEncryptedMessagesSheet.Range("A1:C" & iEncryptedItemCount))
                        .Header = Excel.XlYesNoGuess.xlYes
                        .MatchCase = False
                        .Apply()
                    End With
                End If

                xlEncryptedMessagesSheet.Cells(4, 1) = "Totals"
                xlEncryptedMessagesSheet.Cells(4, 2) = iEncryptedItemCount
                xlEncryptedMessagesSheet.Cells(4, 3) = dblEncrypedItemSize

                If iIrregularItemCount > 0 Then
                    xlIrregularItems.Sort.SortFields.Add(xlIrregularItems.Range("B2:B" & iIrregularItems), Excel.XlSortOn.xlSortOnValues, Excel.XlSortOrder.xlDescending)
                    With xlIrregularItems.Sort
                        .SetRange(xlIrregularItems.Range("A1:C" & iIrregularItems))
                        .Header = Excel.XlYesNoGuess.xlYes
                        .MatchCase = False
                        .Apply()
                    End With
                End If
                xlIrregularItems.Cells(iIrregularItems, 1) = "Totals"
                xlIrregularItems.Cells(iIrregularItems, 2) = iIrregularItemCount
                xlIrregularItems.Cells(iIrregularItems, 3) = dblIrregularItemSize
                If CInt(sExcelVersion) >= 15 Then
                    Dim oIrregularItemsChart As Object
                    Dim xlIrregularItemsChartType As Excel.XlChartType = Excel.XlChartType.xl3DPie
                    'xlIrregularItemsShape = xlIrregularItems.Shapes.AddChart2(286, xlIrregularItemsChartType)
                    Dim xlIrregularItemsRange As Excel.Range
                    If iIrregularItems > 0 Then
                        xlIrregularItemsRange = xlIrregularItems.Range("A1:B" & iIrregularItems - 1)
                    Else
                        xlIrregularItemsRange = xlIrregularItems.Range("A1:B1")
                    End If

                    oIrregularItemsChart = xlIrregularItems.Shapes.AddChart2(286, xlIrregularItemsChartType)
                    oIrregularItemsChart.chart.ChartType = xlIrregularItemsChartType
                    oIrregularItemsChart.Chart.ChartTitle.Text = "Irregular Items"
                    oIrregularItemsChart.Chart.SetSourceData(xlIrregularItemsRange)
                End If
                xlIrregularItems.Columns("A").autofit()
                xlIrregularItems.Columns("B").autofit()
                xlIrregularItems.Columns("C").autofit()
                'xlIrregularItems.Range("A1").Select()
                'xlEncryptedMessagesSheet.Range("A1").Select()
                xlSummarySheet.Activate()

                ImageRange = xlSummarySheet2.Range("A1:A4")
                ImageRange.MergeCells = True

                xlSummarySheet2.Shapes.AddPicture(sImageFile, False, True, 0, 0, 100, 60)
                xlSummarySheet2.Cells(6, 5) = sCurrentCaseVersion
                xlSummarySheet2.Cells(8, 1) = "Case Information"
                xlHighlightRange = xlSummarySheet2.Range("A8:B8")
                With xlHighlightRange.Borders(Excel.XlBordersIndex.xlEdgeLeft)
                    .LineStyle = Excel.XlLineStyle.xlContinuous
                    .ColorIndex = 0
                    .TintAndShade = 0
                    .Weight = Excel.XlBorderWeight.xlThick
                End With
                With xlHighlightRange.Borders(Excel.XlBordersIndex.xlEdgeTop)
                    .LineStyle = Excel.XlLineStyle.xlContinuous
                    .ColorIndex = 0
                    .TintAndShade = 0
                    .Weight = Excel.XlBorderWeight.xlThick
                End With
                With xlHighlightRange.Borders(Excel.XlBordersIndex.xlEdgeBottom)
                    .LineStyle = Excel.XlLineStyle.xlContinuous
                    .ColorIndex = 0
                    .TintAndShade = 0
                    .Weight = Excel.XlBorderWeight.xlThick
                End With
                With xlHighlightRange.Borders(Excel.XlBordersIndex.xlEdgeRight)
                    .LineStyle = Excel.XlLineStyle.xlContinuous
                    .ColorIndex = 0
                    .TintAndShade = 0
                    .Weight = Excel.XlBorderWeight.xlThick
                End With
                With xlHighlightRange.Interior
                    .ThemeColor = Excel.XlThemeColor.xlThemeColorDark2
                    .TintAndShade = 0
                    .PatternTintAndShade = 0
                End With
                With xlHighlightRange.Interior
                    .ThemeColor = Excel.XlThemeColor.xlThemeColorAccent3
                    .TintAndShade = 0
                End With

                xlSummarySheet2.Cells(9, 1) = "Case Name"
                xlSummarySheet2.Cells(9, 2) = sCaseName
                xlSummarySheet2.Cells(10, 1) = "Oldest Case Item"
                xlSummarySheet2.Cells(10, 2) = sOldestTopLevel
                xlSummarySheet2.Cells(11, 1) = "Newest Case Item"
                xlSummarySheet2.Cells(11, 2) = sNewestTopLevel
                xlSummarySheet2.Cells(12, 1) = "Total Case Item Count"
                xlSummarySheet2.Cells(12, 2) = sTotalCaseItemCount
                xlNumberRange = xlSummarySheet2.Cells(12, 2)
                xlNumberRange.NumberFormat = "#,##0"
                xlSummarySheet2.Cells(13, 1) = "Top Level Email Count"
                xlSummarySheet2.Cells(13, 2) = sTopLevelEmailCount
                xlNumberRange = xlSummarySheet2.Cells(13, 2)
                xlNumberRange.NumberFormat = "#,##0"
                xlSummarySheet2.Cells(14, 1) = "Custodians"
                xlSummarySheet2.Cells(14, 2) = sCustodians
                xlSummarySheet2.Cells(15, 1) = "Custodian Count"
                xlSummarySheet2.Cells(15, 2) = sCustodianCount

                xlHighlightRange = xlSummarySheet2.Range("A9:B15")
                With xlHighlightRange.Borders(Excel.XlBordersIndex.xlEdgeLeft)
                    .LineStyle = Excel.XlLineStyle.xlContinuous
                    .ColorIndex = 0
                    .TintAndShade = 0
                    .Weight = Excel.XlBorderWeight.xlThin
                End With
                With xlHighlightRange.Borders(Excel.XlBordersIndex.xlEdgeTop)
                    .LineStyle = Excel.XlLineStyle.xlContinuous
                    .ColorIndex = 0
                    .TintAndShade = 0
                    .Weight = Excel.XlBorderWeight.xlThin
                End With
                With xlHighlightRange.Borders(Excel.XlBordersIndex.xlEdgeBottom)
                    .LineStyle = Excel.XlLineStyle.xlContinuous
                    .ColorIndex = 0
                    .TintAndShade = 0
                    .Weight = Excel.XlBorderWeight.xlThin
                End With
                With xlHighlightRange.Borders(Excel.XlBordersIndex.xlEdgeRight)
                    .LineStyle = Excel.XlLineStyle.xlContinuous
                    .ColorIndex = 0
                    .TintAndShade = 0
                    .Weight = Excel.XlBorderWeight.xlThin
                End With
                With xlHighlightRange.Interior
                    .ThemeColor = Excel.XlThemeColor.xlThemeColorDark2
                    .TintAndShade = 0
                    .PatternTintAndShade = 0
                End With
                With xlHighlightRange.IndentLevel = 1

                End With


                xlSummarySheet2.Cells(16, 1) = "Culling Information"
                xlHighlightRange = xlSummarySheet2.Range("A16:B16")
                With xlHighlightRange.Borders(Excel.XlBordersIndex.xlEdgeLeft)
                    .LineStyle = Excel.XlLineStyle.xlContinuous
                    .ColorIndex = 0
                    .TintAndShade = 0
                    .Weight = Excel.XlBorderWeight.xlThick
                End With
                With xlHighlightRange.Borders(Excel.XlBordersIndex.xlEdgeTop)
                    .LineStyle = Excel.XlLineStyle.xlContinuous
                    .ColorIndex = 0
                    .TintAndShade = 0
                    .Weight = Excel.XlBorderWeight.xlThick
                End With
                With xlHighlightRange.Borders(Excel.XlBordersIndex.xlEdgeBottom)
                    .LineStyle = Excel.XlLineStyle.xlContinuous
                    .ColorIndex = 0
                    .TintAndShade = 0
                    .Weight = Excel.XlBorderWeight.xlThick
                End With
                With xlHighlightRange.Borders(Excel.XlBordersIndex.xlEdgeRight)
                    .LineStyle = Excel.XlLineStyle.xlContinuous
                    .ColorIndex = 0
                    .TintAndShade = 0
                    .Weight = Excel.XlBorderWeight.xlThick
                End With
                With xlHighlightRange.Interior
                    .ThemeColor = Excel.XlThemeColor.xlThemeColorDark2
                    .TintAndShade = 0
                    .PatternTintAndShade = 0
                End With
                With xlHighlightRange.Interior
                    .ThemeColor = Excel.XlThemeColor.xlThemeColorAccent3
                    .TintAndShade = 0
                End With

                xlSummarySheet2.Cells(17, 1) = "Total Item Size"
                xlSummarySheet2.Cells(17, 2) = sCaseFileSize
                xlSummarySheet2.Cells(17, 3) = "Bytes"
                xlNumberRange = xlSummarySheet2.Cells(17, 2)
                xlNumberRange.NumberFormat = "#,##0"
                xlSummarySheet2.Cells(18, 1) = "Original Item Count"
                xlSummarySheet2.Cells(18, 2) = lstOriginalItemCount(lstTotalItem.IndexOf("Total"))
                xlNumberRange = xlSummarySheet2.Cells(18, 2)
                xlNumberRange.NumberFormat = "#,##0"
                xlSummarySheet2.Cells(19, 1) = "Original Item Size"
                xlSummarySheet2.Cells(19, 2) = lstOriginalItemSize(lstTotalItem.IndexOf("Total"))
                xlNumberRange = xlSummarySheet2.Cells(19, 2)
                xlSummarySheet2.Cells(19, 3) = "Bytes"

                xlNumberRange.NumberFormat = "#,##0"

                xlSummarySheet2.Cells(20, 1) = "Duplicate Item Count"
                xlSummarySheet2.Cells(20, 2) = lstDuplicateItemCount(lstTotalItem.IndexOf("Total"))
                xlNumberRange = xlSummarySheet2.Cells(20, 2)
                xlNumberRange.NumberFormat = "#,##0"

                xlSummarySheet2.Cells(21, 1) = "Duplicate Item Size"
                xlSummarySheet2.Cells(21, 2) = lstDuplicateItemSize(lstTotalItem.IndexOf("Total"))
                xlNumberRange = xlSummarySheet2.Cells(21, 2)
                xlSummarySheet2.Cells(21, 3) = "Bytes"
                xlNumberRange.NumberFormat = "#,##0"

                xlSummarySheet2.Cells(22, 1) = "Percentage"
                xlSummarySheet2.Cells(22, 2) = "=B20/B12"
                xlNumberRange = xlSummarySheet2.Cells(22, 2)
                xlNumberRange.NumberFormat = "0.00%"
                xlHighlightRange = xlSummarySheet2.Range("A17:B22")
                With xlHighlightRange.Borders(Excel.XlBordersIndex.xlEdgeLeft)
                    .LineStyle = Excel.XlLineStyle.xlContinuous
                    .ColorIndex = 0
                    .TintAndShade = 0
                    .Weight = Excel.XlBorderWeight.xlThin
                End With
                With xlHighlightRange.Borders(Excel.XlBordersIndex.xlEdgeTop)
                    .LineStyle = Excel.XlLineStyle.xlContinuous
                    .ColorIndex = 0
                    .TintAndShade = 0
                    .Weight = Excel.XlBorderWeight.xlThin
                End With
                With xlHighlightRange.Borders(Excel.XlBordersIndex.xlEdgeBottom)
                    .LineStyle = Excel.XlLineStyle.xlContinuous
                    .ColorIndex = 0
                    .TintAndShade = 0
                    .Weight = Excel.XlBorderWeight.xlThin
                End With
                With xlHighlightRange.Borders(Excel.XlBordersIndex.xlEdgeRight)
                    .LineStyle = Excel.XlLineStyle.xlContinuous
                    .ColorIndex = 0
                    .TintAndShade = 0
                    .Weight = Excel.XlBorderWeight.xlThin
                End With
                With xlHighlightRange.Interior
                    .ThemeColor = Excel.XlThemeColor.xlThemeColorDark2
                    .TintAndShade = 0
                    .PatternTintAndShade = 0
                End With
                With xlHighlightRange.IndentLevel = 1

                End With


                xlSummarySheet2.Cells(23, 1) = "Exceptions"
                xlHighlightRange = xlSummarySheet2.Range("A23:B23")
                With xlHighlightRange.Borders(Excel.XlBordersIndex.xlEdgeLeft)
                    .LineStyle = Excel.XlLineStyle.xlContinuous
                    .ColorIndex = 0
                    .TintAndShade = 0
                    .Weight = Excel.XlBorderWeight.xlThick
                End With
                With xlHighlightRange.Borders(Excel.XlBordersIndex.xlEdgeTop)
                    .LineStyle = Excel.XlLineStyle.xlContinuous
                    .ColorIndex = 0
                    .TintAndShade = 0
                    .Weight = Excel.XlBorderWeight.xlThick
                End With
                With xlHighlightRange.Borders(Excel.XlBordersIndex.xlEdgeBottom)
                    .LineStyle = Excel.XlLineStyle.xlContinuous
                    .ColorIndex = 0
                    .TintAndShade = 0
                    .Weight = Excel.XlBorderWeight.xlThick
                End With
                With xlHighlightRange.Borders(Excel.XlBordersIndex.xlEdgeRight)
                    .LineStyle = Excel.XlLineStyle.xlContinuous
                    .ColorIndex = 0
                    .TintAndShade = 0
                    .Weight = Excel.XlBorderWeight.xlThick
                End With
                With xlHighlightRange.Interior
                    .ThemeColor = Excel.XlThemeColor.xlThemeColorDark2
                    .TintAndShade = 0
                    .PatternTintAndShade = 0
                End With
                With xlHighlightRange.Interior
                    .ThemeColor = Excel.XlThemeColor.xlThemeColorAccent3
                    .TintAndShade = 0
                End With

                xlSummarySheet2.Cells(24, 1) = "Encrypted"
                xlSummarySheet2.Cells(24, 2) = lstIrregularItemsCount(lstIrregularItems.IndexOf("Encrypted"))
                xlNumberRange = xlSummarySheet2.Cells(24, 2)
                xlNumberRange.NumberFormat = "#,##0"
                xlSummarySheet2.Cells(25, 1) = "Decrypted"
                xlSummarySheet2.Cells(25, 2) = lstIrregularItemsCount(lstIrregularItems.IndexOf("Decrypted"))
                xlNumberRange = xlSummarySheet2.Cells(25, 2)
                xlNumberRange.NumberFormat = "#,##0"
                xlSummarySheet2.Cells(26, 1) = "Partially Processed"
                xlSummarySheet2.Cells(26, 2) = lstIrregularItemsCount(lstIrregularItems.IndexOf("Partially_processed"))
                xlNumberRange = xlSummarySheet2.Cells(26, 2)
                xlNumberRange.NumberFormat = "#,##0"
                xlSummarySheet2.Cells(27, 1) = "Corrupted"
                xlSummarySheet2.Cells(27, 2) = lstIrregularItemsCount(lstIrregularItems.IndexOf("Corrupted"))
                xlNumberRange = xlSummarySheet2.Cells(27, 2)
                xlNumberRange.NumberFormat = "#,##0"
                xlSummarySheet2.Cells(28, 1) = "Not Processed"
                xlSummarySheet2.Cells(28, 2) = lstIrregularItemsCount(lstIrregularItems.IndexOf("Not_Processed"))
                xlNumberRange = xlSummarySheet2.Cells(28, 2)
                xlNumberRange.NumberFormat = "#,##0"
                xlSummarySheet2.Cells(29, 1) = "Poisoned"
                xlSummarySheet2.Cells(29, 2) = lstIrregularItemsCount(lstIrregularItems.IndexOf("Poisoned"))
                xlNumberRange = xlSummarySheet2.Cells(29, 2)
                xlNumberRange.NumberFormat = "#,##0"
                xlHighlightRange = xlSummarySheet2.Range("A24:B29")
                With xlHighlightRange.Borders(Excel.XlBordersIndex.xlEdgeLeft)
                    .LineStyle = Excel.XlLineStyle.xlContinuous
                    .ColorIndex = 0
                    .TintAndShade = 0
                    .Weight = Excel.XlBorderWeight.xlThin
                End With
                With xlHighlightRange.Borders(Excel.XlBordersIndex.xlEdgeTop)
                    .LineStyle = Excel.XlLineStyle.xlContinuous
                    .ColorIndex = 0
                    .TintAndShade = 0
                    .Weight = Excel.XlBorderWeight.xlThin
                End With
                With xlHighlightRange.Borders(Excel.XlBordersIndex.xlEdgeBottom)
                    .LineStyle = Excel.XlLineStyle.xlContinuous
                    .ColorIndex = 0
                    .TintAndShade = 0
                    .Weight = Excel.XlBorderWeight.xlThin
                End With
                With xlHighlightRange.Borders(Excel.XlBordersIndex.xlEdgeRight)
                    .LineStyle = Excel.XlLineStyle.xlContinuous
                    .ColorIndex = 0
                    .TintAndShade = 0
                    .Weight = Excel.XlBorderWeight.xlThin
                End With
                With xlHighlightRange.Interior
                    .ThemeColor = Excel.XlThemeColor.xlThemeColorDark2
                    .TintAndShade = 0
                    .PatternTintAndShade = 0
                End With
                With xlHighlightRange.IndentLevel = 1

                End With

                xlSummarySheet2.Columns("A").autofit()
                xlSummarySheet2.Columns("B").autofit()
                xlSummarySheet2.Columns("C").autofit()

                xlWorkbook.SaveAs(txtReportLocation.Text & "\" & sCaseName & "-" & DateTime.Now.ToString("yyyy-MM-dd hh-mm-ss") & ".xlsx")

                xlWorkbook.Close()
                oExcelApp.Quit()

                xlIrregularItems = Nothing
                xlEncryptedMessagesSheet = Nothing
                xlCulling = Nothing
                xlDateRangeGapPivot = Nothing
                xlDateRangeSheet = Nothing
                xlLanguagesSheet = Nothing
                xlMimeTypeSheet = Nothing
                xlItemTypeSheet = Nothing
                xlSummarySheet = Nothing
                xlWorkbook = Nothing
                oExcelApp = Nothing

                row.cells("CollectionStatus").value = "Case Data Exported"
            End If
        Next

        If chkRollUpReporting.Checked = True Then
            Try
                oExcelApp = New Microsoft.Office.Interop.Excel.Application

            Catch ex As Exception
                MessageBox.Show("It appears that Microsoft Excel is not installed on this machine. It is required to have Excel installed in order to export to Excel.", "Dependency missing", MessageBoxButtons.OK)
                Exit Sub
            End Try
            Dim sExcelVersion As String
            Dim iAllCaseRow As Integer
            Dim iItemTypeCounter As Integer
            Dim iMimeTypeCounter As Integer
            Dim iTotalItemCounter As Integer
            Dim iLanguageCounter As Integer
            Dim iIrregularItemsCounter As Integer
            Dim asNuixVersionNumber() As String

            sExcelVersion = oExcelApp.Version.ToString
            Logger(psUCRTLogFile, "Excel Version Number = " & sExcelVersion)
            'oExcelWorkbook = New Excel.Workbook
            Dim xlWorkbook As Excel.Workbook = oExcelApp.Workbooks.Add()
            Dim xlAllCaseData As Excel.Worksheet = CType(xlWorkbook.Sheets.Add(), Excel.Worksheet)
            xlAllCaseData.Name = "All Case Summary"
            oExcelApp.Windows.Application.ActiveWindow.DisplayGridlines = False

            Dim ImageRange As Excel.Range
            ImageRange = xlAllCaseData.Range("D1:D4")
            ImageRange.MergeCells = True
            Dim sImageFile As String = IO.Path.Combine(Application.StartupPath, "Resources\nuix-logo-updated.jpg")

            xlAllCaseData.Shapes.AddPicture(sImageFile, False, True, 120, 0, 100, 60)

            dt = Nothing
            ds = New DataSet
            sqlConnection = New SQLiteConnection("Data Source=" & sReportingDBLocation & "\NuixCaseReports.db3;Version=3;Read Only=True;New=False;Compress=True;")

            mSQL = "select Count(CurrentCaseVersion), CurrentCaseVersion from NuixReportingInfo Group by CurrentCaseVersion"
            sqlCommand = New SQLiteCommand(mSQL, sqlConnection)
            sqlConnection.Open()

            dataReader = sqlCommand.ExecuteReader
            iAllCaseRow = 1
            xlAllCaseData.Cells(iAllCaseRow, 1) = "Nuix Version"
            xlAllCaseData.Cells(iAllCaseRow, 2) = "Number of Cases"
            xlAllCaseData.Range("A1:B1").Font.Name = "Arial Black"

            Dim NuixCaseVersionRange As Excel.Range

            NuixCaseVersionRange = xlAllCaseData.Range("A1:B1")
            With NuixCaseVersionRange.Borders(Excel.XlBordersIndex.xlEdgeLeft)
                .LineStyle = Excel.XlLineStyle.xlContinuous
                .ColorIndex = 0
                .TintAndShade = 0
                .Weight = Excel.XlBorderWeight.xlThick
            End With
            With NuixCaseVersionRange.Borders(Excel.XlBordersIndex.xlEdgeTop)
                .LineStyle = Excel.XlLineStyle.xlContinuous
                .ColorIndex = 0
                .TintAndShade = 0
                .Weight = Excel.XlBorderWeight.xlThick
            End With
            With NuixCaseVersionRange.Borders(Excel.XlBordersIndex.xlEdgeBottom)
                .LineStyle = Excel.XlLineStyle.xlContinuous
                .ColorIndex = 0
                .TintAndShade = 0
                .Weight = Excel.XlBorderWeight.xlThick
            End With
            With NuixCaseVersionRange.Borders(Excel.XlBordersIndex.xlEdgeRight)
                .LineStyle = Excel.XlLineStyle.xlContinuous
                .ColorIndex = 0
                .TintAndShade = 0
                .Weight = Excel.XlBorderWeight.xlThick
            End With

            iAllCaseRow = iAllCaseRow + 1

            While dataReader.Read
                asNuixVersionNumber = Split(dataReader.GetString(1), ".")
                If (CInt(asNuixVersionNumber(0)) < 7) Then
                    xlAllCaseData.Range("A" & iAllCaseRow).Font.ColorIndex = 3
                ElseIf (CInt(asNuixVersionNumber(0)) = 7) And (CInt(asNuixVersionNumber(1)) < 3) Then
                    xlAllCaseData.Range("A" & iAllCaseRow).Font.ColorIndex = 44
                Else
                    xlAllCaseData.Range("A" & iAllCaseRow).Font.ColorIndex = 4
                End If
                xlAllCaseData.Cells(iAllCaseRow, 1) = dataReader.GetString(1)
                xlAllCaseData.Range("A" & iAllCaseRow).IndentLevel = 1
                xlAllCaseData.Cells(iAllCaseRow, 2) = dataReader.GetInt16(0)
                NuixCaseVersionRange = xlAllCaseData.Range("A1:B1")
                With NuixCaseVersionRange.Borders(Excel.XlBordersIndex.xlEdgeLeft)
                    .LineStyle = Excel.XlLineStyle.xlContinuous
                    .ColorIndex = 0
                    .TintAndShade = 0
                    .Weight = Excel.XlBorderWeight.xlThick
                End With
                With NuixCaseVersionRange.Borders(Excel.XlBordersIndex.xlEdgeTop)
                    .LineStyle = Excel.XlLineStyle.xlContinuous
                    .ColorIndex = 0
                    .TintAndShade = 0
                    .Weight = Excel.XlBorderWeight.xlThick
                End With
                With NuixCaseVersionRange.Borders(Excel.XlBordersIndex.xlEdgeBottom)
                    .LineStyle = Excel.XlLineStyle.xlContinuous
                    .ColorIndex = 0
                    .TintAndShade = 0
                    .Weight = Excel.XlBorderWeight.xlThick
                End With
                With NuixCaseVersionRange.Borders(Excel.XlBordersIndex.xlEdgeRight)
                    .LineStyle = Excel.XlLineStyle.xlContinuous
                    .ColorIndex = 0
                    .TintAndShade = 0
                    .Weight = Excel.XlBorderWeight.xlThick
                End With

                iAllCaseRow = iAllCaseRow + 1
                xlAllCaseData.Range("A" & iAllCaseRow & ":B" & iAllCaseRow).Font.Name = "Arial"
            End While

            iAllCaseRow = iAllCaseRow + 1
            xlAllCaseData.Cells(iAllCaseRow, 1) = "Item Type"
            xlAllCaseData.Cells(iAllCaseRow, 2) = "Count"
            xlAllCaseData.Cells(iAllCaseRow, 3) = "Size"
            xlAllCaseData.Range("A" & iAllCaseRow & ":C" & iAllCaseRow).Font.Name = "Arial Black"
            Dim NuixMimeTypeRange As Excel.Range

            NuixMimeTypeRange = xlAllCaseData.Range("A" & iAllCaseRow & ":C" & iAllCaseRow)
            With NuixMimeTypeRange.Borders(Excel.XlBordersIndex.xlEdgeLeft)
                .LineStyle = Excel.XlLineStyle.xlContinuous
                .ColorIndex = 0
                .TintAndShade = 0
                .Weight = Excel.XlBorderWeight.xlThick
            End With
            With NuixMimeTypeRange.Borders(Excel.XlBordersIndex.xlEdgeTop)
                .LineStyle = Excel.XlLineStyle.xlContinuous
                .ColorIndex = 0
                .TintAndShade = 0
                .Weight = Excel.XlBorderWeight.xlThick
            End With
            With NuixMimeTypeRange.Borders(Excel.XlBordersIndex.xlEdgeBottom)
                .LineStyle = Excel.XlLineStyle.xlContinuous
                .ColorIndex = 0
                .TintAndShade = 0
                .Weight = Excel.XlBorderWeight.xlThick
            End With
            With NuixMimeTypeRange.Borders(Excel.XlBordersIndex.xlEdgeRight)
                .LineStyle = Excel.XlLineStyle.xlContinuous
                .ColorIndex = 0
                .TintAndShade = 0
                .Weight = Excel.XlBorderWeight.xlThick
            End With

            For Each ItemType In lstItemType
                iAllCaseRow = iAllCaseRow + 1
                xlAllCaseData.Cells(iAllCaseRow, 1) = ItemType.ToString
                xlAllCaseData.Range("A" & iAllCaseRow).IndentLevel = 1

                xlAllCaseData.Cells(iAllCaseRow, 2) = lstItemTypeCount(iItemTypeCounter)
                xlAllCaseData.Cells(iAllCaseRow, 3) = lstItemTypeSize(iItemTypeCounter)
                xlAllCaseData.Range("A" & iAllCaseRow & ":C" & iAllCaseRow).Font.Name = "Arial"
                iItemTypeCounter = iItemTypeCounter + 1
            Next

            iAllCaseRow = iAllCaseRow + 1
            iAllCaseRow = iAllCaseRow + 1
            xlAllCaseData.Cells(iAllCaseRow, 1) = "Mime Type"
            xlAllCaseData.Cells(iAllCaseRow, 2) = "Count"
            xlAllCaseData.Cells(iAllCaseRow, 3) = "Size"
            xlAllCaseData.Range("A" & iAllCaseRow & ":C" & iAllCaseRow).Font.Name = "Arial Black"
            Dim xlMimeTypeRange As Excel.Range
            xlMimeTypeRange = xlAllCaseData.Range("A" & iAllCaseRow & ":C" & iAllCaseRow)
            With xlMimeTypeRange.Borders(Excel.XlBordersIndex.xlEdgeLeft)
                .LineStyle = Excel.XlLineStyle.xlContinuous
                .ColorIndex = 0
                .TintAndShade = 0
                .Weight = Excel.XlBorderWeight.xlThick
            End With
            With xlMimeTypeRange.Borders(Excel.XlBordersIndex.xlEdgeTop)
                .LineStyle = Excel.XlLineStyle.xlContinuous
                .ColorIndex = 0
                .TintAndShade = 0
                .Weight = Excel.XlBorderWeight.xlThick
            End With
            With xlMimeTypeRange.Borders(Excel.XlBordersIndex.xlEdgeBottom)
                .LineStyle = Excel.XlLineStyle.xlContinuous
                .ColorIndex = 0
                .TintAndShade = 0
                .Weight = Excel.XlBorderWeight.xlThick
            End With
            With xlMimeTypeRange.Borders(Excel.XlBordersIndex.xlEdgeRight)
                .LineStyle = Excel.XlLineStyle.xlContinuous
                .ColorIndex = 0
                .TintAndShade = 0
                .Weight = Excel.XlBorderWeight.xlThick
            End With

            For Each MimeType In lstMimeType
                iAllCaseRow = iAllCaseRow + 1
                xlAllCaseData.Cells(iAllCaseRow, 1) = MimeType.ToString
                xlAllCaseData.Range("A" & iAllCaseRow).IndentLevel = 1

                xlAllCaseData.Cells(iAllCaseRow, 2) = lstMimeTypeCount(iMimeTypeCounter)
                xlAllCaseData.Cells(iAllCaseRow, 3) = lstMimeTypeSize(iMimeTypeCounter)
                xlAllCaseData.Range("A" & iAllCaseRow & ":C" & iAllCaseRow).Font.Name = "Arial"
                iMimeTypeCounter = iMimeTypeCounter + 1
            Next
            iAllCaseRow = iAllCaseRow + 1
            iAllCaseRow = iAllCaseRow + 1
            xlAllCaseData.Cells(iAllCaseRow, 1) = "Kind"
            xlAllCaseData.Cells(iAllCaseRow, 2) = "Total Count"
            xlAllCaseData.Cells(iAllCaseRow, 3) = "Total Size"
            xlAllCaseData.Cells(iAllCaseRow, 4) = "Original Count"
            xlAllCaseData.Cells(iAllCaseRow, 5) = "Original Size"
            xlAllCaseData.Cells(iAllCaseRow, 6) = "Duplicate Count"
            xlAllCaseData.Cells(iAllCaseRow, 7) = "Duplicate Size"
            xlAllCaseData.Range("A" & iAllCaseRow & ":G" & iAllCaseRow).Font.Name = "Arial Black"
            Dim NuixKindRange As Excel.Range
            NuixKindRange = xlAllCaseData.Range("A" & iAllCaseRow & ":G" & iAllCaseRow)
            With NuixKindRange.Borders(Excel.XlBordersIndex.xlEdgeLeft)
                .LineStyle = Excel.XlLineStyle.xlContinuous
                .ColorIndex = 0
                .TintAndShade = 0
                .Weight = Excel.XlBorderWeight.xlThick
            End With
            With NuixKindRange.Borders(Excel.XlBordersIndex.xlEdgeTop)
                .LineStyle = Excel.XlLineStyle.xlContinuous
                .ColorIndex = 0
                .TintAndShade = 0
                .Weight = Excel.XlBorderWeight.xlThick
            End With
            With NuixKindRange.Borders(Excel.XlBordersIndex.xlEdgeBottom)
                .LineStyle = Excel.XlLineStyle.xlContinuous
                .ColorIndex = 0
                .TintAndShade = 0
                .Weight = Excel.XlBorderWeight.xlThick
            End With
            With NuixKindRange.Borders(Excel.XlBordersIndex.xlEdgeRight)
                .LineStyle = Excel.XlLineStyle.xlContinuous
                .ColorIndex = 0
                .TintAndShade = 0
                .Weight = Excel.XlBorderWeight.xlThick
            End With

            For Each TotalItems In lstTotalItem
                iAllCaseRow = iAllCaseRow + 1
                xlAllCaseData.Cells(iAllCaseRow, 1) = TotalItems.ToString
                xlAllCaseData.Range("A" & iAllCaseRow).IndentLevel = 1

                xlAllCaseData.Cells(iAllCaseRow, 2) = lstTotalItemCount(iTotalItemCounter)
                xlAllCaseData.Cells(iAllCaseRow, 3) = lstTotalItemSize(iTotalItemCounter)
                xlAllCaseData.Cells(iAllCaseRow, 4) = lstOriginalItemCount(iTotalItemCounter)
                xlAllCaseData.Cells(iAllCaseRow, 5) = lstOriginalItemSize(iTotalItemCounter)
                xlAllCaseData.Cells(iAllCaseRow, 6) = lstDuplicateItemCount(iTotalItemCounter)
                xlAllCaseData.Cells(iAllCaseRow, 7) = lstDuplicateItemSize(iTotalItemCounter)
                xlAllCaseData.Range("A" & iAllCaseRow & ":G" & iAllCaseRow).Font.Name = "Arial"
                iTotalItemCounter = iTotalItemCounter + 1
            Next
            iAllCaseRow = iAllCaseRow + 1
            iAllCaseRow = iAllCaseRow + 1
            xlAllCaseData.Cells(iAllCaseRow, 1) = "Language"
            xlAllCaseData.Cells(iAllCaseRow, 2) = "Count"
            xlAllCaseData.Cells(iAllCaseRow, 3) = "Size"
            xlAllCaseData.Range("A" & iAllCaseRow & ":C" & iAllCaseRow).Font.Name = "Arial Black"
            Dim NuixLanguageRange As Excel.Range
            NuixLanguageRange = xlAllCaseData.Range("A" & iAllCaseRow & ":C" & iAllCaseRow)
            With NuixLanguageRange.Borders(Excel.XlBordersIndex.xlEdgeLeft)
                .LineStyle = Excel.XlLineStyle.xlContinuous
                .ColorIndex = 0
                .TintAndShade = 0
                .Weight = Excel.XlBorderWeight.xlThick
            End With
            With NuixLanguageRange.Borders(Excel.XlBordersIndex.xlEdgeTop)
                .LineStyle = Excel.XlLineStyle.xlContinuous
                .ColorIndex = 0
                .TintAndShade = 0
                .Weight = Excel.XlBorderWeight.xlThick
            End With
            With NuixLanguageRange.Borders(Excel.XlBordersIndex.xlEdgeBottom)
                .LineStyle = Excel.XlLineStyle.xlContinuous
                .ColorIndex = 0
                .TintAndShade = 0
                .Weight = Excel.XlBorderWeight.xlThick
            End With
            With NuixLanguageRange.Borders(Excel.XlBordersIndex.xlEdgeRight)
                .LineStyle = Excel.XlLineStyle.xlContinuous
                .ColorIndex = 0
                .TintAndShade = 0
                .Weight = Excel.XlBorderWeight.xlThick
            End With
            For Each Language In lstLanguages
                iAllCaseRow = iAllCaseRow + 1
                xlAllCaseData.Cells(iAllCaseRow, 1) = Language.ToString
                xlAllCaseData.Range("A" & iAllCaseRow).IndentLevel = 1

                xlAllCaseData.Cells(iAllCaseRow, 2) = lstLanguagesCount(iLanguageCounter)
                xlAllCaseData.Cells(iAllCaseRow, 3) = lstLanguagesSize(iLanguageCounter)
                xlAllCaseData.Range("A" & iAllCaseRow & ":C" & iAllCaseRow).Font.Name = "Arial"
                iLanguageCounter = iLanguageCounter + 1
            Next
            iAllCaseRow = iAllCaseRow + 1
            iAllCaseRow = iAllCaseRow + 1
            xlAllCaseData.Cells(iAllCaseRow, 1) = "Irregular Item"
            xlAllCaseData.Cells(iAllCaseRow, 2) = "Count"
            xlAllCaseData.Cells(iAllCaseRow, 3) = "Size"
            xlAllCaseData.Range("A" & iAllCaseRow & ":C" & iAllCaseRow).Font.Name = "Arial Black"

            Dim NuixIrregularItemsRange As Excel.Range
            NuixIrregularItemsRange = xlAllCaseData.Range("A" & iAllCaseRow & ":C" & iAllCaseRow)
            With NuixIrregularItemsRange.Borders(Excel.XlBordersIndex.xlEdgeLeft)
                .LineStyle = Excel.XlLineStyle.xlContinuous
                .ColorIndex = 0
                .TintAndShade = 0
                .Weight = Excel.XlBorderWeight.xlThick
            End With
            With NuixIrregularItemsRange.Borders(Excel.XlBordersIndex.xlEdgeTop)
                .LineStyle = Excel.XlLineStyle.xlContinuous
                .ColorIndex = 0
                .TintAndShade = 0
                .Weight = Excel.XlBorderWeight.xlThick
            End With
            With NuixIrregularItemsRange.Borders(Excel.XlBordersIndex.xlEdgeBottom)
                .LineStyle = Excel.XlLineStyle.xlContinuous
                .ColorIndex = 0
                .TintAndShade = 0
                .Weight = Excel.XlBorderWeight.xlThick
            End With
            With NuixIrregularItemsRange.Borders(Excel.XlBordersIndex.xlEdgeRight)
                .LineStyle = Excel.XlLineStyle.xlContinuous
                .ColorIndex = 0
                .TintAndShade = 0
                .Weight = Excel.XlBorderWeight.xlThick
            End With

            For Each IrregularItem In lstIrregularItems
                iAllCaseRow = iAllCaseRow + 1
                xlAllCaseData.Cells(iAllCaseRow, 1) = IrregularItem.ToString
                xlAllCaseData.Range("A" & iAllCaseRow).IndentLevel = 1

                xlAllCaseData.Cells(iAllCaseRow, 2) = lstIrregularItemsCount(iIrregularItemsCounter)
                xlAllCaseData.Cells(iAllCaseRow, 3) = lstIrregularItemsSize(iIrregularItemsCounter)
                xlAllCaseData.Range("A" & iAllCaseRow & ":C" & iAllCaseRow).Font.Name = "Arial"
                iIrregularItemsCounter = iIrregularItemsCounter + 1
            Next

            xlAllCaseData.Columns("A").autofit()
            xlAllCaseData.Columns("B").autofit()
            xlAllCaseData.Columns("C").autofit()
            xlAllCaseData.Columns("D").autofit()
            xlAllCaseData.Columns("E").autofit()
            xlAllCaseData.Columns("F").autofit()
            xlAllCaseData.Columns("G").autofit()


            xlAllCaseData.Columns("B").NumberFormat = "#,##0"
            xlAllCaseData.Columns("C").NumberFormat = "#,##0"
            xlAllCaseData.Columns("D").NumberFormat = "#,##0"
            xlAllCaseData.Columns("E").NumberFormat = "#,##0"
            xlAllCaseData.Columns("F").NumberFormat = "#,##0"
            xlAllCaseData.Columns("G").NumberFormat = "#,##0"

            xlWorkbook.SaveAs(txtReportLocation.Text & "\AllCaseData-" & DateTime.Now.ToString("yyyy-MM-dd hh-mm-ss") & ".xlsx")

            xlWorkbook.Close()
            oExcelApp.Quit()
            xlAllCaseData = Nothing
            xlWorkbook = Nothing
            oExcelApp = Nothing
        End If

        MessageBox.Show("All Excel reports have been exported to the " & txtReportLocation.Text & " directory", "Excel Reports Exported", MessageBoxButtons.OK)

    End Sub

    Private Sub btnSaveConfig_Click(sender As Object, e As EventArgs) Handles btnSaveConfig.Click
        Dim sReportLocation As String
        Dim bStatus As Boolean
        Dim sReportType As String
        Dim sCalculateProcessingSpeeds As String
        Dim sShowSizeIn As String
        Dim sSearchTerm As String
        Dim sNuixLogFileLocation As String
        Dim sNuixConsoleLocation As String
        Dim sNuixLicense As String
        Dim sLicenseType As String

        Dim sNMSAddress As String
        Dim sNMSUserName As String
        Dim sNMSPassword As String
        Dim sRegistryServer As String
        Dim sAppMemory As String
        Dim sOutputFileName As String
        Dim sMachineName As String
        Dim sSettingsFile As String
        Dim bExportSearchResults As Boolean
        Dim sExportSearchResultsLocation As String
        Dim bSearchTermFile As Boolean
        Dim bSearchTerm As Boolean
        Dim bExportOnly As Boolean
        Dim sExportWorkers As String
        Dim sExportWorkerMemory As String
        Dim bIncludeDiskSize As Boolean


        If txtReportLocation.Text <> vbNullString Then
            sReportLocation = txtReportLocation.Text
            sMachineName = System.Net.Dns.GetHostName()
            sOutputFileName = "UCRTConfiguration-" & sMachineName & "-" & DateTime.Now.ToString("yyyy-MM-dd hh-mm-ss") & ".xml"
            sSettingsFile = sReportLocation & "\" & sOutputFileName
            sReportType = cboReportType.Text
            sCalculateProcessingSpeeds = cboCalculateProcessingSpeeds.Text
            sShowSizeIn = cboSizeReporting.Text
            sSearchTerm = txtSearchTerm.Text
            sNuixLogFileLocation = txtNuixLogDir.Text
            sNuixConsoleLocation = txtNuixConsoleLocation.Text
            sNuixLicense = cboNuixLicenseType.Text
            sLicenseType = cboLicenseType.Text
            sNMSAddress = txtNMSAddress.Text
            sNMSUserName = txtNMSUserName.Text
            sNMSPassword = txtNMSInfo.Text
            sRegistryServer = txtRegistryServer.Text
            sAppMemory = numNuixAppMemory.Text
            bExportSearchResults = chkExportSearchResults.Checked
            sExportSearchResultsLocation = txtExportLocation.Text
            bSearchTermFile = radSearchFile.Checked
            bSearchTerm = radSearchTerm.Checked
            bExportOnly = chkExportOnly.Checked
            sExportWorkers = numExportWorkers.Value
            sExportWorkerMemory = numExportWorkerMemory.Value
            bIncludeDiskSize = chkIncludeDiskSize.Checked
        Else
            MessageBox.Show("You must enter a report location to save the config file to.", "Enter Report Location", MessageBoxButtons.OK)
            Exit Sub
        End If

        bStatus = blnBuildNuixSettingsXML(sSettingsFile, sReportLocation, sReportType, sCalculateProcessingSpeeds, sShowSizeIn, sSearchTerm, sNuixLogFileLocation, sNuixConsoleLocation, sNuixLicense, sLicenseType, sNMSAddress, sNMSUserName, "", sRegistryServer, sAppMemory, bExportSearchResults, sExportSearchResultsLocation, bSearchTermFile, bSearchTerm, bExportOnly, sExportWorkers, sExportWorkerMemory, bIncludeDiskSize)

        MessageBox.Show("Configuration file saved to " & sSettingsFile, "Config File Saved", MessageBoxButtons.OK)
    End Sub

    Private Function blnBuildNuixSettingsXML(ByVal sSettingsFile As String, ByVal sReportLocation As String, ByVal sReportType As String, ByVal sCalculateProcessingSpeeds As String, ByVal sShowSizeIn As String, ByVal sSearchTerm As String, ByVal sNuixLogFileLocation As String, ByVal sNuixConsoleLocation As String, ByVal sNuixLicense As String, ByVal sLicenseType As String, ByVal sNMSAddress As String, ByVal sNMSUserName As String, ByVal sNMSInfo As String, ByVal sRegistryServer As String, ByVal sNuixAppMemory As String, ByVal bExportSearchResults As Boolean, ByVal sExportSearchResultsLocation As String, ByVal bSearchTermFile As Boolean, ByVal bSearchTerm As Boolean, ByVal bExportOnly As Boolean, ByVal sExportWorkers As String, ByVal sExportWorkerMemory As String, ByVal bIncludeDiskSpace As Boolean) As Boolean
        Dim SettingsXML As Xml.XmlDocument
        Dim SettingsRoot As Xml.XmlElement

        Dim NuixSettings As Xml.XmlElement
        Dim ReportLocation As Xml.XmlElement
        Dim ReportType As Xml.XmlElement
        Dim CalculateProcessingSpeeds As Xml.XmlElement
        Dim ShowSizeIn As Xml.XmlElement
        Dim SearchTerm As Xml.XmlElement
        Dim NuixLogFileLocation As Xml.XmlElement
        Dim NuixConsoleLocation As Xml.XmlElement
        Dim NuixLicense As Xml.XmlElement
        Dim LicenseType As Xml.XmlElement
        Dim NMSAddress As Xml.XmlElement
        Dim NMSUser As Xml.XmlElement
        Dim NMSInfo As Xml.XmlElement
        Dim RegistryServer As Xml.XmlElement
        Dim AppMemory As Xml.XmlElement
        Dim ExportSearchResults As Xml.XmlElement
        Dim ExportSearchResultsLocation As Xml.XmlElement
        Dim SearchTermFile As Xml.XmlElement
        Dim SearchTermTerm As Xml.XmlElement
        Dim ExportOnly As Xml.XmlElement
        Dim ExportWorkers As Xml.XmlElement
        Dim ExportWorkerMemory As Xml.XmlElement
        Dim IncludeDiskSpace As Xml.XmlElement

        Dim xmlDeclaration As Xml.XmlDeclaration

        blnBuildNuixSettingsXML = False

        SettingsXML = New Xml.XmlDocument
        xmlDeclaration = SettingsXML.CreateXmlDeclaration("1.0", "UTF-8", "yes")
        SettingsRoot = SettingsXML.CreateElement("AllNuixSettings")
        SettingsXML.AppendChild(SettingsRoot)

        SettingsXML.InsertBefore(xmlDeclaration, SettingsRoot)

        NuixSettings = SettingsXML.CreateElement("UCRTSettings")

        SettingsRoot.AppendChild(NuixSettings)
        ReportLocation = SettingsXML.CreateElement("ReportLocation")
        NuixSettings.AppendChild(ReportLocation)
        ReportLocation.InnerText = sReportLocation

        ReportType = SettingsXML.CreateElement("ReportType")
        NuixSettings.AppendChild(ReportType)
        ReportType.InnerText = sReportType

        CalculateProcessingSpeeds = SettingsXML.CreateElement("CalculateProcessingSpeeds")
        NuixSettings.AppendChild(CalculateProcessingSpeeds)
        CalculateProcessingSpeeds.InnerText = sCalculateProcessingSpeeds

        ShowSizeIn = SettingsXML.CreateElement("ShowSizeIn")
        NuixSettings.AppendChild(ShowSizeIn)
        ShowSizeIn.InnerText = sShowSizeIn

        SearchTerm = SettingsXML.CreateElement("SearchTerm")
        NuixSettings.AppendChild(SearchTerm)
        SearchTerm.InnerText = sSearchTerm

        NuixLogFileLocation = SettingsXML.CreateElement("NuixLogFileLocation")
        NuixSettings.AppendChild(NuixLogFileLocation)
        NuixLogFileLocation.InnerText = sNuixLogFileLocation

        NuixConsoleLocation = SettingsXML.CreateElement("NuixConsoleLocation")
        NuixSettings.AppendChild(NuixConsoleLocation)
        NuixConsoleLocation.InnerText = sNuixConsoleLocation

        NuixLicense = SettingsXML.CreateElement("NuixLicense")
        NuixSettings.AppendChild(NuixLicense)
        NuixLicense.InnerText = sNuixLicense

        LicenseType = SettingsXML.CreateElement("LicenseType")
        NuixSettings.AppendChild(LicenseType)
        LicenseType.InnerText = sLicenseType

        NMSAddress = SettingsXML.CreateElement("NMSAddress")
        NuixSettings.AppendChild(NMSAddress)
        NMSAddress.InnerText = sNMSAddress

        NMSUser = SettingsXML.CreateElement("NMSUser")
        NuixSettings.AppendChild(NMSUser)
        NMSUser.InnerText = sNMSUserName


        NMSInfo = SettingsXML.CreateElement("NMSInfo")
        NuixSettings.AppendChild(NMSInfo)
        NMSInfo.InnerText = sNMSInfo

        RegistryServer = SettingsXML.CreateElement("RegistryServer")
        NuixSettings.AppendChild(RegistryServer)
        RegistryServer.InnerText = sRegistryServer

        AppMemory = SettingsXML.CreateElement("AppMemory")
        NuixSettings.AppendChild(AppMemory)
        AppMemory.InnerText = sNuixAppMemory

        ExportSearchResults = SettingsXML.CreateElement("ExportSearchResults")
        NuixSettings.AppendChild(ExportSearchResults)
        ExportSearchResults.InnerText = bExportSearchResults.ToString

        ExportSearchResultsLocation = SettingsXML.CreateElement("ExportSearchResultsLocation")
        NuixSettings.AppendChild(ExportSearchResultsLocation)
        ExportSearchResultsLocation.InnerText = sExportSearchResultsLocation

        SearchTermFile = SettingsXML.CreateElement("SearchTermFile")
        NuixSettings.AppendChild(SearchTermFile)
        SearchTermFile.InnerText = bSearchTermFile.ToString

        SearchTermTerm = SettingsXML.CreateElement("SearchTermTerm")
        NuixSettings.AppendChild(SearchTermTerm)
        SearchTermTerm.InnerText = bSearchTerm.ToString

        ExportOnly = SettingsXML.CreateElement("ExportOnly")
        NuixSettings.AppendChild(ExportOnly)
        ExportOnly.InnerText = bExportOnly.ToString

        ExportWorkers = SettingsXML.CreateElement("ExportWorkers")
        NuixSettings.AppendChild(ExportWorkers)
        ExportWorkers.InnerText = sExportWorkers

        ExportWorkerMemory = SettingsXML.CreateElement("ExportWorkerMemory")
        NuixSettings.AppendChild(ExportWorkerMemory)
        ExportWorkerMemory.InnerText = sExportWorkerMemory

        IncludeDiskSpace = SettingsXML.CreateElement("IncludeDiskSpace")
        NuixSettings.AppendChild(IncludeDiskSpace)
        IncludeDiskSpace.InnerText = bIncludeDiskSpace

        SettingsXML.Save(sSettingsFile)

        blnBuildNuixSettingsXML = True

    End Function

    Private Sub btnLoadConfig_Click(sender As Object, e As EventArgs) Handles btnLoadConfig.Click

        Dim sReportLocation As String
        Dim bStatus As Boolean
        Dim sReportType As String
        Dim sCalculateProcessingSpeeds As String
        Dim sShowSizeIn As String
        Dim sSearchTerm As String
        Dim sNuixLogFileLocation As String
        Dim sNuixConsoleLocation As String
        Dim sNuixLicense As String
        Dim sLicenseType As String

        Dim sNMSLocation As String
        Dim sNMSUserName As String
        Dim sNMSPassword As String
        Dim sRegistryServer As String
        Dim sAppMemory As String
        Dim sSettingsFile As String
        Dim bExportSearchResults As Boolean
        Dim sExportSearchResultsLocation As String
        Dim bSearchTermFile As Boolean
        Dim bSearchTermTerm As Boolean
        Dim bExportOnly As Boolean
        Dim sExportWorkers As String
        Dim sExportWorkerMemory As String
        Dim bIncludeDiskSpace As Boolean
        Dim NuixVersionInfo As FileVersionInfo

        Dim OpenFileDialog1 As New OpenFileDialog

        With OpenFileDialog1
            .Filter = "xml|*.xml"
            .FilterIndex = 1
        End With

        If (OpenFileDialog1.ShowDialog() = Windows.Forms.DialogResult.OK) Then
            sSettingsFile = OpenFileDialog1.FileName.ToString
        Else
            Exit Sub
        End If

        bStatus = blnLoadNuixXMLSettings(sSettingsFile, sReportLocation, sReportType, sCalculateProcessingSpeeds, sShowSizeIn, sSearchTerm, sNuixLogFileLocation, sNuixConsoleLocation, sNuixLicense, sLicenseType, sNMSLocation, sNMSUserName, sNMSPassword, sRegistryServer, sAppMemory, bExportSearchResults, sExportSearchResultsLocation, bSearchTermFile, bSearchTermTerm, bExportOnly, sExportWorkers, sExportWorkerMemory, bIncludeDiskSpace)
        txtReportLocation.Text = sReportLocation
        cboReportType.Text = sReportType
        cboCalculateProcessingSpeeds.Text = sCalculateProcessingSpeeds
        cboSizeReporting.Text = sShowSizeIn
        psShowSizeIn = sShowSizeIn
        txtSearchTerm.Text = sSearchTerm
        txtNuixLogDir.Text = sNuixLogFileLocation
        txtNuixConsoleLocation.Text = sNuixConsoleLocation
        NuixVersionInfo = FileVersionInfo.GetVersionInfo(sNuixConsoleLocation)
        lblNuixConsoleVersion.Text = "Nuix Console Version: " & NuixVersionInfo.ProductMajorPart & "." & NuixVersionInfo.ProductMinorPart & "." & NuixVersionInfo.ProductBuildPart
        cboNuixLicenseType.Text = sNuixLicense
        cboLicenseType.Text = sLicenseType
        txtNMSUserName.Text = sNMSUserName
        txtNMSInfo.Text = sNMSPassword
        txtNMSAddress.Text = sNMSLocation
        txtRegistryServer.Text = sRegistryServer
        numNuixAppMemory.Value = CInt(sAppMemory)
        chkExportSearchResults.Checked = bExportSearchResults
        If bExportSearchResults = True Then
            txtExportLocation.Enabled = True
            btnExportLocation.Enabled = True
            lblExportLocation.Enabled = True
            txtExportLocation.Text = sExportSearchResultsLocation
        Else
            txtExportLocation.Enabled = False
            btnExportLocation.Enabled = False
            lblExportLocation.Enabled = False
        End If
        numExportWorkers.Value = sExportWorkers
        numExportWorkerMemory.Value = sExportWorkerMemory
        radSearchFile.Checked = bSearchTermFile
        radSearchTerm.Checked = bSearchTermTerm
        chkExportOnly.Checked = bExportOnly

    End Sub

    Private Function blnLoadNuixXMLSettings(ByVal sSettingConfigurationFile As String, ByRef sReportLocation As String, ByRef sReportType As String, ByRef sCalculateProcessingSpeeds As String, ByRef sShowSizeIn As String, ByRef sSearchTerm As String, ByRef sNuixLogFileLocation As String, ByRef sNuixConsoleLocation As String, ByRef sNuixLicense As String, ByRef sLicenseType As String, ByRef sNMSAddress As String, ByRef sNMSUserName As String, ByRef sNMSInfo As String, ByRef sRegistryServer As String, ByRef sNuixAppMemory As String, ByRef bExportSearchResults As Boolean, ByRef sExportSearchResults As String, ByRef bSearchTermFile As Boolean, ByRef bSearchTermTerm As Boolean, ByRef bExportOnly As Boolean, ByRef sExportWorkers As String, ByRef sExportWorkerMemory As String, ByRef bIncludeDiskSpace As Boolean) As Boolean
        Dim NuixXMLSettings As Xml.XmlDocument
        Dim oNUIXUCRTSettings As Xml.XmlNode

        blnLoadNuixXMLSettings = False
        NuixXMLSettings = New Xml.XmlDocument

        NuixXMLSettings.Load(sSettingConfigurationFile)

        oNUIXUCRTSettings = NuixXMLSettings.SelectSingleNode("AllNuixSettings/UCRTSettings")
        If oNUIXUCRTSettings.HasChildNodes Then
            For Each Child In oNUIXUCRTSettings.ChildNodes
                If Child.name = "ReportLocation" Then
                    sReportLocation = Child.innertext
                ElseIf Child.name = "ReportType" Then
                    sReportType = Child.innertext
                ElseIf Child.name = "CalculateProcessingSpeeds" Then
                    sCalculateProcessingSpeeds = Child.innertext
                ElseIf Child.name = "ShowSizeIn" Then
                    sShowSizeIn = Child.innertext
                ElseIf Child.name = "SearchTerm" Then
                    sSearchTerm = Child.innertext
                ElseIf Child.name = "NuixLogFileLocation" Then
                    sNuixLogFileLocation = Child.innertext
                ElseIf Child.name = "NuixConsoleLocation" Then
                    sNuixConsoleLocation = Child.innertext
                ElseIf Child.name = "NuixLicense" Then
                    sNuixLicense = Child.innertext
                ElseIf Child.name = "LicenseType" Then
                    sLicenseType = Child.innertext
                ElseIf Child.name = "NMSAddress" Then
                    sNMSAddress = Child.innertext
                ElseIf Child.name = "NMSUser" Then
                    sNMSUserName = Child.innertext
                ElseIf Child.name = "NMSInfo" Then
                    sNMSInfo = Child.innertext
                ElseIf Child.name = "RegistryServer" Then
                    sRegistryServer = Child.innertext
                ElseIf Child.name = "AppMemory" Then
                    sNuixAppMemory = Child.innertext
                ElseIf Child.name = "ExportSearchResults" Then
                    bExportSearchResults = CBool(Child.innertext)
                ElseIf Child.name = "ExportSearchResultsLocation" Then
                    sExportSearchResults = Child.innertext
                ElseIf Child.name = "SearchTermFile" Then
                    bSearchTermFile = CBool(Child.innertext)
                ElseIf Child.name = "SearchTermTerm" Then
                    bSearchTermTerm = CBool(Child.innertext)
                ElseIf Child.name = "ExportOnly" Then
                    bExportOnly = CBool(Child.innertext)
                ElseIf Child.name = "ExportWorkers" Then
                    sExportWorkers = Child.innertext
                ElseIf Child.name = "ExportWorkerMemory" Then
                    sExportWorkerMemory = Child.innertext
                ElseIf Child.name = "IncludeDiskSpace" Then
                    bIncludeDiskSpace = CBool(Child.innertext)
                End If
            Next
        End If

        blnLoadNuixXMLSettings = True
    End Function

    Private Sub btnFileLocation_Click(sender As Object, e As EventArgs) Handles btnFileLocation.Click
        Dim OpenFileDialog1 As New OpenFileDialog
        Dim sFileLocationFile As String

        With OpenFileDialog1
            .Filter = "csv|*.csv"
            .FilterIndex = 1
        End With

        If (OpenFileDialog1.ShowDialog() = Windows.Forms.DialogResult.OK) Then
            sFileLocationFile = OpenFileDialog1.FileName.ToString
            txtCaseFileLocations.Text = sFileLocationFile
        Else
            Exit Sub
        End If
    End Sub

    Private Sub radFileSystem_CheckedChanged(sender As Object, e As EventArgs) Handles radFileSystem.CheckedChanged
        If radFileSystem.Checked = True Then
            lblFileLocation.Enabled = False
            txtCaseFileLocations.Enabled = False
            btnFileLocation.Enabled = False
            panelCaseDirectory.Enabled = True
        Else
            lblFileLocation.Enabled = True
            txtCaseFileLocations.Enabled = True
            btnFileLocation.Enabled = True
            panelCaseDirectory.Enabled = False
        End If
    End Sub

    Private Sub radFile_CheckedChanged(sender As Object, e As EventArgs) Handles radFile.CheckedChanged
        If radFileSystem.Checked = True Then
            lblFileLocation.Enabled = False
            txtCaseFileLocations.Enabled = False
            btnFileLocation.Enabled = False
            panelCaseDirectory.Enabled = True
        Else
            lblFileLocation.Enabled = True
            txtCaseFileLocations.Enabled = True
            btnFileLocation.Enabled = True
            panelCaseDirectory.Enabled = False
        End If
    End Sub

    Private Sub btnStopProcessing_Click(sender As Object, e As EventArgs) Handles btnStopProcessing.Click
        ReportFormUpdate.Abort()
        For Each row In grdCaseInfo.Rows
            If row.cells("CaseGuid").value <> vbNullString Then
                If row.cells("CollectionStatus").value <> "File System and Case Data collected" Then
                    row.cells("CollectionStatus").value = "Collection Stopped..."
                End If
            End If
        Next
    End Sub

    Private Sub btnExportLocation_Click(sender As Object, e As EventArgs) Handles btnExportLocation.Click
        Dim fldrBrowserDialog As New FolderBrowserDialog

        If (fldrBrowserDialog.ShowDialog = System.Windows.Forms.DialogResult.OK) Then

            txtExportLocation.Text = fldrBrowserDialog.SelectedPath

        End If
    End Sub

    Private Sub btnSearchTermFile_Click(sender As Object, e As EventArgs) Handles btnSearchTermFile.Click
        Dim OpenFileDialog1 As New OpenFileDialog
        Dim sFileLocationFile As String

        With OpenFileDialog1
            .Filter = "csv|*.csv"
            .FilterIndex = 1
        End With

        If (OpenFileDialog1.ShowDialog() = Windows.Forms.DialogResult.OK) Then
            sFileLocationFile = OpenFileDialog1.FileName.ToString
            txtSearchTerm.Text = sFileLocationFile
        Else
            Exit Sub
        End If
    End Sub

    Private Sub radSearchTerm_CheckedChanged(sender As Object, e As EventArgs) Handles radSearchTerm.CheckedChanged
        If radSearchFile.Checked = True Then
            txtSearchTerm.Enabled = True
            btnSearchTermFile.Enabled = True
            txtSearchTerm.Text = ""
            chkExportSearchResults.Enabled = True
        ElseIf radSearchTerm.Checked = True Then
            txtSearchTerm.Enabled = True
            btnSearchTermFile.Enabled = False
            txtSearchTerm.Text = ""
            chkExportSearchResults.Enabled = True
        ElseIf radSearchFile.Checked = False And radSearchTerm.Checked = False Then
            txtSearchTerm.Enabled = False
            btnSearchTermFile.Enabled = False
            txtSearchTerm.Text = ""
            chkExportSearchResults.Enabled = False
        End If
    End Sub

    Private Sub radSearchFile_CheckedChanged(sender As Object, e As EventArgs) Handles radSearchFile.CheckedChanged
        If radSearchFile.Checked = True Then
            btnSearchTermFile.Enabled = True
            chkExportSearchResults.Checked = False
        ElseIf radSearchTerm.Checked = True Then
            btnSearchTermFile.Enabled = False
            chkExportSearchResults.Checked = False
        End If

    End Sub

    Private Sub cboSizeReporting_SelectedIndexChanged(sender As Object, e As EventArgs) Handles cboSizeReporting.SelectedIndexChanged
        Dim dlgReturn As DialogResult
        Dim sCurrentShowSizeIn As String
        Dim dCaseSizeOnDisk As Double
        Dim dCaseFileSize As Double
        Dim dCaseAuditSize As Double
        Dim dSearchSize As Double
        Dim dCurrentCaseSizeOnDisk As Double
        Dim dCurrentCaseFileSize As Double
        Dim dCurrentCaseAuditSize As Double
        Dim dCurrentSearchSize As Double

        sCurrentShowSizeIn = psShowSizeIn
        If grdCaseInfo.RowCount > 0 Then
            dlgReturn = MessageBox.Show("Would you like to recalculate the size for all fields in the grid?", "Recalculate Size Fields", MessageBoxButtons.YesNo)
            If dlgReturn = Windows.Forms.DialogResult.Yes Then
                For Each row In grdCaseInfo.Rows
                    dCurrentCaseSizeOnDisk = row.Cells("CaseSizeOnDisk").Value
                    dCurrentCaseFileSize = row.Cells("CaseFileSize").Value
                    dCurrentCaseAuditSize = row.Cells("CaseAuditSize").Value
                    dCurrentSearchSize = row.Cells("SearchSize").Value
                    If sCurrentShowSizeIn = "Bytes" Then
                        dCurrentCaseSizeOnDisk = dCurrentCaseSizeOnDisk
                        dCurrentCaseFileSize = dCurrentCaseFileSize
                        dCurrentCaseAuditSize = dCurrentCaseAuditSize
                        dCurrentSearchSize = dCurrentSearchSize
                    ElseIf sCurrentShowSizeIn = "Megabytes" Then
                        dCurrentCaseSizeOnDisk = dCurrentCaseSizeOnDisk * 1024 * 1024
                        dCurrentCaseFileSize = dCurrentCaseFileSize * 1024 * 1024
                        dCurrentCaseAuditSize = dCurrentCaseAuditSize * 1024 * 1024
                        dCurrentSearchSize = dCurrentSearchSize * 1024 * 1024
                    ElseIf sCurrentShowSizeIn = "Gigabytes" Then
                        dCurrentCaseSizeOnDisk = dCurrentCaseSizeOnDisk * 1024 * 1024 * 1024
                        dCurrentCaseFileSize = dCurrentCaseFileSize * 1024 * 1024 * 1024
                        dCurrentCaseAuditSize = dCurrentCaseAuditSize * 1024 * 1024 * 1024
                        dCurrentSearchSize = dCurrentSearchSize * 1024 * 1024 * 1024
                    End If

            Select Case cboSizeReporting.Text

                Case "Bytes"
                    row.Cells("CaseSizeOnDisk").Value = FormatNumber(dCurrentCaseSizeOnDisk, 2, , TriState.True)
                    row.Cells("CaseFileSize").Value = FormatNumber(dCurrentCaseFileSize, 2, , TriState.True)
                    row.Cells("CaseAuditSize").Value = FormatNumber(dCurrentCaseAuditSize, 2, , TriState.True)
                    row.Cells("SearchSize").Value = FormatNumber(dCurrentSearchSize, 2, TriState.True)

                Case "Megabytes"
                    row.Cells("CaseSizeOnDisk").Value = FormatNumber((dCurrentCaseSizeOnDisk / 1024 / 1024), 2, , TriState.True)
                    row.Cells("CaseFileSize").Value = FormatNumber((dCurrentCaseFileSize / 1024 / 1024), 2, , TriState.True)
                    row.Cells("CaseAuditSize").Value = FormatNumber((dCurrentCaseAuditSize / 1024 / 1024), 2, , TriState.True)
                    row.Cells("SearchSize").Value = FormatNumber((dCurrentSearchSize / 1024 / 1024), 2, , TriState.True)

                Case "Gigabytes"
                    row.Cells("CaseSizeOnDisk").Value = FormatNumber((dCurrentCaseSizeOnDisk / 1024 / 1024 / 1024), 2, , TriState.True)
                    row.Cells("CaseFileSize").Value = FormatNumber((dCurrentCaseFileSize / 1024 / 1024 / 1024), 2, , TriState.True)
                    row.Cells("CaseAuditSize").Value = FormatNumber((dCurrentCaseAuditSize / 1024 / 1024 / 1024), 2, , TriState.True)
                    row.Cells("SearchSize").Value = FormatNumber((dCurrentSearchSize / 1024 / 1024 / 1024), 2, , TriState.True)

            End Select

                Next
                psShowSizeIn = cboSizeReporting.Text
            Else
                Exit Sub
            End If

        End If
    End Sub


    Private Sub cboUpgradeCasees_SelectedIndexChanged(sender As Object, e As EventArgs) Handles cboUpgradeCasees.SelectedIndexChanged
        If cboUpgradeCasees.Text = "Upgrade and Report" Then
            cboCopyMoveCases.Text = "Copy"
            cboCopyMoveCases.Enabled = False
        ElseIf cboUpgradeCasees.Text = "Upgrade Only" Then
            cboCopyMoveCases.Text = "Copy"
            cboCopyMoveCases.Enabled = False
        Else
            cboCopyMoveCases.Text = ""
            cboCopyMoveCases.Enabled = True
        End If
    End Sub
End Class
