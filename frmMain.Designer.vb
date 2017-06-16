<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmMain
#Region "Upgrade Support "
	Private Shared m_vb6FormDefInstance As frmMain
	Private Shared m_InitializingDefInstance As Boolean
	Public Shared Property DefInstance() As frmMain
		Get
			If m_vb6FormDefInstance Is Nothing OrElse m_vb6FormDefInstance.IsDisposed Then
				m_InitializingDefInstance = True
				m_vb6FormDefInstance = New frmMain()
				m_InitializingDefInstance = False
			End If
			Return m_vb6FormDefInstance
		End Get
		Set(ByVal Value As frmMain)
			m_vb6FormDefInstance = Value
		End Set
	End Property
#End Region
#Region "Windows Form Designer generated code "
	Private visualControls() As String = New String() {"components", "ToolTipMain", "_mnuAbout_0", "MainMenu1", "timerRunDBPMProcesses", "cmdError", "cmdPutJobInvoicesIn", "cmdImportQBTables", "cmdRefreshQBTables", "cmdRunDocFinderRefresh", "cmdRunDocFinder", "lstConversionProgress", "chkReloadCustOnly", "cmdNumberDropShips", "cmdPutLeadLoadIn", "cmdLockComputer", "chkProcessCustOnly", "timerMonitorInterval", "cmdAddFranchiseOffices", "Command2", "cmdInsertMaxBillToIntoQB", "chkPauseProcessing", "cmdApplyCreditMomos", "cmdInvPmt", "chkSeeProcessing", "Command1", "cmdAddItems", "cmdAR1Invoice", "cmdFixCustomerAddress", "cmdFixVendorAddress", "cmdAccountToAccount", "cmdImportInvoiceLines", "cmdTryCustomerExport", "cmdCustomerToCustomer", "cmdAddSalesReps", "cmdAddTerms", "cmdVendorToVendor", "cmdDeleteOldFiles", "lblVersion", "lblPaused", "_Label1_2", "_Label1_1", "_Label1_0", "lblListboxStatus", "lblStatus", "Label1", "mnuAbout", "listBoxHelper1"}
	'Required by the Windows Form Designer
	Private components As System.ComponentModel.IContainer
	Public ToolTipMain As System.Windows.Forms.ToolTip
	Private WithEvents _mnuAbout_0 As System.Windows.Forms.ToolStripMenuItem
	Public WithEvents MainMenu1 As System.Windows.Forms.MenuStrip
    Public WithEvents cmdPutJobInvoicesIn As System.Windows.Forms.Button
    Public WithEvents cmdRefreshQBTables As System.Windows.Forms.Button
	Public WithEvents cmdRunDocFinderRefresh As System.Windows.Forms.Button
    Public WithEvents lstConversionProgress As System.Windows.Forms.ListBox
    Public WithEvents cmdNumberDropShips As System.Windows.Forms.Button
    Public WithEvents chkPauseProcessing As System.Windows.Forms.CheckBox
    Public WithEvents chkSeeProcessing As System.Windows.Forms.CheckBox
    Public WithEvents lblVersion As System.Windows.Forms.Label
	Public WithEvents lblPaused As System.Windows.Forms.Label
    Public WithEvents lblListboxStatus As System.Windows.Forms.Label
	Public WithEvents lblStatus As System.Windows.Forms.Label
	Public Label1(2) As System.Windows.Forms.Label
	Public mnuAbout(0) As System.Windows.Forms.ToolStripItem

	'NOTE: The following procedure is required by the Windows Form Designer
	'It can be modified using the Windows Form Designer.
	'Do not modify it using the code editor.
	<System.Diagnostics.DebuggerStepThrough()> _
	 Private Sub InitializeComponent()
        Me.components = New System.ComponentModel.Container()
        Me.ToolTipMain = New System.Windows.Forms.ToolTip(Me.components)
        Me.MainMenu1 = New System.Windows.Forms.MenuStrip()
        Me._mnuAbout_0 = New System.Windows.Forms.ToolStripMenuItem()
        Me.cmdPutJobInvoicesIn = New System.Windows.Forms.Button()
        Me.cmdRefreshQBTables = New System.Windows.Forms.Button()
        Me.cmdRunDocFinderRefresh = New System.Windows.Forms.Button()
        Me.lstConversionProgress = New System.Windows.Forms.ListBox()
        Me.cmdNumberDropShips = New System.Windows.Forms.Button()
        Me.chkPauseProcessing = New System.Windows.Forms.CheckBox()
        Me.chkSeeProcessing = New System.Windows.Forms.CheckBox()
        Me.lblVersion = New System.Windows.Forms.Label()
        Me.lblPaused = New System.Windows.Forms.Label()
        Me.lblListboxStatus = New System.Windows.Forms.Label()
        Me.lblStatus = New System.Windows.Forms.Label()
        Me.btnRunDBPMProcesses = New System.Windows.Forms.Button()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.Label3 = New System.Windows.Forms.Label()
        Me.Label4 = New System.Windows.Forms.Label()
        Me.Panel1 = New System.Windows.Forms.Panel()
        Me.Panel2 = New System.Windows.Forms.Panel()
        Me.chkPauseForErrors = New System.Windows.Forms.CheckBox()
        Me.Label5 = New System.Windows.Forms.Label()
        Me.txtOutput = New System.Windows.Forms.TextBox()
        Me.MainMenu1.SuspendLayout()
        Me.Panel1.SuspendLayout()
        Me.Panel2.SuspendLayout()
        Me.SuspendLayout()
        '
        'MainMenu1
        '
        Me.MainMenu1.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me._mnuAbout_0})
        Me.MainMenu1.Location = New System.Drawing.Point(0, 0)
        Me.MainMenu1.Name = "MainMenu1"
        Me.MainMenu1.Size = New System.Drawing.Size(1141, 24)
        Me.MainMenu1.TabIndex = 39
        '
        '_mnuAbout_0
        '
        Me._mnuAbout_0.Name = "_mnuAbout_0"
        Me._mnuAbout_0.Size = New System.Drawing.Size(52, 20)
        Me._mnuAbout_0.Text = "About"
        '
        'cmdPutJobInvoicesIn
        '
        Me.cmdPutJobInvoicesIn.BackColor = System.Drawing.SystemColors.Control
        Me.cmdPutJobInvoicesIn.Cursor = System.Windows.Forms.Cursors.Default
        Me.cmdPutJobInvoicesIn.ForeColor = System.Drawing.SystemColors.ControlText
        Me.cmdPutJobInvoicesIn.Location = New System.Drawing.Point(852, 276)
        Me.cmdPutJobInvoicesIn.Name = "cmdPutJobInvoicesIn"
        Me.cmdPutJobInvoicesIn.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.cmdPutJobInvoicesIn.Size = New System.Drawing.Size(148, 20)
        Me.cmdPutJobInvoicesIn.TabIndex = 30
        Me.cmdPutJobInvoicesIn.Text = "Put Job Invoices In"
        Me.cmdPutJobInvoicesIn.TextImageRelation = System.Windows.Forms.TextImageRelation.ImageAboveText
        Me.cmdPutJobInvoicesIn.UseVisualStyleBackColor = False
        '
        'cmdRefreshQBTables
        '
        Me.cmdRefreshQBTables.BackColor = System.Drawing.SystemColors.Control
        Me.cmdRefreshQBTables.Cursor = System.Windows.Forms.Cursors.Default
        Me.cmdRefreshQBTables.ForeColor = System.Drawing.SystemColors.ControlText
        Me.cmdRefreshQBTables.Location = New System.Drawing.Point(852, 300)
        Me.cmdRefreshQBTables.Name = "cmdRefreshQBTables"
        Me.cmdRefreshQBTables.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.cmdRefreshQBTables.Size = New System.Drawing.Size(148, 20)
        Me.cmdRefreshQBTables.TabIndex = 21
        Me.cmdRefreshQBTables.Text = "Freshen QB Tables"
        Me.cmdRefreshQBTables.TextImageRelation = System.Windows.Forms.TextImageRelation.ImageAboveText
        Me.cmdRefreshQBTables.UseVisualStyleBackColor = False
        '
        'cmdRunDocFinderRefresh
        '
        Me.cmdRunDocFinderRefresh.BackColor = System.Drawing.SystemColors.Control
        Me.cmdRunDocFinderRefresh.Cursor = System.Windows.Forms.Cursors.Default
        Me.cmdRunDocFinderRefresh.ForeColor = System.Drawing.SystemColors.ControlText
        Me.cmdRunDocFinderRefresh.Location = New System.Drawing.Point(852, 348)
        Me.cmdRunDocFinderRefresh.Name = "cmdRunDocFinderRefresh"
        Me.cmdRunDocFinderRefresh.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.cmdRunDocFinderRefresh.Size = New System.Drawing.Size(148, 20)
        Me.cmdRunDocFinderRefresh.TabIndex = 35
        Me.cmdRunDocFinderRefresh.Text = "Doc Finder Refresh"
        Me.cmdRunDocFinderRefresh.TextImageRelation = System.Windows.Forms.TextImageRelation.ImageAboveText
        Me.cmdRunDocFinderRefresh.UseVisualStyleBackColor = False
        '
        'lstConversionProgress
        '
        Me.lstConversionProgress.BackColor = System.Drawing.SystemColors.Window
        Me.lstConversionProgress.Cursor = System.Windows.Forms.Cursors.Default
        Me.lstConversionProgress.ForeColor = System.Drawing.SystemColors.WindowText
        Me.lstConversionProgress.HorizontalScrollbar = True
        Me.lstConversionProgress.Location = New System.Drawing.Point(8, 80)
        Me.lstConversionProgress.Name = "lstConversionProgress"
        Me.lstConversionProgress.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.lstConversionProgress.ScrollAlwaysVisible = True
        Me.lstConversionProgress.Size = New System.Drawing.Size(817, 511)
        Me.lstConversionProgress.TabIndex = 4
        '
        'cmdNumberDropShips
        '
        Me.cmdNumberDropShips.BackColor = System.Drawing.SystemColors.Control
        Me.cmdNumberDropShips.Cursor = System.Windows.Forms.Cursors.Default
        Me.cmdNumberDropShips.ForeColor = System.Drawing.SystemColors.ControlText
        Me.cmdNumberDropShips.Location = New System.Drawing.Point(852, 324)
        Me.cmdNumberDropShips.Name = "cmdNumberDropShips"
        Me.cmdNumberDropShips.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.cmdNumberDropShips.Size = New System.Drawing.Size(148, 20)
        Me.cmdNumberDropShips.TabIndex = 32
        Me.cmdNumberDropShips.Text = "Number Drop Ships"
        Me.cmdNumberDropShips.TextImageRelation = System.Windows.Forms.TextImageRelation.ImageAboveText
        Me.cmdNumberDropShips.UseVisualStyleBackColor = False
        '
        'chkPauseProcessing
        '
        Me.chkPauseProcessing.BackColor = System.Drawing.SystemColors.ControlDark
        Me.chkPauseProcessing.Cursor = System.Windows.Forms.Cursors.Default
        Me.chkPauseProcessing.ForeColor = System.Drawing.SystemColors.ButtonFace
        Me.chkPauseProcessing.Location = New System.Drawing.Point(8, 128)
        Me.chkPauseProcessing.Name = "chkPauseProcessing"
        Me.chkPauseProcessing.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.chkPauseProcessing.Size = New System.Drawing.Size(153, 17)
        Me.chkPauseProcessing.TabIndex = 22
        Me.chkPauseProcessing.Text = "Pause Timed Processing"
        Me.chkPauseProcessing.UseVisualStyleBackColor = False
        Me.chkPauseProcessing.Visible = False
        '
        'chkSeeProcessing
        '
        Me.chkSeeProcessing.BackColor = System.Drawing.SystemColors.ControlDark
        Me.chkSeeProcessing.Checked = True
        Me.chkSeeProcessing.CheckState = System.Windows.Forms.CheckState.Checked
        Me.chkSeeProcessing.Cursor = System.Windows.Forms.Cursors.Default
        Me.chkSeeProcessing.ForeColor = System.Drawing.SystemColors.ButtonFace
        Me.chkSeeProcessing.Location = New System.Drawing.Point(8, 40)
        Me.chkSeeProcessing.Name = "chkSeeProcessing"
        Me.chkSeeProcessing.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.chkSeeProcessing.Size = New System.Drawing.Size(121, 17)
        Me.chkSeeProcessing.TabIndex = 18
        Me.chkSeeProcessing.Text = "See Processing:"
        Me.chkSeeProcessing.UseVisualStyleBackColor = False
        '
        'lblVersion
        '
        Me.lblVersion.BackColor = System.Drawing.SystemColors.Control
        Me.lblVersion.Cursor = System.Windows.Forms.Cursors.Default
        Me.lblVersion.ForeColor = System.Drawing.SystemColors.ControlText
        Me.lblVersion.Location = New System.Drawing.Point(712, 600)
        Me.lblVersion.Name = "lblVersion"
        Me.lblVersion.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.lblVersion.Size = New System.Drawing.Size(105, 17)
        Me.lblVersion.TabIndex = 38
        Me.lblVersion.Text = "3.0.1"
        Me.lblVersion.TextAlign = System.Drawing.ContentAlignment.TopRight
        '
        'lblPaused
        '
        Me.lblPaused.BackColor = System.Drawing.SystemColors.Control
        Me.lblPaused.Cursor = System.Windows.Forms.Cursors.Default
        Me.lblPaused.Font = New System.Drawing.Font("Microsoft Sans Serif", 18.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblPaused.ForeColor = System.Drawing.Color.Red
        Me.lblPaused.Location = New System.Drawing.Point(320, 24)
        Me.lblPaused.Name = "lblPaused"
        Me.lblPaused.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.lblPaused.Size = New System.Drawing.Size(105, 25)
        Me.lblPaused.TabIndex = 26
        Me.lblPaused.Text = "Paused"
        Me.lblPaused.Visible = False
        '
        'lblListboxStatus
        '
        Me.lblListboxStatus.BackColor = System.Drawing.SystemColors.Control
        Me.lblListboxStatus.Cursor = System.Windows.Forms.Cursors.Default
        Me.lblListboxStatus.ForeColor = System.Drawing.SystemColors.ControlText
        Me.lblListboxStatus.Location = New System.Drawing.Point(8, 600)
        Me.lblListboxStatus.Name = "lblListboxStatus"
        Me.lblListboxStatus.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.lblListboxStatus.Size = New System.Drawing.Size(689, 17)
        Me.lblListboxStatus.TabIndex = 1
        '
        'lblStatus
        '
        Me.lblStatus.BackColor = System.Drawing.SystemColors.Control
        Me.lblStatus.Cursor = System.Windows.Forms.Cursors.Default
        Me.lblStatus.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblStatus.ForeColor = System.Drawing.SystemColors.ControlText
        Me.lblStatus.Location = New System.Drawing.Point(8, 24)
        Me.lblStatus.Name = "lblStatus"
        Me.lblStatus.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.lblStatus.Size = New System.Drawing.Size(521, 25)
        Me.lblStatus.TabIndex = 0
        Me.lblStatus.Text = "Waiting..."
        '
        'btnRunDBPMProcesses
        '
        Me.btnRunDBPMProcesses.Location = New System.Drawing.Point(852, 164)
        Me.btnRunDBPMProcesses.Name = "btnRunDBPMProcesses"
        Me.btnRunDBPMProcesses.Size = New System.Drawing.Size(148, 28)
        Me.btnRunDBPMProcesses.TabIndex = 40
        Me.btnRunDBPMProcesses.Text = "Run All DBPM Processes"
        Me.btnRunDBPMProcesses.UseVisualStyleBackColor = True
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Font = New System.Drawing.Font("Microsoft Sans Serif", 13.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label2.ForeColor = System.Drawing.SystemColors.ButtonFace
        Me.Label2.Location = New System.Drawing.Point(8, 8)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(162, 22)
        Me.Label2.TabIndex = 41
        Me.Label2.Text = "Manual Processing"
        '
        'Label3
        '
        Me.Label3.BackColor = System.Drawing.SystemColors.ActiveCaption
        Me.Label3.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label3.Location = New System.Drawing.Point(852, 228)
        Me.Label3.Name = "Label3"
        Me.Label3.Padding = New System.Windows.Forms.Padding(2)
        Me.Label3.Size = New System.Drawing.Size(148, 40)
        Me.Label3.TabIndex = 42
        Me.Label3.Text = "Executes Single Set of Processes"
        Me.Label3.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'Label4
        '
        Me.Label4.BackColor = System.Drawing.SystemColors.ActiveCaption
        Me.Label4.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label4.Location = New System.Drawing.Point(852, 120)
        Me.Label4.Name = "Label4"
        Me.Label4.Padding = New System.Windows.Forms.Padding(2)
        Me.Label4.Size = New System.Drawing.Size(148, 40)
        Me.Label4.TabIndex = 43
        Me.Label4.Text = "Executes All Processes"
        Me.Label4.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'Panel1
        '
        Me.Panel1.BackColor = System.Drawing.SystemColors.ControlDark
        Me.Panel1.Controls.Add(Me.Label2)
        Me.Panel1.Location = New System.Drawing.Point(836, 80)
        Me.Panel1.Name = "Panel1"
        Me.Panel1.Size = New System.Drawing.Size(180, 296)
        Me.Panel1.TabIndex = 44
        '
        'Panel2
        '
        Me.Panel2.BackColor = System.Drawing.SystemColors.ControlDark
        Me.Panel2.Controls.Add(Me.chkPauseForErrors)
        Me.Panel2.Controls.Add(Me.Label5)
        Me.Panel2.Controls.Add(Me.chkSeeProcessing)
        Me.Panel2.Controls.Add(Me.chkPauseProcessing)
        Me.Panel2.Location = New System.Drawing.Point(836, 388)
        Me.Panel2.Name = "Panel2"
        Me.Panel2.Size = New System.Drawing.Size(180, 204)
        Me.Panel2.TabIndex = 45
        '
        'chkPauseForErrors
        '
        Me.chkPauseForErrors.BackColor = System.Drawing.SystemColors.ControlDark
        Me.chkPauseForErrors.Cursor = System.Windows.Forms.Cursors.Default
        Me.chkPauseForErrors.ForeColor = System.Drawing.SystemColors.ButtonFace
        Me.chkPauseForErrors.Location = New System.Drawing.Point(8, 60)
        Me.chkPauseForErrors.Name = "chkPauseForErrors"
        Me.chkPauseForErrors.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.chkPauseForErrors.Size = New System.Drawing.Size(121, 17)
        Me.chkPauseForErrors.TabIndex = 42
        Me.chkPauseForErrors.Text = "Pause for Errors"
        Me.chkPauseForErrors.UseVisualStyleBackColor = False
        '
        'Label5
        '
        Me.Label5.AutoSize = True
        Me.Label5.Font = New System.Drawing.Font("Microsoft Sans Serif", 13.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label5.ForeColor = System.Drawing.SystemColors.ButtonFace
        Me.Label5.Location = New System.Drawing.Point(8, 8)
        Me.Label5.Name = "Label5"
        Me.Label5.Size = New System.Drawing.Size(72, 22)
        Me.Label5.TabIndex = 41
        Me.Label5.Text = "Options"
        '
        'txtOutput
        '
        Me.txtOutput.Location = New System.Drawing.Point(8, 628)
        Me.txtOutput.Multiline = True
        Me.txtOutput.Name = "txtOutput"
        Me.txtOutput.Size = New System.Drawing.Size(816, 96)
        Me.txtOutput.TabIndex = 46
        '
        'frmMain
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.BackColor = System.Drawing.SystemColors.Control
        Me.ClientSize = New System.Drawing.Size(1141, 733)
        Me.Controls.Add(Me.txtOutput)
        Me.Controls.Add(Me.Panel2)
        Me.Controls.Add(Me.Label4)
        Me.Controls.Add(Me.Label3)
        Me.Controls.Add(Me.btnRunDBPMProcesses)
        Me.Controls.Add(Me.cmdPutJobInvoicesIn)
        Me.Controls.Add(Me.cmdRefreshQBTables)
        Me.Controls.Add(Me.cmdRunDocFinderRefresh)
        Me.Controls.Add(Me.lstConversionProgress)
        Me.Controls.Add(Me.cmdNumberDropShips)
        Me.Controls.Add(Me.lblVersion)
        Me.Controls.Add(Me.lblPaused)
        Me.Controls.Add(Me.lblListboxStatus)
        Me.Controls.Add(Me.lblStatus)
        Me.Controls.Add(Me.MainMenu1)
        Me.Controls.Add(Me.Panel1)
        Me.Cursor = System.Windows.Forms.Cursors.Default
        Me.Location = New System.Drawing.Point(201, 204)
        Me.Name = "frmMain"
        Me.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Text = "DBPM Server - Main"
        Me.MainMenu1.ResumeLayout(False)
        Me.MainMenu1.PerformLayout()
        Me.Panel1.ResumeLayout(False)
        Me.Panel1.PerformLayout()
        Me.Panel2.ResumeLayout(False)
        Me.Panel2.PerformLayout()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Sub ReLoadForm(ByVal addEvents As Boolean)
        InitializemnuAbout()
        InitializeLabel1()
        If addEvents Then
            AddHandler MyBase.Load, AddressOf Me.frmMain_Load
            AddHandler MyBase.Closed, AddressOf Me.frmMain_Closed
        End If
    End Sub
    Sub InitializemnuAbout()
        ReDim mnuAbout(0)
        Me.mnuAbout(0) = _mnuAbout_0
    End Sub
    Sub InitializeLabel1()
        ReDim Label1(2)
        'Me.Label1(2) = _Label1_2
        'Me.Label1(1) = _Label1_1
        'Me.Label1(0) = _Label1_0
    End Sub
    Friend WithEvents btnRunDBPMProcesses As System.Windows.Forms.Button
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents Label4 As System.Windows.Forms.Label
    Friend WithEvents Panel1 As System.Windows.Forms.Panel
    Friend WithEvents Panel2 As System.Windows.Forms.Panel
    Friend WithEvents Label5 As System.Windows.Forms.Label
    Friend WithEvents txtOutput As System.Windows.Forms.TextBox
    Public WithEvents chkPauseForErrors As System.Windows.Forms.CheckBox
#End Region
End Class