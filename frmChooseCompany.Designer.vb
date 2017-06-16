<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmChooseCompany
#Region "Upgrade Support "
	Private Shared m_vb6FormDefInstance As frmChooseCompany
	Private Shared m_InitializingDefInstance As Boolean
	Public Shared Property DefInstance() As frmChooseCompany
		Get
			If m_vb6FormDefInstance Is Nothing OrElse m_vb6FormDefInstance.IsDisposed Then
				m_InitializingDefInstance = True
				m_vb6FormDefInstance = New frmChooseCompany()
				m_InitializingDefInstance = False
			End If
			Return m_vb6FormDefInstance
		End Get
		Set(ByVal Value As frmChooseCompany)
			m_vb6FormDefInstance = Value
		End Set
	End Property
#End Region
#Region "Windows Form Designer generated code "
	Private visualControls() As String = New String() {"components", "ToolTipMain", "cmdSetAsDefaultCompany", "cmdOpenCompany", "lstChooseCompany", "lblSelectingDefaultCompany", "lblInfo", "listBoxHelper1", "listBoxComboBoxHelper1"}
	'Required by the Windows Form Designer
	Private components As System.ComponentModel.IContainer
	Public ToolTipMain As System.Windows.Forms.ToolTip
	Public WithEvents cmdSetAsDefaultCompany As System.Windows.Forms.Button
	Public WithEvents cmdOpenCompany As System.Windows.Forms.Button
	Public WithEvents lstChooseCompany As System.Windows.Forms.ListBox
	Public WithEvents lblSelectingDefaultCompany As System.Windows.Forms.Label
	Public WithEvents lblInfo As System.Windows.Forms.Label
	Private listBoxHelper1 As UpgradeHelpers.Gui.ListBoxHelper
	Private listBoxComboBoxHelper1 As UpgradeHelpers.Gui.ListControlHelper
	'NOTE: The following procedure is required by the Windows Form Designer
	'It can be modified using the Windows Form Designer.
	'Do not modify it using the code editor.
	<System.Diagnostics.DebuggerStepThrough()> _
	 Private Sub InitializeComponent()
		Me.components = New System.ComponentModel.Container()
		Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(frmChooseCompany))
		Me.ToolTipMain = New System.Windows.Forms.ToolTip(Me.components)
		Me.cmdSetAsDefaultCompany = New System.Windows.Forms.Button()
		Me.cmdOpenCompany = New System.Windows.Forms.Button()
		Me.lstChooseCompany = New System.Windows.Forms.ListBox()
		Me.lblSelectingDefaultCompany = New System.Windows.Forms.Label()
		Me.lblInfo = New System.Windows.Forms.Label()
		Me.SuspendLayout()
		Me.listBoxHelper1 = New UpgradeHelpers.Gui.ListBoxHelper(Me.components)
		Me.listBoxComboBoxHelper1 = New UpgradeHelpers.Gui.ListControlHelper(Me.components)
		Me.listBoxComboBoxHelper1.BeginInit()
		' 
		'cmdSetAsDefaultCompany
		' 
		Me.cmdSetAsDefaultCompany.BackColor = System.Drawing.SystemColors.Control
		Me.cmdSetAsDefaultCompany.Cursor = System.Windows.Forms.Cursors.Default
		Me.cmdSetAsDefaultCompany.ForeColor = System.Drawing.SystemColors.ControlText
		Me.cmdSetAsDefaultCompany.Location = New System.Drawing.Point(16, 104)
		Me.cmdSetAsDefaultCompany.Name = "cmdSetAsDefaultCompany"
		Me.cmdSetAsDefaultCompany.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.cmdSetAsDefaultCompany.Size = New System.Drawing.Size(145, 17)
		Me.cmdSetAsDefaultCompany.TabIndex = 4
		Me.cmdSetAsDefaultCompany.Text = "Set As Default Company"
		Me.cmdSetAsDefaultCompany.TextImageRelation = System.Windows.Forms.TextImageRelation.ImageAboveText
		Me.cmdSetAsDefaultCompany.UseVisualStyleBackColor = False
		' 
		'cmdOpenCompany
		' 
		Me.cmdOpenCompany.BackColor = System.Drawing.SystemColors.Control
		Me.cmdOpenCompany.Cursor = System.Windows.Forms.Cursors.Default
		Me.cmdOpenCompany.Enabled = False
		Me.cmdOpenCompany.ForeColor = System.Drawing.SystemColors.ControlText
		Me.cmdOpenCompany.Location = New System.Drawing.Point(112, 144)
		Me.cmdOpenCompany.Name = "cmdOpenCompany"
		Me.cmdOpenCompany.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.cmdOpenCompany.Size = New System.Drawing.Size(97, 25)
		Me.cmdOpenCompany.TabIndex = 1
		Me.cmdOpenCompany.Text = "Open"
		Me.cmdOpenCompany.TextImageRelation = System.Windows.Forms.TextImageRelation.ImageAboveText
		' 
		'lstChooseCompany
		' 
		Me.lstChooseCompany.BackColor = System.Drawing.SystemColors.Window
		Me.lstChooseCompany.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
		Me.lstChooseCompany.CausesValidation = True
		Me.lstChooseCompany.Cursor = System.Windows.Forms.Cursors.Default
		Me.lstChooseCompany.Enabled = True
		Me.lstChooseCompany.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Bold Or System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, 0)
		Me.lstChooseCompany.ForeColor = System.Drawing.SystemColors.WindowText
		Me.lstChooseCompany.IntegralHeight = True
		Me.lstChooseCompany.Location = New System.Drawing.Point(8, 24)
		Me.lstChooseCompany.MultiColumn = False
		Me.lstChooseCompany.Name = "lstChooseCompany"
		Me.lstChooseCompany.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.lstChooseCompany.Size = New System.Drawing.Size(321, 71)
		Me.lstChooseCompany.Sorted = False
		Me.lstChooseCompany.TabIndex = 0
		Me.lstChooseCompany.TabStop = True
		Me.lstChooseCompany.Visible = True
		Me.lstChooseCompany.Items.AddRange(New Object() {"DrummondPrinting", "FrazzledAndBedazzled"})
		' 
		'lblSelectingDefaultCompany
		' 
		Me.lblSelectingDefaultCompany.BackColor = System.Drawing.SystemColors.Control
		Me.lblSelectingDefaultCompany.BorderStyle = System.Windows.Forms.BorderStyle.None
		Me.lblSelectingDefaultCompany.Cursor = System.Windows.Forms.Cursors.Default
		Me.lblSelectingDefaultCompany.ForeColor = System.Drawing.SystemColors.ControlText
		Me.lblSelectingDefaultCompany.Location = New System.Drawing.Point(16, 184)
		Me.lblSelectingDefaultCompany.Name = "lblSelectingDefaultCompany"
		Me.lblSelectingDefaultCompany.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.lblSelectingDefaultCompany.Size = New System.Drawing.Size(305, 17)
		Me.lblSelectingDefaultCompany.TabIndex = 3
		Me.lblSelectingDefaultCompany.Text = "Selecting Default Company in 5 4 3 2 1 ..."
		' 
		'lblInfo
		' 
		Me.lblInfo.BackColor = System.Drawing.SystemColors.Control
		Me.lblInfo.BorderStyle = System.Windows.Forms.BorderStyle.None
		Me.lblInfo.Cursor = System.Windows.Forms.Cursors.Default
		Me.lblInfo.ForeColor = System.Drawing.SystemColors.ControlText
		Me.lblInfo.Location = New System.Drawing.Point(8, 8)
		Me.lblInfo.Name = "lblInfo"
		Me.lblInfo.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.lblInfo.Size = New System.Drawing.Size(321, 17)
		Me.lblInfo.TabIndex = 2
		Me.lblInfo.Text = "Select A Company To Open:"
		' 
		'frmChooseCompany
		' 
		Me.AutoScaleDimensions = New System.Drawing.SizeF(6, 13)
		Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
		Me.BackColor = System.Drawing.SystemColors.Control
		Me.ClientSize = New System.Drawing.Size(338, 213)
		Me.Controls.Add(Me.cmdSetAsDefaultCompany)
		Me.Controls.Add(Me.cmdOpenCompany)
		Me.Controls.Add(Me.lstChooseCompany)
		Me.Controls.Add(Me.lblSelectingDefaultCompany)
		Me.Controls.Add(Me.lblInfo)
		Me.Cursor = System.Windows.Forms.Cursors.Default
		Me.Location = New System.Drawing.Point(4, 23)
		Me.MaximizeBox = True
		Me.MinimizeBox = True
		Me.Name = "frmChooseCompany"
		Me.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.Text = "Choose Company - DBPM"
		Me.listBoxComboBoxHelper1.SetItemData(Me.lstChooseCompany, New Integer() {0, 0})
		listBoxHelper1.SetSelectionMode(Me.lstChooseCompany, System.Windows.Forms.SelectionMode.One)
		Me.listBoxComboBoxHelper1.EndInit()
		Me.ResumeLayout(False)
	End Sub
	Sub ReLoadForm(ByVal addEvents As Boolean)
		If addEvents Then
			AddHandler MyBase.Load, AddressOf Me.frmChooseCompany_Load
			AddHandler MyBase.Closed, AddressOf Me.frmChooseCompany_Closed
		End If
	End Sub
#End Region
End Class