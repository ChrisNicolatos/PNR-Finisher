<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmPriceOptimiser
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
        Me.Label1 = New System.Windows.Forms.Label()
        Me.dgvPNRs = New System.Windows.Forms.DataGridView()
        Me.mnuOptimiser = New System.Windows.Forms.ContextMenuStrip(Me.components)
        Me.mnuOptimiserPNR = New System.Windows.Forms.ToolStripTextBox()
        Me.ToolStripSeparator1 = New System.Windows.Forms.ToolStripSeparator()
        Me.mnuOptimiserIgnore = New System.Windows.Forms.ToolStripMenuItem()
        Me.mnuOptimiserActioned = New System.Windows.Forms.ToolStripMenuItem()
        Me.ToolStripMenuItem1 = New System.Windows.Forms.ToolStripSeparator()
        Me.mnuOptimiserOpenInGDS = New System.Windows.Forms.ToolStripMenuItem()
        Me.lblPCCUser = New System.Windows.Forms.Label()
        Me.cmdRefresh = New System.Windows.Forms.Button()
        Me.ToolStripSeparator2 = New System.Windows.Forms.ToolStripSeparator()
        Me.mnuOptimiserCopyData = New System.Windows.Forms.ToolStripMenuItem()
        CType(Me.dgvPNRs, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.mnuOptimiser.SuspendLayout()
        Me.SuspendLayout()
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label1.Location = New System.Drawing.Point(144, 9)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(178, 13)
        Me.Label1.TabIndex = 0
        Me.Label1.Text = "These PNRs can be optimised"
        '
        'dgvPNRs
        '
        Me.dgvPNRs.AllowUserToAddRows = False
        Me.dgvPNRs.AllowUserToDeleteRows = False
        Me.dgvPNRs.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
            Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.dgvPNRs.AutoSizeColumnsMode = System.Windows.Forms.DataGridViewAutoSizeColumnsMode.AllCells
        Me.dgvPNRs.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.dgvPNRs.ContextMenuStrip = Me.mnuOptimiser
        Me.dgvPNRs.EditMode = System.Windows.Forms.DataGridViewEditMode.EditProgrammatically
        Me.dgvPNRs.Location = New System.Drawing.Point(12, 56)
        Me.dgvPNRs.MultiSelect = False
        Me.dgvPNRs.Name = "dgvPNRs"
        Me.dgvPNRs.ReadOnly = True
        Me.dgvPNRs.SelectionMode = System.Windows.Forms.DataGridViewSelectionMode.CellSelect
        Me.dgvPNRs.ShowEditingIcon = False
        Me.dgvPNRs.Size = New System.Drawing.Size(1450, 274)
        Me.dgvPNRs.TabIndex = 1
        '
        'mnuOptimiser
        '
        Me.mnuOptimiser.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.mnuOptimiserPNR, Me.ToolStripSeparator1, Me.mnuOptimiserIgnore, Me.mnuOptimiserActioned, Me.ToolStripMenuItem1, Me.mnuOptimiserOpenInGDS, Me.ToolStripSeparator2, Me.mnuOptimiserCopyData})
        Me.mnuOptimiser.Name = "mnuOptimiser"
        Me.mnuOptimiser.Size = New System.Drawing.Size(261, 157)
        '
        'mnuOptimiserPNR
        '
        Me.mnuOptimiserPNR.Enabled = False
        Me.mnuOptimiserPNR.Name = "mnuOptimiserPNR"
        Me.mnuOptimiserPNR.Size = New System.Drawing.Size(200, 23)
        '
        'ToolStripSeparator1
        '
        Me.ToolStripSeparator1.Name = "ToolStripSeparator1"
        Me.ToolStripSeparator1.Size = New System.Drawing.Size(257, 6)
        '
        'mnuOptimiserIgnore
        '
        Me.mnuOptimiserIgnore.Name = "mnuOptimiserIgnore"
        Me.mnuOptimiserIgnore.Size = New System.Drawing.Size(260, 22)
        Me.mnuOptimiserIgnore.Text = "Ignore"
        '
        'mnuOptimiserActioned
        '
        Me.mnuOptimiserActioned.Name = "mnuOptimiserActioned"
        Me.mnuOptimiserActioned.Size = New System.Drawing.Size(260, 22)
        Me.mnuOptimiserActioned.Text = "Actioned"
        '
        'ToolStripMenuItem1
        '
        Me.ToolStripMenuItem1.Name = "ToolStripMenuItem1"
        Me.ToolStripMenuItem1.Size = New System.Drawing.Size(257, 6)
        '
        'mnuOptimiserOpenInGDS
        '
        Me.mnuOptimiserOpenInGDS.Name = "mnuOptimiserOpenInGDS"
        Me.mnuOptimiserOpenInGDS.Size = New System.Drawing.Size(260, 22)
        Me.mnuOptimiserOpenInGDS.Text = "Open in GDS"
        '
        'lblPCCUser
        '
        Me.lblPCCUser.AutoSize = True
        Me.lblPCCUser.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblPCCUser.Location = New System.Drawing.Point(144, 30)
        Me.lblPCCUser.Name = "lblPCCUser"
        Me.lblPCCUser.Size = New System.Drawing.Size(178, 13)
        Me.lblPCCUser.TabIndex = 2
        Me.lblPCCUser.Text = "These PNRs can be optimised"
        '
        'cmdRefresh
        '
        Me.cmdRefresh.Location = New System.Drawing.Point(12, 9)
        Me.cmdRefresh.Name = "cmdRefresh"
        Me.cmdRefresh.Size = New System.Drawing.Size(117, 34)
        Me.cmdRefresh.TabIndex = 3
        Me.cmdRefresh.Text = "Refresh"
        Me.cmdRefresh.UseVisualStyleBackColor = True
        '
        'ToolStripSeparator2
        '
        Me.ToolStripSeparator2.Name = "ToolStripSeparator2"
        Me.ToolStripSeparator2.Size = New System.Drawing.Size(257, 6)
        '
        'mnuOptimiserCopyData
        '
        Me.mnuOptimiserCopyData.Name = "mnuOptimiserCopyData"
        Me.mnuOptimiserCopyData.Size = New System.Drawing.Size(260, 22)
        Me.mnuOptimiserCopyData.Text = "Copy data"
        '
        'frmPriceOptimiser
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(1474, 342)
        Me.Controls.Add(Me.cmdRefresh)
        Me.Controls.Add(Me.lblPCCUser)
        Me.Controls.Add(Me.dgvPNRs)
        Me.Controls.Add(Me.Label1)
        Me.Name = "frmPriceOptimiser"
        Me.Text = "Price Optimisation"
        CType(Me.dgvPNRs, System.ComponentModel.ISupportInitialize).EndInit()
        Me.mnuOptimiser.ResumeLayout(False)
        Me.mnuOptimiser.PerformLayout()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub

    Friend WithEvents Label1 As Label
    Friend WithEvents dgvPNRs As DataGridView
    Friend WithEvents lblPCCUser As Label
    Friend WithEvents mnuOptimiser As ContextMenuStrip
    Friend WithEvents mnuOptimiserIgnore As ToolStripMenuItem
    Friend WithEvents mnuOptimiserActioned As ToolStripMenuItem
    Friend WithEvents ToolStripMenuItem1 As ToolStripSeparator
    Friend WithEvents mnuOptimiserOpenInGDS As ToolStripMenuItem
    Friend WithEvents mnuOptimiserPNR As ToolStripTextBox
    Friend WithEvents ToolStripSeparator1 As ToolStripSeparator
    Friend WithEvents cmdRefresh As Button
    Friend WithEvents ToolStripSeparator2 As ToolStripSeparator
    Friend WithEvents mnuOptimiserCopyData As ToolStripMenuItem
End Class
