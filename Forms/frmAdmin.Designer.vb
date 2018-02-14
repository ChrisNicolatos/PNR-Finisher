<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmAdmin
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
        Me.TabControl1 = New System.Windows.Forms.TabControl()
        Me.GDSBackOffice = New System.Windows.Forms.TabPage()
        Me.lblBackOffice = New System.Windows.Forms.Label()
        Me.lblGDS = New System.Windows.Forms.Label()
        Me.dgvBackOffice = New System.Windows.Forms.DataGridView()
        Me.dgvGDS = New System.Windows.Forms.DataGridView()
        Me.PCC = New System.Windows.Forms.TabPage()
        Me.Users = New System.Windows.Forms.TabPage()
        Me.AmadeusReferences = New System.Windows.Forms.TabPage()
        Me.PNRCloseOffEntries = New System.Windows.Forms.TabPage()
        Me.ClientCorporateDeals = New System.Windows.Forms.TabPage()
        Me.ClientAlerts = New System.Windows.Forms.TabPage()
        Me.TabControl1.SuspendLayout()
        Me.GDSBackOffice.SuspendLayout()
        CType(Me.dgvBackOffice, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.dgvGDS, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'TabControl1
        '
        Me.TabControl1.Controls.Add(Me.GDSBackOffice)
        Me.TabControl1.Controls.Add(Me.PCC)
        Me.TabControl1.Controls.Add(Me.Users)
        Me.TabControl1.Controls.Add(Me.AmadeusReferences)
        Me.TabControl1.Controls.Add(Me.PNRCloseOffEntries)
        Me.TabControl1.Controls.Add(Me.ClientCorporateDeals)
        Me.TabControl1.Controls.Add(Me.ClientAlerts)
        Me.TabControl1.Dock = System.Windows.Forms.DockStyle.Fill
        Me.TabControl1.Location = New System.Drawing.Point(0, 0)
        Me.TabControl1.Name = "TabControl1"
        Me.TabControl1.SelectedIndex = 0
        Me.TabControl1.Size = New System.Drawing.Size(1040, 513)
        Me.TabControl1.TabIndex = 0
        '
        'GDSBackOffice
        '
        Me.GDSBackOffice.Controls.Add(Me.lblBackOffice)
        Me.GDSBackOffice.Controls.Add(Me.lblGDS)
        Me.GDSBackOffice.Controls.Add(Me.dgvBackOffice)
        Me.GDSBackOffice.Controls.Add(Me.dgvGDS)
        Me.GDSBackOffice.Location = New System.Drawing.Point(4, 22)
        Me.GDSBackOffice.Name = "GDSBackOffice"
        Me.GDSBackOffice.Padding = New System.Windows.Forms.Padding(3)
        Me.GDSBackOffice.Size = New System.Drawing.Size(1032, 487)
        Me.GDSBackOffice.TabIndex = 2
        Me.GDSBackOffice.Text = "GDS/Back Office"
        Me.GDSBackOffice.UseVisualStyleBackColor = True
        '
        'lblBackOffice
        '
        Me.lblBackOffice.Location = New System.Drawing.Point(58, 212)
        Me.lblBackOffice.Name = "lblBackOffice"
        Me.lblBackOffice.Size = New System.Drawing.Size(514, 13)
        Me.lblBackOffice.TabIndex = 3
        Me.lblBackOffice.Text = "Back Office"
        Me.lblBackOffice.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'lblGDS
        '
        Me.lblGDS.Location = New System.Drawing.Point(58, 18)
        Me.lblGDS.Name = "lblGDS"
        Me.lblGDS.Size = New System.Drawing.Size(514, 13)
        Me.lblGDS.TabIndex = 2
        Me.lblGDS.Text = "GDS"
        Me.lblGDS.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'dgvBackOffice
        '
        Me.dgvBackOffice.AllowUserToAddRows = False
        Me.dgvBackOffice.AllowUserToDeleteRows = False
        Me.dgvBackOffice.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.dgvBackOffice.EditMode = System.Windows.Forms.DataGridViewEditMode.EditProgrammatically
        Me.dgvBackOffice.Location = New System.Drawing.Point(58, 228)
        Me.dgvBackOffice.Name = "dgvBackOffice"
        Me.dgvBackOffice.RowHeadersVisible = False
        Me.dgvBackOffice.Size = New System.Drawing.Size(514, 119)
        Me.dgvBackOffice.TabIndex = 1
        '
        'dgvGDS
        '
        Me.dgvGDS.AllowUserToAddRows = False
        Me.dgvGDS.AllowUserToDeleteRows = False
        Me.dgvGDS.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.dgvGDS.EditMode = System.Windows.Forms.DataGridViewEditMode.EditProgrammatically
        Me.dgvGDS.Location = New System.Drawing.Point(58, 34)
        Me.dgvGDS.Name = "dgvGDS"
        Me.dgvGDS.RowHeadersVisible = False
        Me.dgvGDS.ShowEditingIcon = False
        Me.dgvGDS.Size = New System.Drawing.Size(514, 119)
        Me.dgvGDS.TabIndex = 0
        '
        'PCC
        '
        Me.PCC.Location = New System.Drawing.Point(4, 22)
        Me.PCC.Name = "PCC"
        Me.PCC.Padding = New System.Windows.Forms.Padding(3)
        Me.PCC.Size = New System.Drawing.Size(1032, 487)
        Me.PCC.TabIndex = 0
        Me.PCC.Text = "PCC"
        Me.PCC.UseVisualStyleBackColor = True
        '
        'Users
        '
        Me.Users.Location = New System.Drawing.Point(4, 22)
        Me.Users.Name = "Users"
        Me.Users.Padding = New System.Windows.Forms.Padding(3)
        Me.Users.Size = New System.Drawing.Size(1032, 487)
        Me.Users.TabIndex = 1
        Me.Users.Text = "Users"
        Me.Users.UseVisualStyleBackColor = True
        '
        'AmadeusReferences
        '
        Me.AmadeusReferences.Location = New System.Drawing.Point(4, 22)
        Me.AmadeusReferences.Name = "AmadeusReferences"
        Me.AmadeusReferences.Size = New System.Drawing.Size(1032, 487)
        Me.AmadeusReferences.TabIndex = 3
        Me.AmadeusReferences.Text = "Amadeus References"
        Me.AmadeusReferences.UseVisualStyleBackColor = True
        '
        'PNRCloseOffEntries
        '
        Me.PNRCloseOffEntries.Location = New System.Drawing.Point(4, 22)
        Me.PNRCloseOffEntries.Name = "PNRCloseOffEntries"
        Me.PNRCloseOffEntries.Size = New System.Drawing.Size(1032, 487)
        Me.PNRCloseOffEntries.TabIndex = 4
        Me.PNRCloseOffEntries.Text = "PNR Close Off Entries"
        Me.PNRCloseOffEntries.UseVisualStyleBackColor = True
        '
        'ClientCorporateDeals
        '
        Me.ClientCorporateDeals.Location = New System.Drawing.Point(4, 22)
        Me.ClientCorporateDeals.Name = "ClientCorporateDeals"
        Me.ClientCorporateDeals.Size = New System.Drawing.Size(1032, 487)
        Me.ClientCorporateDeals.TabIndex = 5
        Me.ClientCorporateDeals.Text = "Client Corporate Deals"
        Me.ClientCorporateDeals.UseVisualStyleBackColor = True
        '
        'ClientAlerts
        '
        Me.ClientAlerts.Location = New System.Drawing.Point(4, 22)
        Me.ClientAlerts.Name = "ClientAlerts"
        Me.ClientAlerts.Size = New System.Drawing.Size(1032, 487)
        Me.ClientAlerts.TabIndex = 6
        Me.ClientAlerts.Text = "Client Alerts"
        Me.ClientAlerts.UseVisualStyleBackColor = True
        '
        'frmAdmin
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(1040, 513)
        Me.Controls.Add(Me.TabControl1)
        Me.Name = "frmAdmin"
        Me.Text = "PNR Finisher Administration"
        Me.TabControl1.ResumeLayout(False)
        Me.GDSBackOffice.ResumeLayout(False)
        CType(Me.dgvBackOffice, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.dgvGDS, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)

    End Sub

    Friend WithEvents TabControl1 As TabControl
    Friend WithEvents GDSBackOffice As TabPage
    Friend WithEvents PCC As TabPage
    Friend WithEvents Users As TabPage
    Friend WithEvents AmadeusReferences As TabPage
    Friend WithEvents PNRCloseOffEntries As TabPage
    Friend WithEvents ClientCorporateDeals As TabPage
    Friend WithEvents ClientAlerts As TabPage
    Friend WithEvents dgvGDS As DataGridView
    Friend WithEvents dgvBackOffice As DataGridView
    Friend WithEvents lblBackOffice As Label
    Friend WithEvents lblGDS As Label
End Class
