<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmOSMVessels
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
        Me.Label1 = New System.Windows.Forms.Label()
        Me.txtOSMEditVessel = New System.Windows.Forms.TextBox()
        Me.cmdOSMEditExit = New System.Windows.Forms.Button()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.Label3 = New System.Windows.Forms.Label()
        Me.lstOSMEditCCEmail = New System.Windows.Forms.ListBox()
        Me.lstOSMEditToEmail = New System.Windows.Forms.ListBox()
        Me.txtOSMEditEmailname = New System.Windows.Forms.TextBox()
        Me.Label5 = New System.Windows.Forms.Label()
        Me.txtOSMEditEmail = New System.Windows.Forms.TextBox()
        Me.Label6 = New System.Windows.Forms.Label()
        Me.cmdOSMEditUpdateVessel = New System.Windows.Forms.Button()
        Me.cmdOSMEditUpdateEmail = New System.Windows.Forms.Button()
        Me.cmdOSMAddToEmail = New System.Windows.Forms.Button()
        Me.cmdOSMAddCCEmail = New System.Windows.Forms.Button()
        Me.cmdOSMEditDeleteEmail = New System.Windows.Forms.Button()
        Me.lblEmailType = New System.Windows.Forms.Label()
        Me.Label7 = New System.Windows.Forms.Label()
        Me.cmdOSMAddVessel = New System.Windows.Forms.Button()
        Me.chkOSMVesselInUse = New System.Windows.Forms.CheckBox()
        Me.lstOSMEditVessels = New System.Windows.Forms.ListBox()
        Me.SplitContainer1 = New System.Windows.Forms.SplitContainer()
        Me.SplitContainer2 = New System.Windows.Forms.SplitContainer()
        Me.cmdAddVesselGroup = New System.Windows.Forms.Button()
        Me.Label8 = New System.Windows.Forms.Label()
        Me.SplitContainer3 = New System.Windows.Forms.SplitContainer()
        Me.SplitContainer4 = New System.Windows.Forms.SplitContainer()
        Me.lstVesselGroup = New System.Windows.Forms.CheckedListBox()
        CType(Me.SplitContainer1, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SplitContainer1.Panel1.SuspendLayout()
        Me.SplitContainer1.Panel2.SuspendLayout()
        Me.SplitContainer1.SuspendLayout()
        CType(Me.SplitContainer2, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SplitContainer2.Panel1.SuspendLayout()
        Me.SplitContainer2.Panel2.SuspendLayout()
        Me.SplitContainer2.SuspendLayout()
        CType(Me.SplitContainer3, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SplitContainer3.Panel1.SuspendLayout()
        Me.SplitContainer3.Panel2.SuspendLayout()
        Me.SplitContainer3.SuspendLayout()
        CType(Me.SplitContainer4, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SplitContainer4.Panel1.SuspendLayout()
        Me.SplitContainer4.Panel2.SuspendLayout()
        Me.SplitContainer4.SuspendLayout()
        Me.SuspendLayout()
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(15, 42)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(67, 13)
        Me.Label1.TabIndex = 2
        Me.Label1.Text = "Vessel name"
        '
        'txtOSMEditVessel
        '
        Me.txtOSMEditVessel.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.txtOSMEditVessel.Location = New System.Drawing.Point(95, 39)
        Me.txtOSMEditVessel.Name = "txtOSMEditVessel"
        Me.txtOSMEditVessel.Size = New System.Drawing.Size(523, 20)
        Me.txtOSMEditVessel.TabIndex = 3
        '
        'cmdOSMEditExit
        '
        Me.cmdOSMEditExit.Location = New System.Drawing.Point(440, 89)
        Me.cmdOSMEditExit.Name = "cmdOSMEditExit"
        Me.cmdOSMEditExit.Size = New System.Drawing.Size(141, 23)
        Me.cmdOSMEditExit.TabIndex = 18
        Me.cmdOSMEditExit.Text = "Exit"
        Me.cmdOSMEditExit.UseVisualStyleBackColor = True
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Location = New System.Drawing.Point(14, 42)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(53, 13)
        Me.Label2.TabIndex = 8
        Me.Label2.Text = "emails CC"
        '
        'Label3
        '
        Me.Label3.AutoSize = True
        Me.Label3.Location = New System.Drawing.Point(14, 37)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(54, 13)
        Me.Label3.TabIndex = 5
        Me.Label3.Text = "emails TO"
        '
        'lstOSMEditCCEmail
        '
        Me.lstOSMEditCCEmail.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.lstOSMEditCCEmail.FormattingEnabled = True
        Me.lstOSMEditCCEmail.Location = New System.Drawing.Point(95, 42)
        Me.lstOSMEditCCEmail.Name = "lstOSMEditCCEmail"
        Me.lstOSMEditCCEmail.Size = New System.Drawing.Size(523, 82)
        Me.lstOSMEditCCEmail.TabIndex = 9
        '
        'lstOSMEditToEmail
        '
        Me.lstOSMEditToEmail.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.lstOSMEditToEmail.FormattingEnabled = True
        Me.lstOSMEditToEmail.Location = New System.Drawing.Point(95, 37)
        Me.lstOSMEditToEmail.Name = "lstOSMEditToEmail"
        Me.lstOSMEditToEmail.Size = New System.Drawing.Size(523, 82)
        Me.lstOSMEditToEmail.TabIndex = 6
        '
        'txtOSMEditEmailname
        '
        Me.txtOSMEditEmailname.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.txtOSMEditEmailname.Location = New System.Drawing.Point(95, 20)
        Me.txtOSMEditEmailname.Name = "txtOSMEditEmailname"
        Me.txtOSMEditEmailname.Size = New System.Drawing.Size(523, 20)
        Me.txtOSMEditEmailname.TabIndex = 13
        '
        'Label5
        '
        Me.Label5.AutoSize = True
        Me.Label5.Location = New System.Drawing.Point(15, 23)
        Me.Label5.Name = "Label5"
        Me.Label5.Size = New System.Drawing.Size(63, 13)
        Me.Label5.TabIndex = 12
        Me.Label5.Text = "eMail Name"
        '
        'txtOSMEditEmail
        '
        Me.txtOSMEditEmail.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.txtOSMEditEmail.Location = New System.Drawing.Point(95, 42)
        Me.txtOSMEditEmail.Name = "txtOSMEditEmail"
        Me.txtOSMEditEmail.Size = New System.Drawing.Size(523, 20)
        Me.txtOSMEditEmail.TabIndex = 15
        '
        'Label6
        '
        Me.Label6.AutoSize = True
        Me.Label6.Location = New System.Drawing.Point(15, 45)
        Me.Label6.Name = "Label6"
        Me.Label6.Size = New System.Drawing.Size(32, 13)
        Me.Label6.TabIndex = 14
        Me.Label6.Text = "eMail"
        '
        'cmdOSMEditUpdateVessel
        '
        Me.cmdOSMEditUpdateVessel.Location = New System.Drawing.Point(377, 10)
        Me.cmdOSMEditUpdateVessel.Name = "cmdOSMEditUpdateVessel"
        Me.cmdOSMEditUpdateVessel.Size = New System.Drawing.Size(106, 23)
        Me.cmdOSMEditUpdateVessel.TabIndex = 4
        Me.cmdOSMEditUpdateVessel.Text = "Update Vessel"
        Me.cmdOSMEditUpdateVessel.UseVisualStyleBackColor = True
        '
        'cmdOSMEditUpdateEmail
        '
        Me.cmdOSMEditUpdateEmail.Location = New System.Drawing.Point(95, 89)
        Me.cmdOSMEditUpdateEmail.Name = "cmdOSMEditUpdateEmail"
        Me.cmdOSMEditUpdateEmail.Size = New System.Drawing.Size(141, 23)
        Me.cmdOSMEditUpdateEmail.TabIndex = 16
        Me.cmdOSMEditUpdateEmail.Text = "Update email"
        Me.cmdOSMEditUpdateEmail.UseVisualStyleBackColor = True
        '
        'cmdOSMAddToEmail
        '
        Me.cmdOSMAddToEmail.Location = New System.Drawing.Point(95, 8)
        Me.cmdOSMAddToEmail.Name = "cmdOSMAddToEmail"
        Me.cmdOSMAddToEmail.Size = New System.Drawing.Size(141, 23)
        Me.cmdOSMAddToEmail.TabIndex = 7
        Me.cmdOSMAddToEmail.Text = "Add TO email"
        Me.cmdOSMAddToEmail.UseVisualStyleBackColor = True
        '
        'cmdOSMAddCCEmail
        '
        Me.cmdOSMAddCCEmail.Location = New System.Drawing.Point(95, 13)
        Me.cmdOSMAddCCEmail.Name = "cmdOSMAddCCEmail"
        Me.cmdOSMAddCCEmail.Size = New System.Drawing.Size(141, 23)
        Me.cmdOSMAddCCEmail.TabIndex = 10
        Me.cmdOSMAddCCEmail.Text = "Add CC email"
        Me.cmdOSMAddCCEmail.UseVisualStyleBackColor = True
        '
        'cmdOSMEditDeleteEmail
        '
        Me.cmdOSMEditDeleteEmail.Location = New System.Drawing.Point(266, 89)
        Me.cmdOSMEditDeleteEmail.Name = "cmdOSMEditDeleteEmail"
        Me.cmdOSMEditDeleteEmail.Size = New System.Drawing.Size(141, 23)
        Me.cmdOSMEditDeleteEmail.TabIndex = 17
        Me.cmdOSMEditDeleteEmail.Text = "Delete email"
        Me.cmdOSMEditDeleteEmail.UseVisualStyleBackColor = True
        '
        'lblEmailType
        '
        Me.lblEmailType.AutoSize = True
        Me.lblEmailType.Location = New System.Drawing.Point(92, 4)
        Me.lblEmailType.Name = "lblEmailType"
        Me.lblEmailType.Size = New System.Drawing.Size(10, 13)
        Me.lblEmailType.TabIndex = 11
        Me.lblEmailType.Text = "."
        '
        'Label7
        '
        Me.Label7.AutoSize = True
        Me.Label7.Location = New System.Drawing.Point(95, 65)
        Me.Label7.Name = "Label7"
        Me.Label7.Size = New System.Drawing.Size(235, 13)
        Me.Label7.TabIndex = 19
        Me.Label7.Text = "Multiple email addresses can be separated with ;"
        '
        'cmdOSMAddVessel
        '
        Me.cmdOSMAddVessel.Location = New System.Drawing.Point(95, 10)
        Me.cmdOSMAddVessel.Name = "cmdOSMAddVessel"
        Me.cmdOSMAddVessel.Size = New System.Drawing.Size(118, 23)
        Me.cmdOSMAddVessel.TabIndex = 20
        Me.cmdOSMAddVessel.Text = "Add Vessel"
        Me.cmdOSMAddVessel.UseVisualStyleBackColor = True
        '
        'chkOSMVesselInUse
        '
        Me.chkOSMVesselInUse.AutoSize = True
        Me.chkOSMVesselInUse.Location = New System.Drawing.Point(95, 59)
        Me.chkOSMVesselInUse.Name = "chkOSMVesselInUse"
        Me.chkOSMVesselInUse.Size = New System.Drawing.Size(54, 17)
        Me.chkOSMVesselInUse.TabIndex = 21
        Me.chkOSMVesselInUse.Text = "InUse"
        Me.chkOSMVesselInUse.UseVisualStyleBackColor = True
        '
        'lstOSMEditVessels
        '
        Me.lstOSMEditVessels.Dock = System.Windows.Forms.DockStyle.Fill
        Me.lstOSMEditVessels.DrawMode = System.Windows.Forms.DrawMode.OwnerDrawFixed
        Me.lstOSMEditVessels.FormattingEnabled = True
        Me.lstOSMEditVessels.Location = New System.Drawing.Point(0, 0)
        Me.lstOSMEditVessels.Name = "lstOSMEditVessels"
        Me.lstOSMEditVessels.Size = New System.Drawing.Size(323, 595)
        Me.lstOSMEditVessels.TabIndex = 2
        '
        'SplitContainer1
        '
        Me.SplitContainer1.Dock = System.Windows.Forms.DockStyle.Fill
        Me.SplitContainer1.Location = New System.Drawing.Point(0, 0)
        Me.SplitContainer1.Name = "SplitContainer1"
        '
        'SplitContainer1.Panel1
        '
        Me.SplitContainer1.Panel1.Controls.Add(Me.lstOSMEditVessels)
        '
        'SplitContainer1.Panel2
        '
        Me.SplitContainer1.Panel2.Controls.Add(Me.SplitContainer2)
        Me.SplitContainer1.Size = New System.Drawing.Size(957, 595)
        Me.SplitContainer1.SplitterDistance = 323
        Me.SplitContainer1.TabIndex = 22
        '
        'SplitContainer2
        '
        Me.SplitContainer2.Dock = System.Windows.Forms.DockStyle.Fill
        Me.SplitContainer2.Location = New System.Drawing.Point(0, 0)
        Me.SplitContainer2.Name = "SplitContainer2"
        Me.SplitContainer2.Orientation = System.Windows.Forms.Orientation.Horizontal
        '
        'SplitContainer2.Panel1
        '
        Me.SplitContainer2.Panel1.BackColor = System.Drawing.Color.FromArgb(CType(CType(192, Byte), Integer), CType(CType(255, Byte), Integer), CType(CType(255, Byte), Integer))
        Me.SplitContainer2.Panel1.Controls.Add(Me.lstVesselGroup)
        Me.SplitContainer2.Panel1.Controls.Add(Me.cmdAddVesselGroup)
        Me.SplitContainer2.Panel1.Controls.Add(Me.Label8)
        Me.SplitContainer2.Panel1.Controls.Add(Me.cmdOSMAddVessel)
        Me.SplitContainer2.Panel1.Controls.Add(Me.chkOSMVesselInUse)
        Me.SplitContainer2.Panel1.Controls.Add(Me.Label1)
        Me.SplitContainer2.Panel1.Controls.Add(Me.txtOSMEditVessel)
        Me.SplitContainer2.Panel1.Controls.Add(Me.cmdOSMEditUpdateVessel)
        '
        'SplitContainer2.Panel2
        '
        Me.SplitContainer2.Panel2.Controls.Add(Me.SplitContainer3)
        Me.SplitContainer2.Size = New System.Drawing.Size(630, 595)
        Me.SplitContainer2.SplitterDistance = 184
        Me.SplitContainer2.TabIndex = 0
        '
        'cmdAddVesselGroup
        '
        Me.cmdAddVesselGroup.Location = New System.Drawing.Point(236, 10)
        Me.cmdAddVesselGroup.Name = "cmdAddVesselGroup"
        Me.cmdAddVesselGroup.Size = New System.Drawing.Size(118, 23)
        Me.cmdAddVesselGroup.TabIndex = 24
        Me.cmdAddVesselGroup.Text = "Add Vessel Group"
        Me.cmdAddVesselGroup.UseVisualStyleBackColor = True
        '
        'Label8
        '
        Me.Label8.AutoSize = True
        Me.Label8.Location = New System.Drawing.Point(15, 86)
        Me.Label8.Name = "Label8"
        Me.Label8.Size = New System.Drawing.Size(70, 13)
        Me.Label8.TabIndex = 22
        Me.Label8.Text = "Vessel Group"
        '
        'SplitContainer3
        '
        Me.SplitContainer3.Dock = System.Windows.Forms.DockStyle.Fill
        Me.SplitContainer3.Location = New System.Drawing.Point(0, 0)
        Me.SplitContainer3.Name = "SplitContainer3"
        Me.SplitContainer3.Orientation = System.Windows.Forms.Orientation.Horizontal
        '
        'SplitContainer3.Panel1
        '
        Me.SplitContainer3.Panel1.BackColor = System.Drawing.Color.FromArgb(CType(CType(192, Byte), Integer), CType(CType(255, Byte), Integer), CType(CType(192, Byte), Integer))
        Me.SplitContainer3.Panel1.Controls.Add(Me.lstOSMEditToEmail)
        Me.SplitContainer3.Panel1.Controls.Add(Me.Label3)
        Me.SplitContainer3.Panel1.Controls.Add(Me.cmdOSMAddToEmail)
        '
        'SplitContainer3.Panel2
        '
        Me.SplitContainer3.Panel2.Controls.Add(Me.SplitContainer4)
        Me.SplitContainer3.Size = New System.Drawing.Size(630, 407)
        Me.SplitContainer3.SplitterDistance = 129
        Me.SplitContainer3.TabIndex = 0
        '
        'SplitContainer4
        '
        Me.SplitContainer4.Dock = System.Windows.Forms.DockStyle.Fill
        Me.SplitContainer4.Location = New System.Drawing.Point(0, 0)
        Me.SplitContainer4.Name = "SplitContainer4"
        Me.SplitContainer4.Orientation = System.Windows.Forms.Orientation.Horizontal
        '
        'SplitContainer4.Panel1
        '
        Me.SplitContainer4.Panel1.BackColor = System.Drawing.Color.FromArgb(CType(CType(255, Byte), Integer), CType(CType(224, Byte), Integer), CType(CType(192, Byte), Integer))
        Me.SplitContainer4.Panel1.Controls.Add(Me.lstOSMEditCCEmail)
        Me.SplitContainer4.Panel1.Controls.Add(Me.Label2)
        Me.SplitContainer4.Panel1.Controls.Add(Me.cmdOSMAddCCEmail)
        '
        'SplitContainer4.Panel2
        '
        Me.SplitContainer4.Panel2.BackColor = System.Drawing.Color.FromArgb(CType(CType(255, Byte), Integer), CType(CType(255, Byte), Integer), CType(CType(192, Byte), Integer))
        Me.SplitContainer4.Panel2.Controls.Add(Me.txtOSMEditEmail)
        Me.SplitContainer4.Panel2.Controls.Add(Me.Label7)
        Me.SplitContainer4.Panel2.Controls.Add(Me.cmdOSMEditExit)
        Me.SplitContainer4.Panel2.Controls.Add(Me.lblEmailType)
        Me.SplitContainer4.Panel2.Controls.Add(Me.Label5)
        Me.SplitContainer4.Panel2.Controls.Add(Me.cmdOSMEditDeleteEmail)
        Me.SplitContainer4.Panel2.Controls.Add(Me.txtOSMEditEmailname)
        Me.SplitContainer4.Panel2.Controls.Add(Me.cmdOSMEditUpdateEmail)
        Me.SplitContainer4.Panel2.Controls.Add(Me.Label6)
        Me.SplitContainer4.Size = New System.Drawing.Size(630, 274)
        Me.SplitContainer4.SplitterDistance = 131
        Me.SplitContainer4.TabIndex = 0
        '
        'lstVesselGroup
        '
        Me.lstVesselGroup.FormattingEnabled = True
        Me.lstVesselGroup.Location = New System.Drawing.Point(95, 86)
        Me.lstVesselGroup.Name = "lstVesselGroup"
        Me.lstVesselGroup.Size = New System.Drawing.Size(523, 94)
        Me.lstVesselGroup.TabIndex = 25
        '
        'frmOSMVessels
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(957, 595)
        Me.Controls.Add(Me.SplitContainer1)
        Me.Name = "frmOSMVessels"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterParent
        Me.Text = "OSM Vessels"
        Me.SplitContainer1.Panel1.ResumeLayout(False)
        Me.SplitContainer1.Panel2.ResumeLayout(False)
        CType(Me.SplitContainer1, System.ComponentModel.ISupportInitialize).EndInit()
        Me.SplitContainer1.ResumeLayout(False)
        Me.SplitContainer2.Panel1.ResumeLayout(False)
        Me.SplitContainer2.Panel1.PerformLayout()
        Me.SplitContainer2.Panel2.ResumeLayout(False)
        CType(Me.SplitContainer2, System.ComponentModel.ISupportInitialize).EndInit()
        Me.SplitContainer2.ResumeLayout(False)
        Me.SplitContainer3.Panel1.ResumeLayout(False)
        Me.SplitContainer3.Panel1.PerformLayout()
        Me.SplitContainer3.Panel2.ResumeLayout(False)
        CType(Me.SplitContainer3, System.ComponentModel.ISupportInitialize).EndInit()
        Me.SplitContainer3.ResumeLayout(False)
        Me.SplitContainer4.Panel1.ResumeLayout(False)
        Me.SplitContainer4.Panel1.PerformLayout()
        Me.SplitContainer4.Panel2.ResumeLayout(False)
        Me.SplitContainer4.Panel2.PerformLayout()
        CType(Me.SplitContainer4, System.ComponentModel.ISupportInitialize).EndInit()
        Me.SplitContainer4.ResumeLayout(False)
        Me.ResumeLayout(False)

    End Sub
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents txtOSMEditVessel As System.Windows.Forms.TextBox
    Friend WithEvents cmdOSMEditExit As System.Windows.Forms.Button
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents lstOSMEditCCEmail As System.Windows.Forms.ListBox
    Friend WithEvents lstOSMEditToEmail As System.Windows.Forms.ListBox
    Friend WithEvents txtOSMEditEmailname As System.Windows.Forms.TextBox
    Friend WithEvents Label5 As System.Windows.Forms.Label
    Friend WithEvents txtOSMEditEmail As System.Windows.Forms.TextBox
    Friend WithEvents Label6 As System.Windows.Forms.Label
    Friend WithEvents cmdOSMEditUpdateVessel As System.Windows.Forms.Button
    Friend WithEvents cmdOSMEditUpdateEmail As System.Windows.Forms.Button
    Friend WithEvents cmdOSMAddToEmail As System.Windows.Forms.Button
    Friend WithEvents cmdOSMAddCCEmail As System.Windows.Forms.Button
    Friend WithEvents cmdOSMEditDeleteEmail As System.Windows.Forms.Button
    Friend WithEvents lblEmailType As System.Windows.Forms.Label
    Friend WithEvents Label7 As System.Windows.Forms.Label
    Friend WithEvents cmdOSMAddVessel As System.Windows.Forms.Button
    Friend WithEvents chkOSMVesselInUse As System.Windows.Forms.CheckBox
    Friend WithEvents lstOSMEditVessels As ListBox
    Friend WithEvents SplitContainer1 As SplitContainer
    Friend WithEvents SplitContainer2 As SplitContainer
    Friend WithEvents cmdAddVesselGroup As Button
    Friend WithEvents Label8 As Label
    Friend WithEvents SplitContainer3 As SplitContainer
    Friend WithEvents SplitContainer4 As SplitContainer
    Friend WithEvents lstVesselGroup As CheckedListBox
End Class
