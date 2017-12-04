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
        Me.lstOSMEditVessels = New System.Windows.Forms.ListBox()
        Me.Label4 = New System.Windows.Forms.Label()
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
        Me.SuspendLayout()
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(225, 35)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(67, 13)
        Me.Label1.TabIndex = 2
        Me.Label1.Text = "Vessel name"
        '
        'txtOSMEditVessel
        '
        Me.txtOSMEditVessel.Location = New System.Drawing.Point(305, 32)
        Me.txtOSMEditVessel.Name = "txtOSMEditVessel"
        Me.txtOSMEditVessel.Size = New System.Drawing.Size(312, 20)
        Me.txtOSMEditVessel.TabIndex = 3
        '
        'cmdOSMEditExit
        '
        Me.cmdOSMEditExit.Location = New System.Drawing.Point(476, 507)
        Me.cmdOSMEditExit.Name = "cmdOSMEditExit"
        Me.cmdOSMEditExit.Size = New System.Drawing.Size(141, 23)
        Me.cmdOSMEditExit.TabIndex = 18
        Me.cmdOSMEditExit.Text = "Exit"
        Me.cmdOSMEditExit.UseVisualStyleBackColor = True
        '
        'lstOSMEditVessels
        '
        Me.lstOSMEditVessels.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
            Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.lstOSMEditVessels.DrawMode = System.Windows.Forms.DrawMode.OwnerDrawFixed
        Me.lstOSMEditVessels.FormattingEnabled = True
        Me.lstOSMEditVessels.Location = New System.Drawing.Point(12, 32)
        Me.lstOSMEditVessels.Name = "lstOSMEditVessels"
        Me.lstOSMEditVessels.Size = New System.Drawing.Size(193, 498)
        Me.lstOSMEditVessels.TabIndex = 1
        '
        'Label4
        '
        Me.Label4.AutoSize = True
        Me.Label4.Location = New System.Drawing.Point(12, 16)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(43, 13)
        Me.Label4.TabIndex = 0
        Me.Label4.Text = "Vessels"
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Location = New System.Drawing.Point(225, 269)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(53, 13)
        Me.Label2.TabIndex = 8
        Me.Label2.Text = "emails CC"
        '
        'Label3
        '
        Me.Label3.AutoSize = True
        Me.Label3.Location = New System.Drawing.Point(225, 133)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(54, 13)
        Me.Label3.TabIndex = 5
        Me.Label3.Text = "emails TO"
        '
        'lstOSMEditCCEmail
        '
        Me.lstOSMEditCCEmail.FormattingEnabled = True
        Me.lstOSMEditCCEmail.Location = New System.Drawing.Point(305, 269)
        Me.lstOSMEditCCEmail.Name = "lstOSMEditCCEmail"
        Me.lstOSMEditCCEmail.Size = New System.Drawing.Size(312, 82)
        Me.lstOSMEditCCEmail.TabIndex = 9
        '
        'lstOSMEditToEmail
        '
        Me.lstOSMEditToEmail.FormattingEnabled = True
        Me.lstOSMEditToEmail.Location = New System.Drawing.Point(305, 133)
        Me.lstOSMEditToEmail.Name = "lstOSMEditToEmail"
        Me.lstOSMEditToEmail.Size = New System.Drawing.Size(312, 82)
        Me.lstOSMEditToEmail.TabIndex = 6
        '
        'txtOSMEditEmailname
        '
        Me.txtOSMEditEmailname.Location = New System.Drawing.Point(305, 409)
        Me.txtOSMEditEmailname.Name = "txtOSMEditEmailname"
        Me.txtOSMEditEmailname.Size = New System.Drawing.Size(312, 20)
        Me.txtOSMEditEmailname.TabIndex = 13
        '
        'Label5
        '
        Me.Label5.AutoSize = True
        Me.Label5.Location = New System.Drawing.Point(225, 412)
        Me.Label5.Name = "Label5"
        Me.Label5.Size = New System.Drawing.Size(63, 13)
        Me.Label5.TabIndex = 12
        Me.Label5.Text = "eMail Name"
        '
        'txtOSMEditEmail
        '
        Me.txtOSMEditEmail.Location = New System.Drawing.Point(305, 431)
        Me.txtOSMEditEmail.Name = "txtOSMEditEmail"
        Me.txtOSMEditEmail.Size = New System.Drawing.Size(312, 20)
        Me.txtOSMEditEmail.TabIndex = 15
        '
        'Label6
        '
        Me.Label6.AutoSize = True
        Me.Label6.Location = New System.Drawing.Point(225, 434)
        Me.Label6.Name = "Label6"
        Me.Label6.Size = New System.Drawing.Size(32, 13)
        Me.Label6.TabIndex = 14
        Me.Label6.Text = "eMail"
        '
        'cmdOSMEditUpdateVessel
        '
        Me.cmdOSMEditUpdateVessel.Location = New System.Drawing.Point(476, 74)
        Me.cmdOSMEditUpdateVessel.Name = "cmdOSMEditUpdateVessel"
        Me.cmdOSMEditUpdateVessel.Size = New System.Drawing.Size(141, 23)
        Me.cmdOSMEditUpdateVessel.TabIndex = 4
        Me.cmdOSMEditUpdateVessel.Text = "Update Vessel"
        Me.cmdOSMEditUpdateVessel.UseVisualStyleBackColor = True
        '
        'cmdOSMEditUpdateEmail
        '
        Me.cmdOSMEditUpdateEmail.Location = New System.Drawing.Point(305, 478)
        Me.cmdOSMEditUpdateEmail.Name = "cmdOSMEditUpdateEmail"
        Me.cmdOSMEditUpdateEmail.Size = New System.Drawing.Size(141, 23)
        Me.cmdOSMEditUpdateEmail.TabIndex = 16
        Me.cmdOSMEditUpdateEmail.Text = "Update email"
        Me.cmdOSMEditUpdateEmail.UseVisualStyleBackColor = True
        '
        'cmdOSMAddToEmail
        '
        Me.cmdOSMAddToEmail.Location = New System.Drawing.Point(476, 104)
        Me.cmdOSMAddToEmail.Name = "cmdOSMAddToEmail"
        Me.cmdOSMAddToEmail.Size = New System.Drawing.Size(141, 23)
        Me.cmdOSMAddToEmail.TabIndex = 7
        Me.cmdOSMAddToEmail.Text = "Add TO email"
        Me.cmdOSMAddToEmail.UseVisualStyleBackColor = True
        '
        'cmdOSMAddCCEmail
        '
        Me.cmdOSMAddCCEmail.Location = New System.Drawing.Point(476, 240)
        Me.cmdOSMAddCCEmail.Name = "cmdOSMAddCCEmail"
        Me.cmdOSMAddCCEmail.Size = New System.Drawing.Size(141, 23)
        Me.cmdOSMAddCCEmail.TabIndex = 10
        Me.cmdOSMAddCCEmail.Text = "Add CC email"
        Me.cmdOSMAddCCEmail.UseVisualStyleBackColor = True
        '
        'cmdOSMEditDeleteEmail
        '
        Me.cmdOSMEditDeleteEmail.Location = New System.Drawing.Point(476, 478)
        Me.cmdOSMEditDeleteEmail.Name = "cmdOSMEditDeleteEmail"
        Me.cmdOSMEditDeleteEmail.Size = New System.Drawing.Size(141, 23)
        Me.cmdOSMEditDeleteEmail.TabIndex = 17
        Me.cmdOSMEditDeleteEmail.Text = "Delete email"
        Me.cmdOSMEditDeleteEmail.UseVisualStyleBackColor = True
        '
        'lblEmailType
        '
        Me.lblEmailType.AutoSize = True
        Me.lblEmailType.Location = New System.Drawing.Point(302, 393)
        Me.lblEmailType.Name = "lblEmailType"
        Me.lblEmailType.Size = New System.Drawing.Size(10, 13)
        Me.lblEmailType.TabIndex = 11
        Me.lblEmailType.Text = "."
        '
        'Label7
        '
        Me.Label7.AutoSize = True
        Me.Label7.Location = New System.Drawing.Point(302, 454)
        Me.Label7.Name = "Label7"
        Me.Label7.Size = New System.Drawing.Size(235, 13)
        Me.Label7.TabIndex = 19
        Me.Label7.Text = "Multiple email addresses can be separated with ;"
        '
        'cmdOSMAddVessel
        '
        Me.cmdOSMAddVessel.Location = New System.Drawing.Point(305, 74)
        Me.cmdOSMAddVessel.Name = "cmdOSMAddVessel"
        Me.cmdOSMAddVessel.Size = New System.Drawing.Size(141, 23)
        Me.cmdOSMAddVessel.TabIndex = 20
        Me.cmdOSMAddVessel.Text = "Add Vessel"
        Me.cmdOSMAddVessel.UseVisualStyleBackColor = True
        '
        'chkOSMVesselInUse
        '
        Me.chkOSMVesselInUse.AutoSize = True
        Me.chkOSMVesselInUse.Location = New System.Drawing.Point(305, 51)
        Me.chkOSMVesselInUse.Name = "chkOSMVesselInUse"
        Me.chkOSMVesselInUse.Size = New System.Drawing.Size(54, 17)
        Me.chkOSMVesselInUse.TabIndex = 21
        Me.chkOSMVesselInUse.Text = "InUse"
        Me.chkOSMVesselInUse.UseVisualStyleBackColor = True
        '
        'frmOSMVessels
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(650, 554)
        Me.Controls.Add(Me.chkOSMVesselInUse)
        Me.Controls.Add(Me.cmdOSMAddVessel)
        Me.Controls.Add(Me.Label7)
        Me.Controls.Add(Me.lblEmailType)
        Me.Controls.Add(Me.cmdOSMEditDeleteEmail)
        Me.Controls.Add(Me.cmdOSMAddCCEmail)
        Me.Controls.Add(Me.cmdOSMAddToEmail)
        Me.Controls.Add(Me.cmdOSMEditUpdateEmail)
        Me.Controls.Add(Me.cmdOSMEditUpdateVessel)
        Me.Controls.Add(Me.txtOSMEditEmail)
        Me.Controls.Add(Me.Label6)
        Me.Controls.Add(Me.txtOSMEditEmailname)
        Me.Controls.Add(Me.Label5)
        Me.Controls.Add(Me.lstOSMEditVessels)
        Me.Controls.Add(Me.Label4)
        Me.Controls.Add(Me.Label2)
        Me.Controls.Add(Me.Label3)
        Me.Controls.Add(Me.lstOSMEditCCEmail)
        Me.Controls.Add(Me.lstOSMEditToEmail)
        Me.Controls.Add(Me.cmdOSMEditExit)
        Me.Controls.Add(Me.txtOSMEditVessel)
        Me.Controls.Add(Me.Label1)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog
        Me.Name = "frmOSMVessels"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterParent
        Me.Text = "OSM Vessels"
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents txtOSMEditVessel As System.Windows.Forms.TextBox
    Friend WithEvents cmdOSMEditExit As System.Windows.Forms.Button
    Friend WithEvents lstOSMEditVessels As System.Windows.Forms.ListBox
    Friend WithEvents Label4 As System.Windows.Forms.Label
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
End Class
