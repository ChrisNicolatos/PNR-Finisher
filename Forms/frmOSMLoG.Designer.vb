<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmOSMLoG
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
        Me.optPerPax = New System.Windows.Forms.RadioButton()
        Me.optPerPNR = New System.Windows.Forms.RadioButton()
        Me.fraLOGLayout = New System.Windows.Forms.GroupBox()
        Me.fraReaonForTravel = New System.Windows.Forms.GroupBox()
        Me.optOnSignersBrazil = New System.Windows.Forms.RadioButton()
        Me.optOnSigners = New System.Windows.Forms.RadioButton()
        Me.optOffSigners = New System.Windows.Forms.RadioButton()
        Me.fraPortAgent = New System.Windows.Forms.GroupBox()
        Me.lblOSMMultipleSearchSeparator = New System.Windows.Forms.Label()
        Me.txtOSMAgentsFilter = New System.Windows.Forms.TextBox()
        Me.chkNoPortAgent = New System.Windows.Forms.CheckBox()
        Me.lstPortAgent = New System.Windows.Forms.ListBox()
        Me.lblFileDestination = New System.Windows.Forms.Label()
        Me.txtFileDestination = New System.Windows.Forms.TextBox()
        Me.fileBrowser = New System.Windows.Forms.FolderBrowserDialog()
        Me.cmdFileDestination = New System.Windows.Forms.Button()
        Me.cmdCreatePDF = New System.Windows.Forms.Button()
        Me.cmdExit = New System.Windows.Forms.Button()
        Me.lblPax = New System.Windows.Forms.Label()
        Me.lblSegs = New System.Windows.Forms.Label()
        Me.txtSignedBy = New System.Windows.Forms.TextBox()
        Me.fraSignedBy = New System.Windows.Forms.GroupBox()
        Me.picSignature = New System.Windows.Forms.PictureBox()
        Me.picLogo = New System.Windows.Forms.PictureBox()
        Me.txtTelephone = New System.Windows.Forms.TextBox()
        Me.txtCountry = New System.Windows.Forms.TextBox()
        Me.txtPCArea = New System.Windows.Forms.TextBox()
        Me.txtAddress = New System.Windows.Forms.TextBox()
        Me.txtCompanyName = New System.Windows.Forms.TextBox()
        Me.txtTitle = New System.Windows.Forms.TextBox()
        Me.lstOfficeAddress = New System.Windows.Forms.ListBox()
        Me.lblPNR = New System.Windows.Forms.Label()
        Me.fraLOGLayout.SuspendLayout()
        Me.fraReaonForTravel.SuspendLayout()
        Me.fraPortAgent.SuspendLayout()
        Me.fraSignedBy.SuspendLayout()
        CType(Me.picSignature, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.picLogo, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'optPerPax
        '
        Me.optPerPax.AutoSize = True
        Me.optPerPax.Location = New System.Drawing.Point(15, 29)
        Me.optPerPax.Name = "optPerPax"
        Me.optPerPax.Size = New System.Drawing.Size(131, 17)
        Me.optPerPax.TabIndex = 0
        Me.optPerPax.TabStop = True
        Me.optPerPax.Text = "1 Letter per passenger"
        Me.optPerPax.UseVisualStyleBackColor = True
        '
        'optPerPNR
        '
        Me.optPerPNR.AutoSize = True
        Me.optPerPNR.Location = New System.Drawing.Point(15, 52)
        Me.optPerPNR.Name = "optPerPNR"
        Me.optPerPNR.Size = New System.Drawing.Size(164, 17)
        Me.optPerPNR.TabIndex = 1
        Me.optPerPNR.TabStop = True
        Me.optPerPNR.Text = "1 Letter for all the passengers"
        Me.optPerPNR.UseVisualStyleBackColor = True
        '
        'fraLOGLayout
        '
        Me.fraLOGLayout.Controls.Add(Me.optPerPax)
        Me.fraLOGLayout.Controls.Add(Me.optPerPNR)
        Me.fraLOGLayout.Location = New System.Drawing.Point(12, 12)
        Me.fraLOGLayout.Name = "fraLOGLayout"
        Me.fraLOGLayout.Size = New System.Drawing.Size(216, 86)
        Me.fraLOGLayout.TabIndex = 2
        Me.fraLOGLayout.TabStop = False
        Me.fraLOGLayout.Text = "LoG layout"
        '
        'fraReaonForTravel
        '
        Me.fraReaonForTravel.Controls.Add(Me.optOnSignersBrazil)
        Me.fraReaonForTravel.Controls.Add(Me.optOnSigners)
        Me.fraReaonForTravel.Controls.Add(Me.optOffSigners)
        Me.fraReaonForTravel.Location = New System.Drawing.Point(12, 104)
        Me.fraReaonForTravel.Name = "fraReaonForTravel"
        Me.fraReaonForTravel.Size = New System.Drawing.Size(216, 98)
        Me.fraReaonForTravel.TabIndex = 3
        Me.fraReaonForTravel.TabStop = False
        Me.fraReaonForTravel.Text = "Reason for travel"
        '
        'optOnSignersBrazil
        '
        Me.optOnSignersBrazil.AutoSize = True
        Me.optOnSignersBrazil.Location = New System.Drawing.Point(15, 69)
        Me.optOnSignersBrazil.Name = "optOnSignersBrazil"
        Me.optOnSignersBrazil.Size = New System.Drawing.Size(118, 17)
        Me.optOnSignersBrazil.TabIndex = 2
        Me.optOnSignersBrazil.TabStop = True
        Me.optOnSignersBrazil.Text = "On signers for Brazil"
        Me.optOnSignersBrazil.UseVisualStyleBackColor = True
        '
        'optOnSigners
        '
        Me.optOnSigners.AutoSize = True
        Me.optOnSigners.Location = New System.Drawing.Point(15, 19)
        Me.optOnSigners.Name = "optOnSigners"
        Me.optOnSigners.Size = New System.Drawing.Size(75, 17)
        Me.optOnSigners.TabIndex = 0
        Me.optOnSigners.TabStop = True
        Me.optOnSigners.Text = "On signers"
        Me.optOnSigners.UseVisualStyleBackColor = True
        '
        'optOffSigners
        '
        Me.optOffSigners.AutoSize = True
        Me.optOffSigners.Location = New System.Drawing.Point(15, 44)
        Me.optOffSigners.Name = "optOffSigners"
        Me.optOffSigners.Size = New System.Drawing.Size(75, 17)
        Me.optOffSigners.TabIndex = 1
        Me.optOffSigners.TabStop = True
        Me.optOffSigners.Text = "Off signers"
        Me.optOffSigners.UseVisualStyleBackColor = True
        '
        'fraPortAgent
        '
        Me.fraPortAgent.Controls.Add(Me.lblOSMMultipleSearchSeparator)
        Me.fraPortAgent.Controls.Add(Me.txtOSMAgentsFilter)
        Me.fraPortAgent.Controls.Add(Me.chkNoPortAgent)
        Me.fraPortAgent.Controls.Add(Me.lstPortAgent)
        Me.fraPortAgent.Location = New System.Drawing.Point(248, 19)
        Me.fraPortAgent.Name = "fraPortAgent"
        Me.fraPortAgent.Size = New System.Drawing.Size(418, 170)
        Me.fraPortAgent.TabIndex = 4
        Me.fraPortAgent.TabStop = False
        Me.fraPortAgent.Text = "Port Agent"
        '
        'lblOSMMultipleSearchSeparator
        '
        Me.lblOSMMultipleSearchSeparator.AutoSize = True
        Me.lblOSMMultipleSearchSeparator.Font = New System.Drawing.Font("Microsoft Sans Serif", 6.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(161, Byte))
        Me.lblOSMMultipleSearchSeparator.Location = New System.Drawing.Point(15, 42)
        Me.lblOSMMultipleSearchSeparator.Name = "lblOSMMultipleSearchSeparator"
        Me.lblOSMMultipleSearchSeparator.Size = New System.Drawing.Size(112, 9)
        Me.lblOSMMultipleSearchSeparator.TabIndex = 25
        Me.lblOSMMultipleSearchSeparator.Text = "Multiple search separated with |"
        '
        'txtOSMAgentsFilter
        '
        Me.txtOSMAgentsFilter.Location = New System.Drawing.Point(15, 19)
        Me.txtOSMAgentsFilter.Name = "txtOSMAgentsFilter"
        Me.txtOSMAgentsFilter.Size = New System.Drawing.Size(166, 20)
        Me.txtOSMAgentsFilter.TabIndex = 24
        '
        'chkNoPortAgent
        '
        Me.chkNoPortAgent.AutoSize = True
        Me.chkNoPortAgent.Location = New System.Drawing.Point(15, 142)
        Me.chkNoPortAgent.Name = "chkNoPortAgent"
        Me.chkNoPortAgent.Size = New System.Drawing.Size(93, 17)
        Me.chkNoPortAgent.TabIndex = 1
        Me.chkNoPortAgent.Text = "No Port Agent"
        Me.chkNoPortAgent.UseVisualStyleBackColor = True
        '
        'lstPortAgent
        '
        Me.lstPortAgent.FormattingEnabled = True
        Me.lstPortAgent.Location = New System.Drawing.Point(15, 51)
        Me.lstPortAgent.Name = "lstPortAgent"
        Me.lstPortAgent.Size = New System.Drawing.Size(391, 82)
        Me.lstPortAgent.TabIndex = 0
        '
        'lblFileDestination
        '
        Me.lblFileDestination.AutoSize = True
        Me.lblFileDestination.Location = New System.Drawing.Point(12, 394)
        Me.lblFileDestination.Name = "lblFileDestination"
        Me.lblFileDestination.Size = New System.Drawing.Size(79, 13)
        Me.lblFileDestination.TabIndex = 5
        Me.lblFileDestination.Text = "File Destination"
        '
        'txtFileDestination
        '
        Me.txtFileDestination.Enabled = False
        Me.txtFileDestination.Location = New System.Drawing.Point(97, 390)
        Me.txtFileDestination.Name = "txtFileDestination"
        Me.txtFileDestination.Size = New System.Drawing.Size(518, 20)
        Me.txtFileDestination.TabIndex = 6
        '
        'cmdFileDestination
        '
        Me.cmdFileDestination.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cmdFileDestination.Location = New System.Drawing.Point(621, 390)
        Me.cmdFileDestination.Name = "cmdFileDestination"
        Me.cmdFileDestination.Size = New System.Drawing.Size(45, 20)
        Me.cmdFileDestination.TabIndex = 7
        Me.cmdFileDestination.Text = ". . ."
        Me.cmdFileDestination.UseVisualStyleBackColor = True
        '
        'cmdCreatePDF
        '
        Me.cmdCreatePDF.Location = New System.Drawing.Point(252, 533)
        Me.cmdCreatePDF.Name = "cmdCreatePDF"
        Me.cmdCreatePDF.Size = New System.Drawing.Size(103, 23)
        Me.cmdCreatePDF.TabIndex = 8
        Me.cmdCreatePDF.Text = "Create PDF(s)"
        Me.cmdCreatePDF.UseVisualStyleBackColor = True
        '
        'cmdExit
        '
        Me.cmdExit.Location = New System.Drawing.Point(377, 533)
        Me.cmdExit.Name = "cmdExit"
        Me.cmdExit.Size = New System.Drawing.Size(103, 23)
        Me.cmdExit.TabIndex = 9
        Me.cmdExit.Text = "Exit"
        Me.cmdExit.UseVisualStyleBackColor = True
        '
        'lblPax
        '
        Me.lblPax.BackColor = System.Drawing.Color.Aqua
        Me.lblPax.Font = New System.Drawing.Font("Courier New", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblPax.ForeColor = System.Drawing.Color.Blue
        Me.lblPax.Location = New System.Drawing.Point(262, 429)
        Me.lblPax.Name = "lblPax"
        Me.lblPax.Size = New System.Drawing.Size(288, 89)
        Me.lblPax.TabIndex = 10
        '
        'lblSegs
        '
        Me.lblSegs.BackColor = System.Drawing.Color.Aqua
        Me.lblSegs.Font = New System.Drawing.Font("Courier New", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblSegs.ForeColor = System.Drawing.Color.Blue
        Me.lblSegs.Location = New System.Drawing.Point(595, 429)
        Me.lblSegs.Name = "lblSegs"
        Me.lblSegs.Size = New System.Drawing.Size(288, 89)
        Me.lblSegs.TabIndex = 11
        '
        'txtSignedBy
        '
        Me.txtSignedBy.Location = New System.Drawing.Point(222, 19)
        Me.txtSignedBy.Name = "txtSignedBy"
        Me.txtSignedBy.Size = New System.Drawing.Size(420, 20)
        Me.txtSignedBy.TabIndex = 13
        '
        'fraSignedBy
        '
        Me.fraSignedBy.Controls.Add(Me.picSignature)
        Me.fraSignedBy.Controls.Add(Me.picLogo)
        Me.fraSignedBy.Controls.Add(Me.txtTelephone)
        Me.fraSignedBy.Controls.Add(Me.txtCountry)
        Me.fraSignedBy.Controls.Add(Me.txtPCArea)
        Me.fraSignedBy.Controls.Add(Me.txtAddress)
        Me.fraSignedBy.Controls.Add(Me.txtCompanyName)
        Me.fraSignedBy.Controls.Add(Me.txtTitle)
        Me.fraSignedBy.Controls.Add(Me.lstOfficeAddress)
        Me.fraSignedBy.Controls.Add(Me.txtSignedBy)
        Me.fraSignedBy.Location = New System.Drawing.Point(12, 202)
        Me.fraSignedBy.Name = "fraSignedBy"
        Me.fraSignedBy.Size = New System.Drawing.Size(877, 167)
        Me.fraSignedBy.TabIndex = 14
        Me.fraSignedBy.TabStop = False
        Me.fraSignedBy.Text = "Signed By"
        '
        'picSignature
        '
        Me.picSignature.Location = New System.Drawing.Point(661, 85)
        Me.picSignature.Name = "picSignature"
        Me.picSignature.Size = New System.Drawing.Size(210, 60)
        Me.picSignature.SizeMode = System.Windows.Forms.PictureBoxSizeMode.AutoSize
        Me.picSignature.TabIndex = 24
        Me.picSignature.TabStop = False
        '
        'picLogo
        '
        Me.picLogo.Location = New System.Drawing.Point(661, 19)
        Me.picLogo.Name = "picLogo"
        Me.picLogo.Size = New System.Drawing.Size(210, 60)
        Me.picLogo.SizeMode = System.Windows.Forms.PictureBoxSizeMode.StretchImage
        Me.picLogo.TabIndex = 23
        Me.picLogo.TabStop = False
        '
        'txtTelephone
        '
        Me.txtTelephone.Location = New System.Drawing.Point(222, 139)
        Me.txtTelephone.Name = "txtTelephone"
        Me.txtTelephone.Size = New System.Drawing.Size(420, 20)
        Me.txtTelephone.TabIndex = 22
        '
        'txtCountry
        '
        Me.txtCountry.Location = New System.Drawing.Point(222, 119)
        Me.txtCountry.Name = "txtCountry"
        Me.txtCountry.Size = New System.Drawing.Size(420, 20)
        Me.txtCountry.TabIndex = 21
        '
        'txtPCArea
        '
        Me.txtPCArea.Location = New System.Drawing.Point(222, 99)
        Me.txtPCArea.Name = "txtPCArea"
        Me.txtPCArea.Size = New System.Drawing.Size(420, 20)
        Me.txtPCArea.TabIndex = 20
        '
        'txtAddress
        '
        Me.txtAddress.Location = New System.Drawing.Point(222, 79)
        Me.txtAddress.Name = "txtAddress"
        Me.txtAddress.Size = New System.Drawing.Size(420, 20)
        Me.txtAddress.TabIndex = 19
        '
        'txtCompanyName
        '
        Me.txtCompanyName.Location = New System.Drawing.Point(222, 59)
        Me.txtCompanyName.Name = "txtCompanyName"
        Me.txtCompanyName.Size = New System.Drawing.Size(420, 20)
        Me.txtCompanyName.TabIndex = 18
        '
        'txtTitle
        '
        Me.txtTitle.Location = New System.Drawing.Point(222, 39)
        Me.txtTitle.Name = "txtTitle"
        Me.txtTitle.Size = New System.Drawing.Size(420, 20)
        Me.txtTitle.TabIndex = 17
        '
        'lstOfficeAddress
        '
        Me.lstOfficeAddress.FormattingEnabled = True
        Me.lstOfficeAddress.Location = New System.Drawing.Point(11, 19)
        Me.lstOfficeAddress.Name = "lstOfficeAddress"
        Me.lstOfficeAddress.Size = New System.Drawing.Size(194, 134)
        Me.lstOfficeAddress.TabIndex = 16
        '
        'lblPNR
        '
        Me.lblPNR.BackColor = System.Drawing.Color.Aqua
        Me.lblPNR.Font = New System.Drawing.Font("Courier New", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblPNR.ForeColor = System.Drawing.Color.Blue
        Me.lblPNR.Location = New System.Drawing.Point(20, 429)
        Me.lblPNR.Name = "lblPNR"
        Me.lblPNR.Size = New System.Drawing.Size(197, 89)
        Me.lblPNR.TabIndex = 15
        '
        'frmOSMLoG
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(929, 569)
        Me.Controls.Add(Me.lblPNR)
        Me.Controls.Add(Me.fraSignedBy)
        Me.Controls.Add(Me.lblSegs)
        Me.Controls.Add(Me.lblPax)
        Me.Controls.Add(Me.cmdExit)
        Me.Controls.Add(Me.cmdCreatePDF)
        Me.Controls.Add(Me.cmdFileDestination)
        Me.Controls.Add(Me.txtFileDestination)
        Me.Controls.Add(Me.lblFileDestination)
        Me.Controls.Add(Me.fraPortAgent)
        Me.Controls.Add(Me.fraReaonForTravel)
        Me.Controls.Add(Me.fraLOGLayout)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog
        Me.Name = "frmOSMLoG"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterParent
        Me.Text = "OSM Letter of Guarantee"
        Me.fraLOGLayout.ResumeLayout(False)
        Me.fraLOGLayout.PerformLayout()
        Me.fraReaonForTravel.ResumeLayout(False)
        Me.fraReaonForTravel.PerformLayout()
        Me.fraPortAgent.ResumeLayout(False)
        Me.fraPortAgent.PerformLayout()
        Me.fraSignedBy.ResumeLayout(False)
        Me.fraSignedBy.PerformLayout()
        CType(Me.picSignature, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.picLogo, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub

    Friend WithEvents optPerPax As RadioButton
    Friend WithEvents optPerPNR As RadioButton
    Friend WithEvents fraLOGLayout As GroupBox
    Friend WithEvents fraReaonForTravel As GroupBox
    Friend WithEvents optOnSigners As RadioButton
    Friend WithEvents optOffSigners As RadioButton
    Friend WithEvents fraPortAgent As GroupBox
    Friend WithEvents lstPortAgent As ListBox
    Friend WithEvents lblFileDestination As Label
    Friend WithEvents txtFileDestination As TextBox
    Friend WithEvents fileBrowser As FolderBrowserDialog
    Friend WithEvents cmdFileDestination As Button
    Friend WithEvents cmdCreatePDF As Button
    Friend WithEvents cmdExit As Button
    Friend WithEvents lblPax As Label
    Friend WithEvents lblSegs As Label
    Friend WithEvents chkNoPortAgent As CheckBox
    Friend WithEvents txtSignedBy As TextBox
    Friend WithEvents optOnSignersBrazil As RadioButton
    Friend WithEvents lblOSMMultipleSearchSeparator As Label
    Friend WithEvents txtOSMAgentsFilter As TextBox
    Friend WithEvents fraSignedBy As GroupBox
    Friend WithEvents txtTelephone As TextBox
    Friend WithEvents txtCountry As TextBox
    Friend WithEvents txtPCArea As TextBox
    Friend WithEvents txtAddress As TextBox
    Friend WithEvents txtCompanyName As TextBox
    Friend WithEvents txtTitle As TextBox
    Friend WithEvents lstOfficeAddress As ListBox
    Friend WithEvents picSignature As PictureBox
    Friend WithEvents picLogo As PictureBox
    Friend WithEvents lblPNR As Label
End Class
