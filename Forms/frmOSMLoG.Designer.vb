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
        Me.GroupBox1 = New System.Windows.Forms.GroupBox()
        Me.GroupBox2 = New System.Windows.Forms.GroupBox()
        Me.optOnSigners = New System.Windows.Forms.RadioButton()
        Me.optOffSigners = New System.Windows.Forms.RadioButton()
        Me.GroupBox3 = New System.Windows.Forms.GroupBox()
        Me.chkNoPortAgent = New System.Windows.Forms.CheckBox()
        Me.lstPortAgent = New System.Windows.Forms.ListBox()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.txtFileDestination = New System.Windows.Forms.TextBox()
        Me.fileBrowser = New System.Windows.Forms.FolderBrowserDialog()
        Me.cmdFileDestination = New System.Windows.Forms.Button()
        Me.cmdCreatePDF = New System.Windows.Forms.Button()
        Me.cmdExit = New System.Windows.Forms.Button()
        Me.lblPax = New System.Windows.Forms.Label()
        Me.lblSegs = New System.Windows.Forms.Label()
        Me.txtSignedBy = New System.Windows.Forms.TextBox()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.GroupBox1.SuspendLayout()
        Me.GroupBox2.SuspendLayout()
        Me.GroupBox3.SuspendLayout()
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
        'GroupBox1
        '
        Me.GroupBox1.Controls.Add(Me.optPerPax)
        Me.GroupBox1.Controls.Add(Me.optPerPNR)
        Me.GroupBox1.Location = New System.Drawing.Point(12, 12)
        Me.GroupBox1.Name = "GroupBox1"
        Me.GroupBox1.Size = New System.Drawing.Size(216, 86)
        Me.GroupBox1.TabIndex = 2
        Me.GroupBox1.TabStop = False
        Me.GroupBox1.Text = "LoG layout"
        '
        'GroupBox2
        '
        Me.GroupBox2.Controls.Add(Me.optOnSigners)
        Me.GroupBox2.Controls.Add(Me.optOffSigners)
        Me.GroupBox2.Location = New System.Drawing.Point(12, 104)
        Me.GroupBox2.Name = "GroupBox2"
        Me.GroupBox2.Size = New System.Drawing.Size(216, 86)
        Me.GroupBox2.TabIndex = 3
        Me.GroupBox2.TabStop = False
        Me.GroupBox2.Text = "Reason for travel"
        '
        'optOnSigners
        '
        Me.optOnSigners.AutoSize = True
        Me.optOnSigners.Location = New System.Drawing.Point(15, 29)
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
        Me.optOffSigners.Location = New System.Drawing.Point(15, 52)
        Me.optOffSigners.Name = "optOffSigners"
        Me.optOffSigners.Size = New System.Drawing.Size(75, 17)
        Me.optOffSigners.TabIndex = 1
        Me.optOffSigners.TabStop = True
        Me.optOffSigners.Text = "Off signers"
        Me.optOffSigners.UseVisualStyleBackColor = True
        '
        'GroupBox3
        '
        Me.GroupBox3.Controls.Add(Me.chkNoPortAgent)
        Me.GroupBox3.Controls.Add(Me.lstPortAgent)
        Me.GroupBox3.Location = New System.Drawing.Point(248, 19)
        Me.GroupBox3.Name = "GroupBox3"
        Me.GroupBox3.Size = New System.Drawing.Size(418, 170)
        Me.GroupBox3.TabIndex = 4
        Me.GroupBox3.TabStop = False
        Me.GroupBox3.Text = "Port Agent"
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
        Me.lstPortAgent.Location = New System.Drawing.Point(15, 25)
        Me.lstPortAgent.Name = "lstPortAgent"
        Me.lstPortAgent.Size = New System.Drawing.Size(391, 108)
        Me.lstPortAgent.TabIndex = 0
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(12, 256)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(79, 13)
        Me.Label1.TabIndex = 5
        Me.Label1.Text = "File Destination"
        '
        'txtFileDestination
        '
        Me.txtFileDestination.Enabled = False
        Me.txtFileDestination.Location = New System.Drawing.Point(97, 252)
        Me.txtFileDestination.Name = "txtFileDestination"
        Me.txtFileDestination.Size = New System.Drawing.Size(518, 20)
        Me.txtFileDestination.TabIndex = 6
        '
        'cmdFileDestination
        '
        Me.cmdFileDestination.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cmdFileDestination.Location = New System.Drawing.Point(621, 252)
        Me.cmdFileDestination.Name = "cmdFileDestination"
        Me.cmdFileDestination.Size = New System.Drawing.Size(45, 20)
        Me.cmdFileDestination.TabIndex = 7
        Me.cmdFileDestination.Text = ". . ."
        Me.cmdFileDestination.UseVisualStyleBackColor = True
        '
        'cmdCreatePDF
        '
        Me.cmdCreatePDF.Location = New System.Drawing.Point(252, 395)
        Me.cmdCreatePDF.Name = "cmdCreatePDF"
        Me.cmdCreatePDF.Size = New System.Drawing.Size(103, 23)
        Me.cmdCreatePDF.TabIndex = 8
        Me.cmdCreatePDF.Text = "Create PDF(s)"
        Me.cmdCreatePDF.UseVisualStyleBackColor = True
        '
        'cmdExit
        '
        Me.cmdExit.Location = New System.Drawing.Point(377, 395)
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
        Me.lblPax.Location = New System.Drawing.Point(20, 291)
        Me.lblPax.Name = "lblPax"
        Me.lblPax.Size = New System.Drawing.Size(288, 89)
        Me.lblPax.TabIndex = 10
        '
        'lblSegs
        '
        Me.lblSegs.BackColor = System.Drawing.Color.Aqua
        Me.lblSegs.Font = New System.Drawing.Font("Courier New", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblSegs.ForeColor = System.Drawing.Color.Blue
        Me.lblSegs.Location = New System.Drawing.Point(366, 291)
        Me.lblSegs.Name = "lblSegs"
        Me.lblSegs.Size = New System.Drawing.Size(288, 89)
        Me.lblSegs.TabIndex = 11
        '
        'txtSignedBy
        '
        Me.txtSignedBy.Location = New System.Drawing.Point(97, 208)
        Me.txtSignedBy.Name = "txtSignedBy"
        Me.txtSignedBy.Size = New System.Drawing.Size(518, 20)
        Me.txtSignedBy.TabIndex = 13
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Location = New System.Drawing.Point(12, 212)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(54, 13)
        Me.Label2.TabIndex = 12
        Me.Label2.Text = "Signed by"
        '
        'frmOSMLoG
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(694, 462)
        Me.Controls.Add(Me.txtSignedBy)
        Me.Controls.Add(Me.Label2)
        Me.Controls.Add(Me.lblSegs)
        Me.Controls.Add(Me.lblPax)
        Me.Controls.Add(Me.cmdExit)
        Me.Controls.Add(Me.cmdCreatePDF)
        Me.Controls.Add(Me.cmdFileDestination)
        Me.Controls.Add(Me.txtFileDestination)
        Me.Controls.Add(Me.Label1)
        Me.Controls.Add(Me.GroupBox3)
        Me.Controls.Add(Me.GroupBox2)
        Me.Controls.Add(Me.GroupBox1)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog
        Me.Name = "frmOSMLoG"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterParent
        Me.Text = "OSM Letter of Guarantee"
        Me.GroupBox1.ResumeLayout(False)
        Me.GroupBox1.PerformLayout()
        Me.GroupBox2.ResumeLayout(False)
        Me.GroupBox2.PerformLayout()
        Me.GroupBox3.ResumeLayout(False)
        Me.GroupBox3.PerformLayout()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub

    Friend WithEvents optPerPax As RadioButton
    Friend WithEvents optPerPNR As RadioButton
    Friend WithEvents GroupBox1 As GroupBox
    Friend WithEvents GroupBox2 As GroupBox
    Friend WithEvents optOnSigners As RadioButton
    Friend WithEvents optOffSigners As RadioButton
    Friend WithEvents GroupBox3 As GroupBox
    Friend WithEvents lstPortAgent As ListBox
    Friend WithEvents Label1 As Label
    Friend WithEvents txtFileDestination As TextBox
    Friend WithEvents fileBrowser As FolderBrowserDialog
    Friend WithEvents cmdFileDestination As Button
    Friend WithEvents cmdCreatePDF As Button
    Friend WithEvents cmdExit As Button
    Friend WithEvents lblPax As Label
    Friend WithEvents lblSegs As Label
    Friend WithEvents chkNoPortAgent As CheckBox
    Friend WithEvents txtSignedBy As TextBox
    Friend WithEvents Label2 As Label
End Class
