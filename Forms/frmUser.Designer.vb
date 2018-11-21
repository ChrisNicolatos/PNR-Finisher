<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmUser
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
        Me.lblGDS = New System.Windows.Forms.Label()
        Me.lblPCC = New System.Windows.Forms.Label()
        Me.lblUser = New System.Windows.Forms.Label()
        Me.lblUserName = New System.Windows.Forms.Label()
        Me.lblUserEmail = New System.Windows.Forms.Label()
        Me.lblQForTimeLimit = New System.Windows.Forms.Label()
        Me.lblQForReminder = New System.Windows.Forms.Label()
        Me.txtUsername = New System.Windows.Forms.TextBox()
        Me.txtEmail = New System.Windows.Forms.TextBox()
        Me.txtQueue = New System.Windows.Forms.TextBox()
        Me.txtOPQueue = New System.Windows.Forms.TextBox()
        Me.cmdCancel = New System.Windows.Forms.Button()
        Me.cmdSave = New System.Windows.Forms.Button()
        Me.lblHeader = New System.Windows.Forms.Label()
        Me.lblQHint = New System.Windows.Forms.Label()
        Me.lblCurlyBracket = New System.Windows.Forms.Label()
        Me.SuspendLayout()
        '
        'lblGDS
        '
        Me.lblGDS.AutoSize = True
        Me.lblGDS.Font = New System.Drawing.Font("Microsoft Sans Serif", 10.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblGDS.Location = New System.Drawing.Point(190, 50)
        Me.lblGDS.Name = "lblGDS"
        Me.lblGDS.Size = New System.Drawing.Size(41, 17)
        Me.lblGDS.TabIndex = 0
        Me.lblGDS.Text = "GDS"
        '
        'lblPCC
        '
        Me.lblPCC.AutoSize = True
        Me.lblPCC.Font = New System.Drawing.Font("Microsoft Sans Serif", 10.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblPCC.Location = New System.Drawing.Point(190, 81)
        Me.lblPCC.Name = "lblPCC"
        Me.lblPCC.Size = New System.Drawing.Size(38, 17)
        Me.lblPCC.TabIndex = 1
        Me.lblPCC.Text = "PCC"
        '
        'lblUser
        '
        Me.lblUser.AutoSize = True
        Me.lblUser.Font = New System.Drawing.Font("Microsoft Sans Serif", 10.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblUser.Location = New System.Drawing.Point(190, 112)
        Me.lblUser.Name = "lblUser"
        Me.lblUser.Size = New System.Drawing.Size(42, 17)
        Me.lblUser.TabIndex = 2
        Me.lblUser.Text = "User"
        '
        'lblUserName
        '
        Me.lblUserName.AutoSize = True
        Me.lblUserName.Location = New System.Drawing.Point(12, 147)
        Me.lblUserName.Name = "lblUserName"
        Me.lblUserName.Size = New System.Drawing.Size(60, 13)
        Me.lblUserName.TabIndex = 3
        Me.lblUserName.Text = "User Name"
        '
        'lblUserEmail
        '
        Me.lblUserEmail.AutoSize = True
        Me.lblUserEmail.Location = New System.Drawing.Point(12, 175)
        Me.lblUserEmail.Name = "lblUserEmail"
        Me.lblUserEmail.Size = New System.Drawing.Size(56, 13)
        Me.lblUserEmail.TabIndex = 4
        Me.lblUserEmail.Text = "User email"
        '
        'lblQForTimeLimit
        '
        Me.lblQForTimeLimit.AutoSize = True
        Me.lblQForTimeLimit.Location = New System.Drawing.Point(12, 203)
        Me.lblQForTimeLimit.Name = "lblQForTimeLimit"
        Me.lblQForTimeLimit.Size = New System.Drawing.Size(178, 13)
        Me.lblQForTimeLimit.TabIndex = 5
        Me.lblQForTimeLimit.Text = "Queue for time limit (TK TL / T.TAU)"
        '
        'lblQForReminder
        '
        Me.lblQForReminder.AutoSize = True
        Me.lblQForReminder.Location = New System.Drawing.Point(12, 231)
        Me.lblQForReminder.Name = "lblQForReminder"
        Me.lblQForReminder.Size = New System.Drawing.Size(150, 13)
        Me.lblQForReminder.TabIndex = 6
        Me.lblQForReminder.Text = "Queue for reminder (OP / RB.)"
        '
        'txtUsername
        '
        Me.txtUsername.Location = New System.Drawing.Point(193, 143)
        Me.txtUsername.Name = "txtUsername"
        Me.txtUsername.Size = New System.Drawing.Size(473, 20)
        Me.txtUsername.TabIndex = 7
        '
        'txtEmail
        '
        Me.txtEmail.Location = New System.Drawing.Point(193, 171)
        Me.txtEmail.Name = "txtEmail"
        Me.txtEmail.Size = New System.Drawing.Size(473, 20)
        Me.txtEmail.TabIndex = 8
        '
        'txtQueue
        '
        Me.txtQueue.Location = New System.Drawing.Point(193, 199)
        Me.txtQueue.Name = "txtQueue"
        Me.txtQueue.Size = New System.Drawing.Size(100, 20)
        Me.txtQueue.TabIndex = 9
        '
        'txtOPQueue
        '
        Me.txtOPQueue.Location = New System.Drawing.Point(193, 227)
        Me.txtOPQueue.Name = "txtOPQueue"
        Me.txtOPQueue.Size = New System.Drawing.Size(100, 20)
        Me.txtOPQueue.TabIndex = 10
        '
        'cmdCancel
        '
        Me.cmdCancel.Location = New System.Drawing.Point(320, 291)
        Me.cmdCancel.Name = "cmdCancel"
        Me.cmdCancel.Size = New System.Drawing.Size(75, 23)
        Me.cmdCancel.TabIndex = 11
        Me.cmdCancel.Text = "Cancel"
        Me.cmdCancel.UseVisualStyleBackColor = True
        '
        'cmdSave
        '
        Me.cmdSave.Location = New System.Drawing.Point(193, 291)
        Me.cmdSave.Name = "cmdSave"
        Me.cmdSave.Size = New System.Drawing.Size(75, 23)
        Me.cmdSave.TabIndex = 12
        Me.cmdSave.Text = "Save"
        Me.cmdSave.UseVisualStyleBackColor = True
        '
        'lblHeader
        '
        Me.lblHeader.BackColor = System.Drawing.Color.FromArgb(CType(CType(255, Byte), Integer), CType(CType(128, Byte), Integer), CType(CType(128, Byte), Integer))
        Me.lblHeader.Dock = System.Windows.Forms.DockStyle.Top
        Me.lblHeader.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblHeader.Location = New System.Drawing.Point(0, 0)
        Me.lblHeader.Name = "lblHeader"
        Me.lblHeader.Size = New System.Drawing.Size(700, 23)
        Me.lblHeader.TabIndex = 13
        Me.lblHeader.Text = "Please enter your information to proceed"
        Me.lblHeader.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'lblQHint
        '
        Me.lblQHint.AutoSize = True
        Me.lblQHint.Location = New System.Drawing.Point(341, 217)
        Me.lblQHint.Name = "lblQHint"
        Me.lblQHint.Size = New System.Drawing.Size(229, 13)
        Me.lblQHint.TabIndex = 14
        Me.lblQHint.Text = "Enter queue numbers without Q (e.g. 72, 72C4)"
        '
        'lblCurlyBracket
        '
        Me.lblCurlyBracket.AutoSize = True
        Me.lblCurlyBracket.Font = New System.Drawing.Font("Arial Narrow", 32.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblCurlyBracket.Location = New System.Drawing.Point(299, 194)
        Me.lblCurlyBracket.Name = "lblCurlyBracket"
        Me.lblCurlyBracket.Size = New System.Drawing.Size(35, 52)
        Me.lblCurlyBracket.TabIndex = 15
        Me.lblCurlyBracket.Text = "}"
        '
        'frmUser
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(700, 350)
        Me.Controls.Add(Me.lblCurlyBracket)
        Me.Controls.Add(Me.lblQHint)
        Me.Controls.Add(Me.lblHeader)
        Me.Controls.Add(Me.cmdSave)
        Me.Controls.Add(Me.cmdCancel)
        Me.Controls.Add(Me.txtOPQueue)
        Me.Controls.Add(Me.txtQueue)
        Me.Controls.Add(Me.txtEmail)
        Me.Controls.Add(Me.txtUsername)
        Me.Controls.Add(Me.lblQForReminder)
        Me.Controls.Add(Me.lblQForTimeLimit)
        Me.Controls.Add(Me.lblUserEmail)
        Me.Controls.Add(Me.lblUserName)
        Me.Controls.Add(Me.lblUser)
        Me.Controls.Add(Me.lblPCC)
        Me.Controls.Add(Me.lblGDS)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedToolWindow
        Me.Name = "frmUser"
        Me.Text = "Add User"
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub

    Friend WithEvents lblGDS As Label
    Friend WithEvents lblPCC As Label
    Friend WithEvents lblUser As Label
    Friend WithEvents lblUserName As Label
    Friend WithEvents lblUserEmail As Label
    Friend WithEvents lblQForTimeLimit As Label
    Friend WithEvents lblQForReminder As Label
    Friend WithEvents txtUsername As TextBox
    Friend WithEvents txtEmail As TextBox
    Friend WithEvents txtQueue As TextBox
    Friend WithEvents txtOPQueue As TextBox
    Friend WithEvents cmdCancel As Button
    Friend WithEvents cmdSave As Button
    Friend WithEvents lblHeader As Label
    Friend WithEvents lblQHint As Label
    Friend WithEvents lblCurlyBracket As Label
End Class
