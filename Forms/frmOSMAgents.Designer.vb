<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmOSMAgents
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
        Me.cmdOSMEditDeleteAgent = New System.Windows.Forms.Button()
        Me.cmdOSMAddAgent = New System.Windows.Forms.Button()
        Me.cmdOSMEditUpdateAgent = New System.Windows.Forms.Button()
        Me.txtOSMEditAgentEmail = New System.Windows.Forms.TextBox()
        Me.Label6 = New System.Windows.Forms.Label()
        Me.txtOSMEditAgentName = New System.Windows.Forms.TextBox()
        Me.Label5 = New System.Windows.Forms.Label()
        Me.lstOSMEditAgents = New System.Windows.Forms.ListBox()
        Me.Label4 = New System.Windows.Forms.Label()
        Me.cmdOSMEditAgentExit = New System.Windows.Forms.Button()
        Me.txtOSMEditAgentDetails = New System.Windows.Forms.TextBox()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.SuspendLayout()
        '
        'cmdOSMEditDeleteAgent
        '
        Me.cmdOSMEditDeleteAgent.Location = New System.Drawing.Point(276, 450)
        Me.cmdOSMEditDeleteAgent.Name = "cmdOSMEditDeleteAgent"
        Me.cmdOSMEditDeleteAgent.Size = New System.Drawing.Size(141, 23)
        Me.cmdOSMEditDeleteAgent.TabIndex = 10
        Me.cmdOSMEditDeleteAgent.Text = "Delete agent"
        Me.cmdOSMEditDeleteAgent.UseVisualStyleBackColor = True
        '
        'cmdOSMAddAgent
        '
        Me.cmdOSMAddAgent.Location = New System.Drawing.Point(105, 319)
        Me.cmdOSMAddAgent.Name = "cmdOSMAddAgent"
        Me.cmdOSMAddAgent.Size = New System.Drawing.Size(141, 23)
        Me.cmdOSMAddAgent.TabIndex = 2
        Me.cmdOSMAddAgent.Text = "Add agent"
        Me.cmdOSMAddAgent.UseVisualStyleBackColor = True
        '
        'cmdOSMEditUpdateAgent
        '
        Me.cmdOSMEditUpdateAgent.Location = New System.Drawing.Point(105, 450)
        Me.cmdOSMEditUpdateAgent.Name = "cmdOSMEditUpdateAgent"
        Me.cmdOSMEditUpdateAgent.Size = New System.Drawing.Size(141, 23)
        Me.cmdOSMEditUpdateAgent.TabIndex = 9
        Me.cmdOSMEditUpdateAgent.Text = "Update agent"
        Me.cmdOSMEditUpdateAgent.UseVisualStyleBackColor = True
        '
        'txtOSMEditAgentEmail
        '
        Me.txtOSMEditAgentEmail.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.txtOSMEditAgentEmail.Location = New System.Drawing.Point(105, 404)
        Me.txtOSMEditAgentEmail.Name = "txtOSMEditAgentEmail"
        Me.txtOSMEditAgentEmail.Size = New System.Drawing.Size(523, 20)
        Me.txtOSMEditAgentEmail.TabIndex = 8
        '
        'Label6
        '
        Me.Label6.AutoSize = True
        Me.Label6.Location = New System.Drawing.Point(23, 408)
        Me.Label6.Name = "Label6"
        Me.Label6.Size = New System.Drawing.Size(32, 13)
        Me.Label6.TabIndex = 7
        Me.Label6.Text = "eMail"
        '
        'txtOSMEditAgentName
        '
        Me.txtOSMEditAgentName.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.txtOSMEditAgentName.Location = New System.Drawing.Point(105, 348)
        Me.txtOSMEditAgentName.Name = "txtOSMEditAgentName"
        Me.txtOSMEditAgentName.Size = New System.Drawing.Size(523, 20)
        Me.txtOSMEditAgentName.TabIndex = 4
        '
        'Label5
        '
        Me.Label5.AutoSize = True
        Me.Label5.Location = New System.Drawing.Point(23, 352)
        Me.Label5.Name = "Label5"
        Me.Label5.Size = New System.Drawing.Size(66, 13)
        Me.Label5.TabIndex = 3
        Me.Label5.Text = "Agent Name"
        '
        'lstOSMEditAgents
        '
        Me.lstOSMEditAgents.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
            Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.lstOSMEditAgents.FormattingEnabled = True
        Me.lstOSMEditAgents.Location = New System.Drawing.Point(23, 36)
        Me.lstOSMEditAgents.Name = "lstOSMEditAgents"
        Me.lstOSMEditAgents.Size = New System.Drawing.Size(605, 277)
        Me.lstOSMEditAgents.TabIndex = 1
        '
        'Label4
        '
        Me.Label4.AutoSize = True
        Me.Label4.Location = New System.Drawing.Point(23, 20)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(35, 13)
        Me.Label4.TabIndex = 0
        Me.Label4.Text = "Agent"
        '
        'cmdOSMEditAgentExit
        '
        Me.cmdOSMEditAgentExit.Location = New System.Drawing.Point(487, 511)
        Me.cmdOSMEditAgentExit.Name = "cmdOSMEditAgentExit"
        Me.cmdOSMEditAgentExit.Size = New System.Drawing.Size(141, 23)
        Me.cmdOSMEditAgentExit.TabIndex = 11
        Me.cmdOSMEditAgentExit.Text = "Exit"
        Me.cmdOSMEditAgentExit.UseVisualStyleBackColor = True
        '
        'txtOSMEditAgentDetails
        '
        Me.txtOSMEditAgentDetails.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.txtOSMEditAgentDetails.Location = New System.Drawing.Point(105, 376)
        Me.txtOSMEditAgentDetails.Name = "txtOSMEditAgentDetails"
        Me.txtOSMEditAgentDetails.Size = New System.Drawing.Size(523, 20)
        Me.txtOSMEditAgentDetails.TabIndex = 6
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(23, 380)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(70, 13)
        Me.Label1.TabIndex = 5
        Me.Label1.Text = "Agent Details"
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Location = New System.Drawing.Point(106, 424)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(235, 13)
        Me.Label2.TabIndex = 12
        Me.Label2.Text = "Multiple email addresses can be separated with ;"
        '
        'frmOSMAgents
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(650, 554)
        Me.Controls.Add(Me.Label2)
        Me.Controls.Add(Me.txtOSMEditAgentDetails)
        Me.Controls.Add(Me.Label1)
        Me.Controls.Add(Me.cmdOSMEditDeleteAgent)
        Me.Controls.Add(Me.cmdOSMAddAgent)
        Me.Controls.Add(Me.cmdOSMEditUpdateAgent)
        Me.Controls.Add(Me.txtOSMEditAgentEmail)
        Me.Controls.Add(Me.Label6)
        Me.Controls.Add(Me.txtOSMEditAgentName)
        Me.Controls.Add(Me.Label5)
        Me.Controls.Add(Me.lstOSMEditAgents)
        Me.Controls.Add(Me.Label4)
        Me.Controls.Add(Me.cmdOSMEditAgentExit)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.SizableToolWindow
        Me.Name = "frmOSMAgents"
        Me.Text = "OSM Agents"
        Me.ResumeLayout(false)
        Me.PerformLayout

End Sub
    Friend WithEvents cmdOSMEditDeleteAgent As System.Windows.Forms.Button
    Friend WithEvents cmdOSMAddAgent As System.Windows.Forms.Button
    Friend WithEvents cmdOSMEditUpdateAgent As System.Windows.Forms.Button
    Friend WithEvents txtOSMEditAgentEmail As System.Windows.Forms.TextBox
    Friend WithEvents Label6 As System.Windows.Forms.Label
    Friend WithEvents txtOSMEditAgentName As System.Windows.Forms.TextBox
    Friend WithEvents Label5 As System.Windows.Forms.Label
    Friend WithEvents lstOSMEditAgents As System.Windows.Forms.ListBox
    Friend WithEvents Label4 As System.Windows.Forms.Label
    Friend WithEvents cmdOSMEditAgentExit As System.Windows.Forms.Button
    Friend WithEvents txtOSMEditAgentDetails As System.Windows.Forms.TextBox
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents Label2 As System.Windows.Forms.Label
End Class
