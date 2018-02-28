<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class FrmFileSelect
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
        Me.btnBrowse = New System.Windows.Forms.Button()
        Me.txtFile = New System.Windows.Forms.TextBox()
        Me.btnSubmit = New System.Windows.Forms.Button()
        Me.lblPwd = New System.Windows.Forms.Label()
        Me.TxtEnterPWD = New System.Windows.Forms.TextBox()
        Me.btnBack = New System.Windows.Forms.Button()
        Me.BtnRecover = New System.Windows.Forms.Button()
        Me.SuspendLayout()
        '
        'btnBrowse
        '
        Me.btnBrowse.Location = New System.Drawing.Point(315, 12)
        Me.btnBrowse.Name = "btnBrowse"
        Me.btnBrowse.Size = New System.Drawing.Size(92, 30)
        Me.btnBrowse.TabIndex = 1
        Me.btnBrowse.Text = "&Browse"
        Me.btnBrowse.UseVisualStyleBackColor = True
        '
        'txtFile
        '
        Me.txtFile.Enabled = False
        Me.txtFile.Location = New System.Drawing.Point(12, 12)
        Me.txtFile.Multiline = True
        Me.txtFile.Name = "txtFile"
        Me.txtFile.ReadOnly = True
        Me.txtFile.Size = New System.Drawing.Size(297, 114)
        Me.txtFile.TabIndex = 2
        Me.txtFile.Text = "File"
        '
        'btnSubmit
        '
        Me.btnSubmit.Location = New System.Drawing.Point(315, 145)
        Me.btnSubmit.Name = "btnSubmit"
        Me.btnSubmit.Size = New System.Drawing.Size(93, 30)
        Me.btnSubmit.TabIndex = 4
        Me.btnSubmit.Text = "&Submit"
        Me.btnSubmit.UseVisualStyleBackColor = True
        '
        'lblPwd
        '
        Me.lblPwd.AutoSize = True
        Me.lblPwd.Location = New System.Drawing.Point(9, 145)
        Me.lblPwd.Name = "lblPwd"
        Me.lblPwd.Size = New System.Drawing.Size(81, 13)
        Me.lblPwd.TabIndex = 17
        Me.lblPwd.Text = "Enter Password"
        '
        'TxtEnterPWD
        '
        Me.TxtEnterPWD.Location = New System.Drawing.Point(107, 145)
        Me.TxtEnterPWD.Name = "TxtEnterPWD"
        Me.TxtEnterPWD.PasswordChar = Global.Microsoft.VisualBasic.ChrW(42)
        Me.TxtEnterPWD.Size = New System.Drawing.Size(202, 20)
        Me.TxtEnterPWD.TabIndex = 3
        '
        'btnBack
        '
        Me.btnBack.Location = New System.Drawing.Point(315, 181)
        Me.btnBack.Name = "btnBack"
        Me.btnBack.Size = New System.Drawing.Size(92, 30)
        Me.btnBack.TabIndex = 18
        Me.btnBack.Text = "Back"
        Me.btnBack.UseVisualStyleBackColor = True
        '
        'BtnRecover
        '
        Me.BtnRecover.Location = New System.Drawing.Point(12, 216)
        Me.BtnRecover.Name = "BtnRecover"
        Me.BtnRecover.Size = New System.Drawing.Size(128, 31)
        Me.BtnRecover.TabIndex = 23
        Me.BtnRecover.Text = "Recover Password"
        Me.BtnRecover.UseVisualStyleBackColor = True
        '
        'FrmFileSelect
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(420, 259)
        Me.Controls.Add(Me.BtnRecover)
        Me.Controls.Add(Me.btnBack)
        Me.Controls.Add(Me.lblPwd)
        Me.Controls.Add(Me.TxtEnterPWD)
        Me.Controls.Add(Me.btnSubmit)
        Me.Controls.Add(Me.btnBrowse)
        Me.Controls.Add(Me.txtFile)
        Me.Name = "FrmFileSelect"
        Me.Text = "FrmFileSelect"
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents btnBrowse As System.Windows.Forms.Button
    Friend WithEvents txtFile As System.Windows.Forms.TextBox
    Friend WithEvents btnSubmit As System.Windows.Forms.Button
    Friend WithEvents lblPwd As System.Windows.Forms.Label
    Friend WithEvents TxtEnterPWD As System.Windows.Forms.TextBox
    Friend WithEvents btnBack As System.Windows.Forms.Button
    Friend WithEvents BtnRecover As System.Windows.Forms.Button
End Class
