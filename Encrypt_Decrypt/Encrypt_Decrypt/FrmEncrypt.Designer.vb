<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmEncrypt
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
        Me.BtnChngPwd = New System.Windows.Forms.Button()
        Me.btnEnd = New System.Windows.Forms.Button()
        Me.btnEditCnct = New System.Windows.Forms.Button()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.SuspendLayout()
        '
        'BtnChngPwd
        '
        Me.BtnChngPwd.Location = New System.Drawing.Point(12, 64)
        Me.BtnChngPwd.Name = "BtnChngPwd"
        Me.BtnChngPwd.Size = New System.Drawing.Size(149, 23)
        Me.BtnChngPwd.TabIndex = 7
        Me.BtnChngPwd.Text = "&Change Password"
        Me.BtnChngPwd.UseVisualStyleBackColor = True
        '
        'btnEnd
        '
        Me.btnEnd.Location = New System.Drawing.Point(12, 93)
        Me.btnEnd.Name = "btnEnd"
        Me.btnEnd.Size = New System.Drawing.Size(149, 23)
        Me.btnEnd.TabIndex = 6
        Me.btnEnd.Text = "&End Program"
        Me.btnEnd.UseVisualStyleBackColor = True
        '
        'btnEditCnct
        '
        Me.btnEditCnct.Location = New System.Drawing.Point(12, 35)
        Me.btnEditCnct.Name = "btnEditCnct"
        Me.btnEditCnct.Size = New System.Drawing.Size(149, 23)
        Me.btnEditCnct.TabIndex = 13
        Me.btnEditCnct.Text = "Edit Connection File"
        Me.btnEditCnct.UseVisualStyleBackColor = True
        '
        'Label1
        '
        Me.Label1.Location = New System.Drawing.Point(17, 3)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(223, 32)
        Me.Label1.TabIndex = 14
        Me.Label1.Text = "Maintain encrypted connect string.  Requires password."
        '
        'frmEncrypt
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(253, 125)
        Me.Controls.Add(Me.Label1)
        Me.Controls.Add(Me.btnEditCnct)
        Me.Controls.Add(Me.btnEnd)
        Me.Controls.Add(Me.BtnChngPwd)
        Me.KeyPreview = True
        Me.Name = "frmEncrypt"
        Me.Text = "Encryption"
        Me.ResumeLayout(False)

    End Sub
    Friend WithEvents BtnChngPwd As System.Windows.Forms.Button
    Friend WithEvents btnEnd As System.Windows.Forms.Button
    Friend WithEvents btnEditCnct As System.Windows.Forms.Button
    Friend WithEvents Label1 As System.Windows.Forms.Label

End Class
