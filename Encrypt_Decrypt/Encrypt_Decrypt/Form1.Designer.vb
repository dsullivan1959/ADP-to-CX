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
        Me.Label1 = New System.Windows.Forms.Label()
        Me.GrpSave = New System.Windows.Forms.GroupBox()
        Me.RadioYes = New System.Windows.Forms.RadioButton()
        Me.RadioNo = New System.Windows.Forms.RadioButton()
        Me.GroupBox1 = New System.Windows.Forms.GroupBox()
        Me.RadioEncrypt = New System.Windows.Forms.RadioButton()
        Me.RadioDecrypt = New System.Windows.Forms.RadioButton()
        Me.GrpSave.SuspendLayout()
        Me.GroupBox1.SuspendLayout()
        Me.SuspendLayout()
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(7, 16)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(129, 13)
        Me.Label1.TabIndex = 0
        Me.Label1.Text = "Set for Encrypt or Decrypt"
        '
        'GrpSave
        '
        Me.GrpSave.Controls.Add(Me.RadioNo)
        Me.GrpSave.Controls.Add(Me.RadioYes)
        Me.GrpSave.Location = New System.Drawing.Point(12, 129)
        Me.GrpSave.Name = "GrpSave"
        Me.GrpSave.Size = New System.Drawing.Size(244, 68)
        Me.GrpSave.TabIndex = 3
        Me.GrpSave.TabStop = False
        Me.GrpSave.Text = "Save in this form?"
        '
        'RadioYes
        '
        Me.RadioYes.AutoSize = True
        Me.RadioYes.Location = New System.Drawing.Point(47, 20)
        Me.RadioYes.Name = "RadioYes"
        Me.RadioYes.Size = New System.Drawing.Size(43, 17)
        Me.RadioYes.TabIndex = 0
        Me.RadioYes.TabStop = True
        Me.RadioYes.Text = "Yes"
        Me.RadioYes.UseVisualStyleBackColor = True
        '
        'RadioNo
        '
        Me.RadioNo.AutoSize = True
        Me.RadioNo.Location = New System.Drawing.Point(47, 44)
        Me.RadioNo.Name = "RadioNo"
        Me.RadioNo.Size = New System.Drawing.Size(39, 17)
        Me.RadioNo.TabIndex = 1
        Me.RadioNo.TabStop = True
        Me.RadioNo.Text = "No"
        Me.RadioNo.UseVisualStyleBackColor = True
        '
        'GroupBox1
        '
        Me.GroupBox1.Controls.Add(Me.RadioDecrypt)
        Me.GroupBox1.Controls.Add(Me.RadioEncrypt)
        Me.GroupBox1.Location = New System.Drawing.Point(11, 36)
        Me.GroupBox1.Name = "GroupBox1"
        Me.GroupBox1.Size = New System.Drawing.Size(244, 80)
        Me.GroupBox1.TabIndex = 4
        Me.GroupBox1.TabStop = False
        Me.GroupBox1.Text = "Encrypt or Decrypt?"
        '
        'RadioEncrypt
        '
        Me.RadioEncrypt.AutoSize = True
        Me.RadioEncrypt.Location = New System.Drawing.Point(48, 19)
        Me.RadioEncrypt.Name = "RadioEncrypt"
        Me.RadioEncrypt.Size = New System.Drawing.Size(61, 17)
        Me.RadioEncrypt.TabIndex = 0
        Me.RadioEncrypt.TabStop = True
        Me.RadioEncrypt.Text = "Encrypt"
        Me.RadioEncrypt.UseVisualStyleBackColor = True
        '
        'RadioDecrypt
        '
        Me.RadioDecrypt.AutoSize = True
        Me.RadioDecrypt.Location = New System.Drawing.Point(48, 43)
        Me.RadioDecrypt.Name = "RadioDecrypt"
        Me.RadioDecrypt.Size = New System.Drawing.Size(62, 17)
        Me.RadioDecrypt.TabIndex = 1
        Me.RadioDecrypt.TabStop = True
        Me.RadioDecrypt.Text = "Decrypt"
        Me.RadioDecrypt.UseVisualStyleBackColor = True
        '
        'frmEncrypt
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(284, 213)
        Me.Controls.Add(Me.GroupBox1)
        Me.Controls.Add(Me.GrpSave)
        Me.Controls.Add(Me.Label1)
        Me.Name = "frmEncrypt"
        Me.Text = "Encryption"
        Me.GrpSave.ResumeLayout(False)
        Me.GrpSave.PerformLayout()
        Me.GroupBox1.ResumeLayout(False)
        Me.GroupBox1.PerformLayout()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents GrpSave As System.Windows.Forms.GroupBox
    Friend WithEvents RadioNo As System.Windows.Forms.RadioButton
    Friend WithEvents RadioYes As System.Windows.Forms.RadioButton
    Friend WithEvents GroupBox1 As System.Windows.Forms.GroupBox
    Friend WithEvents RadioDecrypt As System.Windows.Forms.RadioButton
    Friend WithEvents RadioEncrypt As System.Windows.Forms.RadioButton

End Class
