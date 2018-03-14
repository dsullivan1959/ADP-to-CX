Public Class frmEncrypt

    Private sEncrypt As String = ""
    Private bSave As Boolean = False
    Public Shared sWorkingFile As String
    Public Shared sPWDFile As String
    Public Shared sPwd As String = ""
    Public Shared sNewPwd As String = ""
    Private bFileSelected As Boolean = False
    Private bFirstLoad As Boolean = True
    Public Shared sAction As String = ""

    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        sPWDFile = My.Application.Info.DirectoryPath & "\Password.xml"
        Call pwdEncryption(False, "READ", False, sPWDFile)
    End Sub

    Private Sub CheckPWD()
        Dim sPwd As String = ""
    End Sub

    Private Sub BtnChngPwd_Click(sender As Object, e As EventArgs) Handles BtnChngPwd.Click

        sAction = "EditPassword"
        FrmFileSelect.btnBrowse.Enabled = False
        FrmFileSelect.btnSubmit.Enabled = False

        FrmFileSelect.Show()

    End Sub


    Private Sub btnEnd_Click(sender As Object, e As EventArgs) Handles btnEnd.Click
        Call End_program()
    End Sub

    Private Sub btnEditCnct_Click(sender As Object, e As EventArgs) Handles btnEditCnct.Click
        sAction = "EditConnect"
        FrmFileSelect.Show()
    End Sub

 
End Class
