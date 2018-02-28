Public Class FrmFileSelect
    Private bfileselected As Boolean = False
    Private bfirstload As Boolean = True
    Private Shared sWorkingFile = My.Application.Info.DirectoryPath & "\connect - Copy.xml"

    Private Sub txtFile_TextChanged(sender As Object, e As EventArgs) Handles txtFile.TextChanged
        If bFirstLoad = False Then
            If bFileSelected = False Then
                MsgBox("Must select a file", vbOK)
            Else
                Me.btnSubmit.Enabled = True
            End If
        End If
    End Sub

    Private Sub btnBrowse_Click(sender As Object, e As EventArgs) Handles btnBrowse.Click
        Dim fd As OpenFileDialog = New OpenFileDialog

        fd.Title = "Select File"
        fd.InitialDirectory = My.Application.Info.DirectoryPath
        fd.Filter = "All files (*.*)|*.*|All files (*.*)|*.*"
        fd.FilterIndex = 2
        fd.RestoreDirectory = True

        If fd.ShowDialog() = DialogResult.OK Then
            bfileselected = True
            bfirstload = False
            sWorkingFile = fd.FileName
            Me.txtFile.Text = sWorkingFile
        End If
    End Sub

    Private Sub btnSubmit_Click(sender As Object, e As EventArgs) Handles btnSubmit.Click
        Dim iRslt As Integer

        Try
            If TxtEnterPWD.Text <> frmEncrypt.sPwd Then
                MsgBox("Wrong Password " & frmEncrypt.sPwd & " " & TxtEnterPWD.Text, MsgBoxStyle.RetryCancel)
            Else
                If Me.txtFile.Text = "" Then
                    MsgBox("Select a file", MsgBoxStyle.RetryCancel)
                Else
                    If frmEncrypt.sAction = "EditPassword" Then
                        '      If sWorkingFile.Substring(sWorkingFile.Length - 12, 12) <> "Password.xml" Then
                        'MsgBox("wrong file")
                        'Else
                        iRslt = pwdEncryption(False, "DE", False, frmEncrypt.sPWDFile)
                        If iRslt = vbOK Then
                            MsgBox("Process Complete")
                            ' Refresh the PWD from the file
                            Call pwdEncryption(False, "READ", False, frmEncrypt.sPWDFile)

                        End If
                        'End If
                    ElseIf frmEncrypt.sAction = "EditConnect" Then
                        If sWorkingFile.Substring(sWorkingFile.Length - 11, 11) <> "connect.xml" And _
                            sWorkingFile.Substring(sWorkingFile.Length - 18, 18) <> "connect - copy.xml" Then
                            MsgBox("wrong file")
                        Else
                            iRslt = CnctEncryption(sWorkingFile, "DE")
                            If iRslt = vbOK Then
                                MsgBox("Process Complete")
                            End If
                        End If

                    End If
                End If

                Me.TxtEnterPWD.Text = ""
            End If
        Catch ex As Exception
            WriteError("Error in btnSubmit_Click event " & ex.Message)
            Debug.WriteLine(ex.Message)

        Finally
        End Try


    End Sub

    Private Sub FrmFileSelect_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        bfirstload = True
        If frmEncrypt.sAction = "EditConnect" Then
            Me.txtFile.Text = My.Application.Info.DirectoryPath & "\connect - Copy.xml"
        ElseIf frmEncrypt.sAction = "EditPassword" Then
            Me.txtFile.Text = My.Application.Info.DirectoryPath & "\password.xml"
        End If
    End Sub
    Private Sub frmfileselect_unload()
        GC.Collect()
    End Sub

    Private Sub btnBack_Click(sender As Object, e As EventArgs) Handles btnBack.Click
        Me.Close()
    End Sub

    Private Sub Label1_Click(sender As Object, e As EventArgs)

    End Sub

    Private Sub btnChng_Click(sender As Object, e As EventArgs)
        Dim iRslt As Integer = 0
        If TxtEnterPWD.Text <> frmEncrypt.sPwd Then
            MsgBox("Wrong Password", MsgBoxStyle.RetryCancel)
        Else
            Dim sPwdFile As String = My.Application.Info.DirectoryPath & "\Password.xml"
            iRslt = pwdEncryption(False, "DE", False, sPwdFile)
            If iRslt = vbOK Then
                MsgBox("Process Complete")
            End If
        End If
    End Sub


    Private Sub BtnRecover_Click(sender As Object, e As EventArgs) Handles BtnRecover.Click
        'Read the file
        Call MailPWD
        'Email to the email address in the params.ini file

    End Sub


    Public Sub MailPWD()
        Dim RetVal As String
        RetVal = Space(255)
        Dim v As Integer = 0
        Dim crpt As String = ""
        Dim savecrpt As String = ""
        Dim SupportAddress As String = ""
        Dim SmtpMailHost As String = ""
        Dim SmtpPort As String = ""
        Dim sM1 As New SendMessage

        v = NativeMethods.GetPrivateProfileString("CX", "supportaddress", "", RetVal, 255, My.Application.Info.DirectoryPath & "\params.ini")
        '        SupportAddress = CStr(Left(RetVal, v)
        SupportAddress = CStr(Trim(RetVal.Substring(0, v)))
        v = 0
        v = NativeMethods.GetPrivateProfileString("CX", "smptphost", "", RetVal, 255, My.Application.Info.DirectoryPath & "\params.ini")
        '       SmtpMailHost = CStr(Left(RetVal, v))
        SmtpMailHost = Trim(RetVal.Substring(0, v))
        v = 0
        v = NativeMethods.GetPrivateProfileString("CX", "smptport", "", RetVal, 255, My.Application.Info.DirectoryPath & "\params.ini")
        '      SmtpPort = CStr(Left(RetVal, v))
        SmtpPort = Trim(RetVal.Substring(0, v))

        MsgBox("The password will be sent to the support Email address " & SupportAddress)

        sM1.SendEMessage("PWD RECOVERY", "PWD = " & frmEncrypt.sPwd, SupportAddress, SupportAddress, "", SmtpMailHost)
        'Email to the email address in the params.ini file

    End Sub

End Class