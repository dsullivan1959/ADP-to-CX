Imports System.Net.Mail

'Imports EASendMail 'Add EASendMail namespace

Public Class SendMessage

    Public Sub SendEMessage(ByVal subject As String, ByVal messageBody As String, _
        ByVal fromAddress As String, ByVal toAddress As String, ByVal ccAddress As String, sHost As String)

        Dim message As New MailMessage()
        Dim client As New SmtpClient()

        Try
            'Set the sender's address
            message.From = New MailAddress("donotreply@carthage.edu")

            'Allow multiple "To" addresses to be separated by a semi-colon
            If (toAddress.Trim.Length > 0) Then
                For Each addr As String In toAddress.Split(";"c)
                    message.To.Add(New MailAddress(addr))
                Next
            End If

            'Set the subject and message body text
            message.Subject = subject
            message.Body = messageBody

            'Set the SMTP server to be used to send the message
            client.Host = sHost
            client.Port = 25

            'Send the e-mail message
            client.Send(message)
        Catch e As Exception
            WriteError("Error in SendEMessage.  Error = " & e.Message)
        End Try
    End Sub


   


    'Sub SendGMail()

    '    'Send Email using Gmail Account over Implicit SSL on 465 Port
    '    Dim oMail As New SmtpMail("TryIt")
    '    Dim oSmtp As New SmtpClient()
    '    ' Your gmail email address
    '    oMail.From = "gmailid@gmail.com"

    '    ' Set recipient email address, please change it to yours
    '    oMail.To = "support@emailarchitect.net"

    '    ' Set email subject
    '    oMail.Subject = "test email from gmail account"

    '    ' Set email body
    '    oMail.TextBody = "this is a test email sent from VB.NET project with gmail"
    '    'Gmail SMTP server address
    '    Dim oServer As New SmtpServer("smtp.gmail.com")

    '    ' set 465 port
    '    oServer.Port = 465

    '    ' detect SSL/TLS automatically
    '    oServer.ConnectType = SmtpConnectType.ConnectSSLAuto
    '    ' Gmail user authentication should use your
    '    ' Gmail email address as the user name.
    '    ' For example: your email is "gmailid@gmail.com", then the user should be "gmailid@gmail.com"
    '    oServer.User = "gmailid"
    '    oServer.Password = "yourpassword"

    '    Try
    '        debug.writeline("start to send email over SSL ...")
    '        oSmtp.SendMail(oServer, oMail)
    '        debug.writeline("email was sent successfully!")
    '    Catch ep As Exception
    '        debug.writeline("failed to send email with the following error:")
    '        debug.writeline(ep.Message)
    '    End Try
    'End Sub

    'Sub SendGmail1()
    '    'Send Email using Gmail Account over Explicit SSL (TLS) on 25 or 587 Port]
    '    Dim oMail As New SmtpMail("TryIt")
    '    Dim oSmtp As New SmtpClient()
    '    ' Your gmail email address
    '    oMail.From = "gmailid@gmail.com"

    '    ' Set recipient email address, please change it to yours
    '    oMail.To = "support@emailarchitect.net"

    '    ' Set email subject
    '    oMail.Subject = "test email from gmail account"

    '    ' Set email body
    '    oMail.TextBody = "this is a test email sent from VB.NET project with gmail"
    '    'Gmail SMTP server address
    '    Dim oServer As New SmtpServer("smtp.gmail.com")

    '    ' set 587 port, if you want to use 25 port, please change 587 to 25
    '    oServer.Port = 587

    '    ' detect SSL/TLS automatically
    '    oServer.ConnectType = SmtpConnectType.ConnectSSLAuto
    '    ' Gmail user authentication should use your
    '    ' Gmail email address as the user name.
    '    ' For example: your email is "gmailid@gmail.com", then the user should be "gmailid@gmail.com"

    '    oServer.User = "gmailid"
    '    oServer.Password = "yourpassword"

    '    Try
    '        debug.writeline("start to send email over TLS ...")
    '        oSmtp.SendMail(oServer, oMail)
    '        debug.writeline("email was sent successfully!")
    '    Catch ep As Exception
    '        debug.writeline("failed to send email with the following error:")
    '        debug.writeline(ep.Message)
    '    End Try
    'End Sub
End Class
