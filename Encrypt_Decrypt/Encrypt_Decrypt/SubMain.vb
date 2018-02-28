Option Strict Off
Option Explicit On


Imports System
Imports System.Object
Imports System.Configuration
Imports System.Collections.Specialized
Imports System.Net
Imports System.Diagnostics
Imports System.Security.Cryptography
Imports System.Security.Cryptography.Xml
Imports System.Xml
Imports System.IO
Imports System.Text
Imports System.Text.RegularExpressions
Imports System.Data
Imports System.Data.Odbc
Imports System.Data.SqlClient
Imports System.Data.DataSet

Module SubMain

    Dim bSave As Boolean = False
    Dim sEncrypt As String = ""

    Public Sub Main()

    End Sub
    Public Sub End_program()
        GC.Collect()
        frmEncrypt = Nothing
        End
    End Sub


    Public Function CnctEncryption(File As String, sEncrypt As String) As Integer
        Dim RetVal As String
        RetVal = Space(255)
        '  Dim v As Integer
        Dim xmlDoc As New XmlDocument()
        Dim crpt As String = ""
        Dim savecrpt As String = ""
        'Dim ipos As Integer
        '        Dim sDataPath As String
        'Dim AppPath As String


        'AppPath = My.Application.Info.DirectoryPath

        'The connection string will be encrypted in a file within the project
        'sConnectString = "Dsn=" & sDsn & ";Driver=" & sDriver & ";Host=" & sHost & ";Server=" & sServer & ";Service=" & sService & ";Protocol=" & sProtocol & ";Database=" & sDataBase & ";Uid=" & sUid & ";Pwd=" & sPwd & ";"

        '+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
        'This calls the encryption/decryption process
        'to toggle the encryption on or off change the value "crpt" in the Params.ini file:  EN for encrypt, DE for decrypt
        Dim cspParams As New CspParameters()            ' Create a new CspParameters object to specify a key container.
        cspParams.KeyContainerName = "XML_ENC_RSA_KEY"
        ' Create a new RSA key and save it in the container.  This key will encrypt
        ' a symmetric key, which will then be encryped in the XML document.
        Dim rsaKey As New RSACryptoServiceProvider(cspParams)
        Try
            xmlDoc.PreserveWhitespace = True
            '            xmlDoc.Load("connect_Copy.xml")   'In the bin/debug folder
            xmlDoc.Load(File)
           
            If sEncrypt = "EN" Then

                '++++++++++++++++++++++++++++++++++++++++++
                'ENCRYPT
                'Use this to encrypt the connection string file.   Once encrypted, comment out
                'clsEncryption.Encrypt(xmlDoc, "password", "EncryptedElement1", rsaKey, "rsaKey")
                clsEncryption.Encrypt(xmlDoc, "connectionString", "EncryptedElement1", rsaKey, "rsaKey")
                clsEncryption.Encrypt(xmlDoc, "FTPHost", "EncryptedElement2", rsaKey, "rsaKey")
                clsEncryption.Encrypt(xmlDoc, "FTPUser", "EncryptedElement3", rsaKey, "rsaKey")
                clsEncryption.Encrypt(xmlDoc, "FTPPwd", "EncryptedElement4", rsaKey, "rsaKey")
                clsEncryption.Encrypt(xmlDoc, "FTPKey", "EncryptedElement5", rsaKey, "rsaKey")

                '++++++++++++++++++++++++++++++++++++++++++
            ElseIf sEncrypt = "DE" Then
                '++++++++++++++++++++++++++++++++++++++++++
                '++++++++++++++++++++++++++++++++++++++++++
                'DECRYPT
                clsEncryption.Decrypt(xmlDoc, rsaKey, "rsaKey")
                '++++++++++++++++++++++++++++++++++++++++++
                xmlDoc.Save(File)

                Dim myprocess = System.Diagnostics.Process.Start("notepad.exe", File)
                myprocess.WaitForExit()

                MsgBox("Saved")
                xmlDoc.Load(File)

                'clsEncryption.Encrypt(xmlDoc, "password", "EncryptedElement1", rsaKey, "rsaKey")
                clsEncryption.Encrypt(xmlDoc, "connectionString", "EncryptedElement1", rsaKey, "rsaKey")
                clsEncryption.Encrypt(xmlDoc, "FTPHost", "EncryptedElement2", rsaKey, "rsaKey")
                clsEncryption.Encrypt(xmlDoc, "FTPUser", "EncryptedElement3", rsaKey, "rsaKey")
                clsEncryption.Encrypt(xmlDoc, "FTPPwd", "EncryptedElement4", rsaKey, "rsaKey")
                clsEncryption.Encrypt(xmlDoc, "FTPKey", "EncryptedElement5", rsaKey, "rsaKey")

                xmlDoc.Save(File)
                '++++++++++++++++++++++++++++++++++++++++++
            Else
                'do nothing
            End If
            CnctEncryption = vbOK

        Catch e As Exception
            WriteError("Error in cnctEncryption " & e.Message)
            MsgBox("Error = " & e.Message & ", " & e.HelpLink)
            Debug.WriteLine(e.Message & ", " & e.HelpLink)
            CnctEncryption = vbCancel
        End Try

    End Function


    Public Function pwdEncryption(bSave As Boolean, sEncrypt As String, bWrite As Boolean, PWDfile As String) As Integer
        Dim RetVal As String
        RetVal = Space(255)
        Dim xmlDoc As New XmlDocument()
        Dim crpt As String = ""
        Dim savecrpt As String = ""

        '+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
        'This calls the encryption/decryption process
        'to toggle the encryption on or off change the value "crpt" in the Params.ini file:  EN for encrypt, DE for decrypt
        Dim cspParams As New CspParameters()            ' Create a new CspParameters object to specify a key container.
        cspParams.KeyContainerName = "XML_ENC_RSA_KEY"
        ' Create a new RSA key and save it in the container.  This key will encrypt
        ' a symmetric key, which will then be encryped in the XML document.
        Dim rsaKey As New RSACryptoServiceProvider(cspParams)
        Try
            xmlDoc.PreserveWhitespace = True
            xmlDoc.Load(PWDFile)


            If sEncrypt = "EN" Then
                '++++++++++++++++++++++++++++++++++++++++++
                'ENCRYPT
                clsEncryption.Encrypt(xmlDoc, "Password", "EncryptedElement1", rsaKey, "rsaKey")
                '++++++++++++++++++++++++++++++++++++++++++
            ElseIf sEncrypt = "DE" Then
                '++++++++++++++++++++++++++++++++++++++++++
                'DECRYPT
                clsEncryption.Decrypt(xmlDoc, rsaKey, "rsaKey")
                '++++++++++++++++++++++++++++++++++++++++++

                xmlDoc.Save(PWDfile)
                Dim myprocess = System.Diagnostics.Process.Start("notepad.exe", PWDfile)
                myprocess.WaitForExit()
                MsgBox("Saved")

                xmlDoc.Load(PWDfile)
                clsEncryption.Encrypt(xmlDoc, "Password", "EncryptedElement1", rsaKey, "rsaKey")
                xmlDoc.Save(PWDfile)


            ElseIf sEncrypt = "READ" Then
                clsEncryption.Decrypt(xmlDoc, rsaKey, "rsaKey")
            End If

            'Read the password
            Dim elemList As XmlNodeList = xmlDoc.GetElementsByTagName("Password")
            Dim i As Integer
            For i = 0 To elemList.Count - 1
                Debug.WriteLine(elemList(i).InnerXml)
                frmEncrypt.sPwd = Trim(elemList(i).InnerXml)

            Next i

            pwdEncryption = vbOK

        Catch e As Exception
            'WriteError("Error in GetParams " & e.Message)
            Debug.WriteLine(e.Message)
            pwdEncryption = vbCancel
        End Try

    End Function



    Public Function WriteError(ByVal msg As String)
        FileOpen(10, My.Application.Info.DirectoryPath & "\EncryptErrorLog.txt", OpenMode.Append)
        PrintLine(10, msg & ", " & Format(Now, "MM/dd/yyyy HH:MM:ss"))
        '        Dim sMsg As New SendMessage
        '        Call sMsg.SendEMessage("ADP to CX Error", msg, gl.FromAddress, gl.ToAddress, "")
        FileClose(10)
        WriteError = 1
    End Function


    Public Function ConvertString(ByVal sValue As String) As String
        Dim sAns As String
        sAns = Replace(sValue, Chr(39), "''")
        sAns = "'" & sAns & "'"
        ConvertString = sAns
    End Function

End Module
