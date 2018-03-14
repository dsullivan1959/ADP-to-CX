Imports System.Security.Cryptography
Imports System.Xml
Imports System.IO
Imports System.Security.Cryptography.Xml

Public Class clsEncryption
    Implements IDisposable

    Private TripleDes As New TripleDESCryptoServiceProvider

    '*************************
    '** Global Variables
    '*************************

    Dim strFileToEncrypt As String
    Dim strFileToDecrypt As String
    Dim strOutputEncrypt As String
    Dim strOutputDecrypt As String
    Dim fsInput As System.IO.FileStream
    Dim fsOutput As System.IO.FileStream

    '*************************
    '** Create Key
    '*************************
    Public Function CreateKey(ByVal strPassword As String) As Byte()
        Dim bytKey As Byte()
        Dim bytSalt As Byte() = System.Text.Encoding.ASCII.GetBytes("salt")
        '        Dim pdb As New PasswordDeriveBytes(strPassword, bytSalt)
        Dim pdb As New Rfc2898DeriveBytes(strPassword, bytSalt)

        bytKey = pdb.GetBytes(32)

        Return bytKey 'Return the key.
    End Function

    '*************************
    '** Create IV
    '*************************


    Public Function CreateIV(ByVal strPassword As String) As Byte()
        Dim bytIV As Byte()
        Dim bytSalt As Byte() = System.Text.Encoding.ASCII.GetBytes("salt")
        Dim pdb As New Rfc2898DeriveBytes(strPassword, bytSalt)

        bytIV = pdb.GetBytes(16)

        Return bytIV 'Return the IV.
    End Function

    '****************************
    '** Encrypt/Decrypt File
    '****************************

    Public Enum CryptoAction
        'Define the enumeration for CryptoAction.
        ActionEncrypt = 1
        ActionDecrypt = 2
    End Enum

    Public Sub EncryptOrDecryptFile(ByVal strInputFile As String, ByVal strOutputFile As String, ByVal bytKey() As Byte, ByVal bytIV() As Byte, ByVal Direction As CryptoAction)
        Try 'In case of errors.

            'Setup file streams to handle input and output.
            fsInput = New System.IO.FileStream(strInputFile, FileMode.Open, FileAccess.Read)
            fsOutput = New System.IO.FileStream(strOutputFile, FileMode.OpenOrCreate, FileAccess.Write)
            fsOutput.SetLength(0) 'make sure fsOutput is empty

            'Declare variables for encrypt/decrypt process.
            Dim bytBuffer(4096) As Byte 'holds a block of bytes for processing
            Dim lngBytesProcessed As Long = 0 'running count of bytes processed
            Dim lngFileLength As Long = fsInput.Length 'the input file's length
            Dim intBytesInCurrentBlock As Integer 'current bytes being processed
            Dim csCryptoStream As CryptoStream = Nothing
            'Declare your CryptoServiceProvider.
            Dim cspRijndael As New System.Security.Cryptography.RijndaelManaged

            'Determine if ecryption or decryption and setup CryptoStream.
            Select Case Direction
                Case CryptoAction.ActionEncrypt
                    csCryptoStream = New CryptoStream(fsOutput, cspRijndael.CreateEncryptor(bytKey, bytIV), CryptoStreamMode.Write)

                Case CryptoAction.ActionDecrypt
                    csCryptoStream = New CryptoStream(fsOutput, cspRijndael.CreateDecryptor(bytKey, bytIV), CryptoStreamMode.Write)
            End Select

            'Use While to loop until all of the file is processed.
            While lngBytesProcessed < lngFileLength
                'Read file with the input filestream.
                intBytesInCurrentBlock = fsInput.Read(bytBuffer, 0, 4096)
                'Write output file with the cryptostream.
                csCryptoStream.Write(bytBuffer, 0, intBytesInCurrentBlock)
                'Update lngBytesProcessed
                lngBytesProcessed = lngBytesProcessed + CLng(intBytesInCurrentBlock)

            End While

            'Close FileStreams and CryptoStream.
            csCryptoStream.Close()
            fsOutput.Dispose()
            fsInput.Dispose()
        Catch e As Exception
        Finally

        End Try
    End Sub

    'Public Sub KeyMain()  'This was demo code
    '    Try


    '        ' Create a key and save it in a container.  
    '        GenKey_SaveInContainer("MyKeyContainer")

    '        ' Retrieve the key from the container.  
    '        GetKeyFromContainer("MyKeyContainer")

    '        ' Delete the key from the container.  
    '        DeleteKeyFromContainer("MyKeyContainer")

    '        ' Create a key and save it in a container.  
    '        GenKey_SaveInContainer("MyKeyContainer")

    '        ' Delete the key from the container.  
    '        DeleteKeyFromContainer("MyKeyContainer")
    '    Catch e As CryptographicException
    '        debug.writeline(e.Message)
    '    End Try
    'End Sub

    'This is also originally demo code
    Public Sub EncryptElement(ByVal xmlFile As String, sElement As String)
        ' Create an XmlDocument object.
        Dim xmlDoc As New XmlDocument()

        ' Load an XML file into the XmlDocument object.
        Try
            xmlDoc.PreserveWhitespace = True
            xmlDoc.Load(xmlFile)
        Catch e As Exception
            WriteError("Error in clsEncyption.EncryptElement  " & e.Message)
            Debug.WriteLine(e.Message)
        End Try


        ' Create a new CspParameters object to specify
        ' a key container.
        Dim cspParams As New CspParameters()
        cspParams.KeyContainerName = "XML_ENC_RSA_KEY"
        ' Create a new RSA key and save it in the container.  This key will encrypt
        ' a symmetric key, which will then be encryped in the XML document.
        Dim rsaKey As New RSACryptoServiceProvider(cspParams)
        Try
            ' Encrypt the "creditcard" element.
            'Arguments are documentName, Element, ElementID, Algorythm, Key
            'Encrypt(xmlDoc, "creditcard", "EncryptedElement1", rsaKey, "rsaKey")

            'Example of an XML Element
            '<creditcard>
            '  <number>19834209</number>
            '  <expiry>02/02/2002</expiry>
            '</creditcard>

            ' Save the XML document.
            xmlDoc.Save("test.xml")
            ' Display the encrypted XML to the console.
            'debug.writeline("Encrypted XML:")
            'debug.writeline(xmlDoc.OuterXml)
            'Decrypt(xmlDoc, rsaKey, "rsaKey")
            'xmlDoc.Save("test.xml")
            ' Display the encrypted XML to the console.
            'debug.writeline("Decrypted XML:")
            'debug.writeline(xmlDoc.OuterXml)

        Catch e As Exception
            WriteError("Error in clsEncyption.EncryptElement  " & e.Message)
            '            Debug.WriteLine(e.Message)
        Finally
            ' Clear the RSA key.
            rsaKey.Clear()
        End Try


        Console.ReadLine()

    End Sub 'Main


    Public Shared Sub Encrypt(ByVal Doc As XmlDocument, ByVal EncryptionElement As String, ByVal EncryptionElementID As String, ByVal Alg As RSA, ByVal KeyName As String)
        ' Check the arguments.
        If Doc Is Nothing Then
            Throw New ArgumentNullException("Doc")
        End If
        If EncryptionElement Is Nothing Then
            Throw New ArgumentNullException("EncryptionElement")
        End If
        If EncryptionElementID Is Nothing Then
            Throw New ArgumentNullException("EncryptionElementID")
        End If
        If Alg Is Nothing Then
            Throw New ArgumentNullException("Alg")
        End If
        If KeyName Is Nothing Then
            Throw New ArgumentNullException("KeyName")
        End If
        '//////////////////////////////////////////////
        ' Find the specified element in the XmlDocument
        ' object and create a new XmlElemnt object.
        '//////////////////////////////////////////////
        Dim elementToEncrypt As XmlElement = Doc.GetElementsByTagName(EncryptionElement)(0) '

        ' Throw an XmlException if the element was not found.
        If elementToEncrypt Is Nothing Then
            Throw New XmlException("The specified element was not found")
        End If
        Dim sessionKey As RijndaelManaged = Nothing

        Try
            '////////////////////////////////////////////////
            ' Create a new instance of the EncryptedXml class
            ' and use it to encrypt the XmlElement with the
            ' a new random symmetric key.
            '////////////////////////////////////////////////
            ' Create a 256 bit Rijndael key.
            sessionKey = New RijndaelManaged()
            sessionKey.KeySize = 256
            Dim eXml As New EncryptedXml()

            Dim encryptedElement As Byte() = eXml.EncryptData(elementToEncrypt, sessionKey, False)
            '//////////////////////////////////////////////
            ' Construct an EncryptedData object and populate
            ' it with the desired encryption information.
            '//////////////////////////////////////////////
            Dim edElement As New EncryptedData()
            edElement.Type = EncryptedXml.XmlEncElementUrl
            edElement.Id = EncryptionElementID
            ' Create an EncryptionMethod element so that the
            ' receiver knows which algorithm to use for decryption.
            edElement.EncryptionMethod = New EncryptionMethod(EncryptedXml.XmlEncAES256Url)
            ' Encrypt the session key and add it to an EncryptedKey element.
            Dim ek As New EncryptedKey()

            Dim encryptedKey As Byte() = EncryptedXml.EncryptKey(sessionKey.Key, Alg, False)

            ek.CipherData = New CipherData(encryptedKey)

            ek.EncryptionMethod = New EncryptionMethod(EncryptedXml.XmlEncRSA15Url)
            ' Create a new DataReference element
            ' for the KeyInfo element.  This optional
            ' element specifies which EncryptedData
            ' uses this key.  An XML document can have
            ' multiple EncryptedData elements that use
            ' different keys.
            Dim dRef As New DataReference()

            ' Specify the EncryptedData URI.
            dRef.Uri = "#" + EncryptionElementID

            ' Add the DataReference to the EncryptedKey.
            ek.AddReference(dRef)
            ' Add the encrypted key to the
            ' EncryptedData object.
            edElement.KeyInfo.AddClause(New KeyInfoEncryptedKey(ek))
            ' Set the KeyInfo element to specify the
            ' name of the RSA key.
            ' Create a new KeyInfoName element.
            Dim kin As New KeyInfoName()

            ' Specify a name for the key.
            kin.Value = KeyName

            ' Add the KeyInfoName element to the
            ' EncryptedKey object.
            ek.KeyInfo.AddClause(kin)
            ' Add the encrypted element data to the
            ' EncryptedData object.
            edElement.CipherData.CipherValue = encryptedElement
            '//////////////////////////////////////////////////
            ' Replace the element from the original XmlDocument
            ' object with the EncryptedData element.
            '//////////////////////////////////////////////////
            EncryptedXml.ReplaceElement(elementToEncrypt, edElement, False)
        Catch e As Exception
            WriteError("Error in clsEncyption.Encrypt  " & e.Message)

            ' re-throw the exception.
            Throw
        Finally
            If Not (sessionKey Is Nothing) Then
                sessionKey.Clear()
            End If
        End Try

    End Sub 'Encrypt




    Public Shared Sub EncryptTxt(ByVal Doc As XmlDocument, ByVal EncryptionElement As String, ByVal EncryptionElementID As String, ByVal Alg As RSA, ByVal KeyName As String)
        ' Check the arguments.
        If Doc Is Nothing Then
            Throw New ArgumentNullException("Doc")
        End If
        If EncryptionElement Is Nothing Then
            Throw New ArgumentNullException("EncryptionElement")
        End If
        If EncryptionElementID Is Nothing Then
            Throw New ArgumentNullException("EncryptionElementID")
        End If
        If Alg Is Nothing Then
            Throw New ArgumentNullException("Alg")
        End If
        If KeyName Is Nothing Then
            Throw New ArgumentNullException("KeyName")
        End If
        '//////////////////////////////////////////////
        ' Find the specified element in the XmlDocument
        ' object and create a new XmlElemnt object.
        '//////////////////////////////////////////////
        Dim elementToEncrypt As XmlElement = Doc.GetElementsByTagName(EncryptionElement)(0) '

        ' Throw an XmlException if the element was not found.
        If elementToEncrypt Is Nothing Then
            Throw New XmlException("The specified element was not found")
        End If
        Dim sessionKey As RijndaelManaged = Nothing

        Try
            '////////////////////////////////////////////////
            ' Create a new instance of the EncryptedXml class
            ' and use it to encrypt the XmlElement with the
            ' a new random symmetric key.
            '////////////////////////////////////////////////
            ' Create a 256 bit Rijndael key.
            sessionKey = New RijndaelManaged()
            sessionKey.KeySize = 256
            Dim eXml As New EncryptedXml()

            Dim encryptedElement As Byte() = eXml.EncryptData(elementToEncrypt, sessionKey, False)
            '//////////////////////////////////////////////
            ' Construct an EncryptedData object and populate
            ' it with the desired encryption information.
            '//////////////////////////////////////////////
            Dim edElement As New EncryptedData()
            edElement.Type = EncryptedXml.XmlEncElementUrl
            edElement.Id = EncryptionElementID
            ' Create an EncryptionMethod element so that the
            ' receiver knows which algorithm to use for decryption.
            edElement.EncryptionMethod = New EncryptionMethod(EncryptedXml.XmlEncAES256Url)
            ' Encrypt the session key and add it to an EncryptedKey element.
            Dim ek As New EncryptedKey()

            Dim encryptedKey As Byte() = EncryptedXml.EncryptKey(sessionKey.Key, Alg, False)

            ek.CipherData = New CipherData(encryptedKey)

            ek.EncryptionMethod = New EncryptionMethod(EncryptedXml.XmlEncRSA15Url)
            ' Create a new DataReference element
            ' for the KeyInfo element.  This optional
            ' element specifies which EncryptedData
            ' uses this key.  An XML document can have
            ' multiple EncryptedData elements that use
            ' different keys.
            Dim dRef As New DataReference()

            ' Specify the EncryptedData URI.
            dRef.Uri = "#" + EncryptionElementID

            ' Add the DataReference to the EncryptedKey.
            ek.AddReference(dRef)
            ' Add the encrypted key to the
            ' EncryptedData object.
            edElement.KeyInfo.AddClause(New KeyInfoEncryptedKey(ek))
            ' Set the KeyInfo element to specify the
            ' name of the RSA key.
            ' Create a new KeyInfoName element.
            Dim kin As New KeyInfoName()

            ' Specify a name for the key.
            kin.Value = KeyName

            ' Add the KeyInfoName element to the
            ' EncryptedKey object.
            ek.KeyInfo.AddClause(kin)
            ' Add the encrypted element data to the
            ' EncryptedData object.
            edElement.CipherData.CipherValue = encryptedElement
            '//////////////////////////////////////////////////
            ' Replace the element from the original XmlDocument
            ' object with the EncryptedData element.
            '//////////////////////////////////////////////////
            EncryptedXml.ReplaceElement(elementToEncrypt, edElement, False)
        Catch e As Exception
            WriteError("Error in clsEncyption.Encrypt  " & e.Message)

            ' re-throw the exception.
            Throw
        Finally
            If Not (sessionKey Is Nothing) Then
                sessionKey.Clear()
            End If
        End Try

    End Sub 'Encrypt



    Public Shared Sub Decrypt(ByVal Doc As XmlDocument, ByVal Alg As RSA, ByVal KeyName As String)

        Try
            ' Check the arguments.  
            If Doc Is Nothing Then
                Throw New ArgumentNullException("Doc")
            End If
            If Alg Is Nothing Then
                Throw New ArgumentNullException("Alg")
            End If
            If KeyName Is Nothing Then
                Throw New ArgumentNullException("KeyName")
            End If
            ' Create a new EncryptedXml object.
            Dim exml As New EncryptedXml(Doc)

            ' Add a key-name mapping.
            ' This method can only decrypt documents
            ' that present the specified key name.
            exml.AddKeyNameMapping(KeyName, Alg)

            ' Decrypt the element.
            exml.DecryptDocument()
        Catch e As Exception
            WriteError("Error in clsEncyption.Encrypt  " & e.Message)
            ' re-throw the exception.
            Throw
        Finally

        End Try
    End Sub 'Decrypt 






    Public Shared Sub GenKey_SaveInContainer(ByVal ContainerName As String)
        ' Create the CspParameters object and set the key container   
        ' name used to store the RSA key pair.  
        Dim cp As New CspParameters()
        cp.KeyContainerName = ContainerName

        ' Create a new instance of RSACryptoServiceProvider that accesses  
        ' the key container MyKeyContainerName.  
        Dim rsa As New RSACryptoServiceProvider(cp)

        ' Display the key information to the console.  
        Debug.WriteLine("Key added to container:  {0}", rsa.ToXmlString(True))
    End Sub

    Public Shared Sub GetKeyFromContainer(ByVal ContainerName As String)
        ' Create the CspParameters object and set the key container   
        '  name used to store the RSA key pair.  
        Dim cp As New CspParameters()
        cp.KeyContainerName = ContainerName

        ' Create a new instance of RSACryptoServiceProvider that accesses  
        ' the key container MyKeyContainerName.  
        Dim rsa As New RSACryptoServiceProvider(cp)

        ' Display the key information to the console.  
        Debug.WriteLine("Key retrieved from container : {0}", rsa.ToXmlString(True))
    End Sub

    Public Shared Sub DeleteKeyFromContainer(ByVal ContainerName As String)
        ' Create the CspParameters object and set the key container   
        '  name used to store the RSA key pair.  
        Dim cp As New CspParameters()
        cp.KeyContainerName = ContainerName

        ' Create a new instance of RSACryptoServiceProvider that accesses  
        ' the key container.  
        Dim rsa As New RSACryptoServiceProvider(cp)

        ' Delete the key entry in the container.  
        rsa.PersistKeyInCsp = False

        ' Call Clear to release resources and delete the key from the container.  
        rsa.Clear()

        Debug.WriteLine("Key deleted.")
    End Sub


    Sub EncryptKey()
        'Initialize the byte arrays to the public key information.  
        Dim PublicKey As Byte() = {214, 46, 220, 83, 160, 73, 40, 39, 201, 155, 19, 202, 3, 11, 191, 178, 56, 74, 90, 36, 248, 103, 18, 144, 170, 163, 145, 87, 54, 61, 34, 220, 222, 207, 137, 149, 173, 14, 92, 120, 206, 222, 158, 28, 40, 24, 30, 16, 175, 108, 128, 35, 230, 118, 40, 121, 113, 125, 216, 130, 11, 24, 90, 48, 194, 240, 105, 44, 76, 34, 57, 249, 228, 125, 80, 38, 9, 136, 29, 117, 207, 139, 168, 181, 85, 137, 126, 10, 126, 242, 120, 247, 121, 8, 100, 12, 201, 171, 38, 226, 193, 180, 190, 117, 177, 87, 143, 242, 213, 11, 44, 180, 113, 93, 106, 99, 179, 68, 175, 211, 164, 116, 64, 148, 226, 254, 172, 147}

        Dim Exponent As Byte() = {1, 0, 1}

        'Create values to store encrypted symmetric keys.  
        Dim EncryptedSymmetricKey() As Byte
        Dim EncryptedSymmetricIV() As Byte

        'Create a new instance of the RSACryptoServiceProvider class.  
        Dim RSA As New RSACryptoServiceProvider()

        'Create a new instance of the RSAParameters structure.  
        Dim RSAKeyInfo As New RSAParameters()

        'Set RSAKeyInfo to the public key values.   
        RSAKeyInfo.Modulus = PublicKey
        RSAKeyInfo.Exponent = Exponent

        'Import key parameters into RSA.  
        RSA.ImportParameters(RSAKeyInfo)

        'Create a new instance of the RijndaelManaged class.  
        Dim RM As New RijndaelManaged()

        'Encrypt the symmetric key and IV.  
        EncryptedSymmetricKey = RSA.Encrypt(RM.Key, False)
        EncryptedSymmetricIV = RSA.Encrypt(RM.IV, False)
    End Sub




#Region "IDisposable Support"
    Private disposedValue As Boolean ' To detect redundant calls

    ' IDisposable
    Protected Overridable Sub Dispose(disposing As Boolean)
        If Not Me.disposedValue Then
            If disposing Then
                ' TODO: dispose managed state (managed objects).
                fsInput.Dispose()
                fsOutput.Dispose()
                TripleDes.Dispose()
            End If

            fsInput = Nothing
            fsOutput = Nothing
            TripleDes = Nothing
            ' TODO: free unmanaged resources (unmanaged objects) and override Finalize() below.
            ' TODO: set large fields to null.

        End If
        Me.disposedValue = True
    End Sub

    ' TODO: override Finalize() only if Dispose(ByVal disposing As Boolean) above has code to free unmanaged resources.
    'Protected Overrides Sub Finalize()
    '    ' Do not change this code.  Put cleanup code in Dispose(ByVal disposing As Boolean) above.
    '    Dispose(False)
    '    MyBase.Finalize()
    'End Sub

    ' This code added by Visual Basic to correctly implement the disposable pattern.
    Public Sub Dispose() Implements IDisposable.Dispose
        ' Do not change this code.  Put cleanup code in Dispose(disposing As Boolean) above.
        Dispose(True)
        GC.SuppressFinalize(Me)
    End Sub
#End Region

End Class



