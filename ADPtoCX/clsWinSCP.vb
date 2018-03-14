Imports WinSCP
Imports System
Imports System.IO
Imports System.Reflection
Imports System.Reflection.Assembly
Imports System.Reflection.Emit.AssemblyBuilder

Public Class clsWinSCP

    Public Function GetCSV() As Integer

        'Dim tempName As String = Path.GetTempFileName
        'Dim executableName = "WinSCP" & Path.ChangeExtension(Path.GetFileName(tempName), "exe")
        'Dim executablePath = Path.Combine(Path.GetDirectoryName(tempName), executableName)
        'File.Delete(tempName)

        'Dim t As Type = GetType(clsWinSCP)
        'Dim executingAssembly As Assembly = t.Assembly
        'Dim resourceName = executingAssembly.GetName().Name + "." + "WinSCP.exe"


        Try
            ' Setup session options
            Dim sessionOptions As New SessionOptions
            With sessionOptions
                .Protocol = Protocol.Sftp
                .HostName = gl.FTPHost
                .UserName = gl.FTPUser
                .Password = gl.FTPPwd
                .SshHostKeyFingerprint = gl.FTPKey
            End With

            Using session As New Session
                ' Connect
                session.Open(sessionOptions)

                ' Upload files
                Dim transferOptions As New TransferOptions
                transferOptions.TransferMode = TransferMode.Binary

                Dim transferResult As TransferOperationResult

                transferResult = session.GetFiles("/adp/ADPtoCX.csv", gl.FileNamePath & gl.FileIn, False, transferOptions)
                ' Throw on any error
                transferResult.Check()

                ' Print results
                For Each transfer In transferResult.Transfers
                    WriteLog("CSV File sucessfully transferred from FTP site to working folder")
                    'Console.WriteLine("Upload of {0} succeeded", transfer.FileName)
                Next

                ' Verify that the new file is on Teddy
                If File.Exists(gl.FileNamePath & gl.FileIn) Then
                    ' If so Remove the file from the FTP server
                    'session.RemoveFiles("/adp/ADPtoCX.csv")
                Else
                    ' Else, throw an error
                    Throw New Exception("ADPtoCX.csv File did not copy to working location.")
                End If

                ' session.Close()
            End Using

            Return 0
        Catch e As Exception
            'Console.WriteLine("Error: {0}", e)
            WriteError(e.Message)
            Return 1
        Finally
            GC.Collect()
        End Try

    End Function

End Class