Imports WinSCP

Public Class clsWinSCP


    Public Function GetCSV() As Integer

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

                '    session.PutFiles(gl.AppPath, "/home/user/", False, transferOptions)
                transferResult =
                session.GetFiles("/adp/ADPtoCX.csv", gl.FileNamePath & gl.FileIn, False, transferOptions)
                ' Throw on any error
                transferResult.Check()

                ' Print results
                For Each transfer In transferResult.Transfers
                    Console.WriteLine("Upload of {0} succeeded", transfer.FileName)
                Next


                'session.RemoveFiles("/adp/ADPtoCX.csv")

                session.Close()
                GC.Collect()
            End Using

            Return 0
        Catch e As Exception
            'Console.WriteLine("Error: {0}", e)
            WriteError(e.Message)
            Return 1
        End Try

    End Function

End Class