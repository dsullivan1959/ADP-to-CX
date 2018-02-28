Public Class TreeView

    Private Sub TreeView_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Dim fd As OpenFileDialog = New OpenFileDialog
        Dim strFileName As String

        fd.Title = "Select File"
        fd.InitialDirectory = My.Application.Info.DirectoryPath
        fd.Filter = "All files (*.*)|*.*|All files (*.*)|*.*"
        fd.FilterIndex = 2
        fd.RestoreDirectory = True

        If fd.ShowDialog() = DialogResult.OK Then
            strFileName = fd.FileName
        End If

    End Sub
    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnSelect.Click
        'Get a list of drives
        Me.Show()

        Dim drives As System.Collections.ObjectModel.ReadOnlyCollection(Of IO.DriveInfo) = My.Computer.FileSystem.Drives
        Dim rootDir As String = String.Empty
        'Now loop thru each drive and populate the treeview
        For i As Integer = 0 To drives.Count - 1

            rootDir = My.Application.Info.DirectoryPath

            rootDir = drives(i).Name
            'Add this drive as a root node
            TreeView1.Nodes.Add(rootDir)
            'Populate this root node
            PopulateTreeView(rootDir, TreeView1.Nodes(i))
        Next

    End Sub

    Private Sub PopulateTreeView(ByVal dir As String, ByVal parentNode As TreeNode)
        Dim folder As String = String.Empty
        Try
            Dim folders() As String = IO.Directory.GetDirectories(dir)
            If folders.Length <> 0 Then
                Dim childNode As TreeNode = Nothing
                For Each folder In folders
                    childNode = New TreeNode(folder)
                    parentNode.Nodes.Add(childNode)
                    PopulateTreeView(folder, childNode)
                Next
            End If
        Catch ex As UnauthorizedAccessException
            parentNode.Nodes.Add(folder & ": Access Denied")
        End Try
    End Sub
End Class