Public Class clsCSV
    'I would suggest using the TextFieldParser to read in the CSV file.  You can then load the values into a DataTable object and take advatange of its sorting capabilities.  The sorted DataTable can then be written back out to a csv.
    'Here is an example of a simple implementation of the above idea.  This code sorts 200,000 of your example records in about 4.5 seconds - not sure if that is fast enough for your application or not, but this code could probably be sped up and other tactics could also be faster.  But this was meant to be easy to follow for a beginner:



    'Sorts the csv by turning it into a datatable, applying a view filter, and then writing the result back to a csv  
    Public Sub SortCsv(ByVal sourceFile As String, ByVal destinationFile As String, ByVal ParamArray sortColumns() As Integer)
        Dim dt As DataTable = CsvToTable(sourceFile)
        Try
            If sortColumns.Length > 0 Then
                Dim sortStr As String = String.Empty
                For i As Integer = 0 To sortColumns.Length - 1
                    If sortStr.Length > 0 Then sortStr &= ", "
                    sortStr &= "Column" & sortColumns(i).ToString
                Next
                dt.DefaultView.Sort = sortStr

            End If
            TableToCSV(dt.DefaultView.ToTable, destinationFile)
        Catch e As Exception
            WriteError("Error in clsCSV.SortCsv  " & e.Message)
            '            Debug.WriteLine(e.Message)
        End Try
    End Sub

    'Parses a csv into a datatable  
    Public Function CsvToTable(ByVal filePathName As String) As DataTable
        Dim result As New DataTable
        Try
            If System.IO.File.Exists(filePathName) Then
                Dim parser As New Microsoft.VisualBasic.FileIO.TextFieldParser(filePathName)
                parser.Delimiters = New String() {","}
                parser.HasFieldsEnclosedInQuotes = True 'use if data may contain delimiters  
                parser.TextFieldType = FileIO.FieldType.Delimited
                parser.TrimWhiteSpace = True


                While Not parser.EndOfData
                    AddValuesToTable(parser.ReadFields, result)
                End While

                parser.Close()
            End If
            Return result
        Catch e As Exception
            WriteError("Error in clsCSV.CsvToTable  " & e.Message)
            '            Debug.WriteLine(e.Message)
            Return Nothing
        End Try
    End Function

    'Writes a datatable back into a csv  
    Private Sub TableToCSV(ByVal sourceTable As DataTable, ByVal filePathName As String)
        Dim sb As New System.Text.StringBuilder
        Try
            For Each dr As DataRow In sourceTable.Rows
                sb.AppendLine(String.Join(",", Array.ConvertAll(Of Object, String)(dr.ItemArray, _
                Function(o As Object) If(o.ToString.Contains(","), _
                ControlChars.Quote & o.ToString & ControlChars.Quote, o.ToString))))
            Next
            System.IO.File.WriteAllText(filePathName, sb.ToString)
        Catch e As Exception
            WriteError("Error in clsCSV.SourcTable  " & e.Message)
            Debug.WriteLine(e.Message)
        End Try
    End Sub

    'Ensures a datatable can hold an array of values and then adds a new row  
    Private Sub AddValuesToTable(ByVal source() As String, ByVal destination As DataTable)
        Dim existing As Integer = destination.Columns.Count
        Try
            For i As Integer = 0 To source.Length - existing - 1
                destination.Columns.Add("Column" & (existing + 1 + i).ToString, GetType(String))
            Next
            destination.Rows.Add(source)
        Catch e As Exception
            WriteError("Error in clsCSV.AddValuesToTable  " & e.Message)
            Debug.WriteLine(e.Message)
        End Try
    End Sub
End Class
