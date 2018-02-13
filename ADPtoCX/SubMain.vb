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
Imports System.Linq
Imports System.Linq.Queryable


Module SubMain

    Public ConnInformix As New Odbc.OdbcConnection
    Public gl As New clsGlobals


    Public Sub Main()

        Call GetParams()

        Dim cSM As New SendMessage
        'Dim Msg As String = "test email message"
        'Call cSM.SendEMessage("ADP to CX Test", Msg, gl.FromAddress, gl.ToAddress, "")

        '+++++++++++++++++++++++++++++++++++++++++++
        '+++++++++++++++++++++++++++++++++++++++++++
        'To test the email routine...need to talk to Max
        '+++++++++++++++++++++++++++++++++++++++++++
        '+++++++++++++++++++++++++++++++++++++++++++

        Dim d As DateTime = Date.Now.AddDays(-1)
        Dim sFullFileNamePath As String = gl.FileNamePath & gl.FileIn
        Dim sFullOldFileNamePath As String = gl.FileNamePath & gl.OldFile
        Dim sArchiveFile As String = Left(gl.FileIn, Len(gl.FileIn) - 4)
        Dim sFullArchiveFileNamePath As String = gl.FileNamePath & "\Archive" & sArchiveFile & Format(d, "MMddyy hhmmss") & ".csv"

        If File.Exists(sFullFileNamePath) = True Then

            Call QryCSV()

            Try
                '+++++++++++++++++++++++++++++++++++++++++++++++++++++++++
                'Encrypt
                'Do we need to do this at all?  IF no SSN returned?
                'After all the processing is done, encrypt the files
                'Involves reading in the contents of the file, encrypting and then writing back out

                'Dim plainText As String = ""
                'Dim password As String = "password"   'This will be changed to an encrypted value in an ini file

                'plainText = My.Computer.FileSystem.ReadAllText(sFullFileNamePath)
                'Dim wrapper As New EncryptTxt(password)
                'Dim cipherText As String = wrapper.EncryptData(plainText)
                'My.Computer.FileSystem.WriteAllText(sFullFileNamePath, cipherText, False)



                'decrypt
                'cipherText = My.Computer.FileSystem.ReadAllText(My.Computer.FileSystem.SpecialDirectories.MyDocuments & "\cipherText.txt")
                'Dim password As String = InputBox("Enter the password:")
                '            Dim wrapper As New EncryptTxt(password)

                ' DecryptData throws if the wrong password is used.
                'Try
                '    plainText = wrapper.DecryptData(cipherText)
                '    MsgBox("The plain text is: " & plainText)
                'Catch ex As System.Security.Cryptography.CryptographicException
                '    MsgBox("The data could not be decrypted with the password.")
                'End Try

                '+++++++++++++++++++++++++++++++++++++++++++++++++++++++++

                'At end of process, rename file and archive
                'If new file was created, it should not have a "LAST" addition to the file name
                If File.Exists(sFullFileNamePath) Then
                    File.Copy(sFullFileNamePath, sFullArchiveFileNamePath, True)
                    File.Copy(sFullFileNamePath, sFullOldFileNamePath, True)
                    File.Delete(sFullFileNamePath)
                End If

                'IF SSN is included, may need to encrypt and decrypt my files
                'So probably would encrypt at this point, and decrypt just before the file compare.

            Catch e As Exception
                WriteError("Error in Main.  Error = " & e.Message)

                Debug.Print(e.Message)
            Finally
                WriteLog("-----------------")
                'Dim Msg As String = "test email message"
                Call cSM.SendEMessage("ADP to CX Test", "ADP to CX process complete " & CStr(Today), gl.FromAddress, gl.ToAddress, "")
                Call EndProgram()
            End Try
        Else
            WriteLog("No new file")
        End If

    End Sub
    

    Public Sub EndProgram()
        GC.Collect()
        End
    End Sub

    Public Function TrimPhone(ByVal sPhone As String) As Double
        Dim i As Short
        Dim rstStr As String
        rstStr = ""
        For i = 1 To Len(sPhone)
            If IsNumeric(Mid(sPhone, i, 1)) Then
                rstStr = rstStr & Mid(sPhone, i, 1)
            End If
        Next i
        If Len(rstStr) < 10 Then
            TrimPhone = 0.0#
        Else
            TrimPhone = Val(rstStr)
        End If
    End Function

    Public Sub QryCSV()
        Dim c As New clsCSV
        Dim tbl As New DataTable
        Dim oldTbl As New DataTable
        Dim ChangeTbl As New DataTable
        Dim foundRows() As DataRow

        Dim totCnt As Integer = 0
        Dim sFile As String = ""
        Dim sLName As String = ""
        Dim sStatus As String = ""
        Dim sTermDate As String = ""
        Dim sPCN As String = ""
        Dim sPayCode As String = ""
        Dim cols As Integer = 0
        Dim iCnt As Integer = 0
        Dim jcnt As Integer = 0
        Dim rowAsArray As String = ""
        Dim rowAsArray2 As String = ""
        Dim equal As Boolean = True



        Try


            ''+++++++++++++++++++++++++++++++++++++++++++
            ''+++++++++++++++++++++++++++++++++++++++++++
            ''IF SSN is included, may need to encrypt and decrypt the data files
            ''The trick will be when and where to encrypt or decrypt each file
            ''So probably would decrypt at this point, and encrypt in the end program routine.

            ''New file from ADP would probably be encrypted at the end of the program
            ''Last file would need to be decrypted here.
            ''Else, I may need a way to indicate whether the file is encrypted or not

            ''Encrypt
            'Dim plainText As String
            'Dim password As String = "password"   'obviously need to use something else here...
            'Dim wrapper As New EncryptTxt(password)

            ''Here I will read the file
            'Dim ALLText As String = My.Computer.FileSystem.ReadAllText(gl.FileNamePath & gl.FileIn)
            ''This will encrypt the entire file
            'Dim cipherText As String = wrapper.EncryptData(ALLText)
            ''This will write the encrypted data to a file
            'My.Computer.FileSystem.WriteAllText(My.Computer.FileSystem.SpecialDirectories.MyDocuments & "\cipherText.txt", cipherText, False)


            ''To reverse the process - decrypt
            ''Read the encrypted file 
            'cipherText = My.Computer.FileSystem.ReadAllText(My.Computer.FileSystem.SpecialDirectories.MyDocuments & "\cipherText.txt")
            ''Dim password As String = InputBox("Enter the password:")
            ''Dim wrapper As New EncryptTxt(password)

            '' DecryptData throws if the wrong password is used.
            'Try
            '    'this will decrypt the entire file
            '    plainText = wrapper.DecryptData(cipherText)
            '    Catch ex As System.Security.Cryptography.CryptographicException
            '    MsgBox("The data could not be decrypted with the password.")
            'End Try


            ''+++++++++++++++++++++++++++++++++++++++++++
            ''+++++++++++++++++++++++++++++++++++++++++++
            'tbl = c.CsvToTable("\\10.7.118.112\ADP_CX_PROG\bin\DataFiles\ADP to CX.csv")


            tbl = c.CsvToTable(gl.FileNamePath & gl.FileIn)
            oldTbl = c.CsvToTable(gl.FileNamePath & gl.OldFile)


            If tbl.Rows.Count > 0 Then
                frmStatus.Show()
                'frmStatus.Hide()

                frmStatus.txtCount.Text = 0
                frmStatus.ProgressBar1.Value = 0
                frmStatus.ProgressBar1.Minimum = 0
                frmStatus.ProgressBar1.Maximum = tbl.Rows.Count

                For i As Integer = 0 To tbl.Rows.Count - 1
                    rowAsArray = ""
                    rowAsArray2 = ""
                    'If i = 1309 Then
                    '    WriteLog(i)
                    'End If

                    If i = 0 Then
                        'do nothing - ignore header line
                    ElseIf tbl.Rows(i)(65) = "" And tbl.Rows(i)(58) <> "DPW" Then     'No PCN information
                        'do nothing unless DPW student job, check only the CVID info
                    Else


                        Dim row1 As DataRow = tbl.Rows(i)
                        '============

                        'Note:   The select method against the datatable gets confusing with data types.   Best to assign the values to a declared variable
                        sFile = tbl.Rows(i)(0)      'Use ADP file number - not all records have a carthage ID
                        Debug.WriteLine(oldTbl.Rows(i)(1))
                        sLName = tbl.Rows(i)(2)
                        sStatus = tbl.Rows(i)(51)
                        sTermDate = tbl.Rows(i)(50).ToString
                        sPCN = tbl.Rows(i)(65)
                        sPayCode = tbl.Rows(i)(58)

                        'First search last record to see if employee exists
                        Dim expression As String   'This is creating something like a SQL Statement
                        'Because there are a handful of duplicate file numbers, need to check a couple of other fields for the initial match
                        expression = "Column1 = " & ConvertString(sFile) & " and column59 = " & ConvertString(sPayCode) & " and column66 = " & ConvertString(sPCN) ' 65 is the PCN Code, 
                        'Compare two csv files using the file number as the key
                        ' Use the Select method to find all rows matching the filter.
                        foundRows = oldTbl.Select(expression)
                        
                        'If a failure occurs, how will the fix happen?   If I overwrite the "Last" csv, then the change won't be picked up the next day.
                        'One option would be to create a third file with failures, write the the table row to a csv file
                        'The next day, read that file and do something with it, delete the record if success occurs.
                        'Or assume that HR made a mistake and needs to fix manually...based on an email report?

                        If foundRows.Length = 0 Then
                            Debug.WriteLine(tbl.Rows(i)(0).ToString)
                            'This is a brand new record - add to all tables
                            Call ParseADPData(tbl.Rows(i))
                        Else
                            'If there is an employee match, determine if anything has changed in the file
                            'Convert the data row to a string for comparison
                            rowAsArray = String.Join(",", tbl.Rows(i).ItemArray.ToArray())

                            For n As Integer = 0 To oldTbl.Columns.Count - 1
                                rowAsArray2 = rowAsArray2 & foundRows(0)(n)
                                If n < oldTbl.Columns.Count - 1 Then
                                    rowAsArray2 = rowAsArray2 + ","
                                End If
                                Console.WriteLine(foundRows(0)(n))
                            Next


                            If rowAsArray = rowAsArray2 Then
                                'No change - move on to the next record
                            Else
                                'See if the two arrays are the same - no data changed
                                Call ParseADPData(tbl.Rows(i))

                            End If

                        End If

                        '======================

                        iCnt = iCnt + 1
                    End If
                        totCnt = totCnt + 1
                        frmStatus.txtCount.Text = totCnt.ToString & " of " & tbl.Rows.Count.ToString
                        frmStatus.ProgressBar1.Value = totCnt
                        frmStatus.Refresh()
                Next  'Loop for new file to test against old

            End If
            frmStatus.Dispose()
            frmStatus = Nothing

        Catch e As Exception
            WriteError("Error in QryCSV.  Error = " & e.Message)

            Console.Write(e.Message)
        Finally
        End Try
    End Sub


    Private Function AreArraysEqual(Of T)(ByVal a As T(), ByVal b() As T) As Boolean

        'IF 2 NULL REFERENCES WERE PASSED IN, THEN RETURN TRUE, YOU MAY WANT TO RETURN FALSE
        If a Is Nothing AndAlso b Is Nothing Then Return True

        'CHECK THAT THERE IS NOT 1 NULL REFERENCE ARRAY
        If a Is Nothing Or b Is Nothing Then Return False

        'AT THIS POINT NEITHER ARRAY IS NULL
        'IF LENGTHS DON'T MATCH, THEY ARE NOT EQUAL
        If a.Length <> b.Length Then Return False

        'LOOP ARRAYS TO COMPARE CONTENTS
        For i As Integer = 0 To a.GetUpperBound(0)
            'RETURN FALSE AS SOON AS THERE IS NO MATCH
            If Not a(i).Equals(b(i)) Then Return False
        Next

        'IF WE GOT HERE, THE ARRAYS ARE EQUAL
        Return True

    End Function


    Public Sub ParseADPData(fileFields As DataRow)
        ' version for reading text file
        Dim clsCVID As New Cls_CV_IDrec
        Dim clsID As New ClsID_rec
        Dim clsPos As New ClsPos_rec
        Dim clsProf As New clsProfile_rec
        Dim sMl As New SendMessage

        Dim sql As String
        Dim iCarthId As Integer
        Dim Row1 As String

        Dim cADPName As New clsADPName
        Dim cPCN As New ClsPCN
        Dim sCarthNameID As Integer
        Dim sJobCode As String = ""
        Dim sFacStaffFlag As String = ""


        '++++++++++++++++++++++++++++++++++++++++++++++++++++++++
        ' New for provisioning
        '++++++++++++++++++++++++++++++++++++++++++++++++++++++++
        'Try
        '    ' Write the entire row to the ADP_rec table for provisioning
        'Catch ex As Exception

        'End Try
        '++++++++++++++++++++++++++++++++++++++++++++++++++++++++
        '++++++++++++++++++++++++++++++++++++++++++++++++++++++++

        Try
            'CSV is problematic because of the commas within certain fields 
            'Maybe they can submit as a pipe delimited file
            'If not, 
            Row1 = String.Empty

            'Assumption should be that the ID exists in the ID_Rec table
            iCarthId = 0  'initialize test variable
            sql = "Select id_rec.id, id_rec.fullname From id_rec Where id_rec.id = ?"
            Dim cmd As New Odbc.OdbcCommand(sql, ConnInformix)
            If fileFields(1) = "" Then
                Throw New Exception("No Carthage ID, process aborted.  Name = " & RemoveQuotes(Trim(fileFields(6))) & " ADPID = " & fileFields(1).ToString)
            End If
            cmd.Parameters.AddWithValue("@id_rec.id", CInt(fileFields(1)))   'Looking for Carthage ID - Position 0 is the ADP number, position 1 is the Carthage number
            ConnInformix.ConnectionString = gl.ConnectString
            ConnInformix.Open()
            Dim reader As Odbc.OdbcDataReader
            reader = cmd.ExecuteReader()
            If reader.HasRows Then
                Do While (reader.Read())
                    iCarthId = reader.GetDouble(0)   '
                Loop
            Else
                iCarthId = 0
            End If
            reader.Close()
            ConnInformix.Close()
            ConnInformix.Dispose()
            Debug.Print(fileFields(0))

            sCarthNameID = 0

            sJobCode = fileFields(58).ToString
            sFacStaffFlag = GetFacStaffFlag(fileFields(65))
            'K97 is hourly staff, FVW can be faculty or staff.
            'Don't do anything with Faculty except make sure CVID is populated
            'DPW is student - there are no job PCN codes in ADP and CX is the "home" of that info
            'Only check  CVID for DPW

            If (sJobCode = "K97") Or (sJobCode = "FVW" And sFacStaffFlag = "Staff") Then
                clsID.Initialize()
                clsID.ID = Val(fileFields(1))   ' This field is mostly unpopulated.  How do we get the ID?

                'another chance to find the Carthage ID using the name of the individual, gender and birthdate
                clsID.FullName = RemoveQuotes(Trim(fileFields(6)))
                clsID.LName = Trim(RemoveQuotes(fileFields(2)))
                clsID.MidName = Trim(RemoveQuotes(fileFields(4)))
                clsID.FName = Trim(RemoveQuotes(fileFields(3)))

                'GoTo 100   'temporary debug item
                Dim t As String
                t = RemoveQuotes(fileFields(11))
                Debug.Print(Len(t))

                clsID.NameSndx = ""  'I think this is room number - N/A
                clsID.Title = fileFields(5)
                clsID.Suffix = ""  'n/a
                clsID.SSNum = fileFields(43)  'Use XXX-XX-9999 - something required for Provisioning processes
                clsID.PhoneExt = "" 'n/a
                clsID.PrevNameID = 0
                clsID.Mail = ""  'N/A
                clsID.Sol = "" 'N/A
                clsID.Pub = "" 'N/A
                clsID.CorrectAddr = ""  'N/A
                clsID.Decsd = "N"  'Presumably not deceased.  Set to "N"
                clsID.AddDate = fileFields(45)
                clsID.UpDte = fileFields(48)  'Effective Date?
                clsID.Valid = "" 'n/a
                clsID.PurgeDate = "" 'n/a
                clsID.CassCertDate = "" 'n/a
                clsID.CCUserName = "" 'n/a
                clsID.CCPwd = "" 'n/a
                clsID.InquiryNum = 0
                clsID.OffcAddBy = "HR"

                If iCarthId <> 0 Then
                    Debug.Print("ID Exists")
                    clsID.Update_Basic()
                Else
                    'WriteLog("Carthage ID Not recognized")
                    Throw New Exception("No Carthage ID, process aborted.  Name = " & clsID.FullName & " ADPID = " & fileFields(1).ToString)
                    'MAIL Event
                    sMl.SendEMessage("ADP to CX", ("No Carthage ID, process aborted for " & clsID.FullName & "ID Number = " & fileFields(1).ToString), gl.FromAddress, gl.SupportAddress, "")
                    '                    sMl.SendEMessage("ADP to CX", ("No Carthage ID, process aborted.  Name = " & clsID.FullName), gl.FromAddress, gl.SupportAddress, "")
                End If

                'Separate all the Address fields out and do a second insert/update for those.
                'Archive an old address
                'Should double check for changes, don't Archive if address is not the changed part of the record.
                'If address unchanged, only update job_rec fields
                'if address changed, update and add new aa_rec
                'Does not apply to a new record
                clsID.SuffixName = ""  'n/a
                clsID.Addr1 = Trim(fileFields(16))   'csv has an odd vbcrlf
                clsID.Addr2 = Trim(fileFields(17))
                clsID.Addr3 = Trim(fileFields(18))
                clsID.City = Trim(fileFields(19))
                clsID.State = Trim(fileFields(20))
                clsID.Zip = Trim(fileFields(22))
                clsID.Ctry = Trim(fileFields(25))
                clsID.CorrectAddr = "Y"
                clsID.AA = "PERM"
                clsID.Phone = fileFields(27)  'Home Phone in ADP

                Debug.Print(clsID.AA)

                If iCarthId <> 0 Then
                    Dim sAddrChg As String
                    Debug.Print("ID Exists")
                    'Find out if the address has changed

                    sAddrChg = CheckAddress(clsID.ID, clsID.FullName, clsID.Addr1, clsID.Addr2, clsID.Addr3, clsID.City, clsID.State, clsID.Zip, clsID.Ctry)
                    If sAddrChg = "True" Then
                        clsID.Update_Addr()
                    End If
                Else
                    WriteLog("No Carthage ID, process aborted for " & clsID.FullName & " ADPID = " & fileFields(1).ToString)
                    Throw New Exception("No Carthage ID, process aborted.  Name = " & clsID.FullName & " ADPID = " & fileFields(1).ToString)
                    'MAIL Event
                    sMl.SendEMessage("ADP to CX", ("No Carthage ID, process aborted for " & clsID.FullName & "ID Number = " & fileFields(1).ToString), gl.FromAddress, gl.SupportAddress, "")
                    '                    sMl.SendEMessage("ADP to CX", ("No Carthage ID, process aborted.  Name = " & clsID.FullName), gl.FromAddress, gl.SupportAddress, "")
                End If


                '++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
                '++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
                'Deal with personal email???
                If Not IsDBNull(fileFields(15)) And fileFields(15) <> "" Then SetEmail2(fileFields(15), Val(fileFields(1)))
                'Work Email???  'This gets filled in provisioning process, did not plan to touch it originally
                ' filefields(31)
                'aa_rec as eml1
                '++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
                '++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++


                '++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
                '++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
                'Deal with cell phone?
                If Not IsDBNull(fileFields(28)) And fileFields(28) <> "" Then
                    Set_CellPhone(fileFields(28), Val(fileFields(1)), RemoveQuotes(Trim(fileFields(6))))
                    'aa_audit_rec???
                End If
                ' aa_audit_rec???  
                '++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
                '++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

            End If

            If (sJobCode = "K97") Or (sJobCode = "FVW" And sFacStaffFlag = "Staff") Then
                'then deal with profile rec
                'Most of this we won't have at this point.
                'Danger of overwriting existing info. 
                'Update statement will only touch fields that have a value
                Dim ProfileExists As String
                ProfileExists = ValidateField(clsID.ID, "profile_rec", "id", "integer")
                clsProf.Initialize()
                clsProf.ID = clsID.ID  ' Check for duplicate here.
                clsProf.BirthCity = ""
                clsProf.BirthCntry = ""

                Debug.Print(fileFields(8).ToString)

                clsProf.BirthDate = ConvertDATE(fileFields(8))   'FILEFIELD(8) Something is required for provisioning.  Use 1/1/99 to fake out the system.
                If fileFields(8) > "" Then
                    clsProf.Age = Today.Year - CDate(fileFields(8)).Year
                Else
                    clsProf.Age = 0
                End If
                clsProf.BirthState = ""
                clsProf.Citz = ""
                clsProf.ConfirmDate = ""
                clsProf.DecsDate = ""
                clsProf.DenomCode = ""
                clsProf.ChurchID = 0
                clsProf.EthnicCode = ""  ' Maype position 24
                clsProf.GrpNo = 0
                clsProf.HandicapCode = ""
                clsProf.Hispanic = ""  'Position 24
                clsProf.MilitCode = ""
                clsProf.MilitRank = ""
                clsProf.Mrtl = ""
                clsProf.News1ID = 0
                clsProf.News2ID = 0
                clsProf.Occ = ""
                clsProf.Photo = 0
                clsProf.PrivCode = ""
                clsProf.profResDate = ""
                clsProf.ProfLastUpd = Date.Today.ToString("MM-dd-yyyy")
                clsProf.ProfVetChap = ""
                clsProf.ProfVisaDate = ""
                'Need routine to validate race code

                If IsNumeric(fileFields(11)) Then
                    clsProf.Race = GetRaceCode(fileFields(11))  'Valid codes are AM, AP,AS,BA,BL,HI, IS,MU,N,NO,UN,WH,Y 
                Else
                    clsProf.Race = ""
                End If

                If fileFields(13) = "Not Hispanic or Latino" Then
                    clsProf.Hispanic = "N"
                ElseIf fileFields(13) = "HISPANIC OR LATINO" Then
                    clsProf.Hispanic = "N" 'Y,N,U
                Else
                    clsProf.Hispanic = ""
                End If
                clsProf.DecsDate = "01/01/1999"
                clsProf.ResCity = ""
                clsProf.ResCntry = ""
                clsProf.ResState = ""
                clsProf.Sex = Left(fileFields(9), 1)  'this is all we get from ADP, Col 13
                clsProf.Vet = ""
                clsProf.VisaCode = ""
                clsProf.FullName = RemoveQuotes(Trim(fileFields(6)))

                'Add what we can to profile rec
                If ProfileExists = "New" Then
                    clsProf.Insert()
                Else
                    clsProf.Update()
                End If

            End If

            'Provisioning needs the ADPID and may need a dummy SSNumber.
            Dim sCVID As String
            sCVID = ValidateField(Val(fileFields(1)), "cvid_rec", "cx_id", "integer")
            clsCVID.Initialize()
            clsCVID.CXID = Val(fileFields(1))   'Check for duplicate...
            clsCVID.OldId = fileFields(1)
            clsCVID.OldIdNum = Val(fileFields(1))
            clsCVID.ADPID = fileFields(0)
            clsCVID.LDAPName = ""
            clsCVID.LDAPAddDate = "01/01/1999"
            clsCVID.SSN = fileFields(43)
            clsCVID.CXIDst = fileFields(1)
            clsCVID.JobCodePlaceholder = sJobCode  'This is being passed for reporting purposes, not for table update
            clsCVID.FullName = RemoveQuotes(Trim(fileFields(6)))

            If sCVID = "New" Then
                clsCVID.Insert()
            Else
                clsCVID.Update() 'works
            End If

            '++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
            '++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
            'Here I may need to add code to deal with Home Department Code (numeric to function code) and Supervisor
            '  fileFields(89) ex. 405000 corresponds to function code 405
            'Function code is a char(3) field in func table
            'Func_area is a char(4) in pos_rec
            'ValidateField(Left(fileFields(89), 3), "func_table", "func", "string")
            'Validate that first three digits are a valid function code
            'Insert or update into pos_table, func_area field
            '++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
            '++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

            '++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
            '++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
            'As part of job record insert.update...
            'validate ID of supervisor
            'fileFields(91), example "FVW421547"  Must trim off first three characters
            'Supervisor number field is integer
            'insert or update job_rec table, supervisor_no
            'Question.  Do I use this field for all positions?  Even if not in department?
            '++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
            '++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

            'Deal with the main job record..

            If (sJobCode = "K97") Or (sJobCode = "FVW" And sFacStaffFlag = "Staff") Then
                Dim sLeaveDate As String = ""
                Dim sLeaveEnd As String = ""
                sLeaveDate = fileFields(76)

                'Create job record
                Dim iRet As Integer

                'id (1), pcnAgg(65),stDate(48), sEndDate(50), hrStat(61), hrPay(58), ADPTitle(66), Primary(57), FullName(6), Supervisor(91), Func(89) As String
                'EX: 123458, xxx-xxx-xxxxx, 1/1/17, FT or AD, FVW or DPW, Title
                iRet = Process_Job_rec(fileFields(1), CStr(fileFields(65)), fileFields(48), fileFields(50), CStr(fileFields(61)), fileFields(58), fileFields(66), _
                                       fileFields(57), Trim(fileFields(6)), Trim(fileFields(91)), Left(fileFields(89), 3))
                'Col 48 is position effective date, Col 50 is the termination date date

                '++++++++++++++++++++++++++++++++++++++++++++++++++++

                'Handle second or third jobs
                'id As Integer, pcnAgg As String, sStDate As String, sEndDate As String, pstnID As String, FullName As String, jRank As Integer
                If fileFields(77) <> "" Then
                    Call Process_2nd_Job_rec(fileFields(1), fileFields(77), fileFields(79), fileFields(80), CStr(fileFields(78)), Trim(fileFields(6)), 3)
                End If

                If fileFields(81) <> "" Then
                    Call Process_2nd_Job_rec(fileFields(1), fileFields(81), fileFields(83), fileFields(84), CStr(fileFields(82)), Trim(fileFields(6)), 4)
                End If

                If fileFields(85) <> "" Then
                    Call Process_2nd_Job_rec(fileFields(1), fileFields(85), fileFields(87), fileFields(88), CStr(fileFields(86)), Trim(fileFields(6)), 5)
                End If

                '++++++++++++++++++++++++++++++++++++++++++++++++++++
            End If

            Dim aaSCHL As String

            'Add ths "SCHL" record to the AA table IF the individual is employed, else add end date 
            'Use the primary position as the check...
            aaSCHL = AA_SCHL_REC(Val(fileFields(1)), RemoveQuotes(Trim(fileFields(6))), RemoveQuotes(Trim(fileFields(48))), RemoveQuotes(Trim(fileFields(50))))

        Catch e As Exception
            WriteError("Err in ParseADPData.  Error = " & e.Message & " ADP File Number = " & fileFields(0).ToString)
            'MAIL Event
            sMl.SendEMessage("ADP to CX", ("Err in ParseADPData.  Error = " & e.Message & " ADP File Number = " & fileFields(0).ToString), gl.FromAddress, gl.SupportAddress, "")
            '            sMl.SendEMessage("ADP to CX", ("Err in ParseADPData.  Error = " & e.Message & " ADP File Number = " & fileFields(0).ToString & " Name = " & clsID.FullName), gl.FromAddress, gl.ToAddress, "")

            Debug.WriteLine("Error in ParseADPData" & e.Message)

        Finally
            GC.Collect()
        End Try
    End Sub


    '================================================================
    'PCN Aggr  -   FAC-ARHU-THR-ASTP
    'Start Date
    'End Date
    'HrStat    (FT)
    'Positon ID   (FVW)
    'ADP Title     Asst Professor Theatre

    'Job Rec 
    'Need:  TposNo, ID, pcn_aggr, start date, hr_stat, hr_pay, ADPT Job Title
    '================================================================


    Public Function Process_Job_rec(id As Integer, pcnAgg As String, stDate As String, sEndDate As String, hrStat As String, _
                                    hrPay As String, ADPTitle As String, Primary As String, FullName As String, Supervisor As String, Func As String) As Integer
        'Here we have to deal with the job rec issues
        'See if valid in pos_table
        'If ADP doesn't have the correct codes, abort and send notification
        'We will use the PCN_AGGR field in the POS table, split it into the four components and validate against the job_rec table
        'ADM-ATDR-ATH-ADIR matches pcn_01, pcn_02, pcn_03, and pcn_04 
        'ADM is the position type.  Not sure what that relates to in CX.
        'ATDR is the division
        'ATH is the department
        'ADR is the position

        'These are for validating agains the root tables
        Dim clsJR As New ClsJob_rec
        Dim cPCN As New ClsPCN
        Dim sml As New SendMessage

        Dim sDpt, sDiv, sPos, jRecExists As String

        Try
            sDpt = ""
            sDiv = ""
            sPos = ""
            jRecExists = ""
            clsJR.Initialize()
            clsJR.TPosNo = CheckForTPos(pcnAgg)


            If clsJR.TPosNo <> 0 Then
                'Check for duplicates - both ID and Postion as a person can have more than one position
                clsJR.ID = CInt(id)

                '+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
                '                           IMPORTANT
                '+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
                'But now we also need to check for changes.   Need to use the ID, Hire date, effective date, termination date

                jRecExists = CheckForJobRec(clsJR.TPosNo, id, stDate)

                '+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
                '                           IMPORTANT
                '+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

                'if there is a record, use the tpos number to identify the job_rec
                'insert or update will key off of pos_table.tpos_no and id_rec.id

                If jRecExists <> "New" And Not IsDBNull(gl.JobCollection.Item(1)) Then
                    clsJR.JobNo = gl.JobCollection.Item(1)  'Appears to be a sequential identifier, but should be verified along with tpos_no  
                End If

                cPCN.PosType = CStr(pcnAgg)  'Need to work on this, setting up flow initially
                cPCN.Div = (CStr(pcnAgg))  'Need to work on this, setting up flow initially
                cPCN.Dept = (CStr(pcnAgg))  'Need to work on this, setting up flow initially
                'Div and Dept may not exist in these tables.  May only be in the job_rec 
                cPCN.Position = (CStr(pcnAgg))  'Need to work on this, setting up flow initially

                'If these return no matching field, error out and log
                sPos = ValidateField(cPCN.Position, "pos_table", "pcn_04", "string")
                sDiv = ValidateField(cPCN.Div, "div_table", "div", "string")
                sDpt = ValidateField(cPCN.Dept, "dept_table", "dept", "string")

                clsJR.HRDiv = cPCN.Div
                clsJR.HRDept = cPCN.Dept
                clsJR.Descr = GetPosDescr(clsJR.TPosNo)
                '
                '==================================
                If ValidateField(hrPay, "hrpay_table", "hrpay", "string") = "Exists" Then
                    clsJR.HRPay = hrPay
                Else
                    'MAIL Event
                    'sMl.SendEMessage("ADP to CX", ("Error in " & e.Message), gl.fromaddress, gl.toaddress, "")
                    Throw New Exception("Hrpay type does not exist in CX hrpay_table.  Item = " & hrPay)
                    WriteError("HR PAY Value doe not exist in CX.  ID Number =" & clsJR.ID & " Employee = " & FullName & "HRPay Value = " & hrPay)
                End If
                'Need to validate against HRPAY_REC???

                clsJR.BegDate = ConvertDATE(stDate)  'Format for insert?
                If sEndDate <> "" Then clsJR.EndDate = ConvertDATE(sEndDate)
                'clsJR.BJobNo = ""   'Sequential number...
                clsJR.ActiveCtrct = "N"   'Always seems to be "N" in CX - related to EEO
                clsJR.CtrctStat = "N/A"    ' related to EEO 
                clsJR.EGPType = "R"  'Earned Gross Pay.   Always set to R in database (a few blanks)
                clsJR.ExclSrvcHrs = "N"  'Always populated with N
                clsJR.HRStat = (hrStat)  'Match with job Function Code0
                clsJR.FullName = FullName
                If ADPTitle <> "" Then
                    clsJR.JobTitle = ADPTitle   'might not be available for secondary job titles
                Else
                    'Query for title from CX for secondary job title - not necessary for this routine.   Needed for custom fields
                    'Get Title and HR Stat by "SELECT distinct job_rec.title, job_rec.hrstat FROM job_rec Where job_rec.end_date Is Null and tpos_no = clsJR.TPosNo 
                End If
                'clsJR.PseudoSerial = 0  'n/a



                '++++++++++++++++++++++++++
                '  New for provisioning ca. January 2018
                clsJR.SuprvNo = Right(Supervisor, 6)
                '++++++++++++++++++++++++++

                If Primary = "Yes" And sEndDate = "" Then
                    clsJR.TitleRank = 1  'There is now a "primary" flag.  
                ElseIf Primary = "No" And sEndDate = "" Then
                    clsJR.TitleRank = 2
                ElseIf sEndDate <> "" Then
                    clsJR.TitleRank = Nothing
                End If
                'In most cases this will be set to 1
                'If the flag is NOT set to primary, set to 2
                'If there are third or fourth jobs, they will need to be set to 3 or 4 in the Process_Job_Rec2 routine for custom fields

            Else
                'Best to write to a log file the issue.
                'WriteLog("Position " & CheckForTPos(pstnID) & " does not exist in system.   Cannot be asigned to user " & id)
                Throw New Exception("Position does not exist in CX Position Table.  ID = " & id & " Employee = " & FullName & ", Position = " & hrPay & ", PCN_AGG = " & pcnAgg)
            End If


            If jRecExists = "Exists" Then
                'id As Integer, pcnAgg As String, sDate As Date, hrStat As String, pstnID As String, ADPTitle As String
                Debug.Print("Job Rec Exists")
                clsJR.Update()
            ElseIf jRecExists = "New" Then
                Debug.Print("New Job Rec")
                clsJR.Insert()  'Works
            Else
                WriteLog("Carthage ID Not recognized processing job record:  ID = " & id & " Employee = " & FullName)
            End If

            '========================================================
            ' new for provisioning/billing Jan 2018
            ' Add func code to position table -  
            '========================================================
            If ValidateField(Func, "func_table", "func", "string") = "Exists" Then
                Dim clPTbl As New ClsPosTable
                clPTbl.Initialize()
                clPTbl.TPos = clsJR.TPosNo
                clPTbl.FuncArea = Func
                clPTbl.PcnAgg = pcnAgg
                clPTbl.Descr = ADPTitle
                clPTbl.UpdateFunc()
                clPTbl = Nothing
                GC.Collect()
                'Use func and tpos_no
                'IF record exists, then should only need an update to the one field
            End If
            '========================================================
            '========================================================


        Catch ex As Exception
            WriteError("Error in Process_Job_rec: " & ex.Message & " Error = " & ex.Message & ", " & FullName & ", ID = " & id)
            'MAIL Event

            sml.SendEMessage("ADP to CX", "Error in Processing the Job record.  Error = " & ex.Message & ", " & FullName & ", ID = " & id, gl.FromAddress, gl.SupportAddress, "")
            '         sml.SendEMessage("ADP to CX", "Error in Processing the Job record.  Error = " & ex.Message, gl.FromAddress, gl.ToAddress, "")
            Process_Job_rec = 0
        End Try
        Process_Job_rec = 1
    End Function


    Public Function Process_2nd_Job_rec(id As Integer, pcnAgg As String, sStDate As String, sEndDate As String, pstnID As String, FullName As String, jRank As Integer) As Integer
        'Here we have to deal with the secondary job rec issues
        'See if valid in pos_table
        'If ADP doesn't have the correct codes, we will have to search somehow.
        'We will use the PCN_AGGR field in the POS table, split it into the four components and validate against the job_rec table
        'ADM-ATDR-ATH-ADIR matches pcn_01, pcn_02, pcn_03, and pcn_04 
        'ADM is the position type.  Not sure what that relates to in CX.
        'ATDR is the division
        'ATH is the department
        'ADR is the position

        Dim jRec2Exists As String = ""

        'These are for validating agains the root tables
        Dim clsJR2 As New ClsJob_rec
        Dim cPCN2 As New ClsPCN
        Dim sml As New SendMessage

        Dim sDpt, sDiv, sPos As String

        Try
            sDpt = ""
            sDiv = ""
            sPos = ""
            clsJR2.Initialize()
            clsJR2.TPosNo = CheckForTPos(pcnAgg)

            If clsJR2.TPosNo <> 0 Then
                'Check for duplicates - both ID and Postion as a person can have more than one position
                clsJR2.ID = CInt(id)

                '+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
                '                           IMPORTANT
                '+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
                'But now we also need to check for changes.   Need to use the ID, Hire date, effective date, termination date
                jRec2Exists = CheckForJobRec(clsJR2.TPosNo, id, sStDate)
                '+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
                '                           IMPORTANT
                '+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

                'if there is a record, use the tpos number to identify the job_rec
                'insert or update will key off of pos_table.tpos_no and id_rec.id
                If jRec2Exists <> "New" And Not IsDBNull(gl.JobCollection.Item(1)) Then
                    clsJR2.JobNo = gl.JobCollection.Item(1)  'Appears to be a sequential identifier, but should be verified along with tpos_no  
                End If

                cPCN2.PosType = CStr(pcnAgg)  'Need to work on this, setting up flow initially
                cPCN2.Div = (CStr(pcnAgg))  'Need to work on this, setting up flow initially
                cPCN2.Dept = (CStr(pcnAgg))  'Need to work on this, setting up flow initially
                'Div and Dept may not exist in these tables.  May only be in the job_rec 
                cPCN2.Position = (CStr(pcnAgg))  'Need to work on this, setting up flow initially

                'what to do if these return no matching field?   Error out and log?   Insert? Ignore referential integrity and continue?
                sPos = ValidateField(cPCN2.Position, "pos_table", "pcn_04", "string")
                sDiv = ValidateField(cPCN2.Div, "div_table", "div", "string")
                sDpt = ValidateField(cPCN2.Dept, "dept_table", "dept", "string")

                clsJR2.HRDiv = cPCN2.Div
                clsJR2.HRDept = cPCN2.Dept
                clsJR2.Descr = GetPosDescr(clsJR2.TPosNo)

                '==================================
                If ValidateField(pstnID, "hrpay_table", "hrpay", "string") = "Exists" Then
                    clsJR2.HRPay = pstnID
                Else
                    'MAIL Event
                    'sMl.SendEMessage("ADP to CX", ("Error in ParseADPData" & e.Message), "dsullivan@carthage.edu", "dsullivan@carthag.edu", "dsullivan@carthage.edu")
                    Throw New Exception("Hrpay type does not exist in CX hrpay_table.  Item = " & pstnID)
                    WriteError("HR PAY Value doe not exist in CX.  ID Number =" & id & " Employee = " & FullName & "HRPay Value = " & pstnID)

                End If

                clsJR2.BegDate = sStDate  'Format for insert?
                clsJR2.EndDate = sEndDate
                clsJR2.FullName = FullName

                'clsJR2.BJobNo = 0   'sequential number...
                clsJR2.ActiveCtrct = "N"   'Always seems to be "N" in CX -- Related to EEO
                'clsJR.CompBegDate = CDate("1999-01-01")    'Not in ADP
                'clsJR.CompEndDate = CDate("1999-01-01")    'Not in ADP
                'clsJR.AppInactDate = CDate("1999-01-01")     'Not in ADP
                clsJR2.CtrctStat = "N/A"
                clsJR2.EGPType = "R"  'Earned Gross Pay.   Always an R in CX.
                clsJR2.ExclSrvcHrs = "N"  'Always populated with N

                'Query for title from CX for secondary job title
                Dim sql As String
                sql = "SELECT distinct job_rec.job_title, job_rec.hrstat FROM job_rec Where tpos_no = ?"
                Dim cmd As New Odbc.OdbcCommand(sql, ConnInformix)
                cmd.Parameters.AddWithValue("@tpos_no", clsJR2.TPosNo)

                ConnInformix.ConnectionString = gl.ConnectString
                ConnInformix.Open()
                Dim reader As Odbc.OdbcDataReader
                reader = cmd.ExecuteReader()
                If reader.HasRows Then
                    Do While (reader.Read())
                        clsJR2.JobTitle = Trim(reader.GetString(0))
                        clsJR2.HRStat = reader.GetString(1)
                    Loop
                End If
                reader.Close()
                ConnInformix.Close()
                ConnInformix.Dispose()

                'If there are third or fourth jobs, they will need to be set to 3 or 4 
                If sEndDate = "" Then
                    clsJR2.TitleRank = jRank
                Else
                    clsJR2.TitleRank = Nothing
                End If

            Else
                'Best to write to a log file the issue.
                'WriteLog("Position " & CheckForTPos(pstnID) & " does not exist in system.   Cannot be asigned to user " & id)
                Throw New Exception("Secondary position does not exist in CX Position Table.  ID = " & id & " Employee = " & FullName & ", Position = " & pstnID & ", PCN_AGG = " & pcnAgg)
            End If

            If jRec2Exists = "Exists" Then
                'id As Integer, pcnAgg As String, sDate As Date, hrStat As String, pstnID As String, ADPTitle As String
                Debug.Print("Job Rec2 Exists")
                clsJR2.Update()
            ElseIf jRec2Exists = "New" Then
                Debug.Print("New Second Job Rec")
                clsJR2.Insert()  'Works
            Else
                WriteLog("Carthage ID Not recognized processeng second job record:  ID = " & id & " Employee = " & FullName)
            End If

        Catch ex As Exception
            WriteError("Error in Process_2nd_Job_rec: " & ex.Message & "Name = " & FullName & ", ID = " & id & " PCNAgg = " & pcnAgg)
            'MAIL Event
            sml.SendEMessage("ADP to CX", "Error in Process_2nd_Job_rec.  Error = " & ex.Message & "Name = " & FullName & ", ID = " & id & " PCNAgg = " & pcnAgg, gl.FromAddress, gl.ToAddress, "")
            Process_2nd_Job_rec = 0
        End Try
        Process_2nd_Job_rec = 1
        GC.Collect()
    End Function
    Public Function GetTitleRank(id As Integer) As String

        Dim sql As String
        Dim iRank As Integer
        Dim sRank As String = ""
        Try

            sql = "select count(title_rank) as rank from job_rec where id = ? and end_date is null"
            Dim cmd As New Odbc.OdbcCommand(sql, ConnInformix)
            cmd.Parameters.AddWithValue("@id", id)

            ConnInformix.ConnectionString = gl.ConnectString
            ConnInformix.Open()
            Dim reader As Odbc.OdbcDataReader
            reader = cmd.ExecuteReader()
            If reader.HasRows Then
                Do While (reader.Read())
                    If Not reader.IsDBNull(0) Then
                        iRank = Val(reader.GetValue(0)) + 1
                        sRank = iRank.ToString
                    End If
                Loop
            End If
            reader.Close()
            ConnInformix.Close()
            ConnInformix.Dispose()

        Catch ex As Exception
            GetTitleRank = 0
            Debug.WriteLine(ex.Message)
            WriteError("ErrorToString in GetJobRank.  ID = " & id)
        Finally
            If ConnInformix.State = ConnectionState.Open Then
                ConnInformix.Close()
                ConnInformix.Dispose()
            End If
            GC.Collect()
            GetTitleRank = sRank
        End Try
    End Function
    

    Public Function CheckAddress(ID As Integer, FullName As String, A1 As String, A2 As String, A3 As String, Cty As String, St As String, Zp As String, cntry As String) As String
        Dim sql As String = ""
        Dim sql1 As String = ""
        Dim bChange As Boolean = False
        Dim sArchive As String
        Dim adtbl1 As New DataTable
        Dim bArchive As Boolean
        Dim sCurLine1 As String = ""
        Dim sCurCtry As String = ""
        Dim sCurLine2 As String = ""
        Dim sCurLine3 As String = ""
        Dim sCurCity As String = ""
        Dim sCurState As String = ""
        Dim sCurZip As String = ""

        Try

            'Find out if we have a record that needs to be archived
            '===============================
            sql1 = "SELECT id, addr_line1, addr_line2, addr_line3, city, st, zip, ctry From id_rec Where id = ?"
            Dim cmd1 As New Odbc.OdbcCommand(sql1, ConnInformix)
            cmd1.Parameters.AddWithValue("@id", ID)

            ConnInformix.ConnectionString = gl.ConnectString
            ConnInformix.Open()
            Dim reader1 As Odbc.OdbcDataReader
            reader1 = cmd1.ExecuteReader()


            adtbl1.Load(reader1)

            If adtbl1.Rows.Count <> 0 Then

                If Not IsDBNull(adtbl1.Rows(0)(1)) Then sCurLine1 = Trim(adtbl1.Rows(0)(1))
                If Not IsDBNull(adtbl1.Rows(0)(2)) Then sCurLine2 = Trim(adtbl1.Rows(0)(2))
                If Not IsDBNull(adtbl1.Rows(0)(3)) Then sCurLine3 = Trim(adtbl1.Rows(0)(3))
                If Not IsDBNull(adtbl1.Rows(0)(4)) Then sCurCity = Trim(adtbl1.Rows(0)(4))
                If Not IsDBNull(adtbl1.Rows(0)(5)) Then sCurState = Trim(adtbl1.Rows(0)(5))
                If Not IsDBNull(adtbl1.Rows(0)(6)) Then sCurZip = Trim(adtbl1.Rows(0)(6))
                If Not IsDBNull(adtbl1.Rows(0)(7)) Then sCurCtry = Trim(adtbl1.Rows(0)(7))

                ' Only archive if there is a change in city or street.   How much of a change is just an edit?
                If Trim(sCurLine1) = Trim(A1) And Trim(sCurLine2) = Trim(A2) And Trim(sCurLine3) = Trim(A3) And Trim(sCurCity) = Trim(Cty) And Trim(sCurZip) = Trim(Zp) Then
                    bArchive = False
                    bChange = False
                Else
                    bChange = True
                    If sCurLine1 = "" Then
                        bArchive = False
                    Else
                        bArchive = True
                    End If
                End If

            End If
            reader1.Close()
            ConnInformix.Close()
            ConnInformix.Dispose()

            reader1 = Nothing

            If bArchive = True And bChange = True Then
                sArchive = ArchiveAddr(ID, FullName, sCurLine1, sCurLine2, sCurLine3, sCurCity, sCurState, sCurZip, sCurCtry)
                'sArchive = ArchiveAddr(ID, A1, A2, A3, Cty, St, Zp, cntry)
            End If


        Catch ex As Exception
            CheckAddress = bChange.ToString
            Debug.WriteLine(ex.Message)
            WriteError("ErrorToString in CheckAddress.  Name = " & FullName & ", ID = " & ID & "Address = " & Trim(A1 & " " & A2 & " " & A3) & " " & Cty & ", " & St & " " & Zp)
        Finally
            If ConnInformix.State = ConnectionState.Open Then
                ConnInformix.Close()
                ConnInformix.Dispose()
            End If
            GC.Collect()
            CheckAddress = bChange.ToString
        End Try
    End Function

    Public Function ArchiveAddr(ID As Integer, FullName As String, A1 As String, A2 As String, A3 As String, Cty As String, St As String, Zp As String, cntry As String) As String
        'Public Function ArchiveAddr(ID As Integer, A1 As String, Zp As String) As String
        Dim sql As String = ""
        Dim sAA As String = ""
        Dim iAANum As Integer
        Dim ArTbl As New DataTable
        '        Dim ArTbl2 As New DataTable
        Dim bChange As Boolean = False
        Dim sArchAA As String = ""
        Dim sArchAANum As String = ""
        Dim eMsg As New SendMessage

        'Dates are an issue to resolve
        'No Rhyme or reason in the data...
        'Need to see if there is a record of the same type (PERM) and if so add an end date
        'ID, AA and Start date should serve as a key of sorts, and if there is another with the same start date, it will throw Duplicate key error.

        'So in addition to seeing if the same address is in the database, also look for existing record with same aa type and find its start and stop dates.
        'Add End Date to previous record because we would be creating a new alternate address


        Try
            'See if there is already an Archive record with the same address...
            'If it is already there, no need to write anything - very unlikely, but could have moved out and moved back in years later.
            sql = "SELECT id, aa, aa_no, beg_date, line1, Line2, Line3, city, st, zip From aa_rec Where  id = ? "
            If A1 <> "" Then sql = sql & " and line1 = ? "
            If A2 <> "" Then sql = sql & " and line2 = ? "
            If A3 <> "" Then sql = sql & " and line3 = ? "
            If Cty <> "" Then sql = sql & " and city = ? "
            If St <> "" Then sql = sql & " and st = ? "
            If Zp <> "" Then sql = sql & " and zip = ? "
            sql = sql & " and end_date is null"
            sql = sql & " order by beg_date"

            Dim cmd As New Odbc.OdbcCommand(sql, ConnInformix)
            cmd.Parameters.AddWithValue("@id", ID)

            If A1 <> "" Then cmd.Parameters.AddWithValue("@line1", A1)
            If A2 <> "" Then cmd.Parameters.AddWithValue("@line2", A2)
            If A3 <> "" Then cmd.Parameters.AddWithValue("@line3", A3)
            If Cty <> "" Then cmd.Parameters.AddWithValue("@city", Cty)
            If St <> "" Then cmd.Parameters.AddWithValue("@st", St)
            If Zp <> "" Then cmd.Parameters.AddWithValue("@zip", Zp)

            ConnInformix.ConnectionString = gl.ConnectString
            ConnInformix.Open()
            Dim reader As Odbc.OdbcDataReader
            reader = cmd.ExecuteReader()

            'Load data into a new datatable
            ArTbl.Load(reader)
            reader.Close()
            ConnInformix.Close()
            ConnInformix.Dispose()


            sql = "SELECT max(beg_date) as D1, ID, aa, line1, end_date From aa_rec "
            sql = sql & " Where id = ? and aa = ? and end_date is null"
            sql = sql & " group by id, aa, end_date, line1"

            'sql = "SELECT max(beg_date) as D1 From aa_rec Where  id = ? and aa = ? "
            Dim cmd2 As New Odbc.OdbcCommand(sql, ConnInformix)
            cmd2.Parameters.AddWithValue("@id", ID)
            cmd2.Parameters.AddWithValue("@aa", "PREV")
            ConnInformix.ConnectionString = gl.ConnectString
            ConnInformix.Open()
            Dim reader2 As Odbc.OdbcDataReader
            Dim dMaxDate As Date
            Dim dReaderdate As Date
            Dim l1 As String = ""

            reader2 = cmd2.ExecuteReader()
            'Load data into a new datatable
            reader2.Read()


            If reader2.HasRows = False Then
                dReaderdate = Nothing
                l1 = ""
                dMaxDate = Date.Now()
            ElseIf reader2.IsDBNull(0) Then
                dMaxDate = Date.Now()
                dReaderdate = Nothing
                l1 = ""
            Else
                dReaderdate = reader2.GetDate(0)
                l1 = reader2.GetString(3)
                dMaxDate = reader2.GetDate(0)
            End If

            reader2.Close()
            ConnInformix.Close()
            ConnInformix.Dispose()

            If dMaxDate = dReaderdate Then
                'Need to put in the end date or the insert may not work...
                Dim clAA1 As New clsAA
                clAA1.Initialize()
                clAA1.ID = ID
                clAA1.Line1 = l1
                clAA1.AA = "PREV"
                clAA1.Peren = "N"
                clAA1.EndDate = Now().ToString("MM-dd-yyyy")
                clAA1.FullName = FullName
                clAA1.Update()
            End If



            If ArTbl.Rows.Count = 0 Then   'This means that the ID_Rec address will change but nothing exists in aa_rec, so we will insert as 'PREV'
                sAA = "PREV"
                'iAANum = sequence, insert will create
                Dim clAA As New clsAA
                clAA.Initialize()
                clAA.ID = ID
                clAA.AA = "PREV"
                clAA.Line1 = A1
                clAA.Line2 = A2
                clAA.Line3 = A3
                clAA.City = Cty
                clAA.State = St
                clAA.Zip = Zp
                clAA.Cntry = cntry
                clAA.CelCarrier = ""
                clAA.CassCertDate = ""
                clAA.EndDate = ""

                If dMaxDate >= Date.Today Then
                    clAA.BegDate = dMaxDate.AddDays(+1).ToString("MM-dd-yyyy")
                Else
                    clAA.BegDate = Date.Today.ToString("MM-dd-yyyy")
                End If

                clAA.OfcAddBy = "HR"
                clAA.OptOut = ""
                clAA.Peren = "N"
                clAA.Phone = ""
                clAA.PhoneExt = ""
                clAA.FullName = FullName
                clAA.Insert()
                bChange = True
                'Write into AA_rec table as PREV
            Else
                'Something was found in the aa_rec table, decide what to do with it
                'Could be more than one line
                For i As Integer = 0 To ArTbl.Rows.Count - 1
                    sAA = ArTbl.Rows(i)(1)
                    If Not IsDBNull(ArTbl.Rows(i)(2)) Then iAANum = ArTbl.Rows(i)(2)
                    Dim clAA As New clsAA
                    clAA.Initialize()
                    clAA.ID = ID
                    'clAA.AANum = iAANum  'Sequence,  insert will create
                    clAA.Line1 = ArTbl.Rows(i)(1)
                    clAA.Line1 = A1
                    clAA.Line2 = A2
                    clAA.Line3 = A3
                    clAA.City = Cty
                    clAA.State = St
                    clAA.Zip = Zp
                    clAA.Cntry = cntry
                    clAA.CelCarrier = ""
                    clAA.CassCertDate = ""
                    clAA.EndDate = Date.Today.ToString("MM-dd-yyyy")
                    clAA.BegDate = ""
                    clAA.OfcAddBy = "HR"
                    clAA.OptOut = ""
                    clAA.Peren = "N"
                    clAA.Phone = ""
                    clAA.PhoneExt = ""
                    clAA.FullName = FullName


                    If sAA = "PREV" Then
                        clAA.AA = ("PREV")
                        clAA.Update()
                        'Could put and end date here, if necessary
                    ElseIf sAA = "PERM" Then
                        clAA.ChangeType()
                        'Change 'PERM' to 'PREV' 
                        'Then add the new 'PREV' address
                        clAA.AA = ("PREV")
                        clAA.ChangeType()
                        WriteLog("Archived Address.  Full Name= " & FullName & ", ID = " & ID & "Address = " & Trim(A1 & " " & A1 & " " & A2 & " " & A3) & ", " & Cty & " " & St)
                        ' eMsg.SendEMessage("Archived Address", "Full Name = " & FullName & ", ID = " & ID & "Address = " & Trim(A1 & " " & A1 & " " & A2 & " " & A3) & ", " & Cty & " " & St, gl.FromAddress, gl.SupportAddress, "")
                    Else
                        'Do nothing for other AA types
                    End If
                    bChange = True
                Next i
                'Loop
            End If


            reader = Nothing
            reader2 = Nothing
            'ConnInformix.Close()
            'ConnInformix.Dispose()
            ArTbl = Nothing
            GC.Collect()
        Catch ex As Exception
            ArchiveAddr = bChange.ToString
            Debug.WriteLine(ex.Message)
            WriteError("ErrorToString in CheckAddress. Error = " & ex.Message & ", Name = " & FullName & ", ID = " & ID & "Address = " & Trim(A1 & " " & A1 & " " & A2 & " " & A3) & ", " & Cty & " " & St)
        Finally
            If ConnInformix.State = ConnectionState.Open Then
                ConnInformix.Close()
                ConnInformix.Dispose()
            End If
            GC.Collect()
            ArchiveAddr = bChange.ToString
        End Try
    End Function

    Public Function AA_SCHL_REC(ID As Integer, sfullname As String, sBegDate As String, sEndDate As String) As String
        'Public Function ArchiveAddr(ID As Integer, A1 As String, Zp As String) As String
        Dim sql As String = ""
        Dim sAA As String = ""
        '        Dim iAANum As Integer
        Dim ArTbl As New DataTable
        '        Dim ArTbl2 As New DataTable
        Dim sArchAA As String = ""
        Dim sArchAANum As String = ""

        Try
            'See if there is already an Alternate record with the SCHL flag...
            'If it is already there, with no end date, don't update.  Just ID and name.  
            'If it is not there, write
            'If there with end date, write
            'if Change?  End date and write new?

            sql = "SELECT id, aa, aa_no, beg_date, end_date, line1 From aa_rec Where  id = ? and aa = ? "
            'If sfullname <> "" Then sql = sql & " and line1 = ? "
            sql = sql & " and end_date is null"
            sql = sql & " order by beg_date"

            Dim cmd As New Odbc.OdbcCommand(sql, ConnInformix)
            cmd.Parameters.AddWithValue("@id", ID)
            cmd.Parameters.AddWithValue("@aa", "SCHL")

            'If sfullname <> "" Then cmd.Parameters.AddWithValue("@line1", sfullname)

            ConnInformix.ConnectionString = gl.ConnectString
            ConnInformix.Open()
            Dim reader As Odbc.OdbcDataReader
            reader = cmd.ExecuteReader()

            'Load data into a new datatable
            ArTbl.Load(reader)
            reader.Close()
            ConnInformix.Close()
            ConnInformix.Dispose()


            sAA = "SCHL"
            'iAANum = sequence, insert will create
            Dim clAA As New clsAA
            clAA.Initialize()


            clAA.ID = ID
            clAA.AA = "SCHL"
            clAA.Line1 = sfullname
            clAA.Peren = "N"   'What is this field?


            If ArTbl.Rows.Count = 0 Then   'This means that the SCHL entry does not exist, Create it 
                clAA.BegDate = sBegDate
                clAA.OfcAddBy = "HR"
                clAA.Peren = "N"   'What is this field?
                clAA.Insert()
                'Write into AA_rec table as PREV
            ElseIf ArTbl.Rows.Count >= 1 And sEndDate <> "" Then 'And ArTbl.Rows(0)(3) = sBegDate Then
                clAA.BegDate = ArTbl.Rows(0)(3)
                clAA.EndDate = sEndDate
                clAA.Update()
            End If


        Catch ex As Exception
            AA_SCHL_REC = "OK"
            Debug.WriteLine(ex.Message)
            WriteError("ErrorToString in CheckAddress. Error = " & ex.Message & "  ID = " & ID & "Name = " & Trim(sfullname))
        Finally
            If ConnInformix.State = ConnectionState.Open Then
                ConnInformix.Close()
                ConnInformix.Dispose()
            End If
            GC.Collect()
            AA_SCHL_REC = "OK"
        End Try

    End Function




    Public Function ParseCommasInQuotes(arg As String) As String


        Dim retStr As String = ""
        Dim foundEndQuote As Boolean = False
        Dim foundStartQuote As Boolean = False
        Dim iStrt As Integer = 0
        Dim iEnd As Integer = 0
        Dim ipos As Integer = 0

        '44 = comma
        '34 = double quote

        For ipos = 1 To Len(arg)
            If foundEndQuote = True Then
                foundEndQuote = False
                foundStartQuote = False
            End If

            If Mid(arg, ipos, 1) = Chr(34) And foundEndQuote = False And foundStartQuote = True Then
                foundEndQuote = True
            End If

            If Mid(arg, ipos, 1) = Chr(34) And foundStartQuote = False Then
                foundStartQuote = True
            End If

            If Mid(arg, ipos, 1) = Chr(44) And foundStartQuote = True Then
                foundStartQuote = True
            Else
                retStr = retStr & Mid(arg, ipos, 1)
            End If

        Next
        ParseCommasInQuotes = retStr

    End Function

    Public Function ConvertString(ByVal sValue As String) As String
        Dim sAns As String
        sAns = Replace(sValue, Chr(39), "''")
        sAns = "'" & sAns & "'"
        ConvertString = sAns
    End Function

    Public Function Quotes(ByVal sValue As String) As String
        Quotes = Chr(34) & sValue & Chr(34)
    End Function
    Public Function RemoveQuotes(ByVal sValue As String) As String
        Dim sRetVal As String
        Dim QuoteTest As Boolean

        sRetVal = sValue
        QuoteTest = True

        While QuoteTest = True
            If Left(sRetVal, 1) = Chr(34) Then
                sRetVal = Right(sValue, Len(sValue) - 1)
            Else
                QuoteTest = False
            End If
        End While

        QuoteTest = True
        While QuoteTest = True
            If Right(sRetVal, 1) = Chr(34) Then
                sRetVal = Left(sRetVal, Len(sRetVal) - 1)
            Else
                QuoteTest = False
            End If
        End While

        RemoveQuotes = Trim(sRetVal)
    End Function


    Public Sub WriteLog(ByVal msg As String)
        FileOpen(10, My.Application.Info.DirectoryPath & "\ADPLog.txt", OpenMode.Append)
        PrintLine(10, msg & ", " & Format(Now, "MM/dd/yyyy HH:MM:ss"))
        FileClose(10)
    End Sub


    Public Function WriteError(ByVal msg As String)
        FileOpen(10, My.Application.Info.DirectoryPath & "\ErrorLog.txt", OpenMode.Append)
        PrintLine(10, msg & ", " & Format(Now, "MM/dd/yyyy HH:MM:ss"))
        Dim sMsg As New SendMessage
        'Call sMsg.SendEMessage("ADP to CX Error", msg, gl.FromAddress, gl.ToAddress, "")
        FileClose(10)
        WriteError = 1
    End Function



    Private Sub ReadCSVFile()
        Dim fileRows(), fileFields() As String
        Dim Row1 As String


        Row1 = String.Empty
        If File.Exists(gl.FileIn) Then
            Dim fileStream As StreamReader = File.OpenText(gl.FileIn)
            fileRows = fileStream.ReadToEnd().Split(Environment.NewLine)
            For i As Integer = 0 To fileRows.Length - 1
                fileFields = fileRows(i).Split(",")
                Debug.Print(fileFields(0))
                Debug.Print(fileFields(1))
                Debug.Print(fileFields(2))
                If fileFields.Length >= 4 Then
                    Debug.Print(fileFields.Length)
                    Row1 += fileFields(0) & " " & fileFields(1) & " " & fileFields(3) & "<br /> "
                End If
            Next
        Else
            Row1 = gl.FileIn & " not found."
        End If
    End Sub


    Public Sub GetParams()
        Dim RetVal As String
        RetVal = Space(255)
        Dim v As Integer
        Dim xmlDoc As New XmlDocument()
        Dim crpt As String = ""
        Dim savecrpt As String = ""
        Dim ipos As Integer
        '        Dim sDataPath As String

        'Using normal UNC paths such as the one you mentioned works perfectly for me. For example:
        'string[] dirs = Directory.GetFiles(@"\\192.168.1.116\Shared Folder\MyDrive");
        
        gl.AppPath = My.Application.Info.DirectoryPath
        ipos = InStr(gl.AppPath, "Debug") - 1

        gl.FileNamePath = Left(gl.AppPath, ipos) & "DataFiles"

        'general parameters
        v = NativeMethods.GetPrivateProfileString("CX", "FileIn", "", RetVal, 255, My.Application.Info.DirectoryPath & "\params.ini")
        gl.FileIn = CStr(Left(RetVal, v))
        v = NativeMethods.GetPrivateProfileString("CX", "OldFile", "", RetVal, 255, My.Application.Info.DirectoryPath & "\params.ini")
        gl.OldFile = CStr(Left(RetVal, v))

        v = NativeMethods.GetPrivateProfileString("CX", "crpt", "", RetVal, 255, My.Application.Info.DirectoryPath & "\params.ini")
        crpt = CStr(Left(RetVal, v))
        v = NativeMethods.GetPrivateProfileString("CX", "savecrpt", "", RetVal, 255, My.Application.Info.DirectoryPath & "\params.ini")
        savecrpt = CStr(Left(RetVal, v))

        v = NativeMethods.GetPrivateProfileString("CX", "fromaddress", "", RetVal, 255, My.Application.Info.DirectoryPath & "\params.ini")
        gl.FromAddress = CStr(Left(RetVal, v))
        v = NativeMethods.GetPrivateProfileString("CX", "toaddress", "", RetVal, 255, My.Application.Info.DirectoryPath & "\params.ini")
        gl.ToAddress = CStr(Left(RetVal, v))

        v = NativeMethods.GetPrivateProfileString("CX", "supportaddress", "", RetVal, 255, My.Application.Info.DirectoryPath & "\params.ini")
        gl.SupportAddress = CStr(Left(RetVal, v))

        v = NativeMethods.GetPrivateProfileString("CX", "smptphost", "", RetVal, 255, My.Application.Info.DirectoryPath & "\params.ini")
        gl.SmtpMailHost = CStr(Left(RetVal, v))
        v = NativeMethods.GetPrivateProfileString("CX", "smptport", "", RetVal, 255, My.Application.Info.DirectoryPath & "\params.ini")
        gl.SmtpPort = CStr(Left(RetVal, v))

        'The connection string will be encrypted in a file within the project
        'sConnectString = "Dsn=" & sDsn & ";Driver=" & sDriver & ";Host=" & sHost & ";Server=" & sServer & ";Service=" & sService & ";Protocol=" & sProtocol & ";Database=" & sDataBase & ";Uid=" & sUid & ";Pwd=" & sPwd & ";"

        '+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
        'This calls the encryption/decryption process
        'to toggle the encryption on or off change the value "crpt" in the Params.ini file:  EN for encrypt, DE for decrypt

        ' There is a separate program for enctypting and decrypting the connection string file
        ' Run Encrypt_Decrypt.exe to decrypt the file and edit it.   
        ' Run it again to re-encrypt

        Dim cspParams As New CspParameters()            ' Create a new CspParameters object to specify a key container.
        cspParams.KeyContainerName = "XML_ENC_RSA_KEY"
        ' Create a new RSA key and save it in the container.  This key will encrypt
        ' a symmetric key, which will then be encryped in the XML document.
        Dim rsaKey As New RSACryptoServiceProvider(cspParams)
        Try
            xmlDoc.PreserveWhitespace = True
            xmlDoc.Load("connect.xml")   'In the bin/debug folder

            'DECRYPT
            clsEncryption.Decrypt(xmlDoc, rsaKey, "rsaKey")
            'If you want to save in decrypted form
            'xmlDoc.Save("connect.xml")
            'do nothing
            'End If

            'Get the connection string in decrypted form
            Dim elemList As XmlNodeList = xmlDoc.GetElementsByTagName("connectionString")
            Dim i As Integer
            For i = 0 To elemList.Count - 1
                Debug.WriteLine(elemList(i).InnerXml)
                gl.ConnectString = Trim(elemList(i).InnerXml)
            Next i

        Catch e As Exception
            WriteError("Error in GetParams " & e.Message)
            Debug.WriteLine(e.Message)
        End Try

    End Sub




    Public Function ConvertDATE(ByVal sValue As String) As String
        Dim sAns As String
        Dim Y As String
        Dim m As String
        Dim d As String

        Dim iPos As Integer = 0
        Dim iPos2 As Integer = 0

        iPos = InStr(sValue, "/")
        iPos2 = InStrRev(sValue, "/")
        m = Left(sValue, iPos - 1)
        If Len(m) < 2 Then
            m = "0" & m
        End If

        d = Mid(sValue, iPos + 1, (iPos2 - iPos) - 1)
        If Len(d) < 2 Then
            d = "0" & d
        End If

        Y = Right(sValue, Len(sValue) - iPos2)
        'If Val(Y) > 50 Then
        '    Y = "19" & Y
        'Else
        '    Y = "20" & Y
        'End If


        'sAns = Y & "-" & m & "-" & d   'fORMATTING IN THE OLD SQL METHOD IS DIFFERENT FROM THE COMMMAND ADDVALUE.  IT TAKES NORMAL DATES
        sAns = m & "-" & d & "-" & Y
        ConvertDATE = sAns
    End Function


    Public Function ConvertBDATE(ByVal sValue As String) As String
        Dim sAns As String
        Dim Y As String
        Dim m As String
        Dim d As String

        If Len(sValue) < 6 Then
            sValue = "0" & sValue
        End If

        Y = Right(sValue, 2)
        If Val(Y) > 50 Then
            Y = "19" & Y
        Else
            Y = "20" & Y
        End If
        m = Left(sValue, 2)
        d = Mid(sValue, 3, 2)


        sAns = Y & "-" & m & "-" & d
        ConvertBDATE = ConvertString(sAns)
    End Function

    Public Function RemoveLF(ByVal sStr As String) As String
        Dim ipos As Short
        Dim retString As String
        retString = ""
        For ipos = 1 To Len(sStr)
            '        Debug.Print Mid(sStr, iPos, 1)
            '        Debug.Print Asc(Mid(sStr, iPos, 1))
            If Mid(sStr, ipos, 1) <> Chr(10) Then
                retString = retString & Mid(sStr, ipos, 1)
            End If

        Next
        RemoveLF = retString

        ' Debug.Print retString + " - " + sStr
    End Function



    Public Function LPad(ByVal str_Renamed As String, ByVal Ln As Short) As String
        Dim strLen As Short
        Dim pad As Short

        strLen = Len(str_Renamed)
        pad = Ln - strLen


        LPad = New String(" ", pad) & str_Renamed
    End Function


    Private Function DateFormat(p1 As Object, p2 As String) As Date
        Throw New NotImplementedException
    End Function


    Private Function CheckForTPos(sPCNAggr As String) As Integer
        'This reterns the tpos_no field regardless of who it might be attached to
        Dim sql As String
        Dim iTPos As Integer


        iTPos = 0
        Try
            sql = "SELECT tpos_no From pos_table Where pcn_aggr = ?"

            Dim cmd As New Odbc.OdbcCommand(sql, ConnInformix)
            cmd.Parameters.AddWithValue("@pcn_aggr", sPCNAggr)


            ConnInformix.ConnectionString = gl.ConnectString
            ConnInformix.Open()
            Dim reader As Odbc.OdbcDataReader
            reader = cmd.ExecuteReader()
            If reader.HasRows Then
                Do While (reader.Read())
                    iTPos = reader.GetDouble(0)
                Loop
            End If
            reader.Close()
            ConnInformix.Close()
            ConnInformix.Dispose()

        Catch ex As Exception
            CheckForTPos = 0
            WriteError("Error in CheckForTpos:  Error = " & ex.Message)
        Finally
            If ConnInformix.State = ConnectionState.Open Then
                ConnInformix.Close()
                ConnInformix.Dispose()
            End If
            GC.Collect()
            CheckForTPos = iTPos
        End Try
    End Function

    Private Function GetPosDescr(TPos As Integer) As String
        'This reterns the tpos_no field regardless of who it might be attached to
        Dim sql As String
        Dim sPos As String = ""

        Try
            sql = "SELECT descr From pos_table Where tpos_no = ?"

            Dim cmd As New Odbc.OdbcCommand(sql, ConnInformix)
            cmd.Parameters.AddWithValue("@tpos_no", TPos)


            ConnInformix.ConnectionString = gl.ConnectString
            ConnInformix.Open()
            Dim reader As Odbc.OdbcDataReader
            reader = cmd.ExecuteReader()
            If reader.HasRows Then
                Do While (reader.Read())
                    sPos = reader.GetString(0)
                Loop
            End If
            reader.Close()
            ConnInformix.Close()
            ConnInformix.Dispose()

        Catch ex As Exception
            GetPosDescr = ""
            WriteError("Error in GetPosDescr.  Error = " & ex.Message)
        Finally
            If ConnInformix.State = ConnectionState.Open Then
                ConnInformix.Close()
                ConnInformix.Dispose()
            End If
            GC.Collect()
            GetPosDescr = sPos
        End Try
    End Function


    Private Function CheckForJobRec(iPosNo As Integer, id As Integer, stDate As String) As String
        Dim sql As String
        Dim sTest As String

        sTest = ""
        gl.JobCollection.Clear()

        Try
            sql = "SELECT job_no, tpos_no, id From job_rec Where tpos_no = ? and ID = ? and beg_date = ? "
            '? & iPosNo & " And id = " & id & " and end_date is null and hrpay = ? and end_date is null "

            Dim cmd As New Odbc.OdbcCommand(sql, ConnInformix)
            cmd.Parameters.AddWithValue("@tpos_no", iPosNo)
            cmd.Parameters.AddWithValue("@id", id)
            'cmd.Parameters.AddWithValue("@hrpay", hrPay)
            cmd.Parameters.AddWithValue("@beg_date", stDate)

            ConnInformix.ConnectionString = gl.ConnectString
            ConnInformix.Open()
            Dim reader As Odbc.OdbcDataReader

            reader = cmd.ExecuteReader()

            If reader.HasRows Then
                reader.Read()
                'Do While (reader.Read())
                If id = reader.GetValue(2) Then  'ID already exists
                    sTest = "Exists"
                    gl.JobCollection.Add((reader.GetValue(0)), 1)  'Job_n0
                    gl.JobCollection.Add((reader.GetValue(1)), 2)  'tpos_no
                    gl.JobCollection.Add((reader.GetValue(2)), 3)  'id
                Else
                    sTest = "New"
                    gl.JobCollection.Add(0, 1)
                    gl.JobCollection.Add(iPosNo, 2)
                    gl.JobCollection.Add(id, 3)
                End If
                'Loop
            Else
                sTest = "New"
                gl.JobCollection.Add(0, 1)
                gl.JobCollection.Add(iPosNo, 2)
                gl.JobCollection.Add(id, 3)
            End If
            reader.Close()
            ConnInformix.Dispose()

            ConnInformix.Close()
        Catch ex As Exception
            CheckForJobRec = ""
            WriteError("Error in CheckforJobRec.  Error = " & ex.Message)
        Finally
            If ConnInformix.State = ConnectionState.Open Then
                ConnInformix.Close()
                ConnInformix.Dispose()
            End If
            GC.Collect()
            CheckForJobRec = sTest
        End Try
    End Function


    Private Function ValidateField(SearchVal As String, table As String, keyField As String, keyType As String) As String
        Dim sql As String
        Dim sTestVal As String

        sTestVal = ""
        Try

            'This type of query method implicitly handles data types - No need for separate SQL for string key vs numeric key
            sql = "select distinct " & keyField & " from " & table & " Where " & keyField & "=?"
            Dim cmd As New Odbc.OdbcCommand(sql, ConnInformix)
            cmd.Parameters.AddWithValue(keyField, SearchVal)

            ConnInformix.ConnectionString = gl.ConnectString
            ConnInformix.Open()

            Dim reader As Odbc.OdbcDataReader
            reader = cmd.ExecuteReader()

            If reader.HasRows Then
                Do While (reader.Read())
                    If keyType = "string" Then
                        If Trim(SearchVal) = Trim(reader.GetString(0)) Then  'ID already exists
                            sTestVal = "Exists"
                        Else
                            sTestVal = "New"
                        End If
                    Else
                        If SearchVal = reader.GetDouble(0) Then  'ID already exists
                            sTestVal = "Exists"
                        Else
                            sTestVal = "New"
                        End If
                    End If

                Loop
            Else
                sTestVal = "New"  'ID Does not exist
                'Do a Separate check for duplicate name and BDate
            End If
            reader.Close()
            ConnInformix.Close()
            ConnInformix.Dispose()

        Catch ex As Exception
            ValidateField = ""
            WriteError("Error in Validate Field.  Error = " & ex.Message & "Table = " & table & "Field to query = " & keyField & "Search Value = " & SearchVal)
        Finally
            If ConnInformix.State = ConnectionState.Open Then
                ConnInformix.Close()
                ConnInformix.Dispose()
            End If
            ValidateField = sTestVal

        End Try
    End Function
    Private Function CheckDB(id As Integer) As Integer
        Dim sql As String
        Dim stestID As Integer

        Try
            sql = "SELECT id_rec.id, id_rec.lastname, id_rec.firstname FROM id_rec where id_rec.id = ?"
            Dim cmd As New Odbc.OdbcCommand(sql, ConnInformix)
            cmd.Parameters.AddWithValue("id_rec.id", id)

            'may want to us partial first, last and gender due to possible spelling variations
            ConnInformix.ConnectionString = gl.ConnectString
            ConnInformix.Open()
            Dim reader As Odbc.OdbcDataReader
            reader = cmd.ExecuteReader()
            If reader.HasRows Then
                Do While (reader.Read())
                    'If Fname = reader.GetString(0) And lname = reader.GetString(1) And gender = reader.GetString(2) Then  'ID already exists
                    If CInt(reader.GetString(0)) > 0 Then
                        stestID = reader.GetValue(0)
                    Else
                        stestID = 0
                    End If
                Loop
            Else
                stestID = 0  'ID Does not exist
                'Do a Separate check for duplicate name and BDate
            End If
            reader.Close()
            ConnInformix.Close()
            ConnInformix.Dispose()
        Catch ex As Exception
            CheckDB = stestID
            WriteError("Error in CheckForName.  Error = " & ex.Message)
        Finally
            If ConnInformix.State = ConnectionState.Open Then
                ConnInformix.Close()
                ConnInformix.Dispose()
            End If
            GC.Collect()
            CheckDB = stestID
        End Try
    End Function


    Private Function CheckForName(lname As String, fname As String, gender As String, bd As String) As Integer
        Dim sql As String
        Dim sTestID As Integer

        sTestID = 0
        Debug.WriteLine(Convert.ToDateTime(bd))
        Try
            sql = "SELECT id_rec.id, id_rec.lastname, id_rec.firstname, profile_rec.sex, profile_rec.birth_date FROM id_rec, profile_rec where id_rec.id = profile_rec.id"
            sql = sql & " and id_rec.firstname = ? And id_rec.lastname = ? "
            If gender <> "" Then sql = sql & " and profile_rec.sex = ?"
            If bd <> "" Then sql = sql & " and profile_rec.birth_date = ?"

            Dim cmd As New Odbc.OdbcCommand(sql, ConnInformix)
            cmd.Parameters.AddWithValue("id_rec.firstname", fname)
            cmd.Parameters.AddWithValue("id_rec.lastname", lname)
            If gender <> "" Then cmd.Parameters.AddWithValue("profile_rec.sex", Left(gender, 1))
            If bd <> "" Then cmd.Parameters.AddWithValue("profile_rec.birth_date", bd)

            'may want to us partial first, last and gender due to possible spelling variations
            ConnInformix.ConnectionString = gl.ConnectString
            ConnInformix.Open()
            Dim reader As Odbc.OdbcDataReader
            reader = cmd.ExecuteReader()
            If reader.HasRows Then
                Do While (reader.Read())
                    'If Fname = reader.GetString(0) And lname = reader.GetString(1) And gender = reader.GetString(2) Then  'ID already exists
                    If CInt(reader.GetString(0)) > 0 Then
                        sTestID = CInt(reader.GetString(0))
                    Else
                        sTestID = 0
                    End If
                Loop
            Else
                sTestID = 0  'ID Does not exist
                'Do a Separate check for duplicate name and BDate
            End If
            reader.Close()

            ConnInformix.Close()
            ConnInformix.Dispose()
        Catch ex As Exception
            CheckForName = sTestID
            WriteError("Error in CheckForName.  Error = " & ex.Message & " " & lname & "," & fname & " " & bd)
        Finally
            If ConnInformix.State = ConnectionState.Open Then
                ConnInformix.Close()
                ConnInformix.Dispose()
            End If
            GC.Collect()
            CheckForName = sTestID
        End Try
    End Function


    Public Function SetEmail2(Email As String, id As Integer) As String
        Dim sql As String
        Dim sTestVal As String
        Dim cLaa As New clsAA
        Dim DT As DateTime
        Dim sBegDate As String = ""
        sTestVal = ""


        Try
            sql = "SELECT aa_rec.aa, aa_rec.id, aa_rec.line1, aa_rec.aa_no, aa_rec.beg_date FROM aa_rec  "
            sql = sql & "where aa_rec.id = ? and aa_rec.aa = 'EML2' and aa_rec.end_date is null"

            Dim cmd As New Odbc.OdbcCommand(sql, ConnInformix)
            'cmd.Parameters.AddWithValue("aa_rec.aa", "EML2")
            cmd.Parameters.AddWithValue("aa_rec.id", id)


            'may want to us partial first, last and gender due to possible spelling variations
            ConnInformix.ConnectionString = gl.ConnectString
            ConnInformix.Open()
            Dim reader As Odbc.OdbcDataReader
            reader = cmd.ExecuteReader()
            If reader.HasRows Then
                Do While (reader.Read())
                    If (reader.GetString(2)) > "" Then
                        sTestVal = CStr(reader.GetString(2))
                        DT = Convert.ToDateTime(reader.GetString(4))
                        sBegDate = dt.ToString("MM-dd-yyyy")
                    Else
                        sTestVal = ""
                        sBegDate = Date.Today.ToString("MM-dd-yyyy")
                    End If
                Loop
            Else
                sTestVal = "New"  'eMAIL2 Does not exist
                sBegDate = Date.Today.ToString("MM-dd-yyyy")
                'Do a Separate check for duplicate name and BDate
            End If
            reader.Close()
            ConnInformix.Close()
            ConnInformix.Dispose()


            If sTestVal = Email Or sTestVal = "" Or Email = "" Then
                'nothing to do
            ElseIf sTestVal = "New" Then
                cLaa.Initialize()
                cLaa.ID = id
                cLaa.Line1 = Email
                cLaa.AA = "EML2"
                cLaa.BegDate = sBegDate
                cLaa.AddEmail2()
                'insert
            ElseIf sTestVal <> "New" And sTestVal <> Email Then
                cLaa.Initialize()
                cLaa.ID = id
                cLaa.Line1 = Email
                cLaa.AA = "EML2"
                cLaa.BegDate = sBegDate
                cLaa.UpdateEmail2()
                'Update
            End If



        Catch ex As Exception
            SetEmail2 = sTestVal
            WriteError("Error in SetEmail2.  Error = " & ex.Message & " " & id & "," & Email & " EML2")
        Finally
            If ConnInformix.State = ConnectionState.Open Then
                ConnInformix.Close()
                ConnInformix.Dispose()
            End If
            GC.Collect()
            SetEmail2 = sTestVal
        End Try
    End Function


    Public Function Set_CellPhone(Phone As String, id As Integer, fullname As String) As String
        Dim sql As String
        Dim sTestVal As String = ""
        Dim cLaa As New clsAA
        Dim sBegDate As String = ""
        Dim iAANum = 0
        Dim DT As DateTime

        Try
            sql = "SELECT aa_rec.aa, aa_rec.id, aa_rec.phone, aa_rec.aa_no, aa_rec.beg_date FROM aa_rec  "
            sql = sql & "where aa_rec.id = ? and aa_rec.aa = 'CELL' and aa_rec.end_date is null"

            Dim cmd As New Odbc.OdbcCommand(sql, ConnInformix)
            '            cmd.Parameters.AddWithValue("aa_rec.aa", aa)
            cmd.Parameters.AddWithValue("aa_rec.id", id)


            'may want to us partial first, last and gender due to possible spelling variations
            ConnInformix.ConnectionString = gl.ConnectString
            ConnInformix.Open()
            Dim reader As Odbc.OdbcDataReader
            reader = cmd.ExecuteReader()
            If reader.HasRows Then
                Do While (reader.Read())
                    If (reader.GetString(2)) > "" Then
                        iAANum = CInt(reader.GetInt32(3))
                        sTestVal = CStr(reader.GetString(2))
                        DT = Convert.ToDateTime(reader.GetString(4))
                        sBegDate = dt.ToString("MM-dd-yyyy")
                    Else
                        'iAANum is a serial, no insert
                        sTestVal = ""
                        sBegDate = Date.Today.ToString("MM-dd-yyyy")
                    End If
                Loop
            Else
                sTestVal = "New"  'aa Does not exist
                sBegDate = Date.Today.ToString("MM-dd-yyyy")
                'Do a Separate check for duplicate name and BDate
            End If
            reader.Close()
            ConnInformix.Close()
            ConnInformix.Dispose()


            If sTestVal = Phone Or sTestVal = "" Or Phone = "" Then
                'nothing to do
            ElseIf sTestVal = "New" Then
                cLaa.Initialize()
                cLaa.ID = id
                cLaa.Phone = Phone
                cLaa.FullName = fullname  'Only for error reporting
                cLaa.AA = "CELL"
                cLaa.BegDate = sBegDate
                cLaa.AddCell()
                'insert
            ElseIf sTestVal <> "New" And sTestVal <> Phone Then
                cLaa.Initialize()
                cLaa.ID = id
                cLaa.FullName = fullname  'Only for error reporting
                cLaa.Phone = Phone
                cLaa.AA = "CELL"
                cLaa.AANum = iAANum
                cLaa.BegDate = sBegDate
                cLaa.UpdateCell()
                'Update
            End If



        Catch ex As Exception
            Set_CellPhone = sTestVal
            WriteError("Error in Set_CellPhone.  Error = " & ex.Message & " " & id & "," & Phone & " EML2")
        Finally
            If ConnInformix.State = ConnectionState.Open Then
                ConnInformix.Close()
                ConnInformix.Dispose()
            End If
            GC.Collect()
            Set_CellPhone = sTestVal
        End Try
    End Function


    Public Function GetRaceCode(race As Integer) As String
        Select Case race
            Case 1
                GetRaceCode = "WH"
            Case 2  'Black or African American
                GetRaceCode = "BL"
            Case 4 'Asian
                GetRaceCode = "AS"
            Case 5 'Native American or Alaska native
                GetRaceCode = "AM"
            Case 6 '"Native Hawaiian or Other Pacific Islander"
                GetRaceCode = "AP"
            Case 9  '"Two or more races"
                GetRaceCode = "MU"
            Case Else
                GetRaceCode = ""
        End Select
    End Function

    'Public Function GetFacStaffFlag(JobFuncCode As String) As String
    '    Select Case JobFuncCode
    '        Case Is = "AD"  'Exempt Admin FT
    '            GetFacStaffFlag = "Staff"
    '        Case Is = "ADPT"  'Exempt Admin PT
    '            GetFacStaffFlag = "Staff"
    '        Case Is = "FT" 'Full Time faculty
    '            GetFacStaffFlag = "Faculty"
    '        Case Is = "GA" 'Graduate Assistant""
    '            GetFacStaffFlag = "Staff"
    '        Case Is = "HR" 'Hourly Staff
    '            GetFacStaffFlag = "Staff"
    '        Case Is = "HRPT" 'Hourly Staff Part Time
    '            GetFacStaffFlag = "Staff"
    '        Case Is = "PATH"  'Part Time Athletics
    '            GetFacStaffFlag = "Staff"
    '        Case Is = "PT"  'Adjunct Faculty Part Time 
    '            GetFacStaffFlag = "Faculty"
    '        Case Is = "PTGP""  'Adj Faculty GP PT"
    '            GetFacStaffFlag = "Faculty"
    '        Case Is = "PTGP"  'Part Time faculty GPS
    '            GetFacStaffFlag = "Faculty"
    '        Case Is = "TLE"  'Target Language Exp
    '            GetFacStaffFlag = "Staff"
    '        Case Else
    '            GetFacStaffFlag = ""
    '    End Select

    'End Function
    Public Function GetFacStaffFlag(PCN_AGG As String) As String
        Dim sFSFlag As String
        sFSFlag = Left(PCN_AGG, 3)
        Select Case sFSFlag
            Case Is = "ADM"  'Admin 
                GetFacStaffFlag = "Staff"
            Case Is = "EXT"  'Faculty
                GetFacStaffFlag = "Faculty"
            Case Is = "FAC" 'faculty
                GetFacStaffFlag = "Faculty"
            Case Is = "HRL" 'Hourly?
                GetFacStaffFlag = "Staff"
            Case Is = "PAT" 'Athletics?
                GetFacStaffFlag = "Staff"
            Case Is = "PDF"
                GetFacStaffFlag = "Faculty"
            Case Else
                GetFacStaffFlag = ""
        End Select

    End Function


    Public Sub ConnectInformix()

        Try
            'IBM .NET provider
            ConnInformix.ConnectionString = gl.ConnectString
            ConnInformix.Open()
            ConnInformix.Close()
            ConnInformix.Dispose()

            '+++++++++++++++++++++++++++++++++++++++++++++++++++++
            Exit Sub

        Catch e As Exception
            Debug.WriteLine(e)
            WriteError("Error in ConnectInformix..  Error = " & e.Message)
        End Try
    End Sub

    Private Sub ReleaseObject(ByVal obj As Object)
        Try
            System.Runtime.InteropServices.Marshal.ReleaseComObject(obj)
            obj = Nothing
        Catch ex As Exception
            obj = Nothing
        Finally
            GC.Collect()
        End Try
    End Sub

End Module