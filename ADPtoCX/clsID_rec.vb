Option Strict Off
Option Explicit On

Imports System.Data.Odbc

Friend Class ClsID_rec


    Dim iID As Integer 'Primary Key, sequence added by Informix??
    Dim iPrspNo As Integer    'Not Null
    Dim sFullName As String 'Not Null
    Dim sNameSndx As String 'Not Null
    Dim sLName As String 'Not Null
    Dim sFName As String 'Not Null
    Dim sMidName As String 'Not Null
    Dim sSuffixName As String 'Not Null
    Dim sAddr1 As String
    Dim sAddr2 As String
    Dim sAddr3 As String
    Dim sCity As String
    Dim sState As String
    Dim sZip As String 'Not Null
    Dim sCtry As String
    Dim sAA As String
    Dim sTitle As String
    Dim sSuffix As String
    Dim sSSNum As String   'Not Null
    Dim sPhone As String
    Dim sPhoneExt As String
    Dim iPrevNameID As Integer
    Dim sMail As String
    Dim sSol As String
    Dim sPub As String
    Dim sCorrectAddr As String
    Dim sDecsd As String
    Dim dAddDate As String
    Dim sOffcAddBy As String
    Dim dUpdDte As String
    Dim sValid As String
    Dim dPurgeDate As String
    Dim dCassCertDate As String
    Dim sCCUserName As String
    Dim sCCPwd As String
    Dim iInquiryNum As Integer



    Public Property ID() As Integer
        Get
            ID = iID
        End Get
        Set(ByVal Value As Integer)
            iID = Value
        End Set
    End Property

    Public Property PrspNo() As Integer
        Get
            PrspNo = iPrspNo
        End Get
        Set(ByVal Value As Integer)
            iPrspNo = Value
        End Set
    End Property

    Public Property FullName() As String
        Get
            FullName = sFullName
        End Get
        Set(ByVal Value As String)
            sFullName = Value
        End Set
    End Property
    Public Property NameSndx() As String
        Get
            NameSndx = sNameSndx
        End Get
        Set(ByVal Value As String)
            sNameSndx = Value
        End Set
    End Property

    Public Property FName() As String
        Get
            FName = sFName
        End Get
        Set(ByVal Value As String)
            sFName = Value
        End Set
    End Property
    Public Property LName() As String
        Get
            LName = sLName
        End Get
        Set(ByVal Value As String)
            sLName = Value
        End Set
    End Property

    Public Property MidName() As String
        Get
            MidName = sMidName
        End Get
        Set(ByVal Value As String)
            sMidName = Value
        End Set
    End Property

    Public Property SuffixName() As String
        Get
            SuffixName = sSuffixName
        End Get
        Set(ByVal Value As String)
            sSuffixName = Value
        End Set
    End Property

    Public Property Addr1() As String
        Get
            Addr1 = sAddr1
        End Get
        Set(ByVal Value As String)
            sAddr1 = Value
        End Set
    End Property
    Public Property Addr2() As String
        Get
            Addr2 = sAddr2
        End Get
        Set(ByVal Value As String)
            sAddr2 = Value
        End Set
    End Property
    Public Property Addr3() As String
        Get
            Addr3 = sAddr3
        End Get
        Set(ByVal Value As String)
            sAddr3 = Value
        End Set
    End Property

    Public Property City() As String
        Get
            City = sCity
        End Get
        Set(ByVal Value As String)
            sCity = Value
        End Set
    End Property

    Public Property State() As String
        Get
            State = sState
        End Get
        Set(ByVal Value As String)
            sState = Value
        End Set
    End Property

    Public Property Zip() As String
        Get
            Zip = sZip
        End Get
        Set(ByVal Value As String)
            sZip = Value
        End Set
    End Property

    Public Property Ctry() As String
        Get
            Ctry = sCtry
        End Get
        Set(ByVal Value As String)
            sCtry = Value
        End Set
    End Property

    Public Property AA() As String
        Get
            AA = sAA
        End Get
        Set(ByVal Value As String)
            sAA = Value
        End Set
    End Property

    Public Property Title() As String
        Get
            Title = sTitle
        End Get
        Set(ByVal Value As String)
            sTitle = Value
        End Set
    End Property



    Public Property SSNum() As String
        Get
            SSNum = sSSNum
        End Get
        Set(ByVal Value As String)
            sSSNum = Value
        End Set
    End Property
    Public Property Phone() As String
        Get
            Phone = sPhone
        End Get
        Set(ByVal Value As String)
            sPhone = Value
        End Set
    End Property
    Public Property PhoneExt() As String
        Get
            PhoneExt = sPhoneExt
        End Get
        Set(ByVal Value As String)
            sPhoneExt = Value
        End Set
    End Property
    Public Property Suffix() As String
        Get
            Suffix = sSuffix
        End Get
        Set(ByVal Value As String)
            sSuffix = Value
        End Set
    End Property

    Public Property PrevNameID() As Integer
        Get
            PrevNameID = iPrevNameID
        End Get
        Set(ByVal Value As Integer)
            iPrevNameID = Value
        End Set
    End Property

    'Public Property Mail() As String
    '    Get
    '        Mail = sMail
    '    End Get
    '    Set(ByVal Value As String)
    '        sMail = Value
    '    End Set
    'End Property

    Public Property Mail() As String
        Get
            Mail = sMail
        End Get
        Set(ByVal Value As String)
            sMail = Value
        End Set
    End Property

    Public Property Sol() As String
        Get
            Sol = sSol
        End Get
        Set(ByVal Value As String)
            sSol = Value
        End Set
    End Property

    Public Property Pub() As String
        Get
            Pub = sPub
        End Get
        Set(ByVal Value As String)
            sPub = Value
        End Set
    End Property

    Public Property CorrectAddr() As String
        Get
            CorrectAddr = sCorrectAddr
        End Get
        Set(ByVal Value As String)
            sCorrectAddr = Value
        End Set
    End Property

    Public Property Decsd() As String
        Get
            Decsd = sDecsd
        End Get
        Set(ByVal Value As String)
            sDecsd = Value
        End Set
    End Property
    Public Property AddDate() As String
        Get
            AddDate = dAddDate
        End Get
        Set(ByVal Value As String)
            dAddDate = Value
        End Set
    End Property

    Public Property OffcAddBy() As String
        Get
            OffcAddBy = sOffcAddBy
        End Get
        Set(ByVal Value As String)
            sOffcAddBy = Value
        End Set
    End Property
    Public Property UpDte() As String
        Get
            UpDte = dUpdDte
        End Get
        Set(ByVal Value As String)
            dUpdDte = Value
        End Set
    End Property

    Public Property Valid() As String
        Get
            Valid = sValid
        End Get
        Set(ByVal Value As String)
            sValid = Value
        End Set
    End Property
    Public Property PurgeDate() As String
        Get
            PurgeDate = dPurgeDate
        End Get
        Set(ByVal Value As String)
            dPurgeDate = Value
        End Set
    End Property
    Public Property CassCertDate() As String
        Get
            CassCertDate = dCassCertDate
        End Get
        Set(ByVal Value As String)
            dCassCertDate = Value
        End Set
    End Property

    Public Property CCUserName() As String
        Get
            CCUserName = sCCUserName
        End Get
        Set(ByVal Value As String)
            sCCUserName = Value
        End Set
    End Property
    Public Property CCPwd() As String
        Get
            CCPwd = sCCPwd
        End Get
        Set(ByVal Value As String)
            sCCPwd = Value
        End Set
    End Property
    Public Property InquiryNum() As Integer
        Get
            InquiryNum = iInquiryNum
        End Get
        Set(ByVal Value As Integer)
            iInquiryNum = Value
        End Set
    End Property

    Public Sub Insert_Basic()

        Dim sql As String
        sql = " INSERT INTO id_rec (fullname, lastname, firstname, middlename, prev_name_id, upd_date, inquiry_no "
        sql = sql & " ss_no, Decsd, upd_date, ofc_add_by"
        sql = sql & ") values (?,?,?,?,?,?,?,?,?,?,?)"

        Dim cmd As New OdbcCommand(sql, ConnInformix)
        'upd_date, id

        Try

            ConnInformix.ConnectionString = gl.ConnectString
            ConnInformix.Open()

            cmd.CommandType = CommandType.Text
            '            cmd.Parameters.AddWithValue("@id", Me.ID)
            cmd.Parameters.AddWithValue("@fullname", Me.FullName)
            cmd.Parameters.AddWithValue("@lastname", Me.LName)
            cmd.Parameters.AddWithValue("@firstname", Me.FName)
            cmd.Parameters.AddWithValue("@middlename", Me.MidName)
            cmd.Parameters.AddWithValue("@prev_name_id", Me.PrevNameID)
            cmd.Parameters.AddWithValue("@ss_no", Me.SSNum)
            cmd.Parameters.AddWithValue("@decsd", Me.Decsd)
            cmd.Parameters.AddWithValue("@upd_date", Me.UpDte)
            cmd.Parameters.AddWithValue("@ofc_add_by", Me.OffcAddBy)
            cmd.Parameters.AddWithValue("@inquiry_no", Me.InquiryNum)

            '  Dim rowsAffected As Integer = cmd.ExecuteNonQuery()   'Executes query and returns a value
            '  debug.writeline("RowsAffected: {0}", rowsAffected.ToString)
            WriteLog("Update of ID record for " & Me.FName & " " & Me.LName & ", ID = " & Me.ID)

            ConnInformix.Close()
            ConnInformix.Dispose()
        Catch e As Exception
            debug.writeline(e)
            WriteError("Error in clsid_rec.Insert_Basic - Insert.  Error = " & e.Message)
            'MAIL Event
            'sMl.SendEMessage("ADP to CX", (""Error in clsid_rec - Insert_Basic.  Error = " & e.Message), "dsullivan@carthage.edu", "dsullivan@carthag.edu", "dsullivan@carthage.edu")
        End Try
        If ConnInformix.State = ConnectionState.Open Then
            ConnInformix.Close()
            ConnInformix.Dispose()
        End If
    End Sub
    Public Sub Insert_Addr()

        Dim sql As String
        sql = " INSERT INTO id_rec (addr_line1, addr_line2, addr_line3, city, st,  zip, Ctry, AA, correct_addr, upd_date "
        sql = sql & ") values (?,?,?,?,?,?,?,?,?)"

        Dim cmd As New OdbcCommand(sql, ConnInformix)
        'upd_date, id

        Try

            ConnInformix.ConnectionString = gl.ConnectString
            ConnInformix.Open()

            cmd.CommandType = CommandType.Text
            '            cmd.Parameters.AddWithValue("@id", Me.ID)   'ID is a sequence, can not specify
            cmd.Parameters.AddWithValue("@addr_line1", Me.Addr1)
            cmd.Parameters.AddWithValue("@addr_line2", Me.Addr2)
            cmd.Parameters.AddWithValue("@addr_line3", Me.Addr3)
            cmd.Parameters.AddWithValue("@city", Me.City)
            cmd.Parameters.AddWithValue("@st", Me.State)
            cmd.Parameters.AddWithValue("@zip", Me.Zip)
            cmd.Parameters.AddWithValue("@ctry", Me.Ctry)
            cmd.Parameters.AddWithValue("@aa", Me.AA)
            cmd.Parameters.AddWithValue("@correct_addr", Me.CorrectAddr)
            cmd.Parameters.AddWithValue("@upd_date", Me.UpDte)

            ' Dim rowsAffected As Integer = cmd.ExecuteNonQuery()   'Executes query and returns a value
            ' debug.writeline("RowsAffected: {0}", rowsAffected.ToString)
            WriteLog("Insert of address into ID record for " & Me.FName & " " & Me.LName & ", ID = " & Me.ID)

            ConnInformix.Close()
            ConnInformix.Dispose()
        Catch e As Exception
            debug.writeline(e)
            WriteError("Error in clsid_rec.Insert_Addr - Insert.  Error = " & e.Message)
            'MAIL Event
            'sMl.SendEMessage("ADP to CX", (""Error in clsid_rec - Insert_Addr.  Error = " & e.Message), "dsullivan@carthage.edu", "dsullivan@carthag.edu", "dsullivan@carthage.edu")
        End Try
        If ConnInformix.State = ConnectionState.Open Then
            ConnInformix.Close()
            ConnInformix.Dispose()
        End If
    End Sub



    Public Sub Insert()

        Dim sql As String
        sql = " INSERT INTO id_rec (fullname, lastname, firstname, middlename,  "
        sql = sql & "  addr_line1, addr_line2, addr_line3, city, st,  zip, Ctry, AA, ss_no, Decsd, "
        sql = sql & " upd_date, ofc_add_by, correct_addr, prev_name_id, inquiry_no"
        sql = sql & ") values (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)"

        Dim cmd As New OdbcCommand(sql, ConnInformix)
        'upd_date, id

        Try

            ConnInformix.ConnectionString = gl.ConnectString
            ConnInformix.Open()

            cmd.CommandType = CommandType.Text
            '            cmd.Parameters.AddWithValue("@id", Me.ID)
            cmd.Parameters.AddWithValue("@fullname", Me.FullName)
            cmd.Parameters.AddWithValue("@lastname", Me.LName)
            cmd.Parameters.AddWithValue("@firstname", Me.FName)

            cmd.Parameters.AddWithValue("@middlename", Me.MidName)
            cmd.Parameters.AddWithValue("@addr_line1", Me.Addr1)
            cmd.Parameters.AddWithValue("@addr_line2", Me.Addr2)
            cmd.Parameters.AddWithValue("@addr_line3", Me.Addr3)
            cmd.Parameters.AddWithValue("@city", Me.City)
            cmd.Parameters.AddWithValue("@st", Me.State)
            cmd.Parameters.AddWithValue("@zip", Me.Zip)
            cmd.Parameters.AddWithValue("@ctry", Me.Ctry)
            cmd.Parameters.AddWithValue("@aa", Me.AA)
            cmd.Parameters.AddWithValue("@ss_no", Me.SSNum)
            cmd.Parameters.AddWithValue("@decsd", Me.Decsd)
            cmd.Parameters.AddWithValue("@upd_date", Me.UpDte)
            cmd.Parameters.AddWithValue("@ofc_add_by", Me.OffcAddBy)
            cmd.Parameters.AddWithValue("@inquiry_no", Me.InquiryNum)
            cmd.Parameters.AddWithValue("@correct_addr", Me.CorrectAddr)

            'Dim rowsAffected As Integer = cmd.ExecuteNonQuery()   'Executes query and returns a value
            'debug.writeline("RowsAffected: {0}", rowsAffected.ToString)
            WriteLog("Created new ID record for " & Me.FName & " " & Me.LName & ", ID = " & Me.ID)

            ConnInformix.Close()
            ConnInformix.Dispose()
        Catch e As Exception
            debug.writeline(e)
            WriteError("Error in clsid_rec - Insert.  Error = " & e.Message)
            'MAIL Event
            'sMl.SendEMessage("ADP to CX", (""Error in clsid_rec - Insert.  Error = " & e.Message), "dsullivan@carthage.edu", "dsullivan@carthag.edu", "dsullivan@carthage.edu")
        Finally
            If ConnInformix.State = ConnectionState.Open Then
                ConnInformix.Close()
                ConnInformix.Dispose()
            End If
            GC.Collect()
        End Try
    End Sub

    Public Sub Update()

        Dim sql As String
        sql = "Update id_rec set "
        If Me.FullName <> "" Then sql = sql & " fullname = ?"
        If Me.LName <> "" Then sql = sql & ", lastname = ? "
        If Me.FName <> "" Then sql = sql & ", firstname = ? "
        If Me.MidName <> "" Then sql = sql & ", middlename = ? "
        If Me.Addr1 <> "" Then sql = sql & ", addr_line1 = ? "
        If Me.Addr2 <> "" Then sql = sql & ", addr_line2 = ? "
        If Me.City <> "" Then sql = sql & ", city = ? "
        If Me.State <> "" Then sql = sql & ", st = ? "
        If Me.Zip <> "" Then sql = sql & ", zip = ? "
        If Me.Ctry <> "" Then sql = sql & ", ctry = ? "
        If Me.AA <> "" Then sql = sql & ", aa = ? "
        If Me.SSNum <> "" Then sql = sql & ", ss_no = ? "
        If Me.Decsd <> "" Then sql = sql & ", decsd = ? "

        If Me.PrevNameID <> "" Then sql = sql & ", prev_name_id = ? "
        If Me.dUpdDte <> "" Then sql = sql & ", upd_date = ? "
        If Me.InquiryNum <> "" Then sql = sql & ", inquiry_no = ? "


        If Me.AddDate <> "" Then sql = sql & ", add_date = ? "
        If Me.OffcAddBy <> "" Then sql = sql & ", ofc_add_by = ? "
        sql = sql & "WHERE ID = ? "


        Dim cmd As New OdbcCommand(sql, ConnInformix)

        Try

            ConnInformix.ConnectionString = gl.ConnectString
            ConnInformix.Open()

            cmd.CommandType = CommandType.Text

            If Me.FullName > "" Then cmd.Parameters.AddWithValue("@fullname", Me.FullName)
            If Me.LName <> "" Then cmd.Parameters.AddWithValue("@lastname", Me.LName)
            If Me.FName <> "" Then cmd.Parameters.AddWithValue("@firstname", Me.FName)
            If Me.MidName <> "" Then cmd.Parameters.AddWithValue("@middlename", Me.MidName)
            If Me.Addr1 <> "" Then cmd.Parameters.AddWithValue("@addr_line1", Me.Addr1)
            If Me.Addr2 <> "" Then cmd.Parameters.AddWithValue("@addr_line2", Me.Addr2)
            If Me.City <> "" Then cmd.Parameters.AddWithValue("@city", Me.City)
            If Me.State <> "" Then cmd.Parameters.AddWithValue("@st", Me.State)
            If Me.Zip <> "" Then cmd.Parameters.AddWithValue("@zip", Me.Zip)
            If Me.Ctry <> "" Then cmd.Parameters.AddWithValue("@ctry", Me.Ctry)
            If Me.AA <> "" Then cmd.Parameters.AddWithValue("@aa", Me.AA)
            If Me.SSNum <> "" Then cmd.Parameters.AddWithValue("@ss_no", Me.SSNum) 'Ron says full SSN needed for Fin Aid, other downstream processes
            If Me.Decsd <> "" Then cmd.Parameters.AddWithValue("@decsd", Me.Decsd)
            If Me.AddDate <> "" Then cmd.Parameters.AddWithValue("@add_date", Me.AddDate)
            If Me.OffcAddBy <> "" Then cmd.Parameters.AddWithValue("@ofc_add_by", Me.OffcAddBy)
            If Me.InquiryNum <> "" Then cmd.Parameters.AddWithValue("@inquiry_no", Me.InquiryNum)
            If Me.CorrectAddr <> "" Then cmd.Parameters.AddWithValue("@correct_addr", Me.CorrectAddr)

            If Me.dUpdDte <> "" Then cmd.Parameters.AddWithValue("@upd_date", Me.dUpdDte)
            If Me.PrevNameID <> "" Then cmd.Parameters.AddWithValue("@prev_name_id", Me.PrevNameID)

            cmd.Parameters.AddWithValue("@ID", Me.ID)


            'Dim rowsAffected As Integer = cmd.ExecuteNonQuery()   'Executes query and returns a value
            'debug.writeline("RowsAffected: {0}", rowsAffected.ToString)
            ConnInformix.Close()
            ConnInformix.Dispose()
            WriteLog("Update of ID record for " & Me.FName & " " & Me.LName & ", ID = " & Me.ID)

        Catch e As Exception
            debug.writeline(e)
            WriteError("Error in clsid_rec - Update.  Error = " & e.Message)
            'MAIL Event
            'sMl.SendEMessage("ADP to CX", (""Error in clsid_rec - Update.  Error = " & e.Message), "dsullivan@carthage.edu", "dsullivan@carthag.edu", "dsullivan@carthage.edu")
        Finally
            If ConnInformix.State = ConnectionState.Open Then
                ConnInformix.Close()
                ConnInformix.Dispose()
            End If
            GC.Collect()
        End Try
    End Sub

    Public Sub Update_Addr()

        Dim sql As String
        sql = "Update id_rec set "
        If Me.Addr1 <> "" Then sql = sql & " addr_line1 = ? "
        If Me.Addr2 <> "" Then sql = sql & ", addr_line2 = ? "
        If Me.Addr3 <> "" Then sql = sql & ", addr_line3 = ? "
        If Me.City <> "" Then sql = sql & ", city = ? "
        If Me.State <> "" Then sql = sql & ", st = ? "
        If Me.Zip <> "" Then sql = sql & ", zip = ? "
        If Me.Ctry <> "" Then sql = sql & ", ctry = ? "
        If Me.AA <> "" Then sql = sql & ", aa = ? "
        sql = sql & "WHERE ID = ? "


        Dim cmd As New OdbcCommand(sql, ConnInformix)

        Try

            ConnInformix.ConnectionString = gl.ConnectString
            ConnInformix.Open()

            cmd.CommandType = CommandType.Text

            If Me.Addr1 <> "" Then cmd.Parameters.AddWithValue("@addr_line1", Me.Addr1)
            If Me.Addr2 <> "" Then cmd.Parameters.AddWithValue("@addr_line2", Me.Addr2)
            If Me.Addr3 <> "" Then cmd.Parameters.AddWithValue("@addr_line3", Me.Addr2)
            If Me.City <> "" Then cmd.Parameters.AddWithValue("@city", Me.City)
            If Me.State <> "" Then cmd.Parameters.AddWithValue("@st", Me.State)
            If Me.Zip <> "" Then cmd.Parameters.AddWithValue("@zip", Me.Zip)
            If Me.Ctry <> "" Then cmd.Parameters.AddWithValue("@ctry", Me.Ctry)
            If Me.AA <> "" Then cmd.Parameters.AddWithValue("@aa", Me.AA)
            cmd.Parameters.AddWithValue("@ID", Me.ID)


            'Dim rowsAffected As Integer = cmd.ExecuteNonQuery()   'Executes query and returns a value
            'debug.writeline("RowsAffected: {0}", rowsAffected.ToString)
            ConnInformix.Close()
            ConnInformix.Dispose()
            WriteLog("Update of Address info in ID record for " & Me.FullName & ", ID = " & Me.ID & ", Addr = " & Me.Addr1 & " " & Me.Addr2 _
                     & " " & Me.Addr3 & ", " & Me.City & ", " & Me.State & ", " & Me.Zip)

        Catch e As Exception
            debug.writeline(e)
            WriteError("Error in clsid_rec - UpdateAddr.  Error = " & e.Message & ", " & Me.FullName & ", ID = " & Me.ID)
            'MAIL Event
            'sMl.SendEMessage("ADP to CX", (""Error in clsid_rec - UpdateAddr.  Error = " & e.Message & ", " & Me.FullName & ", ID = " & Me.ID), "dsullivan@carthage.edu", "dsullivan@carthag.edu", "dsullivan@carthage.edu")
        Finally
            If ConnInformix.State = ConnectionState.Open Then
                ConnInformix.Close()
                ConnInformix.Dispose()
            End If
            GC.Collect()

        End Try
    End Sub

    Public Sub Update_Basic()

        Dim sql As String


        sql = "Update id_rec set "
        If Me.FullName <> "" Then sql = sql & " fullname = ?"
        If Me.LName <> "" Then sql = sql & ", lastname = ? "
        If Me.FName <> "" Then sql = sql & ", firstname = ? "
        If Me.MidName <> "" Then sql = sql & ", middlename = ? "
        If Me.SSNum <> "" Then sql = sql & ", ss_no = ? "
        If Me.Decsd <> "" Then sql = sql & ", decsd = ? "
        If Me.AddDate <> "" Then sql = sql & ", add_date = ? "
        If Me.UpDte <> "" Then sql = sql & ", upd_date = ? "
        If Me.OffcAddBy <> "" Then sql = sql & ", ofc_add_by = ? "
        sql = sql & "WHERE ID = ? "


        Dim cmd As New OdbcCommand(sql, ConnInformix)

        Try

            If Not Me.ID > 0 Then
                Throw New Exception("ID not valid.  Update clsID_rec failed")
            End If

            ConnInformix.ConnectionString = gl.ConnectString
            ConnInformix.Open()

            cmd.CommandType = CommandType.Text

            If Me.FullName > "" Then cmd.Parameters.AddWithValue("@fullname", Me.FullName)
            If Me.LName <> "" Then cmd.Parameters.AddWithValue("@lastname", Me.LName)
            If Me.FName <> "" Then cmd.Parameters.AddWithValue("@firstname", Me.FName)
            If Me.MidName <> "" Then cmd.Parameters.AddWithValue("@middlename", Me.MidName)
            If Me.SSNum <> "" Then cmd.Parameters.AddWithValue("@ss_no", Me.SSNum)
            If Me.Decsd <> "" Then cmd.Parameters.AddWithValue("@decsd", Me.Decsd)
            If Me.AddDate <> "" Then cmd.Parameters.AddWithValue("@add_date", Me.AddDate)
            If Me.UpDte <> "" Then cmd.Parameters.AddWithValue("@upd_date", Me.UpDte)
            If Me.OffcAddBy <> "" Then cmd.Parameters.AddWithValue("@ofc_add_by", Me.OffcAddBy)
            cmd.Parameters.AddWithValue("@ID", Me.ID)


            'Dim rowsAffected As Integer = cmd.ExecuteNonQuery()   'Executes query and returns a value
            'debug.writeline("RowsAffected: {0}", rowsAffected.ToString)
            ConnInformix.Close()
            ConnInformix.Dispose()
            WriteLog("Update of Basic name related info in ID record for " & Me.FullName & ", ID = " & Me.ID)

        Catch e As Exception
            debug.writeline(e)
            WriteError("Error in clsid_rec - Update_Basic.  Error = " & e.Message & ", " & Me.FullName & ", ID = " & Me.ID)
            'MAIL Event
            'sMl.SendEMessage("ADP to CX", (""Error in clsid_rec - Update_Basic.  Error = " & e.Message), "dsullivan@carthage.edu", "dsullivan@carthag.edu", "dsullivan@carthage.edu")
        Finally
            If ConnInformix.State = ConnectionState.Open Then
                ConnInformix.Close()
                ConnInformix.Dispose()
            End If
            GC.Collect()
        End Try
    End Sub


    Public Sub Initialize()

        iID = 0
        iPrspNo = 0
        sNameSndx = ""
        sFName = ""
        sLName = ""
        sMidName = ""
        sSuffixName = ""
        sAddr1 = ""
        sAddr2 = ""
        sAddr3 = ""
        sCity = ""
        sState = ""
        sState = ""
        sZip = ""
        sCtry = ""
        sAA = ""
        sTitle = ""
        sSuffix = ""
        sSSNum = ""
        sPhone = ""
        sPhoneExt = ""
        iPrevNameID = 0
        sMail = ""
        sSol = ""
        sPub = ""
        sCorrectAddr = ""
        sDecsd = ""
        dAddDate = "01/01/1999"
        sOffcAddBy = ""
        dUpdDte = "01/01/1999"
        sValid = ""
        dPurgeDate = "01/01/1999"
        dCassCertDate = "01/01/1999"
        sCCUserName = ""
        sCCPwd = ""
        iInquiryNum = 0


    End Sub
End Class