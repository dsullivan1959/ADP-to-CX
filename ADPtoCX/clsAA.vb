Option Strict Off
Option Explicit On

Imports System.Data.Odbc

Public Class clsAA

    Dim iID As Integer
    Dim sAA As String
    Dim sBegDate As String
    Dim sEndDate As String
    Dim sPeren As String
    Dim sLine1 As String
    Dim sLine2 As String
    Dim sLine3 As String
    Dim sCity As String
    Dim sState As String
    Dim sZip As String
    Dim sCntry As String
    Dim sPhone As String
    Dim sPhoneExt As String
    Dim sOfcAddBy As String
    Dim sCassCertDate As String
    Dim iAANum As Integer   ' This is a serial, no update
    Dim sCelCarrier As String
    Dim sOptOut As String
    Dim sFullname As String

    Public Property AA() As String
        Get
            AA = sAA
        End Get
        Set(ByVal Value As String)
            sAA = Value
        End Set
    End Property
    Public Property ID() As Integer
        Get
            ID = iID
        End Get
        Set(ByVal Value As Integer)
            iID = Value
        End Set
    End Property

    Public Property BegDate() As String
        Get
            BegDate = sBegDate
        End Get
        Set(ByVal Value As String)
            sBegDate = Value
        End Set
    End Property

    Public Property EndDate() As String
        Get
            EndDate = sEndDate
        End Get
        Set(ByVal Value As String)
            sEndDate = Value
        End Set
    End Property

    Public Property Peren() As String
        Get
            Peren = sPeren
        End Get
        Set(ByVal Value As String)
            sPeren = Value
        End Set
    End Property


    Public Property Line1() As String
        Get
            Line1 = sLine1
        End Get
        Set(ByVal Value As String)
            sLine1 = Value
        End Set
    End Property

    Public Property Line2() As String
        Get
            Line2 = sLine2
        End Get
        Set(ByVal Value As String)
            sLine2 = Value
        End Set
    End Property

    Public Property Line3() As String
        Get
            Line3 = sLine3
        End Get
        Set(ByVal Value As String)
            sLine3 = Value
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
    Public Property Cntry() As String
        Get
            Cntry = sCntry
        End Get
        Set(ByVal Value As String)
            sCntry = Value
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
    Public Property OfcAddBy() As String
        Get
            OfcAddBy = sOfcAddBy
        End Get
        Set(ByVal Value As String)
            sOfcAddBy = Value
        End Set
    End Property
    Public Property CassCertDate() As String
        Get
            CassCertDate = sCassCertDate
        End Get
        Set(ByVal Value As String)
            sCassCertDate = Value
        End Set
    End Property

    Public Property AANum() As Integer
        Get
            AANum = iAANum
        End Get
        Set(ByVal Value As Integer)
            iAANum = Value
        End Set
    End Property
    Public Property CelCarrier() As String
        Get
            CelCarrier = sCelCarrier
        End Get
        Set(ByVal Value As String)
            sCelCarrier = Value
        End Set
    End Property
    Public Property OptOut() As String
        Get
            OptOut = sOptOut
        End Get
        Set(ByVal Value As String)
            sOptOut = Value
        End Set
    End Property

    Public Property FullName() As String
        Get
            FullName = sFullname
        End Get
        Set(ByVal Value As String)
            sFullname = Value
        End Set
    End Property
    'tEST
    '    Public Sub Insert1()

    '        Dim sql As String
    'sql = "INSERT INTO aa_rec(id, aa, beg_date)"
    'sql = sql & "VALUES (11111,?,?)"


    '    sql = "INSERT INTO AA_REC(ID,AA,BEG_DATE) VALUES(1111113, " & ConvertString("PREV") & ", " & ConvertString("01-01-2017") & ")"
    '    Dim cmd As New OdbcCommand(sql, ConnInformix)

    '    Try

    '        ConnInformix.ConnectionString = gl.ConnectString
    '        ConnInformix.Open()

    '        cmd.CommandType = CommandType.Text

    '        'cmd.Parameters.AddWithValue("@id", Me.ID)
    '        'cmd.Parameters.AddWithValue("@aa", Me.AA)
    '        cmd.Parameters.AddWithValue("@beg_date", Me.BegDate)


    '        Dim rowsAffected As Integer = cmd.ExecuteNonQuery()   'Executes query and returns a value
    '        Debug.WriteLine("RowsAffected: {0}", rowsAffected.ToString)

    '        ConnInformix.Close()
    '        WriteLog("Add of aa_Rec completed for ID number " & Me.ID & ", AA = " & Me.AA)

    '    Catch ex As DuplicateNameException
    '        'No need to write anything
    '    Catch e As Exception
    '        Debug.WriteLine(e)
    '        WriteError("Error in aa_rec - Insert.  Error = " & e.Message)
    '        'MAIL Event
    '        'sMl.SendEMessage("ADP to CX", (""Error in aa_rec - insert.  Error = " & e.Message), "dsullivan@carthage.edu", "dsullivan@carthag.edu", "dsullivan@carthage.edu")

    '    End Try
    '    If ConnInformix.State = ConnectionState.Open Then
    '        ConnInformix.Close()
    '    End If
    '    GC.Collect()
    'End Sub

    'tEST




    Public Sub Insert()

        Dim sql As String
        sql = "INSERT INTO aa_rec(id, aa, beg_date, peren, end_date, line1, line2, line3, city, st, zip, ctry, phone, phone_ext, ofc_add_by, cell_carrier, opt_out) "
        sql = sql & " VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)"

        Dim cmd As New OdbcCommand(sql, ConnInformix)

        Try

            ConnInformix.ConnectionString = gl.ConnectString
            ConnInformix.Open()

            cmd.CommandType = CommandType.Text

            cmd.Parameters.AddWithValue("@id", Me.ID)
            cmd.Parameters.AddWithValue("@aa", Me.AA)
            cmd.Parameters.AddWithValue("@beg_date", Me.BegDate)
            cmd.Parameters.AddWithValue("@peren", Me.Peren)
            If Me.EndDate = "" Then
                cmd.Parameters.AddWithValue("@end_date", DBNull.Value)
            Else
                cmd.Parameters.AddWithValue("@end_date", Me.EndDate)
            End If
            cmd.Parameters.AddWithValue("@line1", Me.Line1)
            cmd.Parameters.AddWithValue("@line2", Me.Line2)
            cmd.Parameters.AddWithValue("@line3", Me.Line3)
            cmd.Parameters.AddWithValue("@city", Me.City)
            cmd.Parameters.AddWithValue("@st", Me.State)
            cmd.Parameters.AddWithValue("@zip", Me.Zip)
            cmd.Parameters.AddWithValue("@ctry", Me.Cntry)
            cmd.Parameters.AddWithValue("@phone", Me.Phone)
            cmd.Parameters.AddWithValue("@phone_ext", Me.PhoneExt)
            cmd.Parameters.AddWithValue("@ofc_add_by", Me.OfcAddBy)
            'If Me.CassCertDate = "" Then
            '    cmd.Parameters.AddWithValue("@cass_cert_date", vbNull)
            'Else
            '    cmd.Parameters.AddWithValue("@cass_cert_date", Me.CassCertDate)
            'End If
            cmd.Parameters.AddWithValue("@cell_carrier", Me.CelCarrier)
            cmd.Parameters.AddWithValue("@opt_out", Me.OptOut)

            'Note.   Beg_Date is a unique key.  Can't have two addresses in aa rec with same beg_date


            ' Dim rowsAffected As Integer = cmd.ExecuteNonQuery()   'Executes query and returns a value
            ' Debug.WriteLine("RowsAffected: {0}", rowsAffected.ToString)

            ConnInformix.Close()
            ConnInformix.Dispose()
            WriteLog("Add of alternate address completed for ID number " & Me.FullName & ", " & Me.ID & ", AA = " & Me.AA & ", Detail = " & Me.Line1)

        Catch ex As DuplicateNameException
            'No need to write anything
        Catch e As Exception
            If InStr(1, e.ToString, "duplicate key") = 0 Then
                Debug.WriteLine(" User already exists cannot insert duplicate value.")
            Else
                Debug.WriteLine(e)
                WriteError("Error in aa_rec - Insert.  Error = " & e.Message & ", " & Me.FullName & ", ID = " & Me.ID & ", AA = " & Me.AA & ", Detail = " & Me.Line1)
                'MAIL Event
                'sMl.SendEMessage("ADP to CX", (""Error in aa_rec - insert.  Error = " & e.Message), "dsullivan@carthage.edu", "dsullivan@carthag.edu", "dsullivan@carthage.edu")
            End If
        Finally
            If ConnInformix.State = ConnectionState.Open Then
                ConnInformix.Close()
                ConnInformix.Dispose()
            End If
            GC.Collect()
        End Try
    End Sub



    Public Sub Update()
        'id, aa, end_date, peren, line1, line2, line3, city, st, zip, ctry, phone, phone_ext, ofc_add_by, cass_cert_date, cell_carrier, opt_out)
        Dim sql As String
        sql = "Update aa_rec set "

        If Me.Line1 <> "" Then sql = sql & " line1 = ?"
        If Me.Line2 <> "" Then sql = sql & ", line2 = ?"
        If Me.Line3 <> "" Then sql = sql & ", line3 = ?"
        If Me.City <> "" Then sql = sql & ", city = ?"
        If Me.State <> "" Then sql = sql & ", st = ?"
        If Me.Zip <> "" Then sql = sql & ", zip = ?"
        If Me.Cntry <> "" Then sql = sql & ", ctry = ?"
        If Me.Phone <> "" Then sql = sql & ", phone = ?"
        If Me.PhoneExt <> "" Then sql = sql & ", phone_ext = ?"
        If Me.OfcAddBy <> "" Then sql = sql & ", ofc_add_by = ?"
        '        If Me.BegDate <> "" Then sql = sql & ", beg_date = ?"
        If Me.EndDate <> "" Then sql = sql & ", end_date = ?"
        If Me.Peren <> "" Then sql = sql & ", peren = ?"

        If Me.CassCertDate <> "" Then sql = sql & ", cass_cert_date = ?"

        If Me.CelCarrier <> "" Then sql = sql & ", cell_carrier = ?"
        If Me.OptOut <> "" Then sql = sql & ", opt_out = ?"

        sql = sql & " where id = ?"
        sql = sql & " and aa = ?"

        Dim cmd As New OdbcCommand(sql, ConnInformix)

        Try

            'I believe the preferred method these days is to NOT leave the connection open, but connect, execute and close...
            ConnInformix.ConnectionString = gl.ConnectString
            ConnInformix.Open()
            cmd.CommandType = CommandType.Text
            If Me.Line1 <> "" Then cmd.Parameters.AddWithValue("@line1", Me.Line1)
            If Me.Line2 <> "" Then cmd.Parameters.AddWithValue("@line2", Me.Line2)
            If Me.Line3 <> "" Then cmd.Parameters.AddWithValue("@line3", Me.Line3)
            If Me.City <> "" Then cmd.Parameters.AddWithValue("@city", Me.City)
            If Me.State <> "" Then cmd.Parameters.AddWithValue("@st", Me.State)
            If Me.Zip <> "" Then cmd.Parameters.AddWithValue("@zip", Me.Zip)
            If Me.Cntry <> "" Then cmd.Parameters.AddWithValue("@ctry", Me.Cntry)
            If Me.Phone <> "" Then cmd.Parameters.AddWithValue("@phone", Me.Phone)
            If Me.PhoneExt <> "" Then cmd.Parameters.AddWithValue("@phone_ext", Me.PhoneExt)
            If Me.OfcAddBy <> "" Then cmd.Parameters.AddWithValue("@ofc_add_by", Me.OfcAddBy)
            If Me.CassCertDate <> "" Then cmd.Parameters.AddWithValue("@cass_cert_date", Me.CassCertDate)
            '  If Me.BegDate <> "" Then cmd.Parameters.AddWithValue("@beg_date", Me.BegDate)
            If Me.EndDate <> "" Then cmd.Parameters.AddWithValue("@end_date", Me.EndDate)
            If Me.Peren <> "" Then cmd.Parameters.AddWithValue("@peren", Me.Peren)
            If Me.CelCarrier <> "" Then cmd.Parameters.AddWithValue("@cell_carrier", Me.CelCarrier)
            If Me.OptOut <> "" Then cmd.Parameters.AddWithValue("@opt_out", Me.OptOut)
            cmd.Parameters.AddWithValue("@id", Me.ID)
            cmd.Parameters.AddWithValue("@aa", Me.AA)

            '   Dim rowsAffected As Integer = cmd.ExecuteNonQuery()   'Executes query and returns a value
            '   Debug.WriteLine("RowsAffected: {0}", rowsAffected.ToString)

            ConnInformix.Close()
            ConnInformix.Dispose()
            WriteLog("Update of AA_Rec completed for " & Me.FullName & ", ID = " & Me.ID & ", AA = " & Me.AA & ", Detail = " & Me.Line1)

        Catch e As Exception
            Debug.WriteLine(e)
            WriteError("Error in clsAA - Update.  Error = " & e.Message & ", " & Me.FullName & ", ID = " & Me.ID & ", AA = " & Me.AA & ", Detail = " & Me.Line1)
            'MAIL Event
            'sMl.SendEMessage("ADP to CX", (""Error in aa_rec - Update.  Error = " & e.Message), "dsullivan@carthage.edu", "dsullivan@carthag.edu", "dsullivan@carthage.edu")
        Finally
            If ConnInformix.State = ConnectionState.Open Then
                ConnInformix.Close()
                ConnInformix.Dispose()
            End If
        End Try
    End Sub


    Public Sub ChangeType()
        'id, aa, beg_date, end_date, peren, line1, line2, line3, city, st, zip, ctry, phone, phone_ext, ofc_add_by, cass_cert_date, cell_carrier, opt_out)
        Dim sql As String
        sql = "Update aa_rec set aa = ?"
        sql = sql & " where aa = ?"
        sql = sql & " and id = ?"
        sql = sql & " and aa_no = ?"

        Dim cmd As New OdbcCommand(sql, ConnInformix)

        Try

            'I believe the preferred method these days is to NOT leave the connection open, but connect, execute and close...
            ConnInformix.ConnectionString = gl.ConnectString
            ConnInformix.Open()
            cmd.CommandType = CommandType.Text
            If Me.AA <> "" Then cmd.Parameters.AddWithValue("@aa", Me.AA)
            cmd.Parameters.AddWithValue("@id", Me.ID)
            cmd.Parameters.AddWithValue("@aa_no", Me.AANum)

            '  Dim rowsAffected As Integer = cmd.ExecuteNonQuery()   'Executes query and returns a value
            '  Debug.WriteLine("RowsAffected: {0}", rowsAffected.ToString)

            ConnInformix.Close()
            ConnInformix.Dispose()
            WriteLog("Update of AA_Rec completed for " & Me.FullName & ", ID = " & Me.ID & ", AA = " & Me.AA & ", Detail = " & Me.Line1)

        Catch e As Exception
            Debug.WriteLine(e)
            WriteError("Error in clsAA - Change Type.  Error = " & Err.Description & ", Error Number " & Err.Number & ", " & Me.FullName & ", ID = " & Me.ID & ", AA = " & Me.AA & ", Detail = " & Me.Line1)
            'MAIL Event
            'sMl.SendEMessage("ADP to CX", (""Error in aa_rec - Change Type.  Error = " & e.Message), "dsullivan@carthage.edu", "dsullivan@carthag.edu", "dsullivan@carthage.edu")
        Finally
            If ConnInformix.State = ConnectionState.Open Then
                ConnInformix.Close()
                ConnInformix.Dispose()
            End If
            GC.Collect()
        End Try
    End Sub

    Public Sub AddEmail2()
        Dim sql As String

        sql = "INSERT INTO aa_rec(id, aa, beg_date, line1) VALUES (?,?,?,?)"

        Dim cmd As New OdbcCommand(sql, ConnInformix)

        Try

            'I believe the preferred method these days is to NOT leave the connection open, but connect, execute and close...
            ConnInformix.ConnectionString = gl.ConnectString
            ConnInformix.Open()
            cmd.CommandType = CommandType.Text
            cmd.Parameters.AddWithValue("@id", Me.ID)
            cmd.Parameters.AddWithValue("@aa", Me.AA)
            cmd.Parameters.AddWithValue("@beg_date", Me.BegDate)
            cmd.Parameters.AddWithValue("@line1", Me.Line1)
            cmd.Parameters.AddWithValue("@ofc_add_by", Me.OfcAddBy)


            'Dim rowsAffected As Integer = cmd.ExecuteNonQuery()   'Executes query and returns a value
            'Debug.WriteLine("RowsAffected: {0}", rowsAffected.ToString)

            ConnInformix.Close()
            ConnInformix.Dispose()
            WriteLog("Add of personal email to AA_Rec completed for " & Me.FullName & ", ID = " & Me.ID & ", AA = " & Me.AA & ", Detail = " & Me.Line1)

        Catch e As Exception
            Debug.WriteLine(e)
            WriteError("Error in clsAA - Add Email2.   Error = " & Err.Description & ", Error Number " & Err.Number & ", " & Me.FullName & ", ID = " & Me.ID & ", AA = " & Me.AA & ", Detail = " & Me.Line1)
            'MAIL Event
            'sMl.SendEMessage("ADP to CX", (""Error in aa_rec - Change Type.  Error = " & e.Message), "dsullivan@carthage.edu", "dsullivan@carthag.edu", "dsullivan@carthage.edu")
        Finally
            If ConnInformix.State = ConnectionState.Open Then
                ConnInformix.Close()
                ConnInformix.Dispose()
            End If
            GC.Collect()
        End Try

    End Sub
    Public Sub UpdateEmail2()
        Dim sql As String
        sql = "Update aa_rec set Line1 = ?"
        sql = sql & " where aa = ?"
        sql = sql & " and id = ?"
        sql = sql & " and aa_no = ?"
        sql = sql & " and beg_date = ?"

        Dim cmd As New OdbcCommand(sql, ConnInformix)

        Try

            'I believe the preferred method these days is to NOT leave the connection open, but connect, execute and close...
            ConnInformix.ConnectionString = gl.ConnectString
            ConnInformix.Open()
            cmd.CommandType = CommandType.Text
            If Me.Line1 <> "" Then cmd.Parameters.AddWithValue("@line1", Me.Line1)
            cmd.Parameters.AddWithValue("@id", Me.ID)
            cmd.Parameters.AddWithValue("@aa", Me.AA)
            cmd.Parameters.AddWithValue("@aa_no", Me.AANum)
            cmd.Parameters.AddWithValue("@beg_date", Me.BegDate)

            ' Dim rowsAffected As Integer = cmd.ExecuteNonQuery()   'Executes query and returns a value
            'Debug.WriteLine("RowsAffected: {0}", rowsAffected.ToString)

            ConnInformix.Close()
            ConnInformix.Dispose()
            WriteLog("Update of EMAIL2 completed for " & Me.FullName & ", " & Me.ID & ", AA = " & Me.AA & ", Detail = " & Me.Line1)

        Catch e As Exception
            Debug.WriteLine(e)
            WriteError("Error in clsAA - Change Type.  Error = " & Err.Description & ", Error Number " & Err.Number & ", " & Me.FullName & ", ID = " & Me.ID & ", AA = " & Me.AA & ", Detail = " & Me.Line1)
            'MAIL Event
            'sMl.SendEMessage("ADP to CX", (""Error in aa_rec - Change Type.  Error = " & e.Message), "dsullivan@carthage.edu", "dsullivan@carthag.edu", "dsullivan@carthage.edu")
        Finally
            If ConnInformix.State = ConnectionState.Open Then
                ConnInformix.Close()
                ConnInformix.Dispose()
            End If
            GC.Collect()
        End Try
    End Sub

    Public Sub AddCell()
        Dim sql As String

        sql = "INSERT INTO aa_rec(id, aa, beg_date, phone) VALUES (?,?,?,?)"

        Dim cmd As New OdbcCommand(sql, ConnInformix)

        Try

            'I believe the preferred method these days is to NOT leave the connection open, but connect, execute and close...
            ConnInformix.ConnectionString = gl.ConnectString
            ConnInformix.Open()
            cmd.CommandType = CommandType.Text
            cmd.Parameters.AddWithValue("@id", Me.ID)
            cmd.Parameters.AddWithValue("@aa", Me.AA)
            cmd.Parameters.AddWithValue("@beg_date", Me.BegDate)
            cmd.Parameters.AddWithValue("@phone", Me.Phone)
            cmd.Parameters.AddWithValue("@ofc_add_by", Me.OfcAddBy)


            'Dim rowsAffected As Integer = cmd.ExecuteNonQuery()   'Executes query and returns a value
            'Debug.WriteLine("RowsAffected: {0}", rowsAffected.ToString)

            ConnInformix.Close()
            ConnInformix.Dispose()
            WriteLog("Add of cell phone to AA_Rec completed for " & Me.FullName & ", ID = " & Me.ID & ", AA = " & Me.AA & ", Detail = " & Me.Phone)

        Catch e As Exception
            Debug.WriteLine(e)
            WriteError("Error in clsAA - Add Cell.   Error = " & Err.Description & ", Error Number " & Err.Number & ", " & Me.FullName & ", ID = " & Me.ID & ", AA = " & Me.AA & ", Detail = " & Me.Phone)
            'MAIL Event
            'sMl.SendEMessage("ADP to CX", (""Error in aa_rec - Add Cell.  Error = " & e.Message), "dsullivan@carthage.edu", "dsullivan@carthag.edu", "dsullivan@carthage.edu")
        Finally
            If ConnInformix.State = ConnectionState.Open Then
                ConnInformix.Close()
                ConnInformix.Dispose()
            End If
            GC.Collect()
        End Try

    End Sub
    Public Sub UpdateCell()
        Dim sql As String
        sql = "Update aa_rec set phone = ?"
        sql = sql & " where aa = ?"
        sql = sql & " and id = ?"
        sql = sql & " and aa_no = ?"
        sql = sql & " and beg_date = ?"

        Dim cmd As New OdbcCommand(sql, ConnInformix)

        Try

            'I believe the preferred method these days is to NOT leave the connection open, but connect, execute and close...
            ConnInformix.ConnectionString = gl.ConnectString
            ConnInformix.Open()
            cmd.CommandType = CommandType.Text
            If Me.Phone <> "" Then cmd.Parameters.AddWithValue("@phone", Me.Phone)
            cmd.Parameters.AddWithValue("@aa", Me.AA)
            cmd.Parameters.AddWithValue("@id", Me.ID)
            cmd.Parameters.AddWithValue("@aa_no", Me.AANum)
            cmd.Parameters.AddWithValue("@beg_date", Me.BegDate)

            'Dim rowsAffected As Integer = cmd.ExecuteNonQuery()   'Executes query and returns a value
            'Debug.WriteLine("RowsAffected: {0}", rowsAffected.ToString)

            ConnInformix.Close()
            ConnInformix.Dispose()
            WriteLog("Update of cell phone completed for " & Me.FullName & ", " & Me.ID & ", AA = " & Me.AA & ", Detail = " & Me.Phone)

        Catch e As Exception
            Debug.WriteLine(e)
            WriteError("Error in clsAA - Update Cell.  Error = " & Err.Description & ", Error Number " & Err.Number & ", " & Me.FullName & ", ID = " & Me.ID & ", AA = " & Me.AA & ", Detail = " & Me.Phone)
            'MAIL Event
            'sMl.SendEMessage("ADP to CX", (""Error in aa_rec - Update Phone.  Error = " & e.Message), "dsullivan@carthage.edu", "dsullivan@carthag.edu", "dsullivan@carthage.edu")
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
        sAA = ""
        sBegDate = ""
        sEndDate = ""
        sPeren = ""
        sLine1 = ""
        sLine2 = ""
        sLine3 = ""
        sCity = ""
        sState = ""
        sZip = ""
        sCntry = ""
        sPhone = ""
        sPhoneExt = ""
        sOfcAddBy = ""
        sCassCertDate = ""
        iAANum = 0
        sCelCarrier = ""
        sOptOut = ""
        sFullname = ""



    End Sub
End Class
