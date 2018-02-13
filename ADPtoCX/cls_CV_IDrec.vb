Option Strict Off
Option Explicit On

Imports System.Data.Odbc

Public Class Cls_CV_IDrec

    Dim sOldId As String
    Dim iOldIDNum As Integer
    Dim sADPID As String
    Dim sSSN As String
    Dim iCXID As Integer
    Dim sCXIDst As String
    Dim sLDAPName As String
    Dim dLDAPAddDate As String  'Format seems to be yyyy-mm-dd
    Dim sJobCodePlaceholder As String 'This is only for the log.
    Dim sFullName As String 'This is only for the log.


    Public Property OldId() As String
        Get
            OldId = sOldId
        End Get
        Set(ByVal Value As String)
            sOldId = Value
        End Set
    End Property
    Public Property OldIdNum() As Integer
        Get
            OldIdNum = iOldIDNum
        End Get
        Set(ByVal Value As Integer)
            iOldIDNum = Value
        End Set
    End Property

    Public Property ADPID() As String
        Get
            ADPID = sADPID
        End Get
        Set(ByVal Value As String)
            sADPID = Value
        End Set
    End Property

    Public Property SSN() As String
        Get
            SSN = sSSN
        End Get
        Set(ByVal Value As String)
            sSSN = Value
        End Set
    End Property

    Public Property CXID() As Integer
        Get
            CXID = iCXID
        End Get
        Set(ByVal Value As Integer)
            iCXID = Value
        End Set
    End Property


    Public Property CXIDst() As String
        Get
            CXIDst = sCXIDst
        End Get
        Set(ByVal Value As String)
            sCXIDst = Value
        End Set
    End Property

    Public Property LDAPName() As String
        Get
            LDAPName = sLDAPName

        End Get
        Set(ByVal Value As String)
            sLDAPName = Value
        End Set
    End Property

    Public Property LDAPAddDate() As String
        Get
            LDAPAddDate = dLDAPAddDate
        End Get
        Set(ByVal Value As String)
            dLDAPAddDate = Value
        End Set
    End Property

    Public Property JobCodePlaceholder() As String
        Get
            JobCodePlaceholder = sJobCodePlaceholder
        End Get
        Set(ByVal Value As String)
            sJobCodePlaceholder = Value
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


    Public Sub Insert()

        Dim sql As String
        sql = "INSERT INTO cvid_rec (old_id, old_id_num, adp_id, ssn, cx_id, cx_id_char) values (? ,?, ?, ?, ?, ?) "
        Dim cmd As New OdbcCommand(sql, ConnInformix)

        Try

            ConnInformix.ConnectionString = gl.ConnectString
            ConnInformix.Open()

            cmd.CommandType = CommandType.Text

            cmd.Parameters.AddWithValue("@old_id", Me.OldId)
            cmd.Parameters.AddWithValue("@old_id_num", Me.OldIdNum)
            cmd.Parameters.AddWithValue("@adp_id", Me.ADPID)
            cmd.Parameters.AddWithValue("@ssn", Me.SSN)
            cmd.Parameters.AddWithValue("@cx_id", Me.CXID)
            cmd.Parameters.AddWithValue("@cx_id_char", Me.CXIDst)

            'Dim rowsAffected As Integer = cmd.ExecuteNonQuery()   'Executes query and returns a value
            'debug.writeline("RowsAffected: {0}", rowsAffected.ToString)

            ConnInformix.Close()
            ConnInformix.Dispose()
            WriteLog("Add of CVID_Rec completed for " & Me.FullName & ", ID number " & Me.CXIDst & ", ADP ID = " & Me.ADPID & "Job Rec = " & JobCodePlaceholder)
        Catch ex As DuplicateNameException
            'No need to write anything
        Catch e As Exception
            debug.writeline(e)
            WriteError("Error in clscvid_rec - Insert.  Error = " & e.Message & ", " & Me.FullName & ", ID number " & Me.CXIDst & ", ADP ID = " & Me.ADPID & "Job Rec = " & JobCodePlaceholder)
            'MAIL Event
            'sMl.SendEMessage("ADP to CX", (""Error in clscvid_rec - Insert.  Error = " & e.Messag), "dsullivan@carthage.edu", "dsullivan@carthag.edu", "dsullivan@carthage.edu")
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
        'sql = "Update cvid_rec SET old_id = ?, old_id_num = ?, adp_id = ?, ssn = ?, cx_id_char = ?   WHERE cx_id =  ?"
        sql = "Update cvid_rec set "
        If Me.CXIDst <> "" Then sql = sql & " cx_id_char = ?"
        If Me.OldId <> "" Then sql = sql & ", old_id = ?"
        If Me.OldIdNum <> 0 Then sql = sql & ", old_id_num = ?"
        If Me.ADPID <> "" Then sql = sql & ", adp_id = ?"
        If Me.SSN <> "" Then sql = sql & ", ssn = ?"
        sql = sql & " WHERE cx_id = ?"

        Dim cmd As New OdbcCommand(sql, ConnInformix)
      
        Try

            'I believe the preferred method these days is to NOT leave the connection open, but connect, execute and close...
            ConnInformix.ConnectionString = gl.ConnectString
            ConnInformix.Open()
            cmd.CommandType = CommandType.Text
            If Me.CXIDst <> "" Then cmd.Parameters.AddWithValue("@cx_id_char", Me.CXIDst)
            If Me.OldId <> "" Then cmd.Parameters.AddWithValue("@old_id", Me.OldId)
            If Me.OldIdNum <> 0 Then cmd.Parameters.AddWithValue("@old_id_num", Me.OldIdNum)
            If Me.ADPID <> "" Then cmd.Parameters.AddWithValue("@adp_id", Me.ADPID)
            If Me.SSN <> "" Then cmd.Parameters.AddWithValue("@ssn", Me.SSN)
            cmd.Parameters.AddWithValue("@cx_id", Me.CXID)

            'Dim rowsAffected As Integer = cmd.ExecuteNonQuery()   'Executes query and returns a value
            'debug.writeline("RowsAffected: {0}", rowsAffected.ToString)

            ConnInformix.Close()
            ConnInformix.Dispose()
            WriteLog("Update of CVID_Rec completed for " & Me.FullName & ", ID = " & Me.CXIDst & ", ADP ID inserted as " & Me.ADPID)

        Catch e As Exception
            debug.writeline(e)
            WriteError("Error in clscvid_rec - Update.  Error = " & e.Message & ", " & Me.FullName & ", ID number " & Me.CXIDst & ", ADP ID = " & Me.ADPID & "Job Rec = " & JobCodePlaceholder)
            'MAIL Event
            'sMl.SendEMessage("ADP to CX", (""Error in clscvid_rec - Update.  Error = " & e.Messag), "dsullivan@carthage.edu", "dsullivan@carthag.edu", "dsullivan@carthage.edu")

        Finally
            If ConnInformix.State = ConnectionState.Open Then
                ConnInformix.Close()
                ConnInformix.Dispose()
                GC.Collect()
            End If
        End Try
    End Sub

   


   
   
    Public Sub Initialize()

        sOldId = ""
        iOldIDNum = 0
        sADPID = ""
        sSSN = ""
        iCXID = 0
        sCXIDst = ""
        sLDAPName = ""
        dLDAPAddDate = ""
        sJobCodePlaceholder = ""
        sFullName = ""

    End Sub
End Class