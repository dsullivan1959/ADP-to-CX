Option Strict Off
Option Explicit On

Imports System.Data.Odbc

Friend Class ClsJob_rec


    Dim iJobNo As Integer 'Primary Key, sequence added by Informix??
    Dim iTPosNo As Integer
    Dim sDescr As String '
    Dim sHrClass As String '
    Dim iBJobNo As Integer '
    Dim iID As Integer '
    Dim sHRPay As String

    Dim iSuprvNo As Integer
    Dim sHRStat As String
    Dim iPseudoSerial As Integer
    Dim sEGPType As String
    Dim sHRDiv As String
    Dim sHRDept As String
    Dim dCompBegDate As String
    Dim dCompEndDate As String
    Dim dAppInactDate As String
    Dim dBegDate As String
    Dim dEndDate As String
    Dim sActiveCtrct As String
    Dim sCtrctStat As String
    Dim sExclSrvcHrs As String
    Dim sJobTitle As String
    Dim iTitleRank As Integer
    Dim sFullName As String

    Public Property JobNo() As Integer
        Get
            JobNo = iJobNo
        End Get
        Set(ByVal Value As Integer)
            iJobNo = Value
        End Set
    End Property
    Public Property TPosNo() As Integer
        Get
            TPosNo = iTPosNo
        End Get
        Set(ByVal Value As Integer)
            iTPosNo = Value
        End Set
    End Property
    Public Property Descr() As String
        Get
            Descr = sDescr
        End Get
        Set(ByVal Value As String)
            sDescr = Value
        End Set
    End Property
    Public Property BJobNo() As Integer
        Get
            BJobNo = iBJobNo
        End Get
        Set(ByVal Value As Integer)
            iBJobNo = Value
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
    Public Property HRclass() As String
        Get
            HRclass = sHrClass
        End Get
        Set(ByVal Value As String)
            sHrClass = Value
        End Set
    End Property

    Public Property HRPay() As String
        Get
            HRPay = sHRPay
        End Get
        Set(ByVal Value As String)
            sHRPay = Value
        End Set
    End Property

    Public Property SuprvNo() As Integer
        Get
            SuprvNo = iSuprvNo
        End Get
        Set(ByVal Value As Integer)
            iSuprvNo = Value
        End Set
    End Property
    Public Property HRStat() As String
        Get
            HRStat = sHRStat
        End Get
        Set(ByVal Value As String)
            sHRStat = Value
        End Set
    End Property

    Public Property PseudoSerial() As Integer
        Get
            PseudoSerial = iPseudoSerial
        End Get
        Set(ByVal Value As Integer)
            iPseudoSerial = Value
        End Set
    End Property

    Public Property EGPType() As String
        Get
            EGPType = sEGPType
        End Get
        Set(ByVal Value As String)
            sEGPType = Value
        End Set
    End Property

    Public Property HRDiv() As String
        Get
            HRDiv = sHRDiv
        End Get
        Set(ByVal Value As String)
            sHRDiv = Value
        End Set
    End Property

    Public Property HRDept() As String
        Get
            HRDept = sHRDept
        End Get
        Set(ByVal Value As String)
            sHRDept = Value
        End Set
    End Property

    Public Property CompBegDate() As String
        Get
            CompBegDate = dCompBegDate
        End Get
        Set(ByVal Value As String)
            dCompBegDate = Value
        End Set
    End Property
    Public Property CompEndDate() As String
        Get
            CompEndDate = dCompEndDate
        End Get
        Set(ByVal Value As String)
            dCompEndDate = Value
        End Set
    End Property
    Public Property AppInactDate() As String
        Get
            AppInactDate = dAppInactDate
        End Get
        Set(ByVal Value As String)
            dAppInactDate = Value
        End Set
    End Property
    Public Property BegDate() As String
        Get
            BegDate = dBegDate
        End Get
        Set(ByVal Value As String)
            dBegDate = Value
        End Set
    End Property
    Public Property EndDate() As String
        Get
            EndDate = dEndDate
        End Get
        Set(ByVal Value As String)
            dEndDate = Value
        End Set
    End Property

    Public Property ActiveCtrct() As String
        Get
            ActiveCtrct = sActiveCtrct
        End Get
        Set(ByVal Value As String)
            sActiveCtrct = Value
        End Set
    End Property
    Public Property CtrctStat() As String
        Get
            CtrctStat = sCtrctStat
        End Get
        Set(ByVal Value As String)
            sCtrctStat = Value
        End Set
    End Property
    Public Property ExclSrvcHrs() As String
        Get
            ExclSrvcHrs = sExclSrvcHrs
        End Get
        Set(ByVal Value As String)
            sExclSrvcHrs = Value
        End Set
    End Property

    Public Property JobTitle() As String
        Get
            JobTitle = sJobTitle
        End Get
        Set(ByVal Value As String)
            sJobTitle = Value
        End Set
    End Property

    Public Property TitleRank() As Integer
        Get
            TitleRank = iTitleRank
        End Get
        Set(ByVal Value As Integer)
            iTitleRank = Value
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
        sql = " INSERT INTO job_rec (tpos_no, descr, hrclass, id, hrpay, "
        sql = sql & " supervisor_no, HRStat, egp_type, active_ctrct, ctrct_stat, excl_srvc_hrs,  "
        sql = sql & " hrdiv, hrdept, beg_date, end_date, job_title, title_rank)"
        sql = sql & " values (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)"
        Dim cmd As New OdbcCommand(sql, ConnInformix)

        Try

            'I believe the preferred method these days is to NOT leave the connection open, but connect, execute and close...

            ConnInformix.ConnectionString = gl.ConnectString
            ConnInformix.Open()

            cmd.CommandType = CommandType.Text
            cmd.Parameters.AddWithValue("@tpos_no", Me.TPosNo)
            cmd.Parameters.AddWithValue("@descr", Me.Descr)
            cmd.Parameters.AddWithValue("@hrclass", Me.HRclass)
            cmd.Parameters.AddWithValue("@id", Me.ID)
            cmd.Parameters.AddWithValue("@hrpay", Me.HRPay)
            cmd.Parameters.AddWithValue("@supervisor_no", Me.SuprvNo)
            cmd.Parameters.AddWithValue("@HRStat", Me.HRStat)
            cmd.Parameters.AddWithValue("@egp_type", Me.EGPType)
            cmd.Parameters.AddWithValue("@active_ctrct", Me.ActiveCtrct)
            cmd.Parameters.AddWithValue("@ctrct_stat", Me.CtrctStat)
            cmd.Parameters.AddWithValue("@excl_srvc_hrs", Me.ExclSrvcHrs)
            cmd.Parameters.AddWithValue("@hrdiv", Me.HRDiv)
            cmd.Parameters.AddWithValue("@hrdept", Me.HRDept)
            cmd.Parameters.AddWithValue("@beg_date", Me.BegDate)
            If Me.EndDate = "" Then
                cmd.Parameters.AddWithValue("@end_date", DBNull.Value)
            Else
                cmd.Parameters.AddWithValue("@end_date", Me.EndDate)
            End If
            cmd.Parameters.AddWithValue("@job_title", Me.JobTitle)
            cmd.Parameters.AddWithValue("@title_rank", Me.TitleRank)

            ' Dim rowsAffected As Integer = cmd.ExecuteNonQuery()   'Executes query and returns a value
            '  Debug.WriteLine("RowsAffected: {0}", rowsAffected.ToString)

            ConnInformix.Close()
            ConnInformix.Dispose()
            WriteLog("Add of Job_Rec completed for " & Me.FullName & ",  ID number " & Me.ID & ", Title = " & Me.sJobTitle & ", HRPay = " & Me.HRPay & ", tPos_No = " & Me.TPosNo & ", Begin Date = " & Me.BegDate)
        Catch ex As DuplicateNameException
            'No need to write anything

        Catch e As Exception
            Debug.WriteLine(e)
            WriteError("Error in clJob_rec - Insert.  Error = " & Err.Description & ", Error Number " & Err.Number)
            'MAIL Event
            'sMl.SendEMessage("ADP to CX", (""Error in clsjob_rec - Insert.  Error = " & e.Message), "dsullivan@carthage.edu", "dsullivan@carthag.edu", "dsullivan@carthage.edu")
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

        sql = "Update job_rec set "
        If Me.Descr <> "" Then sql = sql & " descr = ? "
        If Me.JobTitle <> "" Then sql = sql & ", job_title = ? "
        If Me.HRclass <> "" Then sql = sql & ", hrclass = ? "
        If Me.HRPay <> "" Then sql = sql & ", hrpay = ? "
        If Me.HRStat <> "" Then sql = sql & ", hrstat = ? "
        If Me.HRDiv <> "" Then sql = sql & ", hrdiv = ? "
        If Me.sHRDept <> "" Then sql = sql & ", hrdept = ? "
        sql = sql & ", end_date = ? "
        If Me.ActiveCtrct <> "" Then sql = sql & ", active_ctrct = ? "
        If Me.CtrctStat <> "" Then sql = sql & ", ctrct_stat = ? "
        If Me.ExclSrvcHrs <> "" Then sql = sql & ", excl_srvc_hrs = ? "
        'If Me.TitleRank <> 0 Then
        sql = sql & ", title_rank = ? "
        'ElseIf TitleRank = 0 Then
        'sql = sql & ", title_rank = null "
        'End If
        '        If Me.JobNo <> "" Then sql = sql & ", title_rank = ? "   'JobNo is a sequence, don't update that field
        sql = sql & "where id = ? "
        sql = sql & "and tpos_no = ? "
        sql = sql & "and beg_date = ? "

        Dim cmd As New OdbcCommand(sql, ConnInformix)
        Try

            ConnInformix.ConnectionString = gl.ConnectString
            ConnInformix.Open()

            cmd.CommandType = CommandType.Text

            If Me.Descr <> "" Then cmd.Parameters.AddWithValue("@descr", Me.Descr)
            If Me.JobTitle <> "" Then cmd.Parameters.AddWithValue("@job_title", Me.JobTitle)
            If Me.HRclass <> "" Then cmd.Parameters.AddWithValue("@hrclass", Me.HRclass)
            If Me.HRPay <> "" Then cmd.Parameters.AddWithValue("@hrpay", Me.HRPay)
            If Me.HRStat <> "" Then cmd.Parameters.AddWithValue("@HRStat", Me.HRStat)
            If Me.HRDiv <> "" Then cmd.Parameters.AddWithValue("@hrdiv", Me.HRDiv)
            If Me.sHRDept <> "" Then cmd.Parameters.AddWithValue("@hrdept", Me.HRDept)
            If Me.EndDate = "" Then
                cmd.Parameters.AddWithValue("@end_date", DBNull.Value)
            Else
                cmd.Parameters.AddWithValue("@end_date", Me.EndDate)
            End If
            If Me.ActiveCtrct <> "" Then cmd.Parameters.AddWithValue("@active_ctrct", Me.ActiveCtrct)
            If Me.CtrctStat <> "" Then cmd.Parameters.AddWithValue("@ctrct_stat", Me.CtrctStat)
            If Me.ExclSrvcHrs <> "" Then cmd.Parameters.AddWithValue("@excl_srvc_hrs", Me.ExclSrvcHrs)
            If Me.TitleRank <> 0 Then
                cmd.Parameters.AddWithValue("@title_rank", Me.TitleRank)
            ElseIf Me.TitleRank = 0 Then
                cmd.Parameters.AddWithValue("@title_rank", DbNullOrIntegerValue(Me.TitleRank))
            Else
                cmd.Parameters.AddWithValue("@title_rank", DbNullOrIntegerValue(Me.TitleRank))
            End If
            '            If Me.JobNo <> 0 Then cmd.Parameters.AddWithValue("@job_no", Me.JobNo)
            cmd.Parameters.AddWithValue("@id", Me.ID)
            cmd.Parameters.AddWithValue("@tpos_no", Me.TPosNo)
            cmd.Parameters.AddWithValue("@beg_date", Me.BegDate)
            '#11/30/2017#;

            'Dim rowsAffected As Integer = cmd.ExecuteNonQuery()   'Executes query and returns a value
            'Debug.WriteLine("RowsAffected: {0}", rowsAffected.ToString)

            ConnInformix.Close()
            ConnInformix.Dispose()


            WriteLog("Update of Job_Rec completed for " & Me.FullName & ", ID number " & Me.ID & ", Title = " & Me.sJobTitle & ", HRPay = " & Me.HRPay & ", tPos_No = " & Me.TPosNo & ", Begin Date = " & Me.BegDate)

        Catch e As NullReferenceException
            MsgBox("Null value")
            
        Catch e As Exception
            Debug.WriteLine(e)
            WriteError("Error in clsjob_rec - Update.  Error = " & Err.Description & ", Error Number " & Err.Number)
            'MAIL Event
            'sMl.SendEMessage("ADP to CX", (""Error in clsjob_rec - Update.  Error = " & e.Message), "dsullivan@carthage.edu", "dsullivan@carthag.edu", "dsullivan@carthage.edu")
             Finally
            If ConnInformix.State = ConnectionState.Open Then
                ConnInformix.Close()
                ConnInformix.Dispose()
            End If
            GC.Collect()
        End Try
    End Sub

    Public Function DbNullOrIntegerValue(ByVal value As Integer) As Object
        If value = 0 Then
            Return DBNull.Value
        Else
            Return value
        End If
    End Function






    Public Sub Initialize()

        iJobNo = 0 'Primary Key, sequence added by Informix??
        iTPosNo = 0
        sDescr = "" '
        sHrClass = "" '
        iBJobNo = 0 '
        iID = 0 '
        sHRPay = ""

        iSuprvNo = 0
        sHRStat = ""
        iPseudoSerial = 0
        sEGPType = ""
        sHRDiv = ""
        sHRDept = ""
        dCompBegDate = ""
        dCompEndDate = ""
        dAppInactDate = ""
        dBegDate = ""
        dEndDate = ""
        sActiveCtrct = ""
        sCtrctStat = ""
        sExclSrvcHrs = ""
        sJobTitle = ""
        iTitleRank = 0
        sFullName = ""
    End Sub
End Class
