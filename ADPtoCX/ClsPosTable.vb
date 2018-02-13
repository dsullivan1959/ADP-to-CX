Option Strict Off
Option Explicit On

Imports System.Data.Odbc

Friend Class ClsPosTable

    Dim ITPos As Integer 'This is a serial, don't populate
    Dim iClass_no As Integer
    Dim sPcnAgg As String
    Dim sPcn01 As String
    Dim sPcn02 As String
    Dim sPcn03 As String
    Dim sPcn04 As String
    Dim sDescr As String
    Dim sOfc As String
    Dim iCompPlan As Integer
    Dim sFuncArea As String
    Dim iBpos As Integer
    Dim iSupervisor As Integer
    Dim sEEOcode As String
    Dim stsoc2010 As String
    Dim sTenureTrk As String
    Dim dFTE As Double
    Dim sGrpPos As String
    Dim iMaxJobs As Integer
    Dim sHRPay As String
    Dim sAssign As String
    Dim dActiveDate As Date
    Dim dInActiveDate As Date
    Dim iPseudoSerial As Integer
    Dim sPrevPos As String
    Dim sRank As String
    Dim sHRAdjExCl As String



    Public Property TPos() As Integer
        Get
            TPos = ITPos
        End Get
        Set(ByVal Value As Integer)
            ITPos = Value
        End Set
    End Property

    Public Property Class_no() As Integer
        Get
            Class_no = iClass_no
        End Get
        Set(ByVal Value As Integer)
            iClass_no = Value
        End Set
    End Property


    Public Property PcnAgg() As String
        Get
            PcnAgg = sPcnAgg
        End Get
        Set(ByVal Value As String)
            sPcnAgg = Value
        End Set
    End Property
    Public Property Pcn01() As String
        Get
            Pcn01 = sPcn01
        End Get
        Set(ByVal Value As String)
            sPcn01 = Value
        End Set
    End Property

    Public Property Pcn02() As String
        Get
            Pcn02 = sPcn02
        End Get
        Set(ByVal Value As String)
            sPcn02 = Value
        End Set
    End Property
    Public Property Pcn03() As String
        Get
            Pcn03 = sPcn03
        End Get
        Set(ByVal Value As String)
            sPcn03 = Value
        End Set
    End Property
    Public Property Pcn04() As String
        Get
            Pcn04 = sPcn04
        End Get
        Set(ByVal Value As String)
            sPcn04 = Value
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

    Public Property Ofc() As String
        Get
            Ofc = sOfc
        End Get
        Set(ByVal Value As String)
            sOfc = Value
        End Set
    End Property
    Public Property CompPlan() As Integer
        Get
            CompPlan = iCompPlan
        End Get
        Set(ByVal Value As Integer)
            iCompPlan = Value
        End Set
    End Property

    Public Property FuncArea() As String
        Get
            FuncArea = sFuncArea
        End Get
        Set(ByVal Value As String)
            sFuncArea = Value
        End Set
    End Property
    Public Property Bpos() As Integer
        Get
            Bpos = iBpos
        End Get
        Set(ByVal Value As Integer)
            iBpos = Value
        End Set
    End Property
    Public Property Supervisor() As Integer
        Get
            Supervisor = iSupervisor
        End Get
        Set(ByVal Value As Integer)
            iSupervisor = Value
        End Set
    End Property
    Public Property EEOcode() As String
        Get
            EEOcode = sEEOcode
        End Get
        Set(ByVal Value As String)
            sEEOcode = Value
        End Set
    End Property
    Public Property tsoc2010() As String
        Get
            tsoc2010 = stsoc2010
        End Get
        Set(ByVal Value As String)
            stsoc2010 = Value
        End Set
    End Property
    Public Property TenureTrk() As String
        Get
            TenureTrk = sTenureTrk
        End Get
        Set(ByVal Value As String)
            sTenureTrk = Value
        End Set
    End Property
    Public Property FTE() As Double
        Get
            FTE = dFTE
        End Get
        Set(ByVal Value As Double)
            dFTE = Value
        End Set
    End Property
    Public Property GrpPos() As String
        Get
            GrpPos = sGrpPos
        End Get
        Set(ByVal Value As String)
            sGrpPos = Value
        End Set
    End Property

    Public Property MaxJobs() As Integer
        Get
            MaxJobs = iMaxJobs
        End Get
        Set(ByVal Value As Integer)
            iMaxJobs = Value
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
    Public Property Assign() As String
        Get
            Assign = sAssign
        End Get
        Set(ByVal Value As String)
            sAssign = Value
        End Set
    End Property
    Public Property InActiveDate() As Date
        Get
            InActiveDate = dInActiveDate
        End Get
        Set(ByVal Value As Date)
            dInActiveDate = Value
        End Set
    End Property
    Public Property ActiveDate() As Date
        Get
            ActiveDate = dActiveDate
        End Get
        Set(ByVal Value As Date)
            dActiveDate = Value
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
    Public Property PrevPos() As String
        Get
            PrevPos = sPrevPos
        End Get
        Set(ByVal Value As String)
            sPrevPos = Value
        End Set
    End Property
    Public Property HRAdjExCl() As String
        Get
            HRAdjExCl = sHRAdjExCl
        End Get
        Set(ByVal Value As String)
            sHRAdjExCl = Value
        End Set
    End Property
    Public Property Rank() As String
        Get
            Rank = sRank
        End Get
        Set(ByVal Value As String)
            sRank = Value
        End Set
    End Property

    Public Sub Insert()

        Dim sql As String

        Try
            sql = " INSERT INTO pos_table "  'tpos_no is a sequence
            sql = sql & "(tpos_no, class_no, pcn_aggr, pcn_01, pcn_02, pcn_03, pcn_04, descr, ofc, tcompplan_no, func_area, "
            sql = sql & "bpos_no, supervisor_no, eeo_code, tsoc2010, tenure_track, fte, grp_pos, max_jobs, hrpay, assgn, active_date, "
            sql = sql & "inactive_date, pseudo_serial, prev_pos, rank, hradjexcl)"
            sql = sql & " VALUES(?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)"
          

            Dim cmd As New OdbcCommand(sql, ConnInformix)

            cmd.Parameters.AddWithValue("@class_no", Me.Class_no)
            cmd.Parameters.AddWithValue("@pcn_aggr", Me.PcnAgg)
            cmd.Parameters.AddWithValue("@pcn_01", Me.Pcn01)
            cmd.Parameters.AddWithValue("@pcn_02", Me.Pcn02)
            cmd.Parameters.AddWithValue("@pcn_03", Me.Pcn03)
            cmd.Parameters.AddWithValue("@pcn_04", Me.Pcn04)
            cmd.Parameters.AddWithValue("@descr", Me.Descr)
            cmd.Parameters.AddWithValue("@ofc", Me.Ofc)
            cmd.Parameters.AddWithValue("@tcompplan_no", Me.CompPlan)
            cmd.Parameters.AddWithValue("@func_area", Me.FuncArea)
            cmd.Parameters.AddWithValue("@bpos_no", Me.Bpos)
            cmd.Parameters.AddWithValue("@supervisor_no", Me.Supervisor)
            cmd.Parameters.AddWithValue("@eeo_code", Me.EEOcode)
            cmd.Parameters.AddWithValue("@tsoc2010", Me.tsoc2010)
            cmd.Parameters.AddWithValue("@tenure_track", Me.TenureTrk)
            cmd.Parameters.AddWithValue("@fte", Me.FTE)
            cmd.Parameters.AddWithValue("@grp_pos", Me.GrpPos)
            cmd.Parameters.AddWithValue("@max_jobs", Me.MaxJobs)
            cmd.Parameters.AddWithValue("@hrpay", Me.HRPay)
            cmd.Parameters.AddWithValue("@assgn", Me.Assign)
            cmd.Parameters.AddWithValue("@active_date", Me.ActiveDate)
            cmd.Parameters.AddWithValue("@inactive_date", Me.InActiveDate)
            cmd.Parameters.AddWithValue("@pseudo_serial", Me.PseudoSerial)
            cmd.Parameters.AddWithValue("@prev_pos", Me.PrevPos)

            cmd.Parameters.AddWithValue("@rank", Me.Rank)
            cmd.Parameters.AddWithValue("@hradjexcl", Me.HRAdjExCl)

            ConnInformix.Open()

            '            Dim rowsAffected As Integer = cmd.ExecuteNonQuery()
            ConnInformix.Close()
            ConnInformix.Dispose()
            WriteLog("Insert into pos_table, TPOS_NO = " & Me.TPos & ", Descr = " & Me.Descr & " PCN_AGGR = " & Me.PcnAgg)

        Catch e As Exception
            Debug.WriteLine(e)
            WriteError("Error in clsPos_tabel - Insert.  Error = " & Err.Description & ", Error Number " & Err.Number & " TPOS_NO = " & Me.TPos & ", Descr = " & Me.Descr & " PCN_AGGR = " & Me.PcnAgg)
            'MAIL Event
            'sMl.SendEMessage("ADP to CX", (""Error in clsPos_rec - Insert.  Error = " & e.Message), "dsullivan@carthage.edu", "dsullivan@carthag.edu", "dsullivan@carthage.edu")
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
        Try


          

            sql = "Update pos_rec set "
            If Me.Class_no <> "" Then sql = sql & " class_no = ?"
            If Me.PcnAgg <> "" Then sql = sql & ", pcn_aggr = ?"
            If Me.Pcn01 <> "" Then sql = sql & ", pcn_01 = ?"
            If Me.Pcn02 <> "" Then sql = sql & ", pcn_02 = ?"
            If Me.Pcn03 <> "" Then sql = sql & ", pcn_03 = ?"
            If Me.Pcn04 <> "" Then sql = sql & ", pcn_04 = ?"

            If Me.Descr <> "" Then sql = sql & ", descr = ?"
            If Me.Ofc <> "" Then sql = sql & ", ofc = ?"
            If Me.CompPlan <> "" Then sql = sql & ", tcompplan_no = ?"
            If Me.FuncArea <> "" Then sql = sql & ", func_area = ?"
            If Me.Bpos <> "" Then sql = sql & ", bpos_no = ?"
            If Me.Supervisor <> "" Then sql = sql & ", supervisor_no = ?"
            If Me.EEOcode <> "" Then sql = sql & ", eeo_code = ?"
            If Me.tsoc2010 <> "" Then sql = sql & ", tsoc2010 = ?"
            If Me.TenureTrk <> "" Then sql = sql & ", tenure_track = ?"
            If Me.FTE <> "" Then sql = sql & ", fte = ?"
            If Me.GrpPos <> "" Then sql = sql & ", grp_pos = ?"
            If Me.MaxJobs <> "" Then sql = sql & ", max_jobs = ?"
            If Me.HRPay <> "" Then sql = sql & ", hrpay = ?"
            If Me.Assign <> "" Then sql = sql & ", assgn = ?"
            If Me.ActiveDate <> "" Then sql = sql & ", active_date = ?"
            If Me.InActiveDate <> "" Then sql = sql & ", inactive_date = ?"
            If Me.PseudoSerial <> "" Then sql = sql & ", pseudo_serial = ?"
            If Me.PrevPos <> "" Then sql = sql & ", prev_pos = ?"
            If Me.Rank <> "" Then sql = sql & ", rank = ?"
            If Me.HRAdjExCl <> "" Then sql = sql & ", hradjexcl = ?"
            sql = sql & " WHERE tpos_no = ?"

            Dim cmd As New OdbcCommand(sql, ConnInformix)

            If Me.Class_no <> "" Then cmd.Parameters.AddWithValue("@class_no", Me.Class_no)
            If Me.PcnAgg <> "" Then cmd.Parameters.AddWithValue("@pcn_aggr", Me.PcnAgg)
            If Me.Pcn01 <> "" Then cmd.Parameters.AddWithValue("@pcn_01", Me.Pcn01)
            If Me.Pcn02 <> "" Then cmd.Parameters.AddWithValue("@pcn_02", Me.Pcn02)
            If Me.Pcn03 <> "" Then cmd.Parameters.AddWithValue("@pcn_03", Me.Pcn03)
            If Me.Pcn04 <> "" Then cmd.Parameters.AddWithValue("@pcn_04", Me.Pcn04)
            If Me.Descr <> "" Then cmd.Parameters.AddWithValue("@descr", Me.Descr)
            If Me.Ofc <> "" Then cmd.Parameters.AddWithValue("@ofc", Me.Ofc)
            If Me.CompPlan <> "" Then cmd.Parameters.AddWithValue("@tcompplan_no", Me.CompPlan)
            If Me.FuncArea <> "" Then cmd.Parameters.AddWithValue("@func_area", Me.FuncArea)
            If Me.Bpos <> "" Then cmd.Parameters.AddWithValue("@bpos_no", Me.Bpos)
            If Me.Supervisor <> "" Then cmd.Parameters.AddWithValue("@supervisor_no", Me.Supervisor)
            If Me.EEOcode <> "" Then cmd.Parameters.AddWithValue("@eeo_code", Me.EEOcode)
            If Me.tsoc2010 <> "" Then cmd.Parameters.AddWithValue("@tsoc2010", Me.tsoc2010)
            If Me.TenureTrk <> "" Then cmd.Parameters.AddWithValue("@tenure_track", Me.TenureTrk)
            If Me.FTE <> "" Then cmd.Parameters.AddWithValue("@fte", Me.FTE)
            If Me.GrpPos <> "" Then cmd.Parameters.AddWithValue("@grp_pos", Me.GrpPos)
            If Me.MaxJobs <> "" Then cmd.Parameters.AddWithValue("@max_jobs", Me.MaxJobs)
            If Me.HRPay <> "" Then cmd.Parameters.AddWithValue("@hrpay", Me.HRPay)
            If Me.Assign <> "" Then cmd.Parameters.AddWithValue("@assgn", Me.Assign)
            If Me.ActiveDate <> "" Then cmd.Parameters.AddWithValue("@active_date", Me.ActiveDate)
            If Me.InActiveDate <> "" Then cmd.Parameters.AddWithValue("@inactive_date", Me.InActiveDate)
            If Me.PseudoSerial <> "" Then cmd.Parameters.AddWithValue("@pseudo_serial", Me.PseudoSerial)
            If Me.PrevPos <> "" Then cmd.Parameters.AddWithValue("@prev_pos", Me.PrevPos)
            If Me.Rank <> "" Then cmd.Parameters.AddWithValue("@rank", Me.Rank)
            If Me.HRAdjExCl <> "" Then cmd.Parameters.AddWithValue("@hradjexcl", Me.HRAdjExCl)
            cmd.Parameters.AddWithValue("@tpos_no", Me.TPos)

            '
            ConnInformix.ConnectionString = gl.connectstring
            ConnInformix.Open()
            'Dim rowsAffected As Integer = cmd.ExecuteNonQuery()
            'ConnInformix.Close()
            WriteLog("Update pos_table, TPOS_NO = " & Me.TPos & ", Descr = " & Me.Descr & " PCN_AGGR = " & Me.PcnAgg & " TPOS_NO = " & Me.TPos & ", Descr = " & Me.Descr & " PCN_AGGR = " & Me.PcnAgg)
            ConnInformix.Close()
            ConnInformix.Dispose()

        Catch e As NullReferenceException
            MsgBox("Null value")
        Catch ex As DuplicateNameException
            'No need to write anything
        Catch e As Exception
            Debug.WriteLine(e)
            WriteLog("Update pos_table, TPOS_NO = " & Me.TPos & ", Descr = " & Me.Descr & " PCN_AGGR = " & Me.PcnAgg & " TPOS_NO = " & Me.TPos & ", Descr = " & Me.Descr & " PCN_AGGR = " & Me.PcnAgg)
            'MAIL Event
            'sMl.SendEMessage("ADP to CX", (""Error in clsPos_rec - Update.  Error = " & e.Message), "dsullivan@carthage.edu", "dsullivan@carthag.edu", "dsullivan@carthage.edu")
        Finally
            If ConnInformix.State = ConnectionState.Open Then
                ConnInformix.Close()
                ConnInformix.Dispose()
            End If
            GC.Collect()
        End Try
    End Sub



    Public Sub UpdateFunc()

        Dim sql As String
        Try


            sql = "Update pos_table set "

            If Me.FuncArea <> "" Then sql = sql & " func = ?"
            If Me.TPos <> 0 Then sql = sql & " WHERE [tpos_no] = ?"

            Dim cmd As New OdbcCommand(sql, ConnInformix)

            If Me.FuncArea <> "" Then cmd.Parameters.AddWithValue("@func_area", Me.FuncArea)
            If Me.TPos <> 0 Then cmd.Parameters.AddWithValue("@tpos_no", Me.TPos)
            '
            ConnInformix.ConnectionString = gl.ConnectString
            'ConnInformix.Open()
            'Dim rowsAffected As Integer = cmd.ExecuteNonQuery()
            'ConnInformix.Close()
            WriteLog("Update of func_area record for Position " & Me.TPos & ", Descr = " & Me.Descr & ", Function Code = " & Me.FuncArea & ", PCN = " & Me.PcnAgg)
            ConnInformix.Close()
            ConnInformix.Dispose()

        Catch e As NullReferenceException
            MsgBox("Null value")
        Catch ex As DuplicateNameException
            'No need to write anything
        Catch e As Exception
            Debug.WriteLine(e)
            WriteError("Error in clsPos_rec - Update.  Error = " & Err.Description & ", Error Number " & Err.Number)
            'MAIL Event
            'sMl.SendEMessage("ADP to CX", (""Error in clsPos_rec - Update.  Error = " & e.Message), "dsullivan@carthage.edu", "dsullivan@carthag.edu", "dsullivan@carthage.edu")
        Finally
            If ConnInformix.State = ConnectionState.Open Then
                ConnInformix.Close()
                ConnInformix.Dispose()
            End If
            GC.Collect()
        End Try
    End Sub





   





    Public Sub Initialize()

        ITPos = 0
        iClass_no = 0
        sPcnAgg = ""
        sPcn01 = ""
        sPcn02 = ""
        sPcn03 = ""
        sPcn04 = ""
        sDescr = ""
        sOfc = ""
        iCompPlan = 0

        sFuncArea = ""
        iBpos = 0
        iSupervisor = 0
        sEEOcode = ""
        stsoc2010 = ""
        sTenureTrk = ""
        dFTE = 0
        sGrpPos = ""
        iMaxJobs = 0
        sHRPay = ""
        sAssign = ""
        dActiveDate = "01/01/1999"
        dInActiveDate = "01/01/1999"
        iPseudoSerial = 0
        sPrevPos = ""
        sRank = ""
        sHRAdjExCl = ""



    End Sub



End Class
