Option Strict Off
Option Explicit On

Imports System.Data.Odbc

Public Class clsProvision
    Dim sFileNum As String
    Dim sCarthID = ""
    Dim sLName As String
    Dim sFName As String
    Dim sMidName As String
    Dim sSalut As String
    Dim sPayrollName As String
    Dim sPrefName As String
    Dim dBDate As String
    Dim sGender As String
    Dim sMrtStat As String
    Dim sRaceCode As String
    Dim sRaceDescr As String
    Dim sEthn As String
    Dim sEthRaceIDMthd As String
    Dim sPersEmail As String
    Dim sPrimAddr1 As String
    Dim sPrimAddr2 As String
    Dim sPrimAddr3 As String
    Dim sPrimCity As String
    Dim sPrimState As String
    Dim sPrimStateDesc As String
    Dim sPrimZip As String
    Dim sPrimCnty As String
    Dim sPrimCntry As String
    Dim sPrimCntryCode As String
    Dim sUseAsLegal As String

    Dim sHomePhone As String
    Dim sPersMobile As String
    Dim sWorkPhone As String
    Dim sWorkContPhone As String
    Dim sWorkEmail As String
    Dim sUseWrkEmlCont As String

    Dim sLegalAddr1 As String
    Dim sLegalAddr2 As String
    Dim sLegalAddr3 As String
    Dim sLegalCity As String
    Dim sLegalState As String
    Dim sLegalStateDesc As String
    Dim sLegalZip As String
    Dim sLegalCnty As String
    Dim sLegalCntry As String
    Dim sLegalCntryCode As String

    Dim sSSN As String
    Dim dHireDt = ""
    Dim dHireRehireDt = ""
    Dim dRehireDt = ""
    Dim dPosStart = ""
    Dim dPosEffDt = ""
    Dim dPosEffEnd = ""
    Dim dTermDate = ""
    Dim sPosStatus As String
    Dim dStatEffDt = ""
    Dim dStatEffEndDt = ""
    Dim dAdjServDt = ""
    Dim sArchived As String

    Dim sPostnID As String
    Dim sPrimPos As String
    Dim sPayCompCod As String
    Dim sPayCompName As String
    Dim sCIPCode As String
    Dim sWorkCatCode As String
    Dim sWorkCatDescr As String
    Dim sJobTtlCode As String
    Dim sJobTtlDescr As String
    Dim sHomeCostCode As String
    Dim sHomeCostDescr As String
    Dim sJobClassCode As String
    Dim sJobClassDescr As String
    Dim sJobDescr As String
    Dim sJobFuncCode As String
    Dim sJobFuncDescr As String
    Dim sRoomNmbr As String
    Dim sLocCode As String
    Dim sLocDescr As String
    Dim dLeaveStrDt As String
    Dim dLeaveRetDt As String

    Dim sHomeCostNum2 As String
    Dim sPayCompCode2 As String
    Dim dPosEffDt2 As String
    Dim dPosEffEndDt2 As String

    Dim sHomeCostNum3 As String
    Dim sPayCompCode3 As String
    Dim dPosEffDt3 = ""
    Dim dPosEffEndDt3 = ""

    Dim sHomeCostNum4 As String
    Dim sPayCompCode4 As String
    Dim dPosEffDt4 = ""
    Dim dPosEffEndDt4 = ""

    Dim sHomeDptNumCode As String
    Dim sHomeDptDescr As String
    Dim sSuprvID As String
    Dim sSuprvFName As String
    Dim sSuprvLName As String
    Dim dDateStamp As String





    Public Property FileNum() As String
        Get
            FileNum = sFileNum
        End Get
        Set(ByVal Value As String)
            sFileNum = Value
        End Set
    End Property
    Public Property CarthID() As String
        Get
            CarthID = sCarthID
        End Get
        Set(ByVal Value As String)
            sCarthID = Value
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

    Public Property FName() As String
        Get
            FName = sFName
        End Get
        Set(ByVal Value As String)
            sFName = Value
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

    Public Property Salut() As String
        Get
            Salut = sSalut
        End Get
        Set(ByVal Value As String)
            sSalut = Value
        End Set
    End Property

    Public Property PayrollName() As String
        Get
            PayrollName = sPayrollName
        End Get
        Set(ByVal Value As String)
            sPayrollName = Value
        End Set
    End Property

    Public Property PrefName() As String
        Get
            PrefName = sPrefName
        End Get
        Set(ByVal Value As String)
            sPrefName = Value
        End Set
    End Property

    Public Property BDate() As String
        Get
            BDate = dBDate
        End Get
        Set(ByVal Value As String)
            dBDate = Value
        End Set
    End Property

    Public Property Gender() As String
        Get
            Gender = sGender
        End Get
        Set(ByVal Value As String)
            sGender = Value
        End Set
    End Property


    Public Property MrtStat() As String
        Get
            MrtStat = sMrtStat
        End Get
        Set(ByVal Value As String)
            sMrtStat = Value
        End Set
    End Property

    Public Property RaceCode() As String
        Get
            RaceCode = sRaceCode
        End Get
        Set(ByVal Value As String)
            sRaceCode = Value
        End Set
    End Property

    Public Property RaceDescr() As String
        Get
            RaceDescr = sRaceDescr
        End Get
        Set(ByVal Value As String)
            sRaceDescr = Value
        End Set
    End Property

    Public Property Ethn() As String
        Get
            Ethn = sEthn
        End Get
        Set(ByVal Value As String)
            sEthn = Value
        End Set
    End Property

    Public Property EthRaceIDMthd() As String
        Get
            EthRaceIDMthd = sEthRaceIDMthd
        End Get
        Set(ByVal Value As String)
            sEthRaceIDMthd = Value
        End Set
    End Property

    Public Property PersEmail() As String
        Get
            PersEmail = sPersEmail
        End Get
        Set(ByVal Value As String)
            sPersEmail = Value
        End Set
    End Property

    Public Property PrimAddr1() As String
        Get
            PrimAddr1 = sPrimAddr1
        End Get
        Set(ByVal Value As String)
            sPrimAddr1 = Value
        End Set
    End Property

    Public Property PrimAddr2() As String
        Get
            PrimAddr2 = sPrimAddr2
        End Get
        Set(ByVal Value As String)
            sPrimAddr2 = Value
        End Set
    End Property
    Public Property PrimAddr3() As String
        Get
            PrimAddr3 = sPrimAddr3
        End Get
        Set(ByVal Value As String)
            sPrimAddr3 = Value
        End Set
    End Property


    Public Property PrimCity() As String
        Get
            PrimCity = sPrimCity
        End Get
        Set(ByVal Value As String)
            sPrimCity = Value
        End Set
    End Property

    Public Property PrimState() As String
        Get
            PrimState = sPrimState
        End Get
        Set(ByVal Value As String)
            sPrimState = Value
        End Set
    End Property

    Public Property PrimStateDesc() As String
        Get
            PrimStateDesc = sPrimStateDesc
        End Get
        Set(ByVal Value As String)
            sPrimStateDesc = Value
        End Set
    End Property
    Public Property PrimZip() As String
        Get
            PrimZip = sPrimZip
        End Get
        Set(ByVal Value As String)
            sPrimZip = Value
        End Set
    End Property
    Public Property PrimCnty() As String
        Get
            PrimCnty = sPrimCnty
        End Get
        Set(ByVal Value As String)
            sPrimCnty = Value
        End Set
    End Property

    Public Property PrimCntry() As String
        Get
            PrimCntry = sPrimCntry
        End Get
        Set(ByVal Value As String)
            sPrimCntry = Value
        End Set
    End Property
    Public Property PrimCntryCode() As String
        Get
            PrimCntryCode = sPrimCntryCode
        End Get
        Set(ByVal Value As String)
            sPrimCntryCode = Value
        End Set
    End Property

    Public Property UseAsLegal() As String
        Get
            UseAsLegal = sUseAsLegal
        End Get
        Set(ByVal Value As String)
            sUseAsLegal = Value
        End Set
    End Property

    Public Property HomePhone() As String
        Get
            HomePhone = sHomePhone
        End Get
        Set(ByVal Value As String)
            sHomePhone = Value
        End Set
    End Property
    Public Property PersMobile() As String
        Get
            PersMobile = sPersMobile
        End Get
        Set(ByVal Value As String)
            sPersMobile = Value
        End Set
    End Property
    Public Property WorkPhone() As String
        Get
            WorkPhone = sWorkPhone
        End Get
        Set(ByVal Value As String)
            sWorkPhone = Value
        End Set
    End Property
    Public Property WorkContPhone() As String
        Get
            WorkContPhone = sWorkContPhone
        End Get
        Set(ByVal Value As String)
            sWorkContPhone = Value
        End Set
    End Property

    Public Property WorkEmail() As String
        Get
            WorkEmail = sWorkEmail
        End Get
        Set(ByVal Value As String)
            sWorkEmail = Value
        End Set
    End Property
    Public Property UseWrkEmlCont() As String
        Get
            UseWrkEmlCont = sUseWrkEmlCont
        End Get
        Set(ByVal Value As String)
            sUseWrkEmlCont = Value
        End Set
    End Property

    Public Property LegalAddr1() As String
        Get
            LegalAddr1 = sLegalAddr1
        End Get
        Set(ByVal Value As String)
            sLegalAddr1 = Value
        End Set
    End Property
    Public Property LegalAddr2() As String
        Get
            LegalAddr2 = sLegalAddr2
        End Get
        Set(ByVal Value As String)
            sLegalAddr2 = Value
        End Set
    End Property
    Public Property LegalAddr3() As String
        Get
            LegalAddr3 = sLegalAddr3
        End Get
        Set(ByVal Value As String)
            sLegalAddr3 = Value
        End Set
    End Property
    Public Property LegalCity() As String
        Get
            LegalCity = sLegalCity
        End Get
        Set(ByVal Value As String)
            sLegalCity = Value
        End Set
    End Property
    Public Property LegalState() As String
        Get
            LegalState = sLegalState
        End Get
        Set(ByVal Value As String)
            sLegalState = Value
        End Set
    End Property
    Public Property LegalStateDesc() As String
        Get
            LegalStateDesc = sLegalStateDesc
        End Get
        Set(ByVal Value As String)
            sLegalStateDesc = Value
        End Set
    End Property
    Public Property LegalZip() As String
        Get
            LegalZip = sLegalZip
        End Get
        Set(ByVal Value As String)
            sLegalZip = Value
        End Set
    End Property
    Public Property LegalCnty() As String
        Get
            LegalCnty = sLegalCnty
        End Get
        Set(ByVal Value As String)
            sLegalCnty = Value
        End Set
    End Property
    Public Property LegalCntry() As String
        Get
            LegalCntry = sLegalCntry
        End Get
        Set(ByVal Value As String)
            sLegalCntry = Value
        End Set
    End Property
    Public Property LegalCntryCode() As String
        Get
            LegalCntryCode = sLegalCntryCode
        End Get
        Set(ByVal Value As String)
            sLegalCntryCode = Value
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

    Public Property HireDt() As String
        Get
            HireDt = dHireDt
        End Get
        Set(ByVal Value As String)
            dHireDt = Value
        End Set
    End Property
    Public Property HireReHireDt() As String
        Get
            HireReHireDt = dHireRehireDt
        End Get
        Set(ByVal Value As String)
            dHireRehireDt = Value
        End Set
    End Property
    Public Property ReHireDt() As String
        Get
            ReHireDt = dRehireDt
        End Get
        Set(ByVal Value As String)
            dRehireDt = Value
        End Set
    End Property
    Public Property PosStart() As String
        Get
            PosStart = dPosStart
        End Get
        Set(ByVal Value As String)
            dPosStart = Value
        End Set
    End Property
    Public Property PosEffDt() As String
        Get
            PosEffDt = dPosEffDt
        End Get
        Set(ByVal Value As String)
            dPosEffDt = Value
        End Set
    End Property
    Public Property PosEffEnd() As String
        Get
            PosEffEnd = dPosEffEnd
        End Get
        Set(ByVal Value As String)
            dPosEffEnd = Value
        End Set
    End Property
    Public Property TermDate() As String
        Get
            TermDate = dTermDate
        End Get
        Set(ByVal Value As String)
            dTermDate = Value
        End Set
    End Property
    Public Property PosStatus() As String
        Get
            PosStatus = sPosStatus
        End Get
        Set(ByVal Value As String)
            sPosStatus = Value
        End Set
    End Property
    Public Property StatEffDt() As String
        Get
            StatEffDt = dStatEffDt
        End Get
        Set(ByVal Value As String)
            dStatEffDt = Value
        End Set
    End Property
    Public Property StatEffEndDt() As String
        Get
            StatEffEndDt = dStatEffEndDt
        End Get
        Set(ByVal Value As String)
            dStatEffEndDt = Value
        End Set
    End Property
    Public Property AdjServDt() As String
        Get
            AdjServDt = dAdjServDt
        End Get
        Set(ByVal Value As String)
            dAdjServDt = Value
        End Set
    End Property
    Public Property Archived() As String
        Get
            Archived = sArchived
        End Get
        Set(ByVal Value As String)
            sArchived = Value
        End Set
    End Property
    Public Property PostnID() As String
        Get
            PostnID = sPostnID
        End Get
        Set(ByVal Value As String)
            sPostnID = Value
        End Set
    End Property
    Public Property PrimPos() As String
        Get
            PrimPos = sPrimPos
        End Get
        Set(ByVal Value As String)
            sPrimPos = Value
        End Set
    End Property
    Public Property PayCompCod() As String
        Get
            PayCompCod = sPayCompCod
        End Get
        Set(ByVal Value As String)
            sPayCompCod = Value
        End Set
    End Property
    Public Property PayCompName() As String
        Get
            PayCompName = sPayCompName
        End Get
        Set(ByVal Value As String)
            sPayCompName = Value
        End Set
    End Property
    Public Property CipCode() As String
        Get
            CipCode = sCIPCode
        End Get
        Set(ByVal Value As String)
            sCIPCode = Value
        End Set
    End Property
    Public Property WorkCatCode() As String
        Get
            WorkCatCode = sWorkCatCode
        End Get
        Set(ByVal Value As String)
            sWorkCatCode = Value
        End Set
    End Property
    Public Property WorkCatDescr() As String
        Get
            WorkCatDescr = sWorkCatDescr
        End Get
        Set(ByVal Value As String)
            sWorkCatDescr = Value
        End Set
    End Property
    Public Property JobTtlCode() As String
        Get
            JobTtlCode = sJobTtlCode
        End Get
        Set(ByVal Value As String)
            sJobTtlCode = Value
        End Set
    End Property
    Public Property JobTtlDescr() As String
        Get
            JobTtlDescr = sJobTtlDescr
        End Get
        Set(ByVal Value As String)
            sJobTtlDescr = Value
        End Set
    End Property
    Public Property HomeCostCode() As String
        Get
            HomeCostCode = sHomeCostCode
        End Get
        Set(ByVal Value As String)
            sHomeCostCode = Value
        End Set
    End Property
    Public Property HomeCostDescr() As String
        Get
            HomeCostDescr = sHomeCostDescr
        End Get
        Set(ByVal Value As String)
            sHomeCostDescr = Value
        End Set
    End Property
    Public Property JobClassCode() As String
        Get
            JobClassCode = sJobClassCode
        End Get
        Set(ByVal Value As String)
            sJobClassCode = Value
        End Set
    End Property
    Public Property JobClassDescr() As String
        Get
            JobClassDescr = sJobClassDescr
        End Get
        Set(ByVal Value As String)
            sJobClassDescr = Value
        End Set
    End Property
    Public Property JobDescr() As String
        Get
            JobDescr = sJobDescr
        End Get
        Set(ByVal Value As String)
            sJobDescr = Value
        End Set
    End Property
    Public Property JobFuncCode() As String
        Get
            JobFuncCode = sJobFuncCode
        End Get
        Set(ByVal Value As String)
            sJobFuncCode = Value
        End Set
    End Property

    Public Property JobFuncDescr() As String
        Get
            JobFuncDescr = sJobFuncDescr
        End Get
        Set(ByVal Value As String)
            sJobFuncDescr = Value
        End Set
    End Property

    Public Property RoomNmbr() As String
        Get
            RoomNmbr = sRoomNmbr
        End Get
        Set(ByVal Value As String)
            sRoomNmbr = Value
        End Set
    End Property
    Public Property LocCode() As String
        Get
            LocCode = sLocCode
        End Get
        Set(ByVal Value As String)
            sLocCode = Value
        End Set
    End Property
    Public Property LocDescr() As String
        Get
            LocDescr = sLocDescr
        End Get
        Set(ByVal Value As String)
            sLocDescr = Value
        End Set
    End Property
    Public Property LeaveStrDt() As String
        Get
            LeaveStrDt = dLeaveStrDt
        End Get
        Set(ByVal Value As String)
            dLeaveStrDt = Value
        End Set
    End Property
    Public Property LeaveRetDt() As String
        Get
            LeaveRetDt = dLeaveRetDt
        End Get
        Set(ByVal Value As String)
            dLeaveRetDt = Value
        End Set
    End Property

    Public Property HomeCostNum2() As String
        Get
            HomeCostNum2 = sHomeCostNum2
        End Get
        Set(ByVal Value As String)
            sHomeCostNum2 = Value
        End Set
    End Property
    Public Property PayCompCode2() As String
        Get
            PayCompCode2 = sPayCompCode2
        End Get
        Set(ByVal Value As String)
            sPayCompCode2 = Value
        End Set
    End Property
    Public Property PosEffDt2() As String
        Get
            PosEffDt2 = dPosEffDt2
        End Get
        Set(ByVal Value As String)
            dPosEffDt2 = Value
        End Set
    End Property
    Public Property PosEffEndDt2() As String
        Get
            PosEffEndDt2 = dPosEffEndDt2
        End Get
        Set(ByVal Value As String)
            dPosEffEndDt2 = Value
        End Set
    End Property
    Public Property HomeCostNum3() As String
        Get
            HomeCostNum3 = sHomeCostNum3
        End Get
        Set(ByVal Value As String)
            sHomeCostNum3 = Value
        End Set
    End Property
    Public Property PayCompCode3() As String
        Get
            PayCompCode3 = sPayCompCode3
        End Get
        Set(ByVal Value As String)
            sPayCompCode3 = Value
        End Set
    End Property
    Public Property PosEffDt3() As String
        Get
            PosEffDt3 = dPosEffDt3
        End Get
        Set(ByVal Value As String)
            dPosEffDt3 = Value
        End Set
    End Property
    Public Property PosEffEndDt3() As String
        Get
            PosEffEndDt3 = dPosEffEndDt3
        End Get
        Set(ByVal Value As String)
            dPosEffEndDt3 = Value
        End Set
    End Property


    Public Property HomeCostNum4() As String
        Get
            HomeCostNum4 = sHomeCostNum4
        End Get
        Set(ByVal Value As String)
            sHomeCostNum4 = Value
        End Set
    End Property
    Public Property PayCompCode4() As String
        Get
            PayCompCode4 = sPayCompCode4
        End Get
        Set(ByVal Value As String)
            sPayCompCode4 = Value
        End Set
    End Property
    Public Property PosEffDt4() As String
        Get
            PosEffDt4 = dPosEffDt4
        End Get
        Set(ByVal Value As String)
            dPosEffDt4 = Value
        End Set
    End Property
    Public Property PosEffEndDt4() As String
        Get
            PosEffEndDt4 = dPosEffEndDt4
        End Get
        Set(ByVal Value As String)
            dPosEffEndDt4 = Value
        End Set
    End Property

    Public Property HomeDptNumCode() As String
        Get
            HomeDptNumCode = sHomeDptNumCode
        End Get
        Set(ByVal Value As String)
            sHomeDptNumCode = Value
        End Set
    End Property
    Public Property HomeDptDescr() As String
        Get
            HomeDptDescr = sHomeDptDescr
        End Get
        Set(ByVal Value As String)
            sHomeDptDescr = Value
        End Set
    End Property
    Public Property SuprvID() As String
        Get
            SuprvID = sSuprvID
        End Get
        Set(ByVal Value As String)
            sSuprvID = Value
        End Set
    End Property
    Public Property SuprvFName() As String
        Get
            SuprvFName = sSuprvFName
        End Get
        Set(ByVal Value As String)
            sSuprvFName = Value
        End Set
    End Property
    Public Property SuprvLName() As String
        Get
            SuprvLName = sSuprvLName
        End Get
        Set(ByVal Value As String)
            sSuprvLName = Value
        End Set
    End Property
    Public Property Date_Stamp() As String
        Get
            Date_Stamp = dDateStamp
        End Get
        Set(ByVal Value As String)
            dDateStamp = Value
        End Set
    End Property





    Public Sub Insert()

        Dim sql As String
        sql = "INSERT INTO cc_adp_rec (file_no, carthage_id, lastname, firstname, middlename, salutation, fullname, pref_name,"
        sql = sql & " birth_date, gender, marital_status, race, race_descr, hispanic, race_id_method, personal_email, primary_addr_line1,"
        sql = sql & " primary_addr_line2, primary_addr_line3, primary_addr_city, primary_addr_st, primary_addr_state, primary_addr_zip,"
        sql = sql & " primary_addr_county, primary_addr_country, primary_addr_country_code, primary_addr_as_legal, home_phone, cell_phone,"
        sql = sql & " work_phone, work_contact_phone, work_contact_email, work_contact_notification, legal_addr_line1, legal_addr_line2,"
        sql = sql & " legal_addr_line3, legal_addr_city, legal_addr_st, legal_addr_state, legal_addr_zip, legal_addr_county, legal_addr_country,"
        sql = sql & " legal_addr_country_code, ssn, hire_date, hire_rehire_date, rehire_date, position_start_date, position_effective_date,"
        sql = sql & " position_effective_end_date, termination_date, position_status, status_effective_date, status_effective_end_date,"
        sql = sql & " adjusted_service_date, archived_employee, position_id, primary_position, payroll_company_code, payroll_company_name,"
        sql = sql & " cip_code, worker_category_code, worker_category_descr, job_title_code, job_title_descr, home_cost_number_code,"
        sql = sql & " home_cost_number_descr, job_class_code, job_class_descr, job_descr, job_function_code, job_function_descr, room,"
        sql = sql & " bldg, bldg_name, leave_of_absence_start_date, leave_of_absence_return_date, home_cost_number_2, payroll_company_code_2,"
        sql = sql & " position_effective_date_2, position_end_date_2, home_cost_number_3, payroll_company_code_3, position_effective_date_3,"
        sql = sql & " position_end_date_3, home_cost_number_4, payroll_company_code_4, position_effective_date_4, position_end_date_4,"
        sql = sql & " home_depart_num_code, home_depart_num_descr, supervisor_id, supervisor_firstname, supervisor_lastname, date_stamp)"
        sql = sql & " values (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,"
        sql = sql & "?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?) "
        'Add one more ? for datestamp when table structure is updated
        Dim cmd As New OdbcCommand(sql, ConnInformix)

        Try

            ConnInformix.ConnectionString = gl.ConnectString
            ConnInformix.Open()

            cmd.CommandType = CommandType.Text
            cmd.Parameters.AddWithValue("@file_no", Me.FileNum)
            cmd.Parameters.AddWithValue("@carthage_id", Me.CarthID)
            cmd.Parameters.AddWithValue("@lastname", Me.LName)
            cmd.Parameters.AddWithValue("@firstname", Me.FName)
            cmd.Parameters.AddWithValue("@middlename", Me.MidName)
            cmd.Parameters.AddWithValue("@salutation", Me.Salut)
            cmd.Parameters.AddWithValue("@fullname", Me.PayrollName)
            cmd.Parameters.AddWithValue("@pref_name", Me.PrefName)
            cmd.Parameters.AddWithValue("@birth_date", Me.BDate)
            cmd.Parameters.AddWithValue("@gender", Me.Gender)
            cmd.Parameters.AddWithValue("@marital_status", Me.MrtStat)
            cmd.Parameters.AddWithValue("@race", Me.RaceCode)
            cmd.Parameters.AddWithValue("@race_descr", Me.RaceDescr)
            cmd.Parameters.AddWithValue("@hispanic", Me.Ethn)
            cmd.Parameters.AddWithValue("@race_id_method", Me.EthRaceIDMthd)
            cmd.Parameters.AddWithValue("@personal_email", Me.PersEmail)
            cmd.Parameters.AddWithValue("@primary_addr_line1", Me.PrimAddr1)
            cmd.Parameters.AddWithValue("@primary_addr_line2", Me.PrimAddr2)
            cmd.Parameters.AddWithValue("@primary_addr_line3", Me.PrimAddr3)
            cmd.Parameters.AddWithValue("@primary_addr_city", Me.PrimCity)
            cmd.Parameters.AddWithValue("@primary_addr_st", Me.PrimState)
            cmd.Parameters.AddWithValue("@primary_addr_state", Me.PrimStateDesc)
            cmd.Parameters.AddWithValue("@primary_addr_zip", Me.PrimZip)
            cmd.Parameters.AddWithValue("@primary_addr_county", Me.PrimCnty)
            cmd.Parameters.AddWithValue("@primary_addr_country", Me.PrimCntry)
            cmd.Parameters.AddWithValue("@primary_addr_country_code", Me.PrimCntryCode)
            cmd.Parameters.AddWithValue("@primary_addr_as_legal", Me.UseAsLegal)
            cmd.Parameters.AddWithValue("@home_phone", Me.HomePhone)
            cmd.Parameters.AddWithValue("@cell_phone", Me.PersMobile)

            cmd.Parameters.AddWithValue("@work_phone", Me.WorkPhone)
            cmd.Parameters.AddWithValue("@work_contact_phone", Me.WorkContPhone)
            cmd.Parameters.AddWithValue("@work_contact_email", Me.WorkEmail)
            cmd.Parameters.AddWithValue("@work_contact_notification", Me.UseWrkEmlCont)
            cmd.Parameters.AddWithValue("@legal_addr_line1", Me.LegalAddr1)
            cmd.Parameters.AddWithValue("@legal_addr_line2", Me.LegalAddr2)
            cmd.Parameters.AddWithValue("@legal_addr_line3", Me.LegalAddr3)
            cmd.Parameters.AddWithValue("@legal_addr_city", Me.LegalCity)
            cmd.Parameters.AddWithValue("@legal_addr_st", Me.LegalState)
            cmd.Parameters.AddWithValue("@legal_addr_state", Me.LegalStateDesc)
            cmd.Parameters.AddWithValue("@legal_addr_zip", Me.LegalZip)
            cmd.Parameters.AddWithValue("@legal_addr_county", Me.LegalCnty)
            cmd.Parameters.AddWithValue("@legal_addr_country", Me.LegalCntry)
            cmd.Parameters.AddWithValue("@legal_addr_country_code", Me.LegalCntryCode)

            cmd.Parameters.AddWithValue("@ssn", Me.sSSN)
            cmd.Parameters.AddWithValue("@hire_date", Me.HireDt)
            cmd.Parameters.AddWithValue("@hire_rehire_date", Me.HireReHireDt)
            cmd.Parameters.AddWithValue("@rehire_date", Me.ReHireDt)
            cmd.Parameters.AddWithValue("@position_start_date", Me.PosStart)
            cmd.Parameters.AddWithValue("@position_effective_date", Me.PosEffDt)
            cmd.Parameters.AddWithValue("@position_effective_end_date", Me.PosEffEnd)
            cmd.Parameters.AddWithValue("@termination_date", Me.TermDate)
            cmd.Parameters.AddWithValue("@position_status", Me.PosStatus)
            cmd.Parameters.AddWithValue("@status_effective_date", Me.StatEffDt)
            cmd.Parameters.AddWithValue("@status_effective_end_date", Me.StatEffEndDt)

            cmd.Parameters.AddWithValue("@adjusted_service_date", Me.AdjServDt)
            cmd.Parameters.AddWithValue("@archived_employee", Me.Archived)
            cmd.Parameters.AddWithValue("@position_id", Me.PostnID)
            cmd.Parameters.AddWithValue("@primary_position", Me.PrimPos)
            cmd.Parameters.AddWithValue("@payroll_company_code", Me.PayCompCod)
            cmd.Parameters.AddWithValue("@payroll_company_name", Me.PayCompName)
            cmd.Parameters.AddWithValue("@cip_code", Me.CipCode)
            cmd.Parameters.AddWithValue("@worker_category_code", Me.WorkCatCode)
            cmd.Parameters.AddWithValue("@worker_category_descr", Me.WorkCatDescr)
            cmd.Parameters.AddWithValue("@job_title_code", Me.JobTtlCode)

            cmd.Parameters.AddWithValue("@job_title_descr", Me.JobTtlDescr)
            cmd.Parameters.AddWithValue("@home_cost_number_code", Me.HomeCostCode)
            cmd.Parameters.AddWithValue("@home_cost_number_descr", Me.HomeCostDescr)
            cmd.Parameters.AddWithValue("@job_class_code", Me.JobClassCode)
            cmd.Parameters.AddWithValue("@job_class_descr", Me.JobClassDescr)
            cmd.Parameters.AddWithValue("@job_descr", Me.JobDescr)
            cmd.Parameters.AddWithValue("@job_function_code", Me.JobFuncCode)
            cmd.Parameters.AddWithValue("@job_function_descr", Me.JobFuncDescr)
            cmd.Parameters.AddWithValue("@room", Me.RoomNmbr)
            cmd.Parameters.AddWithValue("@bldg", Me.LocCode)

            cmd.Parameters.AddWithValue("@bldg_name", Me.LocDescr)
            cmd.Parameters.AddWithValue("@leave_of_absence_start_date", Me.LeaveStrDt)
            cmd.Parameters.AddWithValue("@leave_of_absence_return_date", Me.LeaveRetDt)
            cmd.Parameters.AddWithValue("@home_cost_number_2", Me.HomeCostNum2)
            cmd.Parameters.AddWithValue("@payroll_company_code_2", Me.PayCompCode2)
            cmd.Parameters.AddWithValue("@position_effective_date_2", Me.PosEffDt2)
            cmd.Parameters.AddWithValue("@position_end_date_2", Me.PosEffEndDt2)
            cmd.Parameters.AddWithValue("@home_cost_number_3", Me.HomeCostNum3)
            cmd.Parameters.AddWithValue("@payroll_company_code_3", Me.PayCompCode3)
            cmd.Parameters.AddWithValue("@position_effective_date_3", Me.PosEffDt3)
            cmd.Parameters.AddWithValue("@position_end_date_3", Me.PosEffEndDt3)

            cmd.Parameters.AddWithValue("@home_cost_number_4", Me.HomeCostNum4)
            cmd.Parameters.AddWithValue("@payroll_company_code_4", Me.PayCompCode4)
            cmd.Parameters.AddWithValue("@position_effective_date_4", Me.PosEffDt4)
            cmd.Parameters.AddWithValue("@position_end_date_4", Me.PosEffEndDt4)
            cmd.Parameters.AddWithValue("@home_depart_num_code", Me.HomeDptNumCode)
            cmd.Parameters.AddWithValue("@home_depart_num_descr", Me.HomeDptDescr)
            cmd.Parameters.AddWithValue("@supervisor_id", Me.SuprvID)
            cmd.Parameters.AddWithValue("@supervisor_firstname", Me.SuprvFName)
            cmd.Parameters.AddWithValue("@supervisor_lastname", Me.SuprvLName)
            cmd.Parameters.AddWithValue("@date_stamp", Me.dDateStamp)


            '            Dim rowsAffected As Integer = cmd.ExecuteNonQuery()   'Executes query and returns a value
            'Dim rowsAffected = 0 = cmd.ExecuteNonQuery()   'Executes query and returns a value
            'Debug.WriteLine("RowsAffected: {0}", rowsAffected.ToString)
            Debug.WriteLine("Insert into ccADP")
            ConnInformix.Close()
            ConnInformix.Dispose()
            WriteLog("Add of cc_adp_rec data completed for " & Me.PayrollName & ", ID number = " & Me.CarthID & ", ADP ID = " & Me.FileNum & "Job Rec = " & HomeCostCode)
        Catch ex As DuplicateNameException
            'No need to write anything
        Catch e As Exception
            '    Debug.WriteLine(e)
            WriteError("Error in clsProvision - Insert.  Error = " & e.Message & ", Name = " & Me.PayrollName & ", ID number = " & Me.CarthID & ", ADP ID = " & Me.FileNum & "Job Rec = " & HomeCostCode)
            '    'MAIL Event
            '    'sMl.SendEMessage("ADP to CX", (""Error in clsProvision - Insert.  Error = " & e.Messag), "dsullivan@carthage.edu", "dsullivan@carthage.edu", "dsullivan@carthage.edu")
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


            sql = "Update cc_adp_rec set "
            If Me.CarthID <> "" Then sql = sql & " carthage_id = ?"
            If Me.LName <> "" Then sql = sql & ", lastname = ?"
            If Me.FName <> "" Then sql = sql & ", firstname = ?"
            If Me.MidName <> "" Then sql = sql & ", middlename = ?"
            If Me.Salut <> "" Then sql = sql & ", salutation = ?"
            If Me.PayrollName <> "" Then sql = sql & ", fullname = ?"
            If Me.PrefName <> "" Then sql = sql & ", pref_name = ?"
            If Me.BDate <> "" Then sql = sql & ", birth_date = ?"
            If Me.Gender <> "" Then sql = sql & ", gender = ?"
            If Me.MrtStat <> "" Then sql = sql & ", marital_status = ?"

            If Me.RaceCode <> "" Then sql = sql & ", race = ?"
            If Me.RaceDescr <> "" Then sql = sql & ", race_descr = ?"
            If Me.Ethn <> "" Then sql = sql & ", hispanic = ?"
            If Me.EthRaceIDMthd <> "" Then sql = sql & ", race_id_method = ?"
            If Me.PersEmail <> "" Then sql = sql & ", personal_email = ?"
            If Me.PrimAddr1 <> "" Then sql = sql & ", primary_addr_line1 = ?"
            If Me.PrimAddr2 <> "" Then sql = sql & ", primary_addr_line2 = ?"
            If Me.PrimAddr3 <> "" Then sql = sql & ", primary_addr_line3 = ?"

            If Me.PrimCity <> "" Then sql = sql & ", primary_addr_city = ?"
            If Me.PrimState <> "" Then sql = sql & ", primary_addr_st = ?"
            If Me.PrimStateDesc <> "" Then sql = sql & ", primary_addr_state = ?"
            If Me.PrimZip <> "" Then sql = sql & ", primary_addr_zip = ?"
            If Me.PrimCnty <> "" Then sql = sql & ", primary_addr_county = ?"
            If Me.PrimCntry <> "" Then sql = sql & ", primary_addr_country = ?"
            If Me.PrimCntryCode <> "" Then sql = sql & ", primary_addr_country_code = ?"
            If Me.UseAsLegal <> "" Then sql = sql & ", primary_addr_as_legal = ?"
            If Me.HomePhone <> "" Then sql = sql & ", home_phone = ?"
            If Me.PersMobile <> "" Then sql = sql & ", cell_phone = ?"

            If Me.WorkPhone <> "" Then sql = sql & ", work_phone = ?"
            If Me.WorkContPhone <> "" Then sql = sql & ", work_contact_phone = ?"
            If Me.WorkEmail <> "" Then sql = sql & ", work_contact_email = ?"
            If Me.UseWrkEmlCont <> "" Then sql = sql & ", work_contact_notification = ?"
            If Me.LegalAddr1 <> "" Then sql = sql & ", legal_addr_line1 = ?"
            If Me.LegalAddr2 <> "" Then sql = sql & ", legal_addr_line2 = ?"
            If Me.LegalAddr3 <> "" Then sql = sql & ", legal_addr_line3 = ?"
            If Me.LegalCity <> "" Then sql = sql & ", legal_addr_city = ?"
            If Me.LegalState <> "" Then sql = sql & ", legal_addr_st = ?"
            If Me.LegalStateDesc <> "" Then sql = sql & ", legal_addr_state = ?"
            If Me.LegalZip <> "" Then sql = sql & ", legal_addr_zip = ?"
            If Me.LegalCnty <> "" Then sql = sql & ", legal_addr_county = ?"
            If Me.LegalCntry <> "" Then sql = sql & ", legal_addr_country = ?"
            If Me.LegalCntryCode <> "" Then sql = sql & ", legal_addr_country_code = ?"

            If Me.sSSN <> "" Then sql = sql & ", ssn = ?"
            If Me.HireDt <> "" Then sql = sql & ", hire_date = ?"
            If Me.HireReHireDt <> "" Then sql = sql & ", hire_rehire_date = ?"
            If Me.ReHireDt <> "" Then sql = sql & ", rehire_date = ?"
            If Me.PosStart <> "" Then sql = sql & ", position_start_date = ?"
            If Me.PosEffDt <> "" Then sql = sql & ", position_effective_date = ?"
            If Me.PosEffEnd <> "" Then sql = sql & ", position_effective_end_date = ?"
            If Me.TermDate <> "" Then sql = sql & ", termination_date = ?"
            If Me.PosStatus <> "" Then sql = sql & ", position_status = ?"
            If Me.StatEffDt <> "" Then sql = sql & ", status_effective_date = ?"
            If Me.StatEffEndDt <> "" Then sql = sql & ", status_effective_end_date = ?"

            If Me.AdjServDt <> "" Then sql = sql & ", adjusted_service_date = ?"
            If Me.Archived <> "" Then sql = sql & ", archived_employee = ?"
            If Me.PostnID <> "" Then sql = sql & ", position_id = ?"
            If Me.PrimPos <> "" Then sql = sql & ", primary_position = ?"
            If Me.PayCompCod <> "" Then sql = sql & ", payroll_company_code = ?"
            If Me.PayCompName <> "" Then sql = sql & ", payroll_company_name = ?"
            If Me.CipCode <> "" Then sql = sql & ", cip_code = ?"
            If Me.WorkCatCode <> "" Then sql = sql & ", worker_category_code = ?"
            If Me.WorkCatDescr <> "" Then sql = sql & ", worker_category_descr = ?"
            If Me.JobTtlCode <> "" Then sql = sql & ", job_title_code = ?"

            If Me.JobTtlDescr <> "" Then sql = sql & ", job_title_descr = ?"
            If Me.HomeCostCode <> "" Then sql = sql & ", home_cost_number_code = ?"
            If Me.HomeCostDescr <> "" Then sql = sql & ", home_cost_number_descr = ?"
            If Me.JobClassCode <> "" Then sql = sql & ", job_class_code = ?"
            If Me.JobClassDescr <> "" Then sql = sql & ", job_class_descr = ?"
            If Me.JobDescr <> "" Then sql = sql & ", job_descr = ?"
            If Me.JobFuncCode <> "" Then sql = sql & ", job_function_code = ?"
            If Me.JobFuncDescr <> "" Then sql = sql & ", job_function_descr = ?"
            If Me.RoomNmbr <> "" Then sql = sql & ", room = ?"
            If Me.LocCode <> "" Then sql = sql & ", bldg = ?"

            If Me.LocDescr <> "" Then sql = sql & ", bldg_name = ?"
            If Me.LeaveStrDt <> "" Then sql = sql & ", leave_of_absence_start_date = ?"
            If Me.LeaveRetDt <> "" Then sql = sql & ", leave_of_absence_return_date = ?"
            If Me.HomeCostNum2 <> "" Then sql = sql & ", home_cost_number_2 = ?"
            If Me.PayCompCode2 <> "" Then sql = sql & ", payroll_company_code_2 = ?"
            If Me.PosEffDt2 <> "" Then sql = sql & ", position_effective_date_2 = ?"
            If Me.PosEffEndDt2 <> "" Then sql = sql & ", position_end_date_2 = ?"
            If Me.HomeCostNum3 <> "" Then sql = sql & ", home_cost_number_3 = ?"
            If Me.PayCompCode3 <> "" Then sql = sql & ", payroll_company_code_3 = ?"
            If Me.PosEffDt3 <> "" Then sql = sql & ", position_effective_date_3 = ?"
            If Me.PosEffEndDt3 <> "" Then sql = sql & ", position_end_date_3 = ?"

            If Me.HomeCostNum4 <> "" Then sql = sql & ", home_cost_number_4 = ?"
            If Me.PayCompCode4 <> "" Then sql = sql & ", payroll_company_code_4 = ?"
            If Me.PosEffDt4 <> "" Then sql = sql & ", position_effective_date_4 = ?"
            If Me.PosEffEndDt4 <> "" Then sql = sql & ", position_end_date_4 = ?"
            If Me.HomeDptNumCode <> "" Then sql = sql & ", home_depart_num_code = ?"
            If Me.HomeDptDescr <> "" Then sql = sql & ", home_depart_num_descr = ?"
            If Me.SuprvID <> "" Then sql = sql & ", supervisor_id = ?"
            If Me.SuprvFName <> "" Then sql = sql & ", supervisor_firstname = ?"
            If Me.SuprvLName <> "" Then sql = sql & ", supervisor_lastname = ?"
            sql = sql & ", DateStamp = ?"

            sql = sql & " WHERE file_no = ?"

            Dim cmd As New OdbcCommand(sql, ConnInformix)
            ConnInformix.ConnectionString = gl.ConnectString
            ConnInformix.Open()

            cmd.CommandType = CommandType.Text




            '    'I believe the preferred method these days is to NOT leave the connection open, but connect, execute and close...
            '    ConnInformix.ConnectionString = gl.ConnectString
            '    ConnInformix.Open()
            '    cmd.CommandType = CommandType.Text
            If Me.CarthID <> "" Then cmd.Parameters.AddWithValue("@carthage_id", Me.CarthID)
            If Me.LName <> "" Then cmd.Parameters.AddWithValue("@lastname", Me.LName)
            If Me.FName <> "" Then cmd.Parameters.AddWithValue("@firstname", Me.FName)
            If Me.MidName <> "" Then cmd.Parameters.AddWithValue("@middlename", Me.MidName)
            If Me.Salut <> "" Then cmd.Parameters.AddWithValue("@salutation", Me.Salut)
            If Me.PayrollName <> "" Then cmd.Parameters.AddWithValue("@fullname", Me.PayrollName)
            If Me.PrefName <> "" Then cmd.Parameters.AddWithValue("@pref_name", Me.PrefName)
            If Me.BDate <> "" Then cmd.Parameters.AddWithValue("@birth_date", Me.BDate)
            If Me.Gender <> "" Then cmd.Parameters.AddWithValue("@gender", Me.Gender)
            If Me.MrtStat <> "" Then cmd.Parameters.AddWithValue("@marital_status", Me.MrtStat)

            If Me.RaceCode <> "" Then cmd.Parameters.AddWithValue("@race", Me.RaceCode)
            If Me.RaceDescr <> "" Then cmd.Parameters.AddWithValue("@race_descr", Me.RaceDescr)
            If Me.Ethn <> "" Then cmd.Parameters.AddWithValue("@hispanic", Me.Ethn)
            If Me.EthRaceIDMthd <> "" Then cmd.Parameters.AddWithValue("@race_id_method", Me.EthRaceIDMthd)
            If Me.PersEmail <> "" Then cmd.Parameters.AddWithValue("@personal_email", Me.PersEmail)
            If Me.PrimAddr1 <> "" Then cmd.Parameters.AddWithValue("@primary_addr_line1", Me.PrimAddr1)
            If Me.PrimAddr2 <> "" Then cmd.Parameters.AddWithValue("@primary_addr_line2", Me.PrimAddr2)
            If Me.PrimAddr3 <> "" Then cmd.Parameters.AddWithValue("@primary_addr_line3", Me.PrimAddr3)

            If Me.PrimCity <> "" Then cmd.Parameters.AddWithValue("@primary_addr_city", Me.PrimCity)
            If Me.PrimState <> "" Then cmd.Parameters.AddWithValue("@primary_addr_st", Me.PrimState)
            If Me.PrimStateDesc <> "" Then cmd.Parameters.AddWithValue("@primary_addr_state", Me.PrimStateDesc)
            If Me.PrimZip <> "" Then cmd.Parameters.AddWithValue("@primary_addr_zip", Me.PrimZip)
            If Me.PrimCnty <> "" Then cmd.Parameters.AddWithValue("@primary_addr_county", Me.PrimCnty)
            If Me.PrimCntry <> "" Then cmd.Parameters.AddWithValue("@primary_addr_country", Me.PrimCntry)
            If Me.PrimCntryCode <> "" Then cmd.Parameters.AddWithValue("@primary_addr_country_code", Me.PrimCntryCode)
            If Me.UseAsLegal <> "" Then cmd.Parameters.AddWithValue("@primary_addr_as_legal", Me.UseAsLegal)
            If Me.HomePhone <> "" Then cmd.Parameters.AddWithValue("@home_phone", Me.HomePhone)
            If Me.PersMobile <> "" Then cmd.Parameters.AddWithValue("@cell_phone", Me.PersMobile)

            If Me.WorkPhone <> "" Then cmd.Parameters.AddWithValue("@work_phone", Me.WorkPhone)
            If Me.WorkContPhone <> "" Then cmd.Parameters.AddWithValue("@work_contact_phone", Me.WorkContPhone)
            If Me.WorkEmail <> "" Then cmd.Parameters.AddWithValue("@work_contact_email", Me.WorkEmail)
            If Me.UseWrkEmlCont <> "" Then cmd.Parameters.AddWithValue("@work_contact_notification", Me.UseWrkEmlCont)
            If Me.LegalAddr1 <> "" Then cmd.Parameters.AddWithValue("@legal_addr_line1", Me.LegalAddr1)
            If Me.LegalAddr2 <> "" Then cmd.Parameters.AddWithValue("@legal_addr_line2", Me.LegalAddr2)
            If Me.LegalAddr3 <> "" Then cmd.Parameters.AddWithValue("@legal_addr_line3", Me.LegalAddr3)
            If Me.LegalCity <> "" Then cmd.Parameters.AddWithValue("@legal_addr_city", Me.LegalCity)
            If Me.LegalState <> "" Then cmd.Parameters.AddWithValue("@legal_addr_st", Me.LegalState)
            If Me.LegalStateDesc <> "" Then cmd.Parameters.AddWithValue("@legal_addr_state", Me.LegalStateDesc)
            If Me.LegalZip <> "" Then cmd.Parameters.AddWithValue("@legal_addr_zip", Me.LegalZip)
            If Me.LegalCnty <> "" Then cmd.Parameters.AddWithValue("@legal_addr_county", Me.LegalCnty)
            If Me.LegalCntry <> "" Then cmd.Parameters.AddWithValue("@legal_addr_country", Me.LegalCntry)
            If Me.LegalCntryCode <> "" Then cmd.Parameters.AddWithValue("@legal_addr_country_code", Me.LegalCntryCode)

            If Me.sSSN <> "" Then cmd.Parameters.AddWithValue("@ssn", Me.sSSN)
            If Me.HireDt <> "" Then cmd.Parameters.AddWithValue("@hire_date", Me.HireDt)
            If Me.HireReHireDt <> "" Then cmd.Parameters.AddWithValue("@hire_rehire_date", Me.HireReHireDt)
            If Me.ReHireDt <> "" Then cmd.Parameters.AddWithValue("@rehire_date", Me.ReHireDt)
            If Me.PosStart <> "" Then cmd.Parameters.AddWithValue("@position_start_date", Me.PosStart)
            If Me.PosEffDt <> "" Then cmd.Parameters.AddWithValue("@position_effective_date", Me.PosEffDt)
            If Me.PosEffEnd <> "" Then cmd.Parameters.AddWithValue("@position_effective_end_date", Me.PosEffEnd)
            If Me.TermDate <> "" Then cmd.Parameters.AddWithValue("@termination_date", Me.TermDate)
            If Me.PosStatus <> "" Then cmd.Parameters.AddWithValue("@position_status", Me.PosStatus)
            If Me.StatEffDt <> "" Then cmd.Parameters.AddWithValue("@status_effective_date", Me.StatEffDt)
            If Me.StatEffEndDt <> "" Then cmd.Parameters.AddWithValue("@status_effective_end_date", Me.StatEffEndDt)

            If Me.AdjServDt <> "" Then cmd.Parameters.AddWithValue("@adjusted_service_date", Me.AdjServDt)
            If Me.Archived <> "" Then cmd.Parameters.AddWithValue("@archived_employee", Me.Archived)
            If Me.PostnID <> "" Then cmd.Parameters.AddWithValue("@position_id", Me.PostnID)
            If Me.PrimPos <> "" Then cmd.Parameters.AddWithValue("@primary_position", Me.PrimPos)
            If Me.PayCompCod <> "" Then cmd.Parameters.AddWithValue("@payroll_company_code", Me.PayCompCod)
            If Me.PayCompName <> "" Then cmd.Parameters.AddWithValue("@payroll_company_name", Me.PayCompName)
            If Me.CipCode <> "" Then cmd.Parameters.AddWithValue("@cip_code", Me.CipCode)
            If Me.WorkCatCode <> "" Then cmd.Parameters.AddWithValue("@worker_category_code", Me.WorkCatCode)
            If Me.WorkCatDescr <> "" Then cmd.Parameters.AddWithValue("@worker_category_descr", Me.WorkCatDescr)
            If Me.JobTtlCode <> "" Then cmd.Parameters.AddWithValue("@job_title_code", Me.JobTtlCode)

            If Me.JobTtlDescr <> "" Then cmd.Parameters.AddWithValue("@job_title_descr", Me.JobTtlDescr)
            If Me.HomeCostCode <> "" Then cmd.Parameters.AddWithValue("@home_cost_number_code", Me.HomeCostCode)
            If Me.HomeCostDescr <> "" Then cmd.Parameters.AddWithValue("@home_cost_number_descr", Me.HomeCostDescr)
            If Me.JobClassCode <> "" Then cmd.Parameters.AddWithValue("@job_class_code", Me.JobClassCode)
            If Me.JobClassDescr <> "" Then cmd.Parameters.AddWithValue("@job_class_descr", Me.CarthID)
            If Me.JobDescr <> "" Then cmd.Parameters.AddWithValue("@job_descr", Me.JobClassDescr)
            If Me.JobFuncCode <> "" Then cmd.Parameters.AddWithValue("@job_function_code", Me.JobFuncCode)
            If Me.JobFuncDescr <> "" Then cmd.Parameters.AddWithValue("@job_function_descr", Me.JobFuncDescr)
            If Me.RoomNmbr <> "" Then cmd.Parameters.AddWithValue("@room", Me.RoomNmbr)
            If Me.LocCode <> "" Then cmd.Parameters.AddWithValue("@bldg", Me.LocCode)

            If Me.LocDescr <> "" Then cmd.Parameters.AddWithValue("@bldg_name", Me.LocDescr)
            If Me.LeaveStrDt <> "" Then cmd.Parameters.AddWithValue("@leave_of_absence_start_date", Me.LeaveStrDt)
            If Me.LeaveRetDt <> "" Then cmd.Parameters.AddWithValue("@leave_of_absence_return_date", Me.LeaveRetDt)
            If Me.HomeCostNum2 <> "" Then cmd.Parameters.AddWithValue("@home_cost_number_2", Me.HomeCostNum2)
            If Me.PayCompCode2 <> "" Then cmd.Parameters.AddWithValue("@payroll_company_code_2", Me.PayCompCode2)
            If Me.PosEffDt2 <> "" Then cmd.Parameters.AddWithValue("@position_effective_date_2", Me.PosEffDt2)
            If Me.PosEffEndDt2 <> "" Then cmd.Parameters.AddWithValue("@position_end_date_2", Me.PosEffEndDt2)
            If Me.HomeCostNum3 <> "" Then cmd.Parameters.AddWithValue("@home_cost_number_3", Me.HomeCostNum3)
            If Me.PayCompCode3 <> "" Then cmd.Parameters.AddWithValue("@payroll_company_code_3", Me.PayCompCode3)
            If Me.PosEffDt3 <> "" Then cmd.Parameters.AddWithValue("@position_effective_date_3", Me.PosEffDt3)
            If Me.PosEffEndDt3 <> "" Then cmd.Parameters.AddWithValue("@position_end_date_3", Me.PosEffEndDt3)

            If Me.HomeCostNum4 <> "" Then cmd.Parameters.AddWithValue("@home_cost_number_4", Me.HomeCostNum4)
            If Me.PayCompCode4 <> "" Then cmd.Parameters.AddWithValue("@payroll_company_code_4", Me.PayCompCode4)
            If Me.PosEffDt4 <> "" Then cmd.Parameters.AddWithValue("@position_effective_date_4", Me.PosEffDt4)
            If Me.PosEffEndDt4 <> "" Then cmd.Parameters.AddWithValue("@position_end_date_4", Me.PosEffEndDt4)
            If Me.HomeDptNumCode <> "" Then cmd.Parameters.AddWithValue("@home_depart_num_code", Me.HomeDptNumCode)
            If Me.HomeDptDescr <> "" Then cmd.Parameters.AddWithValue("@carhome_depart_num_descrthage_id", Me.HomeDptDescr)
            If Me.SuprvID <> "" Then cmd.Parameters.AddWithValue("@supervisor_id", Me.SuprvID)
            If Me.SuprvFName <> "" Then cmd.Parameters.AddWithValue("@supervisor_firstname", Me.SuprvFName)
            If Me.SuprvLName <> "" Then cmd.Parameters.AddWithValue("@supervisor_lastname", Me.SuprvLName)
            cmd.Parameters.AddWithValue("@date_stamp", Me.dDateStamp)

            'cmd.Parameters.AddWithValue("DateStamp", ConvertDATE(Now))
            cmd.Parameters.AddWithValue("@file_no", Me.FileNum)

            '            Dim rowsAffected As Integer = cmd.ExecuteNonQuery()   'Executes query and returns a value
            '            Debug.WriteLine("RowsAffected: {0}", rowsAffected.ToString)

            ConnInformix.Close()
            ConnInformix.Dispose()
            WriteLog("Update of CVID_Rec completed for " & Me.PayrollName & ", ID number = " & Me.CarthID & ", ADP ID = " & Me.FileNum & "Job Rec = " & HomeCostCode)

        Catch e As Exception
            Debug.WriteLine(e)
            WriteError("Error in clsProvision - Update.  Error = " & e.Message & ", Name = " & Me.PayrollName & ", ID number = " & Me.CarthID & ", ADP ID = " & Me.FileNum & "Job Rec = " & HomeCostCode)
            'MAIL Event
            'sMl.SendEMessage("ADP to CX", (""Error in clsProvision - Update.  Error = " & e.Messag), "dsullivan@carthage.edu", "dsullivan@carthag.edu", "dsullivan@carthage.edu")

        Finally
            If ConnInformix.State = ConnectionState.Open Then
                ConnInformix.Close()
                ConnInformix.Dispose()
                GC.Collect()
            End If
        End Try
    End Sub

    Public Sub Initialize()

        sFileNum = ""
        sCarthID = ""
        sLName = ""
        sFName = ""
        sMidName = ""
        sSalut = ""
        sPayrollName = ""
        sPrefName = ""
        dBDate = ""
        sGender = ""
        sMrtStat = ""
        sRaceCode = ""
        sRaceDescr = ""
        sEthn = ""
        sEthRaceIDMthd = ""
        sPersEmail = ""
        sPrimAddr1 = ""
        sPrimAddr2 = ""
        sPrimAddr3 = ""
        sPrimCity = ""
        sPrimState = ""
        sPrimStateDesc = ""
        sPrimZip = ""
        sPrimCnty = ""
        sPrimCntry = ""
        sPrimCntryCode = ""
        sUseAsLegal = ""
        sHomePhone = ""
        sPersMobile = ""
        sWorkPhone = ""
        sWorkContPhone = ""
        sWorkEmail = ""
        sUseWrkEmlCont = ""

        sLegalAddr1 = ""
        sLegalAddr2 = ""
        sLegalAddr3 = ""
        sLegalCity = ""
        sLegalState = ""
        sLegalStateDesc = ""
        sLegalZip = ""
        sLegalCnty = ""
        sLegalCntry = ""
        sLegalCntryCode = ""

        sSSN = ""
        dHireDt = ""
        dHireRehireDt = ""
        dRehireDt = ""
        dPosStart = ""
        dPosEffDt = ""
        dPosEffEnd = ""
        dTermDate = ""
        sPosStatus = ""
        dStatEffDt = ""
        dStatEffEndDt = ""
        dAdjServDt = ""
        sArchived = ""
        sPostnID = ""
        sPrimPos = ""
        sPayCompCod = ""
        sPayCompName = ""
        sCIPCode = ""
        sWorkCatCode = ""
        sWorkCatDescr = ""
        sJobTtlCode = ""
        sJobTtlDescr = ""
        sHomeCostCode = ""
        sHomeCostDescr = ""
        sJobClassCode = ""
        sJobClassDescr = ""
        sJobDescr = ""
        sJobFuncCode = ""
        sJobFuncDescr = ""
        sRoomNmbr = ""
        sLocCode = ""
        sLocDescr = ""
        LeaveStrDt = ""
        LeaveRetDt = ""
        sHomeCostNum2 = ""
        sPayCompCode2 = ""
        dPosEffDt2 = ""
        dPosEffEndDt2 = ""

        sHomeCostNum3 = ""
        sPayCompCode3 = ""
        dPosEffDt3 = ""
        dPosEffEndDt3 = ""

        sHomeCostNum4 = ""
        sPayCompCode4 = ""
        dPosEffDt4 = ""
        dPosEffEndDt4 = ""

        sHomeDptNumCode = ""
        sHomeDptDescr = ""
        sSuprvID = ""
        sSuprvFName = ""
        sSuprvLName = ""
        dDateStamp = ""

    End Sub
End Class
