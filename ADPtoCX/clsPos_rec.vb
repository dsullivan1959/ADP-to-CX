Option Strict Off
Option Explicit On

Imports System.Data.Odbc

Friend Class ClsPos_rec


    Dim sPos As String
    Dim dBegDate As Date
    Dim dEndDate As Date
    Dim sFund As String
    Dim sFunc As String
    Dim sObj As String
    Dim sProj As String
    Dim sDescr As String
    Dim sRank As String
    Dim dbSalMin As Double
    Dim dbSalMax As Double
    Dim dbHrlyMin As Double
    Dim dbHrlyMax As Double
    Dim sOTCode As String
    Dim iLevel1 As Integer
    Dim iLevel2 As Integer
    Dim sCamp As String
    Dim sBldg As String
    Dim sRoom As String
    Dim sExt As String
    Dim sSubs As String
    Dim sSubsID As Integer
    Dim sBalCode As String
    Dim sTotCode As String
    Dim sExcl As String



    Public Property Pos() As String
        Get
            Pos = sPos
        End Get
        Set(ByVal Value As String)
            sPos = Value
        End Set
    End Property

    Public Property BegDate() As Date
        Get
            BegDate = dBegDate
        End Get
        Set(ByVal Value As Date)
            dBegDate = Value
        End Set
    End Property
    Public Property EndDate() As Date
        Get
            EndDate = dEndDate
        End Get
        Set(ByVal Value As Date)
            dEndDate = Value
        End Set
    End Property

    Public Property Fund() As String
        Get
            Fund = sFund
        End Get
        Set(ByVal Value As String)
            sFund = Value
        End Set
    End Property

    Public Property Func() As String
        Get
            Func = sFunc
        End Get
        Set(ByVal Value As String)
            sFunc = Value
        End Set
    End Property
    Public Property Obj() As String
        Get
            Obj = sObj
        End Get
        Set(ByVal Value As String)
            sObj = Value
        End Set
    End Property

    Public Property Proj() As String
        Get
            Proj = sProj
        End Get
        Set(ByVal Value As String)
            sProj = Value
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
    Public Property Rank() As String
        Get
            Rank = sRank
        End Get
        Set(ByVal Value As String)
            sRank = Value
        End Set
    End Property
    Public Property SalMin() As Double
        Get
            SalMin = dbSalMin
        End Get
        Set(ByVal Value As Double)
            dbSalMin = Value
        End Set
    End Property
    Public Property SalMax() As Double
        Get
            SalMax = dbSalMax
        End Get
        Set(ByVal Value As Double)
            dbSalMax = Value
        End Set
    End Property
    Public Property HrlyMin() As Double
        Get
            HrlyMin = dbHrlyMin
        End Get
        Set(ByVal Value As Double)
            dbHrlyMin = Value
        End Set
    End Property
    Public Property HrlyMax() As Double
        Get
            HrlyMax = dbHrlyMax
        End Get
        Set(ByVal Value As Double)
            dbHrlyMax = Value
        End Set
    End Property
    Public Property OTCode() As String
        Get
            OTCode = sOTCode
        End Get
        Set(ByVal Value As String)
            sOTCode = Value
        End Set
    End Property
    Public Property Level1() As Integer
        Get
            Level1 = iLevel1
        End Get
        Set(ByVal Value As Integer)
            iLevel1 = Value
        End Set
    End Property
    Public Property Level2() As Integer
        Get
            Level2 = iLevel2
        End Get
        Set(ByVal Value As Integer)
            iLevel2 = Value
        End Set
    End Property
    Public Property Camp() As String
        Get
            Camp = sCamp
        End Get
        Set(ByVal Value As String)
            sCamp = Value
        End Set
    End Property

    Public Property Bldg() As String
        Get
            Bldg = sBldg
        End Get
        Set(ByVal Value As String)
            sBldg = Value
        End Set
    End Property
    Public Property Room() As String
        Get
            Room = sRoom
        End Get
        Set(ByVal Value As String)
            sRoom = Value
        End Set
    End Property
    Public Property Ext() As String
        Get
            Ext = sExt
        End Get
        Set(ByVal Value As String)
            sExt = Value
        End Set
    End Property
    Public Property Subs() As String
        Get
            Subs = sSubs
        End Get
        Set(ByVal Value As String)
            sSubs = Value
        End Set
    End Property
    Public Property SubsID() As String
        Get
            SubsID = sSubsID
        End Get
        Set(ByVal Value As String)
            sSubsID = Value
        End Set
    End Property
    Public Property BalCode() As String
        Get
            BalCode = sBalCode
        End Get
        Set(ByVal Value As String)
            sBalCode = Value
        End Set
    End Property
    Public Property TotCode() As String
        Get
            TotCode = sTotCode
        End Get
        Set(ByVal Value As String)
            sTotCode = Value
        End Set
    End Property
    Public Property Excl() As String
        Get
            Excl = sExcl
        End Get
        Set(ByVal Value As String)
            sExcl = Value
        End Set
    End Property

    Public Sub Insert()

        Dim sql As String
        
        Try
            sql = " INSERT INTO pos_rec "
            sql = sql & "(pos, beg_date, end_date, fund, func, obj, proj, descr, rank, "
            sql = sql & " sal_min, sal_max, hrly_min, hrly_max, ot_code, level1, level2, camp, bldg, room, "
            sql = sql & " ext, subs, subs_id, bal_code, tot_code, excl)"
            sql = sql & " values (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,? ) "

            Dim cmd As New OdbcCommand(sql, ConnInformix)

            cmd.Parameters.AddWithValue("@POS", Me.Pos)
            cmd.Parameters.AddWithValue("@beg_date", Me.BegDate)
            cmd.Parameters.AddWithValue("@end_date", Me.EndDate)
            cmd.Parameters.AddWithValue("@fund", Me.Fund)
            cmd.Parameters.AddWithValue("@func", Me.Func)
            cmd.Parameters.AddWithValue("@obj", Me.Obj)
            cmd.Parameters.AddWithValue("@proj", Me.Proj)
            cmd.Parameters.AddWithValue("@descr", Me.Descr)
            cmd.Parameters.AddWithValue("@rank", Me.Rank)
            cmd.Parameters.AddWithValue("@sal_min", Me.SalMin)
            cmd.Parameters.AddWithValue("@sal_max", Me.SalMax)
            cmd.Parameters.AddWithValue("@hrly_min", Me.HrlyMin)
            cmd.Parameters.AddWithValue("@hrly_max", Me.HrlyMax)
            cmd.Parameters.AddWithValue("@ot_code", Me.OTCode)
            cmd.Parameters.AddWithValue("@level1", Me.Level1)
            cmd.Parameters.AddWithValue("@level2", Me.Level2)
            cmd.Parameters.AddWithValue("@camp", Me.Camp)
            cmd.Parameters.AddWithValue("@bldg", Me.Bldg)
            cmd.Parameters.AddWithValue("@room", Me.Room)
            cmd.Parameters.AddWithValue("@ext", Me.Ext)
            cmd.Parameters.AddWithValue("@subs", Me.Subs)
            cmd.Parameters.AddWithValue("@subs_id", Me.SubsID)
            cmd.Parameters.AddWithValue("@bal_code", Me.BalCode)
            cmd.Parameters.AddWithValue("@tot_code", Me.TotCode)
            cmd.Parameters.AddWithValue("@excl", Me.Excl)


            ConnInformix.ConnectionString = gl.ConnectString
            ConnInformix.Open()

            '            Dim rowsAffected As Integer = cmd.ExecuteNonQuery()
            ConnInformix.Close()
            ConnInformix.Dispose()
            WriteLog("Create of New Position ID record for Position " & Me.Pos & ", Descr = " & Me.Descr)

        Catch e As Exception
            debug.writeline(e)
            WriteError("Error in clsPos_rec - Insert.  Error = " & Err.Description & ", Error Number " & Err.Number)
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


            sql = "Update pos_rec set  [beg_date] = ?, [end_date] = ?, [fund] = ?, [func] = ?, [obj] = ?, [proj] = ?, [descr] = ?, "
            sql = sql & " [rank] = ?, [sal_min] = ?, [sal_max] = ?, [hrly_min] = ?, [hrly_max] = ?, [ot_code] = ?. [level1] = ?, "
            sql = sql & " [level2] = ?, [camp] = ?,  [bldg] = ?, [room] = ?, [ext] = ?,  [subs] = ?, [subs_id] = ?, "
            sql = sql & " [bal_code] = ?, [tot_code] = ?, [excl] = ?"
            sql = sql & " WHERE [pos] = ?"

            sql = "Update pos_rec set "
            If Me.Fund <> "" Then sql = sql & " fund = ?"
            If Me.Func <> "" Then sql = sql & ", func = ?"
            If Me.Obj <> "" Then sql = sql & ", obj = ?"
            If Me.Proj <> "" Then sql = sql & ", proj = ?"

            If Me.BegDate <> "" Then sql = sql & ", beg_date = ?"
            If Me.EndDate <> "" Then sql = sql & ", end_date = ?"

            If Me.Fund <> "" Then sql = sql & ", fund = ?"
            If Me.Func <> "" Then sql = sql & ", func = ?"
            If Me.Obj <> "" Then sql = sql & ", obj = ?"
            If Me.Proj <> "" Then sql = sql & ", proj = ?"
            If Me.Descr <> "" Then sql = sql & ", descr = ?"
            If Me.Rank <> "" Then sql = sql & ", rank = ?"
            If Me.SalMin <> "" Then sql = sql & ", sal_min = ?"
            If Me.SalMax <> "" Then sql = sql & ", sal_max = ?"
            If Me.HrlyMin <> "" Then sql = sql & ", hrly_min = ?"
            If Me.HrlyMax <> "" Then sql = sql & ", hrly_max = ?"
            If Me.OTCode <> "" Then sql = sql & ", ot_code = ?"
            If Me.Level1 <> "" Then sql = sql & ", level1 = ?"
            If Me.Level2 <> "" Then sql = sql & ", level2 = ?"
            If Me.Camp <> "" Then sql = sql & ", camp = ?"
            If Me.Bldg <> "" Then sql = sql & ", bldg = ?"
            If Me.Room <> "" Then sql = sql & ", room = ?"
            If Me.Ext <> "" Then sql = sql & ", ext = ?"
            'If Me.Subs <> "" Then sql = sql & ", subs = ?"
            'If Me.SubsID <> "" Then sql = sql & ", subs_id = ?"
            'If Me.BalCode <> "" Then sql = sql & ", bal_code = ?"
            'If Me.TotCode <> "" Then sql = sql & ", tot_code = ?"
            'If Me.sExcl <> "" Then sql = sql & ", excl = ?"
            sql = sql & " WHERE [pos] = ?"

            Dim cmd As New OdbcCommand(sql, ConnInformix)

            If Me.Fund <> "" Then cmd.Parameters.AddWithValue("@fund", Me.Fund)
            If Me.Func <> "" Then cmd.Parameters.AddWithValue("@func", Me.Func)
            If Me.Obj <> "" Then cmd.Parameters.AddWithValue("@obj", Me.Obj)
            If Me.Proj <> "" Then cmd.Parameters.AddWithValue("@proj", Me.Proj)
            If Me.BegDate <> "" Then cmd.Parameters.AddWithValue("@beg_date", Me.BegDate)
            If Me.EndDate <> "" Then cmd.Parameters.AddWithValue("@end_date", Me.EndDate)
            If Me.Descr <> "" Then cmd.Parameters.AddWithValue("@descr", Me.Descr)
            If Me.Rank <> "" Then cmd.Parameters.AddWithValue("@rank", Me.Rank)
            If Me.SalMin <> "" Then cmd.Parameters.AddWithValue("@sal_min", Me.SalMin)
            If Me.SalMax <> "" Then cmd.Parameters.AddWithValue("@sal_max", Me.SalMax)
            If Me.HrlyMin <> "" Then cmd.Parameters.AddWithValue("@hrly_min", Me.HrlyMin)
            If Me.HrlyMax <> "" Then cmd.Parameters.AddWithValue("@hrly_max", Me.HrlyMax)
            If Me.OTCode <> "" Then cmd.Parameters.AddWithValue("@ot_code", Me.OTCode)
            If Me.Level1 <> "" Then cmd.Parameters.AddWithValue("@level1", Me.Level1)
            If Me.Level2 <> "" Then cmd.Parameters.AddWithValue("@level2", Me.Level2)
            If Me.Camp <> "" Then cmd.Parameters.AddWithValue("@camp", Me.Camp)
            If Me.Bldg <> "" Then cmd.Parameters.AddWithValue("@bldg", Me.Bldg)
            If Me.Room <> "" Then cmd.Parameters.AddWithValue("@room", Me.Room)
            If Me.Ext <> "" Then cmd.Parameters.AddWithValue("@ext", Me.Ext)
            'If Me.Subs <> "" Then cmd.Parameters.AddWithValue("@subs", Me.subs)
            'If Me.sSubsID <> "" Then cmd.Parameters.AddWithValue("@subs_id", Me.subsid)
            'If Me.BalCode <> "" Then cmd.Parameters.AddWithValue("@bal_code", Me.balcode)
            'If Me.TotCode <> "" Then cmd.Parameters.AddWithValue("@tot_code", Me.totcode)
            'If Me.Excl <> "" Then cmd.Parameters.AddWithValue("@excl", Me.excl)
            cmd.Parameters.AddWithValue("@POS", Me.Pos)

            '
            ConnInformix.ConnectionString = gl.ConnectString
            ConnInformix.Open()
            'Dim rowsAffected As Integer = cmd.ExecuteNonQuery()
            'ConnInformix.Close()
            WriteLog("Update of Position ID record for Position " & Me.Pos & ", Descr = " & Me.Descr)
            ConnInformix.Close()
            ConnInformix.Dispose()

        Catch e As NullReferenceException
            MsgBox("Null value")
        Catch ex As DuplicateNameException
            'No need to write anything
        Catch e As Exception
            debug.writeline(e)
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

        sPos = ""
        dBegDate = ""
        dEndDate = ""
        sFund = ""
        sFunc = ""
        sObj = ""
        sProj = ""
        sDescr = ""
        sRank = ""
        dbSalMin = 0
        dbSalMax = 0
        dbHrlyMin = 0
        dbHrlyMax = 0
        sOTCode = ""
        iLevel1 = 0
        iLevel2 = 0
        sCamp = ""
        sBldg = ""
        sRoom = ""
        sExt = ""
        sSubs = ""
        sSubsID = 0
        sBalCode = ""
        sTotCode = ""
        sExcl = ""

    End Sub
End Class

