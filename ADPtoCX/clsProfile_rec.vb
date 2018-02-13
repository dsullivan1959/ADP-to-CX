Option Strict Off
Option Explicit On

Imports System.Data.Odbc

Friend Class clsProfile_rec

    Dim iID As Integer
    Dim sEthnicCode As String
    Dim sSex As String
    Dim sMrtl As String
    Dim dBirthdate As String
    Dim iAge As Integer
    Dim sBirthCity As String
    Dim sBirthState As String
    Dim sBirthCntry As String
    Dim sResState As String
    Dim sResCity As String
    Dim sCitz As String
    Dim sVisaCode As String
    Dim dProfVisaDate As String
    Dim sProfVisaNo As String
    Dim sResCntry As String
    Dim sVet As String
    Dim sHandicapCode As String
    Dim sDenomCode As String
    Dim iChurchID As Integer
    Dim sOcc As String
    Dim iNews1ID As Integer
    Dim iNews2ID As Integer
    Dim dProfLastUpd As String

    Dim sLang As String
    Dim sprofResCode As String
    Dim dprofResDate As String

    Dim sProfVetChap As String
    Dim iGrpNo As Integer
    Dim sPrivCode As String
    Dim dDecsdDate As String
    Dim sMilitCode As String
    Dim sMilitRank As String
    Dim bPhoto As Byte
    Dim sRace As String
    Dim sHispanic As String
    Dim dConfirmDate As String
    Dim sFullName As String







    Public Property ID() As String
        Get
            ID = iID
        End Get
        Set(ByVal Value As String)
            iID = Value
        End Set
    End Property

    Public Property EthnicCode() As String
        Get
            EthnicCode = sEthnicCode
        End Get
        Set(ByVal Value As String)
            sEthnicCode = Value
        End Set
    End Property

    Public Property Sex() As String
        Get
            Sex = sSex
        End Get
        Set(ByVal Value As String)
            sSex = Value
        End Set
    End Property

    Public Property Mrtl() As String
        Get
            Mrtl = sMrtl
        End Get
        Set(ByVal Value As String)
            sMrtl = Value
        End Set
    End Property

    Public Property BirthDate() As String
        Get
            BirthDate = dBirthdate
        End Get
        Set(ByVal Value As String)
            dBirthdate = Value
        End Set
    End Property

    Public Property Age() As Integer
        Get
            Age = iAge
        End Get
        Set(ByVal Value As Integer)
            iAge = Value
        End Set
    End Property

    Public Property BirthCity() As String
        Get
            BirthCity = sBirthCity
        End Get
        Set(ByVal Value As String)
            sBirthCity = Value
        End Set
    End Property
    Public Property BirthState() As String
        Get
            BirthState = sBirthState
        End Get
        Set(ByVal Value As String)
            sBirthState = Value
        End Set
    End Property
    Public Property BirthCntry() As String
        Get
            BirthCntry = sBirthCntry
        End Get
        Set(ByVal Value As String)
            sBirthCntry = Value
        End Set
    End Property
    Public Property ResState() As String
        Get
            ResState = sResState
        End Get
        Set(ByVal Value As String)
            sResState = Value
        End Set
    End Property
    Public Property ResCity() As String
        Get
            ResCity = sResCity
        End Get
        Set(ByVal Value As String)
            sResCity = Value
        End Set
    End Property
    Public Property Citz() As String
        Get
            Citz = sCitz
        End Get
        Set(ByVal Value As String)
            sCitz = Value
        End Set
    End Property

    Public Property VisaCode() As String
        Get
            VisaCode = sVisaCode
        End Get
        Set(ByVal Value As String)
            sVisaCode = Value
        End Set
    End Property

    Public Property ProfVisaNo() As String
        Get
            ProfVisaNo = sProfVisaNo
        End Get
        Set(ByVal Value As String)
            sProfVisaNo = Value
        End Set
    End Property


    Public Property ProfVisaDate() As String
        Get
            ProfVisaDate = dProfVisaDate
        End Get
        Set(ByVal Value As String)
            dProfVisaDate = Value
        End Set
    End Property
    Public Property ResCntry() As String
        Get
            ResCntry = sResCntry
        End Get
        Set(ByVal Value As String)
            sResCntry = Value
        End Set
    End Property

    Public Property Vet() As String
        Get
            Vet = sVet
        End Get
        Set(ByVal Value As String)
            sVet = Value
        End Set
    End Property
    Public Property HandicapCode() As String
        Get
            HandicapCode = sHandicapCode
        End Get
        Set(ByVal Value As String)
            sHandicapCode = Value
        End Set
    End Property
    Public Property DenomCode() As String
        Get
            DenomCode = sDenomCode
        End Get
        Set(ByVal Value As String)
            sDenomCode = Value
        End Set
    End Property

    Public Property ChurchID() As Integer
        Get
            ChurchID = iChurchID
        End Get
        Set(ByVal Value As Integer)
            iChurchID = Value
        End Set
    End Property
    Public Property Occ() As String
        Get
            Occ = sOcc
        End Get
        Set(ByVal Value As String)
            sOcc = Value
        End Set
    End Property

    Public Property News1ID() As Integer
        Get
            News1ID = iNews1ID
        End Get
        Set(ByVal Value As Integer)
            iNews1ID = Value
        End Set
    End Property

    Public Property News2ID() As Integer
        Get
            News2ID = iNews2ID
        End Get
        Set(ByVal Value As Integer)
            iNews2ID = Value
        End Set
    End Property

    Public Property ProfLastUpd() As String
        Get
            ProfLastUpd = dProfLastUpd
        End Get
        Set(ByVal Value As String)
            dProfLastUpd = Value
        End Set
    End Property

    Public Property ProfVetChap() As String
        Get
            ProfVetChap = sProfVetChap
        End Get
        Set(ByVal Value As String)
            sProfVetChap = Value
        End Set
    End Property
    Public Property GrpNo() As Integer
        Get
            GrpNo = iGrpNo
        End Get
        Set(ByVal Value As Integer)
            iGrpNo = Value
        End Set
    End Property
    Public Property PrivCode() As String
        Get
            PrivCode = sPrivCode
        End Get
        Set(ByVal Value As String)
            sPrivCode = Value
        End Set
    End Property
    Public Property DecsDate() As String
        Get
            DecsDate = dDecsdDate
        End Get
        Set(ByVal Value As String)
            dDecsdDate = Value
        End Set
    End Property
    Public Property MilitCode() As String
        Get
            MilitCode = sMilitCode
        End Get
        Set(ByVal Value As String)
            sMilitCode = Value
        End Set
    End Property
    Public Property MilitRank() As String
        Get
            MilitRank = sMilitRank
        End Get
        Set(ByVal Value As String)
            sMilitRank = Value
        End Set
    End Property
    Public Property Photo() As Byte
        Get
            Photo = bPhoto
        End Get
        Set(ByVal Value As Byte)
            bPhoto = Value
        End Set
    End Property
    Public Property Race() As String
        Get
            Race = sRace
        End Get
        Set(ByVal Value As String)
            sRace = Value
        End Set
    End Property
    Public Property Hispanic() As String
        Get
            Hispanic = sHispanic
        End Get
        Set(ByVal Value As String)
            sHispanic = Value
        End Set
    End Property
    Public Property ConfirmDate() As String
        Get
            ConfirmDate = dConfirmDate
        End Get
        Set(ByVal Value As String)
            dConfirmDate = Value
        End Set
    End Property

    Public Property lang() As String
        Get
            lang = sLang
        End Get
        Set(ByVal Value As String)
            sLang = Value
        End Set
    End Property

    Public Property profResCode() As String
        Get
            profResCode = sprofResCode
        End Get
        Set(ByVal Value As String)
            sprofResCode = Value
        End Set
    End Property
    Public Property profResDate() As String
        Get
            profResDate = dprofResDate
        End Get
        Set(ByVal Value As String)
            dprofResDate = Value
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
        sql = " INSERT INTO profile_rec (id, ethnic_code, sex, race, hispanic, birth_date, news1_id, news2_id, prof_last_upd_date, grp_no) values (?,?,?,?,?,?,?,?,?,?) "
        Dim cmd As New OdbcCommand(sql, ConnInformix)

        Try

            ConnInformix.ConnectionString = gl.ConnectString
            ConnInformix.Open()

            cmd.CommandType = CommandType.Text

            cmd.Parameters.AddWithValue("@id", Me.ID)
            cmd.Parameters.AddWithValue("@ethnic_code", Me.EthnicCode)
            cmd.Parameters.AddWithValue("@sex", Me.Sex)
            cmd.Parameters.AddWithValue("@race", Me.Race)
            cmd.Parameters.AddWithValue("@hispanic", Me.Hispanic)

            cmd.Parameters.AddWithValue("@Bdate", Me.BirthDate)

            cmd.Parameters.AddWithValue("@news1_id", Me.News1ID)
            cmd.Parameters.AddWithValue("@news2_id", Me.News2ID)
            cmd.Parameters.AddWithValue("@prof_last_upd_date", Me.ProfLastUpd)
            cmd.Parameters.AddWithValue("@grp_no", Me.GrpNo)

            ' Dim rowsAffected As Integer = cmd.ExecuteNonQuery()   'Executes query and returns a value
            ' Debug.WriteLine("RowsAffected: {0}", rowsAffected.ToString)
            WriteLog("Create of new Profile ID record for " & Me.FullName & ", ID - " & Me.ID)

            ConnInformix.Close()
        Catch ex As DuplicateNameException
            'No need to write anything
            If ConnInformix.State = ConnectionState.Open Then
                ConnInformix.Close()
                ConnInformix.Dispose()
            End If
        Catch e As Exception
            Debug.WriteLine(e)
            WriteError("Error in clsProfile_rec - Insert.  Error = " & e.Message & ", " & Me.FullName & ", ID = " & Me.ID)
            'MAIL Event
            'sMl.SendEMessage("ADP to CX", (""Error in clsProfile_rec - Insert.  Error = " & e.Message), "dsullivan@carthage.edu", "dsullivan@carthag.edu", "dsullivan@carthage.edu")
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
        '        sql = "Update profile_rec SET ethnic_code = ?,  sex = ? where ID = ?"
        sql = "Update profile_rec set "
        If Me.Sex <> "" Then sql = sql & " sex = ? "
        If Me.EthnicCode <> "" Then sql = sql & ", ethnic_code = ?"
        If Me.Race <> "" Then sql = sql & ", race = ?"
        If Me.Hispanic <> "" Then sql = sql & ", hispanic = ?"
        If Me.BirthDate <> "" Then sql = sql & ", birth_date = ?"
        If Me.ProfLastUpd <> "" Then sql = sql & ", prof_last_upd_date = ?"
        sql = sql & " WHERE ID = ?"
        Dim cmd As New OdbcCommand(sql, ConnInformix)

        Try
            ConnInformix.ConnectionString = gl.ConnectString
            ConnInformix.Open()
            cmd.CommandType = CommandType.Text

            If Me.Sex <> "" Then cmd.Parameters.AddWithValue("@sex", Me.Sex)
            If Me.EthnicCode <> "" Then cmd.Parameters.AddWithValue("@ethnic_code", Me.EthnicCode)
            If Me.Race <> "" Then cmd.Parameters.AddWithValue("@race", Me.Race)
            If Me.Hispanic <> "" Then cmd.Parameters.AddWithValue("@hispanic", Me.Hispanic)
            If Me.BirthDate <> "" Then cmd.Parameters.AddWithValue("@birth_date", Me.BirthDate) 'Ron says necessary for provisioning
            If Me.ProfLastUpd <> "" Then cmd.Parameters.AddWithValue("@prof_last_upd_date", Me.ProfLastUpd)
            cmd.Parameters.AddWithValue("@ID", Me.ID)

            ' Dim rowsAffected As Integer = cmd.ExecuteNonQuery()   'Executes query and returns a value
            ' Debug.WriteLine("RowsAffected: {0}", rowsAffected.ToString)
            WriteLog("Update of Profile record for" & Me.FullName & ", ID =" & Me.ID & ", Race = " & Me.Race & ", Ethnicity = " & Me.EthnicCode)

            ConnInformix.Close()
            ConnInformix.Dispose()
        Catch e As Exception
            Debug.WriteLine(e)
            WriteError("Error in clsProfile_rec - Update.  Error = " & e.Message & ", " & Me.FullName & ", ID = " & Me.ID)
            'MAIL Event
            'sMl.SendEMessage("ADP to CX", (""Error in clsProfile_rec - Update.  Error = " & e.Message), "dsullivan@carthage.edu", "dsullivan@carthag.edu", "dsullivan@carthage.edu")
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
        sEthnicCode = ""
        sSex = ""
        sMrtl = ""
        dBirthdate = ""
        iAge = 0
        sBirthCity = ""
        sBirthState = ""
        sBirthCntry = ""
        sResState = ""
        sResCity = ""
        sCitz = ""
        sVisaCode = ""
        dProfVisaDate = ""
        sProfVisaNo = ""
        sResCntry = ""
        sVet = ""
        sHandicapCode = ""
        sDenomCode = ""
        iChurchID = 0
        sOcc = ""
        iNews1ID = 0
        iNews2ID = 0
        dProfLastUpd = ""
        sLang = ""
        sprofResCode = ""
        dprofResDate = ""
        sProfVetChap = ""
        iGrpNo = 0
        sPrivCode = ""
        dDecsdDate = ""
        sMilitCode = ""
        sMilitRank = ""
        bPhoto = 0
        sRace = ""
        sHispanic = ""
        dConfirmDate = ""
        sFullName = ""

    End Sub
End Class