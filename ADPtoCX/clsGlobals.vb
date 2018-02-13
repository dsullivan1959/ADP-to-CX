Public Class clsGlobals

    Public Shared sSmtpMailHost As String
    Public Shared iSmtpPort As Integer
    Public Shared sConnectString As String
    Public Shared sFileNamePath As String
    Public Shared sFileIn As String
    Public Shared sOldFile As String
    Public Shared sAppPath As String
    Public Shared sToAddress
    Public Shared sSupportAddress
    Public Shared sFromAddress
    Public Shared cJobCollection As New Microsoft.VisualBasic.Collection()

    Public Property SmtpMailHost() As String
        Get
            SmtpMailHost = sSmtpMailHost
        End Get
        Set(ByVal Value As String)
            sSmtpMailHost = Value
        End Set
    End Property

    Public Property SmtpPort() As Integer
        Get
            SmtpPort = iSmtpPort
        End Get
        Set(ByVal Value As Integer)
            iSmtpPort = Value
        End Set
    End Property

    Public Property ConnectString() As String
        Get
            ConnectString = sConnectString
        End Get
        Set(ByVal Value As String)
            sConnectString = Value
        End Set
    End Property

    Public Property FileNamePath() As String
        Get
            FileNamePath = sFileNamePath
        End Get
        Set(ByVal Value As String)
            sFileNamePath = Value
        End Set
    End Property

    Public Property AppPath() As String
        Get
            AppPath = sAppPath
        End Get
        Set(ByVal Value As String)
            sAppPath = Value
        End Set
    End Property

    Public Property FileIn() As String
        Get
            FileIn = sFileIn
        End Get
        Set(ByVal Value As String)
            sFileIn = Value
        End Set
    End Property

    Public Property OldFile() As String
        Get
            OldFile = sOldFile
        End Get
        Set(ByVal Value As String)
            sOldFile = Value
        End Set
    End Property

    Public Property ToAddress() As String
        Get
            ToAddress = sToAddress
        End Get
        Set(ByVal Value As String)
            sToAddress = Value
        End Set
    End Property
    Public Property SupportAddress() As String
        Get
            SupportAddress = sSupportAddress
        End Get
        Set(ByVal Value As String)
            sSupportAddress = Value
        End Set
    End Property
    Public Property FromAddress() As String
        Get
            FromAddress = sFromAddress
        End Get
        Set(ByVal Value As String)
            sFromAddress = Value
        End Set
    End Property


    Public Property JobCollection() As Collection
        Get
            JobCollection = cJobCollection
        End Get
        Set(ByVal Value As Collection)
            cJobCollection = Value
        End Set
    End Property

End Class
