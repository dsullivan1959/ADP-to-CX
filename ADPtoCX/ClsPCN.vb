Public Class ClsPCN

    Dim sPosType As String
    Dim sDiv As String
    Dim sDept As String
    Dim sPosition As String

    'This class has logic for splitting the pcn_aggr field into pcn_01, pcn_02, PCN_03, PC_04
    Public Property PosType() As String
        Get
            PosType = sPosType
        End Get
        Set(ByVal Value As String)
            sPosType = Left(Value, InStr(Value, "-") - 1)
        End Set
    End Property

    Public Property Div() As String
        Get
            Div = sDiv
        End Get
        Set(ByVal Value As String)
            Dim iDash, iDash2 As Integer
            Dim tStr As String
            iDash = 0
            tStr = ""
            iDash = InStr(Value, "-")  'Find the dash between first and second item
            If iDash = 0 Then
                sDiv = ""   ' Data is wrong format
            Else
                tStr = Mid(Value, iDash + 1)
                iDash2 = InStr(tStr, "-")  'find the second "-"
                If iDash2 = 0 Then  'wrong format
                    sDiv = ""
                Else
                    sDiv = Left(tStr, InStr(tStr, "-") - 1)
                End If

            End If
        End Set
    End Property
    Public Property Dept() As String
        Get
            Dept = sDept
        End Get
        Set(ByVal Value As String)
            Dim iDash As Integer
            Dim tStr As String
            Dim x As Integer
            iDash = 0
            tStr = Value
            iDash = InStr(tStr, "-")  'find the "-"

            For x = 1 To 2
                If iDash = 0 Then
                    sDept = ""   ' Data is wrong format
                Else
                    tStr = Mid(tStr, iDash + 1)
                End If
                iDash = InStr(tStr, "-")  'find the next "-"

            Next
            sDept = Left(tStr, InStr(tStr, "-") - 1)

        End Set
    End Property
    Public Property Position() As String
        Get
            Position = sPosition
        End Get
        Set(ByVal Value As String)
            Dim iDash As Integer
            Dim tStr As String
            Dim x As Integer
            iDash = 0
            tStr = Value
            iDash = InStr(tStr, "-")  'find the "-"

            For x = 1 To 3
                If iDash = 0 Then
                    sPosition = ""   ' Data is wrong format
                Else
                    tStr = Mid(tStr, iDash + 1)
                End If
                iDash = InStr(tStr, "-")  'find the next "-"

            Next
            sPosition = tStr
        End Set
    End Property
    Public Sub Initialize()
        sPosType = ""
        sDiv = ""
        sDept = ""
        sPosition = ""
    End Sub

End Class
