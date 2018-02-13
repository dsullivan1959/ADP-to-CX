Public Class clsADPName
    Dim sFullName As String
    Dim sLName As String
    Dim sFName As String
    Dim sMName As String

    'This class is just a convenient place to store logic for breaking down the ADP fullname into separate fields
    Public Property Fullname() As String
        Get
            Fullname = sFullName
        End Get
        Set(ByVal Value As String)
            sFullName = Value
        End Set
    End Property

    Public Property Lname() As String
        Get
            Lname = sLName
        End Get

        Set(ByVal Value As String)
            Dim iComma, iSpace As Integer
            Dim tStr As String
            iComma = 0
            tStr = ""
            iComma = InStr(Value, ",")  'Find the comma between first and last name
            If iComma = 0 Then
                sFName = ""   ' Name is wrong format
            Else
                tStr = Mid(Value, iComma + 2)
                iSpace = InStr(tStr, " ")  'find the space after the start of the first name
                If iSpace = 0 Then  'No middle name
                    sFName = tStr
                Else
                    sFName = Mid(Value, iComma + 2, iSpace - 1)
                End If

            End If

            sLName = Left(Value, InStr(Value, ",") - 1)
        End Set
    End Property
    Public Property Fname() As String
        Get
            Fname = sFName
        End Get
        Set(ByVal Value As String)
            Dim iComma, iSpace As Integer
            Dim tStr As String
            iComma = 0
            tStr = ""
            iComma = InStr(Value, ",")  'Find the comma between first and last name
            If iComma = 0 Then
                sFName = ""   ' Name is wrong format
            Else
                tStr = Mid(Value, iComma + 2)
                iSpace = InStr(tStr, " ")  'find the space after the start of the first name
                If iSpace = 0 Then  'No middle name
                    sFName = tStr
                Else
                    sFName = Mid(Value, iComma + 2, iSpace - 1)
                End If

            End If
        End Set
    End Property
    Public Property MName() As String
        Get
            MName = sMName
        End Get
        Set(ByVal Value As String)
            Dim iComma, iSpace As Integer
            Dim tStr As String
            iComma = 0
            iSpace = 0
            tStr = ""

            iComma = InStr(Value, ",")  'Find the comma between first and last name
            If iComma = 0 Then
                sMName = ""
            Else
                tStr = Mid(Value, iComma + 2)
                iSpace = InStr(tStr, " ")  'find the space after the start of the first name
                If iSpace = 0 Then
                    sMName = ""
                Else
                    sMName = Right(Value, Len(tStr) - iSpace)  'get the remaining part of the name
                End If
            End If

        End Set
    End Property
    Public Sub Initialize()
        sMName = ""
        sLName = ""
        sFName = ""
        sFullName = ""
    End Sub

End Class
