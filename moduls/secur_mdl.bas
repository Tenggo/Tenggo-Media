Attribute VB_Name = "Secur_mdl"
Option Explicit
Dim oTest           As CRijndael
Dim sTemp           As String
Dim bytIn()         As Byte
Dim bytOut()        As Byte
Dim bytPassword()   As Byte
Dim bytClear()      As Byte
Dim lCount          As Long
Dim lLength         As Long
Const saltStr As String = "pandawalima"
Public BoolMaxAgePassword As Boolean
Public StrPasswordMassage As String

Public Function Encrypt(str As String) As String
    Dim outStr As String
    
    Set oTest = New CRijndael
    
    bytIn = str
    bytPassword = saltStr
    bytOut = oTest.EncryptData(bytIn, bytPassword)
    sTemp = ""
    
    For lCount = 0 To UBound(bytOut)
        sTemp = sTemp & Right("0" & Hex(bytOut(lCount)), 2)
    Next
    
    Encrypt = sTemp
End Function
Public Function Decrypt(str As String) As String
Dim outStr As String
    If str = "" Then Decrypt = "": Exit Function
    Set oTest = New CRijndael
    bytIn = str
    bytPassword = saltStr
    lLength = Len(str)
        ReDim bytOut((lLength \ 2) - 1)
        For lCount = 1 To lLength Step 2
            bytOut(lCount \ 2) = CByte("&H" & Mid(str, lCount, 2))
        Next
    bytClear = oTest.DecryptData(bytOut, bytPassword)
    Decrypt = bytClear
End Function

Public Function CanSee(ByVal saUserName As String, ByVal saLinkID As String) As Boolean
'************************************************
' Procedure         : CanSee
' Function          : Mengambil Kode Otoritas PicButton / Icon Bar
' Input Parameter   : -
' OutPut Parameter  : -
' Programmer By     : Tedi
' Date Update       : Jan-2016
' Update By         : Tedi / Kreatif
'************************************************

    Dim strSql As String
    Dim rsTemp As New ADODB.Recordset
    
    strSql = "select SecuritySplit "
    strSql = strSql & "from User_Privillage "
    strSql = strSql & "INNER JOIN User_Menu_Link ON User_Menu_Link.Link_ID=User_Privillage.LinkID "
    strSql = strSql & "WHERE UserName='" & saUserName & "' and Module_ID=3"
    strSql = strSql & "AND User_Menu_Link.LinkName='" & saLinkID & "'"
    
    rsTemp.Open strSql, ConnERP, adOpenForwardOnly, adLockReadOnly, adCmdText: strSql = ""
    
    If Not rsTemp.EOF Then
        If Mid(rsTemp!SecuritySplit, 1, 1) = "0" Or Mid(rsTemp!SecuritySplit, 1, 1) = "x" Then
            CanSee = False
        Else
            CanSee = True
        End If
    Else
        If saUserName = "admin" Then
            CanSee = "111111111111111111"
        Else
            CanSee = "000000000000000000"
        End If
    End If
    
    Call CloseRecordset(rsTemp)

End Function

Public Function GetSecureValue(ByVal saUserName As String, ByVal saLinkID As String) As String
'************************************************
' Procedure         : GetSecureValue
' Function          : Mengambil Semua Nilai Kode Otoritas PicButton / Icon Bar
' Input Parameter   : -
' OutPut Parameter  : -
' Programmer By     : Tedi
' Date Update       : Jan-2016
' Update By         : Tedi / Kreatif
'************************************************

    Dim strSql As String
    Dim rsTemp As New ADODB.Recordset
    
    strSql = "          select SecuritySplit "
    strSql = strSql & "from User_Privillage INNER JOIN User_Menu_Link ON User_Menu_Link.link_ID=User_Privillage.LinkID "
    strSql = strSql & "WHERE UserName='" & saUserName & "' "
    strSql = strSql & "AND User_Menu_Link.LinkName='" & saLinkID & "' "
    strSql = strSql & "AND Module_ID=3 "
    rsTemp.Open strSql, ConnERP, adOpenForwardOnly, adLockReadOnly, adCmdText: strSql = ""
    
    If Not rsTemp.EOF Then
        GetSecureValue = rsTemp!SecuritySplit
    Else
        If UCase(saUserName) = "ADMIN" Then
            GetSecureValue = "1111111111111111111"
        Else
            GetSecureValue = "0000000000000000000"
        End If
    End If
    Call CloseRecordset(rsTemp)
    
End Function
