Attribute VB_Name = "mdl_Other"
Option Explicit
Public strSQL1 As String
Public strSQL2 As String
Public strsql3 As String
Public Sub rIsiItemCombo2(ByRef Combo As Object, ByVal strSql As String)
    Dim rs As New ADODB.Recordset
    ''
    rs.Open strSql, ConnERP, adOpenStatic, adLockReadOnly
    Combo.Clear
    If Not rs.EOF Then rs.MoveFirst
    While Not rs.EOF
        Combo.AddItem Trim(rs(0)) & " - " & Trim(rs(1))
        rs.MoveNext
    Wend
    rs.Close
    Set rs = Nothing
    
End Sub
Public Sub rIsiItemCombo(ByRef Combo As Object, ByVal strSql As String, Optional satu As Boolean, Optional ByRef Combo2 As Object)
    Dim rs As New ADODB.Recordset
    
    rs.Open strSql, ConnERP, adOpenStatic, adLockReadOnly
    Combo.Clear
    While Not rs.EOF
        Combo.AddItem IIf(IsNull(Combo), "", Trim(rs.Fields(0)))
        If satu Then Combo2.AddItem IIf(IsNull(Combo2), "", Trim(rs.Fields(1)))
        rs.MoveNext
    Wend
    rs.Close
    Set rs = Nothing
End Sub
Function rPotongKiri(strNya As String, ByVal strDicari As String, Optional ByVal lngDikurang As Long) As String
    If InStr(1, strNya, strDicari) = 0 Then
        rPotongKiri = strNya
    Else
        rPotongKiri = Left(strNya, InStr(1, strNya, strDicari) - lngDikurang)
    End If
End Function

Function rPotongKanan(strNya As String, ByVal strDicari As String, ByVal lngDikurang As Long) As String
    If InStr(1, strNya, strDicari) = 0 Then
        rPotongKanan = strNya
    Else
        rPotongKanan = Right(strNya, Len(strNya) - InStr(1, strNya, strDicari) - lngDikurang) ' ( 'Left(strNya, InStr(1, strNya, strDicari) - lngDikurang)
    End If
End Function

Public Function rTesEmpty(ByVal strIsi As String, ByVal strMsg As String, ByVal bol As Boolean, Optional Obj As Object) As Boolean
    If strIsi = "" Then
        MsgBox strMsg & " Could Not Be Empty ", vbOKOnly + vbExclamation, strCompany_Name
        If bol Then Obj.SetFocus
        rTesEmpty = False
    Else
        rTesEmpty = True
    End If
End Function
Public Function rTesNumerik(ByVal KeyAscii As Integer) As Integer
    If (KeyAscii < 48 Or KeyAscii > 57) And KeyAscii <> 8 Then
        rTesNumerik = 7
    Else
        rTesNumerik = KeyAscii
    End If
End Function
Public Function rTesPK(ByVal strSql As String, ByVal col As String) As Boolean
    Dim rsTemp As New ADODB.Recordset
    
    rsTemp.Open strSql, ConnERP, adOpenStatic, adLockReadOnly
    If rsTemp.EOF Then
        MsgBox col & " Already Exists " & vbCrLf & " Cannot Insert Duplicate Value ", vbExclamation, strApplication_Name
        rTesPK = False
    Else
        rTesPK = True
    End If
End Function
Public Function rFormatNumerik(ByVal Lost As Boolean, Angka As String) As Variant
    If Lost Then
        If Angka = "" Or Angka = "0" Then
            rFormatNumerik = "0"
        ElseIf CCur(Angka) < 1 Then
            rFormatNumerik = "0"
        Else
            rFormatNumerik = Format(Angka, "###,###,###,###")
        End If
    Else
        If Angka <> "0" Then
            rFormatNumerik = Format(Angka, "#########")
        Else
            rFormatNumerik = "0"
        End If
    End If
End Function
Public Function rTesCinema(ByVal Textnya As String) As Boolean
    Dim rsTemp As New ADODB.Recordset
  
    rsTemp.Open "Select IndukCinema from OtherMediaType", ConnERP, adOpenStatic, adLockReadOnly
    If Mid(Trim(Textnya), 6, 3) = Trim(rsTemp(0)) Then
        rTesCinema = True
    Else
        rTesCinema = False
    End If
End Function
Public Function rTesULI(ByVal Textnya As String) As Boolean
    'Redundant dengan function Is_Special
    Dim rsTemp As New ADODB.Recordset, strSql As String
  
    strSql = "Select Special_Client_flag from vrSelectBrand where Brand_Code = '" & Left(Textnya, 4) & "'"
    rsTemp.Open strSql, ConnERP, adOpenStatic, adLockOptimistic
    If Not rsTemp.EOF Then
        If Trim(rsTemp(0)) = 1 Then
            rTesULI = True
        Else
            rTesULI = False
        End If
    End If
End Function
Public Function NotEOF(ByVal strSql As String) As Boolean
    Dim rsTemp As New ADODB.Recordset
  
    rsTemp.Open strSql, ConnERP, adOpenStatic, adLockReadOnly
    If Not rsTemp.EOF And Not rsTemp.BOF Then
        NotEOF = True
    Else
        NotEOF = False
    End If
End Function


Public Function Get_OT_Media_Quotation_No(strBrandCode As String, IntYear As Integer, StrMediaType As String) As String
'*************************************************************
'Nama Prosedur      : Get_OT_Media_Quotation_No
'Fungsi Prosedur    : Untuk Megenerate Nomor MQ
'Parameter Input    :
'Parameter Output   :
'Tgl Pembuatan      :
'Last Update/By     : 08/01/02
'*************************************************************
    Dim strSql As String
    Dim rsTemp As New ADODB.Recordset
    Dim StrTempMQNumber As String
    Dim strNewMQNumber As String
    
    On Error GoTo errHand
    
    'connerp.BeginTrans
             
    StrTempMQNumber = strBrandCode & "." & StrMediaType & "." & Right(IntYear, 2)
      
    If rsTemp.State = adStateOpen Then rsTemp.Close
    
    strSql = "SELECT No from REUSEABLE_NO_OTHERs where LEFT(NO,11) = '" & StrTempMQNumber & "' order By NO"
    rsTemp.Open strSql, ConnERP, adOpenStatic, adLockReadOnly
    
    If Not rsTemp.EOF And Not rsTemp.BOF Then
        rsTemp.MoveFirst
        
        strNewMQNumber = Trim(rsTemp("No").Value)
        
        strSql = "DELETE from REUSEABLE_NO_OTHERS where NO = '" & strNewMQNumber & "'"
                
    Else
        strSql = "SELECT NO from LAST_NO_OTHERS where LEFT(NO,11)  = '" & StrTempMQNumber & "'"
    
        If rsTemp.State = adStateOpen Then rsTemp.Close
        rsTemp.Open strSql, ConnERP, adOpenStatic, adLockReadOnly
        
        If Not rsTemp.EOF And Not rsTemp.BOF Then
            
            strSql = Right(Trim(rsTemp(0).Value), 6) + 0.0001
            While Len(strSql) < 6
                strSql = strSql & 0
            Wend
            
            strNewMQNumber = Left(Trim(rsTemp("NO").Value), 7) & strSql
            
            strSql = "UPDATE LAST_NO_OTHERS set NO =  '" & strNewMQNumber & "' where NO = '" & Trim(rsTemp(0)) & "'"
            
        Else    'Begining Number
            strNewMQNumber = StrTempMQNumber & "01"
            
            strSql = "INSERT INTO LAST_NO_OTHERS(NO) values('" & Trim(strNewMQNumber) & "')"
            
        End If
    End If
    
    ConnERP.Execute strSql
    
    If rsTemp.State = adStateOpen Then rsTemp.Close
    
    'connerp.CommitTrans
    
    Get_OT_Media_Quotation_No = strNewMQNumber
    
    Exit Function
    
errHand:
    MsgBox Err.Description, vbOK + vbExclamation, strCompany_Name
    'connerp.RollbackTrans
End Function
