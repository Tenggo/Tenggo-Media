Attribute VB_Name = "mdl_Print"
Option Explicit
Public From_PO As Boolean
Public From_Replace As Boolean
Public isFrom_Supplier_TV As Boolean
Public isFrom_Supplier_Print As Boolean
Public isFrom_Supplier_Radio As Boolean
Public isNewNote As Boolean
Public From_Catalog As Boolean
''
Function Get_Material_Print(Brief As String, ib As String, Kode As String) As String
    Dim SqlMtr As String
    Dim rs_Mtr As New ADODB.Recordset
    SqlMtr = "select Material from ib_print_material where client_brief_id='" & Brief & "' and ib_id='" & ib & "' and material_code='" & Kode & "' "
    rs_Mtr.Open SqlMtr, ConnERP, adOpenStatic, adLockReadOnly
    Get_Material_Print = rs_Mtr(0)
    rs_Mtr.Close
    Set rs_Mtr = Nothing
End Function

Function Get_Log_On_Server_Print(strKey As String) As String
    Dim strQuery As String
    Dim recPrintInformation As New ADODB.Recordset
    
    strQuery = "select " & strKey & " from print_information"
    recPrintInformation.Open strQuery, ConnERP, adOpenStatic, adLockReadOnly
    Get_Log_On_Server_Print = Trim(recPrintInformation(0))
    
    recPrintInformation.Close
    Set recPrintInformation = Nothing
End Function


Public Function is_De_Flag(ByVal Brand_code As String) As Boolean
    Dim strSql As String
    Dim rs_brand As New ADODB.Recordset
    
    strSql = "select de_flag from brand where brand_code='" & Trim(Brand_code) & "'"
    rs_brand.Open strSql, ConnERP, adOpenForwardOnly, adLockReadOnly, adCmdText
    
    If Not rs_brand.BOF And Not rs_brand.EOF Then
        If rs_brand(0).Value = 1 Then
            is_De_Flag = True
        Else
            is_De_Flag = False
        End If
    End If
    
    rs_brand.Close
    Set rs_brand = Nothing
    
End Function


Public Function Get_Print_Kode(ByVal Teks As String, ByVal Kode As String) As String
      Get_Print_Kode = Trim(Mid(Trim(Teks), 1, InStr(1, Trim(Teks), Trim(Kode)) - 1))
End Function

Public Function Get_IB_ID_Source_MQ(ByVal Job_Id As String, ByVal CariPanjang As Boolean) As String
      Dim Rs_ib_id As New ADODB.Recordset
      Dim Sql As String
      Sql = "select source_ib from ib_print_quotation_detail where job_id='" & Job_Id & "'"
      Rs_ib_id.Open Sql, ConnERP, adOpenStatic, adLockReadOnly
      
      If CariPanjang Then
          Get_IB_ID_Source_MQ = CStr(Len(Trim(Rs_ib_id(0))))
      Else
            If Not Rs_ib_id.EOF And Not Rs_ib_id.BOF Then
                  If Len(Trim(Rs_ib_id(0))) = 13 Then
                      Get_IB_ID_Source_MQ = Trim(Rs_ib_id(0))
                      Rs_ib_id.Close
                      Set Rs_ib_id = Nothing
                      Exit Function
                  Else
                      
                  End If
            End If
      End If
      Rs_ib_id.Close
      Set Rs_ib_id = Nothing
End Function

Public Function Get_PR_Media_Quotation_No(strBrandCode As String, IntYear As Integer) As String
    Dim Sql As String
    Dim rs_Max As New ADODB.Recordset
    Dim strMediaCode As String
    Dim No As Double
    Dim nomor As String
    Dim rs_Compare As New ADODB.Recordset
    Dim strNewMQNumber As String
    '***************************************
    'prosedur untuk mendapatkan MQ ID Print
    '***************************************
    
    strMediaCode = "015"
    
    'cari dulu di reusable WHERE year and brand
    Sql = "SELECT * FROM Reuseable_IB_ID_mq_Print WHERE year=" & IntYear & " and brand_code = '" & strBrandCode & "' ORDER BY ib_id "
    rs_Max.Open Sql, ConnERP, adOpenStatic, adLockReadOnly, adCmdText
    
    'jika ada StrNewMQNumber = record di reusable
    If rs_Max.RecordCount > 0 Then
    
        strNewMQNumber = rs_Max("ib_id")
        
        'Compare Data reusable dengan last
        'cari di last WHERE year and brand
        Sql = "SELECT last_number FROM Last_IB_ID_mq_Print WHERE year='" & IntYear & "' and brand_code='" & strBrandCode & "'"
        rs_Compare.Open Sql, ConnERP, adOpenStatic, adLockReadOnly, adCmdText
        If rs_Compare.RecordCount = 0 Then
                'jika tidak ada record insert ke last
                Sql = "INSERT INTO Last_IB_ID_mq_Print(brand_code, year, last_number)"
                Sql = Sql & " VALUES ( '" & strBrandCode & "', '" & IntYear & "'," & CInt(Right(Trim(strNewMQNumber), 2)) & ")"
                ConnERP.Execute Sql
        Else
                'jika ada banding apakah last number-nya yang direusable lebih besar
                If rs_Compare(0) < CInt(Right(Trim(strNewMQNumber), 2)) Then
                    Sql = "UPDATE Last_IB_ID_mq_Print set last_number = " & CInt(Right(Trim(strNewMQNumber), 2)) & " WHERE year=" & IntYear & " and brand_code='" & strBrandCode & "'"
                    ConnERP.Execute Sql
                End If
        End If
        
        ' delete di reusable
        Sql = "DELETE FROM Reuseable_IB_ID_mq_Print WHERE year=" & rs_Max("year") & " and Brand_Code= '" & rs_Max("Brand_Code") & "' and IB_ID= '" & rs_Max("IB_ID") & "'"
        ConnERP.Execute Sql
        
        rs_Compare.Close
        Set rs_Compare = Nothing
        
        rs_Max.Close
        Set rs_Max = Nothing
        
        Get_PR_Media_Quotation_No = strNewMQNumber
        
        Exit Function
    End If
    
    rs_Max.Close
    Set rs_Max = Nothing
                    
    'cari nomor WHERE year and brand code
    Sql = "SELECT last_number FROM Last_IB_ID_mq_Print WHERE year=" & IntYear & " and brand_code='" & strBrandCode & "'"
    rs_Max.Open Sql, ConnERP, adOpenStatic, adLockReadOnly
    If rs_Max.RecordCount > 0 Then
         No = rs_Max(0).Value + 1
    Else
         No = 1
    End If
    rs_Max.Close
    Set rs_Max = Nothing
    
    If Len(Trim(No)) = 1 Then nomor = "0" + CStr(No) Else nomor = CStr(No)
        
    strNewMQNumber = strBrandCode & "." & Trim(strMediaCode) & "." & Right(IntYear, 2) & Trim(nomor)
     
'DELETE di last
    Sql = "DELETE Last_IB_ID_mq_Print WHERE brand_code='" & strBrandCode & "' and  year=" & IntYear
    ConnERP.Execute Sql
    
'untuk insert ke last ib print
    Sql = "INSERT INTO Last_IB_ID_mq_Print(brand_code, year, last_number)"
    Sql = Sql & " VALUES ( '" & strBrandCode & "', " & IntYear & "," & CInt(Right(Trim(strNewMQNumber), 2)) & ")"
    ConnERP.Execute Sql
    
    Get_PR_Media_Quotation_No = strNewMQNumber
       
End Function
