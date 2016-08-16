Attribute VB_Name = "Mdl_MediaRadio"
'
Public Function User_Valid(What_Valid As String) As Boolean
    Dim TxtSQl As String
    Dim rs As New ADODB.Recordset
    
    TxtSQl = "select * from Media_Security_Catalog  where position = '" & What_Valid & "' and user_name='" & UserName & "' and Valid_until > getdate()"
    
    rs.Open TxtSQl, ConnERP, adOpenStatic
    
    With rs
        User_Valid = Not .EOF
    End With
    '
    Set rs = Nothing
End Function

Public Function List_ComboBox_Position(ByVal What_Combo As ComboBox, What_Code As String, How_Many_Leght_Code As Integer, Left_Function As Boolean) As Integer
'************************************************************
' Procedure         : ListBox_Position(What_Combo As ComboBox, What_Code As String, How_Many_Leght_Code As Integer, Left_Function As Boolean) As Integer
' Function          : To give listbox position from selected Combobox
' Date              : 10/11/2000
' Parameter Input   : What_Combo As ComboBox, What_Code As String, How_Many_Leght_Code As Integer, Left_Function As Boolean
' Parameter Output  : Integer
' Last Update/By    :
'************************************************************

    Dim ListCounter, ListPos As Integer
    
    ListCounter = What_Combo.ListCount
    
    '* check the paramater if it from left or right function
    Select Case Left_Function
        Case Is = True
            For ListPos = 0 To ListCounter - 1
                What_Combo.ListIndex = ListPos
                If Trim(Left(What_Combo.Text, How_Many_Leght_Code)) = Trim(What_Code) Then
                    List_ComboBox_Position = ListPos
                    Exit Function
                End If
            Next ListPos
        Case Is = False
            For ListPos = 0 To ListCounter - 1
                What_Combo.ListIndex = ListPos
                If Trim(Right(What_Combo.Text, How_Many_Leght_Code)) = Trim(What_Code) Then
                    List_ComboBox_Position = ListPos
                    Exit Function
                End If
            Next ListPos
    End Select
End Function

Public Function Check_A_to_Z_and_0_to_9(ByVal Ascii_Code As Integer) As Boolean
    'Called from :
        'Frm Radio Station Catalog
        
    If (Ascii_Code >= 65 And Ascii_Code <= 90) Or (Ascii_Code >= 97 And Ascii_Code <= 122) Or (Ascii_Code >= 48 And Ascii_Code <= 57) Or Ascii_Code = 8 Then
        Check_A_to_Z_and_0_to_9 = True
    Else
        Check_A_to_Z_and_0_to_9 = False
    End If
End Function

Public Function Get_RD_Media_Quotation_No(strBrandCode As String, IntYear As Integer) As String
    Dim rs As New ADODB.Recordset
    Dim strSql As String
    Dim Str_New_MQ As String
    Dim Str_Number As String
    
    strSql = "SELECT MQ_Number FROM Reuseable_MQ_radio "
    strSql = strSql & " WHERE year = " & IntYear
    strSql = strSql & " AND brand_code = '" & strBrandCode & "'"
    strSql = strSql & " ORDER BY mq_number ASC "
    
    rs.Open strSql, ConnERP, adOpenDynamic, adLockOptimistic
    
    Rem Jika ada di Reuseable table
    
    If Not rs.EOF Then
    
        Str_New_MQ = rs.Fields("mq_number")
        
        strSql = "INSERT INTO MQ_Radio_Running (MQ_Number, Year,Brand_Code) VALUES ( "
        strSql = strSql & " '" & Trim(rs.Fields("mq_number")) & "', "
        strSql = strSql & IntYear & ", "
        strSql = strSql & " '" & strBrandCode & "') "
        ConnERP.Execute strSql
        
        strSql = " DELETE FROM reuseable_MQ_radio "
        strSql = strSql & " WHERE year = " & IntYear
        strSql = strSql & " AND Brand_code = '" & strBrandCode & "'"
        strSql = strSql & " AND MQ_Number ='" & Trim(rs.Fields("mq_number")) & "'"
        ConnERP.Execute strSql
        
        Set rs = Nothing
        
        Get_RD_Media_Quotation_No = Str_New_MQ 'Return Value
        
        Exit Function
    Else
    
        Rem Jika Generate New MQ Number
        Set rs = Nothing
        
        strSql = "SELECT MAX(MQ_Number) as MQ_Number FROM MQ_Radio_Running "
        strSql = strSql & " WHERE year = " & IntYear
        strSql = strSql & " AND brand_code = '" & strBrandCode & "'"
        
        rs.Open strSql, ConnERP, adOpenDynamic, adLockOptimistic
        With rs
            If Not .EOF Then
                Rem Generate New MQ_Radio Number
                Str_Number = Trim(str((Val(Right(Trim(IIf(IsNull(.Fields("MQ_Number")) = True, 0, .Fields("MQ_Number"))), 2)) + 1)))
                Str_Number = IIf(Len(Str_Number) = 1, "0" & Str_Number, Str_Number)
                Str_Number = Right(IntYear, 2) & Str_Number
                
                Str_New_MQ = strBrandCode & ".025." & Str_Number
                                                
                Rem Insert new MQ_Radio Number Into Table MQ_Radio_Running
                strSql = "insert into MQ_Radio_Running (MQ_Number, Year,Brand_Code) values ( "
                strSql = strSql & " '" & Trim(Str_New_MQ) & "', "
                strSql = strSql & IntYear & ", "
                strSql = strSql & " '" & strBrandCode & "') "
                ConnERP.Execute strSql
                
                Get_RD_Media_Quotation_No = Str_New_MQ
                
            End If
        End With
        
        Set rs = Nothing
    End If
    
End Function
