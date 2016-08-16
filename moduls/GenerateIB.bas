Attribute VB_Name = "GenerateIB"
Option Explicit
Public Function Get_Brief_ID(strBrandCode As String, IntYear As Integer)
'*************************************************************
'Function Name      : Get_Brief_ID
'Function Decription: To Get & Generate Client Brief
'Input Parameter    : StrBrandCode, IntYear
'Output Parameter   : Get_Brief_ID
'Created Date/By    : 27 June 2005/Yayan Royani
'Last Update/By     :
'*************************************************************
''
    Dim Out_Param As ADODB.Parameter
    Dim In_Param1 As ADODB.Parameter
    Dim In_Param2 As ADODB.Parameter
    Dim cmd As New ADODB.Command
    
    Dim rsTemp As New ADODB.Recordset
    Dim strSql As String
    Dim strBriefID As String
        
    On Error GoTo errHandling
    
    If strBrandCode = "" Then
        Get_Brief_ID = "Error"
        Exit Function
    End If
    
    
    'read BriefId
    strSql = "SELECT Client_Brief_Id FROM Client_Brief_Media "
    strSql = strSql & " WHERE Brand_code='" & strBrandCode & "'"
    strSql = strSql & " AND Year=" & IntYear
        
    rsTemp.Open strSql, ConnERP, adOpenStatic, adLockReadOnly
    If Not rsTemp.EOF And Not rsTemp.BOF Then
        strBriefID = Trim(rsTemp(0).Value)
    Else
        'Create Brief ID
        cmd.CommandType = adCmdStoredProc
        cmd.CommandText = "Get_New_Brief_Id_Media"
        
        Set In_Param1 = cmd.CreateParameter("Brand_Code", adChar, adParamInput, 10)
        Set In_Param2 = cmd.CreateParameter("Year", adInteger, adParamInput)
        Set Out_Param = cmd.CreateParameter("New_Brief_Id", adChar, adParamOutput, 10)
           
        cmd.Parameters.Append In_Param1
        cmd.Parameters.Append In_Param2
        cmd.Parameters.Append Out_Param

        In_Param1.Value = strBrandCode
        In_Param2.Value = IntYear
                   
        cmd.ActiveConnection = ConnERP
        cmd.Execute
        
        strBriefID = Out_Param.Value
        
        'Create New Brief.
        strSql = "INSERT INTO Client_Brief_Media (Client_Brief_Id,Brand_Code,Extention,Date_of_Previouse_Issue,"
        strSql = strSql & "Date_Issue,[Year],Status,Country,Approved_Team_Leader,Approved_By_CCM,Entered_Date) VALUES("
        strSql = strSql & "'" & strBriefID & "',"
        strSql = strSql & "'" & strBrandCode & "','Auto',"
        strSql = strSql & "Getdate(),Getdate(),"
        strSql = strSql & IntYear & ","
        strSql = strSql & "' ','Indonesia',0,0,"
        strSql = strSql & "Getdate()"
        strSql = strSql & ")"
        ConnERP.Execute strSql
        
    End If
    
    rsTemp.Close
    Set rsTemp = Nothing
    
    'Return Value
    Get_Brief_ID = strBriefID
    
    Exit Function
errHandling:
     MsgBox Err.Description, vbExclamation, strApplication_Name
End Function

Public Function GetApprovalMedium(StrMPMonthlyActivity As String) As String
'*************************************************************
'Function Name      : GetApprovalMedium
'Function Decription: To Get Approved Medium
'Input Parameter    : StrMPMonthlyActivity
'Output Parameter   : GetApprovalMedium
'Created Date/By    : 29 June 2005/Yayan Royani
'Last Update/By     :
'*************************************************************

    Dim strSql As String
    Dim rsTemp As New ADODB.Recordset
    
     'Populate Variable General
    strSql = " SELECT TOP 1 dbo.MP_Medium.medium_code "
    strSql = strSql & " FROM dbo.MP_Medium INNER JOIN"
    strSql = strSql & " dbo.MP_Monthly_Activity ON dbo.MP_Medium.mp_medium_id = dbo.MP_Monthly_Activity.mp_medium_id"
    strSql = strSql & " WHERE dbo.MP_Monthly_Activity.MP_Medium_ID='" & StrMPMonthlyActivity & "'"
        
    'Open Recordset
    rsTemp.Open strSql, ConnERP, adOpenForwardOnly, adLockReadOnly
    
    'Validate Value
    If rsTemp.EOF Then
        rsTemp.Close
        Set rsTemp = Nothing
        MsgBox "Medium Not Found.", vbExclamation, strApplication_Name
        GetApprovalMedium = "Error"
        Exit Function
    End If
    
    'Return Value
    GetApprovalMedium = rsTemp.Fields(0).Value
    'Close Recordset
    rsTemp.Close
    Set rsTemp = Nothing
        
End Function

Public Function Cancel_IB(StrMPMonthlyActivity As String, PlanMonth As Integer, isCreateQuotation As Boolean) As Boolean
    Dim strSql As String
    Dim StrMPMonthlyActivity_old As String
    Dim RsCheck As New ADODB.Recordset
    Dim strMedium As String
    
    Dim strBrandCode As String
    Dim strMPNumber As String
    Dim intPlanYear As Integer
    Dim IsBU1 As Boolean
    
    On Error GoTo errLbl
       
    IsBU1 = False
    
    'Get Old StrMPMonthlyActivity
    strSql = "SELECT Approved_mp_medium_id FROM MP_Monthly_Activity WHERE mp_medium_id='" & StrMPMonthlyActivity & "' AND month_number=" & PlanMonth
    RsCheck.Open strSql, ConnERP, adOpenKeyset, adLockReadOnly, adCmdText
    If Not RsCheck.EOF Then
        StrMPMonthlyActivity_old = IIf(IsNull(RsCheck.Fields(0).Value), "", RsCheck.Fields(0).Value)
    Else
        StrMPMonthlyActivity_old = ""
    End If
    RsCheck.Close
    
    
    If Trim(StrMPMonthlyActivity_old) = "" Then
        Cancel_IB = True
        Exit Function
    End If
    
    'Get Medium
'    StrSQL = "SELECT Medium_Code FROM MP_Medium WHERE mp_medium_id='" & StrMPMonthlyActivity & "'"
'    RsCheck.Open StrSQL, ConnERP, adOpenKeyset, adLockReadOnly, adCmdText
'    If Not RsCheck.EOF Then
'        strMedium = IIf(IsNull(RsCheck.Fields(0).Value), "", RsCheck.Fields(0).Value)
'    Else
'        strMedium = ""
'    End If
'    RsCheck.Close
       
    'Get Medium,Brancode,month,year,MPNumber
    strSql = "SELECT a.mp_number,a.[year],c.original_brand_code,d.medium_code from mp_master a "
    strSql = strSql & "INNER JOIN mp_task b on a.mp_number = b.mp_number "
    strSql = strSql & "INNER JOIN mp_activity c on b.mp_task_id = c.mp_task_id "
    strSql = strSql & "INNER JOIN mp_medium d on c.mp_activity_id = d.mp_activity_id "
    strSql = strSql & "AND d.mp_medium_id = '" & StrMPMonthlyActivity_old & "'"
    RsCheck.Open strSql, ConnERP, 3, 1
    If Not RsCheck.EOF Then
        strBrandCode = RsCheck("original_brand_code").Value
        strMPNumber = RsCheck("mp_number").Value
        intPlanYear = RsCheck("year").Value
        strMedium = RsCheck("medium_code").Value
    Else
        strBrandCode = ""
        strMPNumber = ""
        intPlanYear = Empty
        strMedium = ""
    End If
    RsCheck.Close
    
    IsBU1 = Is_Special_Brand(strBrandCode) 'Get Bu1 or Bu2 Brand
    
    If Trim(StrMPMonthlyActivity_old) <> "" Then
        Select Case strMedium
        Case "TV"
            strSql = "UPDATE IB_TV SET Status=0,Cancel_Date=Getdate(),Cancel_By='" & Clear_String(strLogin_FullName) & "' WHERE mp_medium_id='" & StrMPMonthlyActivity_old & "' AND month_number=" & PlanMonth
            ConnERP.Execute strSql
            
            If isCreateQuotation And IsBU1 Then
                
                'Loop Sebayak monnth yang ada di IB yang di Cancel
                '======================================================
                strSql = "SELECT IB_ID,Month FROM IB_TV_Montly_Budget WHERE  IB_ID IN(SELECT IB_ID FROM IB_TV WHERE mp_medium_id='" & StrMPMonthlyActivity_old & "' AND month_number=" & PlanMonth & ")"
                RsCheck.Open strSql, ConnERP, 3, 1
                Do While RsCheck.EOF
                    'Generete Quotation TV
                    Call CreateMonthlyQuotation_TV(strBrandCode, RsCheck.Fields("Month").Value, intPlanYear, strMPNumber)
                    
                    RsCheck.MoveNext
                Loop
                    
                RsCheck.Close
            End If
            
        Case "RD"
            strSql = "UPDATE IB_Radio SET Status=0,Cancel_Date=Getdate(),Cancel_By='" & Clear_String(strLogin_FullName) & "' WHERE mp_medium_id='" & StrMPMonthlyActivity_old & "' AND month_number=" & PlanMonth
            ConnERP.Execute strSql
            
            'Generete Quotation RD
             If isCreateQuotation And IsBU1 Then
                Call CreateMonthlyQuotation_RD(strBrandCode, PlanMonth, intPlanYear, strMPNumber)
             End If
             
        Case "PR"
            strSql = "UPDATE IB_Print SET Status=0,Cancel_Date=Getdate(),Cancel_By='" & Clear_String(strLogin_FullName) & "' WHERE mp_medium_id='" & StrMPMonthlyActivity_old & "' AND month_number=" & PlanMonth
            ConnERP.Execute strSql
            
            'Generete Quotation PR
            If isCreateQuotation And IsBU1 Then
                Call CreateMonthlyQuotation_PR(strBrandCode, PlanMonth, intPlanYear, strMPNumber)
            End If
             
        Case "OT"
            strSql = "UPDATE IB_Other SET Status=0,Cancel_Date=Getdate(),Cancel_By='" & Clear_String(strLogin_FullName) & "' WHERE mp_medium_id='" & StrMPMonthlyActivity_old & "' AND month_number=" & PlanMonth
            ConnERP.Execute strSql
            
            'Generete Quotation OT
            If isCreateQuotation And IsBU1 Then
                Call CreateMonthlyQuotation_OT(strBrandCode, PlanMonth, intPlanYear, strMPNumber, True)
            End If
        
        Case "CN"
            strSql = "UPDATE IB_Other SET Status=0,Cancel_Date=Getdate(),Cancel_By='" & Clear_String(strLogin_FullName) & "' WHERE mp_medium_id='" & StrMPMonthlyActivity_old & "' AND month_number=" & PlanMonth
            ConnERP.Execute strSql
            
            'Generete Quotation OT
            If isCreateQuotation And IsBU1 Then
                Call CreateMonthlyQuotation_OT(strBrandCode, PlanMonth, intPlanYear, strMPNumber, False)
            End If
        
        End Select
    End If
    
    Cancel_IB = True
    
    Set RsCheck = Nothing
    
    Exit Function
errLbl:
    Cancel_IB = False
End Function

Public Function ShowIB(StrMPMonthlyActivity As String, PlanMonth As Integer) As String
    Dim strSql As String
    Dim RsMPMonthlyActivity As New ADODB.Recordset
    Dim strIBID As String
    Dim rsIB As New ADODB.Recordset
    
    
    strSql = "SELECT A.mp_medium_id,A.Month_Number,A.Approved_mp_medium_id,M.Medium_Code FROM MP_Monthly_Activity A ,MP_Medium M "
    strSql = strSql & " WHERE A.mp_medium_id='" & StrMPMonthlyActivity & "' AND A.Month_Number=" & PlanMonth
    strSql = strSql & " AND A.mp_medium_id=M.mp_medium_id"
    
    RsMPMonthlyActivity.Open strSql, ConnERP, adOpenKeyset, adLockReadOnly
    If RsMPMonthlyActivity.EOF Then
        RsMPMonthlyActivity.Close
        Set RsMPMonthlyActivity = Nothing
        ShowIB = "Related IB not found."
        Exit Function
    End If
    
    'Check apakah sudah ada Approved Medium ID
    If IsNull(RsMPMonthlyActivity.Fields(2).Value) Then
        RsMPMonthlyActivity.Close
        Set RsMPMonthlyActivity = Nothing
        ShowIB = "Related IB not found."
        Exit Function
    End If
     
    
    Select Case RsMPMonthlyActivity.Fields(3).Value
    Case "TV"
        strSql = "SELECT IB_ID FROM IB_TV "
    Case "PR"
        strSql = "SELECT IB_ID FROM IB_Print "
    Case "RD"
        strSql = "SELECT IB_ID FROM IB_Radio "
    Case "OT"
        strSql = "SELECT IB_ID FROM IB_Other "
    Case "CN"
        strSql = "SELECT IB_ID FROM IB_Other "
    End Select
    
    'Jika sudah ada bandingkan apakah Medium id Sama
    If RsMPMonthlyActivity.Fields(0).Value = RsMPMonthlyActivity.Fields(2).Value Then
        'jika sama IB dari Activity tersebut
        strIBID = " Related IB : " & vbCrLf
    Else
        'Jika Beda IB Dari MP Sebelumnya
        strIBID = " Related IB From Previous Media Plan :" & vbCrLf
    End If
    
    strSql = strSql & " WHERE MP_Medium_ID='" & RsMPMonthlyActivity.Fields(2).Value & "' AND Month_Number=" & RsMPMonthlyActivity.Fields(1).Value & " AND Status=1"
    
    rsIB.Open strSql, ConnERP, adOpenKeyset, adLockReadOnly
    If Not rsIB.EOF Then 'IB Found
        Do While Not rsIB.EOF
            strIBID = strIBID & rsIB.Fields(0).Value & vbCrLf
            rsIB.MoveNext
        Loop
        ShowIB = strIBID
    Else 'NO IB Found
        ShowIB = "Related IB not found."
    End If
    
    rsIB.Close
    Set rsIB = Nothing
    
    RsMPMonthlyActivity.Close
    Set RsMPMonthlyActivity = Nothing
        
End Function
