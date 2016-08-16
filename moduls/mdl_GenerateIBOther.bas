Attribute VB_Name = "mdl_GenerateIBOther"
 Option Explicit
 Public Function Generate_IB_Others(StrMPMonthlyActivity As String, PlanMonth As Integer) As String
'*************************************************************
'Function Name      : Generate_IB_Others
'Function Decription: To Generate IB for Others & Cinema
'Input Parameter    : StrMPMonthlyActivity, PlanMonth
'Output Parameter   : Generate_IB_Others
'Created Date/By    : 26 June 2005/Yayan Royani
'Last Update/By     :
'*************************************************************
    ''
    'Declare Local Variable
    Dim strSql As String
    Dim BrandCode As String
    Dim BrandName As String
    Dim BrandVariantCode As String
    Dim BrandVariantName As String
    Dim PlanNo As String
    Dim PlanYear As Integer
    Dim StrPrimaryTarget As String
    Dim StrSecondaryTarget As String
    Dim SecondaryTargetCode As Integer
    Dim ActivityTypeCode As String
    Dim ActivityDescription As String
    Dim MediumCode As String
    Dim MonthlyGrossBudget As Currency
    Dim MonthlyNettBudget As Currency
    Dim MonthlyFeeBudget As Currency
    
    Dim IsOther As Boolean
    Dim IB_ID As String
    Dim Brief_Id As String
    Dim StrDescription As String
    Dim TOTAL As Currency
    
    Dim rsMPInfoGeneral As New ADODB.Recordset
    Dim rsMPInfoOther As New ADODB.Recordset
    Dim rsMPInfoCinema As New ADODB.Recordset
    Dim rsMPInfoCinemaDetail As New ADODB.Recordset
    Dim rsBrandVariant As New ADODB.Recordset
    
    On Error GoTo errLbl
    
    'Cancel Previous IB
    If Not Cancel_IB(StrMPMonthlyActivity, PlanMonth, False) Then
        Generate_IB_Others = "Error"
        Exit Function
    End If
    
    
    'Populate Variable General
    strSql = " SELECT dbo.MP_Master.mp_number,dbo.MP_Master.Year, dbo.MP_Activity.Original_Brand_Code, dbo.MP_Activity.Original_Brand_Name, dbo.MP_Activity.activity_type, dbo.MP_Activity.activity_desc, "
    strSql = strSql & " dbo.MP_Activity.brand_variant_code, dbo.MP_Activity.brand_variant_name, dbo.MP_Activity.target_audience_code, dbo.MP_Activity.target_audience,"
    strSql = strSql & " dbo.MP_Activity.brand_target, dbo.MP_Medium.medium_code, dbo.MP_Monthly_Activity.month_number, dbo.MP_Monthly_Activity.month_name,"
    strSql = strSql & " dbo.MP_Monthly_Activity.Budget , dbo.MP_Monthly_Activity.gross_budget, dbo.MP_Monthly_Activity.MSC_Paid_Value"
    strSql = strSql & " FROM dbo.MP_Master INNER JOIN"
    strSql = strSql & " dbo.MP_Task ON dbo.MP_Master.mp_number = dbo.MP_Task.mp_number INNER JOIN"
    strSql = strSql & " dbo.MP_Activity ON dbo.MP_Task.mp_task_id = dbo.MP_Activity.mp_task_id INNER JOIN"
    strSql = strSql & " dbo.MP_Medium ON dbo.MP_Activity.mp_activity_id = dbo.MP_Medium.mp_activity_id INNER JOIN"
    strSql = strSql & " dbo.MP_Monthly_Activity ON dbo.MP_Medium.mp_medium_id = dbo.MP_Monthly_Activity.mp_medium_id"
    strSql = strSql & " WHERE dbo.MP_Monthly_Activity.MP_Medium_ID='" & StrMPMonthlyActivity & "'"
    strSql = strSql & " AND dbo.MP_Monthly_Activity.Month_Number=" & PlanMonth
    
    rsMPInfoGeneral.Open strSql, ConnERP, adOpenForwardOnly, adLockReadOnly
    If rsMPInfoGeneral.EOF Then
        rsMPInfoGeneral.Close
        Set rsMPInfoGeneral = Nothing
    
        Generate_IB_Others = "Error"
        Exit Function
    End If
    '============================================
    BrandCode = IIf(IsNull(rsMPInfoGeneral.Fields("Original_Brand_Code").Value), "", rsMPInfoGeneral.Fields("Original_Brand_Code").Value)
    BrandName = IIf(IsNull(rsMPInfoGeneral.Fields("Original_Brand_Name").Value), "", rsMPInfoGeneral.Fields("Original_Brand_Name").Value)
    
    BrandVariantCode = IIf(IsNull(rsMPInfoGeneral.Fields("Brand_Variant_Code").Value), "", rsMPInfoGeneral.Fields("Brand_Variant_Code").Value)
    BrandVariantName = IIf(IsNull(rsMPInfoGeneral.Fields("Brand_Variant_Name").Value), "", rsMPInfoGeneral.Fields("Brand_Variant_Name").Value)
    
    strSql = "SELECT Original_Brand_Variant_Code,Original_Brand_Variant_Name FROM Brand_Variant"
    strSql = strSql & " WHERE Brand_Variant_Code='" & BrandVariantCode & "'"
    
    BrandVariantCode = ""
    BrandVariantName = ""
    
    rsBrandVariant.Open strSql, ConnERP, adOpenKeyset, adLockOptimistic
    If Not rsBrandVariant.EOF Then
        BrandVariantCode = IIf(IsNull(rsBrandVariant.Fields("Original_Brand_Variant_Code").Value), "", rsBrandVariant.Fields("Original_Brand_Variant_Code").Value)
        BrandVariantName = IIf(IsNull(rsBrandVariant.Fields("Original_Brand_Variant_Name").Value), "", rsBrandVariant.Fields("Original_Brand_Variant_Name").Value)
    End If
    
    rsBrandVariant.Close
    Set rsBrandVariant = Nothing
    '=======================================
    
    PlanNo = rsMPInfoGeneral.Fields("MP_Number").Value
    PlanYear = rsMPInfoGeneral.Fields("Year").Value
    ActivityTypeCode = IIf(IsNull(rsMPInfoGeneral.Fields("Activity_Type").Value), "", rsMPInfoGeneral.Fields("Activity_Type").Value)
    ActivityDescription = IIf(IsNull(rsMPInfoGeneral.Fields("Activity_Desc").Value), "", rsMPInfoGeneral.Fields("Activity_Desc").Value)
    StrPrimaryTarget = IIf(IsNull(rsMPInfoGeneral.Fields("Brand_Target").Value), "", rsMPInfoGeneral.Fields("Brand_Target").Value)
    StrSecondaryTarget = rsMPInfoGeneral.Fields("Target_Audience").Value
    SecondaryTargetCode = rsMPInfoGeneral.Fields("Target_Audience_Code").Value
    MediumCode = rsMPInfoGeneral.Fields("Medium_Code").Value
    MonthlyGrossBudget = rsMPInfoGeneral.Fields("Gross_Budget").Value
    MonthlyNettBudget = rsMPInfoGeneral.Fields("Budget").Value
    MonthlyFeeBudget = rsMPInfoGeneral.Fields("Msc_Paid_Value").Value
    
    rsMPInfoGeneral.Close
    Set rsMPInfoGeneral = Nothing
    '--------------------------------------------------
            
    If UCase(MediumCode) = "OT" Then
        IsOther = True
    Else
        IsOther = False
    End If
    
    'Get Brief ID
    Brief_Id = Get_Brief_ID(BrandCode, PlanYear)
    'Generate IB ID
    IB_ID = GetIB_ID_Others(BrandCode, PlanYear, IsOther)
    
    
    '------------ Insert Data to IB Table--------------------------
    If IsOther Then 'Other
        'Populate Variable Other & Cinema
        strSql = "SELECT dbo.MP_Plan_Dimension.OT_Description, dbo.mp_other_monthly_budget.nett_budget, dbo.mp_other_monthly_budget.gross_budget"
        strSql = strSql & " FROM dbo.MP_Medium INNER JOIN"
        strSql = strSql & " dbo.MP_Medium_Detail ON dbo.MP_Medium.mp_medium_id = dbo.MP_Medium_Detail.mp_medium_id INNER JOIN"
        strSql = strSql & " dbo.MP_Plan_Dimension ON dbo.MP_Medium_Detail.mp_medium_detail_id = dbo.MP_Plan_Dimension.mp_medium_detail_id INNER JOIN"
        strSql = strSql & " dbo.mp_other_monthly_budget ON dbo.MP_Plan_Dimension.mp_plan_dim_id = dbo.mp_other_monthly_budget.mp_plan_dim_id"
        strSql = strSql & " WHERE dbo.mp_other_monthly_budget.Month_Number=" & PlanMonth
        strSql = strSql & " AND dbo.MP_Medium.MP_Medium_id IN(SELECT dbo.MP_Monthly_Activity.mp_medium_id FROM dbo.MP_Monthly_Activity "
        strSql = strSql & " WHERE dbo.MP_Monthly_Activity.MP_Medium_ID='" & StrMPMonthlyActivity & "')"
        
       
        rsMPInfoOther.Open strSql, ConnERP, adOpenForwardOnly, adLockReadOnly
        If Not rsMPInfoOther.EOF Then
            Do While Not rsMPInfoOther.EOF
                StrDescription = StrDescription & rsMPInfoOther.Fields("OT_Description").Value & vbCrLf
                StrDescription = StrDescription & "Nett Budget : " & rsMPInfoOther.Fields("Nett_Budget").Value & vbCrLf & vbCrLf
                rsMPInfoOther.MoveNext
            Loop
        Else
            StrDescription = "Auto Generate from Media Plan"
        End If
        
        TOTAL = MonthlyNettBudget + MonthlyFeeBudget 'Nett + Fee
        
        rsMPInfoOther.Close
        Set rsMPInfoOther = Nothing
        '-------------------------------------------------
        
        strSql = "INSERT INTO IB_Other(IB_ID,CLIENT_BRIEF_ID,DATE,ENTERED_DATE,ENTERED_BY,PRIMARY_TARGET,SECONDARY_TARGET,"
        strSql = strSql & " PLANN,Approval_Client_Flag,Approval_Date,GRAND_TOTAL,NOTE,BRAND_VARIANT_CODE,BRAND_VARIANT_NAME,PLAN_NO,CLUSTER_CODE,mp_medium_id,month_number) "
        strSql = strSql & " VALUES('" & IB_ID & "','" & Brief_Id & "',"
        strSql = strSql & " Getdate(),Getdate(),'" & Clear_String(strLogin_FullName) & " ','"
        strSql = strSql & StrPrimaryTarget & "','" & StrSecondaryTarget & "','" & Clear_String(StrDescription) & "',"
        strSql = strSql & "1,Getdate()," & TOTAL & ",' ','" & BrandVariantCode & "','"
        strSql = strSql & BrandVariantName & "','" & PlanNo & "','" & SecondaryTargetCode & "',"
        'MP_Medium_Id
        strSql = strSql & "'" & StrMPMonthlyActivity & "',"
        'Month_Number
        strSql = strSql & PlanMonth
        strSql = strSql & ")"
        
        ConnERP.Execute strSql
        
    Else 'Cinema
           
        strSql = "SELECT dbo.MP_Medium_Detail.cinema_code, dbo.MP_Medium_Detail.cinema_name, dbo.MP_Medium_Detail.cinema_studio, "
        strSql = strSql & " dbo.MP_Medium_Detail.mp_medium_detail_id , dbo.MP_Medium_Detail.medium_code"
        strSql = strSql & " FROM dbo.MP_Monthly_Activity INNER JOIN"
        strSql = strSql & " dbo.MP_Medium ON dbo.MP_Monthly_Activity.mp_medium_id = dbo.MP_Medium.mp_medium_id INNER JOIN"
        strSql = strSql & " dbo.MP_Medium_Detail ON dbo.MP_Medium.mp_medium_id = dbo.MP_Medium_Detail.mp_medium_id"
        strSql = strSql & " WHERE dbo.MP_Monthly_Activity.MP_Medium_ID='" & StrMPMonthlyActivity & "'"
        strSql = strSql & " AND dbo.MP_Monthly_Activity.Month_Number=" & PlanMonth
        strSql = strSql & " ORDER BY Cinema_Code"
        
        rsMPInfoCinema.Open strSql, ConnERP, adOpenForwardOnly, adLockReadOnly
        If Not rsMPInfoCinema.EOF Then
            Do While Not rsMPInfoCinema.EOF
                If Trim(rsMPInfoCinema.Fields("Cinema_Code").Value) = "" Then 'Brief
                    StrDescription = StrDescription & GetCinemaDesc(rsMPInfoCinema.Fields("mp_medium_detail_id").Value, PlanMonth, True)
                    'StrDescription = StrDescription & vbCrLf
                Else 'Detail per Cinema
                    StrDescription = StrDescription & GetCinemaDesc(rsMPInfoCinema.Fields("mp_medium_detail_id").Value, PlanMonth, False)
                    'StrDescription = StrDescription & vbCrLf
                End If
                
                rsMPInfoCinema.MoveNext
            Loop
        Else
            StrDescription = StrDescription & "Auto Generate from Media Plan"
        End If
        
        TOTAL = MonthlyNettBudget + MonthlyFeeBudget 'Nett + Fee
        
        rsMPInfoCinema.Close
        Set rsMPInfoCinema = Nothing
       
        
        'Insert Data into IB
        strSql = "INSERT INTO IB_Other(IB_ID,CLIENT_BRIEF_ID,DATE,ENTERED_DATE,ENTERED_BY,PRIMARY_TARGET,SECONDARY_TARGET,"
        strSql = strSql & " PLANN,Approval_Client_Flag,Approval_Date,GRAND_TOTAL,NOTE,BRAND_VARIANT_CODE,BRAND_VARIANT_NAME,PLAN_NO,CLUSTER_CODE,mp_medium_id,month_number) "
        strSql = strSql & " VALUES('" & IB_ID & "','" & Brief_Id & "',"
        strSql = strSql & " Getdate(),Getdate(),'" & Clear_String(strLogin_FullName) & " ','"
        strSql = strSql & StrPrimaryTarget & "','" & StrSecondaryTarget & "','" & Clear_String(StrDescription) & "',"
        strSql = strSql & "1,Getdate()," & TOTAL & ",' ','" & BrandVariantCode & "','"
        strSql = strSql & BrandVariantName & "','" & PlanNo & "','" & SecondaryTargetCode & "',"
        'MP_Medium_Id
        strSql = strSql & "'" & StrMPMonthlyActivity & "',"
        'Month_Number
        strSql = strSql & PlanMonth
        strSql = strSql & ")"
        
        ConnERP.Execute strSql
    
    End If
    
      '-------------- Submit Data to Budget Control Plan & Generate Monthly Quotation If BU1 ------------------
    If Is_Special_Brand(BrandCode) Then
        'SubmitDataToMQTV(IBID)--->Generate Quotation
        
        If IsOther Then
            'Other Medium
            Call CreateMonthlyQuotation_OT(BrandCode, PlanMonth, PlanYear, PlanNo, True)
        Else
            'Cinema
            Call CreateMonthlyQuotation_OT(BrandCode, PlanMonth, PlanYear, PlanNo, False)
        End If
        
        
'        'SubmitdatatoBCMPBU1(Brand_code,Year,month,Medium,Gross,Nett)
'        StrSQL = "INSERT INTO Budget_Control_Detail_BU1(mp_medium_id,Month_Number,Year_Budget,Month_Budget,Medium,Brand_code,Gross_Budget,Nett_Budget) VALUES ("
'        'mp_medium_id,Month_Number
'        StrSQL = StrSQL & "'" & StrMPMonthlyActivity & "'," & PlanMonth & ","
'        'Year_Budget,Month_Number
'        StrSQL = StrSQL & PlanYear & "," & PlanMonth & ","
'        'Medium, Brand_Code
'        StrSQL = StrSQL & "'OT','" & BrandCode & "',"
'        'Gross_Budget, Nett_Bugdet
'        StrSQL = StrSQL & "0," & MonthlyNettBudget
'        StrSQL = StrSQL & ")"
'        ConnERP.Execute StrSQL
    Else
'        'SubmitdatatoBCMPBU2(Brand_code,Year,month,Medium,Gross,Nett)
'        StrSQL = "INSERT INTO Budget_Control_Detail_BU2(mp_medium_id,Month_Number,Year_Budget,Month_Budget,Medium,Brand_code,Gross_Budget,Nett_Budget) VALUES ("
'        'mp_medium_id,Month_Number
'        StrSQL = StrSQL & "'" & StrMPMonthlyActivity & "'," & PlanMonth & ","
'        'Year_Budget,Month_Number
'        StrSQL = StrSQL & PlanYear & "," & PlanMonth & ","
'        'Medium, Brand_Code
'        StrSQL = StrSQL & "'OT','" & BrandCode & "',"
'        'Gross_Budget, Nett_Bugdet
'        StrSQL = StrSQL & "0," & MonthlyNettBudget
'        StrSQL = StrSQL & ")"
'        ConnERP.Execute StrSQL
    End If
    '--------------------------------------------------------------------
        
    'Return Value
    Generate_IB_Others = IB_ID
    
    'Add Cuurent User Job
    If IsOther Then
        Add_Current_User_Job 4, strLogin_FullName, IB_ID, "", "", "", "", "", ""
    Else
        Add_Current_User_Job 5, strLogin_FullName, IB_ID, "", "", "", "", "", ""
    End If
    
    Exit Function
errLbl:
    Generate_IB_Others = "Error"
    
End Function

Public Function GetIB_ID_Others(strBrandCode As String, IntYear As Integer, IsOther As Boolean) As String
'*************************************************************
'Function Name      : GetIB_ID_Others
'Function Decription: To Generate IB ID Others & Cinema
'Input Parameter    : StrBrandCode, IntYear, IsOther
'Output Parameter   : GetIB_ID_Others
'Created Date/By    : 26 June 2005/Yayan Royani
'Last Update/By     :
'*************************************************************
    
    Dim rsTemp As New ADODB.Recordset
    Dim strSql As String
    Dim TempBrief As String
    Dim strTemp As String
                
    On Error GoTo errHandling
    
    If strBrandCode = "" Then
        GetIB_ID_Others = "Error"
        Exit Function
    End If
    
    'Open Transaction
    ConnERP.BeginTrans
    
    'Open Media Type Table for Other
    strSql = "SELECT IndukOther, IndukCinema FROM OtherMediaType "
    rsTemp.Open strSql, ConnERP, adOpenStatic, adLockReadOnly
            
    If IsOther Then
        TempBrief = strBrandCode & "." & rsTemp("IndukOther") & "." & Right(IntYear, 2)
    Else
        TempBrief = strBrandCode & "." & rsTemp("IndukCinema") & "." & Right(IntYear, 2)
    End If
    
    strSql = "SELECT IB_ID FROM REUSEABLE_IB_ID_OTHERS WHERE LEFT(IB_ID,11) = '" & TempBrief & "' ORDER BY IB_ID"
    
    If rsTemp.State = adStateOpen Then rsTemp.Close
    
    rsTemp.Open strSql, ConnERP, adOpenStatic, adLockReadOnly
    If Not rsTemp.EOF And Not rsTemp.BOF Then
        rsTemp.MoveFirst
        GetIB_ID_Others = Trim(rsTemp(0).Value)
        strSql = "DELETE FROM REUSEABLE_IB_ID_OTHERS WHERE IB_ID = '" & Trim(rsTemp(0).Value) & "'"
        ConnERP.Execute strSql
        
        
    Else
        strSql = "SELECT IB_ID FROM LAST_IB_ID_OTHERS WHERE LEFT(IB_ID,11)  = '" & TempBrief & "'"
        
        rsTemp.Close
        
        rsTemp.Open strSql, ConnERP, adOpenStatic, adLockReadOnly
        If Not rsTemp.EOF And Not rsTemp.BOF Then
            strTemp = Right(Trim(rsTemp(0)), 6) + 0.0001
            While Len(strTemp) < 6
                strTemp = strTemp & 0
            Wend
            GetIB_ID_Others = Left(Trim(rsTemp(0)), 7) & strTemp
            TempBrief = Left(Trim(rsTemp(0)), 7) & strTemp
            strSql = "UPDATE LAST_IB_ID_OTHERS SET IB_ID =  '" & TempBrief & "' WHERE IB_ID = '" & Trim(rsTemp(0)) & "'"
            ConnERP.Execute strSql
            
        Else
            GetIB_ID_Others = TempBrief & "01"
            TempBrief = TempBrief & "01"
            strSql = "INSERT INTO LAST_IB_ID_OTHERS(IB_ID) values('" & TempBrief & "')"
            ConnERP.Execute strSql
            
        End If
        
        
    End If
    
    If rsTemp.State = adStateOpen Then rsTemp.Close
    Set rsTemp = Nothing
    
    ConnERP.CommitTrans
    
    Exit Function
errHandling:
    MsgBox Err.Description, vbExclamation, strApplication_Name
    ConnERP.RollbackTrans
End Function

Private Function GetCinemaDesc(strMPMediumDetailID As String, intMonth As Integer, isBrief As Boolean) As String
'*************************************************************
'Function Name      : GetCinemaDesc
'Function Decription: To Get Brief Desc for Cinema
'Input Parameter    : StrMPMediumDetailID, isBrief
'Output Parameter   : GetCinemaDesc
'Created Date/By    : 30 June 2005/Yayan Royani
'Last Update/By     :
'*************************************************************
    Dim strSql As String
    Dim rsCinameInsertion As New ADODB.Recordset
    Dim TmpDesc As String

    If isBrief Then 'Only Brief
        strSql = "SELECT  dbo.MP_Plan_Dimension.mp_plan_dim_id,dbo.MP_Plan_Dimension.OT_Description, dbo.MP_Insertion.week_commencing, dbo.MP_Insertion.[month], dbo.MP_Insertion.week_year, "
        strSql = strSql & " dbo.MP_Insertion.Spot , dbo.MP_Insertion.Nett_Rate, dbo.MP_Insertion.Gross_Rate"
        strSql = strSql & " FROM dbo.MP_Medium_Detail INNER JOIN"
        strSql = strSql & " dbo.MP_Plan_Dimension ON dbo.MP_Medium_Detail.mp_medium_detail_id = dbo.MP_Plan_Dimension.mp_medium_detail_id INNER JOIN"
        strSql = strSql & " dbo.MP_Insertion ON dbo.MP_Plan_Dimension.mp_plan_dim_id = dbo.MP_Insertion.mp_plan_dim_id"
        strSql = strSql & " WHERE dbo.MP_Medium_Detail.mp_medium_detail_id='" & strMPMediumDetailID & "' AND month=" & intMonth
        strSql = strSql & " ORDER BY  dbo.MP_Plan_Dimension.mp_plan_dim_id"
        
        rsCinameInsertion.Open strSql, ConnERP, adOpenForwardOnly, adLockReadOnly
        If Not rsCinameInsertion.EOF Then
            Do While Not rsCinameInsertion.EOF
                TmpDesc = TmpDesc & rsCinameInsertion.Fields("OT_Description").Value
                TmpDesc = TmpDesc & vbCrLf & "Week Commecing : " & Format(rsCinameInsertion.Fields("Week_Commencing").Value, "MM/DD/YYYY")
                TmpDesc = TmpDesc & " Gross : " & Format(rsCinameInsertion.Fields("Gross_Rate").Value, "#,##0")
                TmpDesc = TmpDesc & " Nett  : " & Format(rsCinameInsertion.Fields("Nett_Rate").Value, "#,##0")
                TmpDesc = TmpDesc & vbCrLf
                
                rsCinameInsertion.MoveNext
            Loop
        End If
    Else 'Detail Per Cinema
        strSql = "SELECT dbo.MP_Insertion.week_commencing, dbo.MP_Insertion.[month], dbo.MP_Insertion.week_year, dbo.MP_Insertion.spot, dbo.MP_Insertion.nett_rate, "
        strSql = strSql & " dbo.MP_Insertion.gross_rate, dbo.MP_Plan_Dimension.mp_plan_dim_id, dbo.MP_Medium_Detail.cinema_code,"
        strSql = strSql & " dbo.MP_Medium_Detail.cinema_name, dbo.MP_Medium_Detail.cinema_studio, dbo.MP_Plan_Dimension.cinema_duration,"
        strSql = strSql & " dbo.MP_Plan_Dimension.Version , dbo.MP_Plan_Dimension.Duration, dbo.MP_Medium_Detail.mp_medium_detail_id"
        strSql = strSql & " FROM dbo.MP_Medium_Detail INNER JOIN"
        strSql = strSql & " dbo.MP_Plan_Dimension ON dbo.MP_Medium_Detail.mp_medium_detail_id = dbo.MP_Plan_Dimension.mp_medium_detail_id INNER JOIN"
        strSql = strSql & " dbo.MP_Insertion ON dbo.MP_Plan_Dimension.mp_plan_dim_id = dbo.MP_Insertion.mp_plan_dim_id"
        strSql = strSql & " WHERE MP_Medium_Detail.mp_medium_detail_id='" & strMPMediumDetailID & "' AND MONTH=" & intMonth
        strSql = strSql & " ORDER BY  dbo.MP_Plan_Dimension.mp_plan_dim_id"
        
        rsCinameInsertion.Open strSql, ConnERP, adOpenForwardOnly, adLockReadOnly
        If Not rsCinameInsertion.EOF Then
            Do While Not rsCinameInsertion.EOF
                TmpDesc = TmpDesc
                'MsgBox Len(Trim(rsCinameInsertion.Fields("Cinema_Name").Value))
                TmpDesc = TmpDesc & Trim(rsCinameInsertion.Fields("Cinema_Name").Value) & Space(25 - Len(Trim(rsCinameInsertion.Fields("Cinema_Name").Value))) & vbTab
                TmpDesc = TmpDesc & rsCinameInsertion.Fields("Cinema_studio").Value & Space(5 - Len(Trim(rsCinameInsertion.Fields("Cinema_studio").Value)))
                TmpDesc = TmpDesc & rsCinameInsertion.Fields("Cinema_duration").Value & vbTab
                TmpDesc = TmpDesc & rsCinameInsertion.Fields("Version").Value & vbTab
                TmpDesc = TmpDesc & rsCinameInsertion.Fields("duration").Value & vbTab
                TmpDesc = TmpDesc & rsCinameInsertion.Fields("Week_Commencing").Value & vbTab
                TmpDesc = TmpDesc & rsCinameInsertion.Fields("Spot").Value & vbTab
                TmpDesc = TmpDesc & Format(rsCinameInsertion.Fields("Gross_Rate").Value, "#,##0") & vbTab & vbTab
                TmpDesc = TmpDesc & Format(rsCinameInsertion.Fields("Nett_Rate").Value, "#,##0") & vbTab
                TmpDesc = TmpDesc & vbCrLf
                                
                rsCinameInsertion.MoveNext
            Loop
        End If
    End If
   
    rsCinameInsertion.Close
    Set rsCinameInsertion = Nothing
        

    GetCinemaDesc = TmpDesc
End Function

Public Function CreateMonthlyQuotation_OT(strBrandCode As String, IntMonthQuotation As Integer, IntYearQuotation As Integer, strMPNumber As String, IsOther As Boolean) As Boolean
    'Decalare Local Variable
    Dim strSql As String
    Dim rsTemp As New ADODB.Recordset
    Dim CurMonthlyBudget As Currency
    Dim RsOldMediaQuotation As New ADODB.Recordset
    Dim IsOldmediaQuotation_Exists As Boolean
    Dim Rs_Revision As New ADODB.Recordset
    
    Dim Revision As Integer
    Dim StrJobId As String
    Dim NewMQNumber As String
    Dim StrSourceIB As String
    Dim StrMediaType As String
    
    
    On Error GoTo errLbl
    
    CurMonthlyBudget = 0
    If IsOther Then
        StrMediaType = "065"
        
        strSql = "SELECT IB_ID,Grand_Total FROM IB_Other WHERE "
        strSql = strSql & " LEFT(IB_ID,8)='" & strBrandCode & "." & StrMediaType & "' AND Substring(IB_ID,10,2)='" & Right(IntYearQuotation, 2) & "' AND Status=1 AND mp_medium_id IS NOT NULL AND Month_Number=" & IntMonthQuotation
    Else
        StrMediaType = "035"
        
        strSql = "SELECT IB_ID,Grand_Total FROM IB_Other WHERE "
        strSql = strSql & " LEFT(IB_ID,8)='" & strBrandCode & "." & StrMediaType & "' AND Substring(IB_ID,10,2)='" & Right(IntYearQuotation, 2) & "' AND Status=1 AND mp_medium_id IS NOT NULL AND Month_Number=" & IntMonthQuotation
    End If
    
    'Get Value  Money Selected Month From Monthly IB (IB Lama Tidak Bisa Tidak Keambil)
    'Dari IB Yang statunya Aktif dan Month dan Year-nya Sama
   
    StrSourceIB = ""
    CurMonthlyBudget = 0
    rsTemp.Open strSql, ConnERP, adOpenKeyset, adLockReadOnly, adCmdText
    Do While Not rsTemp.EOF
        CurMonthlyBudget = CurMonthlyBudget + rsTemp.Fields("Grand_Total")
        StrSourceIB = StrSourceIB & rsTemp.Fields("IB_ID")
        
        rsTemp.MoveNext
    Loop
    rsTemp.Close
    
    '------------------ Is MQ Exists ------------------------------
    IsOldmediaQuotation_Exists = False
    
    'Select Where Brand,Media,Month,year
    strSql = "SELECT * FROM IB_Other_Quotation_Detail WHERE YEAR=" & IntYearQuotation & " AND Month=" & IntMonthQuotation
    strSql = strSql & " AND LEFT(Job_ID,8)='" & strBrandCode & "." & StrMediaType & "'"
    
    RsOldMediaQuotation.Open strSql, ConnERP, adOpenKeyset, adLockOptimistic
    
    'Check MQ Monthly
    If Not RsOldMediaQuotation.EOF Then 'If Exist MQ Selected Month ,'Insert To Revision
                                        
        'Generate Revision Value
        strSql = "SELECT MAX(Revision) as LastRevision FROM IB_Other_Quotation_Detail_Revision WHERE Job_Id='" & RsOldMediaQuotation.Fields("Job_Id").Value & "'"
        Rs_Revision.Open strSql, ConnERP, , , adCmdText
        If Not IsNull(Rs_Revision.Fields("LastRevision").Value) Then
            Revision = Rs_Revision.Fields("LastRevision").Value + 1
        Else
            Revision = 1
        End If
        Rs_Revision.Close
        Set Rs_Revision = Nothing
        '***********
        
        StrJobId = RsOldMediaQuotation.Fields("Job_Id").Value
        
        'Save Data
        strSql = ""
        strSql = "INSERT INTO IB_Other_Quotation_Detail_Revision VALUES('"
        'IB_ID
        strSql = strSql & RsOldMediaQuotation.Fields("IB_Id").Value & "','"
        'Job_Id
        strSql = strSql & RsOldMediaQuotation.Fields("Job_Id").Value & "',"
        'Revision
        strSql = strSql & Revision & ","
        'Month
        strSql = strSql & RsOldMediaQuotation.Fields("Month").Value & ","
        'Year
        strSql = strSql & RsOldMediaQuotation.Fields("Year").Value & ","
        'Gross_Cost
        strSql = strSql & RsOldMediaQuotation.Fields("Gross_Cost").Value & ","
        'Nett Cost
        strSql = strSql & RsOldMediaQuotation.Fields("Nett_Cost").Value & ","
        'Media_Sptv_Charge
        strSql = strSql & RsOldMediaQuotation.Fields("Media_Sptv_Charge").Value & ","
        'Other_Charge
        strSql = strSql & RsOldMediaQuotation.Fields("Other_Charge").Value & ","
        'Total_Lintas
        strSql = strSql & RsOldMediaQuotation.Fields("Total_Lintas").Value & ","
        'Agency_Charge
        strSql = strSql & RsOldMediaQuotation.Fields("Agency_Charge").Value & ",'"
        'Job_Number_Agency
        strSql = strSql & RsOldMediaQuotation.Fields("Job_Number_Agency").Value & "',"
        'Grand_Total
        strSql = strSql & RsOldMediaQuotation.Fields("Grand_Total").Value & ")"
        
        ConnERP.Execute strSql
                                        
        IsOldmediaQuotation_Exists = True
    Else
        IsOldmediaQuotation_Exists = False
    End If
    
    RsOldMediaQuotation.Close
    Set RsOldMediaQuotation = Nothing
    '------------------ End if MQ Exists -----------------------
    
    If IsOldmediaQuotation_Exists Then
        'Delete Old Media Quotation
        strSql = "DELETE FROM IB_Other_Quotation_Detail WHERE Job_id='" & StrJobId & "'"
        ConnERP.Execute strSql
    End If
      
    
    If CurMonthlyBudget = 0 Then 'if Value money=0 then
        'Update BC (Tidak Hapus)
        strSql = "UPDATE ULI_Budget_Control SET Budget=0"
        strSql = strSql & " WHERE Job_id='" & StrJobId & "'"
        ConnERP.Execute strSql
        'Update Budget Balannce
        strSql = "UPDATE ULI_Budget_Control SET "
        strSql = strSql & "Budget_Balance=Budget-Money_Spent"
        strSql = strSql & " WHERE Job_id='" & StrJobId & "'"
        ConnERP.Execute strSql
        
    Else 'Else Value Money <> 0
                        
        '----------------- Create MQ -------------------------------
        'Get MQ ID
            NewMQNumber = Get_OT_Media_Quotation_No(strBrandCode, IntYearQuotation, StrMediaType)
        'Jo ID
            StrJobId = strBrandCode & "." & StrMediaType & "." & Right(IntYearQuotation, 2) & Format(IntMonthQuotation, "00")
                
        '------> IB TV Quotation
        strSql = "INSERT INTO IB_Other_Quotation (IB_ID,Month_IB,Year,Date,Entered_By,Media_Plan_No,Source_IB) VALUES("
        strSql = strSql & "'" & NewMQNumber & "'," & IntMonthQuotation & "," & IntYearQuotation & ",Getdate(),'" & Clear_String(strLogin_FullName) & "',"
        strSql = strSql & "'" & strMPNumber & "','" & StrSourceIB & "')"
        
        ConnERP.Execute strSql
        
        '------> IB TV Quotation Detail
        'IB
        strSql = ""
        strSql = "INSERT INTO IB_Other_Quotation_Detail(IB_ID,Job_ID,Month,Year,Gross_Cost,Nett_Cost,Media_Sptv_Charge,Other_Charge,Total_Lintas,Agency_Charge,Job_Number_Agency,Grand_Total,Status) VALUES('"
        'IB_ID(MQ NO)
        strSql = strSql & NewMQNumber & "','"
        'Job Id
        strSql = strSql & StrJobId & "',"
        'Month
        strSql = strSql & IntMonthQuotation & ","
        'Year
        strSql = strSql & IntYearQuotation & ","
        'Gross_Cost
        strSql = strSql & CurMonthlyBudget & ","
        'Nett_Cost
        strSql = strSql & CurMonthlyBudget & ","
        'Media_Sptv_Charge
        strSql = strSql & "0,"
        'Other Charge
        strSql = strSql & "0,"
        'Total_Lintas
        strSql = strSql & CurMonthlyBudget & ","
        'Agency Charge
        strSql = strSql & "0,'"
        'Job Number Agency
        strSql = strSql & "',"
        'Grand Total
        strSql = strSql & CurMonthlyBudget & ","
        'Status
        strSql = strSql & "0)"
            
        ConnERP.Execute strSql
        
        '---------------- End Create MQ ------------------------
        '---------------- Insert /Update BC -------------------
        If IsOldmediaQuotation_Exists Then 'If Exist BC Selected Month
            'Update ULI Budget Control
            strSql = "UPDATE ULI_Budget_Control SET Budget=" & CurMonthlyBudget
            strSql = strSql & ",Budget_Balance=" & CurMonthlyBudget & "-Money_Spent"
            strSql = strSql & " WHERE Job_id='" & StrJobId & "'"
            
            ConnERP.Execute strSql
            
        Else 'BC Not Found
            'Insert Budget Control
            strSql = "INSERT INTO ULI_Budget_Control (Job_Id,Budget,Money_Spent,Plan_Budget,Budget_Balance) VALUES("
            strSql = strSql & "'" & StrJobId & "',"
            strSql = strSql & CurMonthlyBudget & ",0,0," & CurMonthlyBudget & ")"
                        
            ConnERP.Execute strSql
        End If 'End if
        '---------------- End Insert /Update BC -------------------
        
    End If
    
    'Return Value
    CreateMonthlyQuotation_OT = True
    
    'Add Current User Job
    If IsOther Then
        'OT
        Add_Current_User_Job 9, strLogin_FullName, NewMQNumber, "", "", "", "", "", ""
    Else
        'Cinema
        Add_Current_User_Job 10, strLogin_FullName, NewMQNumber, "", "", "", "", "", ""
    End If
    
    Exit Function
errLbl:
    CreateMonthlyQuotation_OT = False
End Function
