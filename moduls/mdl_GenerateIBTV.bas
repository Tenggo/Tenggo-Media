Attribute VB_Name = "mdl_GenerateIBTV"
Option Explicit
Public Function Generate_IB_TV(StrMPMonthlyActivity As String, PlanMonth As Integer) As String
'*************************************************************
'Function Name      : Generate_IB_TV
'Function Decription: To Generate IB for TV
'Input Parameter    : StrMPMonthlyActivity, PlanMonth
'Output Parameter   : Generate_IB_TV
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
        
    Dim IB_ID As String
    Dim Brief_Id As String
    Dim rsMPInfoGeneral As New ADODB.Recordset
    Dim rsObjective As New ADODB.Recordset
    Dim rsObjMaterial As New ADODB.Recordset
    Dim rsCampaign As New ADODB.Recordset
    
    Dim TotalTaprsPerObjective As Double
    Dim TotalNettPerObjective As Currency
    Dim TotalGrossPerObjective As Currency
    Dim TotalMSCPerObjective As Currency
        
    Dim ObjectiveID As Double
    Dim CampaignID As Double
    Dim StrObjID As String
    Dim rsBrandVariant As New ADODB.Recordset
    
    On Error GoTo errLbl
    
    
    'Cancel Previous IB
    If Not Cancel_IB(StrMPMonthlyActivity, PlanMonth, False) Then
        Generate_IB_TV = "Error"
        Exit Function
    End If
    
    StrObjID = ""
    
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
    
        Generate_IB_TV = "Error"
        Exit Function
    End If
    
    'Get Mapping Brand in MP & Activity to IB
    '----------------------------------------------
    BrandCode = IIf(IsNull(rsMPInfoGeneral.Fields("Original_Brand_Code").Value), "", rsMPInfoGeneral.Fields("Original_Brand_Code").Value)
    BrandName = IIf(IsNull(rsMPInfoGeneral.Fields("Original_Brand_Name").Value), "", rsMPInfoGeneral.Fields("Original_Brand_Name").Value)
    
    BrandVariantCode = IIf(IsNull(rsMPInfoGeneral.Fields("Brand_Variant_Code").Value), "", rsMPInfoGeneral.Fields("Brand_Variant_Code").Value)
    BrandVariantName = IIf(IsNull(rsMPInfoGeneral.Fields("Brand_Variant_Name").Value), "", rsMPInfoGeneral.Fields("Brand_Variant_Name").Value)
    
    'Get Original Brand Variant Code,Name
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
    
    '----------------------------------------------
    
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
    
    'Populate Variable TV
    
    'Get Brief ID
    Brief_Id = Get_Brief_ID(BrandCode, PlanYear)
    'Generate IB ID
    IB_ID = GetIB_ID_TV(BrandCode, PlanYear)
        
    '------------ View Tmp IB in Window (Later) ------------------'
    
    '------------ Insert Data to IB Table--------------------------
    
    'Header
        strSql = "INSERT INTO IB_TV "
        strSql = strSql & " VALUES("
        'Client_Brief_Id
        strSql = strSql & "'" & Brief_Id & "',"
        'IB ID
        strSql = strSql & "'" & IB_ID & "',"
        'Revision
        strSql = strSql & "0,"
        'Month
        strSql = strSql & PlanMonth & ","
        'Year
        strSql = strSql & PlanYear & ","
        'Brand_Name
        strSql = strSql & "'" & BrandName & "',"
        'Target Primary
        strSql = strSql & "'" & StrPrimaryTarget & "',"
        'Target Secondary
        strSql = strSql & "'" & StrSecondaryTarget & "',"
        'Consideration (Tidak Ada)
        strSql = strSql & "'',"
        'Other Consideration
        strSql = strSql & "'',"
        'Attachment (Tidak Ada)
        strSql = strSql & "'',"
        'Date Entered
        strSql = strSql & "Getdate(),"
        'Entered By
        strSql = strSql & "'" & strLogin_FullName & "',"
        'Client Approval
        strSql = strSql & "1,"
        'Approval Date
        strSql = strSql & "Getdate(),"
        'Brand Variant Code
        strSql = strSql & "'" & BrandVariantCode & "',"
        'Brand Variant Name
        strSql = strSql & "'" & BrandVariantName & "',"
        'Cluster Code
        strSql = strSql & SecondaryTargetCode & ","
        'Plan No
         strSql = strSql & "'" & PlanNo & "',"
        'Cancel IB ID
        strSql = strSql & "Null,"
        'Cancel By
        strSql = strSql & "Null,"
        'Cancel Date
        strSql = strSql & "Null,"
        'MP_Medium_Id
        strSql = strSql & "'" & StrMPMonthlyActivity & "',"
        'Month_Number
        strSql = strSql & PlanMonth & ","
        'Status
        strSql = strSql & "1"
        strSql = strSql & ")"
        
        ConnERP.Execute strSql
                    
    'IB  Objective (Looping) source from MP_TV_Reach_Frequency
        strSql = "SELECT dbo.MP_Activity.Activity_Type,dbo.MP_TV_Reach_Frequency.mp_tv_rf_id,dbo.MP_TV_Reach_Frequency.frequency_code, dbo.MP_TV_Reach_Frequency.frequency_name, dbo.MP_TV_Reach_Frequency.reach, "
        strSql = strSql & " dbo.MP_TV_Reach_Frequency.week_commencing_start, dbo.MP_TV_Reach_Frequency.week_commencing_end,"
        strSql = strSql & " dbo.MP_TV_Reach_Frequency.month_start, dbo.MP_TV_Reach_Frequency.month_end, dbo.MP_TV_Reach_Frequency.week_year_start,"
        strSql = strSql & " dbo.MP_TV_Reach_Frequency.week_year_end, isnull(dbo.MP_TV_Reach_Frequency.market_code,0) Market_Code ,isnull(dbo.MP_TV_Reach_Frequency.market_name,'NATIONAL') Market_Name,"
        strSql = strSql & " dbo.MP_Monthly_Activity.MSC_Paid,dbo.MP_Monthly_Activity.MSC_Paid_On_Flag"
        strSql = strSql & " FROM dbo.MP_Monthly_Activity INNER JOIN"
        strSql = strSql & " dbo.MP_Medium ON dbo.MP_Monthly_Activity.mp_medium_id = dbo.MP_Medium.mp_medium_id INNER JOIN"
        strSql = strSql & " dbo.MP_Activity ON dbo.MP_Medium.mp_Activity_id = dbo.MP_Activity.mp_Activity_id INNER JOIN"
        strSql = strSql & " dbo.MP_Medium_Detail ON dbo.MP_Medium.mp_medium_id = dbo.MP_Medium_Detail.mp_medium_id INNER JOIN"
        strSql = strSql & " dbo.MP_Plan_Dimension ON dbo.MP_Medium_Detail.mp_medium_detail_id = dbo.MP_Plan_Dimension.mp_medium_detail_id INNER JOIN"
        strSql = strSql & " dbo.MP_TV_Reach_Frequency ON dbo.MP_Plan_Dimension.mp_plan_dim_id = dbo.MP_TV_Reach_Frequency.mp_plan_dim_id"
        strSql = strSql & " WHERE dbo.MP_Monthly_Activity.MP_Medium_ID='" & StrMPMonthlyActivity & "'"
        strSql = strSql & " AND dbo.MP_Monthly_Activity.Month_Number=" & PlanMonth
        strSql = strSql & " AND dbo.MP_TV_Reach_Frequency.Month_Start=" & PlanMonth
        
        rsObjective.Open strSql, ConnERP, adOpenForwardOnly, adLockReadOnly
        
        Do While Not rsObjective.EOF
        
            TotalTaprsPerObjective = 0
            TotalGrossPerObjective = 0
            TotalNettPerObjective = 0
            TotalMSCPerObjective = 0
                
            'Get Tarps,Nett,Msc
            Dim RsTotalPerObj As New ADODB.Recordset
            
            If Trim(StrObjID) <> "" Then
                StrObjID = StrObjID & "," & rsObjective.Fields("mp_tv_rf_id").Value
            Else
                StrObjID = StrObjID & rsObjective.Fields("mp_tv_rf_id").Value
            End If
            
            strSql = "SELECT SUM(Spot) As TotalSpot, SUM(Nett_Rate) AS TotalNett,Sum(Gross_Rate) as TotalGross FROM MP_Insertion WHERE "
            strSql = strSql & "MP_Tv_RF_Id=" & rsObjective.Fields("mp_tv_rf_id").Value
            
            RsTotalPerObj.Open strSql, ConnERP, adOpenForwardOnly, adLockReadOnly
            Do While Not RsTotalPerObj.EOF
            
                TotalTaprsPerObjective = TotalTaprsPerObjective + RsTotalPerObj.Fields("TotalSpot").Value
                TotalGrossPerObjective = TotalGrossPerObjective + RsTotalPerObj.Fields("TotalGross").Value
                TotalNettPerObjective = TotalNettPerObjective + RsTotalPerObj.Fields("TotalNett").Value
                                
                RsTotalPerObj.MoveNext
            Loop
            
            RsTotalPerObj.Close
            
            'Count Fee
            Select Case rsObjective.Fields("MSC_Paid_On_Flag").Value
                Case 0 'On Gross
                     TotalMSCPerObjective = (rsObjective.Fields("MSC_Paid").Value / 100) * TotalGrossPerObjective
                Case 1 'On Nett
                     TotalMSCPerObjective = (rsObjective.Fields("MSC_Paid").Value / 100) * TotalNettPerObjective
                Case 2 'On Gross Rate
                     TotalMSCPerObjective = (rsObjective.Fields("MSC_Paid").Value / 100) * TotalGrossPerObjective
                Case 3 'On Gross Value
                     TotalMSCPerObjective = (rsObjective.Fields("MSC_Paid").Value / 100) * TotalGrossPerObjective
                Case 4 'On Nett Value
                    TotalMSCPerObjective = (rsObjective.Fields("MSC_Paid").Value / 100) * TotalNettPerObjective
            End Select
    
            'Market set Default first
            ObjectiveID = Get_Objective_Id
            
            strSql = "INSERT INTO IB_TV_Objective"
            strSql = strSql & " VALUES("
            'Objecttive ID (Key)
            strSql = strSql & ObjectiveID & ","
            'Client_Brief_Id
            strSql = strSql & "'" & Brief_Id & "',"
            'IB ID (FK)
            strSql = strSql & "'" & IB_ID & "',"
            'week_Commencing (Start + End)
            strSql = strSql & "'" & Format(rsObjective.Fields("Week_Commencing_Start").Value, "MM/DD/YYYY") & " - " & Format(rsObjective.Fields("week_commencing_End").Value, "MM/DD/YYYY") & "',"
            'Campaign_Type
            strSql = strSql & "'" & ActivityTypeCode & "',"
            'Frequency
            strSql = strSql & "'" & Trim(rsObjective.Fields("Frequency_Name").Value) & "',"
            'Reach
            strSql = strSql & rsObjective.Fields("Reach").Value & ","
            'Tarps
            strSql = strSql & TotalTaprsPerObjective & ","
            'Budget_With_MSC
            strSql = strSql & TotalNettPerObjective + TotalMSCPerObjective & ","
            'Row(Tidah dipakai)
            strSql = strSql & "Null,"
            'Campaign_Type_Code
            strSql = strSql & Get_Campaign_Type(rsObjective.Fields("Activity_Type").Value) & ","
            'Frequency_Type_Code
            strSql = strSql & rsObjective.Fields("Frequency_Code").Value & ","
            'W_C_From
            strSql = strSql & "'" & Format(rsObjective.Fields("Week_Commencing_Start").Value, "MM/DD/YYYY") & "',"
            'W_C_To
            strSql = strSql & "'" & Format(rsObjective.Fields("week_commencing_End").Value, "MM/DD/YYYY") & "',"
            'Nett
            strSql = strSql & TotalNettPerObjective & ","
            'MSc
            strSql = strSql & TotalMSCPerObjective & ","
            'Market_Code
            strSql = strSql & rsObjective.Fields("Market_Code").Value & ","
            'Market Name
            strSql = strSql & "'" & rsObjective.Fields("Market_Name").Value & "',"
            'Status
            strSql = strSql & "1,"
            'Replace_By
            strSql = strSql & "Null,"
            'Replace_Date
            strSql = strSql & "Null,"
            'Cancel_By
            strSql = strSql & "Null,"
            'Cancel_Date
            strSql = strSql & "Null"
            strSql = strSql & ")"
            
            ConnERP.Execute strSql
            
            
                 ' IB TV Objective Material
                 '----------------------------------------
                strSql = "SELECT dbo.MP_Plan_Dimension.version, dbo.MP_Plan_Dimension.duration, dbo.MP_Plan_Dimension.mp_plan_dim_id"
                strSql = strSql & " FROM dbo.MP_Plan_Dimension"
                strSql = strSql & " WHERE dbo.MP_Plan_Dimension.mp_plan_dim_id IN ("
                strSql = strSql & " SELECT  mp_plan_dim_id FROM mp_insertion "
                strSql = strSql & " WHERE mp_tv_rf_id = " & rsObjective.Fields("mp_tv_rf_id").Value & ")"
                
                rsObjMaterial.Open strSql, ConnERP, adOpenForwardOnly, adLockReadOnly
                
                Do While Not rsObjMaterial.EOF
                    
                    strSql = "INSERT INTO IB_TV_Objective_Material VALUES("
                    'Client_Bried_Id
                    strSql = strSql & "'" & Brief_Id & "',"
                    'IB ID
                    strSql = strSql & "'" & IB_ID & "',"
                    'Materil_Id (Tidak Dipakai)
                    strSql = strSql & "NULL,"
                    'Objective_Id (Key)
                    strSql = strSql & ObjectiveID & ","
                    'Material_Name (Key)
                    strSql = strSql & "'" & Clear_String(rsObjMaterial.Fields("Version").Value) & "',"
                    'Materil_Duration (Key)
                    strSql = strSql & rsObjMaterial.Fields("Duration").Value & ")"
                    
                    ConnERP.Execute strSql
                
                    'IB Campaign
                    '-------------------------------------------
                        strSql = "SELECT week_commencing, spot, nett_rate, gross_rate, [month], mp_tv_rf_id"
                        strSql = strSql & " FROM dbo.MP_Insertion"
                        strSql = strSql & " WHERE mp_tv_rf_id=" & rsObjective.Fields("mp_tv_rf_id").Value
                        strSql = strSql & " AND mp_plan_dim_id='" & rsObjMaterial.Fields("mp_plan_dim_id").Value & "'"
                        rsCampaign.Open strSql, ConnERP, adOpenForwardOnly, adLockReadOnly
                    
                        Do While Not rsCampaign.EOF
                            CampaignID = Get_Campaign_Id 'Campaign ID
                    
                            strSql = "INSERT INTO IB_TV_Campaign"
                            strSql = strSql & " VALUES("
                            'Campaign_Id (Key)
                            strSql = strSql & CampaignID & ","
                            'Client_Brief_Id
                            strSql = strSql & "'" & Brief_Id & "',"
                            'IB ID
                            strSql = strSql & "'" & IB_ID & "',"
                            'Materil_Id (Tidak Dipakai)
                            strSql = strSql & "NULL,"
                            'Week_Commecing_Date (Tidak Dipakai)
                            strSql = strSql & "NULL,"
                            'Month  (Tidak Dipakai)
                            strSql = strSql & "NULL,"
                            'Year  (Tidak Dipakai)
                            strSql = strSql & "NULL,"
                            'Tarps_Per_Week
                            strSql = strSql & rsCampaign.Fields("spot").Value & ","
                            'Column(Tidak Dipakai)
                            strSql = strSql & "NULL,"
                            'Row(Tidak Dipakai)
                            strSql = strSql & "NULL,"
                            'Objective_Id (FK)
                            strSql = strSql & ObjectiveID & ","
                            'Month_Week_Commencing
                            strSql = strSql & "'" & rsCampaign.Fields("Month").Value & "/1/" & PlanYear & "',"
                            'Week_Commencing
                            strSql = strSql & "'" & rsCampaign.Fields("Week_Commencing").Value & "',"
                            'Material_Name (FK)
                            strSql = strSql & "'" & Clear_String(rsObjMaterial.Fields("Version").Value) & "',"
                            'Material_Duration (FK)
                            strSql = strSql & rsObjMaterial.Fields("Duration").Value & ")"
                            
                            ConnERP.Execute strSql
                                                        
                            rsCampaign.MoveNext 'Next Campign
                        Loop
                        
                        rsCampaign.Close
                        
                    rsObjMaterial.MoveNext 'Next Material
                Loop
                
                rsObjMaterial.Close
                
            rsObjective.MoveNext 'Next Objective
        Loop
        
        rsObjective.Close
        Set rsObjective = Nothing
        
        'Monthly Budget
        
        Dim RsMonthlyBudget As New ADODB.Recordset
        Dim MonthlyFee As Currency
        
        strSql = "SELECT Month,SUM(Spot) As TotalSpot, SUM(Nett_Rate) AS TotalNett,Sum(Gross_Rate) as TotalGross FROM MP_Insertion WHERE "
        strSql = strSql & " MP_Tv_RF_Id IN(" & StrObjID & ")"
        strSql = strSql & " GROUP BY month"
        
        RsMonthlyBudget.Open strSql, ConnERP, adOpenForwardOnly, adLockReadOnly
        Do While Not RsMonthlyBudget.EOF
            'Calculate Fee Monthly
            MonthlyFee = Get_Fee(Left(IB_ID, 4), RsMonthlyBudget.Fields("TotalNett").Value, RsMonthlyBudget.Fields("TotalGross").Value, RsMonthlyBudget.Fields("Month").Value, PlanYear)
                             
            'IB Monthly Budget
            strSql = "INSERT INTO IB_TV_Montly_Budget "
            strSql = strSql & " VALUES("
            'Client_Brief_Id
            strSql = strSql & "'" & Brief_Id & "',"
            'IB ID
            strSql = strSql & "'" & IB_ID & "',"
            'Month
            strSql = strSql & RsMonthlyBudget.Fields("Month").Value & ","
            'Monthly_Budget
            strSql = strSql & RsMonthlyBudget.Fields("Month").Value & ","
            'Budget
            strSql = strSql & RsMonthlyBudget.Fields("TotalNett").Value + MonthlyFee
            strSql = strSql & ")"
            
            ConnERP.Execute strSql
            
            '-------------- Submit Data to Budget Control Plan & Generate Monthly Quotation If BU1 ------------------
            If Is_Special_Brand(BrandCode) Then
                'BU1
                '---------- SubmitDataToMQTV(IBID) -----> Generate Quotation
                Call CreateMonthlyQuotation_TV(BrandCode, RsMonthlyBudget.Fields("Month").Value, PlanYear, PlanNo)
                
'                'SubmitdatatoBCMPBU1(Brand_code,Year,month,Medium,Gross,Nett)
'                StrSQL = "INSERT INTO Budget_Control_Detail_BU1(mp_medium_id,Month_Number,Year_Budget,Month_Budget,Medium,Brand_code,Gross_Budget,Nett_Budget) VALUES ("
'                'mp_medium_id,Month_Number
'                StrSQL = StrSQL & "'" & StrMPMonthlyActivity & "'," & PlanMonth & ","
'                'Year_Budget,Month_Number
'                StrSQL = StrSQL & PlanYear & "," & RsMonthlyBudget.Fields("Month").Value & ","
'                'Medium, Brand_Code
'                StrSQL = StrSQL & "'TV','" & BrandCode & "',"
'                'Gross_Budget, Nett_Bugdet
'                StrSQL = StrSQL & RsMonthlyBudget.Fields("TotalGross").Value & "," & RsMonthlyBudget.Fields("TotalNett").Value
'                StrSQL = StrSQL & ")"
'                ConnERP.Execute StrSQL
                
            Else
'                'SubmitdatatoBCMPBU2(Brand_code,Year,month,Medium,Gross,Nett)
'                StrSQL = "INSERT INTO Budget_Control_Detail_BU2(mp_medium_id,Month_Number,Year_Budget,Month_Budget,Medium,Brand_code,Gross_Budget,Nett_Budget) VALUES ("
'                'mp_medium_id,Month_Number
'                StrSQL = StrSQL & "'" & StrMPMonthlyActivity & "'," & PlanMonth & ","
'                'Year_Budget,Month_Number
'                StrSQL = StrSQL & PlanYear & "," & RsMonthlyBudget.Fields("Month").Value & ","
'                'Medium, Brand_Code
'                StrSQL = StrSQL & "'TV','" & BrandCode & "',"
'                'Gross_Budget, Nett_Bugdet
'                StrSQL = StrSQL & RsMonthlyBudget.Fields("TotalGross").Value & "," & RsMonthlyBudget.Fields("TotalNett").Value
'                StrSQL = StrSQL & ")"
'                ConnERP.Execute StrSQL
            End If
            '--------------------------------------------------------------------
                
            RsMonthlyBudget.MoveNext
        Loop
        
        RsMonthlyBudget.Close
        
    
    'Return Value
    Generate_IB_TV = IB_ID
    
    'Add Cuurent User Job
    Add_Current_User_Job 1, strLogin_FullName, IB_ID, "", "", "", "", "", ""
    
    Exit Function
    
errLbl:
    MsgBox Err.Description, vbExclamation, strApplication_Name
    Generate_IB_TV = "Error"
    
    'Rollback Process
    
End Function

Private Function Get_Campaign_Type(StrCampaign As String) As Integer
    Dim strSql As String
    Dim RsCampType As New ADODB.Recordset
    
    strSql = "SELECT Code FROM Campaign_Type_Catalog WHERE Campaign_Type_Name='" & Trim(StrCampaign) & "'"
    RsCampType.Open strSql, ConnERP, adOpenForwardOnly, adLockReadOnly
    If Not RsCampType.EOF Then
        Get_Campaign_Type = RsCampType.Fields("Code").Value
    Else
        Get_Campaign_Type = 0 'No Data Found
    End If
    
    RsCampType.Close
    Set RsCampType = Nothing
End Function

Public Function GetIB_ID_TV(strBrandCode As String, IntYear As Integer) As String
'*************************************************************
'Function Name      : GetIB_ID_TV
'Function Decription: To Generate IB ID TV
'Input Parameter    : StrBrandCode, IntYear
'Output Parameter   : GetIB_ID_TV
'Created Date/By    : 26 June 2005/Yayan Royani
'Last Update/By     :
'*************************************************************
    Dim Out_Param As ADODB.Parameter
    Dim In_Param1 As ADODB.Parameter
    Dim In_Param2 As ADODB.Parameter
    Dim In_Param3 As ADODB.Parameter
    Dim cmd As New ADODB.Command
    
             
    cmd.CommandType = adCmdStoredProc
    cmd.CommandText = "Get_New_IB_Id_TV"
     
    Set In_Param1 = cmd.CreateParameter("Brand_Code", adChar, adParamInput, 4)
    Set In_Param2 = cmd.CreateParameter("Media_Type", adChar, adParamInput, 3)
    Set In_Param3 = cmd.CreateParameter("Year", adInteger, adParamInput)
    Set Out_Param = cmd.CreateParameter("New_IB_Id", adChar, adParamOutput, 13)
        
    cmd.Parameters.Append In_Param1
    cmd.Parameters.Append In_Param2
    cmd.Parameters.Append In_Param3
    cmd.Parameters.Append Out_Param
    
    In_Param1.Value = strBrandCode
    In_Param2.Value = "055"
    In_Param3.Value = IntYear
         
    cmd.ActiveConnection = ConnERP
    cmd.Execute
    GetIB_ID_TV = Out_Param.Value
         
End Function


Private Function Get_Fee(strBrandCode As String, NettValue As Currency, GrossValue As Currency, intMonth As Integer, IntYear As Integer) As Currency
'*************************************************************
'Function Name      : Get_Fee
'Function Decription: To Get Fee ob Brand montly
'Input Parameter    : StrBrandCode,NettValue,IntMonth ,IntYear
'Output Parameter   : Get_Fee
'Created Date/By    : 18 July 2005/Yayan Royani
'Last Update/By     :
'*************************************************************

    Dim strSql As String
    Dim rsFee As New ADODB.Recordset
    Dim MonthlyFee As Currency
    
    strSql = "SELECT MSC_Paid,MSC_Paid_On_Flag FROM Brand_Fee "
    strSql = strSql & " WHERE Brand_Code='" & strBrandCode & "'"
    strSql = strSql & " AND Month=" & intMonth
    strSql = strSql & " AND Year=" & IntYear
    
    rsFee.Open strSql, ConnERP, adOpenForwardOnly, adLockReadOnly
    If Not rsFee.EOF Then
        Select Case rsFee.Fields("MSC_Paid_On_Flag").Value
            Case 0 'On Gross
                 MonthlyFee = (rsFee.Fields("MSC_Paid").Value / 100) * GrossValue
            Case 1 'On Nett
                 MonthlyFee = (rsFee.Fields("MSC_Paid").Value / 100) * NettValue
            Case 2 'On Gross Rate
                 MonthlyFee = (rsFee.Fields("MSC_Paid").Value / 100) * GrossValue
            Case 3 'On Gross Value
                 MonthlyFee = (rsFee.Fields("MSC_Paid").Value / 100) * GrossValue
            Case 4 'On Nett Value
                MonthlyFee = (rsFee.Fields("MSC_Paid").Value / 100) * NettValue
        End Select
    Else
        MonthlyFee = 0
    End If

    rsFee.Close
    Set rsFee = Nothing
    
    Get_Fee = MonthlyFee
    
    
End Function


Public Function CreateMonthlyQuotation_TV(strBrandCode As String, IntMonthQuotation As Integer, IntYearQuotation As Integer, strMPNumber As String) As Boolean
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
    StrMediaType = "055"
    
    'Get Value  Money Selected Month From Monthly IB (IB Lama Tidak Bisa Tidak Keambil)
    'Dari IB Yang statunya Aktif dan Month dan Year-nya Sama
    strSql = "SELECT IB_ID,Budget FROM IB_TV_Montly_Budget WHERE Month=" & IntMonthQuotation
    strSql = strSql & " AND IB_ID IN (SELECT IB_ID FROM IB_TV WHERE LEFT(IB_ID,4)='" & strBrandCode & "' AND Year=" & IntYearQuotation & " AND Status=1 AND mp_medium_id IS NOT NULL)"
    
    StrSourceIB = ""
    CurMonthlyBudget = 0
    rsTemp.Open strSql, ConnERP, adOpenKeyset, adLockReadOnly, adCmdText
    Do While Not rsTemp.EOF
        CurMonthlyBudget = CurMonthlyBudget + rsTemp.Fields("Budget")
        StrSourceIB = StrSourceIB & rsTemp.Fields("IB_ID")
        
        rsTemp.MoveNext
    Loop
    rsTemp.Close
    
    '------------------ Is MQ Exists ------------------------------
    IsOldmediaQuotation_Exists = False
    
    strSql = "SELECT * FROM IB_TV_Quotation_Detail WHERE YEAR=" & IntYearQuotation & " AND Month=" & IntMonthQuotation
    strSql = strSql & " AND LEFT(Job_Id,4)='" & strBrandCode & "'"
    
    RsOldMediaQuotation.Open strSql, ConnERP, adOpenKeyset, adLockOptimistic
    
    'Check MQ Monthly
    If Not RsOldMediaQuotation.EOF Then 'If Exist MQ Selected Month ,'Insert To Revision
                                        
        'Generate Revision Value
        strSql = "SELECT MAX(Revision) as LastRevision FROM IB_TV_Quotation_Detail_Revision WHERE Job_Id='" & RsOldMediaQuotation.Fields("Job_Id").Value & "'"
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
        strSql = "INSERT INTO IB_TV_Quotation_Detail_Revision VALUES('"
        strSql = strSql & RsOldMediaQuotation.Fields("IB_Id").Value & "','"
        strSql = strSql & RsOldMediaQuotation.Fields("Job_Id").Value & "',"
        strSql = strSql & Revision & ","
        strSql = strSql & RsOldMediaQuotation.Fields("Month").Value & ","
        strSql = strSql & RsOldMediaQuotation.Fields("Year").Value & ","
        strSql = strSql & RsOldMediaQuotation.Fields("Nett_Cost").Value & ","
        strSql = strSql & RsOldMediaQuotation.Fields("Media_Sptv_Charge").Value & ","
        strSql = strSql & RsOldMediaQuotation.Fields("Other_Charge").Value & ","
        strSql = strSql & RsOldMediaQuotation.Fields("Bonus").Value & ","
        strSql = strSql & RsOldMediaQuotation.Fields("Total_Lintas").Value & ","
        strSql = strSql & RsOldMediaQuotation.Fields("Agency_Charge").Value & ",'"
        strSql = strSql & RsOldMediaQuotation.Fields("Job_Number_Agency").Value & "',"
        strSql = strSql & RsOldMediaQuotation.Fields("Grand_Total").Value & ",'"
        strSql = strSql & RsOldMediaQuotation.Fields("Source_IB").Value & "')"
        
        ConnERP.Execute strSql
                                        
        IsOldmediaQuotation_Exists = True 'Delete Media Quotation
    Else
        IsOldmediaQuotation_Exists = False
    End If
    
    RsOldMediaQuotation.Close
    Set RsOldMediaQuotation = Nothing
    '------------------ End if MQ Exists -----------------------
    
    If IsOldmediaQuotation_Exists Then
        'Delete Old Media Quotation
        strSql = "DELETE FROM IB_TV_Quotation_Detail WHERE Job_id='" & StrJobId & "'"
        ConnERP.Execute strSql
    End If
      
    
    If CurMonthlyBudget = 0 Then 'if Value money=0 then
        'Update BC (Tidak Hapus)
        strSql = "UPDATE ULI_Budget_Control SET Budget=0"
        strSql = strSql & " WHERE Job_id='" & StrJobId & "'"
        ConnERP.Execute strSql
        
        strSql = "UPDATE ULI_Budget_Control SET "
        strSql = strSql & "Budget_Balance=Budget-Money_Spent"
        strSql = strSql & " WHERE Job_id='" & StrJobId & "'"
        ConnERP.Execute strSql
        
    Else 'Else Value Money <> 0
                        
        '----------------- Create MQ -------------------------------
        'Get MQ ID
            NewMQNumber = Get_TV_Media_Quotation_No(strBrandCode, IntYearQuotation, StrMediaType)
        'Jo ID
            StrJobId = strBrandCode & "." & StrMediaType & "." & Right(IntYearQuotation, 2) & Format(IntMonthQuotation, "00")
                
        '------> IB TV Quotation
        strSql = "INSERT INTO IB_TV_Quotation (IB_ID,Year,Date,Entered_By,Media_Plan_No) VALUES("
        strSql = strSql & "'" & NewMQNumber & "'," & IntYearQuotation & ",Getdate(),'" & Clear_String(strLogin_FullName) & "',"
        strSql = strSql & "'" & strMPNumber & "')"
        
        ConnERP.Execute strSql
        
        '------> IB TV Quotation Detail
        'IB
        strSql = ""
        strSql = "INSERT INTO IB_TV_Quotation_Detail(IB_ID,Job_ID,Month,Year,Nett_Cost,Media_Sptv_Charge,Other_Charge,Bonus,Total_Lintas,Agency_Charge,Job_Number_Agency,Grand_Total,Source_IB) VALUES('"
        'IB_ID(MQ NO)
        strSql = strSql & NewMQNumber & "','"
        'Job Id
        strSql = strSql & StrJobId & "',"
        'Month
        strSql = strSql & IntMonthQuotation & ","
        'Year
        strSql = strSql & IntYearQuotation & ","
        'Nett_Cost
        strSql = strSql & CurMonthlyBudget & ","
        'Media_Sptv_Charge
        strSql = strSql & "0,"
        'Other Charge
        strSql = strSql & "0,"
        'Bonus
        strSql = strSql & "0,"
        'Total_Lintas
        strSql = strSql & CurMonthlyBudget & ","
        'Agency Charge
        strSql = strSql & "0,'"
        'Job Number Agency
        strSql = strSql & "',"
        'Grand Total
        strSql = strSql & CurMonthlyBudget & ",'"
        'Source IB
        strSql = strSql & StrSourceIB & "')"
            
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
    CreateMonthlyQuotation_TV = True
    
    'Add_Current_User_Job
    Add_Current_User_Job 6, strLogin_FullName, NewMQNumber, "", "", "", "", "", ""
    
    Exit Function
errLbl:
    
    'Rollback Process
    
    CreateMonthlyQuotation_TV = False
        
End Function

