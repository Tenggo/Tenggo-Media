Attribute VB_Name = "mdl_GenerateIBRadio"
Option Explicit
Dim Mat_ID(25) As String
'
 Public Function Generate_IB_Radio(StrMPMonthlyActivity As String, PlanMonth As Integer) As String
'*************************************************************
'Function Name      : Generate_IB_Radio
'Function Decription: To Generate IB for Radio
'Input Parameter    : StrMPMonthlyActivity, PlanMonth
'Output Parameter   : Generate_IB_Radio
'Created Date/By    : 28 June 2005/Diyah
'Last Update/By     :
'*************************************************************
   '
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
    Dim Month_Plan As Integer
    
    Dim IsRadio As Boolean
    Dim IB_ID As String
    Dim Brief_Id As String
    Dim StrDescription As String
    Dim TOTAL As Currency
    Dim Str_Material_Id As String
    Dim int_count_materi As Integer
    
    Dim rsMPInfoGeneral As New ADODB.Recordset
    Dim rsMPInfoRadio As New ADODB.Recordset
    Dim rsMPMaterial As New ADODB.Recordset
    Dim rsMPPlanDetail As New ADODB.Recordset
    Dim rsMPCity As New ADODB.Recordset
    Dim rs_count_materi As New ADODB.Recordset
    Dim rsMPPlan As New ADODB.Recordset
    Dim rsMPPlanDetail_materi As New ADODB.Recordset
    Dim rsBrandVariant As New ADODB.Recordset
    
    On Error GoTo errLbl
    
    
    'Cancel Previous IB
    If Not Cancel_IB(StrMPMonthlyActivity, PlanMonth, False) Then
        Generate_IB_Radio = "Error"
        Exit Function
    End If
        
    'cal generate material id
    Call Initial_Data
    
    'Populate Variable General
    strSql = " SELECT dbo.MP_Master.mp_number,  dbo.MP_Master.[year]," 'dbo.MP_Master.brand_code, dbo.MP_Master.brand_name
    strSql = strSql & " dbo.MP_Activity.original_brand_Code , dbo.MP_Activity.original_brand_name, "
    strSql = strSql & " dbo.MP_Activity.activity_type , dbo.MP_Activity.activity_desc, "
    strSql = strSql & " dbo.MP_Activity.brand_variant_code, dbo.MP_Activity.brand_variant_name, "
    strSql = strSql & " dbo.MP_Activity.target_audience_code , dbo.MP_Activity.Target_Audience, "
    strSql = strSql & " dbo.MP_Activity.brand_target, dbo.MP_Medium.medium_code, dbo.MP_Monthly_Activity.month_number,"
    strSql = strSql & " dbo.MP_Monthly_Activity.month_name , "
    strSql = strSql & " dbo.MP_Monthly_Activity.min_Budget , dbo.MP_Monthly_Activity.gross_budget, dbo.MP_Monthly_Activity.MSC_Paid_Value"
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
    
        Generate_IB_Radio = "Error"
        Exit Function
    End If
    
    BrandCode = rsMPInfoGeneral.Fields("Original_brand_Code") 'Left(StrMPMonthlyActivity, 4)
    BrandName = rsMPInfoGeneral.Fields("Original_brand_Name") 'IIf(IsNull(rsMPInfoGeneral.Fields("Brand_Name").Value), "", rsMPInfoGeneral.Fields("Brand_Name").Value)
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
    '======================================================
    
    PlanNo = rsMPInfoGeneral.Fields("MP_Number").Value
    PlanYear = rsMPInfoGeneral.Fields("Year").Value
    ActivityTypeCode = IIf(IsNull(rsMPInfoGeneral.Fields("Activity_Type").Value), "", rsMPInfoGeneral.Fields("Activity_Type").Value)
    ActivityDescription = IIf(IsNull(rsMPInfoGeneral.Fields("Activity_Desc").Value), "", rsMPInfoGeneral.Fields("Activity_Desc").Value)
    StrPrimaryTarget = IIf(IsNull(rsMPInfoGeneral.Fields("Brand_target").Value), "", rsMPInfoGeneral.Fields("Brand_Target").Value)
    StrSecondaryTarget = rsMPInfoGeneral.Fields("Target_Audience").Value
    SecondaryTargetCode = rsMPInfoGeneral.Fields("Target_Audience_Code").Value
    MediumCode = rsMPInfoGeneral.Fields("Medium_Code").Value
    MonthlyGrossBudget = rsMPInfoGeneral.Fields("Gross_Budget").Value
    MonthlyNettBudget = rsMPInfoGeneral.Fields("Min_Budget").Value
    MonthlyFeeBudget = rsMPInfoGeneral.Fields("Msc_Paid_Value").Value
    
    rsMPInfoGeneral.Close
    Set rsMPInfoGeneral = Nothing
    '--------------------------------------------------
    
    TOTAL = MonthlyNettBudget + MonthlyFeeBudget 'Nett + Fee
            
    'Get Brief ID
    Brief_Id = Get_Brief_ID(BrandCode, PlanYear)
    'Generate IB ID
    IB_ID = Get_New_IB_ID(BrandCode, PlanYear)
    
    'YY Tanda
    'Get IsCity or Is Area
    
    'YY Tanda (Tambahkan Flag,Client Brief Id tidak perlu)
    
    'masukin ke plan detail
    strSql = "SELECT DISTINCT  b.mp_medium_detail_id,b.area_code,b.radio_station_Code,b.radio_station_name,c.isrdbystation"
    strSql = strSql & " FROM MP_medium_Detail b LEFT JOIN mp_plan_dimension c on b.mp_medium_detail_id=c.mp_medium_detail_id "
    strSql = strSql & " WHERE b.mp_medium_id='" & StrMPMonthlyActivity & "' "
    
    rsMPPlan.Open strSql, ConnERP, adOpenStatic, adLockReadOnly
    
    
    'Insert Data to IB_Radio Table
    strSql = "INSERT INTO IB_Radio(CLIENT_BRIEF_ID,IB_ID,Revision,Media_Plan,[Month],[Year],DATE_ENTERED,"
    strSql = strSql & " ENTERED_BY , Brand_Code, TARGET_PRIMARY, TARGET_SECONDARY,Consideration,Attachment, "
    strSql = strSql & " Approved_Flag,Approved_Date,BRAND_VARIANT_CODE,BRAND_VARIANT_NAME,CLUSTER_CODE,PLAN_NO,mp_medium_id,month_number"
    strSql = strSql & " ,[Date],IsCity)"
    strSql = strSql & " VALUES('" & Brief_Id & "','" & IB_ID & "',0,'" & PlanNo & "'," & PlanMonth & "," & PlanYear & ","
    strSql = strSql & " Getdate(),'" & strLogin_FullName & "','" & BrandCode & "',"
    strSql = strSql & "'" & StrPrimaryTarget & "','" & StrSecondaryTarget & "','','',"
    strSql = strSql & "1,Getdate(),'" & BrandVariantCode & "',"
    strSql = strSql & "'" & BrandVariantName & "','" & SecondaryTargetCode & "','" & PlanNo & "',"
    strSql = strSql & "'" & StrMPMonthlyActivity & "',"
    strSql = strSql & PlanMonth
    strSql = strSql & ",Getdate(),"
    If rsMPPlan.Fields("isrdbystation").Value = 1 Then
        strSql = strSql & "0"
    Else 'Area
        strSql = strSql & "1"
    End If
    
    strSql = strSql & ")"
    
    ConnERP.Execute strSql
        
    'data radio plan
    strSql = "INSERT INTO IB_Radio_Plan (Client_Brief_Id,IB_ID,[Month],[Year],Budget) VALUES"
    strSql = strSql & "('" & Brief_Id & "','" & IB_ID & "'," & PlanMonth & "," & PlanYear & "," & TOTAL & ")"
    ConnERP.Execute strSql
    
    
    
    While Not rsMPPlan.EOF
        strSql = "SELECT d.week_commencing,sum(d.spot)as spot,c.version,c.duration"
        strSql = strSql & " FROM MP_Plan_Dimension c"
        strSql = strSql & " LEFT JOIN mp_insertion d on c.mp_plan_dim_id=d.mp_plan_dim_id"
        strSql = strSql & " WHERE c.mp_medium_detail_id='" & rsMPPlan.Fields("mp_medium_detail_id").Value & "' and d.month=" & PlanMonth & ""
        strSql = strSql & " GROUP BY d.week_commencing,c.version,c.duration"
        
        If rsMPPlanDetail_materi.State = adStateOpen Then
            rsMPPlanDetail_materi.Close
        End If
        
        
        rsMPPlanDetail_materi.Open strSql, ConnERP, adOpenStatic, adLockReadOnly
        
        strSql = "SELECT d.week_commencing,sum(d.spot)as spot"
        strSql = strSql & " FROM MP_Plan_Dimension c"
        strSql = strSql & " LEFT JOIN mp_insertion d on c.mp_plan_dim_id=d.mp_plan_dim_id"
        strSql = strSql & " WHERE c.mp_medium_detail_id='" & rsMPPlan.Fields("mp_medium_detail_id").Value & "' and d.month=" & PlanMonth & ""
        strSql = strSql & " GROUP BY d.week_commencing"
        
        Dim count_spot As Integer
        
        If rsMPPlanDetail.State = adStateOpen Then
            rsMPPlanDetail.Close
        End If
        rsMPPlanDetail.Open strSql, ConnERP, adOpenStatic, adLockReadOnly
        
        If rsMPPlanDetail.EOF Then
            'rsMPPlanDetail.Close
            GoTo TidakAdaInsertion
        End If
        
        count_spot = rsMPPlanDetail.Fields("spot").Value / 7
        
        'HITUNG JUMLAH MATERI
        strSql = "SELECT DISTINCT Version,duration FROM MP_Plan_Dimension WHERE mp_Medium_Detail_id='" & rsMPPlan.Fields("mp_medium_detail_id").Value & "'"
        
        If rs_count_materi.State = adStateOpen Then
            rs_count_materi.Close
        End If
        rs_count_materi.Open strSql, ConnERP, adOpenStatic, adLockReadOnly
        
        int_count_materi = rs_count_materi.RecordCount
        rs_count_materi.Close
        
        If rsMPPlan.Fields("isrdbystation").Value = 1 Then
            'jika radionya by station, langsung cari schedulenya
            strSql = "SELECT City_ID FROM Radio_Station WHERE station_Code='" & rsMPPlan.Fields("radio_station_Code").Value & "'"
            
            If rsMPCity.State = adStateOpen Then
                rsMPCity.Close
            End If
            rsMPCity.Open strSql, ConnERP, adOpenStatic, adLockReadOnly
            
            If Not rsMPCity.EOF Then
               ' masukkan ke tabel IB_Print_Plan_Detail
                rsMPPlanDetail.MoveFirst
                While Not rsMPPlanDetail.EOF
                    strSql = "INSERT INTO IB_Radio_Plan_Detail (Client_Brief_Id,IB_ID,[Month],[Year],Schedule,City_Id,Area_Code,Spot,Urban_Flag,Rural_Flag)"
                    strSql = strSql & "VALUES ('" & Brief_Id & "','" & IB_ID & "','" & PlanMonth & "','" & PlanYear & "','" & rsMPPlanDetail.Fields("Week_Commencing").Value & " - " & rsMPPlan.Fields("radio_station_code").Value & "','" & rsMPCity.Fields("City_id").Value & "','" & rsMPPlan.Fields("radio_station_Code").Value & "'," & rsMPPlanDetail.Fields("Spot").Value / 7 & ",1,1)"
                    ConnERP.Execute strSql
                    
                    rsMPPlanDetail.MoveNext
                Wend
                
                'masukkan ke plan detail material
                rsMPPlanDetail_materi.MoveFirst
                While Not rsMPPlanDetail_materi.EOF
                    If IsExist_Materi(IB_ID, Brief_Id, rsMPPlanDetail_materi.Fields("version").Value, rsMPPlanDetail_materi.Fields("duration").Value) = False Then
                        Str_Material_Id = ""
                        Str_Material_Id = Get_New_Material_ID(IB_ID)
                        
                        strSql = "INSERT INTO IB_Radio_Material(Client_Brief_Id,IB_ID,Material_Id,Material_name,Duration) VALUES"
                        strSql = strSql & " ('" & Brief_Id & "','" & IB_ID & "','" & Str_Material_Id & "','" & rsMPPlanDetail_materi.Fields("version").Value & "','" & rsMPPlanDetail_materi.Fields("duration").Value & "')"
                        ConnERP.Execute strSql
                        
                    Else
                        strSql = "SELECT Material_Id FROM IB_Radio_Material WHERE Client_Brief_Id='" & Brief_Id & "' AND IB_id='" & IB_ID & "' AND Material_Name='" & rsMPPlanDetail_materi.Fields("version").Value & "' AND Duration='" & rsMPPlanDetail_materi.Fields("Duration").Value & "'"
                        rsMPMaterial.Open strSql, ConnERP, adOpenStatic, adLockReadOnly
                        
                        If Not rsMPMaterial.EOF Then
                            Str_Material_Id = rsMPMaterial.Fields("Material_Id").Value
                            rsMPMaterial.Close
                        End If
                        
                    End If
                    
                    rsMPPlanDetail.MoveFirst
                    rsMPPlanDetail.Find "week_commencing='" & rsMPPlanDetail_materi.Fields("Week_Commencing").Value & "'"
                    
                    strSql = "INSERT INTO IB_Radio_Plan_Detail_Material (Client_Brief_Id,IB_ID,[Month],[Year],Schedule,City_Id,Material_Id,Material_mix)"
                    strSql = strSql & "VALUES ('" & Brief_Id & "','" & IB_ID & "','" & PlanMonth & "','" & PlanYear & "','" & rsMPPlanDetail_materi.Fields("Week_Commencing").Value & " - " & rsMPPlan.Fields("radio_station_Code").Value & "','" & rsMPCity.Fields("City_id").Value & "','" & Str_Material_Id & "'," & ((rsMPPlanDetail_materi.Fields("spot") / 7) / (rsMPPlanDetail.Fields("spot").Value / 7)) * 100 & ")" '
                    ConnERP.Execute strSql
                    
                    rsMPPlanDetail_materi.MoveNext
                Wend
            End If
        Else
            'kl by area,cari dlu city diarea tersebut
            strSql = " SELECT distinct city_id FROM radio_station WHERE station_Code in"
            strSql = strSql & " (SELECT station_Code FROM radio_area_detail WHERE area_id "
            strSql = strSql & " in( SELECT area_id FROM radio_area_new WHERE area_name='" & rsMPPlan.Fields("area_Code").Value & "'))"
            
            If rsMPCity.State = adStateOpen Then
                rsMPCity.Close
            End If
            rsMPCity.Open strSql, ConnERP, adOpenStatic, adLockReadOnly
            
            While Not rsMPCity.EOF
               ' masukkan ke tabel IB_Print_Plan_Detail
                rsMPPlanDetail.MoveFirst
                While Not rsMPPlanDetail.EOF
                    strSql = "INSERT INTO IB_Radio_Plan_Detail (Client_Brief_Id,IB_ID,[Month],[Year],Schedule,City_Id,Area_Code,Spot,Urban_Flag,Rural_Flag)"
                    strSql = strSql & "VALUES ('" & Brief_Id & "','" & IB_ID & "','" & PlanMonth & "','" & PlanYear & "','" & rsMPPlanDetail.Fields("Week_Commencing").Value & " " & rsMPPlan.Fields("area_code").Value & "','" & rsMPCity.Fields("City_id").Value & "','" & rsMPPlan.Fields("area_Code").Value & "'," & rsMPPlanDetail.Fields("Spot").Value / 7 & ",1,1)"
                    ConnERP.Execute strSql
                    
                    rsMPPlanDetail.MoveNext
                Wend
                'rsMPPlanDetail.MoveFirst
                
                'masukkan ke plan detail material
                rsMPPlanDetail_materi.MoveFirst
                While Not rsMPPlanDetail_materi.EOF
                    If IsExist_Materi(IB_ID, Brief_Id, rsMPPlanDetail_materi.Fields("version").Value, rsMPPlanDetail_materi.Fields("duration").Value) = False Then
                        Str_Material_Id = ""
                        Str_Material_Id = Get_New_Material_ID(IB_ID)
                        
                        strSql = "INSERT INTO IB_Radio_Material(Client_Brief_Id,IB_ID,Material_Id,Material_name,Duration) VALUES"
                        strSql = strSql & " ('" & Brief_Id & "','" & IB_ID & "','" & Str_Material_Id & "','" & rsMPPlanDetail_materi.Fields("version").Value & "','" & rsMPPlanDetail_materi.Fields("duration").Value & "')"
                        ConnERP.Execute strSql
                        
                    Else
                        strSql = "SELECT Material_Id FROM IB_Radio_Material WHERE Client_Brief_Id='" & Brief_Id & "' AND IB_id='" & IB_ID & "' AND Material_Name='" & rsMPPlanDetail_materi.Fields("version").Value & "' AND Duration='" & rsMPPlanDetail_materi.Fields("Duration").Value & "'"
                        rsMPMaterial.Open strSql, ConnERP, adOpenStatic, adLockReadOnly
                        
                        If Not rsMPMaterial.EOF Then
                            Str_Material_Id = rsMPMaterial.Fields("Material_Id").Value
                            rsMPMaterial.Close
                        End If
                        
                    End If
                    
                    rsMPPlanDetail.MoveFirst
                    rsMPPlanDetail.Find "week_commencing='" & rsMPPlanDetail_materi.Fields("Week_Commencing").Value & "'"

                    strSql = "INSERT INTO IB_Radio_Plan_Detail_Material (Client_Brief_Id,IB_ID,[Month],[Year],Schedule,City_Id,Material_Id,Material_mix)"
                    strSql = strSql & "VALUES ('" & Brief_Id & "','" & IB_ID & "','" & PlanMonth & "','" & PlanYear & "','" & rsMPPlanDetail_materi.Fields("Week_Commencing").Value & " " & Trim(rsMPPlan.Fields("area_Code").Value) & "','" & rsMPCity.Fields("City_id").Value & "','" & Str_Material_Id & "'," & ((rsMPPlanDetail_materi.Fields("spot").Value / 7) / (rsMPPlanDetail.Fields("spot").Value / 7)) * 100 & ")" '(rsMPPlanDetail.Fields("Spot").Value) / 7
                    ConnERP.Execute strSql
                    
                    rsMPPlanDetail_materi.MoveNext
                    'MsgBox rsMPPlanDetail_materi.RecordCount
                Wend
                'MsgBox rsMPCity.RecordCount
                rsMPCity.MoveNext
            Wend
        End If
TidakAdaInsertion:
        
        rsMPPlan.MoveNext
    Wend
    
    
    rsMPPlanDetail.Close
    rsMPPlan.Close
    rsMPPlanDetail_materi.Close
    rsMPCity.Close
    
     '-------------- Submit Data to Budget Control Plan & Generate Monthly Quotation If BU1 ------------------
    If Is_Special_Brand(BrandCode) Then
        'SubmitDataToMQTV(IBID)--->Generate Quotation
        
        Call CreateMonthlyQuotation_RD(BrandCode, PlanMonth, PlanYear, PlanNo)
        
'        'SubmitdatatoBCMPBU1(Brand_code,Year,month,Medium,Gross,Nett)
'        StrSQL = "INSERT INTO Budget_Control_Detail_BU1(mp_medium_id,Month_Number,Year_Budget,Month_Budget,Medium,Brand_code,Gross_Budget,Nett_Budget) VALUES ("
'        'mp_medium_id,Month_Number
'        StrSQL = StrSQL & "'" & StrMPMonthlyActivity & "'," & PlanMonth & ","
'        'Year_Budget,Month_Number
'        StrSQL = StrSQL & PlanYear & "," & PlanMonth & ","
'        'Medium, Brand_Code
'        StrSQL = StrSQL & "'RD','" & BrandCode & "',"
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
'        StrSQL = StrSQL & "'RD','" & BrandCode & "',"
'        'Gross_Budget, Nett_Bugdet
'        StrSQL = StrSQL & "0," & MonthlyNettBudget
'        StrSQL = StrSQL & ")"
'        ConnERP.Execute StrSQL
    End If
    '--------------------------------------------------------------------
        
    
    Generate_IB_Radio = IB_ID
    
    
    'Add Cuurent User Job
    Add_Current_User_Job 2, strLogin_FullName, IB_ID, "", "", "", "", "", ""
    
    Exit Function
    
errLbl:
    MsgBox Err.Description, vbExclamation, strApplication_Name
    Generate_IB_Radio = "Error"
End Function

Private Function Get_New_IB_ID(strBrandCode As String, IntYear As Integer) As String
    
    Rem Set Stored Procedure
    Dim rs As New ADODB.Recordset
    Dim Txt_SQL As String
    Dim Param_Year_IN As New ADODB.Parameter
    Dim Param_Brand_Code_IN As New ADODB.Parameter
    Dim Param_New_IB_Out As New ADODB.Parameter
    Dim Param_Media_Type_IN As New ADODB.Parameter
    Dim Cmd_IB As New ADODB.Command


    Cmd_IB.CommandType = adCmdStoredProc
    Cmd_IB.CommandText = "Get_New_IB_ID_Radio"
    
    Set Param_Brand_Code_IN = Cmd_IB.CreateParameter("Brand_Code", adChar, adParamInput, 4)
    Set Param_Media_Type_IN = Cmd_IB.CreateParameter("Media_Type", adChar, adParamInput, 3)
    Set Param_Year_IN = Cmd_IB.CreateParameter("Year", adInteger, adParamInput)
    Set Param_New_IB_Out = Cmd_IB.CreateParameter("New_IB_Id", adChar, adParamOutput, 13)
    
    
    Cmd_IB.Parameters.Append Param_Brand_Code_IN
    Cmd_IB.Parameters.Append Param_Media_Type_IN
    Cmd_IB.Parameters.Append Param_Year_IN
    Cmd_IB.Parameters.Append Param_New_IB_Out
    
    Param_Year_IN.Value = IntYear
    Param_Brand_Code_IN.Value = strBrandCode
    
    Rem Get Media Type COde
    Txt_SQL = "SELECT Media_Type_Code FROM Media_Type WHERE Media_Type_Name ='Radio Media Induk'"
    rs.Open Txt_SQL, ConnERP, adOpenStatic, adLockReadOnly
    
    With rs
        If Not .EOF Then
            Param_Media_Type_IN.Value = .Fields("Media_Type_Code")
        End If
    End With
    Set rs = Nothing
           
    Cmd_IB.ActiveConnection = ConnERP
    Cmd_IB.Execute
    
    Get_New_IB_ID = Param_New_IB_Out.Value
        
End Function

Private Function Get_New_Material_ID(Str_IB_Id As String) As String
    Dim Pos As Integer
    Dim rs As New ADODB.Recordset
    Dim Ada As Boolean
    Dim TxtSQl As String
    
    TxtSQl = "SELECT Material_ID FROM IB_Radio_Material WHERE IB_ID='" & Str_IB_Id & "'"
    rs.Open TxtSQl, ConnERP, adOpenStatic, adLockReadOnly
    
    'Set rs = Frm_IB_Radio.Rs_Materi_Temp.Clone
    With rs
        Pos = 0
        Ada = False
        Do While .EOF = False
            If .Fields("Material_ID") <> Mat_ID(Pos) Then
                Get_New_Material_ID = Mat_ID(Pos)
                If rs.State = adStateOpen Then
                    rs.Close
                End If
                Set rs = Nothing
                Exit Function
            End If
            Pos = Pos + 1
            .MoveNext
        Loop
    End With
    Get_New_Material_ID = Mat_ID(Pos)
    
    If rs.State = adStateOpen Then
        rs.Close
    End If
    Set rs = Nothing

End Function

Private Sub Initial_Data()
    Mat_ID(0) = "A"
    Mat_ID(1) = "B"
    Mat_ID(2) = "C"
    Mat_ID(3) = "D"
    Mat_ID(4) = "E"
    Mat_ID(5) = "F"
    Mat_ID(6) = "G"
    Mat_ID(7) = "H"
    Mat_ID(8) = "I"
    Mat_ID(9) = "J"
    Mat_ID(10) = "K"
    Mat_ID(11) = "L"
    Mat_ID(12) = "M"
    Mat_ID(13) = "N"
    Mat_ID(14) = "O"
    Mat_ID(15) = "P"
    Mat_ID(16) = "Q"
    Mat_ID(17) = "R"
    Mat_ID(18) = "S"
    Mat_ID(19) = "T"
    Mat_ID(20) = "U"
    Mat_ID(21) = "V"
    Mat_ID(22) = "W"
    Mat_ID(23) = "X"
    Mat_ID(24) = "Y"
    Mat_ID(25) = "Z"
End Sub

Private Function IsExist_Materi(strIBID As String, strclient_brief As String, strmateriname As String, IntDuration As Integer) As Boolean
    Dim rs_exist As New ADODB.Recordset
    Dim strSql As String
    
    strSql = "SELECT * FROM IB_Radio_Material WHERE Client_brief_id='" & strclient_brief & "' AND IB_Id='" & strIBID & "' AND Material_Name='" & strmateriname & "' AND Duration=" & IntDuration & ""
    rs_exist.Open strSql, ConnERP, adLockReadOnly, adLockReadOnly
    
    If rs_exist.RecordCount = 0 Then
        IsExist_Materi = False
    Else
        IsExist_Materi = True
    End If
End Function

Public Function CreateMonthlyQuotation_RD(strBrandCode As String, IntMonthQuotation As Integer, IntYearQuotation As Integer, strMPNumber As String) As Boolean
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
    StrMediaType = "025"
    
    'Get Value  Money Selected Month From Monthly IB (IB Lama Tidak Bisa Tidak Keambil)
    'Dari IB Yang statunya Aktif dan Month dan Year-nya Sama
    strSql = "SELECT IB_ID,Budget FROM IB_Radio_Plan WHERE Month=" & IntMonthQuotation & " AND Year=" & IntYearQuotation
    strSql = strSql & " AND IB_ID IN (SELECT IB_ID FROM IB_Radio WHERE LEFT(IB_ID,4)='" & strBrandCode & "' AND Year=" & IntYearQuotation & " AND Status=1 AND mp_medium_id IS NOT NULL)"
    
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
    
    strSql = "SELECT * FROM IB_Radio_Quotation_Detail WHERE YEAR=" & IntYearQuotation & " AND Month=" & IntMonthQuotation
    strSql = strSql & " AND LEFT(Job_Id,4)='" & strBrandCode & "'"
    
    RsOldMediaQuotation.Open strSql, ConnERP, adOpenKeyset, adLockOptimistic
    
    'Check MQ Monthly
    If Not RsOldMediaQuotation.EOF Then 'If Exist MQ Selected Month ,'Insert To Revision
                                        
        'Generate Revision Value
        strSql = "SELECT MAX(Revision) as LastRevision FROM IB_Radio_Quotation_Detail_Revision WHERE Job_Id='" & RsOldMediaQuotation.Fields("Job_Id").Value & "'"
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
        strSql = "INSERT INTO IB_Radio_Quotation_Detail_Revision VALUES('','"
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
        'Nett Cost
        strSql = strSql & RsOldMediaQuotation.Fields("Nett_Cost").Value & ","
        
        strSql = strSql & RsOldMediaQuotation.Fields("Media_Sptv_Charge").Value & ","
        
        strSql = strSql & RsOldMediaQuotation.Fields("Other_Charge").Value & ","
        
        strSql = strSql & RsOldMediaQuotation.Fields("Bonus").Value & ","
        
        strSql = strSql & RsOldMediaQuotation.Fields("Total_Lintas").Value & ","
        
        strSql = strSql & RsOldMediaQuotation.Fields("Agency_Charge").Value & ",'"
        
        strSql = strSql & RsOldMediaQuotation.Fields("Job_Number_Agency").Value & "',"
        'Grand Total
        strSql = strSql & RsOldMediaQuotation.Fields("Grand_Total").Value & ","
        'Budget
        strSql = strSql & RsOldMediaQuotation.Fields("Budget").Value & ",'"
        'Source IB
        strSql = strSql & RsOldMediaQuotation.Fields("Source_IB").Value & "')"
        
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
        strSql = "DELETE FROM IB_Radio_Quotation_Detail WHERE Job_id='" & StrJobId & "'"
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
            NewMQNumber = Get_RD_Media_Quotation_No(strBrandCode, IntYearQuotation)
        'Jo ID
            StrJobId = strBrandCode & "." & StrMediaType & "." & Right(IntYearQuotation, 2) & Format(IntMonthQuotation, "00")
                
        '------> IB TV Quotation
        strSql = "INSERT INTO IB_Radio_Quot (Client_Brief_ID,IB_ID,Month_IB,Year,Date,Entered_By,Plan_No,Approval_Client) VALUES('',"
        strSql = strSql & "'" & NewMQNumber & "',0," & IntYearQuotation & ",Getdate(),'" & Clear_String(strLogin_FullName) & "',"
        strSql = strSql & "'" & strMPNumber & "',0)"
        
        ConnERP.Execute strSql
        
        '------> IB TV Quotation Detail
        'IB
        strSql = ""
        strSql = "INSERT INTO IB_Radio_Quotation_Detail(Client_Brief_ID,IB_ID,Job_ID,Month,Year,Nett_Cost,Media_Sptv_Charge,Other_Charge,Bonus,Total_Lintas,Agency_Charge,Job_Number_Agency,Grand_Total,Budget,Source_IB) VALUES('','"
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
        strSql = strSql & CurMonthlyBudget & ","
        'Budget
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
    CreateMonthlyQuotation_RD = True
    
    'Add Current User Job
    Add_Current_User_Job 7, strLogin_FullName, NewMQNumber, "", "", "", "", "", ""
    
    
    Exit Function
errLbl:
    CreateMonthlyQuotation_RD = False
End Function
