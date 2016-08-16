Attribute VB_Name = "mdl_GenerateIBPrint"
Option Explicit
Dim Mat_ID(25) As String
Public StrClientBriefId As String
Public rsWeekCommencing As ADODB.Recordset
Public keluar As Boolean
'
Public Function Generate_IB_Print(StrMPMonthlyActivity As String, PlanMonth As Integer) As String
    Dim strSql As String
    Dim rsCekData As New ADODB.Recordset
    Dim IntYear As Integer
    Dim strBrandCode As String
    Dim strIBID As String
    Dim Str_Material_Id As String
    Dim StrCreatedBy As String
    '
    'Declare Local Variable
    
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
    Dim MonthBudget As Integer
    Dim MSC_Paid_Value As Currency
    Dim MSC_Bonus_Value As Currency
    Dim MSC_Paid As Double
    Dim MSC_Bonus As Double
    Dim rsMPInfoGeneral As New ADODB.Recordset
    Dim rsMPMaterial As New ADODB.Recordset
    'Dim rsMPPlan As New ADODB.Recordset
    Dim rsMPPlanDetail_materi As New ADODB.Recordset
    Dim int_row As Integer
    Dim rsBrandVariant As New ADODB.Recordset
    Dim MSC_Paid_On_Flag As Integer
    Dim Msc_paid_Value_detail As Currency
    Dim Other_Cost As Currency
    Dim Int_materi As Integer
    Dim Int_Insertion As Integer
    Dim Nett As Currency
    Dim Gross As Currency
    
    'On Error GoTo errLbl
    
    
    'Cancel Previous IB
    If Not Cancel_IB(StrMPMonthlyActivity, PlanMonth, False) Then
        Generate_IB_Print = "Error"
        Exit Function
    End If

    
    Call Initial_Data
    
    'Populate Variable General
    strSql = " SELECT dbo.MP_Master.created_By, dbo.MP_Master.mp_number, dbo.MP_Master.brand_code, dbo.MP_Master.brand_name, dbo.MP_Master.[year],dbo.MP_Activity.activity_type, dbo.MP_Activity.activity_desc, "
    strSql = strSql & " dbo.MP_Activity.brand_variant_code, dbo.MP_Activity.brand_variant_name,dbo.MP_Activity.Original_Brand_code,dbo.MP_Activity.Original_Brand_Name, dbo.MP_Activity.target_audience_code, dbo.MP_Activity.target_audience,dbo.MP_Activity.brand_target,"
    strSql = strSql & " dbo.MP_Activity.brand_target, dbo.MP_Medium.medium_code, dbo.MP_Monthly_Activity.month_number, dbo.MP_Monthly_Activity.month_name,"
    strSql = strSql & " dbo.MP_Monthly_Activity.min_Budget Budget,isnull(dbo.MP_Monthly_Activity.other_cost,0) other_cost , dbo.MP_Monthly_Activity.gross_budget, dbo.MP_Monthly_Activity.MSC_Paid,"
    strSql = strSql & " dbo.MP_Monthly_Activity.MSC_Paid_Value,dbo.MP_Monthly_Activity.MSC_Paid_on_flag, dbo.MP_Monthly_Activity.MSC_Bonus, dbo.MP_Monthly_Activity.MSC_Bonus_On_Flag, dbo.MP_Monthly_Activity.MSC_Bonus_Value"
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
    
        Generate_IB_Print = "Error"
        Exit Function
    End If
    
    StrCreatedBy = IIf(IsNull(rsMPInfoGeneral.Fields("Created_By").Value), "", rsMPInfoGeneral.Fields("Created_By").Value)
    strBrandCode = IIf(IsNull(rsMPInfoGeneral.Fields("Original_Brand_Code").Value), "", rsMPInfoGeneral.Fields("Original_Brand_Code").Value)
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
    '======================================================
    
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
    MSC_Paid_Value = rsMPInfoGeneral.Fields("Msc_Paid_Value").Value
    MSC_Bonus_Value = rsMPInfoGeneral.Fields("Msc_Bonus_Value").Value
    MSC_Paid = rsMPInfoGeneral.Fields("Msc_Paid").Value
    MSC_Bonus = rsMPInfoGeneral.Fields("Msc_Bonus").Value
    MSC_Paid_On_Flag = rsMPInfoGeneral.Fields("MSC_Paid_on_flag").Value
    Other_Cost = rsMPInfoGeneral.Fields("other_cost").Value
    
    rsMPInfoGeneral.Close
    Set rsMPInfoGeneral = Nothing
    '--------------------------------------------------

    'tentukan year
    IntYear = CInt(Mid(StrMPMonthlyActivity, 11, 4))
    MonthBudget = PlanMonth
    
    'generate Client Brief Id
    StrClientBriefId = Get_Brief_ID(strBrandCode, IntYear)
    
    
    'Generate IB ID
    strIBID = GetPrintID(strBrandCode, IntYear)
    

    strSql = " select d.week_commencing as Week_Commencing,c.version as Version ,e.media_name as Media,"
    strSql = strSql & " Case c.print_ismmc"
    strSql = strSql & " when 0 then ltrim(rtrim(c.print_paper_name) + ', '+ cast(c.print_mmc_col as varchar) + ' x ' + cast(c.print_mmc_size as varchar) + ' mm , ' + rtrim(c.print_color_name))"
    strSql = strSql & " when 1 then cast(print_mmc_col as varchar) + ' x ' + cast(print_mmc_size as varchar) + ' mm, ' + rtrim(print_paper_name) + ', ' + rtrim(print_color_name)"
    strSql = strSql & " end Dimension,"
    strSql = strSql & " c.spot_type as Spot_Type,sum(d.spot)as Spot,e.print_code,"
    strSql = strSql & " c.print_size_code,c.print_color_code,c.print_paper_code,"
    strSql = strSql & " c.print_mmc_col , c.print_mmc_size, c.print_min_size, d.Nett_Rate, d.Gross_Rate,c.print_ismmc"
    strSql = strSql & " from MP_Plan_Dimension c"
    strSql = strSql & " left join mp_insertion d on c.mp_plan_dim_id=d.mp_plan_dim_id"
    strSql = strSql & " left join mp_medium_detail e on e.mp_medium_detail_id =  c.mp_medium_detail_id"
    strSql = strSql & " where e.mp_medium_id ='" & StrMPMonthlyActivity & "' and d.month=" & PlanMonth
    strSql = strSql & " group by d.week_commencing,c.spot_type,e.print_code,e.media_name,c.version,c.print_size_code,"
    strSql = strSql & " c.print_paper_code,c.print_paper_name,c.print_color_code,c.print_color_name,c.print_ismmc,c.print_mmc_col,"
    strSql = strSql & " c.print_mmc_size , c.print_min_size, d.Nett_Rate, d.Gross_Rate"

    Set rsWeekCommencing = New ADODB.Recordset
    rsWeekCommencing.CursorLocation = adUseClient
    rsWeekCommencing.Open strSql, ConnERP, adOpenForwardOnly, adLockReadOnly
    
    If Not rsWeekCommencing.EOF Then
        Frm_Date_Commencing.show 1
    End If
        
    If keluar Then     'Cancel
        Unload Frm_Date_Commencing
        Generate_IB_Print = "Cancel"
        Exit Function
    End If
    
    '=========== Insert to IB Print =========
    strSql = "INSERT INTO IB_Print(Client_Brief_Id, IB_Id, Entered_Date,"
    strSql = strSql & " Entered_By, Primary_Target, Secondary_Target, Any_Consideration,"
    strSql = strSql & " Attachment, Approval_Client_Flag, Grand_Total, [Date], Approval_Date,"
    strSql = strSql & " Plan_No, Cluster_Code, Brand_Variant_Code, Brand_Variant_Name,mp_medium_id,month_number) VALUES("
    'Client_Brief_Id
    strSql = strSql & "'" & StrClientBriefId & "',"
    'IB_Id
    strSql = strSql & "'" & strIBID & "',"
    'Entered_Date
    strSql = strSql & "Getdate(),"
    'Entered_By
    strSql = strSql & "'" & StrCreatedBy & "',"
    'Primary_Target
    strSql = strSql & "'" & StrPrimaryTarget & "',"
    'Secondary_Target
    strSql = strSql & "'" & StrSecondaryTarget & "',"
    'Any_Consideration
    strSql = strSql & "'',"
    'Attachment
    strSql = strSql & "'',"
    'Approval_Client_Flag
    strSql = strSql & "1,"
    'Grand_Total
    strSql = strSql & MonthlyNettBudget + MSC_Paid_Value + Other_Cost & ","
    '[Date]
    strSql = strSql & "Getdate(),"
    'Approval_Date
    strSql = strSql & "GetDate(),"
    'Plan_No
    strSql = strSql & "'" & PlanNo & "',"
    'Cluster_Code
    strSql = strSql & SecondaryTargetCode & ","
    'Brand_Variant_Code
    strSql = strSql & "'" & BrandVariantCode & "',"
    'Brand_Variant_Name
    strSql = strSql & "'" & BrandVariantName & "',"
    strSql = strSql & "'" & StrMPMonthlyActivity & "',"
    strSql = strSql & PlanMonth
    strSql = strSql & ")"
        
    ConnERP.Execute strSql
    
    
    ' ================================================

    '============ Insert To IB_Print_Plan ============
    strSql = "INSERT INTO IB_Print_Plan(Client_Brief_Id, IB_Id, [Month], Budget) VALUES("
    'Client_Brief_Id
    strSql = strSql & "'" & StrClientBriefId & "',"
    'IB_Id
    strSql = strSql & "'" & strIBID & "',"
    '[Month]
    strSql = strSql & PlanMonth & ","
    'Budget
    strSql = strSql & Format(MonthlyNettBudget + MSC_Paid_Value + Other_Cost, "####0") & " )"
    
    ConnERP.Execute strSql
    '============================================================


    '======== Insert to IB Print Material & IB Print Plan Title =====

  '  strSQL = "SELECT  Distinct MP_Medium_Detail.MP_Medium_Detail_Id,MP_Medium_Detail.MP_Medium_Id,MP_Medium_Detail.Print_Code,"
   ' strSQL = strSQL & " MP_Medium_Detail.Print_Code FROM MP_Medium_Detail"
    'strSQL = strSQL & " LEFT JOIN "
    'strSQL = strSQL & " MP_Plan_Dimension ON MP_Medium_Detail.MP_Medium_Detail_Id=MP_Plan_Dimension.MP_Medium_Detail_Id "
    'strSQL = strSQL & " WHERE MP_Medium_Detail.MP_Medium_Id='" & StrMPMonthlyActivity & "' "

   ' rsMPPlan.Open strSQL, ConnERP, adOpenStatic, adLockReadOnly

   ' While Not rsMPPlan.EOF

        For int_row = 0 To Frm_Date_Commencing.Grd_Tanggal.Rows - 1
            If Frm_Date_Commencing.Grd_Tanggal.TextMatrix(int_row, Frm_Date_Commencing.Grd_Tanggal.cols - 1) <> "" Then
                If IsExist_Materi(strIBID, StrClientBriefId, Trim(Frm_Date_Commencing.Grd_Tanggal.TextMatrix(int_row, 1))) = False Then

                    Str_Material_Id = ""
                    Str_Material_Id = Get_New_Material_ID(strIBID)

                    strSql = "INSERT INTO IB_Print_Material(Client_Brief_Id, IB_Id,Material_Code, Material) VALUES("
                    ' Client Brief Id
                    strSql = strSql & "'" & StrClientBriefId & "',"
                    ' IB_id
                    strSql = strSql & "'" & strIBID & "',"
                    ' Material Code
                    strSql = strSql & "'" & Str_Material_Id & "',"
                    ' Material Name
                    strSql = strSql & "'" & Trim(Frm_Date_Commencing.Grd_Tanggal.TextMatrix(int_row, 1)) & "')"
                    ConnERP.Execute strSql

                Else
                    strSql = "SELECT Material_Code FROM IB_Print_Material WHERE Client_Brief_Id='" & StrClientBriefId & "' AND IB_id='" & strIBID & "' AND Material='" & Trim(Frm_Date_Commencing.Grd_Tanggal.TextMatrix(int_row, 1)) & "'"
                    rsMPMaterial.Open strSql, ConnERP, adOpenStatic, adLockReadOnly

                    If Not rsMPMaterial.EOF Then
                        Str_Material_Id = rsMPMaterial.Fields("Material_Code").Value
                        rsMPMaterial.Close
                    End If
            
                End If
                    
                If UCase(Left(Trim(Frm_Date_Commencing.Grd_Tanggal.TextMatrix(int_row, 4)), 1)) = "S" Then
                    'Sponsorship
                    Int_materi = 0
                    Int_Insertion = 0
                    
                    If Trim(Frm_Date_Commencing.Grd_Tanggal.TextMatrix(int_row, 10)) <> "" And Trim(Frm_Date_Commencing.Grd_Tanggal.TextMatrix(int_row, 11)) <> "" Then
                        Int_materi = Frm_Date_Commencing.Grd_Tanggal.TextMatrix(int_row, 10) * Frm_Date_Commencing.Grd_Tanggal.TextMatrix(int_row, 11)
                    End If
                    
                    If Int_materi < CInt(Frm_Date_Commencing.Grd_Tanggal.TextMatrix(int_row, 12)) Then
                        Int_materi = CInt(Frm_Date_Commencing.Grd_Tanggal.TextMatrix(int_row, 12))
                    End If
                    
                    If Int_materi = 0 Then
                        Int_materi = 1
                    End If
                    
                    Int_Insertion = Frm_Date_Commencing.Grd_Tanggal.TextMatrix(int_row, 5)
                    
                    Nett = Frm_Date_Commencing.Grd_Tanggal.TextMatrix(int_row, 13) / (Int_Insertion * Int_materi)
                    Gross = Frm_Date_Commencing.Grd_Tanggal.TextMatrix(int_row, 14) / (Int_materi * Int_Insertion)
                Else
                    'Regular
                    Int_materi = 0
                    Int_Insertion = 0
                    
                    If Trim(Frm_Date_Commencing.Grd_Tanggal.TextMatrix(int_row, 10)) <> "" And Trim(Frm_Date_Commencing.Grd_Tanggal.TextMatrix(int_row, 11)) <> "" Then
                        Int_materi = Frm_Date_Commencing.Grd_Tanggal.TextMatrix(int_row, 10) * Frm_Date_Commencing.Grd_Tanggal.TextMatrix(int_row, 11)
                    End If
                    
                    If Int_materi < CInt(Frm_Date_Commencing.Grd_Tanggal.TextMatrix(int_row, 12)) Then
                        Int_materi = CInt(Frm_Date_Commencing.Grd_Tanggal.TextMatrix(int_row, 12))
                    End If
                    
                    If Int_materi = 0 Then
                        Int_materi = 1
                    End If
                    
                    Int_Insertion = Frm_Date_Commencing.Grd_Tanggal.TextMatrix(int_row, 5)
                    
                    Nett = Frm_Date_Commencing.Grd_Tanggal.TextMatrix(int_row, 13) / (Int_Insertion * Int_materi)
                    Gross = Frm_Date_Commencing.Grd_Tanggal.TextMatrix(int_row, 14) / (Int_Insertion * Int_materi)
                End If
                    
                Select Case MSC_Paid_On_Flag
                    Case 1, 4
                        Msc_paid_Value_detail = CCur(Format(Frm_Date_Commencing.Grd_Tanggal.TextMatrix(int_row, Frm_Date_Commencing.Grd_Tanggal.cols - 2) * Int_materi * Nett * MSC_Paid / 100, "####0"))
                    Case 0, 2, 3
                        Msc_paid_Value_detail = CCur(Format(Frm_Date_Commencing.Grd_Tanggal.TextMatrix(int_row, Frm_Date_Commencing.Grd_Tanggal.cols - 2) * Int_materi * Gross * MSC_Paid / 100, "####0"))
                End Select

                If Frm_Date_Commencing.Grd_Tanggal.TextMatrix(int_row, 6) = Frm_Date_Commencing.Grd_Tanggal.TextMatrix(int_row - 1, 6) And Frm_Date_Commencing.Grd_Tanggal.TextMatrix(int_row, 10) = Frm_Date_Commencing.Grd_Tanggal.TextMatrix(int_row - 1, 10) And Frm_Date_Commencing.Grd_Tanggal.TextMatrix(int_row, 11) = Frm_Date_Commencing.Grd_Tanggal.TextMatrix(int_row - 1, 11) And Frm_Date_Commencing.Grd_Tanggal.TextMatrix(int_row, 8) = Frm_Date_Commencing.Grd_Tanggal.TextMatrix(int_row - 1, 8) And Frm_Date_Commencing.Grd_Tanggal.TextMatrix(int_row, 9) = Frm_Date_Commencing.Grd_Tanggal.TextMatrix(int_row - 1, 9) And Frm_Date_Commencing.Grd_Tanggal.TextMatrix(int_row, 1) = Frm_Date_Commencing.Grd_Tanggal.TextMatrix(int_row - 1, 1) Then
                'If 1 = 2 Then
                    '===== Update into IB Print Plan Title Jika key-nya sama =====
                    strSql = "UPDATE IB_Print_Plan_Title "
                    strSql = strSql & " SET Date_W_C_Bonus=''"
                    strSql = strSql & " ,Date_W_C_Paid='" & Trim(Frm_Date_Commencing.Grd_Tanggal.TextMatrix(int_row, Frm_Date_Commencing.Grd_Tanggal.cols - 3)) & "," & Trim(Frm_Date_Commencing.Grd_Tanggal.TextMatrix(int_row - 1, Frm_Date_Commencing.Grd_Tanggal.cols - 3)) & "'"
                    strSql = strSql & " ,Date_W_C='" & Frm_Date_Commencing.Grd_Tanggal.TextMatrix(int_row, Frm_Date_Commencing.Grd_Tanggal.cols - 3) & "," & Frm_Date_Commencing.Grd_Tanggal.TextMatrix(int_row - 1, Frm_Date_Commencing.Grd_Tanggal.cols - 3) & "'"
                    strSql = strSql & " ,Satuan= " & CInt(Frm_Date_Commencing.Grd_Tanggal.TextMatrix(int_row, 5)) + CInt(Frm_Date_Commencing.Grd_Tanggal.TextMatrix(int_row - 1, 5)) & ""
                    strSql = strSql & " ,Insertion=" & CInt(Frm_Date_Commencing.Grd_Tanggal.TextMatrix(int_row, 5)) + CInt(Frm_Date_Commencing.Grd_Tanggal.TextMatrix(int_row - 1, 5)) & ""
                    strSql = strSql & " ,Insertion_Bonus=0"
                    strSql = strSql & " ,Insertion_Paid=" & CInt(Frm_Date_Commencing.Grd_Tanggal.TextMatrix(int_row, Frm_Date_Commencing.Grd_Tanggal.cols - 2)) + CInt(Frm_Date_Commencing.Grd_Tanggal.TextMatrix(int_row - 1, Frm_Date_Commencing.Grd_Tanggal.cols - 2)) & ""
                    strSql = strSql & " ,Total_Incl_MSC= " & CCur(Format(Frm_Date_Commencing.Grd_Tanggal.TextMatrix(int_row, Frm_Date_Commencing.Grd_Tanggal.cols - 2) * Int_materi * Nett + Msc_paid_Value_detail, "####0")) + CCur(Format(Frm_Date_Commencing.Grd_Tanggal.TextMatrix(int_row - 1, Frm_Date_Commencing.Grd_Tanggal.cols - 2) * Int_materi * Nett + Msc_paid_Value_detail, "####0"))
                    strSql = strSql & " ,Total_Nett= " & CCur(Format(Frm_Date_Commencing.Grd_Tanggal.TextMatrix(int_row, Frm_Date_Commencing.Grd_Tanggal.cols - 2) * Int_materi * Nett, "####0")) + CCur(Format(Frm_Date_Commencing.Grd_Tanggal.TextMatrix(int_row - 1, Frm_Date_Commencing.Grd_Tanggal.cols - 2) * Int_materi * Nett, "####0"))
                    strSql = strSql & " ,Total_Gross_Paid=" & CCur(Format(Frm_Date_Commencing.Grd_Tanggal.TextMatrix(int_row, Frm_Date_Commencing.Grd_Tanggal.cols - 2) * Gross * Int_materi, "####0")) + CCur(Format(Frm_Date_Commencing.Grd_Tanggal.TextMatrix(int_row - 1, Frm_Date_Commencing.Grd_Tanggal.cols - 2) * Int_materi * Gross, "####0"))
                    strSql = strSql & " ,Total_Nett_Bonus=0"
                    strSql = strSql & " ,Total_Gross_Bonus=0"
                    strSql = strSql & " WHERE Client_Brief_ID='" & StrClientBriefId & "' AND"
                    strSql = strSql & " IB_Id='" & strIBID & "' AND"
                    strSql = strSql & " Month=" & PlanMonth & " AND"
                    strSql = strSql & " Print_Code='" & Trim(Frm_Date_Commencing.Grd_Tanggal.TextMatrix(int_row, 6)) & "' AND"
                    strSql = strSql & " MM='" & Trim(Frm_Date_Commencing.Grd_Tanggal.TextMatrix(int_row, 11)) & "' AND"
                    strSql = strSql & " CL='" & Trim(Frm_Date_Commencing.Grd_Tanggal.TextMatrix(int_row, 10)) & "' AND"
                    strSql = strSql & " Color='" & Trim(Frm_Date_Commencing.Grd_Tanggal.TextMatrix(int_row, 8)) & "' AND"
                    strSql = strSql & " Paper='" & Trim(Frm_Date_Commencing.Grd_Tanggal.TextMatrix(int_row, 9)) & "'"
    
                    ConnERP.Execute strSql

                Else
                  
                    '===== Insert into IB Print Plan Title =====
                    strSql = "INSERT INTO IB_Print_Plan_Title(Client_Brief_Id, IB_Id,Month,Print_Code,Size,Color,Paper,Material,"
                    strSql = strSql & " Date_W_C_Bonus,Date_W_C_Paid,Date_W_C,satuan,MM,CL,Nett_Rate,Gross_Rate,"
                    strSql = strSql & " Insertion,Insertion_Bonus,Insertion_Paid,Total_Incl_MSC,Total_Nett,"
                    strSql = strSql & " Total_Gross_Paid,Total_Nett_Bonus,Total_Gross_Bonus,Min_Size,Regular_Flag) VALUES("
                    ' Brief ID
                    strSql = strSql & "'" & StrClientBriefId & "',"
                    ' IB ID
                    strSql = strSql & "'" & strIBID & "',"
                    ' Month
                    strSql = strSql & "'" & PlanMonth & "',"
                    ' Print Code
                    strSql = strSql & "'" & Trim(Frm_Date_Commencing.Grd_Tanggal.TextMatrix(int_row, 6)) & "',"
                    ' Size
                    'If rsWeekCommencing("print_ismmc").Value = 1 Then
                    If Trim(Frm_Date_Commencing.Grd_Tanggal.TextMatrix(int_row, 15)) = 1 Then
                        strSql = strSql & "'" & Trim(Frm_Date_Commencing.Grd_Tanggal.TextMatrix(int_row, 10)) & " x " & Trim(Frm_Date_Commencing.Grd_Tanggal.TextMatrix(int_row, 11)) & "',"
                    Else
                        strSql = strSql & "1,"
                    End If
                    ' Color
                    strSql = strSql & "'" & Trim(Frm_Date_Commencing.Grd_Tanggal.TextMatrix(int_row, 8)) & "',"
                    ' Paper
                    strSql = strSql & "'" & Trim(Frm_Date_Commencing.Grd_Tanggal.TextMatrix(int_row, 9)) & "',"
                    ' Material
                    strSql = strSql & "'" & Str_Material_Id & "',"
                    ' Date W C Bonus
                    strSql = strSql & "'',"
                    ' Date W C Paid
                    strSql = strSql & "'" & Trim(Frm_Date_Commencing.Grd_Tanggal.TextMatrix(int_row, Frm_Date_Commencing.Grd_Tanggal.cols - 3)) & "',"
                    ' Date W C
                    strSql = strSql & "'" & Trim(Frm_Date_Commencing.Grd_Tanggal.TextMatrix(int_row, Frm_Date_Commencing.Grd_Tanggal.cols - 3)) & "',"
                    ' satuan
                    strSql = strSql & "'" & Trim(Frm_Date_Commencing.Grd_Tanggal.TextMatrix(int_row, 7)) & "',"
                    ' MM
                    strSql = strSql & "'" & Trim(Frm_Date_Commencing.Grd_Tanggal.TextMatrix(int_row, 11)) & "',"
                    ' CL
                    strSql = strSql & "'" & Trim(Frm_Date_Commencing.Grd_Tanggal.TextMatrix(int_row, 10)) & "',"
                    ' Nett Rate
                    strSql = strSql & Nett & ","
                    ' Gross Rate
                    strSql = strSql & Gross & ","
                    ' Insertion
                    strSql = strSql & " " & CInt(Frm_Date_Commencing.Grd_Tanggal.TextMatrix(int_row, Frm_Date_Commencing.Grd_Tanggal.cols - 2)) & ","
                    ' Insertion Bonus
                    strSql = strSql & "0,"
                    ' Insertion Paid
                     strSql = strSql & " " & CInt(Frm_Date_Commencing.Grd_Tanggal.TextMatrix(int_row, Frm_Date_Commencing.Grd_Tanggal.cols - 2)) & ","
                    ' Total Include MSC
                    'strSql = strSql & CCur(Format(FRm_Date_Commencing.Grd_Tanggal.TextMatrix(int_row, FRm_Date_Commencing.Grd_Tanggal.Cols - 2) * FRm_Date_Commencing.Grd_Tanggal.TextMatrix(int_row, 5) * Nett + Msc_paid_Value_detail, "####0")) & " ,"
                    strSql = strSql & CCur(Format(Frm_Date_Commencing.Grd_Tanggal.TextMatrix(int_row, Frm_Date_Commencing.Grd_Tanggal.cols - 2) * Int_materi * Nett + Msc_paid_Value_detail, "####0")) & " ,"
                    ' Total Nett
                    strSql = strSql & CCur(Format(Frm_Date_Commencing.Grd_Tanggal.TextMatrix(int_row, Frm_Date_Commencing.Grd_Tanggal.cols - 2) * Int_materi * Nett, "####0")) & ","
                    ' Total Gross Paid
                    strSql = strSql & CCur(Format(Frm_Date_Commencing.Grd_Tanggal.TextMatrix(int_row, Frm_Date_Commencing.Grd_Tanggal.cols - 2) * Int_materi * Gross, "####0")) & ","
                    ' Total Nett Bonus
                    strSql = strSql & "0,"
                    ' Total Gross Bonus
                    strSql = strSql & "0,"
                    ' Min Size
                    strSql = strSql & "'" & Frm_Date_Commencing.Grd_Tanggal.TextMatrix(int_row, 12) & "',"
                    ' Reguler Flag
                    strSql = strSql & "'" & UCase(Left(Trim(Frm_Date_Commencing.Grd_Tanggal.TextMatrix(int_row, 4)), 1)) & "')"
    
                    ConnERP.Execute strSql
                    
                    'MsgBox "OK " & Trim(FRm_Date_Commencing.Grd_Tanggal.TextMatrix(int_row, 6))

                End If
            End If
        Next

      '  rsMPPlan.MoveNext
    'Wend

   ' rsMPPlan.Close

    Unload Frm_Date_Commencing
    
    rsWeekCommencing.Close
    Set rsWeekCommencing = Nothing
    
    Generate_IB_Print = strIBID

    '-------------- Submit Data to Budget Control Plan & Generate Monthly Quotation If BU1 ------------------
    If Is_Special_Brand(strBrandCode) Then
        'SubmitDataToMQTV(IBID)--->Generate Quotation
        
        Call CreateMonthlyQuotation_PR(strBrandCode, PlanMonth, PlanYear, PlanNo)
        
'        'SubmitdatatoBCMPBU1(Brand_code,Year,month,Medium,Gross,Nett)
'        StrSQL = "INSERT INTO Budget_Control_Detail_BU1(mp_medium_id,Month_Number,Year_Budget,Month_Budget,Medium,Brand_code,Gross_Budget,Nett_Budget) VALUES ("
'        'mp_medium_id,Month_Number
'        StrSQL = StrSQL & "'" & StrMPMonthlyActivity & "'," & PlanMonth & ","
'        'Year_Budget,Month_Number
'        StrSQL = StrSQL & PlanYear & "," & PlanMonth & ","
'        'Medium, Brand_Code
'        StrSQL = StrSQL & "'PR','" & strBrandCode & "',"
'        'Gross_Budget, Nett_Bugdet
'        StrSQL = StrSQL & "0," & MonthlyNettBudget
'        StrSQL = StrSQL & ")"
'        ConnERP.Execute StrSQL
    Else
        'SubmitdatatoBCMPBU2(Brand_code,Year,month,Medium,Gross,Nett)
'        StrSQL = "INSERT INTO Budget_Control_Detail_BU2(mp_medium_id,Month_Number,Year_Budget,Month_Budget,Medium,Brand_code,Gross_Budget,Nett_Budget) VALUES ("
'        'mp_medium_id,Month_Number
'        StrSQL = StrSQL & "'" & StrMPMonthlyActivity & "'," & PlanMonth & ","
'        'Year_Budget,Month_Number
'        StrSQL = StrSQL & PlanYear & "," & PlanMonth & ","
'        'Medium, Brand_Code
'        StrSQL = StrSQL & "'PR','" & strBrandCode & "',"
'        'Gross_Budget, Nett_Bugdet
'        StrSQL = StrSQL & "0," & MonthlyNettBudget
'        StrSQL = StrSQL & ")"
'        ConnERP.Execute StrSQL
    End If
    '--------------------------------------------------------------------
    
    'Add Cuurent User Job
    Add_Current_User_Job 3, strLogin_FullName, strIBID, "", "", "", "", "", ""
    
    Exit Function
    
errLbl:
    
    Generate_IB_Print = "Error"
    
End Function

Function GetPrintID(strBrand As String, tahun As Integer) As String
        Dim rs_Max As New ADODB.Recordset
        Dim Induk As String
        Dim No As Integer
        Dim nomor As String
        Dim Sql As String
        
        '**************************************************
            'prosedur untuk mendapatkan IB_ID
        '**************************************************
            
        'cari di reusable WHERE year AND brand
        '=================================================
        Sql = "SELECT  * FROM Reuseable_IB_ID_Print WHERE year=" & tahun & " AND brand_code = '" & strBrand & "' ORDER BY ib_id "
        rs_Max.Open Sql, ConnERP, adOpenStatic, adLockReadOnly, adCmdText
        
        'jika ada GetPrintID = record di reusable
        If rs_Max.RecordCount > 0 Then
        
            GetPrintID = rs_Max("ib_id")
            
            ' DELETE di reusable
            Sql = "DELETE FROM Reuseable_IB_ID_Print WHERE year=" & rs_Max("year") & " AND Brand_Code= '" & rs_Max("Brand_Code") & "' AND IB_ID= '" & rs_Max("IB_ID") & "'"
            ConnERP.Execute Sql
            
          
            rs_Max.Close
            Set rs_Max = Nothing
            
            Exit Function
        End If
        
        rs_Max.Close
        Set rs_Max = Nothing
        'END cari di reusable WHERE year AND brand
        '=================================================
        
        'cari Print Media Induk
        Sql = "SELECT  * FROM media_type WHERE rtrim(ltrim(media_type_name)) = 'Print Media Induk'"
        rs_Max.Open Sql, ConnERP, adOpenStatic, adLockReadOnly
        Induk = rs_Max(0)
        rs_Max.Close
        Set rs_Max = Nothing
        
        'cari nomor WHERE year AND brand code
        Sql = "SELECT  last_number FROM Last_IB_ID_Print WHERE year=" & tahun & " AND brand_code='" & strBrand & "'"
        rs_Max.Open Sql, ConnERP, adOpenStatic, adLockReadOnly
        If rs_Max.RecordCount > 0 Then
             No = rs_Max(0) + 1
        Else
             No = 1
        End If
        rs_Max.Close
        Set rs_Max = Nothing
        
        If No > 99 Then
            'translate nomor last
            nomor = GetTranslateNumber(No)
            GetPrintID = strBrand & "." & Trim(Induk) & "." & Mid(StrClientBriefId, 6, 2) & Trim(nomor)
        Else
            If Len(Trim(No)) = 1 Then nomor = "0" + CStr(No) Else nomor = CStr(No)
            GetPrintID = strBrand & "." & Trim(Induk) & "." & Mid(StrClientBriefId, 6, 2) & Trim(nomor)
        End If
        
        'UPDATE di last Number
        If No = 1 Then
            Sql = "INSERT INTO last_ib_id_print(Brand_Code, Year, Last_Number)"
            Sql = Sql & " VALUES('" & strBrand & "'," & tahun & ",1)"
            ConnERP.Execute Sql
        Else
            Sql = "UPDATE last_ib_id_print SET Last_Number=" & No
            Sql = Sql & " WHERE brand_code='" & strBrand & "' AND  year=" & tahun
            ConnERP.Execute Sql
        End If
            
End Function

'dihapus jika disatukan
'==========================================================
Private Function GetTranslateNumber(IntLastNumber As Integer) As String
    '====================================================
    'Convert nilai LastNumber menjadi yang ada di catalog
    '====================================================
    Dim strSql As String
    Dim RsTranslate As New ADODB.Recordset
    
    strSql = "SELECT Translate_To FROM Translate_Number WHERE Number = " & IntLastNumber
    RsTranslate.Open strSql, ConnERP, adOpenForwardOnly, adLockReadOnly
    If Not RsTranslate.EOF Then
        'ada recordnya
        GetTranslateNumber = RsTranslate.Fields("Translate_To").Value
    Else
        'tidak ada recordnya
        
    End If
    RsTranslate.Close
    Set RsTranslate = Nothing
End Function
'==========================================================

Private Function Get_New_Material_ID(strIBID As String) As String
    Dim Pos As Integer
    Dim rs As New ADODB.Recordset
    Dim Ada As Boolean
    Dim Sql As String
    
    
    Sql = "SELECT Material_Code from IB_Print_Material WHERE IB_Id='" & strIBID & "'"
    rs.Open Sql, ConnERP, adOpenStatic, adLockReadOnly
    
    With rs
        Pos = 0
        Ada = False
        Do While .EOF = False
            If .Fields("Material_Code").Value <> Mat_ID(Pos) Then
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
Private Function IsExist_Materi(strIBID As String, strclient_brief As String, strmateriname As String) As Boolean
    Dim rs_exist As New ADODB.Recordset
    Dim strSql As String
    Dim isExist As Boolean

    strSql = "SELECT * FROM IB_Print_Material WHERE Client_brief_id='" & strclient_brief & "' AND IB_Id='" & strIBID & "' AND Material='" & strmateriname & "'"
    rs_exist.Open strSql, ConnERP, adLockReadOnly, adLockReadOnly
    
    If rs_exist.RecordCount = 0 Then
        IsExist_Materi = False
    Else
        IsExist_Materi = True
    End If
End Function



Public Function CreateMonthlyQuotation_PR(strBrandCode As String, IntMonthQuotation As Integer, IntYearQuotation As Integer, strMPNumber As String) As Boolean
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
    StrMediaType = "015"
    
    'Get Value  Money Selected Month From Monthly IB (IB Lama Tidak Bisa Tidak Keambil)
    'Dari IB Yang statunya Aktif dan Month dan Year-nya Sama
    strSql = "SELECT IB_ID,Budget FROM IB_Print_Plan WHERE Month=" & IntMonthQuotation
    strSql = strSql & " AND IB_ID IN (SELECT IB_ID FROM IB_Print WHERE LEFT(IB_ID,4)='" & strBrandCode & "' AND Substring(IB_ID,10,2)='" & Right(IntYearQuotation, 2) & "' AND Status=1 AND mp_medium_id IS NOT NULL)"
    
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
    'Select where Month,Year,Brand
    strSql = "SELECT * FROM IB_Print_Quotation_Detail WHERE YEAR=" & IntYearQuotation & " AND Month=" & IntMonthQuotation
    strSql = strSql & " AND LEFT(Job_Id,4)='" & strBrandCode & "'"
    
    RsOldMediaQuotation.Open strSql, ConnERP, adOpenKeyset, adLockOptimistic
    
    'Check MQ Monthly
    If Not RsOldMediaQuotation.EOF Then 'If Exist MQ Selected Month ,'Insert To Revision
                                        
        'Generate Revision Value
        strSql = "SELECT MAX(Revision) as LastRevision FROM IB_Print_Quotation_Detail_Revision WHERE Job_Id='" & RsOldMediaQuotation.Fields("Job_Id").Value & "'"
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
        strSql = "INSERT INTO IB_Print_Quotation_Detail_Revision VALUES('','"
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
        'Media_Sptv_Charge
        strSql = strSql & RsOldMediaQuotation.Fields("Media_Sptv_Charge").Value & ","
        'Other_Charge
        strSql = strSql & RsOldMediaQuotation.Fields("Other_Charge").Value & ","
        'Bonus
        strSql = strSql & RsOldMediaQuotation.Fields("Bonus").Value & ","
        'Total_Lintas
        strSql = strSql & RsOldMediaQuotation.Fields("Total_Lintas").Value & ","
        'Agency_Charge
        strSql = strSql & RsOldMediaQuotation.Fields("Agency_Charge").Value & ",'"
        'Job_Number_Agency
        strSql = strSql & RsOldMediaQuotation.Fields("Job_Number_Agency").Value & "',"
        'Grand_Total
        strSql = strSql & RsOldMediaQuotation.Fields("Grand_Total").Value & ",'"
        'Source_IB
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
        strSql = "DELETE FROM IB_Print_Quotation_Detail WHERE Job_id='" & StrJobId & "'"
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
            NewMQNumber = Get_PR_Media_Quotation_No(strBrandCode, IntYearQuotation)
        'Jo ID
            StrJobId = strBrandCode & "." & StrMediaType & "." & Right(IntYearQuotation, 2) & Format(IntMonthQuotation, "00")
                
        '------> IB TV Quotation
        strSql = "INSERT INTO IB_Print_Quotation (Client_Brief_ID,IB_ID,Month_IB,Year,Date,Entered_By,Plan_No) VALUES('',"
        strSql = strSql & "'" & NewMQNumber & "'," & IntMonthQuotation & "," & IntYearQuotation & ",Getdate(),'" & Clear_String(strLogin_FullName) & "',"
        strSql = strSql & "'" & strMPNumber & "')"
        
        ConnERP.Execute strSql
        
        '------> IB TV Quotation Detail
        'IB
        strSql = ""
        strSql = "INSERT INTO IB_Print_Quotation_Detail(Client_Brief_ID,IB_ID,Job_ID,Month,Year,Nett_Cost,Media_Sptv_Charge,Other_Charge,Bonus,Total_Lintas,Agency_Charge,Job_Number_Agency,Grand_Total,Source_IB) VALUES('  ','"
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
    CreateMonthlyQuotation_PR = True
    
    'Add Current User Job
    Add_Current_User_Job 8, strLogin_FullName, NewMQNumber, "", "", "", "", "", ""
    
    
    Exit Function
errLbl:
    CreateMonthlyQuotation_PR = False
End Function
