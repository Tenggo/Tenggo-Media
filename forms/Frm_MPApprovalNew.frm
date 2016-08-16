VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "mscomctl.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form Frm_MPApprovalNew 
   Caption         =   "Client Approval"
   ClientHeight    =   5010
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   9840
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   5010
   ScaleWidth      =   9840
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame2 
      Height          =   675
      Left            =   120
      TabIndex        =   2
      Top             =   90
      Width           =   9510
      Begin VB.ComboBox cboMonth 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         ItemData        =   "Frm_MPApprovalNew.frx":0000
         Left            =   960
         List            =   "Frm_MPApprovalNew.frx":0002
         Style           =   2  'Dropdown List
         TabIndex        =   3
         Top             =   225
         Width           =   2385
      End
      Begin VB.Label Label1 
         Caption         =   "Month : "
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   180
         TabIndex        =   6
         Top             =   270
         Width           =   870
      End
      Begin VB.Label Label2 
         Caption         =   "MP Number : "
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   6765
         TabIndex        =   5
         Top             =   255
         Width           =   1155
      End
      Begin VB.Label lblMPNumber 
         Caption         =   "9999.9999.9999"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   240
         Left            =   7965
         TabIndex        =   4
         Top             =   255
         Width           =   1455
      End
   End
   Begin VB.Frame Frame1 
      Height          =   4245
      Left            =   120
      TabIndex        =   0
      Top             =   675
      Width           =   9510
      Begin VB.TextBox txtPassword2 
         Height          =   345
         IMEMode         =   3  'DISABLE
         Left            =   5670
         PasswordChar    =   "X"
         TabIndex        =   15
         Text            =   "Text1"
         Top             =   975
         Width           =   3465
      End
      Begin VB.TextBox txtPassword 
         Height          =   345
         IMEMode         =   3  'DISABLE
         Left            =   6165
         TabIndex        =   14
         Text            =   "Text1"
         Top             =   3720
         Visible         =   0   'False
         Width           =   2370
      End
      Begin VB.CommandButton cmdApprove 
         Caption         =   "Approve"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   525
         Left            =   6735
         TabIndex        =   13
         Top             =   1515
         Width           =   1410
      End
      Begin VB.TextBox txtClientApprovedBY 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   7290
         TabIndex        =   9
         Top             =   2340
         Visible         =   0   'False
         Width           =   2115
      End
      Begin VB.TextBox txtClientNotedBy 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   7290
         TabIndex        =   8
         Top             =   3135
         Visible         =   0   'False
         Width           =   2115
      End
      Begin MSComctlLib.TreeView trvMedium 
         Height          =   3855
         Left            =   150
         TabIndex        =   1
         Top             =   225
         Width           =   5115
         _ExtentX        =   9022
         _ExtentY        =   6800
         _Version        =   393217
         LineStyle       =   1
         Style           =   7
         Checkboxes      =   -1  'True
         BorderStyle     =   1
         Appearance      =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin MSComCtl2.DTPicker DTClientApprovalDate 
         BeginProperty DataFormat 
            Type            =   0
            Format          =   "dddddd"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   2057
            SubFormatType   =   0
         EndProperty
         Height          =   315
         Left            =   7290
         TabIndex        =   7
         Top             =   2715
         Visible         =   0   'False
         Width           =   2115
         _ExtentX        =   3731
         _ExtentY        =   556
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   104398849
         CurrentDate     =   38272
      End
      Begin VB.Label Label6 
         Caption         =   "Enter your password below :"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   225
         Left            =   5685
         TabIndex        =   16
         Top             =   660
         Width           =   2490
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         Caption         =   "Client Approved By :"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   5355
         TabIndex        =   12
         Top             =   2340
         Visible         =   0   'False
         Width           =   1845
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         Caption         =   "Client Approval Date :"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   5355
         TabIndex        =   11
         Top             =   2745
         Visible         =   0   'False
         Width           =   1845
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         Caption         =   "Client Noted By :"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   5355
         TabIndex        =   10
         Top             =   3135
         Visible         =   0   'False
         Width           =   1845
      End
   End
End
Attribute VB_Name = "Frm_MPApprovalNew"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim strMPNumber As String
Dim strSql As String
Dim rsTemp As ADODB.Recordset
Dim strTaskID As String, strActivityID As String
Dim ForceCheck As Boolean
Dim ForceCheckID As Integer
'
Dim fso As Object
Dim FSOInterface As Object
Dim LogFileName As String
    
Private Sub CboMonth_Click()
    Call ViewMedium
End Sub

Private Sub cmdApprove_Click()
    Dim rsTemp As New ADODB.Recordset
    Dim islatest As Integer
    
    
    If isAuthenticated Then
                
        'Check Approval Information (Diabled becouse populated via Web)
        'If Trim(txtClientApprovedBY.Text) = "" Then
        '    MsgBox "a required field is empty! (Client Approved By)"
        '    txtClientApprovedBY.SetFocus
        '    Exit Sub
        'End If
        
        If cboMonth.Text = "---Choose Month---" Then
            MsgBox "Please Choose Month!", vbExclamation, strApplication_Name
            cboMonth.SetFocus
            Exit Sub
        End If
        
        'Check is latest
        rsTemp.Open "select is_latest from mp_master where mp_number = '" & FrmMPInsertion.cboMPNumber.Text & "'", ConnERP, 1, 3
        islatest = 1
        If Not rsTemp.EOF Then
            islatest = rsTemp(0)
        Else
            MsgBox "This Media Plan has been deleted!", vbExclamation, strApplication_Name
        End If
        rsTemp.Close
        
        If islatest <> 1 Then
            MsgBox "Approval Rejected by system! This is not the latest media plan!", vbExclamation, strApplication_Name
            Exit Sub
        End If
        
        'Check Medium & Do Approval
        Call CheckMedium
        
        txtPassword.Text = ""
        txtPassword2.Text = ""
        txtPassword2.SetFocus
    Else
        MsgBox "Invalid Password!", vbExclamation, strApplication_Name
        txtPassword.Text = ""
        txtPassword2.Text = ""
        txtPassword2.SetFocus
    End If
End Sub

Private Sub CheckMedium()
    
    Dim i As Integer
    Dim nApproval As Integer, nMedium As Integer
    Dim rsTemp As New ADODB.Recordset
    Dim strSql As String, strMediumCode As String, isAdaNull As Boolean
    Dim StrIBTemp As String
    Dim StrTask As String
    Dim StrActivity As String
    
    'Create Log File
    Create_Log_File
    
    nApproval = 0
    nMedium = 0
    For i = 1 To trvMedium.Nodes.Count
        If trvMedium.Nodes(i).Checked And Not trvMedium.Nodes(i).Bold Then
            If Mid(trvMedium.Nodes(i).KEY, 6, 4) = "MDUM" Then
                
                nMedium = nMedium + 1
                
                If isInsertionNotEmpty(trvMedium.Nodes(i).KEY, cboMonth.ListIndex) Then
                    
                    StrIBTemp = ""
                    StrActivity = ""
                    StrTask = ""
                     
                    StrActivity = trvMedium.Nodes(i).Parent.Text
                    StrTask = trvMedium.Nodes(i).Parent.Parent.Text
                            
                    'cek objective for tv
                    strSql = "select medium_code from mp_medium where mp_medium_id = '" & trvMedium.Nodes(i).KEY & "'"
                    rsTemp.Open strSql, ConnERP, 1, 3
                        strMediumCode = rsTemp(0)
                    rsTemp.Close
                                       
                    
                    If strMediumCode = "TV" Then 'TV Only (Need To check Link Between)
                        isAdaNull = False
                        'Check Apakah Ada TV Tarps yang belum related ke Objective
                        strSql = "select count(*) from mp_insertion where mp_plan_dim_id in (select mp_plan_dim_id from mp_ids where mp_medium_id = '" & trvMedium.Nodes(i).KEY & "') "
                        strSql = strSql & " and [month] = " & cboMonth.ListIndex & " and mp_tv_rf_id is null"
                        rsTemp.Open strSql, ConnERP, 1, 3
                            If rsTemp(0) <> 0 Then isAdaNull = True
                        rsTemp.Close
                        
                        If Not isAdaNull Then
                            
                            ' Create IB
                                                                                    
                            StrIBTemp = Generate_IB_TV(trvMedium.Nodes(i).KEY, cboMonth.ListIndex)
                           
                            If UCase(StrIBTemp) <> "ERROR" Then
                                nApproval = nApproval + 1
                                
                                Call DoApproval(trvMedium.Nodes(i).KEY, cboMonth.ListIndex)
                                'Write To Log
                                Write_To_Log_File StrTask & " : " & StrActivity & " : Medium: TV : Selected Month: " & cboMonth.Text & " --->  IB ID : " & StrIBTemp
                            Else
                                'Write To Log
                                Write_To_Log_File StrTask & " : " & StrActivity & " : Medium: TV : Selected Month: " & cboMonth.Text & " ---> " & " Generete IB Failed."
                            End If
                        
                        Else
                            'Write To Log There are, Unliked Tarps To Objective
                            Write_To_Log_File StrTask & " : " & StrActivity & " : Medium: TV : Selected Month: " & cboMonth.Text & " ---> Generete IB Failed. There are Unlinked TARPS to Objective."
                            
                        End If
                        
                    Else 'OTher, Cinema,Radio,Print
                        Select Case strMediumCode
                        Case "PR"   'For PR
                            StrIBTemp = Generate_IB_Print(trvMedium.Nodes(i).KEY, cboMonth.ListIndex)
                            If UCase(StrIBTemp) = "ERROR" Then
                               'Write To Log
                                Write_To_Log_File StrTask & " : " & StrActivity & " : Medium: Print : Selected Month: " & cboMonth.Text & " ---> Generete IB Failed. There are Unlinked TARPS to Objective."
                            ElseIf UCase(StrIBTemp) = "CANCEL" Then
                                'Write To Log
                                Write_To_Log_File StrTask & " : " & StrActivity & " : Medium: Print : Selected Month: " & cboMonth.Text & " ---> Approval And Create IB Canceled"
                            Else
                                nApproval = nApproval + 1
                                
                                 Call DoApproval(trvMedium.Nodes(i).KEY, cboMonth.ListIndex)
                                'Write To Log
                                Write_To_Log_File StrTask & " : " & StrActivity & " : Medium: Print : Selected Month: " & cboMonth.Text & " --->  IB ID : " & StrIBTemp
                            End If
                            
                        Case "RD"   'For RD
                        
                            StrIBTemp = Generate_IB_Radio(trvMedium.Nodes(i).KEY, cboMonth.ListIndex)
                            If UCase(StrIBTemp) <> "ERROR" Then
                                nApproval = nApproval + 1
                               
                                Call DoApproval(trvMedium.Nodes(i).KEY, cboMonth.ListIndex)
                                'Write To Log
                                Write_To_Log_File StrTask & " : " & StrActivity & " : Medium: Radio : Selected Month: " & cboMonth.Text & " --->  IB ID : " & StrIBTemp
                            Else
                                'Write To Log
                                Write_To_Log_File StrTask & " : " & StrActivity & " : Medium: Radio : Selected Month: " & cboMonth.Text & " ---> Generete IB Failed. There are Unlinked TARPS to Objective."
                            End If
                            
                        Case "OT" 'For OT
                            StrIBTemp = Generate_IB_Others(trvMedium.Nodes(i).KEY, cboMonth.ListIndex)
                            If UCase(StrIBTemp) <> "ERROR" Then
                                nApproval = nApproval + 1
                                
                                Call DoApproval(trvMedium.Nodes(i).KEY, cboMonth.ListIndex)
                                'Write To Log
                                Write_To_Log_File StrTask & " : " & StrActivity & " : Medium: Other : Selected Month: " & cboMonth.Text & " --->  IB ID : " & StrIBTemp
                            Else
                                'Write To Log
                                Write_To_Log_File StrTask & " : " & StrActivity & " : Medium: Other : Selected Month: " & cboMonth.Text & " ---> Generete IB Failed. There are Unlinked TARPS to Objective."
                            End If
                            
                        Case "CN" 'For Cinema
                            StrIBTemp = Generate_IB_Others(trvMedium.Nodes(i).KEY, cboMonth.ListIndex)
                            If UCase(StrIBTemp) <> "ERROR" Then
                                nApproval = nApproval + 1
                                
                                Call DoApproval(trvMedium.Nodes(i).KEY, cboMonth.ListIndex)
                                'Write To Log
                                Write_To_Log_File StrTask & " : " & StrActivity & " : Medium: Cinema : Selected Month: " & cboMonth.Text & " --->  IB ID : " & StrIBTemp
                            Else
                                'Write To Log
                                Write_To_Log_File StrTask & " : " & StrActivity & " : Medium: Cinema : Selected Month: " & cboMonth.Text & " ---> Generete IB Failed. There are Unlinked TARPS to Objective."
                            End If
                            
                        End Select
                    End If
                 End If
            End If
        End If
    Next
    
    MsgBox nApproval & " Medium(s) approved!", vbExclamation, strApplication_Name
                
    '---------- Show Log File & Close Log file object--------------
    Set fso = Nothing
    Set FSOInterface = Nothing
        
    Shell "Notepad.exe C:\MPLog\" & LogFileName, vbNormalNoFocus
        
    LogFileName = ""
    '--------------------------------------------------------------
    
    'Refresh Form
    Call CboMonth_Click
    
    
    'Yang belum
    '===================================================
    'MP :
        'TV Area
        'Bonus for Print
        'Pajak Tontonan di Cinema
        
    'Link to IB :
        'Output di Modul Generete IB jika error
        'Submit to BC Plan
        'Create Quotation for BU1
    
    'Update Actual :
        'Update Actual for in MP from BC
     '==================================================
End Sub

Private Sub DoApproval(strMediumID As String, intMonth As Integer)
    Dim strSql As String
    
    'Disabled because polpulated via Web
    strSql = "update mp_monthly_activity set approval=1,"
    'StrSQL = StrSQL & "Client_Approved_By='" & txtClientApprovedBY.Text & "',"
    'StrSQL = StrSQL & "Client_Approved_Date='" & CDate(DTClientApprovalDate.Value) & "',"
    'StrSQL = StrSQL & "Client_Noted_By='" & txtClientNotedBy.Text & "',"
    strSql = strSql & "Approved_By='" & strLogin_FullName & "',"
    strSql = strSql & "Approved_Date=getdate(),"
    strSql = strSql & "Approved_mp_medium_id='" & strMediumID & "'"
    strSql = strSql & " WHERE mp_medium_id = '" & strMediumID & "' and month_number=" & intMonth
    ConnERP.Execute strSql
    
    '                           Add/Update MP_Monthly_Quotation
    '============================================================================
    Dim strMPNumber As String
    Dim rsTemp As New ADODB.Recordset
    Dim strMediumCode As String
    
    Dim StrOriginalBrandCode As String
    
    
    'Get MP Number
    strMPNumber = lblMPNumber.Caption
    
    'Get Medium Code
    rsTemp.Open "select medium_code from mp_medium where mp_medium_id='" & strMediumID & "'", ConnERP, 1, 3
        strMediumCode = rsTemp(0)
    rsTemp.Close
    
    'Get StrOriginalBrandCode
    strSql = "SELECT Original_Brand_Code,Original_Brand_Name FROM MP_Activity "
    strSql = strSql & " WHERE MP_Activity_ID IN(SELECT MP_Activity_ID FROM MP_Medium WHERE mp_medium_id='" & strMediumID & "')"
    
    rsTemp.Open strSql, ConnERP, 1, 3
        StrOriginalBrandCode = rsTemp(0).Value
    rsTemp.Close
    
    'cek apakah di quotation sudah ada?? & Get Monthtly_Quotation_ID yang sama Brandocde nya
    Dim isExist As Boolean
    Dim strmp_Monthly_Quot_ID As String
    
    isExist = False
    strmp_Monthly_Quot_ID = ""
    
    strSql = "SELECT mp_monthly_quot_id FROM mp_monthly_quotation "
    strSql = strSql & " WHERE mp_number='" & strMPNumber & "' AND medium_code='" & strMediumCode & "' AND [month]=" & intMonth
    strSql = strSql & " AND LEFT(mp_monthly_quot_id,4)='" & StrOriginalBrandCode & "'"
    
    rsTemp.Open strSql, ConnERP, 1, 3
    If Not rsTemp.EOF Then
        isExist = True
        strmp_Monthly_Quot_ID = rsTemp(0)
    End If
    rsTemp.Close
    
    'Get Total Value Monthly Quotation from mp_monthly_activity
    
    Dim int_total_gross As Double
    Dim int_total_nett As Double
    Dim int_msc As Double
    Dim int_other_cost As Double
    Dim int_bonus_fee As Double
    Dim int_agency_charge As Double
    
    Dim int_sub_total As Double
    Dim int_grand_total As Double
    
    int_total_gross = 0
    int_total_nett = 0
    int_msc = 0
    int_other_cost = 0
    int_bonus_fee = 0
    int_agency_charge = 0
    
    strSql = "select sum(min_budget),sum(gross_budget),sum(msc_paid_value),sum(msc_bonus_value) from mp_monthly_activity "
    strSql = strSql & "where month_number = " & intMonth & " and mp_medium_id = '" & strMediumID & "' "
    
    Dim rsQuotation As New ADODB.Recordset
    rsQuotation.Open strSql, ConnERP, 1, 3
    If Not rsQuotation.EOF Then
        int_total_gross = rsQuotation(1)
        int_total_nett = rsQuotation(0)
        int_msc = rsQuotation(2)
        'int_other_cost = 0
        int_bonus_fee = rsQuotation(3)
        'int_agency_charge = 0
    End If
    rsQuotation.Close
    Set rsQuotation = Nothing
    
    int_sub_total = int_total_nett + int_msc + int_other_cost + int_bonus_fee
    int_grand_total = int_sub_total + int_agency_charge
    
    '==========================================================================================
    If isExist Then
        'update
        strSql = "UPDATE mp_monthly_quotation SET "
        strSql = strSql & "total_gross=total_gross+" & int_total_gross
        strSql = strSql & ", total_nett=total_nett+" & int_total_nett
        strSql = strSql & ", MSC=MSC+" & int_msc
        strSql = strSql & ", other_cost=other_cost+" & int_other_cost
        strSql = strSql & ", bonus_fee=bonus_fee+" & int_bonus_fee
        strSql = strSql & ", sub_total=sub_total+" & int_sub_total
        strSql = strSql & ", agency_charge=agency_charge+" & int_agency_charge
        strSql = strSql & ", grand_total = grand_total+" & int_grand_total
        strSql = strSql & " WHERE mp_monthly_quot_id='" & strmp_Monthly_Quot_ID & "'"
        
        ConnERP.Execute strSql
    Else
        'generate id baru
        
        Dim Int_Year As Integer
        Dim rsRunningNumber As New ADODB.Recordset
        Dim str_running_number As String
        Dim str_param_job_id As String, str_param_job_number_agency As String
        Dim str_job_id As String, str_job_number_agency As String
        Dim strTemp As String
        
        Int_Year = CInt(Mid(strMPNumber, 6, 4))
        
        strTemp = StrOriginalBrandCode & "." & Int_Year & "."
        
        'rsRunningNumber.Open "SELECT ISNULL(MAX(CAST(SUBSTRING(mp_monthly_quot_id,11,4) AS INT)),0)+1 FROM mp_monthly_quotation WHERE mp_monthly_quot_id LIKE '" & Mid(strMPNumber, 1, 10) & "%'", connerp, 1, 3
        rsRunningNumber.Open "SELECT ISNULL(MAX(CAST(SUBSTRING(mp_monthly_quot_id,11,4) AS INT)),0)+1 FROM mp_monthly_quotation WHERE mp_monthly_quot_id LIKE '" & strTemp & "%'", ConnERP, 1, 3
        
        str_running_number = Right("0000" & CStr(rsRunningNumber(0)), 4)
        rsRunningNumber.Close
        
        'strmp_Monthly_Quot_ID = Mid(strMPNumber, 1, 4) & "." & CStr(Int_Year) & "." & str_running_number
        strmp_Monthly_Quot_ID = StrOriginalBrandCode & "." & CStr(Int_Year) & "." & str_running_number
        
        'Generate job_id dan job_number_agency
        Select Case strMediumCode
            Case "TV":
                str_param_job_id = "Television Media Induk"
                str_param_job_number_agency = "Television Club Agency"
            Case "RD":
                str_param_job_id = "Radio Media Induk"
                str_param_job_number_agency = "Radio Club Agency"
            Case "PR":
                str_param_job_id = "Print Media Induk"
                str_param_job_number_agency = "Print Club Agency"
            Case "CN":
                str_param_job_id = "Other Media Induk Cinema"
                str_param_job_number_agency = "Other Club Agency"
            Case "OT":
                str_param_job_id = "Other Media Induk Outdoor"
                str_param_job_number_agency = "Other Club Agency"
        End Select
        
        rsTemp.Open "select media_type_code from media_type where media_type_name = '" & str_param_job_id & "'", ConnERP, 1, 3
            str_job_id = StrOriginalBrandCode & "." & rsTemp(0) & "." & Right(CStr(Int_Year), 2) & Right("0" & CStr(intMonth), 2)
        rsTemp.Close
        
        rsTemp.Open "select media_type_code from media_type where media_type_name = '" & str_param_job_number_agency & "'", ConnERP, 1, 3
            str_job_number_agency = StrOriginalBrandCode & "." & rsTemp(0) & "." & Right(CStr(Int_Year), 2) & Right("0" & CStr(intMonth), 2)
        rsTemp.Close
        
        'insert
        strSql = "INSERT INTO mp_monthly_quotation"
        strSql = strSql & "(mp_monthly_quot_id,[year],mp_number,medium_code,[month],total_gross,total_nett,msc,other_cost,bonus_fee,sub_total,agency_charge,grand_total,"
        strSql = strSql & "job_id,job_number_agency,"
        strSql = strSql & "is_latest,status) VALUES "
        strSql = strSql & "('" & strmp_Monthly_Quot_ID & "'," & Int_Year & ",'" & strMPNumber & "','" & strMediumCode & "'," & intMonth & "," & int_total_gross & "," & int_total_nett & "," & int_msc & "," & int_other_cost & "," & int_bonus_fee & "," & int_sub_total & "," & int_agency_charge & "," & int_grand_total & ","
        strSql = strSql & "'" & str_job_id & "','" & str_job_number_agency & "',"
        strSql = strSql & "1,1)"
        ConnERP.Execute strSql
        
    End If
    '======================== Enf of Update Quotation ==========================
    
End Sub

Private Function isAuthenticated(Optional pvUserName As String) As Boolean
    Dim rsPassword As New ADODB.Recordset
    isAuthenticated = False
    If pvUserName = "" Then pvUserName = strLogin_FullName
    rsPassword.Open "select password from user_id where user_name = '" & Clear_String(pvUserName) & "'", ConnERP, 1, 3
    
    If txtPassword.Text = DecryptPassword(rsPassword(0)) Then
        isAuthenticated = True
    End If
        
    rsPassword.Close
    Set rsPassword = Nothing
End Function

Private Sub Form_Load()
    strMPNumber = FrmMPInsertion.cboMPNumber.Text
    lblMPNumber.Caption = strMPNumber
    rs_Date.Requery
    DTClientApprovalDate.Value = CDate(rs_Date(0))
    txtPassword.Text = ""
    txtPassword2.Text = ""
    ForceCheck = False
    Call LoadMonth
    'Call ViewMedium
End Sub

Private Sub LoadMonth()
    Dim i As Integer
    cboMonth.Clear
    cboMonth.AddItem "---Choose Month---"
    For i = 1 To 12
        cboMonth.AddItem EngMonthName(i)
    Next
    cboMonth.ListIndex = 0
End Sub

Private Sub ViewMedium()
    Me.MousePointer = vbHourglass
    Dim i As Integer
    strSql = " select a.mp_task_id,a.task_desc,b.mp_activity_id,b.activity_type,b.activity_desc,"
    strSql = strSql & "     c.mp_medium_id,c.medium_name, d.approval,d.budget from mp_task a"
    strSql = strSql & " inner join mp_activity b on a.mp_task_id = b.mp_task_id and a.mp_number = '" & strMPNumber & "' "
    strSql = strSql & " inner join mp_medium c on b.mp_activity_id = c.mp_activity_id"
    strSql = strSql & " inner join mp_monthly_activity d on c.mp_medium_id = d.mp_medium_id and d.month_number = " & cboMonth.ListIndex
    strSql = strSql & " order by a.mp_task_id,b.mp_activity_id,c.mp_medium_id"
    Set rsTemp = New ADODB.Recordset
    
    rsTemp.Open strSql, ConnERP, 1, 3
    
    strTaskID = ""
    strActivityID = ""
    
    
    trvMedium.Nodes.Clear
    i = 1
    While Not rsTemp.EOF

        If strTaskID <> rsTemp("mp_task_id") Then
            trvMedium.Nodes.Add , , rsTemp("mp_task_id"), "TASK - " & rsTemp("task_desc")
            trvMedium.Nodes(i).Expanded = True
            i = i + 1
            strTaskID = rsTemp("mp_task_id")
        End If
        
        If strActivityID <> rsTemp("mp_activity_id") Then
            trvMedium.Nodes.Add strTaskID, tvwChild, rsTemp("mp_activity_id"), "ACTIVITY - " & rsTemp("activity_type") & " - " & rsTemp("activity_desc")
            trvMedium.Nodes(i).Expanded = True
            i = i + 1
            strActivityID = rsTemp("mp_activity_id")
        End If
        
        trvMedium.Nodes.Add strActivityID, tvwChild, rsTemp("mp_medium_id"), rsTemp("medium_name") '& " - Rp " & FormatNumber(rsTemp("budget"), 2)
        If rsTemp("approval") = 1 Then
            trvMedium.Nodes(i).Bold = True
            trvMedium.Nodes(i).ForeColor = vbRed
            trvMedium.Nodes(i).Checked = True
            
        End If
        
        i = i + 1
        rsTemp.MoveNext
    Wend
    
    rsTemp.Close
    Me.MousePointer = vbNormal
End Sub

Private Sub trvMedium_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    'related to event nodecheck (link code #911)
    If ForceCheck Then
        trvMedium.Nodes(ForceCheckID).Checked = True
        ForceCheck = False
    End If
End Sub

Private Sub trvMedium_NodeCheck(ByVal Node As MSComctlLib.Node)
    Select Case Mid(Node.KEY, 5, 6)
        Case ".TASK.":
            Call Click_ON_Task(Node.KEY, Node.Checked)
        Case ".ACTV.":
            Call Click_ON_Activity(Node.KEY, Node.Checked)
        Case ".MDUM.":
            If Node.Bold Then 'yg udah approved ga boleh diuncheck
                'Node.Checked = True 'it's not working!! :(
                
                'try this
                ForceCheck = True
                ForceCheckID = Node.Index
                'abis itu.. force-check-nya di event mouse_up (link code #911)
            End If
    End Select
End Sub

Private Sub Click_ON_Activity(strActivityID As String, isChecked As Boolean)
    Dim i As Integer
    For i = 1 To trvMedium.Nodes.Count
        If Mid(trvMedium.Nodes(i).KEY, 5, 6) <> ".TASK." Then 'bukan root
            If trvMedium.Nodes(i).Parent.KEY = strActivityID Then
                If Not trvMedium.Nodes(i).Bold Then trvMedium.Nodes(i).Checked = isChecked
            End If
        End If
    Next
End Sub

Private Sub Click_ON_Task(strTaskID As String, isChecked As Boolean)
    Dim i As Integer
    For i = 1 To trvMedium.Nodes.Count
        If Mid(trvMedium.Nodes(i).KEY, 5, 6) <> ".TASK." Then 'bukan root
            If trvMedium.Nodes(i).Parent.KEY = strTaskID Then
                trvMedium.Nodes(i).Checked = isChecked
                Call Click_ON_Activity(trvMedium.Nodes(i).KEY, isChecked)
            End If
        End If
    Next
End Sub

Private Sub txtPassword2_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
        Case 13:
            Call cmdApprove_Click
        Case 8:
        If txtPassword2.Text <> "" Then
            txtPassword2.Text = Mid(txtPassword2.Text, 1, Len(txtPassword2.Text) - 3)
            txtPassword.Text = Mid(txtPassword.Text, 1, Len(txtPassword.Text) - 1)
        End If
    Case Else
        txtPassword.Text = txtPassword.Text & Chr(KeyAscii)
        txtPassword2.Text = txtPassword2.Text & String(3, Chr(KeyAscii))
    End Select
    txtPassword2.SelStart = Len(txtPassword2.Text)
End Sub

Function isInsertionNotEmpty(strMediumID As String, intMonth As Integer) As Boolean
    Dim rsTemp As New ADODB.Recordset
    Dim strSql As String, strMediumCode As String
    isInsertionNotEmpty = False
    strMediumCode = ""
    strSql = "select medium_code from mp_medium where mp_medium_id = '" & strMediumID & "'"
    rsTemp.Open strSql, ConnERP, 1, 3
    If Not rsTemp.EOF Then strMediumCode = rsTemp(0)
    rsTemp.Close
    If strMediumCode <> "" Then
        Select Case strMediumCode
            Case "OT"
                strSql = "select count(*) from MP_Other_Monthly_Budget where month_number = " & intMonth
            Case Else
                strSql = "select count(*) from MP_Insertion where [month] = " & intMonth
        End Select
        strSql = strSql & " and mp_plan_dim_id in (select mp_plan_dim_id from mp_ids where mp_medium_id = '" & strMediumID & "')"
        rsTemp.Open strSql, ConnERP, 1, 3
        If rsTemp(0) <> 0 Then isInsertionNotEmpty = True
        rsTemp.Close
    End If
End Function


Private Sub Create_Log_File()
'*************************************************************
'Nama Procedure     : Create_Log_File
'Fungsi             : Membuat file text untuk LOG
'Programer          :
'Tgl Pembuatan      : 14 Juni 2004
'Last Update/By     :
'*************************************************************
        
    On Error GoTo errHand
    
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set FSOInterface = fso.CreateFolder("c:\MPLog")
    
    rs_Date.Requery
    LogFileName = "MPAppLog_" & Format(rs_Date(0), "MMDDYY") & "_" & Format(rs_Date(0), "hhmmss") & ".txt"
    Set FSOInterface = fso.CreateTextFile("c:\MPLog\" & LogFileName)
    
    Exit Sub
    
errHand:
    If Err.Number = 58 Then
        Resume Next
    End If
End Sub
Private Sub Write_To_Log_File(StrTeks As String)
'*************************************************************
'Nama Procedure     : Write_To_Log_File
'Fungsi             : Menulis ke LOG
'Programer          :
'Tgl Pembuatan      : 14 Juni 2004
'Last Update/By     :
'*************************************************************
    rs_Date.Requery
    FSOInterface.WriteLine StrTeks
End Sub

