VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Begin VB.Form frm_MPActivityAdd 
   Appearance      =   0  'Flat
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Add Activity"
   ClientHeight    =   3915
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5835
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3915
   ScaleWidth      =   5835
   StartUpPosition =   1  'CenterOwner
   Begin Threed.SSPanel pnl_Main 
      Align           =   1  'Align Top
      Height          =   3150
      Left            =   0
      TabIndex        =   3
      Top             =   750
      Width           =   5835
      _Version        =   65536
      _ExtentX        =   10292
      _ExtentY        =   5556
      _StockProps     =   15
      Caption         =   "SSPanel1"
      BackColor       =   15790320
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BevelOuter      =   1
      Begin VB.Frame Frame1 
         Height          =   3045
         Left            =   105
         TabIndex        =   4
         Top             =   15
         Width           =   5640
         Begin VB.ComboBox cboActivity 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   2055
            Style           =   2  'Dropdown List
            TabIndex        =   16
            Top             =   285
            Width           =   3390
         End
         Begin VB.ComboBox cboBrandVariant 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   2055
            Style           =   2  'Dropdown List
            TabIndex        =   15
            Top             =   1035
            Width           =   2295
         End
         Begin VB.ComboBox cboTargetAudience 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   2055
            Style           =   2  'Dropdown List
            TabIndex        =   14
            Top             =   1395
            Width           =   2295
         End
         Begin VB.TextBox txtBrandTarget 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
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
            Left            =   2055
            TabIndex        =   13
            Text            =   "txtBrandTarget"
            Top             =   1770
            Width           =   3390
         End
         Begin VB.ComboBox cboBrandVariantCode 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   4410
            Style           =   2  'Dropdown List
            TabIndex        =   12
            Top             =   1050
            Visible         =   0   'False
            Width           =   1050
         End
         Begin VB.CheckBox cbkCinema 
            Caption         =   "Cinema"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Left            =   2040
            TabIndex        =   11
            Top             =   2475
            Width           =   960
         End
         Begin VB.CheckBox cbkOther 
            Caption         =   "Other"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   3105
            TabIndex        =   10
            Top             =   2475
            Width           =   855
         End
         Begin VB.CheckBox cbkPrint 
            Caption         =   "Print"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            Left            =   4110
            TabIndex        =   9
            Top             =   2070
            Width           =   735
         End
         Begin VB.CheckBox cbkRadio 
            Caption         =   "Radio"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            Left            =   3105
            TabIndex        =   8
            Top             =   2070
            Width           =   855
         End
         Begin VB.CheckBox cbkTV 
            Caption         =   "TV"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            Left            =   2040
            TabIndex        =   7
            Top             =   2070
            Width           =   615
         End
         Begin VB.ComboBox cboTargetAudienceCode 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   4410
            Style           =   2  'Dropdown List
            TabIndex        =   6
            Top             =   1380
            Visible         =   0   'False
            Width           =   1050
         End
         Begin VB.TextBox txtDescription 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
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
            Left            =   2055
            TabIndex        =   5
            Text            =   "txtDescription"
            Top             =   690
            Width           =   3390
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Type "
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000C0&
            Height          =   195
            Left            =   240
            TabIndex        =   22
            Top             =   300
            Width           =   405
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Brand Variant "
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000C0&
            Height          =   195
            Left            =   255
            TabIndex        =   21
            Top             =   1050
            Width           =   1020
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Target Audience "
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000C0&
            Height          =   195
            Left            =   255
            TabIndex        =   20
            Top             =   1410
            Width           =   1230
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Brand Target "
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000C0&
            Height          =   195
            Left            =   255
            TabIndex        =   19
            Top             =   1785
            Width           =   990
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Medium "
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000C0&
            Height          =   195
            Left            =   255
            TabIndex        =   18
            Top             =   2145
            Width           =   585
         End
         Begin VB.Label Label6 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Description "
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000C0&
            Height          =   195
            Left            =   255
            TabIndex        =   17
            Top             =   705
            Width           =   840
         End
      End
   End
   Begin VB.PictureBox picToolbar 
      Align           =   1  'Align Top
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ClipControls    =   0   'False
      ForeColor       =   &H80000008&
      Height          =   750
      Left            =   0
      ScaleHeight     =   750
      ScaleWidth      =   5835
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   0
      Width           =   5835
      Begin VB.PictureBox picButton 
         Appearance      =   0  'Flat
         BackColor       =   &H00F0F0F0&
         BorderStyle     =   0  'None
         ClipControls    =   0   'False
         ForeColor       =   &H80000008&
         Height          =   750
         Index           =   10
         Left            =   90
         Picture         =   "frm_MPActivityAdd.frx":0000
         ScaleHeight     =   750
         ScaleWidth      =   1500
         TabIndex        =   2
         TabStop         =   0   'False
         Top             =   0
         Width           =   1500
      End
      Begin VB.PictureBox picButton 
         Appearance      =   0  'Flat
         BackColor       =   &H00F0F0F0&
         BorderStyle     =   0  'None
         ClipControls    =   0   'False
         ForeColor       =   &H80000008&
         Height          =   750
         Index           =   23
         Left            =   1620
         Picture         =   "frm_MPActivityAdd.frx":1C02
         ScaleHeight     =   750
         ScaleWidth      =   1500
         TabIndex        =   1
         TabStop         =   0   'False
         Top             =   0
         Width           =   1500
      End
   End
End
Attribute VB_Name = "frm_MPActivityAdd"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
''*****************************************************************************
' Nama Submodul         :  FrmMPActivityAdd
' Fungsi Submodul       :  Untuk Add Activity
' Tabel yg digunakan    :  MP_Master, MP_Task, MP_Activity,Mp_Medium
' Prosedur/Function     :
' Nama Programmer       :  Sistyo
' Tgl Pembuatan         :  10 Agustus 2004
' Last Update           :  11 Agustus 2004/Sistyo
'******************************************************************************
Option Explicit
Dim rsTemp As New ADODB.Recordset

Private Sub cboTargetAudience_Click()
    cboTargetAudienceCode.ListIndex = cboTargetAudience.ListIndex
End Sub

Private Sub Form_Load()
    Call loadcatalog
End Sub

Private Sub db_New()
'<CSCM>
'********************************************************************************
'Procedure Name     : db_New
'Procedure Function : Menambah Activity
'Input Parameter    : ---
'Output Parameter   : ---
'Date               : 3/29/2016
'LastUpdate/By      : Tedi / Kreatif
'Name Before        : cmdAdd_Click
'********************************************************************************
'</CSCM>
    If cboActivity.Text = "" Then
        MsgBox "Please choose Activity Type!", vbExclamation, strApplication_Name
        cboActivity.SetFocus
        Exit Sub
    End If
    If Trim(txtDescription.Text) = "" Then
        MsgBox "Please type Activity Description!", vbExclamation, strApplication_Name
        txtDescription.SetFocus
        Exit Sub
    End If
    If cboBrandVariant.Text = "" Then
        MsgBox "Please choose Brand Variant!", vbExclamation, strApplication_Name
        cboBrandVariant.SetFocus
        Exit Sub
    End If
    If cboTargetAudience.Text = "" Then
        MsgBox "Please choose Target Audience!", vbExclamation, strApplication_Name
        cboTargetAudience.SetFocus
        Exit Sub
    End If
    If Trim(txtBrandTarget.Text) = "" Then
        MsgBox "Please type Brand Target!", vbExclamation, strApplication_Name
        txtBrandTarget.SetFocus
        Exit Sub
    End If
    Call AddActivity
    Unload Me
End Sub

Private Sub db_cancel()
'<CSCM>
'********************************************************************************
'Procedure Name     : cmdCancel_Click
'Procedure Function : ---
'Input Parameter    : ---
'Output Parameter   : ---
'Date               : 3/29/2016
'LastUpdate/By      : Tedi / Kreatif
'Name Before        : cmdCancel_Click
'********************************************************************************
'</CSCM>
    Unload Me
End Sub

Private Sub loadcatalog()
'*****************************************************************************
' Nama Prosedur     :   LoadCatalog
' Fungsi Prosedur   :   loading Catalog
' Parameter  Input  :
' Parameter Output  :
' Tgl Pembuatan     :   09 Agustus 2004
' Last Update/By    :   09 Agustus 2004/Sistyo
'*****************************************************************************

    Dim strSql As String
    Dim idx As Integer 'list index of combo
    
    strSql = "select isnull(brand_variant_code,''),isnull(brand_variant_name,'') from brand_variant "
    strSql = strSql & "where brand_code ='" & Mid(frm_MPEdit.cboMPNum, 1, 4) & "' order by brand_variant_name"
    
    rsTemp.Open strSql, ConnERP, 1, 3
    
    idx = 0
    
    While Not rsTemp.EOF
        
        cboBrandVariantCode.AddItem rsTemp(0), idx
        cboBrandVariant.AddItem rsTemp(1), idx
        idx = idx + 1
        rsTemp.MoveNext
        
    Wend
    
    rsTemp.Close
    
    rsTemp.Open "select isnull(campaign_type_name,'') from campaign_type_catalog", ConnERP, 1, 3
    
    While Not rsTemp.EOF
        
        cboActivity.AddItem rsTemp(0)
        rsTemp.MoveNext
    
    Wend
    
    rsTemp.Close
    
    txtDescription.Text = ""
    
    rsTemp.Open "select cluster,code from cluster", ConnERP, 1, 3
    idx = 0
    While Not rsTemp.EOF
        cboTargetAudience.AddItem rsTemp(0), idx
        cboTargetAudienceCode.AddItem rsTemp(1), idx
        rsTemp.MoveNext
        idx = idx + 1
    Wend
    
    rsTemp.Close
    txtBrandTarget.Text = ""
    
End Sub

Private Sub cboBrandVariant_click()

    cboBrandVariantCode.ListIndex = cboBrandVariant.ListIndex
    
End Sub

Private Sub AddActivity()
'*****************************************************************************
' Nama Prosedur     :   AddActivity
' Fungsi Prosedur   :   Menambah Activity
' Parameter  Input  :
' Parameter Output  :
' Tgl Pembuatan     :   09 Agustus 2004
' Last Update/By    :   09 Agustus 2004/Sistyo
'*****************************************************************************

    Dim strSql As String, strActivity_ID As String, strTask_Id As String, strMedium As String
    Dim strMediumID As String
    Dim pesan
    Dim original_brand_code As String, original_brand_name As String
    'generate Activity ID
        strTask_Id = frm_MPEdit.tdg_Task.Columns(1) ',   .TextMatrix(frm_MPEdit.FGTask.Row, 1)
        
        strSql = "select isnull(max(cast(substring(mp_activity_id,16,4) as int)),0)+1 from mp_activity where mp_task_id like '" & Mid(strTask_Id, 1, 15) & "%'"
    
        rsTemp.Open strSql, ConnERP, 1, 3
            strActivity_ID = Mid(strTask_Id, 1, 4) & ".ACTV." & Mid(strTask_Id, 11, 4) & "." & Right("0000" & CStr(rsTemp(0)), 4) 'Create new Activity_Id
        rsTemp.Close
    'Get original brand
    rsTemp.Open "select original_brand_code,original_brand_name from brand_variant where brand_variant_code = '" & cboBrandVariantCode.Text & "'", ConnERP, 1, 3
    If Not rsTemp.EOF Then
        original_brand_code = rsTemp("original_brand_code")
        original_brand_name = rsTemp("original_brand_name")
    End If
    rsTemp.Close
    
    'insert mp_activity
        strSql = "insert into mp_activity(mp_activity_id,mp_task_id,activity_type,activity_desc,brand_variant_code,"
        strSql = strSql & "brand_variant_name,target_audience,brand_target,target_audience_code,original_brand_code,original_brand_name) values "
        strSql = strSql & "('" & strActivity_ID & "',"
        strSql = strSql & "'" & strTask_Id & "',"
        strSql = strSql & "'" & Clear_String(cboActivity.Text) & "','" & Clear_String(txtDescription.Text) & "','" & cboBrandVariantCode.Text & "',"
        strSql = strSql & "'" & Clear_String(cboBrandVariant.Text) & "','" & Clear_String(cboTargetAudience.Text) & "',"
        strSql = strSql & "'" & Clear_String(txtBrandTarget.Text) & "'," & cboTargetAudienceCode.Text & ",'" & original_brand_code & "','" & Clear_String(original_brand_name) & "')"
        
        ConnERP.Execute strSql
    
    'insert mp_medium
    
    On Error GoTo err_Insert_Medium
        strMedium = ""
        If cbkTV.Value = 1 Then
            strMediumID = NextMPMediumID(strActivity_ID)
            ConnERP.Execute "insert into mp_medium values('" & strMediumID & "','" & strActivity_ID & "','TV','TV')"
            strMedium = strMedium & ", TV"
            Call InsertMonthlyActivity(strMediumID)
        End If
        
        If cbkRadio.Value = 1 Then
            strMediumID = NextMPMediumID(strActivity_ID)
            ConnERP.Execute "insert into mp_medium values('" & strMediumID & "','" & strActivity_ID & "','RD','Radio')"
            strMedium = strMedium & ", Radio"
            Call InsertMonthlyActivity(strMediumID)
        End If
        
        If cbkPrint.Value = 1 Then
            strMediumID = NextMPMediumID(strActivity_ID)
            ConnERP.Execute "insert into mp_medium values('" & strMediumID & "','" & strActivity_ID & "','PR','Print')"
            strMedium = strMedium & ", Print"
            Call InsertMonthlyActivity(strMediumID)
        End If
        
        If cbkCinema.Value = 1 Then
            strMediumID = NextMPMediumID(strActivity_ID)
            ConnERP.Execute "insert into mp_medium values('" & strMediumID & "','" & strActivity_ID & "','CN','Cinema')"
            strMedium = strMedium & ", Cinema"
            Call InsertMonthlyActivity(strMediumID)
        End If
        
        If cbkOther.Value = 1 Then
            strMediumID = NextMPMediumID(strActivity_ID)
            ConnERP.Execute "insert into mp_medium values('" & strMediumID & "','" & strActivity_ID & "','OT','Other')"
            strMedium = strMedium & ", Other"
            Call InsertMonthlyActivity(strMediumID)
        End If
    'end of insert mp_medium
err_Insert_Medium:
    If Err.Number <> 0 Then
        MsgBox Err.Description, vbExclamation, strApplication_Name
        'MsgBox "Can not insert medium because Brand Fee Catalog has not been setup!" & vbCrLf & "Activity will be saved but medium will be ignored!"
        If rsTemp.State = adStateOpen Then
            rsTemp.Close
        End If
    End If
    'update mp_master (update date, update by)
        ConnERP.Execute "update mp_master set last_update_by='" & strLogin_User & "',last_update_date=getdate() where mp_number='" & frm_MPEdit.cboMPNum.Text & "'"
        frm_MPEdit.viewActivity
    'tampilkan activity di grid
'        With frm_MPEdit.FGActivity
'            If .TextMatrix(1, 1) <> "" Then
'                .Rows = .Rows + 1
'            End If
'            .TextMatrix(.Rows - 1, 0) = .Rows - 1
'            .TextMatrix(.Rows - 1, 1) = strActivity_ID
'            .TextMatrix(.Rows - 1, 2) = strTask_Id
'            .TextMatrix(.Rows - 1, 3) = cboActivity.Text
'            .TextMatrix(.Rows - 1, 4) = txtDescription.Text
'            .TextMatrix(.Rows - 1, 5) = cboBrandVariant.Text
'            .TextMatrix(.Rows - 1, 6) = cboTargetAudience.Text
'            .TextMatrix(.Rows - 1, 7) = txtBrandTarget.Text
'            If Len(strMedium) <> 0 Then
'                strMedium = Right(strMedium, Len(strMedium) - 2)
'            End If
'        .TextMatrix(.Rows - 1, .cols - 1) = strMedium
'        End With
    
    'memberitahu user
    pesan = MsgBox("New Activity Created!", vbExclamation, strApplication_Name)
    
End Sub

Private Function NextMPMediumID(strMPACtivityID As String) As String
    Dim rsTemp As New ADODB.Recordset, strSql As String
    strSql = "select isnull(max(cast(substring(mp_medium_id,16,4) as int)),0)+1 from mp_medium where mp_activity_id like '" & Mid(strMPACtivityID, 1, 15) & "%'"
    rsTemp.Open strSql, ConnERP, 1, 3
        NextMPMediumID = Mid(strMPACtivityID, 1, 4) & ".MDUM." & Mid(strMPACtivityID, 11, 4) & "." & Right("0000" & CStr(rsTemp(0)), 4) 'Create new medium_Id
    rsTemp.Close
End Function

Private Sub InsertMonthlyActivity(strMediumID As String)
    'Insert Monthly Activity
    Dim i As Integer, strSql As String
    Dim MSC_Paid As Double
    Dim MSC_Paid_On_Flag As Double
    Dim MSC_Bonus As Double
    Dim MSC_Bonus_On_Flag As Double
    Dim Club_Agency As Double
    Dim Club_Agency_On_Flag As Double
    Dim rsTemp As New ADODB.Recordset
    For i = 1 To 12
        rsTemp.Open "select isnull(msc_paid,0),isnull(MSC_Paid_On_Flag,0),isnull(MSC_Bonus,0),isnull(MSC_Bonus_On_Flag,0),isnull(Club_Agency,0),isnull(Club_Agency_On_Flag,0) from brand_fee where brand_code='" & Mid(strMediumID, 1, 4) & "' and [year]=" & CInt(Mid(strMediumID, 11, 4)) & " and [month]=" & i, ConnERP, 1, 3
        If Not rsTemp.EOF Then
            MSC_Paid = rsTemp(0)
            MSC_Paid_On_Flag = rsTemp(1)
            MSC_Bonus = rsTemp(2)
            MSC_Bonus_On_Flag = rsTemp(3)
            Club_Agency = rsTemp(4)
            Club_Agency_On_Flag = rsTemp(5)
        Else
            MSC_Paid = 0
            MSC_Paid_On_Flag = 0
            MSC_Bonus = 0
            MSC_Bonus_On_Flag = 0
            Club_Agency = 0
            Club_Agency_On_Flag = 0
        End If
        rsTemp.Close
        strSql = "insert into mp_monthly_activity(mp_medium_id,month_number,month_name,[quarter],budget,min_budget,gross_budget,MSC_Paid,MSC_Paid_On_Flag,MSC_Paid_Value,MSC_Bonus,MSC_Bonus_On_Flag,MSC_Bonus_Value,Club_Agency,Club_Agency_On_Flag,Club_Agency_Value,approval) values ('" & strMediumID & "'," & i & ",'" & EngMonthName(i) & "'," & IIf(i Mod 3 = 0, i \ 3, i \ 3 + 1) & ",0,0,0," & MSC_Paid & "," & MSC_Paid_On_Flag & "," & 0 & "," & MSC_Bonus & "," & MSC_Bonus_On_Flag & ",0," & Club_Agency & "," & Club_Agency_On_Flag & ",0,0)"
        ConnERP.Execute strSql
    Next
End Sub

Private Sub picButton_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single) 'TOOLBAR_AI.
'************************************************
' Procedure         : picButton_MouseMove
' Function          : TOOLBAR_AI saat mouse berada di area button.
' Created By        : {73 64 6B}
' Date              : 12-Apr-2015
'************************************************
    picButton_Obj Index, Button, Shift, X, Y, picButton
End Sub

Private Sub picButton_Click(Index As Integer)

'************************************************
' Procedure         : picButton_Click
' Function          : Action utk Navigation dan CRUD.
' Created By        : {73 64 6B}
' Date              : 12-Apr-2015/{73 64 6B} --> Semua coding dan query sudah di optimalkan agar faster, readable, safer, standardable.
'************************************************

    Select Case Index
        Case enButtonType.bieSave  'call db_New.
            Call db_New
        Case Else
            Call db_cancel
    End Select

End Sub
