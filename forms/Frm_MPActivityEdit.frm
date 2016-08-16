VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Begin VB.Form Frm_MPActivityEdit 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Edit Activity"
   ClientHeight    =   3930
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6015
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3930
   ScaleWidth      =   6015
   StartUpPosition =   1  'CenterOwner
   Begin Threed.SSPanel pnl_Main 
      Align           =   1  'Align Top
      Height          =   3120
      Left            =   0
      TabIndex        =   3
      Top             =   750
      Width           =   6015
      _Version        =   65536
      _ExtentX        =   10610
      _ExtentY        =   5503
      _StockProps     =   15
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
         Height          =   2820
         Left            =   135
         TabIndex        =   4
         Top             =   90
         Width           =   5775
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
            Left            =   2160
            Style           =   2  'Dropdown List
            TabIndex        =   16
            Top             =   240
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
            Left            =   2160
            Style           =   2  'Dropdown List
            TabIndex        =   15
            Top             =   975
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
            Left            =   2160
            Style           =   2  'Dropdown List
            TabIndex        =   14
            Top             =   1350
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
            Left            =   2160
            TabIndex        =   13
            Text            =   "txtBrandTarget"
            Top             =   1740
            Width           =   3375
         End
         Begin VB.ComboBox cboBrandVariantCode 
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
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
            Left            =   4560
            Style           =   2  'Dropdown List
            TabIndex        =   12
            Top             =   975
            Visible         =   0   'False
            Width           =   1005
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
            Left            =   3210
            TabIndex        =   11
            Top             =   2460
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
            Height          =   195
            Left            =   4125
            TabIndex        =   10
            Top             =   2160
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
            Height          =   195
            Left            =   3210
            TabIndex        =   9
            Top             =   2160
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
            Height          =   195
            Left            =   2130
            TabIndex        =   8
            Top             =   2160
            Width           =   615
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
            Height          =   195
            Left            =   2130
            TabIndex        =   7
            Top             =   2460
            Width           =   975
         End
         Begin VB.ComboBox cboTargetAudienceCode 
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
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
            Left            =   4545
            Style           =   2  'Dropdown List
            TabIndex        =   6
            Top             =   1335
            Visible         =   0   'False
            Width           =   1005
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
            Left            =   2160
            TabIndex        =   5
            Text            =   "txtDescription"
            Top             =   615
            Width           =   3375
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
            Left            =   90
            TabIndex        =   22
            Top             =   225
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
            Left            =   105
            TabIndex        =   21
            Top             =   975
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
            Left            =   105
            TabIndex        =   20
            Top             =   1350
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
            Left            =   105
            TabIndex        =   19
            Top             =   1755
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
            Left            =   105
            TabIndex        =   18
            Top             =   2130
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
            Left            =   90
            TabIndex        =   17
            Top             =   615
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
      ScaleWidth      =   6015
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   0
      Width           =   6015
      Begin VB.PictureBox picButton 
         Appearance      =   0  'Flat
         BackColor       =   &H00F0F0F0&
         BorderStyle     =   0  'None
         ClipControls    =   0   'False
         ForeColor       =   &H80000008&
         Height          =   750
         Index           =   23
         Left            =   1620
         Picture         =   "Frm_MPActivityEdit.frx":0000
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
         Index           =   10
         Left            =   90
         Picture         =   "Frm_MPActivityEdit.frx":1D07
         ScaleHeight     =   750
         ScaleWidth      =   1500
         TabIndex        =   1
         TabStop         =   0   'False
         Top             =   0
         Width           =   1500
      End
   End
End
Attribute VB_Name = "Frm_MPActivityEdit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
''*****************************************************************************
' Nama Submodul         :  FrmMPActivityEdit
' Fungsi Submodul       :  Untuk Edit Activity
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

Private Sub loadcatalog()
    '*****************************************************************************
    ' Nama Prosedur     :   LoadCatalog
    ' Fungsi Prosedur   :   loading Catalog
    ' Parameter  Input  :
    ' Parameter Output  :
    ' Tgl Pembuatan     :   09 Agustus 2004
    ' Last Update/By    :   09 Agustus 2004/Sistyo
    '*****************************************************************************

    Dim strSql As String, strActivity_ID As String
    Dim idx As Integer
    
    rsTemp.Open "select campaign_type_name from campaign_type_catalog", ConnERP, 1, 3
    While Not rsTemp.EOF
        cboActivity.AddItem rsTemp(0)
        rsTemp.MoveNext
    Wend
    rsTemp.Close
    
    If Trim(frm_MPEdit.tdg_Activity.Columns(3)) <> "" Then
        cboActivity.ListIndex = CheckExitCombo(cboActivity, frm_MPEdit.tdg_Activity.Columns(3))
    End If
    
    txtDescription.Text = ""
    If Trim(frm_MPEdit.tdg_Activity.Columns(4)) <> "" Then
        txtDescription.Text = frm_MPEdit.tdg_Activity.Columns(4)
    End If
    
    strSql = "select brand_variant_code,brand_variant_name from brand_variant "
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
    
    If Trim(frm_MPEdit.tdg_Activity.Columns(5)) <> "" Then
        'cboBrandVariant.Text = frm_MPEdit.tdg_Activity.Columns(5)
        cboBrandVariant.ListIndex = CheckExitCombo(cboBrandVariant, frm_MPEdit.tdg_Activity.Columns(5))
    End If
    
    rsTemp.Open "select cluster,code from cluster", ConnERP, 1, 3
    idx = 0
    While Not rsTemp.EOF
        cboTargetAudience.AddItem rsTemp(0), idx
        cboTargetAudienceCode.AddItem rsTemp(1), idx
        rsTemp.MoveNext
        idx = idx + 1
    Wend
    rsTemp.Close
    
    If Trim(frm_MPEdit.tdg_Activity.Columns(6)) <> "" Then
        'cboTargetAudience.Text = frm_MPEdit.tdg_Activity.Columns(6)
        cboTargetAudience.ListIndex = CheckExitCombo(cboTargetAudience, frm_MPEdit.tdg_Activity.Columns(6))
    End If
        
    txtBrandTarget.Text = frm_MPEdit.tdg_Activity.Columns(7)
    
    strActivity_ID = frm_MPEdit.tdg_Activity.Columns(1)
    
    'Periksa Cek Bok Medium
    rsTemp.Open "select count(*) from mp_medium where medium_code='TV' and mp_activity_id='" & strActivity_ID & "'"
    If rsTemp(0) = 0 Then
        cbkTV.Value = 0
    Else
        cbkTV.Value = 1
    End If
    rsTemp.Close
      
    rsTemp.Open "select count(*) from mp_medium where medium_code='RD' and mp_activity_id='" & strActivity_ID & "'"
    If rsTemp(0) = 0 Then
        cbkRadio.Value = 0
    Else
        cbkRadio.Value = 1
    End If
    rsTemp.Close
    
    rsTemp.Open "select count(*) from mp_medium where medium_code='PR' and mp_activity_id='" & strActivity_ID & "'"
    If rsTemp(0) = 0 Then
        cbkPrint.Value = 0
    Else
        cbkPrint.Value = 1
    End If
    rsTemp.Close
    
    rsTemp.Open "select count(*) from mp_medium where medium_code='CN' and mp_activity_id='" & strActivity_ID & "'"
    If rsTemp(0) = 0 Then
        cbkCinema.Value = 0
    Else
        cbkCinema.Value = 1
    End If
    rsTemp.Close
    
    rsTemp.Open "select count(*) from mp_medium where medium_code='OT' and mp_activity_id='" & strActivity_ID & "'"
    If rsTemp(0) = 0 Then
        cbkOther.Value = 0
    Else
        cbkOther.Value = 1
    End If
    rsTemp.Close
    
    'enable / disable cek bok
    rsTemp.Open "select count(*) from mp_medium_detail where mp_medium_id in (select mp_medium_id from mp_medium where medium_code='TV' and mp_activity_id='" & strActivity_ID & "')"
    If rsTemp(0) <> 0 Then
        cbkTV.Enabled = False
    End If
    rsTemp.Close
    
    rsTemp.Open "select count(*) from mp_medium_detail where mp_medium_id in (select mp_medium_id from mp_medium where medium_code='RD' and mp_activity_id='" & strActivity_ID & "')"
    If rsTemp(0) <> 0 Then
        cbkRadio.Enabled = False
    End If
    rsTemp.Close
    
    rsTemp.Open "select count(*) from mp_medium_detail where mp_medium_id in (select mp_medium_id from mp_medium where medium_code='PR' and mp_activity_id='" & strActivity_ID & "')"
    If rsTemp(0) <> 0 Then
        cbkPrint.Enabled = False
    End If
    rsTemp.Close
    
    rsTemp.Open "select count(*) from mp_medium_detail where mp_medium_id in (select mp_medium_id from mp_medium where medium_code='CN' and mp_activity_id='" & strActivity_ID & "')"
    If rsTemp(0) <> 0 Then
        cbkCinema.Enabled = False
    End If
    rsTemp.Close
    
    rsTemp.Open "select count(*) from mp_medium_detail where mp_medium_id in (select mp_medium_id from mp_medium where medium_code='OT' and mp_activity_id='" & strActivity_ID & "')"
    If rsTemp(0) <> 0 Then
        cbkOther.Enabled = False
    End If
    rsTemp.Close
    
End Sub

Function CheckExitCombo(objCombo As ComboBox, saStrString As String) As String
    '<CSCM>
    '********************************************************************************
    'Procedure Name     : CheckExitCombo
    'Procedure Function : Check apakah Nilai yang dimasukan ada didalam combo
    'Input Parameter    : ---
    'Output Parameter   : ---
    'Date               : 3/29/2016
    'LastUpdate/By      : Tedi / Kreatif
    'Name Before        : -
    '********************************************************************************
    '</CSCM>
    
    On Error GoTo errSet
    
    objCombo.Text = Trim(saStrString)
    CheckExitCombo = objCombo.ListIndex
        
    Exit Function
    
errSet:
    MsgBox "There is no [" & Trim(saStrString) & "] Text Value  On " & objCombo.Name, vbExclamation, strApplication_Name
    CheckExitCombo = objCombo.ListIndex

End Function

Private Sub cboBrandVariant_click()
    cboBrandVariantCode.ListIndex = cboBrandVariant.ListIndex

End Sub

Private Sub saveactivity()
    '*****************************************************************************
    ' Nama Prosedur     :   saveactivity
    ' Fungsi Prosedur   :   Menyimpan Activity
    ' Parameter  Input  :
    ' Parameter Output  :
    ' Tgl Pembuatan     :   09 Agustus 2004
    ' Last Update/By    :   09 Agustus 2004/Sistyo
    '*****************************************************************************
    Dim strSql As String, strActivity_ID As String, strTask_Id As String, strMedium As String, strCurrentMedium As String
    Dim strMediumID As String
    Dim pesan
    Dim rsTemp As New ADODB.Recordset
    Dim original_brand_code As String, original_brand_name As String
    'Get original brand
    rsTemp.Open "select original_brand_code,original_brand_name from brand_variant where brand_variant_code = '" & cboBrandVariantCode.Text & "'", ConnERP, 1, 3
    If Not rsTemp.EOF Then
        original_brand_code = rsTemp("original_brand_code")
        original_brand_name = rsTemp("original_brand_name")
    End If
    rsTemp.Close
    
    'save mp_activity
    strActivity_ID = frm_MPEdit.tdg_Activity.Columns(1)
    strSql = "update mp_activity set activity_type='" & Clear_String(cboActivity.Text) & "', activity_desc='" & Clear_String(txtDescription.Text) & "',brand_variant_code='" & cboBrandVariantCode.Text & "',"
    strSql = strSql & "brand_variant_name='" & Clear_String(cboBrandVariant.Text) & "',target_audience='" & Clear_String(cboTargetAudience.Text) & "',brand_target='" & Clear_String(txtBrandTarget.Text) & "', target_audience_code = " & cboTargetAudienceCode.Text
    strSql = strSql & ",original_brand_code = '" & original_brand_code & "',original_brand_name = '" & Clear_String(original_brand_name) & "'"
    strSql = strSql & " where mp_activity_id='" & strActivity_ID & "'"
    ConnERP.Execute strSql
    
    'save mp_medium
    On Error GoTo err_Insert_Medium
    strCurrentMedium = frm_MPEdit.tdg_Activity.Columns(frm_MPEdit.tdg_Activity.Columns.Count - 1)
    strMedium = ""
    If cbkTV.Value = 1 Then
        If InStr(1, strCurrentMedium, "TV") = 0 Then
            
            rsTemp.Open "select count(*) from mp_medium where medium_code = 'TV' and mp_activity_id='" & strActivity_ID & "'", ConnERP, 1, 3
            If rsTemp(0) = 0 Then
                strMediumID = NextMPMediumID(strActivity_ID)
                ConnERP.Execute "insert into mp_medium values('" & strMediumID & "','" & strActivity_ID & "','TV','TV')"
                Call InsertMonthlyActivity(strMediumID)
            End If
            rsTemp.Close
        End If
        strMedium = strMedium & ", TV"
    Else
        ConnERP.Execute "delete from mp_medium where medium_code='TV' and mp_activity_id='" & strActivity_ID & "'"
    End If
    
    If cbkRadio.Value = 1 Then
        If InStr(1, strCurrentMedium, "Radio") = 0 Then
            
            rsTemp.Open "select count(*) from mp_medium where medium_code = 'RD' and mp_activity_id='" & strActivity_ID & "'", ConnERP, 1, 3
            If rsTemp(0) = 0 Then
                strMediumID = NextMPMediumID(strActivity_ID)
                ConnERP.Execute "insert into mp_medium values('" & strMediumID & "','" & strActivity_ID & "','RD','Radio')"
                Call InsertMonthlyActivity(strMediumID)
            End If
            rsTemp.Close
        End If
        strMedium = strMedium & ", Radio"
    Else
        ConnERP.Execute "delete from mp_medium where medium_code='RD' and mp_activity_id='" & strActivity_ID & "'"
    End If
    
    If cbkPrint.Value = 1 Then
        If InStr(1, strCurrentMedium, "Print") = 0 Then
            
            rsTemp.Open "select count(*) from mp_medium where medium_code = 'PR' and mp_activity_id='" & strActivity_ID & "'", ConnERP, 1, 3
            If rsTemp(0) = 0 Then
                strMediumID = NextMPMediumID(strActivity_ID)
                ConnERP.Execute "insert into mp_medium values('" & strMediumID & "','" & strActivity_ID & "','PR','Print')"
                Call InsertMonthlyActivity(strMediumID)
            End If
            rsTemp.Close
        End If
        strMedium = strMedium & ", Print"
    Else
        ConnERP.Execute "delete from mp_medium where medium_code='PR' and mp_activity_id='" & strActivity_ID & "'"
    End If
    
    If cbkCinema.Value = 1 Then
        If InStr(1, strCurrentMedium, "Cinema") = 0 Then
            
            rsTemp.Open "select count(*) from mp_medium where medium_code = 'CN' and mp_activity_id='" & strActivity_ID & "'", ConnERP, 1, 3
            If rsTemp(0) = 0 Then
                strMediumID = NextMPMediumID(strActivity_ID)
                ConnERP.Execute "insert into mp_medium values('" & strMediumID & "','" & strActivity_ID & "','CN','Cinema')"
                Call InsertMonthlyActivity(strMediumID)
            End If
            rsTemp.Close
        End If
        strMedium = strMedium & ", Cinema"
    Else
        ConnERP.Execute "delete from mp_medium where medium_code='CN' and mp_activity_id='" & strActivity_ID & "'"
    End If
    
    If cbkOther.Value = 1 Then
        If InStr(1, strCurrentMedium, "Other") = 0 Then
        
            rsTemp.Open "select count(*) from mp_medium where medium_code = 'OT' and mp_activity_id='" & strActivity_ID & "'", ConnERP, 1, 3
            If rsTemp(0) = 0 Then
                strMediumID = NextMPMediumID(strActivity_ID)
                ConnERP.Execute "insert into mp_medium values('" & strMediumID & "','" & strActivity_ID & "','OT','Other')"
                Call InsertMonthlyActivity(strMediumID)
            End If
            rsTemp.Close
        End If
        strMedium = strMedium & ", Other"
    Else
        ConnERP.Execute "delete from mp_medium where medium_code='OT' and mp_activity_id='" & strActivity_ID & "'"
    End If
err_Insert_Medium:
    If Err.Number <> 0 Then
        MsgBox Err.Description, vbExclamation, strApplication_Name
        'MsgBox "Can not insert medium because Brand Fee Catalog has not been setup!" & vbCrLf & "Activity will be saved but medium will be ignored!"
        If rsTemp.State = adStateOpen Then
            rsTemp.Close
        End If
    End If
    'update mp_master
    ConnERP.Execute "update mp_master set last_update_by='" & strLogin_User & "',last_update_date = getdate() where mp_number='" & frm_MPEdit.cboMPNum.Text & "'"
    
'    With FrmMPEdit.FGActivity
'        .TextMatrix(.Row, 3) = cboActivity.Text
'        .TextMatrix(.Row, 4) = txtDescription.Text
'        .TextMatrix(.Row, 5) = cboBrandVariant.Text
'        .TextMatrix(.Row, 6) = cboTargetAudience.Text
'        .TextMatrix(.Row, 7) = txtBrandTarget.Text
'        If Len(strMedium) <> 0 Then
'            strMedium = Right(strMedium, Len(strMedium) - 2)
'        End If
'        .TextMatrix(.Row, .cols - 1) = strMedium
'    End With
    
    pesan = MsgBox("Activity Saved!", vbExclamation, strApplication_Name)

End Sub

Private Sub db_save()
    '<CSCM>
    '********************************************************************************
    'Procedure Name     : db_Save
    'Procedure Function : ---
    'Input Parameter    : ---
    'Output Parameter   : ---
    'Date               : 3/29/2016
    'LastUpdate/By      : Tedi / Kreatif
    'Name Before        : cmdSave_Click
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
    Call saveactivity
    Unload Me

End Sub

Private Sub db_Cancel()
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
        Case 10  'call db_New.
            Call db_save
        Case Else
            Call db_Cancel
    End Select

End Sub

