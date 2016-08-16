VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Begin VB.Form Frm_MPCreate 
   BackColor       =   &H00F0F0F0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Create New Media Plan"
   ClientHeight    =   3345
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5310
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3345
   ScaleWidth      =   5310
   StartUpPosition =   1  'CenterOwner
   Begin Threed.SSPanel pnl_Main 
      Align           =   1  'Align Top
      Height          =   2550
      Left            =   0
      TabIndex        =   3
      Top             =   750
      Width           =   5310
      _Version        =   65536
      _ExtentX        =   9366
      _ExtentY        =   4498
      _StockProps     =   15
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
      Begin VB.Frame frame1 
         Height          =   2445
         Left            =   90
         TabIndex        =   4
         Top             =   15
         Width           =   5175
         Begin VB.ComboBox cboYear 
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
            Left            =   1530
            Style           =   2  'Dropdown List
            TabIndex        =   18
            Top             =   1635
            Width           =   3465
         End
         Begin VB.TextBox txtMPNum 
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
            Height          =   285
            Left            =   1545
            Locked          =   -1  'True
            TabIndex        =   17
            Text            =   "(Automatic)"
            Top             =   240
            Width           =   3435
         End
         Begin VB.ComboBox cboBrandName 
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
            Left            =   1530
            Sorted          =   -1  'True
            Style           =   2  'Dropdown List
            TabIndex        =   10
            Top             =   570
            Width           =   3465
         End
         Begin VB.ComboBox cboCountry 
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
            Left            =   1530
            Style           =   2  'Dropdown List
            TabIndex        =   9
            Top             =   1260
            Width           =   3465
         End
         Begin VB.ListBox lstBrandCode 
            BackColor       =   &H80000013&
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   3615
            Sorted          =   -1  'True
            TabIndex        =   8
            Top             =   270
            Visible         =   0   'False
            Width           =   855
         End
         Begin VB.TextBox txtClientName 
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
            Left            =   1530
            TabIndex        =   7
            Top             =   930
            Width           =   3450
         End
         Begin VB.ListBox lstClientName 
            BackColor       =   &H80000013&
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   3720
            Sorted          =   -1  'True
            TabIndex        =   6
            Top             =   1320
            Visible         =   0   'False
            Width           =   855
         End
         Begin VB.TextBox txt_budget 
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
            Left            =   1515
            TabIndex        =   5
            Top             =   1995
            Width           =   3450
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "MP Number  "
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   195
            Left            =   120
            TabIndex        =   16
            Top             =   285
            Width           =   900
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Brand Name  "
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
            Left            =   120
            TabIndex        =   15
            Top             =   615
            Width           =   960
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Client Name  "
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
            Left            =   120
            TabIndex        =   14
            Top             =   945
            Width           =   945
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Country  "
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
            Left            =   120
            TabIndex        =   13
            Top             =   1305
            Width           =   675
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Plan Year  "
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
            Left            =   120
            TabIndex        =   12
            Top             =   1650
            Width           =   765
         End
         Begin VB.Label Label6 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Budget  "
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
            Left            =   120
            TabIndex        =   11
            Top             =   2010
            Width           =   600
         End
      End
   End
   Begin VB.PictureBox picToolbar 
      Align           =   1  'Align Top
      Appearance      =   0  'Flat
      BackColor       =   &H00F0F0F0&
      BorderStyle     =   0  'None
      ClipControls    =   0   'False
      ForeColor       =   &H80000008&
      Height          =   750
      Left            =   0
      ScaleHeight     =   750
      ScaleWidth      =   5310
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   0
      Width           =   5310
      Begin VB.PictureBox picButton 
         Appearance      =   0  'Flat
         BackColor       =   &H00F0F0F0&
         BorderStyle     =   0  'None
         ClipControls    =   0   'False
         ForeColor       =   &H80000008&
         Height          =   750
         Index           =   11
         Left            =   1620
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
         ScaleHeight     =   750
         ScaleWidth      =   1500
         TabIndex        =   1
         TabStop         =   0   'False
         Top             =   0
         Width           =   1500
      End
   End
End
Attribute VB_Name = "Frm_MPCreate"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
''*****************************************************************************
' Nama Submodul         :  Frm_MPCreate
' Fungsi Submodul       :  Untuk Create Media Plan Baru
' Tabel yg digunakan    :  MP_Master (R/W)
' Prosedur/Function     :  loadCatalog, Saveplan
' Nama Programmer       :  Sistyo
' Tgl Pembuatan         :  09 Agustus 2004
' Last Update           :  09 Agustus 2004/Sistyo
'******************************************************************************

Option Explicit
Dim recTemp As New ADODB.Recordset

Private Sub Form_Load()
'<CSCM>
'********************************************************************************
'Procedure Name     : Form_Load
'Procedure Function : ---
'Input Parameter    : ---
'Output Parameter   : ---
'Date               : 3/17/2016
'LastUpdate/By      : Tedi / Kreatif
'Name Before        : -
'********************************************************************************
'</CSCM>
    
    Call loadcatalog
    EnableObject False

End Sub

Private Sub loadcatalog()
'*****************************************************************************
' Nama Prosedur     :   loadCatalog
' Fungsi Prosedur   :   Loading Brand,Client,Country,Year from Catalog
' Parameter  Input  :
' Parameter Output  :
' Tgl Pembuatan     :   09 Agustus 2004
' Last Update/By    :   11 July 2005/Sistyo -->brand.is_media_plan_brand=1
'*****************************************************************************
    Dim intIDX, intCount As Integer
    Dim strSql As String
    Dim strBrand_Filter As String
    
    strBrand_Filter = "select brand_code from media_security_catalog where user_name = '" & strLogin_User & "' and position = 'Planner' and valid_until>=(select getdate())"
    
    strSql = "select  a.brand_code,a.brand_name,b.client_name from brand a inner join client b"
    strSql = strSql & " on a.client_code = b.client_code"
    strSql = strSql & " where a.brand_code in (" & strBrand_Filter & ") and a.is_media_plan_brand=1"
    strSql = strSql & " order by a.brand_code"
    recTemp.Open strSql, ConnERP, 1, 3
    intIDX = 0
    While Not recTemp.EOF
        lstBrandCode.AddItem recTemp(0), intIDX
        cboBrandName.AddItem recTemp(0) & "->" & recTemp(1), intIDX
        lstClientName.AddItem recTemp(2), intIDX
        intIDX = intIDX + 1
        recTemp.MoveNext
    Wend
    recTemp.Close
    
    recTemp.Open "select country from country order by country", ConnERP, 1, 3
    intIDX = 0
    While Not recTemp.EOF
        cboCountry.AddItem recTemp(0), intIDX
        intIDX = intIDX + 1
        recTemp.MoveNext
    Wend
    recTemp.Close
    
    For intCount = (Year(recDate(0)) - 1) To (Year(recDate(0)) + 30)
        cboYear.AddItem intCount
    Next
    cboYear.Text = Year(recDate(0))
    
End Sub

Private Sub SavePlan()
'*****************************************************************************
' Nama Prosedur     :   SavePlan
' Fungsi Prosedur   :   Saving Media Plan
' Parameter  Input  :
' Parameter Output  :
' Tgl Pembuatan     :   09 Agustus 2004
' Last Update/By    :   25 Feb 2005/Sistyo
'*****************************************************************************
    Dim strSql As String
    Dim intPesan As Integer
    Dim strMPNum As String
    Dim strBrandName As String
    
    'Generate MP_Number
    strSql = "select isnull(max(cast(substring(mp_number,11,4) as int)),0)+1 from mp_master where brand_code='" & Mid(cboBrandName.Text, 1, 4) & "' and Year = " & cboYear.Text
    recTemp.Open strSql, ConnERP, 1, 3
    strMPNum = Mid(cboBrandName, 1, 4) & "." & cboYear.Text & "." & Right("0000" & CStr(recTemp(0)), 4)
    recTemp.Close
    strBrandName = Mid(cboBrandName.Text, InStr(1, cboBrandName.Text, ">") + 1, Len(cboBrandName.Text) - InStr(1, cboBrandName.Text, ">"))
    If Trim(txt_budget.Text) = "" Then txt_budget.Text = "0"
        'save new plan
        strSql = "insert into mp_master(mp_number,brand_code,brand_name,client_name,country,[year],created_by,created_date,is_latest,yearly_budget) values ('"
        strSql = strSql & strMPNum & "','" & lstBrandCode.Text & "','"
        strSql = strSql & Clear_String(strBrandName) & "','" & Clear_String(txtClientName.Text) & "','" & Clear_String(cboCountry.Text) & "','"
        strSql = strSql & cboYear.Text & "','" & strLogin_User & "', getdate(),1," & RemoveNumberFormat(txt_budget.Text) & " )"
        
        ConnERP.Execute strSql
    
    txtMPNum.Text = strMPNum
    intPesan = MsgBox("New Media plan created!", vbExclamation, strApplication_Name)
    
End Sub

Private Sub cboBrandName_click()
'<CSCM>
'********************************************************************************
'Procedure Name     : cboBrandName_click
'Procedure Function : ---
'Input Parameter    : ---
'Output Parameter   : ---
'Date               : 3/17/2016
'LastUpdate/By      : Tedi / Kreatif
'Name Before        : -
'********************************************************************************
'</CSCM>

    lstBrandCode.ListIndex = cboBrandName.ListIndex
    lstClientName.ListIndex = cboBrandName.ListIndex
    txtClientName.Text = lstClientName.Text

End Sub

Private Sub db_Create()
'<CSCM>
'********************************************************************************
'Procedure Name     : db_Create
'Procedure Function : Create Media Plan
'Input Parameter    : ---
'Output Parameter   : ---
'Date               : 3/17/2016
'LastUpdate/By      : Tedi / Kreatif
'Name Before        : cmdCreate_Click
'********************************************************************************
'</CSCM>

    Dim strSql As String
    Dim recCheck As New ADODB.Recordset
    
    Dim intPesan As Integer
    
    If cboBrandName.Text = "" Then
        intPesan = MsgBox("Please select Brand!", vbExclamation, strApplication_Name)
        Exit Sub
    End If
    
    If Trim(txtClientName.Text) = "" Then
        intPesan = MsgBox("Please Enter Client Name!", vbExclamation, strApplication_Name)
        Exit Sub
    End If
    
    If cboCountry.Text = "" Then
        intPesan = MsgBox("Please select Country!", vbExclamation, strApplication_Name)
        Exit Sub
    End If
    
    If Not IsFeeAlreadyEntered(Left(cboBrandName.Text, 4), cboYear.Text) Then
        MsgBox "Brand Fee Information not found for selected year !", vbExclamation, strApplication_Name
        Exit Sub
    End If
    
    strSql = "SELECT count(*) FROM MP_Master WHERE Year=" & cboYear.Text & " AND Brand_COde='" & Left(cboBrandName.Text, 4) & "' AND Is_latest=1"
    recCheck.Open strSql, ConnERP, 3, 1
    If recCheck.Fields(0).Value > 0 Then
        Call CloseRecordset(recCheck)
        MsgBox "You can't have more than one active Media Plan in one year per brand.", vbExclamation, strApplication_Name
        Exit Sub
    Else
        recCheck.Close
    End If
    
    Call SavePlan
    
    Select Case objOpener
        Case "Frm_MPEdit"
            frm_MPEdit.cboMPNum.AddItem txtMPNum.Text
            frm_MPEdit.cboMPNum.Text = txtMPNum.Text
        Case "frm_MPInsertion"
            frm_MPInsertion.cboMPNumber.AddItem txtMPNum.Text
            frm_MPInsertion.cboMPNumber.Text = txtMPNum.Text
    End Select
    
    Unload Me
End Sub

Private Sub db_cancel()
'<CSCM>
'********************************************************************************
'Procedure Name     : db_Cancel
'Procedure Function : Cancel Proses
'Input Parameter    : ---
'Output Parameter   : ---
'Date               : 3/17/2016
'LastUpdate/By      : Tedi / Kreatif
'Name Before        : cmdCancel_Click
'********************************************************************************
'</CSCM>

    Unload Me

End Sub

Private Sub txt_budget_GotFocus()
'<CSCM>
'********************************************************************************
'Procedure Name     : txt_budget_GotFocus
'Procedure Function :
'Input Parameter    : ---
'Output Parameter   : ---
'Date               : 3/17/2016
'LastUpdate/By      : Tedi / Kreatif
'Name Before        : -
'********************************************************************************
'</CSCM>
    
    txt_budget.Text = RemoveNumberFormat(txt_budget.Text)

End Sub

Private Sub txt_budget_KeyPress(KeyAscii As Integer)
'<CSCM>
'********************************************************************************
'Procedure Name     : txt_budget_KeyPress
'Procedure Function : Filter untuk hanya memperbolehkan digit angka
'Input Parameter    : ---
'Output Parameter   : ---
'Date               : 3/17/2016
'LastUpdate/By      : Tedi / Kreatif
'Name Before        : -
'********************************************************************************
'</CSCM>

    If (Chr(KeyAscii) < "0" Or Chr(KeyAscii) > "9") And KeyAscii <> 8 And KeyAscii <> 13 And KeyAscii <> 27 Then
            If Chr(KeyAscii) <> "." Then
                KeyAscii = 0
                Beep
            End If
    End If

End Sub

Private Sub txt_budget_LostFocus()
'<CSCM>
'********************************************************************************
'Procedure Name     : txt_budget_LostFocus
'Procedure Function : Seting format Number
'Input Parameter    : ---
'Output Parameter   : ---
'Date               : 3/17/2016
'LastUpdate/By      : Tedi / Kreatif
'Name Before        : -
'********************************************************************************
'</CSCM>

    txt_budget.Text = FormatNumber(Val(txt_budget.Text), 2)
    
End Sub

Sub SetButtonToolbar(ByVal paIsNormalMode As Boolean, picOBJ) 'TOOLBAR_AI.
'<CSCM>
'********************************************************************************
'Procedure Name     : SetButtonToolbar
'Procedure Function : SetLayout Toolbar
'Input Parameter    : ---
'Output Parameter   : ---
'Date               : 3/17/2016
'LastUpdate/By      : Tedi / Kreatif
'Name Before        : -
'********************************************************************************
'</CSCM>


    Dim objElement As Object
    Dim strDummy As String
    
    With picButton(enButtonType.bieSave)    'ADD. 4
        .Enabled = paIsNormalMode
        .Visible = paIsNormalMode
    End With
    With picButton(enButtonType.bieCancel) 'EDIT. 5
        .Enabled = paIsNormalMode
        .Visible = paIsNormalMode
    End With

    For Each objElement In picOBJ
        SetPictureTB objElement.Index, paIsNormalMode, picOBJ
    Next objElement
    'Call SetSecurityCRUDStandar("Duration Catalog", picButton, "1")

End Sub

Private Sub picButton_Click(Index As Integer)

'************************************************
' Procedure         : picButton_Click
' Function          : Action utk Navigation dan CRUD.
' Created By        : {73 64 6B}
' Date              : 12-Apr-2015/{73 64 6B} --> Semua coding dan query sudah di optimalkan agar faster, readable, safer, standardable.
'************************************************
    Dim strCode As String, strFileRpt As String
    'Lock_MainForm True
    Select Case Index
        Case enButtonType.bieSave 'Create.
            Call db_Create
            'tdb_Task_Click
        Case enButtonType.bieCancel 'Cancel.
            Call db_cancel
    End Select

End Sub

Private Sub picButton_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single) 'TOOLBAR_AI.
'************************************************
' Procedure         : picButton_MouseDown
' Function          : TOOLBAR_AI saat mouse ditekan.
' Created By        : {73 64 6B}
' Date              : 12-Apr-2015
'************************************************
    
    picButton(Index).Picture = LoadPicture(SetButtonImageEffect(Index, bieDown)) 'FIRST.

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

Private Sub EnableObject(ByVal paIsEnable As Boolean)
'*****************************************
'Procedure Name     : EnableObject
'Procedure Function : ~ Enable/disable control di frame Entry.
'                     ~ Call SetButtonToolbar utk Toolbar/Statusbar AI (artificial intelligence).
'Input Parameter    : paIsEnable: True=Enable, False=Disable.
'Output Parameter   : -
'Date               : 12-Apr-2015
'LastUpdate/By      : 12-Apr-2015/{73 64 6B}
'*****************************************
    
    Call SetButtonToolbar(Not paIsEnable, picButton) 'TOOLBAR_AI.

End Sub

