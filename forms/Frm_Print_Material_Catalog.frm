VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.ocx"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDatGrd.ocx"
Begin VB.Form Frm_Print_Material_Catalog 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Print Material Catalog"
   ClientHeight    =   4920
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6225
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4920
   ScaleWidth      =   6225
   StartUpPosition =   2  'CenterScreen
   Begin Threed.SSPanel pnl_Main 
      Align           =   1  'Align Top
      Height          =   4170
      Left            =   0
      TabIndex        =   7
      Top             =   750
      Width           =   6225
      _Version        =   65536
      _ExtentX        =   10980
      _ExtentY        =   7355
      _StockProps     =   15
      BackColor       =   15790320
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BevelOuter      =   1
      Begin VB.Frame Frame1 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   3255
         Left            =   105
         TabIndex        =   11
         Top             =   825
         Width           =   6030
         Begin VB.TextBox Txt_Materi_Name 
            Appearance      =   0  'Flat
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
            Height          =   300
            Left            =   1260
            TabIndex        =   15
            Top             =   2400
            Width           =   2535
         End
         Begin VB.CheckBox CbkIsMMC 
            Appearance      =   0  'Flat
            Caption         =   "mmc"
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
            ForeColor       =   &H80000008&
            Height          =   240
            Left            =   4110
            TabIndex        =   14
            Top             =   2445
            Width           =   720
         End
         Begin VB.TextBox txt_column 
            Appearance      =   0  'Flat
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
            Height          =   300
            Left            =   1275
            TabIndex        =   13
            Top             =   2790
            Width           =   825
         End
         Begin VB.TextBox txt_size 
            Appearance      =   0  'Flat
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
            Height          =   300
            Left            =   2625
            TabIndex        =   12
            Top             =   2805
            Width           =   825
         End
         Begin MSDataGridLib.DataGrid DataGrid1 
            Height          =   2010
            Left            =   120
            TabIndex        =   16
            Top             =   225
            Width           =   5775
            _ExtentX        =   10186
            _ExtentY        =   3545
            _Version        =   393216
            AllowUpdate     =   0   'False
            HeadLines       =   1
            RowHeight       =   15
            BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ColumnCount     =   2
            BeginProperty Column00 
               DataField       =   ""
               Caption         =   ""
               BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
                  Type            =   0
                  Format          =   ""
                  HaveTrueFalseNull=   0
                  FirstDayOfWeek  =   0
                  FirstWeekOfYear =   0
                  LCID            =   1033
                  SubFormatType   =   0
               EndProperty
            EndProperty
            BeginProperty Column01 
               DataField       =   ""
               Caption         =   ""
               BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
                  Type            =   0
                  Format          =   ""
                  HaveTrueFalseNull=   0
                  FirstDayOfWeek  =   0
                  FirstWeekOfYear =   0
                  LCID            =   1033
                  SubFormatType   =   0
               EndProperty
            EndProperty
            SplitCount      =   1
            BeginProperty Split0 
               BeginProperty Column00 
               EndProperty
               BeginProperty Column01 
               EndProperty
            EndProperty
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Materi Name"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   300
            TabIndex        =   20
            Top             =   2400
            Width           =   900
         End
         Begin VB.Label Label2 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Size (if mmc)"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   300
            TabIndex        =   19
            Top             =   2835
            Width           =   900
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "col"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   2340
            TabIndex        =   18
            Top             =   2835
            Width           =   195
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "mm"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   3525
            TabIndex        =   17
            Top             =   2850
            Width           =   240
         End
      End
      Begin VB.Frame Frame2 
         Height          =   780
         Left            =   105
         TabIndex        =   8
         Top             =   15
         Width           =   6030
         Begin VB.ComboBox Cbo_Brand 
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
            Left            =   840
            Style           =   2  'Dropdown List
            TabIndex        =   9
            Top             =   285
            Width           =   3420
         End
         Begin VB.Label Label3 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Brand"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   300
            TabIndex        =   10
            Top             =   315
            Width           =   420
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
      ScaleWidth      =   6225
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   0
      Width           =   6225
      Begin VB.PictureBox picButton 
         Appearance      =   0  'Flat
         BackColor       =   &H00F0F0F0&
         BorderStyle     =   0  'None
         ClipControls    =   0   'False
         ForeColor       =   &H80000008&
         Height          =   750
         Index           =   10
         Left            =   6210
         ScaleHeight     =   750
         ScaleWidth      =   1500
         TabIndex        =   6
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
         Index           =   11
         Left            =   7740
         ScaleHeight     =   750
         ScaleWidth      =   1500
         TabIndex        =   5
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
         Index           =   6
         Left            =   3150
         ScaleHeight     =   750
         ScaleWidth      =   1500
         TabIndex        =   4
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
         Left            =   4680
         ScaleHeight     =   750
         ScaleWidth      =   1500
         TabIndex        =   3
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
         Index           =   4
         Left            =   90
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
         Index           =   5
         Left            =   1620
         ScaleHeight     =   750
         ScaleWidth      =   1500
         TabIndex        =   1
         TabStop         =   0   'False
         Top             =   0
         Width           =   1500
      End
   End
End
Attribute VB_Name = "Frm_Print_Material_Catalog"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
''*************************************************************
'Nama Form          : Frm_Print_Material_Catalog
'Fungsi Form        : untuk menambah, edit, delete material Print
'Programer          : sistyo
'Tgl Pembuatan      : 26 oct 2004
'Last Update/By     :
'*************************************************************

Option Explicit

Dim RS_Materi As New ADODB.Recordset
Dim strSql As String
Dim Edit As Boolean

Private Sub CbkIsMMC_Click()
    If CbkIsMMC.Value = 1 Then
        txt_size.Enabled = True
        txt_column.Enabled = True
    Else
        txt_size.Enabled = False
        txt_column.Enabled = False
    End If
End Sub

Private Sub cbo_brand_Click()
    Call Load_Data
End Sub

Private Sub enableText(blnStatus As Boolean)

    If blnStatus = True Then
        Txt_Materi_Name.Enabled = True
        txt_column.Enabled = True
        txt_size.Enabled = True
        Txt_Materi_Name.BackColor = &HFFFFFF
        txt_column.BackColor = &HFFFFFF
        txt_size.BackColor = &HFFFFFF
    Else
        Txt_Materi_Name.Enabled = False
        txt_column.Enabled = False
        txt_size.Enabled = False
        Txt_Materi_Name.BackColor = &HE0E0E0
        txt_column.BackColor = &HE0E0E0
        txt_size.BackColor = &HE0E0E0
    End If

End Sub

Private Sub db_add()
    Call Clear_Form
    Call enableText(True)
    Call EnableObject(True)
    CbkIsMMC.Value = 0
    Txt_Materi_Name.SetFocus
    Edit = False
End Sub

Private Sub db_cancel()
    Call Clear_Form
    Call enableText(False)
    Call EnableObject(False)
End Sub

Private Sub db_close()
    If RS_Materi.State = adStateOpen Then RS_Materi.Close
    Set RS_Materi = Nothing
    Unload Me
End Sub

Private Sub db_delete()
    On Error GoTo Lable
    Dim Tanya As Integer
    Tanya = MsgBox("Are you sure want to delete this Material", vbQuestion + vbYesNo, strCompany_Name)
    If Tanya = vbYes Then
        RS_Materi.Delete
    End If
    Exit Sub
    
Lable:
    MsgBox "Cannot Delete Material.", vbCritical, strCompany_Name
End Sub

Private Sub db_edit()
    On Error GoTo Lable
    If RS_Materi.State = adStateOpen Then
        If Not RS_Materi.EOF Then
            Txt_Materi_Name.Text = RS_Materi.Fields("material_name").Value
            Txt_Materi_Name.Enabled = True
            CbkIsMMC.Value = RS_Materi.Fields("ismmc").Value
            CbkIsMMC.Enabled = True
            txt_size.Text = RS_Materi.Fields("size").Value
            txt_column.Text = RS_Materi.Fields("column").Value
            
            If CbkIsMMC.Value = 1 Then
                txt_size.Enabled = True
                txt_column.Enabled = True
            Else
                txt_size.Enabled = False
                txt_column.Enabled = False
            End If
            
            Edit = True
            Call EnableObject(True)
            Call enableText(True)
        End If
    End If
    Exit Sub
Lable:
    MsgBox "Cannot Edit Material.", vbExclamation, strApplication_Name
    Call enableText(True)
    Call Clear_Form
End Sub

Private Sub db_save()
    Dim pesan
    If Trim(Txt_Materi_Name.Text) = "" Then
        pesan = MsgBox("Please enter material name!", vbExclamation, strApplication_Name)
        Exit Sub
    End If
    If CbkIsMMC.Value = 1 Then
       If Trim(txt_size.Text) = "" Then
            pesan = MsgBox("Please enter mmc size!", vbExclamation, strApplication_Name)
            Exit Sub
       End If
       If Trim(txt_column.Text) = "" Then
            pesan = MsgBox("Please enter mmc column!", vbExclamation, strApplication_Name)
            Exit Sub
       End If
    Else
        txt_size.Text = Empty
        txt_column.Text = Empty
    End If
    Call enableText(False)
    With RS_Materi
        If Not Edit Then
            .AddNew
            .Fields("brand_code").Value = Left(Cbo_Brand.Text, 4)
            .Fields("material_name").Value = Txt_Materi_Name.Text
            .Fields("ismmc").Value = CbkIsMMC.Value
            .Fields("size").Value = Val(txt_size.Text)
            .Fields("column").Value = Val(txt_column.Text)
            .Update
        Else
            .Fields("brand_code").Value = Left(Cbo_Brand.Text, 4)
            .Fields("material_name").Value = Txt_Materi_Name.Text
            .Fields("ismmc").Value = CbkIsMMC.Value
            .Fields("size").Value = Val(txt_size.Text)
            .Fields("column").Value = Val(txt_column.Text)
            .Update
        End If
    End With
    
    Call Clear_Form
    Call EnableObject(False)
    
End Sub

Private Sub Form_Load()
            
    Dim rs_brand As New ADODB.Recordset
    strSql = "SELECT * FROM brand WHERE brand_code IN "
    strSql = strSql & " (SELECT brand_code FROM Media_Security_Catalog WHERE User_name='" & strLogin_User & "' AND position IN ('Planner', 'Buyer', 'Implementor') AND Valid_until > getdate())"
    
    rs_brand.Open strSql, ConnERP, adOpenForwardOnly, adLockReadOnly, adCmdText
    
    Cbo_Brand.Clear
    While Not rs_brand.EOF
        Cbo_Brand.AddItem rs_brand("Brand_Code").Value & " -> " & rs_brand.Fields("Brand_Name").Value
        rs_brand.MoveNext
    Wend
    rs_brand.Close
    Set rs_brand = Nothing
    
    If Cbo_Brand.ListCount > 0 Then
        Cbo_Brand.Text = Cbo_Brand.List(0)
    End If
    
    If objOpener = "FrmMPActivityDetail" Then
        rs_brand.Open "select brand_code,brand_name from brand where brand_code = '" & Left(Frm_MPActivityDetail.txtMPNumber.Text, 4) & "'", ConnERP, 1, 3
        If Not rs_brand.EOF Then
           Cbo_Brand.Text = rs_brand("Brand_Code").Value & " -> " & rs_brand.Fields("Brand_Name").Value
        End If
        rs_brand.Close
    End If
    
    Call Load_Data
    Call enableText(False)
    Call EnableObject(False)

    Edit = False
End Sub

Public Sub Initial_Grid()

    DataGrid1.Columns(0).Width = 0
    DataGrid1.Columns(1).Width = 2000
    DataGrid1.Columns(2).Width = 600
    DataGrid1.Columns(3).Width = 800
    DataGrid1.Columns(4).Width = 800
End Sub

Public Sub Clear_Form()
    CbkIsMMC.Value = 0
    Txt_Materi_Name.Text = Empty
    txt_size.Text = Empty
    txt_column.Text = Empty
    Txt_Materi_Name.Enabled = False
    CbkIsMMC.Enabled = False
    txt_size.Enabled = False
    txt_column.Enabled = False
End Sub

Public Sub Load_Data()
    
        strSql = "SELECT * FROM Print_Material WHERE "
        strSql = strSql & " Brand_Code='" & Left(Cbo_Brand.Text, 4) & "' ORDER BY Material_Name"
        
        If RS_Materi.State = adStateOpen Then
            RS_Materi.Close
            Set RS_Materi = Nothing
        End If
        RS_Materi.CursorLocation = adUseClient
        'RS_Materi.Open StrSQL, ConnERP, adOpenDynamic, adLockPessimistic, adCmdText
        RS_Materi.Open strSql, ConnERP, adOpenKeyset, adLockPessimistic
    
        Set DataGrid1.DataSource = RS_Materi
        Call Initial_Grid
End Sub

Private Sub Form_Unload(Cancel As Integer)
    objOpener = Empty
End Sub

Private Sub txt_column_KeyPress(KeyAscii As Integer)
      If (Chr(KeyAscii) < "0" Or Chr(KeyAscii) > "9") And KeyAscii <> 8 Then
        'MsgBox KeyAscii
        KeyAscii = 0
        Beep
    End If
End Sub

Private Sub Txt_Materi_Name_GotFocus()
    Txt_Materi_Name.SelStart = 0
    Txt_Materi_Name.SelLength = Len(Txt_Materi_Name.Text)
End Sub


Private Sub Txt_Materi_Name_KeyPress(KeyAscii As Integer)
    If Chr(KeyAscii) = ":" Then
        KeyAscii = 0
        Beep
    End If
End Sub

Private Sub txt_size_KeyPress(KeyAscii As Integer)
      If (Chr(KeyAscii) < "0" Or Chr(KeyAscii) > "9") And KeyAscii <> 8 Then
        'MsgBox KeyAscii
        KeyAscii = 0
        Beep
    End If
End Sub

Sub SetButtonToolbar(ByVal paIsNormalMode As Boolean, picOBJ) 'TOOLBAR_AI.
'************************************************
' Procedure         : SetButtonToolbar
' Function          : TOOLBAR_AI.
' Created By        : {73 64 6B}
' Date              : 12-Apr-2015
' LastUpdate/By     : - Rudi
'************************************************

    Dim element
    Dim strDummy As String
    
    With picButton(enButtonType.bieAdd)  'ADD. 4
        .Enabled = paIsNormalMode
        .Visible = paIsNormalMode
    End With
    
    With picButton(enButtonType.bieEdit) 'EDIT. 5
        .Enabled = paIsNormalMode
        .Visible = paIsNormalMode
    End With
    
    With picButton(enButtonType.bieDelete)  'DELETE. 6
        .Enabled = paIsNormalMode
        .Visible = paIsNormalMode
    End With
    
    With picButton(enButtonType.bieClose)      'CLOSE.
        .Enabled = paIsNormalMode
        .Visible = paIsNormalMode
    End With

    With picButton(enButtonType.bieSave)  'SAVE.
        .Enabled = Not paIsNormalMode
        .Visible = Not paIsNormalMode
        .Left = picButton(4).Left
    End With

    With picButton(enButtonType.bieCancel) 'CANCEL.
        .Enabled = Not paIsNormalMode
        .Visible = Not paIsNormalMode
        .Left = picButton(5).Left
    End With
    
    For Each element In picOBJ
        SetPictureTB element.Index, paIsNormalMode, picOBJ
    Next element

End Sub

Sub SetPictureTB(ByVal Index As Integer, ByVal paIsNormalMode As Boolean, picOBJ)
 '*****************************************
'Procedure Name     : SetPictureTB
'Procedure Function :   Creates the SQL statement in ado_Data.recordset.filter
'                       and only filters text currently. It must be modified to filter other data types.
'Input Parameter    : Index,paIsNormalMode,picOBJ
'Output Parameter   :
'Date               : -
'LastUpdate/By      : - Tedi
'*****************************************

   With picOBJ(Index) 'FIRST.
        
        If .Enabled = True Then
            .Picture = LoadPicture(SetButtonImageEffect(Index, bieNormal))
        Else: .Picture = LoadPicture(SetButtonImageEffect(Index, bieDisabled))
        End If
        
    End With
    
End Sub

Sub picButton_Obj(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single, picOBJ) 'TOOLBAR_AI.
'************************************************
' Procedure         : picButton_MouseMove
' Function          : TOOLBAR_AI saat mouse berada di area button.
' Created By        : {73 64 6B}
' Date              : 12-Apr-2015
' addition          : Penambahan picOBJ
'************************************************

    If (X < 0) Or (Y < 0) Or (X > picOBJ(Index).Width) Or (Y > picOBJ(Index).Height) Then 'Dua IF ini jangan diubah keluar CASE agar API-nya jalan.
        ReleaseCapture 'The MOUSE_LEAVE pseudo-event.
        picOBJ(Index).Picture = LoadPicture(SetButtonImageEffect(Index, bieNormal)) 'Back to NORMAL.

    ElseIf GetCapture() <> picOBJ(Index).hwnd Then
        SetCapture picOBJ(Index).hwnd 'The MOUSE_ENTER pseudo-event.
        picOBJ(Index).Picture = LoadPicture(SetButtonImageEffect(Index, bieOver)) 'Set to OVER_EFFECT.
    End If
    
End Sub

Sub EnableObject(ByVal paIsEnable As Boolean)
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

Private Sub picButton_Click(Index As Integer)
'************************************************
' Procedure         : picButton_Click
' Function          : Action utk Navigation dan CRUD.
' Created By        : {73 64 6B}
' Date              : 12-Apr-2015/{73 64 6B} --> Semua coding dan query sudah di optimalkan agar faster, readable, safer, standardable.
'************************************************
    Dim strCode As String, strFileRpt As String
    
    Select Case Index
            
        Case enButtonType.bieAdd  '4 'ADD.
            Call db_add
            
        Case enButtonType.bieEdit  '5 'EDIT.
            Call db_edit
            
        Case enButtonType.bieDelete  '6 'DELETE.
            Call db_delete
            
        Case enButtonType.bieClose  '23 'CLOSE.
            Call db_close
            
        Case enButtonType.bieSave  'SAVE.
            Call db_save
            
        Case enButtonType.bieCancel 'CANCEL.
            Call db_cancel
    End Select

End Sub
