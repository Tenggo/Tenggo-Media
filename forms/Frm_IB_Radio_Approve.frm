VERSION 5.00
Begin VB.Form Frm_IB_Radio_Approve 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Approve"
   ClientHeight    =   2265
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3420
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2265
   ScaleWidth      =   3420
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Cmd_Cancel 
      Caption         =   "&Cancel"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1815
      TabIndex        =   4
      Top             =   1695
      Width           =   975
   End
   Begin VB.CommandButton Cmd_App 
      Caption         =   "&Approve"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   675
      TabIndex        =   3
      Top             =   1695
      Width           =   975
   End
   Begin VB.Frame Frame1 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   533
      TabIndex        =   0
      Top             =   510
      Width           =   2355
      Begin VB.TextBox Txt_Password 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         IMEMode         =   3  'DISABLE
         Left            =   180
         MaxLength       =   15
         PasswordChar    =   "*"
         TabIndex        =   2
         Top             =   345
         Width           =   1980
      End
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Enter Password"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   765
      TabIndex        =   1
      Top             =   150
      Width           =   1935
   End
End
Attribute VB_Name = "Frm_IB_Radio_Approve"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
''************************************************
' Form              : Frm_IB_Radio_Approve
' Function          : untu meng - approve IB Radio
' Created Date      : 5/Feb/2001
' By                : joko
' Last Update       :'
'************************************************

Const Not_Approved = "NOT APPROVED"
Const Approved = "APPROVED"
Public What_Approval As Form

Private Sub Cmd_App_Click()
Dim rs As New ADODB.Recordset
Dim TxtSQl As String

If Txt_Password.Text = "" Then
    MsgBox "Please Insert the Password!", vbCritical, strCompany_Name
    Exit Sub
End If

TxtSQl = "select * from User_id where User_Name = '" & strLogin_User & "' and Password ='" & Encrypt(Txt_Password.Text) & "'"
rs.Open TxtSQl, Conn, adOpenStatic, adLockReadOnly
With rs
    If .EOF = False Then
        What_Approval.Lbl_APP.ForeColor = vbBlack
        What_Approval.Lbl_APP.Caption = Approved
        rs_Date.Requery
        What_Approval.Lbl_Date.Caption = Format(rs_Date(0), "dd/mm/yyyy hh:mm:ss AMPM")
        'What_Approval.Lbl_Time.Caption = Format(Time, "hh:mm:ss")
        MsgBox "Aprroved", vbInformation, StrCompany
    Else
        MsgBox "Wrong Password", vbCritical, StrCompany
    End If
    If .State = adStateOpen Then
        .Close
    End If
End With
Set rs = Nothing
Unload Me
End Sub

Private Sub cmd_cancel_Click()
Unload Me
End Sub

Private Sub Form_Load()
    Frame1.Caption = "HI, " & User
End Sub

Private Sub txt_Password_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    Call Cmd_App_Click
End If
End Sub
