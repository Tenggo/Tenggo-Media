VERSION 5.00
Begin VB.Form Frm_IB_Radio_Approve 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Approve"
   ClientHeight    =   1845
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7155
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1845
   ScaleWidth      =   7155
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
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
      Height          =   1380
      Left            =   120
      TabIndex        =   0
      Top             =   150
      Width           =   6810
      Begin VB.CommandButton Cmd_App 
         Caption         =   "&Approve"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   4320
         TabIndex        =   4
         Top             =   810
         Width           =   975
      End
      Begin VB.CommandButton Cmd_Cancel 
         Caption         =   "&Cancel"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   5460
         TabIndex        =   3
         Top             =   810
         Width           =   975
      End
      Begin VB.TextBox Txt_Password 
         Appearance      =   0  'Flat
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
         IMEMode         =   3  'DISABLE
         Left            =   1590
         PasswordChar    =   "X"
         TabIndex        =   1
         Top             =   285
         Width           =   4875
      End
      Begin VB.Label Label2 
         Caption         =   "Enter password "
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
         Height          =   225
         Left            =   270
         TabIndex        =   2
         Top             =   315
         Width           =   1575
      End
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
rs.Open TxtSQl, ConnERP, adOpenStatic, adLockReadOnly
With rs
    If .EOF = False Then
        What_Approval.Lbl_APP.ForeColor = vbBlack
        What_Approval.Lbl_APP.Caption = Approved
        recDate.Requery
        What_Approval.Lbl_Date.Caption = Format(recDate(0), "dd/mm/yyyy hh:mm:ss AMPM")
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
