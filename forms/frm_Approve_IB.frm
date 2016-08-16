VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Begin VB.Form frm_Approve_IB 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Approve"
   ClientHeight    =   2250
   ClientLeft      =   2760
   ClientTop       =   3750
   ClientWidth     =   3690
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
   ScaleHeight     =   2250
   ScaleWidth      =   3690
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
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
      ScaleWidth      =   3690
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   0
      Width           =   3690
      Begin VB.PictureBox picButton 
         Appearance      =   0  'Flat
         BackColor       =   &H00F0F0F0&
         BorderStyle     =   0  'None
         ClipControls    =   0   'False
         ForeColor       =   &H80000008&
         Height          =   750
         Index           =   17
         Left            =   90
         Picture         =   "frm_Approve_IB.frx":0000
         ScaleHeight     =   750
         ScaleWidth      =   1500
         TabIndex        =   2
         TabStop         =   0   'False
         Top             =   0
         Width           =   1500
      End
   End
   Begin Threed.SSPanel pnlMain 
      Align           =   1  'Align Top
      Height          =   1440
      Left            =   0
      TabIndex        =   0
      Top             =   750
      Width           =   3690
      _Version        =   65536
      _ExtentX        =   6509
      _ExtentY        =   2540
      _StockProps     =   15
      BackColor       =   15790320
      BevelOuter      =   1
      Begin VB.Frame Frame1 
         Height          =   1005
         Left            =   135
         TabIndex        =   3
         Top             =   255
         Width           =   3450
         Begin VB.TextBox Txt_Password 
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
            ForeColor       =   &H000000FF&
            Height          =   315
            IMEMode         =   3  'DISABLE
            Left            =   1380
            MaxLength       =   15
            PasswordChar    =   "*"
            TabIndex        =   4
            Text            =   "a"
            Top             =   555
            Width           =   1935
         End
         Begin VB.Label Lbl_User 
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "a"
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
            Height          =   315
            Left            =   1380
            TabIndex        =   7
            Top             =   195
            Width           =   1935
         End
         Begin VB.Label Lbl_User_Name 
            Caption         =   "User Name "
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
            Height          =   315
            Left            =   135
            TabIndex        =   6
            Top             =   240
            Width           =   1215
         End
         Begin VB.Label Lbl_password 
            Caption         =   "Password "
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
            Height          =   315
            Left            =   135
            TabIndex        =   5
            Top             =   585
            Width           =   1215
         End
      End
      Begin VB.Label Lbl_Enter_Password 
         Alignment       =   2  'Center
         Caption         =   "Enter Your Password....!"
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
         Height          =   300
         Left            =   180
         TabIndex        =   8
         Top             =   60
         Width           =   3255
      End
   End
End
Attribute VB_Name = "frm_Approve_IB"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'
Option Explicit
''
Private Sub Form_Load()
'************************************************
' Procedure         : Form_Load
' Function          : Form Load
' Update  By        : Tedi/Kreatif
' Date              : 15-07-2016
'************************************************
    
    Lbl_User.Caption = UCase(strLogin_User)
    Txt_Password.Text = ""
    
    Lbl_User_Name.Visible = True
    Lbl_User.Visible = True
    Lbl_password.Visible = True
    Txt_Password.Visible = True
    Lbl_Enter_Password.Visible = True
    
    blnStatusPassword = False

End Sub

Private Sub db_Cancel()
'************************************************
' Procedure         : db_Cancel
' Function          : Pembatalan Proses
' Update  By        : Tedi/Kreatif
' Date              : 15-07-2016
' Description       : Before cmd_cancel_Click
'************************************************
    blnStatusPassword = False
    Unload Me
    
End Sub

Private Sub db_OK()
'************************************************
' Procedure         : picButton_Click
' Function          : Proses
' Update  By        : Tedi/Kreatif
' Date              : 15-07-2016
' Description       : Cmd_Ok_Click
'************************************************

    Dim recPasswordUser As New ADODB.Recordset
    
    recPasswordUser.Open "Select Password from User_Id where User_name='" & strLogin_User & "'", ConnERP, adOpenForwardOnly, adLockReadOnly, adCmdText
    
    If Not recPasswordUser.BOF And Not recPasswordUser.EOF Then
        If Txt_Password.Text = Decrypt(Trim(recPasswordUser("Password"))) Then
            blnStatusPassword = True
            
            MsgBox strMsgApproved, vbInformation, strTitleInfo
            
            recDate.Requery
        Else
            blnStatusPassword = False
            
            MsgBox strMsgInvalidPassword, vbCritical, strTitleInfo
            
            'Txt_Password.SetFocus
            
            'SendKeys "{HOME}+{End}"
        End If
    End If
    
    recPasswordUser.Close
    Set recPasswordUser = Nothing
    
    Unload Me
End Sub

Private Sub Form_Activate()
    Me.Txt_Password.SetFocus
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
        Case enButtonType.bieOK '4 'ADD.
            Call db_OK
        Case enButtonType.biecancel  '5 'EDIT.
            Call db_Cancel
    End Select

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

Private Sub picButton_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single) 'TOOLBAR_AI.
'************************************************
' Procedure         : picButton_MouseDown
' Function          : TOOLBAR_AI saat mouse ditekan.
' Created By        : {73 64 6B}
' Date              : 12-Apr-2015
'************************************************
    
    picButton(Index).Picture = LoadPicture(SetButtonImageEffect(Index, bieDown)) 'FIRST.

End Sub
