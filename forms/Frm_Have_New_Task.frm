VERSION 5.00
Begin VB.Form Frm_Have_New_Task 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   1305
   ClientLeft      =   255
   ClientTop       =   1410
   ClientWidth     =   3570
   ClipControls    =   0   'False
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
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1305
   ScaleWidth      =   3570
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame1 
      Height          =   1350
      Left            =   0
      TabIndex        =   0
      Top             =   -75
      Width           =   3525
      Begin VB.CommandButton Cmd_Close 
         Caption         =   "&Close"
         Height          =   390
         Left            =   1245
         TabIndex        =   1
         Top             =   810
         Width           =   1080
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Caption         =   "Please Select Cek My Job On My Menu "
         Height          =   375
         Left            =   150
         TabIndex        =   3
         Top             =   480
         Width           =   3225
      End
      Begin VB.Label lbl_Message 
         Alignment       =   2  'Center
         Caption         =   "You Have New Task !"
         Height          =   315
         Left            =   465
         TabIndex        =   2
         Top             =   210
         Width           =   2130
      End
   End
End
Attribute VB_Name = "Frm_Have_New_Task"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
''
Option Explicit

Private Sub Cmd_Close_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    Dim strSql As String
    Dim rsCount_Task As New ADODB.Recordset
    
    strSql = "SELECT Count(*) FROM Current_User_Job WHERE User_Name='" & strLogin_User & "'"
    strSql = strSql & " AND New_Message=1"
    
    rsCount_Task.Open strSql, ConnERP, adOpenStatic, adLockReadOnly
    
    lbl_Message.Caption = "You Have " & rsCount_Task.Fields(0).Value & " New Task"
    
    rsCount_Task.Close
    Set rsCount_Task = Nothing
    
'    SetWindowPos Me.hWnd, -1, 0, 0, 0, 0, &H1 Or &H2
End Sub



