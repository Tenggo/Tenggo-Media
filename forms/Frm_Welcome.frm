VERSION 5.00
Begin VB.Form Frm_Welcome 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   3135
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6075
   Icon            =   "Frm_Welcome.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3135
   ScaleWidth      =   6075
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Height          =   2925
      Left            =   180
      TabIndex        =   0
      Top             =   75
      Width           =   5775
      Begin VB.Frame Frame2 
         Caption         =   "Hi "
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1485
         Left            =   195
         TabIndex        =   2
         Top             =   735
         Width           =   5370
         Begin VB.Label Label2 
            Alignment       =   1  'Right Justify
            Caption         =   "Last Login : "
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H8000000D&
            Height          =   255
            Left            =   315
            TabIndex        =   6
            Top             =   480
            Width           =   1125
         End
         Begin VB.Label Label4 
            Alignment       =   1  'Right Justify
            Caption         =   "Last Logout : "
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H8000000D&
            Height          =   255
            Left            =   345
            TabIndex        =   5
            Top             =   840
            Width           =   1125
         End
         Begin VB.Label lbllastlogin 
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   255
            Left            =   1455
            TabIndex        =   4
            Top             =   480
            Width           =   3570
         End
         Begin VB.Label lbllastlogout 
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   255
            Left            =   1485
            TabIndex        =   3
            Top             =   840
            Width           =   3600
         End
      End
      Begin VB.CommandButton cmdOK 
         Caption         =   "&OK"
         Default         =   -1  'True
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   425
         Left            =   2205
         TabIndex        =   1
         Top             =   2355
         Width           =   1070
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Caption         =   "Welcome"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   300
         Left            =   555
         TabIndex        =   7
         Top             =   315
         Width           =   4590
      End
   End
End
Attribute VB_Name = "Frm_Welcome"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
''*************************************************************
'Nama Form          : Frm_Welcome
'Fungsi Form        : Menampikan nama user & history user login
'Programer          : Yayan Royani
'Tgl Pembuatan      : 08/Jan/02
'Last Update/By     : 08/Jan/02 / Yayan Royani
'*************************************************************

Option Explicit

Private Sub cmdOK_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    'CenterForm Me
    Frame2.Caption = "Hi " & UCase(strComputer_Name)
    lbllastlogin.Caption = strLastLogin
    lbllastlogout.Caption = strLastLogout
        
End Sub

Private Sub Form_Unload(Cancel As Integer)
    'Load mdi_Main
    mdi_Main.Show
    
End Sub








