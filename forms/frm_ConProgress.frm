VERSION 5.00
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "ieframe.dll"
Begin VB.Form frm_ConProgress 
   ClientHeight    =   990
   ClientLeft      =   120
   ClientTop       =   120
   ClientWidth     =   4005
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "frm_Query_Progress"
   ScaleHeight     =   990
   ScaleWidth      =   4005
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture1 
      BorderStyle     =   0  'None
      Height          =   315
      Left            =   495
      ScaleHeight     =   315
      ScaleWidth      =   3030
      TabIndex        =   0
      Top             =   420
      Width           =   3030
      Begin SHDocVwCtl.WebBrowser WebBrowser1 
         Height          =   2550
         Left            =   -210
         TabIndex        =   1
         Top             =   -1605
         Width           =   3900
         ExtentX         =   6879
         ExtentY         =   4498
         ViewMode        =   0
         Offline         =   0
         Silent          =   0
         RegisterAsBrowser=   0
         RegisterAsDropTarget=   1
         AutoArrange     =   0   'False
         NoClientEdge    =   0   'False
         AlignLeft       =   0   'False
         NoWebView       =   0   'False
         HideFileNames   =   0   'False
         SingleClick     =   0   'False
         SingleSelection =   0   'False
         NoFolders       =   0   'False
         Transparent     =   0   'False
         ViewID          =   "{0057D0E0-3573-11CF-AE69-08002B2E1262}"
         Location        =   "http:///"
      End
   End
   Begin VB.Timer Timer1 
      Interval        =   500
      Left            =   570
      Top             =   270
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Please Wait  ..."
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000011&
      Height          =   270
      Left            =   1140
      TabIndex        =   2
      Top             =   60
      Width           =   1680
   End
End
Attribute VB_Name = "frm_ConProgress"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'<CSCC>
'********************************************************************************
'Submodul Name      : frm_ConProgress
'Submodul Function  : -
'Used Table         : -
'Used SP            : -
'Procedure/Function : -
'Programmer Name    : Tedi
'Date               : 2/25/2016-4:09:42 PM
'Last Update By     : Tedi
'Date Update        : 2/25/2016-4:09:42 PM
'Log Update/By      : -
'********************************************************************************
'</CSCC>
Option Explicit

Private Sub Form_Load()
'<CSCM>
'********************************************************************************
'Submodul Name      : frm_ConProgress
'Submodul Function  : -
'Used Table         : -
'Used SP            : -
'Procedure/Function : -
'Programmer Name    : - Tedi
'Date               : -
'Last Update/By     : -
'Date Update        : -
'Log Update/By      : -
'********************************************************************************
'</CSCM>
    
    WebBrowser1.Navigate App.Path & "\resources\progress.gif"
     
    Timer1.Enabled = True
    Timer1.Interval = 50

End Sub

Private Sub Timer1_Timer()
'<CSCM>
'********************************************************************************
'Submodul Name      : frm_ConProgress
'Submodul Function  : -
'Used Table         : -
'Used SP            : -
'Procedure/Function : -
'Programmer Name    : - Tedi
'Date               : -
'Last Update/By     : -
'Date Update        : -
'Log Update/By      : -
'********************************************************************************
'</CSCM>
    
    If Me.Visible = True Then
    WebBrowser1.SetFocus
    End If

End Sub


