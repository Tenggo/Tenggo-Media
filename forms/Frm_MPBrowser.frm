VERSION 5.00
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "ieframe.dll"
Begin VB.Form Frm_MPBrowser 
   Caption         =   "Form1"
   ClientHeight    =   6420
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8760
   LinkTopic       =   "Form1"
   ScaleHeight     =   6420
   ScaleWidth      =   8760
   StartUpPosition =   3  'Windows Default
   Begin SHDocVwCtl.WebBrowser WebBrowser1 
      Height          =   3075
      Left            =   390
      TabIndex        =   0
      Top             =   105
      Width           =   4725
      ExtentX         =   8334
      ExtentY         =   5424
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
Attribute VB_Name = "Frm_MPBrowser"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'
Private Sub Form_Activate()
    WebBrowser1.Left = 0
    WebBrowser1.Top = 0
    WebBrowser1.FullScreen = False
    
    WebBrowser1.Navigate App.Path & "\BCTMP.HTM"
    SleepX
End Sub

Private Sub Form_Resize()
    WebBrowser1.Height = Me.Height
    WebBrowser1.Width = Me.Width
End Sub


Private Sub Form_Unload(Cancel As Integer)
    WebBrowser1.Stop
End Sub
