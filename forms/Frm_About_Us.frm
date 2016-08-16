VERSION 5.00
Begin VB.Form Frm_About_Us 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "About"
   ClientHeight    =   3555
   ClientLeft      =   2760
   ClientTop       =   3750
   ClientWidth     =   4410
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3555
   ScaleWidth      =   4410
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton OKButton 
      Caption         =   "OK"
      Height          =   375
      Left            =   1728
      TabIndex        =   0
      Top             =   3024
      Width           =   1215
   End
End
Attribute VB_Name = "Frm_About_Us"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
''************************************************
' Form              : Frm_About_Us
' Function          : list Programmer
' Created Date      : 6 Jan 2003
' By                : yayan
' Last Update/ BY   :
'************************************************
Option Explicit

Private Sub Form_Load()
'
End Sub

Private Sub OKButton_Click()
    Unload Me
End Sub
