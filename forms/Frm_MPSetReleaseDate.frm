VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form Frm_MPSetReleaseDate 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Set Release Date"
   ClientHeight    =   555
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4095
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   555
   ScaleWidth      =   4095
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdSet 
      Caption         =   "&Set"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   1965
      TabIndex        =   2
      ToolTipText     =   "create plan"
      Top             =   75
      Width           =   1005
   End
   Begin VB.CommandButton cmdCancel 
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
      Height          =   360
      Left            =   3000
      TabIndex        =   1
      ToolTipText     =   "cancel and return to the main window"
      Top             =   75
      Width           =   1005
   End
   Begin MSComCtl2.DTPicker DTReleaseDate 
      Height          =   300
      Left            =   105
      TabIndex        =   0
      Top             =   105
      Width           =   1725
      _ExtentX        =   3043
      _ExtentY        =   529
      _Version        =   393216
      Format          =   96927745
      CurrentDate     =   38272
   End
End
Attribute VB_Name = "Frm_MPSetReleaseDate"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'
Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdSet_Click()
    ConnERP.Execute "update mp_master set release_date = '" & DTReleaseDate.Value & "' where mp_number = '" & FrmMPInsertion.cboMPNumber.Text & "'"
    FrmMPInsertion.lblReleaseDate.Caption = Format(DTReleaseDate.Value, "dd mmm yyyy")
    FrmMPInsertion.lblReleaseDate.ForeColor = vbBlue
    Unload Me
End Sub

Private Sub Form_Load()
    Dim rsTemp As New ADODB.Recordset
    rsTemp.Open "select release_date from mp_master where mp_number = '" & FrmMPInsertion.cboMPNumber.Text & "'", ConnERP, 1, 3
    If Not rsTemp.EOF Then
        If IsNull(rsTemp(0)) Then
            DTReleaseDate.Value = Now()
        Else
            DTReleaseDate.Value = rsTemp(0)
        End If
    End If
    rsTemp.Close
    Set rsTemp = Nothing
End Sub
