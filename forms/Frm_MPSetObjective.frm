VERSION 5.00
Begin VB.Form Frm_MPSetObjective 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Set Objective"
   ClientHeight    =   615
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5520
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   615
   ScaleWidth      =   5520
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmd_cancel 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   4500
      TabIndex        =   3
      Top             =   105
      Width           =   930
   End
   Begin VB.CommandButton Cmd_set_objective 
      Caption         =   "Set"
      Height          =   375
      Left            =   3510
      TabIndex        =   2
      Top             =   105
      Width           =   930
   End
   Begin VB.ComboBox cbo_mp_tv_rv_id 
      Height          =   315
      Left            =   1410
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   135
      Width           =   2010
   End
   Begin VB.Label Label1 
      Caption         =   "Objective ID : "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Left            =   135
      TabIndex        =   1
      Top             =   180
      Width           =   1350
   End
End
Attribute VB_Name = "Frm_MPSetObjective"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim strMPPlanDimID As String
Dim intWeekYear As Integer
Dim strSql As String
'
Private Sub cmd_cancel_Click()
    Unload Me
End Sub

Private Sub Cmd_set_objective_Click()
    If cbo_mp_tv_rv_id.Text <> "" Then
        strSql = "update mp_insertion set mp_tv_rf_id = " & cbo_mp_tv_rv_id.Text & " where mp_plan_dim_id = '" & strMPPlanDimID & "' and week_year = " & intWeekYear
        ConnERP.Execute strSql
        MsgBox "Objective Set!", vbExclamation, strApplication_Name
        Unload Me
    Else
        MsgBox "Choose Objective ID!", vbExclamation, strApplication_Name
    End If
End Sub


Private Sub Form_Load()
    Dim rsTemp As New ADODB.Recordset
    With frm_MPInsertion.msf_MPInsertion
        strMPPlanDimID = Left(.TextMatrix(.Row, .cols - 1), 19)
        intWeekYear = Trim(.TextMatrix(3, .col))
        strSql = "select mp_tv_rf_id from mp_tv_reach_frequency "
        strSql = strSql & " where mp_plan_dim_id in "
        strSql = strSql & "(select mp_plan_dim_id from mp_ids "
        strSql = strSql & "where mp_plan_dim_id is not null "
        strSql = strSql & "and mp_medium_id = (select mp_medium_id from mp_ids "
        strSql = strSql & "where mp_plan_dim_id = '" & strMPPlanDimID & "')) "
        strSql = strSql & "and week_year_start<= " & intWeekYear & " and week_year_end >= " & intWeekYear & " order by mp_tv_rf_id"
        rsTemp.Open strSql, ConnERP, 1, 3
        While Not rsTemp.EOF
            cbo_mp_tv_rv_id.AddItem rsTemp(0)
            rsTemp.MoveNext
        Wend
        rsTemp.Close
    End With
End Sub


