VERSION 5.00
Begin VB.Form Frm_MPUnlockSecurity 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Enter Supervisor Password"
   ClientHeight    =   3090
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   4875
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3090
   ScaleWidth      =   4875
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.TextBox txtpassword 
      Height          =   300
      IMEMode         =   3  'DISABLE
      Left            =   4230
      PasswordChar    =   "*"
      TabIndex        =   6
      Text            =   "Text1"
      Top             =   2640
      Visible         =   0   'False
      Width           =   285
   End
   Begin VB.TextBox txtUserID 
      Enabled         =   0   'False
      Height          =   300
      Left            =   3870
      TabIndex        =   5
      Top             =   2715
      Visible         =   0   'False
      Width           =   195
   End
   Begin VB.CommandButton cmdCAncel 
      Caption         =   "Cancel"
      Height          =   450
      Left            =   2400
      TabIndex        =   3
      Top             =   2535
      Width           =   1125
   End
   Begin VB.Frame Frame1 
      Height          =   2370
      Left            =   90
      TabIndex        =   7
      Top             =   15
      Width           =   4695
      Begin VB.TextBox txtNote 
         Height          =   1020
         Left            =   1125
         MultiLine       =   -1  'True
         TabIndex        =   1
         Top             =   1155
         Width           =   3420
      End
      Begin VB.TextBox txtSpvName 
         Height          =   300
         Left            =   1125
         Locked          =   -1  'True
         TabIndex        =   4
         Top             =   270
         Width           =   3390
      End
      Begin VB.TextBox txtPassword2 
         Height          =   300
         IMEMode         =   3  'DISABLE
         Left            =   1125
         PasswordChar    =   "X"
         TabIndex        =   0
         Top             =   705
         Width           =   3405
      End
      Begin VB.Label Label3 
         Caption         =   "Note :"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   90
         TabIndex        =   10
         Top             =   1155
         Width           =   960
      End
      Begin VB.Label Label1 
         Caption         =   "Spv. Name :"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   75
         TabIndex        =   9
         Top             =   285
         Width           =   1155
      End
      Begin VB.Label Label2 
         Caption         =   "Password :"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   90
         TabIndex        =   8
         Top             =   705
         Width           =   960
      End
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Height          =   450
      Left            =   1245
      TabIndex        =   2
      Top             =   2550
      Width           =   1125
   End
End
Attribute VB_Name = "Frm_MPUnlockSecurity"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub cmdOK_Click()
    If isAuthenticated(txtUserID.Text) Then
        Call Unlock_Cell
        Unload Me
    Else
        MsgBox "Invalid Password", vbExclamation, strApplication_Name
        txtPassword2.SetFocus
    End If
End Sub
'
Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    Dim strUserID As String
    Dim rsTemp As New ADODB.Recordset
    'get supervisor ID
    rsTemp.Open "select isnull(report_to,user_name) from user_id where user_name = '" & strLogin_User & "'", ConnERP, 1, 3
        strUserID = rsTemp(0)
    rsTemp.Close
    
    rsTemp.Open "select name from user_id where user_name = '" & strUserID & "'", ConnERP, 1, 3
    If Not rsTemp.EOF Then
        txtSpvName.Text = rsTemp(0)
        txtUserID.Text = strUserID
    Else
        txtUserID.Text = strLogin_User
        txtSpvName = Empty
        strUserID = strLogin_User
    End If
    rsTemp.Close
    
    If txtSpvName = Empty Then
        rsTemp.Open "select name from user_id where user_name = '" & strUserID & "'", ConnERP, 1, 3
            txtSpvName.Text = rsTemp(0)
        rsTemp.Close
    End If
    
    txtPassword.Text = ""
    
End Sub

Private Sub txtPassword2_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
            Call cmdOK_Click
    End If
End Sub

Private Function isAuthenticated(Optional pvUserName As String) As Boolean
    Dim rsPassword As New ADODB.Recordset
    isAuthenticated = False
    If pvUserName = "" Then pvUserName = strLogin_User
    rsPassword.Open "select password from user_id where user_name = '" & Clear_String(pvUserName) & "'", ConnERP, 1, 3
    
    If txtPassword2.Text = Decrypt(rsPassword(0)) Then
        isAuthenticated = True
    End If
        
    rsPassword.Close
    Set rsPassword = Nothing
End Function

Private Sub Unlock_Cell()
    Dim rsTemp As New ADODB.Recordset
    
    Dim str_Unlock_Req_By_current As String
    Dim str_MP_Plan_Dim_Id As String
    Dim int_Week_Year As Integer
    Dim dt_Week_Commencing As Date
    Dim str_Unlock_Req_By As String
    Dim str_Unlock_By As String
    Dim dt_Unlock_Date As Date
    Dim str_Note As String
    Dim strSql As String
    With frm_MPInsertion.msf_MPInsertion
        str_MP_Plan_Dim_Id = Left(.TextMatrix(.Row, .cols - 1), 19)
        int_Week_Year = CInt(.TextMatrix(3, .col))
        dt_Week_Commencing = CDate(.TextMatrix(4, .col))
        str_Unlock_Req_By = strLogin_User
        str_Unlock_By = txtUserID.Text
        recDate.Requery
        dt_Unlock_Date = CDate(recDate(0))
        str_Note = Clear_Enter(Clear_String(txtNote.Text))
    End With
    str_Unlock_Req_By_current = ""
    strSql = "select unlock_req_by from mp_unlocked_week where mp_plan_dim_id='" & str_MP_Plan_Dim_Id & "' and week_year=" & int_Week_Year
    rsTemp.Open strSql, ConnERP, 1, 3
    If Not rsTemp.EOF Then
        str_Unlock_Req_By_current = rsTemp(0)
    End If
    rsTemp.Close
    If str_Unlock_Req_By_current <> "" Then
        MsgBox "This cell already unlocked by " & str_Unlock_Req_By_current, vbExclamation, strApplication_Name
    Else
        strSql = "insert into mp_unlocked_week(mp_plan_dim_id,week_year,week_commencing,unlock_req_by,unlock_by,unlock_date,note) "
        strSql = strSql & "values('" & str_MP_Plan_Dim_Id & "'," & int_Week_Year & ",'" & dt_Week_Commencing & "','" & str_Unlock_Req_By & "','" & str_Unlock_By & "','" & dt_Unlock_Date & "','" & str_Note & "')"
        ConnERP.Execute strSql
        MsgBox "Cell Unlocked!", vbExclamation, strApplication_Name
    End If
    With frm_MPInsertion.msf_MPInsertion
        Select Case UCase(Right(.TextMatrix(.Row, .cols - 1), 4))
            Case "EACH"
                .CellBackColor = frm_MPInsertion.Shape_unlocked.FillColor
                .Row = .Row - 1
                .CellBackColor = frm_MPInsertion.Shape_unlocked.FillColor
                .Row = .Row + 2
                .CellBackColor = frm_MPInsertion.Shape_unlocked.FillColor
                .Row = .Row - 1
            Case "FREQ"
                .CellBackColor = frm_MPInsertion.Shape_unlocked.FillColor
                .Row = .Row + 1
                .CellBackColor = frm_MPInsertion.Shape_unlocked.FillColor
                .Row = .Row + 1
                .CellBackColor = frm_MPInsertion.Shape_unlocked.FillColor
                .Row = .Row - 2
            Case Else
            .CellBackColor = frm_MPInsertion.Shape_unlocked.FillColor
            If UCase(Right(.TextMatrix(.Row - 1, .cols - 1), 5)) = "REACH" Then
                .Row = .Row - 1
                .CellBackColor = frm_MPInsertion.Shape_unlocked.FillColor
                .Row = .Row - 1
                .CellBackColor = frm_MPInsertion.Shape_unlocked.FillColor
                .Row = .Row + 2
            End If
        End Select
    End With
    
End Sub

