VERSION 5.00
Begin VB.Form Frm_MPEnterPassword 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Enter Password"
   ClientHeight    =   1545
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6960
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1545
   ScaleWidth      =   6960
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Height          =   1350
      Left            =   135
      TabIndex        =   0
      Top             =   60
      Width           =   6690
      Begin VB.CommandButton cmdOK 
         Caption         =   "OK"
         Height          =   375
         Left            =   4110
         TabIndex        =   6
         Top             =   855
         Width           =   1140
      End
      Begin VB.TextBox txtOKCancel 
         Appearance      =   0  'Flat
         Height          =   300
         Left            =   1620
         TabIndex        =   5
         Text            =   "Text1"
         Top             =   795
         Visible         =   0   'False
         Width           =   915
      End
      Begin VB.CommandButton cmdCancel 
         Caption         =   "Cancel"
         Height          =   375
         Left            =   5430
         TabIndex        =   4
         Top             =   855
         Width           =   1140
      End
      Begin VB.TextBox txtPassword2 
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
         Left            =   1485
         PasswordChar    =   "X"
         TabIndex        =   3
         Top             =   285
         Width           =   4875
      End
      Begin VB.TextBox txtPassword 
         Appearance      =   0  'Flat
         Height          =   300
         IMEMode         =   3  'DISABLE
         Left            =   330
         TabIndex        =   2
         Text            =   "Text1"
         Top             =   630
         Visible         =   0   'False
         Width           =   795
      End
      Begin VB.Label Label6 
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
         Left            =   165
         TabIndex        =   1
         Top             =   315
         Width           =   1575
      End
   End
End
Attribute VB_Name = "Frm_MPEnterPassword"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'
Private Sub cmdCancel_Click()
    txtOKCancel.Text = "Cancel"
    Me.Hide
End Sub

Private Sub cmdOK_Click()
    If isAuthenticated Then
        txtOKCancel.Text = "OK"
        Me.Hide
        frm_MPInsertion.Refresh
    Else
        MsgBox "Invalid Password!", vbExclamation, strApplication_Name
        txtPassword2.Text = ""
        txtPassword.Text = ""
        txtPassword2.SetFocus
    End If
End Sub

Private Sub Form_Load()
    txtPassword2.Text = ""
    txtPassword.Text = ""
End Sub

Private Sub txtPassword2_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
        Case 13:
            Call cmdOK_Click
        Case 8:
        If txtPassword2.Text <> "" Then
            'txtPassword2.Text = Mid(txtPassword2.Text, 1, Len(txtPassword2.Text) - 3)
            'txtPassword.Text = Mid(txtPassword.Text, 1, Len(txtPassword.Text) - 1)
        End If
    Case Else
        txtPassword.Text = txtPassword.Text & Chr(KeyAscii)
        'txtPassword2.Text = txtPassword2.Text & String(3, Chr(KeyAscii))
    End Select
    'txtPassword2.SelStart = Len(txtPassword2.Text)
End Sub

Private Function isAuthenticated(Optional pvUserName As String) As Boolean
    Dim rsPassword As New ADODB.Recordset
    isAuthenticated = False
    If pvUserName = "" Then pvUserName = strLogin_User
    rsPassword.Open "select password from user_id where user_name = '" & Clear_String(pvUserName) & "'", ConnERP, 1, 3
    
    If txtPassword.Text = Decrypt(rsPassword(0)) Then
        isAuthenticated = True
    End If
        
    rsPassword.Close
    Set rsPassword = Nothing
End Function
