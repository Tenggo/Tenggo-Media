VERSION 5.00
Begin VB.Form Frm_Login 
   BorderStyle     =   0  'None
   Caption         =   "Security"
   ClientHeight    =   8820
   ClientLeft      =   0
   ClientTop       =   4515
   ClientWidth     =   12345
   Icon            =   "frm_form.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8820
   ScaleWidth      =   12345
   ShowInTaskbar   =   0   'False
   WindowState     =   2  'Maximized
   Begin VB.PictureBox picRightBar 
      Align           =   4  'Align Right
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   8820
      Left            =   7065
      ScaleHeight     =   8820
      ScaleWidth      =   5280
      TabIndex        =   1
      Top             =   0
      Width           =   5280
      Begin VB.PictureBox picCmd 
         BackColor       =   &H00F39A21&
         BorderStyle     =   0  'None
         Height          =   495
         Index           =   1
         Left            =   690
         ScaleHeight     =   495
         ScaleWidth      =   4215
         TabIndex        =   9
         Top             =   4980
         Width           =   4215
         Begin VB.Label lblCaption 
            BackStyle       =   0  'Transparent
            Caption         =   "Exit"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   270
            Index           =   1
            Left            =   1995
            TabIndex        =   10
            Top             =   120
            Width           =   300
         End
      End
      Begin VB.PictureBox picCmd 
         BackColor       =   &H00F39A21&
         BorderStyle     =   0  'None
         Height          =   510
         Index           =   0
         Left            =   690
         ScaleHeight     =   510
         ScaleWidth      =   4215
         TabIndex        =   7
         Top             =   4305
         Width           =   4215
         Begin VB.Label lblCaption 
            BackStyle       =   0  'Transparent
            Caption         =   "Login"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   270
            Index           =   0
            Left            =   1935
            TabIndex        =   8
            Top             =   150
            Width           =   420
         End
      End
      Begin VB.TextBox txtUserName 
         Height          =   375
         Left            =   2415
         TabIndex        =   6
         Text            =   "bramantd"
         Top             =   2910
         Width           =   2475
      End
      Begin VB.TextBox txtPassword 
         Height          =   375
         IMEMode         =   3  'DISABLE
         Left            =   2415
         PasswordChar    =   "*"
         TabIndex        =   5
         Text            =   "admin123"
         Top             =   3480
         Width           =   2475
      End
      Begin VB.PictureBox picLogoBG 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         Height          =   2505
         Left            =   0
         ScaleHeight     =   2505
         ScaleWidth      =   5235
         TabIndex        =   3
         Top             =   0
         Width           =   5235
         Begin VB.Image imgLogoCo 
            Height          =   2220
            Left            =   1080
            Picture         =   "frm_form.frx":0442
            Top             =   240
            Width           =   3210
         End
      End
      Begin VB.PictureBox picImageLogo 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         Height          =   1095
         Left            =   -15
         ScaleHeight     =   1095
         ScaleWidth      =   5250
         TabIndex        =   2
         Top             =   5925
         Width           =   5250
         Begin VB.Label lblDesc 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "2015 PT. iLab Komunikasi Indonesia. All rights reserved."
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00404040&
            Height          =   210
            Index           =   0
            Left            =   420
            TabIndex        =   4
            Top             =   810
            Width           =   4485
         End
         Begin VB.Image imgLogo 
            Height          =   600
            Left            =   1815
            Picture         =   "frm_form.frx":70AD
            Top             =   165
            Width           =   1695
         End
      End
      Begin VB.Label lblDesc 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Please sign in to continue"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   240
         Index           =   4
         Left            =   1530
         TabIndex        =   11
         Top             =   2445
         Width           =   2175
      End
      Begin VB.Image Image1 
         Height          =   2970
         Left            =   525
         Picture         =   "frm_form.frx":7A2C
         Top             =   2715
         Width           =   4650
      End
   End
   Begin VB.PictureBox picBackGround 
      Align           =   3  'Align Left
      BackColor       =   &H00F39A21&
      BorderStyle     =   0  'None
      Height          =   8820
      Left            =   0
      ScaleHeight     =   8820
      ScaleWidth      =   5865
      TabIndex        =   0
      Top             =   0
      Width           =   5865
      Begin VB.Image imgBG 
         Height          =   1200
         Left            =   0
         Stretch         =   -1  'True
         Top             =   0
         Width           =   1710
      End
   End
End
Attribute VB_Name = "Frm_Login"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'***********************************************************
'Submodul Name      : Form Login
'Submodul Function  : To login to Application.
'Used Table         : -
'Used SP            : -
'Procedure/Function : -
'Programmer Name    : -
'Date               : 12-Apr-2015
'Last Update/By     : Ted
'Date Update        : -
'Log Update/By      : -
'***************************************************************

Private Sub cmdCancel_Click()
'************************************************
' Procedure         : cmdCancel_Click
' Function          : Menutup Aplikasi
' Input Parameter   : -
' Output Parameter  : -
' Programmer By     : -
' Date Update       : Jan-2016
' Update By         : Tedi/ Kreatif
'************************************************

    EndApplication

End Sub

Private Sub cmdOK_Click()
'************************************************
' Procedure         : cmdOK_Click
' Function          : Cek User dan Pass word, jika valid blnLogin = True dan jika tidak valid blnLogin = False
' Input Parameter   : -
' Output Parameter  : -
' Programmer By     : -
' Date Update       : Jan-2016
' Update By         : Tedi/ Kreatif
'************************************************

    Dim strSQL As String
    Dim recTemplate As New ADODB.Recordset

    strUserName = Trim(Clear_String(txtUserName.Text))

    If strUserName = "" Then
        MsgBox "Please enter User Name ...!", vbExclamation, "Login"
        txtUserName.SetFocus
        Exit Sub
    End If
    
    strSQL = "select User_id.* from User_id WHERE User_id.User_Name='" & strUserName & "'"

    recTemplate.CursorLocation = adUseClient
    recTemplate.Open strSQL, connERP, adOpenStatic, adLockReadOnly

    If recTemplate.EOF Then
        MsgBox "User Name not found.", vbCritical, "Access Denied"
        txtUserName.SetFocus
        SendKeys "{Home}+{End}"
        recTemplate.Close
        Set recTemplate = Nothing
        Exit Sub
    Else
        'Check Disable Status
        If recTemplate.Fields("Disable_Flag").Value = 1 Then
            MsgBox "Your account has beed disabled, Please contact your administrator !", vbCritical, "Access Denied."
            Exit Sub
        End If

        'Check Password
        If recTemplate!Password = Encrypt(txtPassword.Text) Then
            blnLogin = True
            strFullName = recTemplate!Name
            strPositionCode = recTemplate!position_Code
            strDivisionCode = recTemplate!Division_code
            strBUCode = recTemplate!Business_Unit_Code

        Else
            MsgBox "Invalid Password, Please try again...!", vbCritical, "Password Checking"
            txtPassword.SetFocus
            'SendKeys "{Home}+{End}"
            Exit Sub
        End If
        Unload Me

        recTemplate.Close
        Set recTemplate = Nothing
    End If
    
End Sub

Private Sub Form_Activate()
'************************************************
' Procedure         : Form_Activate
' Function          : Adjust Size All Control
' Input Parameter   : -
' Output Parameter  : -
' Programmer By     : -
' Date Update       : Jan-2016
' Update By         : Tedi/ Kreatif
'************************************************

    picBackGround.Width = Me.ScaleWidth - picImageLogo.Width
    picImageLogo.Top = picRightBar.Height - (picImageLogo.Height) - 200
    picLogoBG.Width = picRightBar.Width
    imgLogoCo.Left = (picLogoBG.Width / 2) - (imgLogoCo.Width / 2)
    imgLogoCo.Top = (picLogoBG.Height / 2) - (imgLogoCo.Height / 2)
    imgBG.Width = picBackGround.Width
    imgBG.Height = picBackGround.Height

End Sub

Private Sub Form_Load()
'************************************************
' Procedure         : Form_Load
' Function          : Load Form Login
' Input Parameter   : -
' Output Parameter  : -
' Programmer By     : -
' Date Update       : Jan-2016
' Update By         : Tedi/ Kreatif
'************************************************
    
    'Set Back Ground Picture
    GetPictureFromDB imgBG, OurClient.bg
    lblDesc(0).Caption = Chr(169) & " 2015 PT. Ilab Komunikasi Indonesia. All rights reserved."

End Sub

Private Sub lblCaption_Click(Index As Integer)
'************************************************
' Procedure         : lblCaption_Click
' Function          : Klik label Caption Login atau Cancel
' Input Parameter   : -
' Output Parameter  : -
' Programmer By     : -
' Date Update       : Jan-2016
' Update By         : Tedi/ Kreatif
'************************************************

    Select Case Index
        Case 0
            cmdOK_Click
        Case 1
            cmdCancel_Click
    End Select

End Sub

Private Sub picCmd_Click(Index As Integer)
'************************************************
' Procedure         : picCmd_Click
' Function          : Klik picCmd --> Login atau Cancel
' Input Parameter   : -
' Output Parameter  : -
' Programmer By     : -
' Date Update       : Jan-2016
' Update By         : Tedi/ Kreatif
'************************************************
    
    Select Case Index
        Case 0
            cmdOK_Click
        Case 1
            cmdCancel_Click
    End Select

End Sub

Private Sub picCmd_MouseMove(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
'************************************************
' Procedure         : picCmd_MouseMove
' Function          : Saat moise ada disekitar object picCmd
' Input Parameter   : -
' Output Parameter  : -
' Programmer By     : -
' Date Update       : Jan-2016
' Update By         : Tedi/ Kreatif
'************************************************

    ClrColor picCmd, &HF39A21
    picCmd(Index).BackColor = &H316C2

End Sub

Private Sub picForm_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
'************************************************
' Procedure         : picForm_MouseMove
' Function          : Saat moise ada disekitar Form
' Input Parameter   : -
' Output Parameter  : -
' Programmer By     : -
' Date Update       : Jan-2016
' Update By         : Tedi/ Kreatif
'************************************************

    ClrColor picCmd, &HF39A21

End Sub

Private Sub ClrColor(objPic, ByVal picBG)
'************************************************
' Procedure         : cmdCancel_Click
' Function          : Clear Color berdasarkan variable picBG
' Input Parameter   : -
' Output Parameter  : -
' Programmer By     : -
' Date Update       : Jan-2016
' Update By         : Tedi/ Kreatif
'************************************************

    Dim iBackGround As Integer
    For iBackGround = 0 To objPic.Count - 1
        If objPic(iBackGround).BackColor <> picBG Then
            objPic(iBackGround).BackColor = picBG
        End If
    Next iBackGround

End Sub



