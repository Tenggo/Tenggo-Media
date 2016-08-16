VERSION 5.00
Begin VB.Form Frm_Login 
   BorderStyle     =   0  'None
   Caption         =   "Security"
   ClientHeight    =   8820
   ClientLeft      =   0
   ClientTop       =   4515
   ClientWidth     =   12345
   Icon            =   "frm_Login.frx":0000
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
         Text            =   "wibisonog"
         Top             =   2925
         Width           =   2475
      End
      Begin VB.TextBox txtPassword 
         Height          =   375
         IMEMode         =   3  'DISABLE
         Left            =   2415
         PasswordChar    =   "*"
         TabIndex        =   5
         Text            =   "@admin123"
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
            Height          =   2085
            Left            =   705
            Picture         =   "frm_Login.frx":0442
            Top             =   270
            Width           =   3705
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
            Picture         =   "frm_Login.frx":5307
            Top             =   165
            Width           =   1695
         End
      End
      Begin VB.Label lbl_Check_User 
         BackStyle       =   0  'Transparent
         Caption         =   "Label1"
         ForeColor       =   &H00FFFFFF&
         Height          =   180
         Left            =   990
         TabIndex        =   12
         Top             =   7680
         Width           =   3405
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
         Picture         =   "frm_Login.frx":5C86
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
Option Explicit
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
Public Change_Pwd As Boolean
Dim rs As ADODB.Recordset
Dim rsUser_Id As ADODB.Recordset
Public OK As Boolean
Dim strSql As String
Dim strCheck_User As String
Private Declare Function GetComputerName Lib "kernel32" Alias "GetComputerNameA" (ByVal lpBuffer As String, nSize As Long) As Long

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
    
    Dim rsPasswordParameter As New ADODB.Recordset
    Dim rsDatePassword As New ADODB.Recordset
    Dim intPasswordwillexpire As Integer
    Dim DateNow As Date
    Dim DateChangePwd As Date
        
    strLogin_User = txtUserName.Text
    
    If strLogin_User = "" Then
        MsgBox "Enter User Name....!", vbExclamation, strApplication_Name
        Exit Sub
    End If
    
    blnCentral_Service_User = False
    
    Set rs = New ADODB.Recordset
    rs.CursorLocation = adUseClient
    
    strSql = ""
    strSql = strSql & "SELECT User_Id.*,Central_Service_User.user_name AS rpt_User "
    strSql = strSql & "FROM User_Id LEFT JOIN Central_Service_User ON Central_Service_User.user_name=User_Id.user_name  "
    strSql = strSql & "WHERE user_id.User_Name = '" & Clear_String(strLogin_User) & "'"
    
    rs.Open strSql, ConnERP, adOpenKeyset, adLockOptimistic, adCmdText
               
    If rs.RecordCount <> 0 Then
        
        strLogin_FullName = IIf(IsNull(rs!Name), "", rs!Name)
        strLogDivisionCode = rs!Division_Code
        strLogUserDisable = rs!disable_flag
        strLogMobile_Number = IIf(IsNull(rs!Mobile_Phone_Number), "", rs!Mobile_Phone_Number)
        strLogPhone_Number = IIf(IsNull(rs!Phone_Number), "", rs!Phone_Number)
        
        If IsNull(rs!rpt_User) Then
            blnCentral_Service_User = False
        Else
            blnCentral_Service_User = True
        End If
        
        'if user disable
        If strLogUserDisable = 1 Then
            MsgBox "User is disabled, please contact your administrator !", vbExclamation, strApplication_Name
            rs.Close
            Set rs = Nothing
            Exit Sub
        End If
               
        'Check Password
        If Trim$(txtPassword.Text) <> Decrypt(Trim$(rs("Password").Value)) Then
            'MsgBox DecryptPassword(Trim$(rs("Password").Value))
            MsgBox "Invalid Password, try again!", vbExclamation, strApplication_Name
            txtPassword.SetFocus
            'SendKeys "{Home}+{End}"
            rs.Close
            Set rs = Nothing
            Exit Sub
        End If
                        
        'cek apakah ada user lain yang sedang login dgn user tsb
        If UCase(strLogin_User) <> "ADMIN" Then
            strCheck_User = "."
        End If
        
        If strCheck_User = "." Then
            If Not IsNull(rs.Fields("Computer_Name").Value) Then
                If rs.Fields("Computer_Name").Value <> strComputer_Name Then
                    MsgBox "Sorry, Can only have one session at a time", vbExclamation, strApplication_Name
                    rs.Close
                    Set rs = Nothing
                    Exit Sub
                End If
            End If
        End If
                                       
        'Get Last Login & Last Log Out
        If IsNull(rs("Last_Login").Value) Then
            strLastLogin = ""
        Else
            strLastLogin = CStr(Format(rs("Last_Login").Value, "dddd, mmm d yyyy    hh:mm:ss AMPM"))
        End If
        If IsNull(rs("Last_Logout").Value) Then
            strLastLogout = ""
        Else
            strLastLogout = CStr(Format(rs("Last_Logout").Value, "dddd, mmm d yyyy    hh:mm:ss AMPM"))
        End If
        

        
        'Refresh Date
        recDate.Requery
        
        ' Update last login
        rs("Last_Login").Value = recDate(0).Value
        'Computer name
        If strCheck_User = "." Then
            rs("Computer_Name").Value = strComputer_Name
        End If
        
        rs.Update
                      
                
        'cek max age for user password
        BoolMaxAgePassword = False
        
        strSql = "SELECT * FROM Media_Parameter WHERE Parameter_Catalog = 'Password Policy' AND Parameter_Name = 'Maximum Password Age'"
        rsPasswordParameter.Open strSql, ConnERP, adOpenStatic, adLockReadOnly
        
        If Not rsPasswordParameter.EOF Then
            
            If CInt(rsPasswordParameter("value")) <> 0 Then
                strSql = "SELECT * FROM User_Id WHERE User_Name = '" & strLogin_User & "'"
                rsDatePassword.Open strSql, ConnERP, adOpenStatic, adLockReadOnly
                If Not rsDatePassword.EOF Then
                    recDate.Requery
                    'MsgBox Format(DateAdd("d", CInt(rsPasswordParameter("value")), rsDatePassword("Change_Password_date")), "mm/dd/yyyy")
                    DateNow = Format(recDate(0), "mm/dd/yyyy")
                    DateChangePwd = Format(DateAdd("d", CInt(rsPasswordParameter("value")), rsDatePassword("Change_Password_date")), "mm/dd/yyyy")
                    If DateNow > DateChangePwd Then
                        BoolMaxAgePassword = True
                        StrPasswordMassage = "Your Password has expired and must be changed."
                        Frm_Change_Password.show 1
                    Else
                        'Jika Password Akan Expire Dalam 5 hari sampai 1 Hari Muncul
                        'Pesan Apakah Akan Ubah Password
                        intPasswordwillexpire = DateDiff("d", recDate.Fields(0).Value, DateAdd("d", CDbl(rsPasswordParameter("value")), rsDatePassword("Change_Password_date")))
                        If intPasswordwillexpire < 6 Then
                            If MsgBox("Your password will expire in " & intPasswordwillexpire & " days," & vbCrLf & "Do you want to change it now ? ", vbExclamation + vbYesNo, strCompany_Name) = vbYes Then
                                Frm_Change_Password.show 1
                            End If
                        End If
                    End If
                End If
                rsDatePassword.Close
                Set rsDatePassword = Nothing
            End If
        End If
        rsPasswordParameter.Close
        Set rsPasswordParameter = Nothing
                               
            
        strCompany_Code = "IMI"
        strCompany_Name = Get_Company(strCompany_Code)
                
        Unload Me
        'show welcome form
        Frm_Welcome.show
        
    Else
        MsgBox "User Name Not Found, try again!", vbExclamation, strApplication_Name
        txtUserName.SetFocus
        SendKeys "{Home}+{End}"
    End If
    
    rs.Close
    Set rs = Nothing
End Sub
Private Sub cmdOK_Click2()
    
    Change_Pwd = True
    
    ''check for correct password
    strLogin_User = txtUserName.Text
       
    If strLogin_User = "" Then
        MsgBox "Enter User Name....!", vbExclamation, strApplication_Name
        Exit Sub
    End If
    
    Set rs = New ADODB.Recordset
    rs.CursorLocation = adUseClient
    rs.Open "Select User_Id.*,Match_IT_Admin.User_Name as AdminUser,Match_IT_Admin.IsAdmin,Match_IT_Admin.isGroupHead FROM User_ID LEFT JOIN Match_IT_Admin ON Match_IT_Admin.user_name=User_Id.user_name WHERE User_ID.User_Name = '" & Clear_String(strLogin_User) & "' and disable_flag <> 1  ", ConnERP, adOpenKeyset, adLockOptimistic, adCmdText
               
    If rs.RecordCount <> 0 Then
        
        ' ngambil strLogin_User dan DivisionCode dari table User_Id
        ' yang UserNAme nya sama dengan strLogin_User di Login_History
        strSql = ""
        strSql = strSql & " select User_Id.Company_Code, User_Id.Division_Code, User_Id.User_Name, user_id.name, User_Id.Password,"
        strSql = strSql & " User_Id.Last_Login, User_Id.Last_Logout"
        strSql = strSql & " from User_Id "
        strSql = strSql & " where User_Id.User_Name = '" & strLogin_User & "'"
    
        Set rsUser_Id = New ADODB.Recordset
        rsUser_Id.Open strSql, ConnERP, adOpenKeyset, adLockOptimistic, adCmdText
                
        If Not rsUser_Id.EOF And Not rsUser_Id.BOF Then
            strLogin_User = rsUser_Id!user_name
            strLogin_FullName = rsUser_Id!Name
            strLogDivisionCode = rsUser_Id!Division_Code
            strCompany_Code = rsUser_Id!Company_Code
        Else
            MsgBox "Invalid User Name", vbExclamation, "Login"
            txtUserName.SetFocus
            Exit Sub
        End If
        
        rsUser_Id.Close
        Set rsUser_Id = Nothing
        
        '***************************************
        'place code to here to pass the
        'success to the calling sub
        'setting a global var is the easiest
        '****************************************
        
        If Trim(txtPassword.Text) <> Decrypt(Trim(rs("Password").Value)) Then
            MsgBox "Invalid Password, try again!", vbExclamation, strApplication_Name
                txtPassword.SetFocus
            'SendKeys "{Home}+{End}"
            rs.Close
            Exit Sub
        End If
        
'        blnIsAdminUser = False
'        blnIsGroupHead = False
'
'        If Not IsNull(rs.Fields("AdminUser").Value) Then
'
'            If rs.Fields("IsAdmin").Value = 1 Then
'                blnIsAdminUser = True
'            Else
'                blnIsAdminUser = False
'            End If
'
'            If rs.Fields("IsGroupHead").Value = 1 Then
'                blnIsGroupHead = True
'            Else
'                blnIsGroupHead = False
'            End If
'        End If
        
        'Timer1.Enabled = False
                    
        Me.Hide
                        
        mdi_Main.show
        
    Else
        MsgBox "User Name Not Found, try again!", vbExclamation, strApplication_Name
        txtUserName.SetFocus
       ' SendKeys "{Home}+{End}"
    End If
    
    rs.Close
    Set rs = Nothing
  End Sub

Private Sub cmdOK_Click1()
'************************************************
' Procedure         : cmdOK_Click
' Function          : Cek User dan Pass word, jika valid blnLogin = True dan jika tidak valid blnLogin = False
' Input Parameter   : -
' Output Parameter  : -
' Programmer By     : -
' Date Update       : Jan-2016
' Update By         : Tedi/ Kreatif
'************************************************
   
    Dim rsPasswordParameter As New ADODB.Recordset
    Dim rsDatePassword As New ADODB.Recordset
    Dim intPasswordwillexpire As Integer
    Dim DateNow As Date
    Dim DateChangePwd As Date
        
    Dim rec_User As New ADODB.Recordset
    strLogin_User = txtUserName.Text
    
    If strLogin_User = "" Then
        MsgBox "Enter User Name....!", vbExclamation, strApplication_Name
        Exit Sub
    End If
    
    blnCentral_Service_User = False
    
    Set rec_User = New ADODB.Recordset
    rec_User.CursorLocation = adUseClient
    rec_User.Open "SELECT User_Id.*,Central_Service_User.user_name AS rpt_User FROM User_Id LEFT JOIN Central_Service_User ON Central_Service_User.user_name=User_Id.user_name  WHERE user_id.User_Name = '" & Clear_String(strLogin_User) & "'", ConnERP, adOpenKeyset, adLockOptimistic, adCmdText
               
    If rec_User.RecordCount <> 0 Then
        
        strLogin_FullName = IIf(IsNull(rec_User!Name), "", rec_User!Name)
        strLogDivisionCode = rec_User!Division_Code
        strLogUserDisable = rec_User!disable_flag
        strLogMobile_Number = IIf(IsNull(rec_User!Mobile_Phone_Number), "", rec_User!Mobile_Phone_Number)
        strLogPhone_Number = IIf(IsNull(rec_User!Phone_Number), "", rec_User!Phone_Number)
        
        If IsNull(rec_User!rpt_User) Then
            blnCentral_Service_User = False
        Else
            blnCentral_Service_User = True
        End If
        
        'if user disable
        If strLogUserDisable = 1 Then
            MsgBox "User is disabled, please contact your administrator !", vbExclamation, strApplication_Name
            rec_User.Close
            Set rec_User = Nothing
            Exit Sub
        End If
               
        'Check Password
        If Trim$(txtPassword.Text) <> Decrypt(Trim$(rec_User("Password").Value)) Then
            'MsgBox DecryptPassword(Trim$(rs("Password").Value))
            MsgBox "Invalid Password, try again!", vbExclamation, strApplication_Name
            txtPassword.SetFocus
            SendKeys "{Home}+{End}"
            rec_User.Close
            Set rec_User = Nothing
            Exit Sub
        End If
                        
        'cek apakah ada user lain yang sedang login dgn user tsb
        If UCase(strLogin_User) <> "ADMIN" Then
            If Not IsNull(rec_User.Fields("Computer_Name").Value) Then
                If rec_User.Fields("Computer_Name").Value <> strComputer_Name Then
                    MsgBox "Sorry, Can only have one session at a time", vbExclamation, strApplication_Name
                    rec_User.Close
                    Set rec_User = Nothing
                    Exit Sub
                End If
            End If
        End If
                                       
        'Get Last Login & Last Log Out
        If IsNull(rec_User("Last_Login").Value) Then
            strLastLogin = ""
        Else
            strLastLogin = CStr(Format(rec_User("Last_Login").Value, "dddd, mmm d yyyy    hh:mm:ss AMPM"))
        End If
        If IsNull(rec_User("Last_Logout").Value) Then
            strLastLogout = ""
        Else
            strLastLogout = CStr(Format(rec_User("Last_Logout").Value, "dddd, mmm d yyyy    hh:mm:ss AMPM"))
        End If
        
        
        'Refresh Date
        recDate.Requery
        
        ' Update last login
        rec_User("Last_Login").Value = recDate(0).Value
        'Computer name
        If Me.lbl_Check_User.Caption = "." Then
            rec_User("Computer_Name").Value = strComputer_Name
        End If
        
        rec_User.Update
                      
        'cek max age for user password
        BoolMaxAgePassword = False
        
        strSql = "SELECT * FROM Media_Parameter WHERE Parameter_Catalog = 'Password Policy' AND Parameter_Name = 'Maximum Password Age'"
        rsPasswordParameter.Open strSql, ConnERP, adOpenStatic, adLockReadOnly
        
        If Not rsPasswordParameter.EOF Then
            
            If CInt(rsPasswordParameter("value")) <> 0 Then
                strSql = "SELECT * FROM User_Id WHERE User_Name = '" & strLogin_User & "'"
                rsDatePassword.Open strSql, ConnERP, adOpenStatic, adLockReadOnly
                If Not rsDatePassword.EOF Then
                    recDate.Requery
                    'MsgBox Format(DateAdd("d", CInt(rsPasswordParameter("value")), rsDatePassword("Change_Password_date")), "mm/dd/yyyy")
                    DateNow = Format(recDate(0), "mm/dd/yyyy")
                    DateChangePwd = Format(DateAdd("d", CInt(rsPasswordParameter("value")), rsDatePassword("Change_Password_date")), "mm/dd/yyyy")
                    If DateNow > DateChangePwd Then
                        BoolMaxAgePassword = True
                        StrPasswordMassage = "Your Password has expired and must be changed."
                        Frm_Change_Password.show 1
                    Else
                        'Jika Password Akan Expire Dalam 5 hari sampai 1 Hari Muncul
                        'Pesan Apakah Akan Ubah Password
                        intPasswordwillexpire = DateDiff("d", recDate.Fields(0).Value, DateAdd("d", CDbl(rsPasswordParameter("value")), rsDatePassword("Change_Password_date")))
                        If intPasswordwillexpire < 6 Then
                            If MsgBox("Your password will expire in " & intPasswordwillexpire & " days," & vbCrLf & "Do you want to change it now ? ", vbExclamation + vbYesNo, strApplication_Name) = vbYes Then
                                Frm_Change_Password.show 1
                            End If
                        End If
                    End If
                End If
                rsDatePassword.Close
                Set rsDatePassword = Nothing
            End If
        End If
        rsPasswordParameter.Close
        Set rsPasswordParameter = Nothing
                               
            
        strCompany_Code = "IMI"
        strCompany_Name = Get_Company(strCompany_Code)
                
        Unload Me
        'show welcome form
        'Frm_Welcome.show
        
    Else
        MsgBox "User Name Not Found, try again!", vbExclamation, strApplication_Name
        txtUserName.SetFocus
        SendKeys "{Home}+{End}"
    End If
    
    rec_User.Close
    Set rec_User = Nothing
End Sub
'#End Region


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
    GetPictureFromDB imgBG, vntCompany_BackGround
    lblDesc(0).Caption = Chr(169) & " 2015 PT. Ilab Komunikasi Indonesia. All rights reserved."
    'Exit Sub
    Dim rec_Version As New ADODB.Recordset
    Dim rec_Path As New ADODB.Recordset
    Dim strQuery As String
    Dim Ero As ADODB.Error
            
    'On Error GoTo Lable
    
    Me.Caption = "Security"
    
    blnSkipMultiSesion = False
    
    '========= Orginal Connection ==================
'    Time out Connection
'    ConnERP.ConnectionTimeout = 0
'    Time Out Command
'    ConnERP.CommandTimeout = 0
'
'    ConnERP.Open "Provider=SQLOLEDB.1;Persist Security Info=False;User ID=" & strLogin_User & ";password=" & strLogin_Password & ";Initial Catalog=" & strDatabase_Name & ";Data Source=" & strServerName
'    ========= End Orginal Connection ==============
'
    'Refresh Date
    If Not recDate.State = adStateOpen Then
        recDate.Open "Server_Date", ConnERP, adOpenStatic, adLockReadOnly, adCmdStoredProc
    End If
        
'    Set rec_Version = Nothing
    
    Dim recFinanceParameter As New ADODB.Recordset
    '============= Media Parameter ============
    strQuery = "SELECT * FROM App_Media_Par"
    recFinanceParameter.Open strQuery, ConnERP, adOpenStatic, adLockReadOnly
    If Not recFinanceParameter.EOF Then

        strPathPO = IIf(IsNull(recFinanceParameter.Fields("POPath").Value), "", recFinanceParameter.Fields("POPath").Value)
        strPaperType = IIf(IsNull(recFinanceParameter.Fields("POPaperType").Value), "", recFinanceParameter.Fields("POPaperType").Value)
        strPaperTypePOTV = IIf(IsNull(recFinanceParameter.Fields("POTVPaidPaperType").Value), "", recFinanceParameter.Fields("POTVPaidPaperType").Value)
        strPaperTypeCOTV = IIf(IsNull(recFinanceParameter.Fields("COTVPaidPaperType").Value), "", recFinanceParameter.Fields("COTVPaidPaperType").Value)
        strPaperTypeCO = IIf(IsNull(recFinanceParameter.Fields("COPaperType").Value), "", recFinanceParameter.Fields("COPaperType").Value)
        strPOCOCopyNumber = IIf(IsNull(recFinanceParameter.Fields("POCO_Copy_Number").Value), "1", recFinanceParameter.Fields("POCO_Copy_Number").Value)
        recFinanceParameter.MoveNext
        
    End If

    recFinanceParameter.Close
    Set recFinanceParameter = Nothing
    '====================================================


    '=========== Get Report Dir ===================
    'App_Media_Report_Path
    rec_Path.Open "SELECT * FROM App_Media_Report_Path WHERE AppID='" & Left(strAppID, 2) & "'", ConnERP
    If rec_Path.EOF Then
        MsgBox "Report Path not found.", vbExclamation, strApplication_Name
        End
    Else
        strReport_Dir = Trim$(rec_Path.Fields("Report_Path").Value)
    End If
    rec_Path.Close
    
    '=========== End Get Report Dir ===================
    
    '=========== Get SMTP Server Addr, CRM Addr ==============
      
    rec_Path.Open "SELECT * FROM App_Media_Mail WHERE AppID='" & Left(strAppID, 2) & "'", ConnERP
    
    If rec_Path.EOF Then
    
        MsgBox "Email Setting & CRM Address not found.", vbExclamation, strApplication_Name
        End
    
    Else
    
        StrCRMaddress = Trim$(rec_Path.Fields("CRM_Address").Value)
        strSMTP_Server = Trim$(rec_Path.Fields("SMTP_Address").Value)
        '
        strEmail_From = Trim$(rec_Path.Fields("Email_From").Value)
        strEmail_From_Name = Trim$(rec_Path.Fields("Email_From_Name").Value)
        '
        strMP_Email_Subject = Trim$(rec_Path.Fields("MP_Email_Subject").Value)
        strMP_Email_Template_File = Trim$(rec_Path.Fields("MP_Email_Template_File").Value)
        '
        strQuotation_Email_Subject = Trim$(rec_Path.Fields("Quotation_Email_Subject").Value)
        strQuotation_Email_Template_File = Trim$(rec_Path.Fields("Quotation_Email_Template_File").Value)
        '
        strBC_Email_Subject = Trim$(rec_Path.Fields("BC_Email_Subject").Value)
        strBC_Email_Template_File = Trim$(rec_Path.Fields("BC_Email_Template_File").Value)
        '

    End If
    
    rec_Path.Close
    
    '=========== Get SMTP Server Addr, CRM Addr ==============
    Set rec_Path = Nothing
    
    'Get Percent Tax
    Call Load_Taxes_Data
    
    Dim strBuffer As String
    Dim lngLen As Long
    Dim Th As Integer
    'get komputer name
    
    strBuffer = Space(255)
    lngLen = Len(strBuffer)
    If CBool(GetComputerName(strBuffer, lngLen)) Then
        strComputer_Name = Left$(strBuffer, lngLen)
    Else
        strComputer_Name = "Local"
    End If
    
    
    Exit Sub
        
Lable:

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

Private Sub picCmd_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
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

Private Sub picForm_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
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



