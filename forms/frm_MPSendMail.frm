VERSION 5.00
Begin VB.Form frm_MPSendMail 
   Caption         =   "Send Mail"
   ClientHeight    =   6900
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   7065
   LinkTopic       =   "Form1"
   ScaleHeight     =   6900
   ScaleWidth      =   7065
   StartUpPosition =   1  'CenterOwner
   Begin VB.PictureBox picToolbar 
      Align           =   1  'Align Top
      Appearance      =   0  'Flat
      BackColor       =   &H00F0F0F0&
      BorderStyle     =   0  'None
      ClipControls    =   0   'False
      ForeColor       =   &H80000008&
      Height          =   750
      Left            =   0
      ScaleHeight     =   750
      ScaleWidth      =   7065
      TabIndex        =   17
      TabStop         =   0   'False
      Top             =   0
      Width           =   7065
      Begin VB.PictureBox picButton 
         Appearance      =   0  'Flat
         BackColor       =   &H00F0F0F0&
         BorderStyle     =   0  'None
         ClipControls    =   0   'False
         ForeColor       =   &H80000008&
         Height          =   750
         Index           =   71
         Left            =   1605
         Picture         =   "frm_MPSendMail.frx":0000
         ScaleHeight     =   750
         ScaleWidth      =   1500
         TabIndex        =   20
         TabStop         =   0   'False
         Top             =   0
         Width           =   1500
      End
      Begin VB.PictureBox picButton 
         Appearance      =   0  'Flat
         BackColor       =   &H00F0F0F0&
         BorderStyle     =   0  'None
         ClipControls    =   0   'False
         ForeColor       =   &H80000008&
         Height          =   750
         Index           =   23
         Left            =   3120
         Picture         =   "frm_MPSendMail.frx":1E12
         ScaleHeight     =   750
         ScaleWidth      =   1500
         TabIndex        =   19
         TabStop         =   0   'False
         Top             =   15
         Width           =   1500
      End
      Begin VB.PictureBox picButton 
         Appearance      =   0  'Flat
         BackColor       =   &H00F0F0F0&
         BorderStyle     =   0  'None
         ClipControls    =   0   'False
         ForeColor       =   &H80000008&
         Height          =   750
         Index           =   21
         Left            =   90
         Picture         =   "frm_MPSendMail.frx":3B19
         ScaleHeight     =   750
         ScaleWidth      =   1500
         TabIndex        =   18
         TabStop         =   0   'False
         Top             =   0
         Width           =   1500
      End
   End
   Begin VB.Frame Frame2 
      Height          =   990
      Left            =   -15
      TabIndex        =   6
      Top             =   5310
      Width           =   7875
      Begin VB.TextBox Txt_Email_Template 
         Appearance      =   0  'Flat
         Height          =   330
         Left            =   1725
         Locked          =   -1  'True
         TabIndex        =   7
         Text            =   "MP_Mail.htm"
         Top             =   345
         Visible         =   0   'False
         Width           =   2235
      End
      Begin VB.Label Label1 
         Caption         =   "Template file (HTML) : "
         Height          =   285
         Left            =   120
         TabIndex        =   8
         Top             =   360
         Visible         =   0   'False
         Width           =   1605
      End
   End
   Begin VB.Frame Frame1 
      Height          =   4560
      Left            =   0
      TabIndex        =   0
      Top             =   720
      Width           =   7095
      Begin VB.ListBox Lstcc 
         Height          =   1635
         Left            =   3990
         Style           =   1  'Checkbox
         TabIndex        =   15
         Top             =   2625
         Width           =   2925
      End
      Begin VB.Frame Frame3 
         Height          =   1650
         Left            =   3180
         TabIndex        =   12
         Top             =   2580
         Width           =   735
         Begin VB.CommandButton CmdLeftcc 
            Caption         =   "<"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Left            =   150
            Style           =   1  'Graphical
            TabIndex        =   14
            Top             =   885
            Width           =   495
         End
         Begin VB.CommandButton CmdRightcc 
            Caption         =   ">"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Left            =   135
            Style           =   1  'Graphical
            TabIndex        =   13
            Top             =   195
            Width           =   495
         End
      End
      Begin VB.Frame Frame10 
         Height          =   1575
         Left            =   3180
         TabIndex        =   3
         Top             =   555
         Width           =   735
         Begin VB.CommandButton cmdRightOne 
            Caption         =   ">"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Left            =   135
            Style           =   1  'Graphical
            TabIndex        =   5
            Top             =   195
            Width           =   495
         End
         Begin VB.CommandButton cmdLeftOne 
            Caption         =   "<"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Left            =   150
            Style           =   1  'Graphical
            TabIndex        =   4
            Top             =   870
            Width           =   495
         End
      End
      Begin VB.ListBox lstTO 
         Height          =   1635
         Left            =   3975
         Style           =   1  'Checkbox
         TabIndex        =   2
         Top             =   540
         Width           =   2925
      End
      Begin VB.ListBox lstUsers 
         Height          =   3660
         Left            =   120
         Sorted          =   -1  'True
         Style           =   1  'Checkbox
         TabIndex        =   1
         Top             =   555
         Width           =   2970
      End
      Begin VB.Label Label4 
         Caption         =   "cc:"
         Height          =   255
         Left            =   3990
         TabIndex        =   16
         Top             =   2280
         Width           =   2895
      End
      Begin VB.Label Label3 
         Caption         =   "To :"
         Height          =   255
         Left            =   3930
         TabIndex        =   11
         Top             =   195
         Width           =   2895
      End
      Begin VB.Label Label2 
         Caption         =   "List Client who can access selected MP :"
         Height          =   255
         Left            =   120
         TabIndex        =   10
         Top             =   255
         Width           =   2895
      End
   End
   Begin VB.Label Lbl_Status 
      Caption         =   "Lbl_Status"
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   255
      TabIndex        =   9
      Top             =   6420
      Width           =   6405
   End
End
Attribute VB_Name = "frm_MPSendMail"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'************************************************************
' Unit/Module Name  : Frm_Survey_Respondent
' Function          : Add & Delete Survey Respondent
'                   :
' Public Variables  : -
' Used Tables       : Survey_Respondent
'                   :
' Date              :
' Created By        : YY
' Last Update/By    :
'************************************************************
'
Option Explicit

Dim strSql As String
Dim LogFileName As String

Dim fso As Object
Dim FSOInterface As Object
Dim ClientApprovalType As Integer
Dim StrMP_Number As String


Private Sub Cmd_Preview_Mail()
    
    Dim rsWebUser As New ADODB.Recordset
    
    Dim StrOrgBody As String
    Dim StrBody As String
    Dim StrEmailAddress As String
    Dim StrClientName As String
    Dim strBrandName As String
    Dim LenofUserName As Integer
        
    Dim idx As Integer
    
    
    'Checking apakah ada data di list to
    If Me.lstTO.ListCount = 0 Then
        MsgBox "Please select a client to send mail.", vbExclamation, "Missing information", vbExclamation, strApplication_Name
        Exit Sub
    End If
    
    'Read Template File
    StrOrgBody = Readtemplate
    'Read Template File (Check Variable Full Name,User_Name,Password,Link Survey, Link Direct Survey)
                        
    If StrOrgBody = "" Then
        MsgBox "Template file empty. Please select correct template file.", vbCritical, "Survey Generator"
        Exit Sub
    End If
      
    'BrandName
    strBrandName = Get_Brand_Name(Left(frm_MPInsertion.cboMPNumber, 4))
                       
    '------------- Mail Address -------------------
   
    For idx = 0 To Me.lstTO.ListCount - 1 'Loop Selected Respondent
                                   
        LenofUserName = Len(Trim(lstTO.List(idx))) - InStr(1, Trim(lstTO.List(idx)), "(") - 1
            
        strSql = "SELECT * FROM User_Web WHERE User_Name='" & Mid(Trim(lstTO.List(idx)), InStr(1, Trim(lstTO.List(idx)), "(") + 1, LenofUserName) & "'"
            
        'Open Recordset Selected
        rsWebUser.Open strSql, ConnERP, adOpenKeyset, adLockOptimistic
            
        StrClientName = rsWebUser.Fields("Name").Value
        
        If Not IsNull(rsWebUser.Fields("Email").Value) Then
            If Trim(rsWebUser.Fields("Email").Value) <> "" Then
                StrEmailAddress = Trim(rsWebUser.Fields("Email").Value)
                               
            End If
        End If
        rsWebUser.Close
    Next idx
     
    
    'Generate Body
    StrBody = StrOrgBody
    'Full Name
    StrBody = Replace(StrBody, "%LINKLOGO%", strCompany_Logo_Link)
    'Full Name
    StrBody = Replace(StrBody, "%FULLNAME%", StrClientName)
    'Planner
    StrBody = Replace(StrBody, "%PLANNER%", strLogin_FullName)
    'Brand Name
    StrBody = Replace(StrBody, "%BRANDNAME%", strBrandName)
    'MediaPlan Number
    StrBody = Replace(StrBody, "%MP_NUMBER%", StrMP_Number)
    'I-Quest Address
    StrBody = Replace(StrBody, "%IQUESTADDRESS%", StrCRMaddress)
    'Sender
    StrBody = Replace(StrBody, "%SENDER%", strLogin_FullName)
     'Sender Phone+ExtNo
    StrBody = Replace(StrBody, "%SENDERPHONEEXT%", strLogPhone_Number)
     'Sender Modile Nuumber
    StrBody = Replace(StrBody, "%SENDERMOBILE%", strLogMobile_Number)
     
        'Cmd_Preview_Mail.Enabled = False
    'Show Body
        Dim fsotmp As Object
        Dim fsoint As Object
        Dim i As Integer
        
        'Write Temp File
        Set fsotmp = CreateObject("Scripting.FileSystemObject")
        Set fsoint = fsotmp.CreateTextFile(App.Path & "\" & "BCTMP.HTM", True)
        fsoint.WriteLine StrBody
        fsoint.Close
        Set fsoint = Nothing
        Set fsotmp = Nothing
        
        'subDelay 10
        
        'Show To Browser
        Dim FrmBrowser As New Frm_MPBrowser
        FrmBrowser.Caption = "View Email"
        FrmBrowser.show 1
        
        'Cmd_Preview_Mail.Enabled = True
End Sub

Private Sub cmdClose()
    Unload Me
End Sub

Private Sub CmdLeftcc_Click()
    Dim i As Integer
    Dim panjang As Integer
    
    For i = Lstcc.ListCount - 1 To 0 Step -1
        If Lstcc.Selected(i) = True Then
           ' lstTO.AddItem lstUsers.List(i)
            Me.lstUsers.AddItem Lstcc.List(i)
            Lstcc.RemoveItem (i)
        End If
    Next i
End Sub

Private Sub cmdLeftOne_Click()
    Dim i As Integer
    Dim panjang As Integer
    
    For i = lstTO.ListCount - 1 To 0 Step -1
        If lstTO.Selected(i) = True Then
           ' lstTO.AddItem lstUsers.List(i)
            Me.lstUsers.AddItem lstTO.List(i)
            lstTO.RemoveItem (i)
        End If
    Next i
    
End Sub



Private Sub CmdRightcc_Click()
    Dim i As Integer
    Dim panjang As Integer
    
    For i = lstUsers.ListCount - 1 To 0 Step -1
        If lstUsers.Selected(i) = True Then
           ' lstTO.AddItem lstUsers.List(i)
            Me.Lstcc.AddItem lstUsers.List(i)
            lstUsers.RemoveItem (i)
        End If
    Next i
End Sub

Private Sub cmdRightOne_Click()
    Dim i As Integer
    Dim panjang As Integer
    
    For i = lstUsers.ListCount - 1 To 0 Step -1
        If lstUsers.Selected(i) = True Then
           
           'check apakah sudah ada item di list
            If lstTO.ListCount > 0 Then
                MsgBox "This mail can only send to one Client, please use cc if you want to send mail to other client.", vbExclamation, strApplication_Name
                Exit For
            End If
            
            Me.lstTO.AddItem lstUsers.List(i)
            lstUsers.RemoveItem (i)
        End If
    Next i
    
        
End Sub

Private Sub Form_Load()
    Dim strSql As String
    Dim rs_Client As New ADODB.Recordset
    
    strSql = "SELECT A.MP_Approval_type FROM CLient A,Brand B WHERE A.client_code=B.client_code AND B.Brand_code='" & Left(frm_MPInsertion.cboMPNumber, 4) & "'"
        
    'exit sub
    rs_Client.Open strSql, ConnERP
    
    If rs_Client.EOF Then
        ClientApprovalType = 1 'Default
    Else
        ClientApprovalType = rs_Client("MP_Approval_type").Value
    End If
    
    rs_Client.Close
    Set rs_Client = Nothing
    
    Lbl_Status.Caption = ""
    
    StrMP_Number = frm_MPInsertion.cboMPNumber
    
    'If Client 'SAH
        LoadAllUser 'All Web user who have access to this brand
    'Else
        'Load UserBiasa
    'End if
    
End Sub

Private Sub LoadAllUser()
    'User_Id
    Dim rsUser As New ADODB.Recordset
    
    'Set mail to ke Media
    Me.lstTO.Clear
    
    'strSQL = "SELECT user_name, name FROM User_Web WHERE Role IN ('MARKETING','MEDIA','MEDIA2','FINANCE','PLANNER','CCM') AND User_Name IN("
    'strSQL = strSQL & " SELECT user_name FROM User_Web_Security WHERE Brand_Code='" & Left(Frm_MPInsertion.cboMPNumber, 4) & "'"
    'strSQL = strSQL & " AND valid_until>=(select getdate()) )"
    
    If ClientApprovalType = 2 Then
        strSql = "SELECT U.user_name, U.name,S.Position FROM User_Web U, User_Web_Security S"
        strSql = strSql & " WHERE  U.User_Name=S.User_Name AND"
        strSql = strSql & " U.Role IN ('MEDIA')   AND S.Position = 'HEAD'"
        strSql = strSql & " AND Brand_Code='" & Left(frm_MPInsertion.cboMPNumber, 4) & "'"
        strSql = strSql & " AND U.User_Name IN(SELECT user_name FROM User_Web_Security WHERE Brand_Code='" & Left(frm_MPInsertion.cboMPNumber, 4) & "'AND valid_until>=(select getdate()))"
    Else 'Type 1
        strSql = "SELECT U.user_name, U.name,S.Position FROM User_Web U, User_Web_Security S"
        strSql = strSql & " WHERE  U.User_Name=S.User_Name AND"
        strSql = strSql & " U.Role IN ('MARKETING')   AND S.Position = 'HEAD'"
        strSql = strSql & " AND Brand_Code='" & Left(frm_MPInsertion.cboMPNumber, 4) & "'"
        strSql = strSql & " AND U.User_Name IN(SELECT user_name FROM User_Web_Security WHERE Brand_Code='" & Left(frm_MPInsertion.cboMPNumber, 4) & "'AND valid_until>=(select getdate()))"
    End If
    
    rsUser.Open strSql, ConnERP, adOpenForwardOnly, adLockReadOnly, adCmdText
    
    'Clear Previouse Data
    lstUsers.Clear
    
    Do While Not rsUser.EOF
        'lstUsers.AddItem rsUser!Name & " (" & rsUser!user_name & ")"
        lstTO.AddItem rsUser!Name & " (" & rsUser!user_name & ")"
        rsUser.MoveNext
    Loop
    
    rsUser.Close
               
    
    'Yang lainnya ke CC
    Me.Lstcc.Clear
    
    If ClientApprovalType = 2 Then
        strSql = "SELECT U.user_name, U.name,S.Position FROM User_Web U, User_Web_Security S"
        strSql = strSql & " WHERE  U.User_Name=S.User_Name AND"
        strSql = strSql & " (U.Role IN ('MARKETING','MEDIA2','FINANCE','PLANNER','CCM')"
        strSql = strSql & " OR"
        strSql = strSql & " (U.Role IN ('MEDIA')   AND S.Position <> 'HEAD'))"
        strSql = strSql & " AND Brand_Code='" & Left(frm_MPInsertion.cboMPNumber, 4) & "'"
        strSql = strSql & " AND U.User_Name IN(SELECT user_name FROM User_Web_Security WHERE Brand_Code='" & Left(frm_MPInsertion.cboMPNumber, 4) & "'AND valid_until>=(select getdate()))"
    Else 'Type 1
        strSql = "SELECT U.user_name, U.name,S.Position FROM User_Web U, User_Web_Security S"
        strSql = strSql & " WHERE  U.User_Name=S.User_Name AND"
        strSql = strSql & " (U.Role IN ('MEDIA','MEDIA2','FINANCE','PLANNER','CCM')"
        strSql = strSql & " OR"
        strSql = strSql & " (U.Role IN ('MARKETING')   AND S.Position <> 'HEAD'))"
        strSql = strSql & " AND Brand_Code='" & Left(frm_MPInsertion.cboMPNumber, 4) & "'"
        strSql = strSql & " AND U.User_Name IN(SELECT user_name FROM User_Web_Security WHERE Brand_Code='" & Left(frm_MPInsertion.cboMPNumber, 4) & "'AND valid_until>=(select getdate()))"
    End If
    
    rsUser.Open strSql, ConnERP, adOpenForwardOnly, adLockReadOnly, adCmdText
    
    'Clear Previouse Data
    lstUsers.Clear
    
    Do While Not rsUser.EOF
        'lstUsers.AddItem rsUser!Name & " (" & rsUser!user_name & ")"
        Lstcc.AddItem rsUser!Name & " (" & rsUser!user_name & ")"
        rsUser.MoveNext
    Loop
    
    rsUser.Close
        
    Set rsUser = Nothing
    
End Sub



Private Sub Send_Mail()

    Dim rsWebUser As New ADODB.Recordset
    Dim objMail  As New ASPEMAILLib.MailSender
    
    Dim StrOrgBody As String
    Dim StrBody As String
    Dim StrSubject As String
    Dim StrClientName As String
    Dim strBrandName As String
    Dim StrEmailAddress As String
        
    Dim LenofUserName As Integer
    
    
    Dim idx As Integer
    
    
    'Checking apakah ada data di list to
    If Me.lstTO.ListCount = 0 Then
        MsgBox "Please select a client to send mail.", vbExclamation, strApplication_Name
        Exit Sub
    End If
    'Read Template File
    StrOrgBody = Readtemplate
    'Read Template File (Check Variable Full Name,User_Name,Password,Link Survey, Link Direct Survey)
                        
    If StrOrgBody = "" Then
        MsgBox "Template file empty. Please select correct template file.", vbCritical, "Survey Generator", vbExclamation, strApplication_Name
        Exit Sub
    End If
      
    'BrandName
    strBrandName = Get_Brand_Name(Left(frm_MPInsertion.cboMPNumber, 4))
    
     'Create Log File
    Call Create_Log_File
    
    StrSubject = Str_MP_Email_Subject
                   
    '------------- Mail Address -------------------
    StrEmailAddress = ""
    For idx = 0 To Me.lstTO.ListCount - 1 'Loop Selected Respondent
                                   
        LenofUserName = Len(Trim(lstTO.List(idx))) - InStr(1, Trim(lstTO.List(idx)), "(") - 1
            
        strSql = "SELECT * FROM User_Web WHERE User_Name='" & Mid(Trim(lstTO.List(idx)), InStr(1, Trim(lstTO.List(idx)), "(") + 1, LenofUserName) & "'"
            
        'Open Recordset Selected
        rsWebUser.Open strSql, ConnERP, adOpenKeyset, adLockOptimistic
            
        StrClientName = rsWebUser.Fields("Name").Value
        If Not IsNull(rsWebUser.Fields("Email").Value) Then
            If Trim(rsWebUser.Fields("Email").Value) <> "" Then
                StrEmailAddress = Trim(rsWebUser.Fields("Email").Value)
                'Email Address checking ?????????????
                objMail.AddAddress StrEmailAddress
                'Log
                Call Write_To_Log_File(Now & vbTab & "Sending email to : " & rsWebUser.Fields("Name").Value & vbTab & rsWebUser.Fields("Email").Value)
            End If
        End If
        rsWebUser.Close
    Next idx
    '--------------- CC -----------------------
    StrEmailAddress = ""
    For idx = 0 To Me.Lstcc.ListCount - 1 'Loop Selected cc
                                   
        LenofUserName = Len(Trim(Lstcc.List(idx))) - InStr(1, Trim(Lstcc.List(idx)), "(") - 1
            
        strSql = "SELECT * FROM User_Web WHERE User_Name='" & Mid(Trim(Lstcc.List(idx)), InStr(1, Trim(Lstcc.List(idx)), "(") + 1, LenofUserName) & "'"
            
        'Open Recordset Selected
        rsWebUser.Open strSql, ConnERP, adOpenKeyset, adLockOptimistic
        If Not IsNull(rsWebUser.Fields("Email").Value) Then
            If Trim(rsWebUser.Fields("Email").Value) <> "" Then
                StrEmailAddress = Trim(rsWebUser.Fields("Email").Value)
                'Email Address checking ?????????????
                objMail.AddCC StrEmailAddress
                'Log
                Call Write_To_Log_File(Now & vbTab & "cc email to : " & rsWebUser.Fields("Name").Value & vbTab & rsWebUser.Fields("Email").Value)
            End If
        End If
        rsWebUser.Close
    Next idx 'End Loop
    '------------------- End cc ------------------------
    
    'Generate Body
    StrBody = StrOrgBody
    'Full Name
    StrBody = Replace(StrBody, "%FULLNAME%", StrClientName)
    'Planner
    StrBody = Replace(StrBody, "%PLANNER%", strLogin_FullName)
    'Brand Name
    StrBody = Replace(StrBody, "%BRANDNAME%", strBrandName)
    'MediaPlan Number
    StrBody = Replace(StrBody, "%MP_NUMBER%", StrMP_Number)
    'I-Quest Address
    StrBody = Replace(StrBody, "%IQUESTADDRESS%", StrCRMaddress)
    'Sender
    StrBody = Replace(StrBody, "%SENDER%", strLogin_FullName)
     'Sender Phone+ExtNo
    StrBody = Replace(StrBody, "%SENDERPHONEEXT%", strLogPhone_Number)
     'Sender Modile Nuumber
    StrBody = Replace(StrBody, "%SENDERMOBILE%", strLogMobile_Number)
     
    'End Generate Body
               
                    
    'On Error Resume Next
            
    If strSMTP_Server = "" Then
        MsgBox "SMTP Server has not been configured!", vbExclamation, strApplication_Name
        Exit Sub
    End If
        
    objMail.Host = strSMTP_Server
           
    'From
    objMail.From = strEmail_From_Name
    objMail.FromName = strEmail_From_Name
                        
    ' message subject
    objMail.Subject = StrSubject & " Media Plan No : " & StrMP_Number & " (" & Time & ")"
        
    ' message body
    objMail.Body = StrBody
        
    objMail.IsHTML = True
    If strCompany_isGmail = " " Then
'        objMail.Port = strCompany_isPort
'        objMail.UserName = strCompany_UsrEmail
'        objMail.Password = strCompany_PassEmail
    End If
    ' send message
    objMail.Send
        
    'Remove Object from memory
    Set objMail = Nothing
        
    If Err.Number <> 0 Then
        MsgBox Err.Description & vbCrLf & vbCrLf & "Error Code : " & Err.Number, vbExclamation, strApplication_Name
            
        Err.Clear
                                           
        Call Write_To_Log_File(Now & vbTab & "Sending Mail Fail")
    Else
        Call Write_To_Log_File(Now & vbTab & "Sending Mail Done")
    End If
                              
    '================ ==============================
       
    Lbl_Status.Caption = ""
    
    'Close Log File
    Call Close_Log_File
        
End Sub

Private Function Readtemplate() As String
        
    Dim fso As New Scripting.FileSystemObject
    Dim fl As File
    Dim template_path As String
       
    'If IM
    template_path = App.Path & "\temp\email-template.htm" '& Str_MP_Email_Template_File
    'IF BC
    
    'IF UM
    
    If Not fso.FileExists(template_path) Then
        MsgBox "Template File (" & template_path & ") not found!", vbCritical, "File Not Found.", vbExclamation, strApplication_Name
        Exit Function
    End If
    
    Set fl = fso.GetFile(template_path)
           
    Readtemplate = fl.OpenAsTextStream.ReadAll
                 
End Function

Private Sub ChangeVariableLogo()
   Dim intext As String
   Dim outtext As String
   'jjj

   Open App.Path & "\temp\email-template.htm" For Input As #1
   Open App.Path & "\temp\email-template-default.htm" For Output As #2
   
   Do While Not EOF(1)
      Line Input #1, intext
      'MsgBox intext
      If InStr(1, intext, "http://yourlink") > 0 Then
         outtext = Replace(intext, "http://yourlink", strCompany_Logo_Link)
      Else
         outtext = intext
      End If
      
      Print #2, outtext
   Loop
   
   Close

   'Kill App.Path & "\temp\email-template.htm"
   'Name App.Path & "\temp\email-template-default.htm" As App.Path & "\temp\email-template.htm"

End Sub


Public Function SendMailToUser(ByVal StrUserToAddress As String, ByVal StrSubject As String, ByVal StrBody As String, Optional ShowErrorMessage As Boolean) As Boolean
'************************************************************
' Unit/Module Name  : Mdl_Send_Mail
' Function          : Send Mail
' Date              : -
' Created By        : -
' Last Update/By    :
'************************************************************

        Dim objMail  As New ASPEMAILLib.MailSender
        
                
        On Error Resume Next
        
        'SMTP Server
        If SMTP_Server = "" Then
            If ShowErrorMessage Then MsgBox "SMTP Server has not been configured!", vbExclamation, strApplication_Name
            SendMailToUser = False
            Exit Function
        End If
        
        objMail.Host = SMTP_Server
                
        'From
        objMail.From = "initiative@mail.lowe.co.id"
        objMail.FromName = "Initiative"
        
        'To
        objMail.AddAddress StrUserToAddress
        
        ' message subject
        objMail.Subject = StrSubject
        
        ' message body
        objMail.Body = StrBody
        
        objMail.IsHTML = True
        
        ' send message
        objMail.Send
        
        'Remove Object from memory
        Set objMail = Nothing
        
        If Err.Number <> 0 Then
            If ShowErrorMessage Then MsgBox Err.Description & vbCrLf & vbCrLf & "Error Code : " & Err.Number, vbExclamation, strApplication_Name
            Err.Clear
            SendMailToUser = False
             
                    
            Exit Function
        End If
        
        SendMailToUser = True
        
           
End Function
Private Sub Create_Log_File()
'*************************************************************
'Nama Procedure     : Create_Log_File
'Fungsi             : Membuat file text untuk LOG
'Programer          :
'Tgl Pembuatan      : 14 Juni 2004
'Last Update/By     :
'*************************************************************
    
    On Error GoTo errHand
    
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set FSOInterface = fso.CreateFolder(App.Path & "\temp\Logs\Log")
        
    LogFileName = "SendMail" & Format(Now, "MMDDYY") & "_" & Format(Now, "hhmmss") & ".txt"
    Set FSOInterface = fso.CreateTextFile(App.Path & "\temp\Logs" & LogFileName)
    
    Exit Sub
    
errHand:
    If Err.Number = 58 Then
        Resume Next
    End If
End Sub

Private Sub Close_Log_File()
     '---------- Show Log File & Close Log file object--------------
    Set fso = Nothing
    Set FSOInterface = Nothing
        
    Shell "Notepad.exe C:\MpLog\" & LogFileName, vbNormalNoFocus
        
    LogFileName = ""
End Sub

Private Sub Write_To_Log_File(StrTeks As String)
'*************************************************************
'Nama Procedure     : Write_To_Log_File
'Fungsi             : Menulis ke LOG
'Programer          :
'Tgl Pembuatan      : 14 Juni 2004
'Last Update/By     :
'*************************************************************
    
    FSOInterface.WriteLine StrTeks
End Sub


Private Sub subDelay(sngDelay As Single)
Const cSecondsInDay = 86400        ' # of seconds in a day.

Dim sngStart As Single             ' Start time.
Dim sngStop  As Single             ' Stop time.
Dim sngNow   As Single             ' Current time.

sngStart = Timer                   ' Get current timer.
sngStop = sngStart + sngDelay      ' Set up stop time based on
                                   ' delay.

Do
    sngNow = Timer                 ' Get current timer again.
    If sngNow < sngStart Then      ' Has midnight passed?
        sngStop = sngStart - cSecondsInDay  ' If yes, reset end.
    End If
    DoEvents                       ' Let OS process other events.

Loop While sngNow < sngStop        ' Has time elapsed?

End Sub


Private Sub picButton_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single) 'TOOLBAR_AI.
'************************************************
' Procedure         : picButton_MouseMove
' Function          : TOOLBAR_AI saat mouse berada di area button.
' Created By        : {73 64 6B}
' Date              : 12-Apr-2015
'************************************************
    picButton_Obj Index, Button, Shift, X, Y, picButton
End Sub

Private Sub picButton_Click(Index As Integer)

'************************************************
' Procedure         : picButton_Click
' Function          : Action utk Navigation dan CRUD.
' Created By        : {73 64 6B}
' Date              : 12-Apr-2015/{73 64 6B} --> Semua coding dan query sudah di optimalkan agar faster, readable, safer, standardable.
'************************************************

    Select Case Index
        Case enButtonType.biePreview   'call db_New.
            Call Cmd_Preview_Mail
        Case enButtonType.bieSendEmail
            Call Send_Mail
        Case Else
            Call cmdClose
    End Select

End Sub
