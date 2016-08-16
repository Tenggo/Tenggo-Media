VERSION 5.00
Begin VB.Form Frm_Change_Password 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Change Password"
   ClientHeight    =   3150
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5070
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "Frm_Change_Password.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3150
   ScaleWidth      =   5070
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.PictureBox Picture1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3165
      Left            =   -15
      ScaleHeight     =   3105
      ScaleWidth      =   5025
      TabIndex        =   0
      Top             =   -15
      Width           =   5085
      Begin VB.PictureBox Picture2 
         Height          =   480
         Left            =   1140
         ScaleHeight     =   420
         ScaleWidth      =   2820
         TabIndex        =   10
         Top             =   2388
         Width           =   2880
         Begin VB.CommandButton Cmd_Close 
            Cancel          =   -1  'True
            Caption         =   "&Cancel"
            Height          =   420
            Left            =   1410
            TabIndex        =   5
            Top             =   0
            Width           =   1410
         End
         Begin VB.CommandButton Cmd_Change 
            Caption         =   "&OK"
            Default         =   -1  'True
            Height          =   420
            Left            =   0
            TabIndex        =   4
            Top             =   0
            Width           =   1410
         End
      End
      Begin VB.Frame Frame1 
         Height          =   1776
         Left            =   405
         TabIndex        =   6
         Top             =   330
         Width           =   4290
         Begin VB.TextBox Txt_New_Pwd_Type 
            ForeColor       =   &H000000FF&
            Height          =   315
            IMEMode         =   3  'DISABLE
            Left            =   1848
            PasswordChar    =   "*"
            TabIndex        =   3
            Text            =   "Text1"
            ToolTipText     =   "Re - Type Your New Password"
            Top             =   1224
            Width           =   2055
         End
         Begin VB.TextBox Txt_Old_Pwd 
            Height          =   315
            IMEMode         =   3  'DISABLE
            Left            =   1860
            OLEDropMode     =   1  'Manual
            PasswordChar    =   "*"
            TabIndex        =   1
            Text            =   "Text1"
            ToolTipText     =   "Input Your Old Password"
            Top             =   528
            Width           =   2055
         End
         Begin VB.TextBox Txt_New_Pwd 
            ForeColor       =   &H000000FF&
            Height          =   315
            IMEMode         =   3  'DISABLE
            Left            =   1860
            PasswordChar    =   "*"
            TabIndex        =   2
            Text            =   "Text1"
            ToolTipText     =   "Input Your New Password"
            Top             =   876
            Width           =   2055
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            Caption         =   "User name :"
            Height          =   312
            Left            =   576
            TabIndex        =   12
            Top             =   228
            Width           =   1152
         End
         Begin VB.Label Lbl_User 
            AutoSize        =   -1  'True
            Caption         =   "User :"
            Height          =   192
            Left            =   1836
            TabIndex        =   11
            Top             =   228
            Width           =   480
         End
         Begin VB.Label Label4 
            Alignment       =   1  'Right Justify
            Caption         =   "Confirm New :"
            Height          =   312
            Left            =   336
            TabIndex        =   9
            Top             =   1284
            Width           =   1404
         End
         Begin VB.Label Label3 
            Alignment       =   1  'Right Justify
            Caption         =   "New Password :"
            Height          =   312
            Left            =   324
            TabIndex        =   8
            Top             =   912
            Width           =   1428
         End
         Begin VB.Label Label2 
            Alignment       =   1  'Right Justify
            Caption         =   "Old Password :"
            Height          =   312
            Left            =   324
            TabIndex        =   7
            Top             =   552
            Width           =   1428
         End
      End
   End
End
Attribute VB_Name = "Frm_Change_Password"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'''************************************************
' Form                  : Frm_Change_Password
' Function              : Untuk melakukan perubahan terhadap password user
' Created Date          : 7 aug 2002
' By                    :
' Last Update/BY/Notes  : 25 Agustus 2004/Yayan/Last Update By tidak diubah waktu user Ubah Password
'************************************************
Option Explicit

Const StrHurufKecil As String = "abcdefghijklmnopqrstuvwxyz"
Const StrHurufBesar As String = "ABCDEFGHIJKLMNOPQRSTUVWXYZ"
Const StrAngka As String = "1234567890"
Const StrAlpha As String = "~!@#$%^&*"


Private Sub Cmd_Change_Click()
    'On Error GoTo my_error
    Dim rs_Cek_Old_Pwd As New ADODB.Recordset
    Dim strSql As String
    Dim rsParameter As New ADODB.Recordset
    Dim rsPasswordHistory As New ADODB.Recordset
    Dim IntEnforcePassword As Integer
    Dim intIDX As Integer
    Dim StrTempPass As String
    
    'select parameter length password
    strSql = "SELECT * FROM Media_Parameter WHERE Parameter_Catalog='Password Policy' AND Parameter_Name='Minimum Password Length'"
    rsParameter.Open strSql, ConnERP, adOpenForwardOnly, adLockReadOnly
    If Not rsParameter.EOF Then
        If Len(Trim(Me.Txt_New_Pwd.Text)) < CInt(rsParameter("Value")) Then
            MsgBox "Your password must be at least " & CInt(rsParameter("Value")) & " characters long ", vbInformation, strApplication_Name
            rsParameter.Close
            Set rsParameter = Nothing
            Exit Sub
        End If
    Else
        MsgBox "No Parameter for Length password", vbInformation, strApplication_Name
        rsParameter.Close
        Set rsParameter = Nothing
        Exit Sub
    End If
    Call CloseRecordset(rsParameter)
    
    '====================================
           
    'cek new pass en re-type pass
    If Txt_New_Pwd.Text <> Txt_New_Pwd_Type.Text Then
        MsgBox "The password you typed do not match.", vbExclamation, strApplication_Name
        Txt_New_Pwd_Type.SetFocus
        Exit Sub
    End If
    
    'password gak boleh empty
    If Trim(Txt_New_Pwd.Text) = "" Or Trim(Txt_New_Pwd_Type.Text) = "" Then
        MsgBox "Empty String", vbExclamation, strApplication_Name
        Txt_New_Pwd.SetFocus
        Exit Sub
    End If
    
    '================================================================================
    '                   cek password complexity
    '================================================================================
    Dim IntPosPwd As Integer
    Dim BoolStatus As Boolean
    Dim BoolExistParam As Boolean
    Dim StrMsgComplexity As String
    
    StrMsgComplexity = "The Password supplied dos not meet the minimum complexity requirements." & vbCrLf
    StrMsgComplexity = StrMsgComplexity & "Please select another password that meets all of following criteria:" & vbCrLf
    StrMsgComplexity = StrMsgComplexity & "Contains at least English uppercase  characters (A through Z), " & vbCrLf
    StrMsgComplexity = StrMsgComplexity & "English lowercase  characters (a through z), Numerals (0 through 9), " & vbCrLf
    StrMsgComplexity = StrMsgComplexity & "Non-alphabetical characters (such as !$#%)."
    
    strSql = "SELECT rtrim(ltrim(Value)) FROM Media_Parameter WHERE Parameter_Catalog='Password Policy' AND Parameter_Name='Password Complexity'"
    rsParameter.Open strSql, ConnERP, adOpenForwardOnly, adLockReadOnly
    If Not rsParameter.EOF Then
        intIDX = 1
        Do
            If Trim(Mid(rsParameter(0), intIDX, 1)) = "" Then Exit Do
            BoolExistParam = False
            '===================================
            'jika ada huruf kecil di parameter
            '===================================
            If InStr(1, StrHurufKecil, Mid(rsParameter(0), intIDX, 1)) > 0 Then
                If Trim(Mid(rsParameter(0), intIDX, 1)) = "" Then Exit Do
                BoolStatus = False
                BoolExistParam = True
                IntPosPwd = 1
                Do
                    If InStr(1, StrHurufKecil, Mid(Trim(Txt_New_Pwd.Text), IntPosPwd, 1)) > 0 Then
                        'terpenuhi
                        IntPosPwd = Len(Trim(Txt_New_Pwd.Text))
                        BoolStatus = True
                    End If
                    IntPosPwd = IntPosPwd + 1
                Loop Until IntPosPwd > Len(Trim(Txt_New_Pwd.Text))
                
                If Not BoolStatus Then
                    'tidak ada huruf kecil
                    'MsgBox "Your password must contain characters : a - z", vbCritical, "Password Complexity"
                    MsgBox StrMsgComplexity, vbExclamation, strApplication_Name
                    Exit Sub
                End If
                
                intIDX = intIDX + 1
            End If
            
            '===================================
            'jika ada huruf besar di parameter
            '===================================
            If InStr(1, StrHurufBesar, Mid(rsParameter(0), intIDX, 1)) > 0 Then
                If Trim(Mid(rsParameter(0), intIDX, 1)) = "" Then Exit Do
                BoolStatus = False
                BoolExistParam = True
                IntPosPwd = 1
                Do
                    If InStr(1, StrHurufBesar, Mid(Trim(Txt_New_Pwd.Text), IntPosPwd, 1)) > 0 Then
                        'terpenuhi
                        IntPosPwd = Len(Trim(Txt_New_Pwd.Text))
                        BoolStatus = True
                    End If
                    IntPosPwd = IntPosPwd + 1
                Loop Until IntPosPwd > Len(Trim(Txt_New_Pwd.Text))
                
                If Not BoolStatus Then
                    'tidak ada huruf besar
                    'MsgBox "Your password must contain characters : A - Z", vbCritical, "Password Complexity"
                    MsgBox StrMsgComplexity, vbExclamation, strApplication_Name
                    Exit Sub
                End If
                
                intIDX = intIDX + 1
            End If
            
            '===================================
            'jika ada angka di parameter
            '===================================
            If InStr(1, StrAngka, Mid(rsParameter(0), intIDX, 1)) > 0 Then
                If Trim(Mid(rsParameter(0), intIDX, 1)) = "" Then Exit Do
                BoolStatus = False
                BoolExistParam = True
                IntPosPwd = 1
                Do
                    If InStr(1, StrAngka, Mid(Trim(Txt_New_Pwd.Text), IntPosPwd, 1)) > 0 Then
                        'terpenuhi
                        IntPosPwd = Len(Trim(Txt_New_Pwd.Text))
                        BoolStatus = True
                    End If
                    IntPosPwd = IntPosPwd + 1
                Loop Until IntPosPwd > Len(Trim(Txt_New_Pwd.Text))
                
                If Not BoolStatus Then
                    'tidak ada angka
                    'MsgBox "Your password must contain characters : 0 - 9", vbCritical, "Password Complexity"
                    MsgBox StrMsgComplexity, vbExclamation, strApplication_Name
                    Exit Sub
                End If
                
                intIDX = intIDX + 1
            End If
            '==============================================
            'jika ada Special Charecter di parameter
            '===============================================
            If InStr(1, StrAlpha, Mid(rsParameter(0), intIDX, 1)) > 0 Then
                If Trim(Mid(rsParameter(0), intIDX, 1)) = "" Then Exit Do
                BoolStatus = False
                BoolExistParam = True
                IntPosPwd = 1
                Do
                    If InStr(1, StrAlpha, Mid(Trim(Txt_New_Pwd.Text), IntPosPwd, 1)) > 0 Then
                        'terpenuhi
                        IntPosPwd = Len(Trim(Txt_New_Pwd.Text))
                        BoolStatus = True
                    End If
                    IntPosPwd = IntPosPwd + 1
                Loop Until IntPosPwd > Len(Trim(Txt_New_Pwd.Text))
                
                If Not BoolStatus Then
                    'tidak ada alpha numeric
                    'MsgBox "Your password must contain characters : ~,!,#,$,%,^,&,*", vbCritical, "Password Complexity"
                    MsgBox StrMsgComplexity, vbExclamation, strApplication_Name
                    Exit Sub
                End If
                
                intIDX = intIDX + 1
            End If
            '===================================
            
            If Not BoolExistParam Then intIDX = intIDX + 1
        Loop Until intIDX > Len(rsParameter(0))
        
    End If
    rsParameter.Close
    Set rsParameter = Nothing
    '================================================================================
    '                  End cek password complexity
    '================================================================================
    
    strSql = "SELECT Password, isnull(Change_Password_Date,'') Change_Password_Date FROM user_id WHERE User_Name='" & strLogin_User & "'"
    rs_Cek_Old_Pwd.Open strSql, ConnERP, adOpenStatic, adLockReadOnly
    If Not rs_Cek_Old_Pwd.EOF Then
        If Trim(Txt_Old_Pwd.Text) <> Decrypt(rs_Cek_Old_Pwd("Password")) Then
            MsgBox "Invalid Password for this user", vbExclamation, strApplication_Name
            Txt_Old_Pwd.SetFocus
            rs_Cek_Old_Pwd.Close
            Set rs_Cek_Old_Pwd = Nothing
            Exit Sub
        End If
            
        IntEnforcePassword = 0
        'select parameter enforce password history
        strSql = "SELECT isnull(value,0) Value FROM Media_Parameter WHERE Parameter_Catalog='Password Policy' AND Parameter_Name='Enforce Password History'"
        rsParameter.Open strSql, ConnERP, adOpenForwardOnly, adLockReadOnly
        If Not rsParameter.EOF Then
            IntEnforcePassword = CInt(rsParameter("Value"))
        End If
        rsParameter.Close
        Set rsParameter = Nothing
        '====================================
        If IntEnforcePassword = 0 Then
            'do nothing
        Else
            'Check New Password dengan Password Aktual
            If Trim(Txt_Old_Pwd.Text) = Trim(Txt_New_Pwd.Text) Then
                MsgBox "Your new password can not be the same as any of your previous " & IntEnforcePassword & " password", vbExclamation, strApplication_Name
                Exit Sub
            End If
            
            'cek apakah new password udah ada di history
            strSql = "SELECT * FROM Password_History WHERE User_Name='" & strLogin_User & "' Order By Password_No "
            rsPasswordHistory.Open strSql, ConnERP, adOpenForwardOnly, adLockReadOnly
            While Not rsPasswordHistory.EOF
                If Me.Txt_New_Pwd.Text = Decrypt(rsPasswordHistory("Password")) Then
                    MsgBox "Your new password can not be the same as any of your previous " & IntEnforcePassword & " password", vbExclamation, strApplication_Name
                    rsPasswordHistory.Close
                    Set rsPasswordHistory = Nothing
                    
                    rs_Cek_Old_Pwd.Close
                    Set rs_Cek_Old_Pwd = Nothing
                    Exit Sub
                End If
                rsPasswordHistory.MoveNext
            Wend
            rsPasswordHistory.Close
            Set rsPasswordHistory = Nothing
        End If
        
            
        'select parameter minimum password age
        If Txt_Old_Pwd.Text <> "" Then
            strSql = "SELECT * FROM Media_Parameter WHERE Parameter_Catalog='Password Policy' AND Parameter_Name='Minimum Password Age'"
            rsParameter.Open strSql, ConnERP, adOpenForwardOnly, adLockReadOnly
            If Not rsParameter.EOF Then
                If rs_Cek_Old_Pwd("Change_Password_Date") = "" Then
                    'do nothing
                Else
                    recDate.Requery
                    If recDate(0) <= DateAdd("d", CInt(rsParameter("Value")), rs_Cek_Old_Pwd("Change_Password_Date")) Then
                        MsgBox "You Password can not be changed at this time", vbExclamation, strApplication_Name
                        rsParameter.Close
                        Set rsParameter = Nothing
                        
                        rs_Cek_Old_Pwd.Close
                        Set rs_Cek_Old_Pwd = Nothing
                        Exit Sub
                    End If
                End If
            Else
                MsgBox "No Parameter for minimum age password", vbInformation, strApplication_Name
                rsParameter.Close
                Set rsParameter = Nothing
                
                rs_Cek_Old_Pwd.Close
                Set rs_Cek_Old_Pwd = Nothing
                Exit Sub
            End If
            rsParameter.Close
            Set rsParameter = Nothing
        End If
        '====================================
            
        ConnERP.BeginTrans
        
        'update tabel user id
        recDate.Requery
        strSql = "UPDATE user_id SET Password='" & Encrypt(Trim(Txt_New_Pwd.Text)) & "', "
        strSql = strSql & " Change_Password_date='" & recDate(0) & "' "
        strSql = strSql & " WHERE User_Name='" & strLogin_User & "'"
        ConnERP.Execute strSql
        
        'update tabel history
        If Txt_Old_Pwd.Text <> "" Then
            If IntEnforcePassword > 0 Then
                'make recordset temp for password history
                Dim rsPasswordHistoryTemp As New ADODB.Recordset
                
                Set rsPasswordHistoryTemp = Nothing
                Set rsPasswordHistoryTemp = New ADODB.Recordset
                
                With rsPasswordHistoryTemp.Fields
                    .Append "User_Name", adVarChar, 15, adFldMayBeNull
                    .Append "Password_No", adInteger, , adFldMayBeNull
                    .Append "Password", adVarChar, 200, adFldMayBeNull
                End With
                rsPasswordHistoryTemp.CursorLocation = adUseClient
                rsPasswordHistoryTemp.Open , , adOpenDynamic, adLockOptimistic
                
                strSql = "SELECT * FROM Password_History WHERE User_Name='" & strLogin_User & "' ORDER BY Password_No"
                rsPasswordHistory.Open strSql, ConnERP, adOpenForwardOnly, adLockReadOnly
                            
                If Not rsPasswordHistory.EOF Then
                    rsPasswordHistory.MoveFirst
                    While Not rsPasswordHistory.EOF
                        rsPasswordHistoryTemp.AddNew
                        rsPasswordHistoryTemp.Fields("User_Name").Value = rsPasswordHistory.Fields("User_Name").Value
                        rsPasswordHistoryTemp.Fields("Password_No").Value = rsPasswordHistory.Fields("Password_No").Value
                        rsPasswordHistoryTemp.Fields("Password").Value = rsPasswordHistory.Fields("Password").Value
                        rsPasswordHistoryTemp.Update
                        rsPasswordHistory.MoveNext
                    Wend
                    
                    intIDX = 0
                    StrTempPass = ""
                    rsPasswordHistory.MoveFirst
                    While Not rsPasswordHistory.EOF
                        intIDX = intIDX + 1
                        
                        If intIDX = 1 Then
                            'update rec
                            rsPasswordHistoryTemp.Filter = "User_Name ='" & rsPasswordHistory("User_Name") & "' AND Password_No ='" & rsPasswordHistory("Password_No") & "'"
                            If Not rsPasswordHistoryTemp.EOF Then
                                rsPasswordHistoryTemp.Fields("User_Name").Value = strLogin_User
                                rsPasswordHistoryTemp.Fields("Password_No").Value = intIDX
                                rsPasswordHistoryTemp.Fields("Password").Value = Encrypt(Txt_Old_Pwd.Text)
                                rsPasswordHistoryTemp.Update
                            End If
                        Else
                            'update rec
                            
                                rsPasswordHistoryTemp.Filter = "User_Name ='" & rsPasswordHistory("User_Name") & "' AND Password_No ='" & rsPasswordHistory("Password_No") & "'"
                                If Not rsPasswordHistoryTemp.EOF Then
                                    rsPasswordHistoryTemp.Fields("User_Name").Value = strLogin_User
                                    rsPasswordHistoryTemp.Fields("Password_No").Value = intIDX
                                    rsPasswordHistoryTemp.Fields("Password").Value = StrTempPass
                                    rsPasswordHistoryTemp.Update
                                End If
                        End If
                        
                        'simpen ke temp
                        StrTempPass = rsPasswordHistory("Password")
                        
                        rsPasswordHistory.MoveNext
                    Wend
                    
                    
                    'delete data password history history
                    strSql = "DELETE FROM Password_History "
                    strSql = strSql & " WHERE User_Name='" & strLogin_User & "'"
                    ConnERP.Execute strSql
                     
                    'insert to password history
                    rsPasswordHistoryTemp.Filter = ""
                    rsPasswordHistoryTemp.MoveFirst
                    While Not rsPasswordHistoryTemp.EOF
                        strSql = "INSERT INTO Password_History(User_Name, Password_No, Password) VALUES ('" & rsPasswordHistoryTemp.Fields("User_Name").Value & "'," & rsPasswordHistoryTemp.Fields("Password_No").Value & ",'" & rsPasswordHistoryTemp.Fields("Password").Value & "')"
                        ConnERP.Execute strSql
                        rsPasswordHistoryTemp.MoveNext
                    Wend
                    rsPasswordHistoryTemp.Close
                    Set rsPasswordHistoryTemp = Nothing
                    
    '                IntIdx = 0
    '                StrTempPass = ""
    '                While Not rsPasswordHistory.EOF
    '                    IntIdx = IntIdx + 1
    '                    If IntIdx = 1 Then
    '                        'update rec
    '                        strSql = "UPDATE Password_History SET Password='" & EncryptPassword(Txt_Old_Pwd.Text) & "'"
    '                        strSql = strSql & " WHERE User_Name='" & username & "' AND Password_No = " & IntIdx
    '                        ConnERP.Execute strSql
    '                    Else
    '                        'update rec
    '                        strSql = "UPDATE Password_History SET Password='" & StrTempPass & "'"
    '                        strSql = strSql & " WHERE User_Name='" & username & "' AND Password_No = " & IntIdx
    '                        ConnERP.Execute strSql
    '                    End If
    '
    '                    'simpen ke temp
    '                    StrTempPass = rsPasswordHistory("Password")
    '
    '                    rsPasswordHistory.MoveNext
    '                Wend
                    
                    'jika jumlah password history masih < enforce
                    If intIDX < IntEnforcePassword Then
                        strSql = "INSERT INTO Password_History(User_Name, Password_No, Password) VALUES ('" & strLogin_User & "'," & intIDX + 1 & ",'" & StrTempPass & "')"
                        ConnERP.Execute strSql
                    End If
                    
                Else
                    'jika belum ada history
                    'insert ke history
                    strSql = "INSERT INTO Password_History(User_Name, Password_No, Password) VALUES ('" & strLogin_User & "',1,'" & Encrypt(Txt_Old_Pwd.Text) & "')"
                    ConnERP.Execute strSql
                End If
                rsPasswordHistory.Close
                Set rsPasswordHistory = Nothing
            Else
                'jika enforce <= 0 maka data history di hapus
                strSql = "DELETE FROM Password_History WHERE User_Name='" & strLogin_User & "'"
                ConnERP.Execute strSql
            End If
        End If
        ConnERP.CommitTrans
        MsgBox "Your password has been changed.", vbExclamation, strApplication_Name
        Txt_Old_Pwd.Text = ""
        Txt_New_Pwd.Text = ""
        Txt_New_Pwd_Type.Text = ""
        Txt_Old_Pwd.SetFocus
        
        If BoolMaxAgePassword Then
            BoolMaxAgePassword = False
            Unload Me
        Else
            Unload Me
        End If
        
    Else
        MsgBox "Invalid User", vbExclamation, strApplication_Name
        Txt_Old_Pwd.SetFocus
        rs_Cek_Old_Pwd.Close
        Set rs_Cek_Old_Pwd = Nothing
        Exit Sub
    End If
    
    rs_Cek_Old_Pwd.Close
    Set rs_Cek_Old_Pwd = Nothing
    
    Exit Sub
my_error:
    ConnERP.RollbackTrans
    MsgBox Err.Description, vbExclamation, strApplication_Name
End Sub

Private Sub Cmd_Close_Click()
    If BoolMaxAgePassword Then
        BoolMaxAgePassword = False
        Dim rs As New ADODB.Recordset
        Set rs = New ADODB.Recordset
        rs.Open "select * from user_id where user_name='" & strLogin_User & "'", ConnERP, adOpenKeyset, adLockOptimistic, adCmdText
                    
        If Frm_Login.lbl_Check_User.Caption = "." Then
            rs("Computer_Name").Value = Null
        End If
        rs.Update
        rs.Close
        End
    Else
        Unload Me
    End If
    
End Sub

Private Sub Form_Activate()
    Txt_Old_Pwd.SetFocus
End Sub

Private Sub Form_Load()
  RemoveMenus Me, True
  If BoolMaxAgePassword Then
        MsgBox StrPasswordMassage, vbExclamation, strApplication_Name
  End If
  
  Lbl_User.Caption = strLogin_User
  Txt_Old_Pwd.Text = ""
  
  Txt_New_Pwd.Text = ""
  Txt_New_Pwd_Type.Text = ""
   
End Sub



Private Sub Txt_New_Pwd_GotFocus()
    Txt_New_Pwd.SelStart = 0
    Txt_New_Pwd.SelLength = Len(Txt_New_Pwd.Text)
End Sub

Private Sub Txt_New_Pwd_KeyPress(KeyAscii As Integer)
'    If (KeyAscii >= Asc("A") And KeyAscii <= Asc("Z")) Or (KeyAscii >= Asc("a") And KeyAscii <= Asc("z")) Or (KeyAscii >= Asc("0") And KeyAscii <= Asc("9")) Or KeyAscii = 32 Or (KeyAscii >= 24 And KeyAscii <= 27) Or KeyAscii = 13 Or KeyAscii = 8 Or KeyAscii = 27 Then
'    Else
'        KeyAscii = 0
'        Beep
'    End If
    If KeyAscii = 39 Then
        KeyAscii = 0
        Beep
    End If
End Sub

Private Sub Txt_New_Pwd_Type_GotFocus()
    Txt_New_Pwd_Type.SelStart = 0
    Txt_New_Pwd_Type.SelLength = Len(Txt_New_Pwd_Type.Text)
End Sub

Private Sub Txt_New_Pwd_Type_KeyPress(KeyAscii As Integer)
'    If (KeyAscii >= Asc("A") And KeyAscii <= Asc("Z")) Or (KeyAscii >= Asc("a") And KeyAscii <= Asc("z")) Or (KeyAscii >= Asc("0") And KeyAscii <= Asc("9")) Or KeyAscii = 32 Or (KeyAscii >= 24 And KeyAscii <= 27) Or KeyAscii = 13 Or KeyAscii = 8 Or KeyAscii = 27 Then
'    Else
'        KeyAscii = 0
'        Beep
'    End If
    If KeyAscii = 39 Then
        KeyAscii = 0
        Beep
    End If
End Sub

Private Sub Txt_Old_Pwd_GotFocus()
    Txt_Old_Pwd.SelStart = 0
    Txt_Old_Pwd.SelLength = Len(Txt_Old_Pwd.Text)
End Sub

Private Sub Txt_Old_Pwd_KeyPress(KeyAscii As Integer)
'    If (KeyAscii >= Asc("A") And KeyAscii <= Asc("Z")) Or (KeyAscii >= Asc("a") And KeyAscii <= Asc("z")) Or (KeyAscii >= Asc("0") And KeyAscii <= Asc("9")) Or KeyAscii = 32 Or (KeyAscii >= 24 And KeyAscii <= 27) Or KeyAscii = 13 Or KeyAscii = 8 Or KeyAscii = 27 Then
'
'    Else
'        KeyAscii = 0
'        Beep
'    End If

    If KeyAscii = 39 Then
        KeyAscii = 0
        Beep
    End If
End Sub
