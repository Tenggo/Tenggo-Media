VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "mscomctl.OCX"
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.Form Frm_IB_Radio_Plan_Area_New 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Edit Area"
   ClientHeight    =   5190
   ClientLeft      =   45
   ClientTop       =   405
   ClientWidth     =   7830
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5190
   ScaleWidth      =   7830
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox picToolbar 
      Align           =   1  'Align Top
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ClipControls    =   0   'False
      ForeColor       =   &H80000008&
      Height          =   750
      Left            =   0
      ScaleHeight     =   750
      ScaleWidth      =   7830
      TabIndex        =   24
      Top             =   0
      Width           =   7830
      Begin VB.PictureBox picButton 
         Appearance      =   0  'Flat
         BackColor       =   &H00F0F0F0&
         BorderStyle     =   0  'None
         ClipControls    =   0   'False
         ForeColor       =   &H80000008&
         Height          =   750
         Index           =   10
         Left            =   4680
         ScaleHeight     =   750
         ScaleWidth      =   1500
         TabIndex        =   30
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
         Index           =   11
         Left            =   10800
         ScaleHeight     =   750
         ScaleWidth      =   1500
         TabIndex        =   29
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
         Index           =   4
         Left            =   90
         ScaleHeight     =   750
         ScaleWidth      =   1500
         TabIndex        =   28
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
         Index           =   6
         Left            =   3150
         ScaleHeight     =   750
         ScaleWidth      =   1500
         TabIndex        =   27
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
         Index           =   5
         Left            =   1620
         ScaleHeight     =   750
         ScaleWidth      =   1500
         TabIndex        =   26
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
         Left            =   4680
         ScaleHeight     =   750
         ScaleWidth      =   1500
         TabIndex        =   25
         Top             =   0
         Width           =   1500
      End
   End
   Begin VB.ComboBox cboAreaCode 
      Height          =   315
      Left            =   1440
      TabIndex        =   23
      Top             =   1245
      Width           =   1755
   End
   Begin VB.TextBox txtAreaName 
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
      Left            =   1455
      MaxLength       =   15
      TabIndex        =   22
      Top             =   1245
      Width           =   1740
   End
   Begin VB.ComboBox cboBrand 
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
      Left            =   1425
      TabIndex        =   21
      Top             =   870
      Width           =   3885
   End
   Begin Threed.SSPanel pnlMain 
      Align           =   1  'Align Top
      Height          =   4365
      Left            =   0
      TabIndex        =   0
      Top             =   750
      Width           =   7830
      _Version        =   65536
      _ExtentX        =   13811
      _ExtentY        =   7699
      _StockProps     =   15
      BackColor       =   15790320
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BevelOuter      =   1
      Begin VB.Frame fraCities 
         Appearance      =   0  'Flat
         Caption         =   "Cities"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   3090
         Left            =   90
         TabIndex        =   14
         Top             =   975
         Width           =   3510
         Begin VB.ListBox lstCity 
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   2400
            Left            =   105
            TabIndex        =   16
            Top             =   315
            Visible         =   0   'False
            Width           =   3300
         End
         Begin MSComctlLib.TreeView treCity 
            Height          =   2640
            Left            =   105
            TabIndex        =   15
            Top             =   330
            Width           =   3300
            _ExtentX        =   5821
            _ExtentY        =   4657
            _Version        =   393217
            Style           =   7
            Appearance      =   1
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
      End
      Begin VB.Frame fraSelectedCities 
         Appearance      =   0  'Flat
         Caption         =   "Selected Cities"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   3060
         Left            =   4245
         TabIndex        =   12
         Top             =   975
         Width           =   3510
         Begin VB.ListBox lstSelected 
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   2595
            Left            =   105
            TabIndex        =   13
            Top             =   330
            Width           =   3285
         End
      End
      Begin VB.CommandButton cmdRemove 
         Caption         =   "<"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   3645
         TabIndex        =   11
         ToolTipText     =   "Remove Sales Area Cities"
         Top             =   2655
         Width           =   510
      End
      Begin VB.CommandButton cmdInsert 
         Caption         =   ">"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   3645
         TabIndex        =   10
         ToolTipText     =   "Add Sales Area Cities"
         Top             =   2055
         Width           =   510
      End
      Begin VB.CommandButton cmdAddAll 
         Caption         =   ">>"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   3645
         TabIndex        =   9
         ToolTipText     =   "Add All Sales Area Cities"
         Top             =   1455
         Width           =   510
      End
      Begin VB.CommandButton cmdRemoveAll 
         Caption         =   "<<"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   3645
         TabIndex        =   8
         ToolTipText     =   "Remove All Sales Area Cities"
         Top             =   3255
         Width           =   510
      End
      Begin VB.CommandButton cmdAdd 
         Caption         =   "&Add"
         Height          =   495
         Left            =   120
         TabIndex        =   6
         Top             =   4575
         Visible         =   0   'False
         Width           =   1180
      End
      Begin VB.CommandButton cmdEdit 
         Caption         =   "&Edit"
         Height          =   495
         Left            =   1305
         TabIndex        =   5
         Top             =   4575
         Visible         =   0   'False
         Width           =   1180
      End
      Begin VB.CommandButton cmdClose 
         Caption         =   "Cl&ose"
         Height          =   495
         Left            =   5430
         TabIndex        =   3
         Top             =   4575
         Visible         =   0   'False
         Width           =   1180
      End
      Begin VB.CheckBox chkRural 
         Caption         =   "Rural Flag"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   3330
         TabIndex        =   2
         Top             =   540
         Width           =   1215
      End
      Begin VB.CheckBox chkUrban 
         Caption         =   "Urban Flag"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   4635
         TabIndex        =   1
         Top             =   540
         Width           =   1215
      End
      Begin ComctlLib.ProgressBar prgLoadCity 
         Height          =   120
         Left            =   105
         TabIndex        =   7
         Top             =   4140
         Width           =   7620
         _ExtentX        =   13441
         _ExtentY        =   212
         _Version        =   327682
         Appearance      =   1
      End
      Begin VB.CommandButton cmdDelete 
         Caption         =   "&Delete"
         Height          =   495
         Left            =   2490
         TabIndex        =   4
         Top             =   4575
         Visible         =   0   'False
         Width           =   1180
      End
      Begin VB.CommandButton cmdSave 
         Caption         =   "&Save"
         Height          =   495
         Left            =   120
         TabIndex        =   17
         Top             =   4575
         Visible         =   0   'False
         Width           =   1180
      End
      Begin VB.CommandButton cmdCancel 
         Caption         =   "&Cancel"
         Height          =   495
         Left            =   1305
         TabIndex        =   18
         Top             =   4575
         Visible         =   0   'False
         Width           =   1180
      End
      Begin VB.Label lblBrand 
         BackStyle       =   0  'Transparent
         Caption         =   "Brand "
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
         Height          =   255
         Left            =   225
         TabIndex        =   20
         Top             =   105
         Width           =   1170
      End
      Begin VB.Label lblAreaName 
         BackStyle       =   0  'Transparent
         Caption         =   "Area Name "
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
         Height          =   255
         Left            =   225
         TabIndex        =   19
         Top             =   525
         Width           =   1170
      End
   End
End
Attribute VB_Name = "Frm_IB_Radio_Plan_Area_New"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
''************************************************
' Function          : generate template untuk area IB radio
' Last Update       :
'*************************'***********************
Dim blnAddNewFlag As Boolean
Dim blnEditingFlag As Boolean
Dim recCityAreaTemp As New ADODB.Recordset

Private Sub Form_Load()
    LoadBrand cboBrand, strLogin_User, "'planner'"
    
    Call SetButton(False)
    
    If cboAreaCode.ListIndex <> -1 Then
        cmdEdit.Enabled = True
        cmdDelete.Enabled = True
    Else
        cmdEdit.Enabled = False
        cmdDelete.Enabled = False
    End If
    
    blnEditingFlag = False
    
    prgLoadCity.Visible = False
    
    cboBrand.Text = Frm_IB_Radio.cboBrand.Text
    
    Call GetAreaCode
    EnableObject False
    
End Sub

Private Sub LoadCityTV()
    Dim strParent As String
    Dim recQryArea As New ADODB.Recordset
    Dim recQryCity As New ADODB.Recordset
    Dim recCitySelected As New ADODB.Recordset
    
    strParent = "Indonesia"
    
    strQuery = " SELECT * FROM radio_Area WHERE Radio_Area_Code='" & Trim(txtAreaName.Text) & "' AND Brand_Code ='" & Left(cboBrand.Text, 4) & "'"
    recCitySelected.Open strQuery, ConnERP, adOpenStatic, adLockReadOnly

    lstSelected.Clear
    treCity.Nodes.Clear
    treCity.Nodes.Add , , strParent, "Indonesia"
    treCity.Nodes(1).Expanded = True
        
    strQuery = " SELECT * FROM area "
    recQryArea.Open strQuery, ConnERP, adOpenStatic, adLockReadOnly
    
    prgLoadCity.Max = recQryArea.RecordCount
    prgLoadCity.Value = 0
    
    While Not recQryArea.EOF And Not recQryArea.BOF
        treCity.Nodes.Add strParent, tvwChild, Trim(recQryArea("area_name")), recQryArea("area_name")
        strQuery = "SELECT * FROM city WHERE area_id=" & recQryArea("area_id") & " ORDER BY City"
        recQryCity.Open strQuery, ConnERP, adOpenStatic, adLockReadOnly
        
        While Not recQryCity.EOF And Not recQryCity.BOF
            recCitySelected.Filter = "city_id=" & recQryCity("city_id")
            
            If Not recCitySelected.EOF And Not recCitySelected.BOF Then
                lstSelected.AddItem recQryCity("city_id") & " -> " & recQryCity("city")
                chkRural.Value = IIf(IsNull(recCitySelected("Rural_flag")), 0, recCitySelected("Rural_flag"))
                chkUrban.Value = IIf(IsNull(recCitySelected("urban_flag")), 0, recCitySelected("urban_flag"))
            Else
                treCity.Nodes.Add Trim(recQryArea("area_name")), tvwChild, Trim(recQryCity("city_id") & " -> " & recQryCity("city")), recQryCity("city_id") & " -> " & recQryCity("city")
            End If
            
            recCitySelected.Filter = ""
            recQryCity.MoveNext
        Wend
        
        recQryCity.Close
        Set recQryCity = Nothing
        
        recQryArea.MoveNext
        prgLoadCity.Value = prgLoadCity.Value + 1
    Wend
    
    recQryArea.Close
    Set recQryArea = Nothing
    
    recCitySelected.Close
    Set recCitySelected = Nothing
End Sub

Private Sub cboAreaCode_Click()
    prgLoadCity.Visible = True
    txtAreaName.Text = cboAreaCode.Text
    
    If cboAreaCode.ListIndex <> -1 Then
        LoadCityTV
        
        cmdEdit.Enabled = True
        cmdDelete.Enabled = True
        'cmdEdit.SetFocus
    Else
        cmdEdit.Enabled = False
        cmdDelete.Enabled = False
    End If
    
    prgLoadCity.Visible = False
End Sub

Private Sub cboAreaCode_DropDown()
    GetAreaCode
End Sub

Private Sub cboAreaCode_KeyPress(KeyAscii As Integer)
    KeyAscii = 0
End Sub

Private Sub cboBrand_Click()
    If cboBrand.ListIndex <> -1 Then
        treCity.Nodes.Clear
        lstSelected.Clear
        
        Call GetAreaCode
        
        If cboAreaCode.ListIndex <> -1 Then
            cmdEdit.Enabled = True
            cmdDelete.Enabled = True
        Else
            cmdEdit.Enabled = False
            cmdDelete.Enabled = False
        End If
    End If
End Sub

Private Sub cboBrand_KeyPress(KeyAscii As Integer)
    KeyAscii = 0
End Sub

Private Sub cmdAddAll_Click()
    lstSelected.Clear
    treCity.Nodes.Clear
    
    ClearNode
    
    InsertCityToList
End Sub

Private Sub InsertCityToList()
    Dim recLoadCity As New ADODB.Recordset
    
    strQuery = "SELECT * FROM city ORDER BY area_id"
    recLoadCity.Open strQuery, ConnERP, adOpenStatic, adLockReadOnly
    
    While Not recLoadCity.EOF And Not recLoadCity.BOF
        lstSelected.AddItem recLoadCity("city_id") & " -> " & recLoadCity("city")
        recLoadCity.MoveNext
    Wend
    
    recLoadCity.Close
    Set recLoadCity = Nothing
End Sub

Private Sub ClearNode()
    Dim strParent As String
    Dim recQryArea As New ADODB.Recordset

    strParent = "Indonesia"
    treCity.Nodes.Add , , strParent, "Indonesia"
    treCity.Nodes(1).Expanded = True
    
    strQuery = "SELECT * FROM area "
    recQryArea.Open strQuery, ConnERP, adOpenStatic, adLockReadOnly
    
    While Not recQryArea.EOF And Not recQryArea.BOF
        treCity.Nodes.Add strParent, tvwChild, Trim(recQryArea("area_name")), recQryArea("area_name")
        recQryArea.MoveNext
    Wend
    
    recQryArea.Close
    Set recQryArea = Nothing
End Sub

Private Sub cmdAdd_Click()
    EnableObject True
    txtAreaName.Text = Empty
    txtAreaName.Enabled = True
    txtAreaName.SetFocus
    
    chkUrban.Value = 1
    chkRural.Value = 1
    
    Call LoadCityTV
    
    Call SetButton(True)
    
    blnEditingFlag = True
    
    blnAddNewFlag = True
    
    
End Sub

Private Sub cmdCancel_Click()
    
    Call SetButton(False)
    
    Call LoadCityTV
    EnableObject False
    
End Sub

Private Sub cmdClose_Click()
    Unload Me
End Sub

Private Sub cmdDelete_Click()
    If MsgBox(strMsgDeleteConfirm, vbQuestion + vbYesNo, strTitleConfirm) = vbYes Then
        strQuery = "DELETE FROM radio_area WHERE brand_code ='" & Left(Trim(cboBrand.Text), 4) & "' AND Radio_Area_Code ='" & Trim(txtAreaName.Text) & "'"
        ConnERP.Execute strQuery
        
        Call GetAreaCode
        
        Call LoadCityTV
        
        If cboAreaCode.ListCount > 0 Then
            cboAreaCode.ListIndex = 0
        End If
        
        Me.MousePointer = vbDefault
        MsgBox strMsgDeleteDataDone, vbInformation, strTitleInfo
        

    End If
End Sub

Private Sub cmdEdit_Click()
    
    If cboAreaCode.Text = "" Then
        MsgBox "Please area code!", vbCritical, strApplication_Name
        Exit Sub
    End If

    EnableObject True
    Call SetButton(True)
    
    txtAreaName.Enabled = False
    
    blnEditingFlag = True
    
    blnAddNewFlag = False
    
    
End Sub

Private Sub cmdInsert_Click()
    Dim intPosisi As Integer
    
    intPosisi = InStr(1, treCity.Nodes(Trim(treCity.SelectedItem)), "->")
    
    If intPosisi > 0 Then
        lstSelected.AddItem treCity.SelectedItem
        treCity.Nodes.Remove Trim(treCity.SelectedItem)
    End If
End Sub

Private Sub CmdRemoveAll_Click()
    Dim recCariCity As New ADODB.Recordset
    
    Do While lstSelected.ListCount > 0
        lstSelected.ListIndex = 0
        
        strQuery = "SELECT a.area_name FROM area a, city b WHERE a.area_id=b.area_id AND b.city_id=" & Val(lstSelected.Text)
        recCariCity.Open strQuery, ConnERP, adOpenStatic, adLockReadOnly
    
        If Not recCariCity.EOF And Not recCariCity.BOF Then
            treCity.Nodes.Add Trim(recCariCity("area_name")), tvwChild, Trim(lstSelected.Text), lstSelected.Text
            lstSelected.RemoveItem 0
        End If
        
        recCariCity.Close
        Set recCariCity = Nothing
    Loop
End Sub

Private Sub CmdRemove_Click()
    Dim recCariCity As New ADODB.Recordset
    
    If lstSelected.ListIndex <> -1 Then
        strQuery = "SELECT a.area_name FROM area a, city b WHERE a.area_id=b.area_id AND b.city_id=" & Val(lstSelected.Text)
        recCariCity.Open strQuery, ConnERP, adOpenStatic, adLockReadOnly
        
        If Not recCariCity.EOF And Not recCariCity.BOF Then
            treCity.Nodes.Add Trim(recCariCity("area_name")), tvwChild, Trim(lstSelected.Text), lstSelected.Text
            lstSelected.RemoveItem lstSelected.ListIndex
        End If
    End If
End Sub

Private Sub cmdSave_Click()
    If chkRural.Value = 0 And chkUrban.Value = 0 Then
        MsgBox "Select Rural or Urban Flag First", vbExclamation, strTitleMissingInfo
        Exit Sub
    End If
    
    If Len(Trim(txtAreaName.Text)) > 0 Then
        If lstSelected.ListCount > 0 Then
            If blnAddNewFlag = True Then
                If blnCekRadioAreaCode = True Then
                    txtAreaName.BackColor = vbRed
                    txtAreaName.ForeColor = vbWhite
                    MsgBox "Enter a new Area Name ", vbCritical, strTitleMissingInfo
                    txtAreaName.ForeColor = vbWindowText
                    txtAreaName.BackColor = vbWindowBackground
                Else
                    prgLoadCity.Visible = True
                    
                    Call SaveData
                    
                    Call SetButton(False)
                    
                    prgLoadCity.Visible = False
                    
                    Call cboBrand_Click
                    
                    cboAreaCode.ListIndex = cboAreaCode.ListCount - 1
                    
                    Me.MousePointer = vbDefault
                    MsgBox strMsgSaveDataDone, vbInformation, strTitleInfo
                End If
            Else
                prgLoadCity.Visible = True
                
                Call SaveData
                
                Call SetButton(False)
                
                prgLoadCity.Visible = False
                
                Me.MousePointer = vbDefault
                MsgBox strMsgSaveDataDone, vbInformation, strTitleInfo
            End If
        Else
            lstSelected.BackColor = vbRed
            MsgBox "Please select a city before save", vbCritical, strTitleMissingInfo
            lstSelected.BackColor = vbWindowBackground
        End If
    Else
        txtAreaName.BackColor = vbRed
        MsgBox "Please enter a new Radio Area Code", vbCritical, strTitleMissingInfo
        txtAreaName.BackColor = vbWindowBackground
    End If
    
    Call GetAreaCode
    
    If Me.cboAreaCode.ListCount > 0 Then
        cboAreaCode.Text = Me.txtAreaName.Text
    End If
    
    cboAreaCode_Click
    EnableObject False
    
End Sub

Private Sub SetButton(blnEnable As Boolean)
    cmdAdd.Visible = Not blnEnable
    cmdEdit.Visible = Not blnEnable
    cmdSave.Visible = blnEnable
    cmdCancel.Visible = blnEnable
    
    cmdAdd.Enabled = Not blnEnable
    cmdEdit.Enabled = Not blnEnable
    cmdDelete.Enabled = Not blnEnable
    cmdSave.Enabled = blnEnable
    cmdCancel.Enabled = blnEnable
    cmdClose.Enabled = Not blnEnable
    
    cmdRemove.Enabled = blnEnable
    cmdRemoveAll.Enabled = blnEnable
    cmdInsert.Enabled = blnEnable
    cmdAddAll.Enabled = blnEnable
    
    cboBrand.Enabled = Not blnEnable
    cboAreaCode.Visible = Not blnEnable
    
    chkRural.Enabled = blnEnable
    chkUrban.Enabled = blnEnable
End Sub

Private Sub GetAreaCode()
    Dim recAreaCode As New ADODB.Recordset
    
    strQuery = "SELECT DISTINCT Radio_Area_Code FROM Radio_Area WHERE Brand_Code ='" & Left(cboBrand.Text, 4) & "'"
    recAreaCode.Open strQuery, ConnERP, adOpenDynamic, adLockOptimistic
    
    cboAreaCode.Clear
    With recAreaCode
        Do While Not .EOF
            cboAreaCode.AddItem .Fields("Radio_Area_Code").Value
            .MoveNext
        Loop
    End With
    
    Set recAreaCode = Nothing
End Sub

Private Sub SaveData()
    Dim intPosIndex As Integer
    
    On Error GoTo chk
    
    prgLoadCity.Max = lstSelected.ListCount + 3
    
    prgLoadCity.Value = 1
    
    If blnAddNewFlag = False Then
        strQuery = "DELETE FROM radio_area WHERE brand_code ='" & Left(Trim(cboBrand.Text), 4) & "' AND Radio_Area_Code ='" & Trim(txtAreaName.Text) & "'"
        ConnERP.Execute strQuery
        
        prgLoadCity.Value = prgLoadCity.Value + 1
    End If
    
    For intPosIndex = 0 To lstSelected.ListCount - 1
        lstSelected.ListIndex = intPosIndex
        
        strQuery = "INSERT INTO Radio_Area (Radio_Area_Code,Brand_Code,City_ID, urban_flag, rural_flag) VALUES ("
        strQuery = strQuery & "'" & Clear_String(Trim(txtAreaName.Text)) & "', "
        strQuery = strQuery & "'" & Left(Trim(cboBrand.Text), 4) & "', "
        strQuery = strQuery & Val(lstSelected.Text) & ", " & IIf(chkUrban.Value = 1, 1, 0) & ", " & IIf(chkRural.Value = 1, 1, 0) & ")"
        ConnERP.Execute strQuery
        
        prgLoadCity.Value = prgLoadCity.Value + 1
    Next intPosIndex

    prgLoadCity.Value = prgLoadCity.Value + 1
    
    Exit Sub
    
chk:
    MsgBox Err.Number & " " & Err.Description
End Sub

Private Function blnCekRadioAreaCode() As Boolean
    Dim recRadioArea As New ADODB.Recordset
    
    strQuery = "SELECT * FROM Radio_Area WHERE Radio_Area_Code ='" & Clear_String(Trim(txtAreaName.Text)) & "' AND Brand_Code ='" & Left(Trim(cboBrand.Text), 4) & "'"
    recRadioArea.Open strQuery, ConnERP, adOpenDynamic, adLockOptimistic
    
    With recRadioArea
        blnCekRadioAreaCode = Not .EOF
    End With
    
    Set recRadioArea = Nothing
End Function

Private Sub txtAreaName_KeyPress(KeyAscii As Integer)
    If Chr(KeyAscii) = "'" Or Chr(KeyAscii) = "," Or Chr(KeyAscii) = "." Then
        KeyAscii = 8
        Beep
    End If
End Sub

Sub SetButtonToolbar(ByVal paIsNormalMode As Boolean, picOBJ) 'TOOLBAR_AI.
'************************************************
' Procedure         : SetButtonToolbar
' Function          : TOOLBAR_AI.
' Created By        : {73 64 6B}
' Date              : 12-Apr-2015
'LastUpdate/By      : - Rudi
'************************************************

    Dim element
    Dim strDummy As String
    
    With picButton(enButtonType.bieAdd)  'ADD. 4
        .Enabled = paIsNormalMode
        .Visible = paIsNormalMode
    End With
    pnlMain.Enabled = Not paIsNormalMode
    With picButton(enButtonType.bieEdit) 'EDIT. 5
        .Enabled = paIsNormalMode
        .Visible = paIsNormalMode
    End With
    With picButton(enButtonType.bieDelete)  'DELETE. 6
        .Enabled = paIsNormalMode
        .Visible = paIsNormalMode
    End With
    With picButton(enButtonType.bieclose)       'Quit.
        .Enabled = paIsNormalMode
        .Visible = paIsNormalMode
    End With

    With picButton(enButtonType.bieSave)  'SAVE.
        .Enabled = Not paIsNormalMode
        .Visible = Not paIsNormalMode
        .Left = picButton(4).Left
    End With

    With picButton(enButtonType.biecancel) 'CANCEL.
        .Enabled = Not paIsNormalMode
        .Visible = Not paIsNormalMode
        .Left = picButton(5).Left
    End With
    'pnl_Main.Enabled = Not paIsNormalMode
    'cboBrand.Enabled = paIsNormalMode
'    blnEditOrAdd = Not paIsNormalMode
    For Each element In picOBJ
        SetPictureTB element.Index, paIsNormalMode, picOBJ
    Next element
    'Call SetSecurityCRUDStandar("Duration Catalog", picButton, "1")

End Sub

Sub SetPictureTB(ByVal Index As Integer, ByVal paIsNormalMode As Boolean, picOBJ)
 '*****************************************
'Procedure Name     : SetPictureTB
'Procedure Function :   Creates the SQL statement in ado_Data.recordset.filter
'                       and only filters text currently. It must be modified to filter other data types.
'Input Parameter    : Index,paIsNormalMode,picOBJ
'Output Parameter   :
'Date               : -
'LastUpdate/By      : - Tedi
'*****************************************
   With picOBJ(Index) 'FIRST.
        
        If .Enabled = True Then
            .Picture = LoadPicture(SetButtonImageEffect(Index, bieNormal))
        Else: .Picture = LoadPicture(SetButtonImageEffect(Index, bieDisabled))
        End If
    End With
End Sub

Private Sub SetPictureTBEnabled(ByVal Index As Integer, ByVal paIsNormalMode As Boolean)
 '*****************************************
'Procedure Name     : SetPictureTBEnabled
'Procedure Function :   Enable/Disable Button
'Input Parameter    : Index,paIsNormalMode,picOBJ
'Output Parameter   :
'Date               : -
'LastUpdate/By      : - Tedi
'*****************************************
    
    If paIsNormalMode = True Then
        picButton(Index).Picture = LoadPicture(SetButtonImageEffect(Index, bieNormal))
    Else: picButton(Index).Picture = LoadPicture(SetButtonImageEffect(Index, bieDisabled))
    End If
    
End Sub

Sub picButton_Obj(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single, picOBJ) 'TOOLBAR_AI.
'************************************************
' Procedure         : picButton_MouseMove
' Function          : TOOLBAR_AI saat mouse berada di area button.
' Created By        : {73 64 6B}
' Date              : 12-Apr-2015
' addition          : Penambahan picOBJ
'************************************************
    If (X < 0) Or (Y < 0) Or (X > picOBJ(Index).Width) Or (Y > picOBJ(Index).Height) Then 'Dua IF ini jangan diubah keluar CASE agar API-nya jalan.
        ReleaseCapture 'The MOUSE_LEAVE pseudo-event.
        picOBJ(Index).Picture = LoadPicture(SetButtonImageEffect(Index, bieNormal)) 'Back to NORMAL.

    ElseIf GetCapture() <> picOBJ(Index).hwnd Then
        SetCapture picOBJ(Index).hwnd 'The MOUSE_ENTER pseudo-event.
        picOBJ(Index).Picture = LoadPicture(SetButtonImageEffect(Index, bieOver)) 'Set to OVER_EFFECT.
    End If
End Sub

Private Sub EnableObject(ByVal paIsEnable As Boolean)
'*****************************************
'Procedure Name     : EnableObject
'Procedure Function : ~ Enable/disable control di frame Entry.
'                     ~ Call SetButtonToolbar utk Toolbar/Statusbar AI (artificial intelligence).
'Input Parameter    : paIsEnable: True=Enable, False=Disable.
'Output Parameter   : -
'Date               : 12-Apr-2015
'LastUpdate/By      : 12-Apr-2015/{73 64 6B}
'*****************************************
    
    Call SetButtonToolbar(Not paIsEnable, picButton) 'TOOLBAR_AI.

End Sub

Private Sub picButton_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single) 'TOOLBAR_AI.
'************************************************
' Procedure         : picButton_MouseDown
' Function          : TOOLBAR_AI saat mouse ditekan.
' Created By        : {73 64 6B}
' Date              : 12-Apr-2015
'************************************************
    
    picButton(Index).Picture = LoadPicture(SetButtonImageEffect(Index, bieDown)) 'FIRST.

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

Sub AdjustSizeForm()
'************************************************
' Procedure         : Txt_Year_LostFocus
' Function          : Generate IB ID
' Date              : 01/09/2001
' Parameter Input   :
' Parameter Output  :
' Last Update/By    :
'************************************************
    
    Me.Top = 0
    Me.Left = 0
    Me.Width = mdi_Main.ScaleWidth
    Me.Height = mdi_Main.ScaleHeight
    pnlMain.Height = Me.Height - pnlMain.Top - picStatusBar.Height
    fraIB.Width = pnlMain.Width - (fraIB.Left) - 150
    fraPlanMonth.Width = Me.Width - (fraPlanMonth.Left * 2)
    cmdMateri.Left = fraPlanMonth.Width - cmdMateri.Width - 300
    cmdEditPlan.Left = cmdMateri.Left
    msgCity.Width = fraPlanMonth.Width - msgCity.Left - cmdMateri.Width - 450
    fraClientApproval.Width = pnlMain.Width - fraClientApproval.Left - 150
    picApproved.Left = ((fraClientApproval.Width / 2)) - ((picApproved.Width) / 2) 'fraClientApproval.Width - (picApproved.Left * 2)
    txtMediaPlanNo.Width = fraIB.Width - txtMediaPlanNo.Left - 150
    txtPrimaryTarget.Width = txtMediaPlanNo.Width

End Sub

Private Sub picButton_Click(Index As Integer)

'************************************************
' Procedure         : picButton_Click
' Function          : Action utk Navigation dan CRUD.
' Created By        : {73 64 6B}
' Date              : 12-Apr-2015/{73 64 6B} --> Semua coding dan query sudah di optimalkan agar faster, readable, safer, standardable.
'************************************************
    Dim strCode As String, strFileRpt As String
    'Lock_MainForm True
    Select Case Index
        Case enButtonType.bieAdd  '4 'ADD.
            Call cmdAdd_Click
        Case enButtonType.bieEdit  '5 'EDIT.
            Call cmdEdit_Click
        Case enButtonType.bieDelete  '6 'DELETE.
            Call cmdDelete_Click
        Case enButtonType.bieSave  'SAVE.
            Call cmdSave_Click
        Case enButtonType.biecancel 'CANCEL.
            Call cmdCancel_Click
        Case enButtonType.bieclose  'CANCEL.
            Call cmdClose_Click
    End Select
    
End Sub
