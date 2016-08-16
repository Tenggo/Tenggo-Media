VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.ocx"
Begin VB.Form Frm_Radio_Note 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Radio Note"
   ClientHeight    =   4215
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6315
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4215
   ScaleWidth      =   6315
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox picToolbar 
      Align           =   1  'Align Top
      Appearance      =   0  'Flat
      BackColor       =   &H000000FF&
      BorderStyle     =   0  'None
      ClipControls    =   0   'False
      ForeColor       =   &H80000008&
      Height          =   750
      Left            =   0
      ScaleHeight     =   750
      ScaleWidth      =   6315
      TabIndex        =   6
      Top             =   0
      Width           =   6315
      Begin VB.PictureBox picButton 
         Appearance      =   0  'Flat
         BackColor       =   &H00F0F0F0&
         BorderStyle     =   0  'None
         ClipControls    =   0   'False
         ForeColor       =   &H80000008&
         Height          =   750
         Index           =   4
         Left            =   120
         ScaleHeight     =   750
         ScaleWidth      =   1500
         TabIndex        =   12
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
         Left            =   3195
         ScaleHeight     =   750
         ScaleWidth      =   1500
         TabIndex        =   11
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
         Left            =   4725
         ScaleHeight     =   750
         ScaleWidth      =   1500
         TabIndex        =   8
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
         Left            =   1665
         ScaleHeight     =   750
         ScaleWidth      =   1500
         TabIndex        =   7
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
         Left            =   1650
         ScaleHeight     =   750
         ScaleWidth      =   1500
         TabIndex        =   10
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
         Index           =   10
         Left            =   120
         ScaleHeight     =   750
         ScaleWidth      =   1500
         TabIndex        =   9
         Top             =   0
         Width           =   1500
      End
   End
   Begin Threed.SSPanel pnlMain 
      Align           =   1  'Align Top
      Height          =   3315
      Left            =   0
      TabIndex        =   0
      Top             =   750
      Width           =   6315
      _Version        =   65536
      _ExtentX        =   11139
      _ExtentY        =   5847
      _StockProps     =   15
      BackColor       =   15790320
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BevelOuter      =   1
      Begin VB.TextBox txtCode 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   325
         Left            =   1275
         MaxLength       =   15
         TabIndex        =   4
         Top             =   135
         Width           =   2415
      End
      Begin VB.Frame fraNote 
         Appearance      =   0  'Flat
         Caption         =   "Note"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   2535
         Left            =   75
         TabIndex        =   2
         Top             =   555
         Width           =   6000
         Begin VB.TextBox txtNote 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   2070
            Left            =   120
            MaxLength       =   500
            MultiLine       =   -1  'True
            TabIndex        =   3
            Top             =   285
            Width           =   5640
         End
      End
      Begin VB.ComboBox cboCode 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   1275
         TabIndex        =   1
         Top             =   135
         Width           =   2415
      End
      Begin VB.Label lblNoteCode 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Note Code :"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   5
         Top             =   165
         Width           =   1035
      End
   End
End
Attribute VB_Name = "Frm_Radio_Note"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
''*************************************************************
'Fungsi Form        : entry radio note
'Last Update/By     :
'*************************************************************

Option Explicit
Public frmWhatForm As Form
Public strCodeNote As String

Dim blnEditFlag  As Boolean
Dim recNote As New ADODB.Recordset


Private Sub cmdDelete_Click()

End Sub

Private Sub Form_Load()
    Call InitialData
    
    Call ButtonLock(False)
    EnableObject False
End Sub

Private Sub cboCode_Click()
    Call ClearForm
    
    txtCode.Text = ""
    If cboCode.ListIndex <> -1 Then
        Call ShowData
        
        txtCode.Text = Trim(cboCode.Text)
    End If
End Sub

Private Sub dbAdd()
'*****************************************
'Submodul Name      : dbAdd
'Procedure Function : Edit
'Programmer Name    : Tedi
'Date               : 19-11-2015
'Last Update/By     : Tedi
'Date Update        :
'Log Update/By      :
'Previous Name      : cmdAdd_Click
'***************************************************************'*****************************************
    
    cboCode.Visible = False
    
    blnEditFlag = False
    
    Call ButtonLock(True)
    
    Call PrepareTempData
    
    Call ClearForm
    
    EnableObject True
    
    
End Sub

Private Sub dbCancel()
'*****************************************
'Submodul Name      : dbCancel
'Procedure Function : Edit
'Programmer Name    : Tedi
'Date               : 19-11-2015
'Last Update/By     : Tedi
'Date Update        :
'Log Update/By      :
'Previous Name      : cmdEdit_Click
'***************************************************************'*****************************************

    Call ButtonLock(False)
    cboCode.Visible = True
    blnEditFlag = True
    Call InitialData
    EnableObject False

End Sub

Private Sub dbClose()
    Call frmWhatForm.LoadNote

    frmWhatForm.cboNoteCode.Text = strCodeNote
    
    Unload Me
End Sub

Private Sub dbDelete()
    On Error GoTo my_error
    
    If MsgBox(strMsgDeleteConfirm, vbQuestion + vbYesNo, strTitleConfirm) = vbYes Then
        ConnERP.BeginTrans
        ConnERP.Execute "DELETE FROM radio_note WHERE note_code='" & Trim(cboCode.Text) & "'"
        ConnERP.CommitTrans
        
        Call InitialData
    End If
    
    Exit Sub
    
my_error:
    ConnERP.RollbackTrans
    MsgBox Err.Number & " " & Err.Description
End Sub

Private Sub dbEdit()
'*****************************************
'Submodul Name      : dbEdit
'Procedure Function : Edit
'Programmer Name    : Tedi
'Date               : 19-11-2015
'Last Update/By     : Tedi
'Date Update        :
'Log Update/By      :
'Previous Name      : cmdEdit_Click
'***************************************************************'*****************************************

    blnEditFlag = True
    
    Call ButtonLock(True)
    
    Call PrepareTempData
    EnableObject True
    
End Sub

Private Sub dbSave()
'*****************************************
'Submodul Name      : dbCancel
'Procedure Function : Edit
'Programmer Name    : Tedi
'Date               : 19-11-2015
'Last Update/By     : Tedi
'Date Update        :
'Log Update/By      :
'Previous Name      : cmdEdit_Click
'***************************************************************'*****************************************

    If Trim(txtCode.Text) = "" Then
        txtCode.BackColor = vbRed
        
        MsgBox "You must fill Note Code box", vbExclamation, strTitleMissingInfo
        
        txtCode.BackColor = vbWindowBackground
        txtCode.SetFocus
        Exit Sub
    End If
    
    If Trim(txtNote.Text) = "" Then
        txtNote.BackColor = vbRed
        
        MsgBox "You must fill Note box", vbExclamation, strTitleMissingInfo
        
        txtNote.BackColor = vbGreen
        txtNote.SetFocus
        Exit Sub
    End If
    
    If blnCekCode(txtCode.Text) = False Then
        Call SaveData
        
        Call ButtonLock(False)
        
        cboCode.Visible = True
        
        Call ClearForm
        
        Call InitialData
        
        If blnEditFlag = False Then
            cboCode.Text = txtCode.Text
        End If
        
        blnEditFlag = True
    Else
        MsgBox "The Note Code is already exist, enter a unique code", vbExclamation, strTitleExclamation
    End If
    EnableObject True
    
End Sub

Private Sub InitialData()
    Dim recNoteCode As New ADODB.Recordset
    
    strQuery = "SELECT Note_Code FROM Radio_Note ORDER BY Note_Code"
    recNoteCode.Open strQuery, ConnERP, adOpenStatic, adLockReadOnly
    
    cboCode.Clear
    With recNoteCode
        Do While .EOF = False
            cboCode.AddItem .Fields("Note_Code").Value
            .MoveNext
        Loop
    End With
    
    recNoteCode.Close
    Set recNoteCode = Nothing
    
    If strCodeNote <> "" Then
        cboCode.Text = strCodeNote
        txtCode.Text = strCodeNote
        
        Call ShowData
    End If
End Sub

Private Sub ButtonLock(blnEnable As Boolean)
    fraNote.Enabled = blnEnable
'    cmdClose.Enabled = Not blnEnable
'    cmdDelete.Enabled = Not blnEnable
'    cmdAdd.Visible = Not blnEnable
'    cmdEdit.Visible = Not blnEnable
    cboCode.Enabled = Not blnEnable
End Sub

Private Sub SaveData()
    
    With recNote
        If blnEditFlag = False Then
            .AddNew
            .Fields("Note_Code").Value = Trim(txtCode.Text)
        End If
        
        .Fields("Description").Value = Trim(txtNote.Text)
        .Update
    End With
    
    recNote.Close
    Set recNote = Nothing
End Sub

Private Sub PrepareTempData()
    strQuery = "SELECT * FROM Radio_Note WHERE Note_Code ='" & Trim(cboCode.Text) & "'"
    
    If recNote.State = adStateOpen Then
        recNote.Close
    End If
    
    recNote.Open strQuery, ConnERP, adOpenDynamic, adLockOptimistic
End Sub

Private Sub ShowData()
    Dim recRadioNote As New ADODB.Recordset
    
    strQuery = "SELECT * FROM radio_note WHERE note_code='" & Trim(cboCode.Text) & "'"
    recRadioNote.Open strQuery, ConnERP, adOpenStatic, adLockReadOnly
    
    With recRadioNote
        If Not .EOF Then
            txtNote.Text = .Fields("Description").Value
            .MoveNext
        End If
    End With
    
    recRadioNote.Close
    Set recRadioNote = Nothing
    
    strQuery = Empty
End Sub

Private Function blnCekCode(ByRef strNoteCode As String) As Boolean
    Dim recNoteCode As New ADODB.Recordset
    
    If blnEditFlag = False Then
        strQuery = "SELECT * FROM Radio_Note WHERE Note_code='" & Trim(strNoteCode) & "'"
        recNoteCode.Open strQuery, ConnERP, adOpenStatic, adLockReadOnly
        
        With recNoteCode
            blnCekCode = Not .EOF
            
            .Close
        End With
        
        Set recNoteCode = Nothing
        
        strQuery = Empty
    Else
        blnCekCode = False
        Exit Function
    End If
End Function

Private Sub ClearForm()
    txtCode.Text = Empty
    txtNote.Text = Empty
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
    With picButton(enButtonType.bieADD) 'EDIT. 5
        .Enabled = paIsNormalMode
        .Visible = paIsNormalMode
    End With
    With picButton(enButtonType.bieedit) 'EDIT. 5
        .Enabled = paIsNormalMode
        .Visible = paIsNormalMode
    End With
    With picButton(enButtonType.bieSave)  'SAVE.
        .Enabled = Not paIsNormalMode
        .Visible = Not paIsNormalMode
    End With
    With picButton(enButtonType.bieCancel) 'CANCEL.
        .Enabled = Not paIsNormalMode
        .Visible = Not paIsNormalMode
    End With
    With picButton(enButtonType.bieDelete)    'FIND.
        .Enabled = paIsNormalMode
        .Visible = paIsNormalMode
    End With
    With picButton(enButtonType.bieClose)      'Quit.
        .Enabled = paIsNormalMode
        .Visible = paIsNormalMode
    End With

    'blnEditOrAdd = Not paIsNormalMode
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
    pnl_Main.Height = Me.ScaleHeight - picToolbar.Height - picStatusBar.Height
    fra_Deliverable.Height = pnl_Main.Height - (fra_Deliverable.Top + 100)
    SSTab3.Height = fra_Deliverable.Height - (SSTab3.Top) - 150
    txtOther_Recomedation.Height = SSTab3.Height - (txtOther_Recomedation.Top) - 150
    txtAggreed_Channel_shortlist.Height = txtOther_Recomedation.Height
    fra_DeliverableChannel.Height = pnl_Main.Height - (fra_DeliverableChannel.Top + 100)
    fraFilter.Width = pnl_Main.Width - (fraFilter.Left * 2)
    lineFilter.X1 = fraFilter.Width / 2
    lineFilter.X2 = lineFilter.X1
    Fra_Approve.Left = lineFilter.X2 + Label7.Left
    txtYear.Width = lineFilter.X2 - txtYear.Left - 50
    txtClient_Brief_Id.Width = txtYear.Width
    txtExtention.Width = txtYear.Width
    txtStatus.Width = txtYear.Width
    'left part
    lbl_dateofPreviousIssue.Left = lineFilter.X1 + Label7.Left
    dtpDate_Previouse.Left = lbl_dateofPreviousIssue.Left + lbl_dateofPreviousIssue.Width + 50
    dtpDate_Issue.Left = dtpDate_Previouse.Left
    lbl_DateIssue.Left = lbl_dateofPreviousIssue.Left
    lblCountry.Left = lbl_dateofPreviousIssue.Left
    cboCountry.Left = dtpDate_Previouse.Left
    Fra_Approve.Left = dtpDate_Previouse.Left
    fra_DeliverableChannel.Width = pnl_Main.Width - fra_DeliverableChannel.Left - fraFilter.Left
    lstRec_Channel_Selection.Width = fra_DeliverableChannel.Width - (lstRec_Channel_Selection.Left * 2)
    lstRec_Channel_Selection.Height = fra_DeliverableChannel.Height - (lstRec_Channel_Selection.Top) - 200
    chk_All.Top = lstRec_Channel_Selection.Height + lstRec_Channel_Selection.Top + 50
    lbl_CheckAll.Top = chk_All.Top
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
        Case enButtonType.bieADD '4 'EDIT.
            dbEdit
        Case enButtonType.bieedit  '5 'EDIT.
            dbEdit
        Case enButtonType.bieDelete    '8 'FIND.
            Call dbDelete
        Case enButtonType.bieClose   '23 'EXIT.
            Call dbClose
        Case enButtonType.bieSave  '10 SAVE.
            Call dbSave
        Case enButtonType.bieCancel '11 CANCEL.
            Call dbCancel
    End Select

End Sub

Private Sub CheckForClickAll(ByRef ObjListBox As ListBox, ByRef objChkBox As CheckBox, ByVal bol_Temp As Boolean)
'*****************************************
'Submodul Name      : CheckForClickAll
'Procedure Function : Untuk memeriksa kompisisi apakah row di listview tercontreng semua
'                     - Jika node tercontreng semua maka nilai chkAll/objChkBox.Value = 1, jika tidak maka chkAll/objChkBox.Value  = 0
'                       Pemrosesan chkAll.Value diperintahkan dengan code, sehingga perlu diberikan nilai bolean blnNotByClickByList = True
'                     - jika check list row ada yang tidak tercontreng maka nilai bol_Temp=false sebaliknya true
'Used Object        : objListBox,objChkBox,bol_Temp
'Programmer Name    : Tedi
'Date               : 19-11-2015
'Last Update/By     : Tedi
'Date Update        :
'Log Update/By      :
'***************************************************************'*****************************************
    
    If blnNotByClickByList = True Then Exit Sub
    Dim intCheck As Integer
    For intCheck = 0 To ObjListBox.ListCount - 1
        If ObjListBox.Selected(intCheck) = False Then
            'bol_Temp = False
            objChkBox.Value = 0
            Exit Sub
        End If
    Next intCheck
    objChkBox.Value = 1
    'bol_Temp = False
End Sub
