VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFlxGrd.ocx"
Begin VB.Form Frm_MPRadioAreaAdd 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Radio Area"
   ClientHeight    =   6180
   ClientLeft      =   2760
   ClientTop       =   3750
   ClientWidth     =   8565
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6180
   ScaleWidth      =   8565
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
      ScaleWidth      =   8565
      TabIndex        =   7
      TabStop         =   0   'False
      Top             =   0
      Width           =   8565
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
         TabIndex        =   12
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
         Index           =   11
         Left            =   6210
         ScaleHeight     =   750
         ScaleWidth      =   1500
         TabIndex        =   11
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
         Index           =   10
         Left            =   4680
         ScaleHeight     =   750
         ScaleWidth      =   1500
         TabIndex        =   10
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
         Index           =   70
         Left            =   3150
         ScaleHeight     =   750
         ScaleWidth      =   1500
         TabIndex        =   9
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
         Index           =   6
         Left            =   1620
         ScaleHeight     =   750
         ScaleWidth      =   1500
         TabIndex        =   8
         TabStop         =   0   'False
         Top             =   0
         Width           =   1500
      End
   End
   Begin Threed.SSPanel pnl_Main 
      Align           =   1  'Align Top
      Height          =   5430
      Left            =   0
      TabIndex        =   0
      Top             =   750
      Width           =   8565
      _Version        =   65536
      _ExtentX        =   15108
      _ExtentY        =   9578
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
      Begin VB.ComboBox cbo_area_id 
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
         Left            =   5655
         Style           =   2  'Dropdown List
         TabIndex        =   6
         Top             =   4995
         Visible         =   0   'False
         Width           =   345
      End
      Begin VB.TextBox txtAreaName 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   1530
         MaxLength       =   25
         TabIndex        =   4
         Top             =   4995
         Width           =   6960
      End
      Begin VB.Frame Frame1 
         Height          =   4860
         Left            =   90
         TabIndex        =   1
         Top             =   15
         Width           =   8415
         Begin VB.ListBox lst_rd_area 
            Height          =   4545
            Left            =   5985
            TabIndex        =   3
            Top             =   195
            Width           =   2325
         End
         Begin MSFlexGridLib.MSFlexGrid fg_area_detail 
            Height          =   4545
            Left            =   120
            TabIndex        =   2
            Top             =   195
            Width           =   5790
            _ExtentX        =   10213
            _ExtentY        =   8017
            _Version        =   393216
            Cols            =   5
            AllowBigSelection=   -1  'True
            AllowUserResizing=   1
            Appearance      =   0
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
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Save Template As"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   120
         TabIndex        =   5
         Top             =   5025
         Width           =   1290
      End
   End
   Begin VB.Menu mnu_pup 
      Caption         =   "pop up"
      Visible         =   0   'False
      Begin VB.Menu mnu_Delete 
         Caption         =   "&Delete"
      End
      Begin VB.Menu mnu_add 
         Caption         =   "&Add"
      End
      Begin VB.Menu mnu_clear 
         Caption         =   "&Clear"
      End
   End
End
Attribute VB_Name = "Frm_MPRadioAreaAdd"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rsTemp As New ADODB.Recordset
Dim strSql As String
Dim strBrandCode As String
Dim int_current_row As Integer

Private Sub fg_area_detail_Click()
'    Call fg_area_detail_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
End Sub

Private Sub Form_Load()
    strBrandCode = Left(Frm_MPActivityDetail.txtMPNumber.Text, 4)
    Call Load_Radio_Area
    Call init_Grid_Detail
    'cmdSave.Enabled = False
    Call EnableObject(False)
    Call SetButtonSave(False, picButton)
    
End Sub

Private Sub db_save()
    Dim strAreaName As String
    Dim intRecFound As Integer
    
    If fg_area_detail.Rows = 2 Then
        If fg_area_detail.TextMatrix(1, 1) = "" Then
            MsgBox "No Radio Station selected, Nothing to saved!", vbExclamation, strApplication_Name
            Call SetButtonSave(False, picButton)
            'cmdSave.Enabled = False
            Exit Sub
        End If
    End If
    
    strAreaName = Trim(Clear_String(txtAreaName.Text))
    If strAreaName = "" Then
        MsgBox "Please enter new area name!", vbExclamation, strApplication_Name
        Exit Sub
    End If
    rsTemp.Open "select count(*) from radio_area_new where area_name = '" & strAreaName & "' and brand_code = '" & strBrandCode & "'", ConnERP, 1, 3
    intRecFound = rsTemp(0)
    rsTemp.Close
    If intRecFound > 0 Then
        MsgBox "Area Name Already Exist! Please enter another area name!", vbExclamation, strApplication_Name
        Exit Sub
    End If
    
    Dim i As Integer
    Dim strArea_Id As String
    strSql = "select isnull(max(cast(substring(area_id,5,6) as int))+1,1) next_area_id  from radio_area_new"
    rsTemp.Open strSql, ConnERP, 1, 3
        strArea_Id = "area" & Right("000000" & CStr(rsTemp(0)), 6)
    rsTemp.Close
    strSql = "insert into radio_area_new(area_id,area_name,brand_code) values"
    strSql = strSql & "('" & strArea_Id & "','" & strAreaName & "','" & strBrandCode & "')"
    ConnERP.Execute strSql
    
    For i = 1 To fg_area_detail.Rows - 1
        strSql = "insert into radio_area_detail(area_id,station_code) values"
        strSql = strSql & "('" & strArea_Id & "','" & fg_area_detail.TextMatrix(i, 1) & "')"
        ConnERP.Execute strSql
    Next
    
    lst_rd_area.AddItem strAreaName
    cbo_area_id.AddItem strArea_Id
    Call SetButtonSave(False, picButton)
    'cmdSave.Enabled = False
    
    'add new area ke form mpactivitydetail
    Frm_MPActivityDetail.cboRDArea.AddItem strAreaName
    Frm_MPActivityDetail.cboRDStation.AddItem fg_area_detail.Rows - 1
End Sub

Private Sub fg_area_detail_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim i As Integer
    
    With fg_area_detail
        If .MouseRow > 0 Then
            If int_current_row < .Rows Then
                .Row = int_current_row
                For i = 1 To .cols - 1
                    .col = i
                    .CellBackColor = vbWhite
                Next
            End If
            .Row = .MouseRow
            For i = 1 To .cols - 1
                .col = i
                .CellBackColor = vbYellow
            Next
            .col = .MouseCol
            int_current_row = .Row
        End If
    End With
    
End Sub

Private Sub Load_Radio_Area()
    strSql = "select area_id,area_name from radio_area_new where brand_code = '" & strBrandCode & "'"
    rsTemp.Open strSql, ConnERP, 1, 3
    cbo_area_id.Clear
    lst_rd_area.Clear
    While Not rsTemp.EOF
        cbo_area_id.AddItem rsTemp(0)
        lst_rd_area.AddItem rsTemp(1)
        rsTemp.MoveNext
    Wend
    rsTemp.Close
End Sub

Private Sub db_cancel()
    Dim pesan
    'If cmdSave.Enabled Then
    If picButton(enButtonType.bieSave).Enabled = True Then
        pesan = MsgBox("close window without saving?", vbYesNo, strApplication_Name)
        If pesan = 7 Then
            Exit Sub
        End If
    End If
    Unload Me
End Sub

Private Sub lst_rd_area_Click()
    Dim pesan
    'If cmdSave.Enabled Then
    If picButton(enButtonType.bieSave).Enabled = True Then
        pesan = MsgBox("Current Radio Area has not been saved!" & vbCrLf & "Continue Loading Selected Area?", vbYesNo, strApplication_Name)
        If pesan = 7 Then
            Exit Sub
        End If
    End If
    Call Load_Area_Detail(cbo_area_id.List(lst_rd_area.ListIndex))
    Call SetButtonSave(False, picButton)
    'cmdSave.Enabled = False
End Sub

Private Sub Clear_Grid_Detail()
    Dim i As Integer
    With fg_area_detail
        .Rows = 2
        For i = 1 To 5
            .TextMatrix(1, i - 1) = Empty
        Next
    End With
End Sub

Private Sub init_Grid_Detail()
    Dim i As Integer
    With fg_area_detail
        .TextMatrix(0, 0) = "No"
        .TextMatrix(0, 2) = "Station Name"
        .TextMatrix(0, 3) = "City"
        .TextMatrix(0, 4) = "Region"
        .ColWidth(0) = 300
        .ColWidth(1) = 0
        For i = 2 To 4
            .ColWidth(i) = 2000
        Next
        .BackColor = vbWhite
    End With
    int_current_row = 1
End Sub

Private Sub Load_Area_Detail(strArea_Id As String)
    Dim i As Integer
    strSql = "select a.station_code,a.station_name,b.city,c.area_name "
    strSql = strSql & "from radio_station a inner join city b "
    strSql = strSql & "on a.city_id = b.city_id "
    strSql = strSql & "inner join area c on a.area_id = c.area_id "
    strSql = strSql & "inner join radio_area_detail d "
    strSql = strSql & "on a.station_code = d.station_code and d.area_id = '" & strArea_Id & "' "
    strSql = strSql & "order by a.station_name"
    Call Clear_Grid_Detail
    rsTemp.Open strSql, ConnERP, 1, 3
    i = 1
    While Not rsTemp.EOF
        i = i + 1
        fg_area_detail.Rows = i
        fg_area_detail.TextMatrix(i - 1, 0) = i - 1
        fg_area_detail.TextMatrix(i - 1, 1) = rsTemp(0)
        fg_area_detail.TextMatrix(i - 1, 2) = " " & rsTemp(1)
        fg_area_detail.TextMatrix(i - 1, 3) = " " & rsTemp(2)
        fg_area_detail.TextMatrix(i - 1, 4) = " " & rsTemp(3)
        rsTemp.MoveNext
    Wend
    rsTemp.Close
End Sub

Private Sub mnu_Add_Click()
    Frm_MPRadioArea_Selection.show 1
    If Frm_MPRadioArea_Selection.intNodesInserted > 0 Then
        Call SetButtonSave(True, picButton)
    End If
End Sub

Private Sub mnu_clear_Click()
    Call Clear_Grid_Detail
    Call SetButtonSave(False, picButton)
    'cmdSave.Enabled = False
End Sub

Private Sub mnu_Delete_Click()
    Dim i As Integer
    Dim j As Integer
    With fg_area_detail
        If .Rows = 2 Then
            For i = 0 To .cols - 1
                .TextMatrix(.Row, i) = Empty
            Next
        Else
            If .Row < .Rows - 1 Then
                For j = .Row + 1 To .Rows - 1
                    For i = 1 To .cols - 1
                        .TextMatrix(j - 1, i) = .TextMatrix(j, i)
                    Next
                Next
            End If
            .Rows = .Rows - 1
        End If
        If .Rows = 2 And .TextMatrix(1, 1) = "" Then
            Call SetButtonSave(False, picButton)
            'cmdSave.Enabled = False
        Else
            Call SetButtonSave(True, picButton)
            'cmdSave.Enabled = True
        End If
    End With
    
End Sub

Sub SetButtonSave(ByVal blnStatus As Boolean, picOBJ)
'************************************************
' Procedure         : SetPicButton
' Function          : TOOLBAR_AI.
' Created By        : {73 64 6B}
' Date              : 12-Apr-2015
' LastUpdate/By     : - Rudi
'************************************************

    Dim element
    Dim strDummy As String
    
    With picButton(enButtonType.bieSave)  'SAVE
        .Enabled = blnStatus
    End With
    
    For Each element In picOBJ
        SetPictureTB element.Index, blnStatus, picOBJ
    Next element

End Sub

Sub SetButtonDelete(ByVal blnStatus As Boolean, picOBJ)
'************************************************
' Procedure         : SetPicButton
' Function          : TOOLBAR_AI.
' Created By        : {73 64 6B}
' Date              : 12-Apr-2015
' LastUpdate/By     : - Rudi
'************************************************

    Dim element
    Dim strDummy As String
    
    With picButton(enButtonType.bieDelete)  'DELETE. 6
        .Enabled = blnStatus
    End With
    
    For Each element In picOBJ
        SetPictureTB element.Index, blnStatus, picOBJ
    Next element

End Sub

Sub SetButtonToolbar(ByVal paIsNormalMode As Boolean, picOBJ) 'TOOLBAR_AI.
'************************************************
' Procedure         : SetButtonToolbar
' Function          : TOOLBAR_AI.
' Created By        : {73 64 6B}
' Date              : 12-Apr-2015
' LastUpdate/By     : - Rudi
'************************************************

    Dim element
    Dim strDummy As String
    
    With picButton(enButtonType.bieAdd)  'ADD. 4
        .Enabled = paIsNormalMode
        .Visible = paIsNormalMode
    End With
    
    With picButton(enButtonType.bieClear) 'CLEAR. 70
        .Enabled = paIsNormalMode
        .Visible = paIsNormalMode
    End With
    
    With picButton(enButtonType.bieDelete)  'DELETE. 6
        .Enabled = paIsNormalMode
        .Visible = paIsNormalMode
    End With

    With picButton(enButtonType.bieSave)  'SAVE.
        .Enabled = paIsNormalMode
        .Visible = paIsNormalMode
    End With

    With picButton(enButtonType.bieCancel) 'CANCEL.
        .Enabled = paIsNormalMode
        .Visible = paIsNormalMode
    End With
    
    For Each element In picOBJ
        SetPictureTB element.Index, paIsNormalMode, picOBJ
    Next element

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

Sub EnableObject(ByVal paIsEnable As Boolean)
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

Private Sub picButton_Click(Index As Integer)
'************************************************
' Procedure         : picButton_Click
' Function          : Action utk Navigation dan CRUD.
' Created By        : {73 64 6B}
' Date              : 12-Apr-2015/{73 64 6B} --> Semua coding dan query sudah di optimalkan agar faster, readable, safer, standardable.
'************************************************
    Dim strCode As String, strFileRpt As String
    
    Select Case Index
            
        Case enButtonType.bieAdd  '4 'ADD.
            Call mnu_Add_Click
            
        Case enButtonType.bieDelete  '6 'DELETE.
            Call mnu_Delete_Click
            
        Case enButtonType.bieClear   '70 'CLEAR.
            Call mnu_clear_Click
            
        Case enButtonType.bieSave  'SAVE.
            Call db_save
            
        Case enButtonType.bieCancel 'CANCEL.
            Call db_cancel
    End Select

End Sub
