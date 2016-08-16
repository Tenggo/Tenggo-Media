VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "mscomctl.OCX"
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Begin VB.Form Frm_MPRadioArea_Selection 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Select Radio Station"
   ClientHeight    =   5625
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   5040
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5625
   ScaleWidth      =   5040
   StartUpPosition =   2  'CenterScreen
   Begin Threed.SSPanel pnl_Main 
      Align           =   1  'Align Top
      Height          =   4875
      Left            =   0
      TabIndex        =   3
      Top             =   750
      Width           =   5040
      _Version        =   65536
      _ExtentX        =   8890
      _ExtentY        =   8599
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
      Begin MSComctlLib.TreeView trv_rd_station 
         Height          =   4695
         Left            =   90
         TabIndex        =   4
         Top             =   90
         Width           =   4875
         _ExtentX        =   8599
         _ExtentY        =   8281
         _Version        =   393217
         HideSelection   =   0   'False
         LabelEdit       =   1
         LineStyle       =   1
         Style           =   7
         Checkboxes      =   -1  'True
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
   Begin VB.PictureBox picToolbar 
      Align           =   1  'Align Top
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ClipControls    =   0   'False
      ForeColor       =   &H80000008&
      Height          =   750
      Left            =   0
      ScaleHeight     =   750
      ScaleWidth      =   5040
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   0
      Width           =   5040
      Begin VB.PictureBox picButton 
         Appearance      =   0  'Flat
         BackColor       =   &H00F0F0F0&
         BorderStyle     =   0  'None
         ClipControls    =   0   'False
         ForeColor       =   &H80000008&
         Height          =   750
         Index           =   11
         Left            =   1620
         ScaleHeight     =   750
         ScaleWidth      =   1500
         TabIndex        =   2
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
         Index           =   47
         Left            =   90
         ScaleHeight     =   750
         ScaleWidth      =   1500
         TabIndex        =   1
         TabStop         =   0   'False
         Top             =   0
         Width           =   1500
      End
   End
End
Attribute VB_Name = "Frm_MPRadioArea_Selection"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim strSql As String
Public intNodesInserted As Integer
Dim rsTemp As New ADODB.Recordset

Private Sub db_insert()
    Dim i As Integer
    intNodesInserted = 0
    For i = 1 To trv_rd_station.Nodes.Count
        If trv_rd_station.Nodes(i).Checked And Left(trv_rd_station.Nodes(i).KEY, 7) = "Stat-->" Then
            strSql = "select a.station_code,a.station_name,b.city,c.area_name "
            strSql = strSql & "from radio_station a inner join city b "
            strSql = strSql & "on a.city_id = b.city_id and a.station_code = '" & Right(trv_rd_station.Nodes(i).KEY, Len(trv_rd_station.Nodes(i).KEY) - 7) & "' "
            strSql = strSql & "inner join area c on a.area_id = c.area_id "
            rsTemp.Open strSql, ConnERP, 1, 3
            If Not rsTemp.EOF Then
                With Frm_MPRadioAreaAdd.fg_area_detail
                    If .TextMatrix(1, 1) <> "" Then
                        .Rows = .Rows + 1
                    End If
                    .TextMatrix(.Rows - 1, 0) = .Rows - 1
                    .TextMatrix(.Rows - 1, 1) = rsTemp(0)
                    .TextMatrix(.Rows - 1, 2) = " " & rsTemp(1)
                    .TextMatrix(.Rows - 1, 3) = " " & rsTemp(2)
                    .TextMatrix(.Rows - 1, 4) = " " & rsTemp(3)
                    intNodesInserted = intNodesInserted + 1
                End With
            End If
            rsTemp.Close
        End If
    Next
    If intNodesInserted > 0 Then
        Frm_MPRadioAreaAdd.picButton(enButtonType.bieSave).Enabled = True
    End If
    MsgBox intNodesInserted & " Radio Station(s) Inserted!", vbExclamation, strApplication_Name
    Unload Me
    
End Sub

Private Sub Form_Load()
    
    Call EnableObject(False)
    Call load_radio_station
    
End Sub

Private Sub load_radio_station()
    Dim str_area_id As String, str_city_id As String
    Dim strStationCode As String
    strStationCode = ""
    Dim i As Integer
    With Frm_MPRadioAreaAdd.fg_area_detail
        If .Rows > 2 Or .TextMatrix(1, 1) <> "" Then
            For i = 1 To .Rows - 1
                strStationCode = strStationCode & "'" & .TextMatrix(i, 1) & "',"
            Next
            strStationCode = Left(strStationCode, Len(strStationCode) - 1)
            strSql = "select a.area_id,a.area_name,b.city_id,b.city,c.station_code,c.station_name "
            strSql = strSql & "from area a inner join city b "
            strSql = strSql & "on a.area_id = b.area_id "
            strSql = strSql & "inner join radio_station c "
            strSql = strSql & "on b.city_id = c.city_id "
            strSql = strSql & "and c.station_code not in (" & strStationCode & ") "
            strSql = strSql & "order by a.area_id,b.city_id,c.station_code"
        Else
            strSql = "select a.area_id,a.area_name,b.city_id,b.city,c.station_code,c.station_name "
            strSql = strSql & "from area a inner join city b "
            strSql = strSql & "on a.area_id = b.area_id "
            strSql = strSql & "inner join radio_station c "
            strSql = strSql & "on b.city_id = c.city_id "
            strSql = strSql & "order by a.area_id,b.city_id,c.station_code"
        End If
    End With
    rsTemp.Open strSql, ConnERP, 1, 3
    
    str_area_id = Empty
    str_city_id = Empty
    trv_rd_station.Nodes.Clear
    
    While Not rsTemp.EOF
        'Create area node
        If str_area_id <> rsTemp(0) Then
            trv_rd_station.Nodes.Add , , "Area-->" & rsTemp(0), rsTemp(1)
            str_area_id = rsTemp(0)
        End If
        'create city node
        If str_city_id <> rsTemp(2) Then
            trv_rd_station.Nodes.Add "Area-->" & rsTemp(0), tvwChild, "City-->" & rsTemp(2), rsTemp(3)
            str_city_id = rsTemp(2)
        End If
        'adding radio station
        trv_rd_station.Nodes.Add "City-->" & rsTemp(2), tvwChild, "Stat-->" & rsTemp(4), rsTemp(5)
        rsTemp.MoveNext
    Wend
    
    rsTemp.Close
End Sub

Private Sub trv_rd_station_NodeCheck(ByVal Node As MSComctlLib.Node)
    
    Select Case Left(Node.KEY, 7)
        Case "Area-->":
            Call Check_On_Area(Node.Index, Node.Checked)
        Case "City-->":
            Call Check_On_City(Node.Index, Node.Checked)
        Case "Stat-->"
            Call Check_On_Station(Node.Index, Node.Checked)
    End Select
End Sub

Private Sub Check_On_Area(intNode As Integer, isChecked As Boolean)
    Dim i As Integer
    For i = 1 To trv_rd_station.Nodes.Count
        If Left(trv_rd_station.Nodes(i).KEY, 7) = "City-->" Then
            If trv_rd_station.Nodes(i).Parent.KEY = trv_rd_station.Nodes(intNode).KEY Then
                trv_rd_station.Nodes(i).Checked = isChecked
                Call Check_On_City(i, isChecked)
            End If
        End If
    Next
End Sub

Private Sub Check_On_City(intNode As Integer, isChecked As Boolean)
    Dim i As Integer
    For i = 1 To trv_rd_station.Nodes.Count
        If Left(trv_rd_station.Nodes(i).KEY, 7) = "Stat-->" Then
            If trv_rd_station.Nodes(i).Parent.KEY = trv_rd_station.Nodes(intNode).KEY Then
                trv_rd_station.Nodes(i).Checked = isChecked
            End If
        End If
    Next
    If Not isChecked Then
        trv_rd_station.Nodes(intNode).Parent.Checked = isChecked
    End If
End Sub

Private Sub Check_On_Station(intNode As Integer, isChecked As Boolean)
    If Not isChecked Then
        trv_rd_station.Nodes(intNode).Parent.Checked = isChecked
        trv_rd_station.Nodes(trv_rd_station.Nodes(intNode).Parent.Index).Parent.Checked = isChecked
    End If
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

    With picButton(enButtonType.bieInsert)  'INSERT.
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
            
        Case enButtonType.bieInsert  'INSERT.
            Call db_insert
            
        Case enButtonType.bieCancel 'CANCEL.
            Unload Me
            
    End Select

End Sub
