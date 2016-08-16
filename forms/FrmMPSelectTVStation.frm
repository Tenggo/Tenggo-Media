VERSION 5.00
Begin VB.Form Frm_MPSelectTVStation 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "TV Poll"
   ClientHeight    =   5355
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5445
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5355
   ScaleWidth      =   5445
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
      ScaleWidth      =   5445
      TabIndex        =   13
      TabStop         =   0   'False
      Top             =   0
      Width           =   5445
      Begin VB.PictureBox picButton 
         Appearance      =   0  'Flat
         BackColor       =   &H00F0F0F0&
         BorderStyle     =   0  'None
         ClipControls    =   0   'False
         ForeColor       =   &H80000008&
         Height          =   750
         Index           =   23
         Left            =   1620
         Picture         =   "FrmMPSelectTVStation.frx":0000
         ScaleHeight     =   750
         ScaleWidth      =   1500
         TabIndex        =   15
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
         Left            =   90
         Picture         =   "FrmMPSelectTVStation.frx":1D07
         ScaleHeight     =   750
         ScaleWidth      =   1500
         TabIndex        =   14
         TabStop         =   0   'False
         Top             =   0
         Width           =   1500
      End
   End
   Begin VB.Frame Frame1 
      Height          =   4740
      Left            =   -15
      TabIndex        =   0
      Top             =   660
      Width           =   5835
      Begin VB.CommandButton Cmd_Save 
         Caption         =   "&Save"
         Height          =   375
         Left            =   3300
         TabIndex        =   12
         Top             =   4785
         Visible         =   0   'False
         Width           =   1035
      End
      Begin VB.CommandButton Cmd_Close 
         Caption         =   "&Close"
         Height          =   375
         Left            =   4380
         TabIndex        =   11
         Top             =   4785
         Width           =   1035
      End
      Begin VB.TextBox txt_tv_group_name 
         Height          =   345
         Left            =   210
         MaxLength       =   50
         TabIndex        =   9
         Top             =   4230
         Width           =   5040
      End
      Begin VB.CommandButton Cmd_Remove_ALL 
         Caption         =   "<<"
         Height          =   375
         Left            =   2415
         TabIndex        =   8
         Top             =   2670
         Width           =   645
      End
      Begin VB.CommandButton Cmd_Remove_1 
         Caption         =   "<"
         Height          =   375
         Left            =   2415
         TabIndex        =   7
         Top             =   2250
         Width           =   645
      End
      Begin VB.CommandButton Cmd_Select_ALL 
         Caption         =   ">>"
         Height          =   375
         Left            =   2415
         TabIndex        =   6
         Top             =   1305
         Width           =   645
      End
      Begin VB.CommandButton Cmd_Select_1 
         Caption         =   ">"
         Height          =   375
         Left            =   2415
         TabIndex        =   5
         Top             =   885
         Width           =   645
      End
      Begin VB.ListBox lst_selected_tv_station 
         Height          =   3375
         Left            =   3120
         TabIndex        =   3
         Top             =   480
         Width           =   2130
      End
      Begin VB.ListBox lst_tv_station_catalog 
         Height          =   3375
         Left            =   225
         TabIndex        =   1
         Top             =   465
         Width           =   2130
      End
      Begin VB.Label Label3 
         Caption         =   "TV Group Name :"
         ForeColor       =   &H000000C0&
         Height          =   210
         Left            =   210
         TabIndex        =   10
         Top             =   3945
         Width           =   1290
      End
      Begin VB.Label Label2 
         Caption         =   "Selected TV Stations :"
         ForeColor       =   &H000000C0&
         Height          =   255
         Left            =   3135
         TabIndex        =   4
         Top             =   225
         Width           =   1815
      End
      Begin VB.Label Label1 
         Caption         =   "Available TV Stations :"
         Height          =   255
         Left            =   240
         TabIndex        =   2
         Top             =   225
         Width           =   1815
      End
   End
End
Attribute VB_Name = "Frm_MPSelectTVStation"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim strBrandCode As String
'
Private Sub Cmd_Save_Click()
    Dim strStation_Code As String, strSql As String
    Dim i As Integer, rsTemp As New ADODB.Recordset
    If lst_selected_tv_station.ListCount = 0 Then
        MsgBox "Select TV Station!"
        Exit Sub
    End If
    If Trim(txt_tv_group_name.Text) = "" Then
        MsgBox "Input TV Group Name!"
        txt_tv_group_name.SetFocus
        Exit Sub
    End If
    strStation_Code = ""
    For i = 1 To lst_selected_tv_station.ListCount
        strStation_Code = strStation_Code & lst_selected_tv_station.List(i - 1) & ","
    Next
    strStation_Code = Left(strStation_Code, Len(strStation_Code) - 1)
    strSql = "select station_name from tv_station_media_plan where station_code='" & strStation_Code & "' and brand_code in ('ALL','" & strBrandCode & "')"
    rsTemp.Open strSql, ConnERP, 1, 3
    If Not rsTemp.EOF Then
        MsgBox ("A TV Group with the same selection found in Database! TV Group name is " & Chr(34) & rsTemp("station_name") & Chr(34) & ".")
        rsTemp.Close
        Exit Sub
    End If
    If rsTemp.State = adStateOpen Then rsTemp.Close
    strSql = "insert into tv_station_media_plan(station_code,station_name,brand_code) values "
    strSql = strSql & "('" & strStation_Code & "','" & Clear_String(Trim(txt_tv_group_name.Text)) & "','" & strBrandCode & "')"
    ConnERP.Execute strSql
    MsgBox "New TV Group Added!"
    Frm_MPActivityDetail.cboTVStationCode.AddItem strStation_Code
    Frm_MPActivityDetail.cboTVStationName.AddItem Trim(txt_tv_group_name.Text)
End Sub

Private Sub Cmd_Select_1_Click()
    Dim i As Integer
    If lst_tv_station_catalog.SelCount <> 0 Then
        i = lst_tv_station_catalog.ListIndex
        lst_selected_tv_station.AddItem lst_tv_station_catalog.List(i)
        lst_tv_station_catalog.RemoveItem (i)
        If i < lst_tv_station_catalog.ListCount Then
            lst_tv_station_catalog.Selected(i) = True
        End If
    Else
        MsgBox "No TV Station Selected!"
    End If
End Sub

Private Sub Cmd_Select_ALL_Click()
    Dim i As Integer
    For i = 1 To lst_tv_station_catalog.ListCount
        lst_selected_tv_station.AddItem lst_tv_station_catalog.List(0)
        lst_tv_station_catalog.RemoveItem (0)
    Next
End Sub

Private Sub Cmd_Remove_1_Click()
    Dim i As Integer
    If lst_selected_tv_station.SelCount <> 0 Then
        i = lst_selected_tv_station.ListIndex
        lst_tv_station_catalog.AddItem lst_selected_tv_station.List(i)
        lst_selected_tv_station.RemoveItem (i)
        If i < lst_selected_tv_station.ListCount Then
            lst_selected_tv_station.Selected(i) = True
        End If
    Else
        MsgBox "No TV Station Selected!"
    End If
End Sub

Private Sub Cmd_Remove_ALL_Click()
    Dim i As Integer
    For i = 1 To lst_selected_tv_station.ListCount
        lst_tv_station_catalog.AddItem lst_selected_tv_station.List(0)
        lst_selected_tv_station.RemoveItem (0)
    Next
End Sub

Private Sub Cmd_Close_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    strBrandCode = Left(Frm_MPActivityDetail.txtMPNumber.Text, 4)
    Call Load_TV_Station
End Sub

Sub Load_TV_Station()
    Dim rsTemp As New ADODB.Recordset
    Dim strSql As String
    lst_tv_station_catalog.Clear
    strSql = "select station_code from tv_station WHERE Station_Code = On_Air_TV_Station order by station_code"
    rsTemp.Open strSql, ConnERP, 1, 3
    While Not rsTemp.EOF
        lst_tv_station_catalog.AddItem rsTemp("station_code")
        rsTemp.MoveNext
    Wend
    rsTemp.Close
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
        Case enButtonType.bieAdd  'call db_New.
            Call Cmd_Save_Click
        Case Else
            Call Cmd_Close_Click
    End Select

End Sub
