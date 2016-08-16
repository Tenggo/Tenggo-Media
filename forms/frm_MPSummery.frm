VERSION 5.00
Object = "{5A6B4ED6-5C98-4B3E-A352-41BE5E8D78D2}#1.0#0"; "ExPivot.dll"
Object = "{ED7B66F6-C533-48E9-B536-6F3B0E9C5839}#1.0#0"; "ExPrint.dll"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "mscomctl.OCX"
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "ComDlg32.OCX"
Begin VB.Form frm_MPSummery 
   BorderStyle     =   0  'None
   Caption         =   "Coporate Report"
   ClientHeight    =   8460
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   11190
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   8460
   ScaleWidth      =   11190
   ShowInTaskbar   =   0   'False
   Begin VB.PictureBox picStatusBar 
      Align           =   2  'Align Bottom
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00F0F0F0&
      BorderStyle     =   0  'None
      ClipControls    =   0   'False
      ForeColor       =   &H80000008&
      Height          =   330
      Left            =   0
      ScaleHeight     =   330
      ScaleWidth      =   11190
      TabIndex        =   8
      Top             =   8130
      Width           =   11190
      Begin MSComctlLib.ProgressBar ProgressBar1 
         Height          =   240
         Left            =   1065
         TabIndex        =   9
         Top             =   15
         Visible         =   0   'False
         Width           =   3930
         _ExtentX        =   6932
         _ExtentY        =   423
         _Version        =   393216
         Appearance      =   0
      End
      Begin VB.Label lblStatus 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Status"
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   60
         TabIndex        =   10
         Tag             =   "Last Modified Date: "
         Top             =   75
         Width           =   2520
      End
   End
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
      ScaleWidth      =   11190
      TabIndex        =   0
      Top             =   0
      Width           =   11190
      Begin VB.PictureBox picButton 
         BackColor       =   &H00F0F0F0&
         BorderStyle     =   0  'None
         ClipControls    =   0   'False
         Height          =   750
         Index           =   9
         Left            =   9270
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
         Height          =   750
         Index           =   20
         Left            =   1620
         ScaleHeight     =   750
         ScaleWidth      =   1500
         TabIndex        =   6
         Top             =   0
         Width           =   1500
      End
      Begin VB.PictureBox picButton 
         Appearance      =   0  'Flat
         BackColor       =   &H00F0F0F0&
         BorderStyle     =   0  'None
         ClipControls    =   0   'False
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
         Height          =   750
         Index           =   18
         Left            =   3150
         ScaleHeight     =   750
         ScaleWidth      =   1500
         TabIndex        =   5
         Top             =   0
         Width           =   1500
      End
      Begin VB.PictureBox picButton 
         Appearance      =   0  'Flat
         BackColor       =   &H00F0F0F0&
         BorderStyle     =   0  'None
         ClipControls    =   0   'False
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
         Height          =   750
         Index           =   19
         Left            =   4680
         ScaleHeight     =   750
         ScaleWidth      =   1500
         TabIndex        =   4
         Top             =   0
         Width           =   1500
      End
      Begin VB.PictureBox picButton 
         Appearance      =   0  'Flat
         BackColor       =   &H00F0F0F0&
         BorderStyle     =   0  'None
         ClipControls    =   0   'False
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
         Height          =   750
         Index           =   8
         Left            =   6210
         ScaleHeight     =   750
         ScaleWidth      =   1500
         TabIndex        =   3
         Top             =   0
         Width           =   1500
      End
      Begin VB.PictureBox picButton 
         Appearance      =   0  'Flat
         BackColor       =   &H00F0F0F0&
         BorderStyle     =   0  'None
         ClipControls    =   0   'False
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
         Height          =   750
         Index           =   22
         Left            =   7740
         ScaleHeight     =   750
         ScaleWidth      =   1500
         TabIndex        =   2
         Top             =   0
         Width           =   1500
      End
      Begin VB.PictureBox picButton 
         Appearance      =   0  'Flat
         BackColor       =   &H00F0F0F0&
         BorderStyle     =   0  'None
         ClipControls    =   0   'False
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
         Height          =   750
         Index           =   21
         Left            =   90
         ScaleHeight     =   750
         ScaleWidth      =   1500
         TabIndex        =   1
         Top             =   0
         Width           =   1500
      End
   End
   Begin Threed.SSPanel pnl_Main 
      Align           =   1  'Align Top
      Height          =   7140
      Left            =   0
      TabIndex        =   11
      Top             =   750
      Width           =   11190
      _Version        =   65536
      _ExtentX        =   19738
      _ExtentY        =   12594
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
      BorderWidth     =   0
      BevelOuter      =   1
      Begin VB.Frame Frame_View 
         Height          =   570
         Left            =   5340
         TabIndex        =   12
         Top             =   -630
         Width           =   3465
         Begin VB.CheckBox cbkFee 
            Caption         =   "Fee"
            ForeColor       =   &H000000FF&
            Height          =   210
            Left            =   1605
            TabIndex        =   14
            Top             =   225
            Width           =   615
         End
         Begin VB.CheckBox cbkOtherCost 
            Caption         =   "Other Cost"
            ForeColor       =   &H000000FF&
            Height          =   210
            Left            =   2295
            TabIndex        =   13
            Top             =   225
            Width           =   1140
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "View : Nett +"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   150
            TabIndex        =   15
            Top             =   210
            Width           =   1290
         End
      End
      Begin MSComDlg.CommonDialog CommonDialog1 
         Left            =   15
         Top             =   630
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
      Begin EXPIVOTLibCtl.Pivot Pivot1 
         Height          =   6480
         Left            =   135
         OleObjectBlob   =   "frm_MPSummery.frx":0000
         TabIndex        =   16
         Top             =   150
         Width           =   7695
      End
      Begin EXPRINTLibCtl.Print Print1 
         Left            =   8925
         OleObjectBlob   =   "frm_MPSummery.frx":1487
         Top             =   345
      End
   End
End
Attribute VB_Name = "frm_MPSummery"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'***********************************************************
'Submodul Name      : Frm_Rpt_Approval_Timesheet
'Submodul Function  : To preview report Approval Timesheet.
'Used Table         : [Client], [User_ID], [Ts_Date], [Ts_LeaveDate], [Ts_TimesheetDetail_Master], [Ts_Title], [Product].
'Used SP/View       : [vRpt_Time_Utilization_Client]
'Procedure/Function : GetCriteriaFilter, IsValidData, PreviewToPivot, SetButtonToolbar.
'Programmer Name    : {73 64 6B}
'Date               : 05-May-2015
'Last Update/By     : -
'Date Update        : -
'Log Update/By      : -
'***************************************************************
Option Explicit
Dim strSql As String
Dim dbl_ReportID As Double
Dim str_Democode As String
Public rst_empDayPart As New ADODB.Recordset
Dim rec_Report As New ADODB.Recordset
Public str_mp_number As String
Dim rsSummary As New ADODB.Recordset
Dim blnMoveGone As Boolean ' Refresh

Private Sub CreateLvChannel()
'*****************************************
'Procedure Name     : CreateLvChannel
'Procedure Function : Store data to List View Channel.
'Input Parameter    : -
'Output Parameter   : -
'Last Update Date   : Januari 2016
'Last Update By     : Ryo / Kreatif
'*********************************

    Dim li_Channel As ListItem
    Dim rec_Channel As New ADODB.Recordset

    'Open Recordset Brand
    strSql = "SELECT channel_code,Channel FROM ChannelCFG ORDER BY Channel"
    rec_Channel.Open strSql, connMacthIT, adOpenForwardOnly, adLockReadOnly, adCmdText

    'Load To ListView --> lv_Month
    lvw_channel.ListItems.Clear
    
    Do While Not rec_Channel.BOF And Not rec_Channel.EOF
        Set li_Channel = lvw_channel.ListItems.Add
        li_Channel.Text = rec_Channel("Channel").Value
        li_Channel.Tag = rec_Channel("Channel").Value
        rec_Channel.MoveNext
    Loop

    rec_Channel.Close

    Set rec_Channel = Nothing

End Sub

Private Sub CreateLvBrand()
'*****************************************
'Procedure Name     : CreateLvBrand
'Procedure Function : Store data to List View Brand.
'Input Parameter    : -
'Output Parameter   : -
'Last Update Date   : Januari 2016
'Last Update By     : Ryo / Kreatif
'*********************************

    Dim li_Month As ListItem
    Dim rec_Brand As New ADODB.Recordset

    'Open Recordset Brand
      strSql = "SELECT Brand_Code,Brand_Name FROM Brand WHERE Client_Code IN(" & getListClient & ") ORDER BY Brand_Name"
    rec_Brand.Open strSql, ConnERP, adOpenForwardOnly, adLockReadOnly, adCmdText

    'Load To ListView --> lv_Month
    lvw_Brand.ListItems.Clear
    Do While Not rec_Brand.BOF And Not rec_Brand.EOF
        Set li_Month = lvw_Brand.ListItems.Add
        li_Month.Text = rec_Brand("Brand_Name").Value
        li_Month.Tag = rec_Brand("Brand_code").Value
        rec_Brand.MoveNext
    Loop

    CloseRecordset rec_Brand

    Set rec_Brand = Nothing
    
End Sub

Private Sub CreateLvClient()
'*****************************************
'Procedure Name     : CreateLvClient
'Procedure Function : Store data to List View Client.
'Input Parameter    : -
'Output Parameter   : -
'Last Update Date   : Januari 2016
'Last Update By     : Ryo / Kreatif
'*********************************
    
    Dim li_Client As ListItem
    Dim rec_Client As New ADODB.Recordset

    strSql = "SELECT * FROM client WHERE Status='Active' ORDER BY Client_Name"
    rec_Client.Open strSql, ConnERP, adOpenForwardOnly, adLockReadOnly, adCmdText

    'Load To ListView --> lv_Month
    lvw_Client.ListItems.Clear
    Do While Not rec_Client.BOF And Not rec_Client.EOF
        Set li_Client = lvw_Client.ListItems.Add
        li_Client.Text = rec_Client("Client_Name").Value
        li_Client.Tag = rec_Client("Client_Code").Value
        rec_Client.MoveNext
    Loop
    CloseRecordset rec_Client
    
End Sub

Private Function getListChannel() As String
'*****************************************
'Procedure Name     : getListChannel
'Procedure Function : Mengambil Data Channel Yang Di Centang (All)
'Input Parameter    : -
'Output Parameter   : -
'Last Update Date   : Januari 2016
'Last Update By     : Ryo / Kreatif
'*********************************

    Dim int_Count As Integer
    Dim str_Result As String
    
    str_Result = ""
    
    For int_Count = 1 To lvw_channel.ListItems.Count
        If chk_AllMonth.Value = 1 Then
            If str_Result = "" Then
                str_Result = "'" & lvw_channel.ListItems(int_Count).Tag & "',"
                Else
                    str_Result = str_Result & "'" & lvw_channel.ListItems(int_Count).Tag & "',"
            End If
        ElseIf lvw_channel.ListItems(int_Count).Checked Then
            If str_Result = "" Then
                str_Result = "'" & lvw_channel.ListItems(int_Count).Tag & "',"
                Else
                    str_Result = str_Result & "'" & lvw_channel.ListItems(int_Count).Tag & "',"
    
            End If
        End If
    Next int_Count
    If str_Result = "" Then
        getListChannel = "''"
        Else
            getListChannel = Mid(str_Result, 1, Len(str_Result) - 1)
    End If

End Function

Private Function getListClient() As String
'*****************************************
'Procedure Name     : getListClient
'Procedure Function : Mengambil Data cLIENT Yang Di Centang (All)
'Input Parameter    : -
'Output Parameter   : -
'Last Update Date   : Januari 2016
'Last Update By     : Ryo / Kreatif
'*********************************

    Dim int_Count As Integer
    Dim str_Result As String
    
    str_Result = ""
    For int_Count = 1 To lvw_Client.ListItems.Count
        If Chk_AllClient.Value = 1 Then
            If str_Result = "" Then
                str_Result = "'" & lvw_Client.ListItems(int_Count).Tag & "',"
            Else
                str_Result = str_Result & "'" & lvw_Client.ListItems(int_Count).Tag & "',"
            End If
        ElseIf lvw_Client.ListItems(int_Count).Checked Then
            If str_Result = "" Then
                str_Result = "'" & lvw_Client.ListItems(int_Count).Tag & "',"
            Else
                str_Result = str_Result & "'" & lvw_Client.ListItems(int_Count).Tag & "',"

            End If
        End If
    Next int_Count
    
    If str_Result = "" Then
        getListClient = "''"
        Else
            getListClient = Mid(str_Result, 1, Len(str_Result) - 1)
    End If

End Function

Private Function getListBrand() As String
'*****************************************
'Procedure Name     : getListBrand
'Procedure Function : Mengambil Data list Brand Yang Di Centang (All)
'Input Parameter    : -
'Output Parameter   : -
'Date               : 12-Apr-2015/{73 64 6B}
'*****************************************
    
    Dim int_Count As Integer
    Dim str_Result As String
    
    str_Result = ""
    For int_Count = 1 To lvw_Brand.ListItems.Count
        If Chk_AllBrand.Value = 1 Then
            If str_Result = "" Then
                str_Result = "'" & lvw_Brand.ListItems(int_Count).Tag & "',"
                Else
                    str_Result = str_Result & "'" & lvw_Brand.ListItems(int_Count).Tag & "',"
            End If
        ElseIf lvw_Brand.ListItems(int_Count).Checked Then
            If str_Result = "" Then
                str_Result = "'" & lvw_Brand.ListItems(int_Count).Tag & "',"
                Else
                    str_Result = str_Result & "'" & lvw_Brand.ListItems(int_Count).Tag & "',"
    
            End If
        End If
    Next int_Count
    If str_Result = "" Then
        getListBrand = "''"
    Else
        getListBrand = Mid(str_Result, 1, Len(str_Result) - 1)
    End If

End Function

Private Sub CreateLvDemo()
'*****************************************
'Procedure Name     : CreateLvDemo
'Procedure Function : Create Listview Demographic
'Input Parameter    : -
'Output Parameter   : -
'Last Update Date   : Januari 2016
'Last Update By     : Ryo / Kreatif
'*********************************

    Dim rec_Demo As New ADODB.Recordset
    Dim li_Demo As ListItem
    
    strSql = "SELECT * FROM Demographics ORDER BY target_Name"
    rec_Demo.Open strSql, connMacthIT, adOpenForwardOnly, adLockReadOnly, adCmdText
    lvw_demo.ListItems.Clear
    
    Do While Not rec_Demo.BOF And Not rec_Demo.EOF
        Set li_Demo = lvw_demo.ListItems.Add
        li_Demo.Text = rec_Demo("target_Name").Value
        li_Demo.Tag = rec_Demo("target_ID").Value
        rec_Demo.MoveNext
    Loop

    rec_Demo.Close

    Set rec_Demo = Nothing
        
End Sub

Private Function getListDemo() As String
'*****************************************
'Procedure Name     : getListMonthYear
'Procedure Function : Mengambil Data Week Yang Di Centang / All Jika CheckBox di Centang
'Input Parameter    : -
'Output Parameter   : -
'Date               : 12-Apr-2015/{73 64 6B}
'*****************************************
    
    Dim int_Count As Integer
    Dim str_Result As String
    
    str_Result = ""
    
    For int_Count = 1 To lvw_demo.ListItems.Count
        If chk_AllWeek.Value = 1 Then
            
            If str_Result = "" Then
                str_Result = "'" & lvw_demo.ListItems(int_Count).Text & "',"
            Else
                str_Result = str_Result & "'" & lvw_demo.ListItems(int_Count).Text & "',"
            End If
            
        ElseIf lvw_demo.ListItems(int_Count).Checked Then
            
            If str_Result = "" Then
                str_Result = "'" & lvw_demo.ListItems(int_Count).Text & "',"
            Else
                str_Result = str_Result & "'" & lvw_demo.ListItems(int_Count).Text & "',"
            End If
        
        End If
    Next int_Count
    If str_Result = "" Then
        getListDemo = "''"
        Else
            getListDemo = Mid(str_Result, 1, Len(str_Result) - 1)
    End If

End Function

Private Sub cbo_Select_Demo_Change()
'*****************************************
'Procedure Name     : cbo_Select_Demo_Change
'Procedure Function : panggil procedure PreviewToPivot
'Input Parameter    : -
'Output Parameter   : -
'Last Update Date   : Januari 2016
'Last Update By     : Ryo / Kreatif
'*********************************

    PreviewToPivot

End Sub

Private Sub cbo_Select_Demo_Click()
'*****************************************
'Procedure Name     : cbo_Select_Demo_Click
'Procedure Function : panggil procedure PreviewToPivot
'Input Parameter    : -
'Output Parameter   : -
'Last Update Date   : Januari 2016
'Last Update By     : Ryo / Kreatif
'*********************************
    
    PreviewToPivot

End Sub

Private Sub Chk_AllBrand_Click()
'*****************************************
'Procedure Name     : Chk_AllBrand_Click
'Procedure Function : Check All Brand
'Input Parameter    : -
'Output Parameter   : -
'Last Update Date   : Januari 2016
'Last Update By     : Ryo / Kreatif
'*********************************

    If Chk_AllBrand.Value = 1 Then
        releaseCheckBrand (True)
        lvw_Brand.Enabled = False
    Else
        releaseCheckBrand (False)
        lvw_Brand.Enabled = True
    End If
    chk_AllWeek.Value = 0

End Sub

Private Sub Chk_AllClient_Click()
'*****************************************
'Procedure Name     : Chk_AllBrand_Click
'Procedure Function : Check All Brand
'Input Parameter    : -
'Output Parameter   : -
'Last Update Date   : Januari 2016
'Last Update By     : Ryo / Kreatif
'*********************************

    If Chk_AllClient.Value = 1 Then
        releaseCheckClient (True)
        lvw_Client.Enabled = False
    Else
        releaseCheckClient (False)
        lvw_Client.Enabled = True
    End If
    CreateLvBrand
    Chk_AllBrand.Value = 0

End Sub

Private Sub chk_AllMonth_Click()
'*****************************************
'Procedure Name     : chk_AllMonth_Click
'Procedure Function :
'Input Parameter    : -
'Output Parameter   : -
'Last Update Date   : Januari 2016
'Last Update By     : Tedi / Kreatif
'*********************************

   If chk_AllMonth.Value = 1 Then
        releaseCheckMonth (True)
        lvw_channel.Enabled = False
        Else
        releaseCheckMonth (False)
            lvw_channel.Enabled = True
    End If
     chk_AllWeek.Value = 0

End Sub

Private Sub releaseCheckMonth(chk As Boolean)
'*****************************************
'Procedure Name     : releaseCheckMonth
'Procedure Function : Menghapus Tanda Centang pada list view Month
'Input Parameter    : -
'Output Parameter   : -
'Last Update Date   : Januari 2016
'Last Update By     : Ryo / Kreatif
'*********************************

    Dim int_Count As Integer
    For int_Count = 1 To lvw_channel.ListItems.Count
        lvw_channel.ListItems(int_Count).Checked = chk
    Next int_Count

End Sub

Private Sub chk_AllWeek_Click()
'*****************************************
'Procedure Name     : chk_AllWeek_Click
'Procedure Function :
'Input Parameter    : -
'Output Parameter   : -
'Last Update Date   : Januari 2016
'Last Update By     : Ryo / Kreatif
'*********************************

    If chk_AllWeek.Value = 1 Then
        releaseCheckWeek (True)
        lvw_demo.Enabled = False
    Else
        releaseCheckWeek (False)
        lvw_demo.Enabled = True
    End If

End Sub
Private Sub releaseCheckWeek(chk As Boolean)
'*****************************************
'Procedure Name     : releaseCheckWeek
'Procedure Function : Menghapus Tanda Centang pada list view Week
'Input Parameter    : -
'Output Parameter   : -
'Last Update Date   : Januari 2016
'Last Update By     : Ryo / Kreatif

    Dim int_Count As Integer
    
    For int_Count = 1 To lvw_demo.ListItems.Count
        lvw_demo.ListItems(int_Count).Checked = int_Count
    Next int_Count

End Sub

Private Sub releaseCheckBrand(chk As Boolean)
'*****************************************
'Procedure Name     : releaseCheckBrand
'Procedure Function : Realease Listview Brand
'Input Parameter    : -
'Output Parameter   : -
'Last Update Date   : Januari 2016
'Last Update By     : Ryo / Kreatif
'*********************************
    
    Dim int_Count As Integer
    
    For int_Count = 1 To lvw_Brand.ListItems.Count
        lvw_Brand.ListItems(int_Count).Checked = chk
    Next int_Count

End Sub

Private Sub releaseCheckClient(chk As Boolean)
'*****************************************
'Procedure Name     : releaseCheckBrand
'Procedure Function : Realease Listview Brand
'Input Parameter    : -
'Output Parameter   : -
'Last Update Date   : Januari 2016
'Last Update By     : Ryo / Kreatif
'*********************************

    Dim int_Count As Integer
    
    For int_Count = 1 To lvw_Client.ListItems.Count
        lvw_Client.ListItems(int_Count).Checked = chk
    Next int_Count

End Sub


Private Sub Form_Load()
'*****************************************
'Procedure Name     : Form_Load
'Procedure Function : Call EnableObject(False) utk form idle mode.
'Input Parameter    : --
'Output Parameter   : ---
'Last Update Date   : Januari 2016
'Last Update By     : Ryo / Kreatif
'*********************************

    Dim rec_Client As New ADODB.Recordset
    Call SetButtonToolbar(True) 'TOOLBAR_AI.

'    CreateLvClient
'    CreateLvChannel
'    CreateLvDemo
'    create_table_DayPart

    Call PivotVisualDesign(Pivot1)
    'If str_mp_number = "" Then str_mp_number = frmDefault.cboMPNumber.Text
    
    Call show_Summary_Plan(str_mp_number)
   ' Call Set_Default_Pivot_Template
    lblStatus.AutoSize = True
    lblStatus.Tag = "You can start with defining the filter criteria -> Previewing -> Choosing Default Template or User Template -> Redesigning the Pivot Table and then Save Template -> Dragging and dropping/Export data to Excel or just Print it."
    lblStatus.Caption = lblStatus.Tag
    Set_Default_Pivot_Template
    blnMoveGone = False
End Sub

Sub Set_Default_Pivot_Template()
    '== Checking Existing User Template From Database
    Dim rec_Template As New ADODB.Recordset
    
    strSql = ""
    strSql = strSql & "SELECT * "
    strSql = strSql & "FROM TEMP_PIVOT "
    strSql = strSql & "WHERE ltrim(rtrim(username))='" & strLogin_User & "' "
    strSql = strSql & "AND filename='default.SUMP' ORDER BY STAMP DESC "
    
    rec_Template.Open strSql, ConnERP, adOpenStatic, adLockReadOnly
     
'    ProgressBar1.Value = ProgressBar1.Value + 1
    '== Pivot layout
'    Call PivotPreview(Pivot1, rec_Report)
'    ProgressBar1.Value = ProgressBar1.Value + 1
    '== Switch To Template

    If rec_Template.RecordCount < 1 Then
        '-- Default Template
        picButton_Click (1)
    Else
'        rec_Template.MoveFirst
'        '--- Load Last Update Template
        Pivot1.Layout = rec_Template!temp_contain
        'strRefTemplate = Trim(rec_Template!FileName)
        'CekDir "CRPT"
        'createTemplateFile rec_Template
    End If
  
    Call CloseRecordset(rec_Template)
'    ProgressBar1.Visible = False
    Me.MousePointer = 0


End Sub

Sub Set_Reset_Pivot_Template()
    '== Checking Existing User Template From Database
    Dim rec_Template As New ADODB.Recordset
    
    strSql = ""
    strSql = strSql & "SELECT TOP(1) * "
    strSql = strSql & "FROM TEMP_PIVOT "
    strSql = strSql & "WHERE ltrim(rtrim(username))='" & strLogin_User & "' "
    strSql = strSql & "AND filename='reset.SUMP' ORDER BY STAMP DESC "
    
    rec_Template.Open strSql, ConnERP, adOpenStatic, adLockReadOnly
     
'    ProgressBar1.Value = ProgressBar1.Value + 1
    '== Pivot layout
'    Call PivotPreview(Pivot1, rec_Report)
'    ProgressBar1.Value = ProgressBar1.Value + 1
    '== Switch To Template

    If rec_Template.RecordCount < 1 Then
        '-- Default Template
        picButton_Click (1)
    Else
'        rec_Template.MoveFirst
'        '--- Load Last Update Template
        Pivot1.Layout = rec_Template!temp_contain
        'strRefTemplate = Trim(rec_Template!FileName)
        'CekDir "CRPT"
        'createTemplateFile rec_Template
    End If
  
   ' Call CloseRecordset(rec_Template)
'    ProgressBar1.Visible = False
    Me.MousePointer = 0


End Sub


Private Function Get_Report_Id() As Double
'*****************************************
'Procedure Name     : Get_Report_Id
'Procedure Function : --
'Input Parameter    : --
'Output Parameter   : ---
'Last Update Date   : Januari 2016
'Last Update By     : Ryo / Kreatif
'*********************************

    Dim Out_Param_Spot_Id As ADODB.Parameter
    Dim cmd As New ADODB.Command
        
    cmd.CommandType = adCmdStoredProc
    cmd.CommandText = "Get_Spot_id"
    
    
    Set Out_Param_Spot_Id = cmd.CreateParameter("New_Spot_Id", adDouble, adParamOutput)
    
    'Append Parameter
    
    cmd.Parameters.Append Out_Param_Spot_Id
    
    'Execute Command
    cmd.ActiveConnection = connMacthIT
    cmd.Execute
    
    Get_Report_Id = Out_Param_Spot_Id.Value
    
End Function

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
'*****************************************
'Procedure Name     : Form_QueryUnload
'Procedure Function : ~ Lebih fleksibel pake Form_QueryUnload drpd Form_Unload krn ada param Cancel dan UnloadMode yg bisa digunakan utk kondisi tertentu.
'Last Update Date   : Januari 2016
'Last Update By     : Ryo / Kreatif
'*********************************

'    Dim int_Idx As Integer
'
'    For int_Idx = 0 To cbo_Select_Demo.ListCount - 1
'        connMacthIT.Execute "DELETE FROM Temp_Report_Spot WHERE Report_ID=" & cbo_Select_Demo.ItemData(int_Idx)
'        connMacthIT.Execute "DELETE FROM Temp_Report_Reach WHERE Report_ID=" & cbo_Select_Demo.ItemData(int_Idx)
'    Next int_Idx
'    ReleaseCapture 'The MOUSE_LEAVE pseudo-event.
'
'    strSQL = ""
'
'    Cancel = False
    CloseRecordset rsSummary

End Sub

Sub AdjustSizeForm()
'*****************************************
'Procedure Name     : AdjustSizeForm
'Procedure Function : Adjust Size Form
'Input Parameter    : -
'Output Parameter   : -
'Last Update Date   : Januari 2016
'Last Update By     : Ryo / Kreatif
'*********************************
    Dim nWidth As Single, nHeight As Single

    'On Local Error Resume Next
    Me.Top = 0
    Me.Left = 0
    Me.Width = mdi_Main.ScaleWidth
    Me.Height = mdi_Main.ScaleHeight

    nWidth = Me.ScaleWidth: nHeight = Me.ScaleHeight

    With pnl_Main
        .Move .Left, .Top, nWidth - (.Left * 2), nHeight - .Top - picStatusBar.Height
    End With
    'pnl_Filter.Height = pnl_Main.Height - (pnl_Filter.Top * 2)
    'pnlPivot.Left = pnl_Hiding.Left + pnl_Hiding.Width
    'pnlPivot.Width = pnl_Main.Width - pnl_Filter.Left - pnl_Filter.Width - pnl_Hiding.Width
    'pnlPivot.Height = pnl_Filter.Height
    'Pivot1.Width = pnlPivot.Width - (Pivot1.Left * 2)
'    sstab_Data.Width = pnl_Main.Width - (sstab_Data.Left * 2)
'    sstab_Data.Height = pnl_Main.Height - (sstab_Data.Top + 50)
    Pivot1.Width = Me.Width - (Pivot1.Left * 2)
    Pivot1.Height = Me.Height - (Pivot1.Top) - picStatusBar.Height - 900
    
   On Error GoTo 0

End Sub

Private Sub Form_Resize()
'*****************************************
'Procedure Name     : Form_Resize
'Procedure Function : Call AdjustSizeForm to adjust form and object
'Input Parameter    : -
'Output Parameter   : -
'Last Update Date   : Januari 2016
'Last Update By     : Ryo / Kreatif
'*********************************
    
    AdjustSizeForm
    On Local Error GoTo 0
    
End Sub


 
 
Private Sub show_Summary_Plan(strMPNumber As String)
    Dim strSql As String, i As Integer, strTaskID As String, j As Integer

    Dim intTaskNumber As Integer, intTaskRow As Integer, intPrintRow As Integer
    
    Dim TotalTV(13) As Double
    Dim TotalRD(13) As Double
    Dim TotalPR(13) As Double
    Dim TotalCN(13) As Double
    Dim TotalOT(13) As Double
    Dim GrandTotal(13) As Double
'    'set value
'    For i = 1 To 13
'        TotalTV(i - 1) = 0
'        TotalRD(i - 1) = 0
'        TotalPR(i - 1) = 0
'        TotalCN(i - 1) = 0
'        TotalOT(i - 1) = 0
'        GrandTotal(i - 1) = 0
'    Next
    
    strSql = ""
    strSql = strSql & "SELECT mp_task_id,task_desc,brand_variant_code,brand_variant_name,"
    strSql = strSql & " CASE WHEN medium_code='CN' THEN 'CINEMA'"
    strSql = strSql & " WHEN medium_code='RD' THEN 'RADIO'"
    strSql = strSql & " WHEN medium_code='OT' THEN 'OTHER'"
    strSql = strSql & " WHEN medium_code='PR' THEN 'PRINT'"
    strSql = strSql & " WHEN medium_code='TV' THEN 'TV' END AS medium,"
    strSql = strSql & " CASE WHEN month_number=1 THEN 'JANUARY'"
    strSql = strSql & " WHEN month_number=2 THEN 'FEBRUARY'"
    strSql = strSql & " WHEN month_number=3 THEN 'MARCH'"
    strSql = strSql & " WHEN month_number=4 THEN 'APRIL'"
    strSql = strSql & " WHEN month_number=5 THEN 'MAY'"
    strSql = strSql & " WHEN month_number=6 THEN 'JUNE'"
    strSql = strSql & " WHEN month_number=7 THEN 'JULY'"
    strSql = strSql & " WHEN month_number=8 THEN 'AUGUST'"
    strSql = strSql & " WHEN month_number=9 THEN 'SEPTEMBER'"
    strSql = strSql & " WHEN month_number=10 THEN 'OCTOBER'"
    strSql = strSql & " WHEN month_number=11 THEN 'NOVEMBER'"
    strSql = strSql & " WHEN month_number=12 THEN 'DECEMBER'"
    strSql = strSql & " END AS Months, month_number,sum(plan_budget)AS [Plan Budget],"
    strSql = strSql & " max(is_actual) as is_actual,sum(actual_budget) as [Actual Budget],sum(plan_budget)-sum(actual_budget) AS Balance"
    strSql = strSql & " FROM (SELECT d.mp_task_id,d.task_desc,b.medium_code,c.brand_variant_code,c.brand_variant_name,"
    strSql = strSql & " a.month_number,sum(a.nett_only) plan_budget,"
    strSql = strSql & " max(a.is_actual) is_actual ,0 AS actual_budget"
    strSql = strSql & " FROM ("
    strSql = strSql & " select mp_medium_id,month_number,"
    strSql = strSql & " Case IsNull(Total_Actual, -1) when -1 then min_budget  Else Actual_Nett_Paid + Actual_Nett_Bonus end Nett_only,"
    strSql = strSql & " Case IsNull(Total_Actual, -1) when -1 then min_budget + MSC_Paid_Value + MSC_Bonus_Value + Club_Agency_Value  Else Actual_Nett_Paid + Actual_Nett_Bonus + Actual_MSC_Paid + Actual_MSC_Bonus end Nett_plus_fee,"
    strSql = strSql & " Case IsNull(Total_Actual, -1) when -1 then min_budget + MSC_Paid_Value + MSC_Bonus_Value + Club_Agency_Value + isnull(other_cost,0) Else Actual_Nett_Paid + Actual_Nett_Bonus + Actual_MSC_Paid + Actual_MSC_Bonus  + isnull(actual_other_cost,0) end Nett_fee_other,"
    strSql = strSql & " Case IsNull(Total_Actual, -1) when -1 then min_budget  + isnull(other_cost,0) Else Actual_Nett_Paid + Actual_Nett_Bonus + isnull(actual_other_cost,0) end Nett_plus_other,"
    strSql = strSql & " case isnull(total_actual,-1) when -1 then 0 else 1 end is_actual"
    strSql = strSql & " From mp_monthly_activity"
    strSql = strSql & " ) a  INNER JOIN mp_medium b on a.mp_medium_id = b.mp_medium_id"
    strSql = strSql & " INNER JOIN mp_activity c on b.mp_activity_id = c.mp_activity_id"
    strSql = strSql & " INNER JOIN mp_task d on c.mp_task_id = d.mp_task_id and d.mp_number = '" & strMPNumber & "'"
    strSql = strSql & " GROUP BY d.mp_task_id,d.task_desc,b.medium_code,c.brand_variant_code,c.brand_variant_name,a.month_number"
    strSql = strSql & " Union All"
    strSql = strSql & " SELECT d.mp_task_id,d.task_desc,b.medium_code,c.brand_variant_code,c.brand_variant_name,a.month_number,0,0"
    strSql = strSql & " ,sum(isnull(a.actual_nett_paid,0) + isnull(a.actual_nett_bonus,0)) actual_budget"
    strSql = strSql & " FROM mp_monthly_activity a  INNER JOIN mp_medium b on a.mp_medium_id = b.mp_medium_id"
    strSql = strSql & " INNER JOIN mp_activity c on b.mp_activity_id = c.mp_activity_id"
    strSql = strSql & " INNER JOIN mp_task d on c.mp_task_id = d.mp_task_id and d.mp_number = '" & strMPNumber & "'"
    strSql = strSql & " GROUP BY d.mp_task_id,d.task_desc,b.medium_code,c.brand_variant_code,c.brand_variant_name,a.month_number"
    strSql = strSql & " )Z"
    strSql = strSql & " GROUP BY mp_task_id,task_desc,medium_code,brand_variant_code,brand_variant_name,       month_number"
    'rsSummary.Close
    '===============================================================Down List Begin
    strSql = "SELECT"
    strSql = strSql & " mp_task_id,"
    strSql = strSql & " task_desc,"
    strSql = strSql & " brand_variant_code,"
    strSql = strSql & " brand_variant_name,title,"
    strSql = strSql & " CASE"
    strSql = strSql & " WHEN medium_code = 'CN' THEN 'CINEMA'"
    strSql = strSql & " WHEN medium_code = 'RD' THEN 'RADIO'"
    strSql = strSql & " WHEN medium_code = 'OT' THEN 'OTHER'"
    strSql = strSql & " WHEN medium_code = 'PR' THEN 'PRINT'"
    strSql = strSql & " WHEN medium_code = 'TV' THEN 'TV'"
    strSql = strSql & " END AS medium,"
    strSql = strSql & " CASE"
    strSql = strSql & " WHEN month_number = 1 THEN 'JANUARY'"
    strSql = strSql & " WHEN month_number = 2 THEN 'FEBRUARY'"
    strSql = strSql & " WHEN month_number = 3 THEN 'MARCH'"
    strSql = strSql & " WHEN month_number = 4 THEN 'APRIL'"
    strSql = strSql & " WHEN month_number = 5 THEN 'MAY'"
    strSql = strSql & " WHEN month_number = 6 THEN 'JUNE'"
    strSql = strSql & " WHEN month_number = 7 THEN 'JULY'"
    strSql = strSql & " WHEN month_number = 8 THEN 'AUGUST'"
    strSql = strSql & " WHEN month_number = 9 THEN 'SEPTEMBER'"
    strSql = strSql & " WHEN month_number = 10 THEN 'OCTOBER'"
    strSql = strSql & " WHEN month_number = 11 THEN 'NOVEMBER'"
    strSql = strSql & " WHEN month_number = 12 THEN 'DECEMBER'"
    strSql = strSql & " END AS Months,"
    strSql = strSql & " month_number,"
    strSql = strSql & " amount ,"
    strSql = strSql & " is_actual "
    
    strSql = strSql & " FROM (SELECT"
    strSql = strSql & " d.mp_task_id,"
    strSql = strSql & " d.task_desc,"
    strSql = strSql & " b.medium_code,"
    strSql = strSql & " c.brand_variant_code,"
    strSql = strSql & " c.brand_variant_name,"
    strSql = strSql & " a.month_number,"
    strSql = strSql & " SUM(a.nett_only) amount,"
    strSql = strSql & " MAX(a.is_actual) is_actual,"
    strSql = strSql & "'1-Plan' as title,"
    strSql = strSql & " 0,0"
    strSql = strSql & " FROM (SELECT"
    strSql = strSql & " mp_medium_id,"
    strSql = strSql & " month_number,"
    strSql = strSql & " Case IsNull(Total_Actual, -1)"
    strSql = strSql & " WHEN -1 THEN min_budget"
    strSql = strSql & " Else actual_nett_paid +actual_nett_bonus"
    strSql = strSql & " END Nett_only,"
    strSql = strSql & " Case IsNull(Total_Actual, -1)"
    strSql = strSql & " WHEN -1 THEN min_budget + MSC_Paid_Value + MSC_Bonus_Value + Club_Agency_Value"
    strSql = strSql & " Else actual_nett_paid+ actual_nett_bonus + Actual_MSC_Paid + Actual_MSC_Bonus"
    strSql = strSql & " END Nett_plus_fee,"
    strSql = strSql & " Case IsNull(Total_Actual, -1)"
    strSql = strSql & " WHEN -1 THEN min_budget + MSC_Paid_Value + MSC_Bonus_Value + Club_Agency_Value + ISNULL(other_cost, 0)"
    strSql = strSql & " Else actual_nett_paid+ actual_nett_bonus + Actual_MSC_Paid + Actual_MSC_Bonus + IsNull(actual_other_cost, 0)"
    strSql = strSql & " END Nett_fee_other,"
    strSql = strSql & " Case IsNull(Total_Actual, -1)"
    strSql = strSql & " WHEN -1 THEN min_budget + ISNULL(other_cost, 0)"
    strSql = strSql & " Else actual_nett_paid+ actual_nett_bonus + IsNull(actual_other_cost, 0)"
    strSql = strSql & " END Nett_plus_other,"
    strSql = strSql & " Case IsNull(Total_Actual, -1)"
    strSql = strSql & "   WHEN -1 THEN 0"
    strSql = strSql & " ELSE 1"
    strSql = strSql & " END is_actual"
    strSql = strSql & " FROM mp_monthly_activity) a"
    strSql = strSql & " INNER JOIN mp_medium b"
    strSql = strSql & " ON a.mp_medium_id = b.mp_medium_id"
    strSql = strSql & " INNER JOIN mp_activity c"
    strSql = strSql & " ON b.mp_activity_id = c.mp_activity_id"
    strSql = strSql & " INNER JOIN mp_task d"
    strSql = strSql & " ON c.mp_task_id = d.mp_task_id"
    strSql = strSql & " AND d.mp_number = '" & strMPNumber & "'"
    strSql = strSql & " GROUP BY d.mp_task_id,"
    strSql = strSql & " d.task_desc,"
    strSql = strSql & " b.medium_code,"
    strSql = strSql & " c.brand_variant_code,"
    strSql = strSql & " c.brand_variant_name,"
    strSql = strSql & " a.month_number"
    strSql = strSql & " Union All"
    strSql = strSql & " SELECT"
    strSql = strSql & " d.mp_task_id,"
    strSql = strSql & " d.task_desc,"
    strSql = strSql & " b.medium_code,"
    strSql = strSql & " c.brand_variant_code,"
    strSql = strSql & " c.brand_variant_name,"
    strSql = strSql & " a.month_number,0,"
    strSql = strSql & " SUM(ISNULL(a.actual_nett_paid, 0) + ISNULL(a.actual_nett_bonus, 0)) actual_budget,0,'2-Actual' as title"
    strSql = strSql & " FROM mp_monthly_activity a"
    strSql = strSql & " INNER JOIN mp_medium b"
    strSql = strSql & " ON a.mp_medium_id = b.mp_medium_id"
    strSql = strSql & " INNER JOIN mp_activity c"
    strSql = strSql & " ON b.mp_activity_id = c.mp_activity_id"
    strSql = strSql & " INNER JOIN mp_task d"
    strSql = strSql & " ON c.mp_task_id = d.mp_task_id"
    strSql = strSql & " AND d.mp_number = '" & strMPNumber & "'"
    strSql = strSql & " GROUP BY d.mp_task_id,"
    strSql = strSql & " d.task_desc,"
    strSql = strSql & " b.medium_code,"
    strSql = strSql & " c.brand_variant_code,"
    strSql = strSql & " c.brand_variant_name,"
    strSql = strSql & " a.month_number"
    strSql = strSql & " Union All"
    strSql = strSql & " SELECT"
    strSql = strSql & " d.mp_task_id,"
    strSql = strSql & " d.task_desc,"
    strSql = strSql & " b.medium_code,"
    strSql = strSql & " c.brand_variant_code,"
    strSql = strSql & " c.brand_variant_name,"
    strSql = strSql & " a.month_number,"
    
    strSql = strSql & " SUM(COALESCE(Total_Actual, min_budget, actual_nett_paid + actual_nett_bonus) - (IsNull(a.actual_nett_paid, 0) + IsNull(a.actual_nett_bonus, 0))) As Balance"
    strSql = strSql & " ,0,'3-Balance' as title"
    strSql = strSql & " FROM mp_monthly_activity a"
    strSql = strSql & " INNER JOIN mp_medium b"
    strSql = strSql & " ON a.mp_medium_id = b.mp_medium_id"
    strSql = strSql & " INNER JOIN mp_activity c"
    strSql = strSql & " ON b.mp_activity_id = c.mp_activity_id"
    strSql = strSql & " INNER JOIN mp_task d"
    strSql = strSql & " ON c.mp_task_id = d.mp_task_id"
    strSql = strSql & " AND d.mp_number = '" & strMPNumber & "'"
    strSql = strSql & " GROUP BY d.mp_task_id,"
    strSql = strSql & " d.task_desc,"
    strSql = strSql & " b.medium_code,"
    strSql = strSql & " c.brand_variant_code,"
    strSql = strSql & " c.brand_variant_name,"
    strSql = strSql & " a.month_number) Z"
    '===============================================================Down List Begin
    '===============================================================Down List Begin
    strSql = "SELECT"
    strSql = strSql & " mp_task_id,"
    strSql = strSql & " task_desc,"
    strSql = strSql & " brand_variant_code,"
    strSql = strSql & " brand_variant_name,title,"
    strSql = strSql & " CASE"
    strSql = strSql & " WHEN medium_code = 'CN' THEN 'CINEMA'"
    strSql = strSql & " WHEN medium_code = 'RD' THEN 'RADIO'"
    strSql = strSql & " WHEN medium_code = 'OT' THEN 'OTHER'"
    strSql = strSql & " WHEN medium_code = 'PR' THEN 'PRINT'"
    strSql = strSql & " WHEN medium_code = 'TV' THEN 'TV'"
    strSql = strSql & " END AS medium,"
    strSql = strSql & " CASE"
    strSql = strSql & " WHEN month_number = 1 THEN 'JANUARY'"
    strSql = strSql & " WHEN month_number = 2 THEN 'FEBRUARY'"
    strSql = strSql & " WHEN month_number = 3 THEN 'MARCH'"
    strSql = strSql & " WHEN month_number = 4 THEN 'APRIL'"
    strSql = strSql & " WHEN month_number = 5 THEN 'MAY'"
    strSql = strSql & " WHEN month_number = 6 THEN 'JUNE'"
    strSql = strSql & " WHEN month_number = 7 THEN 'JULY'"
    strSql = strSql & " WHEN month_number = 8 THEN 'AUGUST'"
    strSql = strSql & " WHEN month_number = 9 THEN 'SEPTEMBER'"
    strSql = strSql & " WHEN month_number = 10 THEN 'OCTOBER'"
    strSql = strSql & " WHEN month_number = 11 THEN 'NOVEMBER'"
    strSql = strSql & " WHEN month_number = 12 THEN 'DECEMBER'"
    strSql = strSql & " END AS Months,"
    strSql = strSql & " month_number,"
    strSql = strSql & " amount ,"
    strSql = strSql & " is_actual "
    
    strSql = strSql & " FROM (SELECT"
    strSql = strSql & " d.mp_task_id,"
    strSql = strSql & " d.task_desc,"
    strSql = strSql & " b.medium_code,"
    strSql = strSql & " c.brand_variant_code,"
    strSql = strSql & " c.brand_variant_name,"
    strSql = strSql & " a.month_number,"
    strSql = strSql & " SUM(a.nett_only) amount,"
    strSql = strSql & " MAX(a.is_actual) is_actual,"
    
    strSql = strSql & "'1-Plan' as title"
    strSql = strSql & " FROM (SELECT"
    strSql = strSql & " mp_medium_id,"
    strSql = strSql & " month_number,"
    strSql = strSql & " Case IsNull(Total_Actual, -1)"
    strSql = strSql & " WHEN -1 THEN min_budget"
    strSql = strSql & " Else actual_nett_paid +actual_nett_bonus"
    strSql = strSql & " END Nett_only,"
    strSql = strSql & " Case IsNull(Total_Actual, -1)"
    strSql = strSql & " WHEN -1 THEN min_budget + MSC_Paid_Value + MSC_Bonus_Value + Club_Agency_Value"
    strSql = strSql & " Else actual_nett_paid+ actual_nett_bonus + Actual_MSC_Paid + Actual_MSC_Bonus"
    strSql = strSql & " END Nett_plus_fee,"
    strSql = strSql & " Case IsNull(Total_Actual, -1)"
    strSql = strSql & " WHEN -1 THEN min_budget + MSC_Paid_Value + MSC_Bonus_Value + Club_Agency_Value + ISNULL(other_cost, 0)"
    strSql = strSql & " Else actual_nett_paid+ actual_nett_bonus + Actual_MSC_Paid + Actual_MSC_Bonus + IsNull(actual_other_cost, 0)"
    strSql = strSql & " END Nett_fee_other,"
    strSql = strSql & " Case IsNull(Total_Actual, -1)"
    strSql = strSql & " WHEN -1 THEN min_budget + ISNULL(other_cost, 0)"
    strSql = strSql & " Else actual_nett_paid+ actual_nett_bonus + IsNull(actual_other_cost, 0)"
    strSql = strSql & " END Nett_plus_other,"
    strSql = strSql & " Case IsNull(Total_Actual, -1)"
    strSql = strSql & "   WHEN -1 THEN 0"
    strSql = strSql & " ELSE 1"
    strSql = strSql & " END is_actual"
    strSql = strSql & " FROM mp_monthly_activity) a"
    strSql = strSql & " INNER JOIN mp_medium b"
    strSql = strSql & " ON a.mp_medium_id = b.mp_medium_id"
    strSql = strSql & " INNER JOIN mp_activity c"
    strSql = strSql & " ON b.mp_activity_id = c.mp_activity_id"
    strSql = strSql & " INNER JOIN mp_task d"
    strSql = strSql & " ON c.mp_task_id = d.mp_task_id"
    strSql = strSql & " AND d.mp_number = '" & strMPNumber & "'"
    strSql = strSql & " GROUP BY d.mp_task_id,"
    strSql = strSql & " d.task_desc,"
    strSql = strSql & " b.medium_code,"
    strSql = strSql & " c.brand_variant_code,"
    strSql = strSql & " c.brand_variant_name,"
    strSql = strSql & " a.month_number"
    strSql = strSql & " Union All"
    strSql = strSql & " SELECT"
    strSql = strSql & " d.mp_task_id,"
    strSql = strSql & " d.task_desc,"
    strSql = strSql & " b.medium_code,"
    strSql = strSql & " c.brand_variant_code,"
    strSql = strSql & " c.brand_variant_name,"
    strSql = strSql & " a.month_number,"
    strSql = strSql & " SUM(ISNULL(a.actual_nett_paid, 0) + ISNULL(a.actual_nett_bonus, 0)) actual_budget,0,'2-Actual' as title"
    strSql = strSql & " FROM mp_monthly_activity a"
    strSql = strSql & " INNER JOIN mp_medium b"
    strSql = strSql & " ON a.mp_medium_id = b.mp_medium_id"
    strSql = strSql & " INNER JOIN mp_activity c"
    strSql = strSql & " ON b.mp_activity_id = c.mp_activity_id"
    strSql = strSql & " INNER JOIN mp_task d"
    strSql = strSql & " ON c.mp_task_id = d.mp_task_id"
    strSql = strSql & " AND d.mp_number = '" & strMPNumber & "'"
    strSql = strSql & " GROUP BY d.mp_task_id,"
    strSql = strSql & " d.task_desc,"
    strSql = strSql & " b.medium_code,"
    strSql = strSql & " c.brand_variant_code,"
    strSql = strSql & " c.brand_variant_name,"
    strSql = strSql & " a.month_number"
    strSql = strSql & " Union All"
    strSql = strSql & " SELECT"
    strSql = strSql & " d.mp_task_id,"
    strSql = strSql & " d.task_desc,"
    strSql = strSql & " b.medium_code,"
    strSql = strSql & " c.brand_variant_code,"
    strSql = strSql & " c.brand_variant_name,"
    strSql = strSql & " a.month_number,"
    
    strSql = strSql & " SUM(COALESCE(Total_Actual, min_budget, actual_nett_paid + actual_nett_bonus) - (IsNull(a.actual_nett_paid, 0) + IsNull(a.actual_nett_bonus, 0))) As Balance"
    strSql = strSql & " ,0,'3-Balance' as title"
    strSql = strSql & " FROM mp_monthly_activity a"
    strSql = strSql & " INNER JOIN mp_medium b"
    strSql = strSql & " ON a.mp_medium_id = b.mp_medium_id"
    strSql = strSql & " INNER JOIN mp_activity c"
    strSql = strSql & " ON b.mp_activity_id = c.mp_activity_id"
    strSql = strSql & " INNER JOIN mp_task d"
    strSql = strSql & " ON c.mp_task_id = d.mp_task_id"
    strSql = strSql & " AND d.mp_number = '" & strMPNumber & "'"
    strSql = strSql & " GROUP BY d.mp_task_id,"
    strSql = strSql & " d.task_desc,"
    strSql = strSql & " b.medium_code,"
    strSql = strSql & " c.brand_variant_code,"
    strSql = strSql & " c.brand_variant_name,"
    strSql = strSql & " a.month_number) Z"
    '===============================================================Down List Begin
    rsSummary.Open strSql, ConnERP, 1, 3
    Call PivotPreview(Pivot1, rsSummary)
'    rsSummary.Close
'    Set rsSummary = Nothing
End Sub

Private Sub PopulateAddData(dbl_ReportID, str_Democode)
'*****************************************
'Procedure Name     : PopulateAddData
'Procedure Function : Populate Add Data
'Input Parameter    : -
'Output Parameter   : -
'Last Update Date   : Januari 2016
'Last Update By     : Ryo / Kreatif
'*********************************
    
    Dim rec_Spot As New ADODB.Recordset
    Dim int_TmpMarket As Integer
    
    'Open Rec
    strSql = "SELECT *  FROM Temp_Report_Spot WHERE Report_ID=" & dbl_ReportID
    
    rec_Spot.CursorLocation = adUseClient
    rec_Spot.Open strSql, connMacthIT, adOpenKeyset, adLockOptimistic
    
    'Loop
    Do While Not rec_Spot.EOF
        'Rating '(Demographic,Market,Channel,Date,StarTime,EndTime,Program,Product)
        If IsNull(rec_Spot.Fields("Market")) Then
            int_TmpMarket = 0
        Else
            int_TmpMarket = rec_Spot.Fields("Market")
        End If
                
        'Daypart
        rec_Spot.Fields("DayPart").Value = getDaypart(Format(rec_Spot.Fields("ATime"), "HH:MM:SS"))
        
        If IsNull(rec_Spot.Fields("TVR").Value) Then
            rec_Spot.Fields("TARP30").Value = GetTARP30(0, rec_Spot.Fields("Formatted_Duration").Value)
        Else
            rec_Spot.Fields("TARP30").Value = GetTARP30(rec_Spot.Fields("TVR").Value, rec_Spot.Fields("Formatted_Duration").Value)
        End If
        
        'PositionBreak
        If rec_Spot.Fields("PosInBreak").Value = 1 Then
            rec_Spot.Fields("PositionBreak").Value = "First"
        ElseIf rec_Spot.Fields("PosInBreak").Value = rec_Spot.Fields("SpotInBreak").Value Then
            rec_Spot.Fields("PositionBreak").Value = "Last"
        ElseIf rec_Spot.Fields("PosInBreak").Value = 2 Then
            rec_Spot.Fields("PositionBreak").Value = "Second"
        ElseIf rec_Spot.Fields("PosInBreak").Value = rec_Spot.Fields("SpotInBreak").Value - 1 Then
            rec_Spot.Fields("PositionBreak").Value = "Penultimate"
        Else
            rec_Spot.Fields("PositionBreak").Value = "Middle"
        End If
        
        rec_Spot.Update
        
        rec_Spot.MoveNext
        
        If Me.ProgressBar1.Value = 100 Then Me.ProgressBar1.Value = 1
        Me.ProgressBar1.Value = Me.ProgressBar1.Value + 1
            
    'End Loop
    Loop
    
    'Close Recordset
    CloseRecordset rec_Spot
 
End Sub

Private Sub generateData()
'*****************************************
'Procedure Name     : generateData
'Procedure Function : Populate Add Data
'Input Parameter    : -
'Output Parameter   : -
'Last Update Date   : Januari 2016
'Last Update By     : Ryo / Kreatif
'*********************************

    Me.MousePointer = 13
    ProgressBar1.Max = lvw_demo.ListItems.Count - 1
    ProgressBar1.Visible = True
    
    cbo_Select_Demo.Clear
    
    Dim int_IdxDemo As Integer
    Dim rec_Demo As New ADODB.Recordset
    
    For int_IdxDemo = 1 To lvw_demo.ListItems.Count
        If lvw_demo.ListItems(int_IdxDemo).Checked Then
            ProgressBar1.Value = int_IdxDemo
            dbl_ReportID = Get_Report_Id
            
            cbo_Select_Demo.AddItem lvw_demo.ListItems(int_IdxDemo)
            cbo_Select_Demo.ItemData(cbo_Select_Demo.NewIndex) = dbl_ReportID
        
            If rec_Demo.State = adStateOpen Then rec_Demo.Close
                rec_Demo.Open "SELECT * FROM Demographics WHERE Target_ID=" & lvw_demo.ListItems(int_IdxDemo).Tag & "", connMacthIT, adOpenKeyset, adLockOptimistic
            If Not rec_Demo.EOF Then
                str_Democode = Trim(rec_Demo.Fields("Target_Code").Value)
            End If
            CloseRecordset rec_Demo
        
        
            'Get Spot (ReportID)
            Call getSpot
            Call PopulateAddData(dbl_ReportID, str_Democode)

        End If
    Next int_IdxDemo
    
    Me.ProgressBar1.Visible = False
    Me.MousePointer = 0
    
    CloseRecordset rec_Demo
    
    MsgBox "Done !", vbExclamation, strApplication_Name
        
    cbo_Select_Demo.ListIndex = 0

End Sub

Private Sub PreviewToPivot()
'*****************************************
'Procedure Name     : PreviewToPivot
'Procedure Function : Load data to pivot table.
'Input Parameter    : -
'Output Parameter   : -
'Date               : 05-May-2015/{73 64 6B}
'*****************************************
    Dim str_TableAlias() As String, str_FieldName() As String
    Dim str_Criteria As String

    On Local Error GoTo TrapErrorHere
    Erase str_TableAlias: Erase str_FieldName
    str_Criteria = ""

    'Validate.
    If Not IsValidData Then GoTo ClearVarsThenExit

    ProgressBar1.Value = ProgressBar1.Value + 1
    
    strSql = "SELECT Match_Spot_ID,Job_ID + '-' + Job_Number AS Job_Number,"
    strSql = strSql & " Channel,ADate,ATime,ADay,PosInBreak,SpotInbreak,PositionBreak,"
    strSql = strSql & " ASector,AProduct,ACopy,AAdvertiser,AProgram_Name,Genre,Formatted_Duration,DayPart,"
    strSql = strSql & " AMonth, AYear,AStart_Time,AEnd_Time,BreakNumberInProgram,BreakNumberInProgram2130,"
    strSql = strSql & " Spot_Type_Catalog.Spot_Type_Catalog AS SpotType,  "
    strSql = strSql & " SUM(TVR) AS TVR,SUM(TARP30) TARP30,SUM(Exp_Gross) AS Exp_Gross,sum(Exp_Nett) as Exp_Nett"

    strSql = strSql & " ,SUM(isnull(EXP30,0) * 1000) AS ADEX_Arianna"

    strSql = strSql & " FROM Temp_Report_Spot "
    strSql = strSql & " LEFT JOIN Spot_Type_Catalog ON Spot_Type_Catalog.Code=Temp_Report_Spot.Spot_Type"
    strSql = strSql & " WHERE Report_ID=" & cbo_Select_Demo.ItemData(cbo_Select_Demo.ListIndex)

    strSql = strSql & " GROUP BY "
    strSql = strSql & " Match_Spot_ID,Job_ID,Job_Number,"
    strSql = strSql & " Channel,ADate,ATime,ADay,PosInBreak,SpotInbreak,PositionBreak,"
    strSql = strSql & " ASector,AProduct,ACopy,AAdvertiser,AProgram_Name,Genre,Formatted_Duration,DayPart,"
    strSql = strSql & " AMonth, AYear,AStart_Time,AEnd_Time,BreakNumberInProgram,BreakNumberInProgram2130,"
    strSql = strSql & " Spot_Type_Catalog.Spot_Type_Catalog"
    
    ProgressBar1.Value = 5
    
    Set rec_Report = New ADODB.Recordset: rec_Report.CursorLocation = adUseClient
    
    rec_Report.Open strSql, conTemp, adOpenForwardOnly, adLockReadOnly, adCmdText: strSql = ""
    
    ProgressBar1.Value = ProgressBar1.Value + 1
    
    If rec_Report.RecordCount < 1 Then
        MsgBox "Data Is Empty", vbExclamation, strApplication_Name
        Pivot1.DataSource = Nothing
        Pivot1.Refresh
        ProgressBar1.Visible = False
        Me.MousePointer = 0
        Exit Sub
    End If
    

    '== Checking Existing User Template From Database
    Dim rec_Template As New ADODB.Recordset
    
    strSql = ""
    strSql = strSql & "SELECT * "
    strSql = strSql & "FROM TEMP_PIVOT "
    strSql = strSql & "WHERE ltrim(rtrim(username))='" & strUserName & "' "
    strSql = strSql & "AND report_code='RTUC' ORDER BY STAMP DESC "
    
    rec_Template.Open strSql, conTemp, adOpenStatic, adLockReadOnly
     
    ProgressBar1.Value = ProgressBar1.Value + 1
    '== Pivot layout
    Call PivotPreview(Pivot1, rec_Report)
    ProgressBar1.Value = ProgressBar1.Value + 1
    '== Switch To Template
    If rec_Template.RecordCount < 1 Then
        '-- Default Template
        picButton_Click (1)
    Else
        rec_Template.MoveFirst
        '--- Load Last Update Template
        Pivot1.Layout = rec_Template!temp_contain
        strRefTemplate = Trim(rec_Template!FileName)
        CekDir "CRPT"
        createTemplateFile rec_Template
    End If
  
    Call CloseRecordset(rec_Template)
    ProgressBar1.Visible = False
    Me.MousePointer = 0

    strDateFormat = "": strTimeFormat = ""
    On Local Error GoTo 0

    GoTo ClearVarsThenExit

TrapErrorHere:
    MsgBox Err.Description, vbExclamation, strApplication_Name
    Err.Clear

ClearVarsThenExit:
    Erase str_TableAlias: Erase str_FieldName
    str_Criteria = ""

End Sub

Public Sub create_table_DayPart()
'*****************************************
'Procedure Name     : create_table_DayPart
'Procedure Function : create_table_DayPart
'Input Parameter    : -
'Output Parameter   : -
'Last Update Date   : Januari 2016
'Last Update By     : Ryo / Kreatif
'*********************************

    Set rst_empDayPart = Nothing
    Set rst_empDayPart = New ADODB.Recordset
    
    With rst_empDayPart.Fields
        .Append "Brand_code", adVarChar, 4, adFldIsNullable
        .Append "Daypart_Start", adChar, 4, adFldIsNullable
        .Append "Daypart_End", adChar, 4, adFldIsNullable
        .Append "Order_No", adInteger, adFldIsNullable
    End With
    rst_empDayPart.Open
    
End Sub

Private Sub picButton_Click(Index As Integer)
'************************************************
' Procedure         : picButton_Click
' Function          : Action utk Navigation dan CRUD.
'Last Update Date   : Januari 2016
'Last Update By     : Ryo / Kreatif
'*********************************
    
    Dim str_FileUserTemplate As String, str_Values As String
    Dim arr_Values() As String, rec_RptCriteria() As String
    Dim strTemplate As String
    str_FileUserTemplate = "": str_Values = ""

    Select Case Index
    
        Case biePreview  'PREVIEW.
            
            'frm_ConProgress.Show
            'Call SleepX
            'generateData
            Set_Default_Pivot_Template
            'Unload frm_ConProgress
        
        Case bieDefaultTemplate 'DEFAULT_TEMPLATE.
            'Set the layout of pivot to default reset.
            Set_Reset_Pivot_Template
        Case bieUserTemplate 'USER_TEMPLATE.
            'Set the user template.
             str_FileUserTemplate = "Rpt_" & "Corporate_Report" & "_User_Template.txt"

            'Load the template.
            If MsgBox("Do you want to load the user template ?", vbYesNo + vbQuestion, strApplication_Name) = vbYes Then
                Call PivotUserTemplate(Pivot1, CommonDialog1, str_FileUserTemplate, True, "SUMP")
            Else: Call PivotUserTemplate(Pivot1, CommonDialog1, str_FileUserTemplate, False, "SUMP")
            End If

        Case bieSaveTemplate  'SAVE_TEMPLATE.
            'Set the default file template.
            str_FileUserTemplate = "Rpt_" & "Corporate_Report" & "_User_Template.txt"

            'Load the template.
            If MsgBox("Do you want to save this Layout as user template ?", vbYesNo + vbQuestion, strApplication_Name) = vbYes Then
                Call PivotSaveTemplate(Pivot1, strRefTemplate, CommonDialog1, str_FileUserTemplate, True, "SUMP")
            
            End If
            
            str_FileUserTemplate = Pivot1.Export(App.Path & "\templates\sample.csv", "all||;")

        Case biePrint 'PRINT.
            ReDim rec_RptCriteria(3, 1)
            
            'Company
            rec_RptCriteria(0, 0) = ""
            rec_RptCriteria(0, 1) = strCompany_Name
            
            'Blank
            rec_RptCriteria(2, 0) = ""
            rec_RptCriteria(2, 1) = ""
            
            ' Employee
            rec_RptCriteria(3, 0) = "Employee: "
            
            
'            Call PivotPrint(Pivot1, Print1, Me.Caption, rec_RptCriteria, False, CommonDialog1)

        Case bieExportExcel  'EXPORT_EXCEL.
            
            Dim str_hdr As String
'            Call Me.PopupMenu(mdi_Main.mnuPopup())
'
'            ' POPUP_MENU.
'            If pubPopupCmd = POPUP_EXCEL Then
'                ReDim arr_Values(2)
'                arr_Values(0) = strCompany_Name_Name
'                arr_Values(1) = UCase$(Trim$(Me.Caption))
'                arr_Values(2) = Format$(Dt_Period_Start.Value, "DD-MMM-YYYY") & " to " & Format$(Dt_Period_Finish.Value, "DD-MMM-YYYY")
'
'
'                ReDim rec_RptCriteria(0, 1)
'                rec_RptCriteria(0, 0) = "Client"
'                rec_RptCriteria(0, 1) = getListClient & "     Brand : " & getListBrand
'
'                Call PivotExportExcel(Pivot1, arr_Values, rec_RptCriteria)
'
'            ElseIf pubPopupCmd = POPUP_PDF Then
'                ReDim rec_RptCriteria(3, 1)
'
'                'Company
'                rec_RptCriteria(0, 0) = ""
'                rec_RptCriteria(0, 1) = StrCompany_Name
'
'                'Period
'                rec_RptCriteria(1, 0) = "Client: "
'                rec_RptCriteria(1, 1) = getListClient
'
'                'Blank
'                rec_RptCriteria(2, 0) = "MonthYear"
'                rec_RptCriteria(2, 1) = getListDemo
'
'                Call PivotPrint(Pivot1, Print1, str_hdr, rec_RptCriteria, True, CommonDialog1)
'
'            ElseIf pubPopupCmd = POPUP_CSV Then
'                Call PivotExportCSV(Pivot1, CommonDialog1)
'            End If
'            pubPopupCmd = ""
            
        Case bieExit 'EXIT.
            Unload Me
    End Select

    str_FileUserTemplate = "": str_Values = ""
    Erase arr_Values: Erase rec_RptCriteria
    
End Sub

Private Sub picButton_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single) 'TOOLBAR_AI.
 'TOOLBAR_AI.
'************************************************
' Procedure         : picButton_MouseDown
' Function          : TOOLBAR_AI saat mouse ditekan.
' Created By        : {73 64 6B}
' Date              : 05-May-2015
'************************************************

        picButton(Index).Picture = LoadPicture(SetButtonImageEffect(Index, bieDown)) 'PREVIEW.

End Sub

Private Sub picButton_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single) 'TOOLBAR_AI.
'************************************************
' Procedure         : picButton_MouseMove
' Function          : TOOLBAR_AI saat mouse berada di area button.
' Created By        : Rudi / Kreatif
' Date              : 12-Apr-2015
' Last Update Date  : Januari 2016
' Last Update By    : Tedi / Kreatif
'************************************************
    
    picButton_Obj Index, Button, Shift, X, Y, picButton

End Sub


Private Function IsValidData() As Boolean
'*****************************************
'Procedure Name     : IsValidData
'Procedure Function : To cek if data will be entered to table is valid or invalid.
'Input Parameter    : ---
'Output Parameter   : True=Valid, False=Invalid.
'Date               : 05-May-2015
'LastUpdate/By      : {73 64 6B}
'*****************************************
        
    If getListChannel = "''" And getListClient = "''" And getListBrand = "''" Then
        MsgBox "Please select Channel AND/OR Client AND/OR Brand to prosess !", vbExclamation, strApplication_Name
        IsValidData = False
        Exit Function
    End If
    
    IsValidData = True

End Function

Private Sub SetButtonToolbar(ByVal bln_paIsNormalMode As Boolean) 'TOOLBAR_AI.
'************************************************
' Procedure         : SetButtonToolbar
' Function          : TOOLBAR_AI.
' Created By        : Rudi / Kreatif
' Last Update Date  : Januari 2016
' Last Update By    : Rudi / Kreatif
'************************************************

    Dim str_Dummy As String
    
    picToolbar.BackColor = vbButtonFace

    With picButton(biePreview) 'PREVIEW.
        .Enabled = bln_paIsNormalMode
        
        If bln_paIsNormalMode Then
            .Picture = LoadPicture(SetButtonImageEffect(biePreview, bieNormal))
        Else: .Picture = LoadPicture(SetButtonImageEffect(biePreview, bieDisabled))
        End If
        
    End With

    With picButton(bieDefaultTemplate) 'DEFAULT_TEMPLATE.
        .Enabled = bln_paIsNormalMode
        
        If bln_paIsNormalMode Then
            .Picture = LoadPicture(SetButtonImageEffect(bieDefaultTemplate, bieNormal))
        Else: .Picture = LoadPicture(SetButtonImageEffect(bieDefaultTemplate, bieDisabled))
        End If
        
    End With

    With picButton(bieUserTemplate) 'USER_TEMPLATE.
        .Enabled = bln_paIsNormalMode
        
        If bln_paIsNormalMode Then
            .Picture = LoadPicture(SetButtonImageEffect(bieUserTemplate, bieNormal))
        Else: .Picture = LoadPicture(SetButtonImageEffect(bieUserTemplate, bieDisabled))
        End If
        
    End With

    With picButton(bieSaveTemplate) 'SAVE_TEMPLATE.
        .Enabled = bln_paIsNormalMode
        
        If bln_paIsNormalMode Then
            .Picture = LoadPicture(SetButtonImageEffect(bieSaveTemplate, bieNormal))
        Else: .Picture = LoadPicture(SetButtonImageEffect(bieSaveTemplate, bieDisabled))
        End If
    
    End With

    With picButton(biePrint) 'PRINT.
        .Enabled = bln_paIsNormalMode
        .Visible = bln_paIsNormalMode
        
        If bln_paIsNormalMode Then
            .Picture = LoadPicture(SetButtonImageEffect(biePrint, bieNormal))
        Else: .Picture = LoadPicture(SetButtonImageEffect(biePrint, bieDisabled))
        End If
    
    End With

    With picButton(bieExportExcel) 'EXPORT_EXCEL.
        .Enabled = bln_paIsNormalMode
        .Visible = bln_paIsNormalMode
        If bln_paIsNormalMode Then
            .Picture = LoadPicture(SetButtonImageEffect(bieExportExcel, bieNormal))
        Else: .Picture = LoadPicture(SetButtonImageEffect(bieExportExcel, bieDisabled))
        End If
    
    End With
    
    With picButton(bieExit) 'EXIT.
        .Enabled = bln_paIsNormalMode
        .Visible = bln_paIsNormalMode
        
        If bln_paIsNormalMode Then
            .Picture = LoadPicture(SetButtonImageEffect(bieExit, bieNormal))
        Else: .Picture = LoadPicture(SetButtonImageEffect(bieExit, bieDisabled))
        End If
    
    End With
    
    str_Dummy = GetSecureValue(strLogin_User, "Budget Summery")
    
    'Preview
    If Mid(str_Dummy, 2, 1) = "0" Or Trim(Mid(str_Dummy, 2, 1)) = "" Then
        picButton(biePreview).Enabled = False
        picButton(biePreview).Picture = LoadPicture(SetButtonImageEffect(0, bieDisabled))
    End If
    
    'Reset
    If Mid(str_Dummy, 3, 1) = "0" Or Trim(Mid(str_Dummy, 3, 1)) = "" Then
        picButton(bieDefaultTemplate).Enabled = False
        picButton(bieDefaultTemplate).Picture = LoadPicture(SetButtonImageEffect(1, bieDisabled))
    End If
    
    'Load
    If Mid(str_Dummy, 4, 1) = "0" Or Trim(Mid(str_Dummy, 4, 1)) = "" Then
        picButton(bieUserTemplate).Enabled = False
        picButton(bieUserTemplate).Picture = LoadPicture(SetButtonImageEffect(2, bieDisabled))
    End If
    
    'save
    If Mid(str_Dummy, 5, 1) = "0" Or Trim(Mid(str_Dummy, 5, 1)) = "" Then
        picButton(bieSaveTemplate).Enabled = False
        picButton(bieSaveTemplate).Picture = LoadPicture(SetButtonImageEffect(3, bieDisabled))
    End If
    
    'Print
    If Mid(str_Dummy, 6, 1) = "0" Or Trim(Mid(str_Dummy, 6, 1)) = "" Then
        picButton(biePrint).Enabled = False
        picButton(biePrint).Picture = LoadPicture(SetButtonImageEffect(4, bieDisabled))
    End If
    
    'Export
    If Mid(str_Dummy, 7, 1) = "0" Or Trim(Mid(str_Dummy, 7, 1)) = "" Then
        picButton(bieExportExcel).Enabled = False
        picButton(bieExportExcel).Picture = LoadPicture(SetButtonImageEffect(5, bieDisabled))
    End If

End Sub

Private Sub Pivot1_LayoutEndChanging(ByVal Operation As EXPIVOTLibCtl.LayoutChangingEnum)
'*****************************************
'Procedure Name     : Pivot1_LayoutEndChanging
'Procedure Function :
'Input Parameter    : -
'Output Parameter   : -
'Last Update Date   : Januari 2016
'Last Update By     : Ryo / Kreatif
'*********************************
'On Error Resume Next
   If Pivot1.PivotRows = "" Then
   Pivot1.PivotRows = "1,3,4"
   End If
     If Pivot1.PivotColumns = "" Then
   Pivot1.PivotColumns = "sum(7)/5,sum(9)/5,sum(10)/5"
   End If
    Dim str() As String
    If Pivot1.PivotRows <> "" Then
    str = Split(Pivot1.PivotRows, ",")
    Dim clm As String
    clm = "sum[bold,underline,content=currency]"
    Dim brs As String
    brs = Replace(str(0), "[", "")
    brs = Replace(brs, "]", "")
    clm = clm & ",sum(" & brs & ")[bold,underline,content=numeric]"
    Pivot1.PivotTotals = clm
    End If
    If Pivot1.DataColumns.Count > 0 Then
    Dim intCount As Integer

    For intCount = 0 To Pivot1.DataColumns.Count - 1
        Pivot1.BeginUpdate
        Pivot1.DataColumns(intCount).Alignment = RightAlignment
        Pivot1.EndUpdate
    Next
    End If
  '   Pivot1.PivotColumns = "sum(7)/6,sum(10)/6"
  '  x = Pivot1.PivotColumns

   ' Pivot1.PivotColumns = "sum(7)/5,sum(9)/5,sum(10)/5"
  ' Pivot1.ShowViewCompact = exViewCompact
   ' x = Pivot1.Layout
    
    'Pivot1.EndUpdate
    
End Sub

Private Sub pnl_Hiding_Click()
'*****************************************
'Procedure Name     : pnl_Hiding_Click
'Procedure Function : Hiding panel Filter
'Input Parameter    : -
'Output Parameter   : -
'Last Update Date   : Januari 2016
'Last Update By     : Ryo / Kreatif
'*********************************
 If pnl_Filter.Width = 4850 Then
        pnl_Hiding.Left = 0
        pnl_Filter.Width = 0
    Else
        pnl_Filter.Width = 4850
        pnl_Hiding.Left = 4910
    End If
    Call AdjustSizeForm
End Sub




Private Sub Pivot1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If blnMoveGone = True Then Exit Sub
    blnMoveGone = True
    If Pivot1.DataColumns.Count > 0 Then
    Dim intCount As Integer

    For intCount = 0 To Pivot1.DataColumns.Count - 1
        Pivot1.BeginUpdate
        Pivot1.DataColumns(intCount).Alignment = RightAlignment
        Pivot1.EndUpdate
    Next
    End If
End Sub

