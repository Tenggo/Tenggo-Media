VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFlxGrd.ocx"
Begin VB.Form frm_MPTVLayering 
   BorderStyle     =   0  'None
   Caption         =   "TV Layering"
   ClientHeight    =   7170
   ClientLeft      =   -15
   ClientTop       =   555
   ClientWidth     =   12495
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7170
   ScaleWidth      =   12495
   ShowInTaskbar   =   0   'False
   Begin Threed.SSPanel pnl_Main 
      Align           =   1  'Align Top
      Height          =   5580
      Left            =   0
      TabIndex        =   2
      Top             =   750
      Width           =   12495
      _Version        =   65536
      _ExtentX        =   22040
      _ExtentY        =   9842
      _StockProps     =   15
      Caption         =   " "
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
      Begin MSFlexGridLib.MSFlexGrid FGTVLayering 
         Height          =   6825
         Left            =   60
         TabIndex        =   3
         Top             =   60
         Width           =   12210
         _ExtentX        =   21537
         _ExtentY        =   12039
         _Version        =   393216
         Rows            =   5
         Cols            =   4
         FixedRows       =   4
         FixedCols       =   2
         BackColorFixed  =   14737632
         ForeColorFixed  =   -2147483640
         WordWrap        =   -1  'True
         MergeCells      =   1
         AllowUserResizing=   3
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
   Begin VB.PictureBox picToolbar 
      Align           =   1  'Align Top
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ClipControls    =   0   'False
      ForeColor       =   &H80000008&
      Height          =   750
      Left            =   0
      ScaleHeight     =   750
      ScaleWidth      =   12495
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   0
      Width           =   12495
      Begin VB.PictureBox picButton 
         Appearance      =   0  'Flat
         BackColor       =   &H00F0F0F0&
         BorderStyle     =   0  'None
         ClipControls    =   0   'False
         ForeColor       =   &H80000008&
         Height          =   750
         Index           =   23
         Left            =   -15
         Picture         =   "frm_MPTVLayering.frx":0000
         ScaleHeight     =   750
         ScaleWidth      =   1500
         TabIndex        =   1
         TabStop         =   0   'False
         Top             =   0
         Width           =   1500
      End
   End
End
Attribute VB_Name = "frm_MPTVLayering"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim strMPNumber As String
Dim intPlanYear As String
Dim intweekcount As Integer
Dim MarginKanan As Single
Dim MarginBawah As Single
'
Private Sub FGTVLayering_DblClick()
    MsgBox FGTVLayering.Text, vbExclamation, strApplication_Name
End Sub

Private Sub FGTVLayering_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 2 Then
        With FGTVLayering
            .col = .MouseCol
            .Row = .MouseRow
        End With
        'MsgBox "PopupMenu mnu_Popup"
    End If
End Sub

Private Sub Form_Activate()
    On Error GoTo errHandle1
    If frm_MPInsertion.ActiveControl.Name = "cmdExportToExcel" Then
        Me.Hide
        frm_MPInsertion.Refresh
    End If
    Exit Sub
    
err1:
    On Error GoTo errHandle2
    If frm_MediaPlan_View.ActiveControl.Name = "cmdExportToExcel" Then
        Me.Hide
        frm_MediaPlan_View.Refresh
    End If
    Exit Sub
    
errHandle1:
    Err.Clear
    Resume err1
errHandle2:
    Err.Clear
End Sub

Sub Form_Load()
    
    Dim rsMPTask As New ADODB.Recordset
    Dim rsMPActivity As New ADODB.Recordset
    Dim rsMPMediumDetail As New ADODB.Recordset
    Dim rsMPPlanDimension As New ADODB.Recordset
    Dim rsMPInsertion As New ADODB.Recordset
    Dim rsMPTVReachFreq As New ADODB.Recordset
    Dim i As Integer, j As Integer
    Dim intTaskCount As Integer
    Dim intActivityRow As Integer
    Dim intTotalSpot As Double
    Dim TotalPerTask() As Double
    Dim TOTAL As Double
    
    Dim rf_seq As Integer
    
    MarginKanan = Me.Width - (FGTVLayering.Left + FGTVLayering.Width)
    MarginBawah = Me.Height - (FGTVLayering.Top + FGTVLayering.Height)
    
    'MsgBox "mnu_Popup.Visible = False"
    strMPNumber = frm_MPInsertion.cboMPNumber.Text
    If strMPNumber = "" Then strMPNumber = frm_MediaPlan_View.cboMPNumber.Text
    intPlanYear = Mid(strMPNumber, 6, 4)
    resize Me
    intweekcount = 0
    Call initGrid(intPlanYear)
    
    With FGTVLayering
        rsMPTask.Open "select count(*) from mp_task where mp_number = '" & strMPNumber & "'", ConnERP, 1, 3
            ReDim TotalPerTask(rsMPTask(0), intweekcount + 1)
        rsMPTask.Close
        
        rsMPTask.Open "select mp_task_id from mp_task where mp_number = '" & strMPNumber & "'", ConnERP, 1, 3
        intTaskCount = 0
        While Not rsMPTask.EOF
            .Rows = .Rows + 1
            intTaskCount = intTaskCount + 1
            
            .Row = .Rows - 1
            .col = 0
            .Text = "Task " & intTaskCount
            .CellAlignment = 1
            
            rsMPActivity.Open "select * from mp_activity where mp_task_id = '" & rsMPTask("mp_task_id") & "'", ConnERP, 1, 3
            While Not rsMPActivity.EOF
                .Rows = .Rows + 1
                intActivityRow = .Rows - 1
                
                .Row = .Rows - 1
                .col = 0
                .Text = rsMPActivity("activity_desc")
                .CellAlignment = 1
                
                '.Rows = .Rows + 1
                
                '.Row = .Rows - 1
                '.Col = 0
                '.Text = "(" & rsMPActivity("brand_target") & ")"
                '.CellAlignment = 1
                
                rsMPMediumDetail.Open "select * from mp_medium_detail where mp_medium_id in (select mp_medium_id from mp_medium where mp_activity_id = '" & rsMPActivity("mp_activity_id") & "' and medium_code = 'TV') ", ConnERP, 1, 3
                While Not rsMPMediumDetail.EOF
                    rsMPPlanDimension.Open "select * from mp_plan_dimension where mp_medium_detail_id = '" & rsMPMediumDetail("mp_medium_detail_id") & "' ", ConnERP, 1, 3
                    While Not rsMPPlanDimension.EOF
                        
                        .Row = .Rows - 1
                        
                        .col = 1
                        .Text = rsMPPlanDimension("version") & String(.Row, " ")
                        .CellAlignment = 1
                        
                        .col = 3
                        .Text = rsMPMediumDetail("station_name") & String(.Row, " ")
                        .CellAlignment = 1
                        
                        .col = 2
                        .Text = rsMPPlanDimension("duration") & " sec" & String(.Row, " ")
                        .CellAlignment = 1
                        
                        'Print Insertion
                        intTotalSpot = 0
                        rsMPInsertion.Open "select * from mp_insertion where mp_plan_dim_id='" & rsMPPlanDimension("mp_plan_dim_id") & "'", ConnERP, 1, 3
                        While Not rsMPInsertion.EOF
                            .TextMatrix(.Row, rsMPInsertion("week_year") + 3) = FormatNumber(rsMPInsertion("spot"), 0)
                            intTotalSpot = intTotalSpot + rsMPInsertion("Spot")
                            TotalPerTask(intTaskCount - 1, rsMPInsertion("week_year") - 1) = TotalPerTask(intTaskCount - 1, rsMPInsertion("week_year") - 1) + rsMPInsertion("spot")
                            TotalPerTask(intTaskCount - 1, intweekcount) = TotalPerTask(intTaskCount - 1, intweekcount) + rsMPInsertion("spot")
                            rsMPInsertion.MoveNext
                        Wend
                        rsMPInsertion.Close
                        .TextMatrix(.Row, .cols - 1) = FormatNumber(intTotalSpot, 0)
                        
                        .Rows = .Rows + 2
                        
                        .Row = .Rows - 2
                        .Text = "Freq"
                        .CellAlignment = 1
                        
                        .Row = .Rows - 1
                        .Text = "Reach"
                        .CellAlignment = 1
                        
                        'Print Reach & Freq
                        rsMPTVReachFreq.Open "select * from mp_tv_reach_frequency where mp_plan_dim_id = '" & rsMPPlanDimension("mp_plan_dim_id") & "'", ConnERP, 1, 3
                        rf_seq = 1
                        While Not rsMPTVReachFreq.EOF
                            'Print TV RF PAKE MERGING (SAMA DGN DI MPINSERTION)
                            For i = rsMPTVReachFreq("week_year_start") To rsMPTVReachFreq("week_year_end")
                                
                                .col = i + 3
                                .Row = .Rows - 2
                                .Text = String(rf_seq, " ") & Trim(rsMPTVReachFreq("Frequency_name")) & String(rf_seq, " ")
                                .CellBackColor = vbYellow
                                
                                .Row = .Rows - 1
                                .Text = String(rf_seq, " ") & Trim(rsMPTVReachFreq("reach")) & String(rf_seq, " ")
                                .CellBackColor = vbYellow
                                
                            Next
                            .MergeRow(.Rows - 2) = True
                            .MergeRow(.Rows - 1) = True
                            
                            
                            rf_seq = rf_seq + 1
                            rsMPTVReachFreq.MoveNext
                        Wend
                        rsMPTVReachFreq.Close
                        rsMPPlanDimension.MoveNext
                        .Rows = .Rows + 1
                        
                    Wend
                    rsMPPlanDimension.Close
                    rsMPMediumDetail.MoveNext
                    If Not rsMPMediumDetail.EOF Then
                        .Rows = .Rows + 1
                    End If
                Wend
                rsMPMediumDetail.Close
                
                'Print Brand Target dibawah activity desc
                If intActivityRow = .Rows - 1 Then .Rows = .Rows + 1
                .Row = intActivityRow + 1
                .col = 0
                .Text = "(" & rsMPActivity("Brand_target") & ")"
                .CellAlignment = 1
                
                rsMPActivity.MoveNext
                
            Wend
            rsMPActivity.Close
            rsMPTask.MoveNext
        Wend
        rsMPTask.Close
        
        If intTaskCount > 0 Then
            
            'Print Total PerTask
            For i = 1 To intTaskCount
                .Rows = .Rows + 1
                .Row = .Rows - 1
                
                .col = 0
                .Text = "TASK " & i
                .CellAlignment = 1
                
                For j = 1 To intweekcount + 1
                    If TotalPerTask(i - 1, j - 1) <> 0 Then
                        .TextMatrix(.Row, j + 3) = FormatNumber(TotalPerTask(i - 1, j - 1), 0)
                    Else
                        .TextMatrix(.Row, j + 3) = "-"
                    End If
                Next
            Next
            
            'Print Grand Total
            .Rows = .Rows + 1
            .Row = .Rows - 1
            
            .col = 0
            .Text = "TOTAL"
            .CellAlignment = 1
            
            For j = 1 To intweekcount + 1
                TOTAL = 0
                For i = 1 To intTaskCount
                    TOTAL = TOTAL + TotalPerTask(i - 1, j - 1)
                Next
                If TOTAL <> 0 Then
                    .TextMatrix(.Row, j + 3) = FormatNumber(TOTAL, 0)
                Else
                    .TextMatrix(.Row, j + 3) = "-"
                End If
            Next
            
        End If
    End With
    

End Sub

Private Sub initGrid(tahun As String)
'*****************************************************************************
' Nama Submodul         :  LoadWeekCommencing
' Fungsi Submodul       :  Load grid WC untuk insertion
' Nama Programmer       :  Sistyo
' Tgl Pembuatan         :  29 Juni 2005
' Last Update           :  29 Juni 2005/Sistyo
'******************************************************************************
    Dim counter As Integer
    Dim rsTemp As New ADODB.Recordset
    With FGTVLayering

        .cols = 4
        .Rows = 3
        
        .MergeCol(0) = True
        .ColAlignment(0) = 3
        .ColWidth(0) = 2500
        .TextMatrix(0, 0) = "PROJECT" & vbCrLf & "& TARGET"
        .TextMatrix(1, 0) = "PROJECT" & vbCrLf & "& TARGET"
        .TextMatrix(2, 0) = "PROJECT" & vbCrLf & "& TARGET"
        
        .MergeCol(1) = True
        .ColWidth(1) = 2500
        .ColAlignment(1) = 3
        .TextMatrix(0, 1) = "TITLE"
        .TextMatrix(1, 1) = "TITLE"
        .TextMatrix(2, 1) = "TITLE"
        
        .MergeCol(2) = True
        .ColWidth(2) = 600
        .ColAlignment(2) = 3
        .TextMatrix(0, 2) = "DURA" & vbCrLf & "TION"
        .TextMatrix(1, 2) = "DURA" & vbCrLf & "TION"
        .TextMatrix(2, 2) = "DURA" & vbCrLf & "TION"
        
        .MergeCol(3) = True
        .ColWidth(3) = 2500
        .ColAlignment(3) = 3
        .TextMatrix(0, 3) = "STATIONS"
        .TextMatrix(1, 3) = "STATIONS"
        .TextMatrix(2, 3) = "STATIONS"
        
        .MergeRow(0) = True
        .MergeRow(1) = True
        .MergeRow(2) = True
        
        
        rsTemp.Open "select * from week_commencing where [year] = '" & tahun & "' order by cast([month] as int) ", ConnERP, 1, 3
        While Not rsTemp.EOF
            For counter = 1 To rsTemp(7) 'Week Count
                .cols = .cols + 1
                .ColWidth(.cols - 1) = 500
                .ColAlignment(.cols - 1) = 4
                .TextMatrix(0, .cols - 1) = EngMonthName(rsTemp(1)) 'Month
                .TextMatrix(1, .cols - 1) = Day(rsTemp(counter + 1))  'date
                .TextMatrix(2, .cols - 1) = .cols - 4 'Week
            Next
            intweekcount = intweekcount + rsTemp(7)
            rsTemp.MoveNext
        Wend
        
        rsTemp.Close
        .cols = .cols + 1
        .MergeCol(.cols - 1) = True
        .ColAlignment(.cols - 1) = 3
        .ColWidth(.cols - 1) = 2000
        .TextMatrix(0, .cols - 1) = "Total"
        .TextMatrix(1, .cols - 1) = "Total"
        .TextMatrix(2, .cols - 1) = "Total"
        
        .FixedCols = 1
        
    End With
End Sub

Private Sub Form_Resize()
'<CSCM>
'********************************************************************************
'Procedure Name     : Form_Resize
'Procedure Function : ---
'Input Parameter    : ---
'Output Parameter   : ---
'Date               : 3/29/2016
'LastUpdate/By      : Tedi / Kreatif
'Name Before        : -
'********************************************************************************
'</CSCM>

    Me.Height = mdi_Main.Height
    Me.Width = mdi_Main.Width - mdi_Main.picSideBar.Width - 500
    Me.Top = 0
    Me.Left = 0
    picToolbar.Height = 750
    pnl_Main.Height = Me.ScaleHeight - picToolbar.Height
    FGTVLayering.Height = pnl_Main.Height - (FGTVLayering.Top)
    FGTVLayering.Width = pnl_Main.Width - (FGTVLayering.Left * 2)
   
End Sub

Private Sub mnu_freeze_Click()
    With FGTVLayering
        If .col < .FixedCols Then
            .FixedCols = .col
        Else
            .FixedCols = .col + 1
        End If
    End With
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
 
    Unload Me


End Sub

