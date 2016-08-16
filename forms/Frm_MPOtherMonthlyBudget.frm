VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFlxGrd.ocx"
Begin VB.Form Frm_MPOtherMonthlyBudget 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Other Monthly Budget"
   ClientHeight    =   3630
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9690
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3630
   ScaleWidth      =   9690
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.PictureBox picToolbar 
      Align           =   1  'Align Top
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ClipControls    =   0   'False
      ForeColor       =   &H80000008&
      Height          =   750
      Left            =   0
      ScaleHeight     =   750
      ScaleWidth      =   9690
      TabIndex        =   9
      TabStop         =   0   'False
      Top             =   0
      Width           =   9690
      Begin VB.PictureBox picButton 
         Appearance      =   0  'Flat
         BackColor       =   &H00F0F0F0&
         BorderStyle     =   0  'None
         ClipControls    =   0   'False
         ForeColor       =   &H80000008&
         Height          =   750
         Index           =   23
         Left            =   1620
         Picture         =   "Frm_MPOtherMonthlyBudget.frx":0000
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
         Left            =   90
         Picture         =   "Frm_MPOtherMonthlyBudget.frx":1D07
         ScaleHeight     =   750
         ScaleWidth      =   1500
         TabIndex        =   10
         TabStop         =   0   'False
         Top             =   0
         Width           =   1500
      End
   End
   Begin Threed.SSPanel pnl_Main 
      Align           =   1  'Align Top
      Height          =   2835
      Left            =   0
      TabIndex        =   0
      Top             =   750
      Width           =   9690
      _Version        =   65536
      _ExtentX        =   17092
      _ExtentY        =   5001
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
      Begin VB.Frame FrameOTBudget 
         Caption         =   "Monthly Budget"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2055
         Left            =   165
         TabIndex        =   3
         Top             =   45
         Width           =   9360
         Begin VB.TextBox txt_Temp 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0FFFF&
            BorderStyle     =   0  'None
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
            Left            =   4800
            TabIndex        =   4
            Text            =   "Text1"
            Top             =   690
            Visible         =   0   'False
            Width           =   1215
         End
         Begin MSFlexGridLib.MSFlexGrid FG_OT_Monthly_Budget 
            Height          =   1410
            Left            =   225
            TabIndex        =   5
            Top             =   375
            Width           =   8895
            _ExtentX        =   15690
            _ExtentY        =   2487
            _Version        =   393216
            Rows            =   3
            Cols            =   14
            FixedCols       =   2
            BackColorFixed  =   14737632
            Enabled         =   -1  'True
            FocusRect       =   2
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
      Begin VB.CommandButton cmdSave 
         Caption         =   "&Save"
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
         Height          =   360
         Left            =   7485
         TabIndex        =   2
         ToolTipText     =   "create plan"
         Top             =   2220
         Visible         =   0   'False
         Width           =   1005
      End
      Begin VB.CommandButton cmdCancel 
         Caption         =   "&Cancel"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   8520
         TabIndex        =   1
         ToolTipText     =   "cancel and return to the main window"
         Top             =   2220
         Visible         =   0   'False
         Width           =   1005
      End
      Begin VB.Label Label11 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Direct Approval"
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
         Height          =   225
         Left            =   1305
         TabIndex        =   8
         Top             =   2220
         Width           =   1245
      End
      Begin VB.Shape LegendApproval 
         BackColor       =   &H00FFFF80&
         FillColor       =   &H00FFFF80&
         FillStyle       =   0  'Solid
         Height          =   195
         Left            =   1035
         Top             =   2220
         Width           =   180
      End
      Begin VB.Label Label9 
         Caption         =   "Legend :"
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
         Left            =   180
         TabIndex        =   7
         Top             =   2205
         Width           =   750
      End
      Begin VB.Shape LegendWebApproval 
         BackColor       =   &H00FFFF80&
         FillColor       =   &H00FFC0FF&
         FillStyle       =   0  'Solid
         Height          =   195
         Left            =   2655
         Top             =   2220
         Width           =   180
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Web based Approval"
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
         Height          =   225
         Left            =   2925
         TabIndex        =   6
         Top             =   2220
         Width           =   1680
      End
   End
End
Attribute VB_Name = "Frm_MPOtherMonthlyBudget"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'
Option Explicit
Dim strMP_Plan_Dim_Id As String
Dim rsTemp As ADODB.Recordset
Dim EditMode As Boolean
Dim intCol As Integer, intRow As Integer

Private Function bulan_belum_lewat(strYearMonth As String) As Boolean
    Dim strCurrentYearMonth As String
    bulan_belum_lewat = True
    recDate.Requery
    strCurrentYearMonth = CStr(Year(recDate(0))) & Right("0" & CStr(month(recDate(0))), 2)
    
    If CDbl(strYearMonth) < CDbl(strCurrentYearMonth) Then
        'bulan_belum_lewat = False 'LOCK VERSION
        bulan_belum_lewat = True 'UNLOCK VERSION
    End If
End Function

Private Sub cmdSave_Click()
    SaveChanges
End Sub

Private Sub FG_OT_Monthly_Budget_DblClick()
    MarkCellPosition
    If bulan_belum_lewat(Mid(strMP_Plan_Dim_Id, 11, 4) & Right("0" & CStr(EngMonthIndex(FG_OT_Monthly_Budget.TextMatrix(0, intCol))), 2)) Then
        If FG_OT_Monthly_Budget.CellBackColor <> LegendApproval.FillColor And FG_OT_Monthly_Budget.CellBackColor <> LegendWebApproval.FillColor Then
            Call StartTyping("Edit")
        End If
    End If
End Sub

Private Sub FG_OT_Monthly_Budget_KeyDown(KeyCode As Integer, Shift As Integer)
    MarkCellPosition
    If bulan_belum_lewat(Mid(strMP_Plan_Dim_Id, 11, 4) & Right("0" & CStr(EngMonthIndex(FG_OT_Monthly_Budget.TextMatrix(0, intCol))), 2)) Then
        If FG_OT_Monthly_Budget.CellBackColor <> LegendApproval.FillColor And FG_OT_Monthly_Budget.CellBackColor <> LegendWebApproval.FillColor Then
            Select Case KeyCode
                Case 113 'F2
                    Call StartTyping("Edit")
                Case 46 'Delete
                    MarkCellPosition
                    ClearCell
            End Select
        End If
    End If
End Sub

Private Sub FG_OT_Monthly_Budget_KeyPress(KeyAscii As Integer)
    MarkCellPosition
    If Asc("0") <= KeyAscii And Asc("9") >= KeyAscii Then
        If bulan_belum_lewat(Mid(strMP_Plan_Dim_Id, 11, 4) & Right("0" & CStr(EngMonthIndex(FG_OT_Monthly_Budget.TextMatrix(0, intCol))), 2)) Then
            If FG_OT_Monthly_Budget.CellBackColor <> LegendApproval.FillColor And FG_OT_Monthly_Budget.CellBackColor <> LegendWebApproval.FillColor Then
                Call StartTyping("Replace", KeyAscii)
            End If
        End If
    End If
End Sub

Private Sub ClearCell()
    With FG_OT_Monthly_Budget
        'clear nett
        If .TextMatrix(1, intCol) <> "" Then
            .TextMatrix(1, intCol) = Empty
            cmdSave.Enabled = True
        End If
        'clear gross
        If .TextMatrix(2, intCol) <> "" Then
            .TextMatrix(2, intCol) = Empty
            cmdSave.Enabled = True
        End If
    End With
End Sub

Private Sub MarkCellPosition()
    intCol = FG_OT_Monthly_Budget.col
    intRow = FG_OT_Monthly_Budget.Row
End Sub

Private Sub Form_Load()
    Dim i As Integer
    
    cmdSave.Enabled = False
    Set rsTemp = New ADODB.Recordset
    
    With Frm_MPActivityDetail.tdg_OTMedium
        strMP_Plan_Dim_Id = .Columns(0)
    End With
    
    With FG_OT_Monthly_Budget
        .ColWidth(0) = 1000
        .ColWidth(1) = 0
        .TextMatrix(1, 0) = "Gross"
        .TextMatrix(2, 0) = "Nett"
        For i = 2 To 13
            .ColWidth(i) = 1500
            .TextMatrix(0, i) = EngMonthName(i - 1)
            .ColAlignment(i) = 4
        Next
    
        rsTemp.Open "select month_number,gross_budget,nett_budget from mp_other_monthly_budget where mp_plan_dim_id='" & strMP_Plan_Dim_Id & "'", ConnERP, 1, 3
        While Not rsTemp.EOF
            .TextMatrix(1, rsTemp(0) + 1) = FormatNumber(rsTemp(1), 2)
            .TextMatrix(2, rsTemp(0) + 1) = FormatNumber(rsTemp(2), 2)
            rsTemp.MoveNext
        Wend
        rsTemp.Close
        
        rsTemp.Open "select month_number,approval from mp_monthly_activity where mp_medium_id = (select mp_medium_id from mp_ids where mp_plan_dim_id='" & strMP_Plan_Dim_Id & "') ", ConnERP, 1, 3
        While Not rsTemp.EOF
            Select Case rsTemp(1)
            Case 1
                .Row = 1
                .col = rsTemp(0) + 1
                .CellBackColor = LegendApproval.FillColor
                .Row = 2
                .CellBackColor = LegendApproval.FillColor
            Case 2
                .Row = 1
                .col = rsTemp(0) + 1
                .CellBackColor = LegendWebApproval.FillColor
                .Row = 2
                .CellBackColor = LegendWebApproval.FillColor
            End Select
            
            rsTemp.MoveNext
        Wend
        rsTemp.Close
    End With
    
End Sub

Private Sub StartTyping(mode As String, Optional KeyAscii As Integer)

    With FG_OT_Monthly_Budget
        txt_Temp.Text = Empty
        txt_Temp.Top = .CellTop + .Top
        txt_Temp.Left = .CellLeft + .Left
        txt_Temp.Width = .CellWidth
        txt_Temp.Height = .CellHeight
        If mode = "Edit" Then
            If .Text <> "" Then txt_Temp.Text = RemoveNumberFormat(.Text)
            EditMode = True
        Else
            txt_Temp.Text = Chr(KeyAscii)
            EditMode = False
        End If
        txt_Temp.SelStart = Len(txt_Temp.Text)
        txt_Temp.Visible = True
        txt_Temp.SetFocus
    End With
    
End Sub

Private Sub cmdCancel_Click()

    Unload Me
End Sub

Private Sub SaveChanges()
    Dim i As Integer
    Dim pesan
    With FG_OT_Monthly_Budget
        For i = 2 To 13
            ConnERP.Execute "delete from mp_other_monthly_budget where mp_plan_dim_id='" & strMP_Plan_Dim_Id & "' and month_number = " & i - 1
            If .TextMatrix(1, i) <> "" Then
                ConnERP.Execute "insert into mp_other_monthly_budget values ('" & strMP_Plan_Dim_Id & "'," & i - 1 & ",'" & EngMonthName(i - 1) & "'," & RemoveNumberFormat(.TextMatrix(2, i)) & "," & RemoveNumberFormat(.TextMatrix(1, i)) & ")"
            End If
        Next
    End With
    cmdSave.Enabled = False
    pesan = MsgBox("Data saved!", vbInformation + vbOKOnly, strApplication_Name)
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Dim pesan
    If cmdSave.Enabled Then
        pesan = MsgBox("Save Changes?", vbQuestion + vbYesNo, strApplication_Name)
        If pesan = 6 Then
            SaveChanges
        End If
    End If
End Sub



Private Sub txt_Temp_KeyDown(KeyCode As Integer, Shift As Integer)
    With FG_OT_Monthly_Budget
        Select Case KeyCode
            Case 13 'Enter
                .SetFocus
            Case 27 'Escape
                 txt_Temp.Visible = False
                .SetFocus
            Case 37 'panah kiri
                If Not EditMode Then
                    .SetFocus
                    If intCol > .FixedCols Then
                        .col = intCol - 1
                    End If
                End If
            Case 38 'panah atas
                .SetFocus
                If intRow > .FixedRows + 1 Then
                    .Row = intRow - 1
                End If
            Case 39 'panah kanan
                If Not EditMode Then
                    .SetFocus
                    If intCol < .cols - 2 Then
                        .col = intCol + 1
                    End If
                End If
            Case 40 'panah bawah
                .SetFocus
                If intRow < .Rows - 1 Then
                    .Row = intRow + 1
                End If
        End Select
    End With
End Sub

Private Sub txt_Temp_KeyPress(KeyAscii As Integer)
    If (Chr(KeyAscii) < "0" Or Chr(KeyAscii) > "9") And KeyAscii <> 8 And KeyAscii <> 13 And KeyAscii <> 27 Then
        KeyAscii = 0
        Beep
    End If
End Sub

Private Sub txt_Temp_LostFocus()
    With txt_Temp
        If .Visible Then
            If .Text <> "" And FG_OT_Monthly_Budget.TextMatrix(intRow, intCol) <> FormatNumber(Val(.Text), 2) Then
                FG_OT_Monthly_Budget.TextMatrix(intRow, intCol) = FormatNumber(Val(txt_Temp), 2)
                cmdSave.Enabled = True
                If intRow = 1 Then
                    If FG_OT_Monthly_Budget.TextMatrix(2, intCol) = "" Then FG_OT_Monthly_Budget.TextMatrix(2, intCol) = "0.00"
                Else
                    If FG_OT_Monthly_Budget.TextMatrix(1, intCol) = "" Then FG_OT_Monthly_Budget.TextMatrix(1, intCol) = "0.00"
                End If
            End If
            .Visible = False
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

    Select Case Index
        Case enButtonType.bieSave  'call db_New.
            Call cmdSave_Click
        Case Else
            Call cmdCancel_Click
    End Select

End Sub

