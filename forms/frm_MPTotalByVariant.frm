VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frm_MPTotalByVariant 
   Caption         =   "Summary by Variant"
   ClientHeight    =   10005
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11565
   LinkTopic       =   "Form1"
   ScaleHeight     =   10005
   ScaleWidth      =   11565
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmd_export 
      Caption         =   "Export To Excel"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   480
      Left            =   6300
      TabIndex        =   10
      Top             =   150
      Width           =   1470
   End
   Begin VB.Frame Frame_View 
      Height          =   570
      Left            =   7950
      TabIndex        =   2
      Top             =   60
      Width           =   3465
      Begin VB.CheckBox cbkOtherCost 
         Caption         =   "Other Cost"
         ForeColor       =   &H000000FF&
         Height          =   210
         Left            =   2295
         TabIndex        =   9
         Top             =   225
         Width           =   1140
      End
      Begin VB.CheckBox cbkFee 
         Caption         =   "Fee"
         ForeColor       =   &H000000FF&
         Height          =   210
         Left            =   1605
         TabIndex        =   8
         Top             =   225
         Width           =   615
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
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
         TabIndex        =   3
         Top             =   210
         Width           =   1290
      End
   End
   Begin MSFlexGridLib.MSFlexGrid FGTotalByTask_Plan 
      Height          =   2835
      Left            =   135
      TabIndex        =   0
      Top             =   720
      Width           =   11265
      _ExtentX        =   19870
      _ExtentY        =   5001
      _Version        =   393216
      Cols            =   14
      FixedCols       =   0
      BackColor       =   16777215
      GridColor       =   0
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
   Begin MSFlexGridLib.MSFlexGrid FGTotalByTask_Actual 
      Height          =   3240
      Left            =   120
      TabIndex        =   1
      Top             =   3960
      Width           =   11265
      _ExtentX        =   19870
      _ExtentY        =   5715
      _Version        =   393216
      Cols            =   14
      FixedCols       =   0
      BackColor       =   16777215
      GridColor       =   0
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
   Begin MSFlexGridLib.MSFlexGrid FGBalance 
      Height          =   2265
      Left            =   120
      TabIndex        =   6
      Top             =   7620
      Width           =   11265
      _ExtentX        =   19870
      _ExtentY        =   3995
      _Version        =   393216
      Cols            =   14
      FixedCols       =   0
      BackColor       =   16777215
      GridColor       =   0
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
   Begin VB.Label Lbl_Balance 
      Caption         =   "BALANCE :"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   255
      Left            =   120
      TabIndex        =   7
      Top             =   7305
      Width           =   1770
   End
   Begin VB.Label Lbl_BudgetActual 
      Caption         =   "ACTUAL BUDGET :"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   255
      Left            =   120
      TabIndex        =   5
      Top             =   3645
      Width           =   1770
   End
   Begin VB.Label Lbl_BudgetPlan 
      Caption         =   "PLAN BUDGET :"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   255
      Left            =   135
      TabIndex        =   4
      Top             =   405
      Width           =   1770
   End
End
Attribute VB_Name = "frm_MPTotalByVariant"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim MarginKanan As Single
Dim MarginBawah As Single
Dim MarginKiri As Single
Dim MarginAtas As Single
Dim Plan_Actual_Space As Single
Dim Actual_Balance_Space As Single
Dim Actual_label As Single
Dim Balance_label As Single
Dim min_height As Single
Dim min_width As Single
Dim str_mp_number As String
Dim xls As New Excel.Application
Dim xls_wb As New Excel.Workbook
Dim xls_ws As New Excel.Worksheet
'
Private Sub cbkFee_Click()
    Me.MousePointer = vbHourglass
    Call show_Summary_Plan(str_mp_number)
    Call show_Summary_Actual(str_mp_number)
    Call calculate_balance
    Me.MousePointer = vbDefault
End Sub

Private Sub cbkOtherCost_Click()
    Me.MousePointer = vbHourglass
    Call show_Summary_Plan(str_mp_number)
    Call show_Summary_Actual(str_mp_number)
    Call calculate_balance
    Me.MousePointer = vbDefault
End Sub


Private Sub cmd_export_Click()
    Dim int_row As Long
    Dim int_col As Long
    Dim row_plan As Long
    Dim rowactual As Long
    Dim row_balance As Long
    
    Me.MousePointer = vbHourglass
    
    Set xls_wb = xls.Workbooks.Add
    Set xls_ws = xls_wb.Worksheets(1)
    xls.Visible = False
    
    int_row = 2
    
    If FGTotalByTask_Plan.Rows > 0 Then
        
        With xls_ws
            .Cells(int_row, 1) = "PLAN BUDGET"
            .Cells(int_row, 1).Font.Bold = True
            int_row = int_row + 1
            
            For row_plan = 0 To FGTotalByTask_Plan.Rows - 1
                If FGTotalByTask_Plan.TextMatrix(row_plan, 0) = "Sub Total" Or FGTotalByTask_Plan.TextMatrix(row_plan, 0) = "SUMMARY" Or FGTotalByTask_Plan.TextMatrix(row_plan, 0) = "Grand Total" Then
                    .Cells(int_row, 1).Font.Bold = True
                End If
                
                .Cells(int_row, 1) = FGTotalByTask_Plan.TextMatrix(row_plan, 0)
                .Cells(int_row, 2) = FGTotalByTask_Plan.TextMatrix(row_plan, 1)
                .Cells(int_row, 3) = FGTotalByTask_Plan.TextMatrix(row_plan, 2)
                .Cells(int_row, 4) = FGTotalByTask_Plan.TextMatrix(row_plan, 3)
                .Cells(int_row, 5) = FGTotalByTask_Plan.TextMatrix(row_plan, 4)
                .Cells(int_row, 6) = FGTotalByTask_Plan.TextMatrix(row_plan, 5)
                .Cells(int_row, 7) = FGTotalByTask_Plan.TextMatrix(row_plan, 6)
                .Cells(int_row, 8) = FGTotalByTask_Plan.TextMatrix(row_plan, 7)
                .Cells(int_row, 9) = FGTotalByTask_Plan.TextMatrix(row_plan, 8)
                .Cells(int_row, 10) = FGTotalByTask_Plan.TextMatrix(row_plan, 9)
                .Cells(int_row, 11) = FGTotalByTask_Plan.TextMatrix(row_plan, 10)
                .Cells(int_row, 12) = FGTotalByTask_Plan.TextMatrix(row_plan, 11)
                .Cells(int_row, 13) = FGTotalByTask_Plan.TextMatrix(row_plan, 12)
                .Cells(int_row, 14) = FGTotalByTask_Plan.TextMatrix(row_plan, 13)
                .Cells(int_row, 15) = FGTotalByTask_Plan.TextMatrix(row_plan, 14)
                'row_plan = row_plan + 1
                int_row = int_row + 1
            Next
        End With
       End If
       
      If FGTotalByTask_Actual.Rows > 0 Then
              With xls_ws

            int_row = int_row + 2
            .Cells(int_row, 1) = "ACTUAL BUDGET"
            .Cells(int_row, 1).Font.Bold = True
            int_row = int_row + 1
            
            For rowactual = 0 To FGTotalByTask_Actual.Rows - 1
                If FGTotalByTask_Actual.TextMatrix(rowactual, 0) = "Sub Total" Or FGTotalByTask_Actual.TextMatrix(rowactual, 0) = "SUMMARY" Or FGTotalByTask_Actual.TextMatrix(rowactual, 0) = "Grand Total" Then
                    .Cells(int_row, 1).Font.Bold = True
                End If
                .Cells(int_row, 1) = FGTotalByTask_Actual.TextMatrix(rowactual, 0)
                .Cells(int_row, 2) = FGTotalByTask_Actual.TextMatrix(rowactual, 1)
                .Cells(int_row, 3) = FGTotalByTask_Actual.TextMatrix(rowactual, 2)
                .Cells(int_row, 4) = FGTotalByTask_Actual.TextMatrix(rowactual, 3)
                .Cells(int_row, 5) = FGTotalByTask_Actual.TextMatrix(rowactual, 4)
                .Cells(int_row, 6) = FGTotalByTask_Actual.TextMatrix(rowactual, 5)
                .Cells(int_row, 7) = FGTotalByTask_Actual.TextMatrix(rowactual, 6)
                .Cells(int_row, 8) = FGTotalByTask_Actual.TextMatrix(rowactual, 7)
                .Cells(int_row, 9) = FGTotalByTask_Actual.TextMatrix(rowactual, 8)
                .Cells(int_row, 10) = FGTotalByTask_Actual.TextMatrix(rowactual, 9)
                .Cells(int_row, 11) = FGTotalByTask_Actual.TextMatrix(rowactual, 10)
                .Cells(int_row, 12) = FGTotalByTask_Actual.TextMatrix(rowactual, 11)
                .Cells(int_row, 13) = FGTotalByTask_Actual.TextMatrix(rowactual, 12)
                .Cells(int_row, 14) = FGTotalByTask_Actual.TextMatrix(rowactual, 13)
                .Cells(int_row, 15) = FGTotalByTask_Actual.TextMatrix(rowactual, 14)
                'rowactual = rowactual + 1
                int_row = int_row + 1
            Next
            End With
        End If
        
      If FGBalance.Rows > 0 Then
                With xls_ws

            int_row = int_row + 2
            .Cells(int_row, 1) = "BALANCE"
            .Cells(int_row, 1).Font.Bold = True
            int_row = int_row + 1
            
            For row_balance = 0 To FGBalance.Rows - 1
                .Cells(int_row, 1) = FGBalance.TextMatrix(row_balance, 0)
                .Cells(int_row, 2) = FGBalance.TextMatrix(row_balance, 1)
                .Cells(int_row, 3) = FGBalance.TextMatrix(row_balance, 2)
                .Cells(int_row, 4) = FGBalance.TextMatrix(row_balance, 3)
                .Cells(int_row, 5) = FGBalance.TextMatrix(row_balance, 4)
                .Cells(int_row, 6) = FGBalance.TextMatrix(row_balance, 5)
                .Cells(int_row, 7) = FGBalance.TextMatrix(row_balance, 6)
                .Cells(int_row, 8) = FGBalance.TextMatrix(row_balance, 7)
                .Cells(int_row, 9) = FGBalance.TextMatrix(row_balance, 8)
                .Cells(int_row, 10) = FGBalance.TextMatrix(row_balance, 9)
                .Cells(int_row, 11) = FGBalance.TextMatrix(row_balance, 10)
                .Cells(int_row, 12) = FGBalance.TextMatrix(row_balance, 11)
                .Cells(int_row, 13) = FGBalance.TextMatrix(row_balance, 12)
                .Cells(int_row, 14) = FGBalance.TextMatrix(row_balance, 13)
                '.Cells(int_row, 15) = FGBalance.TextMatrix(row_balance, 14)
                'row_balance = row_balance + 1
                int_row = int_row + 1
            Next
        End With
    End If
    
    xls.Columns("A:A").ColumnWidth = 21.29
    xls.Columns("B:N").NumberFormat = "_(* #,##0.0_);_(* (#,##0.0);_(* ""-""??_);_(@_)"
    xls.Columns("B:N").NumberFormat = "_(* #,##0_);_(* (#,##0);_(* ""-""??_);_(@_)"
    xls.Columns("B:N").ColumnWidth = 14.14
    
    xls.Visible = True
    Me.MousePointer = vbNormal
End Sub

Private Sub Form_Load()

    Me.MousePointer = vbHourglass
    
    MarginKanan = Me.Width - (FGTotalByTask_Plan.Left + FGTotalByTask_Plan.Width)
    MarginBawah = Me.Height - (FGBalance.Top + FGBalance.Height)
    MarginKiri = FGTotalByTask_Plan.Left
    MarginAtas = FGTotalByTask_Plan.Top
    
    Plan_Actual_Space = FGTotalByTask_Actual.Top - (FGTotalByTask_Plan.Top + FGTotalByTask_Plan.Height)
    Actual_Balance_Space = FGBalance.Top - (FGTotalByTask_Actual.Top + FGTotalByTask_Actual.Height)
    
    Actual_label = FGTotalByTask_Actual.Top - Lbl_BudgetActual.Top
    Balance_label = FGBalance.Top - Lbl_Balance.Top
    
    min_height = Me.Height - 1000
    min_width = Me.Width - 1000
    
    Call initGrid
    
    str_mp_number = frm_MPInsertion.cboMPNumber.Text
    
    If str_mp_number = "" Then str_mp_number = frm_MediaPlan_View.cboMPNumber.Text
    
    Call show_Summary_Plan(str_mp_number)
    
    Call show_Summary_Actual(str_mp_number)
    
    Call calculate_balance
    
    
    Me.MousePointer = vbDefault
End Sub

Private Sub show_Summary_Plan(strMPNumber As String)

    Dim strSql As String, i As Integer, strTaskID As String, j As Integer
    Dim rsSummary As New ADODB.Recordset
    Dim intTaskNumber As Integer, intTaskRow As Integer, intPrintRow As Integer
    
    Dim TotalTV(13) As Double
    Dim TotalRD(13) As Double
    Dim TotalPR(13) As Double
    Dim TotalCN(13) As Double
    Dim TotalOT(13) As Double
    Dim GrandTotal(13) As Double
    'set value
    For i = 1 To 13
        TotalTV(i - 1) = 0
        TotalRD(i - 1) = 0
        TotalPR(i - 1) = 0
        TotalCN(i - 1) = 0
        TotalOT(i - 1) = 0
        GrandTotal(i - 1) = 0
    Next
    
    'Clear Grid
    FGTotalByTask_Plan.Rows = 1
    
    strSql = "SELECT c.brand_variant_code,c.brand_variant_Name,b.medium_code,a.month_number,"
    
    If cbkFee.Value = 1 And cbkOtherCost.Value = 1 Then 'Nett+fee+otherCost
        strSql = strSql & " sum(a.nett_fee_other) plan_budget"
    Else
        If cbkFee.Value = 1 Then 'Nett+fee
            strSql = strSql & " sum(a.nett_plus_fee) plan_budget"
        Else
            If cbkOtherCost.Value = 1 Then 'Nett + OtherCost
                strSql = strSql & " sum(a.nett_plus_other) plan_budget"
            Else
                'Nett only
                strSql = strSql & "sum(a.nett_only) plan_budget"
            End If
        End If
    End If
    
    strSql = strSql & ", max(a.is_actual) is_actual FROM ("
    
    '==========new view
    strSql = strSql & " select mp_medium_id,month_number, "
    
    strSql = strSql & " Case IsNull(Total_Actual, -1) when -1 then min_budget "
    strSql = strSql & " Else Actual_Nett_Paid + Actual_Nett_Bonus end Nett_only,"
    
    strSql = strSql & " Case IsNull(Total_Actual, -1) when -1 then min_budget + MSC_Paid_Value + MSC_Bonus_Value + Club_Agency_Value "
    strSql = strSql & " Else Actual_Nett_Paid + Actual_Nett_Bonus + Actual_MSC_Paid + Actual_MSC_Bonus end Nett_plus_fee,"
    
    strSql = strSql & " Case IsNull(Total_Actual, -1) when -1 then min_budget + MSC_Paid_Value + MSC_Bonus_Value + Club_Agency_Value + isnull(other_cost,0)"
    strSql = strSql & " Else Actual_Nett_Paid + Actual_Nett_Bonus + Actual_MSC_Paid + Actual_MSC_Bonus  + isnull(actual_other_cost,0) end Nett_fee_other,"
    
    strSql = strSql & " Case IsNull(Total_Actual, -1) when -1 then min_budget  + isnull(other_cost,0)"
    strSql = strSql & " Else Actual_Nett_Paid + Actual_Nett_Bonus + isnull(actual_other_cost,0) end Nett_plus_other,"
    
    strSql = strSql & " case isnull(total_actual,-1) when -1 then 0 else 1 end is_actual From mp_monthly_activity "
    '============
    strSql = strSql & ") a "
    strSql = strSql & " INNER JOIN mp_medium b on a.mp_medium_id = b.mp_medium_id "
    strSql = strSql & " INNER JOIN mp_activity c on b.mp_activity_id = c.mp_activity_id "
    strSql = strSql & " INNER JOIN mp_task d on c.mp_task_id = d.mp_task_id "
    strSql = strSql & "and d.mp_number = '" & strMPNumber & "' "
    
    strSql = strSql & " GROUP BY c.brand_variant_code,c.brand_variant_Name,b.medium_code,a.month_number "
    
    
    rsSummary.Open strSql, ConnERP, 1, 3
    If Not rsSummary.EOF Then
        intTaskNumber = 0
        strTaskID = ""
        
        While Not rsSummary.EOF
        
            If rsSummary(0) <> strTaskID Then
                
                'new Task
                strTaskID = rsSummary(0)
                intTaskNumber = intTaskNumber + 1
                intTaskRow = (7 * intTaskNumber) - 6
                With FGTotalByTask_Plan
                    .Rows = FGTotalByTask_Plan.Rows + 7
                    
                    '.TextMatrix(intTaskRow, 0) = "TASK " & intTaskNumber & " : " & rsSummary(1)
                    .TextMatrix(intTaskRow, 0) = "Brand Variant : " & rsSummary(1)
                    .Row = intTaskRow
                    .col = 0
                    .CellFontBold = True
                    
                    .TextMatrix(intTaskRow + 1, 0) = "TV"
                    .TextMatrix(intTaskRow + 2, 0) = "Radio"
                    .TextMatrix(intTaskRow + 3, 0) = "Print"
                    .TextMatrix(intTaskRow + 4, 0) = "Cinema"
                    .TextMatrix(intTaskRow + 5, 0) = "Other"
                    
                    .TextMatrix(intTaskRow + 6, 0) = "Sub Total"
                    .Row = intTaskRow + 6
                    .col = 0
                    .CellFontBold = True
                    
                 End With
            End If
            Select Case rsSummary(2)
                Case "TV":
                    intPrintRow = intTaskRow + 1
                    TotalTV(rsSummary(3) - 1) = TotalTV(rsSummary(3) - 1) + rsSummary(4)
                    TotalTV(12) = TotalTV(12) + rsSummary(4)
                Case "RD":
                    intPrintRow = intTaskRow + 2
                    TotalRD(rsSummary(3) - 1) = TotalRD(rsSummary(3) - 1) + rsSummary(4)
                    TotalRD(12) = TotalRD(12) + rsSummary(4)
                Case "PR":
                    intPrintRow = intTaskRow + 3
                    TotalPR(rsSummary(3) - 1) = TotalPR(rsSummary(3) - 1) + rsSummary(4)
                    TotalPR(12) = TotalPR(12) + rsSummary(4)
                Case "CN":
                    intPrintRow = intTaskRow + 4
                    TotalCN(rsSummary(3) - 1) = TotalCN(rsSummary(3) - 1) + rsSummary(4)
                    TotalCN(12) = TotalCN(12) + rsSummary(4)
                Case "OT":
                    intPrintRow = intTaskRow + 5
                    TotalOT(rsSummary(3) - 1) = TotalOT(rsSummary(3) - 1) + rsSummary(4)
                    TotalOT(12) = TotalOT(12) + rsSummary(4)
            End Select
            GrandTotal(rsSummary(3) - 1) = GrandTotal(rsSummary(3) - 1) + rsSummary(4)
            GrandTotal(12) = GrandTotal(12) + rsSummary(4)
            
            'plan budget per task per medium (monthly)
            FGTotalByTask_Plan.TextMatrix(intPrintRow, rsSummary(3)) = Format(rsSummary(4), "#,##0")
            
            If rsSummary("is_actual").Value = 1 Then 'is Actual
                FGTotalByTask_Plan.Row = intPrintRow
                FGTotalByTask_Plan.col = rsSummary(3)
                FGTotalByTask_Plan.CellBackColor = vbGreen
            End If
            
            'plan budget per task per medium(1 Year)
            If FGTotalByTask_Plan.TextMatrix(intPrintRow, 13) <> "" Then
                FGTotalByTask_Plan.TextMatrix(intPrintRow, 13) = Format(Val(RemoveNumberFormat(FGTotalByTask_Plan.TextMatrix(intPrintRow, 13) & ".00")) + rsSummary(4), "#,##0")
            Else
                FGTotalByTask_Plan.TextMatrix(intPrintRow, 13) = Format(rsSummary(4), "#,##0")
            End If
            'Sub total per task (monthly)
            If FGTotalByTask_Plan.TextMatrix(intTaskRow + 6, rsSummary(3)) <> "" Then
                FGTotalByTask_Plan.TextMatrix(intTaskRow + 6, rsSummary(3)) = Format(Val(RemoveNumberFormat(FGTotalByTask_Plan.TextMatrix(intTaskRow + 6, rsSummary(3)) & ".00")) + rsSummary(4), "#,##0")
            Else
                FGTotalByTask_Plan.TextMatrix(intTaskRow + 6, rsSummary(3)) = Format(rsSummary(4), "#,##0")
            End If
            'Sub Total per task (1 Year)
            If FGTotalByTask_Plan.TextMatrix(intTaskRow + 6, 13) <> "" Then
                FGTotalByTask_Plan.TextMatrix(intTaskRow + 6, 13) = Format(Val(RemoveNumberFormat(FGTotalByTask_Plan.TextMatrix(intTaskRow + 6, 13) & ".00")) + rsSummary(4), "#,##0")
            Else
                FGTotalByTask_Plan.TextMatrix(intTaskRow + 6, 13) = Format(rsSummary(4), "#,##0")
            End If
            rsSummary.MoveNext
        Wend
        'Grand Total
        With FGTotalByTask_Plan
            .Rows = .Rows + 7
            
            .TextMatrix(.Rows - 7, 0) = "SUMMARY"
            .Row = .Rows - 7
            .col = 0
            .CellFontBold = True
            
            .TextMatrix(.Rows - 6, 0) = "Total TV"
            .TextMatrix(.Rows - 5, 0) = "Total Radio"
            .TextMatrix(.Rows - 4, 0) = "Total Print"
            .TextMatrix(.Rows - 3, 0) = "Total Cinema"
            .TextMatrix(.Rows - 2, 0) = "Total Other"
            
            .TextMatrix(.Rows - 1, 0) = "Grand Total"
            .Row = .Rows - 1
            .col = 0
            .CellFontBold = True
            
            For i = 1 To 13
                .TextMatrix(.Rows - 6, i) = Format(TotalTV(i - 1), "#,##0")
                .TextMatrix(.Rows - 5, i) = Format(TotalRD(i - 1), "#,##0")
                .TextMatrix(.Rows - 4, i) = Format(TotalPR(i - 1), "#,##0")
                .TextMatrix(.Rows - 3, i) = Format(TotalCN(i - 1), "#,##0")
                .TextMatrix(.Rows - 2, i) = Format(TotalOT(i - 1), "#,##0")
                .TextMatrix(.Rows - 1, i) = Format(GrandTotal(i - 1), "#,##0")
            Next
            
            If GrandTotal(12) = 0 Then
                .TextMatrix(.Rows - 6, .cols - 1) = Format(0, "#,##0.00%")
                .TextMatrix(.Rows - 5, .cols - 1) = Format(0, "#,##0.00%")
                .TextMatrix(.Rows - 4, .cols - 1) = Format(0, "#,##0.00%")
                .TextMatrix(.Rows - 3, .cols - 1) = Format(0, "#,##0.00%")
                .TextMatrix(.Rows - 2, .cols - 1) = Format(0, "#,##0.00%")
            Else
                .TextMatrix(.Rows - 6, .cols - 1) = Format(TotalTV(12) / GrandTotal(12), "#,##0.00%")
                .TextMatrix(.Rows - 5, .cols - 1) = Format(TotalRD(12) / GrandTotal(12), "#,##0.00%")
                .TextMatrix(.Rows - 4, .cols - 1) = Format(TotalPR(12) / GrandTotal(12), "#,##0.00%")
                .TextMatrix(.Rows - 3, .cols - 1) = Format(TotalCN(12) / GrandTotal(12), "#,##0.00%")
                .TextMatrix(.Rows - 2, .cols - 1) = Format(TotalOT(12) / GrandTotal(12), "#,##0.00%")
            End If
            
            'mewarnai..
            For i = 1 To intTaskNumber + 1
                .Row = 7 * i
                For j = 1 To .cols - 1
                    .col = j
                    .CellBackColor = vbYellow
                Next
            Next
        End With
    End If
    rsSummary.Close
    Set rsSummary = Nothing
End Sub


Private Sub show_Summary_Actual(strMPNumber As String)
    Dim strSql As String, i As Integer, strTaskID As String, j As Integer
    Dim rsSummary As New ADODB.Recordset
    Dim intTaskNumber As Integer, intTaskRow As Integer, intPrintRow As Integer
    
    Dim TotalTV(13) As Double
    Dim TotalRD(13) As Double
    Dim TotalPR(13) As Double
    Dim TotalCN(13) As Double
    Dim TotalOT(13) As Double
    Dim GrandTotal(13) As Double
    'set value
    For i = 1 To 13
        TotalTV(i - 1) = 0
        TotalRD(i - 1) = 0
        TotalPR(i - 1) = 0
        TotalCN(i - 1) = 0
        TotalOT(i - 1) = 0
        GrandTotal(i - 1) = 0
    Next
    
    'Clear Grid
    FGTotalByTask_Actual.Rows = 1
    
    strSql = "SELECT c.brand_variant_code,c.brand_variant_Name,b.medium_code,a.month_number,"
    
    If cbkFee.Value = 1 And cbkOtherCost.Value = 1 Then 'Nett+fee+otherCost
        strSql = strSql & "sum(isnull(a.actual_nett_paid,0) + isnull(a.actual_nett_bonus,0) + isnull(actual_MSC_Paid,0) + isnull(actual_MSC_Bonus,0) + isnull(actual_other_cost,0)) actual_budget"
    Else
        If cbkFee.Value = 1 Then 'Nett+fee
            strSql = strSql & "sum(isnull(a.actual_nett_paid,0) + isnull(a.actual_nett_bonus,0) + isnull(actual_MSC_Paid,0) + isnull(actual_MSC_Bonus,0)) actual_budget"
        Else
            If cbkOtherCost.Value = 1 Then 'Nett + OtherCost
                strSql = strSql & "sum(isnull(a.actual_nett_paid,0) + isnull(a.actual_nett_bonus,0) + isnull(actual_other_cost,0)) actual_budget"
            Else
                'Nett only
                strSql = strSql & "sum(isnull(a.actual_nett_paid,0) + isnull(a.actual_nett_bonus,0)) actual_budget"
            End If
        End If
    End If
    
    strSql = strSql & " FROM mp_monthly_activity a "
    strSql = strSql & " INNER JOIN mp_medium b on a.mp_medium_id = b.mp_medium_id "
    strSql = strSql & " INNER JOIN mp_activity c on b.mp_activity_id = c.mp_activity_id "
    strSql = strSql & " INNER JOIN mp_task d on c.mp_task_id = d.mp_task_id "
    strSql = strSql & "and d.mp_number = '" & strMPNumber & "' "
    strSql = strSql & " GROUP BY c.brand_variant_code,c.brand_variant_Name,b.medium_code,a.month_number "
    
    rsSummary.Open strSql, ConnERP, 1, 3
    If Not rsSummary.EOF Then
        intTaskNumber = 0
        strTaskID = ""
        While Not rsSummary.EOF
            If rsSummary(0) <> strTaskID Then
                'new Task
                strTaskID = rsSummary(0)
                intTaskNumber = intTaskNumber + 1
                intTaskRow = (7 * intTaskNumber) - 6
                With FGTotalByTask_Actual
                    .Rows = FGTotalByTask_Actual.Rows + 7
                    
                    '.TextMatrix(intTaskRow, 0) = "TASK " & intTaskNumber & " : " & rsSummary(1)
                    .TextMatrix(intTaskRow, 0) = "Brand Variant : " & rsSummary(1)
                    .Row = intTaskRow
                    .col = 0
                    .CellFontBold = True
                    
                    .TextMatrix(intTaskRow + 1, 0) = "TV"
                    .TextMatrix(intTaskRow + 2, 0) = "Radio"
                    .TextMatrix(intTaskRow + 3, 0) = "Print"
                    .TextMatrix(intTaskRow + 4, 0) = "Cinema"
                    .TextMatrix(intTaskRow + 5, 0) = "Other"
                    
                    .TextMatrix(intTaskRow + 6, 0) = "Sub Total"
                    .Row = intTaskRow + 6
                    .col = 0
                    .CellFontBold = True
                    
                 End With
            End If
            Select Case rsSummary(2)
                Case "TV":
                    intPrintRow = intTaskRow + 1
                    TotalTV(rsSummary(3) - 1) = TotalTV(rsSummary(3) - 1) + rsSummary(4)
                    TotalTV(12) = TotalTV(12) + rsSummary(4)
                Case "RD":
                    intPrintRow = intTaskRow + 2
                    TotalRD(rsSummary(3) - 1) = TotalRD(rsSummary(3) - 1) + rsSummary(4)
                    TotalRD(12) = TotalRD(12) + rsSummary(4)
                Case "PR":
                    intPrintRow = intTaskRow + 3
                    TotalPR(rsSummary(3) - 1) = TotalPR(rsSummary(3) - 1) + rsSummary(4)
                    TotalPR(12) = TotalPR(12) + rsSummary(4)
                Case "CN":
                    intPrintRow = intTaskRow + 4
                    TotalCN(rsSummary(3) - 1) = TotalCN(rsSummary(3) - 1) + rsSummary(4)
                    TotalCN(12) = TotalCN(12) + rsSummary(4)
                Case "OT":
                    intPrintRow = intTaskRow + 5
                    TotalOT(rsSummary(3) - 1) = TotalOT(rsSummary(3) - 1) + rsSummary(4)
                    TotalOT(12) = TotalOT(12) + rsSummary(4)
            End Select
            GrandTotal(rsSummary(3) - 1) = GrandTotal(rsSummary(3) - 1) + rsSummary(4)
            GrandTotal(12) = GrandTotal(12) + rsSummary(4)
            
            'plan budget per task per medium (monthly)
            FGTotalByTask_Actual.TextMatrix(intPrintRow, rsSummary(3)) = Format(rsSummary(4), "#,##0")
            'plan budget per task per medium(1 Year)
            If FGTotalByTask_Actual.TextMatrix(intPrintRow, 13) <> "" Then
                FGTotalByTask_Actual.TextMatrix(intPrintRow, 13) = Format(Val(RemoveNumberFormat(FGTotalByTask_Actual.TextMatrix(intPrintRow, 13) & ".00")) + rsSummary(4), "#,##0")
            Else
                FGTotalByTask_Actual.TextMatrix(intPrintRow, 13) = Format(rsSummary(4), "#,##0")
            End If
            'Sub total per task (monthly)
            If FGTotalByTask_Actual.TextMatrix(intTaskRow + 6, rsSummary(3)) <> "" Then
                FGTotalByTask_Actual.TextMatrix(intTaskRow + 6, rsSummary(3)) = Format(Val(RemoveNumberFormat(FGTotalByTask_Actual.TextMatrix(intTaskRow + 6, rsSummary(3)) & ".00")) + rsSummary(4), "#,##0")
            Else
                FGTotalByTask_Actual.TextMatrix(intTaskRow + 6, rsSummary(3)) = Format(rsSummary(4), "#,##0")
            End If
            'Sub Total per task (1 Year)
            If FGTotalByTask_Actual.TextMatrix(intTaskRow + 6, 13) <> "" Then
                FGTotalByTask_Actual.TextMatrix(intTaskRow + 6, 13) = Format(Val(RemoveNumberFormat(FGTotalByTask_Actual.TextMatrix(intTaskRow + 6, 13) & ".00")) + rsSummary(4), "#,##0")
            Else
                FGTotalByTask_Actual.TextMatrix(intTaskRow + 6, 13) = Format(rsSummary(4), "#,##0")
            End If
            rsSummary.MoveNext
        Wend
        'Grand Total
        With FGTotalByTask_Actual
            .Rows = .Rows + 7
            
            .TextMatrix(.Rows - 7, 0) = "SUMMARY"
            .Row = .Rows - 7
            .col = 0
            .CellFontBold = True
            
            .TextMatrix(.Rows - 6, 0) = "Total TV"
            .TextMatrix(.Rows - 5, 0) = "Total Radio"
            .TextMatrix(.Rows - 4, 0) = "Total Print"
            .TextMatrix(.Rows - 3, 0) = "Total Cinema"
            .TextMatrix(.Rows - 2, 0) = "Total Other"
            
            .TextMatrix(.Rows - 1, 0) = "Grand Total"
            .Row = .Rows - 1
            .col = 0
            .CellFontBold = True
            
            For i = 1 To 13
                .TextMatrix(.Rows - 6, i) = Format(TotalTV(i - 1), "#,##0")
                .TextMatrix(.Rows - 5, i) = Format(TotalRD(i - 1), "#,##0")
                .TextMatrix(.Rows - 4, i) = Format(TotalPR(i - 1), "#,##0")
                .TextMatrix(.Rows - 3, i) = Format(TotalCN(i - 1), "#,##0")
                .TextMatrix(.Rows - 2, i) = Format(TotalOT(i - 1), "#,##0")
                .TextMatrix(.Rows - 1, i) = Format(GrandTotal(i - 1), "#,##0")
            Next
            
            If GrandTotal(12) = 0 Then
                .TextMatrix(.Rows - 6, .cols - 1) = Format(0, "#,##0.00%")
                .TextMatrix(.Rows - 5, .cols - 1) = Format(0, "#,##0.00%")
                .TextMatrix(.Rows - 4, .cols - 1) = Format(0, "#,##0.00%")
                .TextMatrix(.Rows - 3, .cols - 1) = Format(0, "#,##0.00%")
                .TextMatrix(.Rows - 2, .cols - 1) = Format(0, "#,##0.00%")
            Else
                .TextMatrix(.Rows - 6, .cols - 1) = Format(TotalTV(12) / GrandTotal(12), "#,##0.00%")
                .TextMatrix(.Rows - 5, .cols - 1) = Format(TotalRD(12) / GrandTotal(12), "#,##0.00%")
                .TextMatrix(.Rows - 4, .cols - 1) = Format(TotalPR(12) / GrandTotal(12), "#,##0.00%")
                .TextMatrix(.Rows - 3, .cols - 1) = Format(TotalCN(12) / GrandTotal(12), "#,##0.00%")
                .TextMatrix(.Rows - 2, .cols - 1) = Format(TotalOT(12) / GrandTotal(12), "#,##0.00%")
            End If
            
            'mewarnai..
            For i = 1 To intTaskNumber + 1
                .Row = 7 * i
                For j = 1 To .cols - 1
                    .col = j
                    .CellBackColor = vbYellow
                Next
            Next
        End With
    End If
    rsSummary.Close
    Set rsSummary = Nothing
End Sub


Private Sub initGrid()
    Dim i As Integer
    'GRID PLAN
    With FGTotalByTask_Plan
        'setting grid
        .Row = 0
        .col = 0
        .Text = "Medium"
        .ColWidth(0) = 2000
        .CellAlignment = 4
        .CellFontBold = True
        .cols = 15
        For i = 1 To 14
            .col = i
            .Text = EngMonthName(i)
            If i = 13 Then .Text = "Total"
            If i = 14 Then .Text = "%"
            .CellAlignment = 4
            .ColWidth(i) = 1500
            .CellFontBold = True
        Next
        .Rows = 1
        .FixedCols = 1
    End With
    
    'GRID ACUTAL
    With FGTotalByTask_Actual
        'setting grid
        .Row = 0
        .col = 0
        .Text = "Medium"
        .ColWidth(0) = 2000
        .CellAlignment = 4
        .CellFontBold = True
        .cols = 15
        For i = 1 To 14
            .col = i
            .Text = EngMonthName(i)
            If i = 13 Then .Text = "Total"
            If i = 14 Then .Text = "%"
            .CellAlignment = 4
            .ColWidth(i) = 1500
            .CellFontBold = True
        Next
        .Rows = 1
        .FixedCols = 1
    End With
    
    'GRID BALANCE
    With FGBalance
        'setting grid
        .Row = 0
        .col = 0
        .Text = "Medium"
        .ColWidth(0) = 2000
        .CellAlignment = 4
        .CellFontBold = True
        .cols = 14
        For i = 1 To 13
            .col = i
            .Text = EngMonthName(i)
            If i = 13 Then .Text = "Total"
            .CellAlignment = 4
            .ColWidth(i) = 1500
            .CellFontBold = True
        Next
        .Rows = 1
        .FixedCols = 1
        
        .Rows = 7
        .TextMatrix(1, 0) = "TV"
        .TextMatrix(2, 0) = "Radio"
        .TextMatrix(3, 0) = "Print"
        .TextMatrix(4, 0) = "Cinema"
        .TextMatrix(5, 0) = "Other"
        .TextMatrix(6, 0) = "TOTAL"
    End With

End Sub

Private Sub Form_Resize()
    Dim resize_height As Single
    If Me.Width < min_width Then Me.Width = min_width
    If Me.Height < min_height Then Me.Height = min_height
    
    FGTotalByTask_Plan.Width = Me.Width - (MarginKiri + MarginKanan)
    FGTotalByTask_Actual.Width = Me.Width - (MarginKiri + MarginKanan)
    FGBalance.Width = Me.Width - (MarginKiri + MarginKanan)
    Frame_View.Left = Me.Width - MarginKanan - Frame_View.Width
    
    resize_height = Me.Height - (MarginAtas + MarginBawah + FGTotalByTask_Actual.Height + FGTotalByTask_Plan.Height + FGBalance.Height + Plan_Actual_Space + Actual_Balance_Space)
    
    FGTotalByTask_Plan.Height = FGTotalByTask_Plan.Height + (resize_height \ 3)
    
    FGTotalByTask_Actual.Top = FGTotalByTask_Plan.Top + FGTotalByTask_Plan.Height + Plan_Actual_Space
    FGTotalByTask_Actual.Height = FGTotalByTask_Actual.Height + (resize_height \ 3)
    Lbl_BudgetActual.Top = FGTotalByTask_Actual.Top - Actual_label
    
    FGBalance.Top = FGTotalByTask_Actual.Top + FGTotalByTask_Actual.Height + Actual_Balance_Space
    FGBalance.Height = FGBalance.Height + (resize_height \ 3)
    Lbl_Balance.Top = FGBalance.Top - Balance_label
End Sub

Private Sub calculate_balance()
    Dim intRow As Integer
    Dim intCol As Integer
    Dim Balance As Double
    With FGBalance
        FGBalance.ForeColor = vbBlack
        For intRow = 1 To 6
            For intCol = 1 To 13
                Balance = CDbl(FGTotalByTask_Plan.TextMatrix(FGTotalByTask_Plan.Rows - (7 - intRow), intCol)) - CDbl(FGTotalByTask_Actual.TextMatrix(FGTotalByTask_Actual.Rows - (7 - intRow), intCol))
                If Balance <> 0 Then
                    .col = intCol
                    .Row = intRow
                    If Balance < 0 Then
                        .CellForeColor = vbRed
                    Else
                        .CellForeColor = &H8000&     'dark green
                    End If
                    .Text = Format(Balance, "#,##0")
                Else
                    .TextMatrix(intRow, intCol) = Format(Balance, "#,##0")
                End If
            Next
        Next
    End With
End Sub
