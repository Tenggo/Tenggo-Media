VERSION 5.00
Begin VB.Form Frm_MPActualBudgetUpdate 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Actual Budget"
   ClientHeight    =   3510
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4830
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3510
   ScaleWidth      =   4830
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Height          =   2970
      Left            =   135
      TabIndex        =   0
      Top             =   390
      Width           =   4500
      Begin VB.CommandButton cmdCancel 
         Caption         =   "Cancel"
         Height          =   420
         Left            =   3345
         TabIndex        =   12
         Top             =   2415
         Width           =   1050
      End
      Begin VB.CommandButton cmdSave 
         Caption         =   "Save"
         Height          =   420
         Left            =   2235
         TabIndex        =   11
         Top             =   2415
         Width           =   1050
      End
      Begin VB.TextBox txtBonusGross 
         Height          =   315
         Left            =   1890
         TabIndex        =   4
         Top             =   1290
         Width           =   2340
      End
      Begin VB.TextBox txtBonusNett 
         Height          =   315
         Left            =   1890
         TabIndex        =   3
         Top             =   1710
         Width           =   2340
      End
      Begin VB.TextBox txtPaidGross 
         Height          =   315
         Left            =   1905
         TabIndex        =   2
         Top             =   270
         Width           =   2340
      End
      Begin VB.TextBox txtPaidNett 
         Height          =   315
         Left            =   1905
         TabIndex        =   1
         Top             =   690
         Width           =   2340
      End
      Begin VB.Label Label6 
         Alignment       =   1  'Right Justify
         Caption         =   "Nett : "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   1215
         TabIndex        =   10
         Top             =   1740
         Width           =   660
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         Caption         =   "Gross : "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   1170
         TabIndex        =   9
         Top             =   1320
         Width           =   660
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         Caption         =   "Nett : "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   1230
         TabIndex        =   8
         Top             =   720
         Width           =   660
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         Caption         =   "Gross : "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   1185
         TabIndex        =   7
         Top             =   300
         Width           =   660
      End
      Begin VB.Label Label2 
         Caption         =   "Bonus"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   240
         TabIndex        =   6
         Top             =   1320
         Width           =   615
      End
      Begin VB.Label Label1 
         Caption         =   "PAID"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   255
         TabIndex        =   5
         Top             =   300
         Width           =   495
      End
   End
   Begin VB.Label Label7 
      Alignment       =   2  'Center
      Caption         =   "Entry Actual Budget"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   555
      TabIndex        =   13
      Top             =   90
      Width           =   3690
   End
End
Attribute VB_Name = "Frm_MPActualBudgetUpdate"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim strMPMediumID As String
Dim intMonth As Integer
'
Dim MSC_Paid As Double
Dim MSC_Paid_On_Flag As Integer

Dim MSC_Bonus As Double
Dim MSC_Bonus_On_Flag As Integer

Dim Total_Actual As Double
Dim Actual_Gross_Paid As Double
Dim Actual_Nett_Paid As Double
Dim Actual_MSC_Paid As Double
Dim Actual_Gross_Bonus As Double
Dim Actual_Nett_Bonus As Double
Dim Actual_MSC_Bonus As Double

Dim strSql As String

Private Sub cmdSave_Click()
    'GET INPUT VALUE
    Actual_Gross_Paid = txtPaidGross.Text
    Actual_Nett_Paid = txtPaidNett.Text
    Actual_Gross_Bonus = txtBonusGross.Text
    Actual_Nett_Bonus = txtBonusNett.Text
    
    'KONFIRMASI JIKA 0
    If Actual_Gross_Paid = 0 And _
        Actual_Nett_Paid = 0 And _
      Actual_Gross_Bonus = 0 And _
       Actual_Nett_Bonus = 0 Then
       
        If MsgBox("You are about inserting zero value into actual budget." & vbCrLf & "are you sure?", vbYesNo, strApplication_Name) = vbNo Then
            Exit Sub
        End If
    End If
    
    'HITUNG FEE
    If MSC_Paid_On_Flag = 1 Then
        Actual_MSC_Paid = (MSC_Paid / 100) * Actual_Nett_Paid
    Else
        Actual_MSC_Paid = (MSC_Paid / 100) * Actual_Gross_Paid
    End If
    
    If MSC_Bonus_On_Flag = 1 Then
        Actual_MSC_Bonus = (MSC_Bonus / 100) * Actual_Nett_Bonus
    Else
        Actual_MSC_Bonus = (MSC_Bonus / 100) * Actual_Gross_Bonus
    End If
    
    Total_Actual = Actual_Nett_Paid + Actual_MSC_Paid + Actual_Nett_Bonus + Actual_MSC_Bonus
    
    Dim rsActual As New ADODB.Recordset
    strSql = "select * from mp_monthly_activity where mp_medium_id = '" & strMPMediumID & "' and month_number = " & intMonth
    rsActual.Open strSql, ConnERP, 1, 3
    If Not rsActual.EOF Then
        'Update Actual Budget
        rsActual("Total_Actual") = Total_Actual
        rsActual("Actual_Gross_Paid") = Actual_Gross_Paid
        rsActual("Actual_Nett_Paid") = Actual_Nett_Paid
        rsActual("Actual_MSC_Paid") = Actual_MSC_Paid
        rsActual("Actual_Gross_Bonus") = Actual_Gross_Bonus
        rsActual("Actual_Nett_Bonus") = Actual_Nett_Bonus
        rsActual("Actual_MSC_Bonus") = Actual_MSC_Bonus
        rsActual.Update
    End If
    rsActual.Close
    Set rsActual = Nothing
    
    '==============Update Nilai Actual yang tampil di MPInsertion=====================
    
    Dim Jumlah_Space As Integer 'Space untuk menghindari merge cell buland yang berbeda
    Dim i As Integer
    Dim strPlanBudget As String 'Jumlah space mencontek jumlah space di plan budget (cell diatasnya actual)
    Dim strActual_Old As String
    Dim strActualTotal_Old As String
    Dim strPLAN_Old As String
    Dim strPLANTotal_Old As String
    
    strActual_Old = FrmMPInsertion.FGMPInsertion.TextMatrix(FrmMPInsertion.FGMPInsertion.Row, FrmMPInsertion.FGMPInsertion.col)
    strActualTotal_Old = FrmMPInsertion.FGMPInsertion.TextMatrix(FrmMPInsertion.FGMPInsertion.Row, FrmMPInsertion.FGMPInsertion.cols - 2)
    
    strPLAN_Old = FrmMPInsertion.FGMPInsertion.TextMatrix(FrmMPInsertion.FGMPInsertion.Row - 1, FrmMPInsertion.FGMPInsertion.col)
    strPLANTotal_Old = FrmMPInsertion.FGMPInsertion.TextMatrix(FrmMPInsertion.FGMPInsertion.Row - 1, FrmMPInsertion.FGMPInsertion.cols - 2)
    
    strPlanBudget = FrmMPInsertion.FGMPInsertion.TextMatrix(FrmMPInsertion.FGMPInsertion.Row - 1, FrmMPInsertion.FGMPInsertion.col)
    Jumlah_Space = 0
    For i = 1 To Len(strPlanBudget)
        If Mid(strPlanBudget, i, 1) = " " Then
            Jumlah_Space = Jumlah_Space + 1
        Else
            Exit For
        End If
    Next
    
    'Cari week awal dari bulan yg akan diupdate
    i = FrmMPInsertion.FGMPInsertion.col
    While EngMonthIndex(FrmMPInsertion.FGMPInsertion.TextMatrix(1, i)) = intMonth
        i = i - 1
    Wend
    i = i + 1 'week awal
    
    'Print Actual Budget
    While EngMonthIndex(FrmMPInsertion.FGMPInsertion.TextMatrix(1, i)) = intMonth
        'Actual
        FrmMPInsertion.FGMPInsertion.TextMatrix(FrmMPInsertion.FGMPInsertion.Row, i) = String(Jumlah_Space, " ") & FormatNumber(Total_Actual, 2) & String(Jumlah_Space, " ")
        'Replace Plan with Actual
        FrmMPInsertion.FGMPInsertion.TextMatrix(FrmMPInsertion.FGMPInsertion.Row - 1, i) = String(Jumlah_Space, " ") & FormatNumber(Total_Actual, 2) & String(Jumlah_Space, " ")
        i = i + 1
    Wend
    
    'Print Sub Total Actual (per year)
    If Trim(strActualTotal_Old) <> "" Then
        If Trim(strActual_Old) <> "" Then
            FrmMPInsertion.FGMPInsertion.TextMatrix(FrmMPInsertion.FGMPInsertion.Row, FrmMPInsertion.FGMPInsertion.cols - 2) = FormatNumber(CDbl(RemoveNumberFormat(strActualTotal_Old)) - CDbl(RemoveNumberFormat(strActual_Old)) + Total_Actual, 2)
        Else
            FrmMPInsertion.FGMPInsertion.TextMatrix(FrmMPInsertion.FGMPInsertion.Row, FrmMPInsertion.FGMPInsertion.cols - 2) = FormatNumber(CDbl(RemoveNumberFormat(strActualTotal_Old)) + Total_Actual, 2)
        End If
    Else
        FrmMPInsertion.FGMPInsertion.TextMatrix(FrmMPInsertion.FGMPInsertion.Row, FrmMPInsertion.FGMPInsertion.cols - 2) = FormatNumber(Total_Actual, 2)
    End If
    
    'Refresh Sub Total Plan (Per year)
    If Trim(strPLANTotal_Old) <> "" Then
        If Trim(strPLAN_Old) <> "" Then
            FrmMPInsertion.FGMPInsertion.TextMatrix(FrmMPInsertion.FGMPInsertion.Row - 1, FrmMPInsertion.FGMPInsertion.cols - 2) = FormatNumber(CDbl(RemoveNumberFormat(strPLANTotal_Old)) - CDbl(RemoveNumberFormat(strPLAN_Old)) + Total_Actual, 2)
        Else
            FrmMPInsertion.FGMPInsertion.TextMatrix(FrmMPInsertion.FGMPInsertion.Row - 1, FrmMPInsertion.FGMPInsertion.cols - 2) = FormatNumber(CDbl(RemoveNumberFormat(strPLANTotal_Old)) + Total_Actual, 2)
        End If
    Else
        FrmMPInsertion.FGMPInsertion.TextMatrix(FrmMPInsertion.FGMPInsertion.Row - 1, FrmMPInsertion.FGMPInsertion.cols - 2) = FormatNumber(Total_Actual, 2)
    End If
    '=====================================================================================
    
    'Warnai yang udah actual
    Dim Start_Block As Double
    Dim End_Block As Double
    Dim Curr_Pos As Double
    Dim strMonth As String
    'Dim i As Integer
    
    With FrmMPInsertion.FGMPInsertion
        End_Block = .Row - 1
        Start_Block = End_Block - 1
        While UCase(Mid(.TextMatrix(Start_Block, 0), 1, 9)) <> "SUB TOTAL" And Start_Block > 4
            Start_Block = Start_Block - 1
        Wend
        Start_Block = Start_Block + 2
        
        strMonth = .TextMatrix(1, .col)
        Curr_Pos = .col
        While .TextMatrix(1, Curr_Pos) = strMonth
            Curr_Pos = Curr_Pos - 1
        Wend
        Curr_Pos = Curr_Pos + 1
            
        While .TextMatrix(1, Curr_Pos) = strMonth
            .col = Curr_Pos
            For i = Start_Block To End_Block
                .Row = i
                .CellBackColor = FrmMPInsertion.LegendActual.FillColor
            Next
            Curr_Pos = Curr_Pos + 1
        Wend
        
    End With
    
    Unload Me
End Sub

Private Sub Form_Load()
    Dim rsActual As New ADODB.Recordset
    strMPMediumID = FrmMPInsertion.FGMPInsertion.TextMatrix(FrmMPInsertion.FGMPInsertion.Row - 1, FrmMPInsertion.FGMPInsertion.cols - 1)
    intMonth = EngMonthIndex(FrmMPInsertion.FGMPInsertion.TextMatrix(1, FrmMPInsertion.FGMPInsertion.col))
    strSql = "select * from mp_monthly_activity where mp_medium_id = '" & strMPMediumID & "' and month_number = " & intMonth
    rsActual.Open strSql, ConnERP, 1, 3
    If Not rsActual.EOF Then
        txtPaidGross.Text = FormatNumber(IIf(IsNull(rsActual.Fields("Actual_Gross_Paid").Value), 0, rsActual.Fields("Actual_Gross_Paid").Value), 2)
        txtPaidNett.Text = FormatNumber(IIf(IsNull(rsActual.Fields("Actual_Nett_Paid").Value), 0, rsActual.Fields("Actual_Nett_Paid").Value), 2)
        txtBonusGross.Text = FormatNumber(IIf(IsNull(rsActual.Fields("Actual_Gross_Bonus").Value), 0, rsActual.Fields("Actual_Gross_Bonus").Value), 2)
        txtBonusNett.Text = FormatNumber(IIf(IsNull(rsActual.Fields("Actual_Nett_Bonus").Value), 0, rsActual.Fields("Actual_Nett_Bonus").Value), 2)
        
        MSC_Paid = rsActual.Fields("MSC_Paid").Value
        MSC_Paid_On_Flag = rsActual.Fields("MSC_Paid_On_Flag").Value
        MSC_Bonus = rsActual.Fields("MSC_Bonus").Value
        MSC_Bonus_On_Flag = rsActual.Fields("MSC_Bonus_On_Flag").Value
        
    End If
    rsActual.Close
    Set rsActual = Nothing
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub txtPaidGross_GotFocus()
    If IsNumeric(txtPaidGross.Text) Then
        txtPaidGross.Text = RemoveNumberFormat(txtPaidGross.Text)
    End If
End Sub

Private Sub txtPaidGross_LostFocus()
    If IsNumeric(txtPaidGross.Text) Then
       txtPaidGross.Text = FormatNumber(txtPaidGross.Text, 2)
    Else
        MsgBox "Invalid Entry!", vbExclamation, strApplication_Name
        txtPaidGross.SetFocus
    End If
End Sub

Private Sub txtPaidNett_GotFocus()
    If IsNumeric(txtPaidNett.Text) Then
        txtPaidNett.Text = RemoveNumberFormat(txtPaidNett.Text)
    End If
End Sub

Private Sub txtPaidNett_LostFocus()
    If IsNumeric(txtPaidNett.Text) Then
       txtPaidNett.Text = FormatNumber(txtPaidNett.Text, 2)
    Else
        MsgBox "Invalid Entry!", vbExclamation, strApplication_Name
        txtPaidNett.SetFocus
    End If
End Sub

Private Sub txtBonusGross_GotFocus()
    If IsNumeric(txtBonusGross.Text) Then
        txtBonusGross.Text = RemoveNumberFormat(txtBonusGross.Text)
    End If
End Sub

Private Sub txtBonusGross_LostFocus()
    If IsNumeric(txtBonusGross.Text) Then
       txtBonusGross.Text = FormatNumber(txtBonusGross.Text, 2)
    Else
        MsgBox "Invalid Entry!", vbExclamation, strApplication_Name
        txtBonusGross.SetFocus
    End If
End Sub

Private Sub txtBonusNett_GotFocus()
    If IsNumeric(txtBonusNett.Text) Then
        txtBonusNett.Text = RemoveNumberFormat(txtBonusNett.Text)
    End If
End Sub

Private Sub txtBonusNett_LostFocus()
    If IsNumeric(txtBonusNett.Text) Then
       txtBonusNett.Text = FormatNumber(txtBonusNett.Text, 2)
    Else
        MsgBox "Invalid Entry!", vbExclamation, strApplication_Name
        txtBonusNett.SetFocus
    End If
End Sub
