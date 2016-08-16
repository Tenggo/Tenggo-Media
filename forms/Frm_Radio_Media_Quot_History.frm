VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form Frm_Radio_Media_Quot_History 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Radio Media Quotation Revision"
   ClientHeight    =   8025
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9645
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "Frm_Radio_Media_Quot_History.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8025
   ScaleWidth      =   9645
   StartUpPosition =   1  'CenterOwner
   Begin Crystal.CrystalReport Crpt 
      Left            =   3420
      Top             =   2655
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      PrintFileLinesPerPage=   60
   End
   Begin VB.PictureBox Picture1 
      Height          =   8010
      Left            =   0
      ScaleHeight     =   7950
      ScaleWidth      =   9570
      TabIndex        =   0
      Top             =   -15
      Width           =   9630
      Begin VB.Frame Fra_Approval 
         Caption         =   "Client Approval"
         ForeColor       =   &H000000FF&
         Height          =   1230
         Left            =   6330
         TabIndex        =   20
         ToolTipText     =   "Double Click to Approve"
         Top             =   5790
         Width           =   3150
         Begin VB.Label Lbl_Approval_Date 
            Alignment       =   2  'Center
            ForeColor       =   &H000000FF&
            Height          =   300
            Left            =   210
            TabIndex        =   22
            Top             =   765
            Width           =   2700
         End
         Begin VB.Label Lbl_Approval_Status 
            Alignment       =   2  'Center
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   11.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   270
            Left            =   285
            TabIndex        =   21
            Top             =   390
            Width           =   2700
         End
      End
      Begin VB.Frame Frame3 
         Caption         =   "Remark"
         Height          =   1230
         Left            =   105
         TabIndex        =   18
         Top             =   5790
         Width           =   4995
         Begin VB.Label Lbl_Remarks 
            BackColor       =   &H00FFFFC0&
            BorderStyle     =   1  'Fixed Single
            Height          =   840
            Left            =   165
            TabIndex        =   19
            Top             =   255
            Width           =   4665
         End
      End
      Begin VB.Frame Frame1 
         Height          =   1530
         Left            =   105
         TabIndex        =   7
         Top             =   0
         Width           =   9390
         Begin VB.ComboBox Cbo_MQ_NO 
            Height          =   315
            Left            =   1590
            Style           =   2  'Dropdown List
            TabIndex        =   12
            Top             =   990
            Width           =   1770
         End
         Begin VB.ComboBox Cbo_Year 
            Height          =   315
            Left            =   1605
            Style           =   2  'Dropdown List
            TabIndex        =   9
            Top             =   645
            Width           =   1530
         End
         Begin VB.ComboBox Cbo_Brand 
            Height          =   315
            Left            =   1605
            Style           =   2  'Dropdown List
            TabIndex        =   8
            Top             =   300
            Width           =   3315
         End
         Begin VB.Label Label4 
            Alignment       =   1  'Right Justify
            Caption         =   "Plan No :"
            Height          =   225
            Left            =   5655
            TabIndex        =   17
            Top             =   360
            Width           =   930
         End
         Begin VB.Label Label5 
            Alignment       =   1  'Right Justify
            Caption         =   "Plan Date :"
            Height          =   225
            Left            =   5670
            TabIndex        =   16
            Top             =   690
            Width           =   930
         End
         Begin VB.Label Lbl_Plan_No 
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   1  'Fixed Single
            Height          =   315
            Left            =   6660
            TabIndex        =   15
            Top             =   315
            Width           =   1875
         End
         Begin VB.Label Lbl_Plan_date 
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   1  'Fixed Single
            Height          =   315
            Left            =   6660
            TabIndex        =   14
            Top             =   660
            Width           =   1875
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            Caption         =   "MQ No :"
            Height          =   210
            Left            =   840
            TabIndex        =   13
            Top             =   1035
            Width           =   705
         End
         Begin VB.Label Label2 
            Alignment       =   1  'Right Justify
            Caption         =   "&Year :"
            Height          =   210
            Left            =   840
            TabIndex        =   11
            Top             =   675
            Width           =   705
         End
         Begin VB.Label Label3 
            Alignment       =   1  'Right Justify
            Caption         =   "&Brand :"
            Height          =   225
            Left            =   900
            TabIndex        =   10
            Top             =   345
            Width           =   645
         End
      End
      Begin VB.Frame Frame2 
         Height          =   4185
         Left            =   90
         TabIndex        =   5
         Top             =   1530
         Width           =   9390
         Begin MSHierarchicalFlexGridLib.MSHFlexGrid Flex_Quot 
            Height          =   3705
            Left            =   300
            TabIndex        =   6
            Top             =   285
            Width           =   8865
            _ExtentX        =   15637
            _ExtentY        =   6535
            _Version        =   393216
            BackColor       =   16777215
            Rows            =   12
            Cols            =   6
            FixedRows       =   0
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
            _NumberOfBands  =   1
            _Band(0).Cols   =   6
         End
      End
      Begin VB.Frame Frame9 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   825
         Left            =   8235
         TabIndex        =   3
         Top             =   7020
         Width           =   1245
         Begin VB.CommandButton Cmd_Close 
            Cancel          =   -1  'True
            Caption         =   "C&lose"
            Height          =   540
            Left            =   135
            Picture         =   "Frm_Radio_Media_Quot_History.frx":0442
            TabIndex        =   4
            Top             =   195
            Width           =   999
         End
      End
      Begin VB.Frame Frame8 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   825
         Left            =   6345
         TabIndex        =   1
         Top             =   7020
         Width           =   1200
         Begin VB.CommandButton Cmd_Print 
            Caption         =   "&Print"
            Height          =   525
            Left            =   150
            Picture         =   "Frm_Radio_Media_Quot_History.frx":074C
            TabIndex        =   2
            Top             =   210
            Width           =   915
         End
      End
   End
End
Attribute VB_Name = "Frm_Radio_Media_Quot_History"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
''*************************************************************
'Nama Form          : Frm_Radio_Media_Quot_History
'Fungsi Form        : mbrowse data history Quotation Radio
'Programer          : joko
'created date       : 16/apr/01
'Last Update/By     :
'*************************************************************
Option Explicit
Dim Index_Row As Integer
Dim Index_Col As Integer
Dim str_month As String
Dim Month_MQ As Integer
Dim Year_MQ As Integer
Dim strSql As String
Dim Rs_View_MQ_Approve As New ADODB.Recordset

Private Sub Initial_Grd()
        Flex_Quot.Height = 0
        For Index_Row = 0 To Flex_Quot.Rows - 1
            Flex_Quot.RowHeight(Index_Row) = 290
            Flex_Quot.Height = Flex_Quot.Height + 310
        Next Index_Row
        Flex_Quot.ColWidth(0) = 2800
        Flex_Quot.cols = 24
        For Index_Col = 1 To 23
            If Index_Col Mod 2 = 0 Then
                 Flex_Quot.ColWidth(Index_Col) = 400
            Else
                 Flex_Quot.ColWidth(Index_Col) = 1600
            End If
        Next
        
        
        
        
        
        For Index_Col = 1 To Flex_Quot.cols - 1
            Flex_Quot.col = Index_Col
            Flex_Quot.Row = 0
            Flex_Quot.CellFontBold = True
            
            Flex_Quot.col = Index_Col
            Flex_Quot.Row = 1
            Flex_Quot.CellFontBold = True
            Flex_Quot.CellBackColor = &HFFFFC0
            
            'Flex_Quot.Col = index_col
'            Flex_Quot.Row = 9
'            Flex_Quot.CellBackColor = &H8000000F
'            Flex_Quot.Col = index_col
            
            Flex_Quot.Row = 10
            Flex_Quot.CellFontBold = True
            Flex_Quot.CellBackColor = &HFFFFC0
            Flex_Quot.CellForeColor = vbRed
        Next Index_Col
        For Index_Col = 2 To 22 Step 2
            For Index_Row = 0 To Flex_Quot.Rows - 1
                Flex_Quot.Row = Index_Row
                Flex_Quot.col = Index_Col
                Flex_Quot.CellBackColor = &H8000000F
            Next Index_Row
        Next
        For Index_Row = 0 To Flex_Quot.Rows - 1
            Flex_Quot.col = 0
            Flex_Quot.Row = Index_Row
            Flex_Quot.CellFontBold = True
        Next
        Flex_Quot.TextMatrix(0, 0) = "Month"
        Flex_Quot.TextMatrix(1, 0) = "Job Id"
        Flex_Quot.TextMatrix(2, 0) = "Nett Cost"
        Flex_Quot.TextMatrix(3, 0) = "Media Supervition Charges"
        Flex_Quot.TextMatrix(4, 0) = "Bonus Fee"
        Flex_Quot.TextMatrix(5, 0) = "Others"
        Flex_Quot.TextMatrix(6, 0) = "Total Lintas"
        Flex_Quot.TextMatrix(7, 0) = "Job Number Club Agency"
        Flex_Quot.TextMatrix(8, 0) = "Club Agency Media Sptv. Charges"
        Flex_Quot.TextMatrix(9, 0) = "Grand Total"
        Flex_Quot.TextMatrix(11, 0) = "Revision"
End Sub



Private Sub cbo_brand_Click()
    Cbo_MQ_NO.Clear
    Clear_Form
    
    If Cbo_Brand.Text <> "" Then
        Cbo_Year.Clear
        Dim Th As Integer
        For Th = 2002 To 2016
          Cbo_Year.AddItem Th
        Next
    End If
End Sub

Private Sub Load_Brand()
Dim rs_brand As New ADODB.Recordset
    Cbo_Brand.Clear
    With rs_brand
        
        strSql = "SELECT a.* FROM brand a, client b WHERE brand_code IN (SELECT brand_code FROM Media_Security_Catalog WHERE User_name='" & strLogin_User & "' AND (position='Implementor' or position='Planner' or position='Buyer' or position='Supervisor' or position='Administrator' or position='Admin' ) and Valid_until > getdate()) and a.client_code=b.client_code"
        .Open strSql, ConnERP, adOpenForwardOnly, adLockReadOnly
        While Not .EOF And Not .BOF
            Cbo_Brand.AddItem .Fields("Brand_Code").Value & " --> " & .Fields("Brand_Name").Value
            .MoveNext
        Wend
            
    End With
    rs_brand.Close
    Set rs_brand = Nothing
End Sub

Private Sub Cbo_MQ_NO_Click()
    Clear_Form
    If Cbo_MQ_NO.Text <> "" Then
      show_data
    End If
End Sub

Private Sub Load_MQ_NO()
    Dim rs_Load_MQ_No As New ADODB.Recordset
    strSql = "SELECT DISTINCT ib_id FROM IB_Radio_Quotation_Detail_Revision WHERE left(job_id,4)='" & Left(Cbo_Brand.Text, 4) & "' AND year= " & CInt(Cbo_Year.Text)
    
    rs_Load_MQ_No.Open strSql, ConnERP, adOpenStatic, adLockReadOnly
    Cbo_MQ_NO.Clear
   ' Cbo_MQ_NO.AddItem "-- ALL --"
    While Not rs_Load_MQ_No.EOF And Not rs_Load_MQ_No.BOF
      Cbo_MQ_NO.AddItem rs_Load_MQ_No(0)
      rs_Load_MQ_No.MoveNext
    Wend
    rs_Load_MQ_No.Close
    Set rs_Load_MQ_No = Nothing
End Sub

Private Sub Cbo_Year_Click()
  
  Cbo_MQ_NO.Clear
  Clear_Form
  
  If Cbo_Year.Text <> "" Then
      Load_MQ_NO
  End If
End Sub

Private Sub Cmd_Close_Click()
    Unload Me
End Sub

Private Sub Cmd_Print_Click()

Dim MSC_Bonus As String
If Cbo_Brand.Text = "" Then
    MsgBox "Select Brand First", vbExclamation, StrCompany
    Exit Sub
End If

If Cbo_Year.Text = "" Then
    MsgBox "Select Year First", vbExclamation, StrCompany
    Exit Sub
End If

If Cbo_MQ_NO.Text = "" Then
    MsgBox "Select MQ No First", vbExclamation, StrCompany
    Exit Sub
End If
Me.MousePointer = vbHourglass

With Crpt
    .Reset
    '.ReportFileName = App.Path + "\report\view_hist_radio.rpt"
    .ReportFileName = Report_Dir + "\Radio\view_hist_radio.rpt"
    'header
    Dim rs_Load_Name_Brand As New ADODB.Recordset
    strSql = "select brand_name from brand where brand_code='" & Left(Cbo_Brand.Text, 4) & "'"
    rs_Load_Name_Brand.Open strSql, ConnERP, adOpenStatic, adLockReadOnly
    
    '.SelectionFormula = "{ib_print_quotation.ib_id}='" & Cbo_MQ_NO.Text & "'"
    .Formulas(0) = "Brand ='" & rs_Load_Name_Brand(0) & "'"
    rs_Load_Name_Brand.Close
    Set rs_Load_Name_Brand = Nothing
    
    strSql = "select client_name from client a,brand  b where b.brand_code='" & Left(Cbo_Brand.Text, 4) & "' and a.client_code=b.client_code "
    rs_Load_Name_Brand.Open strSql, ConnERP, adOpenStatic, adLockReadOnly
    
    .SelectionFormula = "{ib_print_quotation.ib_id}='" & Cbo_MQ_NO.Text & "'"
    .Formulas(30) = "Client ='" & rs_Load_Name_Brand(0) & "'"
    rs_Load_Name_Brand.Close
    Set rs_Load_Name_Brand = Nothing
    
    
    .Formulas(1) = "Year ='" & Cbo_Year.Text & "'"
    .Formulas(2) = "ib_id ='" & Cbo_MQ_NO.Text & "'"
    
    
    
    MSC_Bonus = Format(CDbl(IIf(Flex_Quot.TextMatrix(3, 1) = "", 0, Flex_Quot.TextMatrix(3, 1))) + CDbl(IIf(Flex_Quot.TextMatrix(4, 1) = "", 0, Flex_Quot.TextMatrix(4, 1))), "#,##0")
    If MSC_Bonus = 0 Then MSC_Bonus = ""
    'detail 1
    .Formulas(3) = "month_1 ='" & Flex_Quot.TextMatrix(0, 1) & " " & IIf(Flex_Quot.TextMatrix(0, 1) = "", "", Cbo_Year.Text) & "'"
    .Formulas(4) = "job_no_1 ='" & Flex_Quot.TextMatrix(1, 1) & "'"
    .Formulas(5) = "nett_cost_1 ='" & Flex_Quot.TextMatrix(2, 1) & "'"
    .Formulas(6) = "msc_1 ='" & MSC_Bonus & "'"
    .Formulas(7) = "Other_1 ='" & Flex_Quot.TextMatrix(5, 1) & "'"
    .Formulas(8) = "Total_Lintas_1 ='" & Flex_Quot.TextMatrix(6, 1) & "'"
    .Formulas(9) = "Job_No_Club_1 ='" & Flex_Quot.TextMatrix(7, 1) & "'"
    .Formulas(10) = "Msc_Club_1 ='" & Flex_Quot.TextMatrix(8, 1) & "'"
    .Formulas(11) = "Grand_Total_1 ='" & Flex_Quot.TextMatrix(9, 1) & "'"
    
    
    MSC_Bonus = Format(CDbl(IIf(Flex_Quot.TextMatrix(3, 3) = "", 0, Flex_Quot.TextMatrix(3, 3))) + CDbl(IIf(Flex_Quot.TextMatrix(4, 3) = "", 0, Flex_Quot.TextMatrix(4, 3))), "#,##0")
    If MSC_Bonus = 0 Then MSC_Bonus = ""
    'detail 2
    .Formulas(12) = "month_2 ='" & Flex_Quot.TextMatrix(0, 3) & " " & IIf(Flex_Quot.TextMatrix(0, 3) = "", "", Cbo_Year.Text) & "'"
    .Formulas(13) = "job_no_2 ='" & Flex_Quot.TextMatrix(1, 3) & "'"
    .Formulas(14) = "nett_cost_2 ='" & Flex_Quot.TextMatrix(2, 3) & "'"
    .Formulas(15) = "msc_2 ='" & MSC_Bonus & "'"
    .Formulas(16) = "Other_2 ='" & Flex_Quot.TextMatrix(5, 3) & "'"
    .Formulas(17) = "Total_Lintas_2 ='" & Flex_Quot.TextMatrix(6, 3) & "'"
    .Formulas(18) = "Job_No_Club_2 ='" & Flex_Quot.TextMatrix(7, 3) & "'"
    .Formulas(19) = "Msc_Club_2 ='" & Flex_Quot.TextMatrix(8, 3) & "'"
    .Formulas(20) = "Grand_Total_2 ='" & Flex_Quot.TextMatrix(9, 3) & "'"
    
    MSC_Bonus = Format(CDbl(IIf(Flex_Quot.TextMatrix(3, 5) = "", 0, Flex_Quot.TextMatrix(3, 5))) + CDbl(IIf(Flex_Quot.TextMatrix(4, 5) = "", 0, Flex_Quot.TextMatrix(4, 5))), "#,##0")
    If MSC_Bonus = 0 Then MSC_Bonus = ""
    'detail 3
    .Formulas(21) = "month_3 ='" & Flex_Quot.TextMatrix(0, 5) & " " & IIf(Flex_Quot.TextMatrix(0, 5) = "", "", Cbo_Year.Text) & "'"
    .Formulas(22) = "job_no_3 ='" & Flex_Quot.TextMatrix(1, 5) & "'"
    .Formulas(23) = "nett_cost_3 ='" & Flex_Quot.TextMatrix(2, 5) & "'"
    .Formulas(24) = "msc_3 ='" & MSC_Bonus & "'"
    .Formulas(25) = "Other_3 ='" & Flex_Quot.TextMatrix(5, 5) & "'"
    .Formulas(26) = "Total_Lintas_3 ='" & Flex_Quot.TextMatrix(6, 5) & "'"
    .Formulas(27) = "Job_No_Club_3 ='" & Flex_Quot.TextMatrix(7, 5) & "'"
    .Formulas(28) = "Msc_Club_3 ='" & Flex_Quot.TextMatrix(8, 5) & "'"
    .Formulas(29) = "Grand_Total_3 ='" & Flex_Quot.TextMatrix(9, 5) & "'"
    
    .WindowShowPrintSetupBtn = True
    .WindowShowRefreshBtn = True
    .WindowState = crptMaximized
    .WindowTitle = "-- View MQ Radio Revision --"
    .Connect = "DSN=" & Server_Name & ";UID=" & Login_User & ";PWD=" & Login_Password & ";DSQ=" & Database_Name & ""
    .Action = 1
End With
Me.MousePointer = vbDefault
End Sub

Private Sub Form_Load()
'center Form
    
    RemoveMenus Me, True
'Brand
    
'Load Month
        
        
        
        
        
'load Year
    
    Load_Brand
    
'***********************************
'       Initial Grid
'***********************************
        Initial_Grd
End Sub

Private Sub show_data()

Dim baris As Integer

strSql = "SELECT * FROM ib_radio_quot WHERE ib_id='" & Cbo_MQ_NO.Text & "'"

Rs_View_MQ_Approve.Open strSql, ConnERP, adOpenStatic, adLockReadOnly

If Not Rs_View_MQ_Approve.EOF Then
    Lbl_Plan_No.Caption = Rs_View_MQ_Approve("plan_no")
    Lbl_Plan_date.Caption = Format(Rs_View_MQ_Approve("date"), "mmm dd, yyyy")
    
    Lbl_Remarks.Caption = IIf(IsNull(Rs_View_MQ_Approve("Remarks").Value), "", Rs_View_MQ_Approve("Remarks").Value)
    
    If Rs_View_MQ_Approve("Approval_client") = 1 Then
        Lbl_Approval_Status.Caption = "Approved"
        Lbl_Approval_Date.Caption = Format(Rs_View_MQ_Approve("approved_date"), "dd mmm yy hh:mm:ss AMPM")
        Lbl_Approval_Status.ForeColor = vbBlack
        Lbl_Approval_Date.ForeColor = vbBlack
    Else
        Lbl_Approval_Status.Caption = "UnApproved"
        Lbl_Approval_Date.Caption = ""
        Lbl_Approval_Status.ForeColor = vbRed
        Lbl_Approval_Date.ForeColor = vbRed
    End If
End If
    
Rs_View_MQ_Approve.Close
Set Rs_View_MQ_Approve = Nothing


strSql = " SELECT a.* FROM "
strSql = strSql & " (SELECT IB_ID, Job_Id, Month, Year, Nett_Cost, Media_Sptv_Charge, Bonus, Other_Charge, Total_Lintas, Agency_Charge, Job_Number_Agency, Grand_Total, Source_IB,0 revision "
strSql = strSql & " FROM IB_Radio_Quotation_Detail WHERE"
strSql = strSql & " year=" & CInt(Cbo_Year.Text) & " and ib_id='" & Me.Cbo_MQ_NO.Text & "'"
strSql = strSql & " Union"
strSql = strSql & " SELECT IB_ID, Job_Id, Month, Year, Nett_Cost, Media_Sptv_Charge, Bonus, Other_Charge, Total_Lintas, Agency_Charge, Job_Number_Agency, Grand_Total, Source_IB, revision"
strSql = strSql & " FROM IB_Radio_Quotation_Detail_revision WHERE"
strSql = strSql & " year=" & CInt(Cbo_Year.Text) & " and ib_id='" & Me.Cbo_MQ_NO.Text & "') a"
strSql = strSql & " order by a.month"

Rs_View_MQ_Approve.CursorLocation = adUseClient
Rs_View_MQ_Approve.Open strSql, ConnERP, adOpenStatic, adLockReadOnly

Index_Col = 1

While Not Rs_View_MQ_Approve.EOF
    
    Flex_Quot.TextMatrix(0, Index_Col) = Get_Month_Name(Rs_View_MQ_Approve.Fields("Month").Value)
    Flex_Quot.TextMatrix(1, Index_Col) = Rs_View_MQ_Approve.Fields("job_Id").Value
    Flex_Quot.TextMatrix(2, Index_Col) = Format(Rs_View_MQ_Approve.Fields("Nett_Cost").Value, "##,##0")
    Flex_Quot.TextMatrix(3, Index_Col) = Format(Rs_View_MQ_Approve.Fields("Media_Sptv_Charge").Value, "##,##0")
    Flex_Quot.TextMatrix(4, Index_Col) = Format(Rs_View_MQ_Approve.Fields("Bonus").Value, "##,##0")
    Flex_Quot.TextMatrix(5, Index_Col) = Format(Rs_View_MQ_Approve.Fields("Other_Charge").Value, "##,##0")
    Flex_Quot.TextMatrix(6, Index_Col) = Format(Rs_View_MQ_Approve.Fields("Total_Lintas").Value, "##,##0")
    Flex_Quot.TextMatrix(7, Index_Col) = Rs_View_MQ_Approve.Fields("Job_Number_Agency").Value
    Flex_Quot.TextMatrix(8, Index_Col) = Format(Rs_View_MQ_Approve.Fields("Agency_Charge").Value, "##,##0")
    Flex_Quot.TextMatrix(9, Index_Col) = Format(Rs_View_MQ_Approve.Fields("Grand_Total").Value, "##,##0")
    Flex_Quot.TextMatrix(11, Index_Col) = IIf(Rs_View_MQ_Approve.Fields("revision") = 0, "Actual", Rs_View_MQ_Approve.Fields("revision"))
    
    'give color for actual MQ
    If Rs_View_MQ_Approve.Fields("revision") = 0 Then
        For baris = 0 To 11
                Select Case baris
                Case 1, 10:
                Case Else:  Flex_Quot.Row = baris
                            Flex_Quot.col = Index_Col
                            Flex_Quot.CellBackColor = &HC0FFC0
                End Select
        Next
    End If
    Index_Col = Index_Col + 2
    Rs_View_MQ_Approve.MoveNext
Wend
Rs_View_MQ_Approve.Close
Set Rs_View_MQ_Approve = Nothing

'strSql = "SELECT * FROM IB_Radio_Quotation_Detail WHERE left(job_id,4)='" & Left(Cbo_Brand.Text, 4) & "' and year=" & CInt(Cbo_Year.Text) & " and ib_id='" & Me.Cbo_MQ_NO.Text & "'"
'strSql = strSql & "order by month"
'
'Rs_View_MQ_Approve.CursorLocation = adUseClient
'Rs_View_MQ_Approve.Open strSql, ConnERP , adOpenStatic, adLockReadOnly
'
'Index_Col = 1
'
'While Not Rs_View_MQ_Approve.EOF And Not Rs_View_MQ_Approve.BOF
'            Flex_Quot.TextMatrix(0, Index_Col) = Get_Month_Name(Rs_View_MQ_Approve.Fields("Month").Value)
'            'Flex_Quot.TextMatrix(1, Index_Col) = Rs_View_MQ_Approve.Fields("ib_Id").Value
'            Flex_Quot.TextMatrix(1, Index_Col) = Rs_View_MQ_Approve.Fields("job_Id").Value
'            Flex_Quot.TextMatrix(2, Index_Col) = Format(Rs_View_MQ_Approve.Fields("Nett_Cost").Value, "##,##0")
'            Flex_Quot.TextMatrix(3, Index_Col) = Format(Rs_View_MQ_Approve.Fields("Media_Sptv_Charge").Value, "##,##0")
'            Flex_Quot.TextMatrix(4, Index_Col) = Format(Rs_View_MQ_Approve.Fields("Bonus").Value, "##,##0")
'            Flex_Quot.TextMatrix(5, Index_Col) = Format(Rs_View_MQ_Approve.Fields("Other_Charge").Value, "##,##0")
'            Flex_Quot.TextMatrix(6, Index_Col) = Format(Rs_View_MQ_Approve.Fields("Total_Lintas").Value, "##,##0")
'            Flex_Quot.TextMatrix(7, Index_Col) = Rs_View_MQ_Approve.Fields("Job_Number_Agency").Value
'            Flex_Quot.TextMatrix(8, Index_Col) = Format(Rs_View_MQ_Approve.Fields("Agency_Charge").Value, "##,##0")
'            Flex_Quot.TextMatrix(9, Index_Col) = Format(Rs_View_MQ_Approve.Fields("Grand_Total").Value, "##,##0")
'            Flex_Quot.TextMatrix(11, Index_Col) = "Actual"
'    Rs_View_MQ_Approve.MoveNext
'    'give color for actual MQ
'    For baris = 0 To 11
'            Select Case baris
'            Case 1, 10:
'            Case Else:  Flex_Quot.Row = baris
'                        Flex_Quot.Col = Index_Col
'                        Flex_Quot.CellBackColor = &HC0FFC0
'            End Select
'    Next
'    Index_Col = Index_Col + 2
'Wend
'Rs_View_MQ_Approve.Close
'Set Rs_View_MQ_Approve = Nothing
'
'
'strSql = "SELECT * FROM IB_Radio_Quotation_Detail_Revision WHERE left(job_id,4)='" & Left(Cbo_Brand.Text, 4) & "' and year=" & CInt(Cbo_Year.Text) & " and ib_id='" & Me.Cbo_MQ_NO.Text & "'"
'strSql = strSql & " ORDER BY month"
'
'Rs_View_MQ_Approve.CursorLocation = adUseClient
'Rs_View_MQ_Approve.Open strSql, ConnERP , adOpenStatic, adLockReadOnly
'
'
'
'While Not Rs_View_MQ_Approve.EOF And Not Rs_View_MQ_Approve.BOF
'            Flex_Quot.TextMatrix(0, Index_Col) = Get_Month_Name(Rs_View_MQ_Approve.Fields("Month").Value)
'            'Flex_Quot.TextMatrix(1, Index_Col) = Rs_View_MQ_Approve.Fields("ib_Id").Value
'            Flex_Quot.TextMatrix(1, Index_Col) = Rs_View_MQ_Approve.Fields("job_Id").Value
'            Flex_Quot.TextMatrix(2, Index_Col) = Format(Rs_View_MQ_Approve.Fields("Nett_Cost").Value, "##,##0")
'            Flex_Quot.TextMatrix(3, Index_Col) = Format(Rs_View_MQ_Approve.Fields("Media_Sptv_Charge").Value, "##,##0")
'            Flex_Quot.TextMatrix(4, Index_Col) = Format(Rs_View_MQ_Approve.Fields("Bonus").Value, "##,##0")
'            Flex_Quot.TextMatrix(5, Index_Col) = Format(Rs_View_MQ_Approve.Fields("Other_Charge").Value, "##,##0")
'            Flex_Quot.TextMatrix(6, Index_Col) = Format(Rs_View_MQ_Approve.Fields("Total_Lintas").Value, "##,##0")
'            Flex_Quot.TextMatrix(7, Index_Col) = Rs_View_MQ_Approve.Fields("Job_Number_Agency").Value
'            Flex_Quot.TextMatrix(8, Index_Col) = Format(Rs_View_MQ_Approve.Fields("Agency_Charge").Value, "##,##0")
'            Flex_Quot.TextMatrix(9, Index_Col) = Format(Rs_View_MQ_Approve.Fields("Grand_Total").Value, "##,##0")
'            Flex_Quot.TextMatrix(11, Index_Col) = Rs_View_MQ_Approve.Fields("revision")
'    Rs_View_MQ_Approve.MoveNext
'
'    Index_Col = Index_Col + 2
'Wend
'Rs_View_MQ_Approve.Close
'Set Rs_View_MQ_Approve = Nothing


End Sub

Public Sub Clear_Form()
    For Index_Col = 1 To Flex_Quot.cols - 1 Step 2
        For Index_Row = 0 To Flex_Quot.Rows - 1
            If Index_Row <> 1 And Index_Row <> 10 Then
                Flex_Quot.Row = Index_Row
                Flex_Quot.col = Index_Col
                Flex_Quot.CellBackColor = &H8000000E
            End If
            Flex_Quot.TextMatrix(Index_Row, Index_Col) = ""
        Next Index_Row
    Next Index_Col
    
    Me.Lbl_Plan_date = ""
    Me.Lbl_Plan_No = ""
    Me.Lbl_Remarks = ""
    Lbl_Approval_Status.Caption = ""
    Lbl_Approval_Date.Caption = ""
End Sub

Private Sub Lbl_Approval_Status_dblClick()
    'Check apakah Buyer Brand
    If Not IsValidAccess(strLogin_User, "Implementor", Left(Cbo_Brand.Text, 4)) Then
        MsgBox "Access Denied...", vbCritical, StrCompany
        Exit Sub
    End If
    
    If Lbl_Approval_Status.Caption = "UnApproved" Then
        frm_Approve_Media_Quotation_Revision.Modul = "Radio"
        
        frm_Approve_Media_Quotation_Revision.show 1
    End If
End Sub
