VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form Frm_Radio_MQ_Print 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Radio Media Quotation Print"
   ClientHeight    =   2370
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   4545
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2370
   ScaleWidth      =   4545
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Height          =   1605
      Left            =   120
      TabIndex        =   2
      Top             =   30
      Width           =   4305
      Begin Crystal.CrystalReport CR 
         Left            =   225
         Top             =   825
         _ExtentX        =   741
         _ExtentY        =   741
         _Version        =   348160
         PrintFileLinesPerPage=   60
      End
      Begin VB.ListBox Lst_JOB_ID 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1185
         ItemData        =   "Frm_MQ_Print.frx":0000
         Left            =   930
         List            =   "Frm_MQ_Print.frx":0002
         Style           =   1  'Checkbox
         TabIndex        =   3
         Top             =   255
         Width           =   3255
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Job ID :"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   120
         TabIndex        =   4
         Top             =   225
         Width           =   705
      End
   End
   Begin VB.CommandButton Cmd_Close 
      Caption         =   "&Close"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   3060
      TabIndex        =   1
      Top             =   1770
      Width           =   1125
   End
   Begin VB.CommandButton Cmd_Preview 
      Caption         =   "Pre&View"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   420
      TabIndex        =   0
      Top             =   1785
      Width           =   1125
   End
End
Attribute VB_Name = "Frm_Radio_MQ_Print"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
''*************************************************************
'Nama Form          : Frm_Radio_MQ_Print
'Fungsi Form        : memilih job id Quot Radio yang akan di print
'Programer          : joko
'cerated date       : 16/oct/01
'Last Update/By     :
'*************************************************************

Public What_Cbo_Brand  As ComboBox
Public What_CBO_IB As ComboBox
Public str_job_id As String

Private Sub Cbo_IB_Click()
If cbo_IB.ListIndex <> -1 Then
    Call Load_Job_ID
End If
End Sub

Private Sub Cmd_Close_Click()
    Unload Me
End Sub

Private Sub Cmd_Preview_Click()
    Call Job_Id
    Call Prepare_Print
End Sub



Private Sub Form_Load()
Dim LooP_Year As Integer
Dim Pos_Index As Integer

Rem Cbo_Year
'Cbo_Year.Clear
'For LooP_Year = 2001 To 2016
'    Cbo_Year.AddItem LooP_Year
'Next LooP_Year

Rem Cbo_Brand
'Cbo_Brand.Clear
'With What_Cbo_Brand
'    If .ListCount > 0 Then
'        For Pos_Index = 0 To .ListCount - 1
'            .listIndex = Pos_Index
'            Cbo_Brand.AddItem .Text
'        Next Pos_Index
'    End If
'End With

Rem Cbo_IB_ID
'Cbo_IB.Clear
'With What_CBO_IB
'    If .ListCount > 0 Then
'        For Pos_Index = 0 To .ListCount - 1
'            .listIndex = Pos_Index
'            Cbo_IB.AddItem .Text
'        Next Pos_Index
'    End If
'End With

Call Load_Job_ID
End Sub

Public Sub Load_Job_ID()
Dim TxtSQl As String
Dim rs As New ADODB.Recordset

TxtSQl = "Select job_id from ib_radio_quotation_detail "
TxtSQl = TxtSQl & " where ib_id = " & "'" & Trim(Frm_Radio_Media_Quot.Cbo_MQ.Text) & "'"
rs.Open TxtSQl, Conn, adOpenStatic, adLockReadOnly
With rs
    Lst_JOB_ID.Clear
    Do While .EOF = False
        Lst_JOB_ID.AddItem Trim(.Fields(0))
        .MoveNext
    Loop
End With

Set rs = Nothing
End Sub

Public Sub Job_Id()
Dim TxtSQl As String
Dim Pos_Index, Loop_Index As Integer

str_job_id = ""
With Lst_JOB_ID
    If .ListCount > 0 Then
        For Pos_Index = 0 To .ListCount - 1
            If .Selected(Pos_Index) = True Then
                Loop_Index = Loop_Index + 1
                If Loop_Index = 1 Then
                    .ListIndex = Pos_Index
                    str_job_id = str_job_id & "(IB_Radio_Quotation_Detail.Job_ID ='" & Trim(.Text) & "'"
                ElseIf Loop_Index > 1 Then
                    .ListIndex = Pos_Index
                    str_job_id = str_job_id & " or IB_Radio_Quotation_Detail.Job_ID ='" & Trim(.Text) & "'"
                End If
            End If
        Next Pos_Index
        str_job_id = str_job_id & ")"
        'Debug.Print Str_Job_Id
    End If
End With

'MsgBox Str_Job_Id
End Sub

Public Sub Prepare_Print()
    Dim TxtSQl As String
    Dim rs As New ADODB.Recordset

    TxtSQl = " select   Month_Catalog.Month_Name, Brand.Brand_Name, "
    TxtSQl = TxtSQl & " Client.Client_Name, "
    TxtSQl = TxtSQl & " Company.Company_Name, "
    TxtSQl = TxtSQl & " IB_Radio_Quot.Ib_Id as MediaPlanNo, "
    TxtSQl = TxtSQl & " IB_Radio_Quot.Ib_Id , "
    TxtSQl = TxtSQl & " IB_Radio_Quot.Remarks, "
    TxtSQl = TxtSQl & " IB_Radio_Quot.Date, "
    TxtSQl = TxtSQl & " IB_Radio_Quot.plan_no, "
    TxtSQl = TxtSQl & " IB_Radio_Quotation_Detail.Job_Id, "
    TxtSQl = TxtSQl & " IB_Radio_Quotation_Detail.Month, "
    TxtSQl = TxtSQl & " IB_Radio_Quotation_Detail.Year, "
    TxtSQl = TxtSQl & " IB_Radio_Quotation_Detail.Nett_Cost, "
    TxtSQl = TxtSQl & " IB_Radio_Quotation_Detail.Media_Sptv_Charge, "
    TxtSQl = TxtSQl & " IB_Radio_Quotation_Detail.Other_Charge, "
    TxtSQl = TxtSQl & " IB_Radio_Quotation_Detail.Bonus, "
    TxtSQl = TxtSQl & " IB_Radio_Quotation_Detail.Total_Lintas, "
    TxtSQl = TxtSQl & " IB_Radio_Quotation_Detail.Agency_Charge, "
    TxtSQl = TxtSQl & " IB_Radio_Quotation_Detail.Job_Number_Agency, "
    TxtSQl = TxtSQl & " IB_Radio_Quotation_Detail.Grand_Total "
    TxtSQl = TxtSQl & " from (ib_radio_quot inner join  IB_Radio_Quotation_Detail on IB_Radio_Quotation_Detail.ib_ID = IB_Radio_Quot.ib_ID ) "
    'TxtSQL = TxtSQL & " inner join IB_Radio on ib_radio.Ib_ID = Ib_Radio_quot.ib_id "
    TxtSQl = TxtSQl & " inner join Brand on left(IB_Radio_quot.ib_id,4) = brand.brand_code "
    TxtSQl = TxtSQl & " inner join client on brand.Client_Code = Client.Client_Code "
    TxtSQl = TxtSQl & " inner join company on Brand.Company_Code = Company.Company_Code "
    TxtSQl = TxtSQl & " inner join Month_Catalog on Month_Catalog.Month = ib_radio_quotation_detail.Month "
    TxtSQl = TxtSQl & " where ib_radio_Quot.Ib_ID ='" & Frm_Radio_Media_Quot.Cbo_MQ.Text & "'"
    TxtSQl = TxtSQl & " and  " & str_job_id
    TxtSQl = TxtSQl & " order by IB_Radio_Quotation_Detail.month asc"
    
    'Debug.Print TxtSQL
    rs.Open TxtSQl, Conn, adOpenStatic, adLockReadOnly
    
    Rem Filename .RPT
    CR.ReportFileName = Report_Dir & "\radio\mq_Radio.rpt"
    
    Rem Header Report
    With rs
        If .EOF = False Then
            CR.ParameterFields(39) = "Client;" & .Fields("Client_Name") & ";TRUE"
            CR.ParameterFields(1) = "Brand;" & .Fields("Brand_Name") & ";TRUE"
            CR.ParameterFields(2) = "MediaType;Radio;TRUE"
            CR.ParameterFields(3) = "MediaPlanNo;" & IIf(IsNull(.Fields("Plan_No")) = True, "", .Fields("Plan_No")) & ";TRUE"
            CR.ParameterFields(4) = "Dated;" & Format(CDate(.Fields("Date")), "mmm/dd/yyyy") & ";TRUE"
            CR.ParameterFields(5) = "IBID;" & .Fields("IB_ID") & ";TRUE"
            CR.ParameterFields(6) = "Remarks;" & .Fields("Remarks") & ";TRUE"
            CR.ParameterFields(7) = "PT;PT. INITIATIF MEDIA INDONESIA;TRUE"
            CR.ParameterFields(8) = "Marketing;Marketing Manager;TRUE"
        End If
        
        Rem 1st Month
        If .EOF = False Then
            CR.ParameterFields(9) = "MONTH1;" & .Fields("Month_Name") & ";TRUE"
            CR.ParameterFields(10) = "Year1;" & .Fields("Year") & ";TRUE"
            CR.ParameterFields(11) = "Nett1;" & .Fields("Nett_Cost") & ";TRUE"
            CR.ParameterFields(12) = "MSC1;" & .Fields("Media_Sptv_Charge") + .Fields("bonus") & ";TRUE"
            CR.ParameterFields(13) = "Other1;" & .Fields("Other_Charge") & ";TRUE"
            CR.ParameterFields(14) = "TotalLintas1;" & .Fields("Total_Lintas") & ";TRUE"
            CR.ParameterFields(15) = "JobNoAG1;" & .Fields("Job_Number_Agency") & ";TRUE"
            CR.ParameterFields(16) = "ClubCharge1;" & .Fields("Agency_Charge") & ";TRUE"
            CR.ParameterFields(17) = "GrandTotal1;" & .Fields("Grand_Total") & ";TRUE"
            CR.ParameterFields(36) = "JobID1;" & .Fields("Job_ID") & ";TRUE"
            .MoveNext
        Else
            CR.ParameterFields(9) = "MONTH1;;TRUE"
            CR.ParameterFields(10) = "Year1;;TRUE"
            CR.ParameterFields(11) = "Nett1;0;TRUE"
            CR.ParameterFields(12) = "MSC1;0;TRUE"
            CR.ParameterFields(13) = "Other1;0;TRUE"
            CR.ParameterFields(14) = "TotalLintas1;0;TRUE"
            CR.ParameterFields(15) = "JobNoAG1;;TRUE"
            CR.ParameterFields(16) = "clubcharge1;0;TRUE"
            CR.ParameterFields(17) = "GrandTotal1;0;TRUE"
            CR.ParameterFields(36) = "JobID1;;TRUE"
    
        End If
       
        Rem 2nd Month
        If .EOF = False Then
            CR.ParameterFields(18) = "MONTH2;" & .Fields("Month_Name") & ";TRUE"
            CR.ParameterFields(19) = "Year2;" & .Fields("Year") & ";TRUE"
            CR.ParameterFields(20) = "Nett2;" & .Fields("Nett_Cost") & ";TRUE"
            CR.ParameterFields(21) = "MSC2;" & .Fields("Media_Sptv_Charge") + .Fields("bonus") & ";TRUE"
            CR.ParameterFields(22) = "Other2;" & .Fields("Other_Charge") & ";TRUE"
            CR.ParameterFields(23) = "TotalLintas2;" & .Fields("Total_Lintas") & ";TRUE"
            CR.ParameterFields(24) = "JobNoAG2;" & .Fields("Job_Number_Agency") & ";TRUE"
            CR.ParameterFields(25) = "clubcharge2;" & .Fields("Agency_Charge") & ";TRUE"
            CR.ParameterFields(26) = "GrandTotal2;" & .Fields("Grand_Total") & ";TRUE"
            CR.ParameterFields(37) = "JobID2;" & .Fields("Job_ID") & ";TRUE"
            .MoveNext
        Else
            CR.ParameterFields(18) = "MONTH2;;TRUE"
            CR.ParameterFields(19) = "Year2;;TRUE"
            CR.ParameterFields(20) = "Nett2;0;TRUE"
            CR.ParameterFields(21) = "MSC2;0;TRUE"
            CR.ParameterFields(22) = "Other2;0;TRUE"
            CR.ParameterFields(23) = "TotalLintas2;0;TRUE"
            CR.ParameterFields(24) = "JobNoAG2;;TRUE"
            CR.ParameterFields(25) = "clubcharge2;0;TRUE"
            CR.ParameterFields(26) = "GrandTotal2;0;TRUE"
            CR.ParameterFields(37) = "JobID2;;TRUE"
            
        End If
        
        Rem 3rd Month
        If .EOF = False Then
            CR.ParameterFields(27) = "MONTH3;" & .Fields("Month_Name") & ";TRUE"
            CR.ParameterFields(28) = "Year3;" & .Fields("Year") & ";TRUE"
            CR.ParameterFields(29) = "Nett3;" & .Fields("Nett_Cost") & ";TRUE"
            CR.ParameterFields(30) = "MSC3;" & .Fields("Media_Sptv_Charge") + .Fields("bonus") & ";TRUE"
            CR.ParameterFields(31) = "Other3;" & .Fields("Other_Charge") & ";TRUE"
            CR.ParameterFields(32) = "TotalLintas3;" & .Fields("Total_Lintas") & ";TRUE"
            CR.ParameterFields(33) = "JobNoAG3;" & .Fields("Job_Number_Agency") & ";TRUE"
            CR.ParameterFields(34) = "clubcharge3;" & .Fields("Agency_Charge") & ";TRUE"
            CR.ParameterFields(35) = "GrandTotal3;" & .Fields("Grand_Total") & ";TRUE"
            CR.ParameterFields(38) = "JobID3;" & .Fields("Job_ID") & ";TRUE"
            .MoveNext
        Else
            CR.ParameterFields(27) = "MONTH3;;TRUE"
            CR.ParameterFields(28) = "Year3;;TRUE"
            CR.ParameterFields(29) = "Nett3;0;TRUE"
            CR.ParameterFields(30) = "MSC3;0;TRUE"
            CR.ParameterFields(31) = "Other3;0;TRUE"
            CR.ParameterFields(32) = "TotalLintas3;0;TRUE"
            CR.ParameterFields(33) = "JobNoAG3;;TRUE"
            CR.ParameterFields(34) = "clubcharge3;0;TRUE"
            CR.ParameterFields(35) = "GrandTotal3;0;TRUE"
            CR.ParameterFields(38) = "JobID3;;TRUE"
        End If
    End With
        CR.WindowState = crptMaximized
        CR.WindowShowRefreshBtn = True
        CR.WindowShowPrintSetupBtn = True
        CR.WindowTitle = " -- Implementation Brief Radio Quotation -- "
    CR.Connect = "DSN = " & Server_Name & ";UID = " & Login_User & ";PWD = " & Login_Password & ";DSQ = " & Database_Name
    CR.Action = 1

End Sub
