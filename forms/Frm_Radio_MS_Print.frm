VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form Frm_Radio_MS_Print 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Value Option"
   ClientHeight    =   3000
   ClientLeft      =   1605
   ClientTop       =   1665
   ClientWidth     =   2835
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   9
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3000
   ScaleWidth      =   2835
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame fraValue 
      Appearance      =   0  'Flat
      Caption         =   "Value From"
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
      Height          =   1575
      Left            =   120
      TabIndex        =   3
      Top             =   120
      Width           =   2535
      Begin VB.OptionButton optActual 
         Caption         =   "Actual"
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
         Left            =   300
         TabIndex        =   8
         Top             =   1155
         Width           =   1140
      End
      Begin VB.OptionButton optSchedule 
         Caption         =   "Schedule"
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
         Left            =   300
         TabIndex        =   7
         Top             =   270
         Width           =   1140
      End
      Begin VB.Frame fraSchedule 
         Appearance      =   0  'Flat
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
         Height          =   780
         Left            =   210
         TabIndex        =   4
         Top             =   300
         Width           =   2130
         Begin VB.OptionButton optCity 
            Caption         =   "City"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   315
            TabIndex        =   6
            Top             =   375
            Value           =   -1  'True
            Width           =   705
         End
         Begin VB.OptionButton optDetailSchedule 
            Caption         =   "Detail"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   1080
            TabIndex        =   5
            Top             =   375
            Width           =   800
         End
      End
   End
   Begin VB.CheckBox chkPrintNett 
      Caption         =   "Print Out Nett"
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
      Left            =   240
      TabIndex        =   2
      Top             =   1845
      Width           =   1470
   End
   Begin Crystal.CrystalReport crPrint 
      Left            =   -150
      Top             =   2295
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      PrintFileLinesPerPage=   60
   End
   Begin VB.CommandButton cmdClose 
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
      Height          =   495
      Left            =   1470
      TabIndex        =   1
      Top             =   2280
      Width           =   1180
   End
   Begin VB.CommandButton cmdPreview 
      Caption         =   "Pre&view"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   285
      TabIndex        =   0
      Top             =   2280
      Width           =   1180
   End
End
Attribute VB_Name = "Frm_Radio_MS_Print"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
''*************************************************************
'Fungsi Form        : memilih type report dari Radio Schedule
'Last Update/By     :
'*************************************************************

Option Explicit

Private Sub Form_Load()
    optSchedule.Value = True
    
    If Not Frm_Radio_Media_Quotation.blnJobType Then
        optDetailSchedule.Enabled = False
        optCity.Value = True
    Else
        optDetailSchedule.Enabled = True
    End If
    
    optDetailSchedule_Click
    
    chkPrintNett.Value = 0
End Sub

Private Sub chkPrintNett_Click()
    'set Flag untuk Print Out Nett
    If chkPrintNett.Value = vbChecked Then
        Frm_Radio_Media_Quotation.blnBoolPrintNettFlag = True
    Else
        Frm_Radio_Media_Quotation.blnBoolPrintNettFlag = False
    End If
End Sub

Private Sub cmdClose_Click()
    Unload Me
End Sub

Private Sub cmdPreview_Click()
    If optSchedule.Value = True Then
        If optCity.Value = True Then
            Call Frm_Radio_Media_Quotation.PreparePrint
        Else
            Call Frm_Radio_Media_Quotation.PreparePrintDetail
            
            crPrint.Reset
            crPrint.DataFiles(0) = "C:\ERPTEMPDB\RPT.MDB"
            crPrint.ReportFileName = Report_Dir & "\radio\MS_Radio_Book_Detail.rpt"
            
            If Frm_Radio_Media_Quotation.blnClientSpecial Then
                crPrint.Formulas(1) = "cap_job='JOB ID'"
                crPrint.Formulas(2) = "job_id='" & Frm_Radio_Media_Quotation.cboJobID.Text & "'"
            Else
                crPrint.Formulas(1) = "cap_job='JOB YEAR'"
                crPrint.Formulas(2) = "job_id='" & "20" & Mid(Frm_Radio_Media_Quotation.cboJobID.Text, 10, 2) & "'"
            End If
            
            crPrint.WindowState = crptMaximized
            crPrint.WindowShowRefreshBtn = True
            crPrint.WindowShowPrintSetupBtn = True
            crPrint.RetrieveDataFiles
            crPrint.WindowTitle = " -- Radio Quotation Detail -- "
            crPrint.Action = 1
        End If
        
    ElseIf optActual.Value = True Then
        Frm_Radio_Media_Quotation.PrintScheduleFromPO
    End If
End Sub

Private Sub optDetailSchedule_Click()
    If optCity.Value Then
        'jika yang di click adalah opsi city
        If Frm_Radio_Media_Quotation.blnJobType Then
            'untuk job 020
            chkPrintNett.Value = vbUnchecked
            chkPrintNett.Enabled = False
            
            'set Flag untuk Print Out Nett
            If chkPrintNett.Value = vbChecked Then
                Frm_Radio_Media_Quotation.blnBoolPrintNettFlag = True
            Else
                Frm_Radio_Media_Quotation.blnBoolPrintNettFlag = False
            End If
        Else
            chkPrintNett.Enabled = True
            
            'set Flag untuk Print Out Nett
            If chkPrintNett.Value = vbChecked Then
                Frm_Radio_Media_Quotation.blnBoolPrintNettFlag = True
            Else
                Frm_Radio_Media_Quotation.blnBoolPrintNettFlag = False
            End If
        End If
    Else
        'jika yang di click adalah opsi detail
        chkPrintNett.Value = vbUnchecked
        chkPrintNett.Enabled = False
    End If
End Sub

Private Sub optActual_Click()
    If Frm_Radio_Media_Quotation.blnJobType Then
        'untuk job 020
        chkPrintNett.Value = vbUnchecked
        chkPrintNett.Enabled = False
        
        If chkPrintNett.Value = vbChecked Then
            Frm_Radio_Media_Quotation.blnBoolPrintNettFlag = True
        Else
            Frm_Radio_Media_Quotation.blnBoolPrintNettFlag = False
        End If
    Else
        chkPrintNett.Enabled = True
        
        'set Flag untuk Print Out Nett
        If chkPrintNett.Value = vbChecked Then
            Frm_Radio_Media_Quotation.blnBoolPrintNettFlag = True
        Else
            Frm_Radio_Media_Quotation.blnBoolPrintNettFlag = False
        End If
    End If
        
    optCity.Enabled = False
    optDetailSchedule.Enabled = False
    optCity.Value = False
    optDetailSchedule.Value = False
End Sub

Private Sub optSchedule_Click()
    optCity.Enabled = True
    optDetailSchedule.Enabled = True
    optCity.Value = True
End Sub
