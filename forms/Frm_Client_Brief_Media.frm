VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TabCtl32.Ocx"
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Begin VB.Form Frm_Client_Brief_Media 
   Caption         =   "Client Brief"
   ClientHeight    =   9510
   ClientLeft      =   435
   ClientTop       =   630
   ClientWidth     =   13050
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "Frm_Client_Brief_Media.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   9510
   ScaleWidth      =   13050
   Begin VB.ComboBox cboBrand 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   1710
      TabIndex        =   60
      Text            =   "Cbo_Brand"
      ToolTipText     =   "Select Brand"
      Top             =   945
      Width           =   3945
   End
   Begin VB.TextBox txtYear 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   1695
      MaxLength       =   4
      TabIndex        =   59
      Text            =   "2000"
      Top             =   1320
      Width           =   540
   End
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
      ScaleWidth      =   13050
      TabIndex        =   7
      Top             =   9180
      Width           =   13050
      Begin VB.PictureBox picButton 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00F0F0F0&
         BorderStyle     =   0  'None
         ClipControls    =   0   'False
         ForeColor       =   &H80000008&
         Height          =   345
         Index           =   1
         Left            =   420
         ScaleHeight     =   345
         ScaleWidth      =   300
         TabIndex        =   12
         Top             =   15
         Width           =   300
      End
      Begin VB.PictureBox picButton 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00F0F0F0&
         BorderStyle     =   0  'None
         ClipControls    =   0   'False
         ForeColor       =   &H80000008&
         Height          =   345
         Index           =   0
         Left            =   90
         ScaleHeight     =   345
         ScaleWidth      =   300
         TabIndex        =   11
         Top             =   15
         Width           =   300
      End
      Begin VB.PictureBox picButton 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00F0F0F0&
         BorderStyle     =   0  'None
         ClipControls    =   0   'False
         ForeColor       =   &H80000008&
         Height          =   345
         Index           =   2
         Left            =   750
         ScaleHeight     =   345
         ScaleWidth      =   300
         TabIndex        =   10
         Top             =   15
         Width           =   300
      End
      Begin VB.PictureBox picButton 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00F0F0F0&
         BorderStyle     =   0  'None
         ClipControls    =   0   'False
         ForeColor       =   &H80000008&
         Height          =   345
         Index           =   3
         Left            =   1080
         ScaleHeight     =   345
         ScaleWidth      =   300
         TabIndex        =   9
         Top             =   15
         Width           =   300
      End
      Begin VB.PictureBox picDescColor 
         BorderStyle     =   0  'None
         Height          =   345
         Left            =   9975
         ScaleHeight     =   345
         ScaleWidth      =   1695
         TabIndex        =   8
         Top             =   15
         Width           =   1695
      End
      Begin VB.Label lblLastModifiedDate 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Last Modified Date:                        |"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   1470
         TabIndex        =   14
         Tag             =   "Last Modified Date: "
         Top             =   75
         Width           =   2520
      End
      Begin VB.Label lblLastModifiedBy 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Last Modified by:                                 |"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   4080
         TabIndex        =   13
         Tag             =   "Last Modified by: "
         Top             =   75
         Width           =   2775
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
      ScaleWidth      =   13050
      TabIndex        =   0
      Top             =   0
      Width           =   13050
      Begin VB.PictureBox picButton 
         Appearance      =   0  'Flat
         BackColor       =   &H00F0F0F0&
         BorderStyle     =   0  'None
         ClipControls    =   0   'False
         ForeColor       =   &H80000008&
         Height          =   750
         Index           =   7
         Left            =   4680
         ScaleHeight     =   750
         ScaleWidth      =   1500
         TabIndex        =   61
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
         Index           =   4
         Left            =   90
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
         ForeColor       =   &H80000008&
         Height          =   750
         Index           =   11
         Left            =   9270
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
         ForeColor       =   &H80000008&
         Height          =   750
         Index           =   10
         Left            =   7740
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
         ForeColor       =   &H80000008&
         Height          =   750
         Index           =   9
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
         ForeColor       =   &H80000008&
         Height          =   750
         Index           =   6
         Left            =   3150
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
         ForeColor       =   &H80000008&
         Height          =   750
         Index           =   5
         Left            =   1620
         ScaleHeight     =   750
         ScaleWidth      =   1500
         TabIndex        =   1
         Top             =   0
         Width           =   1500
      End
   End
   Begin Threed.SSPanel pnl_Main 
      Align           =   1  'Align Top
      Height          =   8355
      Left            =   0
      TabIndex        =   15
      Top             =   750
      Width           =   13050
      _Version        =   65536
      _ExtentX        =   23019
      _ExtentY        =   14737
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
      BevelOuter      =   1
      Begin VB.Frame fraInput 
         Caption         =   "Input"
         ForeColor       =   &H000000CD&
         Height          =   1455
         Left            =   105
         TabIndex        =   55
         Top             =   2250
         Width           =   7245
         Begin TabDlg.SSTab SSTab2 
            Height          =   1095
            Left            =   180
            TabIndex        =   56
            Top             =   270
            Width           =   6915
            _ExtentX        =   12197
            _ExtentY        =   1931
            _Version        =   393216
            Tabs            =   2
            Tab             =   1
            TabsPerRow      =   2
            TabHeight       =   520
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            TabCaption(0)   =   "Description of Activity"
            TabPicture(0)   =   "Frm_Client_Brief_Media.frx":0442
            Tab(0).ControlEnabled=   0   'False
            Tab(0).Control(0)=   "txtDescActivity"
            Tab(0).ControlCount=   1
            TabCaption(1)   =   "Activity Budget && Timing"
            TabPicture(1)   =   "Frm_Client_Brief_Media.frx":045E
            Tab(1).ControlEnabled=   -1  'True
            Tab(1).Control(0)=   "txtActivity_Time_Budget"
            Tab(1).Control(0).Enabled=   0   'False
            Tab(1).ControlCount=   1
            Begin VB.TextBox txtDescActivity 
               Height          =   570
               Left            =   -74775
               MultiLine       =   -1  'True
               ScrollBars      =   2  'Vertical
               TabIndex        =   58
               Top             =   405
               Width           =   6510
            End
            Begin VB.TextBox txtActivity_Time_Budget 
               Height          =   570
               Left            =   210
               MultiLine       =   -1  'True
               ScrollBars      =   2  'Vertical
               TabIndex        =   57
               Top             =   405
               Width           =   6510
            End
         End
      End
      Begin VB.Frame fra_Step 
         Caption         =   "Step"
         ForeColor       =   &H000000CD&
         Height          =   1995
         Left            =   135
         TabIndex        =   45
         Top             =   3735
         Width           =   7215
         Begin TabDlg.SSTab SSTab1 
            Height          =   1605
            Left            =   150
            TabIndex        =   46
            Top             =   285
            Width           =   6900
            _ExtentX        =   12171
            _ExtentY        =   2831
            _Version        =   393216
            Tab             =   2
            TabHeight       =   706
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            TabCaption(0)   =   "Target Audience"
            TabPicture(0)   =   "Frm_Client_Brief_Media.frx":047A
            Tab(0).ControlEnabled=   0   'False
            Tab(0).Control(0)=   "Txt_Socio_Demographic"
            Tab(0).Control(1)=   "Txt_Volumetric"
            Tab(0).Control(2)=   "Txt_Attitudinal"
            Tab(0).Control(3)=   "Label13"
            Tab(0).Control(4)=   "Label14"
            Tab(0).Control(5)=   "Label15"
            Tab(0).ControlCount=   6
            TabCaption(1)   =   "Marketing Objective(s)"
            TabPicture(1)   =   "Frm_Client_Brief_Media.frx":0496
            Tab(1).ControlEnabled=   0   'False
            Tab(1).Control(0)=   "txtMarketing_Objective"
            Tab(1).ControlCount=   1
            TabCaption(2)   =   "Communication Objective(s)"
            TabPicture(2)   =   "Frm_Client_Brief_Media.frx":04B2
            Tab(2).ControlEnabled=   -1  'True
            Tab(2).Control(0)=   "txtCommunication_Objective"
            Tab(2).Control(0).Enabled=   0   'False
            Tab(2).ControlCount=   1
            Begin VB.TextBox txtMarketing_Objective 
               Height          =   570
               Left            =   -74790
               MultiLine       =   -1  'True
               ScrollBars      =   2  'Vertical
               TabIndex        =   51
               Top             =   750
               Width           =   6510
            End
            Begin VB.TextBox txtCommunication_Objective 
               Height          =   570
               Left            =   210
               MultiLine       =   -1  'True
               ScrollBars      =   2  'Vertical
               TabIndex        =   50
               Top             =   750
               Width           =   6510
            End
            Begin VB.TextBox Txt_Attitudinal 
               Height          =   645
               Left            =   -72570
               MultiLine       =   -1  'True
               ScrollBars      =   2  'Vertical
               TabIndex        =   49
               Top             =   795
               Width           =   2100
            End
            Begin VB.TextBox Txt_Volumetric 
               Height          =   645
               Left            =   -74790
               MultiLine       =   -1  'True
               ScrollBars      =   2  'Vertical
               TabIndex        =   48
               Top             =   810
               Width           =   2100
            End
            Begin VB.TextBox Txt_Socio_Demographic 
               Height          =   645
               Left            =   -70365
               MultiLine       =   -1  'True
               ScrollBars      =   2  'Vertical
               TabIndex        =   47
               Top             =   795
               Width           =   2100
            End
            Begin VB.Label Label15 
               Caption         =   "Attitudinal"
               Height          =   240
               Left            =   -72585
               TabIndex        =   54
               Top             =   540
               Width           =   1965
            End
            Begin VB.Label Label14 
               Caption         =   "Socio-demographic"
               Height          =   240
               Left            =   -70365
               TabIndex        =   53
               Top             =   540
               Width           =   1965
            End
            Begin VB.Label Label13 
               Caption         =   "Volumetric"
               Height          =   240
               Left            =   -74775
               TabIndex        =   52
               Top             =   540
               Width           =   1965
            End
         End
      End
      Begin VB.Frame fra_DeliverableChannel 
         ForeColor       =   &H000000FF&
         Height          =   4980
         Left            =   7470
         TabIndex        =   41
         Top             =   2250
         Width           =   3270
         Begin VB.CheckBox chk_All 
            Height          =   240
            Left            =   270
            TabIndex        =   62
            Top             =   4575
            Width           =   240
         End
         Begin VB.ListBox lstRec_Channel_Selection 
            Height          =   3885
            ItemData        =   "Frm_Client_Brief_Media.frx":04CE
            Left            =   210
            List            =   "Frm_Client_Brief_Media.frx":04D5
            Style           =   1  'Checkbox
            TabIndex        =   42
            Top             =   630
            Width           =   2880
         End
         Begin VB.Label lbl_CheckAll 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Check All"
            Height          =   285
            Left            =   585
            TabIndex        =   63
            Top             =   4575
            Width           =   2340
         End
         Begin VB.Label Label12 
            Caption         =   "Deliverable"
            ForeColor       =   &H000000CD&
            Height          =   225
            Left            =   195
            TabIndex        =   44
            Top             =   15
            Width           =   990
         End
         Begin VB.Label Label11 
            Caption         =   "Recomended Channel Selection"
            Height          =   255
            Left            =   210
            TabIndex        =   43
            Top             =   285
            Width           =   2940
         End
      End
      Begin VB.Frame fra_Deliverable 
         Caption         =   "Deliverable"
         ForeColor       =   &H000000CD&
         Height          =   1440
         Left            =   135
         TabIndex        =   36
         Top             =   5790
         Width           =   7200
         Begin TabDlg.SSTab SSTab3 
            Height          =   1095
            Left            =   165
            TabIndex        =   37
            Top             =   255
            Width           =   6915
            _ExtentX        =   12197
            _ExtentY        =   1931
            _Version        =   393216
            Tabs            =   2
            TabsPerRow      =   2
            TabHeight       =   520
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            TabCaption(0)   =   "Agreed Channel Shortlist"
            TabPicture(0)   =   "Frm_Client_Brief_Media.frx":04F3
            Tab(0).ControlEnabled=   -1  'True
            Tab(0).Control(0)=   "txtAggreed_Channel_shortlist"
            Tab(0).Control(0).Enabled=   0   'False
            Tab(0).ControlCount=   1
            TabCaption(1)   =   "Other Recomendation"
            TabPicture(1)   =   "Frm_Client_Brief_Media.frx":050F
            Tab(1).ControlEnabled=   0   'False
            Tab(1).Control(0)=   "txtOther_Recomedation"
            Tab(1).ControlCount=   1
            Begin VB.TextBox Text14 
               Height          =   570
               Left            =   -74790
               TabIndex        =   40
               Top             =   405
               Width           =   6510
            End
            Begin VB.TextBox txtAggreed_Channel_shortlist 
               Height          =   570
               Left            =   225
               MultiLine       =   -1  'True
               ScrollBars      =   2  'Vertical
               TabIndex        =   39
               Top             =   405
               Width           =   6510
            End
            Begin VB.TextBox txtOther_Recomedation 
               Height          =   570
               Left            =   -74790
               MultiLine       =   -1  'True
               ScrollBars      =   2  'Vertical
               TabIndex        =   38
               Top             =   405
               Width           =   6510
            End
         End
      End
      Begin VB.Frame fraFilter 
         Height          =   2175
         Left            =   120
         TabIndex        =   16
         Top             =   30
         Width           =   10680
         Begin VB.ComboBox cboCountry 
            Appearance      =   0  'Flat
            Height          =   315
            Left            =   7950
            TabIndex        =   25
            Top             =   915
            Width           =   3945
         End
         Begin VB.TextBox txtExtention 
            Appearance      =   0  'Flat
            Height          =   315
            Left            =   1575
            TabIndex        =   24
            Top             =   1260
            Width           =   3945
         End
         Begin VB.Frame Fra_Approve 
            Height          =   855
            Left            =   5880
            TabIndex        =   19
            Top             =   1215
            Width           =   4635
            Begin VB.CheckBox chkAppTL 
               Height          =   270
               Left            =   165
               TabIndex        =   21
               Top             =   180
               Width           =   270
            End
            Begin VB.CheckBox chkAppCCM 
               Height          =   240
               Left            =   165
               TabIndex        =   20
               Top             =   525
               Width           =   240
            End
            Begin VB.Label Label10 
               Caption         =   "Approved by CCM"
               Height          =   285
               Left            =   540
               TabIndex        =   23
               Top             =   510
               Width           =   2340
            End
            Begin VB.Label Label9 
               Caption         =   "Approved by Team Leader"
               Height          =   285
               Left            =   540
               TabIndex        =   22
               Top             =   210
               Width           =   2325
            End
         End
         Begin VB.TextBox txtStatus 
            Appearance      =   0  'Flat
            Height          =   315
            Left            =   1575
            TabIndex        =   18
            Top             =   1635
            Width           =   3945
         End
         Begin VB.TextBox txtClient_Brief_Id 
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
            Height          =   315
            Left            =   1575
            TabIndex        =   17
            Top             =   900
            Width           =   1620
         End
         Begin MSComCtl2.DTPicker dtpDate_Issue 
            Height          =   315
            Left            =   7920
            TabIndex        =   26
            Top             =   555
            Width           =   1635
            _ExtentX        =   2884
            _ExtentY        =   582
            _Version        =   393216
            Format          =   126877697
            CurrentDate     =   36866
         End
         Begin MSComCtl2.DTPicker dtpDate_Previouse 
            Height          =   315
            Left            =   7920
            TabIndex        =   27
            Top             =   195
            Width           =   1635
            _ExtentX        =   2884
            _ExtentY        =   582
            _Version        =   393216
            Format          =   126877697
            CurrentDate     =   36866
         End
         Begin VB.Line lineFilter 
            BorderColor     =   &H00C0C0C0&
            X1              =   5655
            X2              =   5655
            Y1              =   165
            Y2              =   2040
         End
         Begin VB.Label Label1 
            Caption         =   "Brand "
            ForeColor       =   &H000000CD&
            Height          =   270
            Left            =   180
            TabIndex        =   35
            Top             =   240
            Width           =   1395
         End
         Begin VB.Label lblCountry 
            Caption         =   "Country : "
            Height          =   270
            Left            =   5835
            TabIndex        =   34
            Top             =   945
            Width           =   2040
         End
         Begin VB.Label Label3 
            Caption         =   "Extension "
            Height          =   270
            Left            =   180
            TabIndex        =   33
            Top             =   1305
            Width           =   1395
         End
         Begin VB.Label Label4 
            Caption         =   "Year "
            Height          =   270
            Left            =   180
            TabIndex        =   32
            Top             =   585
            Width           =   1320
         End
         Begin VB.Label lbl_dateofPreviousIssue 
            Caption         =   "Date of Previous Issue "
            Height          =   270
            Left            =   5835
            TabIndex        =   31
            Top             =   240
            Width           =   2040
         End
         Begin VB.Label lbl_DateIssue 
            Caption         =   "Date Issue : "
            Height          =   270
            Left            =   5835
            TabIndex        =   30
            Top             =   585
            Width           =   2040
         End
         Begin VB.Label Label7 
            Caption         =   "Status "
            Height          =   300
            Left            =   180
            TabIndex        =   29
            Top             =   1680
            Width           =   660
         End
         Begin VB.Label Label8 
            Caption         =   "Client Brief Id"
            Height          =   270
            Left            =   195
            TabIndex        =   28
            Top             =   945
            Width           =   1395
         End
      End
   End
End
Attribute VB_Name = "Frm_Client_Brief_Media"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'<CSCC>
'********************************************************************************
'Submodul Name      : Frm_Client_Brief_Media
'Submodul Function  : entry Client Brief Media
'Used Table         : -
'Used SP            : -
'Procedure/Function : -
'Programmer Name    : yayan
'Date               : 8/Jan/2001
'Last Update By     : Tedi
'Date Update        : 3/9/2016-4:21:17 AM
'Log Update/By      : -
'********************************************************************************
'</CSCC>

Option Explicit

Dim recClient_Brief_Media As New ADODB.Recordset
Dim blnEdit_Flag As Boolean
Dim blnEditOrAdd As Boolean 'variable untuk menyatakan apakah sedang dilakukan edit atau delete
Dim blnNoRecord As Boolean
Dim strSql As String
Dim blnNotByClickByList As Boolean

Private Sub cboBrand_Click()
'<CSCM>
'********************************************************************************
'Procedure Name     : CboBrand_Click
'Procedure Function : ---
'Input Parameter    : ---
'Output Parameter   : ---
'Date               : 3/9/2016
'LastUpdate/By      : Tedi / Kreatif
'********************************************************************************
'</CSCM>

    If Trim(cboBrand.Text) = "" Then
        Exit Sub
    End If
    
    'Close Recordset
    If recClient_Brief_Media.State = adStateOpen Then
        recClient_Brief_Media.Close
    End If
    
    'Open Recordset
    strSql = "SELECT * FROM Client_Brief_Media WHERE brand_code='" & Left(cboBrand.Text, 4) & "'"
    recClient_Brief_Media.Open strSql, ConnERP, adOpenDynamic, adLockOptimistic, adCmdText
    
    If Not recClient_Brief_Media.EOF And Not recClient_Brief_Media.BOF Then
        Button_Normal (True)
        ' Call Show Data
        Call viewDetail
        blnNoRecord = False
    Else
        MsgBox "No Record.", vbExclamation, strApplication_Name
        Button_No_Record (True)
        blnNoRecord = True
        Clear_Form
    End If
    
End Sub

Private Sub cboBrand_KeyPress(KeyAscii As Integer)
'<CSCM>
'********************************************************************************
'Procedure Name     : Cbo_Brand_KeyPress
'Procedure Function : ---
'Input Parameter    : ---
'Output Parameter   : ---
'Date               : 3/9/2016
'LastUpdate/By      : Tedi / Kreatif
'********************************************************************************
'</CSCM>

    KeyAscii = 0
    
End Sub

Private Sub cboBrand_LostFocus()
'<CSCM>
'********************************************************************************
'Procedure Name     : Cbo_Brand_LostFocus
'Procedure Function : ---
'Input Parameter    : ---
'Output Parameter   : ---
'Date               : 3/9/2016
'LastUpdate/By      : Tedi / Kreatif
'********************************************************************************
'</CSCM>

    If txtYear.Enabled Then
        txtYear.SetFocus
    End If
    
End Sub

Private Sub db_add()
'************************************************
' Procedure         : db_Add
' Before Name       : Cmd_Add_Click
' Function          : To Add Record
' Date              : 10/18/2000
' Parameter Input   :
' Parameter Output  :
' Last Update/By    : Tedi / 22 Feb 2016
'************************************************
    'Disable Button
    If cboBrand.ListIndex = -1 Then
        MsgBox "Please Select Brand!", vbExclamation, strApplication_Name
        Exit Sub
    End If
    
    EnableObject True
    
    
    Button_Normal (False)
    'Clear Form
    Call Clear_Form
    'Defualt Value
    txtYear.Text = Year(Date)
    txtYear.SetFocus
    
    cboCountry.Text = "Indonesia"
       
    'Add  Record in Creative_Order
    recClient_Brief_Media.AddNew

    Exit Sub
label:
    MsgBox Err.Number & Err.Description & Err.Source & Err.HelpFile, vbExclamation, strApplication_Name
End Sub

Private Sub db_Cancel()
'************************************************
' Procedure         : db_Cancel
' Name Before       : Cmd_Cancel_Click
' Function          : To Cancel Any Changes
' Date              : 10/18/2000
' Parameter Input   :
' Parameter Output  :
' Last Update/By    :
'************************************************

    On Error Resume Next
    
    If blnNoRecord Then
        recClient_Brief_Media.CancelUpdate
        Button_No_Record (True)
        'Cbo_Brand.Enabled = True
        Clear_Form
        Exit Sub
    End If
    
    'Cancel Update
    recClient_Brief_Media.CancelUpdate
    ' Close Transaction
    If Not blnEdit_Flag Then
       ' ConnERP.RollbackTrans
       Cancel_Brief_Id_Media_Running txtClient_Brief_Id.Text
       
       recClient_Brief_Media.MoveLast
    End If
    'Enable Button
    Button_Normal (True)
    
    'Cbo_Brand.Enabled = True
    
    blnEdit_Flag = False
    'Show Actual Data
    viewDetail
    EnableObject False
    
End Sub

Private Sub db_delete()
'************************************************
' Procedure         : db_Delete
' Procedure         : Delete Record
' Function          : To Delete Record
' Date              : 10/19/2000
' Parameter Input   :
' Parameter Output  :
' Last Update Date  : 22 Maret 2016
' Last Update By    : Tedi / Kreatif
' Name before       : Cmd_Delete_Click
'************************************************

    Dim blnAnsware As Integer
    Dim strSql As String
    
    On Error GoTo label
    
    blnAnsware = MsgBox("Do you want to delete?", vbYesNo + vbQuestion + vbExclamation, strApplication_Name)
    If blnAnsware = vbYes Then
        
        ConnERP.BeginTrans
                
        'Delete Dari Channel REC
        strSql = "DELETE FROM Channel_Rec_by_Activity_Deliverable WHERE Client_Brief_Id='" & txtClient_Brief_Id.Text & "'"
        
        ConnERP.Execute strSql
        
        'Delete Record
        
        recClient_Brief_Media.Delete
        
        ConnERP.CommitTrans
        
        recClient_Brief_Media.Requery
        
        Clear_Form
        
        MsgBox "Record Deleted.", vbExclamation, strApplication_Name
        
        If recClient_Brief_Media.EOF Then
            Button_No_Record True
        End If
        
    End If
    
    Exit Sub
label:
    MsgBox Err.Description, vbExclamation, strApplication_Name
    ConnERP.RollbackTrans
    
    recClient_Brief_Media.Requery
    recClient_Brief_Media.Find "Client_Brief_Id='" & txtClient_Brief_Id.Text & "'"
    If Not recClient_Brief_Media.EOF Then
        viewDetail
    Else
        Clear_Form
    End If
    
End Sub

Private Sub db_edit()
'************************************************
' Porcedure         : db_Edit
' Name Before       : Cmd_Edit_Click
' Function          : To Edit Record
' Date              : 10/19/2000
' Parameter Input   :
' Parameter Output  :
' Last Update/By    :
'************************************************
    
    If cboBrand.ListIndex = -1 Then
        MsgBox "Please Select Brand!", vbExclamation, strApplication_Name
        Exit Sub
    End If
    
    blnEdit_Flag = True
    'Disable Button
    EnableObject True
    Button_Normal (False)
    If blnEdit_Flag Then
        'Cbo_Brand.Enabled = False
        txtYear.Enabled = False
    End If
    
End Sub

Private Sub db_Save()
'************************************************
' Procedure         : db_Save
' Name Before       : Cmd_Save_Click
' Function          : To Save Any Change, Add and Edit
' Date              : 10/19/2000
' Parameter Input   :
' Parameter Output  :
' Last Update/By    :
'************************************************
    
    On Error GoTo label
    blnNoRecord = False
    'Vadate Field Value
    If Trim(cboBrand.Text) = "" Then
        MsgBox "Select Brand...!", vbExclamation, strApplication_Name
            Exit Sub
    End If
    
    If dtpDate_Issue.Value < dtpDate_Previouse.Value Then
        MsgBox "Check Date of Previous Issue or Date Issue !", vbExclamation, strApplication_Name
        Exit Sub
    End If
    
    'Assigment Value every Field
    recClient_Brief_Media("Client_Brief_Id") = Trim(txtClient_Brief_Id.Text)
    recClient_Brief_Media("Brand_Code").Value = Trim(Left(cboBrand.Text, 4))
    recClient_Brief_Media("Country").Value = Trim(cboCountry.Text)
    recClient_Brief_Media("Year").Value = Val(txtYear.Text)
    recClient_Brief_Media("Extention").Value = Trim(txtExtention.Text)
    recClient_Brief_Media("Status").Value = Trim(txtStatus.Text)
    recClient_Brief_Media("Date_issue").Value = dtpDate_Issue.Value
    recClient_Brief_Media("Date_of_Previouse_Issue").Value = dtpDate_Previouse.Value
    If chkAppTL.Value = vbChecked Then
        recClient_Brief_Media("Approved_Team_Leader").Value = 1
    Else
        recClient_Brief_Media("Approved_Team_Leader").Value = 0
    End If
    
    If chkAppCCM.Value = vbChecked Then
        recClient_Brief_Media("Approved_By_CCM").Value = 1
    Else
        recClient_Brief_Media("Approved_By_CCM").Value = 0
    End If
    
    recClient_Brief_Media("Description").Value = Trim(txtDescActivity.Text)
    recClient_Brief_Media("Budget_Timing").Value = Trim(txtActivity_Time_Budget.Text)
    recClient_Brief_Media("Volumetric").Value = Trim(Txt_Volumetric.Text)
    recClient_Brief_Media("Attitudinal").Value = Trim(Txt_Attitudinal.Text)
    recClient_Brief_Media("Socio_Demographic").Value = Trim(Txt_Socio_Demographic.Text)
    recClient_Brief_Media("Marketing_Objective").Value = Trim(txtMarketing_Objective.Text)
    recClient_Brief_Media("Communication_Objective").Value = Trim(txtCommunication_Objective.Text)
    recClient_Brief_Media("Agreed_Channel_Shortlist").Value = Trim(txtAggreed_Channel_shortlist.Text)
    recClient_Brief_Media("Other_Recomendation").Value = Trim(txtOther_Recomedation.Text)
    'Save Record
    recClient_Brief_Media.Update
    
    'Save Deliverable
    '****************
    Dim intIndex_List As Integer
    Dim strSql As String
    'For Edit
    If blnEdit_Flag Then
    'Delete Existing Channel
        strSql = "DELETE FROM Channel_Rec_by_Activity_Deliverable WHERE Client_Brief_Id='" & txtClient_Brief_Id.Text & "'"
        ConnERP.Execute strSql
    End If
    'Insert Selected Channel
    For intIndex_List = 0 To lstRec_Channel_Selection.ListCount - 1
        If lstRec_Channel_Selection.Selected(intIndex_List) Then
            strSql = "INSERT INTO Channel_Rec_by_Activity_Deliverable (Channel_Id,Client_Brief_Id) VALUES (" & lstRec_Channel_Selection.ItemData(intIndex_List) & ",'" & txtClient_Brief_Id.Text & "')"
            ConnERP.Execute strSql
        End If
    Next intIndex_List
    
    'LOg Status
    'If Not blnEdit_Flag Then
    '    Save_Log_Status 2, Frm_Client_Brief.txtBrief_Id, Trim(Txt_Job_Number.Text)
    'End If
    
    'Enable Button
    Button_Normal (True)
    
    'Cbo_Brand.Enabled = True
    
    blnEdit_Flag = False
    EnableObject False
    recClient_Brief_Media.Requery
    RemoteMovement txtClient_Brief_Id.Text
    Exit Sub
label:
    If Err.Number = -2147217893 Then
        MsgBox Err.Description, vbExclamation, strApplication_Name
    Else
        MsgBox Err.Description, vbExclamation, strApplication_Name
    End If
    
End Sub



Private Sub db_Find()
'************************************************
' Procedure         : db_Find
' Function          : Mencari Client_Brief_Id
' Date              : 01/08/2001
' Parameter Input   :
' Parameter Output  :
' Programmer By     : Tedi / Kreatif
' Last Update       : 22 Feb 2016
' Last update  By   : Tedi / Kreatif
' Before Description : MsgBox "Under Construction", vbInformation, strApplication_Name
'************************************************
    
    Dim str_arrSplit() As String
    If cboBrand.Text <> "" Then
        frm_Find_Cient_BriefM.strClient_Brief_Id = txtClient_Brief_Id.Text
    Else
        MsgBox "Please select brand before click find!", vbExclamation, strApplication_Name
        Exit Sub
    End If
    frm_Find_Cient_BriefM.strYear = txtYear.Text
    str_arrSplit = Split(cboBrand.Text, "-")
    frm_Find_Cient_BriefM.strBrand = Trim(str_arrSplit(0))
    frm_Find_Cient_BriefM.show vbModal
    
End Sub

Public Sub RemoteMovement(ByVal Client_Brief_Id As String)
    recClient_Brief_Media.MoveFirst
    recClient_Brief_Media.Find "Client_Brief_Id='" & Client_Brief_Id & "'"
    Call viewDetail
End Sub


Private Sub chk_All_Click()
    'If chk_All.Value = 1 Then
        
    'Else
    '    blnNotByClickByList = False
    'End If
    Call CheckForCildren(lstRec_Channel_Selection, chk_All)
End Sub

Private Sub CheckForCildren(ByRef ObjListView As ListBox, ByRef objChk As CheckBox)
'*****************************************
'Procedure Name     : CheckForCildren
'Procedure Function : Untuk memberikan reaksi kepada children root
'Description        : jika node dicontreng maka semua children root akan tercontreng
'Input Parameter    : strArgParent --> string argument untuk nilai node key parent
'                     bolCheck --> boolean untuk nilai checkbox node key parent
'Used Object        : objListView,objChkBox,bol_Temp
'Programmer Name    : Tedi
'Date               : 19-11-2015
'Last Update/By     : Tedi
'Date Update        :
'Log Update/By      :
'***************************************************************'*****************************************
    
    Dim iCheck As Integer
    If blnNotByClickByList = False Then Exit Sub
    For iCheck = 0 To ObjListView.ListCount - 1
'        If objChk.Value = 1 Then
            ObjListView.Selected(iCheck) = objChk.Value
'        Else
'            objListView.Selected(iCheck) = False
'        End If
    Next iCheck
    blnNotByClickByList = False
End Sub

Private Sub chk_All_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    blnNotByClickByList = True
End Sub

Private Sub Form_Load()
'************************************************
' Procedure         : Form_Load
' Function          : To Load Form
' Date              : 01/08/2001
' Parameter Input   :
' Parameter Output  :
' Last Update/By    :
'************************************************
    'Declare Variable
    Dim Brief_Id As String
    Dim rs_brand As New ADODB.Recordset
    Dim Rs_Country As New ADODB.Recordset
    Dim rs_Channel As New ADODB.Recordset
    Dim strSql As String
    
    'Center Form
    EnableObject False
    'CenterForm Me
    RemoveMenus Me, True
        
    'Load Country To Combo
    Set Rs_Country = New ADODB.Recordset
    Rs_Country.Open "select * from Country", ConnERP, adOpenStatic, adLockReadOnly, adCmdText
    cboCountry.Clear
    If Not Rs_Country.EOF And Not Rs_Country.BOF Then
        Rs_Country.MoveFirst
        Do While Not Rs_Country.EOF
            cboCountry.AddItem Trim(Rs_Country!Country)
            Rs_Country.MoveNext
        Loop
    End If
    Rs_Country.Close
    Set Rs_Country = Nothing
    
    'Load Brand To Combo
    Set rs_brand = New ADODB.Recordset
    strSql = "SELECT * FROM brand WHERE brand_code IN (SELECT brand_code FROM Media_Security_Catalog WHERE User_name='" & strLogin_User & "' AND position IN ('Planner','Implementor') and Valid_until > getdate())"
    rs_brand.Open strSql, ConnERP, adOpenStatic, adLockReadOnly, adCmdText
     
    
    cboBrand.Clear
    If Not rs_brand.EOF And Not rs_brand.BOF Then
        rs_brand.MoveFirst
        Do While Not rs_brand.EOF
            cboBrand.AddItem Trim(rs_brand!Brand_code & " - " & rs_brand!Brand_Name)
            rs_brand.MoveNext
        Loop
    Else
        Button_No_Record True
        'Cmd_Add.Enabled = False
        MsgBox "You have no brand.", vbExclamation, strApplication_Name
        rs_brand.Close
        Set rs_brand = Nothing
        Exit Sub
    End If
    rs_brand.Close
    Set rs_brand = Nothing

    'Load Channel To Listbox
    
    rs_Channel.Open "SELECT * FROM Channel", ConnERP, adOpenStatic, adLockReadOnly, adCmdText
    lstRec_Channel_Selection.Clear
    If Not rs_Channel.EOF And Not rs_Channel.BOF Then
        rs_Channel.MoveFirst
        Do While Not rs_Channel.EOF
            lstRec_Channel_Selection.AddItem Trim(rs_Channel!Channel_Desc)
            lstRec_Channel_Selection.ItemData(lstRec_Channel_Selection.NewIndex) = rs_Channel!Channel_Id
            rs_Channel.MoveNext
        Loop
    Else
        MsgBox "Channel Table Empty...!", vbExclamation, strApplication_Name
    End If
    rs_Channel.Close
    Set rs_Channel = Nothing
    

'Open Recodset Client Brief
    'StrSql = "SELECT * FROM Client_Brief_Media WHERE brand_code IN ( SELECT Brand_Code FROM BRAND WHERE AM_Media_Name='" & strLogin_User & "')"
    strSql = "SELECT * FROM Client_Brief_Media WHERE brand_code='" & Left(cboBrand.Text, 4) & "'"
    recClient_Brief_Media.Open strSql, ConnERP, adOpenDynamic, adLockOptimistic, adCmdText
    
    If Not recClient_Brief_Media.EOF And Not recClient_Brief_Media.BOF Then
        Button_Normal (True)
        ' Call Show Data
        Call viewDetail
        blnNoRecord = False
    Else
        'MsgBox "No Record", vbInformation, strApplication_Name
        Button_No_Record (True)
        blnNoRecord = True
    End If
    AdjustSizeForm
End Sub

Private Sub viewDetail()
'************************************************
' Procedure         : viewdetail
' Function          : To Show Data from Recodset to Form
' Date              : 01/9/2001
' Parameter Input   :
' Parameter Output  :
' Last Update/By    :
'************************************************
    
    txtClient_Brief_Id.Text = recClient_Brief_Media("Client_Brief_Id").Value
    cboBrand.Text = recClient_Brief_Media("Brand_Code").Value & " - " & Get_Brand_Name(recClient_Brief_Media("Brand_Code").Value)
    cboCountry.Text = IIf(IsNull(recClient_Brief_Media("Country").Value), "", recClient_Brief_Media("Country").Value)
    txtYear.Text = IIf(IsNull(recClient_Brief_Media("Year").Value), "", recClient_Brief_Media("Year").Value)
    txtExtention.Text = IIf(IsNull(recClient_Brief_Media("Extention").Value), "", recClient_Brief_Media("Extention").Value)
    txtStatus.Text = IIf(IsNull(recClient_Brief_Media("Status").Value), "", recClient_Brief_Media("Status").Value)
    
    dtpDate_Issue.Value = IIf(IsNull(recClient_Brief_Media("Date_issue").Value), Date, recClient_Brief_Media("Date_issue").Value)
    dtpDate_Previouse.Value = IIf(IsNull(recClient_Brief_Media("Date_of_Previouse_Issue").Value), Date, recClient_Brief_Media("Date_of_Previouse_Issue").Value)

    If recClient_Brief_Media("Approved_Team_Leader").Value = 1 Then
        chkAppTL.Value = vbChecked
    Else
        chkAppTL.Value = vbUnchecked
    End If
    
    If recClient_Brief_Media("Approved_By_CCM").Value = 1 Then
        chkAppCCM.Value = vbChecked
    Else
        chkAppCCM.Value = vbUnchecked
    End If
    
    txtDescActivity.Text = IIf(IsNull(recClient_Brief_Media("Description").Value), "", recClient_Brief_Media("Description").Value)
    txtActivity_Time_Budget.Text = IIf(IsNull(recClient_Brief_Media("Budget_Timing").Value), "", recClient_Brief_Media("Budget_Timing").Value)
    'Txt_Target_Audience.Text = IIf(IsNull(recClient_Brief_Media("Target_Audience").Value), "", recClient_Brief_Media("Target_Audience").Value)
    Txt_Volumetric.Text = IIf(IsNull(recClient_Brief_Media("Volumetric").Value), "", recClient_Brief_Media("Volumetric").Value)
    Txt_Attitudinal.Text = IIf(IsNull(recClient_Brief_Media("Attitudinal").Value), "", recClient_Brief_Media("Attitudinal").Value)
    Txt_Socio_Demographic.Text = IIf(IsNull(recClient_Brief_Media("Socio_Demographic").Value), "", recClient_Brief_Media("Socio_Demographic").Value)
    txtMarketing_Objective.Text = IIf(IsNull(recClient_Brief_Media("Marketing_Objective").Value), "", recClient_Brief_Media("Marketing_Objective").Value)
    txtCommunication_Objective.Text = IIf(IsNull(recClient_Brief_Media("Communication_Objective").Value), "", recClient_Brief_Media("Communication_Objective").Value)
    txtAggreed_Channel_shortlist.Text = IIf(IsNull(recClient_Brief_Media("Agreed_Channel_Shortlist").Value), "", recClient_Brief_Media("Agreed_Channel_Shortlist").Value)
    txtOther_Recomedation.Text = IIf(IsNull(recClient_Brief_Media("Other_Recomendation").Value), "", recClient_Brief_Media("Other_Recomendation").Value)
    
    'Load Deliverable
    '****************
    Dim intIndex_List As Integer
    'Clear Previouse
    For intIndex_List = 0 To lstRec_Channel_Selection.ListCount - 1
        lstRec_Channel_Selection.Selected(intIndex_List) = False
    Next intIndex_List
    'Lst_Rec_Channel_Selection.Clear
    'Load
    Dim Rs_Deliverable As New ADODB.Recordset
    
    Rs_Deliverable.Open "SELECT Channel_Id FROM Channel_Rec_by_Activity_Deliverable WHERE Client_Brief_Id='" & recClient_Brief_Media("Client_Brief_Id").Value & "'", ConnERP, adOpenForwardOnly, adLockReadOnly
    Do While Not Rs_Deliverable.BOF And Not Rs_Deliverable.EOF
        For intIndex_List = 0 To lstRec_Channel_Selection.ListCount - 1
            If lstRec_Channel_Selection.ItemData(intIndex_List) = Rs_Deliverable(0).Value Then
                lstRec_Channel_Selection.Selected(intIndex_List) = True
            End If
        Next intIndex_List
        Rs_Deliverable.MoveNext
    Loop
    Rs_Deliverable.Close
    Set Rs_Deliverable = Nothing
    
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    Call CloseRecordset(recClient_Brief_Media)
End Sub

Private Sub Form_Resize()
'************************************************
' Procedure         : Form_Resize
' Function          : Mengatur Posisi Form
' Parameter Input   : -
' Parameter Output  : -
' Last Update       :
' Last Update By    :
'************************************************

    Me.Top = 0
    Me.Left = 0

End Sub

Private Sub Form_Unload(Cancel As Integer)
On Error Resume Next
'************************************************
' Procedure         : Form_Unload
' Function          : To Unload Form
' Date              : 01/09/2001
' Parameter Input   : Cancel
' Parameter Output  :
' Last Update/By    :
'************************************************

    ' Close Recordset
    If Not blnNoRecord Then
        recClient_Brief_Media.Close
        Set recClient_Brief_Media = Nothing
    End If

End Sub

Private Sub Button_Normal(Enable As Boolean)
'************************************************
' Procedure         : Button_Normal
' Function          : To Enable or Disable Button
' Date              : 01/01/2001
' Parameter Input   : Enable
' Parameter Output  :
' Last Update/By    :
'************************************************

    txtClient_Brief_Id.Enabled = False
    cboCountry.Enabled = Not Enable
    txtYear.Enabled = Not Enable
    txtExtention.Enabled = Not Enable
    txtStatus.Enabled = Not Enable
    
    'Cbo_Brand.Enabled = Not Enable
    dtpDate_Issue.Enabled = Not Enable
    dtpDate_Previouse.Enabled = Not Enable
    
    fra_DeliverableChannel.Enabled = Not Enable
    Fra_Approve.Enabled = Not Enable
    'Chk_App_CCM.Enabled = Not Enable
    'Chk_app_TL.Enabled = Not Enable
    
    txtDescActivity.Enabled = Not Enable
    txtActivity_Time_Budget.Enabled = Not Enable
    'Txt_Target_Audience.Enabled = Not Enable
    Txt_Volumetric.Enabled = Not Enable
    Txt_Attitudinal.Enabled = Not Enable
    Txt_Socio_Demographic.Enabled = Not Enable
    txtMarketing_Objective.Enabled = Not Enable
    txtCommunication_Objective.Enabled = Not Enable
    txtAggreed_Channel_shortlist.Enabled = Not Enable
    txtOther_Recomedation.Enabled = Not Enable
    
End Sub

Private Sub Button_No_Record(Enable As Boolean)
'************************************************
' Procedure         : Button_Normal
' Function          : To Enable or Disable Button
' Date              : 10/18/2000
' Parameter Input   : Enable
' Parameter Output  :
' Last Update/By    :
'************************************************

    txtClient_Brief_Id.Enabled = False
    cboCountry.Enabled = False
    txtYear.Enabled = False
    txtExtention.Enabled = False
    txtStatus.Enabled = False
    
    'Cbo_Brand.Enabled = False
    dtpDate_Issue.Enabled = False
    dtpDate_Previouse.Enabled = False
    
    Fra_Approve.Enabled = False
    'Chk_App_CCM.Enabled = False
    'Chk_app_TL.Enabled = False
    fra_DeliverableChannel.Enabled = False
    
    txtDescActivity.Enabled = False
    txtActivity_Time_Budget.Enabled = False
    'Txt_Target_Audience.Enabled = False
    Txt_Volumetric.Enabled = False
    Txt_Attitudinal.Enabled = False
    Txt_Socio_Demographic.Enabled = False
    txtMarketing_Objective.Enabled = False
    txtCommunication_Objective.Enabled = False
    txtAggreed_Channel_shortlist.Enabled = False
    txtOther_Recomedation.Enabled = False
    
End Sub


Private Sub Clear_Form()
'************************************************
' Procedure         : Clear_Form
' Function          : To Claer Form
' Date              : 01/09/2001
' Parameter Input   :
' Parameter Output  :
' Last Update/By    :
'************************************************

    txtClient_Brief_Id.Text = ""
    cboCountry.Text = ""
    txtYear.Text = ""
    txtExtention.Text = ""
    txtStatus.Text = ""
    
    'Cbo_Brand.Text = ""
    dtpDate_Issue.Value = Date
    dtpDate_Previouse.Value = Date
    chkAppCCM.Value = vbUnchecked
    chkAppTL.Value = vbUnchecked
    
    txtActivity_Time_Budget.Text = ""
    txtDescActivity.Text = ""
    txtOther_Recomedation.Text = ""
    txtAggreed_Channel_shortlist.Text = ""
    Txt_Attitudinal.Text = ""
    txtCommunication_Objective.Text = ""
    txtMarketing_Objective.Text = ""
    Txt_Socio_Demographic.Text = ""
    'Txt_Target_Audience.Text = ""
    Txt_Volumetric.Text = ""
    
    Dim Idx_list As Integer
    
    For Idx_list = 0 To lstRec_Channel_Selection.ListCount - 1
        lstRec_Channel_Selection.Selected(Idx_list) = False
    Next Idx_list
    
End Sub

Private Sub lstRec_Channel_Selection_ItemCheck(Item As Integer)
'<CSCM>
'********************************************************************************
'Procedure Name     : Lst_Rec_Channel_Selection_ItemCheck
'Procedure Function : ---
'Input Parameter    : ---
'Output Parameter   : ---
'Date               : 3/9/2016
'LastUpdate/By      : Tedi / Kreatif
'********************************************************************************
'</CSCM>
  
    Call CheckForClickAll(lstRec_Channel_Selection, chk_All, blnNotByClickByList)

End Sub

Private Sub lstRec_Channel_Selection_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
'<CSCM>
'********************************************************************************
'Procedure Name     : Lst_Rec_Channel_Selection_MouseMove
'Procedure Function : ---
'Input Parameter    : ---
'Output Parameter   : ---
'Date               : 3/9/2016
'LastUpdate/By      : Tedi / Kreatif
'********************************************************************************
'</CSCM>
    
    blnNotByClickByList = False

End Sub

Private Sub txtYear_KeyPress(KeyAscii As Integer)
'<CSCM>
'********************************************************************************
'Procedure Name     : Txt_Year_KeyPress
'Procedure Function : ---
'Input Parameter    : ---
'Output Parameter   : ---
'Date               : 3/9/2016
'LastUpdate/By      : Tedi / Kreatif
'********************************************************************************
'</CSCM>

    If (Chr(KeyAscii) < "0" Or Chr(KeyAscii) > "9") And KeyAscii <> 8 Then
        KeyAscii = 0
        Beep
    End If

End Sub

Private Sub txtYear_LostFocus()
'************************************************
' Procedure         : Txt_Year_LostFocus
' Function          : Generate IB ID
' Date              : 01/09/2001
' Parameter Input   :
' Parameter Output  :
' Last Update/By    :
'************************************************

    Dim Brand_code As String
    Dim New_Brief_Id As String
    Dim Year As Integer
    ' check the null text in combo box
      If cboBrand.Text = "" Then
            cboBrand.SetFocus
            MsgBox "Select Brand..!", vbExclamation, strApplication_Name
         Exit Sub
      End If
      If Val(txtYear.Text) = 0 Then
        MsgBox "Please Check Year...", vbExclamation, strApplication_Name
        txtYear.SetFocus
        SendKeys "{Home}+{End}"
        Exit Sub
      End If
      
      'Disable Year and Brand
        txtYear.Enabled = False
        'Cbo_Brand.Enabled = False
                
        'Generate Brief Id
        '****************************
        Dim Out_Param As ADODB.Parameter
        Dim In_Param1 As ADODB.Parameter
        Dim In_Param2 As ADODB.Parameter
        Dim cmd As New ADODB.Command
        
        'Get Last Job Number  From Job_Number_History
            Brand_code = Left(cboBrand.Text, 4)
            Year = Val(txtYear.Text)
            
            cmd.CommandType = adCmdStoredProc
            cmd.CommandText = "Get_New_Brief_Id_Media"
            
            Set In_Param1 = cmd.CreateParameter("Brand_Code", adChar, adParamInput, 10)
            Set In_Param2 = cmd.CreateParameter("Year", adInteger, adParamInput)
            Set Out_Param = cmd.CreateParameter("New_Brief_Id", adChar, adParamOutput, 10)
               
            cmd.Parameters.Append In_Param1
            cmd.Parameters.Append In_Param2
            cmd.Parameters.Append Out_Param

            In_Param1.Value = Brand_code
            In_Param2.Value = Year
           
            
            cmd.ActiveConnection = ConnERP
            cmd.Execute
            
            New_Brief_Id = Out_Param.Value
            
        'Put New Job Number to Txt_Job_Number
        txtClient_Brief_Id.Text = New_Brief_Id
        
End Sub

Sub SetButtonToolbar(ByVal paIsNormalMode As Boolean, picOBJ) 'TOOLBAR_AI.
'************************************************
' Procedure         : SetButtonToolbar
' Function          : TOOLBAR_AI.
' Created By        : {73 64 6B}
' Date              : 12-Apr-2015
'LastUpdate/By      : - Rudi
'************************************************

    Dim element
    Dim strDummy As String
    
    With picButton(enButtonType.bieAdd)  'ADD. 4
        .Enabled = paIsNormalMode
        .Visible = paIsNormalMode
    End With
    With picButton(enButtonType.bieEdit) 'EDIT. 5
        .Enabled = paIsNormalMode
        .Visible = paIsNormalMode
    End With
    With picButton(enButtonType.bieDelete)  'DELETE. 6
        .Enabled = paIsNormalMode
        .Visible = paIsNormalMode
    End With
    With picButton(enButtonType.bieExit)      'Quit.
        .Enabled = paIsNormalMode
        .Visible = paIsNormalMode
    End With

    With picButton(enButtonType.bieSave)  'SAVE.
        .Enabled = Not paIsNormalMode
        .Visible = Not paIsNormalMode
        .Left = picButton(4).Left
    End With
    With picButton(enButtonType.biefind)  'FIND.
        .Enabled = paIsNormalMode
        .Visible = paIsNormalMode
        '.Left = picButton(4).Left
    End With

    With picButton(enButtonType.bieCancel) 'CANCEL.
        .Enabled = Not paIsNormalMode
        .Visible = Not paIsNormalMode
        .Left = picButton(5).Left
    End With
    pnl_Main.Enabled = Not paIsNormalMode
    cboBrand.Enabled = paIsNormalMode
    blnEditOrAdd = Not paIsNormalMode
    For Each element In picOBJ
        SetPictureTB element.Index, paIsNormalMode, picOBJ
    Next element
    'Call SetSecurityCRUDStandar("Duration Catalog", picButton, "1")

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

Private Sub EnableObject(ByVal paIsEnable As Boolean)
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

Sub AdjustSizeForm()
'************************************************
' Procedure         : Txt_Year_LostFocus
' Function          : Generate IB ID
' Date              : 01/09/2001
' Parameter Input   :
' Parameter Output  :
' Last Update/By    :
'************************************************
    
    Me.Top = 0
    Me.Left = 0
    Me.Width = mdi_Main.ScaleWidth
    Me.Height = mdi_Main.ScaleHeight
    pnl_Main.Height = Me.ScaleHeight - picToolbar.Height - picStatusBar.Height
    fra_Deliverable.Height = pnl_Main.Height - (fra_Deliverable.Top + 100)
    SSTab3.Height = fra_Deliverable.Height - (SSTab3.Top) - 150
    txtOther_Recomedation.Height = SSTab3.Height - (txtOther_Recomedation.Top) - 150
    txtAggreed_Channel_shortlist.Height = txtOther_Recomedation.Height
    fra_DeliverableChannel.Height = pnl_Main.Height - (fra_DeliverableChannel.Top + 100)
    fraFilter.Width = pnl_Main.Width - (fraFilter.Left * 2)
    lineFilter.X1 = fraFilter.Width / 2
    lineFilter.X2 = lineFilter.X1
    Fra_Approve.Left = lineFilter.X2 + Label7.Left
    txtYear.Width = lineFilter.X2 - txtYear.Left - 50
    txtClient_Brief_Id.Width = txtYear.Width
    txtExtention.Width = txtYear.Width
    txtStatus.Width = txtYear.Width
    'left part
    lbl_dateofPreviousIssue.Left = lineFilter.X1 + Label7.Left
    dtpDate_Previouse.Left = lbl_dateofPreviousIssue.Left + lbl_dateofPreviousIssue.Width + 50
    dtpDate_Issue.Left = dtpDate_Previouse.Left
    lbl_DateIssue.Left = lbl_dateofPreviousIssue.Left
    lblCountry.Left = lbl_dateofPreviousIssue.Left
    cboCountry.Left = dtpDate_Previouse.Left
    Fra_Approve.Left = dtpDate_Previouse.Left
    fra_DeliverableChannel.Width = pnl_Main.Width - fra_DeliverableChannel.Left - fraFilter.Left
    lstRec_Channel_Selection.Width = fra_DeliverableChannel.Width - (lstRec_Channel_Selection.Left * 2)
    lstRec_Channel_Selection.Height = fra_DeliverableChannel.Height - (lstRec_Channel_Selection.Top) - 200
    chk_All.Top = lstRec_Channel_Selection.Height + lstRec_Channel_Selection.Top + 50
    lbl_CheckAll.Top = chk_All.Top
End Sub

Private Sub picButton_Click(Index As Integer)

'************************************************
' Procedure         : picButton_Click
' Function          : Action utk Navigation dan CRUD.
' Created By        : {73 64 6B}
' Date              : 12-Apr-2015/{73 64 6B} --> Semua coding dan query sudah di optimalkan agar faster, readable, safer, standardable.
'************************************************
    Dim strCode As String, strFileRpt As String
    'Lock_MainForm True
    Select Case Index
        Case enButtonType.bieFirst 'FIRST.
            recClient_Brief_Media.MoveFirst
            viewDetail
            'tdb_Task_Click
        Case enButtonType.biePrev  'PREV.
            recClient_Brief_Media.MovePrevious
            If recClient_Brief_Media.BOF Then
                recClient_Brief_Media.MoveFirst
                MsgBox "Fist Record", vbInformation, strApplication_Name
            End If
            viewDetail
        Case enButtonType.bieNext 'NEXT.
            recClient_Brief_Media.MoveNext
            If recClient_Brief_Media.EOF Then
                recClient_Brief_Media.MoveLast
                MsgBox "Last Record", vbInformation, strApplication_Name
            End If
            viewDetail
        Case enButtonType.bieLast  'LAST.
            If recClient_Brief_Media.EOF Then Exit Sub
            recClient_Brief_Media.MoveLast
            viewDetail
        Case enButtonType.bieAdd  '4 'ADD.
            Call db_add
        Case enButtonType.bieEdit  '5 'EDIT.
            Call db_edit
        Case enButtonType.bieDelete  '6 'DELETE.
            Call db_delete
        Case enButtonType.biefind   '7 'FIND.
            Call db_Find
        Case enButtonType.bieExit  '9 'EXIT.
            Unload Me
        Case enButtonType.bieSave  'SAVE.
            Call db_Save
        Case enButtonType.bieCancel 'CANCEL.
            Call db_Cancel
    End Select

End Sub

Private Sub CheckForClickAll(ByRef ObjListBox As ListBox, ByRef objChkBox As CheckBox, ByVal bol_Temp As Boolean)
'*****************************************
'Submodul Name      : CheckForClickAll
'Procedure Function : Untuk memeriksa kompisisi apakah row di listview tercontreng semua
'                     - Jika node tercontreng semua maka nilai chkAll/objChkBox.Value = 1, jika tidak maka chkAll/objChkBox.Value  = 0
'                       Pemrosesan chkAll.Value diperintahkan dengan code, sehingga perlu diberikan nilai bolean blnNotByClickByList = True
'                     - jika check list row ada yang tidak tercontreng maka nilai bol_Temp=false sebaliknya true
'Used Object        : objListBox,objChkBox,bol_Temp
'Programmer Name    : Tedi
'Date               : 19-11-2015
'Last Update/By     : Tedi
'Date Update        :
'Log Update/By      :
'***************************************************************'*****************************************
    
    If blnNotByClickByList = True Then Exit Sub
    Dim intCheck As Integer
    For intCheck = 0 To ObjListBox.ListCount - 1
        If ObjListBox.Selected(intCheck) = False Then
            'bol_Temp = False
            objChkBox.Value = 0
            Exit Sub
        End If
    Next intCheck
    objChkBox.Value = 1
    'bol_Temp = False
End Sub


