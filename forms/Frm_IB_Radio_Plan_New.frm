VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "mscomctl.OCX"
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFlxGrd.ocx"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDatGrd.ocx"
Begin VB.Form Frm_IB_Radio_Plan_New 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Plan Edit"
   ClientHeight    =   9555
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7680
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9555
   ScaleWidth      =   7680
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
      ScaleWidth      =   7680
      TabIndex        =   51
      Top             =   0
      Width           =   7680
      Begin VB.PictureBox picButton 
         Appearance      =   0  'Flat
         BackColor       =   &H00F0F0F0&
         BorderStyle     =   0  'None
         ClipControls    =   0   'False
         ForeColor       =   &H80000008&
         Height          =   750
         Index           =   23
         Left            =   4680
         ScaleHeight     =   750
         ScaleWidth      =   1500
         TabIndex        =   57
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
         TabIndex        =   56
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
         TabIndex        =   55
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
         TabIndex        =   53
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
         Left            =   10800
         ScaleHeight     =   750
         ScaleWidth      =   1500
         TabIndex        =   52
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
         Left            =   4680
         ScaleHeight     =   750
         ScaleWidth      =   1500
         TabIndex        =   54
         Top             =   0
         Width           =   1500
      End
   End
   Begin Threed.SSPanel pnlMain 
      Height          =   9375
      Left            =   -15
      TabIndex        =   0
      Top             =   750
      Width           =   7605
      _Version        =   65536
      _ExtentX        =   13414
      _ExtentY        =   16536
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
      Begin VB.Frame fraMonthlyBudget 
         Appearance      =   0  'Flat
         Caption         =   "Monthly Budget"
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
         Height          =   2700
         Left            =   165
         TabIndex        =   40
         Top             =   75
         Width           =   7275
         Begin VB.ComboBox cboMonth 
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   885
            TabIndex        =   43
            Top             =   285
            Width           =   2175
         End
         Begin VB.CommandButton cmdDeleteBudget 
            Caption         =   "&Delete Monthly Budget"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Left            =   5835
            TabIndex        =   42
            Top             =   2040
            Width           =   1215
         End
         Begin VB.CommandButton cmdSaveBudget 
            Caption         =   "&Save Monthly Budget"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Left            =   3465
            TabIndex        =   41
            Top             =   2040
            Width           =   1180
         End
         Begin MSDataGridLib.DataGrid dgdMonthlyBudget 
            Height          =   1095
            Left            =   180
            TabIndex        =   44
            Top             =   750
            Width           =   6885
            _ExtentX        =   12144
            _ExtentY        =   1931
            _Version        =   393216
            AllowUpdate     =   0   'False
            HeadLines       =   1
            RowHeight       =   15
            BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ColumnCount     =   2
            BeginProperty Column00 
               DataField       =   ""
               Caption         =   ""
               BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
                  Type            =   0
                  Format          =   ""
                  HaveTrueFalseNull=   0
                  FirstDayOfWeek  =   0
                  FirstWeekOfYear =   0
                  LCID            =   1033
                  SubFormatType   =   0
               EndProperty
            EndProperty
            BeginProperty Column01 
               DataField       =   ""
               Caption         =   ""
               BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
                  Type            =   0
                  Format          =   ""
                  HaveTrueFalseNull=   0
                  FirstDayOfWeek  =   0
                  FirstWeekOfYear =   0
                  LCID            =   1033
                  SubFormatType   =   0
               EndProperty
            EndProperty
            SplitCount      =   1
            BeginProperty Split0 
               BeginProperty Column00 
               EndProperty
               BeginProperty Column01 
               EndProperty
            EndProperty
         End
         Begin MSMask.MaskEdBox medBudget 
            Height          =   330
            Left            =   4890
            TabIndex        =   45
            Top             =   270
            Width           =   2250
            _ExtentX        =   3969
            _ExtentY        =   582
            _Version        =   393216
            Appearance      =   0
            BackColor       =   16777215
            MaxLength       =   12
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            PromptChar      =   "_"
         End
         Begin VB.CommandButton cmdCancelBudget 
            Caption         =   "&Cancel"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Left            =   4635
            TabIndex        =   47
            Top             =   2040
            Width           =   1200
         End
         Begin VB.CommandButton cmdEditBudget 
            Caption         =   "&Edit Monthly Budget"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Left            =   4650
            TabIndex        =   46
            Top             =   2040
            Width           =   1180
         End
         Begin VB.Label lblMonth 
            BackStyle       =   0  'Transparent
            Caption         =   "Month "
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000C0&
            Height          =   255
            Left            =   210
            TabIndex        =   49
            Top             =   330
            Width           =   615
         End
         Begin VB.Label lblBudget 
            BackStyle       =   0  'Transparent
            Caption         =   "Budget "
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000C0&
            Height          =   255
            Left            =   4125
            TabIndex        =   48
            Top             =   315
            Width           =   735
         End
      End
      Begin VB.Frame fraPlanDetail 
         Appearance      =   0  'Flat
         Caption         =   "Plan Detail"
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
         Height          =   5835
         Left            =   165
         TabIndex        =   2
         Top             =   2865
         Width           =   7290
         Begin VB.Frame fraSchedule 
            Appearance      =   0  'Flat
            Caption         =   "Schedule"
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
            Height          =   1080
            Left            =   135
            TabIndex        =   30
            Top             =   195
            Width           =   7005
            Begin VB.TextBox txtSpot 
               Alignment       =   1  'Right Justify
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFFF&
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   325
               Left            =   5310
               MaxLength       =   4
               TabIndex        =   34
               Top             =   630
               Width           =   615
            End
            Begin VB.TextBox txtSchedule 
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFFF&
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   315
               Left            =   1620
               MaxLength       =   30
               TabIndex        =   33
               Top             =   630
               Width           =   2175
            End
            Begin VB.ComboBox cboSchedule 
               Height          =   315
               Left            =   1620
               TabIndex        =   32
               Top             =   630
               Width           =   2175
            End
            Begin VB.ComboBox cboMonthlyBudget 
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   315
               Left            =   1620
               TabIndex        =   31
               Top             =   240
               Width           =   2175
            End
            Begin VB.Label lblSpot 
               BackStyle       =   0  'Transparent
               Caption         =   "Spot / day "
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H000000C0&
               Height          =   255
               Left            =   4275
               TabIndex        =   37
               Top             =   675
               Width           =   975
            End
            Begin VB.Label lblSchedule 
               BackStyle       =   0  'Transparent
               Caption         =   "Schedu&le "
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H000000C0&
               Height          =   255
               Left            =   210
               TabIndex        =   36
               Top             =   675
               Width           =   1050
            End
            Begin VB.Label lblMonthlyBudget 
               BackStyle       =   0  'Transparent
               Caption         =   "Monthly Budget "
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H000000C0&
               Height          =   255
               Left            =   210
               TabIndex        =   35
               Top             =   285
               Width           =   1305
            End
         End
         Begin VB.Frame fraMaterialMix 
            Appearance      =   0  'Flat
            Caption         =   "Material MIX (%)"
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
            Height          =   1830
            Left            =   135
            TabIndex        =   21
            Top             =   3900
            Width           =   7005
            Begin VB.CommandButton cmdDeleteMix 
               Caption         =   "Delete"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   495
               Left            =   5670
               TabIndex        =   25
               Top             =   1125
               Width           =   1180
            End
            Begin VB.CommandButton cmdAddMix 
               Caption         =   "Add"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   495
               Left            =   4485
               TabIndex        =   24
               Top             =   1125
               Width           =   1180
            End
            Begin VB.TextBox txtMix 
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFFF&
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   300
               Left            =   4455
               MaxLength       =   3
               TabIndex        =   23
               Top             =   690
               Width           =   735
            End
            Begin VB.ComboBox cboMaterial 
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   315
               Left            =   4440
               TabIndex        =   22
               Top             =   285
               Width           =   2415
            End
            Begin MSFlexGridLib.MSFlexGrid msgMix 
               Height          =   1245
               Left            =   135
               TabIndex        =   26
               Top             =   345
               Width           =   3270
               _ExtentX        =   5768
               _ExtentY        =   2196
               _Version        =   393216
               FixedCols       =   0
               SelectionMode   =   1
               FormatString    =   "Material Id  | Mix(%)"
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
            Begin MSDataGridLib.DataGrid dgdMix 
               Height          =   1215
               Left            =   135
               TabIndex        =   27
               Top             =   360
               Visible         =   0   'False
               Width           =   3255
               _ExtentX        =   5741
               _ExtentY        =   2143
               _Version        =   393216
               Appearance      =   0
               HeadLines       =   1
               RowHeight       =   15
               BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ColumnCount     =   2
               BeginProperty Column00 
                  DataField       =   ""
                  Caption         =   ""
                  BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
                     Type            =   0
                     Format          =   ""
                     HaveTrueFalseNull=   0
                     FirstDayOfWeek  =   0
                     FirstWeekOfYear =   0
                     LCID            =   1057
                     SubFormatType   =   0
                  EndProperty
               EndProperty
               BeginProperty Column01 
                  DataField       =   ""
                  Caption         =   ""
                  BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
                     Type            =   0
                     Format          =   ""
                     HaveTrueFalseNull=   0
                     FirstDayOfWeek  =   0
                     FirstWeekOfYear =   0
                     LCID            =   1057
                     SubFormatType   =   0
                  EndProperty
               EndProperty
               SplitCount      =   1
               BeginProperty Split0 
                  BeginProperty Column00 
                  EndProperty
                  BeginProperty Column01 
                  EndProperty
               EndProperty
            End
            Begin VB.Label lblMix 
               BackStyle       =   0  'Transparent
               Caption         =   "MIX (%) "
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H000000C0&
               Height          =   255
               Left            =   3555
               TabIndex        =   29
               Top             =   720
               Width           =   855
            End
            Begin VB.Label lblMaterial 
               BackStyle       =   0  'Transparent
               Caption         =   "Mate&rial "
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H000000C0&
               Height          =   255
               Left            =   3555
               TabIndex        =   28
               Top             =   315
               Width           =   855
            End
         End
         Begin VB.Frame fraSalesArea 
            Appearance      =   0  'Flat
            Caption         =   "Sales Area"
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
            Height          =   2520
            Left            =   135
            TabIndex        =   11
            Top             =   1335
            Width           =   7005
            Begin VB.CommandButton Command1 
               Caption         =   "Remove All"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   495
               Left            =   2850
               TabIndex        =   58
               ToolTipText     =   "Remove Sales Area Cities"
               Top             =   1305
               Width           =   870
            End
            Begin VB.ListBox lstCity 
               BackColor       =   &H00FFFFFF&
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               ItemData        =   "Frm_IB_Radio_Plan_New.frx":0000
               Left            =   165
               List            =   "Frm_IB_Radio_Plan_New.frx":0007
               TabIndex        =   19
               Top             =   1380
               Visible         =   0   'False
               Width           =   2610
            End
            Begin VB.CheckBox chkRural 
               Caption         =   "Rural"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Left            =   5460
               TabIndex        =   18
               Top             =   2175
               Width           =   780
            End
            Begin VB.CheckBox chkUrban 
               Caption         =   "Urban"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Left            =   3960
               TabIndex        =   17
               Top             =   2175
               Width           =   900
            End
            Begin VB.CommandButton cmdRemoveCity 
               Caption         =   "Remove City"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   495
               Left            =   2835
               TabIndex        =   16
               ToolTipText     =   "Remove Sales Area Cities"
               Top             =   795
               Width           =   870
            End
            Begin VB.ListBox lstSelectedCity 
               BackColor       =   &H00FFFFFF&
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   1620
               Left            =   3795
               TabIndex        =   15
               Top             =   405
               Width           =   3015
            End
            Begin VB.Frame fraArea 
               Appearance      =   0  'Flat
               Caption         =   "Area"
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
               Height          =   1785
               Left            =   135
               TabIndex        =   12
               Top             =   285
               Width           =   2610
               Begin VB.CommandButton cmdManageArea 
                  Caption         =   "Manage Area"
                  BeginProperty Font 
                     Name            =   "Tahoma"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   495
                  Left            =   675
                  TabIndex        =   14
                  Top             =   1050
                  Width           =   1180
               End
               Begin VB.ComboBox cboRadioArea 
                  BeginProperty Font 
                     Name            =   "Tahoma"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   315
                  Left            =   225
                  TabIndex        =   13
                  Top             =   435
                  Width           =   2175
               End
            End
            Begin VB.Label lblSelectedCity 
               BackStyle       =   0  'Transparent
               Caption         =   "Sele&cted City(s) :"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Left            =   3795
               TabIndex        =   20
               Top             =   135
               Width           =   1815
            End
         End
         Begin VB.CommandButton cmdAdd 
            Caption         =   "&Add"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Left            =   135
            TabIndex        =   10
            Top             =   5865
            Visible         =   0   'False
            Width           =   1180
         End
         Begin VB.CommandButton cmdEdit 
            Caption         =   "&Edit"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Left            =   1320
            TabIndex        =   9
            Top             =   5865
            Visible         =   0   'False
            Width           =   1180
         End
         Begin VB.CommandButton cmdDelete 
            Caption         =   "&Delete"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Left            =   2505
            TabIndex        =   8
            Top             =   5865
            Visible         =   0   'False
            Width           =   1180
         End
         Begin VB.CommandButton cmdClose 
            Caption         =   "Cl&ose"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Left            =   5955
            TabIndex        =   7
            Top             =   5835
            Visible         =   0   'False
            Width           =   1180
         End
         Begin VB.Frame fraSalesStation 
            Appearance      =   0  'Flat
            Caption         =   "Sales Station"
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
            Height          =   2520
            Left            =   135
            TabIndex        =   3
            Top             =   1335
            Width           =   7005
            Begin MSComctlLib.TreeView treStations 
               Height          =   2085
               Left            =   120
               TabIndex        =   4
               Top             =   360
               Width           =   3255
               _ExtentX        =   5741
               _ExtentY        =   3678
               _Version        =   393217
               HideSelection   =   0   'False
               LabelEdit       =   1
               LineStyle       =   1
               Style           =   6
               BorderStyle     =   1
               Appearance      =   1
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
            Begin MSComctlLib.TreeView treSelectedStation 
               Height          =   2085
               Left            =   3600
               TabIndex        =   5
               Top             =   360
               Width           =   3255
               _ExtentX        =   5741
               _ExtentY        =   3678
               _Version        =   393217
               HideSelection   =   0   'False
               LabelEdit       =   1
               LineStyle       =   1
               Style           =   6
               BorderStyle     =   1
               Appearance      =   1
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
            Begin VB.Label lblSelectedCity2 
               BackStyle       =   0  'Transparent
               Caption         =   "Sele&cted Station(s) :"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Left            =   3840
               TabIndex        =   6
               Top             =   135
               Width           =   2175
            End
         End
         Begin VB.CommandButton cmdCancel 
            Caption         =   "&Cancel"
            Height          =   495
            Left            =   1320
            TabIndex        =   38
            Top             =   5850
            Visible         =   0   'False
            Width           =   1180
         End
         Begin VB.CommandButton cmdSave 
            Caption         =   "&Save"
            Height          =   495
            Left            =   135
            TabIndex        =   39
            Top             =   5850
            Visible         =   0   'False
            Width           =   1180
         End
      End
      Begin VB.ListBox lstSelectedStation 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1815
         Left            =   7605
         TabIndex        =   1
         Top             =   4545
         Width           =   2055
      End
      Begin MSDataGridLib.DataGrid dgdPlanDetailTemp 
         Height          =   1095
         Left            =   420
         TabIndex        =   50
         Top             =   5730
         Visible         =   0   'False
         Width           =   6795
         _ExtentX        =   11986
         _ExtentY        =   1931
         _Version        =   393216
         AllowUpdate     =   0   'False
         HeadLines       =   1
         RowHeight       =   15
         BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ColumnCount     =   2
         BeginProperty Column00 
            DataField       =   ""
            Caption         =   ""
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1033
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column01 
            DataField       =   ""
            Caption         =   ""
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1033
               SubFormatType   =   0
            EndProperty
         EndProperty
         SplitCount      =   1
         BeginProperty Split0 
            BeginProperty Column00 
            EndProperty
            BeginProperty Column01 
            EndProperty
         EndProperty
      End
   End
   Begin VB.Menu mnuKanan 
      Caption         =   ""
      Visible         =   0   'False
      Begin VB.Menu mnuRemoveAll 
         Caption         =   "Remove All"
      End
   End
   Begin VB.Menu mnuStationDelete 
      Caption         =   "Delete Station"
      Visible         =   0   'False
      Begin VB.Menu mnuStatDelete 
         Caption         =   "Delete All Station"
      End
   End
   Begin VB.Menu mnuStationAdd 
      Caption         =   "Add Station"
      Visible         =   0   'False
      Begin VB.Menu mnuStatAdd 
         Caption         =   "Add Station"
      End
   End
   Begin VB.Menu mnuStationAll 
      Caption         =   "Select All"
      Visible         =   0   'False
      Begin VB.Menu mnuSelectAllStation 
         Caption         =   "Select All"
      End
   End
End
Attribute VB_Name = "Frm_IB_Radio_Plan_New"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
''************************************************
' Function          : entry plan detal IB radio
' Last Update       :.'
'************************************************
Option Explicit
Const Mix_Max = 100

Dim intJumMix As Integer
Dim blnFillMonthlyBudget As Boolean
Dim blnFillPlanDetail As Boolean
Dim recCityDummy As New ADODB.Recordset
Dim recMixDummy As New ADODB.Recordset
Dim recStationTemp As New ADODB.Recordset 'dw
Dim strTransacProcess As String 'dw
Dim strOldSchedule As String
Dim nodx As Node 'dw

Private Sub Command1_Click()
    Dim intCount As Integer

    Do While True
        If lstSelectedCity.ListCount = 0 Then Exit Do
        lstSelectedCity.ListIndex = 0
        cmdRemoveCity_Click
    Loop
End Sub

Private Sub Form_Load()
    blnFillMonthlyBudget = True
    
    Call Frm_IB_Radio.monthIB(cboMonth)
    
    Call Frm_IB_Radio.monthIB(cboMonthlyBudget)
    
    Call SetButton(False)
    
    'set button
    cmdSaveBudget.Enabled = False
    cmdCancelBudget.Visible = False
    cmdEditBudget.Visible = True
    cmdDeleteBudget.Enabled = True
    
    'set object
    cboMonth.Enabled = False
    cboMonth.Text = ""
    medBudget.Enabled = False
    medBudget.Text = 0
    
    Call ViewGrdBudget
    
    blnFillPlanDetail = False
    
    SwitchSchedule False
    
    FillScheduleList
    
    'dw - Refer from plan detail options selected on IB
    '--------------------------------------------------
    If Frm_IB_Radio.cboPlanDetail.Text = "City" Then
        fraSalesArea.Visible = True
        fraSalesStation.Visible = False
    ElseIf Frm_IB_Radio.cboPlanDetail.Text = "Station" Then
        fraSalesArea.Visible = False
        fraSalesStation.Visible = True
        
        LoadStations
    End If
    '---------------------------------------------------
    If dgdMonthlyBudget.ApproxCount = 0 Then
        fraMonthlyBudget.Enabled = True
        fraSchedule.Enabled = False
        fraSalesStation.Enabled = False
        fraMaterialMix.Enabled = False
    Else
        fraSchedule.Enabled = True
        If txtSpot.Text <> "" Then fraSalesStation.Enabled = True
        If cboSchedule.Text <> "" Then fraSalesStation.Enabled = True
    End If
    Call EnableObject(False)
    'pnlMain.Enabled = False
End Sub

Private Sub RestoreToOriginal()
    Dim intPosisi As Integer
    
    Frm_IB_Radio.recPlanDetailTemp.Filter = ""
    While Not Frm_IB_Radio.recPlanDetailTemp.EOF And Not Frm_IB_Radio.recPlanDetailTemp.BOF
        Frm_IB_Radio.recPlanDetailTemp.Delete
        Frm_IB_Radio.recPlanDetailTemp.MoveNext
    Wend
    
    If Frm_IB_Radio.cboPlanDetail.Text = "City" Then
        
        'city
        recCityDummy.Filter = ""
        While Not recCityDummy.EOF And Not recCityDummy.BOF
            Frm_IB_Radio.recPlanDetailTemp.AddNew
            
            For intPosisi = 0 To 9
                 Frm_IB_Radio.recPlanDetailTemp.Fields(intPosisi).Value = recCityDummy.Fields(intPosisi).Value
            Next
            
            Frm_IB_Radio.recPlanDetailTemp.Update
            
            recCityDummy.MoveNext
        Wend
        
    ElseIf Frm_IB_Radio.cboPlanDetail.Text = "Station" Then 'dw - set recStationtemp to restore to original (cancel proses)
        
        'station
        recStationTemp.Filter = ""
        While Not recStationTemp.EOF And Not recStationTemp.BOF
            Frm_IB_Radio.recPlanDetailTemp.AddNew
            
            For intPosisi = 0 To 9
                 Frm_IB_Radio.recPlanDetailTemp.Fields(intPosisi).Value = recStationTemp.Fields(intPosisi).Value
            Next
            
            Frm_IB_Radio.recPlanDetailTemp.Update
            
            recStationTemp.MoveNext
        Wend
    End If
    
    'mix
    Frm_IB_Radio.recPlanDetailMaterialTemp.Filter = ""
    While Not Frm_IB_Radio.recPlanDetailMaterialTemp.EOF And Not Frm_IB_Radio.recPlanDetailMaterialTemp.BOF
        Frm_IB_Radio.recPlanDetailMaterialTemp.Delete
        Frm_IB_Radio.recPlanDetailMaterialTemp.MoveNext
    Wend
    
    recMixDummy.Filter = ""
    While Not recMixDummy.EOF And Not recMixDummy.BOF
        Frm_IB_Radio.recPlanDetailMaterialTemp.AddNew
        
        For intPosisi = 0 To 7
             Frm_IB_Radio.recPlanDetailMaterialTemp.Fields(intPosisi).Value = recMixDummy.Fields(intPosisi).Value
        Next
        
        Frm_IB_Radio.recPlanDetailMaterialTemp.Update
        
        recMixDummy.MoveNext
    Wend
    
    CloseDummy
End Sub

Private Sub PrepareEditDummy()
    Dim intPosisi As Integer
    
    If Frm_IB_Radio.cboPlanDetail.Text = "City" Then
        CreateCityDummy
        
        Frm_IB_Radio.recPlanDetailTemp.Filter = ""
        While Not Frm_IB_Radio.recPlanDetailTemp.EOF And Not Frm_IB_Radio.recPlanDetailTemp.BOF
            recCityDummy.AddNew
            
            For intPosisi = 0 To 9
                recCityDummy.Fields(intPosisi).Value = Frm_IB_Radio.recPlanDetailTemp.Fields(intPosisi).Value
            Next
            
            recCityDummy.Update
            
            Frm_IB_Radio.recPlanDetailTemp.MoveNext
        Wend
        
    ElseIf Frm_IB_Radio.cboPlanDetail.Text = "Station" Then 'dw - set recStationTemp for edit proses
        CreateStationTemp
        
        Frm_IB_Radio.recPlanDetailTemp.Filter = ""
        While Not Frm_IB_Radio.recPlanDetailTemp.EOF And Not Frm_IB_Radio.recPlanDetailTemp.BOF
            recStationTemp.AddNew
            
            For intPosisi = 0 To 9
                recStationTemp.Fields(intPosisi).Value = Frm_IB_Radio.recPlanDetailTemp.Fields(intPosisi).Value
            Next
            
            recStationTemp.Update
            
            Frm_IB_Radio.recPlanDetailTemp.MoveNext
        Wend
    End If
    
    Set dgdPlanDetailTemp.DataSource = Frm_IB_Radio.recPlanDetailTemp
        
    CreateMixDummy
        
    Frm_IB_Radio.recPlanDetailMaterialTemp.Filter = ""
    While Not Frm_IB_Radio.recPlanDetailMaterialTemp.EOF And Not Frm_IB_Radio.recPlanDetailMaterialTemp.BOF
        recMixDummy.AddNew
        
        For intPosisi = 0 To 7
            recMixDummy.Fields(intPosisi).Value = Frm_IB_Radio.recPlanDetailMaterialTemp.Fields(intPosisi).Value
        Next
        
        recMixDummy.Update
        
        Frm_IB_Radio.recPlanDetailMaterialTemp.MoveNext
    Wend
End Sub

Private Sub CloseDummy()
    If recCityDummy.State = adStateOpen Then
        recCityDummy.Close
        Set recCityDummy = Nothing
    End If

    If recMixDummy.State = adStateOpen Then
        recMixDummy.Close
        Set recMixDummy = Nothing
    End If
    
    If recStationTemp.State = adStateOpen Then 'dw -close recStationTemp
        recStationTemp.Close
        Set recStationTemp = Nothing
    End If
End Sub

Private Sub CreateMixDummy()
    Set recMixDummy = Nothing
    Set recMixDummy = New ADODB.Recordset
    
    With recMixDummy.Fields
        .Append "Client_Brief_Id", adChar, 10, adFldMayBeNull
        .Append "IB_Id", adChar, 13, adFldMayBeNull
        .Append "Month", adInteger, , adFldMayBeNull
        .Append "Year", adInteger, , adFldMayBeNull
        .Append "Schedule", adVarChar, 75, adFldMayBeNull
        .Append "City_Code", adInteger, , adFldMayBeNull
        .Append "Material_Id", adChar, 1, adFldMayBeNull
        .Append "Material_Mix", adDouble, , adFldMayBeNull
    End With
    
    recMixDummy.Open
End Sub

Private Sub CreateCityDummy()
    Set recCityDummy = Nothing
    Set recCityDummy = New ADODB.Recordset

    With recCityDummy.Fields
        .Append "Client_Brief_Id", adChar, 10, adFldMayBeNull
        .Append "IB_Id", adChar, 13, adFldMayBeNull
        .Append "Month", adInteger, , adFldMayBeNull
        .Append "Year", adInteger, , adFldMayBeNull
        .Append "Schedule", adVarChar, 75, adFldMayBeNull
        .Append "City_Code", adInteger, , adFldMayBeNull
        .Append "Area_Code", adChar, 50, adFldMayBeNull
        .Append "Spot", adInteger, , adFldMayBeNull
        .Append "Urban_Flag", adSmallInt, , adFldMayBeNull
        .Append "Rural_Flag", adSmallInt, , adFldMayBeNull
    End With
    
    recCityDummy.Open
End Sub

Private Sub cboMaterial_DropDown()
    Call LoadMateri
End Sub

Private Sub cboMaterial_KeyPress(KeyAscii As Integer)
    KeyAscii = 0
End Sub

Private Sub CboMonth_Click()
    With Frm_IB_Radio.recRadioPlanTemp
        .Filter = ""
                
        strQuery = " IB_ID ='" & Frm_IB_Radio.txtIBID.Text & "'"
        strQuery = strQuery & " AND Year = " & Val(Frm_IB_Radio.cboYear.Text)
        strQuery = strQuery & " AND Month = " & Get_Month_Number(cboMonth.Text)
        .Filter = strQuery
        
        If Not .EOF Then
            medBudget.Text = .Fields("Budget").Value
        Else
            medBudget.Text = 0
        End If
        
        .Filter = ""
        
        Call ViewGrdBudget
    End With
End Sub

Private Sub cboMonth_KeyPress(KeyAscii As Integer)
    KeyAscii = 0
End Sub

Private Sub cboMonthlyBudget_Click()
    ClearForm
    
    cboSchedule.Text = ""
    fraSalesStation.Enabled = True
End Sub

Private Sub cboMonthlyBudget_KeyPress(KeyAscii As Integer)
    KeyAscii = 0
End Sub

Private Sub cboRadioArea_Click()
    Dim recSelected As New ADODB.Recordset
    
    If txtSchedule.Text = "" Then
        MsgBox "You Must Fill Schedule first", vbCritical, strTitleMissingInfo
        cboRadioArea.Clear
        Exit Sub
    End If
    
    If IIf(txtSpot.Text = "", 0, txtSpot.Text) <= 0 Then
        MsgBox "Please insert Spot/Day !", vbCritical, strTitleMissingInfo
        cboRadioArea.Clear
        txtSpot.SetFocus
        Exit Sub
    End If
      
    Me.MousePointer = vbHourglass
    
    lstSelectedCity.Clear
    
    With Frm_IB_Radio.recPlanDetailTemp
        strQuery = "IB_Id='" & Frm_IB_Radio.txtIBID.Text & "' AND "
        strQuery = strQuery & "Month=" & Get_Month_Number(cboMonthlyBudget.Text) & " AND "
        strQuery = strQuery & "Year=" & Val(Frm_IB_Radio.cboYear.Text) & " AND "
        strQuery = strQuery & "Schedule='" & Clear_String(txtSchedule.Text) & "' "
        
        .Filter = strQuery
        
        If .RecordCount > 0 Then
            .MoveFirst
        End If
        
        Do While Not .EOF
            .Delete
            .MoveNext
        Loop
    End With
    
    Rem Selected City
    strQuery = " SELECT radio_Area.*,city.* FROM radio_Area INNER JOIN City ON radio_Area.city_id=city.city_id WHERE Radio_Area_Code='" & Trim(cboRadioArea.Text) & "' AND Brand_Code ='" & Left(Frm_IB_Radio.cboBrand.Text, 4) & "'"
    
    recSelected.CursorLocation = adUseClient
    recSelected.Open strQuery, ConnERP, adOpenStatic, adLockReadOnly
    
    Do While recSelected.EOF = False
        strQuery = "IB_Id='" & Frm_IB_Radio.txtIBID.Text & "' and "
        strQuery = strQuery & "Month=" & Get_Month_Number(cboMonthlyBudget.Text) & " and "
        strQuery = strQuery & "Year=" & Val(Frm_IB_Radio.cboYear.Text) & " and "
        strQuery = strQuery & "Schedule='" & Clear_String(txtSchedule.Text) & "' and "
        strQuery = strQuery & "City_Code=" & recSelected.Fields("city_id").Value & " and "
        strQuery = strQuery & "Area_Code='" & cboRadioArea.Text & "' and "
        strQuery = strQuery & "Spot=" & IIf(txtSpot.Text = "", 0, txtSpot.Text) & " and "
        strQuery = strQuery & "Urban_Flag=" & IIf(IsNull(recSelected.Fields("urban_flag").Value) = True, 0, recSelected.Fields("urban_flag").Value) & " and "
        strQuery = strQuery & "rural_Flag=" & IIf(IsNull(recSelected.Fields("rural_flag").Value) = True, 0, recSelected.Fields("rural_flag").Value)
        Frm_IB_Radio.recPlanDetailTemp.Filter = strQuery
        
        If Not Frm_IB_Radio.recPlanDetailTemp.EOF And Not Frm_IB_Radio.recPlanDetailTemp.BOF Then
            'do nothing
        Else
            Frm_IB_Radio.recPlanDetailTemp.AddNew
            Frm_IB_Radio.recPlanDetailTemp.Fields("month").Value = Get_Month_Number(cboMonthlyBudget.Text)
            Frm_IB_Radio.recPlanDetailTemp.Fields("year").Value = Val(Frm_IB_Radio.cboYear.Text)
            Frm_IB_Radio.recPlanDetailTemp.Fields("IB_Id").Value = Frm_IB_Radio.txtIBID.Text
            Frm_IB_Radio.recPlanDetailTemp.Fields("Schedule").Value = txtSchedule.Text
            Frm_IB_Radio.recPlanDetailTemp.Fields("City_Code").Value = recSelected.Fields("city_id").Value
            Frm_IB_Radio.recPlanDetailTemp.Fields("Area_Code").Value = cboRadioArea.Text
            Frm_IB_Radio.recPlanDetailTemp.Fields("Urban_Flag").Value = recSelected.Fields("urban_flag").Value
            Frm_IB_Radio.recPlanDetailTemp.Fields("rural_Flag").Value = recSelected.Fields("rural_flag").Value
            Frm_IB_Radio.recPlanDetailTemp.Fields("Spot").Value = IIf(txtSpot.Text = "", 0, txtSpot.Text)
            Frm_IB_Radio.recPlanDetailTemp.Update
            lstSelectedCity.AddItem recSelected.Fields("city_id").Value & " " & recSelected.Fields("City").Value
        End If
        
        recSelected.MoveNext
    Loop
    
    Frm_IB_Radio.recPlanDetailTemp.Filter = ""
    
    Set dgdPlanDetailTemp.DataSource = Frm_IB_Radio.recPlanDetailTemp
    Me.MousePointer = vbDefault
    
    recSelected.Close
    Set recSelected = Nothing
End Sub

Private Sub cboRadioArea_DropDown()
    Dim recRadioArea As New ADODB.Recordset
    
    strQuery = "SELECT DISTINCT Radio_Area_Code FROM radio_area "
    strQuery = strQuery & " WHERE brand_code='" & Left(Frm_IB_Radio.cboBrand.Text, 4) & "'"
    recRadioArea.Open strQuery, ConnERP, adOpenStatic, adLockReadOnly
    cboRadioArea.Clear
    
    While Not recRadioArea.EOF And Not recRadioArea.BOF
        cboRadioArea.AddItem recRadioArea.Fields("Radio_Area_Code").Value
        recRadioArea.MoveNext
    Wend
    
    recRadioArea.Close
    Set recRadioArea = Nothing
End Sub

Private Sub FillData()
    Dim StrFilter As String
    Dim recGetCity As New ADODB.Recordset
    Dim blnIsExist As Boolean
    Dim intBarGrd As Integer
        
    Frm_IB_Radio.recPlanDetailTemp.Filter = ""
        
    StrFilter = "IB_Id='" & Frm_IB_Radio.txtIBID.Text & "' and "
    StrFilter = StrFilter & "Month=" & Get_Month_Number(cboMonthlyBudget.Text) & " and "
    StrFilter = StrFilter & "Year=" & Val(Frm_IB_Radio.cboYear.Text) & ""
        
    'isi list
    With Frm_IB_Radio
        If .cboPlanDetail.Text = "City" Then 'dw - must be splited cause they have different format saving
            StrFilter = StrFilter & " and Schedule='" & Clear_String(cboSchedule.Text) & "' "
        ElseIf .cboPlanDetail.Text = "Station" Then
            If cboSchedule.Text <> "" Then
                StrFilter = StrFilter & " and Schedule LIKE '%" & Clear_String(cboSchedule.Text) & "%' "
            End If
        End If
        
        .recPlanDetailTemp.Filter = StrFilter
        
        If Not .recPlanDetailTemp.EOF And Not .recPlanDetailTemp.BOF Then
            .recPlanDetailTemp.MoveFirst
            txtSpot.Text = .recPlanDetailTemp.Fields("spot").Value
            
            If Frm_IB_Radio.cboPlanDetail.Text = "City" Then
                cboRadioArea.Clear
                cboRadioArea.AddItem IIf(IsNull(.recPlanDetailTemp.Fields("area_code").Value), "", .recPlanDetailTemp.Fields("area_code").Value)
                cboRadioArea.Text = IIf(IsNull(.recPlanDetailTemp.Fields("area_code").Value), "", .recPlanDetailTemp.Fields("area_code").Value)
            
                lstSelectedCity.Clear
                
                While Not .recPlanDetailTemp.EOF And Not .recPlanDetailTemp.BOF
                    strQuery = "SELECT city FROM city WHERE city_id=" & .recPlanDetailTemp.Fields("city_code").Value
                    recGetCity.Open strQuery, ConnERP, adOpenStatic, adLockReadOnly
                    
                    lstSelectedCity.AddItem .recPlanDetailTemp.Fields("city_code").Value & " " & recGetCity(0)
                    .recPlanDetailTemp.MoveNext
                    recGetCity.Close
                Wend
                
            ElseIf Frm_IB_Radio.cboPlanDetail.Text = "Station" Then 'dw (01.08.2012)
                LoadStations
            End If
        End If
                
        .recPlanDetailMaterialTemp.Filter = StrFilter
        
        msgMix.Clear
        msgMix.Rows = 2
        msgMix.FormatString = " Material Id | Mix(%) "
        
        If Not .recPlanDetailMaterialTemp.EOF And Not .recPlanDetailMaterialTemp.BOF Then
            .recPlanDetailMaterialTemp.MoveFirst
        End If
        
        While Not .recPlanDetailMaterialTemp.EOF And Not .recPlanDetailMaterialTemp.BOF
            blnIsExist = False
            For intBarGrd = 1 To msgMix.Rows - 1
                If .recPlanDetailMaterialTemp.Fields("material_id").Value = msgMix.TextMatrix(intBarGrd, 0) And .recPlanDetailMaterialTemp.Fields("material_mix").Value = msgMix.TextMatrix(intBarGrd, 1) Then
                    blnIsExist = True
                    Exit For
                End If
            Next
            
            If blnIsExist = False Then
                If msgMix.TextMatrix(1, 0) = "" And msgMix.TextMatrix(1, 1) = "" Then
                    msgMix.TextMatrix(1, 0) = .recPlanDetailMaterialTemp.Fields("material_id").Value
                    msgMix.TextMatrix(1, 1) = .recPlanDetailMaterialTemp.Fields("material_mix").Value
                Else
                    msgMix.AddItem .recPlanDetailMaterialTemp.Fields("material_id").Value & vbTab & .recPlanDetailMaterialTemp.Fields("material_mix").Value
                End If
            End If
            
            .recPlanDetailMaterialTemp.MoveNext
        Wend
    End With
End Sub

Private Sub cboRadioArea_KeyPress(KeyAscii As Integer)
    KeyAscii = 0
End Sub

Private Sub cboSchedule_Click()
    
    ClearForm
    FillData
    fraSalesStation.Enabled = True

End Sub

Private Sub cboSchedule_DropDown()

    FillScheduleList

End Sub

Private Sub cboSchedule_KeyPress(KeyAscii As Integer)

    KeyAscii = 0

End Sub

Private Sub chkRural_Click()
    If blnFillPlanDetail = False Then Exit Sub
    
    If lstSelectedCity.ListIndex <> -1 Then
        Frm_IB_Radio.recPlanDetailTemp.Filter = ""
        
        strQuery = "IB_Id='" & Frm_IB_Radio.txtIBID.Text & "' AND "
        strQuery = strQuery & "Month=" & Get_Month_Number(cboMonthlyBudget.Text) & " AND "
        strQuery = strQuery & "Year=" & Val(Frm_IB_Radio.cboYear.Text) & " AND "
        strQuery = strQuery & "Schedule='" & Clear_String(txtSchedule.Text) & "' AND "
        strQuery = strQuery & "city_code=" & Val(lstSelectedCity.Text)
        
        Frm_IB_Radio.recPlanDetailTemp.Filter = ""
        Frm_IB_Radio.recPlanDetailTemp.Filter = strQuery
        
        If Not Frm_IB_Radio.recPlanDetailTemp.EOF Then
            Frm_IB_Radio.recPlanDetailTemp.Fields("Rural_Flag").Value = chkRural.Value
            Frm_IB_Radio.recPlanDetailTemp.Update
        End If
    End If
End Sub

Private Sub chkUrban_Click()
    If blnFillPlanDetail = False Then Exit Sub
    
    If lstSelectedCity.ListIndex <> -1 Then
        Frm_IB_Radio.recPlanDetailTemp.Filter = ""
        
        strQuery = "IB_Id='" & Frm_IB_Radio.txtIBID.Text & "' AND "
        strQuery = strQuery & "Month=" & Get_Month_Number(cboMonthlyBudget.Text) & " AND "
        strQuery = strQuery & "Year=" & Val(Frm_IB_Radio.cboYear.Text) & " AND "
        strQuery = strQuery & "Schedule='" & Clear_String(txtSchedule.Text) & "' AND "
        strQuery = strQuery & "city_code=" & Val(lstSelectedCity.Text)
        
        Frm_IB_Radio.recPlanDetailTemp.Filter = ""
        Frm_IB_Radio.recPlanDetailTemp.Filter = strQuery
        
        If Not Frm_IB_Radio.recPlanDetailTemp.EOF Then
            Frm_IB_Radio.recPlanDetailTemp.Fields("Urban_Flag").Value = chkUrban.Value
            Frm_IB_Radio.recPlanDetailTemp.Update
        End If
    End If
End Sub

Private Sub cmdAdd_Click()
    strQuery = " IB_ID ='" & Frm_IB_Radio.txtIBID.Text & "'"
    strQuery = strQuery & " AND Year = " & Val(Frm_IB_Radio.cboYear.Text)
    strQuery = strQuery & " AND Month = " & Get_Month_Number(cboMonthlyBudget.Text)
    
    Frm_IB_Radio.recRadioPlanTemp.Filter = ""
    Frm_IB_Radio.recRadioPlanTemp.Filter = strQuery
    
    If Frm_IB_Radio.recRadioPlanTemp.EOF Then
       MsgBox "No Budget for " & cboMonthlyBudget.Text & "  Please entry budget first..", vbExclamation, strTitleMissingInfo
       Frm_IB_Radio.recRadioPlanTemp.Filter = ""
       
       Call ViewGrdBudget
       
       Exit Sub
    End If
    pnlMain.Enabled = True
    Frm_IB_Radio.recRadioPlanTemp.Filter = ""
    
    Call ViewGrdBudget
    
    strTransacProcess = "ADD"
    
    If Frm_IB_Radio.cboPlanDetail.Text = "Station" Then
        LoadStations
    End If
    
    intJumMix = 0
    
    blnFillPlanDetail = True
    
    SwitchSchedule True
    
    blnFillMonthlyBudget = False
    
    Call ClearForm
    
    Call SetButton(True)
    
    txtSchedule.Enabled = True
    EnableObject True
    fraMonthlyBudget.Enabled = False
    
    txtSchedule.SetFocus
End Sub

Private Sub cmdCancelBudget_Click()
    'set object
    cboMonth.Enabled = False
    cboMonth.Text = ""
    medBudget.Enabled = False
    medBudget.Text = 0
    dgdMonthlyBudget.Enabled = True
    
    'set button
    cmdSaveBudget.Enabled = False
    cmdCancelBudget.Visible = False
    cmdEditBudget.Visible = True
    cmdDeleteBudget.Enabled = True
End Sub

Private Sub cmdCancel_Click()
    strTransacProcess = "NO DATA"

    If blnFillMonthlyBudget Then
        RestoreToOriginal
    Else
        DeleteNewData
    End If
    
    fraMonthlyBudget.Enabled = True
    
    blnFillMonthlyBudget = True
    
    blnFillPlanDetail = False
    
    SwitchSchedule False
    
    Call SetButton(False)
    
    txtSchedule.Enabled = True
    
    ClearForm
    
    FillScheduleList 'dw
    'pnlMain.Enabled = False
    EnableObject False
    Call EnableObject(False)
    FillData
    
End Sub

Private Sub DeleteNewData()
    Dim StrFilter As String
    Dim intSta As Integer
    
    StrFilter = "IB_Id='" & Frm_IB_Radio.txtIBID.Text & "' AND "
    StrFilter = StrFilter & "Month=" & Get_Month_Number(cboMonthlyBudget.Text) & " AND "
    StrFilter = StrFilter & "Year=" & Val(Frm_IB_Radio.cboYear.Text)
    
    If Frm_IB_Radio.cboPlanDetail.Text = "City" Then
        StrFilter = StrFilter & " AND Schedule='" & Clear_String(txtSchedule.Text) & "'"
    
        Frm_IB_Radio.recPlanDetailTemp.Filter = ""
        Frm_IB_Radio.recPlanDetailTemp.Filter = StrFilter
        
        While Not Frm_IB_Radio.recPlanDetailTemp.EOF And Not Frm_IB_Radio.recPlanDetailTemp.BOF
            Frm_IB_Radio.recPlanDetailTemp.Delete
            Frm_IB_Radio.recPlanDetailTemp.MoveNext
        Wend
        
        Frm_IB_Radio.recPlanDetailMaterialTemp.Filter = ""
        Frm_IB_Radio.recPlanDetailMaterialTemp.Filter = StrFilter
        
        While Not Frm_IB_Radio.recPlanDetailMaterialTemp.EOF And Not Frm_IB_Radio.recPlanDetailMaterialTemp.BOF
            Frm_IB_Radio.recPlanDetailMaterialTemp.Delete
            Frm_IB_Radio.recPlanDetailMaterialTemp.MoveNext
        Wend
    ElseIf Frm_IB_Radio.cboPlanDetail.Text = "Station" Then
        If lstSelectedStation.ListCount <> 0 Then
            Frm_IB_Radio.recPlanDetailTemp.Filter = ""
            Frm_IB_Radio.recPlanDetailTemp.Filter = StrFilter
            
            While Not Frm_IB_Radio.recPlanDetailTemp.EOF And Not Frm_IB_Radio.recPlanDetailTemp.BOF
                Frm_IB_Radio.recPlanDetailTemp.MoveFirst
            
                For intSta = 0 To lstSelectedStation.ListCount - 1
                    Frm_IB_Radio.recPlanDetailTemp.Filter = StrFilter & " AND Schedule='" & txtSchedule.Text & " - " & Trim(Mid(lstSelectedStation.List(intSta), InStr(lstSelectedStation.List(intSta), "-") + 1, Len(lstSelectedStation.List(intSta)) - InStr(lstSelectedStation.List(intSta), "-") + 1)) & "'"
                    
                    While Not Frm_IB_Radio.recPlanDetailTemp.EOF And Not Frm_IB_Radio.recPlanDetailTemp.BOF
                        Frm_IB_Radio.recPlanDetailTemp.Delete
                        Frm_IB_Radio.recPlanDetailTemp.MoveNext
                    Wend
                Next
            Wend
            
            Frm_IB_Radio.recPlanDetailMaterialTemp.Filter = ""
            Frm_IB_Radio.recPlanDetailMaterialTemp.Filter = StrFilter
            
            While Not Frm_IB_Radio.recPlanDetailMaterialTemp.EOF And Not Frm_IB_Radio.recPlanDetailMaterialTemp.BOF
                Frm_IB_Radio.recPlanDetailMaterialTemp.MoveFirst
                
                For intSta = 0 To lstSelectedStation.ListCount - 1
                    Frm_IB_Radio.recPlanDetailMaterialTemp.Filter = StrFilter & " AND Schedule='" & txtSchedule.Text & " - " & Trim(Mid(lstSelectedStation.List(intSta), InStr(lstSelectedStation.List(intSta), "-") + 1, Len(lstSelectedStation.List(intSta)) - InStr(lstSelectedStation.List(intSta), "-") + 1)) & "'"
                
                    While Not Frm_IB_Radio.recPlanDetailMaterialTemp.EOF And Not Frm_IB_Radio.recPlanDetailMaterialTemp.BOF
                        Frm_IB_Radio.recPlanDetailMaterialTemp.Delete
                        Frm_IB_Radio.recPlanDetailMaterialTemp.MoveNext
                    Wend
                Next
            Wend
        End If
    End If
End Sub

Private Sub cmdClose_Click()
    Unload Me
End Sub

Private Sub cmdDeleteBudget_Click()
    If cboMonth.Text = "" Then
        MsgBox "Please Select Data !", vbExclamation, strTitleMissingInfo
        Exit Sub
    End If
      
    'cek di detail
    strQuery = " IB_ID ='" & Frm_IB_Radio.txtIBID.Text & "'"
    strQuery = strQuery & " AND Year = " & Val(Frm_IB_Radio.cboYear.Text)
    strQuery = strQuery & " AND Month = " & Get_Month_Number(cboMonth.Text)
    
    Frm_IB_Radio.recPlanDetailTemp.Filter = ""
    Frm_IB_Radio.recPlanDetailTemp.Filter = strQuery
    
    If Not Frm_IB_Radio.recPlanDetailTemp.EOF Then
        Frm_IB_Radio.recPlanDetailTemp.Filter = ""
        
        MsgBox "Please Delete Detail City or Station First !", vbExclamation, strTitleInfo
        Exit Sub
    End If
    
    If MsgBox(strMsgDeleteConfirm, vbYesNo, strTitleConfirm) = vbYes Then
        strQuery = " IB_ID ='" & Frm_IB_Radio.txtIBID.Text & "'"
        strQuery = strQuery & " AND Year = " & Val(Frm_IB_Radio.cboYear.Text)
        strQuery = strQuery & " AND Month = " & Get_Month_Number(cboMonth.Text)
        Frm_IB_Radio.recRadioPlanTemp.Filter = ""
        Frm_IB_Radio.recRadioPlanTemp.Filter = strQuery
        
        If Not Frm_IB_Radio.recRadioPlanTemp.EOF Then
            Frm_IB_Radio.recRadioPlanTemp.Delete
        End If
        
        cboMonth.Text = ""
        medBudget.Text = 0
        
        Call ViewGrdBudget
        
        Me.MousePointer = vbDefault
        MsgBox strMsgDeleteDataDone, vbInformation, strTitleInfo
    End If
End Sub

Private Sub cmdDelete_Click()
    If cboSchedule.Text = "" Then
        MsgBox "Please Choose the Schedule First", vbCritical, strTitleMissingInfo
        Exit Sub
    End If

    If MsgBox(strMsgDeleteConfirm, vbYesNo, strTitleConfirm) = vbYes Then
        With Frm_IB_Radio
            .recPlanDetailMaterialTemp.Filter = ""
            strQuery = "IB_Id='" & Frm_IB_Radio.txtIBID.Text & "' AND "
            strQuery = strQuery & "Month=" & Get_Month_Number(cboMonthlyBudget.Text) & " AND "
            strQuery = strQuery & "Year=" & Val(Frm_IB_Radio.cboYear.Text) & " AND "
            
            If Frm_IB_Radio.cboPlanDetail.Text = "City" Then 'dw - is splited cause they have different format
                strQuery = strQuery & "Schedule='" & Clear_String(cboSchedule.Text) & "' "
            ElseIf Frm_IB_Radio.cboPlanDetail.Text = "Station" Then
                strQuery = strQuery & "Schedule LIKE '%" & Clear_String(cboSchedule.Text) & "%' "
            End If
            
            .recPlanDetailMaterialTemp.Filter = strQuery
            
            While Not .recPlanDetailMaterialTemp.EOF And Not .recPlanDetailMaterialTemp.BOF
                .recPlanDetailMaterialTemp.Delete
                .recPlanDetailMaterialTemp.MoveNext
            Wend
            
            .recPlanDetailTemp.Filter = ""
            strQuery = "IB_Id='" & Frm_IB_Radio.txtIBID.Text & "' AND "
            strQuery = strQuery & "Month=" & Get_Month_Number(cboMonthlyBudget.Text) & " AND "
            strQuery = strQuery & "Year=" & Val(Frm_IB_Radio.cboYear.Text) & " AND "
            
            If Frm_IB_Radio.cboPlanDetail.Text = "City" Then 'dw - is splited cause they have different format
                strQuery = strQuery & "Schedule='" & Clear_String(cboSchedule.Text) & "' "
            ElseIf Frm_IB_Radio.cboPlanDetail.Text = "Station" Then
                strQuery = strQuery & "Schedule LIKE '%" & Clear_String(cboSchedule.Text) & "%' "
            End If
                
            .recPlanDetailTemp.Filter = strQuery
            
            While Not .recPlanDetailTemp.EOF And Not .recPlanDetailTemp.BOF
                .recPlanDetailTemp.Delete
                .recPlanDetailTemp.MoveNext
            Wend
            
            ClearForm
            
            cboSchedule.Clear
            
            Me.MousePointer = vbDefault
            MsgBox strMsgDeleteDataDone, vbInformation, strTitleInfo
        End With
    End If
End Sub

Private Sub cmdDeleteMix_Click()
    If msgMix.TextMatrix(1, 0) = "" Or msgMix.TextMatrix(1, 1) = "" Then
        MsgBox "There is no data can be deleted", vbCritical, strTitleMissingInfo
    Else
        If msgMix.Rows > 2 Then
            msgMix.RemoveItem msgMix.Row
        Else
            msgMix.TextMatrix(1, 0) = ""
            msgMix.TextMatrix(1, 1) = ""
        End If
    End If
End Sub

Private Sub cmdEditBudget_Click()
    'untuk menambah/mengubah plan di IB
    If Not Frm_IB_Radio.recRadioPlanTemp.EOF Then
        dgdMonthlyBudget.Enabled = False
        
        cboMonth.Enabled = True
        cboMonth.Text = Get_Month_Name(Frm_IB_Radio.recRadioPlanTemp("Month").Value)
        
        medBudget.Enabled = True
        medBudget.Text = Frm_IB_Radio.recRadioPlanTemp("Budget").Value
    Else
        cboMonth.Enabled = True
        medBudget.Enabled = True
        medBudget.Text = 0
    End If
    
    'set button
    cmdSaveBudget.Enabled = True
    cmdCancelBudget.Visible = True
    cmdEditBudget.Visible = False
    cmdDeleteBudget.Enabled = False
End Sub

Private Sub cmdEdit_Click()
    If cboSchedule.Text = "" Then
        MsgBox "Please Choose the Schedule First", vbCritical, strTitleMissingInfo
        Exit Sub
    End If
    pnlMain.Enabled = True
    PrepareEditDummy
    
    blnFillPlanDetail = True
    SwitchSchedule True
    blnFillMonthlyBudget = True
    
    Call SetButton(True)
    
    fraMonthlyBudget.Enabled = False
    EnableObject True
    txtSchedule.Text = cboSchedule.Text
    txtSchedule.Enabled = False

End Sub


Private Sub cmdRemoveCity_Click()
    Dim StrFilter As String
    
    If lstSelectedCity.ListIndex <> -1 Then
        StrFilter = " month =" & Get_Month_Number(cboMonthlyBudget.Text)
        StrFilter = StrFilter & " and year =" & Val(Frm_IB_Radio.cboYear.Text)
        StrFilter = StrFilter & " and IB_Id='" & Frm_IB_Radio.txtIBID.Text & "'"
        StrFilter = StrFilter & " and CIty_Code=" & Val(lstSelectedCity.Text)
        StrFilter = StrFilter & " and Schedule='" & Clear_String(txtSchedule.Text) & "'"
        
        Frm_IB_Radio.recPlanDetailTemp.Filter = ""
        Frm_IB_Radio.recPlanDetailTemp.Filter = StrFilter
        
        'DELETE CITY
        While Not Frm_IB_Radio.recPlanDetailTemp.EOF
            Frm_IB_Radio.recPlanDetailTemp.Delete
            Frm_IB_Radio.recPlanDetailTemp.MoveNext
        Wend
        
        'material mix
        Frm_IB_Radio.recPlanDetailMaterialTemp.Filter = ""
        Frm_IB_Radio.recPlanDetailMaterialTemp.Filter = StrFilter
        
        While Not Frm_IB_Radio.recPlanDetailMaterialTemp.EOF
            Frm_IB_Radio.recPlanDetailMaterialTemp.Delete
            Frm_IB_Radio.recPlanDetailMaterialTemp.MoveNext
        Wend
        
        lstSelectedCity.RemoveItem lstSelectedCity.ListIndex
        Set dgdPlanDetailTemp.DataSource = Frm_IB_Radio.recPlanDetailTemp
    End If
End Sub

Private Sub cmdSaveBudget_Click()
    'validasi
    If cboMonth.Text = "" Then
        MsgBox "Please Select Month Budget !", vbExclamation, strTitleMissingInfo
        Exit Sub
    End If
    
    If medBudget.Text = "" Or medBudget.Text = 0 Then
        MsgBox "Please Insert Budget !", vbExclamation, strTitleMissingInfo
        Exit Sub
    End If
    
    With Frm_IB_Radio.recRadioPlanTemp
        .Filter = ""
        
        'save data to temp rec
        strQuery = " IB_ID ='" & Frm_IB_Radio.txtIBID.Text & "'"
        strQuery = strQuery & " AND Year = " & Val(Frm_IB_Radio.cboYear.Text)
        strQuery = strQuery & " AND Month = " & Get_Month_Number(cboMonth.Text)
        .Filter = strQuery
        
        If .EOF Then
            .AddNew
        End If
        
       .Fields("IB_ID").Value = Frm_IB_Radio.txtIBID.Text
       .Fields("Year").Value = Frm_IB_Radio.cboYear.Text
       .Fields("Month").Value = Get_Month_Number(cboMonth.Text)
       .Fields("Budget").Value = Val(medBudget.Text)
       .Update
    End With
    
    Call ViewGrdBudget
    
    'set object
    cboMonth.Enabled = False
    cboMonth.Text = ""
    medBudget.Enabled = False
    medBudget.Text = 0
    dgdMonthlyBudget.Enabled = True
    
    'set button
    cmdSaveBudget.Enabled = False
    cmdCancelBudget.Visible = False
    cmdEditBudget.Visible = True
    cmdDeleteBudget.Enabled = True
    
    Me.MousePointer = vbDefault
    MsgBox strMsgSaveDataDone, vbInformation, strTitleInfo
    
    cmdAdd.SetFocus
    fraSchedule.Enabled = True

End Sub

Private Sub cmdSave_Click()
    Dim intbar As Integer
    Dim dblJmlMix As Double
    
    'dw - Validate txtbox
    If txtSchedule.Text = "" Then
        MsgBox "Fill Schedule First", vbCritical, strTitleMissingInfo
        txtSchedule.SetFocus
        Exit Sub
    End If
    
    If txtSpot.Text = "" Then
        MsgBox "Fill Spot First", vbCritical, strTitleMissingInfo
        txtSpot.SetFocus
        Exit Sub
    End If
    
    With Frm_IB_Radio.recPlanDetailTemp
        strQuery = "IB_Id='" & Frm_IB_Radio.txtIBID.Text & "' AND "
        strQuery = strQuery & "Month=" & Get_Month_Number(cboMonthlyBudget.Text) & " AND "
        strQuery = strQuery & "Year=" & Val(Frm_IB_Radio.cboYear.Text) & " "
        
        If Frm_IB_Radio.cboPlanDetail.Text = "City" Then 'dw - is splited cause they have different format saving
            strQuery = strQuery & "AND Schedule='" & Clear_String(txtSchedule.Text) & "' "
        ElseIf Frm_IB_Radio.cboPlanDetail.Text = "Station" Then
            strQuery = strQuery & "AND Schedule LIKE '%" & Clear_String(txtSchedule.Text) & "%' "
        End If
        
        .Filter = strQuery
        
        If Not .EOF And Not .BOF Then
            'do nothing
        Else
            If Frm_IB_Radio.cboPlanDetail.Text = "City" Then
                MsgBox "Fill City First", vbCritical, strTitleMissingInfo
                Exit Sub
            ElseIf Frm_IB_Radio.cboPlanDetail.Text = "Station" Then 'dw
                MsgBox "Select Station First", vbCritical, strTitleMissingInfo
                Exit Sub
            End If
        End If
    End With
    
    If msgMix.TextMatrix(1, 0) = "" And msgMix.TextMatrix(1, 1) = "" Then
        MsgBox "Fill Material Mix First", vbCritical, strTitleMissingInfo
        Exit Sub
    Else
        dblJmlMix = 0
        For intbar = 1 To msgMix.Rows - 1
            dblJmlMix = dblJmlMix + msgMix.TextMatrix(intbar, 1)
        Next
        
        If dblJmlMix < Mix_Max Then
            MsgBox "Material Mix must be in 100", vbCritical, strTitleInfo
            Exit Sub
        End If
    End If
    
    Call SaveData
    
    blnFillMonthlyBudget = True
    
    Call SetButton(False)
    
    cboSchedule_DropDown
    
    cboSchedule.Text = txtSchedule.Text
    
    Call ClearForm
    
    fraMonthlyBudget.Enabled = True
    
    FillData
    
    blnFillPlanDetail = False
    
    SwitchSchedule False
    'pnlMain.Enabled = False
    Me.MousePointer = vbDefault
    MsgBox strMsgSaveDataDone, vbInformation, strTitleInfo
    EnableObject False
    
End Sub

Private Sub cmdAddMix_Click()
    Dim intLp As Integer
    
    If cboMaterial.Text = "" Then
        MsgBox "You Must Fill Material", vbCritical, strTitleMissingInfo
        Exit Sub
    End If
    
    With msgMix
        .FormatString = " Material Id | Mix(%) "
        If .Rows > 0 Then
            intJumMix = 0
            For intLp = 1 To .Rows - 1
              intJumMix = intJumMix + IIf(.TextMatrix(intLp, 1) = "", 0, .TextMatrix(intLp, 1))
            Next
            
            intJumMix = intJumMix + IIf(txtMix.Text = "", 0, txtMix.Text)
            If intJumMix > 100 Then
              MsgBox "Mix More Than 100", vbCritical, strTitleExclamation
              Exit Sub
            End If
            
            If .TextMatrix(1, 0) = "" And .TextMatrix(1, 1) = "" Then
                .TextMatrix(1, 0) = Left(cboMaterial.Text, 1)
                .TextMatrix(1, 1) = IIf(txtMix.Text = "", 0, txtMix.Text)
            Else
                .AddItem Left(cboMaterial.Text, 1) & vbTab & IIf(txtMix.Text = "", 0, txtMix.Text)
            End If
            
            cboMaterial.Clear
            txtMix.Text = 0
       End If
    End With
End Sub

Private Sub cmdManageArea_Click()
    Frm_IB_Radio_Plan_Area_New.show vbModal
End Sub

Private Sub ViewGrdBudget()
    Dim clmGridColumn
    
    Frm_IB_Radio.recRadioPlanTemp.Filter = ""
    Set dgdMonthlyBudget.DataSource = Frm_IB_Radio.recRadioPlanTemp
    
    For Each clmGridColumn In dgdMonthlyBudget.Columns
        Select Case clmGridColumn.DataField
            Case "Budget":
                 clmGridColumn.NumberFormat = "#,##0"
                 clmGridColumn.Alignment = dbgRight
            Case "Month":
                 clmGridColumn.Alignment = dbgRight
            Case Else:
                clmGridColumn.Width = 0
        End Select
    Next
End Sub

Private Sub dgdMonthlyBudget_Click()
    If Not Frm_IB_Radio.recRadioPlanTemp.EOF Then
        cboMonth.Text = Get_Month_Name(Frm_IB_Radio.recRadioPlanTemp("Month").Value)
        medBudget.Text = Frm_IB_Radio.recRadioPlanTemp("Budget").Value
    End If
End Sub

Private Sub FillScheduleList()
    Dim intLst As Integer
    Dim blnIsEksis As Boolean
    
    strQuery = " IB_ID ='" & Frm_IB_Radio.txtIBID.Text & "'"
    strQuery = strQuery & " AND Year = " & Val(Frm_IB_Radio.cboYear.Text)
    strQuery = strQuery & " AND Month = " & Get_Month_Number(cboMonthlyBudget.Text)
    
    Frm_IB_Radio.recPlanDetailTemp.Filter = ""
    Frm_IB_Radio.recPlanDetailTemp.Filter = strQuery
    
    cboSchedule.Clear
    
    While Not Frm_IB_Radio.recPlanDetailTemp.EOF And Not Frm_IB_Radio.recPlanDetailTemp.BOF
        blnIsEksis = False
        For intLst = 0 To cboSchedule.ListCount - 1
            If Frm_IB_Radio.cboPlanDetail.Text = "City" Then 'dw
                If Frm_IB_Radio.recPlanDetailTemp.Fields("schedule").Value = cboSchedule.List(intLst) Then
                    blnIsEksis = True
                    Exit For
                End If
            ElseIf Frm_IB_Radio.cboPlanDetail.Text = "Station" Then 'dw - clear the format to prevent looping
                If Trim(Mid(Frm_IB_Radio.recPlanDetailTemp.Fields("Schedule").Value, 1, InStr(Frm_IB_Radio.recPlanDetailTemp.Fields("Schedule").Value, "-") - 1)) = cboSchedule.List(intLst) Then
                    blnIsEksis = True
                    Exit For
                End If
            End If
        Next
        
        If blnIsEksis = False Then
            If Frm_IB_Radio.cboPlanDetail.Text = "City" Then 'dw
                cboSchedule.AddItem Frm_IB_Radio.recPlanDetailTemp.Fields("schedule").Value
            ElseIf Frm_IB_Radio.cboPlanDetail.Text = "Station" Then 'dw - add to cboSchedule without the format saving
                cboSchedule.AddItem Trim(Mid(Frm_IB_Radio.recPlanDetailTemp.Fields("Schedule").Value, 1, InStr(Frm_IB_Radio.recPlanDetailTemp.Fields("Schedule").Value, "-") - 1))
            End If
        End If
        
        Frm_IB_Radio.recPlanDetailTemp.MoveNext
    Wend
    
    If cboSchedule.ListCount <> 0 Then
        cboSchedule.ListIndex = 0
    End If
    
    Frm_IB_Radio.recPlanDetailTemp.Filter = ""
End Sub

Private Sub SwitchSchedule(blnStatus As Boolean)
    If blnStatus Then
        txtSchedule.Visible = True
        cboSchedule.Visible = False
    Else
        cboSchedule.Visible = True
        txtSchedule.Visible = False
    End If
End Sub

Private Sub LoadMateri()
    If Frm_IB_Radio.recMateriTemp.RecordCount > 0 Then
        Frm_IB_Radio.recMateriTemp.MoveFirst
    End If
    
    cboMaterial.Clear
    Do While Frm_IB_Radio.recMateriTemp.EOF = False
        cboMaterial.AddItem Frm_IB_Radio.recMateriTemp.Fields("Material_ID").Value & " " & Frm_IB_Radio.recMateriTemp.Fields("Material_Name").Value & ":" & Frm_IB_Radio.recMateriTemp.Fields("Duration").Value
        Frm_IB_Radio.recMateriTemp.MoveNext
    Loop
End Sub

Private Sub SetButton(blnEnable As Boolean)
    fraArea.Enabled = blnEnable
    
    cboMonthlyBudget.Enabled = Not blnEnable
    
    cmdAdd.Visible = Not blnEnable
    cmdEdit.Visible = Not blnEnable
    cmdSave.Visible = blnEnable
    cmdCancel.Visible = blnEnable
    cmdDelete.Enabled = Not blnEnable
    txtSpot.Enabled = blnEnable
    
    cboMaterial.Enabled = blnEnable
    cmdAddMix.Enabled = blnEnable
    cmdDeleteMix.Enabled = blnEnable
    cmdRemoveCity.Enabled = blnEnable
    chkUrban.Enabled = blnEnable
    chkRural.Enabled = blnEnable
    
    'dw - Setbutton untuk fraStation
    treStations.Enabled = blnEnable
    treSelectedStation.Enabled = blnEnable
    '--------------------------------
    
    cmdClose.Enabled = Not blnEnable
End Sub

Private Sub Form_Unload(IntCancel As Integer)
    CloseDummy
    
    Call Frm_IB_Radio.CalTotal
End Sub

Private Sub lstSelectedCity_Click()
    Dim StrFilter As String
    
    If lstSelectedCity.ListIndex <> -1 Then
        If Frm_IB_Radio.recPlanDetailTemp.RecordCount > 0 Then
            Frm_IB_Radio.recPlanDetailTemp.MoveFirst
        End If
        
        StrFilter = ""
        
        Frm_IB_Radio.recPlanDetailTemp.Filter = StrFilter
        
        StrFilter = " month =" & Get_Month_Number(cboMonthlyBudget.Text)
        StrFilter = StrFilter & " and year =" & Val(Frm_IB_Radio.cboYear.Text)
        StrFilter = StrFilter & " and IB_Id='" & Frm_IB_Radio.txtIBID.Text & "'"
        StrFilter = StrFilter & " and CIty_Code=" & Val(lstSelectedCity.Text)
        
        If blnFillPlanDetail = True Then
            StrFilter = StrFilter & " and Schedule='" & Clear_String(txtSchedule.Text) & "'"
        Else
            StrFilter = StrFilter & " and Schedule='" & Clear_String(cboSchedule.Text) & "'"
        End If
        
        Frm_IB_Radio.recPlanDetailTemp.Filter = ""
        Frm_IB_Radio.recPlanDetailTemp.Filter = StrFilter
        
        Set dgdPlanDetailTemp.DataSource = Frm_IB_Radio.recPlanDetailTemp
       
        If Frm_IB_Radio.recPlanDetailTemp.EOF = False Then
            chkUrban.Value = IIf(IsNull(Frm_IB_Radio.recPlanDetailTemp.Fields("Urban_Flag").Value) = True, 0, Frm_IB_Radio.recPlanDetailTemp.Fields("Urban_Flag").Value)
            chkRural.Value = IIf(IsNull(Frm_IB_Radio.recPlanDetailTemp.Fields("Rural_flag").Value) = True, 0, Frm_IB_Radio.recPlanDetailTemp.Fields("Rural_flag").Value)
        End If
    End If
End Sub

Private Sub SaveData()
    Dim intLst As Integer
    Dim intGrd As Integer
    
    With Frm_IB_Radio.recPlanDetailMaterialTemp
        strQuery = "IB_Id='" & Frm_IB_Radio.txtIBID.Text & "' AND "
        strQuery = strQuery & "Month=" & Get_Month_Number(cboMonthlyBudget.Text) & " AND "
        strQuery = strQuery & "Year=" & Val(Frm_IB_Radio.cboYear.Text) & " "
        
        If Frm_IB_Radio.cboPlanDetail.Text = "City" Then 'dw - is splited cause they have different format saving
            strQuery = strQuery & "AND Schedule='" & Clear_String(txtSchedule.Text) & "' "
        ElseIf Frm_IB_Radio.cboPlanDetail.Text = "Station" Then
            strQuery = strQuery & "AND Schedule LIKE '%" & Clear_String(txtSchedule.Text) & "%' "
        End If
        
        .Filter = strQuery
                
        If .RecordCount > 0 Then
            .MoveFirst
        End If
        
        While Not .EOF And Not .BOF
            .Delete
            .MoveNext
        Wend
           
        'dw - Save ke recTemp Radio Plan Material first
        If Frm_IB_Radio.cboPlanDetail.Text = "City" Then 'for IB by City
            For intLst = 0 To lstSelectedCity.ListCount - 1
                For intGrd = 1 To msgMix.Rows - 1
                    .AddNew
                    .Fields("IB_Id").Value = Trim(Frm_IB_Radio.txtIBID.Text)
                    .Fields("month").Value = Get_Month_Number(cboMonthlyBudget.Text)
                    .Fields("year").Value = Val(Frm_IB_Radio.cboYear.Text)
                    .Fields("Schedule").Value = Clear_String(txtSchedule.Text)
                    .Fields("City_Code").Value = Val(lstSelectedCity.List(intLst))
                    .Fields("material_Id").Value = msgMix.TextMatrix(intGrd, 0)
                    .Fields("material_Mix").Value = msgMix.TextMatrix(intGrd, 1)
                    .Update
                Next
            Next
        ElseIf Frm_IB_Radio.cboPlanDetail.Text = "Station" Then 'for IB by Station
            For intLst = 0 To lstSelectedStation.ListCount - 1
                For intGrd = 1 To msgMix.Rows - 1
                    .AddNew
                    .Fields("IB_Id").Value = Trim(Frm_IB_Radio.txtIBID.Text)
                    .Fields("month").Value = Get_Month_Number(cboMonthlyBudget.Text)
                    .Fields("year").Value = Val(Frm_IB_Radio.cboYear.Text)
                    .Fields("Schedule").Value = Clear_String(txtSchedule.Text) & " - " & Trim(Mid(lstSelectedStation.List(intLst), InStr(lstSelectedStation.List(intLst), "-") + 1, Len(lstSelectedStation.List(intLst)) - InStr(lstSelectedStation.List(intLst), "-") + 1))
                    .Fields("City_Code").Value = Val(lstSelectedStation.List(intLst))
                    .Fields("material_Id").Value = msgMix.TextMatrix(intGrd, 0)
                    .Fields("material_Mix").Value = msgMix.TextMatrix(intGrd, 1)
                    .Update
                Next
            Next
        End If
    End With
    
    Frm_IB_Radio.recPlanDetailMaterialTemp.Filter = ""
    Set dgdPlanDetailTemp.DataSource = Frm_IB_Radio.recPlanDetailMaterialTemp
End Sub

Private Sub ClearForm()
    lstSelectedCity.Clear
    
    msgMix.Rows = 2
    msgMix.Clear
    
    txtSchedule.Text = Empty
    txtSpot.Text = Empty
    txtMix.Text = Empty
    
    cboRadioArea.Clear
    chkUrban.Value = 0
    chkRural.Value = 0
End Sub

Private Sub lstSelectedCity_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If lstSelectedCity.ListCount > 0 And blnFillPlanDetail = True Then
        If Button = 2 Then
            PopupMenu mnuKanan
        End If
    End If
End Sub

Private Sub MnuRemoveAll_Click()
    Dim StrFilter As String

    StrFilter = " month =" & Get_Month_Number(cboMonthlyBudget.Text)
    StrFilter = StrFilter & " AND year =" & Val(Frm_IB_Radio.cboYear.Text)
    StrFilter = StrFilter & " AND IB_Id='" & Frm_IB_Radio.txtIBID.Text & "'"
    StrFilter = StrFilter & " AND Schedule='" & Clear_String(txtSchedule.Text) & "'"
    
    'CITY
    Frm_IB_Radio.recPlanDetailTemp.Filter = ""
    Frm_IB_Radio.recPlanDetailTemp.Filter = StrFilter
    
    While Not Frm_IB_Radio.recPlanDetailTemp.EOF And Not Frm_IB_Radio.recPlanDetailTemp.BOF
        Frm_IB_Radio.recPlanDetailTemp.Delete
        Frm_IB_Radio.recPlanDetailTemp.MoveNext
    Wend
    
    Frm_IB_Radio.recPlanDetailTemp.Filter = ""
    
    'MATERIAL
    Frm_IB_Radio.recPlanDetailMaterialTemp.Filter = ""
    Frm_IB_Radio.recPlanDetailMaterialTemp.Filter = StrFilter
    
    While Not Frm_IB_Radio.recPlanDetailMaterialTemp.EOF
        Frm_IB_Radio.recPlanDetailMaterialTemp.Delete
        Frm_IB_Radio.recPlanDetailMaterialTemp.MoveNext
    Wend
    
    Frm_IB_Radio.recPlanDetailMaterialTemp.Filter = ""
    
    lstSelectedCity.Clear
    chkUrban.Value = 0
    chkRural.Value = 0
    cboRadioArea.Clear
    msgMix.Rows = 2
    msgMix.Clear
    
    Set dgdPlanDetailTemp.DataSource = Frm_IB_Radio.recPlanDetailTemp
End Sub

Private Sub medBudget_KeyPress(KeyAscii As Integer)
    If (Chr(KeyAscii) < "0" Or Chr(KeyAscii) > "9") And KeyAscii <> 8 And KeyAscii <> 13 And KeyAscii <> 9 And Chr(KeyAscii) <> "." Then
        KeyAscii = 0
        Beep
        Exit Sub
    End If
End Sub

Private Sub treSelectedStation_DblClick() 'dw - want to remove station
    If treSelectedStation.SelectedItem.Children = 0 Then
        mnuStatDelete_Click
    End If
End Sub

Private Sub treStations_DblClick() 'dw - want to add station
    If treStations.SelectedItem.Children = 0 Then
        Call mnuStatAdd_Click
    End If
End Sub

Private Sub txtMix_KeyPress(KeyAscii As Integer)
    If (Chr(KeyAscii) < "0" Or Chr(KeyAscii) > "9") And KeyAscii <> 8 And KeyAscii <> 13 And KeyAscii <> 9 And Chr(KeyAscii) <> "." Then
        KeyAscii = 0
        Beep
        Exit Sub
    End If
End Sub

Private Sub txtSchedule_GotFocus()
    strOldSchedule = UCase(Trim(txtSchedule.Text))
    txtSchedule.SelStart = 0
    txtSchedule.SelLength = Len(txtSchedule.Text)
End Sub

Private Sub txtSchedule_KeyPress(KeyAscii As Integer)
    If Chr(KeyAscii) = "-" Then 'dw - validate txtSchedule for string "-",cause it used for format saving
        KeyAscii = 0
        Beep
        Exit Sub
    End If
End Sub

Private Sub txtSchedule_LostFocus()
    Dim intSta As Integer 'dw
    Dim intSch As Integer 'dw
    Dim strCurrentSchedule As String 'dw

    Frm_IB_Radio.recPlanDetailTemp.Filter = ""
    
    strQuery = "IB_Id='" & Frm_IB_Radio.txtIBID.Text & "' and "
    strQuery = strQuery & "Month=" & Get_Month_Number(cboMonthlyBudget.Text) & " and "
    strQuery = strQuery & "Year=" & Val(Frm_IB_Radio.cboYear.Text)
    
    If Frm_IB_Radio.cboPlanDetail.Text = "City" Then 'dw - is splited cause they have different format
        strQuery = strQuery & " AND Schedule='" & Clear_String(txtSchedule.Text) & "' "
    ElseIf Frm_IB_Radio.cboPlanDetail.Text = "Station" Then
        If txtSchedule.Text <> "" Then
            strQuery = strQuery & " AND Schedule LIKE '%" & Clear_String(txtSchedule.Text) & "%' "
        End If
    End If
    
    Frm_IB_Radio.recPlanDetailTemp.Filter = strQuery
    
    If Not Frm_IB_Radio.recPlanDetailTemp.EOF And Not Frm_IB_Radio.recPlanDetailTemp.BOF Then
        If Not Frm_IB_Radio.recPlanDetailTemp.EOF And Not Frm_IB_Radio.recPlanDetailTemp.BOF Then
            If (blnFillMonthlyBudget = False And strOldSchedule <> UCase(Trim(txtSchedule.Text))) Or (blnFillMonthlyBudget And strOldSchedule <> UCase(Trim(txtSchedule.Text))) Then
                MsgBox "Schedule " & txtSchedule.Text & " for this month is exists", vbCritical, strTitleInfo
                txtSchedule.Text = strOldSchedule
                txtSchedule.SetFocus
                Exit Sub
            End If
        End If
    End If
    
    Frm_IB_Radio.recPlanDetailTemp.Filter = ""
    
    strQuery = "IB_Id='" & Frm_IB_Radio.txtIBID.Text & "' and "
    strQuery = strQuery & "Month=" & Get_Month_Number(cboMonthlyBudget.Text) & " and "
    strQuery = strQuery & "Year=" & Val(Frm_IB_Radio.cboYear.Text)
    
    If Frm_IB_Radio.cboPlanDetail.Text = "City" Then 'dw - is splited cause they have different format
        strQuery = strQuery & " and Schedule='" & strOldSchedule & "'"
    ElseIf Frm_IB_Radio.cboPlanDetail.Text = "Station" Then
        If txtSchedule.Text <> "" And strOldSchedule <> "" Then
            strQuery = strQuery & " and Schedule LIKE '%" & strOldSchedule & "%'"
        End If
    End If
    
    'city
    Frm_IB_Radio.recPlanDetailTemp.Filter = strQuery
    Me.MousePointer = vbHourglass
    
    Do While Not Frm_IB_Radio.recPlanDetailTemp.EOF And Not Frm_IB_Radio.recPlanDetailTemp.BOF
        If Frm_IB_Radio.cboPlanDetail.Text = "City" Then
            Frm_IB_Radio.recPlanDetailTemp.Fields("Schedule").Value = Trim(txtSchedule.Text)
            
            Frm_IB_Radio.recPlanDetailTemp.Update
            Frm_IB_Radio.recPlanDetailTemp.MoveNext

        ElseIf Frm_IB_Radio.cboPlanDetail.Text = "Station" Then 'dw - update with the format saving (with - station_code)
            If lstSelectedStation.ListCount <> 0 Then
                For intSch = 0 To lstSelectedStation.ListCount - 1
                    Frm_IB_Radio.recPlanDetailTemp.Fields("Schedule").Value = Trim(txtSchedule.Text) & " - " & Trim(Mid(lstSelectedStation.List(intSch), InStr(lstSelectedStation.List(intSch), "-") + 1, Len(lstSelectedStation.List(intSch)) - InStr(lstSelectedStation.List(intSch), "-") + 1))
                    
                    Frm_IB_Radio.recPlanDetailTemp.Update
                    Frm_IB_Radio.recPlanDetailTemp.MoveNext
                Next
            Else
                Exit Do
            End If
        End If
    Loop
    
    Frm_IB_Radio.recPlanDetailTemp.Filter = ""
    
    'material mix
    Frm_IB_Radio.recPlanDetailMaterialTemp.Filter = ""
    Frm_IB_Radio.recPlanDetailMaterialTemp.Filter = strQuery
    
    Do While Not Frm_IB_Radio.recPlanDetailMaterialTemp.EOF And Not Frm_IB_Radio.recPlanDetailMaterialTemp.BOF
        If Frm_IB_Radio.cboPlanDetail.Text = "City" Then
            Frm_IB_Radio.recPlanDetailMaterialTemp.Fields("Schedule").Value = Trim(txtSchedule.Text)
            
            Frm_IB_Radio.recPlanDetailMaterialTemp.Update
            Frm_IB_Radio.recPlanDetailMaterialTemp.MoveNext

        ElseIf Frm_IB_Radio.cboPlanDetail.Text = "Station" Then 'dw - update with the format saving (with - station_code)
            If lstSelectedStation.ListCount <> 0 Then
                For intSch = 0 To lstSelectedStation.ListCount - 1
                    Frm_IB_Radio.recPlanDetailMaterialTemp.Fields("Schedule").Value = Trim(txtSchedule.Text) & " - " & Trim(Mid(lstSelectedStation.List(intSch), InStr(lstSelectedStation.List(intSch), "-") + 1, Len(lstSelectedStation.List(intSch)) - InStr(lstSelectedStation.List(intSch), "-") + 1))
                    
                    Frm_IB_Radio.recPlanDetailMaterialTemp.Update
                    Frm_IB_Radio.recPlanDetailMaterialTemp.MoveNext
                Next
            Else
                Exit Do
            End If
        End If
    Loop
    
    Frm_IB_Radio.recPlanDetailMaterialTemp.Filter = ""
    
    Me.MousePointer = vbDefault
End Sub

Private Sub TxtSpot_KeyPress(KeyAscii As Integer)
    If (Chr(KeyAscii) < "0" Or Chr(KeyAscii) > "9") And KeyAscii <> 8 And KeyAscii <> 13 And KeyAscii <> 9 And Chr(KeyAscii) <> "." Then
        KeyAscii = 0
        Beep
        Exit Sub
    End If
    fraSalesStation.Enabled = True
End Sub

Private Sub TxtSpot_LostFocus()
    If txtSpot.Text <> "" Then
        Frm_IB_Radio.recPlanDetailTemp.Filter = ""
        
        strQuery = "IB_Id='" & Frm_IB_Radio.txtIBID.Text & "' and "
        strQuery = strQuery & "Month=" & Get_Month_Number(cboMonthlyBudget.Text) & " and "
        strQuery = strQuery & "Year=" & Val(Frm_IB_Radio.cboYear.Text) & " "
        
        If Frm_IB_Radio.cboPlanDetail.Text = "City" Then 'dw - is splited cause they have different format
            strQuery = strQuery & "and Schedule='" & Clear_String(txtSchedule.Text) & "' "
        ElseIf Frm_IB_Radio.cboPlanDetail.Text = "Station" Then
            strQuery = strQuery & " AND Schedule like '%" & Clear_String(txtSchedule.Text) & "%'"
        End If
        
        Frm_IB_Radio.recPlanDetailTemp.Filter = strQuery
        
        Me.MousePointer = vbHourglass
        
        While Not Frm_IB_Radio.recPlanDetailTemp.EOF And Not Frm_IB_Radio.recPlanDetailTemp.BOF
            Frm_IB_Radio.recPlanDetailTemp.Fields("Spot").Value = txtSpot.Text
            Frm_IB_Radio.recPlanDetailTemp.Update
            Frm_IB_Radio.recPlanDetailTemp.MoveNext
        Wend
        
        Me.MousePointer = vbDefault
    End If
End Sub

Private Sub CreateStationTemp() 'dw - create recStationTemp
    Set recStationTemp = Nothing
    Set recStationTemp = New ADODB.Recordset

    With recStationTemp.Fields
        .Append "Client_Brief_Id", adChar, 10, adFldMayBeNull
        .Append "IB_Id", adChar, 13, adFldMayBeNull
        .Append "Month", adInteger, , adFldMayBeNull
        .Append "Year", adInteger, , adFldMayBeNull
        .Append "Schedule", adVarChar, 75, adFldMayBeNull
        .Append "City_Code", adInteger, , adFldMayBeNull
        .Append "Area_Code", adChar, 50, adFldMayBeNull
        .Append "Spot", adInteger, , adFldMayBeNull
        .Append "Urban_Flag", adSmallInt, , adFldMayBeNull
        .Append "Rural_Flag", adSmallInt, , adFldMayBeNull
    End With
    
    recStationTemp.Open
End Sub

Private Sub LoadStations() 'dw - Load station to treeview
    Dim recArea As New ADODB.Recordset
    Dim recCity As New ADODB.Recordset
    Dim recStation As New ADODB.Recordset
    Dim recSelectedStation As New ADODB.Recordset
    Dim recID As New ADODB.Recordset
    Dim strKeyArea As String
    Dim strKeyCity As String
    Dim Nod As Node
    Dim Nod1 As Node
    Dim nodx As Node
    
    Dim dblRDx As Double
    Dim blnStationCodeFound As Boolean

    Set Nod = Nothing
    Set Nod1 = Nothing
    
    lstSelectedStation.Clear
    
    treStations.Nodes.Clear
    treSelectedStation.Nodes.Clear
    
    recArea.CursorLocation = adUseClient
    recArea.Open "SELECT * FROM Area ORDER BY area_id", ConnERP, adOpenStatic, adLockOptimistic
    
    recCity.CursorLocation = adUseClient
    recCity.Open "SELECT * FROM City ORDER BY city_id", ConnERP, adOpenStatic, adLockOptimistic
    
    recStation.CursorLocation = adUseClient
    recStation.Open "SELECT * FROM Radio_Station ORDER BY station_Code", ConnERP, adOpenStatic, adLockOptimistic

    With recArea
        Do While .EOF = False
            strKeyArea = "A" & .Fields("Area_Id").Value
            
            Set Nod = treStations.Nodes.Add(, , strKeyArea, .Fields!area_name)
            Nod.ForeColor = &HC0&
            Nod.Bold = True
            
            Set Nod1 = treSelectedStation.Nodes.Add(, , strKeyArea, .Fields!area_name)
            Nod1.ForeColor = &HC0&
            Nod1.Bold = True
            
            .MoveNext
        Loop
    End With
        
    With recCity
        Do While .EOF = False
            strKeyCity = "K" & .Fields("City_Id").Value
            
            Set Nod = treStations.Nodes.Add("A" & .Fields("Area_ID").Value, tvwChild, strKeyCity, .Fields!city)
            Nod.ForeColor = vbBlue
            
            Set Nod1 = treSelectedStation.Nodes.Add("A" & .Fields("Area_ID").Value, tvwChild, strKeyCity, .Fields!city)
            Nod1.ForeColor = vbBlue
            
            .MoveNext
        Loop
    End With
    
    strQuery = "SELECT IB_Radio_Plan_Detail.Area_Code,Radio_Station.Station_Name FROM IB_Radio_Plan_Detail INNER JOIN Radio_Station on ib_radio_plan_Detail.Area_Code = Radio_Station.Station_code"
    strQuery = strQuery & " WHERE  ib_radio_plan_detail.IB_ID ='" & Frm_IB_Radio.txtIBID.Text & "'  "
    strQuery = strQuery & " AND ib_radio_plan_detail.month=" & Get_Month_Number(cboMonthlyBudget.Text) & " AND  ib_radio_plan_detail.year= " & Val(Frm_IB_Radio.cboYear.Text)
    strQuery = strQuery & " AND ib_radio_plan_detail.schedule like '%" & Clear_String(cboSchedule.Text) & "%'"
    strQuery = strQuery & " AND Radio_Station.status=1 ORDER BY IB_Radio_Plan_Detail.Area_Code"
        
    recSelectedStation.CursorLocation = adUseClient
    recSelectedStation.Open strQuery, ConnERP, adOpenStatic, adLockReadOnly
    
    If recSelectedStation.EOF And recSelectedStation.BOF Then
        Frm_IB_Radio.recPlanDetailTemp.Filter = ""

        strQuery = "IB_Id='" & Frm_IB_Radio.txtIBID.Text & "' and "
        strQuery = strQuery & "Month=" & Get_Month_Number(cboMonthlyBudget.Text) & " and "
        strQuery = strQuery & "Year=" & Val(Frm_IB_Radio.cboYear.Text)
        
        If Frm_IB_Radio.cboPlanDetail.Text = "City" Then 'dw - is splited cause they have different format
            strQuery = strQuery & " AND Schedule='" & Clear_String(cboSchedule.Text) & "' "
        ElseIf Frm_IB_Radio.cboPlanDetail.Text = "Station" Then
            If cboSchedule.Text <> "" Then
                strQuery = strQuery & " AND Schedule LIKE '%" & Clear_String(cboSchedule.Text) & "%' "
            End If
        End If
        
        Frm_IB_Radio.recPlanDetailTemp.Filter = strQuery
    End If
        
    With recStation
        Do While .EOF = False
            strKeyCity = "K" & .Fields("City_Id").Value
            
            Set Nod = treStations.Nodes.Add(strKeyCity, tvwChild, .Fields("Station_code").Value, .Fields("Station_Name").Value)
                                
            If strTransacProcess <> "ADD" Then
                If Not recSelectedStation.EOF And Not recSelectedStation.BOF Then
                    Do While Not recSelectedStation.EOF And Not recSelectedStation.BOF
                    
                        If .Fields("Station_code").Value = Trim(recSelectedStation.Fields("Area_Code").Value) Then
                             
                            Set Nod1 = treSelectedStation.Nodes.Add(strKeyCity, tvwChild, .Fields("Station_code").Value, .Fields("Station_Name").Value)
                
                            Nod1.ForeColor = vbRed
                            treStations.Nodes.Remove (Nod.Index)
                            
                            lstSelectedStation.AddItem .Fields("City_Id").Value & " - " & .Fields("Station_code").Value
                            
                            recSelectedStation.MoveNext
                        Else
                            Exit Do
                        End If
                    Loop
                Else
                    Do While Not Frm_IB_Radio.recPlanDetailTemp.EOF And Not Frm_IB_Radio.recPlanDetailTemp.BOF
                        If .Fields("Station_code").Value = Trim(Frm_IB_Radio.recPlanDetailTemp.Fields("Area_Code").Value) Then
                         
                            Set Nod1 = treSelectedStation.Nodes.Add(strKeyCity, tvwChild, .Fields("Station_code").Value, .Fields("Station_Name").Value)
                
                            Nod1.ForeColor = vbRed
                            treStations.Nodes.Remove (Nod.Index)
                            
                            lstSelectedStation.AddItem .Fields("City_Id").Value & " - " & .Fields("Station_code").Value
                            
                            Frm_IB_Radio.recPlanDetailTemp.MoveNext
                        Else
                            Exit Do
                        End If
                    Loop
                End If
            End If
            
            .MoveNext
        Loop
    End With
    
    Set recArea = Nothing
    Set recCity = Nothing
    Set recStation = Nothing
    Set recSelectedStation = Nothing
End Sub

Private Sub mnuStatAdd_Click() 'dw - Add station to SelectedStation treeview
    Dim recID As New ADODB.Recordset
    Dim strParent As String
    Dim Nod As Node

    On Error Resume Next

    Set nodx = treStations.SelectedItem
    strParent = ""

    If Frm_IB_Radio.strTransProcess <> "SHOW" Then
    
        strQuery = "SELECT City_Id,Area_Id FROM Radio_Station WHERE Station_Code='" & nodx.KEY & "'"
        recID.Open strQuery, ConnERP, adOpenStatic, adLockReadOnly
        
        strParent = nodx.Parent

        If strParent <> "" Then
            With Frm_IB_Radio.recPlanDetailTemp
                .AddNew
                .Fields("Client_Brief_Id").Value = ""
                .Fields("IB_Id").Value = Frm_IB_Radio.txtIBID.Text
                .Fields("Month").Value = Get_Month_Number(cboMonthlyBudget.Text)
                .Fields("Year").Value = Val(Frm_IB_Radio.cboYear.Text)
                .Fields("Schedule").Value = Clear_String(txtSchedule.Text) & " - " & Trim(nodx.KEY)
                .Fields("City_Code").Value = Trim(recID.Fields("City_Id").Value)
                .Fields("Area_Code").Value = Trim(nodx.KEY) 'Trim(recID.Fields("Area_Id").Value)
                .Fields("Spot").Value = IIf(txtSpot.Text = "", 0, txtSpot.Text)
                .Fields("Urban_Flag").Value = IIf(chkUrban.Value = True, 1, 0)
                .Fields("Rural_Flag").Value = IIf(chkRural.Value = True, 1, 0)
                
                .Update
            End With

            Frm_IB_Radio.recPlanDetailTemp.Filter = ""

            Set Nod = treSelectedStation.Nodes.Add(nodx.Parent.KEY, tvwChild, nodx.KEY, nodx.Text)

            lstSelectedStation.AddItem recID.Fields("City_Id").Value & " - " & Trim(nodx.KEY)
            
            Nod.ForeColor = vbRed
            treStations.Nodes.Remove nodx.Index

        ElseIf strParent = "" Then
            'do nothing
        End If
    End If
    
    recID.Close
    Set recID = Nothing
End Sub

Private Sub mnuStatDelete_Click() 'dw - Remove Station from SelectedStation treeview
    Dim Nod As Node
    Dim strAsk As String
    Dim strParent As String
    Dim strParentDel As String
    Dim strStationCodeDel As String
    Dim strNodeIndex As String
    Dim recID As New ADODB.Recordset
    Dim intLStat As Integer

    On Error GoTo my_error

    Set nodx = treSelectedStation.SelectedItem

    strParent = ""

    On Error Resume Next

    strParent = nodx.Parent
    strParentDel = nodx.Parent
    strStationCodeDel = nodx.KEY
    strNodeIndex = nodx.Index

    'dw - filter City_Id
    strQuery = "SELECT City_Id,Area_Id FROM Radio_Station WHERE Station_Code='" & nodx.KEY & "'"
    recID.Open strQuery, ConnERP, adOpenStatic, adLockReadOnly
    
    'dw - Save to recPlanDetailTemp for IB by Station
    If Frm_IB_Radio.strTransProcess <> "SHOW" Then
        If strParentDel <> "" Then
            With Frm_IB_Radio.recPlanDetailTemp
                .Filter = ""
                .Filter = "IB_ID='" & Frm_IB_Radio.txtIBID.Text & "' AND station_code='" & strStationCodeDel & "'"

                If .RecordCount > 0 Then
                    .MoveFirst

                    Do While .EOF = False
                        If Trim(.Fields(1).Value) = Trim(Frm_IB_Radio.txtIBID.Text) Then
                            If Trim(.Fields(2).Value) = Trim(Get_Month_Number(cboMonthlyBudget.Text)) Then
                                If Trim(.Fields(3).Value) = Val(Frm_IB_Radio.cboYear.Text) Then
                                    If Trim(.Fields(4).Value) = Trim(Clear_String(txtSchedule.Text) & " - " & Trim(strStationCodeDel)) Then
                                        If Trim(.Fields(5).Value) = Trim(recID.Fields("City_Id").Value) Then
                                            .Delete
                                        End If
                                    End If
                                End If
                            End If
                        End If
                            
                        .MoveNext
                    Loop
                End If
            End With

            With Frm_IB_Radio.recPlanDetailMaterialTemp
                .Filter = ""
                
                strQuery = "IB_Id='" & Frm_IB_Radio.txtIBID.Text & "' AND "
                strQuery = strQuery & "Month=" & Get_Month_Number(cboMonthlyBudget.Text) & " AND "
                strQuery = strQuery & "Year=" & Val(Frm_IB_Radio.cboYear.Text) & " AND "
                strQuery = strQuery & "Schedule='" & Clear_String(txtSchedule.Text) & " - " & Trim(strStationCodeDel) & "'"
                
                .Filter = strQuery
                
                If .RecordCount > 0 Then
                    .MoveFirst
                    
                    Do While .EOF = False
                        If Trim(.Fields("IB_Id").Value) = Trim(Frm_IB_Radio.txtIBID.Text) Then
                            If Trim(.Fields("month").Value) = Trim(Get_Month_Number(cboMonthlyBudget.Text)) Then
                                If Trim(.Fields("year").Value) = Val(Frm_IB_Radio.cboYear.Text) Then
                                    If Trim(.Fields("Schedule").Value) = Trim(Clear_String(txtSchedule.Text) & " - " & Trim(strStationCodeDel)) Then
                                        If Trim(.Fields("City_Code").Value) = Trim(recID.Fields("City_Id").Value) Then
                                            If Trim(.Fields("material_Id").Value) = Trim(msgMix.TextMatrix(msgMix.Rows - 1, 0)) Then
                                                .Delete
                                            End If
                                        End If
                                    End If
                                End If
                            End If
                        End If
                        
                        .MoveNext
                    Loop
                End If
            End With

            If Err.Number = 35605 Then
                Err.Clear
            Else
                Set Nod = treStations.Nodes.Add(, tvwChild, strStationCodeDel, nodx.Text)
                                
                treSelectedStation.Nodes.Remove nodx.Index
                                                
                'dw - Remove List Station Selected Deleted in tempObj(lstSelectedStation)
                For intLStat = 0 To lstSelectedStation.ListCount - 1
                    If lstSelectedStation.List(intLStat) = recID.Fields("City_Id").Value & " - " & Trim(strStationCodeDel) Then
                        lstSelectedStation.ListIndex = intLStat
                        lstSelectedStation.RemoveItem (lstSelectedStation.ListIndex)
                    End If
                Next
                                
            End If
        End If
    End If
    
    recID.Close
    Set recID = Nothing
    
    Exit Sub

my_error:
    MsgBox Err.Number & " " & Err.Description, vbCritical, strTitleExclamation
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
    With picButton(enButtonType.bieClose)       'Quit.
        .Enabled = paIsNormalMode
        .Visible = paIsNormalMode
    End With

    With picButton(enButtonType.bieSave)  'SAVE.
        .Enabled = Not paIsNormalMode
        .Visible = Not paIsNormalMode
        .Left = picButton(4).Left
    End With

    With picButton(enButtonType.biecancel) 'CANCEL.
        .Enabled = Not paIsNormalMode
        .Visible = Not paIsNormalMode
        .Left = picButton(5).Left
    End With
    'pnl_Main.Enabled = Not paIsNormalMode
    'cboBrand.Enabled = paIsNormalMode
'    blnEditOrAdd = Not paIsNormalMode
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

Private Sub SetPictureTBEnabled(ByVal Index As Integer, ByVal paIsNormalMode As Boolean)
 '*****************************************
'Procedure Name     : SetPictureTBEnabled
'Procedure Function :   Enable/Disable Button
'Input Parameter    : Index,paIsNormalMode,picOBJ
'Output Parameter   :
'Date               : -
'LastUpdate/By      : - Tedi
'*****************************************
    
    If paIsNormalMode = True Then
        picButton(Index).Picture = LoadPicture(SetButtonImageEffect(Index, bieNormal))
    Else: picButton(Index).Picture = LoadPicture(SetButtonImageEffect(Index, bieDisabled))
    End If
    
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
    pnlMain.Height = Me.Height - pnlMain.Top - picStatusBar.Height
    fraIB.Width = pnlMain.Width - (fraIB.Left) - 150
    fraPlanMonth.Width = Me.Width - (fraPlanMonth.Left * 2)
    cmdMateri.Left = fraPlanMonth.Width - cmdMateri.Width - 300
    cmdEditPlan.Left = cmdMateri.Left
    msgCity.Width = fraPlanMonth.Width - msgCity.Left - cmdMateri.Width - 450
    fraClientApproval.Width = pnlMain.Width - fraClientApproval.Left - 150
    picApproved.Left = ((fraClientApproval.Width / 2)) - ((picApproved.Width) / 2) 'fraClientApproval.Width - (picApproved.Left * 2)
    txtMediaPlanNo.Width = fraIB.Width - txtMediaPlanNo.Left - 150
    txtPrimaryTarget.Width = txtMediaPlanNo.Width

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
        Case enButtonType.bieAdd  '4 'ADD.

            Call cmdAdd_Click
            
        Case enButtonType.bieEdit  '5 'EDIT.
            Call cmdEdit_Click
        Case enButtonType.bieDelete  '6 'DELETE.
            Call cmdDelete_Click
        Case enButtonType.bieSave  'SAVE.
            Call cmdSave_Click
        Case enButtonType.biecancel 'CANCEL.
            Call cmdCancel_Click
        Case enButtonType.bieClose  'CANCEL.
            Call cmdClose_Click
    End Select
    
End Sub
