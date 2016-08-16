VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "mscomctl.OCX"
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFlxGrd.ocx"
Begin VB.Form frm_MPInsertion 
   BorderStyle     =   0  'None
   Caption         =   "Media Plan"
   ClientHeight    =   9105
   ClientLeft      =   480
   ClientTop       =   1860
   ClientWidth     =   14730
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   9105
   ScaleWidth      =   14730
   ShowInTaskbar   =   0   'False
   Begin VB.PictureBox picButton 
      Appearance      =   0  'Flat
      BackColor       =   &H00F0F0F0&
      BorderStyle     =   0  'None
      ClipControls    =   0   'False
      ForeColor       =   &H80000008&
      Height          =   750
      Index           =   63
      Left            =   10800
      ScaleHeight     =   750
      ScaleWidth      =   1500
      TabIndex        =   83
      TabStop         =   0   'False
      ToolTipText     =   "TV Layering"
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
      Index           =   60
      Left            =   3150
      ScaleHeight     =   750
      ScaleWidth      =   1500
      TabIndex        =   80
      TabStop         =   0   'False
      ToolTipText     =   "Copy"
      Top             =   0
      Width           =   1500
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
      ScaleWidth      =   14730
      TabIndex        =   69
      Top             =   8775
      Width           =   14730
      Begin VB.PictureBox pic_Legend 
         BorderStyle     =   0  'None
         Height          =   420
         Left            =   5895
         ScaleHeight     =   420
         ScaleWidth      =   14175
         TabIndex        =   73
         Top             =   45
         Width           =   14175
         Begin VB.Label Label12 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Unlocked Week"
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
            Left            =   4590
            TabIndex        =   40
            Top             =   15
            Width           =   1290
         End
         Begin VB.Shape Shape_unlocked 
            BackColor       =   &H00FFFF80&
            FillColor       =   &H000080FF&
            FillStyle       =   0  'Solid
            Height          =   165
            Left            =   4440
            Top             =   30
            Width           =   180
         End
         Begin VB.Shape LegendApproval 
            BackColor       =   &H00FFFF80&
            FillColor       =   &H00FFFF80&
            FillStyle       =   0  'Solid
            Height          =   165
            Left            =   900
            Top             =   30
            Width           =   180
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
            Left            =   1095
            TabIndex        =   78
            Top             =   15
            Width           =   1230
         End
         Begin VB.Label Label14 
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
            Left            =   2625
            TabIndex        =   77
            Top             =   15
            Width           =   1680
         End
         Begin VB.Shape LegendWebApproval 
            BackColor       =   &H00FFFF80&
            FillColor       =   &H00FFC0FF&
            FillStyle       =   0  'Solid
            Height          =   165
            Left            =   2460
            Top             =   30
            Width           =   180
         End
         Begin VB.Label Label9 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Legend :"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   -1  'True
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   225
            Left            =   0
            TabIndex        =   76
            Top             =   0
            Width           =   810
         End
         Begin VB.Shape LegendTVRF 
            BackColor       =   &H00FFFF80&
            FillColor       =   &H0000FFFF&
            FillStyle       =   0  'Solid
            Height          =   165
            Left            =   5925
            Top             =   30
            Width           =   180
         End
         Begin VB.Label Label15 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "TV Reach && Freq"
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
            Left            =   6090
            TabIndex        =   75
            Top             =   15
            Width           =   1290
         End
         Begin VB.Label Label16 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Actual Budget"
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
            Left            =   7650
            TabIndex        =   74
            Top             =   15
            Width           =   1215
         End
         Begin VB.Shape LegendActual 
            BackColor       =   &H00FFFF80&
            FillColor       =   &H0000C000&
            FillStyle       =   0  'Solid
            Height          =   165
            Left            =   7410
            Top             =   30
            Width           =   180
         End
      End
      Begin VB.PictureBox picDescColor 
         BorderStyle     =   0  'None
         Height          =   345
         Left            =   9975
         ScaleHeight     =   345
         ScaleWidth      =   1695
         TabIndex        =   70
         Top             =   15
         Width           =   1695
      End
      Begin VB.Label lblLastModifiedDate 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Last Modified Date: XX-XXX-XXXX |"
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   90
         TabIndex        =   72
         Tag             =   "Last Modified Date: "
         Top             =   75
         Width           =   2520
      End
      Begin VB.Label lblLastModifiedBy 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Last Modified by: XXXXXXXXXXXXXXX |"
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   2670
         TabIndex        =   71
         Tag             =   "Last Modified by: "
         Top             =   75
         Width           =   2730
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
      ScaleWidth      =   14730
      TabIndex        =   17
      Top             =   0
      Width           =   14730
      Begin VB.PictureBox picButton 
         Appearance      =   0  'Flat
         BackColor       =   &H00F0F0F0&
         BorderStyle     =   0  'None
         ClipControls    =   0   'False
         ForeColor       =   &H80000008&
         Height          =   750
         Index           =   62
         Left            =   7740
         ScaleHeight     =   750
         ScaleWidth      =   1500
         TabIndex        =   82
         TabStop         =   0   'False
         ToolTipText     =   "TV Layering"
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
         Index           =   7
         Left            =   4680
         ScaleHeight     =   750
         ScaleWidth      =   1500
         TabIndex        =   79
         TabStop         =   0   'False
         ToolTipText     =   "Find"
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
         TabIndex        =   23
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
         Index           =   22
         Left            =   12330
         ScaleHeight     =   750
         ScaleWidth      =   1500
         TabIndex        =   22
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
         Index           =   23
         Left            =   13860
         ScaleHeight     =   750
         ScaleWidth      =   1500
         TabIndex        =   21
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
         Index           =   64
         Left            =   9270
         ScaleHeight     =   750
         ScaleWidth      =   1500
         TabIndex        =   20
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
         Index           =   61
         Left            =   6210
         ScaleHeight     =   750
         ScaleWidth      =   1500
         TabIndex        =   19
         ToolTipText     =   "TV Layering"
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
         TabIndex        =   18
         Top             =   0
         Width           =   1500
      End
   End
   Begin VB.Frame FrameViewMonth1 
      BackColor       =   &H80000015&
      BorderStyle     =   0  'None
      Caption         =   "Frame3"
      Height          =   2760
      Left            =   600
      TabIndex        =   0
      Top             =   2040
      Visible         =   0   'False
      Width           =   2415
      Begin VB.Frame FrameViewMonth2 
         BorderStyle     =   0  'None
         Caption         =   "Frame3"
         Height          =   2685
         Left            =   30
         TabIndex        =   1
         Top             =   30
         Width           =   2355
         Begin VB.CommandButton cmdCancel 
            Caption         =   "Cancel"
            Height          =   360
            Left            =   1215
            TabIndex        =   16
            Top             =   2190
            Width           =   960
         End
         Begin VB.CommandButton cmdOK 
            Caption         =   "OK"
            Height          =   360
            Left            =   195
            TabIndex        =   15
            Top             =   2190
            Width           =   960
         End
         Begin VB.CheckBox cbkAllMonth 
            Caption         =   "ALL"
            Height          =   255
            Left            =   195
            TabIndex        =   14
            Top             =   150
            Width           =   1065
         End
         Begin VB.CheckBox cbkViewMonth 
            Caption         =   "December"
            Height          =   285
            Index           =   12
            Left            =   1140
            TabIndex        =   13
            Top             =   1830
            Width           =   1080
         End
         Begin VB.CheckBox cbkViewMonth 
            Caption         =   "November"
            Height          =   285
            Index           =   11
            Left            =   1140
            TabIndex        =   12
            Top             =   1545
            Width           =   1080
         End
         Begin VB.CheckBox cbkViewMonth 
            Caption         =   "October"
            Height          =   285
            Index           =   10
            Left            =   1140
            TabIndex        =   11
            Top             =   1275
            Width           =   1080
         End
         Begin VB.CheckBox cbkViewMonth 
            Caption         =   "September"
            Height          =   285
            Index           =   9
            Left            =   1140
            TabIndex        =   10
            Top             =   1005
            Width           =   1080
         End
         Begin VB.CheckBox cbkViewMonth 
            Caption         =   "August"
            Height          =   285
            Index           =   8
            Left            =   1140
            TabIndex        =   9
            Top             =   735
            Width           =   1080
         End
         Begin VB.CheckBox cbkViewMonth 
            Caption         =   "July"
            Height          =   285
            Index           =   7
            Left            =   1140
            TabIndex        =   8
            Top             =   480
            Width           =   1080
         End
         Begin VB.CheckBox cbkViewMonth 
            Caption         =   "June"
            Height          =   285
            Index           =   6
            Left            =   195
            TabIndex        =   7
            Top             =   1845
            Width           =   1080
         End
         Begin VB.CheckBox cbkViewMonth 
            Caption         =   "May"
            Height          =   285
            Index           =   5
            Left            =   195
            TabIndex        =   6
            Top             =   1560
            Width           =   1080
         End
         Begin VB.CheckBox cbkViewMonth 
            Caption         =   "April"
            Height          =   285
            Index           =   4
            Left            =   195
            TabIndex        =   5
            Top             =   1275
            Width           =   1080
         End
         Begin VB.CheckBox cbkViewMonth 
            Caption         =   "March"
            Height          =   285
            Index           =   3
            Left            =   195
            TabIndex        =   4
            Top             =   1005
            Width           =   1080
         End
         Begin VB.CheckBox cbkViewMonth 
            Caption         =   "February"
            Height          =   285
            Index           =   2
            Left            =   195
            TabIndex        =   3
            Top             =   735
            Width           =   1080
         End
         Begin VB.CheckBox cbkViewMonth 
            Caption         =   "January"
            Height          =   285
            Index           =   1
            Left            =   195
            TabIndex        =   2
            Top             =   480
            Width           =   1080
         End
      End
   End
   Begin Threed.SSPanel pnl_Main 
      Align           =   1  'Align Top
      Height          =   8190
      Left            =   0
      TabIndex        =   24
      Top             =   750
      Width           =   14730
      _Version        =   65536
      _ExtentX        =   25982
      _ExtentY        =   14446
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
      Begin VB.Frame fra_Insertion 
         BackColor       =   &H00F0F0F0&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   8385
         Left            =   75
         TabIndex        =   25
         Top             =   30
         Width           =   14490
         Begin VB.Frame FrameProgressBarCopy 
            Appearance      =   0  'Flat
            BackColor       =   &H80000002&
            BorderStyle     =   0  'None
            Caption         =   "Frame4"
            ForeColor       =   &H80000008&
            Height          =   780
            Left            =   3585
            TabIndex        =   28
            Top             =   3540
            Visible         =   0   'False
            Width           =   7725
            Begin VB.Frame Frame5 
               Appearance      =   0  'Flat
               BackColor       =   &H8000000A&
               BorderStyle     =   0  'None
               Caption         =   "Frame4"
               ForeColor       =   &H80000008&
               Height          =   750
               Left            =   15
               TabIndex        =   29
               Top             =   15
               Width           =   7725
               Begin MSComctlLib.ProgressBar ProgressBarCopy 
                  Height          =   285
                  Left            =   75
                  TabIndex        =   30
                  Top             =   225
                  Width           =   7470
                  _ExtentX        =   13176
                  _ExtentY        =   503
                  _Version        =   393216
                  BorderStyle     =   1
                  Appearance      =   0
                  Max             =   5
                  Scrolling       =   1
               End
            End
         End
         Begin VB.PictureBox Picture7 
            Height          =   360
            Left            =   720
            ScaleHeight     =   300
            ScaleWidth      =   990
            TabIndex        =   26
            Top             =   7215
            Visible         =   0   'False
            Width           =   1050
            Begin VB.CommandButton cmdApproval 
               Caption         =   "A&pproval"
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
               Height          =   300
               Left            =   0
               TabIndex        =   27
               ToolTipText     =   "create new media plan"
               Top             =   0
               Width           =   990
            End
         End
         Begin VB.PictureBox pic_Grd 
            Appearance      =   0  'Flat
            ForeColor       =   &H80000008&
            Height          =   855
            Left            =   150
            ScaleHeight     =   825
            ScaleWidth      =   14190
            TabIndex        =   46
            Top             =   210
            Width           =   14220
            Begin VB.ComboBox cboMPNumber 
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
               Left            =   1200
               Style           =   2  'Dropdown List
               TabIndex        =   58
               Top             =   105
               Width           =   1575
            End
            Begin VB.TextBox txtBrandName 
               Appearance      =   0  'Flat
               BackColor       =   &H00E0E0E0&
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
               Left            =   3540
               Locked          =   -1  'True
               TabIndex        =   57
               Text            =   "txtBrandName"
               Top             =   105
               Width           =   2415
            End
            Begin VB.TextBox txtClientName 
               Appearance      =   0  'Flat
               BackColor       =   &H00E0E0E0&
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
               Left            =   3540
               Locked          =   -1  'True
               TabIndex        =   56
               Text            =   "txtClientName"
               Top             =   465
               Width           =   2415
            End
            Begin VB.TextBox txtCreatedDate 
               Appearance      =   0  'Flat
               BackColor       =   &H00E0E0E0&
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
               Left            =   6840
               Locked          =   -1  'True
               TabIndex        =   55
               Text            =   "txtCreatedDate"
               Top             =   105
               Width           =   1485
            End
            Begin VB.TextBox txtCreatedBy 
               Appearance      =   0  'Flat
               BackColor       =   &H00E0E0E0&
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
               Left            =   6840
               Locked          =   -1  'True
               TabIndex        =   54
               Text            =   "txtCreatedBy"
               Top             =   465
               Width           =   1485
            End
            Begin VB.TextBox txtLastUpdateDate 
               Appearance      =   0  'Flat
               BackColor       =   &H00E0E0E0&
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
               Left            =   9465
               Locked          =   -1  'True
               TabIndex        =   53
               Text            =   "txtLastUpdateDate"
               Top             =   105
               Width           =   1485
            End
            Begin VB.TextBox txtLastUpdateBy 
               Appearance      =   0  'Flat
               BackColor       =   &H00E0E0E0&
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
               Left            =   9465
               Locked          =   -1  'True
               TabIndex        =   52
               Text            =   "txtLastUpdateBy"
               Top             =   465
               Width           =   1485
            End
            Begin VB.TextBox txtViewMonth 
               Appearance      =   0  'Flat
               BackColor       =   &H00E0E0E0&
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
               Left            =   1200
               Locked          =   -1  'True
               TabIndex        =   51
               Text            =   "ALL"
               Top             =   465
               Width           =   1575
            End
            Begin VB.PictureBox picViewMonth 
               Appearance      =   0  'Flat
               BackColor       =   &H80000005&
               BorderStyle     =   0  'None
               ForeColor       =   &H80000008&
               Height          =   240
               Left            =   2520
               Picture         =   "frm_MPInsertion.frx":0000
               ScaleHeight     =   240
               ScaleWidth      =   225
               TabIndex        =   50
               Top             =   480
               Width           =   225
            End
            Begin VB.TextBox txt_remaining_budget 
               Alignment       =   1  'Right Justify
               Appearance      =   0  'Flat
               BackColor       =   &H00E0E0E0&
               BorderStyle     =   0  'None
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   210
               Left            =   11940
               Locked          =   -1  'True
               TabIndex        =   49
               Top             =   495
               Width           =   2145
            End
            Begin VB.PictureBox Picture6 
               Appearance      =   0  'Flat
               BackColor       =   &H80000005&
               ForeColor       =   &H80000008&
               Height          =   330
               Left            =   13665
               ScaleHeight     =   300
               ScaleWidth      =   390
               TabIndex        =   47
               Top             =   105
               Width           =   420
               Begin VB.CommandButton cmdEditGivenBudget 
                  BeginProperty Font 
                     Name            =   "Tahoma"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   405
                  Left            =   -15
                  Picture         =   "frm_MPInsertion.frx":0342
                  Style           =   1  'Graphical
                  TabIndex        =   48
                  ToolTipText     =   "Edit Plan budget"
                  Top             =   -30
                  Width           =   435
               End
            End
            Begin VB.Label Label2 
               Alignment       =   1  'Right Justify
               Caption         =   "Brand "
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
               Left            =   2835
               TabIndex        =   67
               Top             =   135
               Width           =   645
            End
            Begin VB.Label Label1 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               Caption         =   "MP Number "
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   195
               Left            =   285
               TabIndex        =   66
               Top             =   120
               Width           =   855
            End
            Begin VB.Label Label3 
               Alignment       =   1  'Right Justify
               Caption         =   "Client "
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
               Left            =   2820
               TabIndex        =   65
               Top             =   465
               Width           =   645
            End
            Begin VB.Label Label4 
               Alignment       =   1  'Right Justify
               Caption         =   "Created "
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
               Left            =   5475
               TabIndex        =   64
               Top             =   135
               Width           =   1305
            End
            Begin VB.Label Label5 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               Caption         =   "By "
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   195
               Left            =   6195
               TabIndex        =   63
               Top             =   495
               Width           =   225
            End
            Begin VB.Label Label7 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               Caption         =   "By"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   195
               Left            =   8520
               TabIndex        =   62
               Top             =   480
               Width           =   180
            End
            Begin VB.Label Label6 
               Alignment       =   1  'Right Justify
               Caption         =   "Last Update"
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
               Left            =   8430
               TabIndex        =   61
               Top             =   150
               Width           =   945
            End
            Begin VB.Label Label8 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               Caption         =   "View "
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   195
               Left            =   285
               TabIndex        =   60
               Top             =   465
               Width           =   375
            End
            Begin VB.Label Label13 
               Caption         =   "Remaining Budget : "
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   270
               Left            =   11955
               TabIndex        =   59
               Top             =   165
               Width           =   1755
            End
         End
         Begin VB.PictureBox pic_MhsFlex 
            Appearance      =   0  'Flat
            BackColor       =   &H00F0F0F0&
            ForeColor       =   &H80000008&
            Height          =   5985
            Left            =   150
            ScaleHeight     =   5955
            ScaleWidth      =   14190
            TabIndex        =   34
            Top             =   1065
            Width           =   14220
            Begin VB.TextBox txtMPInsertion 
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
               Height          =   375
               Left            =   4575
               TabIndex        =   44
               Text            =   "Text1"
               Top             =   480
               Visible         =   0   'False
               Width           =   975
            End
            Begin VB.ListBox LstHidenCol 
               Height          =   1860
               Left            =   7065
               Style           =   1  'Checkbox
               TabIndex        =   43
               Top             =   210
               Visible         =   0   'False
               Width           =   1815
            End
            Begin VB.Frame FrameProgressBar 
               Appearance      =   0  'Flat
               BackColor       =   &H80000002&
               BorderStyle     =   0  'None
               Caption         =   "FramePleaseWait"
               ForeColor       =   &H00FFFF00&
               Height          =   1290
               Left            =   5805
               TabIndex        =   38
               Top             =   2145
               Visible         =   0   'False
               Width           =   3510
               Begin VB.Frame FrameProgressBar1 
                  BorderStyle     =   0  'None
                  Caption         =   "FramePleaseWait"
                  Height          =   1245
                  Left            =   15
                  TabIndex        =   39
                  Top             =   30
                  Width           =   3480
                  Begin MSComctlLib.ProgressBar ProgressBarExport 
                     Height          =   165
                     Left            =   75
                     TabIndex        =   81
                     Top             =   570
                     Width           =   3330
                     _ExtentX        =   5874
                     _ExtentY        =   291
                     _Version        =   393216
                     Appearance      =   0
                  End
                  Begin VB.Label lblPercent 
                     Caption         =   "s"
                     BeginProperty Font 
                        Name            =   "Tahoma"
                        Size            =   8.25
                        Charset         =   0
                        Weight          =   400
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     ForeColor       =   &H00404040&
                     Height          =   210
                     Left            =   90
                     TabIndex        =   42
                     Top             =   735
                     Width           =   4260
                  End
                  Begin VB.Label lblPleaseWait 
                     BackColor       =   &H80000002&
                     Caption         =   " Exporting to Excel..."
                     BeginProperty Font 
                        Name            =   "Tahoma"
                        Size            =   8.25
                        Charset         =   0
                        Weight          =   400
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     ForeColor       =   &H80000009&
                     Height          =   210
                     Left            =   0
                     TabIndex        =   41
                     Top             =   -30
                     Width           =   3480
                  End
               End
            End
            Begin VB.PictureBox PicButtonPanahBawah 
               Appearance      =   0  'Flat
               BackColor       =   &H80000005&
               BorderStyle     =   0  'None
               ForeColor       =   &H80000008&
               Height          =   210
               Left            =   3195
               Picture         =   "frm_MPInsertion.frx":048C
               ScaleHeight     =   210
               ScaleWidth      =   225
               TabIndex        =   37
               Top             =   1020
               Visible         =   0   'False
               Width           =   225
            End
            Begin VB.ListBox LstFrequency 
               Appearance      =   0  'Flat
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   1980
               Left            =   3255
               TabIndex        =   36
               Top             =   1785
               Visible         =   0   'False
               Width           =   720
            End
            Begin VB.TextBox txtTVREACH 
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
               Height          =   375
               Left            =   5760
               TabIndex        =   35
               Text            =   "Text1"
               Top             =   480
               Visible         =   0   'False
               Width           =   975
            End
            Begin MSFlexGridLib.MSFlexGrid msf_MPInsertion 
               Height          =   6015
               Left            =   15
               TabIndex        =   45
               Top             =   15
               Width           =   10710
               _ExtentX        =   18891
               _ExtentY        =   10610
               _Version        =   393216
               Rows            =   5
               Cols            =   7
               FixedRows       =   4
               FixedCols       =   2
               BackColorFixed  =   12615680
               ForeColorFixed  =   16777215
               GridColor       =   12356167
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
         Begin VB.Frame Frame2 
            Height          =   510
            Left            =   5055
            TabIndex        =   31
            Top             =   4170
            Visible         =   0   'False
            Width           =   1575
            Begin VB.Label lblReleaseDate 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H80000005&
               BackStyle       =   0  'Transparent
               Caption         =   "31 September 2004"
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
               Left            =   0
               TabIndex        =   33
               Top             =   210
               Width           =   1605
            End
            Begin VB.Label Label10 
               Alignment       =   2  'Center
               Caption         =   " Release Date"
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
               Left            =   255
               TabIndex        =   32
               Top             =   15
               Width           =   1020
            End
         End
         Begin VB.Label lbl_show_user_server 
            Height          =   300
            Left            =   12585
            TabIndex        =   68
            Top             =   8475
            Width           =   825
         End
      End
   End
End
Attribute VB_Name = "frm_MPInsertion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
''*****************************************************************************
' Nama Submodul         :  Frm_MPInsertion
' Fungsi Submodul       :  Untuk Edit Insertion
' Nama Programmer       :  Sistyo
' Tgl Pembuatan         :  16 Agustus 2004
' Last Update           :  18 Agustus 2004/Sistyo
'*****************************************************************************
Option Explicit
Dim rsTemp As New ADODB.Recordset
Dim intMouseCol As Double 'posisi colom waktu klik kanan
Dim intRow As Double 'posisi row waktu edit
Dim intCol As Double 'posisi colom waktu edit
Dim FlagMarkStart As Boolean
Dim MarkStartCol As Double
Dim MarkStartRow As Double
Dim EditMode As Boolean
Dim strViewMonth As String
Dim dblUnlockedCellColor As Double

Dim fso As Object
Dim FSOInterface As Object
Dim LogFileName As String

Dim bol_BU2 As Boolean

Private Sub CboMonth_Click()
    Call cboMPNumber_Click
End Sub

Private Sub cbkAllMonth_Click()
    Dim i As Integer
    If cbkAllMonth.Value = 1 Then
        For i = 1 To 12
            cbkViewMonth(i).Enabled = False
        Next
    Else
        For i = 1 To 12
            cbkViewMonth(i).Enabled = True
        Next
    End If
End Sub

Private Sub db_Publish_to_Web()
'<CSCM>
'********************************************************************************
'Procedure Name     : db_Publish_to_Web
'Procedure Function : ---
'Input Parameter    : ---
'Output Parameter   : ---
'Date               : 3/17/2016
'LastUpdate/By      : Tedi / Kreatif
'Name Before        : cmd_Publish_Click
'********************************************************************************
'</CSCM>
    
    If cboMPNumber.Text <> "" Then
        If MsgBox("Are you sure publish this Media Plan on the web?, It will replace current planned budget in Budget Control.", vbYesNo) = 6 Then
            
            Me.MousePointer = vbHourglass
            
            'ConnERP.Execute "UPDATE mp_master SET is_upload_to_web = 1 WHERE mp_number = '" & cboMPNumber.Text & "'"
            
            ConnERP.Execute "UPDATE mp_master SET is_upload_to_web = 1, last_update_by='" & strLogin_User & "',Last_update_date=getdate() WHERE mp_number = '" & cboMPNumber.Text & "'"
            
            Call Update_Budget_Control(cboMPNumber.Text) 'Submit budget to Budget Control Plan
            
            MsgBox "Media Plan has been Published to web!", vbExclamation, strApplication_Name
            
            picButton(biePublishToWeb).Enabled = False
            picButton(bieResendEmail).Enabled = True
            Me.MousePointer = vbDefault
            
            'Send Email To Client
            frm_MPSendMail.show 1
            
        End If
    Else
        MsgBox "Please Select MP Number!"
    End If
End Sub

Private Sub Update_Budget_Control(strMPNumber As String)
    Dim strSql As String, IntYear As Integer, intMonth As Integer
    
    IntYear = CInt(Mid(strMPNumber, 6, 4))
    strSql = "select distinct original_brand_code from mp_activity where mp_activity_id in "
    strSql = strSql & "(select mp_activity_id from mp_ids where mp_number = '" & strMPNumber & "')"
    
    Dim rsbrand As New ADODB.Recordset, rsPlanBudget As New ADODB.Recordset, rsBrandFee As New ADODB.Recordset
    Dim MSC As Double, MSC_Flag As Integer, Bonus_Comm As Double, Bonus_Comm_Flag As Integer
    Dim Club_Agency As Double, Club_Agency_Flag As Integer
    
    rsbrand.Open strSql, ConnERP, 1, 3
    While Not rsbrand.EOF
        If Is_Special_Brand(rsbrand(0)) Then
            'BU1/ULI
            'Delete Budget_Control
            strSql = "Delete from Budget_Control where brand_code='" & rsbrand(0) & "' and [year]=" & IntYear
            ConnERP.Execute strSql
            'Delete Budget_Control_Plan_Detail
            strSql = "Delete from Budget_Control_Plan_Detail where brand_code='" & rsbrand(0) & "' and [year]=" & IntYear
            ConnERP.Execute strSql
            
            strSql = " select month_number,medium_code,sum(budget) budget from"
            strSql = strSql & " ("
            strSql = strSql & " SELECT case b.medium_code"
            strSql = strSql & "     when 'CN' then 'OT'"
            strSql = strSql & "     when 'OT' then 'OT'"
            strSql = strSql & "     when 'RD' then 'RD'"
            strSql = strSql & "     when 'PR' then 'PR'"
            strSql = strSql & "     when 'TV' then 'TV'"
            strSql = strSql & "     end medium_code,"
            strSql = strSql & " a.month_number,"
            strSql = strSql & " case isnull(a.total_actual,-1) when -1 then a.min_budget else a.Actual_Nett_Paid End budget FROM mp_monthly_activity a  "
            strSql = strSql & " INNER JOIN mp_medium b on a.mp_medium_id = b.mp_medium_id  "
            strSql = strSql & " INNER JOIN mp_activity c on b.mp_activity_id = c.mp_activity_id  "
            strSql = strSql & " INNER JOIN mp_task d on c.mp_task_id = d.mp_task_id "
            'parameter here
            strSql = strSql & " and c.original_brand_code = '" & rsbrand(0) & "' "
            strSql = strSql & " and d.mp_number = '" & strMPNumber & "' "
            strSql = strSql & " ) x"
            strSql = strSql & " GROUP BY month_number,medium_code "
            strSql = strSql & " order by month_number,medium_code"
            intMonth = Empty
            rsPlanBudget.Open strSql, ConnERP
            While Not rsPlanBudget.EOF
                If intMonth <> rsPlanBudget("month_number") Then
                    'get brand fee for current month
                    intMonth = rsPlanBudget("month_number")
                    
                    MSC = Empty
                    MSC_Flag = Empty
                    Bonus_Comm = Empty
                    Bonus_Comm_Flag = Empty
                    Club_Agency = Empty
                    Club_Agency_Flag = Empty
                    
                    strSql = " select MSC_Paid,MSC_Paid_On_Flag,MSC_Bonus,MSC_Bonus_On_Flag,Club_Agency,Club_Agency_On_Flag from brand_fee "
                    strSql = strSql & " where brand_code = '" & rsbrand(0) & "' and [month] = " & intMonth & " and [year] = " & IntYear
                    rsBrandFee.Open strSql, ConnERP
                    If Not rsBrandFee.EOF Then
                        MSC = rsBrandFee(0)
                        MSC_Flag = rsBrandFee(1)
                        Bonus_Comm = rsBrandFee(2)
                        Bonus_Comm_Flag = rsBrandFee(3)
                        Club_Agency = rsBrandFee(4)
                        Club_Agency_Flag = rsBrandFee(5)
                    End If
                    rsBrandFee.Close
                End If
                'insert Budget_Control
                strSql = "insert into budget_control(Brand_Code,[Year],[Month],Media_Code,Msc,Msc_Flag,Club_Agency,Club_Agency_Flag,Bonus_Comm,Bonus_Comm_Flag,Status) "
                strSql = strSql & " values('" & rsbrand(0) & "'," & IntYear & "," & intMonth & ",'" & rsPlanBudget("Medium_Code") & "'," & MSC & "," & MSC_Flag & "," & Club_Agency & "," & Club_Agency_Flag & "," & Bonus_Comm & "," & Bonus_Comm_Flag & ",'Open')"
                ConnERP.Execute strSql
                'insert Budget_Control_Plan_Detail
                strSql = " Insert into Budget_Control_Plan_Detail (Brand_Code,[Year],[Month],Media_Code,Plan_Budget)"
                strSql = strSql & " values('" & rsbrand(0) & "'," & IntYear & "," & intMonth & ",'" & rsPlanBudget("medium_code") & "'," & rsPlanBudget("Budget") & ")"
                ConnERP.Execute strSql
                rsPlanBudget.MoveNext
            Wend
            rsPlanBudget.Close
        Else
            'BU2/NON ULI
            'Delete Budget_Control_non_uli
            strSql = "Delete from Budget_Control_Non_ULI where brand_code='" & rsbrand(0) & "' and [year]=" & IntYear
            ConnERP.Execute strSql
            
            strSql = " select month_number,medium_code,sum(budget) budget, "
            strSql = strSql & " sum(gross_budget) gross_budget from"
            strSql = strSql & " ("
            strSql = strSql & " SELECT case b.medium_code"
            strSql = strSql & "     when 'CN' then 'OT'"
            strSql = strSql & "     when 'OT' then 'OT'"
            strSql = strSql & "     when 'RD' then 'RD'"
            strSql = strSql & "     when 'PR' then 'PR'"
            strSql = strSql & "     when 'TV' then 'TV'"
            strSql = strSql & "     end medium_code,"
            strSql = strSql & " a.month_number,"
            strSql = strSql & " case isnull(a.total_actual,-1) when -1 then a.min_budget else a.actual_nett_paid end budget,"
            strSql = strSql & " case isnull(a.total_actual,-1) when -1 then a.gross_budget else a.actual_gross_paid end gross_budget FROM mp_monthly_activity a  "
            strSql = strSql & " INNER JOIN mp_medium b on a.mp_medium_id = b.mp_medium_id  "
            strSql = strSql & " INNER JOIN mp_activity c on b.mp_activity_id = c.mp_activity_id  "
            strSql = strSql & " INNER JOIN mp_task d on c.mp_task_id = d.mp_task_id "
            'parameter here
            strSql = strSql & " and c.original_brand_code = '" & rsbrand(0) & "' "
            strSql = strSql & " and d.mp_number = '" & strMPNumber & "' "
            strSql = strSql & " ) x"
            strSql = strSql & " GROUP BY month_number,medium_code "
            strSql = strSql & " order by month_number,medium_code"
            intMonth = Empty
            rsPlanBudget.Open strSql, ConnERP
            While Not rsPlanBudget.EOF
                If intMonth <> rsPlanBudget("month_number") Then
                    'get brand fee for current month
                    intMonth = rsPlanBudget("month_number")
                    
                    MSC = Empty
                    MSC_Flag = Empty
                    Bonus_Comm = Empty
                    Bonus_Comm_Flag = Empty
                    Club_Agency = Empty
                    Club_Agency_Flag = Empty
                    
                    strSql = " select MSC_Paid,MSC_Paid_On_Flag,MSC_Bonus,MSC_Bonus_On_Flag,Club_Agency,Club_Agency_On_Flag from brand_fee "
                    strSql = strSql & " where brand_code = '" & rsbrand(0) & "' and [month] = " & intMonth & " and [year] = " & IntYear
                    rsBrandFee.Open strSql, ConnERP
                    If Not rsBrandFee.EOF Then
                        MSC = rsBrandFee(0)
                        MSC_Flag = rsBrandFee(1)
                        Bonus_Comm = rsBrandFee(2)
                        Bonus_Comm_Flag = rsBrandFee(3)
                        Club_Agency = rsBrandFee(4)
                        Club_Agency_Flag = rsBrandFee(5)
                    End If
                    rsBrandFee.Close
                End If
                'insert Budget_Control_Non_Uli
                strSql = " insert into budget_control_non_uli(Brand_Code,[Year],[Month],Media_Code,Msc,Msc_Flag,Club_Agency,Club_Agency_Flag,Bonus_Comm,Bonus_Comm_Flag,Budget,Budget_Gross) "
                strSql = strSql & " values('" & rsbrand(0) & "'," & IntYear & "," & intMonth & ",'" & rsPlanBudget("medium_code") & "'," & MSC & "," & MSC_Flag & "," & Club_Agency & "," & Club_Agency_Flag & "," & Bonus_Comm & "," & Bonus_Comm_Flag & "," & rsPlanBudget("Budget") & "," & rsPlanBudget("Gross_Budget") & ")"
                ConnERP.Execute strSql
                rsPlanBudget.MoveNext
            Wend
            rsPlanBudget.Close
        End If
        rsbrand.MoveNext
    Wend
    rsbrand.Close
    Set rsbrand = Nothing
    Set rsPlanBudget = Nothing
    Set rsBrandFee = Nothing
End Sub

Private Sub db_Resend_Mail()
'<CSCM>
'********************************************************************************
'Procedure Name     : db_Resend_Mail
'Procedure Function : Send Email
'Input Parameter    : ---
'Output Parameter   : ---
'Date               : 3/17/2016
'LastUpdate/By      : Tedi / Kreatif
'Name Before        : Cmd_Resend_Mail_Click
'********************************************************************************
'</CSCM>

    frm_MPSendMail.show 1
    
End Sub

Private Sub cmdApproval_Click()
    Frm_MPApprovalNew.show 1
End Sub

Private Sub cmdCancel_Click()
    FrameViewMonth1.Visible = False
End Sub

Private Sub db_Copy()
'<CSCM>
'********************************************************************************
'Procedure Name     : db_Copy
'Procedure Function : ---
'Input Parameter    : ---
'Output Parameter   : ---
'Date               : 3/10/2016
'LastUpdate/By      : Tedi / Kreatif
'Name Before        : CmdCopy_Click
'********************************************************************************
'</CSCM>

    If cboMPNumber.Text = "" Then
        MsgBox "Select MP Number!", vbExclamation, strApplication_Name
        Exit Sub
    End If
    
    If MsgBox("Copy Media Plan?", vbYesNo + vbQuestion, strApplication_Name) = vbYes Then
        Call db_CopyPlan
    End If

End Sub

Private Sub cmdEditGivenBudget_Click()
    Dim rsTemp As New ADODB.Recordset, valid_or_cancel As Boolean
    Dim pesan
    Dim strbudget As String
    If cboMPNumber.Text <> "" Then
        rsTemp.Open "select isnull(yearly_budget,0) from mp_master where mp_number='" & cboMPNumber.Text & "'", ConnERP, 1, 3
        If Not rsTemp.EOF Then
            valid_or_cancel = False
            Do
                pesan = InputBox("Enter Given Budget : ", "Edit Given Budget", FormatNumber(rsTemp(0), 2))
                If pesan <> Empty Then
                    strbudget = RemoveNumberFormat(CStr(pesan))
                    If IsNumeric(pesan) Then
                        ConnERP.Execute "update mp_master set yearly_budget = " & CDbl(pesan) & " where mp_number = '" & cboMPNumber.Text & "'"
                        cboMPNumber.Text = cboMPNumber.Text
                        valid_or_cancel = True
                    Else
                        MsgBox "Invalid Entry!", vbExclamation, strApplication_Name
                    End If
                Else
                    valid_or_cancel = True
                End If
            Loop Until valid_or_cancel
        End If
        rsTemp.Close
    Else
        MsgBox ("Select MP Number!", vbExclamation, strApplication_Name)
    End If
End Sub

Private Sub cmdOK_Click()
    Dim strMonth As String, i As Integer, AllMonth As Boolean
    AllMonth = True
    If cbkAllMonth.Value = 1 Then
        txtViewMonth.Text = "ALL"
        strViewMonth = "ALL"
    Else
        strMonth = ""
        strViewMonth = ""
        For i = 1 To 12
            If cbkViewMonth(i).Value = 1 Then
                strMonth = strMonth & EngMonthName(i) & ","
                strViewMonth = strViewMonth & CStr(i) & ","
            Else
                AllMonth = False
            End If
        Next
        If AllMonth Then
            txtViewMonth.Text = "ALL"
            strViewMonth = "ALL"
        Else
            If Len(strMonth) <> 0 Then
                txtViewMonth.Text = Left(strMonth, Len(strMonth) - 1)
                strViewMonth = "(" & Left(strViewMonth, Len(strViewMonth) - 1) & ")"
            Else
                MsgBox "You must select at least 1 month!", vbExclamation, strApplication_Name
                Exit Sub
            End If
        End If
    End If
    'MsgBox strViewMonth
    FrameViewMonth1.Visible = False
    frm_MPInsertion.Refresh
    Call cboMPNumber_Click
    
End Sub

Private Sub db_SummarybyVariant()
'<CSCM>
'********************************************************************************
'Procedure Name     : db_SummarybyVariant
'Procedure Function : Proses Summary By Variant
'Input Parameter    : ---
'Output Parameter   : ---
'Date               : 3/10/2016
'LastUpdate/By      : Tedi / Kreatif
'Name Before        : cmdSummarybyVariant_Click
'********************************************************************************
'</CSCM>

    Me.MousePointer = vbHourglass
    If cboMPNumber.Text <> "" Then
        frm_MPSummery.str_mp_number = cboMPNumber.Text
        frm_MPSummery.show
'Frm_MPTotalByVariant.Show 1
    Else
        MsgBox "Select MP Number!", vbExclamation, strApplication_Name
    End If

    Me.MousePointer = vbDefault
    
End Sub

Private Sub db_TVLayering()
'<CSCM>
'********************************************************************************
'Procedure Name     : db_TVLayering
'Procedure Function : ---
'Input Parameter    : ---
'Output Parameter   : ---
'Date               : 3/10/2016
'LastUpdate/By      : Tedi / Kreatif
'Name Before        : cmdTVLayering_Click
'********************************************************************************
'</CSCM>
       
    If cboMPNumber.Text <> "" Then
        frm_MPTVLayering.show
    Else
        MsgBox "Select MP Number!", vbExclamation, strApplication_Name
    End If
    
End Sub

Sub DisableResendMail(ByVal blnEnabled As Boolean)
'<CSCM>
'********************************************************************************
'Procedure Name     : DisableResendMail
'Procedure Function : Disable/Enabled icon Resend Mail
'Input Parameter    : ---
'Output Parameter   : ---
'Date               : 3/17/2016
'LastUpdate/By      : Tedi / Kreatif
'Name Before        : -
'********************************************************************************
'</CSCM>

    With picButton(enButtonType.bieResendEmail)  'SummarybyVariant 63.
        .Enabled = Not blnEnabled
    End With
    
End Sub

Sub DisablePublish_To_Web(ByVal blnEnabled As Boolean)
'<CSCM>
'********************************************************************************
'Procedure Name     : DisablePublish_To_Web
'Procedure Function : Disable/Enabled Publish_To_Web
'Input Parameter    : ---
'Output Parameter   : ---
'Date               : 3/17/2016
'LastUpdate/By      : Tedi / Kreatif
'Name Before        : -
'********************************************************************************
'</CSCM>

    With picButton(enButtonType.biePublishToWeb)  'SummarybyVariant 63.
        .Enabled = Not blnEnabled
    End With
    
End Sub

Sub Form_Load()
    Call EnableObject(False)
    Call initform
        
    DisableResendMail True
    
    '=======Resize Form==========
    'Resize Me, "[PicButtonPanahBawah][picViewMonth][FrameViewMonth1][FrameViewMonth2][cmdOK][cmdCancel][cbkAllMonth][cbkViewMonth] "
    picViewMonth.Left = txtViewMonth.Left + txtViewMonth.Width - picViewMonth.Width - 30
    picViewMonth.Top = txtViewMonth.Top + 30
    FrameViewMonth1.Top = fra_Insertion.Top + pic_Grd.Top + txtViewMonth.Top + txtViewMonth.Height + 60
    '============================
    strViewMonth = "ALL"
    txtViewMonth.Text = "ALL"
    cbkAllMonth.Value = 1
    dblUnlockedCellColor = vbRed
    
End Sub

Private Sub db_CopyPlan()
'*****************************************************************************
' Nama Prosedur     :   db_CopyPlan
' Fungsi Prosedur   :   Copy Media Plan
' Parameter  Input  :
' Parameter Output  :
' Tgl Pembuatan     :   08 Sep 2004
' Last Update/By    :   07 July 2005/Sistyo
' Name Before       :   CopyPlan
'*****************************************************************************
    
    Dim strSql As String
    Dim rsTemp As New ADODB.Recordset
    Dim rsMPMaster As New ADODB.Recordset
    Dim rsMPTask As New ADODB.Recordset
    Dim rsMPActivity As New ADODB.Recordset
    Dim rsMPMedium As New ADODB.Recordset, RsMPMonthlyActivity As New ADODB.Recordset
    Dim rsMPMediumDetail As New ADODB.Recordset
    Dim rsMPPlanDimension As New ADODB.Recordset
    Dim rsMPInsertion As New ADODB.Recordset
    Dim rsMPTVReachFreq As New ADODB.Recordset
    Dim rsMPOtherMonthlyBudget As New ADODB.Recordset
    Dim strNewMPNumber As String, strNewTaskID  As String, strNewActivityID As String, strNewMediumID As String
    Dim strNewMediumDetailID As String, strNewPlanDimID As String
    
    Dim i As Integer
    FrameProgressBarCopy.Visible = True
    FrameProgressBarCopy.Refresh
    
    strSql = "select * from mp_master where mp_number='" & cboMPNumber.Text & "'"
    rsMPMaster.Open strSql, ConnERP, 1, 3
    ProgressBarCopy.Max = 1000
    ProgressBarCopy.Min = 1
    ProgressBarCopy.Value = 1
    
    If Not rsMPMaster.EOF Then
'        Call incrProgressBarCopy
        'Generate MP_Number
        strSql = "select isnull(max(cast(substring(mp_number,11,4) as int)),0)+1 from mp_master where mp_number like '" & Mid(cboMPNumber.Text, 1, 10) & "%'"
        rsTemp.Open strSql, ConnERP, 1, 3
            strNewMPNumber = Mid(cboMPNumber, 1, 10) & Right("0000" & CStr(rsTemp(0)), 4)
        rsTemp.Close
        'Copy MPMaster
        strSql = "insert into mp_master select '" & strNewMPNumber & "' mp_number,"
        For i = 1 To rsMPMaster.Fields.Count - 1
            strSql = strSql & rsMPMaster.Fields(i).Name & ","
        Next
        strSql = Mid(strSql, 1, Len(strSql) - 1)
        strSql = strSql & " from mp_master where mp_number = '" & cboMPNumber.Text & "'"
        ConnERP.Execute strSql
        'Update islatest source mp to 0
        strSql = "update mp_master set is_latest=0 where mp_number='" & cboMPNumber.Text & "'"
        ConnERP.Execute strSql
        'update islatest new mp to null (flag copy plan, untuk skip trigger)
        strSql = "update mp_master set is_latest=null where mp_number='" & strNewMPNumber & "'"
        ConnERP.Execute strSql
        
        'MPTask
        strSql = "select * from mp_task where mp_number = '" & cboMPNumber.Text & "'"
        rsMPTask.Open strSql, ConnERP, 1, 3
        
        While Not rsMPTask.EOF
            Call incrProgressBarCopy
            'Generate Task ID
            strSql = "select isnull(max(cast(substring(mp_task_id,16,4) as int)),0)+1 from mp_task where mp_number like '" & Mid(cboMPNumber.Text, 1, 9) & "%'"
            rsTemp.Open strSql, ConnERP, 1, 3
                strNewTaskID = Mid(cboMPNumber.Text, 1, 4) & ".TASK." & Mid(cboMPNumber.Text, 6, 4) & "." & Right("0000" & CStr(rsTemp(0)), 4) 'Create new Task_Id
            rsTemp.Close
            
            strSql = "insert into mp_task select '" & strNewTaskID & "','" & strNewMPNumber & "',"
            For i = 2 To rsMPTask.Fields.Count - 1
                strSql = strSql & rsMPTask.Fields(i).Name & ","
            Next
            strSql = Left(strSql, Len(strSql) - 1)
            strSql = strSql & " from mp_task where mp_task_id = '" & rsMPTask("mp_task_id") & "'"
            ConnERP.Execute strSql
            'MPActivity
            strSql = "select * from mp_activity where mp_task_id = '" & rsMPTask("mp_task_id") & "'"
            rsMPActivity.Open strSql, ConnERP, 1, 3
            While Not rsMPActivity.EOF
                Call incrProgressBarCopy
                'Generate Activity ID
                strSql = "select isnull(max(cast(substring(mp_activity_id,16,4) as int)),0)+1 from mp_activity where mp_task_id like '" & Mid(rsMPTask(0), 1, 15) & "%'"
                rsTemp.Open strSql, ConnERP, 1, 3
                    strNewActivityID = Mid(rsMPTask(0), 1, 4) & ".ACTV." & Mid(rsMPTask(0), 11, 4) & "." & Right("0000" & CStr(rsTemp(0)), 4) 'Create new Activity_Id
                rsTemp.Close
                
                strSql = "insert into mp_activity select '" & strNewActivityID & "','" & strNewTaskID & "',"
                For i = 2 To rsMPActivity.Fields.Count - 1
                    strSql = strSql & rsMPActivity.Fields(i).Name & ","
                Next
                strSql = Left(strSql, Len(strSql) - 1)
                strSql = strSql & " from mp_activity where mp_activity_id = '" & rsMPActivity("mp_activity_id") & "'"
                ConnERP.Execute strSql
                'MPMedium
                strSql = "select * from mp_medium where mp_activity_id = '" & rsMPActivity("mp_activity_id") & "'"
                rsMPMedium.Open strSql, ConnERP, 1, 3
                
                While Not rsMPMedium.EOF
                    Call incrProgressBarCopy
                    'Generate mediumID
                    strNewMediumID = NextMPMediumID(rsMPActivity("mp_activity_id"))
                    strSql = "insert into mp_medium select '" & strNewMediumID & "','" & strNewActivityID & "',"
                    For i = 2 To rsMPMedium.Fields.Count - 1
                        strSql = strSql & rsMPMedium.Fields(i).Name & ","
                    Next
                    strSql = Left(strSql, Len(strSql) - 1)
                    strSql = strSql & " from mp_medium where mp_medium_id = '" & rsMPMedium("mp_medium_id") & "'"
                    ConnERP.Execute strSql
                    
                    'MPMonthlyACtivity
                    strSql = "select * from mp_monthly_activity where mp_medium_id = '" & rsMPMedium("mp_medium_id") & "'"
                    RsMPMonthlyActivity.Open strSql, ConnERP, 1, 3
                    If Not RsMPMonthlyActivity.EOF Then
                        'Call incrProgressBarCopy
                        strSql = "insert into mp_monthly_activity select '" & strNewMediumID & "',"
                        For i = 1 To RsMPMonthlyActivity.Fields.Count - 1
                            strSql = strSql & RsMPMonthlyActivity.Fields(i).Name & ","
                        Next
                        strSql = Left(strSql, Len(strSql) - 1)
                        strSql = strSql & " from mp_monthly_activity where mp_medium_id = '" & RsMPMonthlyActivity("mp_medium_id") & "'"
                        ConnERP.Execute strSql
                    End If
                    RsMPMonthlyActivity.Close
                    
                    'MPMediumDetail
                    strSql = "select * from mp_medium_detail where mp_medium_id = '" & rsMPMedium("mp_medium_id") & "'"
                    rsMPMediumDetail.Open strSql, ConnERP, 1, 3
                    While Not rsMPMediumDetail.EOF
                        Call incrProgressBarCopy
                        'GEnerate MediumDetailID
                        strNewMediumDetailID = NextMPMediumDetailID(rsMPMedium("mp_medium_id"))
                        strSql = "insert into mp_medium_detail select '" & strNewMediumDetailID & "','" & strNewMediumID & "',"
                        For i = 2 To rsMPMediumDetail.Fields.Count - 1
                            strSql = strSql & rsMPMediumDetail.Fields(i).Name & ","
                        Next
                        strSql = Left(strSql, Len(strSql) - 1)
                        strSql = strSql & " from mp_medium_detail where mp_medium_detail_id = '" & rsMPMediumDetail("mp_medium_detail_id") & "'"
                        ConnERP.Execute strSql
                        
                        'MPPlanDimension
                        strSql = "select * from mp_plan_dimension where mp_medium_detail_id = '" & rsMPMediumDetail("mp_medium_detail_id") & "'"
                        rsMPPlanDimension.Open strSql, ConnERP, 1, 3
                        While Not rsMPPlanDimension.EOF
                            Call incrProgressBarCopy
                            'generate new PlanDimID
                            strNewPlanDimID = NextMPPlanDimID(rsMPMediumDetail("mp_medium_detail_id"))
                            strSql = "insert into mp_plan_dimension select '" & strNewPlanDimID & "','" & strNewMediumDetailID & "',"
                            For i = 2 To rsMPPlanDimension.Fields.Count - 1
                                strSql = strSql & rsMPPlanDimension.Fields(i).Name & ","
                            Next
                            strSql = Left(strSql, Len(strSql) - 1)
                            strSql = strSql & " from mp_plan_dimension where mp_plan_dim_id = '" & rsMPPlanDimension("mp_plan_dim_id") & "'"
                            ConnERP.Execute strSql
                            
                            'MPTVReachFreq
                            strSql = "select * from mp_tv_reach_frequency where mp_plan_dim_id = '" & rsMPPlanDimension("mp_plan_dim_id") & "'"
                            rsMPTVReachFreq.Open strSql, ConnERP, 1, 3
                            While Not rsMPTVReachFreq.EOF
                                Call incrProgressBarCopy
                                strSql = "insert into mp_tv_reach_frequency("
                                For i = 1 To rsMPTVReachFreq.Fields.Count - 1
                                    strSql = strSql & rsMPTVReachFreq.Fields(i).Name & ","
                                Next
                                strSql = Left(strSql, Len(strSql) - 1)
                                
                                strSql = strSql & ") select '" & strNewPlanDimID & "',"
                                For i = 2 To rsMPTVReachFreq.Fields.Count - 2
                                    strSql = strSql & rsMPTVReachFreq.Fields(i).Name & ","
                                Next
                                'strSQL = Left(strSQL, Len(strSQL) - 1)
                                strSql = strSql & "mp_tv_rf_id from mp_tv_reach_frequency where mp_plan_dim_id = '" & rsMPPlanDimension("mp_plan_dim_id") & "' and mp_tv_rf_id=" & rsMPTVReachFreq("mp_tv_rf_id")
                                ConnERP.Execute strSql
                                rsMPTVReachFreq.MoveNext
                            Wend
                            rsMPTVReachFreq.Close
                            
                            'MPInsertion
                            strSql = "select * from mp_insertion where mp_plan_dim_id = '" & rsMPPlanDimension("mp_plan_dim_id") & "'"
                            rsMPInsertion.Open strSql, ConnERP, 1, 3
                            While Not rsMPInsertion.EOF
                                Call incrProgressBarCopy
                                strSql = "insert into mp_insertion select '" & strNewPlanDimID & "',"
                                For i = 1 To rsMPInsertion.Fields.Count - 1
                                    strSql = strSql & rsMPInsertion.Fields(i).Name & ","
                                Next
                                strSql = Left(strSql, Len(strSql) - 1)
                                strSql = strSql & " from mp_insertion where mp_plan_dim_id = '" & rsMPPlanDimension("mp_plan_dim_id") & "' and week_year = " & rsMPInsertion("week_year")
                                ConnERP.Execute strSql
                                strSql = "update mp_insertion set mp_tv_rf_id = (select mp_tv_rf_id from mp_tv_reach_frequency where old_mp_tv_rf_id = mp_insertion.mp_tv_rf_id) where mp_plan_dim_id = '" & strNewPlanDimID & "' and week_year = " & rsMPInsertion("week_year")
                                ConnERP.Execute strSql
                                rsMPInsertion.MoveNext
                            Wend
                            rsMPInsertion.Close
                            
                            'MPOtherMonthlyBudget
                            strSql = "select * from mp_other_monthly_budget where mp_plan_dim_id = '" & rsMPPlanDimension("mp_plan_dim_id") & "'"
                            rsMPOtherMonthlyBudget.Open strSql, ConnERP, 1, 3
                            While Not rsMPOtherMonthlyBudget.EOF
                                Call incrProgressBarCopy
                                strSql = "insert into mp_other_monthly_budget select '" & strNewPlanDimID & "',"
                                For i = 1 To rsMPOtherMonthlyBudget.Fields.Count - 1
                                    strSql = strSql & rsMPOtherMonthlyBudget.Fields(i).Name & ","
                                Next
                                strSql = Left(strSql, Len(strSql) - 1)
                                strSql = strSql & " from mp_other_monthly_budget where mp_plan_dim_id = '" & rsMPPlanDimension("mp_plan_dim_id") & "' and month_number = " & rsMPOtherMonthlyBudget("month_number")
                                ConnERP.Execute strSql
                                rsMPOtherMonthlyBudget.MoveNext
                            Wend
                            rsMPOtherMonthlyBudget.Close
                            
                            rsMPPlanDimension.MoveNext
                        Wend
                        rsMPPlanDimension.Close
                        rsMPMediumDetail.MoveNext
                    Wend
                    rsMPMediumDetail.Close
                    rsMPMedium.MoveNext
                Wend
                rsMPMedium.Close
                rsMPActivity.MoveNext
            Wend
            rsMPActivity.Close
            rsMPTask.MoveNext
        Wend
        rsMPTask.Close
    End If
    ProgressBarCopy.Value = 1000
    rsMPMaster.Close
     
    FrameProgressBarCopy.Visible = False
    'Copy selesai..
    
    'Update Monthly_activity (Status Approval sampai N-1)
    'strSQL = "update mp_monthly_activity set approval=0 where mp_medium_id in (select mp_medium_id from mp_ids where mp_number='" & strNewMPNumber & "') and (substring(mp_medium_id,11,4) * 100 + month_number) > (year(getdate()) * 100 + month(getdate())-1)"
    'ConnERP.Execute strSQL
    
    '===================================COPY MP_MONTHLY_QUOTATION===================================
    
    Dim rsMonthlyQuotation As New ADODB.Recordset
    Dim rsRunningNumber As New ADODB.Recordset
    Dim str_running_number As String
    Dim str_mp_monthly_quot_id As String
  
    
    rsMonthlyQuotation.Open "select * from MP_Monthly_Quotation where mp_number = '" & cboMPNumber.Text & "'", ConnERP, 1, 3

    While Not rsMonthlyQuotation.EOF
        
        Call incrProgressBarCopy
        'Generate new Quotation ID
        rsRunningNumber.Open "select isnull(max(cast(substring(mp_monthly_quot_id,11,4) as int)),0)+1 from mp_monthly_quotation where mp_monthly_quot_id like '" & Mid(rsMonthlyQuotation("mp_Monthly_Quot_ID"), 1, 10) & "%'", ConnERP, 1, 3
            str_running_number = Right("0000" & CStr(rsRunningNumber(0)), 4)
        rsRunningNumber.Close
        str_mp_monthly_quot_id = Mid(rsMonthlyQuotation("mp_monthly_quot_id"), 1, 10) & str_running_number
        'Generate strSQL
        strSql = "insert into mp_monthly_quotation select "
        For i = 0 To rsMonthlyQuotation.Fields.Count - 1
            Select Case UCase(rsMonthlyQuotation.Fields(i).Name)
                Case "MP_MONTHLY_QUOT_ID"
                    strSql = strSql & "'" & str_mp_monthly_quot_id & "'"
                Case "MP_NUMBER"
                    strSql = strSql & "'" & strNewMPNumber & "'"
                Case Else
                    strSql = strSql & rsMonthlyQuotation.Fields(i).Name
            End Select
            If i < rsMonthlyQuotation.Fields.Count - 1 Then
                strSql = strSql & ","
            End If
        Next
        strSql = strSql & " from mp_monthly_quotation where mp_Monthly_Quot_ID = '" & rsMonthlyQuotation("mp_monthly_quot_id") & "'"
        ConnERP.Execute strSql
        
        rsMonthlyQuotation.MoveNext
        
    Wend
    rsMonthlyQuotation.Close
    Set rsMonthlyQuotation = Nothing
    Set rsRunningNumber = Nothing
    
    'update is_latest original plan
    ConnERP.Execute "update MP_Monthly_Quotation set is_latest = 0 where mp_number = '" & cboMPNumber.Text & "'"
        
    '===================================END OF COPY MP_MONTHLY_QUOTATION===================================
        
    'Update MP_Master
    strSql = "update mp_master set is_latest = 1, is_upload_to_web = 0 where mp_number = '" & strNewMPNumber & "'"
    ConnERP.Execute strSql
    
    MsgBox "Media Plan Copied!" & vbCrLf & "New MP Number = " & strNewMPNumber, vbExclamation, strApplication_Name
    
    cboMPNumber.RemoveItem cboMPNumber.ListIndex 'trash to history
    cboMPNumber.AddItem strNewMPNumber 'add new plan number
    cboMPNumber.Text = strNewMPNumber 'show new plan number
    Call cboMPNumber_Click
End Sub

Private Sub initform()
'*****************************************************************************
' Nama Submodul         :  initForm
' Fungsi Submodul       :  Inisialisasi Form
' Nama Programmer       :  Sistyo
' Tgl Pembuatan         :  16 Agustus 2004
' Last Update           :  16 Agustus 2004/Sistyo
'******************************************************************************
    Dim baris As Integer, Kolom As Integer, i As Integer
    Dim strBrand_Filter As String
    Dim strBrandSample As String, strSql As String
    
    bol_BU2 = False
    strBrandSample = ""
    
    'hide menu klik kanan
'        MnuFGMPInsertion.Visible = False
    'Load MP Number
        strBrand_Filter = "select brand_code from media_security_catalog where user_name = '" & strLogin_User & "' and position='Planner' and valid_until>=(select getdate())"
        
        rsTemp.Open "select mp_number,brand_code from mp_master where brand_code in (" & strBrand_Filter & ") and is_latest = 1 order by mp_number", ConnERP, 1, 3
        If Not rsTemp.EOF Then
            strBrandSample = rsTemp("brand_code").Value
        End If
        
        While Not rsTemp.EOF
            cboMPNumber.AddItem rsTemp(0)
            rsTemp.MoveNext
            
        Wend
        rsTemp.Close
    'Cek APAKAH USER BU2
        If strBrandSample <> "" Then
            strSql = "select count(*) from client_special where client_code in (select client_code from brand where brand_code = '" & strBrandSample & "')"
            rsTemp.Open strSql, ConnERP, 1, 3
            If rsTemp(0) = 0 Then
                bol_BU2 = True
            End If
            rsTemp.Close
        End If
    
    'Load cboMonth
        
    'Clear Header
        
        lblReleaseDate.Caption = ""
        txtBrandName.Text = ""
        txtClientName.Text = ""
        txtCreatedDate.Text = ""
        txtCreatedBy.Text = ""
        txtLastUpdateDate.Text = ""
        txtLastUpdateBy.Text = ""
        
    With msf_MPInsertion
        'init FGMPInsertion Header
            .MergeCol(0) = True
            .TextMatrix(0, 0) = " "
            .TextMatrix(1, 0) = " "
            .TextMatrix(2, 0) = " "
            .TextMatrix(3, 0) = " "
            .MergeCol(1) = True
            .TextMatrix(0, 1) = "Marketing Task"
            .TextMatrix(1, 1) = "Marketing Task"
            .TextMatrix(2, 1) = "Marketing Task"
            .TextMatrix(3, 1) = "Marketing Task"
            .MergeCol(2) = True
            .TextMatrix(0, 2) = "Channel"
            .TextMatrix(1, 2) = "Channel"
            .TextMatrix(2, 2) = "Channel"
            .TextMatrix(3, 2) = "Channel"
            .MergeCol(3) = True
            .TextMatrix(0, 3) = "Version"
            .TextMatrix(1, 3) = "Version"
            .TextMatrix(2, 3) = "Version"
            .TextMatrix(3, 3) = "Version"
            .MergeCol(4) = True
            .TextMatrix(0, 4) = "Media"
            .TextMatrix(1, 4) = "Media"
            .TextMatrix(2, 4) = "Media"
            .TextMatrix(3, 4) = "Media"
            .MergeCol(5) = True
            .TextMatrix(0, 5) = "Channel Dimension"
            .TextMatrix(1, 5) = "Channel Dimension"
            .TextMatrix(2, 5) = "Channel Dimension"
            .TextMatrix(3, 5) = "Channel Dimension"
            .ColWidth(0) = 1400
            .ColWidth(1) = 3000
            .ColWidth(2) = 1100
            .ColWidth(3) = 1300
            .ColWidth(4) = 1800
            .ColWidth(5) = 2000
            .RowHeight(4) = 0
            
            For baris = 0 To 3
                .Row = baris
                For Kolom = 1 To 5
                    .col = Kolom
                    .CellAlignment = 4
                Next
            Next
            .GridColor = vbBlack
            
        'Load TV Frequency
            rsTemp.Open "select code,frequency_name from frequency_catalog order by code", ConnERP, 1, 3
            While Not rsTemp.EOF
                LstFrequency.AddItem rsTemp(1)
                rsTemp.MoveNext
            Wend
            rsTemp.Close
    End With
    
End Sub


Private Sub cboMPNumber_Click()
'*****************************************************************************
' Nama Submodul         :  cboMPNumber_Click
' Fungsi Submodul       :  view media plan in grid untuk insertion
' Nama Programmer       :  Sistyo
' Tgl Pembuatan         :  16 Agustus 2004
' Last Update           :  19 Agustus 2004/Sistyo
'******************************************************************************
    Me.MousePointer = vbHourglass

    Dim strSql As String, i As Double, counter As Double, insrow As Double
    Dim rsMPMaster As New ADODB.Recordset, rsMPTask As New ADODB.Recordset, rsMPActivity As New ADODB.Recordset
    Dim rsMPMedium As New ADODB.Recordset, rsMPMediumDetail As New ADODB.Recordset, rsBudget As New ADODB.Recordset
    Dim rsMPInsertion As New ADODB.Recordset, rsMPTVReachFrequency As New ADODB.Recordset
    Dim rsMPOtherMonthlyBudget As New ADODB.Recordset, rsFee As New ADODB.Recordset
    Dim arrFee() As Double
    Dim intPeriode As Double, strSpotType As String, strVersion As String
    Dim jumlahActivity As Double, intGrpIns As Double
    
    Dim intStartBlockMedium As Double, intRowBlockMedium As Double
    Dim Remaining_Budget As Double
    Dim intTotalNett As Double
    Dim field_no As Integer
    Dim feePaid As Double, FeeBonus As Double, FeeClub As Double
    
    Dim intTotalBudgetPlan As Double
    Dim intTotalBudgetActual As Double
    '==============================SETTING BUTTON============================
'    cmdEdit.Enabled = True
'    cmdSummary.Enabled = True
'    cmdSummarybyVariant.Enabled = True
'    cmdExportToExcel.Enabled = True
    FlagMarkStart = False
    PicButtonPanahBawah.Visible = False
    LstFrequency.Visible = False
    '==================================
    
    If cboMPNumber.Text <> "" Then
        cmdApproval.Enabled = True
    Else
        cmdApproval.Enabled = False
    End If
    
    'Clear Grid
        With msf_MPInsertion
            .Rows = 6
            For i = 1 To .Rows
                .TextMatrix(.Rows - 1, i - 1) = ""
            Next
        End With
        
    'Load Week Commencing
        Call LoadWeekCommencing(Mid(cboMPNumber.Text, 6, 4))
    
    'Load MP_Master
        strSql = "select mp_number,brand_name,client_name,created_date,created_by,last_update_date,last_update_by,release_date,isnull(yearly_budget,0) yearly_budget,isnull(is_upload_to_web,0) is_upload_to_web,isnull(is_latest,0) is_Latest from mp_master where mp_number = '" & cboMPNumber.Text & "'"
        rsMPMaster.Open strSql, ConnERP, 1, 3
        jumlahActivity = 0
        Remaining_Budget = 0
        If Not rsMPMaster.EOF Then
            'cek is publish to web
                If rsMPMaster("is_upload_to_web") = 1 Then
                    Call DisablePublish_To_Web(True)
                    If rsMPMaster("is_Latest") = 1 Then
                        Call DisableResendMail(False)
                    Else
                        Call DisableResendMail(True)
                    End If
                Else
                    Call DisablePublish_To_Web(False)
                    Call DisableResendMail(True)
                End If
                
            
            'Cek Release Date
                If IsNull(rsMPMaster(7)) Then
                    lblReleaseDate.Caption = "not set"
                    lblReleaseDate.ForeColor = vbRed
                Else
                    lblReleaseDate.Caption = Format(rsMPMaster(7), "dd mmm yyyy")
                    lblReleaseDate.ForeColor = vbBlue
                End If
            'Load Header
                txtBrandName.Text = rsMPMaster(1)
                txtClientName.Text = rsMPMaster(2)
                txtCreatedDate.Text = Format(rsMPMaster(3), "DD MMM YYYY")
                txtCreatedBy.Text = rsMPMaster(4)
                txtLastUpdateDate.Text = Format(rsMPMaster(5), "DD MMM YYYY")
                txtLastUpdateBy.Text = ReplaceNull(rsMPMaster(6))
                Remaining_Budget = rsMPMaster("Yearly_Budget")
                txt_remaining_budget.Text = FormatNumber(Remaining_Budget, 2)
                
                If CDbl(RemoveNumberFormat(txt_remaining_budget.Text)) < 0 Then
                    txt_remaining_budget.ForeColor = vbRed
                Else
                    txt_remaining_budget.ForeColor = vbBlack
                End If
                
            With msf_MPInsertion
                'Load MP_Task
                    
                    rsMPTask.Open "select mp_task_id,Task_desc from mp_task where mp_number = '" & rsMPMaster(0) & "'", ConnERP, 1, 3
                    While Not rsMPTask.EOF
                        If .Rows > 6 Then
                            .Rows = .Rows + 3
                        End If
                        .TextMatrix(.Rows - 1, 0) = "Task"
                        .TextMatrix(.Rows - 1, 1) = rsMPTask(1)
                        'load mp_activity
                            rsMPActivity.Open "select mp_activity_id,activity_type,activity_desc,brand_variant_name,target_audience,brand_target from mp_activity where mp_task_id = '" & rsMPTask(0) & "'", ConnERP, 1, 3
                            While Not rsMPActivity.EOF
                                jumlahActivity = jumlahActivity + 1
                                .Rows = .Rows + 4
                                .TextMatrix(.Rows - 4, 0) = "Activity": .TextMatrix(.Rows - 4, 1) = rsMPActivity(1) & " (" & rsMPActivity(2) & ")"
                                .TextMatrix(.Rows - 3, 0) = "Variant": .TextMatrix(.Rows - 3, 1) = rsMPActivity(3)
                                .TextMatrix(.Rows - 2, 0) = "Target Audience": .TextMatrix(.Rows - 2, 1) = rsMPActivity(4)
                                .TextMatrix(.Rows - 1, 0) = "Brand Target": .TextMatrix(.Rows - 1, 1) = rsMPActivity(5)
                                'load mp_medium
                                    rsMPMedium.Open "select mp_medium_id,medium_name from mp_medium where mp_activity_id = '" & rsMPActivity(0) & "' order by medium_name desc", ConnERP, 1, 3
                                    i = .Rows - 4
                                    While Not rsMPMedium.EOF
                                        If i > .Rows - 2 Then
                                           .Rows = .Rows + 2
                                        End If
                                        intStartBlockMedium = i
                                        .TextMatrix(i, 2) = rsMPMedium(1)
                                        .Row = i
                                        .col = 2
                                        .CellFontBold = True
                                        'Load Fee Component
                                        strSql = "select month_number,MSC_Paid/100 MSC_Paid,case MSC_Paid_On_Flag when 1 then 1 when 2 then 0 when 3 then 0 when 4 then 1 when 0 then 0 end MSC_Paid_On_Flag,"
                                        strSql = strSql & " MSC_Bonus/100 MSC_Bonus,case MSC_Bonus_On_Flag when 1 then 1 when 2 then 0 when 3 then 0 when 4 then 1 when 0 then 0 end MSC_Bonus_On_Flag,"
                                        strSql = strSql & " Club_Agency/100 Club_Agency,case Club_Agency_On_Flag when 1 then 1 when 2 then 0 when 3 then 0 when 4 then 1 when 0 then 0 end  Club_Agency_On_Flag"
                                        strSql = strSql & " from mp_monthly_activity"
                                        strSql = strSql & " where mp_medium_id = '" & rsMPMedium(0) & "' order by month_number"
                                        ReDim arrFee(12, 6)
                                        rsFee.Open strSql, ConnERP, 1, 3
                                        While Not rsFee.EOF
                                            For field_no = 1 To rsFee.Fields.Count - 1
                                                arrFee(rsFee(0), field_no) = rsFee(field_no)
                                            Next
                                            rsFee.MoveNext
                                        Wend
                                        rsFee.Close
                                        'load medium_detail (mp_medium_detail_view join mp_plan_dimension_view)
                                            strSql = "select b.mp_plan_dim_id,b.channel,b.version,a.media,b.channel_dimension,b.nett_rate "
                                            strSql = strSql & "from mp_medium_detail_view a inner join mp_plan_dimension_view b "
                                            strSql = strSql & "on a.mp_medium_detail_id = b.mp_medium_detail_id "
                                            strSql = strSql & "where a.mp_medium_id = '" & rsMPMedium(0) & "' "
                                            'Add by YY 08/28/2008
                                            strSql = strSql & " ORDER BY b.version,a.media"
                                            
                                            rsMPMediumDetail.Open strSql, ConnERP, 1, 3
                                            strSpotType = ""
                                            strVersion = ""
                                            While Not rsMPMediumDetail.EOF
                                                i = i + 1
                                                If i > .Rows - 2 Then
                                                    .Rows = .Rows + 1
                                                End If
                                                .MergeRow(i) = False
                                                'channel / spot type
                                                    If Trim(rsMPMediumDetail(1)) <> strSpotType Then
                                                        .TextMatrix(i, 2) = rsMPMediumDetail(1) & String(i, " ")
                                                        strSpotType = Trim(rsMPMediumDetail(1))
                                                    End If
                                                'version
                                                    If (Trim(rsMPMediumDetail(2)) <> strVersion) Or Trim(.TextMatrix(i, 2) <> "") Then
                                                        .TextMatrix(i, 3) = rsMPMediumDetail(2) & String(i, " ")
                                                        strVersion = Trim(rsMPMediumDetail(2))
                                                    End If
                                                
                                                insrow = 1 'jumlah row untuk masing2 Medium adalah 1, kecuali TV
                                                If rsMPMedium(1) = "TV" Then
                                                    insrow = 3 'Freq,Tarps,Reach
                                                End If
                                                For counter = 1 To insrow
                                                    .Rows = .Rows + counter - 1
                                                    'Media
                                                        .TextMatrix(i + counter - 1, 4) = " " & rsMPMediumDetail(3) & String(i, " ")
                                                    'Channel Dimension
                                                        .TextMatrix(i + counter - 1, 5) = " " & rsMPMediumDetail(4) & String(i, " ")
                                                    If insrow = 3 Then
                                                        Select Case counter
                                                            Case 1
                                                                'menandai current row berisi data tv frequency
                                                                    .TextMatrix(i + counter - 1, .cols - 1) = rsMPMediumDetail(0) & "FREQ" 'mp_plan_dim_id + FREQ
                                                            Case 2
                                                                'menandai current row berisi data tv Reach
                                                                    .TextMatrix(i + counter - 1, .cols - 1) = rsMPMediumDetail(0) & "REACH" 'mp_plan_dim_id + REACH
                                                            Case 3
                                                                .TextMatrix(i + counter - 1, .cols - 1) = rsMPMediumDetail(0)  'mp_plan_dim_id (TARPS)
                                                        End Select
                                                    Else
                                                        If rsMPMedium(1) = "Print" Then 'jika print beri tanda, memudah kan untuk locking krn print beda treatmentnya
                                                            .TextMatrix(i, .cols - 1) = rsMPMediumDetail(0) & "PR"  'mp_plan_dim_id
                                                        Else
                                                            .TextMatrix(i, .cols - 1) = rsMPMediumDetail(0)  'mp_plan_dim_id
                                                        End If
                                                    End If
                                                Next
                                                
                                                'load insertion
                                                    If insrow <> 3 Then
                                                        'Insertion non TV
                                                            If txtViewMonth.Text = "ALL" Then
                                                                rsMPInsertion.Open "select mp_plan_dim_id,week_year,spot,nett_rate,gross_rate,[month],isnull(other_cost,0) other_cost from mp_insertion where mp_plan_dim_id = '" & Left(Trim(msf_MPInsertion.TextMatrix(i, .cols - 1)), 19) & "'", ConnERP, 1, 3
                                                            Else
                                                                rsMPInsertion.Open "select mp_plan_dim_id,week_year,spot,nett_rate,gross_rate,[month],isnull(other_cost,0) other_cost from mp_insertion where mp_plan_dim_id = '" & Left(Trim(msf_MPInsertion.TextMatrix(i, .cols - 1)), 19) & "' and [month] in " & strViewMonth, ConnERP, 1, 3
                                                            End If
                                                            
                                                            intGrpIns = 0
                                                            intTotalNett = 0
                                                            While Not rsMPInsertion.EOF
                                                                .TextMatrix(i, rsMPInsertion(1) + 5) = rsMPInsertion(2)
                                                                intGrpIns = intGrpIns + rsMPInsertion(2)
                                                                'TOTAL KE KANAN per station
                                                                If arrFee(rsMPInsertion("month").Value, 2) = 1 Then
                                                                    feePaid = arrFee(rsMPInsertion("month").Value, 1) * rsMPInsertion("nett_rate").Value
                                                                Else
                                                                    feePaid = arrFee(rsMPInsertion("month").Value, 1) * rsMPInsertion("gross_rate").Value
                                                                End If
                                                                
                                                                If arrFee(rsMPInsertion("month").Value, 4) = 1 Then
                                                                    FeeBonus = arrFee(rsMPInsertion("month").Value, 3) * rsMPInsertion("nett_rate").Value
                                                                Else
                                                                    FeeBonus = arrFee(rsMPInsertion("month").Value, 3) * rsMPInsertion("gross_rate").Value
                                                                End If
                                                                
                                                                If arrFee(rsMPInsertion("month").Value, 6) = 1 Then
                                                                    FeeClub = arrFee(rsMPInsertion("month").Value, 5) * rsMPInsertion("nett_rate").Value
                                                                Else
                                                                    FeeClub = arrFee(rsMPInsertion("month").Value, 5) * rsMPInsertion("gross_rate").Value
                                                                End If
                                                                
                                                                intTotalNett = intTotalNett + rsMPInsertion("nett_rate") + feePaid + FeeBonus + FeeClub + rsMPInsertion("other_cost")
                                                                'End TOTAL KE KANAN per station
                                                                
                                                                rsMPInsertion.MoveNext
                                                            Wend
                                                            rsMPInsertion.Close
                                                    Else
                                                        'Insertion TV (Tarps,Freq,Reach)
                                                            If txtViewMonth.Text = "ALL" Then
                                                                rsMPInsertion.Open "select mp_plan_dim_id,week_year,spot,nett_rate,gross_rate,[month],isnull(other_cost,0) other_cost from mp_insertion where mp_plan_dim_id = '" & Trim(msf_MPInsertion.TextMatrix(i + 2, .cols - 1)) & "'", ConnERP, 1, 3
                                                            Else
                                                                rsMPInsertion.Open "select mp_plan_dim_id,week_year,spot,nett_rate,gross_rate,[month],isnull(other_cost,0) other_cost from mp_insertion where mp_plan_dim_id = '" & Trim(msf_MPInsertion.TextMatrix(i + 2, .cols - 1)) & "' and [month] in " & strViewMonth, ConnERP, 1, 3
                                                            End If
                                                            intGrpIns = 0
                                                            intTotalNett = 0
                                                            While Not rsMPInsertion.EOF
                                                                
                                                                .TextMatrix(i + 2, rsMPInsertion(1) + 5) = rsMPInsertion(2) 'Tarps
                                                                intGrpIns = intGrpIns + rsMPInsertion(2)
                                                                
                                                                'TOTAL KE KANAN per station
                                                                If arrFee(rsMPInsertion("month").Value, 2) = 1 Then
                                                                    feePaid = arrFee(rsMPInsertion("month").Value, 1) * rsMPInsertion("nett_rate").Value
                                                                Else
                                                                    feePaid = arrFee(rsMPInsertion("month").Value, 1) * rsMPInsertion("gross_rate").Value
                                                                End If
                                                                
                                                                If arrFee(rsMPInsertion("month").Value, 4) = 1 Then
                                                                    FeeBonus = arrFee(rsMPInsertion("month").Value, 3) * rsMPInsertion("nett_rate").Value
                                                                Else
                                                                    FeeBonus = arrFee(rsMPInsertion("month").Value, 3) * rsMPInsertion("gross_rate").Value
                                                                End If
                                                                
                                                                If arrFee(rsMPInsertion("month").Value, 6) = 1 Then
                                                                    FeeClub = arrFee(rsMPInsertion("month").Value, 5) * rsMPInsertion("nett_rate").Value
                                                                Else
                                                                    FeeClub = arrFee(rsMPInsertion("month").Value, 5) * rsMPInsertion("gross_rate").Value
                                                                End If
                                                                
                                                                intTotalNett = intTotalNett + rsMPInsertion("nett_rate") + feePaid + FeeBonus + FeeClub + rsMPInsertion("other_cost")
                                                                'End TOTAL KE KANAN per station
                                                                
                                                                rsMPInsertion.MoveNext
                                                            Wend
                                                            rsMPInsertion.Close
                                                            
                                                            .MergeRow(i) = True
                                                            .MergeRow(i + 1) = True
                                                            If txtViewMonth.Text = "ALL" Then
                                                                rsMPTVReachFrequency.Open "select isnull(week_year_start,0),isnull(week_year_end,0),frequency_name,reach from mp_tv_reach_frequency where mp_plan_dim_id = '" & rsMPMediumDetail(0) & "'", ConnERP, 1, 3
                                                            Else
                                                                rsMPTVReachFrequency.Open "select isnull(week_year_start,0),isnull(week_year_end,0),frequency_name,reach from mp_tv_reach_frequency where mp_plan_dim_id = '" & rsMPMediumDetail(0) & "' and (month_start  in " & strViewMonth & " or month_end in " & strViewMonth & ")", ConnERP, 1, 3
                                                            End If
                                                            intPeriode = 1
                                                            While Not rsMPTVReachFrequency.EOF
                                                                If rsMPTVReachFrequency(0) <> 0 And rsMPTVReachFrequency(1) <> 0 Then
                                                                    For counter = rsMPTVReachFrequency(0) To rsMPTVReachFrequency(1)
                                                                        .col = counter + 5
                                                                        .Row = i
                                                                        .CellBackColor = LegendTVRF.FillColor
                                                                        .Text = String(intPeriode, " ") & Trim(rsMPTVReachFrequency(2)) & String(intPeriode, " ")
                                                                        .CellAlignment = 4
                                                                        .Row = i + 1
                                                                        .CellBackColor = LegendTVRF.FillColor '16744703 'merah muda
                                                                        .Text = String(intPeriode, " ") & Trim(rsMPTVReachFrequency(3)) & "%" & String(intPeriode, " ")
                                                                        .CellAlignment = 4
                                                                    Next
                                                                    intPeriode = intPeriode + 1
                                                                End If
                                                                rsMPTVReachFrequency.MoveNext
                                                            Wend
                                                            rsMPTVReachFrequency.Close
                                                    End If
                                                    
                                                    If rsMPMedium(1) = "Other" Then
                                                        If txtViewMonth.Text = "ALL" Then
                                                            rsMPOtherMonthlyBudget.Open "select isnull(nett_budget,0) nett_rate,isnull(gross_budget,0) gross_rate,month_number [month] from mp_other_monthly_budget where mp_plan_dim_id='" & rsMPMediumDetail(0) & "'", ConnERP, 1, 3
                                                        Else
                                                            rsMPOtherMonthlyBudget.Open "select isnull(nett_budget,0) nett_rate,isnull(gross_budget,0) gross_rate,month_number [month] from mp_other_monthly_budget where mp_plan_dim_id='" & rsMPMediumDetail(0) & "' and month_number in " & strViewMonth, ConnERP, 1, 3
                                                        End If
                                                        intTotalNett = 0
                                                        While Not rsMPOtherMonthlyBudget.EOF
                                                            'TOTAL KE KANAN per station
                                                            If arrFee(rsMPOtherMonthlyBudget("month").Value, 2) = 1 Then
                                                                feePaid = arrFee(rsMPOtherMonthlyBudget("month").Value, 1) * rsMPOtherMonthlyBudget("nett_rate").Value
                                                            Else
                                                                feePaid = arrFee(rsMPOtherMonthlyBudget("month").Value, 1) * rsMPOtherMonthlyBudget("gross_rate").Value
                                                            End If
                                                            
                                                            If arrFee(rsMPOtherMonthlyBudget("month").Value, 4) = 1 Then
                                                                FeeBonus = arrFee(rsMPOtherMonthlyBudget("month").Value, 3) * rsMPOtherMonthlyBudget("nett_rate").Value
                                                            Else
                                                                FeeBonus = arrFee(rsMPOtherMonthlyBudget("month").Value, 3) * rsMPOtherMonthlyBudget("gross_rate").Value
                                                            End If
                                                            
                                                            If arrFee(rsMPOtherMonthlyBudget("month").Value, 6) = 1 Then
                                                                FeeClub = arrFee(rsMPOtherMonthlyBudget("month").Value, 5) * rsMPOtherMonthlyBudget("nett_rate").Value
                                                            Else
                                                                FeeClub = arrFee(rsMPOtherMonthlyBudget("month").Value, 5) * rsMPOtherMonthlyBudget("gross_rate").Value
                                                            End If
                                                            
                                                            intTotalNett = intTotalNett + rsMPOtherMonthlyBudget("nett_rate") + feePaid + FeeBonus + FeeClub
                                                            'End TOTAL KE KANAN per station
                                                            rsMPOtherMonthlyBudget.MoveNext
                                                        Wend
                                                        rsMPOtherMonthlyBudget.Close
                                                    End If
                                                    i = i + insrow - 1
                                                'end of load insertion
                                                'GRP/ins
                                                    If intGrpIns = 0 Then
                                                        .TextMatrix(i, .cols - 3) = ""
                                                    Else
                                                        .TextMatrix(i, .cols - 3) = intGrpIns
                                                    End If
                                                'total
                                                    
                                                    .TextMatrix(i, .cols - 2) = FormatNumber(intTotalNett, 2)
                                                rsMPMediumDetail.MoveNext
                                            Wend
                                            rsMPMediumDetail.Close
                                        'end of load mp_medium_detail join mp_plan_dimension
                                        
                                        'print Total Sub Task
                                            i = i + 2
                                            If i > .Rows - 3 Then
                                                .Rows = .Rows + 3
                                            End If
                                            If .TextMatrix(i, 0) <> "" Then
                                                .Rows = .Rows + 3
                                                i = .Rows - 2
                                            End If
                                            .MergeRow(i) = True
                                            .TextMatrix(i, 0) = "Sub Total " & rsMPMedium(1)
                                            .TextMatrix(i, 1) = "Sub Total " & rsMPMedium(1)
                                            
                                            .MergeRow(i + 1) = True
                                            .TextMatrix(i + 1, 0) = "Sub Total " & rsMPMedium(1) & " (Actual)"
                                            .TextMatrix(i + 1, 1) = "Sub Total " & rsMPMedium(1) & " (Actual)"
                                            
                                            'print monthly budget
                                                'rsBudget.Open "select mp_medium_id,month_number,month_name,quarter,min_budget + msc_paid_value + msc_bonus_value + club_agency_value plan_budget,budget,approval,total_actual from mp_monthly_activity where mp_medium_id='" & rsMPMedium(0) & "' order by month_number", ConnERP, 1, 3
                                                rsBudget.Open "select mp_medium_id,month_number,month_name,quarter,case isnull(total_actual,-1) when -1 then min_budget + msc_paid_value + msc_bonus_value + club_agency_value + isnull(other_cost,0) else total_actual end plan_budget,budget,approval,total_actual from mp_monthly_activity where mp_medium_id='" & rsMPMedium(0) & "' order by month_number", ConnERP, 1, 3
                                                intTotalBudgetActual = 0
                                                intTotalBudgetPlan = 0
                                                For counter = 6 To .cols - 4
                                                    If .TextMatrix(1, counter) <> rsBudget("month_name") Then
                                                        Remaining_Budget = Remaining_Budget - rsBudget("plan_budget").Value
                                                        txt_remaining_budget.Text = FormatNumber(Remaining_Budget, 2)
                                                        intTotalBudgetPlan = intTotalBudgetPlan + rsBudget("plan_budget").Value
                                                        intTotalBudgetActual = intTotalBudgetActual + IIf(IsNull(rsBudget("total_Actual").Value), 0, rsBudget("total_Actual").Value)
                                                        rsBudget.MoveNext
                                                    End If
                                                    If InStr(1, txtViewMonth.Text, .TextMatrix(1, counter)) <> 0 Or txtViewMonth.Text = "ALL" Then
                                                        'Plan Budget
                                                        .TextMatrix(i, counter) = String(rsBudget("month_number").Value, " ") & FormatNumber(rsBudget("plan_budget").Value, 2) & String(rsBudget("month_number").Value, " ")
                                                        'Actual Budget
                                                        If Not IsNull(rsBudget("total_Actual").Value) Then
                                                            .TextMatrix(i + 1, counter) = String(rsBudget("month_number").Value, " ") & FormatNumber(rsBudget("total_Actual").Value, 2) & String(rsBudget("month_number").Value, " ")
                                                        Else
                                                            .TextMatrix(i + 1, counter) = String(rsBudget("month_number").Value * 2, " ")
                                                        End If
                                                        
                                                        If Not IsNull(rsBudget("total_actual").Value) Then
                                                            'block yang udah actual
                                                            .col = counter
                                                            For intRowBlockMedium = intStartBlockMedium To i
                                                                .Row = intRowBlockMedium
                                                                .CellBackColor = LegendActual.FillColor
                                                            Next
                                                        Else
                                                            'block yang udah approve
                                                            Select Case rsBudget("Approval").Value
                                                            Case 1
                                                                'block direct approval
                                                                .col = counter
                                                                For intRowBlockMedium = intStartBlockMedium To i
                                                                    .Row = intRowBlockMedium
                                                                    .CellBackColor = LegendApproval.FillColor
                                                                Next
                                                            Case 2
                                                                'block web based approval
                                                                .col = counter
                                                                For intRowBlockMedium = intStartBlockMedium To i
                                                                    .Row = intRowBlockMedium
                                                                    .CellBackColor = LegendWebApproval.FillColor
                                                                Next
                                                            End Select
                                                        End If
                                                    End If
                                                Next
                                                'Total per tahun
                                                intTotalBudgetPlan = intTotalBudgetPlan + rsBudget("plan_budget").Value
                                                intTotalBudgetActual = intTotalBudgetActual + IIf(IsNull(rsBudget("total_Actual").Value), 0, rsBudget("total_Actual").Value)
                                                .TextMatrix(i, .cols - 2) = FormatNumber(intTotalBudgetPlan, 2)
                                                .TextMatrix(i + 1, .cols - 2) = FormatNumber(intTotalBudgetActual, 2)
                                                
                                                Remaining_Budget = Remaining_Budget - rsBudget("plan_budget")
                                                txt_remaining_budget.Text = FormatNumber(Remaining_Budget, 2)
                                                If Remaining_Budget < 0 Then
                                                    txt_remaining_budget.ForeColor = vbRed
                                                Else
                                                    txt_remaining_budget.ForeColor = vbBlack
                                                End If
                                                rsBudget.Close
                                                
                                                'tandai medium id
                                                .TextMatrix(i, .cols - 1) = rsMPMedium(0)
                                            'end of print monthly budget
                                            
                                            'BOLD Tulisan Sub Total Plan
                                            .Row = i
                                            .col = 0
                                            .CellFontBold = True
                                            .col = 1
                                            .CellFontBold = True
                                            'BOLD Tulisan Sub Total Actual
                                            .Row = i + 1
                                            .col = 0
                                            .CellFontBold = True
                                            .col = 1
                                            .CellFontBold = True
                                            
                                        'end of print total sub task
                                        i = i + 3 'next medium
                                        rsMPMedium.MoveNext
                                    Wend
                                    rsMPMedium.Close
                                    
                                 'end of load mp_medium
                                rsMPActivity.MoveNext
                            Wend
                            rsMPActivity.Close
                        
                        'end of load mp_activity
                        rsMPTask.MoveNext
                    Wend
                    rsMPTask.Close
                
                'end of load mp_task
                If jumlahActivity > 0 Then
                    .Rows = .Rows + 1
                End If
            End With
        End If
        rsMPMaster.Close
        
    'end of load mp_master
    Set rsMPMaster = Nothing
    Set rsMPTask = Nothing
    Set rsMPActivity = Nothing
    Set rsMPMedium = Nothing
    Set rsMPMediumDetail = Nothing
    Set rsBudget = Nothing
    Set rsMPInsertion = Nothing
    If txtViewMonth.Text <> "ALL" Then
        For i = 1 To 12
            If cbkViewMonth(i).Value = 1 Then
                Call HighLightMonth(i)
            End If
        Next
    End If
    
    Call HighLightUnlockedCell
    
    'FGMPInsertion.SetFocus
    
    Me.MousePointer = vbNormal
    'cmdSummary.SetFocus
End Sub

Private Sub HighLightUnlockedCell()
    Dim intRow As Integer
    Dim strSql As String
    Dim str_MP_Plan_Dim_Id As String
    With msf_MPInsertion
        For intRow = 0 To .Rows - 1
            If Len(Trim(.TextMatrix(intRow, .cols - 1))) >= 19 Then
                str_MP_Plan_Dim_Id = Left(.TextMatrix(intRow, .cols - 1), 19)
                strSql = "select week_year from mp_unlocked_week where mp_plan_dim_id = '" & str_MP_Plan_Dim_Id & "'"
                rsTemp.Open strSql, ConnERP, 1, 3
                While Not rsTemp.EOF
                    .Row = intRow
                    .col = CInt(rsTemp(0)) + 5
                    .CellBackColor = Shape_unlocked.FillColor
                    rsTemp.MoveNext
                Wend
                rsTemp.Close
            End If
        Next
    End With
End Sub

Private Sub HighLightMonth(intMonth As Double)
    Dim Kolom As Integer, baris As Integer
    
    For Kolom = 6 To msf_MPInsertion.cols - 4
        If EngMonthIndex(msf_MPInsertion.TextMatrix(1, Kolom)) = intMonth Then
            msf_MPInsertion.col = Kolom
            For baris = 1 To 3
                msf_MPInsertion.Row = baris
                msf_MPInsertion.CellBackColor = vbYellow '16744576 biru muda
            Next
            'For baris = 5 To FGMPInsertion.Rows - 1
            '    FGMPInsertion.Row = baris
            '    If FGMPInsertion.Text = "" Then
            '        FGMPInsertion.CellBackColor = 16744576 'biru muda
            '    End If
            'Next
        End If
    Next
End Sub


Private Sub LoadWeekCommencing(tahun As String)
'*****************************************************************************
' Nama Submodul         :  LoadWeekCommencing
' Fungsi Submodul       :  Load grid WC untuk insertion
' Nama Programmer       :  Sistyo
' Tgl Pembuatan         :  18 Agustus 2004
' Last Update           :  18 Agustus 2004/Sistyo
'******************************************************************************
    Dim counter As Integer
    With msf_MPInsertion
        .MergeRow(0) = True
        .MergeRow(1) = True
        .MergeRow(2) = True
        .MergeRow(3) = True
        .cols = 6
        rsTemp.Open "select * from week_commencing where [year] = '" & tahun & "' order by cast([month] as int) ", ConnERP, 1, 3
        While Not rsTemp.EOF
            For counter = 1 To rsTemp(7) 'Week Count
                .cols = .cols + 1
                .ColWidth(.cols - 1) = 500
                .ColAlignment(.cols - 1) = 4
                .TextMatrix(0, .cols - 1) = rsTemp(0) 'Year
                .TextMatrix(1, .cols - 1) = EngMonthName(rsTemp(1)) 'Month
                .TextMatrix(2, .cols - 1) = Day(rsTemp(counter + 1))  'date
                .TextMatrix(3, .cols - 1) = .cols - 6 'Week
                .TextMatrix(4, .cols - 1) = rsTemp(counter + 1)
                If txtViewMonth <> "ALL" Then
                    If InStr(1, txtViewMonth, EngMonthName(rsTemp(1))) = 0 Then
                        .ColWidth(.cols - 1) = 0
                    End If
                End If
            Next
            rsTemp.MoveNext
        Wend
        rsTemp.Close
        .cols = .cols + 3 'grp/ins, Total, mp_plan_dim_ID
        .TextMatrix(1, .cols - 3) = "GRP/ins"
        .ColAlignment(.cols - 3) = 4
        .ColWidth(.cols - 3) = 1000
        .TextMatrix(1, .cols - 2) = "Total"
        .ColAlignment(.cols - 2) = 4
        .ColWidth(.cols - 2) = 2000
        .ColWidth(.cols - 1) = 0   'mp_plan_dim_id
    End With
End Sub

Private Sub msf_MPInsertion_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
'=============================================================================
'Nama Sub Modul : FGMPInsertion_MouseDown
'Fungsi         : Init menu klik kanan
'Programer      : Sistyo
'=============================================================================
    intMouseCol = msf_MPInsertion.MouseCol
    If intMouseCol > 1 Then
        If Button = vbRightButton Then
            If msf_MPInsertion.MouseCol > msf_MPInsertion.FixedCols - 1 Then
                mdi_Main.MnuFreeze.Enabled = True
                mdi_Main.MnuUnFreeze.Enabled = False
            Else
                mdi_Main.MnuFreeze.Enabled = False
                mdi_Main.MnuUnFreeze.Enabled = True
            End If
            mdi_Main.MnuEndPeriode.Enabled = False
            mdi_Main.MnuStartPeriode.Enabled = False
            mdi_Main.MnuClearPeriode.Enabled = False
            mdi_Main.mnu_unlock.Enabled = False
            
            mdi_Main.mnuRateInfo.Enabled = False
            
            mdi_Main.mnu_refresh_rate.Enabled = False
            'mdi_main.mnu_refresh_rate.Enabled = True
            
            mdi_Main.mnu_unapprove.Enabled = False
            mdi_Main.mnu_Approve.Enabled = False
            
            mdi_Main.mnu_Cancel_IB.Enabled = False
            mdi_Main.mnu_Show_Related_IB.Enabled = False
            
            mdi_Main.mnu_set_objective.Enabled = False
            mdi_Main.mnu_view_objective.Enabled = False
            mdi_Main.mnu_view_id.Enabled = False
            
            mdi_Main.mnu_update_actual_budget.Enabled = False
            
                With msf_MPInsertion
                    If .MouseCol > 5 And .MouseCol < .cols - 3 Then
                        .col = .MouseCol
                        .Row = .MouseRow
                        
                        If .Text <> "" And Right(.TextMatrix(.Row, .cols - 1), 4) <> "FREQ" And Right(.TextMatrix(.Row, .cols - 1), 5) <> "REACH" And Len(.TextMatrix(.Row, .cols - 1)) > 18 And Mid(.TextMatrix(.Row, .cols - 1), 6, 4) <> "MDUM" Then
                            mdi_Main.mnuRateInfo.Enabled = True
                            mdi_Main.mnu_view_objective.Enabled = True
                            If isAllowEditCell(.MouseRow, .MouseCol) Then
                                mdi_Main.mnu_set_objective.Enabled = True
                            Else
                                If .CellBackColor = LegendWebApproval.FillColor Then
                                    mdi_Main.mnu_set_objective.Enabled = True
                                End If
                            End If
                        End If
                        
                        If Right(.TextMatrix(.Row, .cols - 1), 4) = "FREQ" Or Right(.TextMatrix(.Row, .cols - 1), 5) = "REACH" Then
                            If .Text <> "" Then
                                mdi_Main.mnu_view_id.Enabled = True
                            End If
                        End If
                        
                        If Mid(.TextMatrix(.Row, .cols - 1), 6, 4) = "MDUM" Then
                            mdi_Main.mnu_Cancel_IB.Enabled = True
                            mdi_Main.mnu_Show_Related_IB.Enabled = True
                            
                            If .CellBackColor = LegendWebApproval.FillColor Or .CellBackColor = LegendApproval.FillColor Then
                                recDate.Requery
                                If .TextMatrix(0, .col) * 100 + EngMonthIndex(.TextMatrix(1, .col)) >= Year(recDate(0)) * 100 + month(recDate(0)) Then
                                    'TIDAK BOLEH UNAPPROVE JIKA SUDAH DI PUBLISH
                                    If picButton(biePublishToWeb).Enabled Then
                                        mdi_Main.mnu_unapprove.Enabled = True
                                    End If
                                End If
                            End If
                            
                            If .CellBackColor <> LegendApproval.FillColor Then
                                mdi_Main.mnu_Approve.Enabled = True
                            End If
                            
                        End If
                        
                        If Mid(.TextMatrix(.Row - 1, .cols - 1), 6, 4) = "MDUM" Then
                            mdi_Main.mnu_update_actual_budget.Enabled = True
                        End If
                        
                        If isAllowEditCell(.MouseRow, .MouseCol) Then
                            If Mid(.TextMatrix(.Row, .cols - 1), 6, 4) = "MDUM" Then
                                mdi_Main.mnu_refresh_rate.Enabled = picButton(biePublishToWeb).Enabled
                            End If
                            
                            If Right(.TextMatrix(.Row, .cols - 1), 5) = "REACH" Or Right(.TextMatrix(.Row, .cols - 1), 4) = "FREQ" Then
                                'EDIT / DELETE PERIODE HANYA BOLEH JIKA BLM DI PUBLISH
                                If picButton(biePublishToWeb).Enabled Then
                                    If FlagMarkStart = True Then
                                        If .col >= MarkStartCol And .Text = "" Then
                                            If .Row = MarkStartRow Or .Row = MarkStartRow + 1 Then
                                                'periode hanya boleh max. 3 bulan
                                                If EngMonthIndex(.TextMatrix(1, .col)) - EngMonthIndex(.TextMatrix(1, MarkStartCol)) < 3 Then
                                                    mdi_Main.MnuEndPeriode.Enabled = True
                                                End If
                                            End If
                                        End If
                                    Else
                                        If .Text = "" Then
                                            mdi_Main.MnuStartPeriode.Enabled = True
                                        End If
                                    End If
                                    If .Text <> "" Then
                                        mdi_Main.MnuClearPeriode.Enabled = True
                                    End If
                                End If
                                '-----
                            End If
                        End If
                    End If
                End With
            PopupMenu mdi_Main.MnuFGMPInsertion
        End If
    End If
    
End Sub

Private Sub msf_MPInsertion_DblClick()
   With msf_MPInsertion
        If .col > 5 And .col < .cols - 3 Then
            If isAllowEditCell(.MouseRow, .MouseCol) Then
                If .TextMatrix(.Row, .cols - 1) <> "" Then
                    intCol = .col
                    intRow = .Row
                    'Entry Insertion hanya boleh jika belum dipunlish ke web
                    MsgBox "If cmd_Publish.Enabled Then", vbExclamation, strApplication_Name
                    If True Then
                        '---
                        If Right(.TextMatrix(.Row, .cols - 1), 5) = "REACH" Then
                            'show reach textbox
                                If .Text <> "" Then
                                    Call EntryReach("Edit")
                                End If
                        Else
                            If Right(.TextMatrix(.Row, .cols - 1), 4) = "FREQ" Then
                                'show cboFreq
                                Call msf_MPInsertion_Click
                            Else
                                If UCase(Mid(.TextMatrix(.Row, 0), 1, 9)) <> "SUB TOTAL" Then
                                    Call StartTyping("Edit")
                                End If
                            End If
                        End If
                        '---
                    End If
                End If
            End If
        End If
    End With
End Sub

Private Sub EntryReach(mode As String, Optional KeyAscii As Integer)
    Dim strMP_Plan_Dim_Id As String, weekStart As Integer, weekEnd As Integer, i As Integer
    Dim Week As Integer, intTbWidth As Integer, flagExist As Boolean
    
    flagExist = False
    intTbWidth = 0
    With msf_MPInsertion
        strMP_Plan_Dim_Id = Mid(.TextMatrix(.Row, .cols - 1), 1, 19)
        Week = .TextMatrix(3, .col)
        rsTemp.Open "select week_year_start,week_year_end from mp_tv_reach_frequency where week_year_start<=" & Week & " and week_year_end>=" & Week & " and mp_plan_dim_id='" & strMP_Plan_Dim_Id & "'"
        If Not rsTemp.EOF Then
            flagExist = True
            weekStart = rsTemp(0)
            weekEnd = rsTemp(1)
        End If
        rsTemp.Close
        If flagExist Then
            For i = weekStart To weekEnd
                intTbWidth = intTbWidth + .ColWidth(5 + i) + .GridLineWidth
            Next
            .col = weekStart + 5
            txtTVREACH.Top = .Top + .CellTop
            txtTVREACH.Left = .Left + .CellLeft
            txtTVREACH.Width = intTbWidth
            txtTVREACH.Height = .CellHeight
            If mode = "Edit" Then
                txtTVREACH.Text = Val(.Text)
            Else
                txtTVREACH.Text = Chr(KeyAscii)
            End If
            txtTVREACH.SelStart = Len(txtTVREACH.Text)
            txtTVREACH.Visible = True
            txtTVREACH.SetFocus
        End If
    End With
End Sub

Private Sub msf_MPInsertion_KeyDown(KeyCode As Integer, Shift As Integer)
    With msf_MPInsertion
        If .col > 5 And .col < .cols - 3 Then
            If isAllowEditCell(.Row, .col) And picButton(biePublishToWeb).Enabled = True Then
                If .TextMatrix(.Row, .cols - 1) <> "" Then
                    intCol = .col
                    intRow = .Row
                    If KeyCode = 113 Then  'F2
                        If Right(.TextMatrix(.Row, .cols - 1), 5) = "REACH" Then
                            'show reach textbox
                                Call EntryReach("Edit")
                        Else
                            If Right(.TextMatrix(.Row, .cols - 1), 4) = "FREQ" Then
                                'show cbo Freq / do nothing
                            Else
                                If UCase(Mid(.TextMatrix(.Row, 0), 1, 9)) <> "SUB TOTAL" Then
                                    Call StartTyping("Edit")
                                End If
                            End If
                        End If
                    End If
                    If KeyCode = 46 Then 'Delete
                        If Right(.TextMatrix(.Row, .cols - 1), 5) = "REACH" Then
                            'show reach textbox
                                If .Text <> "" Then
                                    Call mdi_Main.MnuClearPeriode_Click
                                End If
                        Else
                            If Right(.TextMatrix(.Row, .cols - 1), 4) = "FREQ" Then
                                'Delete Periode
                                    If .Text <> "" Then
                                        Call mdi_Main.MnuClearPeriode_Click
                                    End If
                            Else
                                If UCase(Mid(.TextMatrix(.Row, 0), 1, 9)) <> "SUB TOTAL" Then
                                    If .Text <> "" Then
                                        Call DeleteInsertion
                                    End If
                                End If
                            End If
                        End If
                    End If
                End If
            End If
        End If
    End With
End Sub


Private Sub msf_MPInsertion_KeyPress(KeyAscii As Integer)
 With msf_MPInsertion
        If .col > 5 And .col < .cols - 3 Then
            If isAllowEditCell(.Row, .col) And picButton(enButtonType.biePublishToWeb) = True Then
                If .TextMatrix(.Row, .cols - 1) <> "" Then
                    If KeyAscii > 47 And KeyAscii < 58 Then
                        intCol = .col
                        intRow = .Row
                        If Right(.TextMatrix(.Row, .cols - 1), 5) = "REACH" Then
                            'show reach textbox
                                Call EntryReach("Replace", KeyAscii)
                        Else
                            If Right(.TextMatrix(.Row, .cols - 1), 4) = "FREQ" Then
                                'do nothing
                            Else
                                If UCase(Mid(.TextMatrix(.Row, 0), 1, 9)) <> "SUB TOTAL" Then
                                    Call StartTyping("Replace", KeyAscii)
                                End If
                            End If
                        End If
                    End If
                End If
            End If
        End If
    End With
End Sub

Private Sub lbl_show_user_server_Click()
    MsgBox strLogin_User & "@" & strServerName, vbExclamation, strApplication_Name
End Sub

Private Sub lblReleaseDate_DblClick()
  
    Frm_MPSetReleaseDate.show 1
 
End Sub

Private Sub LstFrequency_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 1 Then
        Dim Id As Double, WeekYearStart As Integer, WeekYearEnd As Integer, i As Integer
    
        Dim flagExist As Boolean
        flagExist = False
        rsTemp.Open "select mp_tv_rf_id,week_year_start,week_year_end from mp_tv_reach_frequency where week_year_start <=" & msf_MPInsertion.TextMatrix(3, intCol) & " and week_year_end >=" & msf_MPInsertion.TextMatrix(3, intCol) & " and mp_plan_dim_id = '" & Mid(msf_MPInsertion.TextMatrix(intRow, msf_MPInsertion.cols - 1), 1, 19) & "'", ConnERP, 1, 3
        If Not rsTemp.EOF Then
            flagExist = True
            Id = rsTemp(0)
            WeekYearStart = rsTemp(1)
            WeekYearEnd = rsTemp(2)
        End If
        rsTemp.Close
        If flagExist Then
            ConnERP.Execute "update mp_tv_reach_frequency set frequency_name='" & LstFrequency.Text & "',frequency_code =" & Val(LstFrequency.Text) & " where mp_tv_rf_id=" & Id
            For i = WeekYearStart To WeekYearEnd
                msf_MPInsertion.col = i + 5
                msf_MPInsertion.Text = LstFrequency.Text
            Next
        End If
        LstFrequency.Visible = False
    End If
End Sub

Private Sub LstFrequency_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim ListIndex As Integer
    If LstFrequency.ListCount <> 0 Then
        ListIndex = GetListBIndex(LstFrequency, Y)
        If ListIndex <> -1 Then
            LstFrequency.ListIndex = ListIndex
        End If
    End If
End Sub

Public Sub Mnu_Approve_Click()
    Me.MousePointer = vbHourglass
    'Approve
    Dim strMediumID As String, intMonth As Integer, strSql As String
    Dim rsTemp As New ADODB.Recordset, strMediumCode As String
    Dim GoThrough As Boolean, isAdaNull As Boolean, StrIBTemp As String
    Dim StrTask As String, StrActivity As String, isApproved As Boolean
    Dim isLogCreated As Boolean, intTV_Obj As Integer, intTV_Obj_Start As Integer, intTV_Obj_End As Integer
    
    isLogCreated = False
    isApproved = False
    strMediumID = msf_MPInsertion.TextMatrix(msf_MPInsertion.Row, msf_MPInsertion.cols - 1)
    intMonth = EngMonthIndex(msf_MPInsertion.TextMatrix(1, msf_MPInsertion.col))
    Frm_MPEnterPassword.show 1
    If Frm_MPEnterPassword.txtOKCancel = "OK" Then
        GoThrough = True
        'cek isLatest?
        rsTemp.Open "select is_latest from mp_master where mp_number = '" & cboMPNumber.Text & "'", ConnERP, 1, 3
        If Not rsTemp.EOF Then
            If rsTemp(0) <> 1 Then
                GoThrough = False
                MsgBox "Approval Rejected by system! This is not the latest media plan!", vbExclamation, strApplication_Name
            End If
        Else
            MsgBox "This Media Plan has been deleted!", vbExclamation, strApplication_Name
            GoThrough = False
        End If
        rsTemp.Close
        
        If GoThrough Then
            'Cek insertion
            If isInsertionNotEmpty(strMediumID, intMonth) Then
                'Create Log Files
                Create_Log_File
                
                isLogCreated = True
                
                'Get Task dan Activity untuk Log file ang get Medium code (TV/RD/OT/ dll)
                StrTask = ""
                StrActivity = ""
                strMediumCode = ""
                strSql = "SELECT a.task_desc,b.activity_type, b.activity_desc,c.medium_code "
                strSql = strSql & " FROM mp_task a inner join mp_activity b "
                strSql = strSql & " ON a.mp_task_id = b.mp_task_id "
                strSql = strSql & " inner join mp_medium c "
                strSql = strSql & " ON b.mp_activity_id = c.mp_activity_id and c.mp_medium_id = '" & strMediumID & "'"
                rsTemp.Open strSql, ConnERP, 1, 3
                    StrTask = rsTemp(0)
                    StrActivity = rsTemp(1) & "-" & rsTemp(2)
                    strMediumCode = rsTemp(3)
                rsTemp.Close
                
                Select Case strMediumCode
                Case "TV":
                    'Check Apakah Ada TV Tarps yang belum related ke Objective
                    isAdaNull = False
                    strSql = "select count(*) from mp_insertion where mp_plan_dim_id in (select mp_plan_dim_id from mp_ids where mp_medium_id = '" & strMediumID & "') "
                    strSql = strSql & " and [month] = " & intMonth & " and mp_tv_rf_id is null"
                    rsTemp.Open strSql, ConnERP, 1, 3
                        If rsTemp(0) <> 0 Then isAdaNull = True
                    rsTemp.Close
                    
                    If Not isAdaNull Then
                        ' Create IB
                        StrIBTemp = Generate_IB_TV(strMediumID, intMonth)
                        
                        If UCase(StrIBTemp) <> "ERROR" Then
                            
                            strSql = "select max(month_end) month_end from mp_tv_reach_frequency "
                            strSql = strSql & " where mp_plan_dim_id in (select mp_plan_dim_id from mp_ids where mp_medium_id = '" & strMediumID & "') "
                            strSql = strSql & " and [month_start] = " & intMonth
                            
                            intTV_Obj_Start = intMonth + 1
                            rsTemp.Open strSql, ConnERP, 1, 3
                                intTV_Obj_End = rsTemp("Month_End").Value
                            rsTemp.Close
                            
                            For intTV_Obj = intTV_Obj_Start To intTV_Obj_End
                                strSql = "update mp_monthly_activity set approved_mp_medium_id_history = isnull(approved_mp_medium_id_history,'') + '" & Right(strMediumID, 5) & "' "
                                strSql = strSql & "where mp_medium_id = '" & strMediumID & "' and month_number = " & intTV_Obj & " and isnull(approved_mp_medium_id_history,'') not like '%" & Right(strMediumID, 5) & "%'"
                                ConnERP.Execute strSql
                            Next
                            
                            Call DoApproval(strMediumID, intMonth)
                            
                            isApproved = True
                            'Write To Log
                            Write_To_Log_File StrTask & " : " & StrActivity & " : Medium: TV : Selected Month: " & EngMonthName(intMonth) & " --->  IB ID : " & StrIBTemp
                        Else
                            'Write To Log
                            Write_To_Log_File StrTask & " : " & StrActivity & " : Medium: TV : Selected Month: " & EngMonthName(intMonth) & " ---> " & " Generete IB Failed."
                        End If
                    Else
                        'Write To Log There are, Unliked Tarps To Objective
                        Write_To_Log_File StrTask & " : " & StrActivity & " : Medium: TV : Selected Month: " & EngMonthName(intMonth) & " ---> Generete IB Failed. There are Unlinked TARPS to Objective."
                    End If
                Case "PR"   'For PR
                    StrIBTemp = Generate_IB_Print(strMediumID, intMonth)
                    If UCase(StrIBTemp) = "ERROR" Then
                       'Write To Log
                        Write_To_Log_File StrTask & " : " & StrActivity & " : Medium: Print : Selected Month: " & EngMonthName(intMonth) & " ---> Generete IB Failed. There are Unlinked TARPS to Objective."
                    ElseIf UCase(StrIBTemp) = "CANCEL" Then
                        'Write To Log
                        Write_To_Log_File StrTask & " : " & StrActivity & " : Medium: Print : Selected Month: " & EngMonthName(intMonth) & " ---> Approval And Create IB Canceled"
                    Else
                        Call DoApproval(strMediumID, intMonth)
                        isApproved = True
                        'Write To Log
                        Write_To_Log_File StrTask & " : " & StrActivity & " : Medium: Print : Selected Month: " & EngMonthName(intMonth) & " --->  IB ID : " & StrIBTemp
                    End If
                        
                Case "RD"   'For RD
                    StrIBTemp = Generate_IB_Radio(strMediumID, intMonth)
                    If UCase(StrIBTemp) <> "ERROR" Then
                        Call DoApproval(strMediumID, intMonth)
                        isApproved = True
                        'Write To Log
                        Write_To_Log_File StrTask & " : " & StrActivity & " : Medium: Radio : Selected Month: " & EngMonthName(intMonth) & " --->  IB ID : " & StrIBTemp
                    Else
                        'Write To Log
                        Write_To_Log_File StrTask & " : " & StrActivity & " : Medium: Radio : Selected Month: " & EngMonthName(intMonth) & " ---> Generete IB Failed."
                    End If
                    
                Case "OT" 'For OT
                    StrIBTemp = Generate_IB_Others(strMediumID, intMonth)
                    If UCase(StrIBTemp) <> "ERROR" Then
                        Call DoApproval(strMediumID, intMonth)
                        isApproved = True
                        'Write To Log
                        Write_To_Log_File StrTask & " : " & StrActivity & " : Medium: Other : Selected Month: " & EngMonthName(intMonth) & " --->  IB ID : " & StrIBTemp
                    Else
                        'Write To Log
                        Write_To_Log_File StrTask & " : " & StrActivity & " : Medium: Other : Selected Month: " & EngMonthName(intMonth) & " ---> Generete IB Failed."
                    End If
                    
                Case "CN" 'For Cinema
                    StrIBTemp = Generate_IB_Others(strMediumID, intMonth)
                    If UCase(StrIBTemp) <> "ERROR" Then
                        Call DoApproval(strMediumID, intMonth)
                        isApproved = True
                        'Write To Log
                        Write_To_Log_File StrTask & " : " & StrActivity & " : Medium: Cinema : Selected Month: " & EngMonthName(intMonth) & " --->  IB ID : " & StrIBTemp
                    Else
                        'Write To Log
                        Write_To_Log_File StrTask & " : " & StrActivity & " : Medium: Cinema : Selected Month: " & EngMonthName(intMonth) & " ---> Generete IB Failed."
                    End If
                End Select
            Else
                MsgBox "Insertion Data is Empty! Approval Canceled!", vbExclamation, strApplication_Name
            End If
            
        End If
    End If
    If isApproved Then
        MsgBox "Approved!", vbInformation, strApplication_Name
        Call cboMPNumber_Click
    End If
    If isLogCreated Then
        '---------- Show Log File & Close Log file object--------------
        Set fso = Nothing
        Set FSOInterface = Nothing
            
        Shell "Notepad.exe C:\MPLog\" & LogFileName, vbNormalNoFocus
            
        LogFileName = ""
        '--------------------------------------------------------------
    End If
    Unload Frm_MPEnterPassword
    Me.MousePointer = vbNormal
End Sub



Public Sub Mnu_Cancel_IB_Click()
    'Cancel IB
    Dim strMediumID As String, intMonth As Integer
    
    strMediumID = msf_MPInsertion.TextMatrix(msf_MPInsertion.Row, msf_MPInsertion.cols - 1)
    intMonth = EngMonthIndex(msf_MPInsertion.TextMatrix(1, msf_MPInsertion.col))
    
    Frm_MPEnterPassword.show 1
    
    If Frm_MPEnterPassword.txtOKCancel = "OK" Then
    
        If Cancel_IB(strMediumID, intMonth, True) Then
            'Send Task To Implementer
            
            MsgBox "Cancel IB Success.", vbExclamation, strApplication_Name
        Else
            MsgBox "Cancel IB Fail.", vbExclamation, strApplication_Name
        End If
    End If
    
    Unload Frm_MPEnterPassword
    
End Sub

Public Sub Mnu_refresh_rate_Click()
    Dim strMediumID As String, strSql As String
    Dim rsInsertion As New ADODB.Recordset, strMonth As String, intMonth As Integer
    strMediumID = msf_MPInsertion.TextMatrix(msf_MPInsertion.Row, msf_MPInsertion.cols - 1)
    strMonth = msf_MPInsertion.TextMatrix(1, msf_MPInsertion.col)
    intMonth = EngMonthIndex(strMonth)
    'MsgBox strMediumId
    strSql = "select * from mp_insertion where mp_plan_dim_id in (select mp_plan_dim_id from mp_ids where mp_medium_id='" & strMediumID & "') and [month] = " & intMonth
    rsInsertion.CursorLocation = adUseClient
    rsInsertion.Open strSql, ConnERP
    While Not rsInsertion.EOF
        'Delete (remove old rate)
        strSql = "delete from mp_insertion where mp_plan_dim_id = '" & rsInsertion(0) & "' and week_year = " & rsInsertion(2) & " and [month] = " & intMonth
        ConnERP.Execute strSql
        'insert (get new rate)
        If IsNull(rsInsertion("mp_tv_rf_id")) Then
            strSql = "insert into mp_insertion(mp_plan_dim_id,[month],week_year,week_commencing,spot,nett_rate,gross_rate) values "
            strSql = strSql & "('" & rsInsertion("mp_plan_dim_id") & "', " & rsInsertion("month") & "," & rsInsertion("week_year") & ",'" & rsInsertion("week_commencing") & "'," & rsInsertion("spot") & ",0,0)"
        Else
            strSql = "insert into mp_insertion(mp_plan_dim_id,[month],week_year,week_commencing,spot,nett_rate,gross_rate,mp_tv_rf_id) values "
            strSql = strSql & "('" & rsInsertion("mp_plan_dim_id") & "', " & rsInsertion("month") & "," & rsInsertion("week_year") & ",'" & rsInsertion("week_commencing") & "'," & rsInsertion("spot") & ",0,0," & rsInsertion("mp_tv_rf_id") & ")"
        End If
        ConnERP.Execute strSql
        rsInsertion.MoveNext
    Wend
    rsInsertion.Close
    Call cboMPNumber_Click
    Me.MousePointer = vbNormal
End Sub

Public Sub Mnu_Set_Objective_Click()
    Frm_MPSetObjective.show 1
End Sub

Public Sub Mnu_Show_Related_IB_Click()
    'Show related IB
    Dim strMediumID As String, intMonth As Integer
    
    strMediumID = msf_MPInsertion.TextMatrix(msf_MPInsertion.Row, msf_MPInsertion.cols - 1)
    intMonth = EngMonthIndex(msf_MPInsertion.TextMatrix(1, msf_MPInsertion.col))
    
    MsgBox ShowIB(strMediumID, intMonth), vbExclamation, strApplication_Name
    
    
End Sub

Public Sub Mnu_unapprove_Click()
    'UnApprove pak dhe
    Dim strMediumID As String, intMonth As Integer
    strMediumID = msf_MPInsertion.TextMatrix(msf_MPInsertion.Row, msf_MPInsertion.cols - 1)
    intMonth = EngMonthIndex(msf_MPInsertion.TextMatrix(1, msf_MPInsertion.col))
    Frm_MPEnterPassword.show 1
    If Frm_MPEnterPassword.txtOKCancel = "OK" Then
        Call DoUnApprove(strMediumID, intMonth)
    End If
    Unload Frm_MPEnterPassword
End Sub

Private Sub DoUnApprove(strMediumID As String, intMonth As Integer)
    Dim strSql As String
    'UnApproved
    strSql = "update mp_monthly_activity set approval=0,"
    strSql = strSql & "Client_Approved_By=null,"
    strSql = strSql & "Client_Approved_Date=null,"
    strSql = strSql & "Client_Noted_By=null,"
    strSql = strSql & "Approved_By=null,"
    strSql = strSql & "Approved_Date=null,"
    'Ditambahkan yang lainnya
    strSql = strSql & "Media_approval=null,"
    strSql = strSql & "Media_Approved_By=null,"
    strSql = strSql & "Media_Approved_Date=null,"
    strSql = strSql & "Media_2_Approval=null,"
    strSql = strSql & "Media_2_Approved_By=null,"
    strSql = strSql & "Media_2_Approved_Date=null"
    
    'strSQL = strSQL & ",Approved_mp_medium_id = null"
    
    strSql = strSql & " WHERE mp_medium_id = '" & strMediumID & "' and month_number=" & intMonth
    ConnERP.Execute strSql
    
    '================================Update Quotation=================================================
    Dim strMPNumber As String
    Dim rsTemp As New ADODB.Recordset
    Dim strMediumCode As String
    
    strMPNumber = cboMPNumber.Text
    
    'Get Medium Code
    rsTemp.Open "select medium_code from mp_medium where mp_medium_id='" & strMediumID & "'", ConnERP, 1, 3
        strMediumCode = rsTemp(0)
    rsTemp.Close
    
    'cek is Exist then get Quotation_ID??
    Dim isExist As Boolean
    Dim strmp_Monthly_Quot_ID As String
    isExist = False
    strmp_Monthly_Quot_ID = ""
    rsTemp.Open "select mp_monthly_quot_id from mp_monthly_quotation where mp_number='" & strMPNumber & "' and medium_code='" & strMediumCode & "' and [month]=" & intMonth, ConnERP, 1, 3
    If Not rsTemp.EOF Then
        isExist = True
        strmp_Monthly_Quot_ID = rsTemp(0)
    End If
    rsTemp.Close
    
    If isExist Then
        'Get Value Pengurang
        Dim int_total_gross As Double
        Dim int_total_nett As Double
        Dim int_msc As Double
        Dim int_other_cost As Double
        Dim int_bonus_fee As Double
        Dim int_agency_charge As Double
        
        Dim int_sub_total As Double
        Dim int_grand_total As Double
        
        int_total_gross = 0
        int_total_nett = 0
        int_msc = 0
        int_other_cost = 0
        int_bonus_fee = 0
        int_agency_charge = 0
        
        strSql = "select sum(min_budget),sum(gross_budget),sum(msc_paid_value),sum(msc_bonus_value) from mp_monthly_activity "
        strSql = strSql & "where month_number = " & intMonth & " and mp_medium_id = '" & strMediumID & "' "
        
        Dim rsQuotation As New ADODB.Recordset
        rsQuotation.Open strSql, ConnERP, 1, 3
        If Not rsQuotation.EOF Then
            int_total_gross = rsQuotation(1)
            int_total_nett = rsQuotation(0)
            int_msc = rsQuotation(2)
            'int_other_cost = 0
            int_bonus_fee = rsQuotation(3)
            'int_agency_charge = 0
        End If
        rsQuotation.Close
        Set rsQuotation = Nothing
        
        int_sub_total = int_total_nett + int_msc + int_other_cost + int_bonus_fee
        int_grand_total = int_sub_total + int_agency_charge
        
        'Update/kurangi Quotation
        strSql = "update mp_monthly_quotation set "
        strSql = strSql & "total_gross=total_gross-" & int_total_gross
        strSql = strSql & ", total_nett=total_nett-" & int_total_nett
        strSql = strSql & ", MSC=MSC-" & int_msc
        strSql = strSql & ", other_cost=other_cost-" & int_other_cost
        strSql = strSql & ", bonus_fee=bonus_fee-" & int_bonus_fee
        strSql = strSql & ", sub_total=sub_total-" & int_sub_total
        strSql = strSql & ", agency_charge=agency_charge-" & int_agency_charge
        strSql = strSql & ", grand_total = grand_total-" & int_grand_total
        strSql = strSql & " where mp_monthly_quot_id='" & strmp_Monthly_Quot_ID & "'"
        
        ConnERP.Execute strSql
    End If
    '================================E/O Update Quotation==============================================
    
    MsgBox "UnApproved", vbExclamation, strApplication_Name
    
    Call cboMPNumber_Click
End Sub

Public Sub Mnu_update_actual_budget_Click()
    Frm_MPActualBudgetUpdate.show 1
End Sub

Public Sub mnu_view_id_Click()
    Dim strMPPlanDimID As String
    Dim intWeekYear As Integer
    Dim strSql As String
    Dim rsTemp As New ADODB.Recordset
    With msf_MPInsertion
        strMPPlanDimID = Left(.TextMatrix(.Row, .cols - 1), 19)
        intWeekYear = Trim(.TextMatrix(3, .col))
        strSql = "select mp_tv_rf_id from mp_tv_reach_frequency where mp_plan_dim_id = '" & strMPPlanDimID & "' and week_year_start<= " & intWeekYear & " and week_year_end >= " & intWeekYear
        rsTemp.Open strSql, ConnERP, 1, 3
        If Not rsTemp.EOF Then
            MsgBox "Objective ID = " & rsTemp(0)
        Else
            MsgBox "Objective not found or record has been deleted!"
        End If
        rsTemp.Close
    End With
End Sub

Public Sub mnu_view_objective_Click()
    Dim strMPPlanDimID As String
    Dim intWeekYear As Integer
    Dim strSql As String
    Dim rsTemp As New ADODB.Recordset
    With msf_MPInsertion
        strMPPlanDimID = Left(.TextMatrix(.Row, .cols - 1), 19)
        intWeekYear = Trim(.TextMatrix(3, .col))
        strSql = "select mp_tv_rf_id,week_commencing_start,week_commencing_end,frequency_name,reach from mp_tv_reach_frequency where mp_tv_rf_id = (select mp_tv_rf_id from mp_insertion where mp_plan_dim_id = '" & strMPPlanDimID & "' and week_year= " & intWeekYear & ")"
        rsTemp.Open strSql, ConnERP, 1, 3
        If Not rsTemp.EOF Then
            If IsNull(rsTemp(0)) Then
                MsgBox "Objective ID not set!", vbExclamation, strApplication_Name
            Else
                MsgBox "Objective ID = " & rsTemp(0) & vbCrLf & "Week Start : " & rsTemp("week_commencing_start") & vbCrLf & "Week End : " & rsTemp("week_commencing_end") & vbCrLf & "Freq. : " & rsTemp("Frequency_name") & vbCrLf & "Reach : " & rsTemp("reach") & "%", vbExclamation, strApplication_Name
            End If
        Else
            MsgBox "Objective not found or record has been deleted!", vbExclamation, strApplication_Name
        End If
        rsTemp.Close
    End With
End Sub

Public Sub MnuClearPeriode_Click()
'*****************************************************************************
' Nama Prosedur     :   MnuClearPeriode_Click
' Fungsi Prosedur   :   Delete TV Reach + Frequency
' Parameter  Input  :
' Parameter Output  :
' Tgl Pembuatan     :   01 nov 2004
' Last Update/By    :   01 nov 2004/Sistyo
'*****************************************************************************

    Dim strMP_Plan_Dim_Id As String, flagExist As Boolean
    Dim Week As Integer, weekStart As Integer, weekEnd As Integer, i As Integer
    Dim pesan
    
    pesan = MsgBox("Confirm Clear Periode?", vbQuestion + vbYesNo, strApplication_Name)
    If pesan = 6 Then
        flagExist = False
        With msf_MPInsertion
            strMP_Plan_Dim_Id = Mid(.TextMatrix(.Row, .cols - 1), 1, 19)
            Week = .TextMatrix(3, .col)
            rsTemp.Open "select week_year_start,week_year_end from mp_tv_reach_frequency where week_year_start<=" & Week & " and week_year_end>=" & Week & " and mp_plan_dim_id='" & strMP_Plan_Dim_Id & "'"
            If Not rsTemp.EOF Then
                flagExist = True
                weekStart = rsTemp(0)
                weekEnd = rsTemp(1)
            End If
            rsTemp.Close
            ConnERP.Execute "delete from mp_tv_reach_frequency where week_year_start<=" & Week & " and week_year_end>=" & Week & " and mp_plan_dim_id='" & strMP_Plan_Dim_Id & "'"
            If flagExist Then
                If Right(.TextMatrix(.Row, .cols - 1), 5) = "REACH" Then
                    .Row = .Row - 1
                End If
                For i = weekStart To weekEnd
                    .col = i + 5
                    .Text = ""
                    .Row = .Row + 2
                    If .CellBackColor = Shape_unlocked.FillColor Then
                        .Row = .Row - 1
                        .Text = ""
                        .CellBackColor = Shape_unlocked.FillColor
                        .Row = .Row - 1
                        .CellBackColor = Shape_unlocked.FillColor
                    Else
                        .Row = .Row - 1
                        .Text = ""
                        .CellBackColor = vbWhite
                        .Row = .Row - 1
                        .CellBackColor = vbWhite
                    End If
                Next
                .Refresh
            End If
        End With
    End If
    
End Sub

Public Sub MnuFreeze_Click()
    Dim i As Integer
    With msf_MPInsertion
        
        If EngMonthIndex(.TextMatrix(1, intMouseCol)) <> -1 Then
            i = intMouseCol
            While .TextMatrix(1, i) = .TextMatrix(1, intMouseCol)
                i = i + 1
            Wend
            msf_MPInsertion.FixedCols = i
        Else
            msf_MPInsertion.FixedCols = intMouseCol + 1
        End If
    End With
End Sub

Public Sub mnuHideCol_Click()
    Dim i As Integer
    Dim strMonth As String, intWeekWidth As Double
    
    With msf_MPInsertion
        If EngMonthIndex(.TextMatrix(1, intMouseCol)) = -1 Then
            LstHidenCol.AddItem "#" & intMouseCol & " : " & .ColWidth(intMouseCol) & " > " & .TextMatrix(0, intMouseCol)
            .ColWidth(intMouseCol) = 0
        Else
            strMonth = .TextMatrix(1, intMouseCol)
            intWeekWidth = .ColWidth(intMouseCol)
            i = intMouseCol
            While .TextMatrix(1, i) = .TextMatrix(1, intMouseCol)
                i = i - 1
            Wend
            i = i + 1
            While .TextMatrix(1, i) = .TextMatrix(1, intMouseCol)
                .ColWidth(i) = 0
                i = i + 1
            Wend
            LstHidenCol.AddItem "#_ : " & intWeekWidth & " > " & strMonth
        End If
        
    End With
End Sub

Public Sub mnuRateInfo_Click()
    Dim strMPPlanDimID As String, strMPMediumID As String, intMonth As Integer, intWeekYear As Integer
    Dim NettRate As Double
    Dim GrossRate As Double
    Dim MSC_Paid_Value As Double
    Dim MSC_Bonus_Value As Double
    Dim Club_Agency_Value As Double
    Dim Fee As Double, Other_Cost As Double
    Dim rsTemp As New ADODB.Recordset
    
    With msf_MPInsertion
        strMPPlanDimID = Left(.TextMatrix(.Row, .cols - 1), 19)
        rsTemp.Open "select mp_medium_id from mp_ids where mp_plan_dim_id='" & strMPPlanDimID & "'", ConnERP, 1, 3
        strMPMediumID = ""
        If Not rsTemp.EOF Then
            strMPMediumID = rsTemp(0)
        End If
        rsTemp.Close
        intMonth = EngMonthIndex(.TextMatrix(1, .col))
        intWeekYear = CInt(.TextMatrix(3, .col))
        NettRate = 0
        GrossRate = 0
        rsTemp.Open "select nett_rate,gross_rate,spot,isnull(other_cost,0) other_cost from mp_insertion where mp_plan_dim_id = '" & strMPPlanDimID & "' and week_year = " & intWeekYear, ConnERP, 1, 3
        If Not rsTemp.EOF Then
            NettRate = rsTemp(0) / rsTemp(2)
            GrossRate = rsTemp(1) / rsTemp(2)
            Other_Cost = rsTemp(3) / rsTemp(2)
        End If
        rsTemp.Close
        MSC_Paid_Value = 0
        MSC_Bonus_Value = 0
        Club_Agency_Value = 0
        rsTemp.Open "select msc_paid,msc_paid_on_flag,msc_bonus,msc_bonus_on_flag,club_agency,club_agency_on_flag from mp_monthly_activity where mp_medium_id='" & strMPMediumID & "' and month_number=" & intMonth, ConnERP, 1, 3
        If Not rsTemp.EOF Then
        
            If rsTemp("msc_paid_on_flag") = 1 Then
                MSC_Paid_Value = rsTemp("msc_paid") * NettRate / 100
            Else
                MSC_Paid_Value = rsTemp("msc_paid") * GrossRate / 100
            End If
            
            If rsTemp("msc_bonus_on_flag") = 1 Then
                MSC_Bonus_Value = rsTemp("msc_bonus") * NettRate / 100
            Else
                MSC_Bonus_Value = rsTemp("msc_bonus") * GrossRate / 100
            End If
            
            If rsTemp("club_agency_on_flag") = 1 Then
                Club_Agency_Value = rsTemp("club_agency") * NettRate / 100
            Else
                Club_Agency_Value = rsTemp("club_agency") * GrossRate / 100
            End If
            
        End If
        rsTemp.Close
        
        Fee = MSC_Paid_Value + MSC_Bonus_Value + Club_Agency_Value
        MsgBox ("RATE/SPOT :" & vbCrLf & "Gross Rate = " & FormatNumber(GrossRate, 2) & vbCrLf & "Nett Rate = " & FormatNumber(NettRate, 2) & vbCrLf & "Fee = " & FormatNumber(Fee, 2) & vbCrLf & "Other Cost = " & FormatNumber(Other_Cost, 2)), vbExclamation, strApplication_Name
    End With
End Sub

Public Sub MnuShowHidenCol_Click()
    Frm_MPInsertionHidenCol.show 1
End Sub

Public Sub MnuStartPeriode_Click()
'*****************************************************************************
' Nama Prosedur     :   MnuStartPeriode_Click
' Fungsi Prosedur   :   Menandai Week Start of periode
' Programer         :   Sistyo
'*****************************************************************************
    Dim pesan, flagExist As Boolean
    Dim fill_color As OLE_COLOR
    flagExist = False
    With msf_MPInsertion
        rsTemp.Open "select mp_tv_rf_id from mp_tv_reach_frequency where week_year_start <= " & .TextMatrix(3, .col) & " and week_year_end >= " & .TextMatrix(3, .col) & " and mp_plan_dim_id='" & Mid(.TextMatrix(.Row, .cols - 1), 1, 19) & "'", ConnERP, 1, 3
        If Not rsTemp.EOF Then
            flagExist = True
        End If
        rsTemp.Close
        If flagExist Then
            pesan = MsgBox("Can not start new periode here!", vbExclamation, strApplication_Name)
        Else
            FlagMarkStart = True
            MarkStartCol = .col
            If .CellBackColor = Shape_unlocked.FillColor Then
                fill_color = Shape_unlocked.FillColor
            Else
                fill_color = LegendTVRF.FillColor
            End If
            .CellBackColor = fill_color
            If Right(.TextMatrix(.Row, .cols - 1), 5) = "REACH" Then
                .Row = .Row - 1
                .CellBackColor = fill_color
                MarkStartRow = .Row
            Else
                MarkStartRow = .Row
                .Row = .Row + 1
                .CellBackColor = fill_color
            End If
        End If
    End With
    
End Sub

Public Sub MnuEndPeriode_Click()
'*****************************************************************************
' Nama Prosedur     :   MnuEndPeriode_Click
' Fungsi Prosedur   :   Menandai Week End of periode, dan update ke database
' Programer         :   Sistyo
'*****************************************************************************
    Dim pesan, strSql As String, flagExist As Boolean
    Dim i As Integer, MarkEndCol As Integer
    Dim fill_color As OLE_COLOR
    flagExist = False
    With msf_MPInsertion
        rsTemp.Open "select mp_tv_rf_id from mp_tv_reach_frequency where week_year_start >= " & .TextMatrix(3, MarkStartCol) & " and week_year_start <= " & .TextMatrix(3, .col) & " and mp_plan_dim_id='" & Mid(.TextMatrix(.Row, .cols - 1), 1, 19) & "'", ConnERP, 1, 3
        If Not rsTemp.EOF Then
            flagExist = True
        End If
        rsTemp.Close
        If flagExist Then
            pesan = MsgBox("Can not End Mark here, Out of Range!", vbExclamation, strApplication_Name)
        Else
            FlagMarkStart = False
            MarkEndCol = .col
            If .CellBackColor = Shape_unlocked.FillColor Then
                fill_color = Shape_unlocked.FillColor
            Else
                fill_color = LegendTVRF
            End If
            
            For i = MarkStartCol To MarkEndCol
                .col = i
                .Row = MarkStartRow
                .Text = "1+"
                .CellBackColor = fill_color
                .Row = MarkStartRow + 1
                .CellBackColor = fill_color
                .Text = "100%"
            Next
            Dim strMarketCode As String, strMarketName As String
            strMarketCode = "null"
            strMarketName = "null"
            strSql = "select market_code,market_name from mp_medium_detail where mp_medium_detail_id = (select mp_medium_detail_id from mp_plan_dimension where mp_plan_dim_id='" & Mid(.TextMatrix(.Row, .cols - 1), 1, 19) & "')"
            rsTemp.Open strSql, ConnERP, 1, 3
            If Not rsTemp.EOF Then
                If Not IsNull(rsTemp("market_code").Value) Then
                    strMarketCode = rsTemp("market_code").Value
                End If
                If Not IsNull(rsTemp("market_name").Value) Then
                    strMarketName = rsTemp("market_name").Value
                End If
            End If
            rsTemp.Close
            
            strSql = "insert into mp_tv_reach_frequency(mp_plan_dim_id,week_year_start,week_year_end,week_commencing_start,week_commencing_end,month_start,month_end,frequency_code,frequency_name,reach,market_code,market_name) values "
            strSql = strSql & "('" & Mid(.TextMatrix(.Row, .cols - 1), 1, 19) & "'," & MarkStartCol - 5 & "," & MarkEndCol - 5 & ",'" & .TextMatrix(4, MarkStartCol) & "','" & .TextMatrix(4, MarkEndCol) & "'," & EngMonthIndex(.TextMatrix(1, MarkStartCol)) & "," & EngMonthIndex(.TextMatrix(1, MarkEndCol)) & ",1,'1+',100," & strMarketCode & ",'" & strMarketName & "')"
            ConnERP.Execute strSql
            .Refresh
        End If
    End With
    
End Sub

Public Sub MnuUnFreeze_Click()
    Dim i As Integer
    With msf_MPInsertion
        If EngMonthIndex(.TextMatrix(1, intMouseCol)) <> -1 Then
            i = intMouseCol
            While .TextMatrix(1, i) = .TextMatrix(1, intMouseCol)
                i = i - 1
            Wend
            msf_MPInsertion.FixedCols = i + 1
        Else
            msf_MPInsertion.FixedCols = intMouseCol
        End If
    End With
End Sub

Private Sub msf_MPInsertion_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    'ToolTipText
    Dim Teks As String
    Teks = Trim(msf_MPInsertion.TextMatrix(msf_MPInsertion.MouseRow, msf_MPInsertion.MouseCol))
    If Teks <> "" Then
        msf_MPInsertion.ToolTipText = Teks
    End If
End Sub

Private Sub PicButtonPanahBawah_Click()
    With LstFrequency
        If .Visible Then
            .Visible = False
        Else
            .Top = PicButtonPanahBawah.Top + PicButtonPanahBawah.Height + 30
            .Left = PicButtonPanahBawah.Left + PicButtonPanahBawah.Width - .Width
            .Visible = True
            .SetFocus
        End If
    End With
End Sub

Private Sub picViewMonth_Click()
    If FrameViewMonth1.Visible Then
        FrameViewMonth1.Visible = False
    Else
        FrameViewMonth1.Visible = True
    End If
End Sub

Private Sub txtMPInsertion_KeyDown(KeyCode As Integer, Shift As Integer)

    Select Case KeyCode
        Case 13 'Enter
            msf_MPInsertion.SetFocus
        Case 27 'Escape
            txtMPInsertion.Visible = False
            msf_MPInsertion.SetFocus
        Case 37 'panah kiri
            If Not EditMode Then
                msf_MPInsertion.SetFocus
                If intCol > msf_MPInsertion.FixedCols Then
                    msf_MPInsertion.col = intCol - 1
                End If
            End If
        Case 38 'panah atas
            msf_MPInsertion.SetFocus
            If intRow > msf_MPInsertion.FixedRows + 1 Then
                msf_MPInsertion.Row = intRow - 1
            End If
        Case 39 'panah kanan
            If Not EditMode Then
                msf_MPInsertion.SetFocus
                If intCol < msf_MPInsertion.cols - 2 Then
                    msf_MPInsertion.col = intCol + 1
                End If
            End If
        Case 40 'panah bawah
            msf_MPInsertion.SetFocus
            If intRow < msf_MPInsertion.Rows - 1 Then
                msf_MPInsertion.Row = intRow + 1
            End If
    End Select
    
End Sub

Private Sub DeleteInsertion()

    Dim pesan
    pesan = MsgBox("Delete Insertion?", vbQuestion + vbYesNo, strApplication_Name)
    If pesan = 6 Then
        ConnERP.Execute "delete from mp_insertion where mp_plan_dim_id = '" & Mid(msf_MPInsertion.TextMatrix(intRow, msf_MPInsertion.cols - 1), 1, 19) & "' and week_year=" & msf_MPInsertion.TextMatrix(3, intCol)
        'update mp_master
            ConnERP.Execute "update mp_master set last_update_by='" & strLogin_User & "',last_update_date = getdate() where mp_number='" & cboMPNumber.Text & "'"
        'GRP/ins
            msf_MPInsertion.TextMatrix(intRow, msf_MPInsertion.cols - 3) = Val(msf_MPInsertion.TextMatrix(intRow, msf_MPInsertion.cols - 3)) - Val(msf_MPInsertion.TextMatrix(intRow, intCol))
            If Val(msf_MPInsertion.TextMatrix(intRow, msf_MPInsertion.cols - 3)) = 0 Then
                msf_MPInsertion.TextMatrix(intRow, msf_MPInsertion.cols - 3) = ""
            End If
            msf_MPInsertion.TextMatrix(intRow, intCol) = ""
        Call RefreshTotal(intRow)
        Call RefreshBudget(intRow, intCol)
    End If
    
End Sub

Private Sub txtMPInsertion_KeyPress(KeyAscii As Integer)
    'Input harus angka
    If (Chr(KeyAscii) < "0" Or Chr(KeyAscii) > "9") And KeyAscii <> 8 And KeyAscii <> 13 And KeyAscii <> 27 Then
            KeyAscii = 0
            Beep
    End If
End Sub

Private Sub txtMPInsertion_LostFocus()
'*****************************************************************************
' Nama Submodul         :  txtMPInsertion_LostFocus
' Fungsi Submodul       :  Mengakhiri Edit Insertion, action : UPDATE, INSERT INSERTION, OR IGNORE CHANGES
' Nama Programmer       :  Sistyo
' Tgl Pembuatan         :  18 Agustus 2004
' Last Update           :  15 November 2005/Sistyo
' Changes History       :  - 15 Nov 2005 : keep Original TV Objective when update insertion
'******************************************************************************
    Dim DataExist As Boolean, oldVal As Long, strSql As String, strMedium As String
    Dim isUpdate As Boolean, old_mp_tv_rf_id As String
    rsTemp.Open "select medium_name from mp_medium where mp_medium_id = (select mp_medium_id from mp_ids where mp_plan_dim_id = '" & Mid(msf_MPInsertion.TextMatrix(intRow, msf_MPInsertion.cols - 1), 1, 19) & "')", ConnERP, 1, 3
        strMedium = rsTemp(0)
    rsTemp.Close
    oldVal = Val(msf_MPInsertion.TextMatrix(intRow, intCol))
    If txtMPInsertion.Visible Then
        txtMPInsertion.Visible = False
        If Val(txtMPInsertion.Text) <> 0 And Val(txtMPInsertion.Text) <> oldVal Then
            'Delete old
            isUpdate = False
                If Trim(msf_MPInsertion.TextMatrix(intRow, intCol)) <> "" Then
                    'Get MP_tv_rf_id
                    isUpdate = True
                    strSql = "select mp_tv_rf_id from mp_insertion where mp_plan_dim_id = '" & Mid(msf_MPInsertion.TextMatrix(intRow, msf_MPInsertion.cols - 1), 1, 19) & "' and week_year = '" & Trim(msf_MPInsertion.TextMatrix(3, intCol)) & "'"
                    rsTemp.Open strSql, ConnERP, 1, 3
                        old_mp_tv_rf_id = IIf(IsNull(rsTemp(0).Value), "NULL", rsTemp(0).Value)
                    rsTemp.Close
                    'Delete old
                    ConnERP.Execute "delete from mp_insertion where mp_plan_dim_id = '" & Mid(msf_MPInsertion.TextMatrix(intRow, msf_MPInsertion.cols - 1), 1, 19) & "' and week_year = '" & Trim(msf_MPInsertion.TextMatrix(3, intCol)) & "'"
                End If
            'insert new
                If isUpdate Then
                    strSql = "insert into mp_insertion(mp_plan_dim_id,[month],week_year,week_commencing,spot,nett_rate,gross_rate,mp_tv_rf_id) values ('" & Mid(msf_MPInsertion.TextMatrix(intRow, msf_MPInsertion.cols - 1), 1, 19) & "', " & EngMonthIndex(Trim(msf_MPInsertion.TextMatrix(1, intCol))) & "," & Trim(msf_MPInsertion.TextMatrix(3, intCol)) & ",'" & msf_MPInsertion.TextMatrix(4, intCol) & "'," & Val(txtMPInsertion.Text) & ",0,0," & old_mp_tv_rf_id & ")"
                Else
                    strSql = "insert into mp_insertion(mp_plan_dim_id,[month],week_year,week_commencing,spot,nett_rate,gross_rate) values ('" & Mid(msf_MPInsertion.TextMatrix(intRow, msf_MPInsertion.cols - 1), 1, 19) & "', " & EngMonthIndex(Trim(msf_MPInsertion.TextMatrix(1, intCol))) & "," & Trim(msf_MPInsertion.TextMatrix(3, intCol)) & ",'" & msf_MPInsertion.TextMatrix(4, intCol) & "'," & Val(txtMPInsertion.Text) & ",0,0)"
                End If
                ConnERP.Execute strSql
            'update mp_master
                ConnERP.Execute "update mp_master set last_update_by='" & strLogin_User & "',last_update_date = getdate() where mp_number='" & cboMPNumber.Text & "'"
            'update flexgrid
                msf_MPInsertion.TextMatrix(intRow, intCol) = Val(txtMPInsertion.Text) 'spot
            'GRP/ins
                msf_MPInsertion.TextMatrix(intRow, msf_MPInsertion.cols - 3) = Val(msf_MPInsertion.TextMatrix(intRow, msf_MPInsertion.cols - 3)) + (Val(txtMPInsertion.Text) - oldVal)
                If Val(msf_MPInsertion.TextMatrix(intRow, msf_MPInsertion.cols - 3)) = 0 Then
                    msf_MPInsertion.TextMatrix(intRow, msf_MPInsertion.cols - 3) = ""
                End If
            Call RefreshTotal(intRow)
            Call RefreshBudget(intRow, intCol)
        End If
    End If
End Sub

Private Sub RefreshTotal(vRow As Double)
'*****************************************************************************
' Nama Submodul         :  RefreshTotal
' Fungsi Submodul       :  Menampilkan Total after Update
' Nama Programmer       :  Sistyo
' Tgl Pembuatan         :  30 Sept 2004
' Last Update           :  30 Sept 2004/Sistyo
'******************************************************************************
    Dim strSql As String, strMediumID As String
    Dim intTotalNett As Double
    Dim field_no As Integer
    Dim feePaid As Double, FeeBonus As Double, FeeClub As Double
    Dim rst As New ADODB.Recordset
    
    With msf_MPInsertion
        'Get Medium ID
        strSql = strSql & "select mp_medium_id from mp_ids "
        strSql = strSql & "where mp_plan_dim_id = '" & Mid(.TextMatrix(vRow, .cols - 1), 1, 19) & "'"
        rst.Open strSql, ConnERP, 1, 3
        If Not rst.EOF Then
            strMediumID = rst("mp_medium_id").Value
        End If
        rst.Close
    
        'Load Fee Component
        strSql = "select month_number,MSC_Paid/100 MSC_Paid,case MSC_Paid_On_Flag when 1 then 1 when 2 then 0 when 3 then 0 when 4 then 1 when 0 then 0 end MSC_Paid_On_Flag,"
        strSql = strSql & " MSC_Bonus/100 MSC_Bonus,case MSC_Bonus_On_Flag when 1 then 1 when 2 then 0 when 3 then 0 when 4 then 1 when 0 then 0 end MSC_Bonus_On_Flag,"
        strSql = strSql & " Club_Agency/100 Club_Agency,case Club_Agency_On_Flag when 1 then 1 when 2 then 0 when 3 then 0 when 4 then 1 when 0 then 0 end  Club_Agency_On_Flag"
        strSql = strSql & " from mp_monthly_activity"
        strSql = strSql & " where mp_medium_id = '" & strMediumID & "' order by month_number"
        ReDim arrFee(12, 6)
        rst.Open strSql, ConnERP, 1, 3
        While Not rst.EOF
            For field_no = 1 To rst.Fields.Count - 1
                arrFee(rst(0), field_no) = rst(field_no).Value
            Next
            rst.MoveNext
        Wend
        rst.Close
        
        'Load Total
        If txtViewMonth.Text = "ALL" Then
            rst.Open "select nett_rate,gross_rate,[month],isnull(other_cost,0) other_cost from mp_insertion where mp_plan_dim_id = '" & Mid(.TextMatrix(vRow, .cols - 1), 1, 19) & "'", ConnERP, 1, 3
        Else
            rst.Open "select nett_rate,gross_rate,[month],isnull(other_cost,0) other_cost from mp_insertion where mp_plan_dim_id = '" & Mid(.TextMatrix(vRow, .cols - 1), 1, 19) & "' and [month] in " & strViewMonth, ConnERP, 1, 3
        End If
        
        intTotalNett = 0
        While Not rst.EOF
            
            'TOTAL KE KANAN per station
            If arrFee(rst("month").Value, 2) = 1 Then
                feePaid = arrFee(rst("month").Value, 1) * rst("nett_rate").Value
            Else
                feePaid = arrFee(rst("month").Value, 1) * rst("gross_rate").Value
            End If
            
            If arrFee(rst("month").Value, 4) = 1 Then
                FeeBonus = arrFee(rst("month").Value, 3) * rst("nett_rate").Value
            Else
                FeeBonus = arrFee(rst("month").Value, 3) * rst("gross_rate").Value
            End If
            
            If arrFee(rst("month").Value, 6) = 1 Then
                FeeClub = arrFee(rst("month").Value, 5) * rst("nett_rate").Value
            Else
                FeeClub = arrFee(rst("month").Value, 5) * rst("gross_rate").Value
            End If
            
            intTotalNett = intTotalNett + rst("nett_rate") + feePaid + FeeBonus + FeeClub + rst("other_cost")
            'End TOTAL KE KANAN per station
            
            rst.MoveNext
        Wend
        .TextMatrix(.Row, .cols - 2) = FormatNumber(intTotalNett, 2)
        rst.Close
        Set rst = Nothing
    End With
End Sub

Private Sub RefreshBudget(vRow As Double, vCol As Double)
'*****************************************************************************
' Nama Submodul         :  RefreshBudget
' Fungsi Submodul       :  Menampilkan Budget after Update
' Nama Programmer       :  Sistyo
' Tgl Pembuatan         :  18 Agustus 2004
' Last Update           :  18 Agustus 2004/Sistyo
'******************************************************************************
    Dim strMP_Plan_Dim_Id As String, strMP_Medium_Id As String, intMonth As Integer
    Dim strMonth As String, intBudgetRow As Integer, intBudgetCol As Integer
    Dim budget_old As Double
    Dim budget_new As Double
    With msf_MPInsertion
        strMP_Plan_Dim_Id = Mid(.TextMatrix(vRow, .cols - 1), 1, 19)
        intMonth = EngMonthIndex(Trim(.TextMatrix(1, vCol)))
        rsTemp.Open "select distinct mp_medium_id from mp_ids where mp_plan_dim_id = '" & strMP_Plan_Dim_Id & "'", ConnERP, 1, 3
        If Not rsTemp.EOF Then
            strMP_Medium_Id = rsTemp(0)
        End If
        rsTemp.Close
        rsTemp.Open "select budget + msc_paid_value + msc_bonus_value + club_agency_value + isnull(other_cost,0) other_cost from mp_monthly_activity where mp_medium_id = '" & strMP_Medium_Id & "' and [month_number] = " & intMonth, ConnERP, 1, 3
        If Not rsTemp.EOF Then
            
                intBudgetRow = vRow
                While Mid(Trim(.TextMatrix(intBudgetRow, 0)), 1, 9) <> "Sub Total"
                    intBudgetRow = intBudgetRow + 1
                Wend
                intBudgetCol = vCol
                
                budget_old = CDbl(RemoveNumberFormat(.TextMatrix(intBudgetRow, intBudgetCol)))
                budget_new = rsTemp(0)
                
                'refresh remaining budget
                txt_remaining_budget.Text = FormatNumber(CDbl(RemoveNumberFormat(txt_remaining_budget.Text)) - (budget_new - budget_old), 2)
                If CDbl(RemoveNumberFormat(txt_remaining_budget.Text)) < 0 Then
                    txt_remaining_budget.ForeColor = vbRed
                Else
                    txt_remaining_budget.ForeColor = vbBlack
                End If
                
                'refresh total per tahun
                If .TextMatrix(intBudgetRow, .cols - 2) <> "" Then
                    .TextMatrix(intBudgetRow, .cols - 2) = FormatNumber(CDbl(RemoveNumberFormat(.TextMatrix(intBudgetRow, .cols - 2))) - budget_old + budget_new, 2)
                Else
                    .TextMatrix(intBudgetRow, .cols - 2) = FormatNumber(budget_new, 2)
                End If
                'tampilkan new budget
                strMonth = .TextMatrix(1, vCol)
                While strMonth = .TextMatrix(1, intBudgetCol)
                    .TextMatrix(intBudgetRow, intBudgetCol) = String(intMonth, " ") & FormatNumber(rsTemp(0)) & String(intMonth, " ")
                    intBudgetCol = intBudgetCol - 1
                Wend
                intBudgetCol = vCol + 1
                While strMonth = .TextMatrix(1, intBudgetCol)
                    .TextMatrix(intBudgetRow, intBudgetCol) = String(intMonth, " ") & FormatNumber(rsTemp(0)) & String(intMonth, " ")
                    intBudgetCol = intBudgetCol + 1
                Wend
                
            'end of tampilkan budget
        End If
        rsTemp.Close
    End With
    
End Sub

Private Sub incrProgressBar()

    ProgressBarExport.Value = ProgressBarExport.Value + 1
    lblPercent.Caption = ProgressBarExport.Value * 100 \ ProgressBarExport.Max & "% Complete..."
    lblPercent.Refresh

End Sub

Private Sub db_ExportToExcel()
'<CSCM>
'********************************************************************************
'Procedure Name     : db_ExportToExcel
'Procedure Function : ---
'Input Parameter    : ---
'Output Parameter   : ---
'Date               : 3/10/2016
'LastUpdate/By      : Tedi / Kreatif
'Name Before        : cmdExportToExcel_Click
'********************************************************************************
'</CSCM>
    If cboMPNumber.Text = "" Then
        MsgBox "Select MP Number!", vbExclamation, strApplication_Name
        Exit Sub
    End If

    Me.MousePointer = vbHourglass
    Call ExportToExcel
    Me.MousePointer = vbNormal
    
End Sub

Private Sub ExportToExcel()
'*****************************************************************************
' Nama Submodul         :  ExportToExcel
' Fungsi Submodul       :  Export To Excel
' Nama Programmer       :  Sistyo
' Tgl Pembuatan         :  18 Agustus 2004
' Last Update           :  18 Agustus 2004/Sistyo
'******************************************************************************
    Const xlTopRow = 2
    Const xlHeaderrows = 5
    Const xlSubHeaderRows = 7
    Const OrigHeaderRows = 5
    Const xlLeftCol = 1
    Dim xlApp As Object, xlWB As Object, xlws As Object
    Dim sCol As Integer, vCol As Integer, vRow As Integer, xlRow As Integer
    Dim intTask As Integer, xlStartRowTask As Integer, strSql As String
    Dim intTVRow As Integer, intTVCol As Integer, Logo_Name As String
    Dim fso, f1, Teks
    Dim pesan
    Dim i As Integer
    
    Set xlApp = CreateObject("Excel.Application")
    Set xlWB = xlApp.Workbooks.Add
    Set xlws = xlWB.Worksheets(1)
    
    
    
    xlws.Name = "Media Plan"
    FrameProgressBar.Visible = True
    FrameProgressBar.Refresh
    lblPleaseWait.Caption = "Exporting Media Plan.."
    lblPercent.Caption = ""
    lblPercent.Refresh
    ProgressBarExport.Max = 28 + msf_MPInsertion.cols - 14 + ((msf_MPInsertion.Rows - 4) * (msf_MPInsertion.cols - 3)) + 29
    ProgressBarExport.Min = 0
    ProgressBarExport.Value = 0
    intTVRow = 0
    
    With xlws
        .Activate
        .Cells.Locked = True
        .Cells.FormulaHidden = False
        'set default font
            .Cells.EntireColumn.Font.Size = 8
            .Cells.EntireColumn.Font.Name = "Sylfaen"
        'SET HEADER BORDER
            .Range(.Cells(xlHeaderrows + (xlTopRow - 1), xlLeftCol).Address & ":" & .Cells(xlHeaderrows + (xlTopRow - 1), msf_MPInsertion.cols - 2 + (xlLeftCol - 1)).Address).Borders(xlEdgeBottom).LineStyle = xlDouble
            .Range(.Cells(xlHeaderrows + (xlTopRow - 1), xlLeftCol).Address & ":" & .Cells(xlHeaderrows + (xlTopRow - 1), msf_MPInsertion.cols - 2 + (xlLeftCol - 1)).Address).Borders(xlEdgeBottom).Weight = xlThick
            .Range(.Cells(xlHeaderrows + (xlTopRow - 1), xlLeftCol).Address & ":" & .Cells(xlHeaderrows + (xlTopRow - 1), msf_MPInsertion.cols - 2 + (xlLeftCol - 1)).Address).Borders(xlEdgeBottom).ColorIndex = xlAutomatic
            
            Call incrProgressBar
        'setting col width
            .Columns(1 + (xlLeftCol - 1)).ColumnWidth = 31.57
            .Columns(2 + (xlLeftCol - 1)).ColumnWidth = 9.29
            .Columns(3 + (xlLeftCol - 1)).ColumnWidth = 19
            .Columns(4 + (xlLeftCol - 1)).ColumnWidth = 22.14
            .Columns(5 + (xlLeftCol - 1)).ColumnWidth = 33.57
            For vCol = 6 To msf_MPInsertion.cols - 4
                .Columns(vCol + (xlLeftCol - 1)).ColumnWidth = 3.14
            Next
            .Columns((msf_MPInsertion.cols - 2) + (xlLeftCol - 1)).ColumnWidth = 12.43
        
            Call incrProgressBar
        'init header
            .Range(.Cells(xlTopRow, xlLeftCol).Address & ":" & .Cells(xlTopRow + 5, xlLeftCol + msf_MPInsertion.cols - 3).Address).Interior.ColorIndex = 2
            .Range(.Cells(xlTopRow, xlLeftCol).Address & ":" & .Cells(xlTopRow + 5, xlLeftCol + msf_MPInsertion.cols - 3).Address).Interior.Pattern = xlSolid
            
            .Range(.Cells(xlTopRow, xlLeftCol).Address & ":" & .Cells(xlHeaderrows + (xlTopRow - 1), 2 + (xlLeftCol - 1)).Address).Merge
            .Range(.Cells(xlTopRow, xlLeftCol).Address & ":" & .Cells(xlHeaderrows + (xlTopRow - 1), 2 + (xlLeftCol - 1)).Address).HorizontalAlignment = 3
            .Range(.Cells(xlTopRow, xlLeftCol).Address & ":" & .Cells(xlHeaderrows + (xlTopRow - 1), 2 + (xlLeftCol - 1)).Address).VerticalAlignment = 2
        
            On Error GoTo NoImage:
            .Range(.Cells(xlTopRow, xlLeftCol).Address & ":" & .Cells(xlHeaderrows + (xlTopRow - 1), 2 + (xlLeftCol - 1)).Address).Locked = False
            Logo_Name = Mid(cboMPNumber.Text, 1, 4) & ".JPG"
            .Pictures.Insert (App.Path & "\Brand_Logo\" & Logo_Name)
            .Pictures(1).ShapeRange.IncrementLeft 26.25
            .Pictures(1).ShapeRange.IncrementTop 28.25
        
NoImage:
            If Err.Number = 1004 Then
                .Range(.Cells(xlTopRow, xlLeftCol).Address & ":" & .Cells(xlHeaderrows + (xlTopRow - 1), 2 + (xlLeftCol - 1)).Address).Value = "[Brand Logo Here]"
            End If
        
            strSql = "select country,[year],mp_number,release_date,prev_release_date,latest_approved_date, "
            strSql = strSql & "client_approved_by,noted_by from mp_master where mp_number = '" & cboMPNumber.Text & "'"
        
            rsTemp.Open strSql, ConnERP, 1, 3
        
            .Cells(2 + (xlTopRow - 1), (3 + xlLeftCol - 1)).Value = "Country :"
            '.Cells(2 + (xlTopRow - 1), (3 + xlLeftCol - 1)).Locked = True
            
            .Cells(3 + (xlTopRow - 1), (3 + xlLeftCol - 1)).Value = rsTemp("country") 'plan country
            '.Cells(3 + (xlTopRow - 1), (3 + xlLeftCol - 1)).Locked = True
            
            .Cells(2 + (xlTopRow - 1), (4 + xlLeftCol - 1)).Value = "Year :"
            '.Cells(2 + (xlTopRow - 1), (4 + xlLeftCol - 1)).Locked = True
            
            .Cells(3 + (xlTopRow - 1), (4 + xlLeftCol - 1)).Value = "'" & rsTemp("year") 'plan Year
            '.Cells(3 + (xlTopRow - 1), (4 + xlLeftCol - 1)).Locked = True
            
            .Cells(2 + (xlTopRow - 1), (6 + xlLeftCol - 1)).Font.Size = 9
            .Cells(2 + (xlTopRow - 1), (6 + xlLeftCol - 1)).Font.ColorIndex = 5 'blue
            .Cells(2 + (xlTopRow - 1), (6 + xlLeftCol - 1)).Font.Bold = True
            .Cells(2 + (xlTopRow - 1), (6 + xlLeftCol - 1)).Value = "Status : Media Plan #" & Val(Right(cboMPNumber.Text, 4)) '?? seq of mp_number
            .Cells(2 + (xlTopRow - 1), (6 + xlLeftCol - 1)).Characters(Start:=1, Length:=8).Font.Bold = False
            '.Cells(2 + (xlTopRow - 1), (6 + xlLeftCol - 1)).Locked = True
            
            .Cells(3 + (xlTopRow - 1), (7 + xlLeftCol - 1)).Font.Size = 8
            .Cells(3 + (xlTopRow - 1), (7 + xlLeftCol - 1)).Font.ColorIndex = 5 'blue
            .Cells(3 + (xlTopRow - 1), (7 + xlLeftCol - 1)).Font.Bold = True
            .Cells(3 + (xlTopRow - 1), (7 + xlLeftCol - 1)).Value = "'(" & cboMPNumber.Text & ")" '?? mp_number
            '.Cells(3 + (xlTopRow - 1), (7 + xlLeftCol - 1)).Locked = True
            
            .Cells(2 + (xlTopRow - 1), (11 + xlLeftCol - 1)).Font.Size = 9
            .Cells(2 + (xlTopRow - 1), (11 + xlLeftCol - 1)).Value = "Release Date: " & rsTemp("release_date")
            '.Cells(2 + (xlTopRow - 1), (11 + xlLeftCol - 1)).Locked = True
            
            .Cells(3 + (xlTopRow - 1), (11 + xlLeftCol - 1)).Font.Size = 9
            .Cells(3 + (xlTopRow - 1), (11 + xlLeftCol - 1)).Value = "Previous Release Date : " & rsTemp("prev_release_date")
            '.Cells(3 + (xlTopRow - 1), (11 + xlLeftCol - 1)).Locked = True
            
            .Cells(4 + (xlTopRow - 1), (11 + xlLeftCol - 1)).Font.Size = 9
            .Cells(4 + (xlTopRow - 1), (11 + xlLeftCol - 1)).Value = "Latest Approval date : " & rsTemp("latest_approved_date")
            '.Cells(4 + (xlTopRow - 1), (11 + xlLeftCol - 1)).Locked = True
            
            .Cells(2 + (xlTopRow - 1), (19 + xlLeftCol - 1)).Font.Size = 9
            .Cells(2 + (xlTopRow - 1), (19 + xlLeftCol - 1)).Value = "Approved By:"
            '.Cells(2 + (xlTopRow - 1), (19 + xlLeftCol - 1)).Locked = True
            
            .Cells(3 + (xlTopRow - 1), (19 + xlLeftCol - 1)).Font.Size = 9
            .Cells(3 + (xlTopRow - 1), (19 + xlLeftCol - 1)).Value = rsTemp("client_approved_by")
            '.Cells(3 + (xlTopRow - 1), (19 + xlLeftCol - 1)).Locked = True
            
            .Cells(2 + (xlTopRow - 1), (24 + xlLeftCol - 1)).Font.Size = 9
            .Cells(2 + (xlTopRow - 1), (24 + xlLeftCol - 1)).Value = "Noted By:"
            '.Cells(2 + (xlTopRow - 1), (24 + xlLeftCol - 1)).Locked = True
            
            .Cells(3 + (xlTopRow - 1), (24 + xlLeftCol - 1)).Font.Size = 9
            .Cells(3 + (xlTopRow - 1), (24 + xlLeftCol - 1)).Value = rsTemp("noted_by")
            '.Cells(3 + (xlTopRow - 1), (24 + xlLeftCol - 1)).Locked = True
            
            rsTemp.Close
            Call incrProgressBar
        
        'init subheader
            .Range(.Cells(xlTopRow + 6, xlLeftCol).Address & ":" & .Cells(xlTopRow + 10, xlLeftCol + 4).Address).Interior.ColorIndex = 2
            .Range(.Cells(xlTopRow + 6, xlLeftCol).Address & ":" & .Cells(xlTopRow + 10, xlLeftCol + 4).Address).Interior.Pattern = xlSolid
        
            .Cells(xlTopRow + 8, xlLeftCol + 4).Font.Bold = True
            .Cells(xlTopRow + 8, xlLeftCol + 4).Font.ColorIndex = 5
            .Cells(xlTopRow + 8, xlLeftCol + 4).HorizontalAlignment = xlRight
            .Cells(xlTopRow + 8, xlLeftCol + 4).Value = "Week"
            '.Cells(xlTopRow + 8, xlLeftCol + 4).Locked = True
            
            Call incrProgressBar
        
        'Bulan
            vRow = xlTopRow + 6 'bulan
            vCol = xlLeftCol + 5 'start bulan
            
            sCol = vCol
            While Trim(msf_MPInsertion.TextMatrix(vRow - xlTopRow - 5, vCol)) = "January"
                vCol = vCol + 1
            Wend
            sRange_Jan = .Cells(1, sCol).Address(True, False, xlA1)
            sRange_Jan = Left(sRange_Jan, InStr(1, sRange_Jan, "$") - 1)
            eRange_Jan = .Cells(1, vCol - 1).Address(True, False, xlA1)
            eRange_Jan = Left(eRange_Jan, InStr(1, eRange_Jan, "$") - 1)
            .Range(sRange_Jan & CStr(vRow) & ":" & eRange_Jan & CStr(vRow)).Merge
            .Range(sRange_Jan & CStr(vRow) & ":" & eRange_Jan & CStr(vRow)).Interior.ColorIndex = 34
            .Range(sRange_Jan & CStr(vRow) & ":" & eRange_Jan & CStr(vRow)).Value = "January"
            '.Range(sRange_Jan & CStr(vRow) & ":" & eRange_Jan & CStr(vRow)).Locked = True
            
            Call incrProgressBar
             
            sCol = vCol
            While Trim(msf_MPInsertion.TextMatrix(vRow - xlTopRow - 5, vCol)) = "February"
                vCol = vCol + 1
            Wend
            sRange_Feb = .Cells(1, sCol).Address(True, False, xlA1)
            sRange_Feb = Left(sRange_Feb, InStr(1, sRange_Feb, "$") - 1)
            eRange_Feb = .Cells(1, vCol - 1).Address(True, False, xlA1)
            eRange_Feb = Left(eRange_Feb, InStr(1, eRange_Feb, "$") - 1)
            .Range(sRange_Feb & CStr(vRow) & ":" & eRange_Feb & CStr(vRow)).Merge
            .Range(sRange_Feb & CStr(vRow) & ":" & eRange_Feb & CStr(vRow)).Interior.ColorIndex = 37
            .Range(sRange_Feb & CStr(vRow) & ":" & eRange_Feb & CStr(vRow)).Value = "February"
            '.Range(sRange_Feb & CStr(vRow) & ":" & eRange_Feb & CStr(vRow)).Locked = True
            
            Call incrProgressBar
            
            sCol = vCol
            While Trim(msf_MPInsertion.TextMatrix(vRow - xlTopRow - 5, vCol)) = "March"
                vCol = vCol + 1
            Wend
            sRange_Mar = .Cells(1, sCol).Address(True, False, xlA1)
            sRange_Mar = Left(sRange_Mar, InStr(1, sRange_Mar, "$") - 1)
            eRange_Mar = .Cells(1, vCol - 1).Address(True, False, xlA1)
            eRange_Mar = Left(eRange_Mar, InStr(1, eRange_Mar, "$") - 1)
            .Range(sRange_Mar & CStr(vRow) & ":" & eRange_Mar & CStr(vRow)).Merge
            
            .Range(sRange_Mar & CStr(vRow) & ":" & eRange_Mar & CStr(vRow)).Interior.ColorIndex = 33
            .Range(sRange_Mar & CStr(vRow) & ":" & eRange_Mar & CStr(vRow)).Value = "March"
            '.Range(sRange_Mar & CStr(vRow) & ":" & eRange_Mar & CStr(vRow)).Locked = True
            
            Call incrProgressBar
            
            sCol = vCol
            While Trim(msf_MPInsertion.TextMatrix(vRow - xlTopRow - 5, vCol)) = "April"
                vCol = vCol + 1
            Wend
            sRange_Apr = .Cells(1, sCol).Address(True, False, xlA1)
            sRange_Apr = Left(sRange_Apr, InStr(1, sRange_Apr, "$") - 1)
            eRange_Apr = .Cells(1, vCol - 1).Address(True, False, xlA1)
            eRange_Apr = Left(eRange_Apr, InStr(1, eRange_Apr, "$") - 1)
            .Range(sRange_Apr & CStr(vRow) & ":" & eRange_Apr & CStr(vRow)).Merge
            .Range(sRange_Apr & CStr(vRow) & ":" & eRange_Apr & CStr(vRow)).Interior.ColorIndex = 34
            .Range(sRange_Apr & CStr(vRow) & ":" & eRange_Apr & CStr(vRow)).Value = "April"
            '.Range(sRange_Apr & CStr(vRow) & ":" & eRange_Apr & CStr(vRow)).Locked = True
            
            Call incrProgressBar

            sCol = vCol
            While Trim(msf_MPInsertion.TextMatrix(vRow - xlTopRow - 5, vCol)) = "May"
                vCol = vCol + 1
            Wend
            sRange_May = .Cells(1, sCol).Address(True, False, xlA1)
            sRange_May = Left(sRange_May, InStr(1, sRange_May, "$") - 1)
            eRange_May = .Cells(1, vCol - 1).Address(True, False, xlA1)
            eRange_May = Left(eRange_May, InStr(1, eRange_May, "$") - 1)
            .Range(sRange_May & CStr(vRow) & ":" & eRange_May & CStr(vRow)).Merge
            .Range(sRange_May & CStr(vRow) & ":" & eRange_May & CStr(vRow)).Interior.ColorIndex = 37
            .Range(sRange_May & CStr(vRow) & ":" & eRange_May & CStr(vRow)).Value = "May"
            '.Range(sRange_May & CStr(vRow) & ":" & eRange_May & CStr(vRow)).Locked = True
            
            Call incrProgressBar

            sCol = vCol
            While Trim(msf_MPInsertion.TextMatrix(vRow - xlTopRow - 5, vCol)) = "June"
                vCol = vCol + 1
            Wend
            sRange_Jun = .Cells(1, sCol).Address(True, False, xlA1)
            sRange_Jun = Left(sRange_Jun, InStr(1, sRange_Jun, "$") - 1)
            eRange_Jun = .Cells(1, vCol - 1).Address(True, False, xlA1)
            eRange_Jun = Left(eRange_Jun, InStr(1, eRange_Jun, "$") - 1)
            .Range(sRange_Jun & CStr(vRow) & ":" & eRange_Jun & CStr(vRow)).Merge
            .Range(sRange_Jun & CStr(vRow) & ":" & eRange_Jun & CStr(vRow)).Interior.ColorIndex = 33
            .Range(sRange_Jun & CStr(vRow) & ":" & eRange_Jun & CStr(vRow)).Value = "June"
            '.Range(sRange_Jun & CStr(vRow) & ":" & eRange_Jun & CStr(vRow)).Locked = True
            
            Call incrProgressBar

            sCol = vCol
            While Trim(msf_MPInsertion.TextMatrix(vRow - xlTopRow - 5, vCol)) = "July"
                vCol = vCol + 1
            Wend
            sRange_Jul = .Cells(1, sCol).Address(True, False, xlA1)
            sRange_Jul = Left(sRange_Jul, InStr(1, sRange_Jul, "$") - 1)
            eRange_Jul = .Cells(1, vCol - 1).Address(True, False, xlA1)
            eRange_Jul = Left(eRange_Jul, InStr(1, eRange_Jul, "$") - 1)
            .Range(sRange_Jul & CStr(vRow) & ":" & eRange_Jul & CStr(vRow)).Merge
            .Range(sRange_Jul & CStr(vRow) & ":" & eRange_Jul & CStr(vRow)).Interior.ColorIndex = 34
            .Range(sRange_Jul & CStr(vRow) & ":" & eRange_Jul & CStr(vRow)).Value = "July"
            '.Range(sRange_Jul & CStr(vRow) & ":" & eRange_Jul & CStr(vRow)).Locked = True
            
            Call incrProgressBar

            sCol = vCol
            While Trim(msf_MPInsertion.TextMatrix(vRow - xlTopRow - 5, vCol)) = "August"
                vCol = vCol + 1
            Wend
            sRange_Aug = .Cells(1, sCol).Address(True, False, xlA1)
            sRange_Aug = Left(sRange_Aug, InStr(1, sRange_Aug, "$") - 1)
            eRange_Aug = .Cells(1, vCol - 1).Address(True, False, xlA1)
            eRange_Aug = Left(eRange_Aug, InStr(1, eRange_Aug, "$") - 1)
            .Range(sRange_Aug & CStr(vRow) & ":" & eRange_Aug & CStr(vRow)).Merge
            .Range(sRange_Aug & CStr(vRow) & ":" & eRange_Aug & CStr(vRow)).Interior.ColorIndex = 37
            .Range(sRange_Aug & CStr(vRow) & ":" & eRange_Aug & CStr(vRow)).Value = "August"
            '.Range(sRange_Aug & CStr(vRow) & ":" & eRange_Aug & CStr(vRow)).Locked = True
            
            Call incrProgressBar

            sCol = vCol
            While Trim(msf_MPInsertion.TextMatrix(vRow - xlTopRow - 5, vCol)) = "September"
                vCol = vCol + 1
            Wend
            sRange_Sep = .Cells(1, sCol).Address(True, False, xlA1)
            sRange_Sep = Left(sRange_Sep, InStr(1, sRange_Sep, "$") - 1)
            eRange_Sep = .Cells(1, vCol - 1).Address(True, False, xlA1)
            eRange_Sep = Left(eRange_Sep, InStr(1, eRange_Sep, "$") - 1)
            .Range(sRange_Sep & CStr(vRow) & ":" & eRange_Sep & CStr(vRow)).Merge
            .Range(sRange_Sep & CStr(vRow) & ":" & eRange_Sep & CStr(vRow)).Interior.ColorIndex = 33
            .Range(sRange_Sep & CStr(vRow) & ":" & eRange_Sep & CStr(vRow)).Value = "September"
            '.Range(sRange_Sep & CStr(vRow) & ":" & eRange_Sep & CStr(vRow)).Locked = True
            
            Call incrProgressBar

            sCol = vCol
            While Trim(msf_MPInsertion.TextMatrix(vRow - xlTopRow - 5, vCol)) = "October"
                vCol = vCol + 1
            Wend
            sRange_Oct = .Cells(1, sCol).Address(True, False, xlA1)
            sRange_Oct = Left(sRange_Oct, InStr(1, sRange_Oct, "$") - 1)
            eRange_Oct = .Cells(1, vCol - 1).Address(True, False, xlA1)
            eRange_Oct = Left(eRange_Oct, InStr(1, eRange_Oct, "$") - 1)
            .Range(sRange_Oct & CStr(vRow) & ":" & eRange_Oct & CStr(vRow)).Merge
            .Range(sRange_Oct & CStr(vRow) & ":" & eRange_Oct & CStr(vRow)).Interior.ColorIndex = 34
            .Range(sRange_Oct & CStr(vRow) & ":" & eRange_Oct & CStr(vRow)).Value = "October"
            '.Range(sRange_Oct & CStr(vRow) & ":" & eRange_Oct & CStr(vRow)).Locked = True
            
            Call incrProgressBar

            sCol = vCol
            While Trim(msf_MPInsertion.TextMatrix(vRow - xlTopRow - 5, vCol)) = "November"
                vCol = vCol + 1
            Wend
            sRange_Nov = .Cells(1, sCol).Address(True, False, xlA1)
            sRange_Nov = Left(sRange_Nov, InStr(1, sRange_Nov, "$") - 1)
            eRange_Nov = .Cells(1, vCol - 1).Address(True, False, xlA1)
            eRange_Nov = Left(eRange_Nov, InStr(1, eRange_Nov, "$") - 1)
            .Range(sRange_Nov & CStr(vRow) & ":" & eRange_Nov & CStr(vRow)).Merge
            .Range(sRange_Nov & CStr(vRow) & ":" & eRange_Nov & CStr(vRow)).Interior.ColorIndex = 37
            .Range(sRange_Nov & CStr(vRow) & ":" & eRange_Nov & CStr(vRow)).Value = "November"
            '.Range(sRange_Nov & CStr(vRow) & ":" & eRange_Nov & CStr(vRow)).Locked = True
            
            Call incrProgressBar

            sCol = vCol
            While Trim(msf_MPInsertion.TextMatrix(vRow - xlTopRow - 5, vCol)) = "December"
                vCol = vCol + 1
            Wend
            sRange_Dec = .Cells(1, sCol).Address(True, False, xlA1)
            sRange_Dec = Left(sRange_Dec, InStr(1, sRange_Dec, "$") - 1)
            eRange_Dec = .Cells(1, vCol - 1).Address(True, False, xlA1)
            eRange_Dec = Left(eRange_Dec, InStr(1, eRange_Dec, "$") - 1)
            .Range(sRange_Dec & CStr(vRow) & ":" & eRange_Dec & CStr(vRow)).Merge
            .Range(sRange_Dec & CStr(vRow) & ":" & eRange_Dec & CStr(vRow)).Interior.ColorIndex = 33
            .Range(sRange_Dec & CStr(vRow) & ":" & eRange_Dec & CStr(vRow)).Value = "December"
            '.Range(sRange_Dec & CStr(vRow) & ":" & eRange_Dec & CStr(vRow)).Locked = True
            
            Call incrProgressBar
            
            'GRP/ins
            .Cells(vRow, msf_MPInsertion.cols - 3 + (xlLeftCol - 1)).Interior.ColorIndex = 33
            .Cells(vRow, msf_MPInsertion.cols - 3 + (xlLeftCol - 1)).Font.Size = 9
            .Cells(vRow, msf_MPInsertion.cols - 3 + (xlLeftCol - 1)).Font.Bold = True
            
            
            .Cells(vRow, msf_MPInsertion.cols - 3 + (xlLeftCol - 1)).Borders(xlEdgeTop).LineStyle = xlDouble
            .Cells(vRow, msf_MPInsertion.cols - 3 + (xlLeftCol - 1)).Borders(xlEdgeTop).Weight = xlThick
            .Cells(vRow, msf_MPInsertion.cols - 3 + (xlLeftCol - 1)).Borders(xlEdgeTop).ColorIndex = 5
        
            .Cells(vRow, msf_MPInsertion.cols - 3 + (xlLeftCol - 1)).Borders(xlEdgeBottom).LineStyle = xlDouble
            .Cells(vRow, msf_MPInsertion.cols - 3 + (xlLeftCol - 1)).Borders(xlEdgeBottom).Weight = xlThick
            .Cells(vRow, msf_MPInsertion.cols - 3 + (xlLeftCol - 1)).Borders(xlEdgeBottom).ColorIndex = 5
            
            .Cells(vRow, msf_MPInsertion.cols - 3 + (xlLeftCol - 1)).HorizontalAlignment = xlCenter
            .Cells(vRow, msf_MPInsertion.cols - 3 + (xlLeftCol - 1)).Value = "GRP/ins"
            '.Cells(vRow, FGMPInsertion.Cols - 3 + (xlLeftCol - 1)).Locked = True
            
            Call incrProgressBar
        'Total
            .Cells(vRow, msf_MPInsertion.cols - 2 + (xlLeftCol - 1)).Font.Name = "Sabon MT"
            .Cells(vRow, msf_MPInsertion.cols - 2 + (xlLeftCol - 1)).Font.Size = 9
            .Cells(vRow, msf_MPInsertion.cols - 2 + (xlLeftCol - 1)).Font.Bold = True
            
            .Cells(vRow, msf_MPInsertion.cols - 2 + (xlLeftCol - 1)).Borders(xlEdgeTop).LineStyle = xlDouble
            .Cells(vRow, msf_MPInsertion.cols - 2 + (xlLeftCol - 1)).Borders(xlEdgeTop).Weight = xlThick
            .Cells(vRow, msf_MPInsertion.cols - 2 + (xlLeftCol - 1)).Borders(xlEdgeTop).ColorIndex = 5
        
            .Cells(vRow, msf_MPInsertion.cols - 2 + (xlLeftCol - 1)).Borders(xlEdgeBottom).LineStyle = xlDouble
            .Cells(vRow, msf_MPInsertion.cols - 2 + (xlLeftCol - 1)).Borders(xlEdgeBottom).Weight = xlThick
            .Cells(vRow, msf_MPInsertion.cols - 2 + (xlLeftCol - 1)).Borders(xlEdgeBottom).ColorIndex = 5
            
            .Cells(vRow, msf_MPInsertion.cols - 2 + (xlLeftCol - 1)).HorizontalAlignment = xlCenter
            .Cells(vRow, msf_MPInsertion.cols - 2 + (xlLeftCol - 1)).Value = "TOTAL"
            '.Cells(vRow, FGMPInsertion.Cols - 2 + (xlLeftCol - 1)).Locked = True
            
            Call incrProgressBar
            
        'date & week
            sCol = xlLeftCol + 5
            .Range(.Cells(xlTopRow + 7, sCol).Address & ":" & .Cells(xlTopRow + 7, msf_MPInsertion.cols - 4 + (xlLeftCol - 1)).Address).Font.Size = 9
            .Range(.Cells(xlTopRow + 7, sCol).Address & ":" & .Cells(xlTopRow + 7, msf_MPInsertion.cols - 4 + (xlLeftCol - 1)).Address).Font.Bold = True
            .Range(.Cells(xlTopRow + 7, sCol).Address & ":" & .Cells(xlTopRow + 7, msf_MPInsertion.cols - 4 + (xlLeftCol - 1)).Address).HorizontalAlignment = xlCenter
            .Range(.Cells(xlTopRow + 7, sCol).Address & ":" & .Cells(xlTopRow + 7, msf_MPInsertion.cols - 4 + (xlLeftCol - 1)).Address).Borders(xlEdgeBottom).LineStyle = xlContinuous
            .Range(.Cells(xlTopRow + 7, sCol).Address & ":" & .Cells(xlTopRow + 7, msf_MPInsertion.cols - 4 + (xlLeftCol - 1)).Address).Borders(xlEdgeBottom).Weight = xlMedium
            .Range(.Cells(xlTopRow + 7, sCol).Address & ":" & .Cells(xlTopRow + 7, msf_MPInsertion.cols - 4 + (xlLeftCol - 1)).Address).Borders(xlEdgeBottom).ColorIndex = 5
            .Range(.Cells(xlTopRow + 7, sCol).Address & ":" & .Cells(xlTopRow + 7, msf_MPInsertion.cols - 4 + (xlLeftCol - 1)).Address).Borders(xlInsideVertical).Weight = xlThin
            .Range(.Cells(xlTopRow + 7, sCol).Address & ":" & .Cells(xlTopRow + 7, msf_MPInsertion.cols - 4 + (xlLeftCol - 1)).Address).Borders(xlInsideVertical).LineStyle = xlContinuous
            .Range(.Cells(xlTopRow + 7, sCol).Address & ":" & .Cells(xlTopRow + 7, msf_MPInsertion.cols - 4 + (xlLeftCol - 1)).Address).Borders(xlInsideVertical).ColorIndex = 5
            
            .Range(.Cells(xlTopRow + 8, sCol).Address & ":" & .Cells(xlTopRow + 8, msf_MPInsertion.cols - 4 + (xlLeftCol - 1)).Address).Font.Size = 10
            .Range(.Cells(xlTopRow + 8, sCol).Address & ":" & .Cells(xlTopRow + 8, msf_MPInsertion.cols - 4 + (xlLeftCol - 1)).Address).Font.Name = "Sabon MT"
            .Range(.Cells(xlTopRow + 8, sCol).Address & ":" & .Cells(xlTopRow + 8, msf_MPInsertion.cols - 4 + (xlLeftCol - 1)).Address).HorizontalAlignment = xlCenter
            .Range(.Cells(xlTopRow + 8, sCol).Address & ":" & .Cells(xlTopRow + 8, msf_MPInsertion.cols - 4 + (xlLeftCol - 1)).Address).Borders(xlEdgeBottom).LineStyle = xlContinuous
            .Range(.Cells(xlTopRow + 8, sCol).Address & ":" & .Cells(xlTopRow + 8, msf_MPInsertion.cols - 4 + (xlLeftCol - 1)).Address).Borders(xlEdgeBottom).Weight = xlHairline
            .Range(.Cells(xlTopRow + 8, sCol).Address & ":" & .Cells(xlTopRow + 8, msf_MPInsertion.cols - 4 + (xlLeftCol - 1)).Address).Borders(xlEdgeBottom).ColorIndex = xlAutomatic
            .Range(.Cells(xlTopRow + 8, sCol).Address & ":" & .Cells(xlTopRow + 8, msf_MPInsertion.cols - 4 + (xlLeftCol - 1)).Address).Borders(xlInsideVertical).Weight = xlHairline
            .Range(.Cells(xlTopRow + 8, sCol).Address & ":" & .Cells(xlTopRow + 8, msf_MPInsertion.cols - 4 + (xlLeftCol - 1)).Address).Borders(xlInsideVertical).LineStyle = xlContinuous
            .Range(.Cells(xlTopRow + 8, sCol).Address & ":" & .Cells(xlTopRow + 8, msf_MPInsertion.cols - 4 + (xlLeftCol - 1)).Address).Borders(xlInsideVertical).ColorIndex = xlAutomatic
        
            For vCol = 6 To msf_MPInsertion.cols - 4
                .Cells(xlTopRow + 7, sCol).Value = msf_MPInsertion.TextMatrix(2, vCol)
                '.Cells(xlTopRow + 7, sCol).Locked = True
                .Cells(xlTopRow + 8, sCol).Value = sCol - xlLeftCol - 4
                '.Cells(xlTopRow + 8, sCol).Locked = True
                sCol = sCol + 1
                
                Call incrProgressBar
            Next
            
        'Month Font & Borders Setting
            
            .Range(sRange_Jan & CStr(vRow) & ":" & eRange_Dec & CStr(vRow)).Font.Size = 9
            .Range(sRange_Jan & CStr(vRow) & ":" & eRange_Dec & CStr(vRow)).Font.Bold = True
            .Range(sRange_Jan & CStr(vRow) & ":" & eRange_Dec & CStr(vRow)).HorizontalAlignment = xlCenter
            
            .Range(sRange_Jan & CStr(vRow) & ":" & eRange_Dec & CStr(vRow)).Borders(xlEdgeTop).LineStyle = xlDouble
            .Range(sRange_Jan & CStr(vRow) & ":" & eRange_Dec & CStr(vRow)).Borders(xlEdgeTop).Weight = xlThick
            .Range(sRange_Jan & CStr(vRow) & ":" & eRange_Dec & CStr(vRow)).Borders(xlEdgeTop).ColorIndex = 5
        
            .Range(sRange_Jan & CStr(vRow) & ":" & eRange_Dec & CStr(vRow)).Borders(xlEdgeBottom).LineStyle = xlDouble
            .Range(sRange_Jan & CStr(vRow) & ":" & eRange_Dec & CStr(vRow)).Borders(xlEdgeBottom).Weight = xlThick
            .Range(sRange_Jan & CStr(vRow) & ":" & eRange_Dec & CStr(vRow)).Borders(xlEdgeBottom).ColorIndex = 5
            
            Call incrProgressBar
                
        
            
        'marketing, media, channel, dimension ,etc
            
            .Range(.Cells(xlTopRow + 11, xlLeftCol).Address & ":" & .Cells(xlTopRow + 11, xlLeftCol + 4).Address).Interior.ColorIndex = 34
            
            .Range(.Cells(xlTopRow + 11, xlLeftCol).Address & ":" & .Cells(xlTopRow + 11, xlLeftCol + 4).Address).Borders(xlEdgeTop).LineStyle = xlContinuous
            .Range(.Cells(xlTopRow + 11, xlLeftCol).Address & ":" & .Cells(xlTopRow + 11, xlLeftCol + 4).Address).Borders(xlEdgeTop).Weight = xlHairline
                    
            .Range(.Cells(xlTopRow + 11, xlLeftCol).Address & ":" & .Cells(xlTopRow + 11, xlLeftCol + 4).Address).Borders(xlEdgeBottom).LineStyle = xlContinuous
            .Range(.Cells(xlTopRow + 11, xlLeftCol).Address & ":" & .Cells(xlTopRow + 11, xlLeftCol + 4).Address).Borders(xlEdgeBottom).Weight = xlHairline
            
            .Cells(xlTopRow + 11, xlLeftCol).Value = "Marketing Task(Activity/Variant)"
            .Cells(xlTopRow + 11, xlLeftCol + 1).Value = "Channel"
            .Cells(xlTopRow + 11, xlLeftCol + 2).Value = "Version"
            .Cells(xlTopRow + 11, xlLeftCol + 3).Value = "Media"
            .Cells(xlTopRow + 11, xlLeftCol + 4).Value = "Channel Dimension"
            
            Call incrProgressBar
            
        'Exporting data
            .Range(.Cells(xlTopRow + 12, 6 + (xlLeftCol - 1)).Address & ":" & .Cells(msf_MPInsertion.Rows + xlHeaderrows + xlSubHeaderRows + xlTopRow - OrigHeaderRows, msf_MPInsertion.cols - 2 + (xlLeftCol - 1)).Address).HorizontalAlignment = xlCenter
            intTask = 0
            xlRow = 5 + xlTopRow + 7
            For vRow = 5 To msf_MPInsertion.Rows - 1
                
                If Mid(msf_MPInsertion.TextMatrix(vRow, 0), 1, 9) = "Sub Total" Then
                    .Range(.Cells(xlRow, xlLeftCol + 1).Address & ":" & .Cells(xlRow, xlLeftCol + 4).Address).Merge
                    .Range(.Cells(xlRow, xlLeftCol + 1).Address & ":" & .Cells(xlRow, xlLeftCol + 4).Address).Interior.ColorIndex = 33
                    .Range(.Cells(xlRow, xlLeftCol + 1).Address & ":" & .Cells(xlRow, xlLeftCol + 4).Address).HorizontalAlignment = xlCenter
                    .Range(.Cells(xlRow, xlLeftCol + 1).Address & ":" & .Cells(xlRow, xlLeftCol + 4).Address).Value = msf_MPInsertion.TextMatrix(vRow, 0) & " - Task " & intTask
                    .Cells(xlRow, xlLeftCol).Font.ColorIndex = 2 'putih
                    .Cells(xlRow, xlLeftCol).Value = msf_MPInsertion.TextMatrix(vRow, 0) & " - Task " & intTask
                    
                    .Range(sRange_Jan & CStr(xlRow) & ":" & .Cells(xlRow, xlLeftCol + msf_MPInsertion.cols - 3).Address).Interior.ColorIndex = 6
                    .Range(sRange_Jan & CStr(xlRow) & ":" & .Cells(xlRow, xlLeftCol + msf_MPInsertion.cols - 3).Address).Borders(xlEdgeTop).LineStyle = xlContinuous
                    .Range(sRange_Jan & CStr(xlRow) & ":" & .Cells(xlRow, xlLeftCol + msf_MPInsertion.cols - 3).Address).Borders(xlEdgeTop).Weight = xlThin
                    .Range(sRange_Jan & CStr(xlRow) & ":" & .Cells(xlRow, xlLeftCol + msf_MPInsertion.cols - 3).Address).Borders(xlEdgeBottom).LineStyle = xlContinuous
                    .Range(sRange_Jan & CStr(xlRow) & ":" & .Cells(xlRow, xlLeftCol + msf_MPInsertion.cols - 3).Address).Borders(xlEdgeBottom).Weight = xlThin
                    .Range(sRange_Jan & CStr(xlRow) & ":" & .Cells(xlRow, xlLeftCol + msf_MPInsertion.cols - 3).Address).NumberFormat = "_(#,#0.00_);_(#,#0.00_);_(-_)"
                    .Range(sRange_Jan & CStr(xlRow) & ":" & .Cells(xlRow, xlLeftCol + msf_MPInsertion.cols - 3).Address).HorizontalAlignment = xlCenter
                    .Range(sRange_Jan & CStr(xlRow) & ":" & .Cells(xlRow, xlLeftCol + msf_MPInsertion.cols - 3).Address).Font.Name = "Arial"
                    .Range(sRange_Jan & CStr(xlRow) & ":" & .Cells(xlRow, xlLeftCol + msf_MPInsertion.cols - 3).Address).Font.Size = 10
                    
                    .Range(sRange_Jan & CStr(xlRow) & ":" & eRange_Jan & CStr(xlRow)).Merge
                    .Range(sRange_Feb & CStr(xlRow) & ":" & eRange_Feb & CStr(xlRow)).Merge
                    .Range(sRange_Mar & CStr(xlRow) & ":" & eRange_Mar & CStr(xlRow)).Merge
                    .Range(sRange_Apr & CStr(xlRow) & ":" & eRange_Apr & CStr(xlRow)).Merge
                    .Range(sRange_May & CStr(xlRow) & ":" & eRange_May & CStr(xlRow)).Merge
                    .Range(sRange_Jun & CStr(xlRow) & ":" & eRange_Jun & CStr(xlRow)).Merge
                    .Range(sRange_Jul & CStr(xlRow) & ":" & eRange_Jul & CStr(xlRow)).Merge
                    .Range(sRange_Aug & CStr(xlRow) & ":" & eRange_Aug & CStr(xlRow)).Merge
                    .Range(sRange_Sep & CStr(xlRow) & ":" & eRange_Sep & CStr(xlRow)).Merge
                    .Range(sRange_Oct & CStr(xlRow) & ":" & eRange_Oct & CStr(xlRow)).Merge
                    .Range(sRange_Nov & CStr(xlRow) & ":" & eRange_Nov & CStr(xlRow)).Merge
                    .Range(sRange_Dec & CStr(xlRow) & ":" & eRange_Dec & CStr(xlRow)).Merge
                    
                Else
                    If Trim(msf_MPInsertion.TextMatrix(vRow, 0)) = "Task" Then
                        If intTask > 0 Then
                            xlRow = xlRow - 1
                            'Total By Task
                                Call TotalByTask(xlws, xlRow, xlLeftCol, xlStartRowTask, intTask, "")
                                xlRow = xlRow + 1
                                Call TotalByTask(xlws, xlRow, xlLeftCol, xlStartRowTask, intTask, "Actual")
                        End If
                        intTask = intTask + 1
                        xlStartRowTask = xlRow
                        .Cells(xlRow, xlLeftCol).Value = Trim(msf_MPInsertion.TextMatrix(vRow, 0)) & " " & intTask & " : " & Trim(msf_MPInsertion.TextMatrix(vRow, 1))
                    Else
                        If Trim(msf_MPInsertion.TextMatrix(vRow, 0)) <> "" Then
                            .Cells(xlRow, xlLeftCol).Value = Trim(msf_MPInsertion.TextMatrix(vRow, 0)) & " : " & Trim(msf_MPInsertion.TextMatrix(vRow, 1))
                        End If
                    End If
                End If

                For vCol = 2 To msf_MPInsertion.cols - 2
                    If Mid(msf_MPInsertion.TextMatrix(vRow, 0), 1, 9) = "Sub Total" Then
                        If vCol < 6 Then
                            vCol = 6
                            ProgressBarExport.Value = ProgressBarExport.Value + 4 'skip
                        End If
                    End If
                    
                    If Trim(msf_MPInsertion.TextMatrix(vRow, vCol)) <> "" Then
                        If vCol > 1 And vCol < 6 Then 'buat rata kiri, rata atas
                            .Cells(xlRow, vCol + (xlLeftCol - 1)).HorizontalAlignment = xlLeft
                            .Cells(xlRow, vCol + (xlLeftCol - 1)).VerticalAlignment = xlTop
                            .Cells(xlRow, vCol + (xlLeftCol - 1)).WrapText = True
                        End If
                        .Cells(xlRow, vCol + (xlLeftCol - 1)).Value = Trim(msf_MPInsertion.TextMatrix(vRow, vCol))
                        
                        If Mid(msf_MPInsertion.TextMatrix(vRow, 0), 1, 9) <> "Sub Total" Then 'untuk sub total udah di lock dan udah merge, jd gak bisa di lock lagi
                            '.Cells(xlRow, vCol + (xlLeftCol - 1)).Locked = True
                        End If
                        
                        If vCol = 2 Then
                            .Cells(xlRow, vCol + (xlLeftCol - 1)).Font.Bold = True
                            '.Cells(xlRow, vCol + (xlLeftCol - 1)).Locked = True
                            Select Case Trim(msf_MPInsertion.TextMatrix(vRow, vCol))
                                Case "Reguler"
                                    .Cells(xlRow, vCol + (xlLeftCol - 1)).Font.Bold = False
                                Case "Sponsorship/Program"
                                    .Cells(xlRow, vCol + (xlLeftCol - 1)).Font.Bold = False
                                    .Cells(xlRow + 1, vCol + (xlLeftCol - 1)).Font.Bold = False
                                    .Cells(xlRow, vCol + (xlLeftCol - 1)).Value = "Sponsorship/"
                                    .Cells(xlRow + 1, vCol + (xlLeftCol - 1)).Value = "Program"
                            End Select
                        End If
                    End If
                    Call incrProgressBar
                Next
                If intTVRow = 2 Then
                    'Merge TV Station
                        .Cells(xlRow, xlLeftCol + 3).Value = ""
                        .Cells(xlRow - 1, xlLeftCol + 3).Value = ""
                        .Range(.Cells(xlRow - 2, xlLeftCol + 3).Address & ":" & .Cells(xlRow, xlLeftCol + 3).Address).Merge
                    'Merge Dimension
                        .Cells(xlRow, xlLeftCol + 4).Value = ""
                        .Cells(xlRow - 1, xlLeftCol + 4).Value = ""
                        .Range(.Cells(xlRow - 2, xlLeftCol + 4).Address & ":" & .Cells(xlRow, xlLeftCol + 4).Address).Merge
                        intTVRow = 0
                End If
                If Right(msf_MPInsertion.TextMatrix(vRow, msf_MPInsertion.cols - 1), 5) = "REACH" Then
                    'Merging TV Reach & FRequency
                        If strViewMonth = "ALL" Then
                            rsTemp.Open "select week_year_start,week_year_end from mp_tv_reach_frequency where mp_plan_dim_id='" & Left(msf_MPInsertion.TextMatrix(vRow, msf_MPInsertion.cols - 1), 19) & "' order by week_year_start", ConnERP, 1, 3
                        Else
                            rsTemp.Open "select week_year_start,week_year_end from mp_tv_reach_frequency where mp_plan_dim_id='" & Left(msf_MPInsertion.TextMatrix(vRow, msf_MPInsertion.cols - 1), 19) & "' and (month_start in (" & strViewMonth & ")  or month_end in (" & strViewMonth & "))order by week_year_start", ConnERP, 1, 3
                        End If
                        
                        While Not rsTemp.EOF
                            For intTVCol = rsTemp(0) + 1 To rsTemp(1)
                                .Cells(xlRow - 1, intTVCol + 5 + xlLeftCol - 1).Value = ""
                                .Cells(xlRow, intTVCol + 5 + xlLeftCol - 1).Value = ""
                            Next
                            .Range(.Cells(xlRow - 1, rsTemp(0) + 5 + xlLeftCol - 1).Address & ":" & .Cells(xlRow - 1, rsTemp(1) + 5 + xlLeftCol - 1).Address).Merge
                            .Range(.Cells(xlRow - 1, rsTemp(0) + 5 + xlLeftCol - 1).Address & ":" & .Cells(xlRow - 1, rsTemp(1) + 5 + xlLeftCol - 1).Address).Borders.LineStyle = xlContinuous
                            .Range(.Cells(xlRow - 1, rsTemp(0) + 5 + xlLeftCol - 1).Address & ":" & .Cells(xlRow - 1, rsTemp(1) + 5 + xlLeftCol - 1).Address).Borders.Weight = xlThin
                            .Range(.Cells(xlRow - 1, rsTemp(0) + 5 + xlLeftCol - 1).Address & ":" & .Cells(xlRow - 1, rsTemp(1) + 5 + xlLeftCol - 1).Address).Borders.ColorIndex = 5
                            .Range(.Cells(xlRow - 1, rsTemp(0) + 5 + xlLeftCol - 1).Address & ":" & .Cells(xlRow - 1, rsTemp(1) + 5 + xlLeftCol - 1).Address).Interior.ColorIndex = 34
                            .Range(.Cells(xlRow, rsTemp(0) + 5 + xlLeftCol - 1).Address & ":" & .Cells(xlRow, rsTemp(1) + 5 + xlLeftCol - 1).Address).Merge
                            .Range(.Cells(xlRow, rsTemp(0) + 5 + xlLeftCol - 1).Address & ":" & .Cells(xlRow, rsTemp(1) + 5 + xlLeftCol - 1).Address).Borders.LineStyle = xlContinuous
                            .Range(.Cells(xlRow, rsTemp(0) + 5 + xlLeftCol - 1).Address & ":" & .Cells(xlRow, rsTemp(1) + 5 + xlLeftCol - 1).Address).Borders.Weight = xlThin
                            .Range(.Cells(xlRow, rsTemp(0) + 5 + xlLeftCol - 1).Address & ":" & .Cells(xlRow, rsTemp(1) + 5 + xlLeftCol - 1).Address).Borders.ColorIndex = 5
                            .Range(.Cells(xlRow, rsTemp(0) + 5 + xlLeftCol - 1).Address & ":" & .Cells(xlRow, rsTemp(1) + 5 + xlLeftCol - 1).Address).Interior.ColorIndex = 34
                            rsTemp.MoveNext
                        Wend
                        rsTemp.Close
                        intTVRow = 2
                End If
            xlRow = xlRow + 1
            
            Next
        'Total By Task
            If intTask = 1 Then
                xlRow = xlRow - 1
            End If
            Call TotalByTask(xlws, xlRow, xlLeftCol, xlStartRowTask, intTask, "")
                xlRow = xlRow + 1
            Call TotalByTask(xlws, xlRow, xlLeftCol, xlStartRowTask, intTask, "Actual")
        'Grand Total
            xlRow = xlRow - 1
        Call GrandTotal(xlws, xlRow, xlLeftCol, xlTopRow, "Plan")
            xlRow = xlRow + 8
        Call GrandTotal(xlws, xlRow, xlLeftCol, xlTopRow, "Actual")
            xlRow = xlRow + 8
        Call CalculateBalance(xlws, xlRow, xlLeftCol, xlTopRow)
        
        'SETTING BORDERS
        
        'Month Sparator Borders
            .Range(sRange_Jan & CStr(xlTopRow + 6) & ":" & eRange_Jan & CStr(xlRow + 7)).Borders(xlEdgeLeft).LineStyle = xlContinuous
            .Range(sRange_Jan & CStr(xlTopRow + 6) & ":" & eRange_Jan & CStr(xlRow + 7)).Borders(xlEdgeLeft).Weight = xlThick
            .Range(sRange_Jan & CStr(xlTopRow + 6) & ":" & eRange_Jan & CStr(xlRow + 7)).Borders(xlEdgeLeft).ColorIndex = 5
            .Range(sRange_Jan & CStr(xlTopRow + 6) & ":" & eRange_Jan & CStr(xlRow + 7)).Borders(xlEdgeRight).LineStyle = xlContinuous
            .Range(sRange_Jan & CStr(xlTopRow + 6) & ":" & eRange_Jan & CStr(xlRow + 7)).Borders(xlEdgeRight).Weight = xlThick
            .Range(sRange_Jan & CStr(xlTopRow + 6) & ":" & eRange_Jan & CStr(xlRow + 7)).Borders(xlEdgeRight).ColorIndex = 5
            
            .Range(sRange_Feb & CStr(xlTopRow + 6) & ":" & eRange_Feb & CStr(xlRow + 7)).Borders(xlEdgeRight).LineStyle = xlContinuous
            .Range(sRange_Feb & CStr(xlTopRow + 6) & ":" & eRange_Feb & CStr(xlRow + 7)).Borders(xlEdgeRight).Weight = xlThick
            .Range(sRange_Feb & CStr(xlTopRow + 6) & ":" & eRange_Feb & CStr(xlRow + 7)).Borders(xlEdgeRight).ColorIndex = 5
            
            .Range(sRange_Mar & CStr(xlTopRow + 6) & ":" & eRange_Mar & CStr(xlRow + 7)).Borders(xlEdgeRight).LineStyle = xlContinuous
            .Range(sRange_Mar & CStr(xlTopRow + 6) & ":" & eRange_Mar & CStr(xlRow + 7)).Borders(xlEdgeRight).Weight = xlThick
            .Range(sRange_Mar & CStr(xlTopRow + 6) & ":" & eRange_Mar & CStr(xlRow + 7)).Borders(xlEdgeRight).ColorIndex = 5
            
            .Range(sRange_Apr & CStr(xlTopRow + 6) & ":" & eRange_Apr & CStr(xlRow + 7)).Borders(xlEdgeRight).LineStyle = xlContinuous
            .Range(sRange_Apr & CStr(xlTopRow + 6) & ":" & eRange_Apr & CStr(xlRow + 7)).Borders(xlEdgeRight).Weight = xlThick
            .Range(sRange_Apr & CStr(xlTopRow + 6) & ":" & eRange_Apr & CStr(xlRow + 7)).Borders(xlEdgeRight).ColorIndex = 5
            
            .Range(sRange_May & CStr(xlTopRow + 6) & ":" & eRange_May & CStr(xlRow + 7)).Borders(xlEdgeRight).LineStyle = xlContinuous
            .Range(sRange_May & CStr(xlTopRow + 6) & ":" & eRange_May & CStr(xlRow + 7)).Borders(xlEdgeRight).Weight = xlThick
            .Range(sRange_May & CStr(xlTopRow + 6) & ":" & eRange_May & CStr(xlRow + 7)).Borders(xlEdgeRight).ColorIndex = 5
            
            .Range(sRange_Jun & CStr(xlTopRow + 6) & ":" & eRange_Jun & CStr(xlRow + 7)).Borders(xlEdgeRight).LineStyle = xlContinuous
            .Range(sRange_Jun & CStr(xlTopRow + 6) & ":" & eRange_Jun & CStr(xlRow + 7)).Borders(xlEdgeRight).Weight = xlThick
            .Range(sRange_Jun & CStr(xlTopRow + 6) & ":" & eRange_Jun & CStr(xlRow + 7)).Borders(xlEdgeRight).ColorIndex = 5
            
            .Range(sRange_Jul & CStr(xlTopRow + 6) & ":" & eRange_Jul & CStr(xlRow + 7)).Borders(xlEdgeRight).LineStyle = xlContinuous
            .Range(sRange_Jul & CStr(xlTopRow + 6) & ":" & eRange_Jul & CStr(xlRow + 7)).Borders(xlEdgeRight).Weight = xlThick
            .Range(sRange_Jul & CStr(xlTopRow + 6) & ":" & eRange_Jul & CStr(xlRow + 7)).Borders(xlEdgeRight).ColorIndex = 5
            
            .Range(sRange_Aug & CStr(xlTopRow + 6) & ":" & eRange_Aug & CStr(xlRow + 7)).Borders(xlEdgeRight).LineStyle = xlContinuous
            .Range(sRange_Aug & CStr(xlTopRow + 6) & ":" & eRange_Aug & CStr(xlRow + 7)).Borders(xlEdgeRight).Weight = xlThick
            .Range(sRange_Aug & CStr(xlTopRow + 6) & ":" & eRange_Aug & CStr(xlRow + 7)).Borders(xlEdgeRight).ColorIndex = 5
            
            .Range(sRange_Sep & CStr(xlTopRow + 6) & ":" & eRange_Sep & CStr(xlRow + 7)).Borders(xlEdgeRight).LineStyle = xlContinuous
            .Range(sRange_Sep & CStr(xlTopRow + 6) & ":" & eRange_Sep & CStr(xlRow + 7)).Borders(xlEdgeRight).Weight = xlThick
            .Range(sRange_Sep & CStr(xlTopRow + 6) & ":" & eRange_Sep & CStr(xlRow + 7)).Borders(xlEdgeRight).ColorIndex = 5
            
            .Range(sRange_Oct & CStr(xlTopRow + 6) & ":" & eRange_Oct & CStr(xlRow + 7)).Borders(xlEdgeRight).LineStyle = xlContinuous
            .Range(sRange_Oct & CStr(xlTopRow + 6) & ":" & eRange_Oct & CStr(xlRow + 7)).Borders(xlEdgeRight).Weight = xlThick
            .Range(sRange_Oct & CStr(xlTopRow + 6) & ":" & eRange_Oct & CStr(xlRow + 7)).Borders(xlEdgeRight).ColorIndex = 5
            
            .Range(sRange_Nov & CStr(xlTopRow + 6) & ":" & eRange_Nov & CStr(xlRow + 7)).Borders(xlEdgeRight).LineStyle = xlContinuous
            .Range(sRange_Nov & CStr(xlTopRow + 6) & ":" & eRange_Nov & CStr(xlRow + 7)).Borders(xlEdgeRight).Weight = xlThick
            .Range(sRange_Nov & CStr(xlTopRow + 6) & ":" & eRange_Nov & CStr(xlRow + 7)).Borders(xlEdgeRight).ColorIndex = 5
            
            .Range(sRange_Dec & CStr(xlTopRow + 6) & ":" & eRange_Dec & CStr(xlRow + 7)).Borders(xlEdgeRight).LineStyle = xlContinuous
            .Range(sRange_Dec & CStr(xlTopRow + 6) & ":" & eRange_Dec & CStr(xlRow + 7)).Borders(xlEdgeRight).Weight = xlThick
            .Range(sRange_Dec & CStr(xlTopRow + 6) & ":" & eRange_Dec & CStr(xlRow + 7)).Borders(xlEdgeRight).ColorIndex = 5
            
        'GRP / INS Border
            .Range(.Cells(xlTopRow + 6, msf_MPInsertion.cols - 3 + (xlLeftCol - 1)).Address & ":" & .Cells(xlRow + 6, msf_MPInsertion.cols - 3 + (xlLeftCol - 1)).Address).Borders(xlEdgeRight).LineStyle = xlContinuous
            .Range(.Cells(xlTopRow + 6, msf_MPInsertion.cols - 3 + (xlLeftCol - 1)).Address & ":" & .Cells(xlRow + 6, msf_MPInsertion.cols - 3 + (xlLeftCol - 1)).Address).Borders(xlEdgeRight).Weight = xlThick
            .Range(.Cells(xlTopRow + 6, msf_MPInsertion.cols - 3 + (xlLeftCol - 1)).Address & ":" & .Cells(xlRow + 6, msf_MPInsertion.cols - 3 + (xlLeftCol - 1)).Address).Borders(xlEdgeRight).ColorIndex = 5
            
        'Setting Outer Border
            
            .Range(.Cells(xlTopRow, xlLeftCol).Address & ":" & .Cells(xlRow + 8, msf_MPInsertion.cols - 2 + (xlLeftCol - 1)).Address).Borders(xlEdgeLeft).LineStyle = xlDouble
            .Range(.Cells(xlTopRow, xlLeftCol).Address & ":" & .Cells(xlRow + 8, msf_MPInsertion.cols - 2 + (xlLeftCol - 1)).Address).Borders(xlEdgeLeft).Weight = xlThick
            .Range(.Cells(xlTopRow, xlLeftCol).Address & ":" & .Cells(xlRow + 8, msf_MPInsertion.cols - 2 + (xlLeftCol - 1)).Address).Borders(xlEdgeLeft).ColorIndex = xlAutomatic
            
            .Range(.Cells(xlTopRow, xlLeftCol).Address & ":" & .Cells(xlRow + 8, msf_MPInsertion.cols - 2 + (xlLeftCol - 1)).Address).Borders(xlEdgeTop).LineStyle = xlDouble
            .Range(.Cells(xlTopRow, xlLeftCol).Address & ":" & .Cells(xlRow + 8, msf_MPInsertion.cols - 2 + (xlLeftCol - 1)).Address).Borders(xlEdgeTop).Weight = xlThick
            .Range(.Cells(xlTopRow, xlLeftCol).Address & ":" & .Cells(xlRow + 8, msf_MPInsertion.cols - 2 + (xlLeftCol - 1)).Address).Borders(xlEdgeTop).ColorIndex = xlAutomatic
            
            .Range(.Cells(xlTopRow, xlLeftCol).Address & ":" & .Cells(xlRow + 8, msf_MPInsertion.cols - 2 + (xlLeftCol - 1)).Address).Borders(xlEdgeRight).LineStyle = xlDouble
            .Range(.Cells(xlTopRow, xlLeftCol).Address & ":" & .Cells(xlRow + 8, msf_MPInsertion.cols - 2 + (xlLeftCol - 1)).Address).Borders(xlEdgeRight).Weight = xlThick
            .Range(.Cells(xlTopRow, xlLeftCol).Address & ":" & .Cells(xlRow + 8, msf_MPInsertion.cols - 2 + (xlLeftCol - 1)).Address).Borders(xlEdgeRight).ColorIndex = xlAutomatic
            
            .Range(.Cells(xlTopRow, xlLeftCol).Address & ":" & .Cells(xlRow + 8, msf_MPInsertion.cols - 2 + (xlLeftCol - 1)).Address).Borders(xlEdgeBottom).LineStyle = xlDouble
            .Range(.Cells(xlTopRow, xlLeftCol).Address & ":" & .Cells(xlRow + 8, msf_MPInsertion.cols - 2 + (xlLeftCol - 1)).Address).Borders(xlEdgeBottom).Weight = xlThick
            .Range(.Cells(xlTopRow, xlLeftCol).Address & ":" & .Cells(xlRow + 8, msf_MPInsertion.cols - 2 + (xlLeftCol - 1)).Address).Borders(xlEdgeBottom).ColorIndex = xlAutomatic
        
        'Hight light Total per quarter
            'Plan
            .Range(sRange_Jan & CStr(xlRow - 10) & ":" & eRange_Dec & CStr(xlRow - 9)).Interior.ColorIndex = 6
            .Range(sRange_Jan & CStr(xlRow - 10) & ":" & eRange_Dec & CStr(xlRow - 9)).Interior.Pattern = xlSolid
            'Actual
            .Range(sRange_Jan & CStr(xlRow - 2) & ":" & eRange_Dec & CStr(xlRow - 1)).Interior.ColorIndex = 6
            .Range(sRange_Jan & CStr(xlRow - 2) & ":" & eRange_Dec & CStr(xlRow - 1)).Interior.Pattern = xlSolid
            'Balance
            .Range(sRange_Jan & CStr(xlRow + 6) & ":" & eRange_Dec & CStr(xlRow + 7)).Interior.ColorIndex = 6
            .Range(sRange_Jan & CStr(xlRow + 6) & ":" & eRange_Dec & CStr(xlRow + 7)).Interior.Pattern = xlSolid
        
        Call incrProgressBar
        'page setup
            On Error GoTo skip_page_setup
            .PageSetup.LeftMargin = xlApp.InchesToPoints(0.25): Call incrProgressBar
            .PageSetup.RightMargin = xlApp.InchesToPoints(0.25): Call incrProgressBar
            .PageSetup.TopMargin = xlApp.InchesToPoints(0.25): Call incrProgressBar
            .PageSetup.BottomMargin = xlApp.InchesToPoints(0.25): Call incrProgressBar
            .PageSetup.HeaderMargin = xlApp.InchesToPoints(0.5): Call incrProgressBar
            .PageSetup.FooterMargin = xlApp.InchesToPoints(0.5): Call incrProgressBar
            .PageSetup.CenterHorizontally = True: Call incrProgressBar
            .PageSetup.Orientation = xlLandscape: Call incrProgressBar
            .PageSetup.PaperSize = xlPaperA4: Call incrProgressBar
            .PageSetup.Zoom = 40: Call incrProgressBar
skip_page_setup:
        'Protect Sheet
            .Protect "27 October", DrawingObjects:=False, Contents:=True, Scenarios:=True, AllowFormattingCells:=True, AllowFormattingColumns:=True, AllowFormattingRows:=True
        
        ProgressBarExport.Value = ProgressBarExport.Max
        lblPercent.Caption = "100% Complete..."
        lblPercent.Refresh
    End With
    
    If xlWB.Worksheets.Count < 2 Then
        Set xlws = xlWB.Worksheets.Add
    Else
        Set xlws = xlWB.Worksheets(2)
    End If
        
    xlws.Name = "TV Layering"
    Call ExportTVLayering(xlws)
    
    If xlWB.Worksheets.Count < 3 Then
        Set xlws = xlWB.Worksheets.Add
    Else
        Set xlws = xlWB.Worksheets(3)
    End If
        
    xlws.Name = "Summary"
    Call ExportSummary(xlws)
    
    pesan = MsgBox("Exporting Complete!", vbExclamation, strApplication_Name)
    
    FrameProgressBar.Visible = False
    xlApp.Visible = True
    
    Set xlApp = Nothing
    Set xlWB = Nothing
    Set xlws = Nothing
    
End Sub

Private Sub ExportSummary(xlws As Object)
    Dim strSql As String, rsTemp As New ADODB.Recordset, tot_rec As Integer
    Dim rsTask As New ADODB.Recordset
    Dim rsSummary As New ADODB.Recordset
    Dim i As Integer, Current_Task_Row As Integer, Task_Num As Integer, Total_Row_Pos As Integer
    Dim intStartRow As Single
    Dim intTotal_Row_Pos_Plan
    Dim intTotal_Row_Pos_Actual
    Dim intRow As Single, intCol As Single
    
    Dim BgColorMonth(12) As Integer
        BgColorMonth(0) = 40
        BgColorMonth(1) = 36
        BgColorMonth(2) = 35
        BgColorMonth(3) = 34
        BgColorMonth(4) = 40
        BgColorMonth(5) = 36
        BgColorMonth(6) = 35
        BgColorMonth(7) = 34
        BgColorMonth(8) = 40
        BgColorMonth(9) = 36
        BgColorMonth(10) = 35
        BgColorMonth(11) = 34
    
    'GET TOTAL ROW
    strSql = "select a.mp_task_id,c.medium_code from mp_task a"
    strSql = strSql & " inner join mp_activity b"
    strSql = strSql & " on a.mp_number = '" & cboMPNumber.Text & "' and a.mp_task_id = b.mp_task_id"
    strSql = strSql & " inner join mp_medium c"
    strSql = strSql & " on b.mp_activity_id = c.mp_activity_id"
    strSql = strSql & " group by a.mp_task_id,c.medium_code"
    rsTemp.CursorLocation = adUseClient
    rsTemp.Open strSql, ConnERP
        tot_rec = rsTemp.RecordCount
    rsTemp.Close
    ProgressBarExport.Max = tot_rec * 12
    ProgressBarExport.Min = 0
    
    
    'EXPORT BUDGET PLAN
    lblPleaseWait.Caption = "Exporting Summary... [PLAN]"
    lblPercent.Caption = ""
    lblPercent.Refresh
    ProgressBarExport.Value = 0
    intStartRow = 2
    With xlws
        'Print Header
        .Cells(intStartRow - 1, 1).Value = "PLAN"
        
        .Cells(intStartRow, 1).Value = "Medium"
        .Cells(intStartRow, 1).Interior.ColorIndex = 36
        For i = 1 To 12
            .Cells(intStartRow, i + 1).Value = EngMonthName(i)
            .Cells(intStartRow, i + 1).Interior.ColorIndex = BgColorMonth(i - 1)
        Next
        .Cells(intStartRow, 14).Value = "TOTAL"
        .Cells(intStartRow, 14).Interior.ColorIndex = 36
        .Cells(intStartRow, 15).Value = "   %   "
        .Cells(intStartRow, 15).Interior.ColorIndex = 36
        
        .Range("A" & CStr(intStartRow) & ":O" & CStr(intStartRow)).HorizontalAlignment = xlCenter
        .Range("A" & CStr(intStartRow) & ":O" & CStr(intStartRow)).Font.Bold = True
        
        Current_Task_Row = intStartRow + 1
        Task_Num = 1
        strSql = "select mp_task_id,task_desc from mp_task where mp_number = '" & cboMPNumber.Text & "'"
        rsTask.CursorLocation = adUseClient
        rsTask.Open strSql, ConnERP
        Total_Row_Pos = 7 * rsTask.RecordCount + 2 + intStartRow - 1
        intTotal_Row_Pos_Plan = Total_Row_Pos
        
        'Print footer
        .Cells(Total_Row_Pos, 1).Value = "TOTAL"
        .Cells(Total_Row_Pos, 1).Font.ColorIndex = 5
        .Cells(Total_Row_Pos + 1, 1).Value = "TV"
        .Cells(Total_Row_Pos + 2, 1).Value = "Radio"
        .Cells(Total_Row_Pos + 3, 1).Value = "Print"
        .Cells(Total_Row_Pos + 4, 1).Value = "Cinema"
        .Cells(Total_Row_Pos + 5, 1).Value = "Other"
        .Cells(Total_Row_Pos + 6, 1).Value = "Grand Total"
        .Range(.Cells(Total_Row_Pos + 6, 1).Address & ":" & .Cells(Total_Row_Pos + 6, 14).Address).Font.Bold = True
        
        While Not rsTask.EOF
            
            .Cells(Current_Task_Row, 1).Value = "TASK " & Task_Num & " : " & rsTask(1)
            .Cells(Current_Task_Row, 1).Font.ColorIndex = 5
            .Cells(Current_Task_Row + 1, 1).Value = "TV"
            .Cells(Current_Task_Row + 2, 1).Value = "Radio"
            .Cells(Current_Task_Row + 3, 1).Value = "Print"
            .Cells(Current_Task_Row + 4, 1).Value = "Cinema"
            .Cells(Current_Task_Row + 5, 1).Value = "Other"
            .Cells(Current_Task_Row + 6, 1).Value = "Sub Total"
            .Range(.Cells(Current_Task_Row + 6, 1).Address & ":" & .Cells(Current_Task_Row + 6, 14).Address).Font.Bold = True
            
            strSql = "select b.medium_code,c.month_number,sum(c.NETT_ONLY) as budget, sum(c.NETT_PLUS_FEE) as NETT_PLUS_FEE,sum(c.NETT_FEE_OTHER) as NETT_FEE_OTHER, max(c.IS_ACTUAL) IS_ACTUAL"
            'LAST POS
            
            strSql = strSql & " from mp_activity a inner join mp_medium b"
            strSql = strSql & " on a.mp_activity_id = b.mp_activity_id"
            strSql = strSql & " and a.mp_task_id = '" & rsTask(0) & "'"
            strSql = strSql & " inner join ("
             
             '==========new view
            strSql = strSql & " select mp_medium_id,month_number, Case IsNull(Total_Actual, -1) when -1 then min_budget "
            strSql = strSql & " Else Actual_Nett_Paid + Actual_Nett_Bonus end Nett_only,"
            
            strSql = strSql & " Case IsNull(Total_Actual, -1) when -1 then "
            strSql = strSql & " min_budget + MSC_Paid_Value + MSC_Bonus_Value + Club_Agency_Value "
            strSql = strSql & " Else Actual_Nett_Paid + Actual_Nett_Bonus + Actual_MSC_Paid + Actual_MSC_Bonus end Nett_plus_fee,"
            
            strSql = strSql & " Case IsNull(Total_Actual, -1) when -1 then "
            strSql = strSql & " min_budget + MSC_Paid_Value + MSC_Bonus_Value + Club_Agency_Value + isnull(other_cost,0) "
            strSql = strSql & " Else Actual_Nett_Paid + Actual_Nett_Bonus + Actual_MSC_Paid + Actual_MSC_Bonus + isnull(Actual_other_cost,0) end Nett_fee_other,"
            
            strSql = strSql & " case isnull(total_actual,-1) when -1 then 0 else 1 end is_actual From mp_monthly_activity "
            '============
            'strSQL = strSQL & " mp_monthly_activity "
            
            strSql = strSql & " ) c on b.mp_medium_id = c.mp_medium_id"
            strSql = strSql & " group by b.medium_code,c.month_number"
            rsSummary.Open strSql, ConnERP, 1, 3
            While Not rsSummary.EOF
                Select Case rsSummary("medium_code")
                    Case "TV"
                        'month
                        .Cells(Current_Task_Row + 1, rsSummary(1) + 1).Value = rsSummary("NETT_FEE_OTHER")
                        If rsSummary("is_actual").Value = 1 Then
                            .Cells(Current_Task_Row + 1, rsSummary(1) + 1).Interior.ColorIndex = 35
                            .Cells(Current_Task_Row + 1, rsSummary(1) + 1).Interior.Pattern = xlSolid
                        End If
                        'year
                        .Cells(Current_Task_Row + 1, 14).Value = .Cells(Current_Task_Row + 1, 14).Value + rsSummary("NETT_FEE_OTHER")
                        'total month
                        .Cells(Total_Row_Pos + 1, rsSummary(1) + 1).Value = .Cells(Total_Row_Pos + 1, rsSummary(1) + 1).Value + rsSummary("NETT_FEE_OTHER")
                        'total year
                        .Cells(Total_Row_Pos + 1, 14).Value = .Cells(Total_Row_Pos + 1, 14).Value + rsSummary("NETT_FEE_OTHER")
                    Case "RD"
                        'month
                        .Cells(Current_Task_Row + 2, rsSummary(1) + 1).Value = rsSummary("NETT_FEE_OTHER")
                        If rsSummary("is_actual").Value = 1 Then
                            .Cells(Current_Task_Row + 2, rsSummary(1) + 1).Interior.ColorIndex = 35
                            .Cells(Current_Task_Row + 2, rsSummary(1) + 1).Interior.Pattern = xlSolid
                        End If
                        'year
                        .Cells(Current_Task_Row + 2, 14).Value = .Cells(Current_Task_Row + 2, 14).Value + rsSummary("NETT_FEE_OTHER")
                        'total month
                        .Cells(Total_Row_Pos + 2, rsSummary(1) + 1).Value = .Cells(Total_Row_Pos + 2, rsSummary(1) + 1).Value + rsSummary("NETT_FEE_OTHER")
                        'total year
                        .Cells(Total_Row_Pos + 2, 14).Value = .Cells(Total_Row_Pos + 2, 14).Value + rsSummary("NETT_FEE_OTHER")
                    Case "PR"
                        'month
                        .Cells(Current_Task_Row + 3, rsSummary(1) + 1).Value = rsSummary("NETT_FEE_OTHER")
                        If rsSummary("is_actual").Value = 1 Then
                            .Cells(Current_Task_Row + 3, rsSummary(1) + 1).Interior.ColorIndex = 35
                            .Cells(Current_Task_Row + 3, rsSummary(1) + 1).Interior.Pattern = xlSolid
                        End If
                        'year
                        .Cells(Current_Task_Row + 3, 14).Value = .Cells(Current_Task_Row + 3, 14).Value + rsSummary("NETT_FEE_OTHER")
                        'total month
                        .Cells(Total_Row_Pos + 3, rsSummary(1) + 1).Value = .Cells(Total_Row_Pos + 3, rsSummary(1) + 1).Value + rsSummary("NETT_FEE_OTHER")
                        'total year
                        .Cells(Total_Row_Pos + 3, 14).Value = .Cells(Total_Row_Pos + 3, 14).Value + rsSummary("NETT_FEE_OTHER")
                    Case "CN"
                        'month
                        .Cells(Current_Task_Row + 4, rsSummary(1) + 1).Value = rsSummary("NETT_FEE_OTHER")
                        If rsSummary("is_actual").Value = 1 Then
                            .Cells(Current_Task_Row + 4, rsSummary(1) + 1).Interior.ColorIndex = 35
                            .Cells(Current_Task_Row + 4, rsSummary(1) + 1).Interior.Pattern = xlSolid
                        End If
                        'year
                        .Cells(Current_Task_Row + 4, 14).Value = .Cells(Current_Task_Row + 4, 14).Value + rsSummary("NETT_FEE_OTHER")
                        'total month
                        .Cells(Total_Row_Pos + 4, rsSummary(1) + 1).Value = .Cells(Total_Row_Pos + 4, rsSummary(1) + 1).Value + rsSummary("NETT_FEE_OTHER")
                        'total year
                        .Cells(Total_Row_Pos + 4, 14).Value = .Cells(Total_Row_Pos + 4, 14).Value + rsSummary("NETT_FEE_OTHER")
                    Case "OT"
                        'month
                        .Cells(Current_Task_Row + 5, rsSummary(1) + 1).Value = rsSummary("NETT_FEE_OTHER")
                        If rsSummary("is_actual").Value = 1 Then
                            .Cells(Current_Task_Row + 5, rsSummary(1) + 1).Interior.ColorIndex = 35
                            .Cells(Current_Task_Row + 5, rsSummary(1) + 1).Interior.Pattern = xlSolid
                        End If
                        'year
                        .Cells(Current_Task_Row + 5, 14).Value = .Cells(Current_Task_Row + 5, 14).Value + rsSummary("NETT_FEE_OTHER")
                        'total month
                        .Cells(Total_Row_Pos + 5, rsSummary(1) + 1).Value = .Cells(Total_Row_Pos + 5, rsSummary(1) + 1).Value + rsSummary("NETT_FEE_OTHER")
                        'total year
                        .Cells(Total_Row_Pos + 5, 14).Value = .Cells(Total_Row_Pos + 5, 14).Value + rsSummary("NETT_FEE_OTHER")
                End Select
                'Sub Total month
                .Cells(Current_Task_Row + 6, rsSummary(1) + 1).Value = .Cells(Current_Task_Row + 6, rsSummary(1) + 1).Value + rsSummary("NETT_FEE_OTHER")
                'Sub Total year
                .Cells(Current_Task_Row + 6, 14).Value = .Cells(Current_Task_Row + 6, 14).Value + rsSummary("NETT_FEE_OTHER")
                'Grand Total Month
                .Cells(Total_Row_Pos + 6, rsSummary(1) + 1).Value = .Cells(Total_Row_Pos + 6, rsSummary(1) + 1).Value + rsSummary("NETT_FEE_OTHER")
                'Grand Total Year
                .Cells(Total_Row_Pos + 6, 14).Value = .Cells(Total_Row_Pos + 6, 14).Value + rsSummary("NETT_FEE_OTHER")
                
                rsSummary.MoveNext
                Call incrProgressBar
            Wend
            rsSummary.Close
            Current_Task_Row = Current_Task_Row + 7
            Task_Num = Task_Num + 1
            rsTask.MoveNext
        Wend
        rsTask.Close
        
        .Cells.Style = "Comma"
        .Cells.Font.Name = "Tahoma"
        .Cells.Font.Size = 8
        
        'inner borders
        With .Range(.Cells(intStartRow, 1).Address & ":" & .Cells(Total_Row_Pos + 6, 15).Address).Borders
            .LineStyle = xlContinuous
            .Weight = xlThin
            .ColorIndex = 40
        End With
        
        'blok warna kolom percent
        .Range("O" & CStr(intStartRow) & ":" & .Cells(Total_Row_Pos + 6, 15).Address).Interior.ColorIndex = 36
        .Range("O" & CStr(intStartRow) & ":" & .Cells(Total_Row_Pos + 6, 15).Address).NumberFormat = "#0.00%"
        
        'set formula kolom %
        '----------------------- Cells.FormulaR1C1 Example ---------------------------
        'R = curent row, R[x] = current row + x, R[-x] = curent row - x. C = Column
        'contoh :   RC = current row, current col
        '        RC[1] = current row, current col+1
        '       R[-1]C = current row - 1, curent col
        '-----------------------------------------------------------------------------
        .Cells(Total_Row_Pos + 1, 15).FormulaR1C1 = "=RC[-1]/R[5]C[-1]"
        .Cells(Total_Row_Pos + 2, 15).FormulaR1C1 = "=RC[-1]/R[4]C[-1]"
        .Cells(Total_Row_Pos + 3, 15).FormulaR1C1 = "=RC[-1]/R[3]C[-1]"
        .Cells(Total_Row_Pos + 4, 15).FormulaR1C1 = "=RC[-1]/R[2]C[-1]"
        .Cells(Total_Row_Pos + 5, 15).FormulaR1C1 = "=RC[-1]/R[1]C[-1]"
        
        'draw outer border
        With .Range(.Cells(intStartRow, 1).Address & ":" & .Cells(Total_Row_Pos + 6, 15).Address).Borders(xlEdgeLeft)
            .LineStyle = xlContinuous
            .Weight = xlThick
            .ColorIndex = 5
        End With
        With .Range(.Cells(intStartRow, 1).Address & ":" & .Cells(Total_Row_Pos + 6, 15).Address).Borders(xlEdgeRight)
            .LineStyle = xlContinuous
            .Weight = xlThick
            .ColorIndex = 5
        End With
        With .Range(.Cells(intStartRow, 1).Address & ":" & .Cells(Total_Row_Pos + 6, 15).Address).Borders(xlEdgeTop)
            .LineStyle = xlContinuous
            .Weight = xlThick
            .ColorIndex = 5
        End With
        With .Range(.Cells(intStartRow, 1).Address & ":" & .Cells(Total_Row_Pos + 6, 15).Address).Borders(xlEdgeBottom)
            .LineStyle = xlContinuous
            .Weight = xlThick
            .ColorIndex = 5
        End With
        
        '.Cells.EntireColumn.AutoFit
        '.Cells.EntireRow.AutoFit
        
        '.Activate
        '.Cells(2, 2).Select
        'ActiveWindow.FreezePanes = True
    End With
    
'===============================================================================================================

    'EXPORT BUDGET ACTUAL
    lblPleaseWait.Caption = "Exporting Summary... [ACTUAL]"
    lblPercent.Caption = ""
    lblPercent.Refresh
    ProgressBarExport.Value = 0
    intStartRow = Total_Row_Pos + 9
    With xlws
        'Print Header
        .Cells(intStartRow - 1, 1).Value = "ACTUAL"
        
        .Cells(intStartRow, 1).Value = "Medium"
        .Cells(intStartRow, 1).Interior.ColorIndex = 36
        For i = 1 To 12
            .Cells(intStartRow, i + 1).Value = EngMonthName(i)
            .Cells(intStartRow, i + 1).Interior.ColorIndex = BgColorMonth(i - 1)
        Next
        .Cells(intStartRow, 14).Value = "TOTAL"
        .Cells(intStartRow, 14).Interior.ColorIndex = 36
        .Cells(intStartRow, 15).Value = "   %   "
        .Cells(intStartRow, 15).Interior.ColorIndex = 36
        
        .Range("A" & CStr(intStartRow) & ":O" & CStr(intStartRow)).HorizontalAlignment = xlCenter
        .Range("A" & CStr(intStartRow) & ":O" & CStr(intStartRow)).Font.Bold = True
        
        Current_Task_Row = intStartRow + 1
        Task_Num = 1
        strSql = "select mp_task_id,task_desc from mp_task where mp_number = '" & cboMPNumber.Text & "'"
        rsTask.CursorLocation = adUseClient
        rsTask.Open strSql, ConnERP
        Total_Row_Pos = 7 * rsTask.RecordCount + 2 + intStartRow - 1
        intTotal_Row_Pos_Actual = Total_Row_Pos
        
        'Print footer
        .Cells(Total_Row_Pos, 1).Value = "TOTAL"
        .Cells(Total_Row_Pos, 1).Font.ColorIndex = 5
        .Cells(Total_Row_Pos + 1, 1).Value = "TV"
        .Cells(Total_Row_Pos + 2, 1).Value = "Radio"
        .Cells(Total_Row_Pos + 3, 1).Value = "Print"
        .Cells(Total_Row_Pos + 4, 1).Value = "Cinema"
        .Cells(Total_Row_Pos + 5, 1).Value = "Other"
        .Cells(Total_Row_Pos + 6, 1).Value = "Grand Total"
        .Range(.Cells(Total_Row_Pos + 6, 1).Address & ":" & .Cells(Total_Row_Pos + 6, 14).Address).Font.Bold = True
        
        While Not rsTask.EOF
            
            .Cells(Current_Task_Row, 1).Value = "TASK " & Task_Num & " : " & rsTask(1)
            .Cells(Current_Task_Row, 1).Font.ColorIndex = 5
            .Cells(Current_Task_Row + 1, 1).Value = "TV"
            .Cells(Current_Task_Row + 2, 1).Value = "Radio"
            .Cells(Current_Task_Row + 3, 1).Value = "Print"
            .Cells(Current_Task_Row + 4, 1).Value = "Cinema"
            .Cells(Current_Task_Row + 5, 1).Value = "Other"
            .Cells(Current_Task_Row + 6, 1).Value = "Sub Total"
            .Range(.Cells(Current_Task_Row + 6, 1).Address & ":" & .Cells(Current_Task_Row + 6, 14).Address).Font.Bold = True
            
            strSql = "select b.medium_code,c.month_number,sum(isnull(c.total_actual,0)) as nett_fee_other "
            strSql = strSql & " from mp_activity a inner join mp_medium b"
            strSql = strSql & " on a.mp_activity_id = b.mp_activity_id"
            strSql = strSql & " and a.mp_task_id = '" & rsTask(0) & "'"
            strSql = strSql & " inner join mp_monthly_activity c on b.mp_medium_id = c.mp_medium_id"
            strSql = strSql & " group by b.medium_code,c.month_number"
            rsSummary.Open strSql, ConnERP, 1, 3
            
            While Not rsSummary.EOF
                Select Case rsSummary("medium_code")
                    Case "TV"
                        'month
                        .Cells(Current_Task_Row + 1, rsSummary(1) + 1).Value = rsSummary("NETT_FEE_OTHER")
                       
                        'year
                        .Cells(Current_Task_Row + 1, 14).Value = .Cells(Current_Task_Row + 1, 14).Value + rsSummary("NETT_FEE_OTHER")
                        'total month
                        .Cells(Total_Row_Pos + 1, rsSummary(1) + 1).Value = .Cells(Total_Row_Pos + 1, rsSummary(1) + 1).Value + rsSummary("NETT_FEE_OTHER")
                        'total year
                        .Cells(Total_Row_Pos + 1, 14).Value = .Cells(Total_Row_Pos + 1, 14).Value + rsSummary("NETT_FEE_OTHER")
                    Case "RD"
                        'month
                        .Cells(Current_Task_Row + 2, rsSummary(1) + 1).Value = rsSummary("NETT_FEE_OTHER")
                        
                        'year
                        .Cells(Current_Task_Row + 2, 14).Value = .Cells(Current_Task_Row + 2, 14).Value + rsSummary("NETT_FEE_OTHER")
                        'total month
                        .Cells(Total_Row_Pos + 2, rsSummary(1) + 1).Value = .Cells(Total_Row_Pos + 2, rsSummary(1) + 1).Value + rsSummary("NETT_FEE_OTHER")
                        'total year
                        .Cells(Total_Row_Pos + 2, 14).Value = .Cells(Total_Row_Pos + 2, 14).Value + rsSummary("NETT_FEE_OTHER")
                    Case "PR"
                        'month
                        .Cells(Current_Task_Row + 3, rsSummary(1) + 1).Value = rsSummary("NETT_FEE_OTHER")
                        
                        'year
                        .Cells(Current_Task_Row + 3, 14).Value = .Cells(Current_Task_Row + 3, 14).Value + rsSummary("NETT_FEE_OTHER")
                        'total month
                        .Cells(Total_Row_Pos + 3, rsSummary(1) + 1).Value = .Cells(Total_Row_Pos + 3, rsSummary(1) + 1).Value + rsSummary("NETT_FEE_OTHER")
                        'total year
                        .Cells(Total_Row_Pos + 3, 14).Value = .Cells(Total_Row_Pos + 3, 14).Value + rsSummary("NETT_FEE_OTHER")
                    Case "CN"
                        'month
                        .Cells(Current_Task_Row + 4, rsSummary(1) + 1).Value = rsSummary("NETT_FEE_OTHER")
                        
                        'year
                        .Cells(Current_Task_Row + 4, 14).Value = .Cells(Current_Task_Row + 4, 14).Value + rsSummary("NETT_FEE_OTHER")
                        'total month
                        .Cells(Total_Row_Pos + 4, rsSummary(1) + 1).Value = .Cells(Total_Row_Pos + 4, rsSummary(1) + 1).Value + rsSummary("NETT_FEE_OTHER")
                        'total year
                        .Cells(Total_Row_Pos + 4, 14).Value = .Cells(Total_Row_Pos + 4, 14).Value + rsSummary("NETT_FEE_OTHER")
                    Case "OT"
                        'month
                        .Cells(Current_Task_Row + 5, rsSummary(1) + 1).Value = rsSummary("NETT_FEE_OTHER")
                        
                        'year
                        .Cells(Current_Task_Row + 5, 14).Value = .Cells(Current_Task_Row + 5, 14).Value + rsSummary("NETT_FEE_OTHER")
                        'total month
                        .Cells(Total_Row_Pos + 5, rsSummary(1) + 1).Value = .Cells(Total_Row_Pos + 5, rsSummary(1) + 1).Value + rsSummary("NETT_FEE_OTHER")
                        'total year
                        .Cells(Total_Row_Pos + 5, 14).Value = .Cells(Total_Row_Pos + 5, 14).Value + rsSummary("NETT_FEE_OTHER")
                End Select
                'Sub Total month
                .Cells(Current_Task_Row + 6, rsSummary(1) + 1).Value = .Cells(Current_Task_Row + 6, rsSummary(1) + 1).Value + rsSummary("NETT_FEE_OTHER")
                'Sub Total year
                .Cells(Current_Task_Row + 6, 14).Value = .Cells(Current_Task_Row + 6, 14).Value + rsSummary("NETT_FEE_OTHER")
                'Grand Total Month
                .Cells(Total_Row_Pos + 6, rsSummary(1) + 1).Value = .Cells(Total_Row_Pos + 6, rsSummary(1) + 1).Value + rsSummary("NETT_FEE_OTHER")
                'Grand Total Year
                .Cells(Total_Row_Pos + 6, 14).Value = .Cells(Total_Row_Pos + 6, 14).Value + rsSummary("NETT_FEE_OTHER")
                
                rsSummary.MoveNext
                Call incrProgressBar
            Wend
            rsSummary.Close
            Current_Task_Row = Current_Task_Row + 7
            Task_Num = Task_Num + 1
            rsTask.MoveNext
        Wend
        rsTask.Close
        
        .Cells.Style = "Comma"
        .Cells.Font.Name = "Tahoma"
        .Cells.Font.Size = 8
        
        'inner borders
        With .Range(.Cells(intStartRow, 1).Address & ":" & .Cells(Total_Row_Pos + 6, 15).Address).Borders
            .LineStyle = xlContinuous
            .Weight = xlThin
            .ColorIndex = 40
        End With
        
        'blok warna kolom percent
        .Range("O" & CStr(intStartRow) & ":" & .Cells(Total_Row_Pos + 6, 15).Address).Interior.ColorIndex = 36
        .Range("O" & CStr(intStartRow) & ":" & .Cells(Total_Row_Pos + 6, 15).Address).NumberFormat = "#0.00%"
        
        'set formula kolom %
        '----------------------- Cells.FormulaR1C1 Example ---------------------------
        'R = curent row, R[x] = current row + x, R[-x] = curent row - x. C = Column
        'contoh :   RC = current row, current col
        '        RC[1] = current row, current col+1
        '       R[-1]C = current row - 1, curent col
        '-----------------------------------------------------------------------------
        .Cells(Total_Row_Pos + 1, 15).FormulaR1C1 = "=RC[-1]/R[5]C[-1]"
        .Cells(Total_Row_Pos + 2, 15).FormulaR1C1 = "=RC[-1]/R[4]C[-1]"
        .Cells(Total_Row_Pos + 3, 15).FormulaR1C1 = "=RC[-1]/R[3]C[-1]"
        .Cells(Total_Row_Pos + 4, 15).FormulaR1C1 = "=RC[-1]/R[2]C[-1]"
        .Cells(Total_Row_Pos + 5, 15).FormulaR1C1 = "=RC[-1]/R[1]C[-1]"
        
        'draw outer border
        With .Range(.Cells(intStartRow, 1).Address & ":" & .Cells(Total_Row_Pos + 6, 15).Address).Borders(xlEdgeLeft)
            .LineStyle = xlContinuous
            .Weight = xlThick
            .ColorIndex = 5
        End With
        With .Range(.Cells(intStartRow, 1).Address & ":" & .Cells(Total_Row_Pos + 6, 15).Address).Borders(xlEdgeRight)
            .LineStyle = xlContinuous
            .Weight = xlThick
            .ColorIndex = 5
        End With
        With .Range(.Cells(intStartRow, 1).Address & ":" & .Cells(Total_Row_Pos + 6, 15).Address).Borders(xlEdgeTop)
            .LineStyle = xlContinuous
            .Weight = xlThick
            .ColorIndex = 5
        End With
        With .Range(.Cells(intStartRow, 1).Address & ":" & .Cells(Total_Row_Pos + 6, 15).Address).Borders(xlEdgeBottom)
            .LineStyle = xlContinuous
            .Weight = xlThick
            .ColorIndex = 5
        End With
        
        .Cells.EntireColumn.AutoFit
        .Cells.EntireRow.AutoFit
        
        'Format %
        .Range("O1:" & .Cells(Total_Row_Pos + 6, 15).Address).NumberFormat = "#0.00%"
       
    End With
    
'======================================================BALANCE (PLAN - ACTUAL)==============================
    'EXPORT Balance
    lblPleaseWait.Caption = "Exporting Summary... [BALANCE]"
    lblPercent.Caption = ""
    lblPercent.Refresh
    ProgressBarExport.Max = 14 * 6
    ProgressBarExport.Value = 0
    intStartRow = Total_Row_Pos + 9
    Total_Row_Pos = intStartRow
    With xlws
        'Print Header
        .Cells(intStartRow - 1, 1).Value = "BALANCE"
        
        .Cells(intStartRow, 1).Value = "Medium"
        .Cells(intStartRow, 1).Interior.ColorIndex = 36
        For i = 1 To 12
            .Cells(intStartRow, i + 1).Value = EngMonthName(i)
            .Cells(intStartRow, i + 1).Interior.ColorIndex = BgColorMonth(i - 1)
        Next
        .Cells(intStartRow, 14).Value = "TOTAL"
        .Cells(intStartRow, 14).Interior.ColorIndex = 36
        
        .Range("A" & CStr(intStartRow) & ":N" & CStr(intStartRow)).HorizontalAlignment = xlCenter
        .Range("A" & CStr(intStartRow) & ":N" & CStr(intStartRow)).Font.Bold = True
        
        For intRow = 1 To 6
            For intCol = 1 To 14
                If intCol > 1 Then
                    .Cells(intStartRow + intRow, intCol).Formula = "=" & .Cells(intTotal_Row_Pos_Plan + intRow, intCol).Address & " - " & .Cells(intTotal_Row_Pos_Actual + intRow, intCol).Address
                Else
                    .Cells(intStartRow + intRow, intCol).Value = .Cells(intTotal_Row_Pos_Plan + intRow, intCol).Value
                End If
                Call incrProgressBar
            Next
        Next
        
        .Range("A" & CStr(intStartRow + 6) & ":N" & CStr(intStartRow + 6)).Font.Bold = True
        
        'inner borders
        With .Range(.Cells(intStartRow, 1).Address & ":" & .Cells(intStartRow + 6, 15).Address).Borders
            .LineStyle = xlContinuous
            .Weight = xlThin
            .ColorIndex = 40
        End With
        
        'draw outer border
        With .Range(.Cells(intStartRow, 1).Address & ":" & .Cells(intStartRow + 6, 14).Address).Borders(xlEdgeLeft)
            .LineStyle = xlContinuous
            .Weight = xlThick
            .ColorIndex = 5
        End With
        With .Range(.Cells(intStartRow, 1).Address & ":" & .Cells(intStartRow + 6, 14).Address).Borders(xlEdgeRight)
            .LineStyle = xlContinuous
            .Weight = xlThick
            .ColorIndex = 5
        End With
        With .Range(.Cells(intStartRow, 1).Address & ":" & .Cells(intStartRow + 6, 14).Address).Borders(xlEdgeTop)
            .LineStyle = xlContinuous
            .Weight = xlThick
            .ColorIndex = 5
        End With
        With .Range(.Cells(intStartRow, 1).Address & ":" & .Cells(intStartRow + 6, 14).Address).Borders(xlEdgeBottom)
            .LineStyle = xlContinuous
            .Weight = xlThick
            .ColorIndex = 5
        End With
        
        .Activate
        'Protect Sheet
        .Protect "27 October", DrawingObjects:=False, Contents:=True, Scenarios:=True, AllowFormattingCells:=True, AllowFormattingColumns:=True, AllowFormattingRows:=True
        .Cells(3, 2).Select
        
    End With
End Sub

Private Sub ExportTVLayering(xlws As Object)
    Dim TVL_Row As Integer
    Dim TVL_Col As Integer
    Dim strCurrentMonth As String, strCurrentRF As String
    Dim intWeekFrom As Integer, IntWeekTo As Integer
    Dim jumlahTask As Integer, strTemp As String
    
    frm_MPTVLayering.show
    
    jumlahTask = 0
    On Error Resume Next
    Do
        jumlahTask = jumlahTask + 1
        strTemp = frm_MPTVLayering.FGTVLayering.TextMatrix(frm_MPTVLayering.FGTVLayering.Rows - jumlahTask, 0)
        If Err.Number <> 0 Then
            Err.Clear
            Exit Do
        End If
    Loop Until strTemp = ""
    jumlahTask = jumlahTask - 1
    
    lblPleaseWait.Caption = "Exporting TV Layering.."
    lblPercent.Caption = ""
    lblPercent.Refresh
    ProgressBarExport.Max = frm_MPTVLayering.FGTVLayering.Rows
    ProgressBarExport.Min = 0
    ProgressBarExport.Value = 0
    
    With xlws
        '1. print header
        
        '1.a print header kiri
        TVL_Row = 1
        For TVL_Col = 1 To 4
            .Cells(TVL_Row, TVL_Col).Value = Trim(frm_MPTVLayering.FGTVLayering.TextMatrix(TVL_Row - 1, TVL_Col - 1))
            .Range(.Cells(1, TVL_Col).Address & ":" & .Cells(3, TVL_Col).Address).Merge
            .Range(.Cells(1, TVL_Col).Address & ":" & .Cells(3, TVL_Col).Address).Interior.ColorIndex = 34
            .Cells(frm_MPTVLayering.FGTVLayering.Rows - jumlahTask, TVL_Col).Interior.ColorIndex = 34
        Next
        
        '1.b print header kanan (bulan)
        
        Dim BgColorMonth(12) As Integer
        BgColorMonth(0) = 40
        BgColorMonth(1) = 36
        BgColorMonth(2) = 35
        BgColorMonth(3) = 34
        BgColorMonth(4) = 40
        BgColorMonth(5) = 36
        BgColorMonth(6) = 35
        BgColorMonth(7) = 34
        BgColorMonth(8) = 40
        BgColorMonth(9) = 36
        BgColorMonth(10) = 35
        BgColorMonth(11) = 34
        
        strCurrentMonth = ""
        For TVL_Col = 5 To frm_MPTVLayering.FGTVLayering.cols
            If Trim(frm_MPTVLayering.FGTVLayering.TextMatrix(TVL_Row - 1, TVL_Col - 1)) <> strCurrentMonth Then
                If strCurrentMonth <> "" Then 'merge current month
                    IntWeekTo = TVL_Col - 1
                    .Range(.Cells(TVL_Row, intWeekFrom).Address & ":" & .Cells(TVL_Row, IntWeekTo).Address).Merge
                    .Range(.Cells(TVL_Row, intWeekFrom).Address & ":" & .Cells(TVL_Row, IntWeekTo).Address).Interior.ColorIndex = BgColorMonth(EngMonthIndex(strCurrentMonth) - 1)
                    .Range(.Cells(frm_MPTVLayering.FGTVLayering.Rows - jumlahTask, intWeekFrom).Address & ":" & .Cells(frm_MPTVLayering.FGTVLayering.Rows - jumlahTask, IntWeekTo).Address).Interior.ColorIndex = BgColorMonth(EngMonthIndex(strCurrentMonth) - 1)
                    'draw border pemisah bulan
                    
                    With .Range(.Cells(4, intWeekFrom).Address & ":" & .Cells(frm_MPTVLayering.FGTVLayering.Rows - jumlahTask, IntWeekTo).Address).Borders(xlEdgeRight)
                        .LineStyle = xlContinuous
                        .Weight = xlThin
                        .ColorIndex = 40
                    End With
                    With .Range(.Cells(4, intWeekFrom).Address & ":" & .Cells(frm_MPTVLayering.FGTVLayering.Rows - jumlahTask, IntWeekTo).Address).Borders(xlEdgeLeft)
                        .LineStyle = xlContinuous
                        .Weight = xlThin
                        .ColorIndex = 40
                    End With
                    With .Range(.Cells(4, intWeekFrom).Address & ":" & .Cells(frm_MPTVLayering.FGTVLayering.Rows - jumlahTask, IntWeekTo).Address).Borders(xlEdgeTop)
                        .LineStyle = xlContinuous
                        .Weight = xlThin
                        .ColorIndex = 40
                    End With
                    With .Range(.Cells(4, intWeekFrom).Address & ":" & .Cells(frm_MPTVLayering.FGTVLayering.Rows - jumlahTask, IntWeekTo).Address).Borders(xlEdgeBottom)
                        .LineStyle = xlContinuous
                        .Weight = xlThin
                        .ColorIndex = 40
                    End With
                    
                    'draw pattern
'                    With .Range(.Cells(4, intWeekFrom).Address & ":" & .Cells(frm_MPTVLayering.FGTVLayering.Rows - jumlahTask, intWeekTo).Address).Interior
'                        .ColorIndex = BgColorMonth(EngMonthIndex(strCurrentMonth) - 1)
'                        .Pattern = xlGray8
'                        .PatternColorIndex = xlAutomatic
'                    End With
                
                End If
                'print next month
                strCurrentMonth = Trim(frm_MPTVLayering.FGTVLayering.TextMatrix(TVL_Row - 1, TVL_Col - 1))
                .Cells(TVL_Row, TVL_Col).Value = strCurrentMonth
                intWeekFrom = TVL_Col
            End If
        Next
        '1.c merge header kolom total
        .Range(.Cells(1, frm_MPTVLayering.FGTVLayering.cols).Address & ":" & .Cells(3, frm_MPTVLayering.FGTVLayering.cols).Address).Merge
        .Range(.Cells(1, frm_MPTVLayering.FGTVLayering.cols).Address & ":" & .Cells(3, frm_MPTVLayering.FGTVLayering.cols).Address).Interior.ColorIndex = 34
        .Cells(frm_MPTVLayering.FGTVLayering.Rows - jumlahTask, frm_MPTVLayering.FGTVLayering.cols).Interior.ColorIndex = 34
        
        Call incrProgressBar
        
        '1.d print header kanan(week dan tanggal)
        For TVL_Row = 2 To 3
            For TVL_Col = 5 To frm_MPTVLayering.FGTVLayering.cols - 1
                .Cells(TVL_Row, TVL_Col).Value = Trim(frm_MPTVLayering.FGTVLayering.TextMatrix(TVL_Row - 1, TVL_Col - 1))
            Next
            Call incrProgressBar
        Next
        
        '1.e set teks alignment untuk header
        .Range(.Cells(1, 1).Address & ":" & .Cells(3, frm_MPTVLayering.FGTVLayering.cols).Address).HorizontalAlignment = xlCenter
        .Range(.Cells(1, 1).Address & ":" & .Cells(3, frm_MPTVLayering.FGTVLayering.cols).Address).VerticalAlignment = xlTop
        
        '1.f set border header
        .Range(.Cells(1, 1).Address & ":" & .Cells(3, frm_MPTVLayering.FGTVLayering.cols).Address).Borders.LineStyle = xlContinuous
        .Range(.Cells(1, 1).Address & ":" & .Cells(3, frm_MPTVLayering.FGTVLayering.cols).Address).Borders.Weight = xlThin
        .Range(.Cells(1, 1).Address & ":" & .Cells(3, frm_MPTVLayering.FGTVLayering.cols).Address).Borders.ColorIndex = 40
        
        '1.g set border untuk footer
        .Range(.Cells(frm_MPTVLayering.FGTVLayering.Rows - jumlahTask + 1, 1).Address & ":" & .Cells(frm_MPTVLayering.FGTVLayering.Rows, frm_MPTVLayering.FGTVLayering.cols).Address).Borders.LineStyle = xlContinuous
        .Range(.Cells(frm_MPTVLayering.FGTVLayering.Rows - jumlahTask + 1, 1).Address & ":" & .Cells(frm_MPTVLayering.FGTVLayering.Rows, frm_MPTVLayering.FGTVLayering.cols).Address).Borders.Weight = xlThin
        .Range(.Cells(frm_MPTVLayering.FGTVLayering.Rows - jumlahTask + 1, 1).Address & ":" & .Cells(frm_MPTVLayering.FGTVLayering.Rows, frm_MPTVLayering.FGTVLayering.cols).Address).Borders.ColorIndex = 40
        
        '2.a print detail
        For TVL_Row = 4 To frm_MPTVLayering.FGTVLayering.Rows
            If Trim(frm_MPTVLayering.FGTVLayering.TextMatrix(TVL_Row - 1, 2)) <> "Reach" And Trim(frm_MPTVLayering.FGTVLayering.TextMatrix(TVL_Row - 1, 2)) <> "Freq" Then
                For TVL_Col = 1 To frm_MPTVLayering.FGTVLayering.cols
                    .Cells(TVL_Row, TVL_Col).Value = Trim(frm_MPTVLayering.FGTVLayering.TextMatrix(TVL_Row - 1, TVL_Col - 1))
                Next
            Else
                'Reach & Freq
                For TVL_Col = 1 To 4
                    .Cells(TVL_Row, TVL_Col).Value = Trim(frm_MPTVLayering.FGTVLayering.TextMatrix(TVL_Row - 1, TVL_Col - 1))
                Next
                strCurrentRF = ""
                For TVL_Col = 5 To frm_MPTVLayering.FGTVLayering.cols
                    If frm_MPTVLayering.FGTVLayering.TextMatrix(TVL_Row - 1, TVL_Col - 1) <> strCurrentRF Then
                        If Trim(strCurrentRF) <> "" Then 'merge current rf
                            IntWeekTo = TVL_Col - 1
                            .Range(.Cells(TVL_Row, intWeekFrom).Address & ":" & .Cells(TVL_Row, IntWeekTo).Address).Merge
                            .Range(.Cells(TVL_Row, intWeekFrom).Address & ":" & .Cells(TVL_Row, IntWeekTo).Address).Interior.ColorIndex = 6
                            .Range(.Cells(TVL_Row, intWeekFrom).Address & ":" & .Cells(TVL_Row, IntWeekTo).Address).Borders.LineStyle = xlContinuous
                            .Range(.Cells(TVL_Row, intWeekFrom).Address & ":" & .Cells(TVL_Row, IntWeekTo).Address).Borders.Weight = xlThin
                            .Range(.Cells(TVL_Row, intWeekFrom).Address & ":" & .Cells(TVL_Row, IntWeekTo).Address).Borders.ColorIndex = 40
                        End If
                        'print next rf
                        strCurrentRF = frm_MPTVLayering.FGTVLayering.TextMatrix(TVL_Row - 1, TVL_Col - 1)
                        .Cells(TVL_Row, TVL_Col).Value = strCurrentRF
                        intWeekFrom = TVL_Col
                    End If
                Next
            End If
            Call incrProgressBar
        Next
        
        '2.b set teks alignment untuk tarps
        .Range(.Cells(3, 5).Address & ":" & .Cells(frm_MPTVLayering.FGTVLayering.Rows, frm_MPTVLayering.FGTVLayering.cols).Address).HorizontalAlignment = xlCenter
        .Range(.Cells(3, 5).Address & ":" & .Cells(frm_MPTVLayering.FGTVLayering.Rows, frm_MPTVLayering.FGTVLayering.cols).Address).VerticalAlignment = xlTop
        
        '3 set default font
        With .Cells.EntireColumn.Font
            .Name = "Tahoma"
            .Size = 8
        End With
        '4 set col & row height (auto fit)
        .Cells.EntireColumn.AutoFit
        .Cells.EntireRow.AutoFit
        
        '5 draw outer border
        With .Range(.Cells(1, 1).Address & ":" & .Cells(frm_MPTVLayering.FGTVLayering.Rows, frm_MPTVLayering.FGTVLayering.cols).Address).Borders(xlEdgeLeft)
            .LineStyle = xlContinuous
            .Weight = xlThick
            .ColorIndex = 5
        End With
        With .Range(.Cells(1, 1).Address & ":" & .Cells(frm_MPTVLayering.FGTVLayering.Rows, frm_MPTVLayering.FGTVLayering.cols).Address).Borders(xlEdgeRight)
            .LineStyle = xlContinuous
            .Weight = xlThick
            .ColorIndex = 5
        End With
        With .Range(.Cells(1, 1).Address & ":" & .Cells(frm_MPTVLayering.FGTVLayering.Rows, frm_MPTVLayering.FGTVLayering.cols).Address).Borders(xlEdgeTop)
            .LineStyle = xlContinuous
            .Weight = xlThick
            .ColorIndex = 5
        End With
        With .Range(.Cells(1, 1).Address & ":" & .Cells(frm_MPTVLayering.FGTVLayering.Rows, frm_MPTVLayering.FGTVLayering.cols).Address).Borders(xlEdgeBottom)
            .LineStyle = xlContinuous
            .Weight = xlThick
            .ColorIndex = 5
        End With
        .Activate
        'Protect Sheet
        .Protect "27 October", DrawingObjects:=False, Contents:=True, Scenarios:=True, AllowFormattingCells:=True, AllowFormattingColumns:=True, AllowFormattingRows:=True
        
    End With
    
    Unload frm_MPTVLayering
End Sub

Private Sub TotalByTask(xlws As Object, xlRow As Integer, xlLeftCol As Integer, xlStartRowTask As Integer, intTask As Integer, strActual As String)
'*****************************************************************************
' Nama Submodul         :  TotalByTask
' Fungsi Submodul       :  Print Total By Task
' Nama Programmer       :  Sistyo
' Tgl Pembuatan         :  18 Agustus 2004
' Last Update           :  18 Agustus 2004/Sistyo
'******************************************************************************
    If strActual = "Actual" Then
        strActual = " (Actual)"
    Else
        strActual = ""
    End If
    With xlws
        .Rows(CStr(xlRow) & ":" & CStr(xlRow + 6)).Insert Shift:=xlDown
        xlRow = xlRow + 7
         
        .Range(.Cells(xlRow - 7, xlLeftCol).Address & ":" & .Cells(xlRow - 3, xlLeftCol + msf_MPInsertion.cols - 3).Address).Interior.ColorIndex = 34

        .Range(.Cells(xlRow - 7, xlLeftCol).Address & ":" & .Cells(xlRow - 2, xlLeftCol + msf_MPInsertion.cols - 3).Address).Borders(xlEdgeTop).LineStyle = xlContinuous
        .Range(.Cells(xlRow - 7, xlLeftCol).Address & ":" & .Cells(xlRow - 2, xlLeftCol + msf_MPInsertion.cols - 3).Address).Borders(xlEdgeTop).Weight = xlThin
        .Range(.Cells(xlRow - 7, xlLeftCol).Address & ":" & .Cells(xlRow - 2, xlLeftCol + msf_MPInsertion.cols - 3).Address).Borders(xlEdgeTop).ColorIndex = xlAutomatic
        
        .Range(.Cells(xlRow - 7, xlLeftCol).Address & ":" & .Cells(xlRow - 2, xlLeftCol + msf_MPInsertion.cols - 3).Address).Borders(xlEdgeBottom).LineStyle = xlContinuous
        .Range(.Cells(xlRow - 7, xlLeftCol).Address & ":" & .Cells(xlRow - 2, xlLeftCol + msf_MPInsertion.cols - 3).Address).Borders(xlEdgeBottom).Weight = xlThin
        .Range(.Cells(xlRow - 7, xlLeftCol).Address & ":" & .Cells(xlRow - 2, xlLeftCol + msf_MPInsertion.cols - 3).Address).Borders(xlEdgeBottom).ColorIndex = xlAutomatic
        
        .Range(.Cells(xlRow - 7, xlLeftCol).Address & ":" & .Cells(xlRow - 2, xlLeftCol + msf_MPInsertion.cols - 3).Address).Borders(xlInsideHorizontal).LineStyle = xlContinuous
        .Range(.Cells(xlRow - 7, xlLeftCol).Address & ":" & .Cells(xlRow - 2, xlLeftCol + msf_MPInsertion.cols - 3).Address).Borders(xlInsideHorizontal).Weight = xlThin
        .Range(.Cells(xlRow - 7, xlLeftCol).Address & ":" & .Cells(xlRow - 2, xlLeftCol + msf_MPInsertion.cols - 3).Address).Borders(xlInsideHorizontal).ColorIndex = xlAutomatic
        
        .Range(.Cells(xlRow - 7, xlLeftCol).Address & ":" & .Cells(xlRow - 2, xlLeftCol + msf_MPInsertion.cols - 3).Address).Font.Size = 9
        
        .Cells(xlRow - 7, xlLeftCol).Value = "Television" & strActual
        .Cells(xlRow - 7, xlLeftCol).Font.ColorIndex = 1
        .Cells(xlRow - 7, xlLeftCol + 3).Value = "Sub Total TV" & strActual & " - Task " & CStr(intTask)
        .Cells(xlRow - 7, xlLeftCol + 3).Font.ColorIndex = 34
        
        .Cells(xlRow - 6, xlLeftCol).Value = "Radio" & strActual
        .Cells(xlRow - 6, xlLeftCol).Font.ColorIndex = 1
        .Cells(xlRow - 6, xlLeftCol + 3).Value = "Sub Total Radio" & strActual & " - Task " & CStr(intTask)
        .Cells(xlRow - 6, xlLeftCol + 3).Font.ColorIndex = 34
        
        .Cells(xlRow - 5, xlLeftCol).Value = "Print" & strActual
        .Cells(xlRow - 5, xlLeftCol).Font.ColorIndex = 1
        .Cells(xlRow - 5, xlLeftCol + 3).Value = "Sub Total Print" & strActual & " - Task " & CStr(intTask)
        .Cells(xlRow - 5, xlLeftCol + 3).Font.ColorIndex = 34
        
        .Cells(xlRow - 4, xlLeftCol).Value = "Cinema" & strActual
        .Cells(xlRow - 4, xlLeftCol).Font.ColorIndex = 1
        .Cells(xlRow - 4, xlLeftCol + 3).Value = "Sub Total Cinema" & strActual & " - Task " & CStr(intTask)
        .Cells(xlRow - 4, xlLeftCol + 3).Font.ColorIndex = 34
        
        .Cells(xlRow - 3, xlLeftCol).Value = "Other" & strActual
        .Cells(xlRow - 3, xlLeftCol).Font.ColorIndex = 1
        .Cells(xlRow - 3, xlLeftCol + 3).Value = "Sub Total Other" & strActual & " - Task " & CStr(intTask)
        .Cells(xlRow - 3, xlLeftCol + 3).Font.ColorIndex = 34
        
        .Cells(xlRow - 2, xlLeftCol).Value = "Total Netto" & strActual & " Task " & intTask
        .Cells(xlRow - 2, xlLeftCol).Font.ColorIndex = 1
        
        .Range(sRange_Jan & CStr(xlRow - 7) & ":" & eRange_Jan & CStr(xlRow - 7)).Merge
        .Range(sRange_Jan & CStr(xlRow - 7)).Formula = "=sumif(" & .Cells(xlStartRowTask, xlLeftCol).Address & ":" & .Cells(xlRow - 7 - 1, xlLeftCol).Address & "," & .Cells(xlRow - 7, xlLeftCol + 3).Address & "," & sRange_Jan & CStr(xlStartRowTask) & ":" & eRange_Jan & CStr(xlRow - 7 - 1) & ")"
        .Range(sRange_Jan & CStr(xlRow - 7)).NumberFormat = "_(#,#0.00_);_(#,#0.00_);_(-_)"
        
        .Range(sRange_Jan & CStr(xlRow - 6) & ":" & eRange_Jan & CStr(xlRow - 6)).Merge
        .Range(sRange_Jan & CStr(xlRow - 6)).Formula = "=sumif(" & .Cells(xlStartRowTask, xlLeftCol).Address & ":" & .Cells(xlRow - 7 - 1, xlLeftCol).Address & "," & .Cells(xlRow - 6, xlLeftCol + 3).Address & "," & sRange_Jan & CStr(xlStartRowTask) & ":" & eRange_Jan & CStr(xlRow - 7 - 1) & ")"
        .Range(sRange_Jan & CStr(xlRow - 6)).NumberFormat = "_(#,#0.00_);_(#,#0.00_);_(-_)"
        
        .Range(sRange_Jan & CStr(xlRow - 5) & ":" & eRange_Jan & CStr(xlRow - 5)).Merge
        .Range(sRange_Jan & CStr(xlRow - 5)).Formula = "=sumif(" & .Cells(xlStartRowTask, xlLeftCol).Address & ":" & .Cells(xlRow - 7 - 1, xlLeftCol).Address & "," & .Cells(xlRow - 5, xlLeftCol + 3).Address & "," & sRange_Jan & CStr(xlStartRowTask) & ":" & eRange_Jan & CStr(xlRow - 7 - 1) & ")"
        .Range(sRange_Jan & CStr(xlRow - 5)).NumberFormat = "_(#,#0.00_);_(#,#0.00_);_(-_)"
        
        .Range(sRange_Jan & CStr(xlRow - 4) & ":" & eRange_Jan & CStr(xlRow - 4)).Merge
        .Range(sRange_Jan & CStr(xlRow - 4)).Formula = "=sumif(" & .Cells(xlStartRowTask, xlLeftCol).Address & ":" & .Cells(xlRow - 7 - 1, xlLeftCol).Address & "," & .Cells(xlRow - 4, xlLeftCol + 3).Address & "," & sRange_Jan & CStr(xlStartRowTask) & ":" & eRange_Jan & CStr(xlRow - 7 - 1) & ")"
        .Range(sRange_Jan & CStr(xlRow - 4)).NumberFormat = "_(#,#0.00_);_(#,#0.00_);_(-_)"
        
        .Range(sRange_Jan & CStr(xlRow - 3) & ":" & eRange_Jan & CStr(xlRow - 3)).Merge
        .Range(sRange_Jan & CStr(xlRow - 3)).Formula = "=sumif(" & .Cells(xlStartRowTask, xlLeftCol).Address & ":" & .Cells(xlRow - 7 - 1, xlLeftCol).Address & "," & .Cells(xlRow - 3, xlLeftCol + 3).Address & "," & sRange_Jan & CStr(xlStartRowTask) & ":" & eRange_Jan & CStr(xlRow - 7 - 1) & ")"
        .Range(sRange_Jan & CStr(xlRow - 3)).NumberFormat = "_(#,#0.00_);_(#,#0.00_);_(-_)"
        
        .Range(sRange_Jan & CStr(xlRow - 2) & ":" & eRange_Jan & CStr(xlRow - 2)).Merge 'Total Netto Task X - Jan
        .Range(sRange_Jan & CStr(xlRow - 2)).Formula = "=sum(" & sRange_Jan & CStr(xlRow - 3) & ":" & eRange_Jan & CStr(xlRow - 7) & ")"
        .Range(sRange_Jan & CStr(xlRow - 2)).NumberFormat = "_(#,#0.00_);_(#,#0.00_);_(-_)"
        
        .Range(sRange_Feb & CStr(xlRow - 7) & ":" & eRange_Feb & CStr(xlRow - 7)).Merge
        .Range(sRange_Feb & CStr(xlRow - 7)).Formula = "=sumif(" & .Cells(xlStartRowTask, xlLeftCol).Address & ":" & .Cells(xlRow - 7 - 1, xlLeftCol).Address & "," & .Cells(xlRow - 7, xlLeftCol + 3).Address & "," & sRange_Feb & CStr(xlStartRowTask) & ":" & eRange_Feb & CStr(xlRow - 7 - 1) & ")"
        .Range(sRange_Feb & CStr(xlRow - 7)).NumberFormat = "_(#,#0.00_);_(#,#0.00_);_(-_)"
        
        .Range(sRange_Feb & CStr(xlRow - 6) & ":" & eRange_Feb & CStr(xlRow - 6)).Merge
        .Range(sRange_Feb & CStr(xlRow - 6)).Formula = "=sumif(" & .Cells(xlStartRowTask, xlLeftCol).Address & ":" & .Cells(xlRow - 7 - 1, xlLeftCol).Address & "," & .Cells(xlRow - 6, xlLeftCol + 3).Address & "," & sRange_Feb & CStr(xlStartRowTask) & ":" & eRange_Feb & CStr(xlRow - 7 - 1) & ")"
        .Range(sRange_Feb & CStr(xlRow - 6)).NumberFormat = "_(#,#0.00_);_(#,#0.00_);_(-_)"
        
        .Range(sRange_Feb & CStr(xlRow - 5) & ":" & eRange_Feb & CStr(xlRow - 5)).Merge
        .Range(sRange_Feb & CStr(xlRow - 5)).Formula = "=sumif(" & .Cells(xlStartRowTask, xlLeftCol).Address & ":" & .Cells(xlRow - 7 - 1, xlLeftCol).Address & "," & .Cells(xlRow - 5, xlLeftCol + 3).Address & "," & sRange_Feb & CStr(xlStartRowTask) & ":" & eRange_Feb & CStr(xlRow - 7 - 1) & ")"
        .Range(sRange_Feb & CStr(xlRow - 5)).NumberFormat = "_(#,#0.00_);_(#,#0.00_);_(-_)"
        
        .Range(sRange_Feb & CStr(xlRow - 4) & ":" & eRange_Feb & CStr(xlRow - 4)).Merge
        .Range(sRange_Feb & CStr(xlRow - 4)).Formula = "=sumif(" & .Cells(xlStartRowTask, xlLeftCol).Address & ":" & .Cells(xlRow - 7 - 1, xlLeftCol).Address & "," & .Cells(xlRow - 4, xlLeftCol + 3).Address & "," & sRange_Feb & CStr(xlStartRowTask) & ":" & eRange_Feb & CStr(xlRow - 7 - 1) & ")"
        .Range(sRange_Feb & CStr(xlRow - 4)).NumberFormat = "_(#,#0.00_);_(#,#0.00_);_(-_)"
        
        .Range(sRange_Feb & CStr(xlRow - 3) & ":" & eRange_Feb & CStr(xlRow - 3)).Merge
        .Range(sRange_Feb & CStr(xlRow - 3)).Formula = "=sumif(" & .Cells(xlStartRowTask, xlLeftCol).Address & ":" & .Cells(xlRow - 7 - 1, xlLeftCol).Address & "," & .Cells(xlRow - 3, xlLeftCol + 3).Address & "," & sRange_Feb & CStr(xlStartRowTask) & ":" & eRange_Feb & CStr(xlRow - 7 - 1) & ")"
        .Range(sRange_Feb & CStr(xlRow - 3)).NumberFormat = "_(#,#0.00_);_(#,#0.00_);_(-_)"
        
        .Range(sRange_Feb & CStr(xlRow - 2) & ":" & eRange_Feb & CStr(xlRow - 2)).Merge 'Total Netto Task X - Feb
        .Range(sRange_Feb & CStr(xlRow - 2)).Formula = "=sum(" & sRange_Feb & CStr(xlRow - 3) & ":" & eRange_Feb & CStr(xlRow - 7) & ")"
        .Range(sRange_Feb & CStr(xlRow - 2)).NumberFormat = "_(#,#0.00_);_(#,#0.00_);_(-_)"
        
        .Range(sRange_Mar & CStr(xlRow - 7) & ":" & eRange_Mar & CStr(xlRow - 7)).Merge
        .Range(sRange_Mar & CStr(xlRow - 7)).Formula = "=sumif(" & .Cells(xlStartRowTask, xlLeftCol).Address & ":" & .Cells(xlRow - 7 - 1, xlLeftCol).Address & "," & .Cells(xlRow - 7, xlLeftCol + 3).Address & "," & sRange_Mar & CStr(xlStartRowTask) & ":" & eRange_Mar & CStr(xlRow - 7 - 1) & ")"
        .Range(sRange_Mar & CStr(xlRow - 7)).NumberFormat = "_(#,#0.00_);_(#,#0.00_);_(-_)"
        
        .Range(sRange_Mar & CStr(xlRow - 6) & ":" & eRange_Mar & CStr(xlRow - 6)).Merge
        .Range(sRange_Mar & CStr(xlRow - 6)).Formula = "=sumif(" & .Cells(xlStartRowTask, xlLeftCol).Address & ":" & .Cells(xlRow - 7 - 1, xlLeftCol).Address & "," & .Cells(xlRow - 6, xlLeftCol + 3).Address & "," & sRange_Mar & CStr(xlStartRowTask) & ":" & eRange_Mar & CStr(xlRow - 7 - 1) & ")"
        .Range(sRange_Mar & CStr(xlRow - 6)).NumberFormat = "_(#,#0.00_);_(#,#0.00_);_(-_)"
        
        .Range(sRange_Mar & CStr(xlRow - 5) & ":" & eRange_Mar & CStr(xlRow - 5)).Merge
        .Range(sRange_Mar & CStr(xlRow - 5)).Formula = "=sumif(" & .Cells(xlStartRowTask, xlLeftCol).Address & ":" & .Cells(xlRow - 7 - 1, xlLeftCol).Address & "," & .Cells(xlRow - 5, xlLeftCol + 3).Address & "," & sRange_Mar & CStr(xlStartRowTask) & ":" & eRange_Mar & CStr(xlRow - 7 - 1) & ")"
        .Range(sRange_Mar & CStr(xlRow - 5)).NumberFormat = "_(#,#0.00_);_(#,#0.00_);_(-_)"
        
        .Range(sRange_Mar & CStr(xlRow - 4) & ":" & eRange_Mar & CStr(xlRow - 4)).Merge
        .Range(sRange_Mar & CStr(xlRow - 4)).Formula = "=sumif(" & .Cells(xlStartRowTask, xlLeftCol).Address & ":" & .Cells(xlRow - 7 - 1, xlLeftCol).Address & "," & .Cells(xlRow - 4, xlLeftCol + 3).Address & "," & sRange_Mar & CStr(xlStartRowTask) & ":" & eRange_Mar & CStr(xlRow - 7 - 1) & ")"
        .Range(sRange_Mar & CStr(xlRow - 4)).NumberFormat = "_(#,#0.00_);_(#,#0.00_);_(-_)"
        
        .Range(sRange_Mar & CStr(xlRow - 3) & ":" & eRange_Mar & CStr(xlRow - 3)).Merge
        .Range(sRange_Mar & CStr(xlRow - 3)).Formula = "=sumif(" & .Cells(xlStartRowTask, xlLeftCol).Address & ":" & .Cells(xlRow - 7 - 1, xlLeftCol).Address & "," & .Cells(xlRow - 3, xlLeftCol + 3).Address & "," & sRange_Mar & CStr(xlStartRowTask) & ":" & eRange_Mar & CStr(xlRow - 7 - 1) & ")"
        .Range(sRange_Mar & CStr(xlRow - 3)).NumberFormat = "_(#,#0.00_);_(#,#0.00_);_(-_)"
        
        .Range(sRange_Mar & CStr(xlRow - 2) & ":" & eRange_Mar & CStr(xlRow - 2)).Merge 'Total Netto Task X - Mar
        .Range(sRange_Mar & CStr(xlRow - 2)).Formula = "=sum(" & sRange_Mar & CStr(xlRow - 3) & ":" & eRange_Mar & CStr(xlRow - 7) & ")"
        .Range(sRange_Mar & CStr(xlRow - 2)).NumberFormat = "_(#,#0.00_);_(#,#0.00_);_(-_)"
        
        .Range(sRange_Apr & CStr(xlRow - 7) & ":" & eRange_Apr & CStr(xlRow - 7)).Merge
        .Range(sRange_Apr & CStr(xlRow - 7)).Formula = "=sumif(" & .Cells(xlStartRowTask, xlLeftCol).Address & ":" & .Cells(xlRow - 7 - 1, xlLeftCol).Address & "," & .Cells(xlRow - 7, xlLeftCol + 3).Address & "," & sRange_Apr & CStr(xlStartRowTask) & ":" & eRange_Apr & CStr(xlRow - 7 - 1) & ")"
        .Range(sRange_Apr & CStr(xlRow - 7)).NumberFormat = "_(#,#0.00_);_(#,#0.00_);_(-_)"
        
        .Range(sRange_Apr & CStr(xlRow - 6) & ":" & eRange_Apr & CStr(xlRow - 6)).Merge
        .Range(sRange_Apr & CStr(xlRow - 6)).Formula = "=sumif(" & .Cells(xlStartRowTask, xlLeftCol).Address & ":" & .Cells(xlRow - 7 - 1, xlLeftCol).Address & "," & .Cells(xlRow - 6, xlLeftCol + 3).Address & "," & sRange_Apr & CStr(xlStartRowTask) & ":" & eRange_Apr & CStr(xlRow - 7 - 1) & ")"
        .Range(sRange_Apr & CStr(xlRow - 6)).NumberFormat = "_(#,#0.00_);_(#,#0.00_);_(-_)"
        
        .Range(sRange_Apr & CStr(xlRow - 5) & ":" & eRange_Apr & CStr(xlRow - 5)).Merge
        .Range(sRange_Apr & CStr(xlRow - 5)).Formula = "=sumif(" & .Cells(xlStartRowTask, xlLeftCol).Address & ":" & .Cells(xlRow - 7 - 1, xlLeftCol).Address & "," & .Cells(xlRow - 5, xlLeftCol + 3).Address & "," & sRange_Apr & CStr(xlStartRowTask) & ":" & eRange_Apr & CStr(xlRow - 7 - 1) & ")"
        .Range(sRange_Apr & CStr(xlRow - 5)).NumberFormat = "_(#,#0.00_);_(#,#0.00_);_(-_)"
        
        .Range(sRange_Apr & CStr(xlRow - 4) & ":" & eRange_Apr & CStr(xlRow - 4)).Merge
        .Range(sRange_Apr & CStr(xlRow - 4)).Formula = "=sumif(" & .Cells(xlStartRowTask, xlLeftCol).Address & ":" & .Cells(xlRow - 7 - 1, xlLeftCol).Address & "," & .Cells(xlRow - 4, xlLeftCol + 3).Address & "," & sRange_Apr & CStr(xlStartRowTask) & ":" & eRange_Apr & CStr(xlRow - 7 - 1) & ")"
        .Range(sRange_Apr & CStr(xlRow - 4)).NumberFormat = "_(#,#0.00_);_(#,#0.00_);_(-_)"
        
        .Range(sRange_Apr & CStr(xlRow - 3) & ":" & eRange_Apr & CStr(xlRow - 3)).Merge
        .Range(sRange_Apr & CStr(xlRow - 3)).Formula = "=sumif(" & .Cells(xlStartRowTask, xlLeftCol).Address & ":" & .Cells(xlRow - 7 - 1, xlLeftCol).Address & "," & .Cells(xlRow - 3, xlLeftCol + 3).Address & "," & sRange_Apr & CStr(xlStartRowTask) & ":" & eRange_Apr & CStr(xlRow - 7 - 1) & ")"
        .Range(sRange_Apr & CStr(xlRow - 3)).NumberFormat = "_(#,#0.00_);_(#,#0.00_);_(-_)"
        
        .Range(sRange_Apr & CStr(xlRow - 2) & ":" & eRange_Apr & CStr(xlRow - 2)).Merge 'Total Netto Task X - Apr
        .Range(sRange_Apr & CStr(xlRow - 2)).Formula = "=sum(" & sRange_Apr & CStr(xlRow - 3) & ":" & eRange_Apr & CStr(xlRow - 7) & ")"
        .Range(sRange_Apr & CStr(xlRow - 2)).NumberFormat = "_(#,#0.00_);_(#,#0.00_);_(-_)"
        
        .Range(sRange_May & CStr(xlRow - 7) & ":" & eRange_May & CStr(xlRow - 7)).Merge
        .Range(sRange_May & CStr(xlRow - 7)).Formula = "=sumif(" & .Cells(xlStartRowTask, xlLeftCol).Address & ":" & .Cells(xlRow - 7 - 1, xlLeftCol).Address & "," & .Cells(xlRow - 7, xlLeftCol + 3).Address & "," & sRange_May & CStr(xlStartRowTask) & ":" & eRange_May & CStr(xlRow - 7 - 1) & ")"
        .Range(sRange_May & CStr(xlRow - 7)).NumberFormat = "_(#,#0.00_);_(#,#0.00_);_(-_)"
        
        .Range(sRange_May & CStr(xlRow - 6) & ":" & eRange_May & CStr(xlRow - 6)).Merge
        .Range(sRange_May & CStr(xlRow - 6)).Formula = "=sumif(" & .Cells(xlStartRowTask, xlLeftCol).Address & ":" & .Cells(xlRow - 7 - 1, xlLeftCol).Address & "," & .Cells(xlRow - 6, xlLeftCol + 3).Address & "," & sRange_May & CStr(xlStartRowTask) & ":" & eRange_May & CStr(xlRow - 7 - 1) & ")"
        .Range(sRange_May & CStr(xlRow - 6)).NumberFormat = "_(#,#0.00_);_(#,#0.00_);_(-_)"
        
        .Range(sRange_May & CStr(xlRow - 5) & ":" & eRange_May & CStr(xlRow - 5)).Merge
        .Range(sRange_May & CStr(xlRow - 5)).Formula = "=sumif(" & .Cells(xlStartRowTask, xlLeftCol).Address & ":" & .Cells(xlRow - 7 - 1, xlLeftCol).Address & "," & .Cells(xlRow - 5, xlLeftCol + 3).Address & "," & sRange_May & CStr(xlStartRowTask) & ":" & eRange_May & CStr(xlRow - 7 - 1) & ")"
        .Range(sRange_May & CStr(xlRow - 5)).NumberFormat = "_(#,#0.00_);_(#,#0.00_);_(-_)"
        
        .Range(sRange_May & CStr(xlRow - 4) & ":" & eRange_May & CStr(xlRow - 4)).Merge
        .Range(sRange_May & CStr(xlRow - 4)).Formula = "=sumif(" & .Cells(xlStartRowTask, xlLeftCol).Address & ":" & .Cells(xlRow - 7 - 1, xlLeftCol).Address & "," & .Cells(xlRow - 4, xlLeftCol + 3).Address & "," & sRange_May & CStr(xlStartRowTask) & ":" & eRange_May & CStr(xlRow - 7 - 1) & ")"
        .Range(sRange_May & CStr(xlRow - 4)).NumberFormat = "_(#,#0.00_);_(#,#0.00_);_(-_)"
        
        .Range(sRange_May & CStr(xlRow - 3) & ":" & eRange_May & CStr(xlRow - 3)).Merge
        .Range(sRange_May & CStr(xlRow - 3)).Formula = "=sumif(" & .Cells(xlStartRowTask, xlLeftCol).Address & ":" & .Cells(xlRow - 7 - 1, xlLeftCol).Address & "," & .Cells(xlRow - 3, xlLeftCol + 3).Address & "," & sRange_May & CStr(xlStartRowTask) & ":" & eRange_May & CStr(xlRow - 7 - 1) & ")"
        .Range(sRange_May & CStr(xlRow - 3)).NumberFormat = "_(#,#0.00_);_(#,#0.00_);_(-_)"
        
        .Range(sRange_May & CStr(xlRow - 2) & ":" & eRange_May & CStr(xlRow - 2)).Merge 'Total Netto Task X - May
        .Range(sRange_May & CStr(xlRow - 2)).Formula = "=sum(" & sRange_May & CStr(xlRow - 3) & ":" & eRange_May & CStr(xlRow - 7) & ")"
        .Range(sRange_May & CStr(xlRow - 2)).NumberFormat = "_(#,#0.00_);_(#,#0.00_);_(-_)"
        
        .Range(sRange_Jun & CStr(xlRow - 7) & ":" & eRange_Jun & CStr(xlRow - 7)).Merge
        .Range(sRange_Jun & CStr(xlRow - 7)).Formula = "=sumif(" & .Cells(xlStartRowTask, xlLeftCol).Address & ":" & .Cells(xlRow - 7 - 1, xlLeftCol).Address & "," & .Cells(xlRow - 7, xlLeftCol + 3).Address & "," & sRange_Jun & CStr(xlStartRowTask) & ":" & eRange_Jun & CStr(xlRow - 7 - 1) & ")"
        .Range(sRange_Jun & CStr(xlRow - 7)).NumberFormat = "_(#,#0.00_);_(#,#0.00_);_(-_)"
        
        .Range(sRange_Jun & CStr(xlRow - 6) & ":" & eRange_Jun & CStr(xlRow - 6)).Merge
        .Range(sRange_Jun & CStr(xlRow - 6)).Formula = "=sumif(" & .Cells(xlStartRowTask, xlLeftCol).Address & ":" & .Cells(xlRow - 7 - 1, xlLeftCol).Address & "," & .Cells(xlRow - 6, xlLeftCol + 3).Address & "," & sRange_Jun & CStr(xlStartRowTask) & ":" & eRange_Jun & CStr(xlRow - 7 - 1) & ")"
        .Range(sRange_Jun & CStr(xlRow - 6)).NumberFormat = "_(#,#0.00_);_(#,#0.00_);_(-_)"
        
        .Range(sRange_Jun & CStr(xlRow - 5) & ":" & eRange_Jun & CStr(xlRow - 5)).Merge
        .Range(sRange_Jun & CStr(xlRow - 5)).Formula = "=sumif(" & .Cells(xlStartRowTask, xlLeftCol).Address & ":" & .Cells(xlRow - 7 - 1, xlLeftCol).Address & "," & .Cells(xlRow - 5, xlLeftCol + 3).Address & "," & sRange_Jun & CStr(xlStartRowTask) & ":" & eRange_Jun & CStr(xlRow - 7 - 1) & ")"
        .Range(sRange_Jun & CStr(xlRow - 5)).NumberFormat = "_(#,#0.00_);_(#,#0.00_);_(-_)"
        
        .Range(sRange_Jun & CStr(xlRow - 4) & ":" & eRange_Jun & CStr(xlRow - 4)).Merge
        .Range(sRange_Jun & CStr(xlRow - 4)).Formula = "=sumif(" & .Cells(xlStartRowTask, xlLeftCol).Address & ":" & .Cells(xlRow - 7 - 1, xlLeftCol).Address & "," & .Cells(xlRow - 4, xlLeftCol + 3).Address & "," & sRange_Jun & CStr(xlStartRowTask) & ":" & eRange_Jun & CStr(xlRow - 7 - 1) & ")"
        .Range(sRange_Jun & CStr(xlRow - 4)).NumberFormat = "_(#,#0.00_);_(#,#0.00_);_(-_)"
        
        .Range(sRange_Jun & CStr(xlRow - 3) & ":" & eRange_Jun & CStr(xlRow - 3)).Merge
        .Range(sRange_Jun & CStr(xlRow - 3)).Formula = "=sumif(" & .Cells(xlStartRowTask, xlLeftCol).Address & ":" & .Cells(xlRow - 7 - 1, xlLeftCol).Address & "," & .Cells(xlRow - 3, xlLeftCol + 3).Address & "," & sRange_Jun & CStr(xlStartRowTask) & ":" & eRange_Jun & CStr(xlRow - 7 - 1) & ")"
        .Range(sRange_Jun & CStr(xlRow - 3)).NumberFormat = "_(#,#0.00_);_(#,#0.00_);_(-_)"
        
        .Range(sRange_Jun & CStr(xlRow - 2) & ":" & eRange_Jun & CStr(xlRow - 2)).Merge 'Total Netto Task X - Jun
        .Range(sRange_Jun & CStr(xlRow - 2)).Formula = "=sum(" & sRange_Jun & CStr(xlRow - 3) & ":" & eRange_Jun & CStr(xlRow - 7) & ")"
        .Range(sRange_Jun & CStr(xlRow - 2)).NumberFormat = "_(#,#0.00_);_(#,#0.00_);_(-_)"
        
        .Range(sRange_Jul & CStr(xlRow - 7) & ":" & eRange_Jul & CStr(xlRow - 7)).Merge
        .Range(sRange_Jul & CStr(xlRow - 7)).Formula = "=sumif(" & .Cells(xlStartRowTask, xlLeftCol).Address & ":" & .Cells(xlRow - 7 - 1, xlLeftCol).Address & "," & .Cells(xlRow - 7, xlLeftCol + 3).Address & "," & sRange_Jul & CStr(xlStartRowTask) & ":" & eRange_Jul & CStr(xlRow - 7 - 1) & ")"
        .Range(sRange_Jul & CStr(xlRow - 7)).NumberFormat = "_(#,#0.00_);_(#,#0.00_);_(-_)"
        
        .Range(sRange_Jul & CStr(xlRow - 6) & ":" & eRange_Jul & CStr(xlRow - 6)).Merge
        .Range(sRange_Jul & CStr(xlRow - 6)).Formula = "=sumif(" & .Cells(xlStartRowTask, xlLeftCol).Address & ":" & .Cells(xlRow - 7 - 1, xlLeftCol).Address & "," & .Cells(xlRow - 6, xlLeftCol + 3).Address & "," & sRange_Jul & CStr(xlStartRowTask) & ":" & eRange_Jul & CStr(xlRow - 7 - 1) & ")"
        .Range(sRange_Jul & CStr(xlRow - 6)).NumberFormat = "_(#,#0.00_);_(#,#0.00_);_(-_)"
        
        .Range(sRange_Jul & CStr(xlRow - 5) & ":" & eRange_Jul & CStr(xlRow - 5)).Merge
        .Range(sRange_Jul & CStr(xlRow - 5)).Formula = "=sumif(" & .Cells(xlStartRowTask, xlLeftCol).Address & ":" & .Cells(xlRow - 7 - 1, xlLeftCol).Address & "," & .Cells(xlRow - 5, xlLeftCol + 3).Address & "," & sRange_Jul & CStr(xlStartRowTask) & ":" & eRange_Jul & CStr(xlRow - 7 - 1) & ")"
        .Range(sRange_Jul & CStr(xlRow - 5)).NumberFormat = "_(#,#0.00_);_(#,#0.00_);_(-_)"
        
        .Range(sRange_Jul & CStr(xlRow - 4) & ":" & eRange_Jul & CStr(xlRow - 4)).Merge
        .Range(sRange_Jul & CStr(xlRow - 4)).Formula = "=sumif(" & .Cells(xlStartRowTask, xlLeftCol).Address & ":" & .Cells(xlRow - 7 - 1, xlLeftCol).Address & "," & .Cells(xlRow - 4, xlLeftCol + 3).Address & "," & sRange_Jul & CStr(xlStartRowTask) & ":" & eRange_Jul & CStr(xlRow - 7 - 1) & ")"
        .Range(sRange_Jul & CStr(xlRow - 4)).NumberFormat = "_(#,#0.00_);_(#,#0.00_);_(-_)"
        
        .Range(sRange_Jul & CStr(xlRow - 3) & ":" & eRange_Jul & CStr(xlRow - 3)).Merge
        .Range(sRange_Jul & CStr(xlRow - 3)).Formula = "=sumif(" & .Cells(xlStartRowTask, xlLeftCol).Address & ":" & .Cells(xlRow - 7 - 1, xlLeftCol).Address & "," & .Cells(xlRow - 3, xlLeftCol + 3).Address & "," & sRange_Jul & CStr(xlStartRowTask) & ":" & eRange_Jul & CStr(xlRow - 7 - 1) & ")"
        .Range(sRange_Jul & CStr(xlRow - 3)).NumberFormat = "_(#,#0.00_);_(#,#0.00_);_(-_)"
        
        .Range(sRange_Jul & CStr(xlRow - 2) & ":" & eRange_Jul & CStr(xlRow - 2)).Merge 'Total Netto Task X - Jul
        .Range(sRange_Jul & CStr(xlRow - 2)).Formula = "=sum(" & sRange_Jul & CStr(xlRow - 3) & ":" & eRange_Jul & CStr(xlRow - 7) & ")"
        .Range(sRange_Jul & CStr(xlRow - 2)).NumberFormat = "_(#,#0.00_);_(#,#0.00_);_(-_)"
        
        .Range(sRange_Aug & CStr(xlRow - 7) & ":" & eRange_Aug & CStr(xlRow - 7)).Merge
        .Range(sRange_Aug & CStr(xlRow - 7)).Formula = "=sumif(" & .Cells(xlStartRowTask, xlLeftCol).Address & ":" & .Cells(xlRow - 7 - 1, xlLeftCol).Address & "," & .Cells(xlRow - 7, xlLeftCol + 3).Address & "," & sRange_Aug & CStr(xlStartRowTask) & ":" & eRange_Aug & CStr(xlRow - 7 - 1) & ")"
        .Range(sRange_Aug & CStr(xlRow - 7)).NumberFormat = "_(#,#0.00_);_(#,#0.00_);_(-_)"
        
        .Range(sRange_Aug & CStr(xlRow - 6) & ":" & eRange_Aug & CStr(xlRow - 6)).Merge
        .Range(sRange_Aug & CStr(xlRow - 6)).Formula = "=sumif(" & .Cells(xlStartRowTask, xlLeftCol).Address & ":" & .Cells(xlRow - 7 - 1, xlLeftCol).Address & "," & .Cells(xlRow - 6, xlLeftCol + 3).Address & "," & sRange_Aug & CStr(xlStartRowTask) & ":" & eRange_Aug & CStr(xlRow - 7 - 1) & ")"
        .Range(sRange_Aug & CStr(xlRow - 6)).NumberFormat = "_(#,#0.00_);_(#,#0.00_);_(-_)"
        
        .Range(sRange_Aug & CStr(xlRow - 5) & ":" & eRange_Aug & CStr(xlRow - 5)).Merge
        .Range(sRange_Aug & CStr(xlRow - 5)).Formula = "=sumif(" & .Cells(xlStartRowTask, xlLeftCol).Address & ":" & .Cells(xlRow - 7 - 1, xlLeftCol).Address & "," & .Cells(xlRow - 5, xlLeftCol + 3).Address & "," & sRange_Aug & CStr(xlStartRowTask) & ":" & eRange_Aug & CStr(xlRow - 7 - 1) & ")"
        .Range(sRange_Aug & CStr(xlRow - 5)).NumberFormat = "_(#,#0.00_);_(#,#0.00_);_(-_)"
        
        .Range(sRange_Aug & CStr(xlRow - 4) & ":" & eRange_Aug & CStr(xlRow - 4)).Merge
        .Range(sRange_Aug & CStr(xlRow - 4)).Formula = "=sumif(" & .Cells(xlStartRowTask, xlLeftCol).Address & ":" & .Cells(xlRow - 7 - 1, xlLeftCol).Address & "," & .Cells(xlRow - 4, xlLeftCol + 3).Address & "," & sRange_Aug & CStr(xlStartRowTask) & ":" & eRange_Aug & CStr(xlRow - 7 - 1) & ")"
        .Range(sRange_Aug & CStr(xlRow - 4)).NumberFormat = "_(#,#0.00_);_(#,#0.00_);_(-_)"
        
        .Range(sRange_Aug & CStr(xlRow - 3) & ":" & eRange_Aug & CStr(xlRow - 3)).Merge
        .Range(sRange_Aug & CStr(xlRow - 3)).Formula = "=sumif(" & .Cells(xlStartRowTask, xlLeftCol).Address & ":" & .Cells(xlRow - 7 - 1, xlLeftCol).Address & "," & .Cells(xlRow - 3, xlLeftCol + 3).Address & "," & sRange_Aug & CStr(xlStartRowTask) & ":" & eRange_Aug & CStr(xlRow - 7 - 1) & ")"
        .Range(sRange_Aug & CStr(xlRow - 3)).NumberFormat = "_(#,#0.00_);_(#,#0.00_);_(-_)"
        
        .Range(sRange_Aug & CStr(xlRow - 2) & ":" & eRange_Aug & CStr(xlRow - 2)).Merge 'Total Netto Task X - Aug
        .Range(sRange_Aug & CStr(xlRow - 2)).Formula = "=sum(" & sRange_Aug & CStr(xlRow - 3) & ":" & eRange_Aug & CStr(xlRow - 7) & ")"
        .Range(sRange_Aug & CStr(xlRow - 2)).NumberFormat = "_(#,#0.00_);_(#,#0.00_);_(-_)"
        
        .Range(sRange_Sep & CStr(xlRow - 7) & ":" & eRange_Sep & CStr(xlRow - 7)).Merge
        .Range(sRange_Sep & CStr(xlRow - 7)).Formula = "=sumif(" & .Cells(xlStartRowTask, xlLeftCol).Address & ":" & .Cells(xlRow - 7 - 1, xlLeftCol).Address & "," & .Cells(xlRow - 7, xlLeftCol + 3).Address & "," & sRange_Sep & CStr(xlStartRowTask) & ":" & eRange_Sep & CStr(xlRow - 7 - 1) & ")"
        .Range(sRange_Sep & CStr(xlRow - 7)).NumberFormat = "_(#,#0.00_);_(#,#0.00_);_(-_)"
        
        .Range(sRange_Sep & CStr(xlRow - 6) & ":" & eRange_Sep & CStr(xlRow - 6)).Merge
        .Range(sRange_Sep & CStr(xlRow - 6)).Formula = "=sumif(" & .Cells(xlStartRowTask, xlLeftCol).Address & ":" & .Cells(xlRow - 7 - 1, xlLeftCol).Address & "," & .Cells(xlRow - 6, xlLeftCol + 3).Address & "," & sRange_Sep & CStr(xlStartRowTask) & ":" & eRange_Sep & CStr(xlRow - 7 - 1) & ")"
        .Range(sRange_Sep & CStr(xlRow - 6)).NumberFormat = "_(#,#0.00_);_(#,#0.00_);_(-_)"
        
        .Range(sRange_Sep & CStr(xlRow - 5) & ":" & eRange_Sep & CStr(xlRow - 5)).Merge
        .Range(sRange_Sep & CStr(xlRow - 5)).Formula = "=sumif(" & .Cells(xlStartRowTask, xlLeftCol).Address & ":" & .Cells(xlRow - 7 - 1, xlLeftCol).Address & "," & .Cells(xlRow - 5, xlLeftCol + 3).Address & "," & sRange_Sep & CStr(xlStartRowTask) & ":" & eRange_Sep & CStr(xlRow - 7 - 1) & ")"
        .Range(sRange_Sep & CStr(xlRow - 5)).NumberFormat = "_(#,#0.00_);_(#,#0.00_);_(-_)"
        
        .Range(sRange_Sep & CStr(xlRow - 4) & ":" & eRange_Sep & CStr(xlRow - 4)).Merge
        .Range(sRange_Sep & CStr(xlRow - 4)).Formula = "=sumif(" & .Cells(xlStartRowTask, xlLeftCol).Address & ":" & .Cells(xlRow - 7 - 1, xlLeftCol).Address & "," & .Cells(xlRow - 4, xlLeftCol + 3).Address & "," & sRange_Sep & CStr(xlStartRowTask) & ":" & eRange_Sep & CStr(xlRow - 7 - 1) & ")"
        .Range(sRange_Sep & CStr(xlRow - 4)).NumberFormat = "_(#,#0.00_);_(#,#0.00_);_(-_)"
        
        .Range(sRange_Sep & CStr(xlRow - 3) & ":" & eRange_Sep & CStr(xlRow - 3)).Merge
        .Range(sRange_Sep & CStr(xlRow - 3)).Formula = "=sumif(" & .Cells(xlStartRowTask, xlLeftCol).Address & ":" & .Cells(xlRow - 7 - 1, xlLeftCol).Address & "," & .Cells(xlRow - 3, xlLeftCol + 3).Address & "," & sRange_Sep & CStr(xlStartRowTask) & ":" & eRange_Sep & CStr(xlRow - 7 - 1) & ")"
        .Range(sRange_Sep & CStr(xlRow - 3)).NumberFormat = "_(#,#0.00_);_(#,#0.00_);_(-_)"
        
        .Range(sRange_Sep & CStr(xlRow - 2) & ":" & eRange_Sep & CStr(xlRow - 2)).Merge 'Total Netto Task X - Sep
        .Range(sRange_Sep & CStr(xlRow - 2)).Formula = "=sum(" & sRange_Sep & CStr(xlRow - 3) & ":" & eRange_Sep & CStr(xlRow - 7) & ")"
        .Range(sRange_Sep & CStr(xlRow - 2)).NumberFormat = "_(#,#0.00_);_(#,#0.00_);_(-_)"
        
        .Range(sRange_Oct & CStr(xlRow - 7) & ":" & eRange_Oct & CStr(xlRow - 7)).Merge
        .Range(sRange_Oct & CStr(xlRow - 7)).Formula = "=sumif(" & .Cells(xlStartRowTask, xlLeftCol).Address & ":" & .Cells(xlRow - 7 - 1, xlLeftCol).Address & "," & .Cells(xlRow - 7, xlLeftCol + 3).Address & "," & sRange_Oct & CStr(xlStartRowTask) & ":" & eRange_Oct & CStr(xlRow - 7 - 1) & ")"
        .Range(sRange_Oct & CStr(xlRow - 7)).NumberFormat = "_(#,#0.00_);_(#,#0.00_);_(-_)"
        
        .Range(sRange_Oct & CStr(xlRow - 6) & ":" & eRange_Oct & CStr(xlRow - 6)).Merge
        .Range(sRange_Oct & CStr(xlRow - 6)).Formula = "=sumif(" & .Cells(xlStartRowTask, xlLeftCol).Address & ":" & .Cells(xlRow - 7 - 1, xlLeftCol).Address & "," & .Cells(xlRow - 6, xlLeftCol + 3).Address & "," & sRange_Oct & CStr(xlStartRowTask) & ":" & eRange_Oct & CStr(xlRow - 7 - 1) & ")"
        .Range(sRange_Oct & CStr(xlRow - 6)).NumberFormat = "_(#,#0.00_);_(#,#0.00_);_(-_)"
        
        .Range(sRange_Oct & CStr(xlRow - 5) & ":" & eRange_Oct & CStr(xlRow - 5)).Merge
        .Range(sRange_Oct & CStr(xlRow - 5)).Formula = "=sumif(" & .Cells(xlStartRowTask, xlLeftCol).Address & ":" & .Cells(xlRow - 7 - 1, xlLeftCol).Address & "," & .Cells(xlRow - 5, xlLeftCol + 3).Address & "," & sRange_Oct & CStr(xlStartRowTask) & ":" & eRange_Oct & CStr(xlRow - 7 - 1) & ")"
        .Range(sRange_Oct & CStr(xlRow - 5)).NumberFormat = "_(#,#0.00_);_(#,#0.00_);_(-_)"
        
        .Range(sRange_Oct & CStr(xlRow - 4) & ":" & eRange_Oct & CStr(xlRow - 4)).Merge
        .Range(sRange_Oct & CStr(xlRow - 4)).Formula = "=sumif(" & .Cells(xlStartRowTask, xlLeftCol).Address & ":" & .Cells(xlRow - 7 - 1, xlLeftCol).Address & "," & .Cells(xlRow - 4, xlLeftCol + 3).Address & "," & sRange_Oct & CStr(xlStartRowTask) & ":" & eRange_Oct & CStr(xlRow - 7 - 1) & ")"
        .Range(sRange_Oct & CStr(xlRow - 4)).NumberFormat = "_(#,#0.00_);_(#,#0.00_);_(-_)"
        
        .Range(sRange_Oct & CStr(xlRow - 3) & ":" & eRange_Oct & CStr(xlRow - 3)).Merge
        .Range(sRange_Oct & CStr(xlRow - 3)).Formula = "=sumif(" & .Cells(xlStartRowTask, xlLeftCol).Address & ":" & .Cells(xlRow - 7 - 1, xlLeftCol).Address & "," & .Cells(xlRow - 3, xlLeftCol + 3).Address & "," & sRange_Oct & CStr(xlStartRowTask) & ":" & eRange_Oct & CStr(xlRow - 7 - 1) & ")"
        .Range(sRange_Oct & CStr(xlRow - 3)).NumberFormat = "_(#,#0.00_);_(#,#0.00_);_(-_)"
        
        .Range(sRange_Oct & CStr(xlRow - 2) & ":" & eRange_Oct & CStr(xlRow - 2)).Merge 'Total Netto Task X - Oct
        .Range(sRange_Oct & CStr(xlRow - 2)).Formula = "=sum(" & sRange_Oct & CStr(xlRow - 3) & ":" & eRange_Oct & CStr(xlRow - 7) & ")"
        .Range(sRange_Oct & CStr(xlRow - 2)).NumberFormat = "_(#,#0.00_);_(#,#0.00_);_(-_)"
        
        .Range(sRange_Nov & CStr(xlRow - 7) & ":" & eRange_Nov & CStr(xlRow - 7)).Merge
        .Range(sRange_Nov & CStr(xlRow - 7)).Formula = "=sumif(" & .Cells(xlStartRowTask, xlLeftCol).Address & ":" & .Cells(xlRow - 7 - 1, xlLeftCol).Address & "," & .Cells(xlRow - 7, xlLeftCol + 3).Address & "," & sRange_Nov & CStr(xlStartRowTask) & ":" & eRange_Nov & CStr(xlRow - 7 - 1) & ")"
        .Range(sRange_Nov & CStr(xlRow - 7)).NumberFormat = "_(#,#0.00_);_(#,#0.00_);_(-_)"
        
        .Range(sRange_Nov & CStr(xlRow - 6) & ":" & eRange_Nov & CStr(xlRow - 6)).Merge
        .Range(sRange_Nov & CStr(xlRow - 6)).Formula = "=sumif(" & .Cells(xlStartRowTask, xlLeftCol).Address & ":" & .Cells(xlRow - 7 - 1, xlLeftCol).Address & "," & .Cells(xlRow - 6, xlLeftCol + 3).Address & "," & sRange_Nov & CStr(xlStartRowTask) & ":" & eRange_Nov & CStr(xlRow - 7 - 1) & ")"
        .Range(sRange_Nov & CStr(xlRow - 6)).NumberFormat = "_(#,#0.00_);_(#,#0.00_);_(-_)"
        
        .Range(sRange_Nov & CStr(xlRow - 5) & ":" & eRange_Nov & CStr(xlRow - 5)).Merge
        .Range(sRange_Nov & CStr(xlRow - 5)).Formula = "=sumif(" & .Cells(xlStartRowTask, xlLeftCol).Address & ":" & .Cells(xlRow - 7 - 1, xlLeftCol).Address & "," & .Cells(xlRow - 5, xlLeftCol + 3).Address & "," & sRange_Nov & CStr(xlStartRowTask) & ":" & eRange_Nov & CStr(xlRow - 7 - 1) & ")"
        .Range(sRange_Nov & CStr(xlRow - 5)).NumberFormat = "_(#,#0.00_);_(#,#0.00_);_(-_)"
        
        .Range(sRange_Nov & CStr(xlRow - 4) & ":" & eRange_Nov & CStr(xlRow - 4)).Merge
        .Range(sRange_Nov & CStr(xlRow - 4)).Formula = "=sumif(" & .Cells(xlStartRowTask, xlLeftCol).Address & ":" & .Cells(xlRow - 7 - 1, xlLeftCol).Address & "," & .Cells(xlRow - 4, xlLeftCol + 3).Address & "," & sRange_Nov & CStr(xlStartRowTask) & ":" & eRange_Nov & CStr(xlRow - 7 - 1) & ")"
        .Range(sRange_Nov & CStr(xlRow - 4)).NumberFormat = "_(#,#0.00_);_(#,#0.00_);_(-_)"
        
        .Range(sRange_Nov & CStr(xlRow - 3) & ":" & eRange_Nov & CStr(xlRow - 3)).Merge
        .Range(sRange_Nov & CStr(xlRow - 3)).Formula = "=sumif(" & .Cells(xlStartRowTask, xlLeftCol).Address & ":" & .Cells(xlRow - 7 - 1, xlLeftCol).Address & "," & .Cells(xlRow - 3, xlLeftCol + 3).Address & "," & sRange_Nov & CStr(xlStartRowTask) & ":" & eRange_Nov & CStr(xlRow - 7 - 1) & ")"
        .Range(sRange_Nov & CStr(xlRow - 3)).NumberFormat = "_(#,#0.00_);_(#,#0.00_);_(-_)"
        
        .Range(sRange_Nov & CStr(xlRow - 2) & ":" & eRange_Nov & CStr(xlRow - 2)).Merge 'Total Netto Task X - Nov
        .Range(sRange_Nov & CStr(xlRow - 2)).Formula = "=sum(" & sRange_Nov & CStr(xlRow - 3) & ":" & eRange_Nov & CStr(xlRow - 7) & ")"
        .Range(sRange_Nov & CStr(xlRow - 2)).NumberFormat = "_(#,#0.00_);_(#,#0.00_);_(-_)"
        
        .Range(sRange_Dec & CStr(xlRow - 7) & ":" & eRange_Dec & CStr(xlRow - 7)).Merge
        .Range(sRange_Dec & CStr(xlRow - 7)).Formula = "=sumif(" & .Cells(xlStartRowTask, xlLeftCol).Address & ":" & .Cells(xlRow - 7 - 1, xlLeftCol).Address & "," & .Cells(xlRow - 7, xlLeftCol + 3).Address & "," & sRange_Dec & CStr(xlStartRowTask) & ":" & eRange_Dec & CStr(xlRow - 7 - 1) & ")"
        .Range(sRange_Dec & CStr(xlRow - 7)).NumberFormat = "_(#,#0.00_);_(#,#0.00_);_(-_)"
        
        .Range(sRange_Dec & CStr(xlRow - 6) & ":" & eRange_Dec & CStr(xlRow - 6)).Merge
        .Range(sRange_Dec & CStr(xlRow - 6)).Formula = "=sumif(" & .Cells(xlStartRowTask, xlLeftCol).Address & ":" & .Cells(xlRow - 7 - 1, xlLeftCol).Address & "," & .Cells(xlRow - 6, xlLeftCol + 3).Address & "," & sRange_Dec & CStr(xlStartRowTask) & ":" & eRange_Dec & CStr(xlRow - 7 - 1) & ")"
        .Range(sRange_Dec & CStr(xlRow - 6)).NumberFormat = "_(#,#0.00_);_(#,#0.00_);_(-_)"
        
        .Range(sRange_Dec & CStr(xlRow - 5) & ":" & eRange_Dec & CStr(xlRow - 5)).Merge
        .Range(sRange_Dec & CStr(xlRow - 5)).Formula = "=sumif(" & .Cells(xlStartRowTask, xlLeftCol).Address & ":" & .Cells(xlRow - 7 - 1, xlLeftCol).Address & "," & .Cells(xlRow - 5, xlLeftCol + 3).Address & "," & sRange_Dec & CStr(xlStartRowTask) & ":" & eRange_Dec & CStr(xlRow - 7 - 1) & ")"
        .Range(sRange_Dec & CStr(xlRow - 5)).NumberFormat = "_(#,#0.00_);_(#,#0.00_);_(-_)"
        
        .Range(sRange_Dec & CStr(xlRow - 4) & ":" & eRange_Dec & CStr(xlRow - 4)).Merge
        .Range(sRange_Dec & CStr(xlRow - 4)).Formula = "=sumif(" & .Cells(xlStartRowTask, xlLeftCol).Address & ":" & .Cells(xlRow - 7 - 1, xlLeftCol).Address & "," & .Cells(xlRow - 4, xlLeftCol + 3).Address & "," & sRange_Dec & CStr(xlStartRowTask) & ":" & eRange_Dec & CStr(xlRow - 7 - 1) & ")"
        .Range(sRange_Dec & CStr(xlRow - 4)).NumberFormat = "_(#,#0.00_);_(#,#0.00_);_(-_)"
        
        .Range(sRange_Dec & CStr(xlRow - 3) & ":" & eRange_Dec & CStr(xlRow - 3)).Merge
        .Range(sRange_Dec & CStr(xlRow - 3)).Formula = "=sumif(" & .Cells(xlStartRowTask, xlLeftCol).Address & ":" & .Cells(xlRow - 7 - 1, xlLeftCol).Address & "," & .Cells(xlRow - 3, xlLeftCol + 3).Address & "," & sRange_Dec & CStr(xlStartRowTask) & ":" & eRange_Dec & CStr(xlRow - 7 - 1) & ")"
        .Range(sRange_Dec & CStr(xlRow - 3)).NumberFormat = "_(#,#0.00_);_(#,#0.00_);_(-_)"
        
        .Range(sRange_Dec & CStr(xlRow - 2) & ":" & eRange_Dec & CStr(xlRow - 2)).Merge 'Total Netto Task X - Dec
        .Range(sRange_Dec & CStr(xlRow - 2)).Formula = "=sum(" & sRange_Dec & CStr(xlRow - 3) & ":" & eRange_Dec & CStr(xlRow - 7) & ")"
        .Range(sRange_Dec & CStr(xlRow - 2)).NumberFormat = "_(#,#0.00_);_(#,#0.00_);_(-_)"
        
        .Cells(xlRow - 7, xlLeftCol + msf_MPInsertion.cols - 3).Formula = "=sumif(" & .Cells(xlStartRowTask, xlLeftCol).Address & ":" & .Cells(xlRow - 7 - 1, xlLeftCol).Address & "," & .Cells(xlRow - 7, xlLeftCol + 3).Address & "," & .Cells(xlStartRowTask, xlLeftCol + msf_MPInsertion.cols - 3).Address & ":" & .Cells(xlRow - 7 - 1, xlLeftCol + msf_MPInsertion.cols - 3).Address & ")"
        .Cells(xlRow - 7, xlLeftCol + msf_MPInsertion.cols - 3).NumberFormat = "_(#,#0.00_);_(#,#0.00_);_(-_)"
        
        .Cells(xlRow - 6, xlLeftCol + msf_MPInsertion.cols - 3).Formula = "=sumif(" & .Cells(xlStartRowTask, xlLeftCol).Address & ":" & .Cells(xlRow - 7 - 1, xlLeftCol).Address & "," & .Cells(xlRow - 6, xlLeftCol + 3).Address & "," & .Cells(xlStartRowTask, xlLeftCol + msf_MPInsertion.cols - 3).Address & ":" & .Cells(xlRow - 7 - 1, xlLeftCol + msf_MPInsertion.cols - 3).Address & ")"
        .Cells(xlRow - 6, xlLeftCol + msf_MPInsertion.cols - 3).NumberFormat = "_(#,#0.00_);_(#,#0.00_);_(-_)"
        
        .Cells(xlRow - 5, xlLeftCol + msf_MPInsertion.cols - 3).Formula = "=sumif(" & .Cells(xlStartRowTask, xlLeftCol).Address & ":" & .Cells(xlRow - 7 - 1, xlLeftCol).Address & "," & .Cells(xlRow - 5, xlLeftCol + 3).Address & "," & .Cells(xlStartRowTask, xlLeftCol + msf_MPInsertion.cols - 3).Address & ":" & .Cells(xlRow - 7 - 1, xlLeftCol + msf_MPInsertion.cols - 3).Address & ")"
        .Cells(xlRow - 5, xlLeftCol + msf_MPInsertion.cols - 3).NumberFormat = "_(#,#0.00_);_(#,#0.00_);_(-_)"
        
        .Cells(xlRow - 4, xlLeftCol + msf_MPInsertion.cols - 3).Formula = "=sumif(" & .Cells(xlStartRowTask, xlLeftCol).Address & ":" & .Cells(xlRow - 7 - 1, xlLeftCol).Address & "," & .Cells(xlRow - 4, xlLeftCol + 3).Address & "," & .Cells(xlStartRowTask, xlLeftCol + msf_MPInsertion.cols - 3).Address & ":" & .Cells(xlRow - 7 - 1, xlLeftCol + msf_MPInsertion.cols - 3).Address & ")"
        .Cells(xlRow - 4, xlLeftCol + msf_MPInsertion.cols - 3).NumberFormat = "_(#,#0.00_);_(#,#0.00_);_(-_)"
        
        .Cells(xlRow - 3, xlLeftCol + msf_MPInsertion.cols - 3).Formula = "=sumif(" & .Cells(xlStartRowTask, xlLeftCol).Address & ":" & .Cells(xlRow - 7 - 1, xlLeftCol).Address & "," & .Cells(xlRow - 3, xlLeftCol + 3).Address & "," & .Cells(xlStartRowTask, xlLeftCol + msf_MPInsertion.cols - 3).Address & ":" & .Cells(xlRow - 7 - 1, xlLeftCol + msf_MPInsertion.cols - 3).Address & ")"
        .Cells(xlRow - 3, xlLeftCol + msf_MPInsertion.cols - 3).NumberFormat = "_(#,#0.00_);_(#,#0.00_);_(-_)"
        
        .Cells(xlRow - 2, xlLeftCol + msf_MPInsertion.cols - 3).Formula = "=sum(" & .Cells(xlRow - 3, xlLeftCol + msf_MPInsertion.cols - 3).Address & ":" & .Cells(xlRow - 7, xlLeftCol + msf_MPInsertion.cols - 3).Address & ")"
        .Cells(xlRow - 2, xlLeftCol + msf_MPInsertion.cols - 3).NumberFormat = "_(#,#0.00_);_(#,#0.00_);_(-_)"
    
    End With
End Sub

Private Sub GrandTotal(xlws As Object, xlRow As Integer, xlLeftCol As Integer, xlTopRow As Integer, strActual)
'*****************************************************************************
' Nama Submodul         :  GrandTotal
' Fungsi Submodul       :  Print Grand Total
' Nama Programmer       :  Sistyo
' Tgl Pembuatan         :  18 Agustus 2004
' Last Update           :  18 Agustus 2004/Sistyo
'******************************************************************************
    If strActual = "Actual" Then
        strActual = " (Actual)"
    Else
        strActual = ""
    End If
    With xlws
        .Rows(CStr(xlRow) & ":" & CStr(xlRow + 7)).Insert Shift:=xlDown
        .Range(.Cells(xlRow + 1, xlLeftCol).Address & ":" & eRange_Dec & CStr(xlRow + 5)).Interior.ColorIndex = 2
        .Range(.Cells(xlRow + 1, xlLeftCol).Address & ":" & eRange_Dec & CStr(xlRow + 5)).Borders(xlEdgeTop).LineStyle = xlContinuous
        .Range(.Cells(xlRow + 1, xlLeftCol).Address & ":" & eRange_Dec & CStr(xlRow + 5)).Borders(xlEdgeTop).ColorIndex = xlAutomatic
        .Range(.Cells(xlRow + 1, xlLeftCol).Address & ":" & eRange_Dec & CStr(xlRow + 5)).Borders(xlEdgeTop).Weight = xlThin
        .Range(.Cells(xlRow + 1, xlLeftCol).Address & ":" & eRange_Dec & CStr(xlRow + 5)).Interior.ColorIndex = 2
        .Range(.Cells(xlRow + 1, xlLeftCol).Address & ":" & eRange_Dec & CStr(xlRow + 5)).Borders(xlEdgeBottom).LineStyle = xlContinuous
        .Range(.Cells(xlRow + 1, xlLeftCol).Address & ":" & eRange_Dec & CStr(xlRow + 5)).Borders(xlEdgeBottom).ColorIndex = xlAutomatic
        .Range(.Cells(xlRow + 1, xlLeftCol).Address & ":" & eRange_Dec & CStr(xlRow + 5)).Borders(xlEdgeBottom).Weight = xlThin
        .Range(.Cells(xlRow + 1, xlLeftCol).Address & ":" & eRange_Dec & CStr(xlRow + 5)).Interior.ColorIndex = 2
        .Range(.Cells(xlRow + 1, xlLeftCol).Address & ":" & eRange_Dec & CStr(xlRow + 5)).Borders(xlInsideHorizontal).LineStyle = xlContinuous
        .Range(.Cells(xlRow + 1, xlLeftCol).Address & ":" & eRange_Dec & CStr(xlRow + 5)).Borders(xlInsideHorizontal).ColorIndex = xlAutomatic
        .Range(.Cells(xlRow + 1, xlLeftCol).Address & ":" & eRange_Dec & CStr(xlRow + 5)).Borders(xlInsideHorizontal).Weight = xlThin
        
        .Range(.Cells(xlRow + 1, xlLeftCol).Address & ":" & eRange_Dec & CStr(xlRow + 8)).Font.Size = 9
        .Range(.Cells(xlRow + 1, xlLeftCol).Address & ":" & eRange_Dec & CStr(xlRow + 8)).Font.ColorIndex = 5
        .Range(.Cells(xlRow + 1, xlLeftCol).Address & ":" & eRange_Dec & CStr(xlRow + 8)).Font.Bold = True
        
        '.Range(.Cells(xlRow + 1, xlLeftCol).Address & ":" & .Cells(xlRow + 5, xlLeftCol).Address).Locked = True
        .Cells(xlRow + 1, xlLeftCol).Value = "Television" & strActual
        .Cells(xlRow + 2, xlLeftCol).Value = "Radio" & strActual
        .Cells(xlRow + 3, xlLeftCol).Value = "Print" & strActual
        .Cells(xlRow + 4, xlLeftCol).Value = "Cinema" & strActual
        .Cells(xlRow + 5, xlLeftCol).Value = "Other" & strActual
        .Cells(xlRow + 6, xlLeftCol + 4).Value = "Grand Total / Month" & strActual
        .Cells(xlRow + 7, xlLeftCol + 4).Value = "Grand Total / Quarter" & strActual
        
        '.Range(sRange_Jan & CStr(xlRow + 1) & ":" & eRange_Dec & CStr(xlRow + 5)).Locked = True
        
        'Border grand
        .Range(sRange_Jan & CStr(xlRow + 1) & ":" & .Cells(xlRow + 7, msf_MPInsertion.cols - 2 + xlLeftCol - 3).Address).Borders.LineStyle = xlContinuous
        .Range(sRange_Jan & CStr(xlRow + 1) & ":" & .Cells(xlRow + 7, msf_MPInsertion.cols - 2 + xlLeftCol - 3).Address).Borders.Weight = xlThick
        .Range(sRange_Jan & CStr(xlRow + 1) & ":" & .Cells(xlRow + 7, msf_MPInsertion.cols - 2 + xlLeftCol - 3).Address).Borders.ColorIndex = 5
        
        .Range(sRange_Jan & CStr(xlRow + 1) & ":" & eRange_Dec & CStr(xlRow + 8)).NumberFormat = "_(#,#0.00_);_(#,#0.00_);_(-_)"
        
        .Range(sRange_Jan & CStr(xlRow + 1) & ":" & eRange_Jan & CStr(xlRow + 1)).Merge
        .Range(sRange_Jan & CStr(xlRow + 1) & ":" & eRange_Jan & CStr(xlRow + 1)).Formula = "=sumif(" & .Cells(xlTopRow + 12, xlLeftCol).Address & ":" & .Cells(xlRow - 2, xlLeftCol).Address & "," & .Cells(xlRow + 1, xlLeftCol).Address & "," & sRange_Jan & CStr(xlTopRow + 12) & ":" & eRange_Jan & CStr(xlRow - 2) & ")"
        .Range(sRange_Jan & CStr(xlRow + 2) & ":" & eRange_Jan & CStr(xlRow + 2)).Merge
        .Range(sRange_Jan & CStr(xlRow + 2) & ":" & eRange_Jan & CStr(xlRow + 2)).Formula = "=sumif(" & .Cells(xlTopRow + 12, xlLeftCol).Address & ":" & .Cells(xlRow - 2, xlLeftCol).Address & "," & .Cells(xlRow + 2, xlLeftCol).Address & "," & sRange_Jan & CStr(xlTopRow + 12) & ":" & eRange_Jan & CStr(xlRow - 2) & ")"
        .Range(sRange_Jan & CStr(xlRow + 3) & ":" & eRange_Jan & CStr(xlRow + 3)).Merge
        .Range(sRange_Jan & CStr(xlRow + 3) & ":" & eRange_Jan & CStr(xlRow + 3)).Formula = "=sumif(" & .Cells(xlTopRow + 12, xlLeftCol).Address & ":" & .Cells(xlRow - 2, xlLeftCol).Address & "," & .Cells(xlRow + 3, xlLeftCol).Address & "," & sRange_Jan & CStr(xlTopRow + 12) & ":" & eRange_Jan & CStr(xlRow - 2) & ")"
        .Range(sRange_Jan & CStr(xlRow + 4) & ":" & eRange_Jan & CStr(xlRow + 4)).Merge
        .Range(sRange_Jan & CStr(xlRow + 4) & ":" & eRange_Jan & CStr(xlRow + 4)).Formula = "=sumif(" & .Cells(xlTopRow + 12, xlLeftCol).Address & ":" & .Cells(xlRow - 2, xlLeftCol).Address & "," & .Cells(xlRow + 4, xlLeftCol).Address & "," & sRange_Jan & CStr(xlTopRow + 12) & ":" & eRange_Jan & CStr(xlRow - 2) & ")"
        .Range(sRange_Jan & CStr(xlRow + 5) & ":" & eRange_Jan & CStr(xlRow + 5)).Merge
        .Range(sRange_Jan & CStr(xlRow + 5) & ":" & eRange_Jan & CStr(xlRow + 5)).Formula = "=sumif(" & .Cells(xlTopRow + 12, xlLeftCol).Address & ":" & .Cells(xlRow - 2, xlLeftCol).Address & "," & .Cells(xlRow + 5, xlLeftCol).Address & "," & sRange_Jan & CStr(xlTopRow + 12) & ":" & eRange_Jan & CStr(xlRow - 2) & ")"
        .Range(sRange_Jan & CStr(xlRow + 6) & ":" & eRange_Jan & CStr(xlRow + 6)).Merge
        .Range(sRange_Jan & CStr(xlRow + 6) & ":" & eRange_Jan & CStr(xlRow + 6)).Formula = "=sum(" & sRange_Jan & CStr(xlRow + 1) & ":" & eRange_Jan & CStr(xlRow + 5) & ")"
        
        .Range(sRange_Feb & CStr(xlRow + 1) & ":" & eRange_Feb & CStr(xlRow + 1)).Merge
        .Range(sRange_Feb & CStr(xlRow + 1) & ":" & eRange_Feb & CStr(xlRow + 1)).Formula = "=sumif(" & .Cells(xlTopRow + 12, xlLeftCol).Address & ":" & .Cells(xlRow - 2, xlLeftCol).Address & "," & .Cells(xlRow + 1, xlLeftCol).Address & "," & sRange_Feb & CStr(xlTopRow + 12) & ":" & eRange_Feb & CStr(xlRow - 2) & ")"
        .Range(sRange_Feb & CStr(xlRow + 2) & ":" & eRange_Feb & CStr(xlRow + 2)).Merge
        .Range(sRange_Feb & CStr(xlRow + 2) & ":" & eRange_Feb & CStr(xlRow + 2)).Formula = "=sumif(" & .Cells(xlTopRow + 12, xlLeftCol).Address & ":" & .Cells(xlRow - 2, xlLeftCol).Address & "," & .Cells(xlRow + 2, xlLeftCol).Address & "," & sRange_Feb & CStr(xlTopRow + 12) & ":" & eRange_Feb & CStr(xlRow - 2) & ")"
        .Range(sRange_Feb & CStr(xlRow + 3) & ":" & eRange_Feb & CStr(xlRow + 3)).Merge
        .Range(sRange_Feb & CStr(xlRow + 3) & ":" & eRange_Feb & CStr(xlRow + 3)).Formula = "=sumif(" & .Cells(xlTopRow + 12, xlLeftCol).Address & ":" & .Cells(xlRow - 2, xlLeftCol).Address & "," & .Cells(xlRow + 3, xlLeftCol).Address & "," & sRange_Feb & CStr(xlTopRow + 12) & ":" & eRange_Feb & CStr(xlRow - 2) & ")"
        .Range(sRange_Feb & CStr(xlRow + 4) & ":" & eRange_Feb & CStr(xlRow + 4)).Merge
        .Range(sRange_Feb & CStr(xlRow + 4) & ":" & eRange_Feb & CStr(xlRow + 4)).Formula = "=sumif(" & .Cells(xlTopRow + 12, xlLeftCol).Address & ":" & .Cells(xlRow - 2, xlLeftCol).Address & "," & .Cells(xlRow + 4, xlLeftCol).Address & "," & sRange_Feb & CStr(xlTopRow + 12) & ":" & eRange_Feb & CStr(xlRow - 2) & ")"
        .Range(sRange_Feb & CStr(xlRow + 5) & ":" & eRange_Feb & CStr(xlRow + 5)).Merge
        .Range(sRange_Feb & CStr(xlRow + 5) & ":" & eRange_Feb & CStr(xlRow + 5)).Formula = "=sumif(" & .Cells(xlTopRow + 12, xlLeftCol).Address & ":" & .Cells(xlRow - 2, xlLeftCol).Address & "," & .Cells(xlRow + 5, xlLeftCol).Address & "," & sRange_Feb & CStr(xlTopRow + 12) & ":" & eRange_Feb & CStr(xlRow - 2) & ")"
        .Range(sRange_Feb & CStr(xlRow + 6) & ":" & eRange_Feb & CStr(xlRow + 6)).Merge
        .Range(sRange_Feb & CStr(xlRow + 6) & ":" & eRange_Feb & CStr(xlRow + 6)).Formula = "=sum(" & sRange_Feb & CStr(xlRow + 1) & ":" & eRange_Feb & CStr(xlRow + 5) & ")"
        
        .Range(sRange_Mar & CStr(xlRow + 1) & ":" & eRange_Mar & CStr(xlRow + 1)).Merge
        .Range(sRange_Mar & CStr(xlRow + 1) & ":" & eRange_Mar & CStr(xlRow + 1)).Formula = "=sumif(" & .Cells(xlTopRow + 12, xlLeftCol).Address & ":" & .Cells(xlRow - 2, xlLeftCol).Address & "," & .Cells(xlRow + 1, xlLeftCol).Address & "," & sRange_Mar & CStr(xlTopRow + 12) & ":" & eRange_Mar & CStr(xlRow - 2) & ")"
        .Range(sRange_Mar & CStr(xlRow + 2) & ":" & eRange_Mar & CStr(xlRow + 2)).Merge
        .Range(sRange_Mar & CStr(xlRow + 2) & ":" & eRange_Mar & CStr(xlRow + 2)).Formula = "=sumif(" & .Cells(xlTopRow + 12, xlLeftCol).Address & ":" & .Cells(xlRow - 2, xlLeftCol).Address & "," & .Cells(xlRow + 2, xlLeftCol).Address & "," & sRange_Mar & CStr(xlTopRow + 12) & ":" & eRange_Mar & CStr(xlRow - 2) & ")"
        .Range(sRange_Mar & CStr(xlRow + 3) & ":" & eRange_Mar & CStr(xlRow + 3)).Merge
        .Range(sRange_Mar & CStr(xlRow + 3) & ":" & eRange_Mar & CStr(xlRow + 3)).Formula = "=sumif(" & .Cells(xlTopRow + 12, xlLeftCol).Address & ":" & .Cells(xlRow - 2, xlLeftCol).Address & "," & .Cells(xlRow + 3, xlLeftCol).Address & "," & sRange_Mar & CStr(xlTopRow + 12) & ":" & eRange_Mar & CStr(xlRow - 2) & ")"
        .Range(sRange_Mar & CStr(xlRow + 4) & ":" & eRange_Mar & CStr(xlRow + 4)).Merge
        .Range(sRange_Mar & CStr(xlRow + 4) & ":" & eRange_Mar & CStr(xlRow + 4)).Formula = "=sumif(" & .Cells(xlTopRow + 12, xlLeftCol).Address & ":" & .Cells(xlRow - 2, xlLeftCol).Address & "," & .Cells(xlRow + 4, xlLeftCol).Address & "," & sRange_Mar & CStr(xlTopRow + 12) & ":" & eRange_Mar & CStr(xlRow - 2) & ")"
        .Range(sRange_Mar & CStr(xlRow + 5) & ":" & eRange_Mar & CStr(xlRow + 5)).Merge
        .Range(sRange_Mar & CStr(xlRow + 5) & ":" & eRange_Mar & CStr(xlRow + 5)).Formula = "=sumif(" & .Cells(xlTopRow + 12, xlLeftCol).Address & ":" & .Cells(xlRow - 2, xlLeftCol).Address & "," & .Cells(xlRow + 5, xlLeftCol).Address & "," & sRange_Mar & CStr(xlTopRow + 12) & ":" & eRange_Mar & CStr(xlRow - 2) & ")"
        .Range(sRange_Mar & CStr(xlRow + 6) & ":" & eRange_Mar & CStr(xlRow + 6)).Merge
        .Range(sRange_Mar & CStr(xlRow + 6) & ":" & eRange_Mar & CStr(xlRow + 6)).Formula = "=sum(" & sRange_Mar & CStr(xlRow + 1) & ":" & eRange_Mar & CStr(xlRow + 5) & ")"
        
        .Range(sRange_Apr & CStr(xlRow + 1) & ":" & eRange_Apr & CStr(xlRow + 1)).Merge
        .Range(sRange_Apr & CStr(xlRow + 1) & ":" & eRange_Apr & CStr(xlRow + 1)).Formula = "=sumif(" & .Cells(xlTopRow + 12, xlLeftCol).Address & ":" & .Cells(xlRow - 2, xlLeftCol).Address & "," & .Cells(xlRow + 1, xlLeftCol).Address & "," & sRange_Apr & CStr(xlTopRow + 12) & ":" & eRange_Apr & CStr(xlRow - 2) & ")"
        .Range(sRange_Apr & CStr(xlRow + 2) & ":" & eRange_Apr & CStr(xlRow + 2)).Merge
        .Range(sRange_Apr & CStr(xlRow + 2) & ":" & eRange_Apr & CStr(xlRow + 2)).Formula = "=sumif(" & .Cells(xlTopRow + 12, xlLeftCol).Address & ":" & .Cells(xlRow - 2, xlLeftCol).Address & "," & .Cells(xlRow + 2, xlLeftCol).Address & "," & sRange_Apr & CStr(xlTopRow + 12) & ":" & eRange_Apr & CStr(xlRow - 2) & ")"
        .Range(sRange_Apr & CStr(xlRow + 3) & ":" & eRange_Apr & CStr(xlRow + 3)).Merge
        .Range(sRange_Apr & CStr(xlRow + 3) & ":" & eRange_Apr & CStr(xlRow + 3)).Formula = "=sumif(" & .Cells(xlTopRow + 12, xlLeftCol).Address & ":" & .Cells(xlRow - 2, xlLeftCol).Address & "," & .Cells(xlRow + 3, xlLeftCol).Address & "," & sRange_Apr & CStr(xlTopRow + 12) & ":" & eRange_Apr & CStr(xlRow - 2) & ")"
        .Range(sRange_Apr & CStr(xlRow + 4) & ":" & eRange_Apr & CStr(xlRow + 4)).Merge
        .Range(sRange_Apr & CStr(xlRow + 4) & ":" & eRange_Apr & CStr(xlRow + 4)).Formula = "=sumif(" & .Cells(xlTopRow + 12, xlLeftCol).Address & ":" & .Cells(xlRow - 2, xlLeftCol).Address & "," & .Cells(xlRow + 4, xlLeftCol).Address & "," & sRange_Apr & CStr(xlTopRow + 12) & ":" & eRange_Apr & CStr(xlRow - 2) & ")"
        .Range(sRange_Apr & CStr(xlRow + 5) & ":" & eRange_Apr & CStr(xlRow + 5)).Merge
        .Range(sRange_Apr & CStr(xlRow + 5) & ":" & eRange_Apr & CStr(xlRow + 5)).Formula = "=sumif(" & .Cells(xlTopRow + 12, xlLeftCol).Address & ":" & .Cells(xlRow - 2, xlLeftCol).Address & "," & .Cells(xlRow + 5, xlLeftCol).Address & "," & sRange_Apr & CStr(xlTopRow + 12) & ":" & eRange_Apr & CStr(xlRow - 2) & ")"
        .Range(sRange_Apr & CStr(xlRow + 6) & ":" & eRange_Apr & CStr(xlRow + 6)).Merge
        .Range(sRange_Apr & CStr(xlRow + 6) & ":" & eRange_Apr & CStr(xlRow + 6)).Formula = "=sum(" & sRange_Apr & CStr(xlRow + 1) & ":" & eRange_Apr & CStr(xlRow + 5) & ")"
        
        .Range(sRange_May & CStr(xlRow + 1) & ":" & eRange_May & CStr(xlRow + 1)).Merge
        .Range(sRange_May & CStr(xlRow + 1) & ":" & eRange_May & CStr(xlRow + 1)).Formula = "=sumif(" & .Cells(xlTopRow + 12, xlLeftCol).Address & ":" & .Cells(xlRow - 2, xlLeftCol).Address & "," & .Cells(xlRow + 1, xlLeftCol).Address & "," & sRange_May & CStr(xlTopRow + 12) & ":" & eRange_May & CStr(xlRow - 2) & ")"
        .Range(sRange_May & CStr(xlRow + 2) & ":" & eRange_May & CStr(xlRow + 2)).Merge
        .Range(sRange_May & CStr(xlRow + 2) & ":" & eRange_May & CStr(xlRow + 2)).Formula = "=sumif(" & .Cells(xlTopRow + 12, xlLeftCol).Address & ":" & .Cells(xlRow - 2, xlLeftCol).Address & "," & .Cells(xlRow + 2, xlLeftCol).Address & "," & sRange_May & CStr(xlTopRow + 12) & ":" & eRange_May & CStr(xlRow - 2) & ")"
        .Range(sRange_May & CStr(xlRow + 3) & ":" & eRange_May & CStr(xlRow + 3)).Merge
        .Range(sRange_May & CStr(xlRow + 3) & ":" & eRange_May & CStr(xlRow + 3)).Formula = "=sumif(" & .Cells(xlTopRow + 12, xlLeftCol).Address & ":" & .Cells(xlRow - 2, xlLeftCol).Address & "," & .Cells(xlRow + 3, xlLeftCol).Address & "," & sRange_May & CStr(xlTopRow + 12) & ":" & eRange_May & CStr(xlRow - 2) & ")"
        .Range(sRange_May & CStr(xlRow + 4) & ":" & eRange_May & CStr(xlRow + 4)).Merge
        .Range(sRange_May & CStr(xlRow + 4) & ":" & eRange_May & CStr(xlRow + 4)).Formula = "=sumif(" & .Cells(xlTopRow + 12, xlLeftCol).Address & ":" & .Cells(xlRow - 2, xlLeftCol).Address & "," & .Cells(xlRow + 4, xlLeftCol).Address & "," & sRange_May & CStr(xlTopRow + 12) & ":" & eRange_May & CStr(xlRow - 2) & ")"
        .Range(sRange_May & CStr(xlRow + 5) & ":" & eRange_May & CStr(xlRow + 5)).Merge
        .Range(sRange_May & CStr(xlRow + 5) & ":" & eRange_May & CStr(xlRow + 5)).Formula = "=sumif(" & .Cells(xlTopRow + 12, xlLeftCol).Address & ":" & .Cells(xlRow - 2, xlLeftCol).Address & "," & .Cells(xlRow + 5, xlLeftCol).Address & "," & sRange_May & CStr(xlTopRow + 12) & ":" & eRange_May & CStr(xlRow - 2) & ")"
        .Range(sRange_May & CStr(xlRow + 6) & ":" & eRange_May & CStr(xlRow + 6)).Merge
        .Range(sRange_May & CStr(xlRow + 6) & ":" & eRange_May & CStr(xlRow + 6)).Formula = "=sum(" & sRange_May & CStr(xlRow + 1) & ":" & eRange_May & CStr(xlRow + 5) & ")"
        
        .Range(sRange_Jun & CStr(xlRow + 1) & ":" & eRange_Jun & CStr(xlRow + 1)).Merge
        .Range(sRange_Jun & CStr(xlRow + 1) & ":" & eRange_Jun & CStr(xlRow + 1)).Formula = "=sumif(" & .Cells(xlTopRow + 12, xlLeftCol).Address & ":" & .Cells(xlRow - 2, xlLeftCol).Address & "," & .Cells(xlRow + 1, xlLeftCol).Address & "," & sRange_Jun & CStr(xlTopRow + 12) & ":" & eRange_Jun & CStr(xlRow - 2) & ")"
        .Range(sRange_Jun & CStr(xlRow + 2) & ":" & eRange_Jun & CStr(xlRow + 2)).Merge
        .Range(sRange_Jun & CStr(xlRow + 2) & ":" & eRange_Jun & CStr(xlRow + 2)).Formula = "=sumif(" & .Cells(xlTopRow + 12, xlLeftCol).Address & ":" & .Cells(xlRow - 2, xlLeftCol).Address & "," & .Cells(xlRow + 2, xlLeftCol).Address & "," & sRange_Jun & CStr(xlTopRow + 12) & ":" & eRange_Jun & CStr(xlRow - 2) & ")"
        .Range(sRange_Jun & CStr(xlRow + 3) & ":" & eRange_Jun & CStr(xlRow + 3)).Merge
        .Range(sRange_Jun & CStr(xlRow + 3) & ":" & eRange_Jun & CStr(xlRow + 3)).Formula = "=sumif(" & .Cells(xlTopRow + 12, xlLeftCol).Address & ":" & .Cells(xlRow - 2, xlLeftCol).Address & "," & .Cells(xlRow + 3, xlLeftCol).Address & "," & sRange_Jun & CStr(xlTopRow + 12) & ":" & eRange_Jun & CStr(xlRow - 2) & ")"
        .Range(sRange_Jun & CStr(xlRow + 4) & ":" & eRange_Jun & CStr(xlRow + 4)).Merge
        .Range(sRange_Jun & CStr(xlRow + 4) & ":" & eRange_Jun & CStr(xlRow + 4)).Formula = "=sumif(" & .Cells(xlTopRow + 12, xlLeftCol).Address & ":" & .Cells(xlRow - 2, xlLeftCol).Address & "," & .Cells(xlRow + 4, xlLeftCol).Address & "," & sRange_Jun & CStr(xlTopRow + 12) & ":" & eRange_Jun & CStr(xlRow - 2) & ")"
        .Range(sRange_Jun & CStr(xlRow + 5) & ":" & eRange_Jun & CStr(xlRow + 5)).Merge
        .Range(sRange_Jun & CStr(xlRow + 5) & ":" & eRange_Jun & CStr(xlRow + 5)).Formula = "=sumif(" & .Cells(xlTopRow + 12, xlLeftCol).Address & ":" & .Cells(xlRow - 2, xlLeftCol).Address & "," & .Cells(xlRow + 5, xlLeftCol).Address & "," & sRange_Jun & CStr(xlTopRow + 12) & ":" & eRange_Jun & CStr(xlRow - 2) & ")"
        .Range(sRange_Jun & CStr(xlRow + 6) & ":" & eRange_Jun & CStr(xlRow + 6)).Merge
        .Range(sRange_Jun & CStr(xlRow + 6) & ":" & eRange_Jun & CStr(xlRow + 6)).Formula = "=sum(" & sRange_Jun & CStr(xlRow + 1) & ":" & eRange_Jun & CStr(xlRow + 5) & ")"
        
        .Range(sRange_Jul & CStr(xlRow + 1) & ":" & eRange_Jul & CStr(xlRow + 1)).Merge
        .Range(sRange_Jul & CStr(xlRow + 1) & ":" & eRange_Jul & CStr(xlRow + 1)).Formula = "=sumif(" & .Cells(xlTopRow + 12, xlLeftCol).Address & ":" & .Cells(xlRow - 2, xlLeftCol).Address & "," & .Cells(xlRow + 1, xlLeftCol).Address & "," & sRange_Jul & CStr(xlTopRow + 12) & ":" & eRange_Jul & CStr(xlRow - 2) & ")"
        .Range(sRange_Jul & CStr(xlRow + 2) & ":" & eRange_Jul & CStr(xlRow + 2)).Merge
        .Range(sRange_Jul & CStr(xlRow + 2) & ":" & eRange_Jul & CStr(xlRow + 2)).Formula = "=sumif(" & .Cells(xlTopRow + 12, xlLeftCol).Address & ":" & .Cells(xlRow - 2, xlLeftCol).Address & "," & .Cells(xlRow + 2, xlLeftCol).Address & "," & sRange_Jul & CStr(xlTopRow + 12) & ":" & eRange_Jul & CStr(xlRow - 2) & ")"
        .Range(sRange_Jul & CStr(xlRow + 3) & ":" & eRange_Jul & CStr(xlRow + 3)).Merge
        .Range(sRange_Jul & CStr(xlRow + 3) & ":" & eRange_Jul & CStr(xlRow + 3)).Formula = "=sumif(" & .Cells(xlTopRow + 12, xlLeftCol).Address & ":" & .Cells(xlRow - 2, xlLeftCol).Address & "," & .Cells(xlRow + 3, xlLeftCol).Address & "," & sRange_Jul & CStr(xlTopRow + 12) & ":" & eRange_Jul & CStr(xlRow - 2) & ")"
        .Range(sRange_Jul & CStr(xlRow + 4) & ":" & eRange_Jul & CStr(xlRow + 4)).Merge
        .Range(sRange_Jul & CStr(xlRow + 4) & ":" & eRange_Jul & CStr(xlRow + 4)).Formula = "=sumif(" & .Cells(xlTopRow + 12, xlLeftCol).Address & ":" & .Cells(xlRow - 2, xlLeftCol).Address & "," & .Cells(xlRow + 4, xlLeftCol).Address & "," & sRange_Jul & CStr(xlTopRow + 12) & ":" & eRange_Jul & CStr(xlRow - 2) & ")"
        .Range(sRange_Jul & CStr(xlRow + 5) & ":" & eRange_Jul & CStr(xlRow + 5)).Merge
        .Range(sRange_Jul & CStr(xlRow + 5) & ":" & eRange_Jul & CStr(xlRow + 5)).Formula = "=sumif(" & .Cells(xlTopRow + 12, xlLeftCol).Address & ":" & .Cells(xlRow - 2, xlLeftCol).Address & "," & .Cells(xlRow + 5, xlLeftCol).Address & "," & sRange_Jul & CStr(xlTopRow + 12) & ":" & eRange_Jul & CStr(xlRow - 2) & ")"
        .Range(sRange_Jul & CStr(xlRow + 6) & ":" & eRange_Jul & CStr(xlRow + 6)).Merge
        .Range(sRange_Jul & CStr(xlRow + 6) & ":" & eRange_Jul & CStr(xlRow + 6)).Formula = "=sum(" & sRange_Jul & CStr(xlRow + 1) & ":" & eRange_Jul & CStr(xlRow + 5) & ")"
        
        .Range(sRange_Aug & CStr(xlRow + 1) & ":" & eRange_Aug & CStr(xlRow + 1)).Merge
        .Range(sRange_Aug & CStr(xlRow + 1) & ":" & eRange_Aug & CStr(xlRow + 1)).Formula = "=sumif(" & .Cells(xlTopRow + 12, xlLeftCol).Address & ":" & .Cells(xlRow - 2, xlLeftCol).Address & "," & .Cells(xlRow + 1, xlLeftCol).Address & "," & sRange_Aug & CStr(xlTopRow + 12) & ":" & eRange_Aug & CStr(xlRow - 2) & ")"
        .Range(sRange_Aug & CStr(xlRow + 2) & ":" & eRange_Aug & CStr(xlRow + 2)).Merge
        .Range(sRange_Aug & CStr(xlRow + 2) & ":" & eRange_Aug & CStr(xlRow + 2)).Formula = "=sumif(" & .Cells(xlTopRow + 12, xlLeftCol).Address & ":" & .Cells(xlRow - 2, xlLeftCol).Address & "," & .Cells(xlRow + 2, xlLeftCol).Address & "," & sRange_Aug & CStr(xlTopRow + 12) & ":" & eRange_Aug & CStr(xlRow - 2) & ")"
        .Range(sRange_Aug & CStr(xlRow + 3) & ":" & eRange_Aug & CStr(xlRow + 3)).Merge
        .Range(sRange_Aug & CStr(xlRow + 3) & ":" & eRange_Aug & CStr(xlRow + 3)).Formula = "=sumif(" & .Cells(xlTopRow + 12, xlLeftCol).Address & ":" & .Cells(xlRow - 2, xlLeftCol).Address & "," & .Cells(xlRow + 3, xlLeftCol).Address & "," & sRange_Aug & CStr(xlTopRow + 12) & ":" & eRange_Aug & CStr(xlRow - 2) & ")"
        .Range(sRange_Aug & CStr(xlRow + 4) & ":" & eRange_Aug & CStr(xlRow + 4)).Merge
        .Range(sRange_Aug & CStr(xlRow + 4) & ":" & eRange_Aug & CStr(xlRow + 4)).Formula = "=sumif(" & .Cells(xlTopRow + 12, xlLeftCol).Address & ":" & .Cells(xlRow - 2, xlLeftCol).Address & "," & .Cells(xlRow + 4, xlLeftCol).Address & "," & sRange_Aug & CStr(xlTopRow + 12) & ":" & eRange_Aug & CStr(xlRow - 2) & ")"
        .Range(sRange_Aug & CStr(xlRow + 5) & ":" & eRange_Aug & CStr(xlRow + 5)).Merge
        .Range(sRange_Aug & CStr(xlRow + 5) & ":" & eRange_Aug & CStr(xlRow + 5)).Formula = "=sumif(" & .Cells(xlTopRow + 12, xlLeftCol).Address & ":" & .Cells(xlRow - 2, xlLeftCol).Address & "," & .Cells(xlRow + 5, xlLeftCol).Address & "," & sRange_Aug & CStr(xlTopRow + 12) & ":" & eRange_Aug & CStr(xlRow - 2) & ")"
        .Range(sRange_Aug & CStr(xlRow + 6) & ":" & eRange_Aug & CStr(xlRow + 6)).Merge
        .Range(sRange_Aug & CStr(xlRow + 6) & ":" & eRange_Aug & CStr(xlRow + 6)).Formula = "=sum(" & sRange_Aug & CStr(xlRow + 1) & ":" & eRange_Aug & CStr(xlRow + 5) & ")"
        
        .Range(sRange_Sep & CStr(xlRow + 1) & ":" & eRange_Sep & CStr(xlRow + 1)).Merge
        .Range(sRange_Sep & CStr(xlRow + 1) & ":" & eRange_Sep & CStr(xlRow + 1)).Formula = "=sumif(" & .Cells(xlTopRow + 12, xlLeftCol).Address & ":" & .Cells(xlRow - 2, xlLeftCol).Address & "," & .Cells(xlRow + 1, xlLeftCol).Address & "," & sRange_Sep & CStr(xlTopRow + 12) & ":" & eRange_Sep & CStr(xlRow - 2) & ")"
        .Range(sRange_Sep & CStr(xlRow + 2) & ":" & eRange_Sep & CStr(xlRow + 2)).Merge
        .Range(sRange_Sep & CStr(xlRow + 2) & ":" & eRange_Sep & CStr(xlRow + 2)).Formula = "=sumif(" & .Cells(xlTopRow + 12, xlLeftCol).Address & ":" & .Cells(xlRow - 2, xlLeftCol).Address & "," & .Cells(xlRow + 2, xlLeftCol).Address & "," & sRange_Sep & CStr(xlTopRow + 12) & ":" & eRange_Sep & CStr(xlRow - 2) & ")"
        .Range(sRange_Sep & CStr(xlRow + 3) & ":" & eRange_Sep & CStr(xlRow + 3)).Merge
        .Range(sRange_Sep & CStr(xlRow + 3) & ":" & eRange_Sep & CStr(xlRow + 3)).Formula = "=sumif(" & .Cells(xlTopRow + 12, xlLeftCol).Address & ":" & .Cells(xlRow - 2, xlLeftCol).Address & "," & .Cells(xlRow + 3, xlLeftCol).Address & "," & sRange_Sep & CStr(xlTopRow + 12) & ":" & eRange_Sep & CStr(xlRow - 2) & ")"
        .Range(sRange_Sep & CStr(xlRow + 4) & ":" & eRange_Sep & CStr(xlRow + 4)).Merge
        .Range(sRange_Sep & CStr(xlRow + 4) & ":" & eRange_Sep & CStr(xlRow + 4)).Formula = "=sumif(" & .Cells(xlTopRow + 12, xlLeftCol).Address & ":" & .Cells(xlRow - 2, xlLeftCol).Address & "," & .Cells(xlRow + 4, xlLeftCol).Address & "," & sRange_Sep & CStr(xlTopRow + 12) & ":" & eRange_Sep & CStr(xlRow - 2) & ")"
        .Range(sRange_Sep & CStr(xlRow + 5) & ":" & eRange_Sep & CStr(xlRow + 5)).Merge
        .Range(sRange_Sep & CStr(xlRow + 5) & ":" & eRange_Sep & CStr(xlRow + 5)).Formula = "=sumif(" & .Cells(xlTopRow + 12, xlLeftCol).Address & ":" & .Cells(xlRow - 2, xlLeftCol).Address & "," & .Cells(xlRow + 5, xlLeftCol).Address & "," & sRange_Sep & CStr(xlTopRow + 12) & ":" & eRange_Sep & CStr(xlRow - 2) & ")"
        .Range(sRange_Sep & CStr(xlRow + 6) & ":" & eRange_Sep & CStr(xlRow + 6)).Merge
        .Range(sRange_Sep & CStr(xlRow + 6) & ":" & eRange_Sep & CStr(xlRow + 6)).Formula = "=sum(" & sRange_Sep & CStr(xlRow + 1) & ":" & eRange_Sep & CStr(xlRow + 5) & ")"
        
        .Range(sRange_Oct & CStr(xlRow + 1) & ":" & eRange_Oct & CStr(xlRow + 1)).Merge
        .Range(sRange_Oct & CStr(xlRow + 1) & ":" & eRange_Oct & CStr(xlRow + 1)).Formula = "=sumif(" & .Cells(xlTopRow + 12, xlLeftCol).Address & ":" & .Cells(xlRow - 2, xlLeftCol).Address & "," & .Cells(xlRow + 1, xlLeftCol).Address & "," & sRange_Oct & CStr(xlTopRow + 12) & ":" & eRange_Oct & CStr(xlRow - 2) & ")"
        .Range(sRange_Oct & CStr(xlRow + 2) & ":" & eRange_Oct & CStr(xlRow + 2)).Merge
        .Range(sRange_Oct & CStr(xlRow + 2) & ":" & eRange_Oct & CStr(xlRow + 2)).Formula = "=sumif(" & .Cells(xlTopRow + 12, xlLeftCol).Address & ":" & .Cells(xlRow - 2, xlLeftCol).Address & "," & .Cells(xlRow + 2, xlLeftCol).Address & "," & sRange_Oct & CStr(xlTopRow + 12) & ":" & eRange_Oct & CStr(xlRow - 2) & ")"
        .Range(sRange_Oct & CStr(xlRow + 3) & ":" & eRange_Oct & CStr(xlRow + 3)).Merge
        .Range(sRange_Oct & CStr(xlRow + 3) & ":" & eRange_Oct & CStr(xlRow + 3)).Formula = "=sumif(" & .Cells(xlTopRow + 12, xlLeftCol).Address & ":" & .Cells(xlRow - 2, xlLeftCol).Address & "," & .Cells(xlRow + 3, xlLeftCol).Address & "," & sRange_Oct & CStr(xlTopRow + 12) & ":" & eRange_Oct & CStr(xlRow - 2) & ")"
        .Range(sRange_Oct & CStr(xlRow + 4) & ":" & eRange_Oct & CStr(xlRow + 4)).Merge
        .Range(sRange_Oct & CStr(xlRow + 4) & ":" & eRange_Oct & CStr(xlRow + 4)).Formula = "=sumif(" & .Cells(xlTopRow + 12, xlLeftCol).Address & ":" & .Cells(xlRow - 2, xlLeftCol).Address & "," & .Cells(xlRow + 4, xlLeftCol).Address & "," & sRange_Oct & CStr(xlTopRow + 12) & ":" & eRange_Oct & CStr(xlRow - 2) & ")"
        .Range(sRange_Oct & CStr(xlRow + 5) & ":" & eRange_Oct & CStr(xlRow + 5)).Merge
        .Range(sRange_Oct & CStr(xlRow + 5) & ":" & eRange_Oct & CStr(xlRow + 5)).Formula = "=sumif(" & .Cells(xlTopRow + 12, xlLeftCol).Address & ":" & .Cells(xlRow - 2, xlLeftCol).Address & "," & .Cells(xlRow + 5, xlLeftCol).Address & "," & sRange_Oct & CStr(xlTopRow + 12) & ":" & eRange_Oct & CStr(xlRow - 2) & ")"
        .Range(sRange_Oct & CStr(xlRow + 6) & ":" & eRange_Oct & CStr(xlRow + 6)).Merge
        .Range(sRange_Oct & CStr(xlRow + 6) & ":" & eRange_Oct & CStr(xlRow + 6)).Formula = "=sum(" & sRange_Oct & CStr(xlRow + 1) & ":" & eRange_Oct & CStr(xlRow + 5) & ")"
        
        .Range(sRange_Nov & CStr(xlRow + 1) & ":" & eRange_Nov & CStr(xlRow + 1)).Merge
        .Range(sRange_Nov & CStr(xlRow + 1) & ":" & eRange_Nov & CStr(xlRow + 1)).Formula = "=sumif(" & .Cells(xlTopRow + 12, xlLeftCol).Address & ":" & .Cells(xlRow - 2, xlLeftCol).Address & "," & .Cells(xlRow + 1, xlLeftCol).Address & "," & sRange_Nov & CStr(xlTopRow + 12) & ":" & eRange_Nov & CStr(xlRow - 2) & ")"
        .Range(sRange_Nov & CStr(xlRow + 2) & ":" & eRange_Nov & CStr(xlRow + 2)).Merge
        .Range(sRange_Nov & CStr(xlRow + 2) & ":" & eRange_Nov & CStr(xlRow + 2)).Formula = "=sumif(" & .Cells(xlTopRow + 12, xlLeftCol).Address & ":" & .Cells(xlRow - 2, xlLeftCol).Address & "," & .Cells(xlRow + 2, xlLeftCol).Address & "," & sRange_Nov & CStr(xlTopRow + 12) & ":" & eRange_Nov & CStr(xlRow - 2) & ")"
        .Range(sRange_Nov & CStr(xlRow + 3) & ":" & eRange_Nov & CStr(xlRow + 3)).Merge
        .Range(sRange_Nov & CStr(xlRow + 3) & ":" & eRange_Nov & CStr(xlRow + 3)).Formula = "=sumif(" & .Cells(xlTopRow + 12, xlLeftCol).Address & ":" & .Cells(xlRow - 2, xlLeftCol).Address & "," & .Cells(xlRow + 3, xlLeftCol).Address & "," & sRange_Nov & CStr(xlTopRow + 12) & ":" & eRange_Nov & CStr(xlRow - 2) & ")"
        .Range(sRange_Nov & CStr(xlRow + 4) & ":" & eRange_Nov & CStr(xlRow + 4)).Merge
        .Range(sRange_Nov & CStr(xlRow + 4) & ":" & eRange_Nov & CStr(xlRow + 4)).Formula = "=sumif(" & .Cells(xlTopRow + 12, xlLeftCol).Address & ":" & .Cells(xlRow - 2, xlLeftCol).Address & "," & .Cells(xlRow + 4, xlLeftCol).Address & "," & sRange_Nov & CStr(xlTopRow + 12) & ":" & eRange_Nov & CStr(xlRow - 2) & ")"
        .Range(sRange_Nov & CStr(xlRow + 5) & ":" & eRange_Nov & CStr(xlRow + 5)).Merge
        .Range(sRange_Nov & CStr(xlRow + 5) & ":" & eRange_Nov & CStr(xlRow + 5)).Formula = "=sumif(" & .Cells(xlTopRow + 12, xlLeftCol).Address & ":" & .Cells(xlRow - 2, xlLeftCol).Address & "," & .Cells(xlRow + 5, xlLeftCol).Address & "," & sRange_Nov & CStr(xlTopRow + 12) & ":" & eRange_Nov & CStr(xlRow - 2) & ")"
        .Range(sRange_Nov & CStr(xlRow + 6) & ":" & eRange_Nov & CStr(xlRow + 6)).Merge
        .Range(sRange_Nov & CStr(xlRow + 6) & ":" & eRange_Nov & CStr(xlRow + 6)).Formula = "=sum(" & sRange_Nov & CStr(xlRow + 1) & ":" & eRange_Nov & CStr(xlRow + 5) & ")"
        
        .Range(sRange_Dec & CStr(xlRow + 1) & ":" & eRange_Dec & CStr(xlRow + 1)).Merge
        .Range(sRange_Dec & CStr(xlRow + 1) & ":" & eRange_Dec & CStr(xlRow + 1)).Formula = "=sumif(" & .Cells(xlTopRow + 12, xlLeftCol).Address & ":" & .Cells(xlRow - 2, xlLeftCol).Address & "," & .Cells(xlRow + 1, xlLeftCol).Address & "," & sRange_Dec & CStr(xlTopRow + 12) & ":" & eRange_Dec & CStr(xlRow - 2) & ")"
        .Range(sRange_Dec & CStr(xlRow + 2) & ":" & eRange_Dec & CStr(xlRow + 2)).Merge
        .Range(sRange_Dec & CStr(xlRow + 2) & ":" & eRange_Dec & CStr(xlRow + 2)).Formula = "=sumif(" & .Cells(xlTopRow + 12, xlLeftCol).Address & ":" & .Cells(xlRow - 2, xlLeftCol).Address & "," & .Cells(xlRow + 2, xlLeftCol).Address & "," & sRange_Dec & CStr(xlTopRow + 12) & ":" & eRange_Dec & CStr(xlRow - 2) & ")"
        .Range(sRange_Dec & CStr(xlRow + 3) & ":" & eRange_Dec & CStr(xlRow + 3)).Merge
        .Range(sRange_Dec & CStr(xlRow + 3) & ":" & eRange_Dec & CStr(xlRow + 3)).Formula = "=sumif(" & .Cells(xlTopRow + 12, xlLeftCol).Address & ":" & .Cells(xlRow - 2, xlLeftCol).Address & "," & .Cells(xlRow + 3, xlLeftCol).Address & "," & sRange_Dec & CStr(xlTopRow + 12) & ":" & eRange_Dec & CStr(xlRow - 2) & ")"
        .Range(sRange_Dec & CStr(xlRow + 4) & ":" & eRange_Dec & CStr(xlRow + 4)).Merge
        .Range(sRange_Dec & CStr(xlRow + 4) & ":" & eRange_Dec & CStr(xlRow + 4)).Formula = "=sumif(" & .Cells(xlTopRow + 12, xlLeftCol).Address & ":" & .Cells(xlRow - 2, xlLeftCol).Address & "," & .Cells(xlRow + 4, xlLeftCol).Address & "," & sRange_Dec & CStr(xlTopRow + 12) & ":" & eRange_Dec & CStr(xlRow - 2) & ")"
        .Range(sRange_Dec & CStr(xlRow + 5) & ":" & eRange_Dec & CStr(xlRow + 5)).Merge
        .Range(sRange_Dec & CStr(xlRow + 5) & ":" & eRange_Dec & CStr(xlRow + 5)).Formula = "=sumif(" & .Cells(xlTopRow + 12, xlLeftCol).Address & ":" & .Cells(xlRow - 2, xlLeftCol).Address & "," & .Cells(xlRow + 5, xlLeftCol).Address & "," & sRange_Dec & CStr(xlTopRow + 12) & ":" & eRange_Dec & CStr(xlRow - 2) & ")"
        .Range(sRange_Dec & CStr(xlRow + 6) & ":" & eRange_Dec & CStr(xlRow + 6)).Merge
        .Range(sRange_Dec & CStr(xlRow + 6) & ":" & eRange_Dec & CStr(xlRow + 6)).Formula = "=sum(" & sRange_Dec & CStr(xlRow + 1) & ":" & eRange_Dec & CStr(xlRow + 5) & ")"
            
        'Total per Quarter
        .Range(.Cells(xlRow + 7, xlLeftCol).Address & ":" & .Cells(xlRow + 9, msf_MPInsertion.cols - 2 + xlLeftCol - 1).Address).Borders(xlInsideVertical).LineStyle = xlNone
            
        .Range(sRange_Jan & CStr(xlRow + 7) & ":" & .Cells(xlRow + 7, msf_MPInsertion.cols - 2 + xlLeftCol - 1).Address).Borders(xlEdgeTop).LineStyle = xlContinuous
        .Range(sRange_Jan & CStr(xlRow + 7) & ":" & .Cells(xlRow + 7, msf_MPInsertion.cols - 2 + xlLeftCol - 1).Address).Borders(xlEdgeTop).Weight = xlThick
        .Range(sRange_Jan & CStr(xlRow + 7) & ":" & .Cells(xlRow + 7, msf_MPInsertion.cols - 2 + xlLeftCol - 1).Address).Borders(xlEdgeTop).ColorIndex = 5
        
        .Range(sRange_Jan & CStr(xlRow + 7) & ":" & eRange_Mar & CStr(xlRow + 7)).Merge 'Q1
        .Range(sRange_Jan & CStr(xlRow + 7) & ":" & eRange_Mar & CStr(xlRow + 7)).Formula = "=sum(" & sRange_Jan & CStr(xlRow + 1) & ":" & eRange_Mar & CStr(xlRow + 5) & ")"
        .Range(sRange_Apr & CStr(xlRow + 7) & ":" & eRange_Jun & CStr(xlRow + 7)).Merge 'Q2
        .Range(sRange_Apr & CStr(xlRow + 7) & ":" & eRange_Jun & CStr(xlRow + 7)).Formula = "=sum(" & sRange_Apr & CStr(xlRow + 1) & ":" & eRange_Jun & CStr(xlRow + 5) & ")"
        .Range(sRange_Jul & CStr(xlRow + 7) & ":" & eRange_Sep & CStr(xlRow + 7)).Merge 'Q3
        .Range(sRange_Jul & CStr(xlRow + 7) & ":" & eRange_Sep & CStr(xlRow + 7)).Formula = "=sum(" & sRange_Jul & CStr(xlRow + 1) & ":" & eRange_Sep & CStr(xlRow + 5) & ")"
        .Range(sRange_Oct & CStr(xlRow + 7) & ":" & eRange_Dec & CStr(xlRow + 7)).Merge 'Q4
        .Range(sRange_Oct & CStr(xlRow + 7) & ":" & eRange_Dec & CStr(xlRow + 7)).Formula = "=sum(" & sRange_Oct & CStr(xlRow + 1) & ":" & eRange_Dec & CStr(xlRow + 5) & ")"
        
        'Grand Total 1 tahun
        .Cells(xlRow + 7, xlLeftCol + msf_MPInsertion.TextMatrix(3, msf_MPInsertion.cols - 4) + 6).Formula = "=sum(" & sRange_Jan & CStr(xlRow + 7) & ":" & eRange_Dec & CStr(xlRow + 7) & ")"
        
        
        .Range(sRange_Jan & CStr(xlRow + 7) & ":" & eRange_Dec & CStr(xlRow + 7)).Borders(xlEdgeBottom).LineStyle = xlContinuous
        .Range(sRange_Jan & CStr(xlRow + 7) & ":" & eRange_Dec & CStr(xlRow + 7)).Borders(xlEdgeBottom).Weight = xlThick
        .Range(sRange_Jan & CStr(xlRow + 7) & ":" & eRange_Dec & CStr(xlRow + 7)).Borders(xlEdgeBottom).ColorIndex = 5
        .Range(sRange_Jan & CStr(xlRow + 7) & ":" & eRange_Dec & CStr(xlRow + 7)).Borders(xlEdgeLeft).LineStyle = xlContinuous
        .Range(sRange_Jan & CStr(xlRow + 7) & ":" & eRange_Dec & CStr(xlRow + 7)).Borders(xlEdgeLeft).Weight = xlThick
        .Range(sRange_Jan & CStr(xlRow + 7) & ":" & eRange_Dec & CStr(xlRow + 7)).Borders(xlEdgeLeft).ColorIndex = 5
        .Range(sRange_Jan & CStr(xlRow + 7) & ":" & eRange_Dec & CStr(xlRow + 7)).Borders(xlEdgeRight).LineStyle = xlContinuous
        .Range(sRange_Jan & CStr(xlRow + 7) & ":" & eRange_Dec & CStr(xlRow + 7)).Borders(xlEdgeRight).Weight = xlThick
        .Range(sRange_Jan & CStr(xlRow + 7) & ":" & eRange_Dec & CStr(xlRow + 7)).Borders(xlEdgeRight).ColorIndex = 5
        .Range(sRange_Jan & CStr(xlRow + 7) & ":" & eRange_Dec & CStr(xlRow + 7)).Borders(xlInsideVertical).LineStyle = xlContinuous
        .Range(sRange_Jan & CStr(xlRow + 7) & ":" & eRange_Dec & CStr(xlRow + 7)).Borders(xlInsideVertical).Weight = xlThick
        .Range(sRange_Jan & CStr(xlRow + 7) & ":" & eRange_Dec & CStr(xlRow + 7)).Borders(xlInsideVertical).ColorIndex = 5
    End With
End Sub

Private Sub txtTVREACH_Change()
    If Val(txtTVREACH) > 100 Then
        txtTVREACH.Text = Mid(txtTVREACH.Text, 1, 2)
        txtTVREACH.SelStart = 2
    End If
End Sub

Private Sub txtTVREACH_KeyDown(KeyCode As Integer, Shift As Integer)

    Select Case KeyCode
        Case 13 'Enter
            msf_MPInsertion.SetFocus
        Case 27 'Escape
            txtTVREACH.Visible = False
            msf_MPInsertion.SetFocus
        Case 37 'panah kiri
            If Not EditMode Then
                msf_MPInsertion.SetFocus
                If intCol > msf_MPInsertion.FixedCols Then
                    msf_MPInsertion.col = intCol - 1
                End If
            End If
        Case 38 'panah atas
            msf_MPInsertion.SetFocus
            If intRow > msf_MPInsertion.FixedRows + 1 Then
                msf_MPInsertion.Row = intRow - 1
            End If
        Case 39 'panah kanan
            If Not EditMode Then
                msf_MPInsertion.SetFocus
                If intCol < msf_MPInsertion.cols - 2 Then
                    msf_MPInsertion.col = intCol + 1
                End If
            End If
        Case 40 'panah bawah
            msf_MPInsertion.SetFocus
            If intRow < msf_MPInsertion.Rows - 1 Then
                msf_MPInsertion.Row = intRow + 1
            End If
    End Select

End Sub

Private Sub txtTVREACH_KeyPress(KeyAscii As Integer)
'Input harus angka
    If (Chr(KeyAscii) < "0" Or Chr(KeyAscii) > "9") And KeyAscii <> 8 And KeyAscii <> 13 And KeyAscii <> 27 Then
            KeyAscii = 0
            Beep
    End If
End Sub

Private Sub txtTVREACH_LostFocus()
    Dim intWeekStart As Integer
    Dim intWeekEnd As Integer
    Dim intMp_tv_rf_id As Double
    Dim flagExist As Boolean
    Dim nSpace As Integer
    Dim i As Integer
    
    flagExist = False
    
    If txtTVREACH.Visible Then
        txtTVREACH.Visible = False
        rsTemp.Open "select mp_tv_rf_id,week_year_start,week_year_end from mp_tv_reach_frequency where mp_plan_dim_id='" & Left(msf_MPInsertion.TextMatrix(intRow, msf_MPInsertion.cols - 1), 19) & "' and week_year_start<=" & msf_MPInsertion.TextMatrix(3, intCol) & " and week_year_end>=" & msf_MPInsertion.TextMatrix(3, intCol), ConnERP, 1, 3
        If Not rsTemp.EOF Then
            flagExist = True
            intMp_tv_rf_id = rsTemp(0)
            intWeekStart = rsTemp(1)
            intWeekEnd = rsTemp(2)
        End If
        rsTemp.Close
        nSpace = 1
        While Mid(msf_MPInsertion.TextMatrix(intRow, intCol), nSpace, 1) = " "
            nSpace = nSpace + 1
        Wend
        nSpace = nSpace - 1
        If flagExist Then
            ConnERP.Execute "update mp_tv_reach_frequency set reach=" & Val(txtTVREACH.Text) & " where mp_tv_rf_id=" & intMp_tv_rf_id
            For i = intWeekStart To intWeekEnd
                msf_MPInsertion.TextMatrix(intRow, i + 5) = String(nSpace, " ") & CStr(Val(txtTVREACH.Text)) & "%" & String(nSpace, " ")
            Next
        End If
    Else
        'Do Nothing
    End If
End Sub

Private Sub db_add()
'<CSCM>
'********************************************************************************
'Procedure Name     : db_Add
'Procedure Function : Create Media Plan Insertion
'Input Parameter    : ---
'Output Parameter   : ---
'Date               : 3/10/2016
'LastUpdate/By      : Tedi / Kreatif
'********************************************************************************
'</CSCM>

    objOpener = "frm_MPInsertion"
    Frm_MPCreate.show 1

End Sub

Private Sub incrProgressBarCopy()
    
    If ProgressBarCopy.Value + 1 >= ProgressBarCopy.Max Then
        ProgressBarCopy.Value = ProgressBarCopy.Max
    Else
        ProgressBarCopy.Value = ProgressBarCopy.Value + 1
    End If

End Sub

Private Sub db_edit()
'<CSCM>
'********************************************************************************
'Procedure Name     : db_Edit
'Procedure Function : ---
'Input Parameter    : ---
'Output Parameter   : ---
'Date               : 3/10/2016
'LastUpdate/By      : Tedi / Kreatif
'Name Before        : cmdEdit_Click
'********************************************************************************
'</CSCM>

    'Is MP already Published to Web
    If picButton(biePublishToWeb).Enabled = False Then
        MsgBox "You can't edit This Media Plan, it has been published to Web.", vbExclamation, strApplication_Name
        Exit Sub
    End If
    frm_MPEdit.show 1
    
End Sub

Private Sub db_Find()
'<CSCM>
'********************************************************************************
'Procedure Name     : db_Find
'Procedure Function : Untuk Pencarian Data
'Input Parameter    : ---
'Output Parameter   : ---
'Date               : 3/10/2016
'LastUpdate/By      : Tedi / Kreatif
'Name Before        : cmdSearch_Click
'********************************************************************************
'</CSCM>

    frm_MPSearch.show 1

End Sub

Private Sub db_BudgetSummary()
'<CSCM>
'********************************************************************************
'Procedure Name     : db_BudgetSummary
'Procedure Function : ---
'Input Parameter    : ---
'Output Parameter   : ---
'Date               : 3/10/2016
'LastUpdate/By      : Tedi / Kreatif
'Name Before        : cmdSummary_Click
'********************************************************************************
'</CSCM>

    Me.MousePointer = vbHourglass
    If cboMPNumber.Text <> "" Then
        frm_MPTotalByTask.str_mp_number = cboMPNumber.Text
        frm_MPTotalByTask.show 1
        'frm_MPSummery.Show
    Else
        MsgBox "Select MP Number!", vbExclamation, strApplication_Name
    End If
    Me.MousePointer = vbDefault
    
End Sub

Private Sub msf_MPInsertion_Click()
    PicButtonPanahBawah.Visible = False
    LstFrequency.Visible = False
    With msf_MPInsertion
        If .col > 5 And .col < .cols - 3 Then
            If isAllowEditCell(.MouseRow, .MouseCol) Then
                intRow = .Row
                intCol = .col
                If Right(.TextMatrix(.Row, .cols - 1), 4) = "FREQ" Then
                    If .Text <> "" Then
                        'Show button List Freq
                        PicButtonPanahBawah.Top = .Top + .CellTop
                        PicButtonPanahBawah.Left = .Left + .CellLeft + .CellWidth - PicButtonPanahBawah.Width
                        PicButtonPanahBawah.Visible = True
                    End If
                End If
            End If
        End If
    End With
    
End Sub

Private Sub msf_MPInsertion_SelChange()
    With msf_MPInsertion
        If Not .ColIsVisible(.col) Then
            .LeftCol = .LeftCol + 1
        End If
    End With
End Sub

Private Sub StartTyping(mode As String, Optional KeyAscii As Integer)

    With msf_MPInsertion
        txtMPInsertion.Top = .CellTop + .Top
        txtMPInsertion.Left = .CellLeft + .Left
        txtMPInsertion.Width = .CellWidth
        txtMPInsertion.Height = .CellHeight
        If mode = "Edit" Then
            txtMPInsertion.Text = msf_MPInsertion.Text
            EditMode = True
        Else
            txtMPInsertion.Text = Chr(KeyAscii)
            EditMode = False
        End If
        txtMPInsertion.SelStart = Len(txtMPInsertion.Text)
        txtMPInsertion.Visible = True
        txtMPInsertion.SetFocus
    End With
    
End Sub

Private Sub db_close()
    Unload Me
End Sub

Private Function isAllowEditCell(baris As Integer, Kolom As Integer) As Boolean
    
    Dim DtselectedDate As Date, DtLocked As Date
    Dim intMonth As Integer, IntYear As Integer
    Dim strSql As String
    
    msf_MPInsertion.Row = baris
    msf_MPInsertion.col = Kolom
    
    If msf_MPInsertion.CellBackColor = LegendApproval.FillColor Or msf_MPInsertion.CellBackColor = LegendWebApproval.FillColor Then 'Sudah Approved!
        isAllowEditCell = False
        Exit Function
    End If
    
    If EngMonthIndex(msf_MPInsertion.TextMatrix(1, Kolom)) = -1 Then
        Exit Function
    End If
    
    intMonth = EngMonthIndex(msf_MPInsertion.TextMatrix(1, Kolom))
    
    IntYear = CInt(msf_MPInsertion.TextMatrix(0, Kolom))
    
    recDate.Requery
    'If Right(FGMPInsertion.TextMatrix(baris, FGMPInsertion.Cols - 1), 2) = "PR" Then
    If False Then 'locking bukan monthly (sebelumnya pakai monthly, jadi di skip ajah pake if false)
        intMonth = intMonth - 1
        If intMonth = 0 Then
            intMonth = 12
            IntYear = IntYear - 1
        End If
        If IntYear > Year(recDate(0)) Then
            isAllowEditCell = True
        Else
            If IntYear = Year(recDate(0)) Then
                If intMonth > month(recDate(0)) Then
                    isAllowEditCell = True
                Else
                    isAllowEditCell = False
                End If
            Else
                isAllowEditCell = False
            End If
        End If
    Else
        
        DtselectedDate = CDate(msf_MPInsertion.TextMatrix(4, Kolom))
        If bol_BU2 Then
            DtLocked = DateAdd("d", 0 - Weekday(recDate(0)), recDate(0)) 'BU2 : Current weeks masih unlocked
        Else
            DtLocked = DateAdd("d", 14 - Weekday(recDate(0)), recDate(0)) 'BU1 : Curr Week + 1 sudah locked
        End If
        
        If DtselectedDate > DtLocked Then
            isAllowEditCell = True
        Else
            isAllowEditCell = False
            'isAllowEditCell = True 'locking pake weekly (sementara disable dulu (always true) per 6 july 2005)
        End If
    End If
    'isAllowEditCell = True 'skip this for Lock Mode
    If msf_MPInsertion.CellBackColor = Shape_unlocked.FillColor Then
        mdi_Main.mnu_unlock.Enabled = True
        mdi_Main.mnu_unlock.Caption = "Lock Cell"
        isAllowEditCell = True
    End If
    
    If Not isAllowEditCell Then
        If msf_MPInsertion.TextMatrix(baris, msf_MPInsertion.cols - 1) <> "" And _
            Mid(msf_MPInsertion.TextMatrix(baris, msf_MPInsertion.cols - 1), 6, 4) <> "MDUM" Then
            If msf_MPInsertion.CellBackColor <> LegendApproval.FillColor And _
                msf_MPInsertion.CellBackColor <> LegendWebApproval.FillColor And _
                msf_MPInsertion.CellBackColor <> LegendActual.FillColor Then
                mdi_Main.mnu_unlock.Enabled = True
                mdi_Main.mnu_unlock.Caption = "Unlock Cell"
            End If
        End If
    End If
End Function

Private Function NextMPMediumID(strMPACtivityID As String) As String
    Dim rsTemp As New ADODB.Recordset, strSql As String
    strSql = "select isnull(max(cast(substring(mp_medium_id,16,4) as int)),0)+1 from mp_medium where mp_activity_id like '" & Mid(strMPACtivityID, 1, 15) & "%'"
    rsTemp.Open strSql, ConnERP, 1, 3
        NextMPMediumID = Mid(strMPACtivityID, 1, 4) & ".MDUM." & Mid(strMPACtivityID, 11, 4) & "." & Right("0000" & CStr(rsTemp(0)), 4) 'Create new medium_Id
    rsTemp.Close
End Function

Private Function NextMPMediumDetailID(strMPMediumID As String) As String
    Dim rsTemp As New ADODB.Recordset, strSql As String
    strSql = "select isnull(max(cast(substring(mp_medium_detail_id,16,4) as int)),0)+1 from mp_medium_detail where mp_medium_id like '" & Mid(strMPMediumID, 1, 15) & "%'"
    rsTemp.Open strSql, ConnERP, 1, 3
        NextMPMediumDetailID = Mid(strMPMediumID, 1, 4) & ".MDUD." & Mid(strMPMediumID, 11, 4) & "." & Right("0000" & CStr(rsTemp(0)), 4) 'Create new medium_Detail_Id
    rsTemp.Close
End Function

Private Function NextMPPlanDimID(strMPMediumDetailID As String) As String
    Dim rsTemp As New ADODB.Recordset, strSql As String
    strSql = "select isnull(max(cast(substring(mp_plan_dim_id,16,4) as int)),0)+1 from mp_plan_dimension where mp_medium_detail_id like '" & Mid(strMPMediumDetailID, 1, 15) & "%'"
    rsTemp.Open strSql, ConnERP, 1, 3
        NextMPPlanDimID = Mid(strMPMediumDetailID, 1, 4) & ".MPDM." & Mid(strMPMediumDetailID, 11, 4) & "." & Right("0000" & CStr(rsTemp(0)), 4) 'Create new plan dimension Id
    rsTemp.Close
End Function

Public Sub Mnu_unlock_Click()
    If mdi_Main.mnu_unlock.Caption = "Lock Cell" Then
        'Lock Cell
        Call Lock_Cell
    Else
        'Unlock Cell
        Frm_MPUnlockSecurity.show 1
    End If
End Sub

Private Sub Lock_Cell()
    Dim str_MP_Plan_Dim_Id As String
    Dim int_Week_Year As Integer
    Dim dt_Week_Commencing As Date
    Dim str_Unlock_Req_By As String
    Dim str_Unlock_By As String
    Dim dt_Unlock_Date As Date
    Dim str_Note As String
    Dim strSql As String
    With msf_MPInsertion
        str_MP_Plan_Dim_Id = Left(.TextMatrix(.Row, .cols - 1), 19)
        int_Week_Year = CInt(.TextMatrix(3, .col))
        'insert to history
        strSql = "insert into mp_unlocked_week_history(mp_plan_dim_id,week_year,week_commencing,unlock_req_by,unlock_by,unlock_date,note) "
        strSql = strSql & "select mp_plan_dim_id,week_year,week_commencing,unlock_req_by,unlock_by,unlock_date,note from mp_unlocked_week "
        strSql = strSql & "where mp_plan_dim_id = '" & str_MP_Plan_Dim_Id & "' and week_year = " & int_Week_Year
        ConnERP.Execute strSql
        
        'delete actual
        strSql = "delete from mp_unlocked_week where mp_plan_dim_id = '" & str_MP_Plan_Dim_Id & "' and week_year = " & int_Week_Year
        ConnERP.Execute strSql
    End With
    
    With msf_MPInsertion
        Select Case UCase(Right(.TextMatrix(.Row, .cols - 1), 4))
            Case "EACH"
                If Trim(.Text) <> "" Then
                    .CellBackColor = LegendTVRF.FillColor
                Else
                    .CellBackColor = .BackColor
                End If
                .Row = .Row - 1
                If Trim(.Text) <> "" Then
                    .CellBackColor = LegendTVRF.FillColor
                Else
                    .CellBackColor = .BackColor
                End If
                .Row = .Row + 2
                .CellBackColor = .BackColor
                .Row = .Row - 1
            Case "FREQ"
                If Trim(.Text) <> "" Then
                    .CellBackColor = LegendTVRF.FillColor
                Else
                    .CellBackColor = .BackColor
                End If
                .Row = .Row + 1
                If Trim(.Text) <> "" Then
                    .CellBackColor = LegendTVRF.FillColor
                Else
                    .CellBackColor = .BackColor
                End If
                .Row = .Row + 1
                .CellBackColor = .BackColor
                .Row = .Row - 2
            Case Else
                .CellBackColor = .BackColor
                If UCase(Right(.TextMatrix(.Row - 1, .cols - 1), 5)) = "REACH" Then
                    .Row = .Row - 1
                    If Trim(.Text) <> "" Then
                        .CellBackColor = LegendTVRF.FillColor
                    Else
                        .CellBackColor = .BackColor
                    End If
                    
                    .Row = .Row - 1
                    If Trim(.Text) <> "" Then
                        .CellBackColor = LegendTVRF.FillColor
                    Else
                        .CellBackColor = .BackColor
                    End If
                    
                    .Row = .Row + 2
                End If
        End Select
    End With
End Sub

Function isInsertionNotEmpty(strMediumID As String, intMonth As Integer) As Boolean
    Dim rsTemp As New ADODB.Recordset
    Dim strSql As String, strMediumCode As String
    isInsertionNotEmpty = False
    strMediumCode = ""
    strSql = "select medium_code from mp_medium where mp_medium_id = '" & strMediumID & "'"
    rsTemp.Open strSql, ConnERP, 1, 3
    If Not rsTemp.EOF Then strMediumCode = rsTemp(0)
    rsTemp.Close
    If strMediumCode <> "" Then
        Select Case strMediumCode
            Case "OT"
                strSql = "select count(*) from MP_Other_Monthly_Budget where month_number = " & intMonth
            Case Else
                strSql = "select count(*) from MP_Insertion where [month] = " & intMonth
        End Select
        strSql = strSql & " and mp_plan_dim_id in (select mp_plan_dim_id from mp_ids where mp_medium_id = '" & strMediumID & "')"
        rsTemp.Open strSql, ConnERP, 1, 3
        If rsTemp(0) <> 0 Then isInsertionNotEmpty = True
        rsTemp.Close
    End If
End Function

Private Sub Create_Log_File()
'*************************************************************
'Nama Procedure     : Create_Log_File
'Fungsi             : Membuat file text untuk LOG
'Programer          :
'Tgl Pembuatan      : 14 Juni 2004
'Last Update/By     :
'*************************************************************
        
    On Error GoTo errHand
    
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set FSOInterface = fso.CreateFolder("c:\MPLog")
    
    recDate.Requery
    LogFileName = "MPAppLog_" & Format(recDate(0), "MMDDYY") & "_" & Format(recDate(0), "hhmmss") & ".txt"
    Set FSOInterface = fso.CreateTextFile("c:\MPLog\" & LogFileName)
    
    Exit Sub
    
errHand:
    If Err.Number = 58 Then
        Resume Next
    End If
End Sub

Private Sub Write_To_Log_File(StrTeks As String)
'*************************************************************
'Nama Procedure     : Write_To_Log_File
'Fungsi             : Menulis ke LOG
'Programer          :
'Tgl Pembuatan      : 14 Juni 2004
'Last Update/By     :
'*************************************************************
    recDate.Requery
    FSOInterface.WriteLine StrTeks
End Sub

Private Sub DoApproval(strMediumID As String, intMonth As Integer)
    Dim strSql As String
    
    'Disabled because polpulated via Web
    strSql = "update mp_monthly_activity set approval=1,"
    'StrSQL = StrSQL & "Client_Approved_By='" & txtClientApprovedBY.Text & "',"
    'StrSQL = StrSQL & "Client_Approved_Date='" & CDate(DTClientApprovalDate.Value) & "',"
    'StrSQL = StrSQL & "Client_Noted_By='" & txtClientNotedBy.Text & "',"
    strSql = strSql & "Approved_By='" & strLogin_User & "',"
    strSql = strSql & "Approved_Date=getdate(),"
    strSql = strSql & "Approved_mp_medium_id='" & strMediumID & "',"
    strSql = strSql & "Approved_mp_medium_id_history = replace(isnull(Approved_mp_medium_id_history,''),'" & Right(strMediumID, 5) & "','') + '" & Right(strMediumID, 5) & "'"
    strSql = strSql & " WHERE mp_medium_id = '" & strMediumID & "' and month_number=" & intMonth
    ConnERP.Execute strSql
    
    '                           Add/Update MP_Monthly_Quotation
    '============================================================================
    Dim strMPNumber As String
    Dim rsTemp As New ADODB.Recordset
    Dim strMediumCode As String
    
    Dim StrOriginalBrandCode As String
    
    
    'Get MP Number
    strMPNumber = cboMPNumber.Text
    
    'Get Medium Code
    rsTemp.Open "select medium_code from mp_medium where mp_medium_id='" & strMediumID & "'", ConnERP, 1, 3
        strMediumCode = rsTemp(0)
    rsTemp.Close
    
    'Get StrOriginalBrandCode
    strSql = "SELECT Original_Brand_Code,Original_Brand_Name FROM MP_Activity "
    strSql = strSql & " WHERE MP_Activity_ID IN(SELECT MP_Activity_ID FROM MP_Medium WHERE mp_medium_id='" & strMediumID & "')"
    
    rsTemp.Open strSql, ConnERP, 1, 3
        StrOriginalBrandCode = rsTemp(0).Value
    rsTemp.Close
    
    'cek apakah di quotation sudah ada?? & Get Monthtly_Quotation_ID yang sama Brandocde nya
    Dim isExist As Boolean
    Dim strmp_Monthly_Quot_ID As String
    
    isExist = False
    strmp_Monthly_Quot_ID = ""
    
    strSql = "SELECT mp_monthly_quot_id FROM mp_monthly_quotation "
    strSql = strSql & " WHERE mp_number='" & strMPNumber & "' AND medium_code='" & strMediumCode & "' AND [month]=" & intMonth
    strSql = strSql & " AND LEFT(mp_monthly_quot_id,4)='" & StrOriginalBrandCode & "'"
    
    rsTemp.Open strSql, ConnERP, 1, 3
    If Not rsTemp.EOF Then
        isExist = True
        strmp_Monthly_Quot_ID = rsTemp(0)
    End If
    rsTemp.Close
    
    'Get Total Value Monthly Quotation from mp_monthly_activity
    
    Dim int_total_gross As Double
    Dim int_total_nett As Double
    Dim int_msc As Double
    Dim int_other_cost As Double
    Dim int_bonus_fee As Double
    Dim int_agency_charge As Double
    
    Dim int_sub_total As Double
    Dim int_grand_total As Double
    
    int_total_gross = 0
    int_total_nett = 0
    int_msc = 0
    int_other_cost = 0
    int_bonus_fee = 0
    int_agency_charge = 0
    
    strSql = "select sum(min_budget),sum(gross_budget),sum(msc_paid_value),sum(msc_bonus_value) from mp_monthly_activity "
    strSql = strSql & "where month_number = " & intMonth & " and mp_medium_id = '" & strMediumID & "' "
    
    Dim rsQuotation As New ADODB.Recordset
    rsQuotation.Open strSql, ConnERP, 1, 3
    If Not rsQuotation.EOF Then
        int_total_gross = rsQuotation(1)
        int_total_nett = rsQuotation(0)
        int_msc = rsQuotation(2)
        'int_other_cost = 0
        int_bonus_fee = rsQuotation(3)
        'int_agency_charge = 0
    End If
    rsQuotation.Close
    Set rsQuotation = Nothing
    
    int_sub_total = int_total_nett + int_msc + int_other_cost + int_bonus_fee
    int_grand_total = int_sub_total + int_agency_charge
    
    '==========================================================================================
    If isExist Then
        'update
        strSql = "UPDATE mp_monthly_quotation SET "
        strSql = strSql & "total_gross=total_gross+" & int_total_gross
        strSql = strSql & ", total_nett=total_nett+" & int_total_nett
        strSql = strSql & ", MSC=MSC+" & int_msc
        strSql = strSql & ", other_cost=other_cost+" & int_other_cost
        strSql = strSql & ", bonus_fee=bonus_fee+" & int_bonus_fee
        strSql = strSql & ", sub_total=sub_total+" & int_sub_total
        strSql = strSql & ", agency_charge=agency_charge+" & int_agency_charge
        strSql = strSql & ", grand_total = grand_total+" & int_grand_total
        strSql = strSql & " WHERE mp_monthly_quot_id='" & strmp_Monthly_Quot_ID & "'"
        
        ConnERP.Execute strSql
    Else
        'generate id baru
        
        Dim Int_Year As Integer
        Dim rsRunningNumber As New ADODB.Recordset
        Dim str_running_number As String
        Dim str_param_job_id As String, str_param_job_number_agency As String
        Dim str_job_id As String, str_job_number_agency As String
        Dim strTemp As String
        
        Int_Year = CInt(Mid(strMPNumber, 6, 4))
        
        strTemp = StrOriginalBrandCode & "." & Int_Year & "."
        
        'rsRunningNumber.Open "SELECT ISNULL(MAX(CAST(SUBSTRING(mp_monthly_quot_id,11,4) AS INT)),0)+1 FROM mp_monthly_quotation WHERE mp_monthly_quot_id LIKE '" & Mid(strMPNumber, 1, 10) & "%'", ConnERP, 1, 3
        rsRunningNumber.Open "SELECT ISNULL(MAX(CAST(SUBSTRING(mp_monthly_quot_id,11,4) AS INT)),0)+1 FROM mp_monthly_quotation WHERE mp_monthly_quot_id LIKE '" & strTemp & "%'", ConnERP, 1, 3
        
        str_running_number = Right("0000" & CStr(rsRunningNumber(0)), 4)
        rsRunningNumber.Close
        
        'strmp_Monthly_Quot_ID = Mid(strMPNumber, 1, 4) & "." & CStr(Int_Year) & "." & str_running_number
        strmp_Monthly_Quot_ID = StrOriginalBrandCode & "." & CStr(Int_Year) & "." & str_running_number
        
        'Generate job_id dan job_number_agency
        Select Case strMediumCode
            Case "TV":
                str_param_job_id = "Television Media Induk"
                str_param_job_number_agency = "Television Club Agency"
            Case "RD":
                str_param_job_id = "Radio Media Induk"
                str_param_job_number_agency = "Radio Club Agency"
            Case "PR":
                str_param_job_id = "Print Media Induk"
                str_param_job_number_agency = "Print Club Agency"
            Case "CN":
                str_param_job_id = "Other Media Induk Cinema"
                str_param_job_number_agency = "Other Club Agency"
            Case "OT":
                str_param_job_id = "Other Media Induk Outdoor"
                str_param_job_number_agency = "Other Club Agency"
        End Select
        
        rsTemp.Open "select media_type_code from media_type where media_type_name = '" & str_param_job_id & "'", ConnERP, 1, 3
            str_job_id = StrOriginalBrandCode & "." & rsTemp(0) & "." & Right(CStr(Int_Year), 2) & Right("0" & CStr(intMonth), 2)
        rsTemp.Close
        
        rsTemp.Open "select media_type_code from media_type where media_type_name = '" & str_param_job_number_agency & "'", ConnERP, 1, 3
            str_job_number_agency = StrOriginalBrandCode & "." & rsTemp(0) & "." & Right(CStr(Int_Year), 2) & Right("0" & CStr(intMonth), 2)
        rsTemp.Close
        
        'insert
        strSql = "INSERT INTO mp_monthly_quotation"
        strSql = strSql & "(mp_monthly_quot_id,[year],mp_number,medium_code,[month],total_gross,total_nett,msc,other_cost,bonus_fee,sub_total,agency_charge,grand_total,"
        strSql = strSql & "job_id,job_number_agency,"
        strSql = strSql & "is_latest,status) VALUES "
        strSql = strSql & "('" & strmp_Monthly_Quot_ID & "'," & Int_Year & ",'" & strMPNumber & "','" & strMediumCode & "'," & intMonth & "," & int_total_gross & "," & int_total_nett & "," & int_msc & "," & int_other_cost & "," & int_bonus_fee & "," & int_sub_total & "," & int_agency_charge & "," & int_grand_total & ","
        strSql = strSql & "'" & str_job_id & "','" & str_job_number_agency & "',"
        strSql = strSql & "1,1)"
        ConnERP.Execute strSql
        
    End If
    '======================== Enf of Update Quotation ==========================
    
End Sub


Private Sub CalculateBalance(xlws As Object, xlRow As Integer, xlLeftCol As Integer, xlTopRow As Integer)
'*****************************************************************************
' Nama Submodul         :  GrandTotal
' Fungsi Submodul       :  Print Grand Total
' Nama Programmer       :  Sistyo
' Tgl Pembuatan         :  18 Agustus 2004
' Last Update           :  18 Agustus 2004/Sistyo
'******************************************************************************
    
    With xlws
        .Rows(CStr(xlRow) & ":" & CStr(xlRow + 7)).Insert Shift:=xlDown
        .Range(.Cells(xlRow + 1, xlLeftCol).Address & ":" & eRange_Dec & CStr(xlRow + 5)).Interior.ColorIndex = 2
        .Range(.Cells(xlRow + 1, xlLeftCol).Address & ":" & eRange_Dec & CStr(xlRow + 5)).Borders(xlEdgeTop).LineStyle = xlContinuous
        .Range(.Cells(xlRow + 1, xlLeftCol).Address & ":" & eRange_Dec & CStr(xlRow + 5)).Borders(xlEdgeTop).ColorIndex = xlAutomatic
        .Range(.Cells(xlRow + 1, xlLeftCol).Address & ":" & eRange_Dec & CStr(xlRow + 5)).Borders(xlEdgeTop).Weight = xlThin
        .Range(.Cells(xlRow + 1, xlLeftCol).Address & ":" & eRange_Dec & CStr(xlRow + 5)).Interior.ColorIndex = 2
        .Range(.Cells(xlRow + 1, xlLeftCol).Address & ":" & eRange_Dec & CStr(xlRow + 5)).Borders(xlEdgeBottom).LineStyle = xlContinuous
        .Range(.Cells(xlRow + 1, xlLeftCol).Address & ":" & eRange_Dec & CStr(xlRow + 5)).Borders(xlEdgeBottom).ColorIndex = xlAutomatic
        .Range(.Cells(xlRow + 1, xlLeftCol).Address & ":" & eRange_Dec & CStr(xlRow + 5)).Borders(xlEdgeBottom).Weight = xlThin
        .Range(.Cells(xlRow + 1, xlLeftCol).Address & ":" & eRange_Dec & CStr(xlRow + 5)).Interior.ColorIndex = 2
        .Range(.Cells(xlRow + 1, xlLeftCol).Address & ":" & eRange_Dec & CStr(xlRow + 5)).Borders(xlInsideHorizontal).LineStyle = xlContinuous
        .Range(.Cells(xlRow + 1, xlLeftCol).Address & ":" & eRange_Dec & CStr(xlRow + 5)).Borders(xlInsideHorizontal).ColorIndex = xlAutomatic
        .Range(.Cells(xlRow + 1, xlLeftCol).Address & ":" & eRange_Dec & CStr(xlRow + 5)).Borders(xlInsideHorizontal).Weight = xlThin
        
        .Range(.Cells(xlRow + 1, xlLeftCol).Address & ":" & eRange_Dec & CStr(xlRow + 8)).Font.Size = 9
        .Range(.Cells(xlRow + 1, xlLeftCol).Address & ":" & eRange_Dec & CStr(xlRow + 8)).Font.ColorIndex = 5
        .Range(.Cells(xlRow + 1, xlLeftCol).Address & ":" & eRange_Dec & CStr(xlRow + 8)).Font.Bold = True
        
        .Cells(xlRow, xlLeftCol).Value = "BALANCE"
        .Cells(xlRow + 1, xlLeftCol).Value = "Television"
        .Cells(xlRow + 2, xlLeftCol).Value = "Radio"
        .Cells(xlRow + 3, xlLeftCol).Value = "Print"
        .Cells(xlRow + 4, xlLeftCol).Value = "Cinema"
        .Cells(xlRow + 5, xlLeftCol).Value = "Other"
        .Cells(xlRow + 6, xlLeftCol + 4).Value = "Grand Total / Month"
        .Cells(xlRow + 7, xlLeftCol + 4).Value = "Grand Total / Quarter"
        
        
        
        'Border grand
        .Range(sRange_Jan & CStr(xlRow + 1) & ":" & .Cells(xlRow + 7, msf_MPInsertion.cols - 2 + xlLeftCol - 3).Address).Borders.LineStyle = xlContinuous
        .Range(sRange_Jan & CStr(xlRow + 1) & ":" & .Cells(xlRow + 7, msf_MPInsertion.cols - 2 + xlLeftCol - 3).Address).Borders.Weight = xlThick
        .Range(sRange_Jan & CStr(xlRow + 1) & ":" & .Cells(xlRow + 7, msf_MPInsertion.cols - 2 + xlLeftCol - 3).Address).Borders.ColorIndex = 5
        
        .Range(sRange_Jan & CStr(xlRow + 1) & ":" & eRange_Dec & CStr(xlRow + 8)).NumberFormat = "_(#,#0.00_);_(#,#0.00_);_(-_)"
        
        .Range(sRange_Jan & CStr(xlRow + 1) & ":" & eRange_Jan & CStr(xlRow + 1)).Merge
        .Range(sRange_Jan & CStr(xlRow + 1) & ":" & eRange_Jan & CStr(xlRow + 1)).Formula = "=" & sRange_Jan & CStr(xlRow - 15) & " - " & sRange_Jan & CStr(xlRow - 7)
        .Range(sRange_Jan & CStr(xlRow + 2) & ":" & eRange_Jan & CStr(xlRow + 2)).Merge
        .Range(sRange_Jan & CStr(xlRow + 2) & ":" & eRange_Jan & CStr(xlRow + 2)).Formula = "=" & sRange_Jan & CStr(xlRow - 14) & " - " & sRange_Jan & CStr(xlRow - 6)
        .Range(sRange_Jan & CStr(xlRow + 3) & ":" & eRange_Jan & CStr(xlRow + 3)).Merge
        .Range(sRange_Jan & CStr(xlRow + 3) & ":" & eRange_Jan & CStr(xlRow + 3)).Formula = "=" & sRange_Jan & CStr(xlRow - 13) & " - " & sRange_Jan & CStr(xlRow - 5)
        .Range(sRange_Jan & CStr(xlRow + 4) & ":" & eRange_Jan & CStr(xlRow + 4)).Merge
        .Range(sRange_Jan & CStr(xlRow + 4) & ":" & eRange_Jan & CStr(xlRow + 4)).Formula = "=" & sRange_Jan & CStr(xlRow - 12) & " - " & sRange_Jan & CStr(xlRow - 4)
        .Range(sRange_Jan & CStr(xlRow + 5) & ":" & eRange_Jan & CStr(xlRow + 5)).Merge
        .Range(sRange_Jan & CStr(xlRow + 5) & ":" & eRange_Jan & CStr(xlRow + 5)).Formula = "=" & sRange_Jan & CStr(xlRow - 11) & " - " & sRange_Jan & CStr(xlRow - 3)
        .Range(sRange_Jan & CStr(xlRow + 6) & ":" & eRange_Jan & CStr(xlRow + 6)).Merge
        .Range(sRange_Jan & CStr(xlRow + 6) & ":" & eRange_Jan & CStr(xlRow + 6)).Formula = "=" & sRange_Jan & CStr(xlRow - 10) & " - " & sRange_Jan & CStr(xlRow - 2)
        
        .Range(sRange_Feb & CStr(xlRow + 1) & ":" & eRange_Feb & CStr(xlRow + 1)).Merge
        .Range(sRange_Feb & CStr(xlRow + 1) & ":" & eRange_Feb & CStr(xlRow + 1)).Formula = "=" & sRange_Feb & CStr(xlRow - 15) & " - " & sRange_Feb & CStr(xlRow - 7)
        .Range(sRange_Feb & CStr(xlRow + 2) & ":" & eRange_Feb & CStr(xlRow + 2)).Merge
        .Range(sRange_Feb & CStr(xlRow + 2) & ":" & eRange_Feb & CStr(xlRow + 2)).Formula = "=" & sRange_Feb & CStr(xlRow - 14) & " - " & sRange_Feb & CStr(xlRow - 6)
        .Range(sRange_Feb & CStr(xlRow + 3) & ":" & eRange_Feb & CStr(xlRow + 3)).Merge
        .Range(sRange_Feb & CStr(xlRow + 3) & ":" & eRange_Feb & CStr(xlRow + 3)).Formula = "=" & sRange_Feb & CStr(xlRow - 13) & " - " & sRange_Feb & CStr(xlRow - 5)
        .Range(sRange_Feb & CStr(xlRow + 4) & ":" & eRange_Feb & CStr(xlRow + 4)).Merge
        .Range(sRange_Feb & CStr(xlRow + 4) & ":" & eRange_Feb & CStr(xlRow + 4)).Formula = "=" & sRange_Feb & CStr(xlRow - 12) & " - " & sRange_Feb & CStr(xlRow - 4)
        .Range(sRange_Feb & CStr(xlRow + 5) & ":" & eRange_Feb & CStr(xlRow + 5)).Merge
        .Range(sRange_Feb & CStr(xlRow + 5) & ":" & eRange_Feb & CStr(xlRow + 5)).Formula = "=" & sRange_Feb & CStr(xlRow - 11) & " - " & sRange_Feb & CStr(xlRow - 3)
        .Range(sRange_Feb & CStr(xlRow + 6) & ":" & eRange_Feb & CStr(xlRow + 6)).Merge
        .Range(sRange_Feb & CStr(xlRow + 6) & ":" & eRange_Feb & CStr(xlRow + 6)).Formula = "=" & sRange_Feb & CStr(xlRow - 10) & " - " & sRange_Feb & CStr(xlRow - 2)
        
      
        .Range(sRange_Mar & CStr(xlRow + 1) & ":" & eRange_Mar & CStr(xlRow + 1)).Merge
        .Range(sRange_Mar & CStr(xlRow + 1) & ":" & eRange_Mar & CStr(xlRow + 1)).Formula = "=" & sRange_Mar & CStr(xlRow - 15) & " - " & sRange_Mar & CStr(xlRow - 7)
        .Range(sRange_Mar & CStr(xlRow + 2) & ":" & eRange_Mar & CStr(xlRow + 2)).Merge
        .Range(sRange_Mar & CStr(xlRow + 2) & ":" & eRange_Mar & CStr(xlRow + 2)).Formula = "=" & sRange_Mar & CStr(xlRow - 14) & " - " & sRange_Mar & CStr(xlRow - 6)
        .Range(sRange_Mar & CStr(xlRow + 3) & ":" & eRange_Mar & CStr(xlRow + 3)).Merge
        .Range(sRange_Mar & CStr(xlRow + 3) & ":" & eRange_Mar & CStr(xlRow + 3)).Formula = "=" & sRange_Mar & CStr(xlRow - 13) & " - " & sRange_Mar & CStr(xlRow - 5)
        .Range(sRange_Mar & CStr(xlRow + 4) & ":" & eRange_Mar & CStr(xlRow + 4)).Merge
        .Range(sRange_Mar & CStr(xlRow + 4) & ":" & eRange_Mar & CStr(xlRow + 4)).Formula = "=" & sRange_Mar & CStr(xlRow - 12) & " - " & sRange_Mar & CStr(xlRow - 4)
        .Range(sRange_Mar & CStr(xlRow + 5) & ":" & eRange_Mar & CStr(xlRow + 5)).Merge
        .Range(sRange_Mar & CStr(xlRow + 5) & ":" & eRange_Mar & CStr(xlRow + 5)).Formula = "=" & sRange_Mar & CStr(xlRow - 11) & " - " & sRange_Mar & CStr(xlRow - 3)
        .Range(sRange_Mar & CStr(xlRow + 6) & ":" & eRange_Mar & CStr(xlRow + 6)).Merge
        .Range(sRange_Mar & CStr(xlRow + 6) & ":" & eRange_Mar & CStr(xlRow + 6)).Formula = "=" & sRange_Mar & CStr(xlRow - 10) & " - " & sRange_Mar & CStr(xlRow - 2)
        
        
        .Range(sRange_Apr & CStr(xlRow + 1) & ":" & eRange_Apr & CStr(xlRow + 1)).Merge
        .Range(sRange_Apr & CStr(xlRow + 1) & ":" & eRange_Apr & CStr(xlRow + 1)).Formula = "=" & sRange_Apr & CStr(xlRow - 15) & " - " & sRange_Apr & CStr(xlRow - 7)
        .Range(sRange_Apr & CStr(xlRow + 2) & ":" & eRange_Apr & CStr(xlRow + 2)).Merge
        .Range(sRange_Apr & CStr(xlRow + 2) & ":" & eRange_Apr & CStr(xlRow + 2)).Formula = "=" & sRange_Apr & CStr(xlRow - 14) & " - " & sRange_Apr & CStr(xlRow - 6)
        .Range(sRange_Apr & CStr(xlRow + 3) & ":" & eRange_Apr & CStr(xlRow + 3)).Merge
        .Range(sRange_Apr & CStr(xlRow + 3) & ":" & eRange_Apr & CStr(xlRow + 3)).Formula = "=" & sRange_Apr & CStr(xlRow - 13) & " - " & sRange_Apr & CStr(xlRow - 5)
        .Range(sRange_Apr & CStr(xlRow + 4) & ":" & eRange_Apr & CStr(xlRow + 4)).Merge
        .Range(sRange_Apr & CStr(xlRow + 4) & ":" & eRange_Apr & CStr(xlRow + 4)).Formula = "=" & sRange_Apr & CStr(xlRow - 12) & " - " & sRange_Apr & CStr(xlRow - 4)
        .Range(sRange_Apr & CStr(xlRow + 5) & ":" & eRange_Apr & CStr(xlRow + 5)).Merge
        .Range(sRange_Apr & CStr(xlRow + 5) & ":" & eRange_Apr & CStr(xlRow + 5)).Formula = "=" & sRange_Apr & CStr(xlRow - 11) & " - " & sRange_Apr & CStr(xlRow - 3)
        .Range(sRange_Apr & CStr(xlRow + 6) & ":" & eRange_Apr & CStr(xlRow + 6)).Merge
        .Range(sRange_Apr & CStr(xlRow + 6) & ":" & eRange_Apr & CStr(xlRow + 6)).Formula = "=" & sRange_Apr & CStr(xlRow - 10) & " - " & sRange_Apr & CStr(xlRow - 2)
        
      
        .Range(sRange_May & CStr(xlRow + 1) & ":" & eRange_May & CStr(xlRow + 1)).Merge
        .Range(sRange_May & CStr(xlRow + 1) & ":" & eRange_May & CStr(xlRow + 1)).Formula = "=" & sRange_May & CStr(xlRow - 15) & " - " & sRange_May & CStr(xlRow - 7)
        .Range(sRange_May & CStr(xlRow + 2) & ":" & eRange_May & CStr(xlRow + 2)).Merge
        .Range(sRange_May & CStr(xlRow + 2) & ":" & eRange_May & CStr(xlRow + 2)).Formula = "=" & sRange_May & CStr(xlRow - 14) & " - " & sRange_May & CStr(xlRow - 6)
        .Range(sRange_May & CStr(xlRow + 3) & ":" & eRange_May & CStr(xlRow + 3)).Merge
        .Range(sRange_May & CStr(xlRow + 3) & ":" & eRange_May & CStr(xlRow + 3)).Formula = "=" & sRange_May & CStr(xlRow - 13) & " - " & sRange_May & CStr(xlRow - 5)
        .Range(sRange_May & CStr(xlRow + 4) & ":" & eRange_May & CStr(xlRow + 4)).Merge
        .Range(sRange_May & CStr(xlRow + 4) & ":" & eRange_May & CStr(xlRow + 4)).Formula = "=" & sRange_May & CStr(xlRow - 12) & " - " & sRange_May & CStr(xlRow - 4)
        .Range(sRange_May & CStr(xlRow + 5) & ":" & eRange_May & CStr(xlRow + 5)).Merge
        .Range(sRange_May & CStr(xlRow + 5) & ":" & eRange_May & CStr(xlRow + 5)).Formula = "=" & sRange_May & CStr(xlRow - 11) & " - " & sRange_May & CStr(xlRow - 3)
        .Range(sRange_May & CStr(xlRow + 6) & ":" & eRange_May & CStr(xlRow + 6)).Merge
        .Range(sRange_May & CStr(xlRow + 6) & ":" & eRange_May & CStr(xlRow + 6)).Formula = "=" & sRange_May & CStr(xlRow - 10) & " - " & sRange_May & CStr(xlRow - 2)
        
      
        .Range(sRange_Jun & CStr(xlRow + 1) & ":" & eRange_Jun & CStr(xlRow + 1)).Merge
        .Range(sRange_Jun & CStr(xlRow + 1) & ":" & eRange_Jun & CStr(xlRow + 1)).Formula = "=" & sRange_Jun & CStr(xlRow - 15) & " - " & sRange_Jun & CStr(xlRow - 7)
        .Range(sRange_Jun & CStr(xlRow + 2) & ":" & eRange_Jun & CStr(xlRow + 2)).Merge
        .Range(sRange_Jun & CStr(xlRow + 2) & ":" & eRange_Jun & CStr(xlRow + 2)).Formula = "=" & sRange_Jun & CStr(xlRow - 14) & " - " & sRange_Jun & CStr(xlRow - 6)
        .Range(sRange_Jun & CStr(xlRow + 3) & ":" & eRange_Jun & CStr(xlRow + 3)).Merge
        .Range(sRange_Jun & CStr(xlRow + 3) & ":" & eRange_Jun & CStr(xlRow + 3)).Formula = "=" & sRange_Jun & CStr(xlRow - 13) & " - " & sRange_Jun & CStr(xlRow - 5)
        .Range(sRange_Jun & CStr(xlRow + 4) & ":" & eRange_Jun & CStr(xlRow + 4)).Merge
        .Range(sRange_Jun & CStr(xlRow + 4) & ":" & eRange_Jun & CStr(xlRow + 4)).Formula = "=" & sRange_Jun & CStr(xlRow - 12) & " - " & sRange_Jun & CStr(xlRow - 4)
        .Range(sRange_Jun & CStr(xlRow + 5) & ":" & eRange_Jun & CStr(xlRow + 5)).Merge
        .Range(sRange_Jun & CStr(xlRow + 5) & ":" & eRange_Jun & CStr(xlRow + 5)).Formula = "=" & sRange_Jun & CStr(xlRow - 11) & " - " & sRange_Jun & CStr(xlRow - 3)
        .Range(sRange_Jun & CStr(xlRow + 6) & ":" & eRange_Jun & CStr(xlRow + 6)).Merge
        .Range(sRange_Jun & CStr(xlRow + 6) & ":" & eRange_Jun & CStr(xlRow + 6)).Formula = "=" & sRange_Jun & CStr(xlRow - 10) & " - " & sRange_Jun & CStr(xlRow - 2)
        
      
        .Range(sRange_Jul & CStr(xlRow + 1) & ":" & eRange_Jul & CStr(xlRow + 1)).Merge
        .Range(sRange_Jul & CStr(xlRow + 1) & ":" & eRange_Jul & CStr(xlRow + 1)).Formula = "=" & sRange_Jul & CStr(xlRow - 15) & " - " & sRange_Jul & CStr(xlRow - 7)
        .Range(sRange_Jul & CStr(xlRow + 2) & ":" & eRange_Jul & CStr(xlRow + 2)).Merge
        .Range(sRange_Jul & CStr(xlRow + 2) & ":" & eRange_Jul & CStr(xlRow + 2)).Formula = "=" & sRange_Jul & CStr(xlRow - 14) & " - " & sRange_Jul & CStr(xlRow - 6)
        .Range(sRange_Jul & CStr(xlRow + 3) & ":" & eRange_Jul & CStr(xlRow + 3)).Merge
        .Range(sRange_Jul & CStr(xlRow + 3) & ":" & eRange_Jul & CStr(xlRow + 3)).Formula = "=" & sRange_Jul & CStr(xlRow - 13) & " - " & sRange_Jul & CStr(xlRow - 5)
        .Range(sRange_Jul & CStr(xlRow + 4) & ":" & eRange_Jul & CStr(xlRow + 4)).Merge
        .Range(sRange_Jul & CStr(xlRow + 4) & ":" & eRange_Jul & CStr(xlRow + 4)).Formula = "=" & sRange_Jul & CStr(xlRow - 12) & " - " & sRange_Jul & CStr(xlRow - 4)
        .Range(sRange_Jul & CStr(xlRow + 5) & ":" & eRange_Jul & CStr(xlRow + 5)).Merge
        .Range(sRange_Jul & CStr(xlRow + 5) & ":" & eRange_Jul & CStr(xlRow + 5)).Formula = "=" & sRange_Jul & CStr(xlRow - 11) & " - " & sRange_Jul & CStr(xlRow - 3)
        .Range(sRange_Jul & CStr(xlRow + 6) & ":" & eRange_Jul & CStr(xlRow + 6)).Merge
        .Range(sRange_Jul & CStr(xlRow + 6) & ":" & eRange_Jul & CStr(xlRow + 6)).Formula = "=" & sRange_Jul & CStr(xlRow - 10) & " - " & sRange_Jul & CStr(xlRow - 2)
        
      
        .Range(sRange_Aug & CStr(xlRow + 1) & ":" & eRange_Aug & CStr(xlRow + 1)).Merge
        .Range(sRange_Aug & CStr(xlRow + 1) & ":" & eRange_Aug & CStr(xlRow + 1)).Formula = "=" & sRange_Aug & CStr(xlRow - 15) & " - " & sRange_Aug & CStr(xlRow - 7)
        .Range(sRange_Aug & CStr(xlRow + 2) & ":" & eRange_Aug & CStr(xlRow + 2)).Merge
        .Range(sRange_Aug & CStr(xlRow + 2) & ":" & eRange_Aug & CStr(xlRow + 2)).Formula = "=" & sRange_Aug & CStr(xlRow - 14) & " - " & sRange_Aug & CStr(xlRow - 6)
        .Range(sRange_Aug & CStr(xlRow + 3) & ":" & eRange_Aug & CStr(xlRow + 3)).Merge
        .Range(sRange_Aug & CStr(xlRow + 3) & ":" & eRange_Aug & CStr(xlRow + 3)).Formula = "=" & sRange_Aug & CStr(xlRow - 13) & " - " & sRange_Aug & CStr(xlRow - 5)
        .Range(sRange_Aug & CStr(xlRow + 4) & ":" & eRange_Aug & CStr(xlRow + 4)).Merge
        .Range(sRange_Aug & CStr(xlRow + 4) & ":" & eRange_Aug & CStr(xlRow + 4)).Formula = "=" & sRange_Aug & CStr(xlRow - 12) & " - " & sRange_Aug & CStr(xlRow - 4)
        .Range(sRange_Aug & CStr(xlRow + 5) & ":" & eRange_Aug & CStr(xlRow + 5)).Merge
        .Range(sRange_Aug & CStr(xlRow + 5) & ":" & eRange_Aug & CStr(xlRow + 5)).Formula = "=" & sRange_Aug & CStr(xlRow - 11) & " - " & sRange_Aug & CStr(xlRow - 3)
        .Range(sRange_Aug & CStr(xlRow + 6) & ":" & eRange_Aug & CStr(xlRow + 6)).Merge
        .Range(sRange_Aug & CStr(xlRow + 6) & ":" & eRange_Aug & CStr(xlRow + 6)).Formula = "=" & sRange_Aug & CStr(xlRow - 10) & " - " & sRange_Aug & CStr(xlRow - 2)
        
      
        .Range(sRange_Sep & CStr(xlRow + 1) & ":" & eRange_Sep & CStr(xlRow + 1)).Merge
        .Range(sRange_Sep & CStr(xlRow + 1) & ":" & eRange_Sep & CStr(xlRow + 1)).Formula = "=" & sRange_Sep & CStr(xlRow - 15) & " - " & sRange_Sep & CStr(xlRow - 7)
        .Range(sRange_Sep & CStr(xlRow + 2) & ":" & eRange_Sep & CStr(xlRow + 2)).Merge
        .Range(sRange_Sep & CStr(xlRow + 2) & ":" & eRange_Sep & CStr(xlRow + 2)).Formula = "=" & sRange_Sep & CStr(xlRow - 14) & " - " & sRange_Sep & CStr(xlRow - 6)
        .Range(sRange_Sep & CStr(xlRow + 3) & ":" & eRange_Sep & CStr(xlRow + 3)).Merge
        .Range(sRange_Sep & CStr(xlRow + 3) & ":" & eRange_Sep & CStr(xlRow + 3)).Formula = "=" & sRange_Sep & CStr(xlRow - 13) & " - " & sRange_Sep & CStr(xlRow - 5)
        .Range(sRange_Sep & CStr(xlRow + 4) & ":" & eRange_Sep & CStr(xlRow + 4)).Merge
        .Range(sRange_Sep & CStr(xlRow + 4) & ":" & eRange_Sep & CStr(xlRow + 4)).Formula = "=" & sRange_Sep & CStr(xlRow - 12) & " - " & sRange_Sep & CStr(xlRow - 4)
        .Range(sRange_Sep & CStr(xlRow + 5) & ":" & eRange_Sep & CStr(xlRow + 5)).Merge
        .Range(sRange_Sep & CStr(xlRow + 5) & ":" & eRange_Sep & CStr(xlRow + 5)).Formula = "=" & sRange_Sep & CStr(xlRow - 11) & " - " & sRange_Sep & CStr(xlRow - 3)
        .Range(sRange_Sep & CStr(xlRow + 6) & ":" & eRange_Sep & CStr(xlRow + 6)).Merge
        .Range(sRange_Sep & CStr(xlRow + 6) & ":" & eRange_Sep & CStr(xlRow + 6)).Formula = "=" & sRange_Sep & CStr(xlRow - 10) & " - " & sRange_Sep & CStr(xlRow - 2)
        
      
        .Range(sRange_Oct & CStr(xlRow + 1) & ":" & eRange_Oct & CStr(xlRow + 1)).Merge
        .Range(sRange_Oct & CStr(xlRow + 1) & ":" & eRange_Oct & CStr(xlRow + 1)).Formula = "=" & sRange_Oct & CStr(xlRow - 15) & " - " & sRange_Oct & CStr(xlRow - 7)
        .Range(sRange_Oct & CStr(xlRow + 2) & ":" & eRange_Oct & CStr(xlRow + 2)).Merge
        .Range(sRange_Oct & CStr(xlRow + 2) & ":" & eRange_Oct & CStr(xlRow + 2)).Formula = "=" & sRange_Oct & CStr(xlRow - 14) & " - " & sRange_Oct & CStr(xlRow - 6)
        .Range(sRange_Oct & CStr(xlRow + 3) & ":" & eRange_Oct & CStr(xlRow + 3)).Merge
        .Range(sRange_Oct & CStr(xlRow + 3) & ":" & eRange_Oct & CStr(xlRow + 3)).Formula = "=" & sRange_Oct & CStr(xlRow - 13) & " - " & sRange_Oct & CStr(xlRow - 5)
        .Range(sRange_Oct & CStr(xlRow + 4) & ":" & eRange_Oct & CStr(xlRow + 4)).Merge
        .Range(sRange_Oct & CStr(xlRow + 4) & ":" & eRange_Oct & CStr(xlRow + 4)).Formula = "=" & sRange_Oct & CStr(xlRow - 12) & " - " & sRange_Oct & CStr(xlRow - 4)
        .Range(sRange_Oct & CStr(xlRow + 5) & ":" & eRange_Oct & CStr(xlRow + 5)).Merge
        .Range(sRange_Oct & CStr(xlRow + 5) & ":" & eRange_Oct & CStr(xlRow + 5)).Formula = "=" & sRange_Oct & CStr(xlRow - 11) & " - " & sRange_Oct & CStr(xlRow - 3)
        .Range(sRange_Oct & CStr(xlRow + 6) & ":" & eRange_Oct & CStr(xlRow + 6)).Merge
        .Range(sRange_Oct & CStr(xlRow + 6) & ":" & eRange_Oct & CStr(xlRow + 6)).Formula = "=" & sRange_Oct & CStr(xlRow - 10) & " - " & sRange_Oct & CStr(xlRow - 2)
        
      
        .Range(sRange_Nov & CStr(xlRow + 1) & ":" & eRange_Nov & CStr(xlRow + 1)).Merge
        .Range(sRange_Nov & CStr(xlRow + 1) & ":" & eRange_Nov & CStr(xlRow + 1)).Formula = "=" & sRange_Nov & CStr(xlRow - 15) & " - " & sRange_Nov & CStr(xlRow - 7)
        .Range(sRange_Nov & CStr(xlRow + 2) & ":" & eRange_Nov & CStr(xlRow + 2)).Merge
        .Range(sRange_Nov & CStr(xlRow + 2) & ":" & eRange_Nov & CStr(xlRow + 2)).Formula = "=" & sRange_Nov & CStr(xlRow - 14) & " - " & sRange_Nov & CStr(xlRow - 6)
        .Range(sRange_Nov & CStr(xlRow + 3) & ":" & eRange_Nov & CStr(xlRow + 3)).Merge
        .Range(sRange_Nov & CStr(xlRow + 3) & ":" & eRange_Nov & CStr(xlRow + 3)).Formula = "=" & sRange_Nov & CStr(xlRow - 13) & " - " & sRange_Nov & CStr(xlRow - 5)
        .Range(sRange_Nov & CStr(xlRow + 4) & ":" & eRange_Nov & CStr(xlRow + 4)).Merge
        .Range(sRange_Nov & CStr(xlRow + 4) & ":" & eRange_Nov & CStr(xlRow + 4)).Formula = "=" & sRange_Nov & CStr(xlRow - 12) & " - " & sRange_Nov & CStr(xlRow - 4)
        .Range(sRange_Nov & CStr(xlRow + 5) & ":" & eRange_Nov & CStr(xlRow + 5)).Merge
        .Range(sRange_Nov & CStr(xlRow + 5) & ":" & eRange_Nov & CStr(xlRow + 5)).Formula = "=" & sRange_Nov & CStr(xlRow - 11) & " - " & sRange_Nov & CStr(xlRow - 3)
        .Range(sRange_Nov & CStr(xlRow + 6) & ":" & eRange_Nov & CStr(xlRow + 6)).Merge
        .Range(sRange_Nov & CStr(xlRow + 6) & ":" & eRange_Nov & CStr(xlRow + 6)).Formula = "=" & sRange_Nov & CStr(xlRow - 10) & " - " & sRange_Nov & CStr(xlRow - 2)
        
       
        .Range(sRange_Dec & CStr(xlRow + 1) & ":" & eRange_Dec & CStr(xlRow + 1)).Merge
        .Range(sRange_Dec & CStr(xlRow + 1) & ":" & eRange_Dec & CStr(xlRow + 1)).Formula = "=" & sRange_Dec & CStr(xlRow - 15) & " - " & sRange_Dec & CStr(xlRow - 7)
        .Range(sRange_Dec & CStr(xlRow + 2) & ":" & eRange_Dec & CStr(xlRow + 2)).Merge
        .Range(sRange_Dec & CStr(xlRow + 2) & ":" & eRange_Dec & CStr(xlRow + 2)).Formula = "=" & sRange_Dec & CStr(xlRow - 14) & " - " & sRange_Dec & CStr(xlRow - 6)
        .Range(sRange_Dec & CStr(xlRow + 3) & ":" & eRange_Dec & CStr(xlRow + 3)).Merge
        .Range(sRange_Dec & CStr(xlRow + 3) & ":" & eRange_Dec & CStr(xlRow + 3)).Formula = "=" & sRange_Dec & CStr(xlRow - 13) & " - " & sRange_Dec & CStr(xlRow - 5)
        .Range(sRange_Dec & CStr(xlRow + 4) & ":" & eRange_Dec & CStr(xlRow + 4)).Merge
        .Range(sRange_Dec & CStr(xlRow + 4) & ":" & eRange_Dec & CStr(xlRow + 4)).Formula = "=" & sRange_Dec & CStr(xlRow - 12) & " - " & sRange_Dec & CStr(xlRow - 4)
        .Range(sRange_Dec & CStr(xlRow + 5) & ":" & eRange_Dec & CStr(xlRow + 5)).Merge
        .Range(sRange_Dec & CStr(xlRow + 5) & ":" & eRange_Dec & CStr(xlRow + 5)).Formula = "=" & sRange_Dec & CStr(xlRow - 11) & " - " & sRange_Dec & CStr(xlRow - 3)
        .Range(sRange_Dec & CStr(xlRow + 6) & ":" & eRange_Dec & CStr(xlRow + 6)).Merge
        .Range(sRange_Dec & CStr(xlRow + 6) & ":" & eRange_Dec & CStr(xlRow + 6)).Formula = "=" & sRange_Dec & CStr(xlRow - 10) & " - " & sRange_Dec & CStr(xlRow - 2)
        
          
        'Total per Quarter
        .Range(.Cells(xlRow + 7, xlLeftCol).Address & ":" & .Cells(xlRow + 9, msf_MPInsertion.cols - 2 + xlLeftCol - 1).Address).Borders(xlInsideVertical).LineStyle = xlNone

        .Range(sRange_Jan & CStr(xlRow + 7) & ":" & .Cells(xlRow + 7, msf_MPInsertion.cols - 2 + xlLeftCol - 1).Address).Borders(xlEdgeTop).LineStyle = xlContinuous
        .Range(sRange_Jan & CStr(xlRow + 7) & ":" & .Cells(xlRow + 7, msf_MPInsertion.cols - 2 + xlLeftCol - 1).Address).Borders(xlEdgeTop).Weight = xlThick
        .Range(sRange_Jan & CStr(xlRow + 7) & ":" & .Cells(xlRow + 7, msf_MPInsertion.cols - 2 + xlLeftCol - 1).Address).Borders(xlEdgeTop).ColorIndex = 5
        
        .Range(sRange_Jan & CStr(xlRow + 7) & ":" & eRange_Mar & CStr(xlRow + 7)).Merge 'Q1
        .Range(sRange_Jan & CStr(xlRow + 7) & ":" & eRange_Mar & CStr(xlRow + 7)).Formula = "=sum(" & sRange_Jan & CStr(xlRow + 1) & ":" & eRange_Mar & CStr(xlRow + 5) & ")"
        .Range(sRange_Apr & CStr(xlRow + 7) & ":" & eRange_Jun & CStr(xlRow + 7)).Merge 'Q2
        .Range(sRange_Apr & CStr(xlRow + 7) & ":" & eRange_Jun & CStr(xlRow + 7)).Formula = "=sum(" & sRange_Apr & CStr(xlRow + 1) & ":" & eRange_Jun & CStr(xlRow + 5) & ")"
        .Range(sRange_Jul & CStr(xlRow + 7) & ":" & eRange_Sep & CStr(xlRow + 7)).Merge 'Q3
        .Range(sRange_Jul & CStr(xlRow + 7) & ":" & eRange_Sep & CStr(xlRow + 7)).Formula = "=sum(" & sRange_Jul & CStr(xlRow + 1) & ":" & eRange_Sep & CStr(xlRow + 5) & ")"
        .Range(sRange_Oct & CStr(xlRow + 7) & ":" & eRange_Dec & CStr(xlRow + 7)).Merge 'Q4
        .Range(sRange_Oct & CStr(xlRow + 7) & ":" & eRange_Dec & CStr(xlRow + 7)).Formula = "=sum(" & sRange_Oct & CStr(xlRow + 1) & ":" & eRange_Dec & CStr(xlRow + 5) & ")"

        'Grand Total 1 tahun
        .Cells(xlRow + 7, xlLeftCol + msf_MPInsertion.TextMatrix(3, msf_MPInsertion.cols - 4) + 6).Formula = "=sum(" & sRange_Jan & CStr(xlRow + 7) & ":" & eRange_Dec & CStr(xlRow + 7) & ")"


        .Range(sRange_Jan & CStr(xlRow + 7) & ":" & eRange_Dec & CStr(xlRow + 7)).Borders(xlEdgeBottom).LineStyle = xlContinuous
        .Range(sRange_Jan & CStr(xlRow + 7) & ":" & eRange_Dec & CStr(xlRow + 7)).Borders(xlEdgeBottom).Weight = xlThick
        .Range(sRange_Jan & CStr(xlRow + 7) & ":" & eRange_Dec & CStr(xlRow + 7)).Borders(xlEdgeBottom).ColorIndex = 5
        .Range(sRange_Jan & CStr(xlRow + 7) & ":" & eRange_Dec & CStr(xlRow + 7)).Borders(xlEdgeLeft).LineStyle = xlContinuous
        .Range(sRange_Jan & CStr(xlRow + 7) & ":" & eRange_Dec & CStr(xlRow + 7)).Borders(xlEdgeLeft).Weight = xlThick
        .Range(sRange_Jan & CStr(xlRow + 7) & ":" & eRange_Dec & CStr(xlRow + 7)).Borders(xlEdgeLeft).ColorIndex = 5
        .Range(sRange_Jan & CStr(xlRow + 7) & ":" & eRange_Dec & CStr(xlRow + 7)).Borders(xlEdgeRight).LineStyle = xlContinuous
        .Range(sRange_Jan & CStr(xlRow + 7) & ":" & eRange_Dec & CStr(xlRow + 7)).Borders(xlEdgeRight).Weight = xlThick
        .Range(sRange_Jan & CStr(xlRow + 7) & ":" & eRange_Dec & CStr(xlRow + 7)).Borders(xlEdgeRight).ColorIndex = 5
        .Range(sRange_Jan & CStr(xlRow + 7) & ":" & eRange_Dec & CStr(xlRow + 7)).Borders(xlInsideVertical).LineStyle = xlContinuous
        .Range(sRange_Jan & CStr(xlRow + 7) & ":" & eRange_Dec & CStr(xlRow + 7)).Borders(xlInsideVertical).Weight = xlThick
        .Range(sRange_Jan & CStr(xlRow + 7) & ":" & eRange_Dec & CStr(xlRow + 7)).Borders(xlInsideVertical).ColorIndex = 5
    End With
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

Sub SetButtonToolbar(ByVal paIsNormalMode As Boolean, picOBJ) 'TOOLBAR_AI.
'************************************************
' Procedure         : SetButtonToolbar
' Function          : TOOLBAR_AI.
' Created By        : {73 64 6B}
' Date              : 12-Apr-2015
'************************************************
Dim element
    With picButton(enButtonType.bieAdd)  'ADD. 4
        .Enabled = paIsNormalMode
        .Visible = paIsNormalMode
    End With
    With picButton(enButtonType.bieEdit) 'EDIT. 5
        .Enabled = paIsNormalMode
        .Visible = paIsNormalMode
    End With
    With picButton(enButtonType.biefind)  'EDIT. 5
        .Enabled = paIsNormalMode
        .Visible = paIsNormalMode
    End With
    With picButton(enButtonType.bieTVLayering)   'TV Layering. 61
        .Enabled = paIsNormalMode
        .Visible = paIsNormalMode
    End With
    With picButton(enButtonType.bieSummary)  'Budget Summery 62.
        .Enabled = paIsNormalMode
        .Visible = paIsNormalMode
    End With
    With picButton(enButtonType.bieResendEmail)  'SummarybyVariant 63.
        .Enabled = paIsNormalMode
        .Visible = paIsNormalMode
    End With
    With picButton(enButtonType.bieExportExcel)  'Export Excel
        .Enabled = paIsNormalMode
        .Visible = paIsNormalMode
    End With
    With picButton(enButtonType.bieClose)   'EXIT.
        .Enabled = paIsNormalMode
        .Visible = paIsNormalMode
    End With

    With picButton(enButtonType.biePublishToWeb)  'SAVE.
        .Enabled = paIsNormalMode
        .Visible = paIsNormalMode

    End With


    'fr_UserDetail.Enabled = Not paIsNormalMode
    'fr_Privillage.Enabled = Not paIsNormalMode
    'pnlFilter.Enabled = paIsNormalMode
    'fraView.Enabled = paIsNormalMode
For Each element In picOBJ
    SetPictureTB element.Index, paIsNormalMode, picOBJ
Next element
End Sub

Private Sub Form_Resize()
    AdjustSizeForm
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
Private Sub picButton_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single) 'TOOLBAR_AI.
'************************************************
' Procedure         : picButton_MouseDown
' Function          : TOOLBAR_AI saat mouse ditekan.
' Created By        : {73 64 6B}
' Date              : 12-Apr-2015
'************************************************
    picButton(Index).Picture = LoadPicture(SetButtonImageEffect(Index, bieDown)) 'FIRST.

End Sub

Private Sub picButton_Click(Index As Integer)
'************************************************
' Procedure         : picButton_Click
' Function          : Action utk Navigation dan CRUD.
' Created By        : {73 64 6B}
' Date              : 12-Apr-2015/{73 64 6B} --> Semua coding dan query sudah di optimalkan agar faster, readable, safer, standardable.
'************************************************
   
   ' nCurrentPos = 'rsRecord.AbsolutePosition

    Select Case Index
        Case enButtonType.bieFirst 'FIRST.
            'Call Rst_MoveFirst
        Case enButtonType.biePrev  'PREV.
            'Call db_MovePrev
        Case enButtonType.bieNext 'NEXT.
            'Call db_MoveNext
        Case enButtonType.bieLast  'LAST.
            'Call db_MoveLast
        Case enButtonType.bieAdd  '4 'ADD.
            Call db_add
         Case enButtonType.bieEdit  '5 'EDIT.
            Call db_edit
        Case enButtonType.bieCopy  'SAVE.
            Call db_Copy
        Case enButtonType.biefind   '6 'DELETE.
            Call db_Find
        Case enButtonType.bieTVLayering '61 TV Layering
            Call db_TVLayering
            
        Case enButtonType.bieSummary
        'MsgBox "dijadikan satu pakai expivot -->BudgetSummary gabung SummarybyVariant"
            db_SummarybyVariant
            'Call db_BudgetSummary
        Case enButtonType.bieResendEmail
            Call db_Resend_Mail
        Case enButtonType.biePublishToWeb
            Call db_Publish_to_Web
        Case enButtonType.bieExportExcel
            Call db_ExportToExcel
        Case enButtonType.bieClose    '7 'EXIT.
            Call db_close

        Case enButtonType.bieCancel 'CANCEL.
            'Call db_cancel
        Case enButtonType.biePrint  'CANCEL.
            'Call db_print
                     
                     'Call ShowData
    End Select

End Sub

Sub AdjustSizeForm()
'<CSCM>
'********************************************************************************
'Procedure Name     : AdjustSizeForm
'Procedure Function : Control resizer
'Input Parameter    : ---
'Output Parameter   : ---
'Date               : 4/2/2016
'LastUpdate/By      : Tedi / Kreatif
'Name Before        : -
'********************************************************************************
'</CSCM>
 
    Dim nWidth As Single, nHeight As Single

    On Local Error Resume Next
    Me.Top = 0
    Me.Left = 0
    Width = mdi_Main.ScaleWidth
    Height = mdi_Main.ScaleHeight
    
    pnl_Main.Width = Me.Width - (pnl_Main.Left * 2)
    pnl_Main.Height = Me.Height - (pnl_Main.Top) - 400
    
    fra_Insertion.Width = pnl_Main.Width - (fra_Insertion.Left * 2)
    fra_Insertion.Height = pnl_Main.Height - (fra_Insertion.Top * 2)
 
    pic_Grd.Width = fra_Insertion.Width - (pic_Grd.Left * 2)
 
    pic_MhsFlex.Width = fra_Insertion.Width - (pic_MhsFlex.Left * 2)
    pic_MhsFlex.Height = fra_Insertion.Height - (pic_MhsFlex.Top) - 200
 
    msf_MPInsertion.Width = pic_MhsFlex.Width - (msf_MPInsertion.Left * 2)
    msf_MPInsertion.Height = pic_MhsFlex.Height - (msf_MPInsertion.Top * 2)

    On Local Error GoTo 0
    
End Sub

