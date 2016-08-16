VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TabCtl32.Ocx"
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFlxGrd.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "ComDlg32.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MShflxgd.ocx"
Begin VB.Form Frm_New_IB_TV 
   Caption         =   "Unapproved Timesheet Summary"
   ClientHeight    =   9480
   ClientLeft      =   645
   ClientTop       =   3090
   ClientWidth     =   11955
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   9480
   ScaleWidth      =   11955
   ShowInTaskbar   =   0   'False
   WindowState     =   2  'Maximized
   Begin Threed.SSPanel pnlCopy_IB 
      Height          =   7065
      Left            =   9330
      TabIndex        =   74
      Top             =   2985
      Visible         =   0   'False
      Width           =   5055
      _Version        =   65536
      _ExtentX        =   8916
      _ExtentY        =   12462
      _StockProps     =   15
      BackColor       =   15790320
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BevelOuter      =   1
      Begin Threed.SSFrame fraCopy_IB 
         Height          =   5145
         Left            =   735
         TabIndex        =   75
         Top             =   300
         Width           =   3900
         _Version        =   65536
         _ExtentX        =   6879
         _ExtentY        =   9075
         _StockProps     =   14
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Begin VB.Frame Frame5 
            Height          =   1395
            Left            =   45
            TabIndex        =   78
            Top             =   255
            Width           =   3795
            Begin VB.OptionButton Opt_New 
               Appearance      =   0  'Flat
               BackColor       =   &H80000004&
               Caption         =   "Create New IB"
               ForeColor       =   &H80000008&
               Height          =   360
               Left            =   1275
               TabIndex        =   80
               Top             =   330
               Value           =   -1  'True
               Width           =   2280
            End
            Begin VB.OptionButton Opt_Copy 
               Appearance      =   0  'Flat
               BackColor       =   &H80000004&
               Caption         =   "Copy From IB"
               ForeColor       =   &H80000008&
               Height          =   360
               Left            =   1275
               TabIndex        =   79
               Top             =   735
               Width           =   2280
            End
         End
         Begin VB.CommandButton Cmd_Ok 
            Caption         =   "&Ok"
            Height          =   480
            Left            =   60
            TabIndex        =   77
            Top             =   4560
            Width           =   3780
         End
         Begin VB.ListBox Lst_IB_TV 
            Appearance      =   0  'Flat
            Height          =   2730
            Left            =   60
            Style           =   1  'Checkbox
            TabIndex        =   76
            Top             =   1770
            Width           =   3765
         End
      End
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
      ScaleWidth      =   11955
      TabIndex        =   8
      Top             =   9150
      Width           =   11955
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
         TabIndex        =   13
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
         Index           =   2
         Left            =   750
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
         Index           =   3
         Left            =   1080
         ScaleHeight     =   345
         ScaleWidth      =   300
         TabIndex        =   10
         Top             =   15
         Width           =   300
      End
      Begin VB.PictureBox picDescColor 
         BorderStyle     =   0  'None
         Height          =   345
         Left            =   9975
         ScaleHeight     =   345
         ScaleWidth      =   1695
         TabIndex        =   9
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
         Left            =   1470
         TabIndex        =   15
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
         Left            =   4080
         TabIndex        =   14
         Tag             =   "Last Modified by: "
         Top             =   75
         Width           =   2730
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
      ScaleWidth      =   11955
      TabIndex        =   1
      Top             =   0
      Width           =   11955
      Begin VB.PictureBox picButton 
         Appearance      =   0  'Flat
         BackColor       =   &H00F0F0F0&
         BorderStyle     =   0  'None
         ClipControls    =   0   'False
         ForeColor       =   &H80000008&
         Height          =   750
         Index           =   8
         Left            =   4680
         ScaleHeight     =   750
         ScaleWidth      =   1500
         TabIndex        =   72
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
         TabIndex        =   7
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
         Index           =   9
         Left            =   6210
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
         Index           =   11
         Left            =   9270
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
         Index           =   4
         Left            =   90
         ScaleHeight     =   750
         ScaleWidth      =   1500
         TabIndex        =   2
         Top             =   -15
         Width           =   1500
      End
   End
   Begin Threed.SSPanel pnlMain 
      Height          =   8190
      Left            =   75
      TabIndex        =   0
      Top             =   810
      Width           =   10905
      _Version        =   65536
      _ExtentX        =   19235
      _ExtentY        =   14446
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
      BorderWidth     =   0
      BevelOuter      =   1
      Begin VB.Frame Frame11 
         BackColor       =   &H00F0F0F0&
         Height          =   945
         Left            =   7245
         TabIndex        =   66
         Top             =   6090
         Width           =   2136
         Begin VB.TextBox Txt_Enterd_By 
            Appearance      =   0  'Flat
            Enabled         =   0   'False
            Height          =   285
            Left            =   675
            TabIndex        =   67
            Top             =   555
            Width           =   1368
         End
         Begin MSComCtl2.DTPicker DT_Date 
            Height          =   315
            Left            =   675
            TabIndex        =   68
            Top             =   195
            Width           =   1365
            _ExtentX        =   2408
            _ExtentY        =   556
            _Version        =   393216
            Enabled         =   0   'False
            Format          =   105775105
            CurrentDate     =   36805
         End
         Begin VB.Label Label2 
            BackColor       =   &H00F0F0F0&
            Caption         =   "By "
            Height          =   195
            Left            =   120
            TabIndex        =   70
            Top             =   570
            Width           =   540
         End
         Begin VB.Label Label9 
            AutoSize        =   -1  'True
            BackColor       =   &H00F0F0F0&
            Caption         =   "Date "
            Height          =   195
            Left            =   120
            TabIndex        =   69
            Top             =   240
            Width           =   390
         End
      End
      Begin VB.Frame Fra_Approval 
         BackColor       =   &H00F0F0F0&
         Caption         =   "Client Approval"
         ForeColor       =   &H000000FF&
         Height          =   855
         Left            =   9390
         TabIndex        =   62
         ToolTipText     =   "Double Click to Approve"
         Top             =   6090
         Width           =   1980
         Begin VB.Label Lbl_Approval_Status 
            Alignment       =   2  'Center
            BackColor       =   &H00F0F0F0&
            Caption         =   "Unapproved"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   300
            Left            =   30
            TabIndex        =   64
            Top             =   225
            Width           =   1815
         End
         Begin VB.Label Lbl_Approval_Date 
            Alignment       =   2  'Center
            BackColor       =   &H00F0F0F0&
            Caption         =   "Date"
            ForeColor       =   &H000000FF&
            Height          =   285
            Left            =   60
            TabIndex        =   63
            Top             =   510
            Width           =   2145
         End
      End
      Begin VB.Frame Frame6 
         BackColor       =   &H00F0F0F0&
         Caption         =   "Montly Budget"
         ForeColor       =   &H000000FF&
         Height          =   660
         Left            =   30
         TabIndex        =   48
         Top             =   5400
         Width           =   10530
         Begin VB.TextBox Txt_Month_2 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            Height          =   315
            Left            =   4740
            TabIndex        =   51
            Top             =   225
            Width           =   1830
         End
         Begin VB.TextBox Txt_Month_3 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            Height          =   315
            Left            =   8535
            TabIndex        =   50
            Top             =   210
            Width           =   1830
         End
         Begin VB.TextBox Txt_Month_1 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            Height          =   315
            Left            =   1290
            TabIndex        =   49
            Top             =   225
            Width           =   1830
         End
         Begin VB.Label Lbl_Month_3 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00F0F0F0&
            Caption         =   "September"
            Height          =   330
            Left            =   7290
            TabIndex        =   57
            Top             =   270
            Width           =   1065
         End
         Begin VB.Label Lbl_Month_2 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00F0F0F0&
            Caption         =   "September"
            Height          =   330
            Left            =   3525
            TabIndex        =   56
            Top             =   270
            Width           =   1065
         End
         Begin VB.Label Lbl_Month_1 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00F0F0F0&
            Caption         =   "September"
            Height          =   330
            Left            =   90
            TabIndex        =   55
            Top             =   270
            Width           =   1065
         End
         Begin VB.Label L2 
            BackColor       =   &H00F0F0F0&
            Caption         =   ":"
            Height          =   315
            Left            =   4650
            TabIndex        =   54
            Top             =   270
            Width           =   60
         End
         Begin VB.Label L1 
            BackColor       =   &H00F0F0F0&
            Caption         =   ":"
            Height          =   315
            Left            =   1215
            TabIndex        =   53
            Top             =   255
            Width           =   60
         End
         Begin VB.Label L3 
            BackColor       =   &H00F0F0F0&
            Caption         =   ":"
            Height          =   315
            Left            =   8460
            TabIndex        =   52
            Top             =   255
            Width           =   60
         End
      End
      Begin VB.Frame Frame3 
         BackColor       =   &H00F0F0F0&
         Caption         =   "Campaign Outline"
         ForeColor       =   &H000000FF&
         Height          =   1800
         Left            =   30
         TabIndex        =   41
         Top             =   3585
         Width           =   10515
         Begin VB.TextBox Txt_Spot 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   2655
            MaxLength       =   4
            TabIndex        =   46
            Top             =   975
            Visible         =   0   'False
            Width           =   1035
         End
         Begin VB.CommandButton Cmd_Cancel_C_O 
            Caption         =   "Cancel"
            Enabled         =   0   'False
            Height          =   345
            Left            =   9015
            TabIndex        =   45
            Top             =   972
            Width           =   1320
         End
         Begin VB.CommandButton Cmd_Save_C_O 
            Caption         =   "Save"
            Enabled         =   0   'False
            Height          =   360
            Left            =   9015
            TabIndex        =   44
            Top             =   612
            Width           =   1320
         End
         Begin VB.CommandButton Cmd_Edit_C_O 
            Caption         =   "Edit"
            Enabled         =   0   'False
            Height          =   360
            Left            =   9015
            TabIndex        =   43
            Top             =   252
            Width           =   1320
         End
         Begin VB.CommandButton Cmd_Delete_C_O 
            Caption         =   "Delete"
            Height          =   324
            Left            =   9012
            TabIndex        =   42
            Top             =   1332
            Width           =   1320
         End
         Begin MSFlexGridLib.MSFlexGrid Flex_W_C 
            Height          =   1455
            Left            =   1155
            TabIndex        =   65
            Top             =   270
            Width           =   9045
            _ExtentX        =   15954
            _ExtentY        =   2566
            _Version        =   393216
            Rows            =   3
            Cols            =   14
            FixedRows       =   2
            FixedCols       =   0
            Enabled         =   -1  'True
            AllowUserResizing=   1
            Appearance      =   0
         End
         Begin VB.Label Label1 
            BackColor       =   &H00F0F0F0&
            Caption         =   "Material Duration  "
            Height          =   390
            Left            =   435
            TabIndex        =   47
            Top             =   780
            Width           =   1020
            WordWrap        =   -1  'True
         End
      End
      Begin VB.Frame Frame2 
         BackColor       =   &H00F0F0F0&
         Caption         =   "Television Objective"
         ForeColor       =   &H000000FF&
         Height          =   2160
         Left            =   30
         TabIndex        =   36
         Top             =   1410
         Width           =   10500
         Begin VB.CommandButton Cmd_Add_T_O 
            Caption         =   "Add"
            Height          =   425
            Left            =   9045
            TabIndex        =   39
            Top             =   570
            Width           =   1290
         End
         Begin VB.CommandButton Cmd_Edit_T_O 
            Caption         =   "Edit"
            Height          =   425
            Left            =   9045
            TabIndex        =   38
            Top             =   1005
            Width           =   1290
         End
         Begin VB.CommandButton Cmd_Delete_T_O 
            Caption         =   "Delete"
            Height          =   425
            Left            =   9045
            TabIndex        =   37
            Top             =   1425
            Width           =   1290
         End
         Begin MSHierarchicalFlexGridLib.MSHFlexGrid Flex_T_O 
            Height          =   1620
            Left            =   165
            TabIndex        =   40
            Top             =   390
            Width           =   8730
            _ExtentX        =   15399
            _ExtentY        =   2858
            _Version        =   393216
            Cols            =   6
            BackColorBkg    =   8421504
            Appearance      =   0
            _NumberOfBands  =   1
            _Band(0).Cols   =   6
         End
      End
      Begin VB.Frame Frame4 
         BackColor       =   &H00F0F0F0&
         ForeColor       =   &H000000FF&
         Height          =   1350
         Left            =   4830
         TabIndex        =   23
         Top             =   45
         Width           =   6570
         Begin VB.CommandButton cmd_Find 
            Caption         =   "..."
            Height          =   315
            Left            =   2925
            TabIndex        =   73
            Top             =   915
            Width           =   360
         End
         Begin VB.TextBox txt_Primary 
            Appearance      =   0  'Flat
            Height          =   315
            Left            =   4830
            TabIndex        =   29
            Top             =   250
            Width           =   1380
         End
         Begin VB.ComboBox Cbo_Cluster 
            Appearance      =   0  'Flat
            Height          =   315
            Left            =   4830
            TabIndex        =   28
            Top             =   590
            Width           =   1410
         End
         Begin VB.ComboBox cbo_Client_Brief_Id 
            Appearance      =   0  'Flat
            Height          =   315
            Left            =   1500
            Style           =   2  'Dropdown List
            TabIndex        =   27
            Top             =   585
            Width           =   1800
         End
         Begin VB.TextBox txt_IB_Id 
            Appearance      =   0  'Flat
            Enabled         =   0   'False
            Height          =   315
            Left            =   1500
            TabIndex        =   26
            Top             =   920
            Width           =   1395
         End
         Begin VB.ComboBox Cbo_Month 
            Appearance      =   0  'Flat
            Height          =   315
            Left            =   1500
            Style           =   2  'Dropdown List
            TabIndex        =   25
            Top             =   250
            Width           =   1800
         End
         Begin VB.TextBox Txt_Plan_No 
            Appearance      =   0  'Flat
            Height          =   315
            Left            =   4830
            MaxLength       =   19
            TabIndex        =   24
            Top             =   920
            Width           =   1380
         End
         Begin VB.Label Label5 
            BackColor       =   &H00F0F0F0&
            Caption         =   "Secondary "
            Height          =   315
            Left            =   3770
            TabIndex        =   35
            Top             =   620
            Width           =   1040
         End
         Begin VB.Label Label4 
            BackColor       =   &H00F0F0F0&
            Caption         =   "Primary "
            Height          =   330
            Left            =   3770
            TabIndex        =   34
            Top             =   310
            Width           =   1040
         End
         Begin VB.Label Label8 
            AutoSize        =   -1  'True
            BackColor       =   &H00F0F0F0&
            Caption         =   "IB Id "
            Height          =   195
            Left            =   180
            TabIndex        =   33
            Top             =   980
            Width           =   390
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            BackColor       =   &H00F0F0F0&
            Caption         =   "Client Brief Id "
            Height          =   195
            Left            =   180
            TabIndex        =   32
            Top             =   620
            Width           =   1020
         End
         Begin VB.Label Label13 
            AutoSize        =   -1  'True
            BackColor       =   &H00F0F0F0&
            Caption         =   "Starting Month "
            Height          =   195
            Left            =   180
            TabIndex        =   31
            Top             =   310
            Width           =   1110
         End
         Begin VB.Label Label10 
            BackColor       =   &H00F0F0F0&
            Caption         =   "Plan No "
            Height          =   330
            Left            =   3770
            TabIndex        =   30
            Top             =   980
            Width           =   1040
         End
      End
      Begin VB.Frame Frame1 
         BackColor       =   &H00F0F0F0&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1356
         Left            =   30
         TabIndex        =   16
         Top             =   45
         Width           =   4785
         Begin VB.ComboBox Cbo_Brand 
            Appearance      =   0  'Flat
            Height          =   315
            Left            =   1590
            Style           =   2  'Dropdown List
            TabIndex        =   19
            Top             =   250
            Width           =   2910
         End
         Begin VB.ComboBox Cbo_Brand_Variant 
            Appearance      =   0  'Flat
            Height          =   315
            Left            =   1590
            Style           =   2  'Dropdown List
            TabIndex        =   18
            Top             =   590
            Width           =   2910
         End
         Begin VB.ComboBox Cbo_Year 
            Appearance      =   0  'Flat
            Height          =   315
            Left            =   1590
            Style           =   2  'Dropdown List
            TabIndex        =   17
            Top             =   920
            Width           =   1050
         End
         Begin VB.Label Label6 
            AutoSize        =   -1  'True
            BackColor       =   &H00F0F0F0&
            Caption         =   "&Brand "
            Height          =   195
            Left            =   285
            TabIndex        =   22
            Top             =   310
            Width           =   465
         End
         Begin VB.Label Label7 
            BackColor       =   &H00F0F0F0&
            Caption         =   "Brand Variant "
            Height          =   195
            Left            =   285
            TabIndex        =   21
            Top             =   630
            Width           =   1260
         End
         Begin VB.Label Label14 
            AutoSize        =   -1  'True
            BackColor       =   &H00F0F0F0&
            Caption         =   "Year "
            Height          =   195
            Left            =   285
            TabIndex        =   20
            Top             =   980
            Width           =   375
         End
      End
      Begin MSComDlg.CommonDialog CommonDialog1 
         Left            =   0
         Top             =   30
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
      Begin TabDlg.SSTab SSTab1 
         Height          =   1185
         Left            =   30
         TabIndex        =   58
         Top             =   6090
         Width           =   7185
         _ExtentX        =   12674
         _ExtentY        =   2090
         _Version        =   393216
         Tab             =   2
         TabHeight       =   706
         BackColor       =   15790320
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         TabCaption(0)   =   "Programe Type Considerations"
         TabPicture(0)   =   "Frm_New_IB_TV1.frx":0000
         Tab(0).ControlEnabled=   0   'False
         Tab(0).Control(0)=   "Txt_Prog_Consideration"
         Tab(0).ControlCount=   1
         TabCaption(1)   =   "Other Considerations"
         TabPicture(1)   =   "Frm_New_IB_TV1.frx":001C
         Tab(1).ControlEnabled=   0   'False
         Tab(1).Control(0)=   "Txt_Other_Consideration"
         Tab(1).ControlCount=   1
         TabCaption(2)   =   "Attachments"
         TabPicture(2)   =   "Frm_New_IB_TV1.frx":0038
         Tab(2).ControlEnabled=   -1  'True
         Tab(2).Control(0)=   "Txt_Attachments"
         Tab(2).Control(0).Enabled=   0   'False
         Tab(2).ControlCount=   1
         Begin VB.TextBox Txt_Attachments 
            Appearance      =   0  'Flat
            Height          =   495
            Left            =   180
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   61
            Top             =   540
            Width           =   6680
         End
         Begin VB.TextBox Txt_Prog_Consideration 
            Appearance      =   0  'Flat
            Height          =   500
            Left            =   -74835
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   60
            Top             =   540
            Width           =   6680
         End
         Begin VB.TextBox Txt_Other_Consideration 
            Appearance      =   0  'Flat
            Height          =   500
            Left            =   -74820
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   59
            Top             =   540
            Width           =   6585
         End
      End
      Begin VB.Label Lbl_Status 
         Alignment       =   2  'Center
         BackColor       =   &H00F0F0F0&
         ForeColor       =   &H000000FF&
         Height          =   420
         Left            =   7260
         TabIndex        =   71
         Top             =   7080
         Width           =   4140
      End
   End
End
Attribute VB_Name = "Frm_New_IB_TV"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public toDay As Date
Dim Ib_Copy_1 As String
Dim Ib_Copy_2 As String
Dim Ib_Copy_3 As String

'variabel bulan
Dim Month_Copy_1 As Integer
Dim Month_Copy_2 As Integer
Dim Month_Copy_3 As Integer
Dim mNew As Boolean
Dim Edit_Flag As Boolean
Dim Edit_WC_Flag As Boolean
Dim Index_Col As Integer
Dim Index_Row As Integer
Dim First_Time As Boolean
Dim Max_Col As Integer
Dim strSql As String
Public Partial_Brand As Boolean
Dim isCopy As Boolean
Public Actual_Year As Integer
Public Actual_Month As Integer
Dim rPost As Integer
Dim add_flag As Boolean
Dim Bulan1 As Integer
Dim Bulan2 As Integer
Dim Bulan3 As Integer

Dim No_Record As Boolean
Dim Week_Count_Month1 As Integer
Dim Week_Count_Month2 As Integer
Dim Week_Count_Month3 As Integer
Dim counter As Integer

Public Week_Commencing As String
Public Campaign_Type As String
    
'Real Recordset
Public rs_IB_TV As New ADODB.Recordset
Dim Rs_Camp_Outline As New ADODB.Recordset
Dim Rs_Television_Obj As New ADODB.Recordset

' Dim Rs_Material_Mix As New ADODB.Recordset

Dim Rs_TV_Material As New ADODB.Recordset
Dim Rs_Montly_Budget As New ADODB.Recordset

'Media Type
Dim Media_Type_Code As String

'Temp Recordset

Public rsTemp_Objective As New ADODB.Recordset

'Public rsTemp_Mix As New ADODB.Recordset

Public Rs_Materi_Temp As New ADODB.Recordset
Public Rs_Temp_Camp_Outline As New ADODB.Recordset
Public BoolObjStatusEdit As Boolean

Private Sub cbo_brand_Click()
'*******************************************
' Procedure     : cbo_brand_Click
' Function      : cbo_brand_Click event Handler
' Last Update   : 18/03/2016
'*****************************************
    Dim Rs_Brand_Variant As New ADODB.Recordset
    
    
    
    Rs_Brand_Variant.Open "SELECT  *  FROM Brand_Variant " _
                        & "WHERE Brand_Code='" & Left(Cbo_Brand.Text, 4) & "'", ConnERP, adOpenStatic, adLockPessimistic
                     
    Cbo_Brand_Variant.Clear
    While Not Rs_Brand_Variant.EOF
        Cbo_Brand_Variant.AddItem Rs_Brand_Variant.Fields("Brand_Variant_Code").Value & " --> " & Rs_Brand_Variant.Fields("Brand_Variant_Name").Value
        Rs_Brand_Variant.MoveNext
    Wend
    CloseRecordset Rs_Brand_Variant
   
    If Cbo_Brand_Variant.ListCount <> 0 Then
        Cbo_Brand_Variant.Text = Cbo_Brand_Variant.List(0)
    End If
    
    
    If Get_Percent_CASC(Left(Cbo_Brand.Text, 4)) = 0 Then
        Partial_Brand = False
        'MsgBox "Full Brand"
    Else
        Partial_Brand = True
        'MsgBox "Partial Brand"
    End If
    
    If Not First_Time Then
        load_IB_TV
    End If
End Sub
Public Sub Show_Temp_Objective()
'*******************************************
' Procedure     : Show_Temp_Objective
' Function      : Menampilkan temporary objective pada Flexgrid Flex_T_O
' Last Update   : 18/03/2016
'*****************************************
    Dim Index_Array As Integer
    With rsTemp_Objective
        Flex_T_O.Rows = 1
        .Filter = ""
        
        If .RecordCount <> 0 Then
            .MoveFirst
        End If
    
        Do While Not .EOF And Not .BOF
        'Show To Grid
        
        'Add Row
            Flex_T_O.Rows = Flex_T_O.Rows + 1
            Flex_T_O.FixedRows = 1
        
            For Index_Array = 1 To 7
                Frm_New_IB_TV.Flex_T_O.Row = Flex_T_O.Rows - 1
                Frm_New_IB_TV.Flex_T_O.col = Index_Array
                
                If .Fields("Status").Value = 1 Then
                    Frm_New_IB_TV.Flex_T_O.CellBackColor = vbWhite
                    Frm_New_IB_TV.Flex_T_O.CellForeColor = vbBlack
                    Else
                        Frm_New_IB_TV.Flex_T_O.CellBackColor = vbRed
                        Frm_New_IB_TV.Flex_T_O.CellForeColor = vbWhite
                End If
            Next
        
            Flex_T_O.TextMatrix(Flex_T_O.Rows - 1, 0) = .Fields("Objective_Id").Value    'Objective id
            Flex_T_O.TextMatrix(Flex_T_O.Rows - 1, 1) = Format(.Fields("W_C_from").Value, "MMM/dd/yyyy") & " - " & Format(.Fields("W_C_To").Value, "MMM/dd/yyyy")  'Week Commencing
            Flex_T_O.TextMatrix(Flex_T_O.Rows - 1, 2) = .Fields("Campaign_Type").Value  'Campaign Type
            Flex_T_O.TextMatrix(Flex_T_O.Rows - 1, 3) = .Fields("Frequency").Value 'Frequency
            Flex_T_O.TextMatrix(Flex_T_O.Rows - 1, 4) = .Fields("Reach").Value 'Reach
            Flex_T_O.TextMatrix(Flex_T_O.Rows - 1, 5) = .Fields("Tarps").Value 'Tarps
            Flex_T_O.TextMatrix(Flex_T_O.Rows - 1, 6) = Format(.Fields("Nett").Value, "##,##0") 'Nett
            Flex_T_O.TextMatrix(Flex_T_O.Rows - 1, 7) = Format(.Fields("MSC").Value, "##,##0") 'MSC
        
            .MoveNext
        Loop
    End With
End Sub
Private Sub Cbo_Brand_Variant_Click()
'*******************************************
' Procedure     : Cbo_Brand_Variant_Click
' Function      : Cbo_Brand_Variant Click Event Handler
' Last Update   : 18/03/2016
'*****************************************
    If Not First_Time Then
       load_IB_TV
    End If
End Sub
Private Sub clearCopyIB()
'*******************************************
' Procedure     : clearCopyIB
' Function      : Clearing Copy IB reference
' Last Update   : 18/03/2016
'*****************************************
    Ib_Copy_1 = ""
    Ib_Copy_2 = ""
    Ib_Copy_3 = ""
    
    Month_Copy_1 = 0
    Month_Copy_2 = 0
    Month_Copy_3 = 0
    
    Opt_New.Value = True
    Lst_IB_TV.Clear
    Lst_IB_TV.Enabled = False
End Sub
Private Sub Cbo_Client_Brief_Id_KeyPress(KeyAscii As Integer)
'*******************************************
' Procedure     : Cbo_Client_Brief_Id_KeyPress
' Function      : Cbo_Client_Brief_Id KeyPress Event Handler
' Last Update   : 18/03/2016
'*****************************************
    KeyAscii = 0
End Sub
Private Sub Copy_From_Other_IB()
'*******************************************
' Procedure     : Copy_From_Other_IB
' Function      : Copy IB Target Audience Dari IB yang sudah ada
' Last Update   : 18/03/2016
'*****************************************
        Dim rs_Copy_Budget As New ADODB.Recordset
        Dim rs_Copy_Materi As New ADODB.Recordset
        
        Dim idx As Integer
        
        Dim Rs_Copy_target As New ADODB.Recordset
        Dim strIBID(4) As String
        Dim IntMonthCopy(4) As Integer
        
        'isi target audiance
        strSql = "Select  * From IB_TV  WHERE ib_id='" & IIf(Ib_Copy_1 = "", IIf(Ib_Copy_2 = "", IIf(Ib_Copy_3 = "", "", Ib_Copy_3), Ib_Copy_2), Ib_Copy_1) & "'"
        Rs_Copy_target.Open strSql, ConnERP, adOpenStatic, adLockPessimistic
       
        If Not Rs_Copy_target.EOF Then
            txt_Primary.Text = IIf(IsNull(Rs_Copy_target("Target_Primary")), "", Rs_Copy_target("Target_Primary"))
            Txt_Plan_No.Text = IIf(IsNull(Rs_Copy_target("Plan_No")), "", Rs_Copy_target("Plan_No"))
            Cbo_Cluster.Text = IIf(IsNull(Rs_Copy_target("Target_Secondary")), "", Rs_Copy_target("Target_Secondary"))
            Cbo_Cluster.ListIndex = IIf(IsNull(Rs_Copy_target("Cluster_code")), 0, Rs_Copy_target("Cluster_code")) - 1
        End If
        CloseRecordset Rs_Copy_target
        'masukkan data ke Monthly budget
        'month 1
        If Ib_Copy_1 <> "" Then
            strIBID(1) = Ib_Copy_1
            IntMonthCopy(1) = Month_Copy_1
            
            strSql = "SELECT  *  FROM ib_tv_montly_budget WHERE ib_id='" & Ib_Copy_1 & "' AND Month=" & Month_Copy_1
            rs_Copy_Budget.Open strSql, ConnERP, adOpenStatic, adLockPessimistic
            
            If Not rs_Copy_Budget.EOF Then
                Lbl_Month_1.Caption = Get_Month_Name(rs_Copy_Budget("month"))
                Txt_Month_1.Text = Format(rs_Copy_Budget.Fields("budget").Value, "##,##0")
            End If
            CloseRecordset rs_Copy_Budget
        End If
        
        'month 2
        If Ib_Copy_2 <> "" Then
            strIBID(2) = Ib_Copy_2
            IntMonthCopy(2) = Month_Copy_2
            
            strSql = "SELECT  *  FROM ib_tv_montly_budget WHERE ib_id='" & Ib_Copy_2 & "' AND Month=" & Month_Copy_2
            rs_Copy_Budget.Open strSql, ConnERP, adOpenStatic, adLockPessimistic
           
            If Not rs_Copy_Budget.EOF Then
                Lbl_Month_2.Caption = Get_Month_Name(rs_Copy_Budget("month"))
                Txt_Month_2.Text = Format(rs_Copy_Budget.Fields("budget").Value, "##,##0")
            End If
            CloseRecordset rs_Copy_Budget
        End If
        
        'month 3
        If Ib_Copy_3 <> "" Then
            strIBID(3) = Ib_Copy_3
            IntMonthCopy(3) = Month_Copy_3
            
            strSql = "SELECT  *  FROM ib_tv_montly_budget WHERE ib_id='" & Ib_Copy_3 & "' AND Month=" & Month_Copy_3
            rs_Copy_Budget.Open strSql, ConnERP, adOpenStatic, adLockPessimistic
            
            If Not rs_Copy_Budget.EOF Then
                    Lbl_Month_3.Caption = Get_Month_Name(rs_Copy_Budget("month"))
                    Txt_Month_3.Text = Format(rs_Copy_Budget.Fields("budget").Value, "##,##0")
            End If
            CloseRecordset rs_Copy_Budget
        End If
        '========================================
        
        'Television Objective
        Dim rs_Copy_Obj As New ADODB.Recordset
        Dim Termasuk As Boolean
        Dim Str_Sel As String
        Dim rs_WC As New ADODB.Recordset
        Dim Tgl_awal As Date
        Dim Tgl_Akhir As Date
        Dim DblObjectiveId As Double
        Dim DblCampaignID As Double
        Dim Rs_Copy_Camp As New ADODB.Recordset
        
        For idx = 1 To 3
        
            'month 1
            If strIBID(idx) <> "" Then
                
                
                strSql = " SELECT  *  FROM IB_TV_Objective  WHERE ib_id='" & strIBID(idx) & "'"
                rs_Copy_Obj.Open strSql, ConnERP, adOpenStatic, adLockPessimistic
                
                While Not rs_Copy_Obj.EOF
                    'get tanggal awal dan akhir wc
                    Termasuk = False
                    strSql = "SELECT week_1, isnull(week_5,week_4) FROM Week_Commencing WHERE month= " & IntMonthCopy(idx) & " AND year= " & Cbo_Year.Text
                    rs_WC.Open strSql, ConnERP, adOpenStatic, adLockReadOnly
                    If Not rs_WC.EOF Then
                        Tgl_awal = rs_WC(0)
                        Tgl_Akhir = rs_WC(1)
                        
                        Do
                            If rs_Copy_Obj("W_C_From") = Tgl_awal Then
                                Termasuk = True
                                Exit Do
                            End If
                            Tgl_awal = DateAdd("d", 7, Tgl_awal)
                        Loop Until Tgl_awal > Tgl_Akhir
                    End If
                    CloseRecordset rs_WC
                    
                    If Termasuk Then
                        DblObjectiveId = Get_Objective_Id
                        rsTemp_Objective.AddNew
                        rsTemp_Objective("Objective_Id") = DblObjectiveId
                        rsTemp_Objective("IB_Id") = txt_IB_Id.Text
                        rsTemp_Objective("Week_Commencing").Value = rs_Copy_Obj("W_C_From") & " - " & rs_Copy_Obj("W_C_To")
                        rsTemp_Objective("Campaign_Type") = rs_Copy_Obj("Campaign_Type")
                        rsTemp_Objective("Frequency") = rs_Copy_Obj("Frequency")
                        rsTemp_Objective("Reach") = rs_Copy_Obj("Reach")
                        rsTemp_Objective("Tarps") = rs_Copy_Obj("Tarps")
                        rsTemp_Objective("Budget_With_MSC") = rs_Copy_Obj("Budget_With_MSC")
                        rsTemp_Objective("Campaign_Type_Code") = rs_Copy_Obj("Campaign_Type_Code")
                        rsTemp_Objective("Frequency_Code") = rs_Copy_Obj("Frequency_Code")
                        rsTemp_Objective("W_C_From") = rs_Copy_Obj("W_C_From")
                        rsTemp_Objective("W_C_To") = rs_Copy_Obj("W_C_To")
                        rsTemp_Objective("Nett") = rs_Copy_Obj("Nett")
                        rsTemp_Objective("MSC") = rs_Copy_Obj("MSC")
                        rsTemp_Objective("Market_Code") = rs_Copy_Obj("Market_Code")
                        rsTemp_Objective("Market_Name") = rs_Copy_Obj("Market_Name")
                        rsTemp_Objective("Status") = 1 'rs_Copy_Obj("Status")
                        rsTemp_Objective.Update
                                            
                                            
                        'Material Objective
                        strSql = "SELECT  *  FROM IB_TV_Objective_Material  WHERE Objective_Id = " & rs_Copy_Obj.Fields("Objective_Id")
                        rs_Copy_Materi.Open strSql, ConnERP, adOpenStatic, adLockPessimistic
                    
                        While Not rs_Copy_Materi.EOF
                            Rs_Materi_Temp.AddNew
                            Rs_Materi_Temp.Fields("client_brief_id").Value = cbo_Client_Brief_Id.Text
                            Rs_Materi_Temp.Fields("IB_ID").Value = txt_IB_Id.Text
                            Rs_Materi_Temp.Fields("Objective_Id").Value = DblObjectiveId
                            Rs_Materi_Temp.Fields("Material_Name").Value = rs_Copy_Materi("Material_Name")
                            Rs_Materi_Temp.Fields("Material_Duration").Value = rs_Copy_Materi("Material_Duration")
                            Rs_Materi_Temp.Update
                            rs_Copy_Materi.MoveNext
                                'idx = idx + 1
                        Wend
                        CloseRecordset rs_Copy_Materi
                        '===============================
                        
                        'CAMPAIGN
                        strSql = "SELECT  *  FROM IB_TV_Campaign  WHERE ObJective_Id = " & rs_Copy_Obj.Fields("Objective_Id")
                        Rs_Copy_Camp.Open strSql, ConnERP, adOpenStatic, adLockPessimistic
                        While Not Rs_Copy_Camp.EOF
                                DblCampaignID = Get_Campaign_Id
                                Rs_Temp_Camp_Outline.AddNew
                                Rs_Temp_Camp_Outline.Fields("Campaign_Id").Value = DblCampaignID
                                Rs_Temp_Camp_Outline.Fields("IB_ID").Value = txt_IB_Id.Text
                                Rs_Temp_Camp_Outline.Fields("Tarps_Per_Week").Value = Rs_Copy_Camp("Tarps_Per_Week")
                                Rs_Temp_Camp_Outline.Fields("Objective_Id").Value = DblObjectiveId
                                Rs_Temp_Camp_Outline.Fields("Month_Week_commencing").Value = Rs_Copy_Camp("Month_Week_commencing")
                                Rs_Temp_Camp_Outline.Fields("Week_Commencing").Value = Rs_Copy_Camp("Week_Commencing")
                                Rs_Temp_Camp_Outline.Fields("Material_Name").Value = Rs_Copy_Camp("Material_Name")
                                Rs_Temp_Camp_Outline.Fields("Material_Duration").Value = Rs_Copy_Camp("Material_Duration")
                                Rs_Temp_Camp_Outline.Update
                                                    
                            Rs_Copy_Camp.MoveNext
                        Wend
                        CloseRecordset Rs_Copy_Camp
                        
                        '==================
                    End If
                    rs_Copy_Obj.MoveNext
                Wend
               CloseRecordset rs_Copy_Obj
            End If
        Next idx
        
        Show_Temp_Objective
        
        Show_Campaign_Outline
        
 
End Sub
Public Sub Show_Campaign_Outline()
'*******************************************
' Procedure     : Show_Campaign_Outline
' Function      : Menampilkan Data Campaign
' Last Update   : 18/03/2016
'*****************************************
    Dim intRow As Integer
    Dim intCol As Integer
    Dim DateMWC As Date
    Dim DateWC As Date
    'Campaign Outline
    
    Flex_W_C.Rows = 2
    For intCol = 1 To Flex_W_C.cols - 1
        Flex_W_C.TextMatrix(0, intCol) = ""
        Flex_W_C.TextMatrix(1, intCol) = ""
    Next intCol
    
    Initilize_Flex_Grid Actual_Month, Actual_Year
    
    
    Rs_Materi_Temp.Filter = ""
    While Not Rs_Materi_Temp.EOF
        Flex_W_C.AddItem ""
        intRow = Flex_W_C.Rows - 1
        Flex_W_C.TextMatrix(intRow, 0) = Rs_Materi_Temp.Fields("Objective_Id").Value & ":" & Rs_Materi_Temp.Fields("Material_name").Value & ":" & Rs_Materi_Temp.Fields("Material_Duration").Value
            
        strSql = "Objective_Id= " & Rs_Materi_Temp.Fields("Objective_Id").Value
        strSql = strSql & " AND Material_Name = '" & Clear_String(Rs_Materi_Temp.Fields("Material_Name").Value) & "'"
        strSql = strSql & " AND Material_Duration = " & CDbl(Rs_Materi_Temp.Fields("Material_Duration").Value)
        
        Rs_Temp_Camp_Outline.Filter = ""
        Rs_Temp_Camp_Outline.Filter = strSql
        While Not Rs_Temp_Camp_Outline.EOF
            For intCol = 1 To Flex_W_C.cols - 1
                If Flex_W_C.TextMatrix(1, intCol) <> "" Then
                    DateMWC = Get_Month_Number(Flex_W_C.TextMatrix(0, intCol)) & "/" & "01" & "/" & Actual_Year
                    DateWC = Get_Actual_Month(CInt(Flex_W_C.TextMatrix(1, intCol)), CInt(Flex_W_C.TextMatrix(1, intCol)), Get_Month_Number(Flex_W_C.TextMatrix(0, intCol)), Actual_Year) & "/" & Flex_W_C.TextMatrix(1, intCol) & "/" & Get_Actual_Year(CInt(Flex_W_C.TextMatrix(1, intCol)), CInt(Flex_W_C.TextMatrix(1, intCol)), Get_Month_Number(Flex_W_C.TextMatrix(0, intCol)), Actual_Year)
                    
                    If DateMWC = Rs_Temp_Camp_Outline.Fields("Month_Week_commencing").Value And _
                        DateWC = Rs_Temp_Camp_Outline.Fields("Week_Commencing").Value Then
                        Flex_W_C.TextMatrix(intRow, intCol) = Rs_Temp_Camp_Outline.Fields("Tarps_Per_Week").Value
                    End If
                    
                End If
            Next
            
            Rs_Temp_Camp_Outline.MoveNext
        Wend
        
        Rs_Temp_Camp_Outline.Filter = ""
        Rs_Materi_Temp.MoveNext
    Wend
    Rs_Materi_Temp.Filter = ""
End Sub
Private Sub showCopyIB()
'*******************************************
' Procedure     : showCopyIB
' Function      : Menampilkan panel Copy IB
' Last Update   : 18/03/2016
'*****************************************
    pnlCopy_IB.Visible = True
    pnlMain.Visible = False
    clearCopyIB
End Sub
Private Sub hideCopyIB()
'*******************************************
' Procedure     : hideCopyIB
' Function      : Menyembunyikan panel Copy IB
' Last Update   : 18/03/2016
'*****************************************
    clearCopyIB
    pnlCopy_IB.Visible = False
    pnlMain.Visible = True

End Sub
Private Sub cbo_Client_Brief_Id_LostFocus()
 '************************************************
' Procedure         : txt_IB_Id_GotFocus
' Function          : Generate Brief Id
' Date              : 01/01/2001
' Parameter Input   : Enable
' Parameter Output  :
' Last Update/By    : 11/02/2002
'************************************************
    Dim New_IB_Id As String
    Dim tahun As Integer
    Dim Bulan As String
    Dim Brand_code  As String
    Dim Rs_Check_IV_TV As New ADODB.Recordset
   
    If cbo_Client_Brief_Id.Text = "" Then
       If cbo_Client_Brief_Id.ListCount <> 0 Then
            MsgBox "Please Select  Client Brief Id..", vbExclamation, APPLICATION_NAME
            cbo_Client_Brief_Id.SetFocus
            Exit Sub
        Else
            Exit Sub
        End If
    End If
    
    'generate Brief Id
    Actual_Year = Cbo_Year.Text
    Actual_Month = Get_Month_Number(Cbo_Month.Text)
    
    
    If Get_WeekCommencing(Actual_Month, Actual_Year) = "" Then
        MsgBox "Week Commencing Not Found, Please Contact IT Department !", vbCritical, APPLICATION_NAME
        Exit Sub
    End If
    
    

    If mNew Then
        Dim Out_Param As ADODB.Parameter
        Dim In_Param1 As ADODB.Parameter
        Dim In_Param2 As ADODB.Parameter
        Dim In_Param3 As ADODB.Parameter
        Dim cmd As New ADODB.Command
        
        Brand_code = Left(Cbo_Brand.Text, 4)
        tahun = Val(Cbo_Year.Text)
                        
        txt_IB_Id.Text = getID("IBI", Brand_code, Media_Type_Code, tahun)
            
        showCopyIB

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
Sub SetButtonToolbar(ByVal paIsNormalMode As Boolean, picOBJ) 'TOOLBAR_AI.
'************************************************
' Procedure         : SetButtonToolbar
' Function          : TOOLBAR_AI.
' Created By        : {73 64 6B}
' Date              : 12-Apr-2015
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
    With picButton(enButtonType.bieExit)    'Quit.
        .Enabled = paIsNormalMode
        .Visible = paIsNormalMode
    End With
     With picButton(enButtonType.biePrint)      'Quit.
        .Enabled = paIsNormalMode
        .Visible = paIsNormalMode
    End With
    With picButton(enButtonType.bieSave)  'SAVE.
        .Enabled = Not paIsNormalMode
        .Visible = Not paIsNormalMode
        .Left = picButton(4).Left
    End With
    With picButton(enButtonType.bieCancel) 'CANCEL.
        .Enabled = Not paIsNormalMode
        .Visible = Not paIsNormalMode
        .Left = picButton(5).Left
    End With
For Each element In picOBJ
    SetPictureTB element.Index, paIsNormalMode, picOBJ
Next element
    
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
Sub SetPictureTB(ByVal Index As Integer, ByVal paIsNormalMode As Boolean, picOBJ)
   With picOBJ(Index) 'FIRST.
        
        If .Enabled = True Then
            .Picture = LoadPicture(SetButtonImageEffect(Index, bieNormal))
        Else: .Picture = LoadPicture(SetButtonImageEffect(Index, bieDisabled))
        End If
    End With
End Sub
Public Sub Television_Objective_Button(Tampil As Boolean)
'*******************************************
' Procedure     : Television_Objective_Button
' Function      : Button Authorize
' Last Update   : 18/03/2016
'*****************************************
    Cmd_Add_T_O.Enabled = Tampil
    Cmd_Delete_T_O.Enabled = Tampil
    Cmd_Edit_T_O.Enabled = Tampil
End Sub

Private Sub Cbo_Month_Click()
'*******************************************
' Procedure     : Cbo_Month_Click
' Function      : Cbo_Month Click Event Handler
' Last Update   : 18/03/2016
'*****************************************

 If Not First_Time Then
      If mNew Then
        Dim month As Integer
        Dim Year As Integer
        
        Year = CInt(Cbo_Year.Text)
        month = CInt(Get_Month_Number(Cbo_Month.Text))
        Initilize_Flex_Grid month, Year
       End If
    End If
End Sub

Private Sub Cbo_Year_Click()
'*******************************************
' Procedure     : Cbo_Year_Click
' Function      : Cbo_Year Click Event Handler
' Last Update   : 18/03/2016
'*****************************************
    If Not First_Time Then
       load_IB_TV
    End If
End Sub

Private Sub Cmd_Add_T_O_Click()
'*******************************************
' Procedure     : Cmd_Add_T_O_Click
' Function      : Cmd_Add_T_O Click Event Handler
' Last Update   : 18/03/2016
'*****************************************

    Frm_Get_Week_Commencing.DblObjId = Get_Objective_Id
    Frm_Get_Week_Commencing.BoolAddFlag = True
    Frm_Get_Week_Commencing.show 1
End Sub

Private Sub Cmd_Delete_T_O_Click()
'*******************************************
' Procedure     : Cmd_Delete_T_O_Click
' Function      : Cmd_Delete_T_O Click Event Handler
' Last Update   : 18/03/2016
'*****************************************
    Dim Week_Commencing As String
    Dim Campaign_Type As String
    Dim Ada_Pilihan As Boolean
    Dim DblObjIdDel As Double
    Dim rsCekObj As New ADODB.Recordset
    
   
    Ada_Pilihan = False
    If mNew Or Edit_Flag Then
        For Index_Row = 1 To Flex_T_O.Rows - 1
        
        'Cari data yang baris biru
            Flex_T_O.col = 1
            Flex_T_O.Row = Index_Row
            If Flex_T_O.CellBackColor = &HFFC0C0 Then
                DblObjIdDel = CDbl(Flex_T_O.TextMatrix(Index_Row, 0))
                Ada_Pilihan = True
                Exit For
            End If
        Next Index_Row
        
        If Not Ada_Pilihan Then
            MsgBox "Select objective to delete !", vbExclamation, APPLICATION_NAME
            Exit Sub
        End If

        If MsgBox("Are you sure to delete this objective including campaign ?", vbYesNo + vbQuestion, APPLICATION_NAME) = vbYes Then
            With Rs_Temp_Camp_Outline
                .Filter = ""
                .Filter = "Objective_Id =" & DblObjIdDel
                
                While Not .EOF
                    .Delete
                    .MoveNext
                Wend
                .Filter = ""
            End With
            
            With Rs_Materi_Temp
                .Filter = ""
                .Filter = "Objective_Id =" & DblObjIdDel
                
                While Not .EOF
                    .Delete
                    .MoveNext
                Wend
                .Filter = ""
            End With
            
            With rsTemp_Objective
                .Filter = ""
                .Filter = "Objective_Id =" & DblObjIdDel
                
                While Not .EOF
                    .Delete
                    .MoveNext
                Wend
                .Filter = ""
            End With
            
            Show_Temp_Objective
            Show_Campaign_Outline
            
        End If
    End If
            
End Sub

Private Sub Cmd_Edit_T_O_Click()
'*******************************************
' Procedure     : Cmd_Edit_T_O_Click
' Function      : Cmd_Edit_T_O Click Event Handler
' Last Update   : 18/03/2016
'*****************************************
    Dim Ada_Pilihan As Boolean
    
    Ada_Pilihan = False
    If mNew Or Edit_Flag Then
        For Index_Row = 1 To Flex_T_O.Rows - 1
        
        'Cari data yang baris biru
            Flex_T_O.col = 1
            Flex_T_O.Row = Index_Row
            If Flex_T_O.CellBackColor = &HFFC0C0 Then
                Frm_Get_Week_Commencing.DblObjId = Flex_T_O.TextMatrix(Index_Row, 0)
                Ada_Pilihan = True
                Exit For
            End If
        Next Index_Row
    End If
    
   If Ada_Pilihan Then
        Frm_Get_Week_Commencing.BoolAddFlag = False
        Frm_Get_Week_Commencing.show 1
   Else
        MsgBox "Select objective to edit !", vbExclamation, APPLICATION_NAME
   End If
End Sub

Private Sub Cmd_Find_Click()
'*******************************************
' Procedure     : cmd_Find_Click
' Function      : cmd_Find Click Event Handler
' Last Update   : 18/03/2016
'*****************************************
    txt_IB_Id.Enabled = True
    txt_IB_Id.SetFocus
End Sub

Private Sub Cmd_Ok_Click()
'*******************************************
' Procedure     : Cmd_Ok_Click
' Function      : Cmd_Ok Click Event Handler
' Last Update   : 18/03/2016
'*****************************************
    Dim idx As Integer
    Dim Count_Select As Integer
    
    Dim IsError As Boolean
    IsError = False
    Count_Select = 0
    If Me.Opt_Copy.Value Then
           isCopy = True
            
    Else
        isCopy = False
    End If
     Cbo_Month.Enabled = False
    cbo_Client_Brief_Id.Enabled = False
    Television_Objective_Button True
    
    If isCopy Then
      Copy_From_Other_IB
        Cmd_Edit_C_O.Enabled = True
         
         Call EnableObject(True)
    End If
    hideCopyIB
End Sub

Private Sub Flex_T_O_Click()
'*******************************************
' Procedure     : Flex_T_O_Click
' Function      : Flex_T_O Click Event Handler
' Last Update   : 18/03/2016
'*****************************************
    Dim set_Row As Integer
    
    set_Row = Flex_T_O.RowSel
    For Index_Col = 1 To Flex_T_O.cols - 1
        For Index_Row = 1 To Flex_T_O.Rows - 1
            Flex_T_O.col = Index_Col
            Flex_T_O.Row = Index_Row
            If Index_Row = set_Row Then
                If Trim(Flex_T_O.TextMatrix(Index_Row, 0)) <> "" Then
                    Flex_T_O.CellBackColor = &HFFC0C0
                End If
                
                Else
                    If Flex_T_O.CellForeColor = vbBlack Then
                        If Trim(Flex_T_O.TextMatrix(Index_Row, 0)) <> "" Then
                            Flex_T_O.CellBackColor = vbWhite
                        End If
                        
                        Else
                            If Trim(Flex_T_O.TextMatrix(Index_Row, 0)) <> "" Then
                                Flex_T_O.CellBackColor = vbRed
                            End If
                    End If
            End If
        Next Index_Row
    Next Index_Col
End Sub

Private Sub Flex_T_O_DblClick()
'*******************************************
' Procedure     : Flex_T_O_DblClick
' Function      : Flex_T_O DblClick Event Handler
' Last Update   : 18/03/2016
'*****************************************
    Dim Ada_Pilihan As Boolean
    

    If (Not mNew) And (Not Edit_Flag) Then
        
    If Lbl_Approval_Status.Caption = "UnApproved" Then Exit Sub
    
    Ada_Pilihan = False
    For Index_Row = 1 To Flex_T_O.Rows - 1
        
            'Cari data yang baris biru
        Flex_T_O.col = 1
        Flex_T_O.Row = Index_Row
        If Flex_T_O.CellBackColor = &HFFC0C0 Then
            Frm_Get_Week_Commencing.DblObjId = Flex_T_O.TextMatrix(Index_Row, 0)
            Ada_Pilihan = True
            Exit For
        End If
    Next Index_Row
       
    If Ada_Pilihan Then
            
        BoolObjStatusEdit = True
            
        Initialize_Temp_Table
            
        strSql = ""
        strSql = " SELECT  *  FROM IB_TV_Objective  WHERE "
        strSql = strSql & " IB_ID='" & txt_IB_Id.Text & "' AND Objective_Id = " & Flex_T_O.TextMatrix(Index_Row, 0)
            
        Rs_Television_Obj.Open strSql, ConnERP, adOpenStatic, adLockPessimistic
            
        Do While Not Rs_Television_Obj.EOF
        
            rsTemp_Objective.AddNew
            rsTemp_Objective.Fields("Objective_Id").Value = Rs_Television_Obj.Fields("Objective_Id").Value
            rsTemp_Objective.Fields("Client_Brief_Id").Value = Rs_Television_Obj.Fields("Client_Brief_Id").Value
            rsTemp_Objective.Fields("IB_Id").Value = Rs_Television_Obj.Fields("IB_Id").Value
            rsTemp_Objective.Fields("Week_Commencing").Value = Rs_Television_Obj.Fields("Week_Commencing").Value
            rsTemp_Objective.Fields("Campaign_Type").Value = Rs_Television_Obj.Fields("Campaign_Type").Value
            rsTemp_Objective.Fields("Frequency").Value = Rs_Television_Obj.Fields("Frequency").Value
            rsTemp_Objective.Fields("Reach").Value = Rs_Television_Obj.Fields("Reach").Value
            rsTemp_Objective.Fields("Tarps").Value = Rs_Television_Obj.Fields("Tarps").Value
            rsTemp_Objective.Fields("Budget_With_MSC").Value = Rs_Television_Obj.Fields("Budget_With_MSC").Value
            rsTemp_Objective.Fields("Campaign_Type_Code").Value = Rs_Television_Obj.Fields("Campaign_Type_Code").Value
            rsTemp_Objective.Fields("Frequency_Code").Value = Rs_Television_Obj.Fields("Frequency_Code").Value
            rsTemp_Objective.Fields("W_C_From").Value = Rs_Television_Obj.Fields("W_C_From").Value
            rsTemp_Objective.Fields("W_C_To").Value = Rs_Television_Obj.Fields("W_C_To").Value
            rsTemp_Objective.Fields("Nett").Value = Rs_Television_Obj.Fields("Nett").Value
            rsTemp_Objective.Fields("MSC").Value = Rs_Television_Obj.Fields("MSC").Value
            rsTemp_Objective.Fields("Market_Code").Value = Rs_Television_Obj.Fields("Market_Code").Value
            rsTemp_Objective.Fields("Market_Name").Value = Rs_Television_Obj.Fields("Market_Name").Value
            rsTemp_Objective.Fields("Status").Value = Rs_Television_Obj.Fields("Status").Value
            rsTemp_Objective.Update
                                
            Rs_Television_Obj.MoveNext
        Loop
            
        CloseRecordset Rs_Television_Obj
                       
        '===============================================
        'Material
        '===============================================
        strSql = ""
        strSql = " SELECT  *  FROM IB_TV_Objective_Material  WHERE  "
        strSql = strSql & " Objective_Id  IN (SELECT Objective_Id FROM Ib_TV_Objective WHERE "
        strSql = strSql & " IB_ID='" & txt_IB_Id.Text & "') ORDER BY Objective_Id  "
                
        Rs_TV_Material.Open strSql, ConnERP, adOpenStatic, adLockPessimistic
                
        Do While Not Rs_TV_Material.EOF
            Rs_Materi_Temp.AddNew
            Rs_Materi_Temp.Fields("Objective_Id").Value = Rs_TV_Material.Fields("Objective_Id").Value
            Rs_Materi_Temp.Fields("Client_Brief_Id").Value = Rs_TV_Material.Fields("Client_Brief_Id").Value
            Rs_Materi_Temp.Fields("IB_Id").Value = Rs_TV_Material.Fields("IB_Id").Value
            Rs_Materi_Temp.Fields("Material_Name").Value = Rs_TV_Material.Fields("Material_Name").Value
            Rs_Materi_Temp.Fields("Material_Duration").Value = Rs_TV_Material.Fields("Material_Duration").Value
            Rs_Materi_Temp.Update
                    
            Rs_TV_Material.MoveNext
        Loop
              
        CloseRecordset Rs_TV_Material
    
        '=================================
        'Campaign Outline
        '=================================
        strSql = ""
        strSql = " SELECT  *  FROM IB_TV_Campaign  WHERE Objective_Id IN (SELECT Objective_Id FROM IB_TV_Objective WHERE"
        strSql = strSql & " IB_Id='" & txt_IB_Id.Text & "') ORDER BY Campaign_Id"
            
        Set Rs_Camp_Outline = Nothing
            
        Rs_Camp_Outline.Open strSql, ConnERP, adOpenStatic, adLockPessimistic
        With Rs_Camp_Outline
            Do While Not .EOF
                Rs_Temp_Camp_Outline.AddNew
                Rs_Temp_Camp_Outline.Fields("Campaign_Id").Value = Rs_Camp_Outline.Fields("Campaign_Id").Value
                Rs_Temp_Camp_Outline.Fields("Client_Brief_ID").Value = Rs_Camp_Outline.Fields("Client_Brief_ID").Value
                Rs_Temp_Camp_Outline.Fields("IB_Id").Value = Rs_Camp_Outline.Fields("IB_Id").Value
                Rs_Temp_Camp_Outline.Fields("Tarps_Per_Week").Value = Rs_Camp_Outline.Fields("Tarps_Per_Week").Value
                Rs_Temp_Camp_Outline.Fields("Objective_Id").Value = Rs_Camp_Outline.Fields("Objective_Id").Value
                Rs_Temp_Camp_Outline.Fields("Month_Week_commencing").Value = Rs_Camp_Outline.Fields("Month_Week_commencing").Value
                Rs_Temp_Camp_Outline.Fields("Week_Commencing").Value = Rs_Camp_Outline.Fields("Week_Commencing").Value
                Rs_Temp_Camp_Outline.Fields("Material_Name").Value = Rs_Camp_Outline.Fields("Material_Name").Value
                Rs_Temp_Camp_Outline.Fields("Material_Duration").Value = Rs_Camp_Outline.Fields("Material_Duration").Value
                Rs_Temp_Camp_Outline.Update
                    
                .MoveNext
            Loop
            .Close
        End With
            
        Frm_Get_Week_Commencing.BoolAddFlag = False
            
        Frm_Get_Week_Commencing.show 1
            
        BoolObjStatusEdit = False
            
        Deinitialize_Temp_Table
       Else
            MsgBox "Select Week Commencing to Edit !", vbExclamation, APPLICATION_NAME
    End If
End If
End Sub

Public Sub Initialize_Temp_Table()
'*******************************************
' Procedure     : Initialize_Temp_Table
' Function      : Initialize Temporary Table
' Last Update   : 18/03/2016
'*****************************************
    ' Temp TV Objective
    Set rsTemp_Objective = Nothing
    Set rsTemp_Objective = New ADODB.Recordset
    With rsTemp_Objective.Fields
        .Append "Objective_ID", adDouble, , adFldIsNullable
        .Append "Client_Brief_Id", adVarChar, 13, adFldIsNullable
        .Append "IB_Id", adVarChar, 13, adFldIsNullable
        .Append "Week_Commencing", adVarChar, 23, adFldIsNullable
        .Append "Campaign_Type", adVarChar, 30, adFldIsNullable
        .Append "Frequency", adVarChar, 10, adFldIsNullable
        .Append "Reach", adDecimal, , adFldIsNullable
        .Append "TARPS", adInteger, , adFldIsNullable
        .Append "Budget_With_MSC", adCurrency, , adFldIsNullable
        .Append "Campaign_Type_Code", adSmallInt, , adFldIsNullable
        .Append "Frequency_Code", adSmallInt, , adFldIsNullable
        .Append "W_C_From", adDate, , adFldIsNullable
        .Append "W_C_To", adDate, , adFldIsNullable
        
        .Append "Nett", adCurrency, , adFldIsNullable
        .Append "MSC", adCurrency, , adFldIsNullable
        .Append "Market_Code", adSmallInt, , adFldIsNullable
        .Append "Market_Name", adVarChar, 50, adFldIsNullable
        .Append "Status", adSmallInt, , adFldIsNullable
        
    End With
    rsTemp_Objective.Open
    'Temp Materi
    Set Rs_Materi_Temp = Nothing
    Set Rs_Materi_Temp = New ADODB.Recordset

    With Rs_Materi_Temp.Fields
        .Append "Objective_Id", adBigInt, , adFldMayBeNull
        .Append "Client_Brief_Id", adVarChar, 13, adFldIsNullable
        .Append "IB_Id", adVarChar, 13, adFldIsNullable
        .Append "Material_Name", adVarChar, 50, adFldMayBeNull
        .Append "Material_Duration", adVarChar, 5, adFldMayBeNull
        
    End With
    Rs_Materi_Temp.Open

    ' Temp Campaign Outline
    Set Rs_Temp_Camp_Outline = New ADODB.Recordset
    
    With Rs_Temp_Camp_Outline
        .Fields.Append "Campaign_Id", adBigInt, , adFldIsNullable
        .Fields.Append "Client_Brief_ID", adChar, 10, adFldIsNullable
        .Fields.Append "IB_Id", adChar, 13, adFldIsNullable
        .Fields.Append "Tarps_Per_Week", adInteger, , adFldIsNullable
        .Fields.Append "Objective_Id", adBigInt, , adFldIsNullable
        .Fields.Append "Month_Week_commencing", adDate, , adFldIsNullable
        .Fields.Append "Week_Commencing", adDate, , adFldIsNullable
        .Fields.Append "Material_Name", adVarChar, 50, adFldMayBeNull
        .Fields.Append "Material_Duration", adDecimal, 9, adFldMayBeNull
        
    End With
    Rs_Temp_Camp_Outline.Open
End Sub
Public Sub Deinitialize_Temp_Table()
'*******************************************
' Procedure     : Deinitialize_Temp_Table
' Function      : Hapus Temporary Table
' Last Update   : 18/03/2016
'*****************************************
    CloseRecordset rsTemp_Objective
    CloseRecordset Rs_Materi_Temp
    CloseRecordset Rs_Temp_Camp_Outline
End Sub

Private Sub Form_Load()
'*******************************************
' Procedure     : Form_Load
' Function      : Form Load Event Handler
' Last Update   : 18/03/2016
'*****************************************
Dim str_month As String
    Call EnableObject(False)
    Edit_WC_Flag = False
    First_Time = True
    recDate.Requery
    toDay = recDate(0).Value
    Cbo_Month.AddItem "January"
    Cbo_Month.AddItem "February"
    Cbo_Month.AddItem "March"
    Cbo_Month.AddItem "April"
    Cbo_Month.AddItem "May"
    Cbo_Month.AddItem "June"
    Cbo_Month.AddItem "July"
    Cbo_Month.AddItem "August"
    Cbo_Month.AddItem "September"
    Cbo_Month.AddItem "October"
    Cbo_Month.AddItem "November"
    Cbo_Month.AddItem "December"
    str_month = Get_Month_Name(month(toDay))
    LoadYear Cbo_Year
    Cbo_Year.Text = Year(toDay)
    Cbo_Month.Text = str_month
    Index_Row = 0
    Initilize_Header_Flex_Grid
    loadCboBrand Cbo_Brand, " AND brand_code IN (SELECT brand_code FROM Media_Security_Catalog WHERE User_name='" & UserName & "' AND position IN ('Planner','Implementor') and Valid_until > getdate()) "
    If Cbo_Brand.ListCount <> 0 Then
        Cbo_Brand.Text = Cbo_Brand.List(0)
    End If
    loadCboCluster Cbo_Cluster, ""
    Media_Type_Code = getMediaTypeCode(" AND UPPER(Media_Type_Name)='" & "TELEVISION MEDIA INDUK'")
    load_IB_TV
     '  Cbo_Brand.SetFocus
    First_Time = False
End Sub

Private Sub load_IB_TV()
'*******************************************
' Procedure     : load_IB_TV
' Function      : Loading data IB TV ke dalam Recordset IB_TV
' Last Update   : 18/03/2016
'*****************************************
    If rs_IB_TV.State = 1 Then rs_IB_TV.Close
    rs_IB_TV.Open "Select * From IB_TV  " _
                & "WHERE  Left(IB_ID,4)='" & Left(Cbo_Brand.Text, 4) & "'" _
                & " AND Brand_Variant_Code ='" & Left(Cbo_Brand_Variant.Text, 5) & "' " _
                & "AND Year = " & Cbo_Year.Text _
                   , ConnERP, adOpenStatic, adLockPessimistic, adCmdText
    Clear_Form
    If Not rs_IB_TV.EOF And Not rs_IB_TV.BOF Then
        LoadToForm
    End If
End Sub
Private Sub Clear_Form()
'*************************************************
' Procedure         : Clear_Form
' Function          : To Claer Form
' Date              : 01/09/2001
' Parameter Input   :
' Parameter Output  :
' Last Update/By    :
'************************************************
    txt_IB_Id.Text = ""
    DT_Date.Value = Date
    Cbo_Cluster.Text = ""
    Txt_Enterd_By.Text = ""
    txt_Primary.Text = ""
    Txt_Plan_No.Text = ""
    
    Lbl_Approval_Status.Caption = ""
    Lbl_Approval_Date.Caption = ""
          lblLastModifiedDate.Caption = lblLastModifiedDate.Tag & ""
    lblLastModifiedBy.Caption = lblLastModifiedBy.Tag & ""

    'Television Objective
    Flex_T_O.Rows = 2
    For Index_Col = 0 To 7
        Flex_T_O.Row = 1
        Flex_T_O.col = Index_Col
        Flex_T_O.TextMatrix(1, Index_Col) = ""
        Flex_T_O.CellBackColor = vbWhite
    Next Index_Col
    
    'Campaign Outline
    Flex_W_C.Rows = 2
    For Index_Col = 1 To Flex_W_C.cols - 1
        Flex_W_C.TextMatrix(0, Index_Col) = ""
        Flex_W_C.TextMatrix(1, Index_Col) = ""
    Next Index_Col
    
    'Montly Budget
    Lbl_Month_1.Caption = ""
    Lbl_Month_2.Caption = ""
    Lbl_Month_3.Caption = ""
    
    Txt_Month_1.Text = ""
    Txt_Month_2.Text = ""
    Txt_Month_3.Text = ""
    
    'Consideration
    Txt_Other_Consideration.Text = ""
    Txt_Attachments.Text = ""
    Txt_Prog_Consideration.Text = ""
    
    Lbl_Status.Caption = ""
    Lbl_Status.ToolTipText = ""
End Sub

Public Sub LoadToForm()
'*******************************************
' Procedure     : LoadToForm
' Function      : Menampilkan data IB TV ke dalam From
' Last Update   : 18/03/2016
'*****************************************
    Dim Index_Array As Integer
    Dim intCol As Integer
    Dim DateMWC As Date
    Dim DateWC As Date
    Dim usrStamp As String
    Dim strUstamp() As String
    
    Me.MousePointer = vbHourglass
    Disable_Form
    Clear_Form
    mNew = False
    
    'Initialize Flex_Grid
    If rs_IB_TV.BOF Or rs_IB_TV.EOF Then Exit Sub
    
    Initilize_Flex_Grid rs_IB_TV.Fields("Month").Value, rs_IB_TV.Fields("Year").Value
    'Clear Texr
    Me.Txt_Month_1.Text = ""
    Me.Txt_Month_2.Text = ""
    Me.Txt_Month_3.Text = ""
    
    Actual_Year = rs_IB_TV.Fields("Year").Value
    Actual_Month = rs_IB_TV.Fields("Month").Value
    
    cbo_Client_Brief_Id.Clear
    cbo_Client_Brief_Id.AddItem rs_IB_TV.Fields("Client_Brief_Id").Value
    cbo_Client_Brief_Id.Text = rs_IB_TV.Fields("Client_Brief_Id").Value
    
    txt_IB_Id.Text = rs_IB_TV.Fields("IB_ID").Value
    
    DT_Date.Value = rs_IB_TV.Fields("Date_Entered").Value
    Cbo_Month.Text = Get_Month_Name(rs_IB_TV.Fields("Month").Value)
    
   
    Txt_Enterd_By.Text = rs_IB_TV.Fields("Entered_By").Value
    txt_Primary.Text = rs_IB_TV.Fields("Target_Primary").Value
    
    'plan no
    Txt_Plan_No.Text = IIf(IsNull(rs_IB_TV.Fields("Plan_No").Value), "", rs_IB_TV.Fields("Plan_No").Value)
    Cbo_Cluster.Text = rs_IB_TV.Fields("Target_Secondary").Value
    Cbo_Cluster.ListIndex = IIf(IsNull(rs_IB_TV.Fields("Cluster_code").Value), 1, rs_IB_TV.Fields("Cluster_code").Value) - 1

    Txt_Attachments.Text = rs_IB_TV.Fields("Attachment").Value
    Txt_Prog_Consideration.Text = rs_IB_TV.Fields("Consideration").Value
    Txt_Other_Consideration.Text = rs_IB_TV.Fields("Other_Consideration").Value
    usrStamp = getUserStamp("IB_TV", "WHERE IB_ID='" & txt_IB_Id.Text & "'")
    If usrStamp <> "" Then
        strUstamp = Split(usrStamp, "|")
        
      lblLastModifiedDate.Caption = lblLastModifiedDate.Tag & strUstamp(1) & " |"
    lblLastModifiedBy.Caption = lblLastModifiedBy.Tag & strUstamp(0) & " |"
    End If

    'Approval Status
    If rs_IB_TV.Fields("Client_Approval").Value = 1 Then
        Lbl_Approval_Status.ForeColor = vbBlack
        Lbl_Approval_Date.ForeColor = vbBlack
        Lbl_Approval_Status.Caption = "Approved"
        Lbl_Approval_Date.Caption = rs_IB_TV.Fields("Approval_Date").Value
        Fra_Approval.ToolTipText = "Approved"
        Lbl_Approval_Status.ToolTipText = "Approved"
        
        'cek apakah objective nya cancel semua
        Dim rsCekObj As New ADODB.Recordset
        Dim dtaCekObj As New cls_data
        rsCekObj.Open " SELECT  *  FROM IB_TV_Objective  WHERE Status = 1 AND IB_Id='" & txt_IB_Id.Text & "'", ConnERP, adOpenStatic, adLockPessimistic
        
        
    Else
        Lbl_Approval_Status.ForeColor = vbRed
        Lbl_Approval_Date.ForeColor = vbRed
        Lbl_Approval_Status.Caption = "UnApproved"
        Lbl_Approval_Date.Caption = ""
        Fra_Approval.ToolTipText = "Double Click to Approve"
        Lbl_Approval_Status.ToolTipText = "Double Click to Approve"
      '  Cmd_Cancel_IB.Enabled = False
    End If
    
    'Status
    If rs_IB_TV.Fields("Status").Value = 1 Then
        Lbl_Status.Caption = ""
        Lbl_Status.ToolTipText = ""
    Else
        Lbl_Status.Caption = "Canceled"
        Lbl_Status.ToolTipText = rs_IB_TV.Fields("Cancel_By").Value & " (" & rs_IB_TV.Fields("Cancel_Date").Value & ")"
    End If
    
    '*********** Television Objective *******************
    Flex_T_O.Rows = 1
    Rs_Television_Obj.Open " SELECT  *  FROM IB_TV_Objective  WHERE IB_ID='" & txt_IB_Id.Text & "' ORDER BY Objective_Id", ConnERP, adOpenStatic, adLockPessimistic
    
    Do While Not Rs_Television_Obj.EOF And Not Rs_Television_Obj.BOF

        'Add Row
        Flex_T_O.Rows = Flex_T_O.Rows + 1
        Flex_T_O.FixedRows = 1
        For Index_Array = 1 To 7
           Frm_New_IB_TV.Flex_T_O.Row = Frm_New_IB_TV.Flex_T_O.Rows - 1
           Frm_New_IB_TV.Flex_T_O.col = Index_Array
           If Rs_Television_Obj.Fields("Status") = 1 Then
                Frm_New_IB_TV.Flex_T_O.CellBackColor = vbWhite
                Frm_New_IB_TV.Flex_T_O.CellForeColor = vbBlack
                Else
                    Frm_New_IB_TV.Flex_T_O.CellBackColor = vbRed
                    Frm_New_IB_TV.Flex_T_O.CellForeColor = vbWhite
           End If
        Next Index_Array

        Flex_T_O.TextMatrix(Flex_T_O.Rows - 1, 0) = IIf(IsNull(Rs_Television_Obj.Fields("Objective_Id").Value), "", Rs_Television_Obj.Fields("Objective_Id").Value)
        Flex_T_O.TextMatrix(Flex_T_O.Rows - 1, 1) = Format(IIf(IsNull(Rs_Television_Obj.Fields("W_C_From").Value), "1/1/1900", Rs_Television_Obj.Fields("W_C_From").Value), "MMM/dd/yyyy") & " - " & Format(IIf(IsNull(Rs_Television_Obj.Fields("W_C_To").Value), "1/1/1900", Rs_Television_Obj.Fields("W_C_To").Value), "MMM/dd/yyyy") 'Week Commencing
        Flex_T_O.TextMatrix(Flex_T_O.Rows - 1, 2) = IIf(IsNull(Rs_Television_Obj.Fields("Campaign_Type").Value), "", Rs_Television_Obj.Fields("Campaign_Type").Value) 'IIf(IsNull(Rs_Television_Obj.Fields("Campaign_Type_Name").Value), "", Rs_Television_Obj.Fields("Campaign_Type_Name").Value) 'Campaign Type
        Flex_T_O.TextMatrix(Flex_T_O.Rows - 1, 3) = IIf(IsNull(Rs_Television_Obj.Fields("Frequency").Value), "", Rs_Television_Obj.Fields("Frequency").Value) 'IIf(IsNull(Rs_Television_Obj.Fields("Frequency_Name").Value), "", Rs_Television_Obj.Fields("Frequency_Name").Value) 'Frequency
        Flex_T_O.TextMatrix(Flex_T_O.Rows - 1, 4) = Rs_Television_Obj.Fields("Reach").Value 'Reach
        Flex_T_O.TextMatrix(Flex_T_O.Rows - 1, 5) = Rs_Television_Obj.Fields("Tarps").Value 'Tarps
        Flex_T_O.TextMatrix(Flex_T_O.Rows - 1, 6) = Format(IIf(IsNull(Rs_Television_Obj.Fields("Nett").Value), 0, Rs_Television_Obj.Fields("Nett").Value), "##,##0")   'Nett
        Flex_T_O.TextMatrix(Flex_T_O.Rows - 1, 7) = Format(IIf(IsNull(Rs_Television_Obj.Fields("MSC").Value), 0, Rs_Television_Obj.Fields("MSC").Value), "##,##0")  'MSC
'        '************ Material Mix *****************'
'

        Rs_Television_Obj.MoveNext
    Loop

    CloseRecordset Rs_Television_Obj

    '*************** IB TV Material ********************
    'Open Recodset Material
    strSql = " SELECT  *  FROM IB_TV_Objective_Material  WHERE Objective_Id IN (SELECT Objective_Id FROM IB_TV_Objective WHERE "
    strSql = strSql & " IB_ID ='" & txt_IB_Id.Text & "') ORDER BY Objective_Id"
    Rs_TV_Material.Open strSql, ConnERP, adOpenStatic, adLockPessimistic
    
    Flex_W_C.Rows = 2
    Do While Not Rs_TV_Material.BOF And Not Rs_TV_Material.EOF
        Flex_W_C.Rows = Flex_W_C.Rows + 1
        Index_Row = Flex_W_C.Rows - 1
        Flex_W_C.TextMatrix(Index_Row, 0) = Rs_TV_Material.Fields("Objective_Id").Value & ":" & Rs_TV_Material.Fields("Material_Name").Value & ":" & Rs_TV_Material.Fields("Material_Duration").Value

            'Campaign Outline
            '=========================================
        strSql = " SELECT * FROM IB_TV_Campaign WHERE Objective_Id =" & Rs_TV_Material.Fields("Objective_Id").Value
        strSql = strSql & " AND Material_Name = '" & Clear_String(Rs_TV_Material.Fields("Material_Name").Value) & "'"
        strSql = strSql & " AND Material_Duration = " & Rs_TV_Material.Fields("Material_Duration").Value & " ORDER BY Objective_Id, Campaign_Id"
        Rs_Camp_Outline.Open strSql, ConnERP, adOpenStatic, adLockPessimistic
            
        With Rs_Camp_Outline
            Do While Not .BOF And Not .EOF
                    'Show To Grid
                For intCol = 1 To Flex_W_C.cols - 1
                    If Flex_W_C.TextMatrix(1, intCol) <> "" Then
                        DateMWC = Get_Month_Number(Flex_W_C.TextMatrix(0, intCol)) & "/" & "01" & "/" & rs_IB_TV.Fields("Year").Value
                        DateWC = Get_Actual_Month(CInt(Flex_W_C.TextMatrix(1, intCol)), CInt(Flex_W_C.TextMatrix(1, intCol)), Get_Month_Number(Flex_W_C.TextMatrix(0, intCol)), rs_IB_TV.Fields("Year").Value) & "/" & Flex_W_C.TextMatrix(1, intCol) & "/" & Get_Actual_Year(CInt(Flex_W_C.TextMatrix(1, intCol)), CInt(Flex_W_C.TextMatrix(1, intCol)), Get_Month_Number(Flex_W_C.TextMatrix(0, intCol)), rs_IB_TV.Fields("Year").Value)
                            
                        If DateMWC = Rs_Camp_Outline.Fields("Month_Week_commencing").Value And _
                            DateWC = Rs_Camp_Outline.Fields("Week_Commencing").Value Then
                                Flex_W_C.TextMatrix(Index_Row, intCol) = Rs_Camp_Outline.Fields("Tarps_Per_Week").Value
                        End If
                            
                    End If
                Next
                
                .MoveNext
            Loop

            .Close
        End With
        Rs_TV_Material.MoveNext
    Loop
    CloseRecordset Rs_TV_Material

'

    '************ Montly Budget ******************
    'Montly Budget
    Lbl_Month_1.Caption = ""
    Lbl_Month_2.Caption = ""
    Lbl_Month_3.Caption = ""
    Rs_Montly_Budget.Open "Select Month,budget From IB_TV_Montly_Budget WHERE Client_Brief_ID='" & cbo_Client_Brief_Id.Text & "' and IB_ID='" & txt_IB_Id.Text & "'", ConnERP, adOpenKeyset, adLockReadOnly, adCmdText
    
    If Not Rs_Montly_Budget.BOF And Not Rs_Montly_Budget.EOF Then
        Rs_Montly_Budget.MoveFirst
        Lbl_Month_1.Caption = Get_Month_Name(Rs_Montly_Budget.Fields("Month").Value)
        Txt_Month_1.Text = Format(Rs_Montly_Budget.Fields("budget").Value, "##,##0")
        Rs_Montly_Budget.MoveNext
        If Not Rs_Montly_Budget.EOF Then
            Txt_Month_2.Text = Format(Rs_Montly_Budget.Fields("Budget").Value, "##,##0")
            Lbl_Month_2.Caption = Get_Month_Name(Rs_Montly_Budget.Fields("Month").Value)
        Else
            GoTo Close_Recodset
        End If
        Rs_Montly_Budget.MoveNext
        If Not Rs_Montly_Budget.EOF Then
            Txt_Month_3.Text = Format(Rs_Montly_Budget.Fields("budget").Value, "##,##0")
            Lbl_Month_3.Caption = Get_Month_Name(Rs_Montly_Budget.Fields("Month").Value)
        Else
            GoTo Close_Recodset
        End If
    End If
   
Close_Recodset:
    CloseRecordset Rs_Montly_Budget
    Me.MousePointer = 0
    
End Sub
Public Sub Initilize_Flex_Grid(Set_Month As Integer, Set_Year As Integer)
'*******************************************
' Procedure     : Initilize_Flex_Grid
' Function      : Initilize Flex_Grid
' Last Update   : 18/03/2016
'*****************************************
    If Get_WeekCommencing(Set_Month, Set_Year) = "" Then
        MsgBox "Week Commencing Not Found, Please Contact IT Department !", vbCritical, APPLICATION_NAME
        Exit Sub
    End If

    If Set_Month = 12 Then
        Max_Col = Len(Get_WeekCommencing(Set_Month, Set_Year)) / 2
        Flex_W_C.cols = Max_Col + 1
        Bulan1 = Set_Month
      
        Week_Count_Month1 = Len(Get_WeekCommencing(Bulan1, Set_Year))
        Week_Count_Month1 = Week_Count_Month1 / 2
    
    
    ElseIf Set_Month = 11 Then
    'Max_Col = 8
        Max_Col = Len(Get_WeekCommencing(Set_Month, Set_Year)) / 2
        Max_Col = Max_Col + Len(Get_WeekCommencing(Set_Month + 1, Set_Year)) / 2
    
        Flex_W_C.cols = Max_Col + 1
        Bulan1 = Set_Month
        Bulan2 = Set_Month + 1
    
    
        Week_Count_Month1 = Len(Get_WeekCommencing(Bulan1, Set_Year))
        Week_Count_Month1 = Week_Count_Month1 / 2
        Week_Count_Month2 = Len(Get_WeekCommencing(Bulan2, Set_Year))
        Week_Count_Month2 = Week_Count_Month2 / 2
       
    Else
        Max_Col = 14
    'Max_Col = 13
    
        Flex_W_C.cols = Max_Col + 1
    
        Bulan1 = Set_Month
        Bulan2 = Set_Month + 1
        Bulan3 = Set_Month + 2
    
        Week_Count_Month1 = Len(Get_WeekCommencing(Bulan1, Set_Year))
        Week_Count_Month1 = Week_Count_Month1 / 2
        Week_Count_Month2 = Len(Get_WeekCommencing(Bulan2, Set_Year))
        Week_Count_Month2 = Week_Count_Month2 / 2
        Week_Count_Month3 = Len(Get_WeekCommencing(Bulan3, Set_Year))
        Week_Count_Month3 = Week_Count_Month3 / 2

    End If

    Flex_W_C.MergeRow(0) = False

    With Flex_W_C

        For Index_Col = 1 To Max_Col
           
            .ColWidth(Index_Col) = 550
            .Row = 0
            .col = Index_Col
            
            If Set_Month = 11 Then
                If Index_Col <= Week_Count_Month1 Then
                     .Text = Get_Month_Name(Bulan1)
                    Else
                        .Text = Get_Month_Name(Bulan2)
                End If
                
            ElseIf Set_Month = 12 Then
                If Index_Col <= Week_Count_Month1 Then
                     .Text = Get_Month_Name(Bulan1)
                End If
                
            Else
                If Index_Col <= Week_Count_Month1 Then
                    .Text = Get_Month_Name(Bulan1)
                ElseIf Index_Col <= (Week_Count_Month1 + Week_Count_Month2) Then
                     .Text = Get_Month_Name(Bulan2)
                ElseIf Index_Col > (Week_Count_Month1 + Week_Count_Month2) Then
                     .Text = Get_Month_Name(Bulan3)
                End If
                
            End If
        Next Index_Col
    
    
    .MergeCells = flexMergeRestrictColumns
         .MergeRow(0) = True
         If Set_Month = 11 Then
             .ColAlignment(1) = 3
             .ColAlignment(Week_Count_Month1 + 1) = 3
         ElseIf Set_Month = 12 Then
             .ColAlignment(1) = 3
         Else
             .ColAlignment(1) = 3
             .ColAlignment(Week_Count_Month1 + 1) = 3
             .ColAlignment(Week_Count_Month1 + Week_Count_Month2 + 1) = 3
        End If
        'Week Commencing
        '+++++++++++++++++++++++++++++++++++++++++++++++++++++
         Dim str_week As String
         Dim Index_coll As Integer
        
         Index_coll = 1
         'Month 1
        
        .RowHeight(1) = 325
         
        If Set_Month = 11 Then
        
             L1.Visible = True
             Txt_Month_1.Visible = True
             Lbl_Month_1.Visible = True
             L2.Visible = True
             Txt_Month_2.Visible = True
             Lbl_Month_2.Visible = True
             L3.Visible = False
             Txt_Month_3.Visible = False
             Lbl_Month_3.Visible = False
                     
             str_week = Get_WeekCommencing(Bulan1, Set_Year)
             Do While str_week <> ""
                 .Row = 1
                 .col = Index_coll
                 .Text = Mid(str_week, 1, 2)
                 str_week = Right(str_week, Len(str_week) - 2)
                 Index_coll = Index_coll + 1
             Loop
             ' Month 2
             
             str_week = Get_WeekCommencing(Bulan2, Set_Year)
             Do While str_week <> ""
                 .Row = 1
                 .col = Index_coll
                 .Text = Mid(str_week, 1, 2)
                 str_week = Right(str_week, Len(str_week) - 2)
                 Index_coll = Index_coll + 1
             Loop
        ElseIf Set_Month = 12 Then
             'Montly Budget
             L1.Visible = True
             Txt_Month_1.Visible = True
             Lbl_Month_1.Visible = True
             L2.Visible = False
             L3.Visible = False
             Txt_Month_2.Visible = False
             Txt_Month_3.Visible = False
             Lbl_Month_2.Visible = False
             Lbl_Month_3.Visible = False
                     
             str_week = Get_WeekCommencing(Bulan1, Set_Year)
             Do While str_week <> ""
                 .Row = 1
                 .col = Index_coll
                 .Text = Mid(str_week, 1, 2)
                 str_week = Right(str_week, Len(str_week) - 2)
                 Index_coll = Index_coll + 1
             Loop
        Else
             'Montly Budget
             L1.Visible = True
             Txt_Month_1.Visible = True
             Lbl_Month_1.Visible = True
             L2.Visible = True
             L3.Visible = True
             Txt_Month_2.Visible = True
             Txt_Month_3.Visible = True
             Lbl_Month_2.Visible = True
             Lbl_Month_3.Visible = True
        
             str_week = Get_WeekCommencing(Bulan1, Set_Year)
             Do While str_week <> ""
                 .Row = 1
                 .col = Index_coll
                 .Text = Mid(str_week, 1, 2)
                 str_week = Right(str_week, Len(str_week) - 2)
                 Index_coll = Index_coll + 1
             Loop
             ' Month 2
             
             str_week = Get_WeekCommencing(Bulan2, Set_Year)
             Do While str_week <> ""
                 .Row = 1
                 .col = Index_coll
                 .Text = Mid(str_week, 1, 2)
                 str_week = Right(str_week, Len(str_week) - 2)
                 Index_coll = Index_coll + 1
             Loop
              
              ' Month 3
              
             str_week = Get_WeekCommencing(Bulan3, Set_Year)
             Do While str_week <> ""
                 .Row = 1
                 .col = Index_coll
                 .Text = Mid(str_week, 1, 2)
                 str_week = Right(str_week, Len(str_week) - 2)
                 Index_coll = Index_coll + 1
             Loop
              '++++++++++++++++++++++++++++++++++++++++++++++++++++
         End If
    End With
    
    Lbl_Month_1.Caption = Get_Month_Name(Bulan1)
    Lbl_Month_2.Caption = Get_Month_Name(Bulan2)
    Lbl_Month_3.Caption = Get_Month_Name(Bulan3)



End Sub
Public Sub Initilize_Header_Flex_Grid()
'*******************************************
' Procedure     : Initilize_Header_Flex_Grid
' Function      : Initilize Flex_Grid Header
' Last Update   : 18/03/2016
'*****************************************
    With Flex_T_O
        .cols = 8
        .col = 0
        .Row = 0
        .Text = "ID"
        
        .col = 1
        .Row = 0
        .Text = "Week Commencing"
        .col = 2
        .Row = 0
        .Text = "Campaign Type"
        .col = 3
        .Row = 0
        .Text = "Frequency"
        .col = 4
        .Row = 0
        .Text = "Reach %"
        .col = 5
        .Row = 0
        .Text = "Tarps"
        
        .col = 6
        .Row = 0
        .Text = "Nett"
        
        .col = 7
        .Row = 0
        .Text = "MSC"
        .ColWidth(0) = 500
        .ColWidth(1) = 2600
        .ColWidth(2) = 2200
        .ColWidth(5) = 600
        .ColWidth(6) = 1500
        .ColWidth(7) = 1500
    End With
    
    With Flex_W_C
        .col = 0
        .Row = 0
        .ColWidth(0) = 1800
        .Text = "Month"
        .col = 0
        .Row = 1
        .Text = "Week Commencing"
    End With
    
End Sub

Private Sub Form_Resize()
'*******************************************
' Procedure     : Form_Resize
' Function      : Form Resize Event Handler
' Last Update   : 18/03/2016
'*****************************************
 Dim nWidth As Single, nHeight As Single

    On Local Error Resume Next
    nWidth = Me.ScaleWidth: nHeight = Me.ScaleHeight

    With pnlMain
        .Move .Left, .Top, nWidth - (.Left * 2), nHeight - .Top - picStatusBar.Height - 20
    End With
     With pnlCopy_IB
     .Width = nWidth - 40
     .Height = nHeight - 40
     .Left = 20
     .Top = 20
     
     
    End With
    fraCopy_IB.Left = (pnlCopy_IB.Width - fraCopy_IB.Width) / 2
    fraCopy_IB.Top = (pnlCopy_IB.Height - fraCopy_IB.Height) / 2
    Frame4.Width = pnlMain.Width - Frame4.Left - 20
    
    Frame2.Width = pnlMain.Width - (Frame2.Left * 2)
    Frame3.Width = Frame2.Width
    Flex_T_O.Width = Frame2.Width - Flex_T_O.Left - 1200
    Cmd_Add_T_O.Width = Frame2.Width - (Flex_T_O.Left + Flex_T_O.Width) - 120
    Cmd_Add_T_O.Left = Flex_T_O.Left + Flex_T_O.Width + 40
    Cmd_Edit_T_O.Width = Cmd_Add_T_O.Width
    Cmd_Edit_T_O.Left = Cmd_Add_T_O.Left
    Cmd_Delete_T_O.Width = Cmd_Add_T_O.Width
    Cmd_Delete_T_O.Left = Cmd_Add_T_O.Left
    Flex_W_C.Width = Frame3.Width - Flex_W_C.Left - 1200
    Cmd_Edit_C_O.Width = Frame3.Width - (Flex_W_C.Left + Flex_W_C.Width) - 120
    Cmd_Edit_C_O.Left = Flex_W_C.Left + Flex_W_C.Width + 40
    Cmd_Save_C_O.Width = Cmd_Edit_C_O.Width
    Cmd_Save_C_O.Left = Cmd_Edit_C_O.Left
    Cmd_Cancel_C_O.Width = Cmd_Edit_C_O.Width
    Cmd_Cancel_C_O.Left = Cmd_Edit_C_O.Left
    Cmd_Delete_C_O.Width = Cmd_Edit_C_O.Width
    Cmd_Delete_C_O.Left = Cmd_Edit_C_O.Left
    SSTab1.Height = pnlMain.Height - SSTab1.Top - 60
    SSTab1.Width = ((pnlMain.Width - SSTab1.Left) / 3) * 2 - 500
    Frame11.Left = SSTab1.Left + SSTab1.Width + 60
    Lbl_Status.Left = Frame11.Left
    
    Fra_Approval.Left = Frame11.Left + Frame11.Width + 60
    Fra_Approval.Width = pnlMain.Width - Fra_Approval.Left - 60
    Fra_Approval.Height = Frame11.Height
    Lbl_Status.Width = Frame11.Width + Fra_Approval.Width + 60
    Lbl_Status.Height = SSTab1.Height - Frame11.Height
    
    Txt_Prog_Consideration.Height = SSTab1.Height - (Txt_Prog_Consideration.Top + 80)
    Txt_Other_Consideration.Height = SSTab1.Height - (Txt_Other_Consideration.Top + 80)
    Txt_Attachments.Height = SSTab1.Height - (Txt_Attachments.Top + 80)
    Txt_Prog_Consideration.Width = SSTab1.Width - 360
    Txt_Other_Consideration.Width = SSTab1.Width - 360
    Txt_Attachments.Width = SSTab1.Width - 360
    Lbl_Approval_Status.Left = 160
    Lbl_Approval_Status.Width = Fra_Approval.Width - 320
    Lbl_Approval_Date.Width = Fra_Approval.Width - 320
    Lbl_Approval_Date.Left = Lbl_Approval_Status.Left
    txt_Primary.Width = Frame4.Width - txt_Primary.Left - 120
    Cbo_Cluster.Width = Frame4.Width - txt_Primary.Left - 120
    Txt_Plan_No.Width = Frame4.Width - txt_Primary.Left - 120
    Frame6.Width = Frame2.Width
    Cbo_Brand.SetFocus
        On Local Error GoTo 0
End Sub

Private Sub Lst_IB_TV_Click()
'*******************************************
' Procedure     : Lst_IB_TV_Click
' Function      : Lst_IB_TV Click Event Handler
' Last Update   : 18/03/2016
'*****************************************
    Dim IsError As Boolean
           Dim Count_Select As Integer
           Dim idx As Integer
           Cmd_Ok.Enabled = False
           IsError = False
           Ib_Copy_1 = ""
           Ib_Copy_2 = ""
           Ib_Copy_3 = ""
            
           Month_Copy_1 = 0
           Month_Copy_2 = 0
           Month_Copy_3 = 0
           
           For idx = 0 To Me.Lst_IB_TV.ListCount - 1
                If Me.Lst_IB_TV.Selected(idx) Then
                    'hitung list yang terpilih
                    Count_Select = Count_Select + 1
                    If Count_Select <= 3 Then
                        Select Case Count_Select
                        Case 1: Ib_Copy_1 = Left(Lst_IB_TV.List(idx), 13)
                                Month_Copy_1 = Get_Month_Number(Trim(Mid(Lst_IB_TV.List(idx), InStr(1, Lst_IB_TV.List(idx), ">") + 1, Len(Lst_IB_TV.List(idx)) - (InStr(1, Lst_IB_TV.List(idx), ">")))))
                        Case 2: Ib_Copy_2 = Left(Lst_IB_TV.List(idx), 13)
                                Month_Copy_2 = Get_Month_Number(Trim(Mid(Lst_IB_TV.List(idx), InStr(1, Lst_IB_TV.List(idx), ">") + 1, Len(Lst_IB_TV.List(idx)) - (InStr(1, Lst_IB_TV.List(idx), ">")))))
                                'cocokkan
                                If Month_Copy_1 = Month_Copy_2 Then
                                    IsError = True
                                    Exit For
                                End If
                        Case 3: Ib_Copy_3 = Left(Lst_IB_TV.List(idx), 13)
                                Month_Copy_3 = Get_Month_Number(Trim(Mid(Lst_IB_TV.List(idx), InStr(1, Lst_IB_TV.List(idx), ">") + 1, Len(Lst_IB_TV.List(idx)) - (InStr(1, Lst_IB_TV.List(idx), ">")))))
                                If Month_Copy_1 = Month_Copy_3 Or Month_Copy_2 = Month_Copy_3 Then
                                    IsError = True
                                    Exit For
                                End If
                        End Select
                        
                    End If
                End If
           Next
           
           If IsError Then
                MsgBox "Month is Double", vbCritical, APPLICATION_NAME
                
                Exit Sub
           End If
           If Count_Select > 3 Then
                MsgBox "You Can't select more than 3 Item", vbCritical, APPLICATION_NAME
                Exit Sub
           End If
           If Count_Select = 0 Then
                MsgBox "No IB ID Selectted", vbCritical, APPLICATION_NAME
                Exit Sub
           End If
           Cmd_Ok.Enabled = True
End Sub

Private Sub Opt_Copy_Click()
'*******************************************
' Procedure     : Opt_Copy_Click
' Function      : Opt_Copy Click Event Handler
' Last Update   : 18/03/2016
'*****************************************
    Dim RS_data As New ADODB.Recordset
    
    Me.Lst_IB_TV.Clear
    Lst_IB_TV.Enabled = True
    Cmd_Ok.Enabled = False
     
    strSql = "SELECT  *  FROM ib_tv_montly_budget WHERE ib_id in (SELECT ib_id FROM ib_tv WHERE Month=" & Get_Month_Number(Cbo_Month.Text) & " AND year = " & Val(Cbo_Year.Text) & " AND left(ib_id,4)='" & Left(Cbo_Brand.Text, 4) & "') ORDER BY MOnth, IB_ID "
    RS_data.Open strSql, ConnERP, adOpenStatic, adLockPessimistic
    While Not RS_data.EOF
        Lst_IB_TV.AddItem RS_data("ib_id") & " -> " & Get_Month_Name(RS_data("month"))
        RS_data.MoveNext
    Wend
    CloseRecordset RS_data
End Sub

Private Sub Opt_New_Click()
'*******************************************
' Procedure     : Opt_New_Click
' Function      : Opt_New Click Event Handler
' Last Update   : 18/03/2016
'*****************************************
    Lst_IB_TV.Clear
    Lst_IB_TV.Enabled = False
    Cmd_Ok.Enabled = True
End Sub

Private Sub picButton_Click(Index As Integer)
'*******************************************
' Procedure     : picButton_Click
' Function      : picButton Click Event Handler
' Last Update   : 18/03/2016
'*****************************************
    Select Case Index
        Case 0 'PREV
            If mNew = True Or Edit_Flag = True Then Exit Sub
            If rs_IB_TV.RecordCount < 1 Then Exit Sub
            
            rs_IB_TV.MoveFirst
            rPost = rs_IB_TV.AbsolutePosition
            LoadToForm
            
        Case 1 'PREV
            If mNew = True Or Edit_Flag = True Then Exit Sub
            If rs_IB_TV.RecordCount < 1 Then Exit Sub
            If rPost = 1 Then
                MsgBox "Already On First Record Position", vbCritical, APPLICATION_NAME
                Exit Sub
            End If
            rs_IB_TV.MovePrevious
            rPost = rs_IB_TV.AbsolutePosition
            LoadToForm
            
        Case 2 'NEXT
            If mNew = True Or Edit_Flag = True Then Exit Sub
            If rs_IB_TV.RecordCount < 1 Then Exit Sub
            If rPost = rs_IB_TV.RecordCount Then
                MsgBox "Already On Last Record Position", vbCritical, APPLICATION_NAME
                Exit Sub
            End If
            rs_IB_TV.MoveNext
            rPost = rs_IB_TV.AbsolutePosition
            LoadToForm
            
        Case 3 'LAST
            If mNew = True Or Edit_Flag = True Then Exit Sub
            If rs_IB_TV.RecordCount < 1 Then Exit Sub
             If rPost = rs_IB_TV.RecordCount Then
                MsgBox "Already On Last Record Position", vbCritical, APPLICATION_NAME
                Exit Sub
            End If
            rs_IB_TV.MoveLast
            rPost = rs_IB_TV.AbsolutePosition
            LoadToForm
            
        Case 4 'ADD
            db_add
        Case 5
            Edit_Mode
        Case 8 'PRINT
            print_IB_TV
        Case 9 'CLOSE
            Unload Me
        Case 10
            db_save
        Case 11 'CANCEL
            If mNew Then
                If cbo_Client_Brief_Id.ListCount <> 0 Then
                    If txt_IB_Id.Text <> "" Then
                        CancelID "IBI", txt_IB_Id.Text, Left(Cbo_Brand.Text, 4), Cbo_Year.Text
                    End If
                End If
            End If
            Deinitialize_Temp_Table
            Clear_Form
            Call EnableObject(False)
            LoadToForm
            Me.MousePointer = 0
        End Select
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
    Cbo_Brand.Enabled = Enable
    Cbo_Brand_Variant.Enabled = Enable
    cbo_Client_Brief_Id.Enabled = Not Enable
    Cbo_Cluster.Enabled = Not Enable
    Cbo_Month.Enabled = Not Enable
    Cbo_Year.Enabled = Enable
    txt_Primary.Enabled = Not Enable
    Txt_Plan_No.Enabled = Not Enable
    Txt_Month_1.Enabled = Not Enable
    Txt_Month_2.Enabled = Not Enable
    Txt_Month_3.Enabled = Not Enable
    Txt_Other_Consideration.Locked = Enable
    Txt_Attachments.Locked = Enable
    Txt_Prog_Consideration.Locked = Enable
    
    Frame2.Enabled = True
    Frame3.Enabled = True
    Cmd_Edit_C_O.Enabled = False
    
    

End Sub
Sub Edit_Mode()
'*******************************************
' Procedure     : Edit_Mode
' Function      : prosedur edit data
' Last Update   : 18/03/2016
'**********************************************
    Dim Answare As Integer
    Dim Index_Record As Integer
    
    If rs_IB_TV.EOF Or rs_IB_TV.BOF Then Exit Sub
    If rs_IB_TV.Fields("Client_Approval") = 1 Then
        MsgBox "Approved Implementation Brief, Cannot be Edited", vbCritical, APPLICATION_NAME
    Else
        Edit_Flag = True
        
        Actual_Month = Get_Month_Number(Cbo_Month.Text)
               
        Button_Normal False
        Cbo_Month.Enabled = False
        cbo_Client_Brief_Id.Enabled = False
        Me.Cmd_Add_T_O.Enabled = True
        Me.Cmd_Delete_T_O.Enabled = True
        Me.Cmd_Edit_T_O.Enabled = True
        Me.Cmd_Cancel_C_O.Enabled = True
        Me.Cmd_Edit_C_O.Enabled = True
        Me.Cmd_Save_C_O.Enabled = True
        
        Initialize_Temp_Table
        Dim myData As New cls_data
        strSql = " SELECT  *  FROM IB_TV_Objective  WHERE IB_ID='" & txt_IB_Id.Text & "' ORDER BY Objective_Id"
        Rs_Television_Obj.Open strSql, ConnERP, adOpenStatic, adLockPessimistic
        
        
        Do While Not Rs_Television_Obj.EOF
            rsTemp_Objective.AddNew
            rsTemp_Objective.Fields("Objective_Id").Value = Rs_Television_Obj.Fields("Objective_Id").Value
            rsTemp_Objective.Fields("Client_Brief_Id").Value = Rs_Television_Obj.Fields("Client_Brief_Id").Value
            rsTemp_Objective.Fields("IB_Id").Value = Rs_Television_Obj.Fields("IB_Id").Value
            rsTemp_Objective.Fields("Week_Commencing").Value = Rs_Television_Obj.Fields("Week_Commencing").Value
            rsTemp_Objective.Fields("Campaign_Type").Value = Rs_Television_Obj.Fields("Campaign_Type").Value
            rsTemp_Objective.Fields("Frequency").Value = Rs_Television_Obj.Fields("Frequency").Value
            rsTemp_Objective.Fields("Reach").Value = Rs_Television_Obj.Fields("Reach").Value
            rsTemp_Objective.Fields("Tarps").Value = Rs_Television_Obj.Fields("Tarps").Value
            rsTemp_Objective.Fields("Budget_With_MSC").Value = Rs_Television_Obj.Fields("Budget_With_MSC").Value
            rsTemp_Objective.Fields("Campaign_Type_Code").Value = Rs_Television_Obj.Fields("Campaign_Type_Code").Value
            rsTemp_Objective.Fields("Frequency_Code").Value = Rs_Television_Obj.Fields("Frequency_Code").Value
            rsTemp_Objective.Fields("W_C_From").Value = Rs_Television_Obj.Fields("W_C_From").Value
            rsTemp_Objective.Fields("W_C_To").Value = Rs_Television_Obj.Fields("W_C_To").Value
            rsTemp_Objective.Fields("Nett").Value = Rs_Television_Obj.Fields("Nett").Value
            rsTemp_Objective.Fields("MSC").Value = Rs_Television_Obj.Fields("MSC").Value
            rsTemp_Objective.Fields("Market_Code").Value = Rs_Television_Obj.Fields("Market_Code").Value
            rsTemp_Objective.Fields("Market_Name").Value = Rs_Television_Obj.Fields("Market_Name").Value
            rsTemp_Objective.Fields("Status").Value = Rs_Television_Obj.Fields("Status").Value
            rsTemp_Objective.Update
            
            
            Rs_Television_Obj.MoveNext
        Loop
        CloseRecordset Rs_Television_Obj
        
        '********************************
        'Material
        '========
        strSql = " SELECT  *  FROM IB_TV_Objective_Material  WHERE IB_ID='" & txt_IB_Id.Text & "' ORDER BY Objective_Id  "
            
        Rs_TV_Material.Open strSql, ConnERP, adOpenStatic, adLockPessimistic
            
        
        Do While Not Rs_TV_Material.EOF
            Rs_Materi_Temp.AddNew
            Rs_Materi_Temp.Fields("Objective_Id").Value = Rs_TV_Material.Fields("Objective_Id").Value
            Rs_Materi_Temp.Fields("Client_Brief_Id").Value = Rs_TV_Material.Fields("Client_Brief_Id").Value
            Rs_Materi_Temp.Fields("IB_Id").Value = Rs_TV_Material.Fields("IB_Id").Value
            Rs_Materi_Temp.Fields("Material_Name").Value = Rs_TV_Material.Fields("Material_Name").Value
            Rs_Materi_Temp.Fields("Material_Duration").Value = Rs_TV_Material.Fields("Material_Duration").Value
                
            Rs_Materi_Temp.Update
            Rs_TV_Material.MoveNext
        Loop
        CloseRecordset Rs_TV_Material
   '========================================

        'Campaign Outline
        
        strSql = " SELECT  *  FROM IB_TV_Campaign  WHERE Objective_Id IN (SELECT Objective_Id FROM IB_TV_Objective WHERE"
        strSql = strSql & " IB_Id='" & txt_IB_Id.Text & "') ORDER BY Campaign_Id"
        Set Rs_Camp_Outline = Nothing
        
        Rs_Camp_Outline.Open strSql, ConnERP, adOpenStatic, adLockPessimistic
        With Rs_Camp_Outline
            Do While Not .EOF
                Rs_Temp_Camp_Outline.AddNew
                Rs_Temp_Camp_Outline.Fields("Campaign_Id").Value = Rs_Camp_Outline.Fields("Campaign_Id").Value
                Rs_Temp_Camp_Outline.Fields("Client_Brief_ID").Value = Rs_Camp_Outline.Fields("Client_Brief_ID").Value
                Rs_Temp_Camp_Outline.Fields("IB_Id").Value = Rs_Camp_Outline.Fields("IB_Id").Value
                
                Rs_Temp_Camp_Outline.Fields("Tarps_Per_Week").Value = Rs_Camp_Outline.Fields("Tarps_Per_Week").Value
                Rs_Temp_Camp_Outline.Fields("Objective_Id").Value = Rs_Camp_Outline.Fields("Objective_Id").Value
                Rs_Temp_Camp_Outline.Fields("Month_Week_commencing").Value = Rs_Camp_Outline.Fields("Month_Week_commencing").Value
                Rs_Temp_Camp_Outline.Fields("Week_Commencing").Value = Rs_Camp_Outline.Fields("Week_Commencing").Value
                Rs_Temp_Camp_Outline.Fields("Material_Name").Value = Rs_Camp_Outline.Fields("Material_Name").Value
                Rs_Temp_Camp_Outline.Fields("Material_Duration").Value = Rs_Camp_Outline.Fields("Material_Duration").Value
                
                
                Rs_Temp_Camp_Outline.Update
                .MoveNext
            Loop
            .Close
        End With
        Call EnableObject(True)
'=================================================
    End If
     
End Sub
Sub Disable_Form()
'*******************************************
' Procedure     : Disable_Form
' Function      : disabling form object
' Last Update   : 18/03/2016
'**********************************************
    Cbo_Month.Enabled = False
    cbo_Client_Brief_Id.Enabled = False
    txt_IB_Id.Enabled = False
    txt_Primary.Enabled = False
    Cbo_Cluster.Enabled = False
    Txt_Plan_No.Enabled = False
    Frame2.Enabled = False
    Frame3.Enabled = False
    Frame6.Enabled = False
    SSTab1.Enabled = False
End Sub
Sub Enable_Form()
'*******************************************
' Procedure     : Enable_Form
' Function      : enabling form object
' Last Update   : 18/03/2016
'**********************************************
    Cbo_Month.Enabled = True
    cbo_Client_Brief_Id.Enabled = True
    txt_Primary.Enabled = True
    Cbo_Cluster.Enabled = True
    Txt_Plan_No.Enabled = True
    Frame2.Enabled = True
    Frame3.Enabled = True
    Frame6.Enabled = True
    SSTab1.Enabled = True
End Sub
Private Sub print_IB_TV()
'*******************************************
' Procedure     : print_IB_TV
' Function      : cetak IB TV
' Last Update   : 18/03/2016
'**********************************************
 Dim Special_Brand As Boolean
    'Declare
    Dim Rs_Temp As New ADODB.Recordset
    
    'Validasi
    Special_Brand = Is_Special_Brand(Left(Cbo_Brand.Text, 4))
    'Temp Recordset
    
    Set Rs_Temp = Nothing
    Rs_Temp.Open "Select getdate()", ConnERP, , , adCmdText
    
    'Logo & Title
    If AppID = "IM" Then
        'Logo
        Rpt_TV_IB.Sections("Section4").Controls("Image_Logo").Visible = True
        'Lbl_Title1 .2340
        Rpt_TV_IB.Sections("Section4").Controls("Lbl_Title1").Left = 2340
        Rpt_TV_IB.Sections("Section4").Controls("Lbl_Title1").Alignment = 2 'Center
        'Lbl_Title2 .2340
        Rpt_TV_IB.Sections("Section4").Controls("Lbl_Title2").Left = 2340
        Rpt_TV_IB.Sections("Section4").Controls("Lbl_Title2").Alignment = 2 'Center
    Else
        Rpt_TV_IB.Sections("Section4").Controls("Image_Logo").Visible = False
        'Lbl_Title1 .2340
        Rpt_TV_IB.Sections("Section4").Controls("Lbl_Title1").Left = 0
        Rpt_TV_IB.Sections("Section4").Controls("Lbl_Title1").Alignment = 0 'Left
        'Lbl_Title2 .2340
        Rpt_TV_IB.Sections("Section4").Controls("Lbl_Title2").Left = 0
        Rpt_TV_IB.Sections("Section4").Controls("Lbl_Title2").Alignment = 0 'Left
    End If
    
    'Header
    '=========
        'Brand
        Rpt_TV_IB.Sections("Section4").Controls("Lbl_Brand").Caption = Get_Brand_Name(Left(Cbo_Brand.Text, 4))
        'Brand Variant
        If IsNull(rs_IB_TV.Fields("Brand_Variant_Name").Value) Then
            Rpt_TV_IB.Sections("Section4").Controls("Lbl_Brand_Variant").Caption = ""
        Else
            Rpt_TV_IB.Sections("Section4").Controls("Lbl_Brand_Variant").Caption = rs_IB_TV.Fields("Brand_Variant_Name").Value
        End If
        'Date
        Rpt_TV_IB.Sections("Section4").Controls("Lbl_Date").Caption = Format(Date, "mmm dd yyyy")
        'IB ID
        Rpt_TV_IB.Sections("Section4").Controls("Lbl_IB_Id").Caption = txt_IB_Id.Text
    'Detail
    '=========
        '---> Target Group
            'Primary Target
            Rpt_TV_IB.Sections("Section4").Controls("Lbl_Primary_Target").Caption = Me.txt_Primary.Text
            'Secontary Taget
            Rpt_TV_IB.Sections("Section4").Controls("Lbl_Secondary_Target").Caption = Cbo_Cluster.Text
            If Special_Brand Then
                'Lable Primary
                Rpt_TV_IB.Sections("Section4").Controls("Lbl_Primary").Caption = "Primary"
                'Lable Secondary
                Rpt_TV_IB.Sections("Section4").Controls("Lbl_Secondary").Caption = "Secondary"
                'Approval Client
                 Rpt_TV_IB.Sections("Section5").Controls("Lbl_Approval_Client").Caption = "Marketing Manager"
            Else
                'Lable Primary
                Rpt_TV_IB.Sections("Section4").Controls("Lbl_Primary").Caption = "Primary"
                'Lable Secondary
                Rpt_TV_IB.Sections("Section4").Controls("Lbl_Secondary").Caption = "Secondary"
                'Approval Client
                 Rpt_TV_IB.Sections("Section5").Controls("Lbl_Approval_Client").Caption = "Marketing"
            End If
            
        '---> Television Objective
        '==================================================
            Dim Rs_TO As New ADODB.Recordset
            Dim Index_TO As Integer
            Dim str_week As String
            
            strSql = " SELECT  *  FROM IB_TV_Objective  WHERE IB_Id='" & txt_IB_Id.Text & "' ORDER BY Objective_Id"
            Rs_TO.Open strSql, ConnERP, adOpenStatic, adLockPessimistic
            'Loop TO
            Index_TO = 0
            Do While Not Rs_TO.EOF
                'Show TO
                    'Week
'                    If Left(Trim(Rs_TO.Fields("Week_Commencing").Value), 6) = Right(Trim(Rs_TO.Fields("Week_Commencing").Value), 6) Then
                    If Rs_TO.Fields("W_C_From").Value = Rs_TO.Fields("W_C_To").Value Then
                        str_week = Format(Rs_TO.Fields("W_C_From").Value, "MMM dd")
                    Else
                        str_week = Format(Rs_TO.Fields("W_C_From").Value, "MMM dd") & " - " & Format(Rs_TO.Fields("W_C_To").Value, "MMM dd")
                    End If
                    
                     Rpt_TV_IB.Sections("Section4").Controls("Lbl_Week_Comm_" & Index_TO).Caption = str_week
                    'Camp Type
                     Rpt_TV_IB.Sections("Section4").Controls("Lbl_Camp_Type_" & Index_TO).Caption = Rs_TO.Fields("Campaign_Type").Value
                    'Freq
                     Rpt_TV_IB.Sections("Section4").Controls("Lbl_Frequency_" & Index_TO).Caption = Rs_TO.Fields("Frequency").Value
                    'Reach
                    Rpt_TV_IB.Sections("Section4").Controls("Lbl_Reach_" & Index_TO).Caption = Rs_TO.Fields("Reach").Value & "%"
                    'Tarps
                    Rpt_TV_IB.Sections("Section4").Controls("Lbl_Tarps_" & Index_TO).Caption = Rs_TO.Fields("Tarps").Value
                    'Budget
                    Rpt_TV_IB.Sections("Section4").Controls("Lbl_Budget_Burst_" & Index_TO).Caption = Format(Rs_TO.Fields("Budget_With_MSC").Value, "##,##0")
                    '========================
                                    
                    'Tampilkan Line
                    Rpt_TV_IB.Sections("Section4").Controls("Shape_WC_" & Index_TO + 1).Visible = True
                    Rpt_TV_IB.Sections("Section4").Controls("Shape_CT_" & Index_TO + 1).Visible = True
                    Rpt_TV_IB.Sections("Section4").Controls("Shape_F_" & Index_TO + 1).Visible = True
                    Rpt_TV_IB.Sections("Section4").Controls("Shape_R_" & Index_TO + 1).Visible = True
                    Rpt_TV_IB.Sections("Section4").Controls("Shape_T_" & Index_TO + 1).Visible = True
                    Rpt_TV_IB.Sections("Section4").Controls("Shape_B_" & Index_TO + 1).Visible = True
                    
                    
                    Rs_TO.MoveNext
                    Index_TO = Index_TO + 1
                    'Batas TO 10
                    If Index_TO > 13 Then
                        Exit Do
                    End If
            'End Loop TO
            Loop
            
            If Index_TO > 7 Then
                Rpt_TV_IB.Sections("Section4").Controls("Shape_W_2").Visible = True
                Rpt_TV_IB.Sections("Section4").Controls("Shape_W_3").Visible = True
                Rpt_TV_IB.Sections("Section4").Controls("Shape_W_4").Visible = True
                Rpt_TV_IB.Sections("Section4").Controls("Shape_W_5").Visible = True
                Rpt_TV_IB.Sections("Section4").Controls("Shape_W_6").Visible = True
                
                Rpt_TV_IB.Sections("Section4").Controls("Lbl_WC_2").Visible = True
                Rpt_TV_IB.Sections("Section4").Controls("Lbl_TOH_1").Visible = True
                Rpt_TV_IB.Sections("Section4").Controls("Lbl_TOH_2").Visible = True
                Rpt_TV_IB.Sections("Section4").Controls("Lbl_TOH_3").Visible = True
                Rpt_TV_IB.Sections("Section4").Controls("Lbl_TOH_4").Visible = True
                Rpt_TV_IB.Sections("Section4").Controls("Lbl_TOH_5").Visible = True
            End If
            
            'Close Rs_To
            Rs_TO.Close
            Set Rs_TO = Nothing
            '----> Pengatutan Posisi Top jika W/C > 5
            
            
        '======================= End Television Objective =========
            
            
        '---> Campaign Outline
        '=================================================
        Dim RS_Materi As New ADODB.Recordset
        Dim Rs_Co As New ADODB.Recordset
        Dim Index_Materi As Integer
        Dim Year_Ib As Integer
        Dim Month_1 As Integer
        Dim Month_2 As Integer
        Dim Month_3 As Integer
        Dim Week_Count_Month1 As Integer
        Dim Week_Count_Month2 As Integer
        Dim Week_Count_Month3 As Integer
        Dim Str_Week_Comm_1 As String
        Dim Str_Week_Comm_2 As String
        Dim Str_Week_Comm_3 As String
        Dim Week_Position As Integer
        
        Dim RsTempCampaign As New ADODB.Recordset
        
        With RsTempCampaign.Fields
            .Append "Objective_id", adVarChar, 100, adFldIsNullable
                        
            .Append "Tarps_1", adInteger, , adFldIsNullable
            .Append "Tarps_2", adInteger, , adFldIsNullable
            .Append "Tarps_3", adInteger, , adFldIsNullable
            .Append "Tarps_4", adInteger, , adFldIsNullable
            .Append "Tarps_5", adInteger, , adFldIsNullable
            
            .Append "Tarps_6", adInteger, , adFldIsNullable
            .Append "Tarps_7", adInteger, , adFldIsNullable
            .Append "Tarps_8", adInteger, , adFldIsNullable
            .Append "Tarps_9", adInteger, , adFldIsNullable
            .Append "Tarps_10", adInteger, , adFldIsNullable
            
            .Append "Tarps_11", adInteger, , adFldIsNullable
            .Append "Tarps_12", adInteger, , adFldIsNullable
            .Append "Tarps_13", adInteger, , adFldIsNullable
            .Append "Tarps_14", adInteger, , adFldIsNullable
            
            
        End With
        
        RsTempCampaign.Open
        
        'Generate Tampilan CO
        '========================================
        
        Year_Ib = rs_IB_TV.Fields("Year").Value
        Month_1 = rs_IB_TV.Fields("Month").Value
        
        If Month_1 = 12 Then
            Month_2 = 0
            Month_3 = 0
            
            Str_Week_Comm_1 = Get_WeekCommencing(Month_1, Year_Ib)
            Week_Count_Month1 = Len(Str_Week_Comm_1) / 2
            Week_Count_Month2 = 4
            Week_Count_Month3 = 4
            
        ElseIf Month_1 = 11 Then
            Month_2 = Month_1 + 1
            Month_3 = 0
            
            Str_Week_Comm_1 = Get_WeekCommencing(Month_1, Year_Ib)
            Week_Count_Month1 = Len(Str_Week_Comm_1) / 2
            Str_Week_Comm_2 = Get_WeekCommencing(Month_2, Year_Ib)
            Week_Count_Month2 = Len(Str_Week_Comm_2) / 2
            
            Week_Count_Month3 = 4
        Else
            Month_2 = Month_1 + 1
            Month_3 = Month_2 + 1
            
            Str_Week_Comm_1 = Get_WeekCommencing(Month_1, Year_Ib)
            Str_Week_Comm_2 = Get_WeekCommencing(Month_2, Year_Ib)
            Str_Week_Comm_3 = Get_WeekCommencing(Month_3, Year_Ib)
            Week_Count_Month1 = Len(Str_Week_Comm_1) / 2
            Week_Count_Month2 = Len(Str_Week_Comm_2) / 2
            Week_Count_Month3 = Len(Str_Week_Comm_3) / 2
        End If
        
        'Tampil Month1
        Week_Position = 1
        If Month_1 <> 0 Then
            Rpt_TV_IB.Sections("Section2").Controls("Lbl_Month_CO_1").Caption = Get_Month_Name(Month_1)
            
            'Width Lable
            Rpt_TV_IB.Sections("Section2").Controls("Lbl_Month_CO_1").Width = Week_Count_Month1 * 335
            
            'Width Frame
            Rpt_TV_IB.Sections("Section2").Controls("Shape_Month_1").Width = Week_Count_Month1 * 335
             
            'Tampillkan Week Comm

            Do While Str_Week_Comm_1 <> ""
                Rpt_TV_IB.Sections("Section2").Controls("Lbl_Week_" & Week_Position).Caption = Mid(Str_Week_Comm_1, 1, 2)
                Str_Week_Comm_1 = Right(Str_Week_Comm_1, Len(Str_Week_Comm_1) - 2)
                Week_Position = Week_Position + 1
            Loop
           
        End If
        
        'Tampil Month2
        If Month_2 <> 0 Then
            Rpt_TV_IB.Sections("Section2").Controls("Lbl_Month_CO_2").Caption = Get_Month_Name(Month_2)
            
            'Width Lable
            Rpt_TV_IB.Sections("Section2").Controls("Lbl_Month_CO_2").Width = Week_Count_Month2 * 335
            'Width Frame
            Rpt_TV_IB.Sections("Section2").Controls("Shape_Month_2").Width = Week_Count_Month2 * 335
             
            'Tampillkan Week Comm

            Do While Str_Week_Comm_2 <> ""
                Rpt_TV_IB.Sections("Section2").Controls("Lbl_Week_" & Week_Position).Caption = Mid(Str_Week_Comm_2, 1, 2)
                Str_Week_Comm_2 = Right(Str_Week_Comm_2, Len(Str_Week_Comm_2) - 2)
                Week_Position = Week_Position + 1
            Loop
            
        Else
            'Width Lable
            Rpt_TV_IB.Sections("Section2").Controls("Lbl_Month_CO_2").Width = 5 * 335
            'Width Frame
            Rpt_TV_IB.Sections("Section2").Controls("Shape_Month_2").Width = 5 * 335
             
        End If
        
        'Tampil Month3
        If Month_3 <> 0 Then
            Rpt_TV_IB.Sections("Section2").Controls("Lbl_Month_CO_3").Caption = Get_Month_Name(Month_3)
           
             'Width Lable
            Rpt_TV_IB.Sections("Section2").Controls("Lbl_Month_CO_3").Width = Week_Count_Month3 * 335
            'Width Frame
            Rpt_TV_IB.Sections("Section2").Controls("Shape_Month_3").Width = Week_Count_Month3 * 335
             
            'Tampillkan Week Comm

            Do While Str_Week_Comm_3 <> ""
                Rpt_TV_IB.Sections("Section2").Controls("Lbl_Week_" & Week_Position).Caption = Mid(Str_Week_Comm_3, 1, 2)
                Str_Week_Comm_3 = Right(Str_Week_Comm_3, Len(Str_Week_Comm_3) - 2)
                Week_Position = Week_Position + 1
            Loop
        Else
            If Month_2 = 0 Then
                 'Width Lable
                Rpt_TV_IB.Sections("Section2").Controls("Lbl_Month_CO_3").Width = 4 * 335
                'Width Frame
                Rpt_TV_IB.Sections("Section2").Controls("Shape_Month_3").Width = 4 * 335
            
            
            Else
                'Week_Count_Month2
                Rpt_TV_IB.Sections("Section2").Controls("Lbl_Month_CO_3").Width = (14 - (Week_Count_Month2 + Week_Count_Month1)) * 335
                'Width Frame
                Rpt_TV_IB.Sections("Section2").Controls("Shape_Month_3").Width = (14 - (Week_Count_Month2 + Week_Count_Month1)) * 335
                
                
            End If
        End If
        
        'Left Dari Frame & Lable
        Rpt_TV_IB.Sections("Section2").Controls("Lbl_Month_CO_2").Left = Rpt_TV_IB.Sections("Section2").Controls("Lbl_Month_CO_1").Left + Rpt_TV_IB.Sections("Section2").Controls("Lbl_Month_CO_1").Width
        Rpt_TV_IB.Sections("Section2").Controls("Lbl_Month_CO_3").Left = Rpt_TV_IB.Sections("Section2").Controls("Lbl_Month_CO_2").Left + Rpt_TV_IB.Sections("Section2").Controls("Lbl_Month_CO_2").Width
        'Fame
        Rpt_TV_IB.Sections("Section2").Controls("Shape_Month_2").Left = Rpt_TV_IB.Sections("Section2").Controls("Shape_Month_1").Left + Rpt_TV_IB.Sections("Section2").Controls("Shape_Month_1").Width
        Rpt_TV_IB.Sections("Section2").Controls("Shape_Month_3").Left = Rpt_TV_IB.Sections("Section2").Controls("Shape_Month_2").Left + Rpt_TV_IB.Sections("Section2").Controls("Shape_Month_2").Width
        
        If (Week_Count_Month1 + Week_Count_Month2 + Week_Count_Month3) = 13 Then
             Rpt_TV_IB.Sections("Section2").Controls("ShapeCO_1").Visible = False
             Rpt_TV_IB.Sections("Section1").Controls("ShapeCO_2").Visible = False
        End If
        '================== End Tampilan CO ===========
        
        Dim StrField As String
        
        'Open Material
        strSql = "SELECT * FROM IB_TV_Material WHERE Client_Brief_ID='" & cbo_Client_Brief_Id.Text & "' AND IB_ID='" & txt_IB_Id.Text & "'"
        
        RS_Materi.Open strSql, ConnERP, adOpenStatic, adLockPessimistic
        'Loop Marterial
'        Index_Materi = 0
        Do While Not RS_Materi.EOF
            'assign to recordset temp
            RsTempCampaign.AddNew
            RsTempCampaign.Fields("Objective_id").Value = RS_Materi.Fields("Objective_Id").Value & ":" & RS_Materi.Fields("Material_Name").Value & ", " & RS_Materi.Fields("Material_Duration").Value
            
            For Week_Position = 1 To 14
                If Trim(Rpt_TV_IB.Sections("Section2").Controls("Lbl_Week_" & Week_Position).Caption) <> "" Then
                    strSql = " SELECT  * FROM IB_TV_Campaign  WHERE "
                    strSql = strSql & " Objective_Id = " & RS_Materi.Fields("Objective_Id").Value
                    strSql = strSql & " AND Material_Name='" & Clear_String(RS_Materi.Fields("Material_Name").Value) & "'"
                    strSql = strSql & " AND Material_Duration = " & RS_Materi.Fields("Material_Duration").Value
                    If Week_Position <= Week_Count_Month1 Then
                        strSql = strSql & " AND Month(Month_Week_commencing)=" & Month_1
                        
                    ElseIf Week_Position > Week_Count_Month1 And Week_Position <= (Week_Count_Month1 + Week_Count_Month2) Then
                        strSql = strSql & " AND Month(Month_Week_commencing)=" & Month_2
                    Else
                        strSql = strSql & " AND Month(Month_Week_commencing)=" & Month_3
                    End If
                    strSql = strSql & " AND Day(Week_commencing)=" & Trim(Rpt_TV_IB.Sections("Section2").Controls("Lbl_Week_" & Week_Position).Caption)
                   Rs_Co.Open strSql, ConnERP, adOpenStatic, adLockPessimistic
                    
                    StrField = "Tarps_" & Week_Position
                    

                    Do While Not Rs_Co.EOF
                        
                        RsTempCampaign.Fields(StrField).Value = Rs_Co.Fields("Tarps_Per_Week").Value
                        Rs_Co.MoveNext
                    Loop
                    Rs_Co.Close
                End If
            Next Week_Position
            
            RS_Materi.MoveNext
            RsTempCampaign.Update
        Loop
       
        CloseRecordset RS_Materi
        
        Set Rpt_TV_IB.DataSource = RsTempCampaign
        
        'Assign Recordsource
        Rpt_TV_IB.Sections("Section1").Controls("Txt_Material_0").DataField = "Objective_id"
        
        Rpt_TV_IB.Sections("Section1").Controls("Txt_CO_01").DataField = "Tarps_1"
        Rpt_TV_IB.Sections("Section1").Controls("Txt_CO_02").DataField = "Tarps_2"
        Rpt_TV_IB.Sections("Section1").Controls("Txt_CO_03").DataField = "Tarps_3"
        Rpt_TV_IB.Sections("Section1").Controls("Txt_CO_04").DataField = "Tarps_4"
        
        Rpt_TV_IB.Sections("Section1").Controls("Txt_CO_05").DataField = "Tarps_5"
        Rpt_TV_IB.Sections("Section1").Controls("Txt_CO_06").DataField = "Tarps_6"
        Rpt_TV_IB.Sections("Section1").Controls("Txt_CO_07").DataField = "Tarps_7"
        Rpt_TV_IB.Sections("Section1").Controls("Txt_CO_08").DataField = "Tarps_8"
        
        Rpt_TV_IB.Sections("Section1").Controls("Txt_CO_09").DataField = "Tarps_9"
        Rpt_TV_IB.Sections("Section1").Controls("Txt_CO_10").DataField = "Tarps_10"
        Rpt_TV_IB.Sections("Section1").Controls("Txt_CO_11").DataField = "Tarps_11"
        Rpt_TV_IB.Sections("Section1").Controls("Txt_CO_12").DataField = "Tarps_12"
        
        Rpt_TV_IB.Sections("Section1").Controls("Txt_CO_13").DataField = "Tarps_13"
        Rpt_TV_IB.Sections("Section1").Controls("Txt_CO_14").DataField = "Tarps_14"
        
        '================ End Campaign Outline ==========
        
        '---> Budget Split By Month
            'Month
            Rpt_TV_IB.Sections("Section4").Controls("Lbl_Month_1").Caption = " " & Me.Lbl_Month_1.Caption
            Rpt_TV_IB.Sections("Section4").Controls("Lbl_Month_2").Caption = " " & Me.Lbl_Month_2.Caption
            Rpt_TV_IB.Sections("Section4").Controls("Lbl_Month_3").Caption = " " & Me.Lbl_Month_3.Caption
            'Budget
            Rpt_TV_IB.Sections("Section4").Controls("Lbl_Budget_1").Caption = Me.Txt_Month_1.Text
            Rpt_TV_IB.Sections("Section4").Controls("Lbl_Budget_2").Caption = Me.Txt_Month_2.Text
            Rpt_TV_IB.Sections("Section4").Controls("Lbl_Budget_3").Caption = Me.Txt_Month_3.Text
            
        
        
        '---> Program Type Consideration
         Rpt_TV_IB.Sections("Section5").Controls("Lbl_Program_Consideration").Caption = Me.Txt_Prog_Consideration.Text
        '---> Other Consideration
         Rpt_TV_IB.Sections("Section5").Controls("Lbl_Other_Consideration").Caption = Me.Txt_Other_Consideration.Text
        '---> Attachment
         Rpt_TV_IB.Sections("Section5").Controls("Lbl_Attachment").Caption = Me.Txt_Attachments.Text
    Rpt_TV_IB.show 1
End Sub
Private Sub db_add()
'*******************************************
' Procedure     : db_add
' Function      : prosedur add data
' Last Update   : 18/03/2016
'**********************************************
    On Error GoTo errAdd
    If Cbo_Brand.Text = "" Then
        MsgBox "Select Brand...!", vbExclamation, APPLICATION_NAME
        Exit Sub
    End If
   
    Call Clear_Form
    Call EnableObject(True)
    Enable_Form
    cbo_Client_Brief_Id.Clear
    loadCbo_ClientBriefID cbo_Client_Brief_Id, " AND Brand_Code='" & Left(Cbo_Brand.Text, 4) & "' AND Substring(Client_Brief_Id,6,2) = '" & Right(Cbo_Year.Text, 2) & "'"
    Initialize_Temp_Table
    mNew = True
    Exit Sub
   
errAdd:
   MsgBox "[db_add] error Number :" & Err.Number & " - " & Err.Description, vbCritical, Str_Company_Name
  
End Sub

Private Sub picButton_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
'************************************************
' Procedure         : picButton_MouseDown
' Function          : TOOLBAR_AI saat mouse ditekan.
' Created By        : {73 64 6B}
' Date              : 12-Apr-2015
'************************************************
    picButton(Index).Picture = LoadPicture(SetButtonImageEffect(Index, bieDown)) 'FIRST.

End Sub

Private Sub picButton_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
'************************************************
' Procedure         : picButton_MouseMove
' Function          : TOOLBAR_AI saat mouse berada di area button.
' Created By        : {73 64 6B}
' Date              : 12-Apr-2015
'************************************************
    picButton_Obj Index, Button, Shift, X, Y, picButton

End Sub

Private Sub txt_IB_Id_KeyPress(KeyAscii As Integer)
'************************************************
' Procedure         : txt_IB_Id_KeyPress
' Function          : txt_IB_Id KeyPress Event Handler.
' Created By        : {73 64 6B}
' Date              : 12-Apr-2015
'************************************************
    If mNew Then KeyAscii = 0
End Sub

Private Sub txt_IB_Id_LostFocus()
'************************************************
' Procedure         : txt_IB_Id_LostFocus
' Function          : txt_IB_Id LostFocus Event Handler.
' Created By        : {73 64 6B}
' Date              : 12-Apr-2015
'************************************************
    If txt_IB_Id.Enabled = False Then Exit Sub
    If mNew Then Exit Sub
    
    Find_IB_ID
End Sub
Private Sub Find_IB_ID()
'************************************************
' Procedure         : Find_IB_ID
' Function          : Mencari Nilai IB ID.
' Created By        : {73 64 6B}
' Date              : 12-Apr-2015
'************************************************
    If rs_IB_TV.RecordCount < 1 Then
        MsgBox "No Data Record To Find", vbInformation, APPLICATION_NAME
        Exit Sub
    End If
    Dim intIB As Integer
    With rs_IB_TV
    .MoveFirst
    For intIB = 1 To .RecordCount
        If Trim(!IB_ID) = Trim(txt_IB_Id.Text) Then
            rPost = intIB
            .AbsolutePosition = rPost
            LoadToForm
            Exit Sub
        End If
        
        .MoveNext
    Next
   
    MsgBox "Invalid IB ID", vbInformation, APPLICATION_NAME
    .AbsolutePosition = rPost
    LoadToForm
    Exit Sub

    End With

End Sub
Sub db_save()
    Dim Total_Montly_Budget As Currency
    Dim Total_Busrt_Budget As Currency

    On Error GoTo ErrLable

    If add_flag Then
    'Validasi Field
        
        Dim VrnBrandVariant As Variant
        
        
        'Primary Target
        If txt_Primary.Text = "" Then
            MsgBox "Primary Target Empty", vbCritical, "Missing Information"
            txt_Primary.SetFocus
            Exit Sub
        End If
        
        'Secondary Target
        If Cbo_Cluster.Text = "" Or Cbo_Cluster.ListIndex = -1 Then
            MsgBox "Cluster Empty", vbCritical, "Missing Information"
            Cbo_Cluster.SetFocus
            Exit Sub
        End If
        
        'Media Plan Number
        If Trim(Txt_Plan_No.Text) = "" Then
            MsgBox "Please Insert Plan Number !", vbCritical, "Missing Information"
            Txt_Plan_No.SetFocus
            Exit Sub
        End If
        
        
        'Television Objective
        rsTemp_Objective.Filter = ""
        If rsTemp_Objective.BOF And rsTemp_Objective.EOF Then
            MsgBox "Television Objective Empty", vbCritical, "Missing Information"
            Exit Sub
        End If
      
        'Campign outline
        Rs_Temp_Camp_Outline.Filter = ""
        If Rs_Temp_Camp_Outline.BOF And Rs_Temp_Camp_Outline.EOF Then
            MsgBox "Campaign Outline Empty", vbCritical, "Missing Information"
            Exit Sub
        End If
        
        'Montly Budget
        If Txt_Month_1.Text = "" And Txt_Month_2.Text = "" And Txt_Month_3.Text = "" Then
            MsgBox "Montly Budget Empty", vbCritical, "Missing Information"
            Exit Sub
        End If
        
        'Total Montly Budget and total Budget every Burst
        Total_Montly_Budget = CCur(Format(IIf(Txt_Month_1.Text = "", "0", Txt_Month_1.Text), "####0")) + CCur(Format(IIf(Txt_Month_2.Text = "", "0", Txt_Month_2.Text), "####0")) + CCur(Format(IIf(Txt_Month_3.Text = "", "0", Txt_Month_3.Text), "####0"))
        Total_Busrt_Budget = 0
        For Index_Row = 1 To Flex_T_O.Rows - 1
            Total_Busrt_Budget = Total_Busrt_Budget + CCur(IIf(Flex_T_O.TextMatrix(Index_Row, 6) = "", "0", Format(Flex_T_O.TextMatrix(Index_Row, 6), "####0")))
            Total_Busrt_Budget = Total_Busrt_Budget + CCur(IIf(Flex_T_O.TextMatrix(Index_Row, 7) = "", "0", Format(Flex_T_O.TextMatrix(Index_Row, 7), "####0")))
        Next Index_Row

        If Total_Montly_Budget <> Total_Busrt_Budget Then
            MsgBox "Total Montly Budget doesn't match with Total Busrt Budget", vbCritical, APPLICATION_NAME
            Exit Sub
        End If
        '=====================
    
    '************* IB TV ************************
    
    'Begin Transaction
        ConnERP.BeginTrans
        
        rs_IB_TV.AddNew
        'Assigment
        rs_IB_TV.Fields("Client_Brief_Id").Value = cbo_Client_Brief_Id.Text
        rs_IB_TV.Fields("IB_Id").Value = txt_IB_Id.Text
        
        rs_IB_TV.Fields("Year").Value = Cbo_Year.Text
        rs_IB_TV.Fields("Month").Value = Get_Month_Number(Cbo_Month.Text)
        
        rs_IB_TV.Fields("Date_Entered").Value = DT_Date.Value
        rs_IB_TV.Fields("Entered_By").Value = Txt_Enterd_By.Text
        
        rs_IB_TV.Fields("Target_Primary").Value = txt_Primary.Text
        rs_IB_TV.Fields("Target_Secondary").Value = Cbo_Cluster.Text
        rs_IB_TV.Fields("Attachment").Value = Txt_Attachments.Text
        rs_IB_TV.Fields("Consideration").Value = Txt_Prog_Consideration.Text
        rs_IB_TV.Fields("Other_Consideration").Value = Txt_Other_Consideration.Text
        
        
        VrnBrandVariant = Split(Cbo_Brand_Variant.Text, "-->")
        rs_IB_TV.Fields("Brand_Variant_Code").Value = Trim(VrnBrandVariant(0))
        rs_IB_TV.Fields("Brand_Variant_Name").Value = Trim(VrnBrandVariant(1))
        
        VrnBrandVariant = Split(Cbo_Brand.Text, "-->")
        rs_IB_TV.Fields("Brand_Name").Value = Trim(VrnBrandVariant(1))
        rs_IB_TV.Fields("Cluster_Code").Value = Cbo_Cluster.ItemData(Cbo_Cluster.ListIndex)
        
        If add_flag Then
            rs_IB_TV.Fields("Client_Approval").Value = 0
        End If
        
        rs_IB_TV.Fields("Plan_No").Value = Txt_Plan_No.Text
        rs_IB_TV.Fields("user_id").Value = UserName
        rs_IB_TV.Fields("stamp").Value = Now()
    'Save Data
        rs_IB_TV.Update
     
    '************* Montly Budget ***************
    'Assigment -- > Save
    If add_flag Then
        If Txt_Month_1.Visible Then
            If Txt_Month_1.Text <> "" Then
                strSql = ""
                strSql = "INSERT INTO IB_TV_Montly_Budget VALUES('"
                strSql = strSql & cbo_Client_Brief_Id.Text & "','"
                strSql = strSql & txt_IB_Id.Text & "',"
                strSql = strSql & Get_Month_Number(Lbl_Month_1.Caption) & ","
                strSql = strSql & Get_Month_Number(Cbo_Month.Text) & ","
                strSql = strSql & Format(Txt_Month_1.Text, "######0") & ")"
                ConnERP.Execute strSql
            End If
        End If
        If Txt_Month_2.Visible Then
            If Txt_Month_2.Text <> "" Then
                strSql = ""
                strSql = "INSERT INTO IB_TV_Montly_Budget VALUES('"
                strSql = strSql & cbo_Client_Brief_Id.Text & "','"
                strSql = strSql & txt_IB_Id.Text & "',"
                strSql = strSql & Get_Month_Number(Lbl_Month_2.Caption) & ","
                strSql = strSql & Get_Month_Number(Cbo_Month.Text) & ","
                strSql = strSql & Format(Txt_Month_2.Text, "######0") & ")"
                ConnERP.Execute strSql
            End If
        End If
        If Txt_Month_3.Visible Then
            If Txt_Month_3.Text <> "" Then
                strSql = ""
                strSql = "INSERT INTO IB_TV_Montly_Budget VALUES('"
                strSql = strSql & cbo_Client_Brief_Id.Text & "','"
                strSql = strSql & txt_IB_Id.Text & "',"
                strSql = strSql & Get_Month_Number(Lbl_Month_3.Caption) & ","
                strSql = strSql & Get_Month_Number(Cbo_Month.Text) & ","
                strSql = strSql & Format(Txt_Month_3.Text, "######0") & ")"
                ConnERP.Execute strSql
            End If
        End If
    End If
    
    
    
    
    '************* TV Objective *****************
    'when Add
    'Dim Index_Row As Integer
        If add_flag Then
        'Assigment
        Index_Row = 0
        With rsTemp_Objective
            .Filter = ""
            .MoveFirst
            Do While Not .BOF And Not .EOF
                strSql = ""
                Index_Row = Index_Row + 1
                strSql = "INSERT INTO IB_TV_Objective(Objective_Id, Client_Brief_Id, IB_Id, Week_Commencing, Campaign_Type, "
                strSql = strSql & " Frequency, Reach, Tarps, "
                strSql = strSql & " Budget_With_MSC, Campaign_Type_Code, "
                strSql = strSql & " Frequency_Code, W_C_From, W_C_To, "
                strSql = strSql & " Nett, MSC, Market_Code, Market_Name, Status) VALUES("
                
                strSql = strSql & .Fields("Objective_Id").Value & ",'"
                strSql = strSql & .Fields("Client_Brief_Id").Value & "','"
                strSql = strSql & .Fields("IB_ID").Value & "','"
                strSql = strSql & .Fields("Week_Commencing").Value & "','"
                strSql = strSql & .Fields("Campaign_Type").Value & "','"
                strSql = strSql & .Fields("Frequency").Value & "',"
                strSql = strSql & .Fields("Reach").Value & ","
                strSql = strSql & .Fields("Tarps").Value & ","
                strSql = strSql & .Fields("Budget_With_MSC").Value & ","
                strSql = strSql & .Fields("Campaign_Type_Code").Value & ","
                strSql = strSql & .Fields("Frequency_Code").Value & ",'"
                strSql = strSql & .Fields("W_C_From").Value & "','"
                strSql = strSql & .Fields("W_C_To").Value & "',"
                strSql = strSql & .Fields("Nett").Value & ","
                strSql = strSql & .Fields("MSC").Value & ","
                strSql = strSql & .Fields("Market_Code").Value & ",'"
                strSql = strSql & .Fields("Market_Name").Value & "',"
                strSql = strSql & .Fields("Status").Value & ")"
'                StrSQL = StrSQL & .Fields("Replace_By").Value & "',"
'                StrSQL = StrSQL & .Fields("Replace_Date").Value & "',"
'                StrSQL = StrSQL & .Fields("Cancel_By").Value & "',"
'                StrSQL = StrSQL & .Fields("Cancel_Date").Value & "')"
                
                ConnERP.Execute strSql
                .MoveNext
            Loop
        'Save Data
        End With
        End If
    
    '*************  IB TV Material **************
    'Assigment --> 'Save Data
    If add_flag Then
        With Rs_Materi_Temp
        .Filter = ""
        .MoveFirst
        Do While Not .BOF And Not .EOF
            strSql = ""
            strSql = "INSERT INTO IB_TV_Objective_Material(Client_Brief_Id, IB_Id, Objective_Id, Material_Name, Material_Duration) VALUES('"
            strSql = strSql & .Fields("Client_Brief_Id").Value & "','"
            strSql = strSql & .Fields("IB_Id").Value & "',"
            strSql = strSql & .Fields("Objective_Id").Value & ",'"
            strSql = strSql & Trim$(Clear_String(.Fields("Material_Name").Value)) & "','"
            strSql = strSql & .Fields("Material_Duration").Value & "')"
            ConnERP.Execute strSql
            .MoveNext
        Loop
        End With
    End If
    
    '************* Campaign Outline *************
    'Assigment
    'when Add
        If add_flag Then
        'Assigment
        With Rs_Temp_Camp_Outline
            .Filter = ""
            .MoveFirst
            Do While Not .BOF And Not .EOF
                strSql = ""
                strSql = "INSERT INTO IB_TV_Campaign(Campaign_Id, Client_Brief_ID, IB_Id, Tarps_Per_Week,"
                strSql = strSql & " Objective_Id, Month_Week_commencing, Week_Commencing, "
                strSql = strSql & " Material_Name, Material_Duration) VALUES("
                strSql = strSql & .Fields("Campaign_Id").Value & ",'"
                strSql = strSql & .Fields("Client_Brief_ID").Value & "','"
                strSql = strSql & .Fields("Ib_Id").Value & "',"
                strSql = strSql & .Fields("Tarps_Per_Week").Value & ","
                strSql = strSql & .Fields("Objective_Id").Value & ",'"
                strSql = strSql & .Fields("Month_Week_commencing").Value & "','"
                strSql = strSql & .Fields("Week_Commencing").Value & "','"
                strSql = strSql & Trim$(Clear_String(.Fields("Material_Name").Value)) & "',"
                strSql = strSql & .Fields("Material_Duration").Value & ")"
                
                ConnERP.Execute strSql
                .MoveNext
            Loop
        'Save Data
        End With
        End If
    'Flag
    
    'Show Data
        Deinitialize_Temp_Table
        
    'Enable Button
        add_flag = False
        
        Button_Normal (True)
        Television_Objective_Button False
        
        
'===================================================
'           Save Waktu Edit
'===================================================
    Else
        'Validasi Field
        'Primary Target
        If txt_Primary.Text = "" Then
            MsgBox "Primary Target Empty", vbCritical, "Missing Information"
            txt_Primary.SetFocus
            Exit Sub
        End If
        
        'Secondary Target
        If Cbo_Cluster.Text = "" Or Cbo_Cluster.ListIndex = -1 Then
            MsgBox "Cluster Empty", vbCritical, "Missing Information"
            Cbo_Cluster.SetFocus
            Exit Sub
        End If
        
        'Media Plan Number
        If Trim(Txt_Plan_No.Text) = "" Then
            MsgBox "Please Insert Plan Number !", vbCritical, "Missing Information"
            Txt_Plan_No.SetFocus
            Exit Sub
        End If
        
        'Television Objective
        
        'Campaign Outline
        
        'Montly Budget
        If Txt_Month_1.Text = "" And Txt_Month_2.Text = "" And Txt_Month_3.Text = "" Then
            MsgBox "Montly Budget Empty", vbCritical, APPLICATION_NAME
            Exit Sub
        End If
        
        'Total Montly Budget and total Budget every Burst
        Total_Montly_Budget = CCur(Format(IIf(Txt_Month_1.Text = "", "0", Txt_Month_1.Text), "####0")) + CCur(Format(IIf(Txt_Month_2.Text = "", "0", Txt_Month_2.Text), "####0")) + CCur(Format(IIf(Txt_Month_3.Text = "", "0", Txt_Month_3.Text), "####0"))
        Total_Busrt_Budget = 0
        For Index_Row = 1 To Flex_T_O.Rows - 1
            Total_Busrt_Budget = Total_Busrt_Budget + CCur(IIf(Flex_T_O.TextMatrix(Index_Row, 6) = "", "0", Format(Flex_T_O.TextMatrix(Index_Row, 6), "####0")))
            Total_Busrt_Budget = Total_Busrt_Budget + CCur(IIf(Flex_T_O.TextMatrix(Index_Row, 7) = "", "0", Format(Flex_T_O.TextMatrix(Index_Row, 7), "####0")))
        Next Index_Row
        If Total_Montly_Budget <> Total_Busrt_Budget Then
            MsgBox "Total Montly Budget does not equal with Total Busrt Budget", vbCritical, APPLICATION_NAME
            Exit Sub
        End If
        
        'Begin Transaction
        ConnERP.BeginTrans
        
        'Save Data
        'Assigment IB TV
        rs_IB_TV.Fields("Target_Primary").Value = txt_Primary.Text
        rs_IB_TV.Fields("Cluster_Code").Value = Cbo_Cluster.ItemData(Cbo_Cluster.ListIndex)
        rs_IB_TV.Fields("Target_Secondary").Value = Cbo_Cluster.Text
        rs_IB_TV.Fields("Attachment").Value = Txt_Attachments.Text
        rs_IB_TV.Fields("Consideration").Value = Txt_Prog_Consideration.Text
        rs_IB_TV.Fields("Other_Consideration").Value = Txt_Other_Consideration.Text
        'rs_Date.Requery
        rs_IB_TV.Fields("Date_Entered").Value = getDate
        rs_IB_TV.Fields("Entered_By").Value = FullName
        rs_IB_TV.Fields("Plan_No").Value = Txt_Plan_No.Text
         rs_IB_TV.Fields("user_id").Value = UserName
        rs_IB_TV.Fields("stamp").Value = Now()
        rs_IB_TV.Update
        
        'Delete Existing Data
        
            'Campaign outline
            strSql = "DELETE FROM ib_tv_Campaign WHERE "
            strSql = strSql & " Objective_id IN (SELECT Objective_id FROM IB_TV_Objective"
            strSql = strSql & " WHERE IB_ID='" & txt_IB_Id.Text & "')"
            ConnERP.Execute strSql
            
            'Material Ib
            strSql = "DELETE FROM IB_TV_Objective_Material "
            strSql = strSql & " WHERE Objective_id IN (SELECT Objective_id FROM IB_TV_Objective"
            strSql = strSql & " WHERE IB_ID='" & txt_IB_Id.Text & "')"
            ConnERP.Execute strSql
                        
            'Television Objective
            strSql = "DELETE FROM ib_tv_objective WHERE IB_ID='" & txt_IB_Id.Text & "'"
            ConnERP.Execute strSql

'=========================  Add Existing Data ================
        With rsTemp_Objective
            .Filter = ""
            .MoveFirst
            Do While Not .BOF And Not .EOF
                strSql = ""
                Index_Row = Index_Row + 1
                strSql = "INSERT INTO IB_TV_Objective(Objective_Id, Client_Brief_Id, IB_Id, Week_Commencing, Campaign_Type, "
                strSql = strSql & " Frequency, Reach, Tarps, "
                strSql = strSql & " Budget_With_MSC, Campaign_Type_Code, "
                strSql = strSql & " Frequency_Code, W_C_From, W_C_To, "
                strSql = strSql & " Nett, MSC, Market_Code, Market_Name, Status) VALUES("
                
                strSql = strSql & .Fields("Objective_Id").Value & ",'"
                strSql = strSql & .Fields("Client_Brief_Id").Value & "','"
                strSql = strSql & .Fields("IB_ID").Value & "','"
                strSql = strSql & .Fields("Week_Commencing").Value & "','"
                strSql = strSql & .Fields("Campaign_Type").Value & "','"
                strSql = strSql & .Fields("Frequency").Value & "',"
                strSql = strSql & .Fields("Reach").Value & ","
                strSql = strSql & .Fields("Tarps").Value & ","
                strSql = strSql & .Fields("Budget_With_MSC").Value & ","
                strSql = strSql & .Fields("Campaign_Type_Code").Value & ","
                strSql = strSql & .Fields("Frequency_Code").Value & ",'"
                strSql = strSql & .Fields("W_C_From").Value & "','"
                strSql = strSql & .Fields("W_C_To").Value & "',"
                strSql = strSql & .Fields("Nett").Value & ","
                strSql = strSql & .Fields("MSC").Value & ","
                strSql = strSql & .Fields("Market_Code").Value & ",'"
                strSql = strSql & .Fields("Market_Name").Value & "',"
                strSql = strSql & .Fields("Status").Value & ")"
'                StrSQL = StrSQL & .Fields("Replace_By").Value & "',"
'                StrSQL = StrSQL & .Fields("Replace_Date").Value & "',"
'                StrSQL = StrSQL & .Fields("Cancel_By").Value & "',"
'                StrSQL = StrSQL & .Fields("Cancel_Date").Value & "')"
                
                ConnERP.Execute strSql
                .MoveNext
            Loop
        'Save Data
        End With
        
        With Rs_Materi_Temp
        .Filter = ""
        .MoveFirst
        Do While Not .BOF And Not .EOF
            strSql = ""
            strSql = "INSERT INTO IB_TV_Objective_Material(Client_Brief_Id, IB_Id, Objective_Id, Material_Name, Material_Duration) VALUES('"
            strSql = strSql & .Fields("Client_Brief_Id").Value & "','"
            strSql = strSql & .Fields("IB_Id").Value & "',"
            strSql = strSql & .Fields("Objective_Id").Value & ",'"
            strSql = strSql & Trim$(Clear_String(.Fields("Material_Name").Value)) & "','"
            strSql = strSql & .Fields("Material_Duration").Value & "')"
            ConnERP.Execute strSql
            .MoveNext
        Loop
        End With
        
        With Rs_Temp_Camp_Outline
            .Filter = ""
            .MoveFirst
            Do While Not .BOF And Not .EOF
                strSql = ""
                strSql = "INSERT INTO IB_TV_Campaign(Campaign_Id, Client_Brief_ID, IB_Id, Tarps_Per_Week,"
                strSql = strSql & " Objective_Id, Month_Week_commencing, Week_Commencing, "
                strSql = strSql & " Material_Name, Material_Duration) VALUES("
                strSql = strSql & .Fields("Campaign_Id").Value & ",'"
                strSql = strSql & .Fields("Client_Brief_ID").Value & "','"
                strSql = strSql & .Fields("Ib_Id").Value & "',"
                strSql = strSql & .Fields("Tarps_Per_Week").Value & ","
                strSql = strSql & .Fields("Objective_Id").Value & ",'"
                strSql = strSql & .Fields("Month_Week_commencing").Value & "','"
                strSql = strSql & .Fields("Week_Commencing").Value & "','"
                strSql = strSql & Trim$(Clear_String(.Fields("Material_Name").Value)) & "',"
                strSql = strSql & .Fields("Material_Duration").Value & ")"
                
                ConnERP.Execute strSql
                .MoveNext
            Loop
        'Save Data
        End With
        
        

        
        'Montly Budget **************
        'Delete Previouse Data
        
        strSql = ""
        strSql = "DELETE FROM IB_TV_Montly_Budget WHERE IB_ID='" & Me.txt_IB_Id.Text & "'"
        ConnERP.Execute strSql
         If Txt_Month_1.Visible Then
            If Txt_Month_1.Text <> "" Then
                strSql = ""
                strSql = "INSERT INTO IB_TV_Montly_Budget VALUES('"
                strSql = strSql & cbo_Client_Brief_Id.Text & "','"
                strSql = strSql & txt_IB_Id.Text & "',"
                strSql = strSql & Get_Month_Number(Lbl_Month_1.Caption) & ","
                strSql = strSql & Get_Month_Number(Cbo_Month.Text) & ","
                strSql = strSql & Format(Txt_Month_1.Text, "######0") & ")"
                ConnERP.Execute strSql
            End If
        End If
        
        If Txt_Month_2.Visible Then
Balik_Ke_2:
            If Txt_Month_2.Text <> "" Then
                If Lbl_Month_2.Caption = "" Then
                    Lbl_Month_2.Caption = Get_Month_Name(Val(InputBox("Insert Month that you added ( 1 - 12 )!", APPLICATION_NAME, Get_Month_Number(Lbl_Month_1.Caption) + 1)))
                End If
                If Get_Month_Number(Lbl_Month_2.Caption) = 0 Then
                    GoTo Balik_Ke_2
                End If
                strSql = ""
                strSql = "INSERT INTO IB_TV_Montly_Budget VALUES('"
                strSql = strSql & cbo_Client_Brief_Id.Text & "','"
                strSql = strSql & txt_IB_Id.Text & "',"
                strSql = strSql & Get_Month_Number(Lbl_Month_2.Caption) & ","
                strSql = strSql & Get_Month_Number(Cbo_Month.Text) & ","
                strSql = strSql & Format(Txt_Month_2.Text, "######0") & ")"
                ConnERP.Execute strSql
            End If
        End If
        If Txt_Month_3.Visible Then
Balik_Ke_3:
            If Txt_Month_3.Text <> "" Then
                If Lbl_Month_3.Caption = "" Then
                    Lbl_Month_3.Caption = Get_Month_Name(Val(InputBox("Insert Month that you added ( 1 - 12 ) !", APPLICATION_NAME, Get_Month_Number(Lbl_Month_2.Caption) + 1)))
                End If
                
                If Get_Month_Number(Lbl_Month_3.Caption) = 0 Then
                    GoTo Balik_Ke_3
                End If
                
                strSql = ""
                strSql = "INSERT INTO IB_TV_Montly_Budget VALUES('"
                strSql = strSql & cbo_Client_Brief_Id.Text & "','"
                strSql = strSql & txt_IB_Id.Text & "',"
                strSql = strSql & Get_Month_Number(Lbl_Month_3.Caption) & ","
                strSql = strSql & Get_Month_Number(Cbo_Month.Text) & ","
                strSql = strSql & Format(Txt_Month_3.Text, "######0") & ")"
                ConnERP.Execute strSql
            
            End If
        End If
        'Edit Flag
        Edit_Flag = False
        
        'Button
        Button_Normal (True)
        Television_Objective_Button False
        
        'De initialize temp table
        Deinitialize_Temp_Table
    End If
    
       
    'End Transaction
    ConnERP.CommitTrans
    
    Dim Actual_IB_Id As String
    Actual_IB_Id = Me.txt_IB_Id.Text
    
    rs_IB_TV.Requery
    
    No_Record = False
    
    'Find Actual Record
    rs_IB_TV.Find "IB_Id='" & Actual_IB_Id & "'"
    If Not rs_IB_TV.EOF Then
        LoadToForm
        Call picButton_Click(11)
    End If
    Exit Sub
    
    
ErrLable:
 On Error Resume Next
    If Err.Number = -2147217893 Then
        MsgBox Err.Description, vbCritical, APPLICATION_NAME
        
        ConnERP.RollbackTrans
        rs_IB_TV.Requery
        
        Call picButton_Click(11)
    ElseIf Err.Number = -2147217873 Then
        MsgBox "Double Data", vbCritical, APPLICATION_NAME
        ConnERP.RollbackTrans
        rs_IB_TV.Requery
        
        Call picButton_Click(11)
    Else
        MsgBox Err.Number
        MsgBox Err.Description
        
        ConnERP.RollbackTrans
        
        rs_IB_TV.Requery
        MsgBox Err.Description & vbCrLf & "Please Contact IT Department. ", vbCritical, APPLICATION_NAME
        
        Call picButton_Click(11)
    End If
    
End Sub

