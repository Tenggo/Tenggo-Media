VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "threed32.ocx"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form Frm_Radio_Media_Quot 
   BorderStyle     =   0  'None
   Caption         =   "Radio Media Quotation"
   ClientHeight    =   9360
   ClientLeft      =   0
   ClientTop       =   -45
   ClientWidth     =   13935
   Icon            =   "Frm_Radio_Media_Quot.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   9360
   ScaleWidth      =   13935
   ShowInTaskbar   =   0   'False
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
      ScaleWidth      =   13935
      TabIndex        =   43
      TabStop         =   0   'False
      Top             =   9030
      Width           =   13935
      Begin VB.PictureBox picDescColor 
         BorderStyle     =   0  'None
         Height          =   345
         Left            =   9975
         ScaleHeight     =   345
         ScaleWidth      =   1695
         TabIndex        =   48
         TabStop         =   0   'False
         Top             =   15
         Width           =   1695
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
         TabIndex        =   47
         TabStop         =   0   'False
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
         TabIndex        =   46
         TabStop         =   0   'False
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
         TabIndex        =   45
         TabStop         =   0   'False
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
         Index           =   1
         Left            =   420
         ScaleHeight     =   345
         ScaleWidth      =   300
         TabIndex        =   44
         TabStop         =   0   'False
         Top             =   15
         Width           =   300
      End
      Begin VB.Label lblLastModifiedBy 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Last Modified By:                                 |"
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
         Left            =   5145
         TabIndex        =   50
         Tag             =   "Last Modified by: "
         Top             =   75
         Width           =   2775
      End
      Begin VB.Label lblLastModifiedDate 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
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
         TabIndex        =   49
         Tag             =   "Last Modified Date: "
         Top             =   75
         Width           =   2520
      End
   End
   Begin Threed.SSPanel pnl_Main 
      Align           =   1  'Align Top
      Height          =   7830
      Left            =   0
      TabIndex        =   12
      Top             =   750
      Width           =   13935
      _Version        =   65536
      _ExtentX        =   24580
      _ExtentY        =   13811
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
      Begin VB.CommandButton Cmd_Approved_MQ 
         Caption         =   "Approve Media Quotation List"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   9945
         TabIndex        =   69
         Top             =   5055
         Visible         =   0   'False
         Width           =   3120
      End
      Begin VB.CommandButton Cmd_History_Revision 
         Caption         =   "&History Revision Media Quotation"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   10665
         TabIndex        =   68
         Top             =   1875
         Visible         =   0   'False
         Width           =   2970
      End
      Begin VB.Frame Frame9 
         Height          =   825
         Left            =   9195
         TabIndex        =   64
         Top             =   3630
         Visible         =   0   'False
         Width           =   3855
         Begin VB.CommandButton Cmd_Close 
            Caption         =   "C&lose"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   525
            Left            =   2820
            Picture         =   "Frm_Radio_Media_Quot.frx":0442
            TabIndex        =   67
            Top             =   180
            Width           =   915
         End
         Begin VB.CommandButton Cmd_Print 
            Caption         =   "&Print"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   525
            Left            =   1815
            Picture         =   "Frm_Radio_Media_Quot.frx":074C
            TabIndex        =   66
            Top             =   180
            Width           =   915
         End
         Begin VB.CommandButton Cmd_Pulish_to_Web 
            Caption         =   "Publish to Web"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   525
            Left            =   150
            Picture         =   "Frm_Radio_Media_Quot.frx":0A56
            TabIndex        =   65
            Top             =   180
            Width           =   1290
         End
      End
      Begin VB.Frame Frame7 
         Height          =   825
         Left            =   9510
         TabIndex        =   57
         Top             =   2700
         Visible         =   0   'False
         Width           =   2460
         Begin VB.PictureBox Picture1 
            Height          =   528
            Left            =   108
            ScaleHeight     =   465
            ScaleWidth      =   2175
            TabIndex        =   58
            Top             =   195
            Width           =   2232
            Begin VB.CommandButton Cmd_Cancel 
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
               Height          =   495
               Left            =   732
               TabIndex        =   63
               Top             =   0
               Width           =   720
            End
            Begin VB.CommandButton Cmd_Save 
               Caption         =   "&Save"
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
               Left            =   0
               TabIndex        =   62
               Top             =   0
               Width           =   720
            End
            Begin VB.CommandButton Cmd_Edit 
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
               Height          =   465
               Left            =   744
               TabIndex        =   61
               Top             =   0
               Width           =   720
            End
            Begin VB.CommandButton Cmd_Add 
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
               Height          =   465
               Left            =   15
               TabIndex        =   60
               Top             =   0
               Width           =   720
            End
            Begin VB.CommandButton Cmd_Delete 
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
               Height          =   465
               Left            =   1464
               TabIndex        =   59
               Top             =   0
               Width           =   705
            End
         End
      End
      Begin VB.Frame Frame10 
         Height          =   825
         Left            =   9525
         TabIndex        =   51
         Top             =   1920
         Visible         =   0   'False
         Width           =   2400
         Begin VB.PictureBox Picture3 
            Height          =   570
            Left            =   135
            ScaleHeight     =   510
            ScaleWidth      =   2085
            TabIndex        =   52
            Top             =   180
            Width           =   2145
            Begin VB.CommandButton Cmd_Last 
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   495
               Left            =   1590
               Picture         =   "Frm_Radio_Media_Quot.frx":0D60
               Style           =   1  'Graphical
               TabIndex        =   56
               Top             =   0
               Width           =   495
            End
            Begin VB.CommandButton Cmd_Next 
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   495
               Left            =   1060
               Picture         =   "Frm_Radio_Media_Quot.frx":0EAA
               Style           =   1  'Graphical
               TabIndex        =   55
               Top             =   0
               Width           =   495
            End
            Begin VB.CommandButton Cmd_Previous 
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   495
               Left            =   530
               Picture         =   "Frm_Radio_Media_Quot.frx":0FF4
               Style           =   1  'Graphical
               TabIndex        =   54
               Top             =   0
               Width           =   495
            End
            Begin VB.CommandButton Cmd_First 
               DownPicture     =   "Frm_Radio_Media_Quot.frx":113E
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   495
               Left            =   0
               Picture         =   "Frm_Radio_Media_Quot.frx":1288
               Style           =   1  'Graphical
               TabIndex        =   53
               Top             =   0
               Width           =   495
            End
         End
      End
      Begin VB.Frame Fra_Approval 
         Caption         =   "Client Approval"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   1590
         Left            =   5850
         TabIndex        =   39
         ToolTipText     =   "Double Click to Approve"
         Top             =   5775
         Width           =   3150
         Begin VB.Label Lbl_Date 
            Alignment       =   1  'Right Justify
            BackColor       =   &H8000000D&
            BackStyle       =   0  'Transparent
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   315
            Left            =   240
            TabIndex        =   42
            Top             =   960
            Width           =   1335
         End
         Begin VB.Label Lbl_APP 
            Alignment       =   2  'Center
            BackColor       =   &H8000000D&
            BackStyle       =   0  'Transparent
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   11.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   315
            Left            =   240
            TabIndex        =   41
            Top             =   405
            Width           =   2670
         End
         Begin VB.Label Lbl_Time 
            BackColor       =   &H8000000D&
            BackStyle       =   0  'Transparent
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   315
            Left            =   1650
            TabIndex        =   40
            Top             =   960
            Width           =   1260
         End
      End
      Begin VB.Frame Frame3 
         Caption         =   "Remarks"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1590
         Left            =   105
         TabIndex        =   37
         Top             =   5775
         Width           =   5610
         Begin VB.TextBox Txt_Remarks 
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
            Height          =   1170
            Left            =   120
            MaxLength       =   50
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   38
            Top             =   285
            Width           =   5355
         End
      End
      Begin VB.Frame Frame2 
         Height          =   3870
         Left            =   105
         TabIndex        =   32
         Top             =   1830
         Width           =   8925
         Begin VB.TextBox Txt_Job_No 
            Height          =   285
            Left            =   270
            MaxLength       =   3
            TabIndex        =   35
            Top             =   630
            Visible         =   0   'False
            Width           =   1080
         End
         Begin VB.TextBox Txt_Editing 
            Alignment       =   1  'Right Justify
            Height          =   285
            Left            =   270
            TabIndex        =   34
            Top             =   315
            Visible         =   0   'False
            Width           =   1035
         End
         Begin VB.ComboBox Cbo_Month_MQ 
            Appearance      =   0  'Flat
            Height          =   315
            Left            =   3240
            Style           =   2  'Dropdown List
            TabIndex        =   33
            Top             =   360
            Visible         =   0   'False
            Width           =   1215
         End
         Begin MSHierarchicalFlexGridLib.MSHFlexGrid Flex_Quot 
            Height          =   3270
            Left            =   240
            TabIndex        =   36
            Top             =   240
            Width           =   8505
            _ExtentX        =   15002
            _ExtentY        =   5768
            _Version        =   393216
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
      Begin VB.Frame Frame11 
         Height          =   1725
         Left            =   9855
         TabIndex        =   25
         Top             =   15
         Width           =   4005
         Begin VB.TextBox Txt_Enterd_By 
            Appearance      =   0  'Flat
            Enabled         =   0   'False
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
            Left            =   1245
            TabIndex        =   27
            Top             =   225
            Width           =   2550
         End
         Begin VB.TextBox Txt_Plan_No 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0FFC0&
            Enabled         =   0   'False
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
            Left            =   1245
            MaxLength       =   14
            TabIndex        =   26
            Top             =   1200
            Width           =   2550
         End
         Begin MSComCtl2.DTPicker DT_Date 
            Height          =   315
            Left            =   1245
            TabIndex        =   28
            Top             =   705
            Width           =   2580
            _ExtentX        =   4551
            _ExtentY        =   556
            _Version        =   393216
            Enabled         =   0   'False
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Format          =   240058369
            CurrentDate     =   36805
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Entered By"
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
            Left            =   300
            TabIndex        =   31
            Top             =   285
            Width           =   795
         End
         Begin VB.Label Label9 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Date"
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
            Left            =   300
            TabIndex        =   30
            Top             =   750
            Width           =   345
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Plan No."
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
            Left            =   300
            TabIndex        =   29
            Top             =   1260
            Width           =   600
         End
      End
      Begin VB.Frame Frame1 
         Height          =   1725
         Left            =   105
         TabIndex        =   13
         Top             =   15
         Width           =   8925
         Begin VB.TextBox Txt_MQ 
            Appearance      =   0  'Flat
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Left            =   1575
            TabIndex        =   20
            Top             =   720
            Visible         =   0   'False
            Width           =   1770
         End
         Begin VB.ComboBox Cbo_IB_ID 
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
            Left            =   3585
            Style           =   2  'Dropdown List
            TabIndex        =   19
            Top             =   1200
            Visible         =   0   'False
            Width           =   1785
         End
         Begin VB.TextBox Txt_CB_ID 
            Appearance      =   0  'Flat
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Left            =   1575
            TabIndex        =   18
            Top             =   1185
            Visible         =   0   'False
            Width           =   1770
         End
         Begin VB.ComboBox Cbo_MQ 
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
            Left            =   1575
            Style           =   2  'Dropdown List
            TabIndex        =   17
            Top             =   735
            Width           =   1785
         End
         Begin VB.ComboBox Cbo_Brand 
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
            Left            =   1575
            Style           =   2  'Dropdown List
            TabIndex        =   16
            Top             =   285
            Width           =   3795
         End
         Begin VB.ComboBox Cbo_Year 
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
            Left            =   4320
            Style           =   2  'Dropdown List
            TabIndex        =   15
            Top             =   735
            Width           =   1050
         End
         Begin VB.ComboBox Cbo_CB 
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
            Left            =   1575
            Style           =   2  'Dropdown List
            TabIndex        =   14
            Top             =   1185
            Visible         =   0   'False
            Width           =   1785
         End
         Begin VB.Label Label6 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "&Brand"
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
            Left            =   300
            TabIndex        =   24
            Top             =   345
            Width           =   420
         End
         Begin VB.Label Label3 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Client Brief Id"
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
            Left            =   300
            TabIndex        =   23
            Top             =   1245
            Visible         =   0   'False
            Width           =   975
         End
         Begin VB.Label Label8 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "&Year"
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
            Index           =   0
            Left            =   3840
            TabIndex        =   22
            Top             =   765
            Width           =   330
         End
         Begin VB.Label Label8 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "&Meda Quot. No."
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
            Index           =   1
            Left            =   300
            TabIndex        =   21
            Top             =   780
            Width           =   1155
         End
      End
      Begin Crystal.CrystalReport CR 
         Left            =   0
         Top             =   0
         _ExtentX        =   741
         _ExtentY        =   741
         _Version        =   348160
         PrintFileLinesPerPage=   60
      End
   End
   Begin VB.PictureBox picToolbar 
      Align           =   1  'Align Top
      Appearance      =   0  'Flat
      BackColor       =   &H8000000D&
      BorderStyle     =   0  'None
      ClipControls    =   0   'False
      ForeColor       =   &H80000008&
      Height          =   750
      Left            =   0
      ScaleHeight     =   750
      ScaleWidth      =   13935
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   0
      Width           =   13935
      Begin VB.PictureBox picButton 
         Appearance      =   0  'Flat
         BackColor       =   &H00F0F0F0&
         BorderStyle     =   0  'None
         ClipControls    =   0   'False
         ForeColor       =   &H80000008&
         Height          =   750
         Index           =   66
         Left            =   4680
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
         Index           =   67
         Left            =   6210
         ScaleHeight     =   750
         ScaleWidth      =   1500
         TabIndex        =   10
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
         Index           =   5
         Left            =   1620
         ScaleHeight     =   750
         ScaleWidth      =   1500
         TabIndex        =   9
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
         Index           =   6
         Left            =   3150
         ScaleHeight     =   750
         ScaleWidth      =   1500
         TabIndex        =   8
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
         Index           =   23
         Left            =   12330
         ScaleHeight     =   750
         ScaleWidth      =   1500
         TabIndex        =   7
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
         Left            =   13860
         ScaleHeight     =   750
         ScaleWidth      =   1500
         TabIndex        =   6
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
         Index           =   11
         Left            =   15390
         ScaleHeight     =   750
         ScaleWidth      =   1500
         TabIndex        =   5
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
         TabIndex        =   4
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
         Index           =   68
         Left            =   7740
         ScaleHeight     =   750
         ScaleWidth      =   1500
         TabIndex        =   3
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
         Index           =   64
         Left            =   10800
         ScaleHeight     =   750
         ScaleWidth      =   1500
         TabIndex        =   2
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
         Index           =   8
         Left            =   9270
         ScaleHeight     =   750
         ScaleWidth      =   1500
         TabIndex        =   1
         TabStop         =   0   'False
         Top             =   0
         Width           =   1500
      End
   End
End
Attribute VB_Name = "Frm_Radio_Media_Quot"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
''*************************************************************
'Nama Form          : Frm_Radio_Media_Quot
'Fungsi Form        : entry Quotation Radio
'Programer          : joko
'created date       : 10/apr/01
'Last Update/By    : - 31/Dec/2003 /  - Implementor bisa add yg lain view only
'*************************************************************
Option Explicit

Const Not_Approved = "NOT APPROVED"
Const Approved = "APPROVED"

Dim Rs_Media_Quotation As New ADODB.Recordset

Public IB_ID_1 As String
Public IB_ID_2 As String
Public IB_ID_3 As String

Dim Month_Ib As Integer
Dim Year_Ib As Integer

Private Type Code_Rad
    Media_Induk As String * 3
    Media_Agency As String * 3
End Type

Dim Media_Radio_Code As Code_Rad

Dim Add_flag As Boolean
Dim Sukses_flag As Boolean

Dim strSql As String

Private Sub cbo_brand_Click()
    If Cbo_Brand.ListIndex <> -1 Then
        
        Call Initial_Grid
        
        'Call Get_Brand_Info
        Call Load_IB_ID
        
        Button_Lock False
        
        Cmd_Edit.Enabled = False
        Cmd_Delete.Enabled = False
                
        Txt_CB_ID.Text = ""
        Cmd_History_Revision.Enabled = False
        Call setButtonHistory(False, picButton)
        
    End If
End Sub

Private Sub Cbo_IB_Id_Click()
    Dim rs As New ADODB.Recordset
    Dim TxtSQl As String
    
    If Cbo_IB_ID.ListIndex <> -1 Then
        If Rs_Media_Quotation.State = adStateOpen Then
            Rs_Media_Quotation.Close
        End If
    
        Rs_Media_Quotation.Open "Select * From IB_Radio_Quot Where Left(Client_Brief_Id,4)='" & Left(Cbo_Brand.Text, 4) & "'", ConnERP, adOpenKeyset, adLockOptimistic, adCmdText
    
        Rem Get Cleint_Brief_ID
        TxtSQl = "Select * from IB_Radio where ib_id ='" & Cbo_IB_ID.Text & "'"
        rs.Open TxtSQl, ConnERP, adOpenStatic, adLockReadOnly
        With rs
            If .EOF = False Then
                Txt_CB_ID.Text = .Fields("Client_Brief_Id")
                Month_Ib = .Fields("Month")
                Year_Ib = .Fields("Year")
            End If
        End With
        Set rs = Nothing
        
        TxtSQl = "select * from IB_Radio_Quotation_Detail_Revision where ib_id='" & Cbo_IB_ID.Text & "'"
        rs.Open TxtSQl, ConnERP, adOpenStatic, adLockReadOnly
        With rs
            Cmd_History_Revision.Enabled = Not .EOF
        End With
        Set rs = Nothing
        
        Call Clear_Form
        Call Initial_Grid
        Call show_data
        
        If Cmd_History_Revision.Enabled = True Then
            MsgBox "There is History Revision for current Media Quotation", vbInformation, strLogin_User
        End If
        If Lbl_APP.Caption = Not_Approved Then
            Cmd_Delete.Enabled = True
        Else
            Cmd_Delete.Enabled = False
        End If
        
    End If
End Sub

Private Sub Load_Brand()
    Dim TxtSQl As String
    Dim rs As New ADODB.Recordset
    
    TxtSQl = "SELECT * FROM brand inner join Client on client.client_code = brand.client_code "
    TxtSQl = TxtSQl & " WHERE brand_code IN (SELECT brand_code FROM Media_Security_Catalog WHERE User_name='" & strLogin_User & "' "
    TxtSQl = TxtSQl & " AND position IN ('Implementor','Buyer', 'Admin', 'Supervisor', 'Planner', 'Administrator') and Valid_until > getdate()) and client.special_client_Flag =1"
    rs.Open TxtSQl, ConnERP, adOpenStatic, adLockReadOnly
    
    With rs
        
        If .EOF = False Then
            Do While .EOF = False
                Cbo_Brand.AddItem .Fields("Brand_Code") & " - " & .Fields("Brand_Name")
                .MoveNext
            Loop
            'Select 1st Brand
            If Cbo_Brand.ListCount > 0 Then
                Cbo_Brand.ListIndex = 0
            End If
        Else
            MsgBox "Sorry, you don't have any brand", vbCritical, strApplication_Name
        End If
    End With
    
    Set rs = Nothing
End Sub


Private Sub Cbo_Month_MQ_Click()
    If Cbo_Month_MQ.ListIndex <> -1 Then
        Call Clear_Form
        If Trim(Cbo_Month_MQ) = "-None-" Then
            If Flex_Quot.col = 1 Then
                Call Initial_Grid
                Call Load_Month
            ElseIf Flex_Quot.col = 3 Then
                If Flex_Quot.TextMatrix(0, 3) <> "" Then
                    Cbo_Month_MQ.AddItem Flex_Quot.TextMatrix(0, 3), (Get_Month_Number(Flex_Quot.TextMatrix(0, 3)))
                End If
                Flex_Quot.TextMatrix(0, 3) = "" '"Month"
                Flex_Quot.TextMatrix(1, 3) = "" '"Job Id"
                Flex_Quot.TextMatrix(2, 3) = "" '"Nett Cost"
                Flex_Quot.TextMatrix(3, 3) = "" '"Media Supervision Charges"
                Flex_Quot.TextMatrix(4, 3) = "" '"Bonus Fee"
                Flex_Quot.TextMatrix(5, 3) = "" '"Others"
                Flex_Quot.TextMatrix(6, 3) = "" '"Total Lintas"
                Flex_Quot.TextMatrix(7, 3) = "" '"Job Number Club Agency"
                Flex_Quot.TextMatrix(8, 3) = "" '"Club Agency Media Sptv. Charges"
                Flex_Quot.TextMatrix(9, 3) = "" '"Grand Total"
                Flex_Quot.TextMatrix(11, 3) = "" ' "Budget"
                Flex_Quot.TextMatrix(0, 5) = "" '"Month"
                Flex_Quot.TextMatrix(1, 5) = "" '"Job Id"
                Flex_Quot.TextMatrix(2, 5) = "" '"Nett Cost"
                Flex_Quot.TextMatrix(3, 5) = "" '"Media Supervision Charges"
                Flex_Quot.TextMatrix(4, 5) = "" '"Bonus Fee"
                Flex_Quot.TextMatrix(5, 5) = "" '"Others"
                Flex_Quot.TextMatrix(6, 5) = "" '"Total Lintas"
                Flex_Quot.TextMatrix(7, 5) = "" '"Job Number Club Agency"
                Flex_Quot.TextMatrix(8, 5) = "" '"Club Agency Media Sptv. Charges"
                Flex_Quot.TextMatrix(9, 5) = "" '"Grand Total"
                Flex_Quot.TextMatrix(11, 5) = "" ' "Budget"
            ElseIf Flex_Quot.col = 5 Then
                If Flex_Quot.TextMatrix(0, 5) <> "" Then
                    Cbo_Month_MQ.AddItem Flex_Quot.TextMatrix(0, 5), (Get_Month_Number(Flex_Quot.TextMatrix(0, 5)))
                End If
                Flex_Quot.TextMatrix(0, 5) = "" '"Month"
                Flex_Quot.TextMatrix(1, 5) = "" '"Job Id"
                Flex_Quot.TextMatrix(2, 5) = "" '"Nett Cost"
                Flex_Quot.TextMatrix(3, 5) = "" '"Media Supervision Charges"
                Flex_Quot.TextMatrix(4, 5) = "" '"Bonus Fee"
                Flex_Quot.TextMatrix(5, 5) = "" '"Others"
                Flex_Quot.TextMatrix(6, 5) = "" '"Total Lintas"
                Flex_Quot.TextMatrix(7, 5) = "" '"Job Number Club Agency"
                Flex_Quot.TextMatrix(8, 5) = "" '"Club Agency Media Sptv. Charges"
                Flex_Quot.TextMatrix(9, 5) = "" '"Grand Total"
                Flex_Quot.TextMatrix(11, 5) = "" ' "Budget"
            End If
            Cbo_Month_MQ.Visible = False
        Else
            Frm_Radio_MQ_IB.show vbModal
        End If
    End If
End Sub

Private Sub Cbo_MQ_Click()
    Dim rs As New ADODB.Recordset
    Dim TxtSQl As String
        
    If Cbo_MQ.ListIndex <> -1 Then
        Txt_MQ.Text = Cbo_MQ.Text
        
        If Rs_Media_Quotation.State = adStateOpen Then
            Rs_Media_Quotation.Close
        End If
                    
        TxtSQl = "SELECT * FROM IB_Radio_Quot WHERE IB_ID='" & Cbo_MQ.Text & "'"
        Rs_Media_Quotation.Open TxtSQl, ConnERP, adOpenKeyset, adLockOptimistic, adCmdText
        
        Year_Ib = Val(Cbo_Year.Text)
        
        'Check Revision
        TxtSQl = "SELECT * FROM IB_Radio_Quotation_Detail_Revision WHERE ib_id='" & Cbo_MQ.Text & "'"
        rs.Open TxtSQl, ConnERP, adOpenStatic, adLockReadOnly
        
        With rs
            Cmd_History_Revision.Enabled = Not .EOF
            Call setButtonHistory(Not .EOF, picButton)
        End With
        
        Set rs = Nothing
        'End Check Revision
        
        
        Call Clear_Form
        Call Initial_Grid
        Call show_data
        
    End If
End Sub

Private Sub Cbo_Year_Click()
    If Cbo_Year.ListIndex <> -1 Then
        
        Clear_Form
                
        Call Initial_Grid
        
        Call Load_IB_ID
    End If
End Sub

Private Sub db_add()
    'Check apakah Implementor Brand
    If Not IsValidAccess(strLogin_User, "Implementor", Left(Cbo_Brand.Text, 4)) Then
        MsgBox "Access Denied...", vbCritical, "Access Denied"
        Exit Sub
    End If

    If Cbo_Brand.Text = "" Then
        MsgBox "Brand is empty", vbExclamation, strLogin_User
        Exit Sub
    End If
    
    Add_flag = True
    
    Txt_Enterd_By.Text = strLogin_User
    recDate.Requery
    DT_Date.Value = recDate.Fields(0)
    
    Call Clear_Form 'Clear Form
    
    Call Prepare_Temp
    
    Call Initial_Grid 'Grid setting
    
    'Call Get_New_MQ 'Generate Media Quotation Number
    Txt_MQ.Text = Get_RD_Media_Quotation_No(Left(Cbo_Brand.Text, 4), Cbo_Year.Text)
   
    Call Load_Month
    
    Call Button_Lock(True)
    Call setButtonHistory(True, picButton)
End Sub

Private Sub db_Approved_MQ()
    Frm_Radio_View_MQ_Approve.show 1
End Sub

Private Sub db_Cancel()

    Cbo_Month_MQ.Visible = False
        
    If Add_flag Then
        Call Cancel_MQ_Number
        Add_flag = False
    End If
    
    Call Button_Lock(False)
    
    Call Initial_Grid
    
    Call Load_Month
    
    Call Cbo_MQ_Click
    
    Sukses_flag = False

End Sub

Private Sub db_delete()
    Dim Tanya As String
    Dim Index_Col As Integer
    Dim rs_Cek_Schedule As New ADODB.Recordset
    
    'Check apakah Implementor Brand
    If Not IsValidAccess(strLogin_User, "Implementor", Left(Cbo_Brand.Text, 4)) Then
        MsgBox "Access Denied...", vbCritical, "Access Denied"
        Exit Sub
    End If
    
    'Check Related Data
    If Trim(Flex_Quot.TextMatrix(1, 1)) <> "" Then
        strSql = "SELECT * FROM Montly_Radio_Quotation WHERE job_id = '" & Flex_Quot.TextMatrix(1, 1) & "'"
        rs_Cek_Schedule.Open strSql, ConnERP, adOpenStatic, adLockReadOnly
        If Not rs_Cek_Schedule.EOF And Not rs_Cek_Schedule.BOF Then
            MsgBox "Sorry, Can't Delete this Quotation, Schedule has been created", vbExclamation, strLogin_User
            rs_Cek_Schedule.Close
            Exit Sub
        End If
        rs_Cek_Schedule.Close
    End If
    
    If Trim(Flex_Quot.TextMatrix(1, 3)) <> "" Then
        strSql = "SELECT * FROM Montly_Radio_Quotation WHERE job_id = '" & Flex_Quot.TextMatrix(1, 3) & "'"
        rs_Cek_Schedule.Open strSql, ConnERP, adOpenStatic, adLockReadOnly
        If Not rs_Cek_Schedule.EOF And Not rs_Cek_Schedule.BOF Then
            MsgBox "Sorry, Can't Delete this Quotation, Schedule has been created", vbExclamation, strLogin_User
            rs_Cek_Schedule.Close
            Exit Sub
        End If
        rs_Cek_Schedule.Close
    End If
    
    If Trim(Flex_Quot.TextMatrix(1, 5)) <> "" Then
        strSql = "SELECT * FROM Montly_Radio_Quotation WHERE job_id = '" & Flex_Quot.TextMatrix(1, 5) & "'"
        rs_Cek_Schedule.Open strSql, ConnERP, adOpenStatic, adLockReadOnly
        If Not rs_Cek_Schedule.EOF And Not rs_Cek_Schedule.BOF Then
            MsgBox "Sorry, Can't Delete this Quotation, Schedule has been created", vbExclamation, strLogin_User
            rs_Cek_Schedule.Close
            Exit Sub
        End If
        rs_Cek_Schedule.Close
    End If
    'End Check Related Data
    
    'Confirm Delete
    Tanya = MsgBox("Are you sure, you want to delete this Quotation", vbQuestion + vbYesNo, "Confirm Delete")
    If Tanya = vbYes Then
    
        'Delete Quotation Detail,Revision,Ib_Quot
        Call Delete
        
        'Delete Budget Control
        For Index_Col = 1 To 5 Step 2
            strSql = "DELETE FROM uli_budget_control WHERE job_id='" & Flex_Quot.TextMatrix(1, Index_Col) & "'"
            ConnERP.Execute strSql
        Next
            
        'Refresh Form
        Call cbo_brand_Click
    End If

End Sub

Private Sub db_edit()

    'Check apakah Implementor Brand
    If Not IsValidAccess(strLogin_User, "Implementor", Left(Cbo_Brand.Text, 4)) Then
        MsgBox "Access Denied...", vbCritical, "Access Denied."
        Exit Sub
    End If
    
    If Lbl_APP.Caption <> Approved Then
        Call Prepare_Temp
        Call Button_Lock(True)
        Add_flag = False
        Txt_Enterd_By.Text = strLogin_FullName
        Call EnableObject(True)
    Else
        MsgBox "Can not Edit Approved Implementation Brief Quotation", vbExclamation, strLogin_User
    End If

End Sub

Private Sub db_History_Revision()
    Frm_Radio_Media_Quot_History.show vbModal
End Sub

Private Sub db_print()
    If Cbo_MQ.Text = "" Then
        MsgBox "Please Select a Media Quotation !", vbExclamation, "Missing Information"
        Exit Sub
    End If
    With Frm_Radio_MQ_Print
        Set .What_Cbo_Brand = Frm_Radio_Media_Quot.Cbo_Brand
        Set .What_CBO_IB = Frm_Radio_Media_Quot.Cbo_MQ
    End With
    
    Frm_Radio_MQ_Print.show vbModal
    
End Sub

Private Sub db_Publish_to_Web()
     Dim strSql As String
    
    On Error GoTo errLbl
    
    'Is Media Quotation Opened
    If Trim(Cbo_MQ.Text) = "" Then
        MsgBox "Please Select a Job Quotation !", vbExclamation, "Missing Information"
        Exit Sub
    End If
'Validate by Implenter Only
    If Not IsValidAccess(strLogin_User, "Implementor", Left(Cbo_Brand.Text, 4)) Then
        MsgBox "Access Denied, Media Quotation can only publish by Implementer.", vbCritical, "Access Denied."
        Exit Sub
    End If
    
'If Alredy Published
    If Rs_Media_Quotation.Fields("Is_Upload_to_Web").Value = 1 Then
        MsgBox "Media Quotation Already Published to Web.", vbCritical, "Access Denied"
        Exit Sub
    End If
'If Already Approved
    If Rs_Media_Quotation.Fields("Approval_Client").Value = 1 Then
        MsgBox "Media Quotation Already Approved by Client.", vbCritical, "Access Denied"
        Exit Sub
    End If
    
'Is Media Quotation Has Detail Quotation
    If Flex_Quot.TextMatrix(1, 1) = "" Then
        'MsgBox Flex_Quot.TextMatrix(1, 1)
        MsgBox "Media Quotation doesn't have Detail Quotation.", vbCritical, "Access Denied"
        Exit Sub
    End If
'Set Value Publish to Web, Log
    strSql = "UPDATE IB_Radio_Quot SET "
    strSql = strSql & " Is_Upload_to_Web=1,"
    strSql = strSql & " Upload_to_Web_Date=getdate(),"
    strSql = strSql & " Upload_to_Web_By='" & strLogin_FullName & "'"
    strSql = strSql & " WHERE IB_ID='" & Cbo_MQ.Text & "'"
    
    ConnERP.Execute strSql

    MsgBox "Published to Web done.", vbInformation, "Information"
    
    Rs_Media_Quotation.Requery
    
    show_data

'Populate Table BU1 Media_Quotation

    Exit Sub
errLbl:
    MsgBox "Error : " & Err.Description, vbCritical, "Error"
End Sub

Private Sub db_save()
    Dim Last_Index As Integer
    
    Call Save_Data
    'Edit_Flag = False
    
    If Sukses_flag = True Then
    
        Call Button_Lock(False)
        
        If Cbo_MQ.ListCount > 0 Then
            Last_Index = Cbo_MQ.ListIndex
        End If
        
        If Add_flag = True Then
            Call Load_IB_ID
        End If
        
        If Add_flag = True Then
            Cbo_MQ.ListIndex = Cbo_MQ.ListCount - 1
        Else
            Cbo_MQ.ListIndex = Last_Index
        End If
        
        Add_flag = False
        
    End If
    
     'Show Data
    'show_data
    
End Sub

Private Sub Flex_Quot_Click()
    With Flex_Quot
        If .col > 1 Then
            If .col = 3 Then
                If Trim(.TextMatrix(0, 1)) = "" Then
                    Exit Sub
                End If
            End If
            If .col = 5 Then
                If Trim(.TextMatrix(0, 1)) = "" Or Trim(.TextMatrix(0, 3)) = "" Then
                    Exit Sub
                End If
            End If
        End If
        
        If Flex_Quot.Row = 0 Then
            If .col = 1 Then
            'If .Col = 1 Or .Col = 3 Or .Col = 5 Then
                'Yang boleh hnya Colom 1
                If .Row = 0 Then
                    If Cbo_Month_MQ.Visible = False Then
                        Cbo_Month_MQ.Visible = True
                        Cbo_Month_MQ.Move .Left + .CellLeft, .Top + .CellTop, .CellWidth ', .CellHeight'                            Cbo_Month_MQ.Text = .TextMatrix(.Row, .Col)
                        Cbo_Month_MQ.SetFocus
                    End If
                End If
            End If
        End If
    End With
End Sub

Private Sub Flex_Quot_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        'Hanya Kolom dan Baris Tertentu
        If Flex_Quot.Row = 2 Or Flex_Quot.Row = 4 Or Flex_Quot.Row = 0 Then
            If Flex_Quot.Row = 7 Then
                'Get Club Agency
                If Get_Media_Fee(Left(Cbo_Brand.Text, 4), Get_Month_Number(Flex_Quot.TextMatrix(0, Flex_Quot.col)), Cbo_Year.Text, False, strCLUB_AGENCY) = 0 Then
                'If Brand_Info.Club_Agency_SC = 0 Then
                    Exit Sub
                End If
            End If
            If Flex_Quot.col = 1 Then
            'If Flex_Quot.Col = 1 Or Flex_Quot.Col = 3 Or Flex_Quot.Col = 5 Then
                    'Yang Boleh hanya Kolom 1
                With Flex_Quot
                    'If Any Month
                    If .TextMatrix(0, .col) <> "" Then
                        'If Not Job No
                        If .Row <> 7 Then
                            If Txt_Editing.Visible = False Then Txt_Editing.Visible = True
                            Txt_Editing.Height = .CellHeight
                            Txt_Editing.Move .Left + .CellLeft, .Top + .CellTop, .CellWidth ', .CellHeight
                            Txt_Editing.Text = Format(.TextMatrix(.Row, .col), "####0")
                            Txt_Editing.SetFocus
                        Else
                            If Txt_Job_No.Visible = False Then Txt_Job_No.Visible = True
                            Txt_Job_No.Height = .CellHeight
                            Txt_Job_No.Move .Left + .CellLeft, .Top + .CellTop, .CellWidth ', .CellHeight
                            Txt_Job_No.Text = .TextMatrix(.Row, .col)
                            Txt_Job_No.SetFocus
                        End If
                    End If
                    If .Row = 0 Then
                            If Cbo_Month_MQ.Visible = False Then Cbo_Month_MQ.Visible = True
                            'Cbo_Month_MQ.Height = .CellHeight
                            Cbo_Month_MQ.Move .Left + .CellLeft, .Top + .CellTop, .CellWidth ', .CellHeight
    '                            Cbo_Month_MQ.Text = .TextMatrix(.Row, .Col)
                            Cbo_Month_MQ.SetFocus
                    End If
                End With
            End If
        End If
    End If

End Sub

Private Sub Form_Load()
    Dim Pos_Year As Integer
    
    Call AdjustSizeForm
    Call EnableObject(False)
    
    For Pos_Year = 2002 To 2016
        Cbo_Year.AddItem Pos_Year
    Next Pos_Year
    
    DT_Date.Value = Date
    
    Add_flag = False
    
    Call Load_Code
    Call Load_Brand
    Call Initial_Grid
    Call Button_Lock(False)
    
    'Set First View Button
    Cmd_Edit.Enabled = False
    Cmd_Delete.Enabled = False
    Call setButtonHistory(False, picButton)
    
    recDate.Requery
    Cbo_Year.Text = Year(recDate(0))
    
    'Load Month
    Cbo_Month_MQ.AddItem "-None-"
    Cbo_Month_MQ.AddItem "January"
    Cbo_Month_MQ.AddItem "February"
    Cbo_Month_MQ.AddItem "March"
    Cbo_Month_MQ.AddItem "April"
    Cbo_Month_MQ.AddItem "May"
    Cbo_Month_MQ.AddItem "June"
    Cbo_Month_MQ.AddItem "July"
    Cbo_Month_MQ.AddItem "August"
    Cbo_Month_MQ.AddItem "September"
    Cbo_Month_MQ.AddItem "October"
    Cbo_Month_MQ.AddItem "November"
    Cbo_Month_MQ.AddItem "December"

End Sub

Private Sub Load_IB_ID()
    Dim rs As New ADODB.Recordset
    Dim TxtSQl As String
    
    TxtSQl = "SELECT IB_ID FROM IB_Radio_Quot WHERE left(IB_ID,4) ='" & Left(Cbo_Brand.Text, 4) & "' AND Year =" & Val(Cbo_Year.Text) & " ORDER BY IB_ID ASC "
    
    rs.Open TxtSQl, ConnERP, adOpenStatic, adLockReadOnly
    Cbo_MQ.Clear
    With rs
        Do While .EOF = False
            Cbo_MQ.AddItem Trim(.Fields("IB_ID"))
            .MoveNext
        Loop
    End With
End Sub

Private Sub Get_Brand_Info()
    Dim rs As New ADODB.Recordset
    Dim TxtSQl As String

    TxtSQl = "select * from Brand where brand_code ='" & Left(Cbo_Brand.Text, 4) & "'"
    rs.Open TxtSQl, ConnERP, adOpenStatic, adLockReadOnly

    With rs
        Do While .EOF = False
            Brand_Info.ULI = IIf(Trim(.Fields("Client_Code")) = "ULI", True, False)
            Brand_Info.MSC = IIf(IsNull(.Fields("MSC")) = True, 0, .Fields("MSC")) / 100
            Select Case .Fields("MSC_Nett_Flag")
                Case Is = 1
                    Brand_Info.MSC_Nett_Flag = True
                Case Is = 2
                    Brand_Info.MSC_Nett_Flag = False
                Case Is = 3
                    Brand_Info.MSC_Nett_Flag = False
                Case Is = 4
                    Brand_Info.MSC_Nett_Flag = True
            End Select
            'Brand_Info.MSC_Nett_Flag = IIf(.Fields("MSC_Nett_Flag") = 1, True, False)
            Brand_Info.PSC = IIf(IsNull(.Fields("PSC")) = True, 0, .Fields("PSC")) / 100
            Brand_Info.PSC_Nett_Flag = IIf(.Fields("PSC_Nett_Flag") = 1, True, False)
            Brand_Info.Media_Agency_Bonus = IIf(IsNull(.Fields("Media_Agency_Bonus")) = True, 0, .Fields("Media_Agency_Bonus")) / 100
            Brand_Info.Media_Agency_Bonus_Nett_Flag = IIf(.Fields("Media_Agency_Bonus_Nett_Flag") = 1, True, False)
            Brand_Info.Club_Agency_Flag = IIf(.Fields("Club_Agency_Flag") = 1, True, False)
            Brand_Info.Club_Agency_SC = IIf(IsNull(.Fields("Club_Agency_SC")) = True, 0, .Fields("Club_Agency_SC")) / 100
            .MoveNext
        Loop
    End With
    
    If rs.State = adStateOpen Then
        rs.Close
    End If
    
    Set rs = Nothing

End Sub

Private Sub show_data()

    Dim TxtSQl As String
    Dim rs As New ADODB.Recordset
    Dim Pos_Col As Integer, Pos_Row As Integer
    
    
    'Show Detail Per Month
    TxtSQl = "SELECT * FROM Ib_radio_quotation_detail WHERE IB_ID ='" & Trim(Cbo_MQ.Text) & "'"
    rs.Open TxtSQl, ConnERP, adOpenStatic, adLockReadOnly

    Pos_Col = 0
    Pos_Row = 0
    With Flex_Quot
        Cmd_Edit.Enabled = Not rs.EOF
        Cmd_Delete.Enabled = Not rs.EOF
        Do While rs.EOF = False
            Pos_Col = Pos_Col + 2
            .TextMatrix(0, Pos_Col - 1) = Get_Month_Name(rs.Fields("month"))
            .TextMatrix(1, Pos_Col - 1) = rs.Fields("job_id")
            .TextMatrix(2, Pos_Col - 1) = Format(rs.Fields("Nett_Cost"), "##,##0")
            .TextMatrix(3, Pos_Col - 1) = Format(rs.Fields("Media_Sptv_Charge"), "##,##0")
            .TextMatrix(4, Pos_Col - 1) = Format(rs.Fields("Bonus"), "##,##0")
            .TextMatrix(5, Pos_Col - 1) = Format(rs.Fields("Other_Charge"), "##,##0")
            .TextMatrix(6, Pos_Col - 1) = Format(rs.Fields("Total_Lintas"), "##,##0")
            .TextMatrix(7, Pos_Col - 1) = rs.Fields("Job_Number_Agency")
            .TextMatrix(8, Pos_Col - 1) = Format(rs.Fields("Agency_Charge"), "##,##0")
            .TextMatrix(9, Pos_Col - 1) = Format(rs.Fields("Grand_total"), "##,##0")
            .TextMatrix(11, Pos_Col - 1) = Format(rs.Fields("Budget"), "##,##0")
            If Pos_Col = 2 Then
                IB_ID_1 = IIf(IsNull(rs.Fields("Source_IB")) = True, "", rs.Fields("Source_IB"))
            End If
            If Pos_Col = 4 Then
                IB_ID_2 = IIf(IsNull(rs.Fields("Source_IB")) = True, "", rs.Fields("Source_IB"))
            End If
            If Pos_Col = 6 Then
                IB_ID_3 = IIf(IsNull(rs.Fields("Source_IB")) = True, "", rs.Fields("Source_IB"))
            End If

            rs.MoveNext
        Loop
    End With
    If rs.State = adStateOpen Then
        rs.Close
    End If
    Set rs = Nothing
    'End Show Detail Per Month
    
    'Show Header
    With Rs_Media_Quotation
        Txt_Enterd_By.Text = .Fields("Entered_By")
        DT_Date.Value = .Fields!Date
        Txt_Remarks.Text = IIf(IsNull(.Fields!Remarks) = True, "", .Fields!Remarks)
        Txt_Plan_No.Text = IIf(IsNull(.Fields!Plan_No) = True, "", .Fields!Plan_No)
        If .Fields!Approval_Client = 0 Then
            Lbl_APP.Caption = Not_Approved
            Lbl_Date.Caption = ""
            Lbl_Time.Caption = ""
            Cmd_Edit.Enabled = True
        Else
            Lbl_APP.Caption = Approved
            Lbl_Date.Caption = Format(.Fields!Approved_Date, "mm/dd/yyyy")
            Lbl_Time.Caption = Format(.Fields!Approved_Date, "hh:mm:ss AM/PM")
            Cmd_Edit.Enabled = False
        End If
        
        'Publish to Web
        If .Fields("Is_Upload_to_Web").Value = 1 Then
            Cmd_Pulish_to_Web.Enabled = False
        Else
            Cmd_Pulish_to_Web.Enabled = True
        End If
        
    End With
    
    If Lbl_APP.Caption = Not_Approved Then
        Cmd_Delete.Enabled = True
    Else
        Cmd_Delete.Enabled = False
    End If
    'End Show Header
    
End Sub

Private Sub Prepare_Temp()
    Dim TxtSQl As String
    
    If Rs_Media_Quotation.State = adStateOpen Then
        Rs_Media_Quotation.Close
    End If
    Set Rs_Media_Quotation = Nothing
    
    TxtSQl = "select * from IB_Radio_quot where IB_ID ='" & Cbo_MQ.Text & "'"
    
    'Rs_Media_Quotation.CursorLocation = adUseClient
    Rs_Media_Quotation.Open TxtSQl, ConnERP, adOpenKeyset, adLockOptimistic
    
    'MsgBox Rs_Media_Quotation.RecordCount
End Sub

Private Sub Initial_Grid()
    Dim Index_Row As Integer
    Dim Index_Col As Integer
    
    Flex_Quot.Clear
'Flex_Quot.Cols = 6
'Flex_Quot.Rows = 12
'Flex_Quot.FixedCols = 1

'***********************************
'       Initial Grid
'***********************************
        Flex_Quot.Height = 0
        For Index_Row = 0 To Flex_Quot.Rows - 1
            Flex_Quot.RowHeight(Index_Row) = 290
            Flex_Quot.Height = Flex_Quot.Height + 290
        Next Index_Row
        
        Flex_Quot.ColWidth(0) = 2600
        Flex_Quot.ColWidth(1) = 1600
        Flex_Quot.ColWidth(2) = 400
        Flex_Quot.ColWidth(3) = 1600
        Flex_Quot.ColWidth(4) = 400
        Flex_Quot.ColWidth(5) = 1600
        For Index_Row = 0 To Flex_Quot.Rows - 1
            Flex_Quot.Row = Index_Row
            Flex_Quot.col = 2
            Flex_Quot.CellBackColor = &H8000000F
            Flex_Quot.col = 4
            Flex_Quot.CellBackColor = &H8000000F
        Next Index_Row
        
        For Index_Col = 1 To Flex_Quot.cols - 1
            Flex_Quot.col = Index_Col
            Flex_Quot.Row = 10
            Flex_Quot.CellBackColor = &H8000000F
            Flex_Quot.col = Index_Col
            Flex_Quot.Row = 11
            Flex_Quot.CellFontBold = True
            '2
            If Index_Col = 1 Or Index_Col = 3 Or Index_Col = 5 Then
                Flex_Quot.col = Index_Col
                Flex_Quot.Row = 2
                Flex_Quot.CellBackColor = &HC0FFC0
                '4
                Flex_Quot.col = Index_Col
                Flex_Quot.Row = 4
                Flex_Quot.CellBackColor = &HC0FFC0
            End If
        Next Index_Col
        'Bold Grid
        '&H00C0FFC0&
        Flex_Quot.col = 0
        Flex_Quot.Row = 0
        Flex_Quot.CellFontBold = True
        Flex_Quot.col = 0
        Flex_Quot.Row = 1
        Flex_Quot.CellFontBold = True
        Flex_Quot.col = 0
        Flex_Quot.Row = 6
        Flex_Quot.CellFontBold = True
        Flex_Quot.col = 0
        Flex_Quot.Row = 9
        Flex_Quot.CellFontBold = True
        
        Flex_Quot.col = 1
        Flex_Quot.Row = 0
        Flex_Quot.CellFontBold = True
        Flex_Quot.col = 1
        Flex_Quot.Row = 1
        Flex_Quot.CellFontBold = True
        Flex_Quot.col = 1
        Flex_Quot.Row = 6
        Flex_Quot.CellFontBold = True
        Flex_Quot.col = 1
        Flex_Quot.Row = 9
        Flex_Quot.CellFontBold = True
        
        Flex_Quot.col = 3
        Flex_Quot.Row = 0
        Flex_Quot.CellFontBold = True
        Flex_Quot.col = 3
        Flex_Quot.Row = 1
        Flex_Quot.CellFontBold = True
        Flex_Quot.col = 3
        Flex_Quot.Row = 6
        Flex_Quot.CellFontBold = True
        Flex_Quot.col = 3
        Flex_Quot.Row = 9
        Flex_Quot.CellFontBold = True
        
        Flex_Quot.col = 5
        Flex_Quot.Row = 0
        Flex_Quot.CellFontBold = True
        Flex_Quot.col = 5
        Flex_Quot.Row = 1
        Flex_Quot.CellFontBold = True
        Flex_Quot.col = 5
        Flex_Quot.Row = 6
        Flex_Quot.CellFontBold = True
        Flex_Quot.col = 5
        Flex_Quot.Row = 9
        Flex_Quot.CellFontBold = True
        Flex_Quot.col = 0
        Flex_Quot.Row = 11
        Flex_Quot.CellFontBold = True
        
        Flex_Quot.TextMatrix(0, 0) = "Month"
        Flex_Quot.TextMatrix(1, 0) = "Job Id"
        Flex_Quot.TextMatrix(2, 0) = "Nett Cost"
        Flex_Quot.TextMatrix(3, 0) = "Media Supervision Charges"
        Flex_Quot.TextMatrix(4, 0) = "Bonus Fee"
        Flex_Quot.TextMatrix(5, 0) = "Others"
        Flex_Quot.TextMatrix(6, 0) = "Total Lintas"
        Flex_Quot.TextMatrix(7, 0) = "Job Number Club Agency"
        Flex_Quot.TextMatrix(8, 0) = "Club Agency Media Sptv. Charges"
        Flex_Quot.TextMatrix(9, 0) = "Grand Total"
        Flex_Quot.TextMatrix(11, 0) = "Budget"
End Sub


Private Sub Load_Plan()
    Dim rs As New ADODB.Recordset
    Dim TxtSQl As String
    Dim Pos_Col As Integer
    Dim RS_Code As New ADODB.Recordset
    Dim Radio_Code As String
    
    TxtSQl = "select * from Media_Type where Media_Type_Name ='Radio Media Induk'"
    RS_Code.Open TxtSQl, ConnERP, adOpenStatic, adLockReadOnly
    With RS_Code
        If .EOF = False Then
            Radio_Code = .Fields(0)
        End If
    End With
    Set RS_Code = Nothing
    
    TxtSQl = " select * from IB_Radio_Plan where ib_id ='" & Cbo_IB_ID.Text & "' order by month asc"
    rs.Open TxtSQl, ConnERP, adOpenStatic, adLockReadOnly
    ''Debug.Print TxtSQL
    With rs
        Do While .EOF = False
            Pos_Col = Pos_Col + 2
            Flex_Quot.TextMatrix(0, Pos_Col - 1) = Get_Month_Name(.Fields("month"))
            
            Rem Get New Job_ID
            Dim New_Job As String
            Dim New_Month As String
            
            If Len(.Fields("month")) = 1 Then
                New_Month = Trim("0" & Trim(str(.Fields("month"))))
            Else
                New_Month = .Fields("month")
            End If
            
            New_Job = Left(Txt_CB_ID.Text, 4) & "." & Trim(Radio_Code) & "." & Right(Year_Ib, 2) & New_Month
            Flex_Quot.TextMatrix(1, Pos_Col - 1) = New_Job
            Flex_Quot.TextMatrix(11, Pos_Col - 1) = Format(.Fields("Budget"), "#,##0")
            .MoveNext
        Loop
    End With
End Sub



Private Sub Form_Unload(Cancel As Integer)
    If Rs_Media_Quotation.State = adStateOpen Then
        Rs_Media_Quotation.Close
        Set Rs_Media_Quotation = Nothing
    End If
    
End Sub

Private Sub Fra_Approval_DblClick()
    Dim TxtSQl As String
    Dim strSql As String
    Dim rs As New ADODB.Recordset
    Dim RS_Detail_Old As New ADODB.Recordset
    Dim RS_Detail_New As New ADODB.Recordset
    Dim RS_Detail_Add As New ADODB.Recordset
    
    'Check apakah Implementor Brand
    If Not IsValidAccess(strLogin_User, "Implementor", Left(Cbo_Brand.Text, 4)) Then
        MsgBox "Access Denied...", vbCritical, "Access Denied"
        Exit Sub
    End If
    
    If Lbl_APP.Caption = Not_Approved Then
            If Cbo_MQ.ListIndex <> -1 Then
                Set Frm_IB_Radio_Approve.What_Approval = Frm_Radio_Media_Quot
                
                Frm_IB_Radio_Approve.show vbModal
                
                If Lbl_APP.Caption <> Not_Approved Then
                    TxtSQl = "UPDATE Ib_Radio_Quot SET Approval_Client =1, Approved_Date = '" & CDate(Lbl_Date.Caption) & "' WHERE Ib_ID ='" & Cbo_MQ.Text & "'"
                    ConnERP.Execute TxtSQl
                    
                    '********************
                    'Bandingkan IB Bugdet
                    TxtSQl = "SELECT * FROM IB_Radio_Quotation_Detail WHERE ib_id = '" & Trim(Txt_MQ.Text) & "'"
                    RS_Detail_New.Open TxtSQl, ConnERP, adOpenDynamic, adLockOptimistic
                    
                    Do While RS_Detail_New.EOF = False
                             
                            'Delete Previouse Approved Media Quotation
                            TxtSQl = "DELETE FROM IB_Radio_Quotation_Detail_approved WHERE job_id = '" & Trim(RS_Detail_New.Fields("job_id").Value) & "' "
                            ConnERP.Execute TxtSQl
                            
                            'Insert New Approved Media Quotation
                            strSql = ""
                            strSql = "INSERT INTO IB_Radio_Quotation_Detail_Approved VALUES('"
                            strSql = strSql & RS_Detail_New.Fields("Client_Brief_Id").Value & "','"
                            strSql = strSql & RS_Detail_New.Fields("IB_Id").Value & "','"
                            strSql = strSql & RS_Detail_New.Fields("Job_Id").Value & "',"
                            strSql = strSql & RS_Detail_New.Fields("Month").Value & ","
                            strSql = strSql & RS_Detail_New.Fields("Year").Value & ","
                            strSql = strSql & RS_Detail_New.Fields("Nett_Cost").Value & ","
                            strSql = strSql & RS_Detail_New.Fields("Media_Sptv_Charge").Value & ","
                            strSql = strSql & RS_Detail_New.Fields("Other_Charge").Value & ","
                            strSql = strSql & RS_Detail_New.Fields("Bonus").Value & ","
                            strSql = strSql & RS_Detail_New.Fields("Total_Lintas").Value & ","
                            strSql = strSql & RS_Detail_New.Fields("Agency_Charge").Value & ",'"
                            strSql = strSql & RS_Detail_New.Fields("Job_Number_Agency").Value & "',"
                            strSql = strSql & RS_Detail_New.Fields("Grand_Total").Value & ", "
                            strSql = strSql & RS_Detail_New.Fields("Budget").Value & ", "
                            strSql = strSql & "'" & RS_Detail_New.Fields("Source_IB").Value & "'"
                            strSql = strSql & " )"
                            ConnERP.Execute strSql
                        'End If
                        
                        RS_Detail_New.MoveNext
                    Loop
                End If
            End If
        
    End If
End Sub

Private Sub Txt_Job_No_GotFocus()
    If Txt_Job_No.Text <> "" Then
        SendKeys "{Home}+{End}"
    End If
End Sub

Private Sub Txt_Job_No_KeyPress(KeyAscii As Integer)
    If (Chr(KeyAscii) < "0" Or Chr(KeyAscii) > "9") And KeyAscii <> 8 And KeyAscii <> 13 And KeyAscii <> 9 And Chr(KeyAscii) <> "." Then
        'MsgBox KeyAscii
        KeyAscii = 0
        Beep
        Exit Sub
    End If
    If KeyAscii = 13 Or KeyAscii = 9 Then
        If Flex_Quot.col > 0 Then
            If Txt_Job_No.Text = "" Then
                Flex_Quot.TextMatrix(Flex_Quot.Row, Flex_Quot.col) = ""
            Else
                Flex_Quot.TextMatrix(Flex_Quot.Row, Flex_Quot.col) = Left(Cbo_Brand.Text, 4) & "." & Media_Radio_Code.Media_Agency & "." & Txt_Job_No.Text
            End If
            Txt_Job_No.Text = ""
            Txt_Job_No.Visible = False
        End If
    End If
End Sub

Private Sub Txt_Job_No_LostFocus()
    Txt_Job_No.Visible = False
    Flex_Quot.SetFocus
End Sub

Private Sub Load_Code()
    Dim rs As New ADODB.Recordset
    Dim TxtSQl As String
    
    TxtSQl = "SELECT * FROM Media_Type WHERE Media_Type_Name ='Radio Media Induk'"
    rs.Open TxtSQl, ConnERP, adOpenStatic, adLockReadOnly
    With rs
        If .EOF = False Then
            Media_Radio_Code.Media_Induk = .Fields(0)
        End If
    End With
    Set rs = Nothing
    
    TxtSQl = "SELECT * FROM Media_Type WHERE Media_Type_Name ='Radio Club Agency'"
    rs.Open TxtSQl, ConnERP, adOpenStatic, adLockReadOnly
    With rs
        If .EOF = False Then
            Media_Radio_Code.Media_Agency = .Fields(0)
        End If
    End With

End Sub

Private Sub Txt_Editing_GotFocus()
    If Txt_Editing.Text <> "" Then
        SendKeys "{Home}+{End}"
    End If
End Sub

Private Sub Txt_Editing_KeyPress(KeyAscii As Integer)
    If Chr(KeyAscii) = "m" Or Chr(KeyAscii) = "t" Or Chr(KeyAscii) = "b" Then
        Txt_Editing.Text = Txt_Editing.Text & Chr(KeyAscii)
        Call Converting_Money(Txt_Editing)
    End If
        If (Chr(KeyAscii) < "0" Or Chr(KeyAscii) > "9") And KeyAscii <> 8 And KeyAscii <> 13 And KeyAscii <> 9 And KeyAscii <> 46 Then
            'MsgBox KeyAscii
            KeyAscii = 0
            Beep
            Exit Sub
        End If
        
    If KeyAscii = 13 Or KeyAscii = 9 Then
        If Flex_Quot.col > 0 Then
            If Val(Txt_Editing.Text) = 0 Then
                Flex_Quot.TextMatrix(Flex_Quot.Row, Flex_Quot.col) = 0
            Else
                Flex_Quot.TextMatrix(Flex_Quot.Row, Flex_Quot.col) = Format(Val(Txt_Editing.Text), "##0,0")
            End If
            Txt_Editing.Text = ""
            Txt_Editing.Visible = False
        End If
    End If
    
    Cmd_Save.Enabled = True


End Sub

Private Sub Txt_Editing_LostFocus()
    Dim Nett_Total As Currency
    Dim MSC As Currency
    Dim Bonus_Fee As Currency
    Dim Others As Currency
    Dim Year_Month As String
    Dim Total_Lintas As Currency
    Dim Club_MSC As Currency
    Dim Grand_Total As Currency
    
    Txt_Editing.Visible = False
    
    Flex_Quot.SetFocus
    
    'Calculate
    '================================================
    
    'Get Fee
    '----------
    'Select Fee Untuk Bulan Yang Di Proses ----?
    'Paid Fee
    'Bonus Fee (Tidak Dihitung Dari Presesntasi)
    'Club Agency
        
    'Nett Total
    If Flex_Quot.TextMatrix(2, Flex_Quot.col) = 0 Then
        Nett_Total = 0
    Else
        Nett_Total = Flex_Quot.TextMatrix(2, Flex_Quot.col)
    End If
          
    'Media Supv. Charge
    MSC = Get_Media_Fee(Left(Cbo_Brand.Text, 4), Get_Month_Number(Flex_Quot.TextMatrix(0, Flex_Quot.col)), Cbo_Year.Text, False, StrMSC_PAID)
    MSC = (MSC / 100) * Nett_Total
    Flex_Quot.TextMatrix(3, Flex_Quot.col) = Format(MSC, "##,##0")
    
    'Total Lintas
    If Val(Flex_Quot.TextMatrix(4, Flex_Quot.col)) = 0 Then
        Bonus_Fee = 0
        Flex_Quot.TextMatrix(4, Flex_Quot.col) = 0
    Else
        Bonus_Fee = CCur(Format(Flex_Quot.TextMatrix(4, Flex_Quot.col), "####0"))
    End If
    
    'Others
    Flex_Quot.TextMatrix(5, Flex_Quot.col) = 0
        
    'Club Agenccy MSC
    'Club_MSC = (Brand_Info.Club_Agency_SC) * Nett_Total
    Club_MSC = Get_Media_Fee(Left(Cbo_Brand.Text, 4), Get_Month_Number(Flex_Quot.TextMatrix(0, Flex_Quot.col)), Cbo_Year.Text, False, strCLUB_AGENCY)
    Club_MSC = (Club_MSC / 100) * Nett_Total
    
    If Club_MSC <> 0 Then
        'Job No Club Agency
        Year_Month = Right(Flex_Quot.TextMatrix(1, Flex_Quot.col), 3)
        Flex_Quot.TextMatrix(7, Flex_Quot.col) = Left(Cbo_Brand.Text, 4) & "." & Media_Radio_Code.Media_Agency & "." & Year_Month
        'Club MSC
        Flex_Quot.TextMatrix(8, Flex_Quot.col) = Format(Club_MSC, "##,##0")
    Else
        Flex_Quot.TextMatrix(8, Flex_Quot.col) = 0
    End If
        
    'Cek Nett
    If Nett_Total = 0 Then
        Exit Sub
    End If
    
    'Grand Total
    Total_Lintas = Nett_Total + MSC + Bonus_Fee + Others
    Grand_Total = Total_Lintas + Club_MSC
    
    'Sisa Tambah Ke Others
    If Grand_Total < Val(Format(Flex_Quot.TextMatrix(11, Flex_Quot.col), "####0")) Then
        Others = Val(Format(Flex_Quot.TextMatrix(11, Flex_Quot.col), "####0")) - Grand_Total
        Flex_Quot.TextMatrix(5, Flex_Quot.col) = Format(Others, "##,##0")
        Total_Lintas = Nett_Total + MSC + Bonus_Fee + Others
        Grand_Total = Total_Lintas + Club_MSC
    End If
    Flex_Quot.TextMatrix(6, Flex_Quot.col) = Format(Total_Lintas, "##,##0")
    Flex_Quot.TextMatrix(9, Flex_Quot.col) = Format(Grand_Total, "##,##0")
    
    Rem Month 1st
    If Val(Format(Flex_Quot.TextMatrix(9, 1), "####0")) > Val(Format(Flex_Quot.TextMatrix(11, 1), "####0")) Then
        Flex_Quot.col = 1
        Flex_Quot.Row = 9
        Flex_Quot.CellBackColor = vbRed
        MsgBox "Grand Total For " & Flex_Quot.TextMatrix(0, 1) & " Higher than Budget", vbCritical, strLogin_User
        Cmd_Save.Enabled = False
        Flex_Quot.col = 1
        Flex_Quot.Row = 9
        Flex_Quot.CellBackColor = vbWhite
    End If
    
    Rem Month 2nd
    If Val(Format(Flex_Quot.TextMatrix(9, 3), "####0")) > Val(Format(Flex_Quot.TextMatrix(11, 3), "####0")) Then
        Flex_Quot.col = 3
        Flex_Quot.Row = 9
        Flex_Quot.CellBackColor = vbRed
        MsgBox "Grand Total For " & Flex_Quot.TextMatrix(0, 3) & " Higher than Budget", vbCritical, strLogin_User
            Flex_Quot.col = 3
            Flex_Quot.Row = 9
            Flex_Quot.CellBackColor = vbWhite
        Cmd_Save.Enabled = False
    End If
    
    Rem Month 3rd
    If Val(Format(Flex_Quot.TextMatrix(9, 5), "####0")) > Val(Format(Flex_Quot.TextMatrix(11, 5), "####0")) Then
        Flex_Quot.col = 5
        Flex_Quot.Row = 9
        Flex_Quot.CellBackColor = vbRed
        MsgBox "Grand Total For " & Flex_Quot.TextMatrix(0, 5) & " Higher than Budget", vbCritical, strLogin_User
            Flex_Quot.col = 5
            Flex_Quot.Row = 9
            Flex_Quot.CellBackColor = vbWhite
        Cmd_Save.Enabled = False
    End If

    
End Sub

Private Sub Button_Lock(Enable As Boolean)
    Cbo_Year.Enabled = Not Enable
    Flex_Quot.Enabled = Enable
    Cbo_Brand.Enabled = Not Enable
    Cbo_MQ.Enabled = Not Enable
    Cmd_Add.Visible = Not Enable
    Cmd_Edit.Visible = Not Enable
    Cmd_Delete.Enabled = Not Enable
    Cmd_Close.Enabled = Not Enable
    Cmd_Print.Enabled = Not Enable
    Cmd_Next.Enabled = Not Enable
    Cmd_Previous.Enabled = Not Enable
    Cmd_Last.Enabled = Not Enable
    Cmd_First.Enabled = Not Enable
    Txt_Remarks.Enabled = Enable
    Cmd_Save.Enabled = Enable
    Cmd_Cancel.Enabled = Enable
    Txt_MQ.Visible = Enable
    Cbo_MQ.Visible = Not Enable
    Txt_Plan_No.Enabled = Enable
    'Cmd_Approved_MQ.Enabled = Not Enable
    Cmd_Pulish_to_Web.Enabled = Not Enable
    Call EnableObject(Enable)

End Sub

Private Sub Clear_Form()
    Lbl_APP.ForeColor = vbRed
    Lbl_APP.Caption = Not_Approved
    Lbl_Date.Caption = ""
    Lbl_Time.Caption = ""
    Txt_Remarks.Text = ""
    Txt_Enterd_By.Text = strLogin_FullName
End Sub

Private Sub Save_Data()
    Dim Answare As Integer
    Dim strSql As String
    Dim Nett_Total As Currency
    Dim MSC As Currency
    Dim Bonus_Fee As Currency
    Dim Others As Currency
    Dim Revision As Integer
    Dim Total_Lintas As Currency
    Dim Club_MSC As Currency
    Dim Grand_Total As Currency
    Dim Str_Message As String
    Dim Any_Message As Boolean
    Dim Add_flag_rev As Boolean
    Dim Index_Col As Integer
    Dim Index_Row As Integer
       
    On Error GoTo Label_Err
    
    ConnERP.BeginTrans
    
    Add_flag_rev = False
    
    
    If Add_flag Then
        Any_Message = False
        Str_Message = "There is Revision Media Quotation for "            'On Error GoTo label_err
        
            '*********************************************
            ' Cek untuk Revisi
            '*********************************************
            
            Dim Rs_Detail_Revision As New ADODB.Recordset
            Dim StrSQL_Revision As String
            
            'Generate Where
            StrSQL_Revision = "SELECT * FROM IB_Radio_Quotation_Detail WHERE JOb_ID='" & Flex_Quot.TextMatrix(1, 1) & "'"
            If Flex_Quot.TextMatrix(1, 3) <> "" Then
                StrSQL_Revision = StrSQL_Revision & " OR Job_ID='" & Flex_Quot.TextMatrix(1, 3) & "'"
            End If
            If Flex_Quot.TextMatrix(1, 5) <> "" Then
                StrSQL_Revision = StrSQL_Revision & " OR Job_ID='" & Flex_Quot.TextMatrix(1, 5) & "'"
            End If
            
            Rs_Detail_Revision.Open StrSQL_Revision, ConnERP, adOpenKeyset, adLockOptimistic, adCmdText
            
            Do While Not Rs_Detail_Revision.EOF
                'Transfer ke Revision Table
                '**************************
                'Generate Revision Value
                Dim Str_Max As String
                Dim Rs_Revision As New ADODB.Recordset
                
                Str_Max = "SELECT MAX(Revision) as LastRevision FROM IB_Radio_Quotation_Detail_Revision WHERE Job_Id='" & Rs_Detail_Revision.Fields("Job_Id").Value & "'"
                Rs_Revision.Open Str_Max, ConnERP, , , adCmdText
                If Not IsNull(Rs_Revision.Fields("LastRevision").Value) Then
                    Revision = Rs_Revision.Fields("LastRevision").Value + 1
                Else
                    Revision = 1
                End If
                Rs_Revision.Close
                Set Rs_Revision = Nothing
                                
                Rem Jika ada revisi maka Add_Flag =False
                Add_flag_rev = True
                
                '***********
                'Save Data
                strSql = ""
                strSql = "INSERT INTO IB_Radio_Quotation_Detail_Revision VALUES('"
                strSql = strSql & Rs_Detail_Revision.Fields("Client_Brief_Id").Value & "','"
                strSql = strSql & Rs_Detail_Revision.Fields("IB_Id").Value & "','"
                strSql = strSql & Rs_Detail_Revision.Fields("Job_Id").Value & "',"
                strSql = strSql & Revision & ","
                strSql = strSql & Rs_Detail_Revision.Fields("Month").Value & ","
                strSql = strSql & Rs_Detail_Revision.Fields("Year").Value & ","
                strSql = strSql & Rs_Detail_Revision.Fields("Nett_Cost").Value & ","
                strSql = strSql & Rs_Detail_Revision.Fields("Media_Sptv_Charge").Value & ","
                strSql = strSql & Rs_Detail_Revision.Fields("Other_Charge").Value & ","
                strSql = strSql & Rs_Detail_Revision.Fields("Bonus").Value & ","
                strSql = strSql & Rs_Detail_Revision.Fields("Total_Lintas").Value & ","
                strSql = strSql & Rs_Detail_Revision.Fields("Agency_Charge").Value & ",'"
                strSql = strSql & Rs_Detail_Revision.Fields("Job_Number_Agency").Value & "',"
                strSql = strSql & Rs_Detail_Revision.Fields("Grand_Total").Value & ", "
                strSql = strSql & Rs_Detail_Revision.Fields("Budget").Value & ", "
                strSql = strSql & "'" & Rs_Detail_Revision.Fields("Source_IB").Value & "'"
                strSql = strSql & " )"
                ConnERP.Execute strSql
                
                
                'Delete
                '**************************
                Str_Message = Str_Message & Get_Month_Name(Rs_Detail_Revision.Fields("Month").Value) & " "
                Rs_Detail_Revision.Delete
                Rs_Detail_Revision.MoveNext
                Any_Message = True
                
            Loop
            Rs_Detail_Revision.Close
            Set Rs_Detail_Revision = Nothing
            
            'Pesan Ada Revision untuk bulan ?
            If Any_Message Then
                
                'Pesan
                MsgBox Str_Message, vbInformation, strLogin_User
                Any_Message = False
            End If
        End If
    '************************************************
        
        '*********************************************
        'Month 1
        If Flex_Quot.TextMatrix(0, 1) <> "" Then
            'Cek Nett Total
            If Flex_Quot.TextMatrix(2, 1) = "" Then
                MsgBox "Nett Cost For " & Flex_Quot.TextMatrix(0, 1) & " Empty . ", vbCritical, strLogin_User
                ConnERP.RollbackTrans
                Rs_Media_Quotation.Requery
                Exit Sub
            End If
            
            'Cek Job No Club Aggency 'Cek Value Club AG MSC
            If Val(Flex_Quot.TextMatrix(8, 1)) <> 0 Then
                If Trim(Flex_Quot.TextMatrix(7, 1)) = "" Then
                    MsgBox "Please insert Job Number Club Agency !", vbExclamation, strLogin_User
                    ConnERP.RollbackTrans
                    Rs_Media_Quotation.Requery
                    Exit Sub
                End If
            End If
            'Cek Grand Total
                '<
                If Val(Format(Flex_Quot.TextMatrix(9, 1), "####0")) < Val(Format(Flex_Quot.TextMatrix(11, 1), "####0")) Then
                    Answare = MsgBox("Grand Total For " & Flex_Quot.TextMatrix(0, 1) & " Lower than Budget, Do You Want to Save it ?", vbQuestion + vbYesNo, strLogin_User)
                    If Answare = vbNo Then
                        ConnERP.RollbackTrans
                        Rs_Media_Quotation.Requery
                        Exit Sub
                    End If
                    'Save ?
                ElseIf Val(Format(Flex_Quot.TextMatrix(9, 1), "####0")) > Val(Format(Flex_Quot.TextMatrix(11, 1), "####0")) Then
                '>
                    MsgBox "Grand Total For " & Flex_Quot.TextMatrix(0, 1) & " Higher than Budget", vbCritical, strLogin_User
                    ConnERP.RollbackTrans
                    Rs_Media_Quotation.Requery
                    Exit Sub
                End If
                
        End If
        'Month 2
        If Flex_Quot.TextMatrix(0, 3) <> "" Then
            'Cek Nett Total
            If Flex_Quot.TextMatrix(2, 3) = "" Then
                MsgBox "Nett Cost For " & Flex_Quot.TextMatrix(0, 3) & " Empty . ", vbCritical, strLogin_User
                ConnERP.RollbackTrans
                Rs_Media_Quotation.Requery
                Exit Sub
            End If
            'Cek Job No Club Aggency 'Cek Value Club AG MSC
            If Val(Flex_Quot.TextMatrix(8, 3)) <> 0 Then
                If Trim(Flex_Quot.TextMatrix(7, 3)) = "" Then
                    MsgBox "Please insert Job Number Club Agency !", vbExclamation, strLogin_User
                    ConnERP.RollbackTrans
                    Rs_Media_Quotation.Requery
                    Exit Sub
                End If
            End If
            'Cek Grand Total
                '<
                If Val(Format(Flex_Quot.TextMatrix(9, 3), "####0")) < Val(Format(Flex_Quot.TextMatrix(11, 3), "####0")) Then
                    Answare = MsgBox("Grand Total For " & Flex_Quot.TextMatrix(0, 3) & " Lower than Budget, Do You Want to Save it ?", vbQuestion + vbYesNo, strLogin_User)
                    If Answare = vbNo Then
                        ConnERP.RollbackTrans
                        Rs_Media_Quotation.Requery
                        Exit Sub
                    End If
                    'Save ?
                ElseIf Val(Format(Flex_Quot.TextMatrix(9, 3), "####0")) > Val(Format(Flex_Quot.TextMatrix(11, 3), "####0")) Then
                '>
                    MsgBox "Grand Total For " & Flex_Quot.TextMatrix(0, 3) & " Higher than Budget", vbCritical, strLogin_User
                    ConnERP.RollbackTrans
                    Rs_Media_Quotation.Requery
                    Exit Sub
                End If
        
        End If
        'Month 3
        If Flex_Quot.TextMatrix(0, 5) <> "" Then
            'Cek Nett Total
            If Flex_Quot.TextMatrix(2, 5) = "" Then
                MsgBox "Nett Cost For " & Flex_Quot.TextMatrix(0, 5) & " Empty . ", vbCritical, strLogin_User
                ConnERP.RollbackTrans
                Rs_Media_Quotation.Requery
                Exit Sub
            End If
            'Cek Job No Club Aggency 'Cek Value Club AG MSC
            If Val(Flex_Quot.TextMatrix(8, 5)) <> 0 Then
                If Trim(Flex_Quot.TextMatrix(7, 5)) = "" Then
                    MsgBox "Please insert Job Number Club Agency !", vbExclamation, strLogin_User
                    ConnERP.RollbackTrans
                    Rs_Media_Quotation.Requery
                    Exit Sub
                End If
            End If
            'Cek Grand Total
                '<
                If Val(Format(Flex_Quot.TextMatrix(9, 5), "####0")) < Val(Format(Flex_Quot.TextMatrix(11, 5), "####0")) Then
                    Answare = MsgBox("Grand Total For " & Flex_Quot.TextMatrix(0, 5) & " Lower than Budget, Do You Want to Save it ?", vbQuestion + vbYesNo, strLogin_User)
                    If Answare = vbNo Then
                        ConnERP.RollbackTrans
                        Rs_Media_Quotation.Requery
                        Exit Sub
                    End If
                    'Save ?
                ElseIf Val(Format(Flex_Quot.TextMatrix(9, 5), "####0")) > Val(Format(Flex_Quot.TextMatrix(11, 5), "####0")) Then
                '>
                    MsgBox "Grand Total For " & Flex_Quot.TextMatrix(0, 5) & " Higher than Budget", vbCritical, strLogin_User
                    ConnERP.RollbackTrans
                    Rs_Media_Quotation.Requery
                    Exit Sub
                End If
        
        End If
    
    '**********************
    'Assign Value
    '======================
    If Add_flag Then
        Rs_Media_Quotation.AddNew
        Rs_Media_Quotation.Fields("Client_Brief_Id").Value = Txt_CB_ID.Text
        Rs_Media_Quotation.Fields("IB_ID").Value = Trim(Txt_MQ.Text)
        Rs_Media_Quotation.Fields("Month_IB").Value = Month_Ib
        Rs_Media_Quotation.Fields("Year").Value = Val(Cbo_Year.Text)
        Rs_Media_Quotation.Fields("Date").Value = DT_Date.Value
        Rs_Media_Quotation.Fields("Entered_By").Value = Me.Txt_Enterd_By.Text
        Rs_Media_Quotation.Fields("Approval_Client").Value = 0
    End If
    Rs_Media_Quotation.Fields("Plan_No").Value = (Txt_Plan_No.Text)
    Rs_Media_Quotation.Fields("Remarks").Value = (Txt_Remarks.Text)
    Rs_Media_Quotation.Fields("Entered_By").Value = (Txt_Enterd_By.Text)
    'Save record
    Rs_Media_Quotation.Update
    
    ' Save Detail Record
            'Month 1, 2, 3
    '**********************************
    '       Just Edit Delete Old Record
    
    'Media Quotation Detail
    '==========================================================
            For Index_Col = 1 To Flex_Quot.cols - 1 Step 2
                If Flex_Quot.TextMatrix(0, Index_Col) <> "" Then
                    strSql = "DELETE FROM ib_Radio_Quotation_Detail WHERE Job_ID='" & Flex_Quot.TextMatrix(1, Index_Col) & "'"
                    ConnERP.Execute strSql
                End If
            Next Index_Col
    
    
    
    '***********************************
    For Index_Col = 1 To 5 Step 2
        If Flex_Quot.TextMatrix(0, Index_Col) <> "" Then
            Nett_Total = 0
            MSC = 0
            Bonus_Fee = 0
            Others = 0
            Total_Lintas = 0
            Club_MSC = 0
            Grand_Total = 0
        
        'Assign
            'Nett
            If Flex_Quot.TextMatrix(2, Index_Col) = "" Then
                Nett_Total = 0
            Else
                Nett_Total = Flex_Quot.TextMatrix(2, Index_Col)
            End If
            
            'Media Spv Charge (Ambil Dari Flex Grid)
            'MSC = Brand_Info.MSC * Nett_Total
            
            'Bonus Fee
            If Flex_Quot.TextMatrix(4, Index_Col) = "" Then
                Bonus_Fee = 0
            Else
                Bonus_Fee = Flex_Quot.TextMatrix(4, Index_Col)
            End If
            
            'Others Cost
            If Flex_Quot.TextMatrix(5, Index_Col) = "" Then
                Others = 0
            Else
                Others = Flex_Quot.TextMatrix(5, Index_Col)
            End If
            
            'Total Lintas (Ambil Dari Flex Grid)
            'Total_Lintas = Nett_Total + MSC + Bonus_Fee + Others
            
            'Club Agency
            If Flex_Quot.TextMatrix(8, Index_Col) = "" Then
                Club_MSC = 0
            Else
                Club_MSC = Flex_Quot.TextMatrix(8, Index_Col)
            End If
            
            'Grand Total (Ambil Dari Flex Grid)
            'Grand_Total = Total_Lintas + Club_MSC
            
            
        'Execute
            strSql = ""
            strSql = "INSERT INTO IB_Radio_Quotation_Detail VALUES('"
            'strSql = strSql & Txt_CB_ID.Text & "','"
            strSql = strSql & "','"
            strSql = strSql & Trim(Txt_MQ.Text) & "','"
            strSql = strSql & Flex_Quot.TextMatrix(1, Index_Col) & "',"
            strSql = strSql & Get_Month_Number(Flex_Quot.TextMatrix(0, Index_Col)) & ","
            strSql = strSql & Val(Cbo_Year.Text) & ","
            'Nett
            strSql = strSql & Nett_Total & ","
            'MSC
            strSql = strSql & CCur(Flex_Quot.TextMatrix(3, Index_Col)) & ","
            'MsgBox CCur(Flex_Quot.TextMatrix(3, index_Col))
            'Others Cost
            strSql = strSql & CCur(Flex_Quot.TextMatrix(5, Index_Col)) & ","
            'Bonus Fee
            If Flex_Quot.TextMatrix(4, Index_Col) <> "" Then
                strSql = strSql & CCur(Flex_Quot.TextMatrix(4, Index_Col)) & ","
            Else
                strSql = strSql & "0 ,"
            End If
            'Total IMI
            strSql = strSql & CCur(Flex_Quot.TextMatrix(6, Index_Col)) & ","
            'Club Agency
            strSql = strSql & Club_MSC & ",'"
            'Club Agency Job Number
            strSql = strSql & Flex_Quot.TextMatrix(7, Index_Col) & "',"
            'Grand Total
            strSql = strSql & CCur(Flex_Quot.TextMatrix(9, Index_Col)) & ", "
            'Budget
            strSql = strSql & CCur(Flex_Quot.TextMatrix(11, Index_Col))
            'Source IB
            If Index_Col = 1 Then
                strSql = strSql & ", '" & Trim(IB_ID_1) & "' "
            End If
            If Index_Col = 3 Then
                strSql = strSql & ", '" & Trim(IB_ID_2) & "' "
            End If
            If Index_Col = 5 Then
                strSql = strSql & ", '" & Trim(IB_ID_3) & "' "
            End If
            
            strSql = strSql & ")"
    '        'Debug.Print Strsql
            ConnERP.Execute strSql
        End If
    Next Index_Col
    
            '*********************
            Rem ULI Budget Control
            '*********************
                If Add_flag And Not Add_flag_rev Then
                    If Flex_Quot.TextMatrix(0, 1) <> "" Then
                        strSql = " insert into uli_budget_control (Client_Brief_Id,Job_Id,Budget,budget_Balance) values("
                        strSql = strSql & "'" & Trim(Txt_CB_ID.Text) & "', "
                        strSql = strSql & "'" & Trim(Flex_Quot.TextMatrix(1, 1)) & "', "
                        strSql = strSql & CDbl(Format(Flex_Quot.TextMatrix(6, 1), "##0")) & ","
                        strSql = strSql & CDbl(Format(Flex_Quot.TextMatrix(6, 1), "##0")) & ")"
                        
                        ConnERP.Execute strSql
                    End If
                    If Flex_Quot.TextMatrix(0, 3) <> "" Then
                        strSql = " insert into uli_budget_control (Client_Brief_Id,Job_Id,Budget,budget_Balance) values("
                        strSql = strSql & "'" & Trim(Txt_CB_ID.Text) & "', "
                        strSql = strSql & "'" & Trim(Flex_Quot.TextMatrix(1, 3)) & "', "
                        strSql = strSql & CDbl(Format(Flex_Quot.TextMatrix(6, 3), "##0")) & ","
                        strSql = strSql & CDbl(Format(Flex_Quot.TextMatrix(6, 3), "##0")) & ")"
                        ConnERP.Execute strSql
                    End If
                    If Flex_Quot.TextMatrix(0, 5) <> "" Then
                        strSql = " insert into uli_budget_control (Client_Brief_Id,Job_Id,Budget,budget_Balance) values("
                        strSql = strSql & "'" & Trim(Txt_CB_ID.Text) & "', "
                        strSql = strSql & "'" & Trim(Flex_Quot.TextMatrix(1, 5)) & "', "
                        strSql = strSql & CDbl(Format(Flex_Quot.TextMatrix(6, 5), "##0")) & ","
                        strSql = strSql & CDbl(Format(Flex_Quot.TextMatrix(6, 5), "##0")) & ")"
                        ConnERP.Execute strSql
                    End If
                Else
                    'Ini Harus Di Check Ulang Apakah Mengupdate BC
                                                
                    Dim rs_Money_spent As New ADODB.Recordset
                         
                                                
                    If Flex_Quot.TextMatrix(0, 1) <> "" Then
                        strSql = "select * from uli_budget_control where job_id='" & Trim(Flex_Quot.TextMatrix(1, 1)) & "' "
                        rs_Money_spent.Open strSql, ConnERP, adOpenStatic, adLockReadOnly
                        
                        If Not rs_Money_spent.EOF Then
                                strSql = " UPDATE uli_budget_control SET "
                                strSql = strSql & " budget = " & CDbl(Format(Flex_Quot.TextMatrix(6, 1), "##0"))
                                'StrSQL = StrSQL & " , budget_balance = " & CDbl(Format(Flex_Quot.TextMatrix(6, 1), "##0"))
                                strSql = strSql & " WHERE "
                                strSql = strSql & "  Job_id = '" & Trim(Flex_Quot.TextMatrix(1, 1)) & "'"
                                ConnERP.Execute strSql
                                
                                'Betulin Balance
                                strSql = "UPDATE uli_budget_control SET budget_balance = budget - money_spent WHERE job_id='" & Flex_Quot.TextMatrix(1, 1) & "'"
                                ConnERP.Execute strSql
                        Else
                                strSql = " insert into uli_budget_control (Client_Brief_Id,Job_Id,Budget,budget_Balance) values("
                                strSql = strSql & "'" & Trim(Txt_CB_ID.Text) & "', "
                                strSql = strSql & "'" & Trim(Flex_Quot.TextMatrix(1, 1)) & "', "
                                strSql = strSql & CDbl(Format(Flex_Quot.TextMatrix(6, 1), "##0")) & ","
                                strSql = strSql & CDbl(Format(Flex_Quot.TextMatrix(6, 1), "##0")) & ")"
                                ConnERP.Execute strSql
                        End If
                        rs_Money_spent.Close
                        Set rs_Money_spent = Nothing
                    End If
                    
                    If Flex_Quot.TextMatrix(0, 3) <> "" Then
                        strSql = "SELECT * FROM uli_budget_control WHERE job_id='" & Trim(Flex_Quot.TextMatrix(1, 3)) & "' "
                        rs_Money_spent.Open strSql, ConnERP, adOpenStatic, adLockReadOnly
                        If Not rs_Money_spent.EOF Then
                                strSql = " UPDATE uli_budget_control SET "
                                strSql = strSql & " budget = " & CDbl(Format(Flex_Quot.TextMatrix(6, 3), "##0"))
                                'StrSQL = StrSQL & " , budget_balance = " & CDbl(Format(Flex_Quot.TextMatrix(6, 1), "##0"))
                                strSql = strSql & " WHERE "
                                strSql = strSql & "  Job_id = '" & Trim(Flex_Quot.TextMatrix(1, 3)) & "'"
                                ConnERP.Execute strSql
                                
                                'Betulin Balance
                                strSql = "UPDATE uli_budget_control set budget_balance = budget - money_spent where job_id='" & Flex_Quot.TextMatrix(1, 3) & "'"
                                ConnERP.Execute strSql
                        Else
                                strSql = " insert into uli_budget_control (Client_Brief_Id,Job_Id,Budget,budget_Balance) values("
                                strSql = strSql & "'" & Trim(Txt_CB_ID.Text) & "', "
                                strSql = strSql & "'" & Trim(Flex_Quot.TextMatrix(1, 3)) & "', "
                                strSql = strSql & CDbl(Format(Flex_Quot.TextMatrix(6, 3), "##0")) & ","
                                strSql = strSql & CDbl(Format(Flex_Quot.TextMatrix(6, 3), "##0")) & ")"
                                ConnERP.Execute strSql
                        End If
                        rs_Money_spent.Close
                        Set rs_Money_spent = Nothing
                    End If
                    
                    If Flex_Quot.TextMatrix(0, 5) <> "" Then
                        strSql = "SELECT * FROM uli_budget_control WHERE job_id='" & Trim(Flex_Quot.TextMatrix(1, 5)) & "' "
                        rs_Money_spent.Open strSql, ConnERP, adOpenStatic, adLockReadOnly
                        If Not rs_Money_spent.EOF Then
                                strSql = " UPDATE uli_budget_control SET "
                                strSql = strSql & " budget = " & CDbl(Format(Flex_Quot.TextMatrix(6, 5), "##0"))
                                'StrSQL = StrSQL & " , budget_balance = " & CDbl(Format(Flex_Quot.TextMatrix(6, 1), "##0"))
                                strSql = strSql & " WHERE "
                                strSql = strSql & "  Job_id = '" & Trim(Flex_Quot.TextMatrix(1, 5)) & "'"
                                ConnERP.Execute strSql
                                
                                'Betulin Balance
                                strSql = "UPDATE uli_budget_control SET budget_balance = budget - money_spent where job_id='" & Flex_Quot.TextMatrix(1, 5) & "'"
                                ConnERP.Execute strSql
                        Else
                                strSql = " insert into uli_budget_control (Client_Brief_Id,Job_Id,Budget,budget_Balance) values("
                                strSql = strSql & "'" & Trim(Txt_CB_ID.Text) & "', "
                                strSql = strSql & "'" & Trim(Flex_Quot.TextMatrix(1, 5)) & "', "
                                strSql = strSql & CDbl(Format(Flex_Quot.TextMatrix(6, 5), "##0")) & ","
                                strSql = strSql & CDbl(Format(Flex_Quot.TextMatrix(6, 5), "##0")) & ")"
                                ConnERP.Execute strSql
                        End If
                        rs_Money_spent.Close
                        Set rs_Money_spent = Nothing
                    End If
                End If

    
    ConnERP.CommitTrans
    
    Sukses_flag = True
           
    Exit Sub

Label_Err:

    Sukses_flag = False
    
    ConnERP.RollbackTrans
    
    Rs_Media_Quotation.Requery
    
    If Abs(Err.Number) = CDbl("2147217900") Then
        'delete Data
        MsgBox "Ada Revisi IB", vbInformation, strLogin_User
    Else
        MsgBox "Another Error.... " & Abs(Err.Number) & " " & Err.Description, vbCritical, strLogin_User
    End If

End Sub

Private Sub Delete()
    Dim Param_IN As New ADODB.Parameter
    Dim Cmd_SP As New ADODB.Command
    On Error GoTo my_error
    
    ConnERP.BeginTrans
    Cmd_SP.CommandType = adCmdStoredProc
    Cmd_SP.CommandText = "Delete_Radio_QUot"
    
    Set Param_IN = Cmd_SP.CreateParameter("IB_ID", adChar, adParamInput, 13)
    Cmd_SP.Parameters.Append Param_IN
    Param_IN.Value = Trim(Cbo_MQ.Text)
    
    Cmd_SP.ActiveConnection = ConnERP
    Cmd_SP.Execute
    ConnERP.CommitTrans
    Exit Sub
    
my_error:
    ConnERP.RollbackTrans
    MsgBox Err.Number & " " & Err.Description
End Sub

Private Sub Prepare_Print()
    Dim TxtSQl As String
    Dim rs As New ADODB.Recordset
    
    TxtSQl = " select   Month_Catalog.Month_Name, Brand.Brand_Name, "
    TxtSQl = TxtSQl & " Client.Client_Name, "
    TxtSQl = TxtSQl & " Company.Company_Name, "
    TxtSQl = TxtSQl & " IB_Radio.Media_Plan, IB_Radio.IB_ID, "
    TxtSQl = TxtSQl & " IB_Radio_Quot.Remarks, "
    TxtSQl = TxtSQl & " IB_Radio_Quot.Date, "
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
    TxtSQl = TxtSQl & " inner join IB_Radio on ib_radio.Ib_ID = Ib_Radio_quot.ib_id "
    TxtSQl = TxtSQl & " inner join Brand on IB_Radio.brand_code = brand.brand_code "
    TxtSQl = TxtSQl & " inner join client on brand.Client_Code = Client.Client_Code "
    TxtSQl = TxtSQl & " inner join company on Brand.Company_Code = Company.Company_Code "
    TxtSQl = TxtSQl & " inner join Month_Catalog on Month_Catalog.Month = ib_radio_quotation_detail.Month "
    TxtSQl = TxtSQl & " where ib_radio_Quot.Ib_ID ='" & Cbo_MQ.Text & "'"
    TxtSQl = TxtSQl & " order by IB_Radio_Quotation_Detail.month asc"
    
    ''Debug.Print TxtSQL
    rs.Open TxtSQl, ConnERP, adOpenStatic, adLockReadOnly
    
    Rem Filename .RPT
    CR.ReportFileName = Report_Dir & "\Radio\mq_Radio.rpt"
    
    Rem Header Report
    With rs
        If .EOF = False Then
            CR.ParameterFields(39) = "Client;" & .Fields("Client_Name") & ";TRUE"
            CR.ParameterFields(1) = "Brand;" & .Fields("Brand_Name") & ";TRUE"
            CR.ParameterFields(2) = "MediaType;Radio;TRUE"
            CR.ParameterFields(3) = "MediaPlanNo;" & .Fields("Media_Plan") & ";TRUE"
            CR.ParameterFields(4) = "Dated;" & .Fields("Date") & ";TRUE"
            CR.ParameterFields(5) = "IBID;" & .Fields("IB_ID") & ";TRUE"
            CR.ParameterFields(6) = "Remarks;" & .Fields("Remarks") & ";TRUE"
            CR.ParameterFields(7) = "PT;" & .Fields("Company_Name") & ";TRUE"
            CR.ParameterFields(8) = "Marketing;Marketing Manager;TRUE"
        End If
        
        Rem 1st Month
        If .EOF = False Then
            CR.ParameterFields(9) = "MONTH1;" & .Fields("Month_Name") & ";TRUE"
            CR.ParameterFields(10) = "Year1;" & .Fields("Year") & ";TRUE"
            CR.ParameterFields(11) = "Nett1;" & .Fields("Nett_Cost") & ";TRUE"
            CR.ParameterFields(12) = "MSC1;" & .Fields("Media_Sptv_Charge") & ";TRUE"
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
            CR.ParameterFields(21) = "MSC2;" & .Fields("Media_Sptv_Charge") & ";TRUE"
            CR.ParameterFields(22) = "Other2;" & .Fields("Other_Charge") & ";TRUE"
            CR.ParameterFields(23) = "TotalLintas2;" & .Fields("Total_Lintas") & ";TRUE"
            CR.ParameterFields(24) = "JobNoAG2;" & .Fields("Media_Sptv_Charge") & ";TRUE"
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
            CR.ParameterFields(30) = "MSC3;" & .Fields("Media_Sptv_Charge") & ";TRUE"
            CR.ParameterFields(31) = "Other3;" & .Fields("Other_Charge") & ";TRUE"
            CR.ParameterFields(32) = "TotalLintas3;" & .Fields("Total_Lintas") & ";TRUE"
            CR.ParameterFields(33) = "JobNoAG3;" & .Fields("Media_Sptv_Charge") & ";TRUE"
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
        
    CR.Connect = "DSN =" & Server_Name & ";UID = " & Login_User & ";DSQ = " & Database_Name & "; PWD =" & Login_Password
    CR.Action = 1

End Sub



Private Sub Cancel_MQ_Number()
    Dim rs As New ADODB.Recordset
    Dim TxtSQl As String

    Rem Cancel MQ_Number and Insert Into Reuseable_MQ_Radio Table
    TxtSQl = " Delete from MQ_Radio_Running "
    TxtSQl = TxtSQl & " where year = " & Val(Cbo_Year.Text)
    TxtSQl = TxtSQl & " and brand_code = '" & Left(Cbo_Brand.Text, 4) & "'"
    TxtSQl = TxtSQl & " and MQ_Number ='" & Trim(Txt_MQ.Text) & "'"
    ConnERP.Execute TxtSQl
    
    TxtSQl = "insert into reuseable_mq_radio (MQ_Number, Year,Brand_Code) values ( "
    TxtSQl = TxtSQl & " '" & Trim(Txt_MQ.Text) & "', "
    TxtSQl = TxtSQl & Val(Cbo_Year.Text) & ", "
    TxtSQl = TxtSQl & " '" & Left(Cbo_Brand.Text, 4) & "') "
    ConnERP.Execute TxtSQl
End Sub

Private Sub Load_CB()
    Dim rs As New ADODB.Recordset
    Dim TxtSQl As String
    
    TxtSQl = " select client_Brief_ID from Client_Brief_Media where Brand_Code ='" & Left(Cbo_Brand.Text, 4) & "'"
    rs.Open TxtSQl, ConnERP, adOpenStatic, adLockPessimistic
    With rs
        Cbo_CB.Clear
        Do While .EOF = False
            Cbo_CB.AddItem .Fields("Client_brief_ID").Value
            .MoveNext
        Loop
    End With

End Sub

Private Sub Load_Month()
    Cbo_Month_MQ.Clear
    Cbo_Month_MQ.AddItem "-None-"
    
    'If What_Month <> "January" Then
        Cbo_Month_MQ.AddItem "January"
    'End If
    
    'If What_Month <> "January" Then
        Cbo_Month_MQ.AddItem "February"
    'End If
    
    'If What_Month <> "January" Then
        Cbo_Month_MQ.AddItem "March"
    'End If
    
    'If What_Month <> "January" Then
        Cbo_Month_MQ.AddItem "April"
    'End If
    
    'If What_Month <> "January" Then
        Cbo_Month_MQ.AddItem "May"
    'End If
    
    'If What_Month <> "January" Then
        Cbo_Month_MQ.AddItem "June"
    'End If
    
    'If What_Month <> "January" Then
        Cbo_Month_MQ.AddItem "July"
    'End If
    
    'If What_Month <> "January" Then
        Cbo_Month_MQ.AddItem "August"
    'End If
    
    'If What_Month <> "January" Then
        Cbo_Month_MQ.AddItem "September"
    'End If
    'If What_Month <> "January" Then
        Cbo_Month_MQ.AddItem "October"
   'End If
    
    'If What_Month <> "January" Then
        Cbo_Month_MQ.AddItem "November"
    'End If
    
    'If What_Month <> "January" Then
        Cbo_Month_MQ.AddItem "December"
    'End If
End Sub

Sub SetButtonToolbar(ByVal paIsNormalMode As Boolean, picOBJ) 'TOOLBAR_AI.
'************************************************
' Procedure         : SetButtonToolbar
' Function          : TOOLBAR_AI.
' Created By        : {73 64 6B}
' Date              : 12-Apr-2015
' LastUpdate/By     : - Rudi
'************************************************

    Dim element
    Dim strDummy As String
    
    With picButton(enButtonType.bieFirst)  'FIRST
        .Enabled = paIsNormalMode
    End With
    
    With picButton(enButtonType.biePrev)  'PREVIOUS
        .Enabled = paIsNormalMode
    End With
    
    With picButton(enButtonType.bieNext)  'NEXT
        .Enabled = paIsNormalMode
    End With
    
    With picButton(enButtonType.bieLast)  'LAST
        .Enabled = paIsNormalMode
    End With
    
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
    
    With picButton(enButtonType.bieClose)      'CLOSE.
        .Enabled = paIsNormalMode
        .Visible = paIsNormalMode
    End With

    With picButton(enButtonType.bieSave)  'SAVE.
        .Enabled = Not paIsNormalMode
        .Visible = Not paIsNormalMode
        .Left = picButton(4).Left
    End With
    
    With picButton(enButtonType.bieApprove)  'APPROVE.
        .Enabled = paIsNormalMode
        .Visible = paIsNormalMode
    End With
    
    With picButton(enButtonType.bieApprovedQuotationList)  'APPROVE.
        .Enabled = paIsNormalMode
        .Visible = paIsNormalMode
    End With
    
    With picButton(enButtonType.bieRevisionHistory)  'APPROVE.
        .Enabled = paIsNormalMode
        .Visible = paIsNormalMode
    End With
    
    With picButton(enButtonType.biePublishToWeb)  'publish.
        .Enabled = paIsNormalMode
        .Visible = paIsNormalMode
    End With
    
    With picButton(enButtonType.biePrint)  'PRINT.
        .Enabled = paIsNormalMode
        .Visible = paIsNormalMode
    End With

    With picButton(enButtonType.biecancel) 'CANCEL.
        .Enabled = Not paIsNormalMode
        .Visible = Not paIsNormalMode
        .Left = picButton(5).Left
    End With
    
    For Each element In picOBJ
        SetPictureTB element.Index, paIsNormalMode, picOBJ
    Next element

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

Sub EnableObject(ByVal paIsEnable As Boolean)
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

Private Sub picButton_Click(Index As Integer)
'************************************************
' Procedure         : picButton_Click
' Function          : Action utk Navigation dan CRUD.
' Created By        : {73 64 6B}
' Date              : 12-Apr-2015/{73 64 6B} --> Semua coding dan query sudah di optimalkan agar faster, readable, safer, standardable.
'************************************************
    Dim strCode As String, strFileRpt As String
    
    Select Case Index
        Case enButtonType.bieFirst 'FIRST.
            On Error Resume Next
            If Cbo_MQ.ListCount > 0 Then
                Cbo_MQ.ListIndex = 0
            End If
            
        Case enButtonType.biePrev  'PREV.
            On Error Resume Next
            If Cbo_MQ.ListCount > 0 Then
                If Cbo_MQ.ListIndex < Cbo_MQ.ListCount And Cbo_MQ.ListIndex > -1 Then
                    If Cbo_MQ.ListIndex > 0 Then
                        Cbo_MQ.ListIndex = Cbo_MQ.ListIndex - 1
                    Else
                        MsgBox "First Record", vbInformation, strCompany_Name
                    End If
                End If
            End If
            
        Case enButtonType.bieNext 'NEXT.
            On Error Resume Next
            If Cbo_MQ.ListCount > 0 Then
                If Cbo_MQ.ListIndex < Cbo_MQ.ListCount Then
                    If Cbo_MQ.ListIndex < Cbo_MQ.ListCount - 1 Then
                        Cbo_MQ.ListIndex = Cbo_MQ.ListIndex + 1
                    Else
                        MsgBox "Last record", vbInformation, strCompany_Name
                    End If
                End If
            End If
            
        Case enButtonType.bieLast  'LAST.
            On Error Resume Next
            If Cbo_MQ.ListCount > 0 Then
                Cbo_MQ.ListIndex = Cbo_MQ.ListCount - 1
            End If
            
        Case enButtonType.bieAdd  '4 'ADD.
            Call db_add
            
        Case enButtonType.bieEdit  '5 'EDIT.
            Call db_edit
            
        Case enButtonType.bieDelete  '6 'DELETE.
            Call db_delete
        
        Case enButtonType.biePrint   '8 'PRINT.
            Call db_print
        
        Case enButtonType.bieApprove   '66 'APPROVE.
            Call Fra_Approval_DblClick
            
        Case enButtonType.bieRevisionHistory
            Call db_History_Revision
            
        Case enButtonType.bieApprovedQuotationList
            Call db_Approved_MQ
            
        Case enButtonType.biePublishToWeb   '64 'PUBLISH TO WEB.
            Call db_Publish_to_Web
            
        Case enButtonType.bieClose  '23 'CLOSE.
            Unload Me
            
        Case enButtonType.bieSave  'SAVE.
            Call db_save
            
        Case enButtonType.biecancel 'CANCEL.
            Call db_Cancel
    End Select

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
    pnl_Main.Height = Me.Height - pnl_Main.Top - picStatusBar.Height
    Frame1.Width = ((pnl_Main.Width / 4) * 3) - Frame1.Left
    Frame11.Left = Frame1.Left + Frame1.Width + 100
    Frame11.Width = pnl_Main.Width - Frame11.Left - Frame1.Left
    Frame2.Width = pnl_Main.Width - (Frame2.Left * 2)
    Flex_Quot.Width = Frame2.Width - (Flex_Quot.Left * 2)
    Frame3.Width = ((pnl_Main.Width / 4) * 3) - Frame3.Left
    Fra_Approval.Left = Frame3.Left + Frame3.Width + 100
    Fra_Approval.Width = pnl_Main.Width - Fra_Approval.Left - Frame3.Left
    Lbl_APP.Width = Fra_Approval.Width - (Lbl_APP.Left * 2)
    Lbl_Date.Width = (Fra_Approval.Width / 2) - Lbl_Date.Left - 50
    Lbl_Time.Left = Lbl_Date.Left + Lbl_Date.Width + 100
    Lbl_Time.Width = Lbl_Date.Width
    Txt_Remarks.Width = Frame3.Width - (Txt_Remarks.Left * 2)
End Sub

Sub setButtonHistory(ByVal blnStatus As Boolean, picOBJ) 'TOOLBAR_AI.
'************************************************
' Procedure         : SetButtonToolbar
' Function          : TOOLBAR_AI.
' Created By        : {73 64 6B}
' Date              : 12-Apr-2015
' LastUpdate/By     : - Rudi
'************************************************

    Dim element
    Dim strDummy As String
    
    With picButton(enButtonType.bieRevisionHistory)  'APPROVE.
        .Enabled = blnStatus
    End With
    
    For Each element In picOBJ
        SetPictureTB element.Index, blnStatus, picOBJ
    Next element

End Sub
