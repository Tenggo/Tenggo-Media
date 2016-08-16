VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Object = "{B8E5842E-102B-4289-9D57-3B3F5B5E15D3}#16.4#0"; "CODEJO~1.OCX"
Begin VB.MDIForm mdi_Main 
   BackColor       =   &H8000000C&
   ClientHeight    =   7695
   ClientLeft      =   1860
   ClientTop       =   2010
   ClientWidth     =   11400
   Icon            =   "mdi_Main.frx":0000
   LinkTopic       =   "MDIForm1"
   LockControls    =   -1  'True
   WindowState     =   2  'Maximized
   Begin VB.PictureBox picTopBar 
      Align           =   1  'Align Top
      BorderStyle     =   0  'None
      Height          =   1065
      Left            =   0
      ScaleHeight     =   1065
      ScaleWidth      =   11400
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   0
      Width           =   11400
      Begin VB.PictureBox picSidebars 
         BackColor       =   &H00E3E3E3&
         BorderStyle     =   0  'None
         Height          =   330
         Left            =   0
         ScaleHeight     =   330
         ScaleWidth      =   3435
         TabIndex        =   26
         TabStop         =   0   'False
         Top             =   720
         Width           =   3435
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Favourite View"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00404040&
            Height          =   255
            Left            =   90
            TabIndex        =   27
            Top             =   15
            Width           =   2655
         End
      End
      Begin VB.PictureBox picTabs 
         BackColor       =   &H00E3E3E3&
         BorderStyle     =   0  'None
         FillColor       =   &H00F1E6DC&
         Height          =   360
         Left            =   3435
         ScaleHeight     =   360
         ScaleWidth      =   18735
         TabIndex        =   9
         TabStop         =   0   'False
         Top             =   690
         Width           =   18735
         Begin VB.PictureBox picTab 
            BackColor       =   &H00FAE196&
            BorderStyle     =   0  'None
            Height          =   285
            Index           =   0
            Left            =   45
            ScaleHeight     =   285
            ScaleWidth      =   1695
            TabIndex        =   24
            TabStop         =   0   'False
            Top             =   45
            Visible         =   0   'False
            Width           =   1695
            Begin VB.Label lblTab 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "aaaa"
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
               Left            =   50
               TabIndex        =   25
               Top             =   30
               Width           =   360
            End
         End
         Begin VB.PictureBox picTab 
            BackColor       =   &H00BFBFBF&
            BorderStyle     =   0  'None
            ForeColor       =   &H00E0E0E0&
            Height          =   285
            Index           =   1
            Left            =   1785
            ScaleHeight     =   285
            ScaleWidth      =   1695
            TabIndex        =   22
            TabStop         =   0   'False
            Top             =   45
            Visible         =   0   'False
            Width           =   1695
            Begin VB.Label lblTab 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "aaaa"
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
               Left            =   50
               TabIndex        =   23
               Top             =   30
               Width           =   360
            End
         End
         Begin VB.PictureBox picTab 
            BackColor       =   &H00BFBFBF&
            BorderStyle     =   0  'None
            Height          =   285
            Index           =   2
            Left            =   3525
            ScaleHeight     =   285
            ScaleWidth      =   1695
            TabIndex        =   20
            TabStop         =   0   'False
            Top             =   45
            Visible         =   0   'False
            Width           =   1695
            Begin VB.Label lblTab 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "aaaa"
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
               Index           =   2
               Left            =   50
               TabIndex        =   21
               Top             =   15
               Width           =   360
            End
         End
         Begin VB.PictureBox picTab 
            BackColor       =   &H00BFBFBF&
            BorderStyle     =   0  'None
            Height          =   285
            Index           =   3
            Left            =   5265
            ScaleHeight     =   285
            ScaleWidth      =   1695
            TabIndex        =   18
            TabStop         =   0   'False
            Top             =   45
            Visible         =   0   'False
            Width           =   1695
            Begin VB.Label lblTab 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "aaaa"
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
               Index           =   3
               Left            =   50
               TabIndex        =   19
               Top             =   15
               Width           =   360
            End
         End
         Begin VB.PictureBox picTab 
            BackColor       =   &H00BFBFBF&
            BorderStyle     =   0  'None
            Height          =   285
            Index           =   4
            Left            =   7005
            ScaleHeight     =   285
            ScaleWidth      =   1695
            TabIndex        =   16
            TabStop         =   0   'False
            Top             =   45
            Visible         =   0   'False
            Width           =   1695
            Begin VB.Label lblTab 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "aaaa"
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
               Index           =   4
               Left            =   50
               TabIndex        =   17
               Top             =   15
               Width           =   360
            End
         End
         Begin VB.PictureBox picTab 
            BackColor       =   &H00BFBFBF&
            BorderStyle     =   0  'None
            Height          =   285
            Index           =   5
            Left            =   8745
            ScaleHeight     =   285
            ScaleWidth      =   1695
            TabIndex        =   14
            TabStop         =   0   'False
            Top             =   45
            Visible         =   0   'False
            Width           =   1695
            Begin VB.Label lblTab 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "aaaa"
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
               Index           =   5
               Left            =   50
               TabIndex        =   15
               Top             =   15
               Width           =   360
            End
         End
         Begin VB.PictureBox picTab 
            BackColor       =   &H00BFBFBF&
            BorderStyle     =   0  'None
            Height          =   285
            Index           =   6
            Left            =   10485
            ScaleHeight     =   285
            ScaleWidth      =   1695
            TabIndex        =   12
            TabStop         =   0   'False
            Top             =   45
            Visible         =   0   'False
            Width           =   1695
            Begin VB.Label lblTab 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "aaaa"
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
               Index           =   6
               Left            =   50
               TabIndex        =   13
               Top             =   15
               Width           =   360
            End
         End
         Begin VB.PictureBox picTab 
            BackColor       =   &H00BFBFBF&
            BorderStyle     =   0  'None
            Height          =   285
            Index           =   7
            Left            =   12225
            ScaleHeight     =   285
            ScaleWidth      =   1695
            TabIndex        =   10
            TabStop         =   0   'False
            Top             =   45
            Visible         =   0   'False
            Width           =   1695
            Begin VB.Label lblTab 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "aaaa"
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
               Index           =   7
               Left            =   50
               TabIndex        =   11
               Top             =   15
               Width           =   360
            End
         End
      End
      Begin VB.PictureBox picLine 
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   0  'None
         Height          =   15
         Index           =   0
         Left            =   0
         ScaleHeight     =   15
         ScaleWidth      =   32460
         TabIndex        =   8
         TabStop         =   0   'False
         Top             =   0
         Width           =   32465
      End
      Begin VB.PictureBox Picture1 
         BackColor       =   &H00BFBFBF&
         BorderStyle     =   0  'None
         Height          =   705
         Left            =   -15
         ScaleHeight     =   705
         ScaleWidth      =   2145
         TabIndex        =   5
         TabStop         =   0   'False
         Top             =   0
         Width           =   2145
         Begin VB.PictureBox picButton 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0E0FF&
            BorderStyle     =   0  'None
            ClipControls    =   0   'False
            ForeColor       =   &H80000008&
            Height          =   675
            Index           =   25
            Left            =   15
            Picture         =   "mdi_Main.frx":0442
            ScaleHeight     =   675
            ScaleWidth      =   1050
            TabIndex        =   7
            TabStop         =   0   'False
            Top             =   15
            Width           =   1050
         End
         Begin VB.PictureBox picButton 
            Appearance      =   0  'Flat
            BackColor       =   &H00F0F0F0&
            BorderStyle     =   0  'None
            ClipControls    =   0   'False
            ForeColor       =   &H80000008&
            Height          =   675
            Index           =   28
            Left            =   1080
            ScaleHeight     =   675
            ScaleWidth      =   1050
            TabIndex        =   6
            TabStop         =   0   'False
            Top             =   15
            Width           =   1050
         End
      End
      Begin VB.PictureBox picLine 
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   0  'None
         Height          =   15
         Index           =   1
         Left            =   0
         ScaleHeight     =   15
         ScaleWidth      =   32460
         TabIndex        =   4
         TabStop         =   0   'False
         Top             =   690
         Width           =   32465
      End
   End
   Begin VB.PictureBox picSideBar 
      Align           =   3  'Align Left
      BackColor       =   &H00F0F0F0&
      Height          =   6255
      Left            =   0
      ScaleHeight     =   6195
      ScaleWidth      =   3390
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   1065
      Width           =   3450
      Begin VB.PictureBox picLogoBottom 
         BackColor       =   &H00565656&
         BorderStyle     =   0  'None
         ClipControls    =   0   'False
         Height          =   750
         Left            =   0
         ScaleHeight     =   750
         ScaleWidth      =   3420
         TabIndex        =   2
         TabStop         =   0   'False
         Top             =   6375
         Width           =   3420
         Begin VB.Image Image1 
            Height          =   570
            Left            =   1200
            Picture         =   "mdi_Main.frx":1AEC
            Top             =   165
            Width           =   810
         End
      End
      Begin XtremeTaskPanel.TaskPanel wndTaskPanel 
         Height          =   6495
         Left            =   -15
         TabIndex        =   1
         Top             =   105
         Width           =   3495
         _Version        =   1048580
         _ExtentX        =   6165
         _ExtentY        =   11456
         _StockProps     =   64
         VisualTheme     =   9
         ItemLayout      =   2
         HotTrackStyle   =   1
         Begin MSComctlLib.ImageList imlTaskPanelIcons1 
            Left            =   0
            Top             =   0
            _ExtentX        =   1005
            _ExtentY        =   1005
            BackColor       =   5658198
            ImageWidth      =   16
            ImageHeight     =   16
            MaskColor       =   5658198
            _Version        =   393216
            BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
               NumListImages   =   32
               BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "mdi_Main.frx":5A51
                  Key             =   ""
               EndProperty
               BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "mdi_Main.frx":5DA5
                  Key             =   ""
               EndProperty
               BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "mdi_Main.frx":60F9
                  Key             =   ""
               EndProperty
               BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "mdi_Main.frx":644D
                  Key             =   ""
               EndProperty
               BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "mdi_Main.frx":67A1
                  Key             =   ""
               EndProperty
               BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "mdi_Main.frx":6AF5
                  Key             =   ""
               EndProperty
               BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "mdi_Main.frx":6E49
                  Key             =   ""
               EndProperty
               BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "mdi_Main.frx":719D
                  Key             =   ""
               EndProperty
               BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "mdi_Main.frx":74F1
                  Key             =   ""
               EndProperty
               BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "mdi_Main.frx":7845
                  Key             =   ""
               EndProperty
               BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "mdi_Main.frx":7B99
                  Key             =   ""
               EndProperty
               BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "mdi_Main.frx":7EED
                  Key             =   ""
               EndProperty
               BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "mdi_Main.frx":8241
                  Key             =   ""
               EndProperty
               BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "mdi_Main.frx":8595
                  Key             =   ""
               EndProperty
               BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "mdi_Main.frx":88E9
                  Key             =   ""
               EndProperty
               BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "mdi_Main.frx":8C3D
                  Key             =   ""
               EndProperty
               BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "mdi_Main.frx":8F91
                  Key             =   ""
               EndProperty
               BeginProperty ListImage18 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "mdi_Main.frx":92E5
                  Key             =   ""
               EndProperty
               BeginProperty ListImage19 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "mdi_Main.frx":9639
                  Key             =   ""
               EndProperty
               BeginProperty ListImage20 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "mdi_Main.frx":998D
                  Key             =   ""
               EndProperty
               BeginProperty ListImage21 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "mdi_Main.frx":9CE1
                  Key             =   ""
               EndProperty
               BeginProperty ListImage22 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "mdi_Main.frx":A035
                  Key             =   ""
               EndProperty
               BeginProperty ListImage23 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "mdi_Main.frx":A389
                  Key             =   ""
               EndProperty
               BeginProperty ListImage24 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "mdi_Main.frx":A6DD
                  Key             =   ""
               EndProperty
               BeginProperty ListImage25 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "mdi_Main.frx":AA31
                  Key             =   ""
               EndProperty
               BeginProperty ListImage26 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "mdi_Main.frx":AD85
                  Key             =   ""
               EndProperty
               BeginProperty ListImage27 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "mdi_Main.frx":B0D9
                  Key             =   ""
               EndProperty
               BeginProperty ListImage28 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "mdi_Main.frx":B42D
                  Key             =   ""
               EndProperty
               BeginProperty ListImage29 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "mdi_Main.frx":B781
                  Key             =   ""
               EndProperty
               BeginProperty ListImage30 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "mdi_Main.frx":BAD5
                  Key             =   ""
               EndProperty
               BeginProperty ListImage31 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "mdi_Main.frx":BE29
                  Key             =   ""
               EndProperty
               BeginProperty ListImage32 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "mdi_Main.frx":C17D
                  Key             =   ""
               EndProperty
            EndProperty
         End
         Begin MSComctlLib.ImageList imlTaskPanelIcons 
            Left            =   735
            Top             =   450
            _ExtentX        =   1005
            _ExtentY        =   1005
            BackColor       =   5658198
            ImageWidth      =   16
            ImageHeight     =   16
            MaskColor       =   5658198
            _Version        =   393216
            BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
               NumListImages   =   29
               BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "mdi_Main.frx":C4D1
                  Key             =   ""
               EndProperty
               BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "mdi_Main.frx":C825
                  Key             =   ""
               EndProperty
               BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "mdi_Main.frx":CB79
                  Key             =   ""
               EndProperty
               BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "mdi_Main.frx":CECD
                  Key             =   ""
               EndProperty
               BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "mdi_Main.frx":D221
                  Key             =   ""
               EndProperty
               BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "mdi_Main.frx":D575
                  Key             =   ""
               EndProperty
               BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "mdi_Main.frx":D8C9
                  Key             =   ""
               EndProperty
               BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "mdi_Main.frx":DC1D
                  Key             =   ""
               EndProperty
               BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "mdi_Main.frx":DF71
                  Key             =   ""
               EndProperty
               BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "mdi_Main.frx":E2C5
                  Key             =   ""
               EndProperty
               BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "mdi_Main.frx":E619
                  Key             =   ""
               EndProperty
               BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "mdi_Main.frx":E96D
                  Key             =   ""
               EndProperty
               BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "mdi_Main.frx":ECC1
                  Key             =   ""
               EndProperty
               BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "mdi_Main.frx":F015
                  Key             =   ""
               EndProperty
               BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "mdi_Main.frx":F369
                  Key             =   ""
               EndProperty
               BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "mdi_Main.frx":F6BD
                  Key             =   ""
               EndProperty
               BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "mdi_Main.frx":FA11
                  Key             =   ""
               EndProperty
               BeginProperty ListImage18 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "mdi_Main.frx":FD65
                  Key             =   ""
               EndProperty
               BeginProperty ListImage19 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "mdi_Main.frx":100B9
                  Key             =   ""
               EndProperty
               BeginProperty ListImage20 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "mdi_Main.frx":1040D
                  Key             =   ""
               EndProperty
               BeginProperty ListImage21 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "mdi_Main.frx":10761
                  Key             =   ""
               EndProperty
               BeginProperty ListImage22 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "mdi_Main.frx":10AB5
                  Key             =   ""
               EndProperty
               BeginProperty ListImage23 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "mdi_Main.frx":10E09
                  Key             =   ""
               EndProperty
               BeginProperty ListImage24 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "mdi_Main.frx":1115D
                  Key             =   ""
               EndProperty
               BeginProperty ListImage25 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "mdi_Main.frx":114B1
                  Key             =   ""
               EndProperty
               BeginProperty ListImage26 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "mdi_Main.frx":11805
                  Key             =   ""
               EndProperty
               BeginProperty ListImage27 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "mdi_Main.frx":11B59
                  Key             =   ""
               EndProperty
               BeginProperty ListImage28 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "mdi_Main.frx":11EAD
                  Key             =   ""
               EndProperty
               BeginProperty ListImage29 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "mdi_Main.frx":12201
                  Key             =   ""
               EndProperty
            EndProperty
         End
      End
   End
   Begin VB.Timer Tmr_Cek_Task 
      Interval        =   15000
      Left            =   0
      Top             =   0
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   28
      Top             =   7320
      Width           =   11400
      _ExtentX        =   20108
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   7
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Enabled         =   0   'False
            Object.Width           =   8008
            MinWidth        =   8008
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   3
            AutoSize        =   1
            Enabled         =   0   'False
            Object.Width           =   2469
            MinWidth        =   2469
            TextSave        =   "INS"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   2
            Object.Width           =   2469
            MinWidth        =   2469
            TextSave        =   "NUM"
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   1
            Enabled         =   0   'False
            Object.Width           =   2469
            MinWidth        =   2469
            TextSave        =   "CAPS"
         EndProperty
         BeginProperty Panel5 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Object.Visible         =   0   'False
         EndProperty
         BeginProperty Panel6 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   6
            TextSave        =   "3/8/2016"
         EndProperty
         BeginProperty Panel7 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            Alignment       =   2
            AutoSize        =   1
            TextSave        =   "1:08 AM"
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Menu Mnu_Production 
      Caption         =   "&Production"
      Visible         =   0   'False
   End
   Begin VB.Menu Mnu_New_Media 
      Caption         =   "&Media"
      Begin VB.Menu mnuAccountManagement 
         Caption         =   "&Account Management"
         Enabled         =   0   'False
         Visible         =   0   'False
         Begin VB.Menu Mnu_Client_Brief 
            Caption         =   "Client Brief Production"
         End
      End
      Begin VB.Menu Garis1 
         Caption         =   "-"
         Visible         =   0   'False
      End
      Begin VB.Menu Mnu_Media_Planner 
         Caption         =   "Media &Planner"
         Begin VB.Menu Mnu_Client_Brief_Media 
            Caption         =   "Client Brief Media"
         End
         Begin VB.Menu Mnu_Media_Plan 
            Caption         =   "Media Plan"
         End
         Begin VB.Menu Mnu_Implentation_Brief 
            Caption         =   "Implementation Brief"
            Begin VB.Menu Mnu_IB_Television 
               Caption         =   "Television"
            End
            Begin VB.Menu Mnu_IB_Radio 
               Caption         =   "Radio"
            End
            Begin VB.Menu Mnu_IB_Print 
               Caption         =   "Print"
            End
            Begin VB.Menu Mnu_IB_Others 
               Caption         =   "Others"
            End
         End
         Begin VB.Menu mnu_Retainer_Fee_Entry 
            Caption         =   "Retainer Fee Entry"
         End
      End
      Begin VB.Menu Garis2 
         Caption         =   "-"
      End
      Begin VB.Menu Mnu_Media_Implemetator 
         Caption         =   "Media &Implementer"
         Begin VB.Menu Mnu_View_Media_Plan 
            Caption         =   "Media Plan (Read Only)"
         End
         Begin VB.Menu Mnu_New_TV 
            Caption         =   "New Television"
            Begin VB.Menu Mnu_New_TV_MQ 
               Caption         =   "Television Media Quotation"
            End
            Begin VB.Menu Mnu_New_TV_Implementation 
               Caption         =   "Television Implementation"
            End
            Begin VB.Menu Mnu_New_TV_Finalize_PO 
               Caption         =   "Finalize Purchase Order"
            End
            Begin VB.Menu Mnu_New_TV_Schedule_Org 
               Caption         =   "Original TV Schedule"
            End
            Begin VB.Menu Mnu_New_TV_Material 
               Caption         =   "TV Material Catalog"
            End
            Begin VB.Menu Mnu_Transfer_TVS 
               Caption         =   "Transfer TV Shcedule to Excel"
            End
            Begin VB.Menu Mnu_New_TV_Preemption_log 
               Caption         =   "Preemption Log"
            End
         End
         Begin VB.Menu Mnu_Mid_Radio 
            Caption         =   "Radio"
            Begin VB.Menu Mnu_I_B_Quotation_Radio 
               Caption         =   "Radio Media Quotation"
            End
            Begin VB.Menu Mnu_Radio_Schedule 
               Caption         =   "Radio Monthly Media Schedule"
            End
            Begin VB.Menu Mnu_Radio_Purcashe_Order 
               Caption         =   "Purchase Order"
            End
            Begin VB.Menu Mnu_Radio_Cancelation_Order 
               Caption         =   "Cancel Order"
            End
            Begin VB.Menu Garis19 
               Caption         =   "-"
            End
            Begin VB.Menu Mnu_Radio_Station_Address_Lable 
               Caption         =   "Radio Station Address Label"
            End
            Begin VB.Menu Mnu_Radio_Address_Lable_Per_Job 
               Caption         =   "Radio Station Address Label Per Job"
            End
         End
         Begin VB.Menu Mnu_MIB_Print 
            Caption         =   "Print"
            Begin VB.Menu Mnu_I_B_Quotation_Print 
               Caption         =   "Print Media Quotation"
            End
            Begin VB.Menu Mnu_Print_Schedule 
               Caption         =   "Print Schedule"
            End
            Begin VB.Menu mnu_PO_Print 
               Caption         =   "Purchase Order"
            End
            Begin VB.Menu mnu_Replace_Order_Print 
               Caption         =   "Replace Order "
            End
            Begin VB.Menu Cancel_Order_Print 
               Caption         =   "Cancel Order"
            End
         End
         Begin VB.Menu Mnu_MIB_Others 
            Caption         =   "Others"
            Begin VB.Menu Mnu_Other_IB_Quotation 
               Caption         =   "Other Media Quotation"
            End
            Begin VB.Menu mnu_Other_Monthly_Media_Schedule 
               Caption         =   "Other Monthly Media Schedule"
            End
            Begin VB.Menu Mnu_Puchase_Order_Other 
               Caption         =   "Purchase Order"
            End
            Begin VB.Menu Mnu_Cancel_Order_Other 
               Caption         =   "Cancel Order"
            End
         End
         Begin VB.Menu Garis22 
            Caption         =   "-"
         End
         Begin VB.Menu Mnu_Travel_Expenses 
            Caption         =   "Additional Cost"
            Begin VB.Menu Mnu_Add_Cost_Quot 
               Caption         =   "Additional Cost Quotation"
            End
            Begin VB.Menu Mnu_Add_Cost_PO 
               Caption         =   "Additional Cost PO"
            End
            Begin VB.Menu Mnu_Add_Cost_CO 
               Caption         =   "Additional Cost CO"
            End
         End
         Begin VB.Menu Garis55 
            Caption         =   "-"
         End
         Begin VB.Menu Mnu_Client_PO 
            Caption         =   "Client PO"
            Begin VB.Menu Mnu_Client_Purchase_Order 
               Caption         =   "Client Purchase Order"
            End
            Begin VB.Menu Mnu_Client_Purchase_Order_line 
               Caption         =   "Client Purchase Order Line"
            End
            Begin VB.Menu Mnu_Client_Purchase_Order_by_Client 
               Caption         =   "Client Purchase Order by Client"
            End
            Begin VB.Menu Mnu_Client_Purchase_Order_BC 
               Caption         =   "Client Purchase Order Budget Control"
            End
            Begin VB.Menu Mnu_client_PO_jobNo_detail 
               Caption         =   "Client Purchase Order Job Number Detail"
            End
         End
      End
      Begin VB.Menu Garis5 
         Caption         =   "-"
      End
      Begin VB.Menu Mnu_Budget_Control 
         Caption         =   "Budget Control"
         Begin VB.Menu Mnu_View_Budget_Control 
            Caption         =   "Status of Budget BU-1"
         End
         Begin VB.Menu Mnu_Budget_Control_Report 
            Caption         =   "Budget Control BU-1"
         End
         Begin VB.Menu Garis17 
            Caption         =   "-"
         End
         Begin VB.Menu Mnu_New_SOB_BU_2 
            Caption         =   "New Status of Budget BU-2"
         End
         Begin VB.Menu Mnu_New_BC_BU_2 
            Caption         =   "New Budget Control BU-2"
         End
         Begin VB.Menu Garis41 
            Caption         =   "-"
         End
         Begin VB.Menu Mnu_New_SOB_BU_2_Monthly 
            Caption         =   "New Status of Budget BU-2 - Monthly "
         End
         Begin VB.Menu Mnu_New_BC_BU_2_Monthly 
            Caption         =   "New Budget Control BU-2 - Monthly"
         End
         Begin VB.Menu Garis411 
            Caption         =   "-"
            Visible         =   0   'False
         End
         Begin VB.Menu Mnu_Status_Of_Budget_Non_ULI 
            Caption         =   "Status of Budget BU-2"
            Visible         =   0   'False
         End
         Begin VB.Menu Mnu_Budget_Control_Non_ULI 
            Caption         =   "Budget Control BU-2"
            Visible         =   0   'False
         End
      End
      Begin VB.Menu Garis7 
         Caption         =   "-"
      End
      Begin VB.Menu Mnu_Rate_Card 
         Caption         =   "&Rate Card"
         Begin VB.Menu Mnu_Rate_Program_Television 
            Caption         =   "Television"
            Begin VB.Menu Mnu_New_TV_Prg_Catalog 
               Caption         =   "New TV Program Catalog"
            End
            Begin VB.Menu Mnu_New_TV_Prg_Rate 
               Caption         =   "New TV Program Rate"
            End
            Begin VB.Menu Mnu_New_TV_Discount 
               Caption         =   "Cash Discount"
            End
            Begin VB.Menu Mnu_New_TV_Bonus 
               Caption         =   "Bonus Scheme"
            End
            Begin VB.Menu Mnu_New_TV_Rate_by_Client 
               Caption         =   "Generate TV Rate Card by Client"
            End
            Begin VB.Menu Mnu_New_TV_Rate_by_Client_Download 
               Caption         =   "Download TV Rate Card by Client"
            End
            Begin VB.Menu Garis42 
               Caption         =   "-"
            End
            Begin VB.Menu Mnu_New_TV_CPRP_Catalog 
               Caption         =   "CPRP Catalog"
            End
            Begin VB.Menu Garis18 
               Caption         =   "-"
            End
            Begin VB.Menu Mnu_ACNielsen_Movie_Code 
               Caption         =   "ACNielsen Movie Code"
            End
         End
         Begin VB.Menu Mnu_Rate_Program_Radio 
            Caption         =   "Radio"
            Begin VB.Menu Sts_Rate 
               Caption         =   "Station Rate"
            End
         End
         Begin VB.Menu Mnu_Rate_Program_Print 
            Caption         =   "Print"
            Begin VB.Menu mnu_Print_Rate 
               Caption         =   "Print Rate"
            End
         End
         Begin VB.Menu Mnu_Rate_Program_Others 
            Caption         =   "Others"
            Begin VB.Menu Mnu_Cinema_Rate 
               Caption         =   "Cinema Rate"
            End
         End
      End
      Begin VB.Menu Garis6 
         Caption         =   "-"
      End
      Begin VB.Menu mnu_Report_Media 
         Caption         =   "Report"
         Begin VB.Menu mnu_Report_Media_TV 
            Caption         =   "TV"
            Begin VB.Menu Mnu_New_Tv_Billing_Report_By_Station 
               Caption         =   "Tv Billing Report By Station"
            End
            Begin VB.Menu Mnu_New_Tv_Billing_Report_By_Brand 
               Caption         =   "Tv Billing Report By Brand"
            End
            Begin VB.Menu Mnu_Report_Tv_Actual_vs_Budget_New 
               Caption         =   "TV Actual vs Plan Budget Report by Brand"
            End
            Begin VB.Menu Mnu_Summary_Blocking_New 
               Caption         =   "Summary Blocking"
            End
         End
         Begin VB.Menu mnu_Report_Media_radio 
            Caption         =   "Radio"
            Begin VB.Menu Mnu_Radio_Billing_Report_By_Station 
               Caption         =   "Radio Billing Report By Station"
            End
            Begin VB.Menu Mnu_Radio_Billing_Report_By_Station_By_Brand 
               Caption         =   "Radio Billing Report By Station By Brand"
            End
         End
         Begin VB.Menu mnu_Report_Media_Print 
            Caption         =   "Print"
            Begin VB.Menu mnu_Print_Buying_Activities 
               Caption         =   "Print Buying Activities"
            End
            Begin VB.Menu Mnu_Media_Owners_Draft 
               Caption         =   "Media Owners Draft"
            End
            Begin VB.Menu Mnu_Media_Print_Insertion 
               Caption         =   "Media Print Insertion"
            End
            Begin VB.Menu Mnu_Print_Billing_Report_By_Media 
               Caption         =   "Print Billing Report By Media"
            End
            Begin VB.Menu Mnu_Print_Billing_Report_By_Brand 
               Caption         =   "Print Billing Report By Brand"
            End
         End
         Begin VB.Menu mnu_Report_Media_Others 
            Caption         =   "Others"
            Begin VB.Menu Mnu_Other_Billing_Report_By_Supplier 
               Caption         =   "Other Billing Report By Supplier"
            End
            Begin VB.Menu Mnu_Other_Billing_Report_By_Brand 
               Caption         =   "Other Billing Report By Brand"
            End
         End
         Begin VB.Menu Garis37 
            Caption         =   "-"
         End
         Begin VB.Menu Mnu_New_Media_Monthly_Billing 
            Caption         =   "New Media Monthly Billing"
         End
         Begin VB.Menu Mnu_Media_Monthly_Billing 
            Caption         =   "Media Monthly Billing"
         End
         Begin VB.Menu Garis8 
            Caption         =   "-"
         End
         Begin VB.Menu Mnu_Billing_Report_by_Month 
            Caption         =   "Billing Report by Month BU-1"
         End
         Begin VB.Menu Mnu_Billing_Report_by_Medium 
            Caption         =   "Billing Report by Medium BU-1"
         End
         Begin VB.Menu Garis24 
            Caption         =   "-"
         End
         Begin VB.Menu Mnu_New_Billing_Report_by_Month_BU2 
            Caption         =   "New Billing Report By Month BU-2"
         End
         Begin VB.Menu Mnu_New_Billing_Report_by_Medium_BU2 
            Caption         =   "New Billing Report by Medium BU-2"
         End
      End
      Begin VB.Menu Garis14 
         Caption         =   "-"
         Visible         =   0   'False
      End
      Begin VB.Menu Mnu_Administrator_Media 
         Caption         =   "Administrator"
         Begin VB.Menu Mnu_Week_Commencing 
            Caption         =   "Entry Week Commencing"
         End
      End
      Begin VB.Menu Garis25 
         Caption         =   "-"
      End
      Begin VB.Menu Mnu_Media_Kit_New 
         Caption         =   "Media Kit"
      End
      Begin VB.Menu Garis39 
         Caption         =   "-"
      End
      Begin VB.Menu Mnu_Week_Commencing_View 
         Caption         =   "Week Commencing"
      End
   End
   Begin VB.Menu mnu_Finance 
      Caption         =   "&Finance"
      Visible         =   0   'False
   End
   Begin VB.Menu mnucatalog 
      Caption         =   "&Catalog"
      Begin VB.Menu mnucatalog_general 
         Caption         =   "&General"
         Begin VB.Menu mnucatalog_Client 
            Caption         =   "Client"
         End
         Begin VB.Menu mnucatalogBrang 
            Caption         =   "Brand"
         End
         Begin VB.Menu mnuProducts 
            Caption         =   "Product"
         End
         Begin VB.Menu mnucatalog_Brandcategory 
            Caption         =   "Product Category"
         End
         Begin VB.Menu Garis9 
            Caption         =   "-"
         End
         Begin VB.Menu mnucatalog_supplier 
            Caption         =   "Supplier"
         End
         Begin VB.Menu Garis10 
            Caption         =   "-"
         End
         Begin VB.Menu mnucatalog_media_type 
            Caption         =   "Media Type"
         End
         Begin VB.Menu Garis16 
            Caption         =   "-"
         End
         Begin VB.Menu Mnu_Province_City 
            Caption         =   "Province and City"
         End
      End
      Begin VB.Menu Garis11 
         Caption         =   "-"
      End
      Begin VB.Menu mnucatalog_Media 
         Caption         =   "Media"
         Begin VB.Menu mnucatalog_television 
            Caption         =   "Television"
         End
         Begin VB.Menu mnucatalog_radio 
            Caption         =   "Radio"
         End
         Begin VB.Menu mnucatalog_print_ad 
            Caption         =   "Print Ad"
            Begin VB.Menu Mnu_Media_Print 
               Caption         =   "Media Print"
            End
            Begin VB.Menu Mnu_Print_CPS 
               Caption         =   "Color, Paper and Size Catalog"
            End
         End
         Begin VB.Menu mnucatalog_others 
            Caption         =   "Others (Cinema)"
         End
      End
      Begin VB.Menu Garis50 
         Caption         =   "-"
      End
      Begin VB.Menu Mnu_Audit_Trail_View 
         Caption         =   "Audit Trail View"
      End
   End
   Begin VB.Menu mnu_My_Menu 
      Caption         =   "My Menu"
      Begin VB.Menu Mnu_Check_My_Job 
         Caption         =   "Check My Job"
      End
      Begin VB.Menu Mnu_Change_Password 
         Caption         =   "Change Password"
      End
      Begin VB.Menu mnu_Exit 
         Caption         =   "&Exit"
      End
   End
   Begin VB.Menu Mnu_Utility 
      Caption         =   "Utility"
      Begin VB.Menu Mnu_Access_Right 
         Caption         =   "User Access Right"
      End
      Begin VB.Menu Mnu_Set_Security_Media 
         Caption         =   "Set Security Media"
      End
      Begin VB.Menu Mnu_Administrator_Catalog 
         Caption         =   "Administrator Catalog"
      End
      Begin VB.Menu Garis15 
         Caption         =   "-"
      End
      Begin VB.Menu Mnu_I_Quest_Usr_Mgr 
         Caption         =   "I-Quest User Manager"
      End
      Begin VB.Menu Garis27 
         Caption         =   "-"
         Visible         =   0   'False
      End
      Begin VB.Menu Mnu_Setup 
         Caption         =   "Setup"
      End
   End
   Begin VB.Menu LogOff 
      Caption         =   "&Log Off"
      Enabled         =   0   'False
      Visible         =   0   'False
   End
   Begin VB.Menu Mnu_Help 
      Caption         =   "&Help"
      Begin VB.Menu Mnu_Contents 
         Caption         =   "Contents"
      End
      Begin VB.Menu Garis12 
         Caption         =   "-"
      End
      Begin VB.Menu Mnu_About_Erp 
         Caption         =   "About "
      End
      Begin VB.Menu Garis13 
         Caption         =   "-"
         Visible         =   0   'False
      End
      Begin VB.Menu Mnu_About_Us 
         Caption         =   "About Us"
         Visible         =   0   'False
      End
   End
End
Attribute VB_Name = "mdi_Main"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'<CSCC>
'********************************************************************************
'Submodul Name      : mdi_Main
'Submodul Function  : {MemberName}
'Used Table         : -
'Used SP            : -
'Procedure/Function : -
'Programmer Name    : Tedi
'Date               : 3/8/2016-1:05:14 AM
'Last Update By     : Tedi
'Date Update        : 3/8/2016-1:05:14 AM
'Log Update/By      : -
'********************************************************************************
'</CSCC>
'************************************************************
'Aplication Name    : MEDIA
'Version            : 2.0.0
'*************************************************************
Option Explicit

Const ID_THEME_OFFICE2000 = 140
Const ID_THEME_OFFICE2003 = 141
Const ID_THEME_NATIVE = 142
Const ID_THEME_OFFICE2000_PLAIN = 143
Const ID_THEME_OFFICEXP_PLAIN = 144
Const ID_THEME_OFFICE2003_PLAIN = 145
Const ID_THEME_NATIVE_PLAIN = 146

Const ID_TASKITEM_ADMIN = 1
Const ID_MEDIA_PLAN = 10
Const ID_TASKITEM_SEARCH = 3
Const ID_TASK_REPORT_BY_CLIENT = 19
Const ID_TASK_REPORT_BY_BRAND = 5
Const ID_TASK_REPORT_ACTUAL = 6
'Frm_Rpt_Approval_Timesheet
Const ID_TASK_REPORT_APPROVAL = 7
'Frm_Rpt_Unapprove_Timesheet
Const ID_TASK_REPORT_NOT_APPROVAL = 8
Const ID_TASK_REPORT_MISS_TIME = 9
Const ID_TASK_REPORT_EXPORT = 10
Const ID_TASK_MONTHLYCLOSING = 12
Const ID_TASK_GENERALDATE = 13
Const ID_TASK_TIMESHEETCALC = 14
Const ID_TASK_CATALOG = 15
Const ID_CAT_DIVISION = 0
Const ID_CAT_TASK = 1
Const ID_CAT_TITLE = 2

Dim ID_TASK_FAVOURITE As Integer '= 11
Dim arrCatalog() As String
Const FCONTROL = 8
Dim intDefaultTab As Integer
Public frmDefault As Form
Dim strSql As String

Private Sub Cancel_Order_Print_Click()
    Frm_Print_Cancel_PO.Show 1
End Sub

Private Sub LogOff_Click()
    FrmLogin.Show vbModal
End Sub

Private Sub MDIForm_Load()
'<CSCM>
'********************************************************************************
'Procedure Name     : MDIForm_Load
'Procedure Function : ---
'Input Parameter    : ---
'Output Parameter   : ---
'Date               : 3/8/2016
'LastUpdate/By      : Tedi / Kreatif
'********************************************************************************
'</CSCM>
    
    'Set Status Bar info
    StatusBar1.Panels(1).Text = FullName & " --> [" & Server_Name & " : " & Database_Name & " ]"
    
    Me.Caption = StrCompany
       
    Load_Week_Commencing 'Load Week Commencing Table
    
    'Load TV Program Rate Period
    Load_TV_Program_Period
    
    'MsgBox Get_ProgramRatePeriod(" TV", "M")
    
    Setting_Menu 'Set Menu for Current User
    
    '------------------ Load Task for Current USer ---------------------
    Dim Rs_Pop As New ADODB.Recordset
    
    strSql = " SELECT user_name FROM Current_user_job WHERE user_name= '" & UserName & "' "
    strSql = strSql & " AND New_Message=1"
    
    Rs_Pop.Open strSql, Conn, adOpenStatic, adLockReadOnly
    If Not Rs_Pop.EOF And Not Rs_Pop.BOF Then
        'show frm_pop_up_menu
        frm_Pop_Up_Menu.Show 1
    End If
    Rs_Pop.Close
    Set Rs_Pop = Nothing
    '-------------------------------------------------------------------
    StatusBar1.Panels(1).Text = VarPub.Login_FullName & " - [" & VarPub.Server_Name & " ]"
    CreateTaskPanel

End Sub


Private Sub Setting_Menu()
    Dim rsAccessRight As New ADODB.Recordset
    Dim Ctl As Control
    Dim sCtlType As String
    Dim Menu_Name As String
    Dim Index
            
    strSql = ""
    strSql = strSql & " select Menu_Media_AccessRight.Menu_Id, Menu_Name from Menu_Media_AccessRight "
    strSql = strSql & " INNER JOIN Menu_Media ON"
    strSql = strSql & " Menu_Media_AccessRight.Menu_Id = Menu_Media.Menu_Id Where "
    strSql = strSql & " User_Name = '" & UserName & "'"
    
    rsAccessRight.Open strSql, Conn, adOpenStatic, adLockReadOnly, adCmdText

    For Each Ctl In Mdi_Master.Controls
        If TypeOf Ctl Is Menu Then
            If UCase(Left(Ctl.Name, 5)) <> "GARIS" Then
                Ctl.Enabled = False
            End If
        End If
    Next
    
    If Not rsAccessRight.EOF Or Not rsAccessRight.BOF Then
        'Loop Valid Menu
        Do While Not rsAccessRight.EOF
            'Loop Menu
            For Each Ctl In Mdi_Master.Controls
                If UCase(Ctl.Name) = UCase(rsAccessRight.Fields("Menu_name").Value) Then
                        Ctl.Enabled = True
                End If
            Next
            'Next Valid Menu
            rsAccessRight.MoveNext
        Loop
    End If
    
    rsAccessRight.Close
    
End Sub
Private Sub MDIForm_Unload(Cancel As Integer)
    Dim rs As ADODB.Recordset
            
    'Update Last Lof Off
    Set rs = New ADODB.Recordset
    rs.Open "SELECT * FROM User_ID WHERE user_name='" & UserName & "'", Conn, adOpenKeyset, adLockOptimistic, adCmdText
    rs_Date.Requery
                    
    'Last Log Out
    rs("Last_Logout").Value = rs_Date(0).Value
    If Not Is_SkipMultiSesion Then
        rs("Computer_Name").Value = Null
    End If
    rs.Update
    
    'close Recordset
    rs.Close
    Set rs = Nothing
    
    'Close Connection
    Conn.Close
    Set Conn = Nothing
    
    'End Application
    End
End Sub

Private Sub Mnu_About_Erp_Click()
   FrmAbout.Show 1
   
End Sub

Private Sub Mnu_About_Us_Click()
       Frm_About_Us.Show 1
End Sub

Private Sub Mnu_Access_Right_Click()
    Frm_User_Access_Right.Show 1
End Sub

Private Sub Mnu_ACNielsen_Movie_Code_Click()
    Frm_Movie_Code.Show 1
End Sub

Private Sub Mnu_Add_Cost_CO_Click()
    Frm_CO_Add_Cost.Show 1
End Sub

Private Sub Mnu_Add_Cost_PO_Click()
    Frm_PO_Add_Cost.Show 1
End Sub

Private Sub Mnu_Add_Cost_Quot_Click()
    Frm_Travel_Expenses.Show 1
End Sub

Private Sub Mnu_Administrator_Catalog_Click()
    Frm_Administator_Catalog.Show 1
End Sub

Private Sub Mnu_Audit_Trail_View_Click()
    frm_audit_trail_export.Show 1
End Sub


Private Sub Mnu_Billing_Report_by_Medium_Click()
    Frm_Report_Billing_By_Medium.Show 1
End Sub



Private Sub Mnu_Billing_Report_by_Month_Click()
    Frm_Report_Billing_by_Month.Show 1
End Sub





Private Sub Mnu_Budget_Control_Non_ULI_Click()
    Frm_Budget_Control_Media_Non_Uli.Show 1
End Sub

Private Sub Mnu_Budget_Control_Report_Click()
    Frm_Budget_Control_Media.Show 1
End Sub





Private Sub Mnu_Cancel_Order_Other_Click()
    Frm_PO_Other_Cancel.Show 1
End Sub

Private Sub Mnu_Change_Password_Click()
    Frm_Change_Password.Show 1
End Sub


Private Sub Mnu_Check_My_Job_Click()
    frm_Pop_Up_Menu.Show 1
End Sub









Private Sub Mnu_Cinema_Rate_Click()
    Frm_Cinema_Catalog.Show 1
End Sub



Private Sub Mnu_Client_Brief_Media_Click()
    Frm_Client_Brief_Media.Show 1
End Sub

Private Sub Mnu_client_PO_jobNo_detail_Click()
    Frm_client_PO_jobNo_detail.Show 1
End Sub

Private Sub Mnu_Client_Purchase_Order_BC_Click()
    frm_Client_PO_Budget_Control.Show 1
End Sub

Private Sub Mnu_Client_Purchase_Order_by_Client_Click()
    frm_Client_Purchase_Order_by_Client.Show 1
End Sub

Private Sub Mnu_Client_Purchase_Order_Click()
    'frm_Client_Purchase_Order_by_Client.show 1
    frm_Client_Purchase_Order.Show 1
End Sub

Private Sub Mnu_Client_Purchase_Order_line_Click()
    frm_client_purchase_order_Line.Show 1
End Sub




Private Sub Mnu_Contents_Click()
    SendKeys "{F1}"
   
End Sub









Private Sub mnu_Exit_Click()
    Unload Me
    End
End Sub




Private Sub Mnu_I_B_Quotation_Print_Click()
    Frm_Print_Media_Quotation.Show 1
End Sub

Private Sub Mnu_I_B_Quotation_Radio_Click()
    Frm_Radio_Media_Quot.Show vbModal
End Sub



Private Sub Mnu_I_Quest_Usr_Mgr_Click()
    frm_iquest_user_manager.Show 1
End Sub

Private Sub Mnu_IB_Others_Click()
    Frm_IB_Others.Show 1
End Sub

Private Sub Mnu_IB_Print_Click()
    Frm_IB_Print.Show 1
End Sub

Private Sub Mnu_IB_Radio_Click()
    Frm_IB_Radio.Show vbModal
End Sub

Private Sub Mnu_IB_Television_Click()
   
    Frm_New_IB_TV.Show 1 'Show IB TV form
    
End Sub


Private Sub Mnu_Media_Kit_New_Click()
    frm_Media_Kit_New.Show 1
End Sub


Private Sub Mnu_Media_Owners_Draft_Click()
    Frm_Print_Media_Draft.Show 1
End Sub

Private Sub Mnu_Media_Plan_Click()
    Dim pesan As Integer
    
    If Screen.Width < 15360 And Screen.Height < 11520 Then
        pesan = MsgBox("Recomended viewed in 1024 X 768 or Higher Screen Resolution." & vbCrLf & "Click OK to Continue, otherwise Click Cancel", vbOKCancel + vbInformation)
        If pesan = 2 Then
            Exit Sub
        End If
    End If
    
    FrmMPInsertion.Show 1
    
End Sub

Private Sub Mnu_Media_Print_Click()
    Frm_Print_Catalog.Show 1
End Sub

Private Sub Mnu_Media_Print_Insertion_Click()
    Frm_Report_Insertion_Media_Print.Show 1
End Sub



Private Sub Mnu_New_BC_BU_2_Click()
    Frm_New_BC_BU_2.Show 1
End Sub

Private Sub Mnu_New_BC_BU_2_Monthly_Click()
    Frm_New_BC_BU_2_Monthly.Show 1
End Sub

Private Sub Mnu_New_Billing_Report_by_Medium_BU2_Click()
    Frm_Report_Billing_By_Medium_BU2.Show 1
End Sub

Private Sub Mnu_New_Billing_Report_by_Month_BU2_Click()
    Frm_Report_Billing_By_Month_BU2.Show 1
End Sub

Private Sub Mnu_New_Media_Monthly_Billing_Click()
    Frm_Report_Media_Monthy_Billing_New.Show 1
End Sub

Private Sub Mnu_New_SOB_BU_2_Click()
    Frm_New_SOB_BU2.Show 1
End Sub

Private Sub Mnu_New_SOB_BU_2_Monthly_Click()
    Frm_New_SOB_BU2_monthly.Show 1
End Sub



Private Sub Mnu_New_Tv_Billing_Report_By_Brand_Click()
    Frm_Report_Billing_TV_By_Brand_New.Show 1
End Sub

Private Sub Mnu_New_Tv_Billing_Report_By_Station_Click()
    Frm_Report_Billing_By_TV_Station_New.Show 1
End Sub

Private Sub Mnu_New_TV_Bonus_Click()
    frm_tv_Rate_Card_Bonus_Scheme.Show 1
End Sub





Private Sub Mnu_New_TV_CPRP_Catalog_Click()
    Frm_TV_CPRP_Catalog.Show 1
End Sub

Private Sub Mnu_New_TV_Discount_Click()
    'frm_TV_RateCard_Scheme.show 1
        frm_TV_RateCard_Cash_Discount.Show 1
End Sub

Private Sub Mnu_New_TV_Finalize_PO_Click()
    Frm_New_TV_Schedule_Master_Finalize.Show 1
End Sub

Private Sub Mnu_New_TV_Implementation_Click()
    'Resize Screen
    LScreenWidth = Screen.Width / Screen.TwipsPerPixelX
    LScreenHeight = Screen.Height / Screen.TwipsPerPixelY
    
    'resize screen
    'ChangeRes 1024, 768
    
    'Load Form
    Frm_New_TV_Schedule_Master.Show 1
End Sub

Private Sub Mnu_New_TV_Material_Click()
    Frm_TV_Material_Catalog.Show 1
End Sub

Private Sub Mnu_New_TV_MQ_Click()
    Frm_Television_Media_Quotation.Show 1
End Sub

Private Sub Mnu_New_TV_Preemption_log_Click()
    Frm_New_TV_Preemption_Log.Show 1
End Sub

Private Sub Mnu_New_TV_Prg_Catalog_Click()
    'Frm_TV_Program_New.show 1
    Frm_TV_RateCard_Program.Show 1
End Sub

Private Sub Mnu_New_TV_Prg_Rate_Click()
    'Frm_TV_Program_Rate_New.show 1
    frm_TV_RateCard.Show 1
End Sub

Private Sub Mnu_New_TV_Rate_by_Client_Click()
    frm_TV_RateCard_Generate_Per_Client.Show 1
End Sub

Private Sub Mnu_New_TV_Rate_by_Client_Download_Click()
    frm_TV_RateCard_Download_Per_Client.Show 1
End Sub

Private Sub Mnu_New_TV_Schedule_Org_Click()
    'Resize Screen
    LScreenWidth = Screen.Width / Screen.TwipsPerPixelX
    LScreenHeight = Screen.Height / Screen.TwipsPerPixelY
    
    'resize screen
    'ChangeRes 1024, 768
    
    'Load Form
    Frm_New_TV_Schedule_Master_Original.Show 1
End Sub

Private Sub Mnu_Other_Billing_Report_By_Brand_Click()
    Frm_Report_Other_Billing_by_Brand.Show 1
End Sub

Private Sub Mnu_Other_Billing_Report_By_Supplier_Click()
    Frm_Report_Other_Billing_by_Media.Show 1
End Sub

Private Sub Mnu_Other_IB_Quotation_Click()
    Frm_Other_Media_Quotation.Show 1
End Sub

Private Sub mnu_Other_Monthly_Media_Schedule_Click()
    Frm_Other_Schedule.Show 1
End Sub











Private Sub mnu_PO_Print_Click()
    Frm_PO_Print.Show 1
End Sub



Private Sub Mnu_Print_Billing_Report_By_Brand_Click()
    Frm_Report_Print_Billing_by_Brand.Show 1
End Sub

Private Sub mnu_Print_Buying_Activities_Click()
    Frm_Print_Media_Buying_Activities.Show 1
End Sub

Private Sub Mnu_Print_CPS_Click()
    Frm_Print_PSC_Catalog.Show 1
End Sub

Private Sub mnu_Print_Rate_Click()
    Frm_Print_Rate.Show 1
End Sub

Private Sub Mnu_Print_Schedule_Click()
    Frm_Print_schedule.Show 1
End Sub









Private Sub Mnu_Province_City_Click()
    Frm_Area_City_Catalog.Show 1
End Sub

Private Sub Mnu_Puchase_Order_Other_Click()
    Frm_PO_Other.Show 1
End Sub

Private Sub Mnu_Radio_Address_Lable_Per_Job_Click()
    frm_Radio_Station_Address_From_Job_Number.Show 1
End Sub

Private Sub Mnu_Radio_Billing_Report_By_Station_By_Brand_Click()
    Frm_Radio_Billing_Report_By_Station_By_Brand.Show 1
End Sub

Private Sub Mnu_Radio_Billing_Report_By_Station_Click()
    Frm_Radio_Billing_Report_By_Station.Show 1
End Sub

Private Sub Mnu_Radio_Cancelation_Order_Click()
    Dim Cek_User As Boolean
    
    Cek_User = User_Valid("Implementor")
    
    If Not Cek_User Then
        Cek_User = User_Valid("Admin")
    End If
    
    If Cek_User = True Then
        Frm_Radio_Cancelation_Order.Show vbModal
    Else
        MsgBox "You don't have any access to this menu", vbCritical, StrCompany
    End If
End Sub

Private Sub Mnu_Radio_Purcashe_Order_Click()

    Dim Cek_User As Boolean
    
    Cek_User = User_Valid("Implementor")
    
    If Not Cek_User Then
        Cek_User = User_Valid("Admin")
    End If
    
    If Cek_User = True Then
        Frm_PO_Media_Radio.Show vbModal
    Else
        MsgBox "You don't have any access to this menu", vbCritical, StrCompany
    End If

End Sub

Private Sub Mnu_Radio_Schedule_Click()
    If User_Valid("Implementor") = True Then
        Frm_Radio_Media_Quotation.Show vbModal
    Else
        MsgBox "You don't have any access to this menu", vbCritical, StrCompany
    End If
End Sub

Private Sub Mnu_Radio_Station_Address_Lable_Click()
    frm_Radio_Station_Address.Show 1
End Sub

Private Sub mnu_Replace_Order_Print_Click()
    Frm_Print_Replace_order.Show 1
End Sub





Private Sub Mnu_Report_Tv_Actual_vs_Budget_New_Click()
    frm_Report_Tv_Actual_vs_Budget_New.Show
End Sub

Private Sub mnu_Retainer_Fee_Entry_Click()
    Frm_Retainer_Fee_Entry.Show 1
End Sub



Private Sub Mnu_Set_Security_Media_Click()
    'frm_SecurityMedia.show 1
    frm_securityMedia_by_client.Show 1
End Sub

Private Sub Mnu_Setup_Click()
    Frm_Setup_Media_Parameter.Show 1
End Sub

Private Sub Mnu_Status_Of_Budget_Non_ULI_Click()
    Frm_View_Budget_Control_Non_Uli.Show 1
End Sub



Private Sub Mnu_Summary_Blocking_New_Click()
    Frm_Summary_Blocking_New.Show 1
End Sub



Private Sub Mnu_Transfer_TVS_Click()
    frm_New_TV_Schedule_Transfer_XLS.Show 1
End Sub





Private Sub Mnu_View_Budget_Control_Click()
    Frm_View_Budget_Control.Show 1
End Sub

Private Sub Mnu_View_Media_Plan_Click()
    Frm_MediaPlan_View.Show 1
End Sub

Private Sub Mnu_Week_Commencing_Click()
    Frm_Week_Commencing.Show 1
End Sub

Private Sub Mnu_Week_Commencing_View_Click()
    Frm_Week_Commencing_View.Show 1
End Sub

Private Sub mnucatalog_Brandcategory_Click()
    'Frm_Product_Category.show 1
    Frm_Brand_Category_Catalog.Show 1
End Sub

Private Sub mnucatalog_Client_Click()
    Frm_Client_Catalog.Show 1
End Sub

Private Sub mnucatalog_media_type_Click()
    Frm_Media_Type_Catalog.Show 1
End Sub

Private Sub mnucatalog_others_Click()
    Frm_Cinema_Catalog.Show 1
End Sub

Private Sub mnucatalog_radio_Click()
    Frm_Radio_Station_Catalog.Show vbModal
End Sub

Private Sub mnucatalog_supplier_Click()
    Frm_Supplier_Catalog.Show 1
End Sub

Private Sub mnucatalog_television_Click()
    Frm_TV_Catalog.Show 1
End Sub

Private Sub mnucatalogBrang_Click()
    frm_Brand_Catalog.Show 1
End Sub

Private Sub mnuProducts_Click()
    Frm_Brand_Product.Show 1
End Sub

Private Sub mnu_Print_Billing_Report_By_Media_Click()
    Frm_Report_Print_Billing_by_Media.Show 1
End Sub

Private Sub Sts_Rate_Click()
    Frm_Radio_Rate.Show vbModal
End Sub


Private Function Has_New_Message() As Boolean
    Dim rsCek_Message As New ADODB.Recordset
    
    strSql = "SELECT * FROM Current_User_Job WHERE User_Name='" & UserName & "' AND New_Message=1"
    
    rsCek_Message.Open strSql, Conn, adOpenStatic, adLockReadOnly
    
    If Not (rsCek_Message.EOF And rsCek_Message.BOF) Then
        Has_New_Message = True
    Else
        Has_New_Message = False
    End If
    
    rsCek_Message.Close
    Set rsCek_Message = Nothing
    
End Function



Private Sub Tmr_Cek_Task_Timer()
    Tmr_Cek_Task.Enabled = False

    If Has_New_Message Then
        
        Frm_Have_New_Task.Show 1

        'Update task yang sudah ditampilkan ke 0
        strSql = "UPDATE Current_User_Job SET New_Message=0 WHERE User_Name='" & UserName & "' AND New_Message=1"

        Conn.Execute strSql
    End If

    Tmr_Cek_Task.Enabled = True

End Sub


Public Sub CreateTaskPanel()
'*****************************************
'Procedure Name     : CreateTaskPanel
'Procedure Function : Generate Task Panel /Side Bar
'Input Parameter    : -
'Output Parameter   : -
'*****************************************
    Dim Group As TaskPanelGroup
    Dim Item As TaskPanelGroupItem
    Dim rst_Temp As New ADODB.Recordset
    Dim str_Sql As String
    Dim strComposite As String
    Dim xItem As TaskPanelItem
    Dim strSecure As String
    Dim int_Count As Integer
    str_Sql = "select link_view from user_shortcut WHERE user_id='" & VarPub.Login_User & "' AND group_view='1' AND module_id='4' "
     rst_Temp.Open str_Sql, VarPub.ConnERP, adOpenForwardOnly, adLockReadOnly, adCmdText: str_Sql = ""
     If Not rst_Temp.EOF Then
        strComposite = rst_Temp!link_view
     Else
        strComposite = ""
     End If
     Call CloseRecordset(rst_Temp)
     
     wndTaskPanel.Groups.Clear
     
     str_Sql = "select * from User_Menu_Link where Modul_id=4 AND shownlink=1 order by Modul_ID,ParentID,oderby" 'ParentID,oderby"
     'MsgBox InputBox("", "", str_Sql)
     rst_Temp.Open str_Sql, VarPub.ConnERP, adOpenForwardOnly, adLockReadOnly, adCmdText: str_Sql = ""
     int_Count = 0
     While Not rst_Temp.EOF
        int_Count = int_Count + 1
        If rst_Temp!link_id = rst_Temp!ParentID Then
            If CanSee(VarPub.Login_User, Trim(rst_Temp!LinkName)) = True Then
                Set Group = wndTaskPanel.Groups.Add(0, Trim(rst_Temp!LinkName))
                Group.Tag = rst_Temp!link_id
            End If
        Else

            If Mid(strComposite, rst_Temp!Icon, 1) = "1" Then
                If CanSee(VarPub.Login_User, Trim(rst_Temp!LinkName)) = True Then
                    Set xItem = Group.Items.Add(rst_Temp!link_id, Trim(rst_Temp!LinkName), xtpTaskItemTypeLink, rst_Temp!Icon)
                    xItem.Tag = rst_Temp!link_id
                End If
            End If
        End If
        rst_Temp.MoveNext
     Wend
    
    ID_TASK_FAVOURITE = int_Count
    Set Item = Group.Items.Add(ID_TASK_FAVOURITE, "Favourite View", xtpTaskItemTypeLink, 17)
    Call CloseRecordset(rst_Temp)
    Item.Tooltip = "Customize Favourite Shortcut"

    wndTaskPanel.SetImageList imlTaskPanelIcons
    wndTaskPanel.SetMargins 2, 5, 5, 5, 5
End Sub

