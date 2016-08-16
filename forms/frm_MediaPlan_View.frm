VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFlxGrd.ocx"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.Form frm_MediaPlan_View 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Media Plan"
   ClientHeight    =   9105
   ClientLeft      =   525
   ClientTop       =   2535
   ClientWidth     =   14730
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9105
   ScaleWidth      =   14730
   Begin VB.Frame FrameViewMonth1 
      BackColor       =   &H80000015&
      BorderStyle     =   0  'None
      Caption         =   "Frame3"
      Height          =   2760
      Left            =   600
      TabIndex        =   28
      Top             =   2040
      Visible         =   0   'False
      Width           =   2415
      Begin VB.Frame FrameViewMonth2 
         BorderStyle     =   0  'None
         Caption         =   "Frame3"
         Height          =   2685
         Left            =   30
         TabIndex        =   29
         Top             =   30
         Width           =   2355
         Begin VB.CommandButton cmdCancel 
            Caption         =   "Cancel"
            Height          =   360
            Left            =   1215
            TabIndex        =   44
            Top             =   2190
            Width           =   960
         End
         Begin VB.CommandButton cmdOK 
            Caption         =   "OK"
            Height          =   360
            Left            =   195
            TabIndex        =   43
            Top             =   2190
            Width           =   960
         End
         Begin VB.CheckBox cbkAllMonth 
            Caption         =   "ALL"
            Height          =   255
            Left            =   195
            TabIndex        =   42
            Top             =   150
            Width           =   1065
         End
         Begin VB.CheckBox cbkViewMonth 
            Caption         =   "December"
            Height          =   285
            Index           =   12
            Left            =   1140
            TabIndex        =   41
            Top             =   1830
            Width           =   1080
         End
         Begin VB.CheckBox cbkViewMonth 
            Caption         =   "November"
            Height          =   285
            Index           =   11
            Left            =   1140
            TabIndex        =   40
            Top             =   1545
            Width           =   1080
         End
         Begin VB.CheckBox cbkViewMonth 
            Caption         =   "October"
            Height          =   285
            Index           =   10
            Left            =   1140
            TabIndex        =   39
            Top             =   1275
            Width           =   1080
         End
         Begin VB.CheckBox cbkViewMonth 
            Caption         =   "September"
            Height          =   285
            Index           =   9
            Left            =   1140
            TabIndex        =   38
            Top             =   1005
            Width           =   1080
         End
         Begin VB.CheckBox cbkViewMonth 
            Caption         =   "August"
            Height          =   285
            Index           =   8
            Left            =   1140
            TabIndex        =   37
            Top             =   735
            Width           =   1080
         End
         Begin VB.CheckBox cbkViewMonth 
            Caption         =   "July"
            Height          =   285
            Index           =   7
            Left            =   1140
            TabIndex        =   36
            Top             =   480
            Width           =   1080
         End
         Begin VB.CheckBox cbkViewMonth 
            Caption         =   "June"
            Height          =   285
            Index           =   6
            Left            =   195
            TabIndex        =   35
            Top             =   1845
            Width           =   1080
         End
         Begin VB.CheckBox cbkViewMonth 
            Caption         =   "May"
            Height          =   285
            Index           =   5
            Left            =   195
            TabIndex        =   34
            Top             =   1560
            Width           =   1080
         End
         Begin VB.CheckBox cbkViewMonth 
            Caption         =   "April"
            Height          =   285
            Index           =   4
            Left            =   195
            TabIndex        =   33
            Top             =   1275
            Width           =   1080
         End
         Begin VB.CheckBox cbkViewMonth 
            Caption         =   "March"
            Height          =   285
            Index           =   3
            Left            =   195
            TabIndex        =   32
            Top             =   1005
            Width           =   1080
         End
         Begin VB.CheckBox cbkViewMonth 
            Caption         =   "February"
            Height          =   285
            Index           =   2
            Left            =   195
            TabIndex        =   31
            Top             =   735
            Width           =   1080
         End
         Begin VB.CheckBox cbkViewMonth 
            Caption         =   "January"
            Height          =   285
            Index           =   1
            Left            =   195
            TabIndex        =   30
            Top             =   480
            Width           =   1080
         End
      End
   End
   Begin VB.Frame FrameInsertion 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   8955
      Left            =   105
      TabIndex        =   1
      Top             =   0
      Width           =   14490
      Begin VB.PictureBox Picture9 
         Height          =   360
         Left            =   10170
         ScaleHeight     =   300
         ScaleWidth      =   3150
         TabIndex        =   49
         Top             =   7995
         Width           =   3210
         Begin VB.CommandButton cmdExportToExcel 
            Caption         =   "E&xport"
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Left            =   2295
            TabIndex        =   53
            ToolTipText     =   "Export Media Plan to MS Excel"
            Top             =   0
            Width           =   855
         End
         Begin VB.CommandButton cmdSummary 
            Caption         =   "Summary"
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
            Left            =   1275
            TabIndex        =   52
            ToolTipText     =   "View Summary of Budget"
            Top             =   0
            Width           =   1005
         End
         Begin VB.CommandButton cmdTVLayering 
            Caption         =   "TV Layering"
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
            TabIndex        =   50
            ToolTipText     =   "View TV Layering"
            Top             =   0
            Width           =   1260
         End
      End
      Begin VB.PictureBox Picture5 
         Height          =   360
         Left            =   13485
         ScaleHeight     =   300
         ScaleWidth      =   855
         TabIndex        =   24
         Top             =   7995
         Width           =   915
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
            Height          =   300
            Left            =   0
            TabIndex        =   25
            ToolTipText     =   "Close This Window"
            Top             =   0
            Width           =   855
         End
      End
      Begin VB.PictureBox Picture3 
         Appearance      =   0  'Flat
         ForeColor       =   &H80000008&
         Height          =   6780
         Left            =   135
         ScaleHeight     =   6750
         ScaleWidth      =   14190
         TabIndex        =   10
         Top             =   1065
         Width           =   14220
         Begin VB.Frame FrameProgressBar 
            Appearance      =   0  'Flat
            BackColor       =   &H80000002&
            BorderStyle     =   0  'None
            Caption         =   "FramePleaseWait"
            ForeColor       =   &H00FFFF00&
            Height          =   1290
            Left            =   5805
            TabIndex        =   13
            Top             =   2145
            Visible         =   0   'False
            Width           =   3510
            Begin VB.Frame FrameProgressBar1 
               BorderStyle     =   0  'None
               Caption         =   "FramePleaseWait"
               Height          =   1245
               Left            =   15
               TabIndex        =   14
               Top             =   30
               Width           =   3480
               Begin ComctlLib.ProgressBar ProgressBarExport 
                  Height          =   180
                  Left            =   75
                  TabIndex        =   15
                  Top             =   540
                  Width           =   3330
                  _ExtentX        =   5874
                  _ExtentY        =   318
                  _Version        =   327682
                  Appearance      =   0
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
                  TabIndex        =   17
                  Top             =   -30
                  Width           =   3480
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
                  TabIndex        =   16
                  Top             =   735
                  Width           =   4260
               End
            End
         End
         Begin VB.ListBox LstHidenCol 
            Height          =   1860
            Left            =   7065
            Style           =   1  'Checkbox
            TabIndex        =   12
            Top             =   210
            Visible         =   0   'False
            Width           =   1815
         End
         Begin MSFlexGridLib.MSFlexGrid FGMPInsertion 
            Height          =   6795
            Left            =   -30
            TabIndex        =   11
            Top             =   -30
            Width           =   14235
            _ExtentX        =   25109
            _ExtentY        =   11986
            _Version        =   393216
            Rows            =   5
            Cols            =   7
            FixedRows       =   4
            FixedCols       =   2
            BackColorFixed  =   12356167
            ForeColorFixed  =   -2147483640
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
      Begin VB.PictureBox Picture1 
         Appearance      =   0  'Flat
         ForeColor       =   &H80000008&
         Height          =   855
         Left            =   135
         ScaleHeight     =   825
         ScaleWidth      =   14190
         TabIndex        =   2
         Top             =   210
         Width           =   14220
         Begin VB.PictureBox picViewMonth 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   240
            Left            =   2535
            Picture         =   "frm_MediaPlan_View.frx":0000
            ScaleHeight     =   240
            ScaleWidth      =   225
            TabIndex        =   58
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
            TabIndex        =   47
            Top             =   495
            Width           =   2145
         End
         Begin VB.TextBox txtViewMonth 
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
            TabIndex        =   27
            Text            =   "ALL"
            Top             =   450
            Width           =   1575
         End
         Begin VB.TextBox txtLastUpdateBy 
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
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
            Height          =   195
            Left            =   9465
            Locked          =   -1  'True
            TabIndex        =   23
            Text            =   "txtLastUpdateBy"
            Top             =   495
            Width           =   1485
         End
         Begin VB.TextBox txtLastUpdateDate 
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
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
            Height          =   195
            Left            =   9465
            Locked          =   -1  'True
            TabIndex        =   22
            Text            =   "txtLastUpdateDate"
            Top             =   150
            Width           =   1485
         End
         Begin VB.TextBox txtCreatedBy 
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
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
            Height          =   195
            Left            =   6690
            Locked          =   -1  'True
            TabIndex        =   21
            Text            =   "txtCreatedBy"
            Top             =   495
            Width           =   1485
         End
         Begin VB.TextBox txtCreatedDate 
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
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
            Height          =   195
            Left            =   6690
            Locked          =   -1  'True
            TabIndex        =   20
            Text            =   "txtCreatedDate"
            Top             =   150
            Width           =   1485
         End
         Begin VB.TextBox txtClientName 
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
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
            Height          =   195
            Left            =   3540
            Locked          =   -1  'True
            TabIndex        =   19
            Text            =   "txtClientName"
            Top             =   495
            Width           =   2505
         End
         Begin VB.TextBox txtBrandName 
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
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
            Height          =   195
            Left            =   3540
            Locked          =   -1  'True
            TabIndex        =   18
            Text            =   "txtBrandName"
            Top             =   150
            Width           =   1560
         End
         Begin VB.ComboBox cboMPNumber 
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
            Height          =   315
            Left            =   1200
            Style           =   2  'Dropdown List
            TabIndex        =   0
            Top             =   90
            Width           =   1575
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
            TabIndex        =   48
            Top             =   165
            Width           =   1755
         End
         Begin VB.Label Label8 
            Alignment       =   1  'Right Justify
            Caption         =   "View :"
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
            Left            =   270
            TabIndex        =   26
            Top             =   465
            Width           =   885
         End
         Begin VB.Label Label6 
            Alignment       =   1  'Right Justify
            Caption         =   "Last Update :"
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
            Left            =   8235
            TabIndex        =   8
            Top             =   150
            Width           =   1170
         End
         Begin VB.Label Label7 
            Alignment       =   1  'Right Justify
            Caption         =   "By :"
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
            Left            =   8250
            TabIndex        =   9
            Top             =   480
            Width           =   1170
         End
         Begin VB.Label Label5 
            Alignment       =   1  'Right Justify
            Caption         =   "By :"
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
            Left            =   5325
            TabIndex        =   7
            Top             =   480
            Width           =   1305
         End
         Begin VB.Label Label4 
            Alignment       =   1  'Right Justify
            Caption         =   "Created :"
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
            Left            =   5325
            TabIndex        =   6
            Top             =   135
            Width           =   1305
         End
         Begin VB.Label Label3 
            Alignment       =   1  'Right Justify
            Caption         =   "Client :"
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
            Left            =   2835
            TabIndex        =   5
            Top             =   480
            Width           =   645
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            Caption         =   "MP Number :"
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
            Left            =   45
            TabIndex        =   4
            Top             =   120
            Width           =   1095
         End
         Begin VB.Label Label2 
            Alignment       =   1  'Right Justify
            Caption         =   "Brand :"
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
            Left            =   2835
            TabIndex        =   3
            Top             =   135
            Width           =   645
         End
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
         Left            =   7785
         TabIndex        =   57
         Top             =   8565
         Width           =   1215
      End
      Begin VB.Shape LegendActual 
         BackColor       =   &H00FFFF80&
         FillColor       =   &H0000C000&
         FillStyle       =   0  'Solid
         Height          =   165
         Left            =   7545
         Top             =   8580
         Width           =   180
      End
      Begin VB.Label lbl_show_user_server 
         Height          =   300
         Left            =   12585
         TabIndex        =   56
         Top             =   8475
         Width           =   825
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
         Left            =   6240
         TabIndex        =   55
         Top             =   8550
         Width           =   1290
      End
      Begin VB.Shape LegendTVRF 
         BackColor       =   &H00FFFF80&
         FillColor       =   &H0000FFFF&
         FillStyle       =   0  'Solid
         Height          =   165
         Left            =   6075
         Top             =   8565
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
         Left            =   150
         TabIndex        =   54
         Top             =   8535
         Width           =   810
      End
      Begin VB.Shape LegendWebApproval 
         BackColor       =   &H00FFFF80&
         FillColor       =   &H00FFC0FF&
         FillStyle       =   0  'Solid
         Height          =   165
         Left            =   2610
         Top             =   8565
         Width           =   180
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
         Left            =   2775
         TabIndex        =   51
         Top             =   8550
         Width           =   1680
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
         Left            =   1245
         TabIndex        =   46
         Top             =   8550
         Width           =   1230
      End
      Begin VB.Shape LegendApproval 
         BackColor       =   &H00FFFF80&
         FillColor       =   &H00FFFF80&
         FillStyle       =   0  'Solid
         Height          =   165
         Left            =   1050
         Top             =   8565
         Width           =   180
      End
      Begin VB.Shape Shape_unlocked 
         BackColor       =   &H00FFFF80&
         FillColor       =   &H000080FF&
         FillStyle       =   0  'Solid
         Height          =   165
         Left            =   4590
         Top             =   8565
         Width           =   180
      End
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
         Left            =   4740
         TabIndex        =   45
         Top             =   8550
         Width           =   1290
      End
   End
   Begin VB.Menu MnuFGMPInsertion 
      Caption         =   "MnuFGMPInsertion"
      Begin VB.Menu MnuFreeze 
         Caption         =   "Freeze Col"
      End
      Begin VB.Menu MnuUnFreeze 
         Caption         =   "UnFreeze Col"
      End
      Begin VB.Menu mnuHideCol 
         Caption         =   "Hide Col"
      End
      Begin VB.Menu mnuShowHidenCol 
         Caption         =   "Show Hiden Col"
      End
      Begin VB.Menu mdiv1 
         Caption         =   "-"
      End
      Begin VB.Menu mnu_view_objective 
         Caption         =   "View Objective"
      End
      Begin VB.Menu mnu_view_id 
         Caption         =   "View ID"
      End
      Begin VB.Menu mdiv3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuRateInfo 
         Caption         =   "Rate Info"
      End
   End
End
Attribute VB_Name = "frm_MediaPlan_View"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
''*****************************************************************************
' Nama Submodul         :  Frm_MediaPlan_View
' Fungsi Submodul       :  Media Plan Read Only
' Nama Programmer       :  Sistyo
' Tgl Pembuatan         :  19 October 2005
' Last Update           :  19 October 2005/Sistyo'
'*****************************************************************************
Option Explicit
Dim rsTemp As New ADODB.Recordset
Dim intMouseCol As Double 'posisi colom waktu klik kanan
Dim intRow As Double 'posisi row waktu edit
Dim intCol As Double 'posisi colom waktu edit
Dim strViewMonth As String
Dim dblUnlockedCellColor As Double

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

Private Sub cmdCancel_Click()
    FrameViewMonth1.Visible = False
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

Private Sub cmdTVLayering_Click()
    If cboMPNumber.Text <> "" Then
        frm_MPTVLayering.show 1
    Else
        MsgBox "Select MP Number!", vbExclamation, strApplication_Name
    End If
End Sub

Sub Form_Load()
    Call initform
    '=======Resize Form==========
    resize Me, "[PicButtonPanahBawah][picViewMonth][FrameViewMonth1][FrameViewMonth2][cmdOK][cmdCancel][cbkAllMonth][cbkViewMonth] "
    picViewMonth.Left = txtViewMonth.Left + txtViewMonth.Width - picViewMonth.Width - 30
    picViewMonth.Top = txtViewMonth.Top + 30
    FrameViewMonth1.Top = FrameInsertion.Top + Picture1.Top + txtViewMonth.Top + txtViewMonth.Height + 60
    '============================
    strViewMonth = "ALL"
    txtViewMonth.Text = "ALL"
    cbkAllMonth.Value = 1
    dblUnlockedCellColor = vbRed
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
    
    'hide menu klik kanan
        MnuFGMPInsertion.Visible = False
    'Load MP Number
        strBrand_Filter = "select distinct brand_code from media_security_catalog where user_name = '" & strLogin_User & "' and position in ('Planner','Implementor','Buyer') and valid_until>=(select getdate())"
        
        rsTemp.Open "select mp_number from mp_master where brand_code in (" & strBrand_Filter & ") and is_latest = 1 order by mp_number", ConnERP, 1, 3
        'rsTemp.Open "select mp_number from mp_master where brand_code in ('4042') and is_latest = 1 order by mp_number", ConnERP, 1, 3
        While Not rsTemp.EOF
            cboMPNumber.AddItem rsTemp(0)
            rsTemp.MoveNext
            
        Wend
        rsTemp.Close
        
    'Load cboMonth
        
    'Clear Header
        
        txtBrandName.Text = ""
        txtClientName.Text = ""
        txtCreatedDate.Text = ""
        txtCreatedBy.Text = ""
        txtLastUpdateDate.Text = ""
        txtLastUpdateBy.Text = ""
        
    With FGMPInsertion
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
    
    cmdSummary.Enabled = True
    cmdExportToExcel.Enabled = True
    
    
    
    '==================================
    
    
    
    'Clear Grid
        With FGMPInsertion
            .Rows = 6
            For i = 1 To .Rows
                .TextMatrix(.Rows - 1, i - 1) = ""
            Next
        End With
        
    'Load Week Commencing
        Call LoadWeekCommencing(Mid(cboMPNumber.Text, 6, 4))
    
    'Load MP_Master
        strSql = "select mp_number,brand_name,client_name,created_date,created_by,last_update_date,last_update_by,release_date,isnull(yearly_budget,0) yearly_budget,isnull(is_upload_to_web,0) is_upload_to_web from mp_master where mp_number = '" & cboMPNumber.Text & "'"
        rsMPMaster.Open strSql, ConnERP, 1, 3
        jumlahActivity = 0
        Remaining_Budget = 0
        If Not rsMPMaster.EOF Then
            
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
                
            With FGMPInsertion
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
                                                                rsMPInsertion.Open "select mp_plan_dim_id,week_year,spot,nett_rate,gross_rate,[month],isnull(other_cost,0) other_cost from mp_insertion where mp_plan_dim_id = '" & Left(Trim(FGMPInsertion.TextMatrix(i, .cols - 1)), 19) & "'", ConnERP, 1, 3
                                                            Else
                                                                rsMPInsertion.Open "select mp_plan_dim_id,week_year,spot,nett_rate,gross_rate,[month],isnull(other_cost,0) other_cost from mp_insertion where mp_plan_dim_id = '" & Left(Trim(FGMPInsertion.TextMatrix(i, .cols - 1)), 19) & "' and [month] in " & strViewMonth, ConnERP, 1, 3
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
                                                                rsMPInsertion.Open "select mp_plan_dim_id,week_year,spot,nett_rate,gross_rate,[month],isnull(other_cost,0) other_cost from mp_insertion where mp_plan_dim_id = '" & Trim(FGMPInsertion.TextMatrix(i + 2, .cols - 1)) & "'", ConnERP, 1, 3
                                                            Else
                                                                rsMPInsertion.Open "select mp_plan_dim_id,week_year,spot,nett_rate,gross_rate,[month],isnull(other_cost,0) other_cost from mp_insertion where mp_plan_dim_id = '" & Trim(FGMPInsertion.TextMatrix(i + 2, .cols - 1)) & "' and [month] in " & strViewMonth, ConnERP, 1, 3
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
    
    'FGMPInsertion.SetFocus
    
    Me.MousePointer = vbNormal
    'cmdSummary.SetFocus
End Sub


Private Sub HighLightMonth(intMonth As Double)
    Dim Kolom As Integer, baris As Integer
    
    For Kolom = 6 To FGMPInsertion.cols - 4
        If EngMonthIndex(FGMPInsertion.TextMatrix(1, Kolom)) = intMonth Then
            FGMPInsertion.col = Kolom
            For baris = 1 To 3
                FGMPInsertion.Row = baris
                FGMPInsertion.CellBackColor = vbYellow '16744576 biru muda
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
    With FGMPInsertion
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

Private Sub FGMPInsertion_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
'=============================================================================
'Nama Sub Modul : FGMPInsertion_MouseDown
'Fungsi         : Init menu klik kanan
'Programer      : Sistyo
'=============================================================================
    intMouseCol = FGMPInsertion.MouseCol
    If intMouseCol > 1 Then
        If Button = vbRightButton Then
            If FGMPInsertion.MouseCol > FGMPInsertion.FixedCols - 1 Then
                MnuFreeze.Enabled = True
                MnuUnFreeze.Enabled = False
            Else
                MnuFreeze.Enabled = False
                MnuUnFreeze.Enabled = True
            End If
            
            mnuRateInfo.Enabled = False
            mnu_view_objective.Enabled = False
            mnu_view_id.Enabled = False
            
            With FGMPInsertion
                If .MouseCol > 5 And .MouseCol < .cols - 3 Then
                    .col = .MouseCol
                    .Row = .MouseRow
                    
                    If .Text <> "" And Right(.TextMatrix(.Row, .cols - 1), 4) <> "FREQ" And Right(.TextMatrix(.Row, .cols - 1), 5) <> "REACH" And Len(.TextMatrix(.Row, .cols - 1)) > 18 And Mid(.TextMatrix(.Row, .cols - 1), 6, 4) <> "MDUM" Then
                        mnuRateInfo.Enabled = True
                        mnu_view_objective.Enabled = True
                    End If
                    
                    If Right(.TextMatrix(.Row, .cols - 1), 4) = "FREQ" Or Right(.TextMatrix(.Row, .cols - 1), 5) = "REACH" Then
                        If .Text <> "" Then
                            mnu_view_id.Enabled = True
                        End If
                    End If
                    
                End If
            End With
            PopupMenu MnuFGMPInsertion
        End If
    End If
    
End Sub



Private Sub lbl_show_user_server_Click()
    MsgBox strLogin_User & "@" & strServerName, vbExclamation, strApplication_Name
End Sub

Private Sub mnu_view_id_Click()
    Dim strMPPlanDimID As String
    Dim intWeekYear As Integer
    Dim strSql As String
    Dim rsTemp As New ADODB.Recordset
    With FGMPInsertion
        strMPPlanDimID = Left(.TextMatrix(.Row, .cols - 1), 19)
        intWeekYear = Trim(.TextMatrix(3, .col))
        strSql = "select mp_tv_rf_id from mp_tv_reach_frequency where mp_plan_dim_id = '" & strMPPlanDimID & "' and week_year_start<= " & intWeekYear & " and week_year_end >= " & intWeekYear
        rsTemp.Open strSql, ConnERP, 1, 3
        If Not rsTemp.EOF Then
            MsgBox "Objective ID = " & rsTemp(0), vbExclamation, strApplication_Name
        Else
            MsgBox "Objective not found or record has been deleted!", vbExclamation, strApplication_Name
        End If
        rsTemp.Close
    End With
End Sub

Private Sub mnu_view_objective_Click()
    Dim strMPPlanDimID As String
    Dim intWeekYear As Integer
    Dim strSql As String
    Dim rsTemp As New ADODB.Recordset
    With FGMPInsertion
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
            MsgBox "Objective not found or record has been deleted!"
        End If
        rsTemp.Close
    End With
End Sub

Private Sub MnuFreeze_Click()
    Dim i As Integer
    With FGMPInsertion
        
        If EngMonthIndex(.TextMatrix(1, intMouseCol)) <> -1 Then
            i = intMouseCol
            While .TextMatrix(1, i) = .TextMatrix(1, intMouseCol)
                i = i + 1
            Wend
            FGMPInsertion.FixedCols = i
        Else
            FGMPInsertion.FixedCols = intMouseCol + 1
        End If
    End With
End Sub

Private Sub mnuHideCol_Click()
    Dim i As Integer
    Dim strMonth As String, intWeekWidth As Double
    
    With FGMPInsertion
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

Private Sub mnuRateInfo_Click()
    Dim strMPPlanDimID As String, strMPMediumID As String, intMonth As Integer, intWeekYear As Integer
    Dim NettRate As Double
    Dim GrossRate As Double
    Dim MSC_Paid_Value As Double
    Dim MSC_Bonus_Value As Double
    Dim Club_Agency_Value As Double
    Dim Fee As Double
    Dim rsTemp As New ADODB.Recordset
    
    With FGMPInsertion
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
        rsTemp.Open "select nett_rate,gross_rate,spot from mp_insertion where mp_plan_dim_id = '" & strMPPlanDimID & "' and week_year = " & intWeekYear, ConnERP, 1, 3
        If Not rsTemp.EOF Then
            NettRate = rsTemp(0) / rsTemp(2)
            GrossRate = rsTemp(1) / rsTemp(2)
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
        MsgBox ("RATE/SPOT :" & vbCrLf & "Gross Rate = " & FormatNumber(GrossRate, 2) & vbCrLf & "Nett Rate = " & FormatNumber(NettRate, 2) & vbCrLf & "Fee = " & FormatNumber(Fee, 2)), vbExclamation, strApplication_Name
    End With
End Sub

Private Sub MnuShowHidenCol_Click()
    Frm_MPInsertionHidenCol.show 1
End Sub

Private Sub MnuUnFreeze_Click()
    Dim i As Integer
    With FGMPInsertion
        If EngMonthIndex(.TextMatrix(1, intMouseCol)) <> -1 Then
            i = intMouseCol
            While .TextMatrix(1, i) = .TextMatrix(1, intMouseCol)
                i = i - 1
            Wend
            FGMPInsertion.FixedCols = i + 1
        Else
            FGMPInsertion.FixedCols = intMouseCol
        End If
    End With
End Sub

Private Sub FGMPInsertion_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    'ToolTipText
    Dim Teks As String
    Teks = Trim(FGMPInsertion.TextMatrix(FGMPInsertion.MouseRow, FGMPInsertion.MouseCol))
    If Teks <> "" Then
        FGMPInsertion.ToolTipText = Teks
    End If
End Sub

Private Sub picViewMonth_Click()
    If FrameViewMonth1.Visible Then
        FrameViewMonth1.Visible = False
    Else
        FrameViewMonth1.Visible = True
    End If
End Sub

Private Sub incrProgressBar()

    ProgressBarExport.Value = ProgressBarExport.Value + 1
    lblPercent.Caption = ProgressBarExport.Value * 100 \ ProgressBarExport.Max & "% Complete..."
    lblPercent.Refresh

End Sub

Private Sub cmdExportToExcel_Click()
    Me.MousePointer = vbHourglass
    Dim pesan
    pesan = MsgBox("Export To Excel...?", vbQuestion + vbYesNo, strApplication_Name)
    If pesan = 6 Then
        Call ExportToExcel
    End If
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
    ProgressBarExport.Max = 28 + FGMPInsertion.cols - 14 + ((FGMPInsertion.Rows - 4) * (FGMPInsertion.cols - 3)) + 29
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
            .Range(.Cells(xlHeaderrows + (xlTopRow - 1), xlLeftCol).Address & ":" & .Cells(xlHeaderrows + (xlTopRow - 1), FGMPInsertion.cols - 2 + (xlLeftCol - 1)).Address).Borders(xlEdgeBottom).LineStyle = xlDouble
            .Range(.Cells(xlHeaderrows + (xlTopRow - 1), xlLeftCol).Address & ":" & .Cells(xlHeaderrows + (xlTopRow - 1), FGMPInsertion.cols - 2 + (xlLeftCol - 1)).Address).Borders(xlEdgeBottom).Weight = xlThick
            .Range(.Cells(xlHeaderrows + (xlTopRow - 1), xlLeftCol).Address & ":" & .Cells(xlHeaderrows + (xlTopRow - 1), FGMPInsertion.cols - 2 + (xlLeftCol - 1)).Address).Borders(xlEdgeBottom).ColorIndex = xlAutomatic
            
            Call incrProgressBar
        'setting col width
            .Columns(1 + (xlLeftCol - 1)).ColumnWidth = 31.57
            .Columns(2 + (xlLeftCol - 1)).ColumnWidth = 9.29
            .Columns(3 + (xlLeftCol - 1)).ColumnWidth = 19
            .Columns(4 + (xlLeftCol - 1)).ColumnWidth = 22.14
            .Columns(5 + (xlLeftCol - 1)).ColumnWidth = 33.57
            For vCol = 6 To FGMPInsertion.cols - 4
                .Columns(vCol + (xlLeftCol - 1)).ColumnWidth = 3.14
            Next
            .Columns((FGMPInsertion.cols - 2) + (xlLeftCol - 1)).ColumnWidth = 12.43
        
            Call incrProgressBar
        'init header
            .Range(.Cells(xlTopRow, xlLeftCol).Address & ":" & .Cells(xlTopRow + 5, xlLeftCol + FGMPInsertion.cols - 3).Address).Interior.ColorIndex = 2
            .Range(.Cells(xlTopRow, xlLeftCol).Address & ":" & .Cells(xlTopRow + 5, xlLeftCol + FGMPInsertion.cols - 3).Address).Interior.Pattern = xlSolid
            
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
            While Trim(FGMPInsertion.TextMatrix(vRow - xlTopRow - 5, vCol)) = "January"
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
            While Trim(FGMPInsertion.TextMatrix(vRow - xlTopRow - 5, vCol)) = "February"
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
            While Trim(FGMPInsertion.TextMatrix(vRow - xlTopRow - 5, vCol)) = "March"
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
            While Trim(FGMPInsertion.TextMatrix(vRow - xlTopRow - 5, vCol)) = "April"
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
            While Trim(FGMPInsertion.TextMatrix(vRow - xlTopRow - 5, vCol)) = "May"
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
            While Trim(FGMPInsertion.TextMatrix(vRow - xlTopRow - 5, vCol)) = "June"
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
            While Trim(FGMPInsertion.TextMatrix(vRow - xlTopRow - 5, vCol)) = "July"
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
            While Trim(FGMPInsertion.TextMatrix(vRow - xlTopRow - 5, vCol)) = "August"
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
            While Trim(FGMPInsertion.TextMatrix(vRow - xlTopRow - 5, vCol)) = "September"
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
            While Trim(FGMPInsertion.TextMatrix(vRow - xlTopRow - 5, vCol)) = "October"
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
            While Trim(FGMPInsertion.TextMatrix(vRow - xlTopRow - 5, vCol)) = "November"
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
            While Trim(FGMPInsertion.TextMatrix(vRow - xlTopRow - 5, vCol)) = "December"
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
            .Cells(vRow, FGMPInsertion.cols - 3 + (xlLeftCol - 1)).Interior.ColorIndex = 33
            .Cells(vRow, FGMPInsertion.cols - 3 + (xlLeftCol - 1)).Font.Size = 9
            .Cells(vRow, FGMPInsertion.cols - 3 + (xlLeftCol - 1)).Font.Bold = True
            
            
            .Cells(vRow, FGMPInsertion.cols - 3 + (xlLeftCol - 1)).Borders(xlEdgeTop).LineStyle = xlDouble
            .Cells(vRow, FGMPInsertion.cols - 3 + (xlLeftCol - 1)).Borders(xlEdgeTop).Weight = xlThick
            .Cells(vRow, FGMPInsertion.cols - 3 + (xlLeftCol - 1)).Borders(xlEdgeTop).ColorIndex = 5
        
            .Cells(vRow, FGMPInsertion.cols - 3 + (xlLeftCol - 1)).Borders(xlEdgeBottom).LineStyle = xlDouble
            .Cells(vRow, FGMPInsertion.cols - 3 + (xlLeftCol - 1)).Borders(xlEdgeBottom).Weight = xlThick
            .Cells(vRow, FGMPInsertion.cols - 3 + (xlLeftCol - 1)).Borders(xlEdgeBottom).ColorIndex = 5
            
            .Cells(vRow, FGMPInsertion.cols - 3 + (xlLeftCol - 1)).HorizontalAlignment = xlCenter
            .Cells(vRow, FGMPInsertion.cols - 3 + (xlLeftCol - 1)).Value = "GRP/ins"
            '.Cells(vRow, FGMPInsertion.Cols - 3 + (xlLeftCol - 1)).Locked = True
            
            Call incrProgressBar
        'Total
            .Cells(vRow, FGMPInsertion.cols - 2 + (xlLeftCol - 1)).Font.Name = "Sabon MT"
            .Cells(vRow, FGMPInsertion.cols - 2 + (xlLeftCol - 1)).Font.Size = 9
            .Cells(vRow, FGMPInsertion.cols - 2 + (xlLeftCol - 1)).Font.Bold = True
            
            .Cells(vRow, FGMPInsertion.cols - 2 + (xlLeftCol - 1)).Borders(xlEdgeTop).LineStyle = xlDouble
            .Cells(vRow, FGMPInsertion.cols - 2 + (xlLeftCol - 1)).Borders(xlEdgeTop).Weight = xlThick
            .Cells(vRow, FGMPInsertion.cols - 2 + (xlLeftCol - 1)).Borders(xlEdgeTop).ColorIndex = 5
        
            .Cells(vRow, FGMPInsertion.cols - 2 + (xlLeftCol - 1)).Borders(xlEdgeBottom).LineStyle = xlDouble
            .Cells(vRow, FGMPInsertion.cols - 2 + (xlLeftCol - 1)).Borders(xlEdgeBottom).Weight = xlThick
            .Cells(vRow, FGMPInsertion.cols - 2 + (xlLeftCol - 1)).Borders(xlEdgeBottom).ColorIndex = 5
            
            .Cells(vRow, FGMPInsertion.cols - 2 + (xlLeftCol - 1)).HorizontalAlignment = xlCenter
            .Cells(vRow, FGMPInsertion.cols - 2 + (xlLeftCol - 1)).Value = "TOTAL"
            '.Cells(vRow, FGMPInsertion.Cols - 2 + (xlLeftCol - 1)).Locked = True
            
            Call incrProgressBar
            
        'date & week
            sCol = xlLeftCol + 5
            .Range(.Cells(xlTopRow + 7, sCol).Address & ":" & .Cells(xlTopRow + 7, FGMPInsertion.cols - 4 + (xlLeftCol - 1)).Address).Font.Size = 9
            .Range(.Cells(xlTopRow + 7, sCol).Address & ":" & .Cells(xlTopRow + 7, FGMPInsertion.cols - 4 + (xlLeftCol - 1)).Address).Font.Bold = True
            .Range(.Cells(xlTopRow + 7, sCol).Address & ":" & .Cells(xlTopRow + 7, FGMPInsertion.cols - 4 + (xlLeftCol - 1)).Address).HorizontalAlignment = xlCenter
            .Range(.Cells(xlTopRow + 7, sCol).Address & ":" & .Cells(xlTopRow + 7, FGMPInsertion.cols - 4 + (xlLeftCol - 1)).Address).Borders(xlEdgeBottom).LineStyle = xlContinuous
            .Range(.Cells(xlTopRow + 7, sCol).Address & ":" & .Cells(xlTopRow + 7, FGMPInsertion.cols - 4 + (xlLeftCol - 1)).Address).Borders(xlEdgeBottom).Weight = xlMedium
            .Range(.Cells(xlTopRow + 7, sCol).Address & ":" & .Cells(xlTopRow + 7, FGMPInsertion.cols - 4 + (xlLeftCol - 1)).Address).Borders(xlEdgeBottom).ColorIndex = 5
            .Range(.Cells(xlTopRow + 7, sCol).Address & ":" & .Cells(xlTopRow + 7, FGMPInsertion.cols - 4 + (xlLeftCol - 1)).Address).Borders(xlInsideVertical).Weight = xlThin
            .Range(.Cells(xlTopRow + 7, sCol).Address & ":" & .Cells(xlTopRow + 7, FGMPInsertion.cols - 4 + (xlLeftCol - 1)).Address).Borders(xlInsideVertical).LineStyle = xlContinuous
            .Range(.Cells(xlTopRow + 7, sCol).Address & ":" & .Cells(xlTopRow + 7, FGMPInsertion.cols - 4 + (xlLeftCol - 1)).Address).Borders(xlInsideVertical).ColorIndex = 5
            
            .Range(.Cells(xlTopRow + 8, sCol).Address & ":" & .Cells(xlTopRow + 8, FGMPInsertion.cols - 4 + (xlLeftCol - 1)).Address).Font.Size = 10
            .Range(.Cells(xlTopRow + 8, sCol).Address & ":" & .Cells(xlTopRow + 8, FGMPInsertion.cols - 4 + (xlLeftCol - 1)).Address).Font.Name = "Sabon MT"
            .Range(.Cells(xlTopRow + 8, sCol).Address & ":" & .Cells(xlTopRow + 8, FGMPInsertion.cols - 4 + (xlLeftCol - 1)).Address).HorizontalAlignment = xlCenter
            .Range(.Cells(xlTopRow + 8, sCol).Address & ":" & .Cells(xlTopRow + 8, FGMPInsertion.cols - 4 + (xlLeftCol - 1)).Address).Borders(xlEdgeBottom).LineStyle = xlContinuous
            .Range(.Cells(xlTopRow + 8, sCol).Address & ":" & .Cells(xlTopRow + 8, FGMPInsertion.cols - 4 + (xlLeftCol - 1)).Address).Borders(xlEdgeBottom).Weight = xlHairline
            .Range(.Cells(xlTopRow + 8, sCol).Address & ":" & .Cells(xlTopRow + 8, FGMPInsertion.cols - 4 + (xlLeftCol - 1)).Address).Borders(xlEdgeBottom).ColorIndex = xlAutomatic
            .Range(.Cells(xlTopRow + 8, sCol).Address & ":" & .Cells(xlTopRow + 8, FGMPInsertion.cols - 4 + (xlLeftCol - 1)).Address).Borders(xlInsideVertical).Weight = xlHairline
            .Range(.Cells(xlTopRow + 8, sCol).Address & ":" & .Cells(xlTopRow + 8, FGMPInsertion.cols - 4 + (xlLeftCol - 1)).Address).Borders(xlInsideVertical).LineStyle = xlContinuous
            .Range(.Cells(xlTopRow + 8, sCol).Address & ":" & .Cells(xlTopRow + 8, FGMPInsertion.cols - 4 + (xlLeftCol - 1)).Address).Borders(xlInsideVertical).ColorIndex = xlAutomatic
        
            For vCol = 6 To FGMPInsertion.cols - 4
                .Cells(xlTopRow + 7, sCol).Value = FGMPInsertion.TextMatrix(2, vCol)
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
            .Range(.Cells(xlTopRow + 12, 6 + (xlLeftCol - 1)).Address & ":" & .Cells(FGMPInsertion.Rows + xlHeaderrows + xlSubHeaderRows + xlTopRow - OrigHeaderRows, FGMPInsertion.cols - 2 + (xlLeftCol - 1)).Address).HorizontalAlignment = xlCenter
            intTask = 0
            xlRow = 5 + xlTopRow + 7
            For vRow = 5 To FGMPInsertion.Rows - 1
                
                If Mid(FGMPInsertion.TextMatrix(vRow, 0), 1, 9) = "Sub Total" Then
                    .Range(.Cells(xlRow, xlLeftCol + 1).Address & ":" & .Cells(xlRow, xlLeftCol + 4).Address).Merge
                    .Range(.Cells(xlRow, xlLeftCol + 1).Address & ":" & .Cells(xlRow, xlLeftCol + 4).Address).Interior.ColorIndex = 33
                    .Range(.Cells(xlRow, xlLeftCol + 1).Address & ":" & .Cells(xlRow, xlLeftCol + 4).Address).HorizontalAlignment = xlCenter
                    .Range(.Cells(xlRow, xlLeftCol + 1).Address & ":" & .Cells(xlRow, xlLeftCol + 4).Address).Value = FGMPInsertion.TextMatrix(vRow, 0) & " - Task " & intTask
                    .Cells(xlRow, xlLeftCol).Font.ColorIndex = 2 'putih
                    .Cells(xlRow, xlLeftCol).Value = FGMPInsertion.TextMatrix(vRow, 0) & " - Task " & intTask
                    
                    .Range(sRange_Jan & CStr(xlRow) & ":" & .Cells(xlRow, xlLeftCol + FGMPInsertion.cols - 3).Address).Interior.ColorIndex = 6
                    .Range(sRange_Jan & CStr(xlRow) & ":" & .Cells(xlRow, xlLeftCol + FGMPInsertion.cols - 3).Address).Borders(xlEdgeTop).LineStyle = xlContinuous
                    .Range(sRange_Jan & CStr(xlRow) & ":" & .Cells(xlRow, xlLeftCol + FGMPInsertion.cols - 3).Address).Borders(xlEdgeTop).Weight = xlThin
                    .Range(sRange_Jan & CStr(xlRow) & ":" & .Cells(xlRow, xlLeftCol + FGMPInsertion.cols - 3).Address).Borders(xlEdgeBottom).LineStyle = xlContinuous
                    .Range(sRange_Jan & CStr(xlRow) & ":" & .Cells(xlRow, xlLeftCol + FGMPInsertion.cols - 3).Address).Borders(xlEdgeBottom).Weight = xlThin
                    .Range(sRange_Jan & CStr(xlRow) & ":" & .Cells(xlRow, xlLeftCol + FGMPInsertion.cols - 3).Address).NumberFormat = "_(#,#0.00_);_(#,#0.00_);_(-_)"
                    .Range(sRange_Jan & CStr(xlRow) & ":" & .Cells(xlRow, xlLeftCol + FGMPInsertion.cols - 3).Address).HorizontalAlignment = xlCenter
                    .Range(sRange_Jan & CStr(xlRow) & ":" & .Cells(xlRow, xlLeftCol + FGMPInsertion.cols - 3).Address).Font.Name = "Arial"
                    .Range(sRange_Jan & CStr(xlRow) & ":" & .Cells(xlRow, xlLeftCol + FGMPInsertion.cols - 3).Address).Font.Size = 10
                    
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
                    If Trim(FGMPInsertion.TextMatrix(vRow, 0)) = "Task" Then
                        If intTask > 0 Then
                            xlRow = xlRow - 1
                            'Total By Task
                                Call TotalByTask(xlws, xlRow, xlLeftCol, xlStartRowTask, intTask, "")
                                xlRow = xlRow + 1
                                Call TotalByTask(xlws, xlRow, xlLeftCol, xlStartRowTask, intTask, "Actual")
                        End If
                        intTask = intTask + 1
                        xlStartRowTask = xlRow
                        .Cells(xlRow, xlLeftCol).Value = Trim(FGMPInsertion.TextMatrix(vRow, 0)) & " " & intTask & " : " & Trim(FGMPInsertion.TextMatrix(vRow, 1))
                    Else
                        If Trim(FGMPInsertion.TextMatrix(vRow, 0)) <> "" Then
                            .Cells(xlRow, xlLeftCol).Value = Trim(FGMPInsertion.TextMatrix(vRow, 0)) & " : " & Trim(FGMPInsertion.TextMatrix(vRow, 1))
                        End If
                    End If
                End If

                For vCol = 2 To FGMPInsertion.cols - 2
                    If Mid(FGMPInsertion.TextMatrix(vRow, 0), 1, 9) = "Sub Total" Then
                        If vCol < 6 Then
                            vCol = 6
                            ProgressBarExport.Value = ProgressBarExport.Value + 4 'skip
                        End If
                    End If
                    
                    If Trim(FGMPInsertion.TextMatrix(vRow, vCol)) <> "" Then
                        If vCol > 1 And vCol < 6 Then 'buat rata kiri, rata atas
                            .Cells(xlRow, vCol + (xlLeftCol - 1)).HorizontalAlignment = xlLeft
                            .Cells(xlRow, vCol + (xlLeftCol - 1)).VerticalAlignment = xlTop
                            .Cells(xlRow, vCol + (xlLeftCol - 1)).WrapText = True
                        End If
                        .Cells(xlRow, vCol + (xlLeftCol - 1)).Value = Trim(FGMPInsertion.TextMatrix(vRow, vCol))
                        
                        If Mid(FGMPInsertion.TextMatrix(vRow, 0), 1, 9) <> "Sub Total" Then 'untuk sub total udah di lock dan udah merge, jd gak bisa di lock lagi
                            '.Cells(xlRow, vCol + (xlLeftCol - 1)).Locked = True
                        End If
                        
                        If vCol = 2 Then
                            .Cells(xlRow, vCol + (xlLeftCol - 1)).Font.Bold = True
                            '.Cells(xlRow, vCol + (xlLeftCol - 1)).Locked = True
                            Select Case Trim(FGMPInsertion.TextMatrix(vRow, vCol))
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
                If Right(FGMPInsertion.TextMatrix(vRow, FGMPInsertion.cols - 1), 5) = "REACH" Then
                    'Merging TV Reach & FRequency
                        If strViewMonth = "ALL" Then
                            rsTemp.Open "select week_year_start,week_year_end from mp_tv_reach_frequency where mp_plan_dim_id='" & Left(FGMPInsertion.TextMatrix(vRow, FGMPInsertion.cols - 1), 19) & "' order by week_year_start", ConnERP, 1, 3
                        Else
                            rsTemp.Open "select week_year_start,week_year_end from mp_tv_reach_frequency where mp_plan_dim_id='" & Left(FGMPInsertion.TextMatrix(vRow, FGMPInsertion.cols - 1), 19) & "' and (month_start in (" & strViewMonth & ")  or month_end in (" & strViewMonth & "))order by week_year_start", ConnERP, 1, 3
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
            .Range(.Cells(xlTopRow + 6, FGMPInsertion.cols - 3 + (xlLeftCol - 1)).Address & ":" & .Cells(xlRow + 6, FGMPInsertion.cols - 3 + (xlLeftCol - 1)).Address).Borders(xlEdgeRight).LineStyle = xlContinuous
            .Range(.Cells(xlTopRow + 6, FGMPInsertion.cols - 3 + (xlLeftCol - 1)).Address & ":" & .Cells(xlRow + 6, FGMPInsertion.cols - 3 + (xlLeftCol - 1)).Address).Borders(xlEdgeRight).Weight = xlThick
            .Range(.Cells(xlTopRow + 6, FGMPInsertion.cols - 3 + (xlLeftCol - 1)).Address & ":" & .Cells(xlRow + 6, FGMPInsertion.cols - 3 + (xlLeftCol - 1)).Address).Borders(xlEdgeRight).ColorIndex = 5
            
        'Setting Outer Border
            
            .Range(.Cells(xlTopRow, xlLeftCol).Address & ":" & .Cells(xlRow + 8, FGMPInsertion.cols - 2 + (xlLeftCol - 1)).Address).Borders(xlEdgeLeft).LineStyle = xlDouble
            .Range(.Cells(xlTopRow, xlLeftCol).Address & ":" & .Cells(xlRow + 8, FGMPInsertion.cols - 2 + (xlLeftCol - 1)).Address).Borders(xlEdgeLeft).Weight = xlThick
            .Range(.Cells(xlTopRow, xlLeftCol).Address & ":" & .Cells(xlRow + 8, FGMPInsertion.cols - 2 + (xlLeftCol - 1)).Address).Borders(xlEdgeLeft).ColorIndex = xlAutomatic
            
            .Range(.Cells(xlTopRow, xlLeftCol).Address & ":" & .Cells(xlRow + 8, FGMPInsertion.cols - 2 + (xlLeftCol - 1)).Address).Borders(xlEdgeTop).LineStyle = xlDouble
            .Range(.Cells(xlTopRow, xlLeftCol).Address & ":" & .Cells(xlRow + 8, FGMPInsertion.cols - 2 + (xlLeftCol - 1)).Address).Borders(xlEdgeTop).Weight = xlThick
            .Range(.Cells(xlTopRow, xlLeftCol).Address & ":" & .Cells(xlRow + 8, FGMPInsertion.cols - 2 + (xlLeftCol - 1)).Address).Borders(xlEdgeTop).ColorIndex = xlAutomatic
            
            .Range(.Cells(xlTopRow, xlLeftCol).Address & ":" & .Cells(xlRow + 8, FGMPInsertion.cols - 2 + (xlLeftCol - 1)).Address).Borders(xlEdgeRight).LineStyle = xlDouble
            .Range(.Cells(xlTopRow, xlLeftCol).Address & ":" & .Cells(xlRow + 8, FGMPInsertion.cols - 2 + (xlLeftCol - 1)).Address).Borders(xlEdgeRight).Weight = xlThick
            .Range(.Cells(xlTopRow, xlLeftCol).Address & ":" & .Cells(xlRow + 8, FGMPInsertion.cols - 2 + (xlLeftCol - 1)).Address).Borders(xlEdgeRight).ColorIndex = xlAutomatic
            
            .Range(.Cells(xlTopRow, xlLeftCol).Address & ":" & .Cells(xlRow + 8, FGMPInsertion.cols - 2 + (xlLeftCol - 1)).Address).Borders(xlEdgeBottom).LineStyle = xlDouble
            .Range(.Cells(xlTopRow, xlLeftCol).Address & ":" & .Cells(xlRow + 8, FGMPInsertion.cols - 2 + (xlLeftCol - 1)).Address).Borders(xlEdgeBottom).Weight = xlThick
            .Range(.Cells(xlTopRow, xlLeftCol).Address & ":" & .Cells(xlRow + 8, FGMPInsertion.cols - 2 + (xlLeftCol - 1)).Address).Borders(xlEdgeBottom).ColorIndex = xlAutomatic
        
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
            .PageSetup.Zoom = 30: Call incrProgressBar
skip_page_setup:
        'Protect Sheet
            .Protect "27 October", DrawingObjects:=False, Contents:=True, Scenarios:=True, AllowFormattingCells:=True, AllowFormattingColumns:=True, AllowFormattingRows:=True
        
        ProgressBarExport.Value = ProgressBarExport.Max
        lblPercent.Caption = "100% Complete..."
        lblPercent.Refresh
    End With
    
    Set xlws = xlWB.Worksheets(2)
    xlws.Name = "TV Layering"
    Call ExportTVLayering(xlws)
    
    Set xlws = xlWB.Worksheets(3)
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
            
            strSql = "select b.medium_code,c.month_number,sum(c.min_budget) as budget, sum(c.msc_paid_value + c.msc_bonus_value + c.club_agency_value) as fee "
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
                        .Cells(Current_Task_Row + 1, rsSummary(1) + 1).Value = rsSummary("budget") + rsSummary("fee")
                        'year
                        .Cells(Current_Task_Row + 1, 14).Value = .Cells(Current_Task_Row + 1, 14).Value + rsSummary("budget") + rsSummary("fee")
                        'total month
                        .Cells(Total_Row_Pos + 1, rsSummary(1) + 1).Value = .Cells(Total_Row_Pos + 1, rsSummary(1) + 1).Value + rsSummary("budget") + rsSummary("fee")
                        'total year
                        .Cells(Total_Row_Pos + 1, 14).Value = .Cells(Total_Row_Pos + 1, 14).Value + rsSummary("budget") + rsSummary("fee")
                    Case "RD"
                        'month
                        .Cells(Current_Task_Row + 2, rsSummary(1) + 1).Value = rsSummary("budget") + rsSummary("fee")
                        'year
                        .Cells(Current_Task_Row + 2, 14).Value = .Cells(Current_Task_Row + 2, 14).Value + rsSummary("budget") + rsSummary("fee")
                        'total month
                        .Cells(Total_Row_Pos + 2, rsSummary(1) + 1).Value = .Cells(Total_Row_Pos + 2, rsSummary(1) + 1).Value + rsSummary("budget") + rsSummary("fee")
                        'total year
                        .Cells(Total_Row_Pos + 2, 14).Value = .Cells(Total_Row_Pos + 2, 14).Value + rsSummary("budget") + rsSummary("fee")
                    Case "PR"
                        'month
                        .Cells(Current_Task_Row + 3, rsSummary(1) + 1).Value = rsSummary("budget") + rsSummary("fee")
                        'year
                        .Cells(Current_Task_Row + 3, 14).Value = .Cells(Current_Task_Row + 3, 14).Value + rsSummary("budget") + rsSummary("fee")
                        'total month
                        .Cells(Total_Row_Pos + 3, rsSummary(1) + 1).Value = .Cells(Total_Row_Pos + 3, rsSummary(1) + 1).Value + rsSummary("budget") + rsSummary("fee")
                        'total year
                        .Cells(Total_Row_Pos + 3, 14).Value = .Cells(Total_Row_Pos + 3, 14).Value + rsSummary("budget") + rsSummary("fee")
                    Case "CN"
                        'month
                        .Cells(Current_Task_Row + 4, rsSummary(1) + 1).Value = rsSummary("budget") + rsSummary("fee")
                        'year
                        .Cells(Current_Task_Row + 4, 14).Value = .Cells(Current_Task_Row + 4, 14).Value + rsSummary("budget") + rsSummary("fee")
                        'total month
                        .Cells(Total_Row_Pos + 4, rsSummary(1) + 1).Value = .Cells(Total_Row_Pos + 4, rsSummary(1) + 1).Value + rsSummary("budget") + rsSummary("fee")
                        'total year
                        .Cells(Total_Row_Pos + 4, 14).Value = .Cells(Total_Row_Pos + 4, 14).Value + rsSummary("budget") + rsSummary("fee")
                    Case "OT"
                        'month
                        .Cells(Current_Task_Row + 5, rsSummary(1) + 1).Value = rsSummary("budget") + rsSummary("fee")
                        'year
                        .Cells(Current_Task_Row + 5, 14).Value = .Cells(Current_Task_Row + 5, 14).Value + rsSummary("budget") + rsSummary("fee")
                        'total month
                        .Cells(Total_Row_Pos + 5, rsSummary(1) + 1).Value = .Cells(Total_Row_Pos + 5, rsSummary(1) + 1).Value + rsSummary("budget") + rsSummary("fee")
                        'total year
                        .Cells(Total_Row_Pos + 5, 14).Value = .Cells(Total_Row_Pos + 5, 14).Value + rsSummary("budget") + rsSummary("fee")
                End Select
                'Sub Total month
                .Cells(Current_Task_Row + 6, rsSummary(1) + 1).Value = .Cells(Current_Task_Row + 6, rsSummary(1) + 1).Value + rsSummary("budget") + rsSummary("fee")
                'Sub Total year
                .Cells(Current_Task_Row + 6, 14).Value = .Cells(Current_Task_Row + 6, 14).Value + rsSummary("budget") + rsSummary("fee")
                'Grand Total Month
                .Cells(Total_Row_Pos + 6, rsSummary(1) + 1).Value = .Cells(Total_Row_Pos + 6, rsSummary(1) + 1).Value + rsSummary("budget") + rsSummary("fee")
                'Grand Total Year
                .Cells(Total_Row_Pos + 6, 14).Value = .Cells(Total_Row_Pos + 6, 14).Value + rsSummary("budget") + rsSummary("fee")
                
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
            
            strSql = "select b.medium_code,c.month_number,sum(isnull(c.total_actual,0)) as budget, 0 as fee "
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
                        .Cells(Current_Task_Row + 1, rsSummary(1) + 1).Value = rsSummary("budget") + rsSummary("fee")
                        'year
                        .Cells(Current_Task_Row + 1, 14).Value = .Cells(Current_Task_Row + 1, 14).Value + rsSummary("budget") + rsSummary("fee")
                        'total month
                        .Cells(Total_Row_Pos + 1, rsSummary(1) + 1).Value = .Cells(Total_Row_Pos + 1, rsSummary(1) + 1).Value + rsSummary("budget") + rsSummary("fee")
                        'total year
                        .Cells(Total_Row_Pos + 1, 14).Value = .Cells(Total_Row_Pos + 1, 14).Value + rsSummary("budget") + rsSummary("fee")
                    Case "RD"
                        'month
                        .Cells(Current_Task_Row + 2, rsSummary(1) + 1).Value = rsSummary("budget") + rsSummary("fee")
                        'year
                        .Cells(Current_Task_Row + 2, 14).Value = .Cells(Current_Task_Row + 2, 14).Value + rsSummary("budget") + rsSummary("fee")
                        'total month
                        .Cells(Total_Row_Pos + 2, rsSummary(1) + 1).Value = .Cells(Total_Row_Pos + 2, rsSummary(1) + 1).Value + rsSummary("budget") + rsSummary("fee")
                        'total year
                        .Cells(Total_Row_Pos + 2, 14).Value = .Cells(Total_Row_Pos + 2, 14).Value + rsSummary("budget") + rsSummary("fee")
                    Case "PR"
                        'month
                        .Cells(Current_Task_Row + 3, rsSummary(1) + 1).Value = rsSummary("budget") + rsSummary("fee")
                        'year
                        .Cells(Current_Task_Row + 3, 14).Value = .Cells(Current_Task_Row + 3, 14).Value + rsSummary("budget") + rsSummary("fee")
                        'total month
                        .Cells(Total_Row_Pos + 3, rsSummary(1) + 1).Value = .Cells(Total_Row_Pos + 3, rsSummary(1) + 1).Value + rsSummary("budget") + rsSummary("fee")
                        'total year
                        .Cells(Total_Row_Pos + 3, 14).Value = .Cells(Total_Row_Pos + 3, 14).Value + rsSummary("budget") + rsSummary("fee")
                    Case "CN"
                        'month
                        .Cells(Current_Task_Row + 4, rsSummary(1) + 1).Value = rsSummary("budget") + rsSummary("fee")
                        'year
                        .Cells(Current_Task_Row + 4, 14).Value = .Cells(Current_Task_Row + 4, 14).Value + rsSummary("budget") + rsSummary("fee")
                        'total month
                        .Cells(Total_Row_Pos + 4, rsSummary(1) + 1).Value = .Cells(Total_Row_Pos + 4, rsSummary(1) + 1).Value + rsSummary("budget") + rsSummary("fee")
                        'total year
                        .Cells(Total_Row_Pos + 4, 14).Value = .Cells(Total_Row_Pos + 4, 14).Value + rsSummary("budget") + rsSummary("fee")
                    Case "OT"
                        'month
                        .Cells(Current_Task_Row + 5, rsSummary(1) + 1).Value = rsSummary("budget") + rsSummary("fee")
                        'year
                        .Cells(Current_Task_Row + 5, 14).Value = .Cells(Current_Task_Row + 5, 14).Value + rsSummary("budget") + rsSummary("fee")
                        'total month
                        .Cells(Total_Row_Pos + 5, rsSummary(1) + 1).Value = .Cells(Total_Row_Pos + 5, rsSummary(1) + 1).Value + rsSummary("budget") + rsSummary("fee")
                        'total year
                        .Cells(Total_Row_Pos + 5, 14).Value = .Cells(Total_Row_Pos + 5, 14).Value + rsSummary("budget") + rsSummary("fee")
                End Select
                'Sub Total month
                .Cells(Current_Task_Row + 6, rsSummary(1) + 1).Value = .Cells(Current_Task_Row + 6, rsSummary(1) + 1).Value + rsSummary("budget") + rsSummary("fee")
                'Sub Total year
                .Cells(Current_Task_Row + 6, 14).Value = .Cells(Current_Task_Row + 6, 14).Value + rsSummary("budget") + rsSummary("fee")
                'Grand Total Month
                .Cells(Total_Row_Pos + 6, rsSummary(1) + 1).Value = .Cells(Total_Row_Pos + 6, rsSummary(1) + 1).Value + rsSummary("budget") + rsSummary("fee")
                'Grand Total Year
                .Cells(Total_Row_Pos + 6, 14).Value = .Cells(Total_Row_Pos + 6, 14).Value + rsSummary("budget") + rsSummary("fee")
                
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
    'EXPORT BUDGET ACTUAL
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
    
    frm_MPTVLayering.show 1
    
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
'                    With .Range(.Cells(4, intWeekFrom).Address & ":" & .Cells(Frm_MPTVLayering.FGTVLayering.Rows - jumlahTask, intWeekTo).Address).Interior
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
         
        .Range(.Cells(xlRow - 7, xlLeftCol).Address & ":" & .Cells(xlRow - 3, xlLeftCol + FGMPInsertion.cols - 3).Address).Interior.ColorIndex = 34

        .Range(.Cells(xlRow - 7, xlLeftCol).Address & ":" & .Cells(xlRow - 2, xlLeftCol + FGMPInsertion.cols - 3).Address).Borders(xlEdgeTop).LineStyle = xlContinuous
        .Range(.Cells(xlRow - 7, xlLeftCol).Address & ":" & .Cells(xlRow - 2, xlLeftCol + FGMPInsertion.cols - 3).Address).Borders(xlEdgeTop).Weight = xlThin
        .Range(.Cells(xlRow - 7, xlLeftCol).Address & ":" & .Cells(xlRow - 2, xlLeftCol + FGMPInsertion.cols - 3).Address).Borders(xlEdgeTop).ColorIndex = xlAutomatic
        
        .Range(.Cells(xlRow - 7, xlLeftCol).Address & ":" & .Cells(xlRow - 2, xlLeftCol + FGMPInsertion.cols - 3).Address).Borders(xlEdgeBottom).LineStyle = xlContinuous
        .Range(.Cells(xlRow - 7, xlLeftCol).Address & ":" & .Cells(xlRow - 2, xlLeftCol + FGMPInsertion.cols - 3).Address).Borders(xlEdgeBottom).Weight = xlThin
        .Range(.Cells(xlRow - 7, xlLeftCol).Address & ":" & .Cells(xlRow - 2, xlLeftCol + FGMPInsertion.cols - 3).Address).Borders(xlEdgeBottom).ColorIndex = xlAutomatic
        
        .Range(.Cells(xlRow - 7, xlLeftCol).Address & ":" & .Cells(xlRow - 2, xlLeftCol + FGMPInsertion.cols - 3).Address).Borders(xlInsideHorizontal).LineStyle = xlContinuous
        .Range(.Cells(xlRow - 7, xlLeftCol).Address & ":" & .Cells(xlRow - 2, xlLeftCol + FGMPInsertion.cols - 3).Address).Borders(xlInsideHorizontal).Weight = xlThin
        .Range(.Cells(xlRow - 7, xlLeftCol).Address & ":" & .Cells(xlRow - 2, xlLeftCol + FGMPInsertion.cols - 3).Address).Borders(xlInsideHorizontal).ColorIndex = xlAutomatic
        
        .Range(.Cells(xlRow - 7, xlLeftCol).Address & ":" & .Cells(xlRow - 2, xlLeftCol + FGMPInsertion.cols - 3).Address).Font.Size = 9
        
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
        
        .Cells(xlRow - 7, xlLeftCol + FGMPInsertion.cols - 3).Formula = "=sumif(" & .Cells(xlStartRowTask, xlLeftCol).Address & ":" & .Cells(xlRow - 7 - 1, xlLeftCol).Address & "," & .Cells(xlRow - 7, xlLeftCol + 3).Address & "," & .Cells(xlStartRowTask, xlLeftCol + FGMPInsertion.cols - 3).Address & ":" & .Cells(xlRow - 7 - 1, xlLeftCol + FGMPInsertion.cols - 3).Address & ")"
        .Cells(xlRow - 7, xlLeftCol + FGMPInsertion.cols - 3).NumberFormat = "_(#,#0.00_);_(#,#0.00_);_(-_)"
        
        .Cells(xlRow - 6, xlLeftCol + FGMPInsertion.cols - 3).Formula = "=sumif(" & .Cells(xlStartRowTask, xlLeftCol).Address & ":" & .Cells(xlRow - 7 - 1, xlLeftCol).Address & "," & .Cells(xlRow - 6, xlLeftCol + 3).Address & "," & .Cells(xlStartRowTask, xlLeftCol + FGMPInsertion.cols - 3).Address & ":" & .Cells(xlRow - 7 - 1, xlLeftCol + FGMPInsertion.cols - 3).Address & ")"
        .Cells(xlRow - 6, xlLeftCol + FGMPInsertion.cols - 3).NumberFormat = "_(#,#0.00_);_(#,#0.00_);_(-_)"
        
        .Cells(xlRow - 5, xlLeftCol + FGMPInsertion.cols - 3).Formula = "=sumif(" & .Cells(xlStartRowTask, xlLeftCol).Address & ":" & .Cells(xlRow - 7 - 1, xlLeftCol).Address & "," & .Cells(xlRow - 5, xlLeftCol + 3).Address & "," & .Cells(xlStartRowTask, xlLeftCol + FGMPInsertion.cols - 3).Address & ":" & .Cells(xlRow - 7 - 1, xlLeftCol + FGMPInsertion.cols - 3).Address & ")"
        .Cells(xlRow - 5, xlLeftCol + FGMPInsertion.cols - 3).NumberFormat = "_(#,#0.00_);_(#,#0.00_);_(-_)"
        
        .Cells(xlRow - 4, xlLeftCol + FGMPInsertion.cols - 3).Formula = "=sumif(" & .Cells(xlStartRowTask, xlLeftCol).Address & ":" & .Cells(xlRow - 7 - 1, xlLeftCol).Address & "," & .Cells(xlRow - 4, xlLeftCol + 3).Address & "," & .Cells(xlStartRowTask, xlLeftCol + FGMPInsertion.cols - 3).Address & ":" & .Cells(xlRow - 7 - 1, xlLeftCol + FGMPInsertion.cols - 3).Address & ")"
        .Cells(xlRow - 4, xlLeftCol + FGMPInsertion.cols - 3).NumberFormat = "_(#,#0.00_);_(#,#0.00_);_(-_)"
        
        .Cells(xlRow - 3, xlLeftCol + FGMPInsertion.cols - 3).Formula = "=sumif(" & .Cells(xlStartRowTask, xlLeftCol).Address & ":" & .Cells(xlRow - 7 - 1, xlLeftCol).Address & "," & .Cells(xlRow - 3, xlLeftCol + 3).Address & "," & .Cells(xlStartRowTask, xlLeftCol + FGMPInsertion.cols - 3).Address & ":" & .Cells(xlRow - 7 - 1, xlLeftCol + FGMPInsertion.cols - 3).Address & ")"
        .Cells(xlRow - 3, xlLeftCol + FGMPInsertion.cols - 3).NumberFormat = "_(#,#0.00_);_(#,#0.00_);_(-_)"
        
        .Cells(xlRow - 2, xlLeftCol + FGMPInsertion.cols - 3).Formula = "=sum(" & .Cells(xlRow - 3, xlLeftCol + FGMPInsertion.cols - 3).Address & ":" & .Cells(xlRow - 7, xlLeftCol + FGMPInsertion.cols - 3).Address & ")"
        .Cells(xlRow - 2, xlLeftCol + FGMPInsertion.cols - 3).NumberFormat = "_(#,#0.00_);_(#,#0.00_);_(-_)"
    
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
        .Range(sRange_Jan & CStr(xlRow + 1) & ":" & .Cells(xlRow + 7, FGMPInsertion.cols - 2 + xlLeftCol - 3).Address).Borders.LineStyle = xlContinuous
        .Range(sRange_Jan & CStr(xlRow + 1) & ":" & .Cells(xlRow + 7, FGMPInsertion.cols - 2 + xlLeftCol - 3).Address).Borders.Weight = xlThick
        .Range(sRange_Jan & CStr(xlRow + 1) & ":" & .Cells(xlRow + 7, FGMPInsertion.cols - 2 + xlLeftCol - 3).Address).Borders.ColorIndex = 5
        
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
        .Range(.Cells(xlRow + 7, xlLeftCol).Address & ":" & .Cells(xlRow + 9, FGMPInsertion.cols - 2 + xlLeftCol - 1).Address).Borders(xlInsideVertical).LineStyle = xlNone
            
        .Range(sRange_Jan & CStr(xlRow + 7) & ":" & .Cells(xlRow + 7, FGMPInsertion.cols - 2 + xlLeftCol - 1).Address).Borders(xlEdgeTop).LineStyle = xlContinuous
        .Range(sRange_Jan & CStr(xlRow + 7) & ":" & .Cells(xlRow + 7, FGMPInsertion.cols - 2 + xlLeftCol - 1).Address).Borders(xlEdgeTop).Weight = xlThick
        .Range(sRange_Jan & CStr(xlRow + 7) & ":" & .Cells(xlRow + 7, FGMPInsertion.cols - 2 + xlLeftCol - 1).Address).Borders(xlEdgeTop).ColorIndex = 5
        
        .Range(sRange_Jan & CStr(xlRow + 7) & ":" & eRange_Mar & CStr(xlRow + 7)).Merge 'Q1
        .Range(sRange_Jan & CStr(xlRow + 7) & ":" & eRange_Mar & CStr(xlRow + 7)).Formula = "=sum(" & sRange_Jan & CStr(xlRow + 1) & ":" & eRange_Mar & CStr(xlRow + 5) & ")"
        .Range(sRange_Apr & CStr(xlRow + 7) & ":" & eRange_Jun & CStr(xlRow + 7)).Merge 'Q2
        .Range(sRange_Apr & CStr(xlRow + 7) & ":" & eRange_Jun & CStr(xlRow + 7)).Formula = "=sum(" & sRange_Apr & CStr(xlRow + 1) & ":" & eRange_Jun & CStr(xlRow + 5) & ")"
        .Range(sRange_Jul & CStr(xlRow + 7) & ":" & eRange_Sep & CStr(xlRow + 7)).Merge 'Q3
        .Range(sRange_Jul & CStr(xlRow + 7) & ":" & eRange_Sep & CStr(xlRow + 7)).Formula = "=sum(" & sRange_Jul & CStr(xlRow + 1) & ":" & eRange_Sep & CStr(xlRow + 5) & ")"
        .Range(sRange_Oct & CStr(xlRow + 7) & ":" & eRange_Dec & CStr(xlRow + 7)).Merge 'Q4
        .Range(sRange_Oct & CStr(xlRow + 7) & ":" & eRange_Dec & CStr(xlRow + 7)).Formula = "=sum(" & sRange_Oct & CStr(xlRow + 1) & ":" & eRange_Dec & CStr(xlRow + 5) & ")"
        
        'Grand Total 1 tahun
        .Cells(xlRow + 7, xlLeftCol + FGMPInsertion.TextMatrix(3, FGMPInsertion.cols - 4) + 6).Formula = "=sum(" & sRange_Jan & CStr(xlRow + 7) & ":" & eRange_Dec & CStr(xlRow + 7) & ")"
        
        
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

Private Sub cmdSummary_Click()
    Me.MousePointer = vbHourglass
    frm_MPTotalByTask.show 1
    Me.MousePointer = vbDefault
End Sub

Private Sub FGMPInsertion_SelChange()
    With FGMPInsertion
        If Not .ColIsVisible(.col) Then
            .LeftCol = .LeftCol + 1
        End If
    End With
End Sub

Private Sub cmdClose_Click()
'Close Form
    Unload Me
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
        .Range(sRange_Jan & CStr(xlRow + 1) & ":" & .Cells(xlRow + 7, FGMPInsertion.cols - 2 + xlLeftCol - 3).Address).Borders.LineStyle = xlContinuous
        .Range(sRange_Jan & CStr(xlRow + 1) & ":" & .Cells(xlRow + 7, FGMPInsertion.cols - 2 + xlLeftCol - 3).Address).Borders.Weight = xlThick
        .Range(sRange_Jan & CStr(xlRow + 1) & ":" & .Cells(xlRow + 7, FGMPInsertion.cols - 2 + xlLeftCol - 3).Address).Borders.ColorIndex = 5
        
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
        .Range(.Cells(xlRow + 7, xlLeftCol).Address & ":" & .Cells(xlRow + 9, FGMPInsertion.cols - 2 + xlLeftCol - 1).Address).Borders(xlInsideVertical).LineStyle = xlNone

        .Range(sRange_Jan & CStr(xlRow + 7) & ":" & .Cells(xlRow + 7, FGMPInsertion.cols - 2 + xlLeftCol - 1).Address).Borders(xlEdgeTop).LineStyle = xlContinuous
        .Range(sRange_Jan & CStr(xlRow + 7) & ":" & .Cells(xlRow + 7, FGMPInsertion.cols - 2 + xlLeftCol - 1).Address).Borders(xlEdgeTop).Weight = xlThick
        .Range(sRange_Jan & CStr(xlRow + 7) & ":" & .Cells(xlRow + 7, FGMPInsertion.cols - 2 + xlLeftCol - 1).Address).Borders(xlEdgeTop).ColorIndex = 5
        
        .Range(sRange_Jan & CStr(xlRow + 7) & ":" & eRange_Mar & CStr(xlRow + 7)).Merge 'Q1
        .Range(sRange_Jan & CStr(xlRow + 7) & ":" & eRange_Mar & CStr(xlRow + 7)).Formula = "=sum(" & sRange_Jan & CStr(xlRow + 1) & ":" & eRange_Mar & CStr(xlRow + 5) & ")"
        .Range(sRange_Apr & CStr(xlRow + 7) & ":" & eRange_Jun & CStr(xlRow + 7)).Merge 'Q2
        .Range(sRange_Apr & CStr(xlRow + 7) & ":" & eRange_Jun & CStr(xlRow + 7)).Formula = "=sum(" & sRange_Apr & CStr(xlRow + 1) & ":" & eRange_Jun & CStr(xlRow + 5) & ")"
        .Range(sRange_Jul & CStr(xlRow + 7) & ":" & eRange_Sep & CStr(xlRow + 7)).Merge 'Q3
        .Range(sRange_Jul & CStr(xlRow + 7) & ":" & eRange_Sep & CStr(xlRow + 7)).Formula = "=sum(" & sRange_Jul & CStr(xlRow + 1) & ":" & eRange_Sep & CStr(xlRow + 5) & ")"
        .Range(sRange_Oct & CStr(xlRow + 7) & ":" & eRange_Dec & CStr(xlRow + 7)).Merge 'Q4
        .Range(sRange_Oct & CStr(xlRow + 7) & ":" & eRange_Dec & CStr(xlRow + 7)).Formula = "=sum(" & sRange_Oct & CStr(xlRow + 1) & ":" & eRange_Dec & CStr(xlRow + 5) & ")"

        'Grand Total 1 tahun
        .Cells(xlRow + 7, xlLeftCol + FGMPInsertion.TextMatrix(3, FGMPInsertion.cols - 4) + 6).Formula = "=sum(" & sRange_Jan & CStr(xlRow + 7) & ":" & eRange_Dec & CStr(xlRow + 7) & ")"


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

