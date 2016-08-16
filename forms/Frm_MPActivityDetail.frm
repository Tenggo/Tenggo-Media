VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TabCtl32.Ocx"
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Object = "{562E3E04-2C31-4ECE-83F4-4017EEE51D40}#8.0#0"; "todg8.ocx"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSAdoDc.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFlxGrd.ocx"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.Form Frm_MPActivityDetail 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Activity Detail"
   ClientHeight    =   9240
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   9585
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9240
   ScaleWidth      =   9585
   StartUpPosition =   1  'CenterOwner
   Begin VB.PictureBox picToolbar 
      Align           =   1  'Align Top
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ClipControls    =   0   'False
      ForeColor       =   &H80000008&
      Height          =   750
      Left            =   0
      ScaleHeight     =   750
      ScaleWidth      =   9585
      TabIndex        =   12
      TabStop         =   0   'False
      Top             =   0
      Width           =   9585
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
         TabStop         =   0   'False
         Top             =   -15
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
         TabIndex        =   17
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
         TabIndex        =   16
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
         Left            =   4680
         ScaleHeight     =   750
         ScaleWidth      =   1500
         TabIndex        =   15
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
         Left            =   6210
         ScaleHeight     =   750
         ScaleWidth      =   1500
         TabIndex        =   14
         TabStop         =   0   'False
         Top             =   15
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
         Left            =   7740
         ScaleHeight     =   750
         ScaleWidth      =   1500
         TabIndex        =   13
         TabStop         =   0   'False
         Top             =   0
         Width           =   1500
      End
   End
   Begin Threed.SSPanel pnl 
      Align           =   1  'Align Top
      Height          =   8610
      Left            =   0
      TabIndex        =   0
      Top             =   750
      Width           =   9585
      _Version        =   65536
      _ExtentX        =   16907
      _ExtentY        =   15187
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
      Alignment       =   1
      Begin TabDlg.SSTab SSTabMedium 
         Height          =   6885
         Left            =   120
         TabIndex        =   19
         Top             =   1455
         Width           =   9375
         _ExtentX        =   16536
         _ExtentY        =   12144
         _Version        =   393216
         Tabs            =   5
         Tab             =   4
         TabsPerRow      =   5
         TabHeight       =   520
         ShowFocusRect   =   0   'False
         ForeColor       =   -2147483630
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         TabCaption(0)   =   "       TV         "
         TabPicture(0)   =   "Frm_MPActivityDetail.frx":0000
         Tab(0).ControlEnabled=   0   'False
         Tab(0).Control(0)=   "FrameTV"
         Tab(0).ControlCount=   1
         TabCaption(1)   =   "      Radio      "
         TabPicture(1)   =   "Frm_MPActivityDetail.frx":001C
         Tab(1).ControlEnabled=   0   'False
         Tab(1).Control(0)=   "Label44"
         Tab(1).Control(1)=   "fra_RD"
         Tab(1).Control(2)=   "fra_RDByStation"
         Tab(1).Control(3)=   "OptArea"
         Tab(1).Control(4)=   "OptStation"
         Tab(1).ControlCount=   5
         TabCaption(2)   =   "      Print       "
         TabPicture(2)   =   "Frm_MPActivityDetail.frx":0038
         Tab(2).ControlEnabled=   0   'False
         Tab(2).Control(0)=   "FramePR"
         Tab(2).ControlCount=   1
         TabCaption(3)   =   "      Cinema      "
         TabPicture(3)   =   "Frm_MPActivityDetail.frx":0054
         Tab(3).ControlEnabled=   0   'False
         Tab(3).Control(0)=   "frame_CN_detail"
         Tab(3).Control(1)=   "optCNDetail"
         Tab(3).Control(2)=   "optCNBrief"
         Tab(3).Control(3)=   "Label9"
         Tab(3).ControlCount=   4
         TabCaption(4)   =   "       Other       "
         TabPicture(4)   =   "Frm_MPActivityDetail.frx":0070
         Tab(4).ControlEnabled=   -1  'True
         Tab(4).Control(0)=   "Frame3"
         Tab(4).Control(0).Enabled=   0   'False
         Tab(4).ControlCount=   1
         Begin VB.Frame frame_CN_detail 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   6105
            Left            =   -74895
            TabIndex        =   60
            Top             =   675
            Width           =   9135
            Begin VB.Frame frame_CN_brief 
               Height          =   6090
               Left            =   0
               TabIndex        =   61
               Top             =   0
               Width           =   9120
               Begin VB.TextBox txtCNBrief 
                  Appearance      =   0  'Flat
                  BackColor       =   &H00E0E0E0&
                  Height          =   1455
                  Left            =   150
                  MultiLine       =   -1  'True
                  TabIndex        =   68
                  Top             =   480
                  Width           =   8835
               End
               Begin VB.PictureBox Picture4 
                  Height          =   360
                  Left            =   6090
                  ScaleHeight     =   300
                  ScaleWidth      =   2850
                  TabIndex        =   64
                  Top             =   2025
                  Visible         =   0   'False
                  Width           =   2910
                  Begin VB.CommandButton CmdDeleteCN2 
                     Caption         =   "&Delete"
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
                     Left            =   1725
                     TabIndex        =   67
                     ToolTipText     =   "delete supplier"
                     Top             =   0
                     Width           =   855
                  End
                  Begin VB.CommandButton CmdAddCN2 
                     Caption         =   "&Add"
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
                     TabIndex        =   66
                     ToolTipText     =   "add supplier"
                     Top             =   0
                     Width           =   855
                  End
                  Begin VB.CommandButton cmdEditCN2 
                     Caption         =   "&Edit"
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
                     Left            =   870
                     TabIndex        =   65
                     ToolTipText     =   "delete supplier"
                     Top             =   0
                     Width           =   855
                  End
               End
               Begin VB.TextBox txtCNRateGross2 
                  Alignment       =   1  'Right Justify
                  Appearance      =   0  'Flat
                  BackColor       =   &H00E0E0E0&
                  BeginProperty DataFormat 
                     Type            =   1
                     Format          =   "0"
                     HaveTrueFalseNull=   0
                     FirstDayOfWeek  =   0
                     FirstWeekOfYear =   0
                     LCID            =   1033
                     SubFormatType   =   1
                  EndProperty
                  BeginProperty Font 
                     Name            =   "Tahoma"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   285
                  Left            =   1530
                  TabIndex        =   63
                  Text            =   "0.00"
                  Top             =   2100
                  Width           =   1410
               End
               Begin VB.TextBox txtCNRate2 
                  Alignment       =   1  'Right Justify
                  Appearance      =   0  'Flat
                  BackColor       =   &H00E0E0E0&
                  BeginProperty DataFormat 
                     Type            =   1
                     Format          =   "0"
                     HaveTrueFalseNull=   0
                     FirstDayOfWeek  =   0
                     FirstWeekOfYear =   0
                     LCID            =   1033
                     SubFormatType   =   1
                  EndProperty
                  BeginProperty Font 
                     Name            =   "Tahoma"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   285
                  Left            =   1530
                  TabIndex        =   62
                  Text            =   "0.00"
                  Top             =   2460
                  Width           =   1410
               End
               Begin TrueOleDBGrid80.TDBGrid tdg_FGCNMedium2 
                  Height          =   3135
                  Left            =   150
                  TabIndex        =   195
                  Top             =   2805
                  Width           =   8835
                  _ExtentX        =   15584
                  _ExtentY        =   5530
                  _LayoutType     =   4
                  _RowHeight      =   -2147483647
                  _WasPersistedAsPixels=   0
                  Columns(0)._VlistStyle=   0
                  Columns(0)._MaxComboItems=   5
                  Columns(0).Caption=   "No"
                  Columns(0).DataField=   "No"
                  Columns(0)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
                  Columns(1)._VlistStyle=   0
                  Columns(1)._MaxComboItems=   5
                  Columns(1).Caption=   "mp_task_id"
                  Columns(1).DataField=   "mp_task_id"
                  Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
                  Columns(2)._VlistStyle=   0
                  Columns(2)._MaxComboItems=   5
                  Columns(2).Caption=   "Description"
                  Columns(2).DataField=   "task_desc"
                  Columns(2)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
                  Columns.Count   =   3
                  Splits(0)._UserFlags=   0
                  Splits(0).AnchorRightColumn=   -1  'True
                  Splits(0).MarqueeStyle=   3
                  Splits(0).RecordSelectorWidth=   503
                  Splits(0)._SavedRecordSelectors=   -1  'True
                  Splits(0)._GSX_SAVERECORDSELECTORS=   0
                  Splits(0).FetchRowStyle=   -1  'True
                  Splits(0).AlternatingRowStyle=   -1  'True
                  Splits(0).DividerColor=   12632256
                  Splits(0).SpringMode=   0   'False
                  Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
                  Splits(0)._ColumnProps(0)=   "Columns.Count=3"
                  Splits(0)._ColumnProps(1)=   "Column(0).Width=1032"
                  Splits(0)._ColumnProps(2)=   "Column(0).DividerStyle=2"
                  Splits(0)._ColumnProps(3)=   "Column(0).DividerColor=0"
                  Splits(0)._ColumnProps(4)=   "Column(0)._WidthInPix=926"
                  Splits(0)._ColumnProps(5)=   "Column(0)._EditAlways=0"
                  Splits(0)._ColumnProps(6)=   "Column(0)._ColStyle=2"
                  Splits(0)._ColumnProps(7)=   "Column(0).Order=1"
                  Splits(0)._ColumnProps(8)=   "Column(1).Width=2725"
                  Splits(0)._ColumnProps(9)=   "Column(1).DividerColor=0"
                  Splits(0)._ColumnProps(10)=   "Column(1)._WidthInPix=2646"
                  Splits(0)._ColumnProps(11)=   "Column(1)._EditAlways=0"
                  Splits(0)._ColumnProps(12)=   "Column(1).Visible=0"
                  Splits(0)._ColumnProps(13)=   "Column(1).Order=2"
                  Splits(0)._ColumnProps(14)=   "Column(2).Width=14499"
                  Splits(0)._ColumnProps(15)=   "Column(2).DividerColor=0"
                  Splits(0)._ColumnProps(16)=   "Column(2)._WidthInPix=14420"
                  Splits(0)._ColumnProps(17)=   "Column(2)._EditAlways=0"
                  Splits(0)._ColumnProps(18)=   "Column(2).Order=3"
                  Splits.Count    =   1
                  PrintInfos(0)._StateFlags=   3
                  PrintInfos(0).Name=   "piInternal 0"
                  PrintInfos(0).PageHeaderFont=   "Size=8.25,Charset=0,Weight=400,Underline=0,Italic=0,Strikethrough=0,Name=MS Sans Serif"
                  PrintInfos(0).PageFooterFont=   "Size=8.25,Charset=0,Weight=400,Underline=0,Italic=0,Strikethrough=0,Name=MS Sans Serif"
                  PrintInfos(0).PageHeaderHeight=   0
                  PrintInfos(0).PageFooterHeight=   0
                  PrintInfos.Count=   1
                  AllowUpdate     =   0   'False
                  DefColWidth     =   0
                  HeadLines       =   1
                  FootLines       =   1
                  MultipleLines   =   0
                  CellTipsWidth   =   0
                  MultiSelect     =   0
                  DeadAreaBackColor=   12632256
                  RowDividerColor =   12632256
                  RowSubDividerColor=   12632256
                  DirectionAfterEnter=   1
                  DirectionAfterTab=   1
                  MaxRows         =   250000
                  ViewColumnCaptionWidth=   0
                  ViewColumnWidth =   0
                  _PropDict       =   "_ExtentX,2003,3;_ExtentY,2004,3;_LayoutType,512,2;_RowHeight,16,3;_StyleDefs,513,0;_WasPersistedAsPixels,516,2"
                  _StyleDefs(0)   =   "_StyleRoot:id=0,.parent=-1,.alignment=3,.valignment=0,.bgcolor=&H80000005&"
                  _StyleDefs(1)   =   ":id=0,.fgcolor=&H80000008&,.wraptext=0,.locked=0,.transparentBmp=0"
                  _StyleDefs(2)   =   ":id=0,.fgpicPosition=0,.bgpicMode=0,.appearance=0,.borderSize=0,.ellipsis=0"
                  _StyleDefs(3)   =   ":id=0,.borderColor=&H80000005&,.borderType=0,.bold=0,.fontsize=825,.italic=0"
                  _StyleDefs(4)   =   ":id=0,.underline=0,.strikethrough=0,.charset=0"
                  _StyleDefs(5)   =   ":id=0,.fontname=MS Sans Serif"
                  _StyleDefs(6)   =   "Style:id=1,.parent=0,.namedParent=33,.fgcolor=&H0&,.borderColor=&H80000008&"
                  _StyleDefs(7)   =   "CaptionStyle:id=4,.parent=2,.namedParent=37"
                  _StyleDefs(8)   =   "HeadingStyle:id=2,.parent=1,.namedParent=34,.bgcolor=&HBC8A47&"
                  _StyleDefs(9)   =   ":id=2,.fgcolor=&HFFFFFF&"
                  _StyleDefs(10)  =   "FooterStyle:id=3,.parent=1,.namedParent=35"
                  _StyleDefs(11)  =   "InactiveStyle:id=5,.parent=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
                  _StyleDefs(12)  =   "SelectedStyle:id=6,.parent=1,.namedParent=36,.fgcolor=&HFFFFFF&"
                  _StyleDefs(13)  =   "EditorStyle:id=7,.parent=1,.fgcolor=&HFFFF&"
                  _StyleDefs(14)  =   "HighlightRowStyle:id=8,.parent=1,.namedParent=38,.bgcolor=&HD69A69&"
                  _StyleDefs(15)  =   "EvenRowStyle:id=9,.parent=1,.namedParent=39,.bgcolor=&HF8EDDE&"
                  _StyleDefs(16)  =   "OddRowStyle:id=10,.parent=1,.namedParent=40,.fgcolor=&H646464&"
                  _StyleDefs(17)  =   "RecordSelectorStyle:id=11,.parent=2,.namedParent=41"
                  _StyleDefs(18)  =   "FilterBarStyle:id=12,.parent=1,.namedParent=42"
                  _StyleDefs(19)  =   "Splits(0).Style:id=13,.parent=1,.fgcolor=&H80000014&"
                  _StyleDefs(20)  =   "Splits(0).CaptionStyle:id=22,.parent=4"
                  _StyleDefs(21)  =   "Splits(0).HeadingStyle:id=14,.parent=2"
                  _StyleDefs(22)  =   "Splits(0).FooterStyle:id=15,.parent=3"
                  _StyleDefs(23)  =   "Splits(0).InactiveStyle:id=16,.parent=5"
                  _StyleDefs(24)  =   "Splits(0).SelectedStyle:id=18,.parent=6,.bgcolor=&H80000011&"
                  _StyleDefs(25)  =   ":id=18,.fgcolor=&H80000007&"
                  _StyleDefs(26)  =   "Splits(0).EditorStyle:id=17,.parent=7,.fgcolor=&H575757&"
                  _StyleDefs(27)  =   "Splits(0).HighlightRowStyle:id=19,.parent=8,.bgcolor=&H8000000D&,.wraptext=-1"
                  _StyleDefs(28)  =   "Splits(0).EvenRowStyle:id=20,.parent=9,.fgcolor=&H0&"
                  _StyleDefs(29)  =   "Splits(0).OddRowStyle:id=21,.parent=10"
                  _StyleDefs(30)  =   "Splits(0).RecordSelectorStyle:id=23,.parent=11"
                  _StyleDefs(31)  =   "Splits(0).FilterBarStyle:id=24,.parent=12,.wraptext=-1"
                  _StyleDefs(32)  =   "Splits(0).Columns(0).Style:id=28,.parent=13,.alignment=1"
                  _StyleDefs(33)  =   "Splits(0).Columns(0).HeadingStyle:id=25,.parent=14"
                  _StyleDefs(34)  =   "Splits(0).Columns(0).FooterStyle:id=26,.parent=15"
                  _StyleDefs(35)  =   "Splits(0).Columns(0).EditorStyle:id=27,.parent=17"
                  _StyleDefs(36)  =   "Splits(0).Columns(1).Style:id=32,.parent=13"
                  _StyleDefs(37)  =   "Splits(0).Columns(1).HeadingStyle:id=29,.parent=14"
                  _StyleDefs(38)  =   "Splits(0).Columns(1).FooterStyle:id=30,.parent=15"
                  _StyleDefs(39)  =   "Splits(0).Columns(1).EditorStyle:id=31,.parent=17"
                  _StyleDefs(40)  =   "Splits(0).Columns(2).Style:id=46,.parent=13"
                  _StyleDefs(41)  =   "Splits(0).Columns(2).HeadingStyle:id=43,.parent=14"
                  _StyleDefs(42)  =   "Splits(0).Columns(2).FooterStyle:id=44,.parent=15"
                  _StyleDefs(43)  =   "Splits(0).Columns(2).EditorStyle:id=45,.parent=17"
                  _StyleDefs(44)  =   "Named:id=33:Normal"
                  _StyleDefs(45)  =   ":id=33,.parent=0"
                  _StyleDefs(46)  =   "Named:id=34:Heading"
                  _StyleDefs(47)  =   ":id=34,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
                  _StyleDefs(48)  =   ":id=34,.wraptext=-1"
                  _StyleDefs(49)  =   "Named:id=35:Footing"
                  _StyleDefs(50)  =   ":id=35,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
                  _StyleDefs(51)  =   "Named:id=36:Selected"
                  _StyleDefs(52)  =   ":id=36,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
                  _StyleDefs(53)  =   "Named:id=37:Caption"
                  _StyleDefs(54)  =   ":id=37,.parent=34,.alignment=2"
                  _StyleDefs(55)  =   "Named:id=38:HighlightRow"
                  _StyleDefs(56)  =   ":id=38,.parent=33,.bgcolor=&HFF0000&,.fgcolor=&H8000000E&,.borderColor=&HFF2B2B&"
                  _StyleDefs(57)  =   "Named:id=39:EvenRow"
                  _StyleDefs(58)  =   ":id=39,.parent=33,.bgcolor=&HFFFF00&"
                  _StyleDefs(59)  =   "Named:id=40:OddRow"
                  _StyleDefs(60)  =   ":id=40,.parent=33"
                  _StyleDefs(61)  =   "Named:id=41:RecordSelector"
                  _StyleDefs(62)  =   ":id=41,.parent=34"
                  _StyleDefs(63)  =   "Named:id=42:FilterBar"
                  _StyleDefs(64)  =   ":id=42,.parent=33,.fgcolor=&H80000005&"
               End
               Begin VB.Label Label17 
                  Alignment       =   2  'Center
                  Caption         =   "Description"
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
                  Height          =   270
                  Left            =   3990
                  TabIndex        =   72
                  Top             =   180
                  Width           =   1215
               End
               Begin VB.Label Label21 
                  Caption         =   "Gross"
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
                  Left            =   3015
                  TabIndex        =   71
                  Top             =   2190
                  Width           =   420
               End
               Begin VB.Label Label23 
                  Caption         =   "Nett"
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
                  Left            =   3015
                  TabIndex        =   70
                  Top             =   2565
                  Width           =   375
               End
               Begin VB.Label Label32 
                  Alignment       =   1  'Right Justify
                  Caption         =   "Rate/insertion :"
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
                  Height          =   375
                  Left            =   120
                  TabIndex        =   69
                  Top             =   2085
                  Width           =   1335
               End
            End
            Begin VB.ComboBox cboCNName 
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
               Left            =   1545
               Style           =   2  'Dropdown List
               TabIndex        =   89
               Top             =   240
               Width           =   2655
            End
            Begin VB.ComboBox cboCNCode 
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
               Left            =   7110
               Style           =   2  'Dropdown List
               TabIndex        =   88
               Top             =   3300
               Visible         =   0   'False
               Width           =   1650
            End
            Begin VB.ComboBox cboCNSpotType 
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
               Left            =   6060
               Style           =   2  'Dropdown List
               TabIndex        =   87
               Top             =   225
               Width           =   2910
            End
            Begin VB.ComboBox cboCNVersion 
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
               Left            =   6060
               Style           =   2  'Dropdown List
               TabIndex        =   86
               Top             =   600
               Width           =   2265
            End
            Begin VB.PictureBox Picture5 
               Height          =   360
               Left            =   6390
               ScaleHeight     =   300
               ScaleWidth      =   2565
               TabIndex        =   82
               Top             =   1380
               Visible         =   0   'False
               Width           =   2625
               Begin VB.CommandButton cmdAddCN 
                  Caption         =   "&Add"
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
                  TabIndex        =   85
                  ToolTipText     =   "add cinema"
                  Top             =   0
                  Width           =   855
               End
               Begin VB.CommandButton cmdDeleteCN 
                  Caption         =   "&Delete"
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
                  Left            =   1710
                  TabIndex        =   84
                  ToolTipText     =   "delete cinema"
                  Top             =   0
                  Width           =   855
               End
               Begin VB.CommandButton cmdEditCN 
                  Caption         =   "&Edit"
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
                  Left            =   855
                  TabIndex        =   83
                  ToolTipText     =   "delete cinema"
                  Top             =   0
                  Width           =   855
               End
            End
            Begin VB.ComboBox cboCNMaterialDuration 
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
               Left            =   5220
               Style           =   2  'Dropdown List
               TabIndex        =   81
               Top             =   3285
               Visible         =   0   'False
               Width           =   1755
            End
            Begin VB.ComboBox cboCNStudio 
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
               Left            =   1545
               Style           =   2  'Dropdown List
               TabIndex        =   80
               Top             =   600
               Width           =   2655
            End
            Begin VB.ComboBox cboCNMaterialJenisDurasi 
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
               Left            =   3090
               Style           =   2  'Dropdown List
               TabIndex        =   79
               Top             =   3285
               Visible         =   0   'False
               Width           =   2055
            End
            Begin VB.TextBox txtCNRate 
               Alignment       =   1  'Right Justify
               Appearance      =   0  'Flat
               BackColor       =   &H00E0E0E0&
               BeginProperty DataFormat 
                  Type            =   1
                  Format          =   "0"
                  HaveTrueFalseNull=   0
                  FirstDayOfWeek  =   0
                  FirstWeekOfYear =   0
                  LCID            =   1033
                  SubFormatType   =   1
               EndProperty
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
               Height          =   285
               Left            =   1530
               TabIndex        =   78
               Text            =   "0.00"
               Top             =   1365
               Width           =   1410
            End
            Begin VB.TextBox txtCNDuration 
               Alignment       =   1  'Right Justify
               Appearance      =   0  'Flat
               BackColor       =   &H00E0E0E0&
               BeginProperty DataFormat 
                  Type            =   1
                  Format          =   "0"
                  HaveTrueFalseNull=   0
                  FirstDayOfWeek  =   0
                  FirstWeekOfYear =   0
                  LCID            =   1033
                  SubFormatType   =   1
               EndProperty
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   285
               Left            =   6075
               Locked          =   -1  'True
               TabIndex        =   77
               Text            =   "0"
               Top             =   990
               Width           =   585
            End
            Begin VB.ComboBox cboCNJenisDurasi 
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
               Left            =   6735
               Style           =   2  'Dropdown List
               TabIndex        =   76
               Top             =   975
               Width           =   2250
            End
            Begin VB.PictureBox Picture10 
               Height          =   285
               Left            =   8385
               ScaleHeight     =   225
               ScaleWidth      =   525
               TabIndex        =   74
               Top             =   615
               Width           =   585
               Begin VB.CommandButton cmdNewCNMaterial 
                  Caption         =   "+"
                  BeginProperty Font 
                     Name            =   "Tahoma"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   225
                  Left            =   0
                  TabIndex        =   75
                  Top             =   0
                  Width           =   525
               End
            End
            Begin VB.TextBox txtCNRateGross 
               Alignment       =   1  'Right Justify
               Appearance      =   0  'Flat
               BackColor       =   &H00E0E0E0&
               BeginProperty DataFormat 
                  Type            =   1
                  Format          =   "0"
                  HaveTrueFalseNull=   0
                  FirstDayOfWeek  =   0
                  FirstWeekOfYear =   0
                  LCID            =   1033
                  SubFormatType   =   1
               EndProperty
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
               Height          =   285
               Left            =   1545
               TabIndex        =   73
               Text            =   "0.00"
               Top             =   1020
               Width           =   1410
            End
            Begin TrueOleDBGrid80.TDBGrid tdg_FGCNMedium 
               Height          =   3900
               Left            =   135
               TabIndex        =   196
               Top             =   2085
               Width           =   8820
               _ExtentX        =   15558
               _ExtentY        =   6879
               _LayoutType     =   4
               _RowHeight      =   -2147483647
               _WasPersistedAsPixels=   0
               Columns(0)._VlistStyle=   0
               Columns(0)._MaxComboItems=   5
               Columns(0).Caption=   "No"
               Columns(0).DataField=   "No"
               Columns(0)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
               Columns(1)._VlistStyle=   0
               Columns(1)._MaxComboItems=   5
               Columns(1).Caption=   "mp_task_id"
               Columns(1).DataField=   "mp_task_id"
               Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
               Columns(2)._VlistStyle=   0
               Columns(2)._MaxComboItems=   5
               Columns(2).Caption=   "Description"
               Columns(2).DataField=   "task_desc"
               Columns(2)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
               Columns.Count   =   3
               Splits(0)._UserFlags=   0
               Splits(0).AnchorRightColumn=   -1  'True
               Splits(0).MarqueeStyle=   3
               Splits(0).RecordSelectorWidth=   503
               Splits(0)._SavedRecordSelectors=   -1  'True
               Splits(0)._GSX_SAVERECORDSELECTORS=   0
               Splits(0).FetchRowStyle=   -1  'True
               Splits(0).AlternatingRowStyle=   -1  'True
               Splits(0).DividerColor=   12632256
               Splits(0).SpringMode=   0   'False
               Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
               Splits(0)._ColumnProps(0)=   "Columns.Count=3"
               Splits(0)._ColumnProps(1)=   "Column(0).Width=1032"
               Splits(0)._ColumnProps(2)=   "Column(0).DividerStyle=2"
               Splits(0)._ColumnProps(3)=   "Column(0).DividerColor=0"
               Splits(0)._ColumnProps(4)=   "Column(0)._WidthInPix=926"
               Splits(0)._ColumnProps(5)=   "Column(0)._EditAlways=0"
               Splits(0)._ColumnProps(6)=   "Column(0)._ColStyle=2"
               Splits(0)._ColumnProps(7)=   "Column(0).Order=1"
               Splits(0)._ColumnProps(8)=   "Column(1).Width=2725"
               Splits(0)._ColumnProps(9)=   "Column(1).DividerColor=0"
               Splits(0)._ColumnProps(10)=   "Column(1)._WidthInPix=2646"
               Splits(0)._ColumnProps(11)=   "Column(1)._EditAlways=0"
               Splits(0)._ColumnProps(12)=   "Column(1).Visible=0"
               Splits(0)._ColumnProps(13)=   "Column(1).Order=2"
               Splits(0)._ColumnProps(14)=   "Column(2).Width=14499"
               Splits(0)._ColumnProps(15)=   "Column(2).DividerColor=0"
               Splits(0)._ColumnProps(16)=   "Column(2)._WidthInPix=14420"
               Splits(0)._ColumnProps(17)=   "Column(2)._EditAlways=0"
               Splits(0)._ColumnProps(18)=   "Column(2).Order=3"
               Splits.Count    =   1
               PrintInfos(0)._StateFlags=   3
               PrintInfos(0).Name=   "piInternal 0"
               PrintInfos(0).PageHeaderFont=   "Size=8.25,Charset=0,Weight=400,Underline=0,Italic=0,Strikethrough=0,Name=MS Sans Serif"
               PrintInfos(0).PageFooterFont=   "Size=8.25,Charset=0,Weight=400,Underline=0,Italic=0,Strikethrough=0,Name=MS Sans Serif"
               PrintInfos(0).PageHeaderHeight=   0
               PrintInfos(0).PageFooterHeight=   0
               PrintInfos.Count=   1
               AllowUpdate     =   0   'False
               DefColWidth     =   0
               HeadLines       =   1
               FootLines       =   1
               MultipleLines   =   0
               CellTipsWidth   =   0
               MultiSelect     =   0
               DeadAreaBackColor=   12632256
               RowDividerColor =   12632256
               RowSubDividerColor=   12632256
               DirectionAfterEnter=   1
               DirectionAfterTab=   1
               MaxRows         =   250000
               ViewColumnCaptionWidth=   0
               ViewColumnWidth =   0
               _PropDict       =   "_ExtentX,2003,3;_ExtentY,2004,3;_LayoutType,512,2;_RowHeight,16,3;_StyleDefs,513,0;_WasPersistedAsPixels,516,2"
               _StyleDefs(0)   =   "_StyleRoot:id=0,.parent=-1,.alignment=3,.valignment=0,.bgcolor=&H80000005&"
               _StyleDefs(1)   =   ":id=0,.fgcolor=&H80000008&,.wraptext=0,.locked=0,.transparentBmp=0"
               _StyleDefs(2)   =   ":id=0,.fgpicPosition=0,.bgpicMode=0,.appearance=0,.borderSize=0,.ellipsis=0"
               _StyleDefs(3)   =   ":id=0,.borderColor=&H80000005&,.borderType=0,.bold=0,.fontsize=825,.italic=0"
               _StyleDefs(4)   =   ":id=0,.underline=0,.strikethrough=0,.charset=0"
               _StyleDefs(5)   =   ":id=0,.fontname=MS Sans Serif"
               _StyleDefs(6)   =   "Style:id=1,.parent=0,.namedParent=33,.fgcolor=&H0&,.borderColor=&H80000008&"
               _StyleDefs(7)   =   "CaptionStyle:id=4,.parent=2,.namedParent=37"
               _StyleDefs(8)   =   "HeadingStyle:id=2,.parent=1,.namedParent=34,.bgcolor=&HBC8A47&"
               _StyleDefs(9)   =   ":id=2,.fgcolor=&HFFFFFF&"
               _StyleDefs(10)  =   "FooterStyle:id=3,.parent=1,.namedParent=35"
               _StyleDefs(11)  =   "InactiveStyle:id=5,.parent=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
               _StyleDefs(12)  =   "SelectedStyle:id=6,.parent=1,.namedParent=36,.fgcolor=&HFFFFFF&"
               _StyleDefs(13)  =   "EditorStyle:id=7,.parent=1,.fgcolor=&HFFFF&"
               _StyleDefs(14)  =   "HighlightRowStyle:id=8,.parent=1,.namedParent=38,.bgcolor=&HD69A69&"
               _StyleDefs(15)  =   "EvenRowStyle:id=9,.parent=1,.namedParent=39,.bgcolor=&HF8EDDE&"
               _StyleDefs(16)  =   "OddRowStyle:id=10,.parent=1,.namedParent=40,.fgcolor=&H646464&"
               _StyleDefs(17)  =   "RecordSelectorStyle:id=11,.parent=2,.namedParent=41"
               _StyleDefs(18)  =   "FilterBarStyle:id=12,.parent=1,.namedParent=42"
               _StyleDefs(19)  =   "Splits(0).Style:id=13,.parent=1,.fgcolor=&H80000014&"
               _StyleDefs(20)  =   "Splits(0).CaptionStyle:id=22,.parent=4"
               _StyleDefs(21)  =   "Splits(0).HeadingStyle:id=14,.parent=2"
               _StyleDefs(22)  =   "Splits(0).FooterStyle:id=15,.parent=3"
               _StyleDefs(23)  =   "Splits(0).InactiveStyle:id=16,.parent=5"
               _StyleDefs(24)  =   "Splits(0).SelectedStyle:id=18,.parent=6,.bgcolor=&H80000011&"
               _StyleDefs(25)  =   ":id=18,.fgcolor=&H80000007&"
               _StyleDefs(26)  =   "Splits(0).EditorStyle:id=17,.parent=7,.fgcolor=&H575757&"
               _StyleDefs(27)  =   "Splits(0).HighlightRowStyle:id=19,.parent=8,.bgcolor=&H8000000D&,.wraptext=-1"
               _StyleDefs(28)  =   "Splits(0).EvenRowStyle:id=20,.parent=9,.fgcolor=&H0&"
               _StyleDefs(29)  =   "Splits(0).OddRowStyle:id=21,.parent=10"
               _StyleDefs(30)  =   "Splits(0).RecordSelectorStyle:id=23,.parent=11"
               _StyleDefs(31)  =   "Splits(0).FilterBarStyle:id=24,.parent=12,.wraptext=-1"
               _StyleDefs(32)  =   "Splits(0).Columns(0).Style:id=28,.parent=13,.alignment=1"
               _StyleDefs(33)  =   "Splits(0).Columns(0).HeadingStyle:id=25,.parent=14"
               _StyleDefs(34)  =   "Splits(0).Columns(0).FooterStyle:id=26,.parent=15"
               _StyleDefs(35)  =   "Splits(0).Columns(0).EditorStyle:id=27,.parent=17"
               _StyleDefs(36)  =   "Splits(0).Columns(1).Style:id=32,.parent=13"
               _StyleDefs(37)  =   "Splits(0).Columns(1).HeadingStyle:id=29,.parent=14"
               _StyleDefs(38)  =   "Splits(0).Columns(1).FooterStyle:id=30,.parent=15"
               _StyleDefs(39)  =   "Splits(0).Columns(1).EditorStyle:id=31,.parent=17"
               _StyleDefs(40)  =   "Splits(0).Columns(2).Style:id=46,.parent=13"
               _StyleDefs(41)  =   "Splits(0).Columns(2).HeadingStyle:id=43,.parent=14"
               _StyleDefs(42)  =   "Splits(0).Columns(2).FooterStyle:id=44,.parent=15"
               _StyleDefs(43)  =   "Splits(0).Columns(2).EditorStyle:id=45,.parent=17"
               _StyleDefs(44)  =   "Named:id=33:Normal"
               _StyleDefs(45)  =   ":id=33,.parent=0"
               _StyleDefs(46)  =   "Named:id=34:Heading"
               _StyleDefs(47)  =   ":id=34,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
               _StyleDefs(48)  =   ":id=34,.wraptext=-1"
               _StyleDefs(49)  =   "Named:id=35:Footing"
               _StyleDefs(50)  =   ":id=35,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
               _StyleDefs(51)  =   "Named:id=36:Selected"
               _StyleDefs(52)  =   ":id=36,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
               _StyleDefs(53)  =   "Named:id=37:Caption"
               _StyleDefs(54)  =   ":id=37,.parent=34,.alignment=2"
               _StyleDefs(55)  =   "Named:id=38:HighlightRow"
               _StyleDefs(56)  =   ":id=38,.parent=33,.bgcolor=&HFF0000&,.fgcolor=&H8000000E&,.borderColor=&HFF2B2B&"
               _StyleDefs(57)  =   "Named:id=39:EvenRow"
               _StyleDefs(58)  =   ":id=39,.parent=33,.bgcolor=&HFFFF00&"
               _StyleDefs(59)  =   "Named:id=40:OddRow"
               _StyleDefs(60)  =   ":id=40,.parent=33"
               _StyleDefs(61)  =   "Named:id=41:RecordSelector"
               _StyleDefs(62)  =   ":id=41,.parent=34"
               _StyleDefs(63)  =   "Named:id=42:FilterBar"
               _StyleDefs(64)  =   ":id=42,.parent=33,.fgcolor=&H80000005&"
            End
            Begin VB.Label Label24 
               Caption         =   "Cinema "
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
               Height          =   375
               Left            =   210
               TabIndex        =   97
               Top             =   240
               Width           =   1215
            End
            Begin VB.Label Label26 
               Caption         =   "Spot Type  "
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
               Height          =   315
               Left            =   4740
               TabIndex        =   96
               Top             =   240
               Width           =   1215
            End
            Begin VB.Label Label27 
               Caption         =   "Version  "
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
               Height          =   375
               Left            =   4740
               TabIndex        =   95
               Top             =   585
               Width           =   1215
            End
            Begin VB.Label lblCNDuration 
               Caption         =   "Duration  "
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
               Height          =   270
               Left            =   4755
               TabIndex        =   94
               Top             =   990
               Width           =   1215
            End
            Begin VB.Label Label31 
               Caption         =   "Studio  "
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
               Height          =   375
               Left            =   225
               TabIndex        =   93
               Top             =   615
               Width           =   1215
            End
            Begin VB.Label Label34 
               Caption         =   "Rate/insertion "
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
               Height          =   225
               Left            =   225
               TabIndex        =   92
               Top             =   1020
               Width           =   1215
            End
            Begin VB.Label Label42 
               Caption         =   "Gross"
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
               Left            =   3015
               TabIndex        =   91
               Top             =   1110
               Width           =   420
            End
            Begin VB.Label Label43 
               Caption         =   "Nett"
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
               Left            =   3015
               TabIndex        =   90
               Top             =   1455
               Width           =   420
            End
         End
         Begin VB.Frame Frame3 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   6390
            Left            =   120
            TabIndex        =   53
            Top             =   360
            Width           =   9135
            Begin MSAdodcLib.Adodc adoOT 
               Height          =   480
               Left            =   2985
               Top             =   2940
               Visible         =   0   'False
               Width           =   1200
               _ExtentX        =   2117
               _ExtentY        =   847
               ConnectMode     =   0
               CursorLocation  =   3
               IsolationLevel  =   -1
               ConnectionTimeout=   15
               CommandTimeout  =   30
               CursorType      =   3
               LockType        =   3
               CommandType     =   8
               CursorOptions   =   0
               CacheSize       =   50
               MaxRecords      =   0
               BOFAction       =   0
               EOFAction       =   0
               ConnectStringType=   1
               Appearance      =   1
               BackColor       =   -2147483643
               ForeColor       =   -2147483640
               Orientation     =   0
               Enabled         =   -1
               Connect         =   ""
               OLEDBString     =   ""
               OLEDBFile       =   ""
               DataSourceName  =   ""
               OtherAttributes =   ""
               UserName        =   ""
               Password        =   ""
               RecordSource    =   ""
               Caption         =   ""
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               _Version        =   393216
            End
            Begin VB.PictureBox Picture6 
               Height          =   360
               Left            =   6375
               ScaleHeight     =   300
               ScaleWidth      =   2565
               TabIndex        =   55
               Top             =   1485
               Visible         =   0   'False
               Width           =   2625
               Begin VB.CommandButton cmdAddOT 
                  Caption         =   "&Add"
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
                  TabIndex        =   58
                  ToolTipText     =   "add supplier"
                  Top             =   -225
                  Width           =   855
               End
               Begin VB.CommandButton cmdDeleteOT 
                  Caption         =   "&Delete"
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
                  Left            =   1710
                  TabIndex        =   57
                  ToolTipText     =   "delete supplier"
                  Top             =   -210
                  Width           =   855
               End
               Begin VB.CommandButton CmdEditOT 
                  Caption         =   "&Edit"
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
                  Left            =   840
                  TabIndex        =   56
                  ToolTipText     =   "delete supplier"
                  Top             =   -240
                  Width           =   855
               End
            End
            Begin VB.TextBox txtOTDescrition 
               Appearance      =   0  'Flat
               BackColor       =   &H00E0E0E0&
               Height          =   1260
               Left            =   135
               MultiLine       =   -1  'True
               TabIndex        =   54
               Top             =   555
               Width           =   8835
            End
            Begin TrueOleDBGrid80.TDBGrid tdg_OTMedium 
               Bindings        =   "Frm_MPActivityDetail.frx":008C
               Height          =   4290
               Left            =   135
               TabIndex        =   194
               Top             =   2025
               Width           =   8820
               _ExtentX        =   15558
               _ExtentY        =   7567
               _LayoutType     =   4
               _RowHeight      =   -2147483647
               _WasPersistedAsPixels=   0
               Columns(0)._VlistStyle=   0
               Columns(0)._MaxComboItems=   5
               Columns(0).Caption=   "mp_plan_dim_id"
               Columns(0).DataField=   "Plan Dimention ID"
               Columns(0)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
               Columns(1)._VlistStyle=   0
               Columns(1)._MaxComboItems=   5
               Columns(1).Caption=   "Detail ID"
               Columns(1).DataField=   "mp_medium_detail_id"
               Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
               Columns(2)._VlistStyle=   0
               Columns(2)._MaxComboItems=   5
               Columns(2).Caption=   "Description"
               Columns(2).DataField=   "Description"
               Columns(2)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
               Columns(3)._VlistStyle=   80
               Columns(3)._MaxComboItems=   5
               Columns(3).ValueItems(0)._DefaultItem=   0
               Columns(3).ValueItems(0).Value=   "1"
               Columns(3).ValueItems(0).Value.vt=   8
               Columns(3).ValueItems(0).DisplayValue.CLSID=   "{0BE35204-8F91-11CE-9DE3-00AA004BB851}"
               Columns(3).ValueItems(0).DisplayValue(0)=   "bHQAAO4RAABCTe4RAAAAAAAANgAAACgAAABsAAAADgAAAAEAGAAAAAAAuBEAAAAAAAAAAAAAAAAA"
               Columns(3).ValueItems(0).DisplayValue(1)=   "AAAAAAD39/d4e3p4e3t4e3t4e3p4e3t4e3t5e3t5e3t4fHt4e3t5e3t5e3x5fHx5fHx5e3t4fHx5"
               Columns(3).ValueItems(0).DisplayValue(2)=   "e3x6fHt5fHx6fHt6fHx5fHx6fXx5fHx6fHx6fH15fH15fHx6fHx6fHx6fXx6fX17fX16fXx6fX16"
               Columns(3).ValueItems(0).DisplayValue(3)=   "fX17fX17fX17fX57fX17fX17fX17fn17fX57fX57fn18fn57fn57fn57fX58fn58fn58fn58fn98"
               Columns(3).ValueItems(0).DisplayValue(4)=   "fn98f358f398fn58f358f358f398fn58f398f398f398fn99f398f398gIB9f4B9gH99gH99gH99"
               Columns(3).ValueItems(0).DisplayValue(5)=   "gIB9gH99f4B+f4B9gIB9gH9+f39+gIB+gIB+gIB+gIB9gIB+gYB+gYB+gIF+gIB/gYB+gYB+gYB+"
               Columns(3).ValueItems(0).DisplayValue(6)=   "gYB+gYF/gIF/gYB/gIF/gYF/gYF/gYF+gYF/gYF/gYJ/gYJ/gYGAgoFxdHT39/f19fXz8/Pz8/Pz"
               Columns(3).ValueItems(0).DisplayValue(7)=   "8/Pz8/Pr5NjWvpfEnFu1gSzt5drSto/Mq3Xg0bfx7+zz8/Pz8/Pz8/P09PT09PT19PT29fX09PTz"
               Columns(3).ValueItems(0).DisplayValue(8)=   "8/Pz8/Pz8/P29vb29vb19fXz8/P19fX29fX39fX39vb19fX09PT19fX29fX19fXz8/P19fX29vb0"
               Columns(3).ValueItems(0).DisplayValue(9)=   "9PT19fX29vb29fX29fX29fX29vb29vb09fX19PT39vb29vb19fX39vb29vb4+PiysrJUVlZaXFy+"
               Columns(3).ValueItems(0).DisplayValue(10)=   "v7/7+/vz8/P09PT09PT09PT09PT29vb39vb39vb29fX19fX09PT09PT19fX29vb29vb29vb29vb2"
               Columns(3).ValueItems(0).DisplayValue(11)=   "9vbz8/P19fX29vb29vb29vb29vb///+ur69GSUlITExESEhWWVmrra3y8vL4+Pj19fX29fX29fX2"
               Columns(3).ValueItems(0).DisplayValue(12)=   "9fX29fX09PT19fX39vb39vb4+Pjt7e1xdHT39/f19fXz8/Pz8/Py8vHp4dPNrXizfieqbgqpawXs"
               Columns(3).ValueItems(0).DisplayValue(13)=   "49jHpHutcxO+kUfcyKjx8O3y8vHz8/Pz8/Pz8/P29vb4+Pj29vbz8/Pz8/Pz8/P29vb4+Pj29vb0"
               Columns(3).ValueItems(0).DisplayValue(14)=   "9PT39vb39/f29vb39/f29vb09PT19fX4+Pj29vbz8/P29vb49/f29fX19fX29vb39/f29vb29vb3"
               Columns(3).ValueItems(0).DisplayValue(15)=   "9/f29vb09PT19fX39/f29vb19fX4+Pj4+Pj29vbf399cYGAOExOOj4/////z8/P09PT09PT09PT1"
               Columns(3).ValueItems(0).DisplayValue(16)=   "9fX39/f39/f29vb39/f39/f29vb19fX39/f4+Pj39/f39/f4+Pj39/f09PT29vb39/f29vb39/f3"
               Columns(3).ValueItems(0).DisplayValue(17)=   "9/f5+fnf4ODCw8POz8+foKAjJiZNUFDp6en9/f339/f29vb29vb39vb29vb09PT19fX39/f39/f5"
               Columns(3).ValueItems(0).DisplayValue(18)=   "+fnt7e1wdHT39/f19fXz8/Pz8/Pm2sjOr3yvdhmvdyrXwKfx6+P7+/n7+/nn282veCq7jUDh0bjm"
               Columns(3).ValueItems(0).DisplayValue(19)=   "2sjz8/P19fX29vaho6NMT0+YmZn6+vr5+fnq6uqVlpZgY2OPkZHf39+XmJhXWlpSVlZaXV2lpqbm"
               Columns(3).ValueItems(0).DisplayValue(20)=   "5uamp6dYW1uTlZXh4eFwc3NcX1+4uLjX19eAgoJSVVVUV1dpampmaGiBg4PNzc2dnp5dX1+AgYG+"
               Columns(3).ValueItems(0).DisplayValue(21)=   "vr5vcXFqbGzV1tb///94enoNERFeYWHT1NT39/f09PT39/fv7++ur69jZmZXWlpVWFhXWlpnaWm+"
               Columns(3).ValueItems(0).DisplayValue(22)=   "v7/o6OiWmJhTVlZdYGBnaWlfYmKYmprp6uqIiYlXWlpYW1tdX19TV1eytLTj4uKLjY1nampXWVkl"
               Columns(3).ValueItems(0).DisplayValue(23)=   "KCgzODjBwsLU09OBg4NVWFhSVVVQU1NzdXW+vb2lpqZWWVlOUVF5e3vs7OxwdHT39/f19fXz8/Pz"
               Columns(3).ValueItems(0).DisplayValue(24)=   "8/PRtYa4hjSpawW7j1jn2821g0Pi0sHSt5n7+/ni0sGqbgrKqHDRtYbz8/P09PT39/d8fn4BBQVv"
               Columns(3).ValueItems(0).DisplayValue(25)=   "cXH////////s7OxqbGwfIyNgYmKIiYk8Pz9CRUWEhYVSVVUzNzeVlpaKi4sTFhZpamrZ2to2OjoY"
               Columns(3).ValueItems(0).DisplayValue(26)=   "HBydnp6ZmZkxNDQ1OTmPkZF9fn4dICBPUlK8vLx2eHgaHh5NT0+np6czNjYtLy/Ly8vPz89RVFQZ"
               Columns(3).ValueItems(0).DisplayValue(27)=   "HR0zNjacnJz8/Pzz8/P4+Pjs7e2PkZEiJSUoKyuOj49iZWUgIyNjZWWKioo9Pz8zNjZ1d3dQVFQg"
               Columns(3).ValueItems(0).DisplayValue(28)=   "JCRlZ2egoqI5PDwzNjaTlZVRVVUCBweMjo6lpqYbHh5NUFCTlJQ4Ojo7Pj6oqKhmaGgrLi6Cg4Oz"
               Columns(3).ValueItems(0).DisplayValue(29)=   "tLSusLDExcWxsbFMTk4jJydqbGy5urru7u5wdHT39/f19fXz8/Pz8/O/lE2tchGpawWpawWpawWp"
               Columns(3).ValueItems(0).DisplayValue(30)=   "awXi0sHn2837+/nx6+OpawW5iDm/k03z8/P09PT39/eFh4cVGRlxc3O6u7uChISfoKBxcnIzNjZp"
               Columns(3).ValueItems(0).DisplayValue(31)=   "ampJS0sfIyOVl5f///+OkJAJDg5maGiVlpYjJiZzdHTb29tERkYoLCygoaGAgoItMTFkZmbo6em0"
               Columns(3).ValueItems(0).DisplayValue(32)=   "tbUlKChbXV3BwcGAg4MoLCxWWFisrKxAREQ6PT3U1NSTlJQkKCgzNjYuMTF2eHj6+vr19fX4+Pjt"
               Columns(3).ValueItems(0).DisplayValue(33)=   "7e2Vl5cvMzNGSEjf39+oqakvMzMzNzdOUFAuMTFcXl7BwsJ9f38xNTVlZ2dmaWkvMzN0dnbx8fF6"
               Columns(3).ValueItems(0).DisplayValue(34)=   "fX0QFBR0d3dwc3MhJCSbnJze3t5BQ0M9QECmp6c1OjofIiJgYmKBg4OFh4eeoKCMjY08QEBDRka6"
               Columns(3).ValueItems(0).DisplayValue(35)=   "urr////u7u5wdHT39/f19fXz8/Pz8/OyfCKpawWpawWpawXXwKf7+/n7+/n7+/n7+/nHpHupawWs"
               Columns(3).ValueItems(0).DisplayValue(36)=   "cA+yeyLz8/P09PT39/eFh4ccICBVV1daXFwUGRlHSUlxc3M3OjppampMT08hJSWNj4/7+vqMjo4L"
               Columns(3).ValueItems(0).DisplayValue(37)=   "Dw9naWmVlpYgJCRwcnLk5eVAREQlKSm1traRkpIwMzNpbGzt7e2ztLQlKSlaXFy+vr59gIAqLi5f"
               Columns(3).ValueItems(0).DisplayValue(38)=   "YWGsrKxAQ0M7Pj7AwMBvcXEkJydlZ2dNUFBMT0/IyMj29vb5+fnt7e2Wl5cwNDQuMTFpbGxaXFw2"
               Columns(3).ValueItems(0).DisplayValue(39)=   "OTmChISKi4srLi5bXV3BwcF+gYExNTVmaGhqbGwuMjJmaGjs7Ox7fn4PFBR5fHx0eHgcICCWl5ff"
               Columns(3).ValueItems(0).DisplayValue(40)=   "399CREQ9QECmp6c2OTkgJCSBhIRrbW0cHx9AQkJ5enpIS0tGSUnAwcH////v7+9xdHT39/f19fXz"
               Columns(3).ValueItems(0).DisplayValue(41)=   "8/Pz8/OyfCKpawWpawXBmWv7+/n28+7s49jMrYupawWnagWpawWscQ+yfCLz8/P19fX4+PiFh4ce"
               Columns(3).ValueItems(0).DisplayValue(42)=   "IiIxNTU6PT05PT0qLi5KTk4xNDRsbm6ZmppER0c2OTlqbGxFSEhBRESjo6OIiYkSFhZAQ0N6fHwi"
               Columns(3).ValueItems(0).DisplayValue(43)=   "JSUtMDCztLRvcXEkKChDR0d7fn5xc3M0NzdARERqbm44PT0nKyuPkJCzs7M/QkI9Pz93eXk3Ozs/"
               Columns(3).ValueItems(0).DisplayValue(44)=   "QkK0tLR9f38NERF2eHj5+fn5+fnt7e2WmJgwNDQuMTFrbW1BRUUlKCi8vb21tbUVGBhOUVG8vb10"
               Columns(3).ValueItems(0).DisplayValue(45)=   "d3ceIiJlZ2ewsbFJTEwmKSlzdXVQU1MVGRmVl5eys7MoLCw/Q0N1d3cgJCQsLy+pqal3eXkxNDR7"
               Columns(3).ValueItems(0).DisplayValue(46)=   "fX1pamouMTFzdXV+gIAvMjI0ODhhZGSgoqLt7e1xdHT39/f19fXz8/Pz8/O/lE2tchGpawXBmWv7"
               Columns(3).ValueItems(0).DisplayValue(47)=   "+/nx6+Pi0sG7j1jXwKfFo3upawW5iTm/lE3z8/P19fX39/eFh4cbHx8cICBOUlKanJw4PDwfJCQo"
               Columns(3).ValueItems(0).DisplayValue(48)=   "LCxvcXHg4OCur69xc3NpbGxydXW3t7ft7Oy0tbV3enp8fn5vcXF2eHiSlJTHx8djZWUiJiY7Pz9i"
               Columns(3).ValueItems(0).DisplayValue(49)=   "ZWViZGQ3Ojo7Pj5tcHBucXGIiorW1ta6u7s7Pz88Pz+Bg4OAgoKdnp7v7u7GxsZmaGidnp74+Pj5"
               Columns(3).ValueItems(0).DisplayValue(50)=   "+Pjt7e2VmJgvMzNGSEjh4OB7fX0MEBCTlZXKysp4eXmWmJjU1NSsra16fHynqKjw8PCsrKx3eXlt"
               Columns(3).ValueItems(0).DisplayValue(51)=   "cHA/Q0MWGhqYmprw8PCen59ucXFrbm5xdHSBhITOzs7c3NyXmZlrbm5rbW2HiYnJycmRk5MmKiox"
               Columns(3).ValueItems(0).DisplayValue(52)=   "NDRMUFCJi4vs7OxxdHT39/f19fXz8/Pz8/PRtIa4hjSpawWpagXMrYv28+77+/n7+/nXwKe1g0Oq"
               Columns(3).ValueItems(0).DisplayValue(53)=   "bgrKqHDRtYbz8/P09PT39/d+f38IDAwXGxt5enrw8PCDhYUaHh4WGhpkZWXe3t7/////////////"
               Columns(3).ValueItems(0).DisplayValue(54)=   "///6+vr19fX5+fn////////////////////6+vqHh4ceISFdX1/p6em0tbUlKChWWFjZ2Nj/////"
               Columns(3).ValueItems(0).DisplayValue(55)=   "///39/e5ubk4PDw5PDzR0ND////8/Pz19fX4+Pj////7+/vz8/P5+fnt7u6PkZEiJSUoKyuPkZFN"
               Columns(3).ValueItems(0).DisplayValue(56)=   "UFAZHh6PkpL29vb////8/Pz39/f6+vr////6+vr19fX7+/v///////9ydXUKDw+SlJT////8/Pz/"
               Columns(3).ValueItems(0).DisplayValue(57)=   "///////////+/v739/f39/f9/f3////////9/f3///+0tbUvMzM4Ozu5urr////u7u5wdHT39/f1"
               Columns(3).ValueItems(0).DisplayValue(58)=   "9fXz8/Pz8/Pm2sjOr3yvdhmpawWpawWpawXs49jHpHupawWpawW7jUDh0bjm2sjz8/P09PT29vai"
               Columns(3).ValueItems(0).DisplayValue(59)=   "o6NNUFBrbW22trb7+/vX2NhiZWVYW1uQkZHk5OT5+fn09PT09PT09PT09PT09PT09PT09PT09PTz"
               Columns(3).ValueItems(0).DisplayValue(60)=   "8/Pz8/P29vb09PSmpqZfYmKJi4vq6uqmp6cZHBxJTEzPzs729vb29vbz8/O2trYtMDAuMDDHx8f/"
               Columns(3).ValueItems(0).DisplayValue(61)=   "///09PT09PTz8/Pz8/P09PTz8/P29vbu7++usLBkZ2dXWlpVWFhWWVl+f3/T0tL8/Pzz8/Pz8/P0"
               Columns(3).ValueItems(0).DisplayValue(62)=   "9PT09PTz8/Pz8/Pz8/P09PT19fX9/f1maGgAAQGMjo7////z8/Pz8/Pz8/P09PT09PT09PT09PTz"
               Columns(3).ValueItems(0).DisplayValue(63)=   "8/Pz8/Pz8/P09PT7+/vFxsZrbW1xdHTGxsb4+Pju7u5xdHT39/f19fXz8/Pz8/Py8vHp4dPNrXiz"
               Columns(3).ValueItems(0).DisplayValue(64)=   "fieqbgqpawXs49jHpHutcxO+kUfcyKjx8O3y8vHz8/Pz8/Pz8/P29vb39/f39/f19fX09PT19fX4"
               Columns(3).ValueItems(0).DisplayValue(65)=   "9/f39/f29vb09PT09PT09PTz8/Pz8/P09PT09PTz8/P09PTz8/P09PT09PTz8/Pz8/P19fX39/f2"
               Columns(3).ValueItems(0).DisplayValue(66)=   "9vb+/v61tbVPUVFzdXXX19f29vb39/f09PTExcVdYGBeYGDT09P9/f3z8/Pz8/Pz8/Pz8/P09PT0"
               Columns(3).ValueItems(0).DisplayValue(67)=   "9PT09PT09PT19fX4+Pj39/f29vb39/f39/f19fX09PT09PT09PTz8/Pz8/Pz8/P09PT09PTz8/P0"
               Columns(3).ValueItems(0).DisplayValue(68)=   "9PT7+vqJi4s5PT2lpqb////09PT09PT09PT09PTz8/P09PT09PT09PT09PTz8/Pz8/Pz8/P19fX4"
               Columns(3).ValueItems(0).DisplayValue(69)=   "9/f39vb19fX29vbt7e1xdHT39/f19fXz8/Pz8/Pz8/Pz8/Pr5NjWvpfEnFu1gSyxeR28jUHMq3Xg"
               Columns(3).ValueItems(0).DisplayValue(70)=   "0bfx7+zz8/Pz8/Pz8/Pz8/Pz8/P29fX39vb29vb19fX09PTz8/P29fX29vb19fX09PTz8/Pz8/Pz"
               Columns(3).ValueItems(0).DisplayValue(71)=   "8/Pz8/Pz8/Pz8/Pz8/P09PTz8/P09PT09PTz8/P09PT19fX29fX19PTz8/Py8vLy8vLy8vLz8/P0"
               Columns(3).ValueItems(0).DisplayValue(72)=   "9PTz8/Pz8/P09PTx8fHy8vLz8/Pz8/Pz8/P09PT09PTz8/P09PT09PTz8/Pz8/P09PT39vb29vb2"
               Columns(3).ValueItems(0).DisplayValue(73)=   "9fX29vb19fX09PTz8/Pz8/P09PTz8/Pz8/Pz8/Pz8/P09PT09PT09PTz8/Py8vLx8fHz8/P09PTz"
               Columns(3).ValueItems(0).DisplayValue(74)=   "8/Pz8/Pz8/Pz8/Pz8/Pz8/Pz8/Pz8/Pz8/Pz8/Pz8/Pz8/P09PT29vb19fX09PT19fXt7e1wdHT3"
               Columns(3).ValueItems(0).DisplayValue(75)=   "9/f19fX09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT0"
               Columns(3).ValueItems(0).DisplayValue(76)=   "9PT09PT19fX19fX09PT19fX19fX09PT09PT09PT09PT09PT09PT19fX19fX19fX09PT09PT09PT0"
               Columns(3).ValueItems(0).DisplayValue(77)=   "9PT09PT09PT09PT09PT09PT09PT19fX19fX39/f5+fn4+Pj29vb19fX09PT09PT39/f5+fn4+Pj1"
               Columns(3).ValueItems(0).DisplayValue(78)=   "9fX09PT09PT19fX19fX19fX19fX19fX09PT09PT09PT19fX19fX09PT09PT09PT09PT09PT09PT0"
               Columns(3).ValueItems(0).DisplayValue(79)=   "9PT19fX19fX09PT09PT09PT09PT09PT09PT39/f5+fn39vbz9PT09PT19fX19fX09PT09PT09PT1"
               Columns(3).ValueItems(0).DisplayValue(80)=   "9fX19fX19fX09PT09PT09PT09PT09PT19fX19fX39/fu7u7R0dE="
               Columns(3).ValueItems(0).DisplayValue.vt=   9
               Columns(3).ValueItems(0)._PropDict=   "_DefaultItem,517,2"
               Columns(3).ValueItems(1)._DefaultItem=   0
               Columns(3).ValueItems(1).Value=   "2"
               Columns(3).ValueItems(1).Value.vt=   8
               Columns(3).ValueItems(1).DisplayValue=   "2"
               Columns(3).ValueItems(1).DisplayValue.vt=   8
               Columns(3).ValueItems(1)._PropDict=   "_DefaultItem,517,2"
               Columns(3).ValueItems.Count=   2
               Columns(3).Caption=   "command"
               Columns(3).DataField=   "command"
               Columns(3)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
               Columns.Count   =   4
               Splits(0)._UserFlags=   0
               Splits(0).AnchorRightColumn=   -1  'True
               Splits(0).MarqueeStyle=   3
               Splits(0).RecordSelectorWidth=   503
               Splits(0)._SavedRecordSelectors=   -1  'True
               Splits(0)._GSX_SAVERECORDSELECTORS=   0
               Splits(0).FetchRowStyle=   -1  'True
               Splits(0).AlternatingRowStyle=   -1  'True
               Splits(0).DividerColor=   12632256
               Splits(0).SpringMode=   0   'False
               Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
               Splits(0)._ColumnProps(0)=   "Columns.Count=4"
               Splits(0)._ColumnProps(1)=   "Column(0).Width=1032"
               Splits(0)._ColumnProps(2)=   "Column(0).DividerStyle=2"
               Splits(0)._ColumnProps(3)=   "Column(0).DividerColor=0"
               Splits(0)._ColumnProps(4)=   "Column(0)._WidthInPix=926"
               Splits(0)._ColumnProps(5)=   "Column(0)._EditAlways=0"
               Splits(0)._ColumnProps(6)=   "Column(0)._ColStyle=2"
               Splits(0)._ColumnProps(7)=   "Column(0).Order=1"
               Splits(0)._ColumnProps(8)=   "Column(1).Width=2725"
               Splits(0)._ColumnProps(9)=   "Column(1).DividerColor=0"
               Splits(0)._ColumnProps(10)=   "Column(1)._WidthInPix=2646"
               Splits(0)._ColumnProps(11)=   "Column(1)._EditAlways=0"
               Splits(0)._ColumnProps(12)=   "Column(1).Visible=0"
               Splits(0)._ColumnProps(13)=   "Column(1).Order=2"
               Splits(0)._ColumnProps(14)=   "Column(2).Width=11192"
               Splits(0)._ColumnProps(15)=   "Column(2).DividerColor=0"
               Splits(0)._ColumnProps(16)=   "Column(2)._WidthInPix=11113"
               Splits(0)._ColumnProps(17)=   "Column(2)._EditAlways=0"
               Splits(0)._ColumnProps(18)=   "Column(2).Order=3"
               Splits(0)._ColumnProps(19)=   "Column(3).Width=2725"
               Splits(0)._ColumnProps(20)=   "Column(3).DividerColor=0"
               Splits(0)._ColumnProps(21)=   "Column(3)._WidthInPix=2646"
               Splits(0)._ColumnProps(22)=   "Column(3)._EditAlways=0"
               Splits(0)._ColumnProps(23)=   "Column(3).Order=4"
               Splits.Count    =   1
               PrintInfos(0)._StateFlags=   3
               PrintInfos(0).Name=   "piInternal 0"
               PrintInfos(0).PageHeaderFont=   "Size=8.25,Charset=0,Weight=400,Underline=0,Italic=0,Strikethrough=0,Name=MS Sans Serif"
               PrintInfos(0).PageFooterFont=   "Size=8.25,Charset=0,Weight=400,Underline=0,Italic=0,Strikethrough=0,Name=MS Sans Serif"
               PrintInfos(0).PageHeaderHeight=   0
               PrintInfos(0).PageFooterHeight=   0
               PrintInfos.Count=   1
               AllowUpdate     =   0   'False
               DefColWidth     =   0
               HeadLines       =   1
               FootLines       =   1
               MultipleLines   =   0
               CellTipsWidth   =   0
               MultiSelect     =   0
               DeadAreaBackColor=   12632256
               RowDividerColor =   12632256
               RowSubDividerColor=   12632256
               DirectionAfterEnter=   1
               DirectionAfterTab=   1
               MaxRows         =   250000
               ViewColumnCaptionWidth=   0
               ViewColumnWidth =   0
               _PropDict       =   "_ExtentX,2003,3;_ExtentY,2004,3;_LayoutType,512,2;_RowHeight,16,3;_StyleDefs,513,0;_WasPersistedAsPixels,516,2"
               _StyleDefs(0)   =   "_StyleRoot:id=0,.parent=-1,.alignment=3,.valignment=0,.bgcolor=&H80000005&"
               _StyleDefs(1)   =   ":id=0,.fgcolor=&H80000008&,.wraptext=0,.locked=0,.transparentBmp=0"
               _StyleDefs(2)   =   ":id=0,.fgpicPosition=0,.bgpicMode=0,.appearance=0,.borderSize=0,.ellipsis=0"
               _StyleDefs(3)   =   ":id=0,.borderColor=&H80000005&,.borderType=0,.bold=0,.fontsize=825,.italic=0"
               _StyleDefs(4)   =   ":id=0,.underline=0,.strikethrough=0,.charset=0"
               _StyleDefs(5)   =   ":id=0,.fontname=MS Sans Serif"
               _StyleDefs(6)   =   "Style:id=1,.parent=0,.namedParent=33,.fgcolor=&H0&,.borderColor=&H80000008&"
               _StyleDefs(7)   =   "CaptionStyle:id=4,.parent=2,.namedParent=37"
               _StyleDefs(8)   =   "HeadingStyle:id=2,.parent=1,.namedParent=34,.bgcolor=&HBC8A47&"
               _StyleDefs(9)   =   ":id=2,.fgcolor=&HFFFFFF&"
               _StyleDefs(10)  =   "FooterStyle:id=3,.parent=1,.namedParent=35"
               _StyleDefs(11)  =   "InactiveStyle:id=5,.parent=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
               _StyleDefs(12)  =   "SelectedStyle:id=6,.parent=1,.namedParent=36,.fgcolor=&HFFFFFF&"
               _StyleDefs(13)  =   "EditorStyle:id=7,.parent=1,.fgcolor=&HFFFF&"
               _StyleDefs(14)  =   "HighlightRowStyle:id=8,.parent=1,.namedParent=38,.bgcolor=&HD69A69&"
               _StyleDefs(15)  =   "EvenRowStyle:id=9,.parent=1,.namedParent=39,.bgcolor=&HF8EDDE&"
               _StyleDefs(16)  =   "OddRowStyle:id=10,.parent=1,.namedParent=40,.fgcolor=&H646464&"
               _StyleDefs(17)  =   "RecordSelectorStyle:id=11,.parent=2,.namedParent=41"
               _StyleDefs(18)  =   "FilterBarStyle:id=12,.parent=1,.namedParent=42"
               _StyleDefs(19)  =   "Splits(0).Style:id=13,.parent=1,.fgcolor=&H80000014&"
               _StyleDefs(20)  =   "Splits(0).CaptionStyle:id=22,.parent=4"
               _StyleDefs(21)  =   "Splits(0).HeadingStyle:id=14,.parent=2"
               _StyleDefs(22)  =   "Splits(0).FooterStyle:id=15,.parent=3"
               _StyleDefs(23)  =   "Splits(0).InactiveStyle:id=16,.parent=5"
               _StyleDefs(24)  =   "Splits(0).SelectedStyle:id=18,.parent=6,.bgcolor=&H80000011&"
               _StyleDefs(25)  =   ":id=18,.fgcolor=&H80000007&"
               _StyleDefs(26)  =   "Splits(0).EditorStyle:id=17,.parent=7,.fgcolor=&H575757&"
               _StyleDefs(27)  =   "Splits(0).HighlightRowStyle:id=19,.parent=8,.bgcolor=&H8000000D&,.wraptext=-1"
               _StyleDefs(28)  =   "Splits(0).EvenRowStyle:id=20,.parent=9,.fgcolor=&H0&"
               _StyleDefs(29)  =   "Splits(0).OddRowStyle:id=21,.parent=10"
               _StyleDefs(30)  =   "Splits(0).RecordSelectorStyle:id=23,.parent=11"
               _StyleDefs(31)  =   "Splits(0).FilterBarStyle:id=24,.parent=12,.wraptext=-1"
               _StyleDefs(32)  =   "Splits(0).Columns(0).Style:id=28,.parent=13,.alignment=1"
               _StyleDefs(33)  =   "Splits(0).Columns(0).HeadingStyle:id=25,.parent=14"
               _StyleDefs(34)  =   "Splits(0).Columns(0).FooterStyle:id=26,.parent=15"
               _StyleDefs(35)  =   "Splits(0).Columns(0).EditorStyle:id=27,.parent=17"
               _StyleDefs(36)  =   "Splits(0).Columns(1).Style:id=32,.parent=13"
               _StyleDefs(37)  =   "Splits(0).Columns(1).HeadingStyle:id=29,.parent=14"
               _StyleDefs(38)  =   "Splits(0).Columns(1).FooterStyle:id=30,.parent=15"
               _StyleDefs(39)  =   "Splits(0).Columns(1).EditorStyle:id=31,.parent=17"
               _StyleDefs(40)  =   "Splits(0).Columns(2).Style:id=46,.parent=13"
               _StyleDefs(41)  =   "Splits(0).Columns(2).HeadingStyle:id=43,.parent=14"
               _StyleDefs(42)  =   "Splits(0).Columns(2).FooterStyle:id=44,.parent=15"
               _StyleDefs(43)  =   "Splits(0).Columns(2).EditorStyle:id=45,.parent=17"
               _StyleDefs(44)  =   "Splits(0).Columns(3).Style:id=50,.parent=13"
               _StyleDefs(45)  =   "Splits(0).Columns(3).HeadingStyle:id=47,.parent=14"
               _StyleDefs(46)  =   "Splits(0).Columns(3).FooterStyle:id=48,.parent=15"
               _StyleDefs(47)  =   "Splits(0).Columns(3).EditorStyle:id=49,.parent=17"
               _StyleDefs(48)  =   "Named:id=33:Normal"
               _StyleDefs(49)  =   ":id=33,.parent=0"
               _StyleDefs(50)  =   "Named:id=34:Heading"
               _StyleDefs(51)  =   ":id=34,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
               _StyleDefs(52)  =   ":id=34,.wraptext=-1"
               _StyleDefs(53)  =   "Named:id=35:Footing"
               _StyleDefs(54)  =   ":id=35,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
               _StyleDefs(55)  =   "Named:id=36:Selected"
               _StyleDefs(56)  =   ":id=36,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
               _StyleDefs(57)  =   "Named:id=37:Caption"
               _StyleDefs(58)  =   ":id=37,.parent=34,.alignment=2"
               _StyleDefs(59)  =   "Named:id=38:HighlightRow"
               _StyleDefs(60)  =   ":id=38,.parent=33,.bgcolor=&HFF0000&,.fgcolor=&H8000000E&,.borderColor=&HFF2B2B&"
               _StyleDefs(61)  =   "Named:id=39:EvenRow"
               _StyleDefs(62)  =   ":id=39,.parent=33,.bgcolor=&HFFFF00&"
               _StyleDefs(63)  =   "Named:id=40:OddRow"
               _StyleDefs(64)  =   ":id=40,.parent=33"
               _StyleDefs(65)  =   "Named:id=41:RecordSelector"
               _StyleDefs(66)  =   ":id=41,.parent=34"
               _StyleDefs(67)  =   "Named:id=42:FilterBar"
               _StyleDefs(68)  =   ":id=42,.parent=33,.fgcolor=&H80000005&"
            End
            Begin VB.Label Label13 
               Alignment       =   2  'Center
               Caption         =   "Description"
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
               Height          =   270
               Left            =   4005
               TabIndex        =   59
               Top             =   255
               Width           =   1215
            End
         End
         Begin VB.OptionButton OptStation 
            Caption         =   "Station"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Left            =   -72450
            TabIndex        =   52
            Top             =   495
            Width           =   1005
         End
         Begin VB.OptionButton OptArea 
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
            Height          =   240
            Left            =   -73350
            TabIndex        =   51
            Top             =   495
            Value           =   -1  'True
            Width           =   780
         End
         Begin VB.Frame FrameTV 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   6360
            Left            =   -74880
            TabIndex        =   22
            Top             =   360
            Width           =   9135
            Begin VB.PictureBox Picture8 
               BorderStyle     =   0  'None
               Height          =   405
               Left            =   8055
               Picture         =   "Frm_MPActivityDetail.frx":00A0
               ScaleHeight     =   405
               ScaleWidth      =   450
               TabIndex        =   39
               Top             =   585
               Width           =   450
               Begin VB.CommandButton cmdNewTVMaterial 
                  Caption         =   "+"
                  BeginProperty Font 
                     Name            =   "Tahoma"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   330
                  Left            =   -15
                  Style           =   1  'Graphical
                  TabIndex        =   40
                  Top             =   -15
                  Width           =   330
               End
            End
            Begin VB.TextBox txtTVDuration 
               Alignment       =   1  'Right Justify
               Appearance      =   0  'Flat
               BackColor       =   &H00E0E0E0&
               BeginProperty DataFormat 
                  Type            =   1
                  Format          =   "0"
                  HaveTrueFalseNull=   0
                  FirstDayOfWeek  =   0
                  FirstWeekOfYear =   0
                  LCID            =   1033
                  SubFormatType   =   1
               EndProperty
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   285
               Left            =   5775
               Locked          =   -1  'True
               TabIndex        =   38
               Text            =   "0"
               Top             =   960
               Width           =   690
            End
            Begin VB.TextBox txtTVRate 
               Alignment       =   1  'Right Justify
               Appearance      =   0  'Flat
               BackColor       =   &H00E0E0E0&
               BeginProperty DataFormat 
                  Type            =   1
                  Format          =   "0"
                  HaveTrueFalseNull=   0
                  FirstDayOfWeek  =   0
                  FirstWeekOfYear =   0
                  LCID            =   1033
                  SubFormatType   =   1
               EndProperty
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   285
               Left            =   1590
               TabIndex        =   37
               Text            =   "0.00"
               Top             =   1680
               Width           =   1605
            End
            Begin VB.ComboBox cboTVProgram 
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
               Left            =   6465
               TabIndex        =   36
               Top             =   3300
               Visible         =   0   'False
               Width           =   2235
            End
            Begin VB.ComboBox cboTVStationCode 
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
               Left            =   1575
               Style           =   2  'Dropdown List
               TabIndex        =   35
               Top             =   240
               Width           =   2235
            End
            Begin VB.ComboBox cboTVStationName 
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
               Left            =   5775
               Style           =   2  'Dropdown List
               TabIndex        =   34
               Top             =   255
               Width           =   2235
            End
            Begin VB.ComboBox cboTVSpotType 
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
               Left            =   1575
               Style           =   2  'Dropdown List
               TabIndex        =   33
               Top             =   975
               Width           =   2235
            End
            Begin VB.ComboBox cboTVVersion 
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
               Left            =   5775
               Style           =   2  'Dropdown List
               TabIndex        =   32
               Top             =   615
               Width           =   2235
            End
            Begin VB.ComboBox cboMaterialDuration 
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
               Left            =   5145
               Style           =   2  'Dropdown List
               TabIndex        =   31
               Top             =   3345
               Visible         =   0   'False
               Width           =   1170
            End
            Begin VB.PictureBox Picture1 
               Height          =   360
               Left            =   6345
               ScaleHeight     =   300
               ScaleWidth      =   2565
               TabIndex        =   27
               Top             =   1350
               Visible         =   0   'False
               Width           =   2625
               Begin VB.CommandButton cmdDeleteTV 
                  Caption         =   "&Delete"
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
                  Left            =   1710
                  TabIndex        =   30
                  ToolTipText     =   "delete tv station"
                  Top             =   0
                  Width           =   855
               End
               Begin VB.CommandButton cmdAddTV 
                  Caption         =   "&Add"
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
                  TabIndex        =   29
                  ToolTipText     =   "add tv station"
                  Top             =   0
                  Width           =   855
               End
               Begin VB.CommandButton CmdEditTV 
                  Caption         =   "&Edit"
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
                  Left            =   855
                  TabIndex        =   28
                  ToolTipText     =   "delete tv station"
                  Top             =   0
                  Width           =   855
               End
            End
            Begin VB.TextBox txtTVRateGross 
               Alignment       =   1  'Right Justify
               Appearance      =   0  'Flat
               BackColor       =   &H00E0E0E0&
               BeginProperty DataFormat 
                  Type            =   1
                  Format          =   "0"
                  HaveTrueFalseNull=   0
                  FirstDayOfWeek  =   0
                  FirstWeekOfYear =   0
                  LCID            =   1033
                  SubFormatType   =   1
               EndProperty
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   285
               Left            =   1590
               TabIndex        =   26
               Text            =   "0.00"
               Top             =   1350
               Width           =   1605
            End
            Begin VB.CommandButton CmdNewTVStation 
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   330
               Left            =   3825
               Picture         =   "Frm_MPActivityDetail.frx":01A2
               Style           =   1  'Graphical
               TabIndex        =   25
               Top             =   240
               Width           =   390
            End
            Begin VB.ComboBox cboTVMarketName 
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
               Left            =   1575
               Style           =   2  'Dropdown List
               TabIndex        =   24
               Top             =   600
               Width           =   2235
            End
            Begin VB.ComboBox cboTVMarketCode 
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
               Left            =   2385
               Style           =   2  'Dropdown List
               TabIndex        =   23
               Top             =   4425
               Visible         =   0   'False
               Width           =   2175
            End
            Begin TrueOleDBGrid80.TDBGrid tdg_TVMedium 
               Height          =   3960
               Left            =   120
               TabIndex        =   197
               Top             =   2235
               Width           =   8850
               _ExtentX        =   15610
               _ExtentY        =   6985
               _LayoutType     =   4
               _RowHeight      =   -2147483647
               _WasPersistedAsPixels=   0
               Columns(0)._VlistStyle=   0
               Columns(0)._MaxComboItems=   5
               Columns(0).Caption=   "No"
               Columns(0).DataField=   "No"
               Columns(0)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
               Columns(1)._VlistStyle=   0
               Columns(1)._MaxComboItems=   5
               Columns(1).Caption=   "mp_task_id"
               Columns(1).DataField=   "mp_task_id"
               Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
               Columns(2)._VlistStyle=   0
               Columns(2)._MaxComboItems=   5
               Columns(2).Caption=   "Description"
               Columns(2).DataField=   "task_desc"
               Columns(2)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
               Columns.Count   =   3
               Splits(0)._UserFlags=   0
               Splits(0).AnchorRightColumn=   -1  'True
               Splits(0).MarqueeStyle=   3
               Splits(0).RecordSelectorWidth=   503
               Splits(0)._SavedRecordSelectors=   -1  'True
               Splits(0)._GSX_SAVERECORDSELECTORS=   0
               Splits(0).FetchRowStyle=   -1  'True
               Splits(0).AlternatingRowStyle=   -1  'True
               Splits(0).DividerColor=   12632256
               Splits(0).SpringMode=   0   'False
               Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
               Splits(0)._ColumnProps(0)=   "Columns.Count=3"
               Splits(0)._ColumnProps(1)=   "Column(0).Width=1032"
               Splits(0)._ColumnProps(2)=   "Column(0).DividerStyle=2"
               Splits(0)._ColumnProps(3)=   "Column(0).DividerColor=0"
               Splits(0)._ColumnProps(4)=   "Column(0)._WidthInPix=926"
               Splits(0)._ColumnProps(5)=   "Column(0)._EditAlways=0"
               Splits(0)._ColumnProps(6)=   "Column(0)._ColStyle=2"
               Splits(0)._ColumnProps(7)=   "Column(0).Order=1"
               Splits(0)._ColumnProps(8)=   "Column(1).Width=2725"
               Splits(0)._ColumnProps(9)=   "Column(1).DividerColor=0"
               Splits(0)._ColumnProps(10)=   "Column(1)._WidthInPix=2646"
               Splits(0)._ColumnProps(11)=   "Column(1)._EditAlways=0"
               Splits(0)._ColumnProps(12)=   "Column(1).Visible=0"
               Splits(0)._ColumnProps(13)=   "Column(1).Order=2"
               Splits(0)._ColumnProps(14)=   "Column(2).Width=1640"
               Splits(0)._ColumnProps(15)=   "Column(2).DividerColor=0"
               Splits(0)._ColumnProps(16)=   "Column(2)._WidthInPix=1561"
               Splits(0)._ColumnProps(17)=   "Column(2)._EditAlways=0"
               Splits(0)._ColumnProps(18)=   "Column(2).Order=3"
               Splits.Count    =   1
               PrintInfos(0)._StateFlags=   3
               PrintInfos(0).Name=   "piInternal 0"
               PrintInfos(0).PageHeaderFont=   "Size=8.25,Charset=0,Weight=400,Underline=0,Italic=0,Strikethrough=0,Name=MS Sans Serif"
               PrintInfos(0).PageFooterFont=   "Size=8.25,Charset=0,Weight=400,Underline=0,Italic=0,Strikethrough=0,Name=MS Sans Serif"
               PrintInfos(0).PageHeaderHeight=   0
               PrintInfos(0).PageFooterHeight=   0
               PrintInfos.Count=   1
               AllowUpdate     =   0   'False
               DefColWidth     =   0
               HeadLines       =   1
               FootLines       =   1
               MultipleLines   =   0
               CellTipsWidth   =   0
               MultiSelect     =   0
               DeadAreaBackColor=   12632256
               ScrollTrack     =   -1  'True
               ScrollTips      =   -1  'True
               RowDividerColor =   12632256
               RowSubDividerColor=   12632256
               DirectionAfterEnter=   1
               DirectionAfterTab=   1
               MaxRows         =   250000
               ViewColumnCaptionWidth=   0
               ViewColumnWidth =   0
               _PropDict       =   "_ExtentX,2003,3;_ExtentY,2004,3;_LayoutType,512,2;_RowHeight,16,3;_StyleDefs,513,0;_WasPersistedAsPixels,516,2"
               _StyleDefs(0)   =   "_StyleRoot:id=0,.parent=-1,.alignment=3,.valignment=0,.bgcolor=&H80000005&"
               _StyleDefs(1)   =   ":id=0,.fgcolor=&H80000008&,.wraptext=0,.locked=0,.transparentBmp=0"
               _StyleDefs(2)   =   ":id=0,.fgpicPosition=0,.bgpicMode=0,.appearance=0,.borderSize=0,.ellipsis=0"
               _StyleDefs(3)   =   ":id=0,.borderColor=&H80000005&,.borderType=0,.bold=0,.fontsize=825,.italic=0"
               _StyleDefs(4)   =   ":id=0,.underline=0,.strikethrough=0,.charset=0"
               _StyleDefs(5)   =   ":id=0,.fontname=MS Sans Serif"
               _StyleDefs(6)   =   "Style:id=1,.parent=0,.namedParent=33,.fgcolor=&H0&,.borderColor=&H80000008&"
               _StyleDefs(7)   =   "CaptionStyle:id=4,.parent=2,.namedParent=37"
               _StyleDefs(8)   =   "HeadingStyle:id=2,.parent=1,.namedParent=34,.bgcolor=&HBC8A47&"
               _StyleDefs(9)   =   ":id=2,.fgcolor=&HFFFFFF&"
               _StyleDefs(10)  =   "FooterStyle:id=3,.parent=1,.namedParent=35"
               _StyleDefs(11)  =   "InactiveStyle:id=5,.parent=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
               _StyleDefs(12)  =   "SelectedStyle:id=6,.parent=1,.namedParent=36,.fgcolor=&HFFFFFF&"
               _StyleDefs(13)  =   "EditorStyle:id=7,.parent=1,.fgcolor=&HFFFF&"
               _StyleDefs(14)  =   "HighlightRowStyle:id=8,.parent=1,.namedParent=38,.bgcolor=&HD69A69&"
               _StyleDefs(15)  =   "EvenRowStyle:id=9,.parent=1,.namedParent=39,.bgcolor=&HF8EDDE&"
               _StyleDefs(16)  =   "OddRowStyle:id=10,.parent=1,.namedParent=40,.fgcolor=&H646464&"
               _StyleDefs(17)  =   "RecordSelectorStyle:id=11,.parent=2,.namedParent=41"
               _StyleDefs(18)  =   "FilterBarStyle:id=12,.parent=1,.namedParent=42"
               _StyleDefs(19)  =   "Splits(0).Style:id=13,.parent=1,.fgcolor=&H80000014&"
               _StyleDefs(20)  =   "Splits(0).CaptionStyle:id=22,.parent=4"
               _StyleDefs(21)  =   "Splits(0).HeadingStyle:id=14,.parent=2"
               _StyleDefs(22)  =   "Splits(0).FooterStyle:id=15,.parent=3"
               _StyleDefs(23)  =   "Splits(0).InactiveStyle:id=16,.parent=5"
               _StyleDefs(24)  =   "Splits(0).SelectedStyle:id=18,.parent=6,.bgcolor=&H80000011&"
               _StyleDefs(25)  =   ":id=18,.fgcolor=&H80000007&"
               _StyleDefs(26)  =   "Splits(0).EditorStyle:id=17,.parent=7,.fgcolor=&H575757&"
               _StyleDefs(27)  =   "Splits(0).HighlightRowStyle:id=19,.parent=8,.bgcolor=&H8000000D&,.wraptext=-1"
               _StyleDefs(28)  =   "Splits(0).EvenRowStyle:id=20,.parent=9,.fgcolor=&H0&"
               _StyleDefs(29)  =   "Splits(0).OddRowStyle:id=21,.parent=10"
               _StyleDefs(30)  =   "Splits(0).RecordSelectorStyle:id=23,.parent=11"
               _StyleDefs(31)  =   "Splits(0).FilterBarStyle:id=24,.parent=12,.wraptext=-1"
               _StyleDefs(32)  =   "Splits(0).Columns(0).Style:id=28,.parent=13,.alignment=1"
               _StyleDefs(33)  =   "Splits(0).Columns(0).HeadingStyle:id=25,.parent=14"
               _StyleDefs(34)  =   "Splits(0).Columns(0).FooterStyle:id=26,.parent=15"
               _StyleDefs(35)  =   "Splits(0).Columns(0).EditorStyle:id=27,.parent=17"
               _StyleDefs(36)  =   "Splits(0).Columns(1).Style:id=32,.parent=13"
               _StyleDefs(37)  =   "Splits(0).Columns(1).HeadingStyle:id=29,.parent=14"
               _StyleDefs(38)  =   "Splits(0).Columns(1).FooterStyle:id=30,.parent=15"
               _StyleDefs(39)  =   "Splits(0).Columns(1).EditorStyle:id=31,.parent=17"
               _StyleDefs(40)  =   "Splits(0).Columns(2).Style:id=46,.parent=13"
               _StyleDefs(41)  =   "Splits(0).Columns(2).HeadingStyle:id=43,.parent=14"
               _StyleDefs(42)  =   "Splits(0).Columns(2).FooterStyle:id=44,.parent=15"
               _StyleDefs(43)  =   "Splits(0).Columns(2).EditorStyle:id=45,.parent=17"
               _StyleDefs(44)  =   "Named:id=33:Normal"
               _StyleDefs(45)  =   ":id=33,.parent=0"
               _StyleDefs(46)  =   "Named:id=34:Heading"
               _StyleDefs(47)  =   ":id=34,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
               _StyleDefs(48)  =   ":id=34,.wraptext=-1"
               _StyleDefs(49)  =   "Named:id=35:Footing"
               _StyleDefs(50)  =   ":id=35,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
               _StyleDefs(51)  =   "Named:id=36:Selected"
               _StyleDefs(52)  =   ":id=36,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
               _StyleDefs(53)  =   "Named:id=37:Caption"
               _StyleDefs(54)  =   ":id=37,.parent=34,.alignment=2"
               _StyleDefs(55)  =   "Named:id=38:HighlightRow"
               _StyleDefs(56)  =   ":id=38,.parent=33,.bgcolor=&HFF0000&,.fgcolor=&H8000000E&,.borderColor=&HFF2B2B&"
               _StyleDefs(57)  =   "Named:id=39:EvenRow"
               _StyleDefs(58)  =   ":id=39,.parent=33,.bgcolor=&HFFFF00&"
               _StyleDefs(59)  =   "Named:id=40:OddRow"
               _StyleDefs(60)  =   ":id=40,.parent=33"
               _StyleDefs(61)  =   "Named:id=41:RecordSelector"
               _StyleDefs(62)  =   ":id=41,.parent=34"
               _StyleDefs(63)  =   "Named:id=42:FilterBar"
               _StyleDefs(64)  =   ":id=42,.parent=33,.fgcolor=&H80000005&"
            End
            Begin VB.Label Label37 
               Caption         =   "sec"
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
               Left            =   6540
               TabIndex        =   50
               Top             =   975
               Width           =   375
            End
            Begin VB.Label lbl_CPRP 
               Caption         =   "Rate/insertion "
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
               Height          =   375
               Left            =   180
               TabIndex        =   49
               Top             =   1335
               Width           =   1215
            End
            Begin VB.Label Label6 
               Caption         =   "Station Code  "
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
               Height          =   375
               Left            =   180
               TabIndex        =   48
               Top             =   240
               Width           =   1215
            End
            Begin VB.Label Label7 
               Caption         =   "Station Name "
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
               Height          =   375
               Left            =   4440
               TabIndex        =   47
               Top             =   255
               Width           =   1245
            End
            Begin VB.Label Label8 
               Caption         =   "Spot Type  "
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
               Height          =   375
               Left            =   180
               TabIndex        =   46
               Top             =   975
               Width           =   1215
            End
            Begin VB.Label Lbl_Version_TV 
               Caption         =   "Version  "
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
               Height          =   285
               Left            =   4440
               TabIndex        =   45
               Top             =   615
               Width           =   1215
            End
            Begin VB.Label Label10 
               Caption         =   "Duration  "
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
               Height          =   375
               Left            =   4440
               TabIndex        =   44
               Top             =   960
               Width           =   1215
            End
            Begin VB.Label Label25 
               Caption         =   "Nett"
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
               Left            =   3270
               TabIndex        =   43
               Top             =   1755
               Width           =   615
            End
            Begin VB.Label Label28 
               Caption         =   "Gross"
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
               Left            =   3270
               TabIndex        =   42
               Top             =   1395
               Width           =   660
            End
            Begin VB.Label Label54 
               Caption         =   "Market  "
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
               Height          =   375
               Left            =   180
               TabIndex        =   41
               Top             =   600
               Width           =   1215
            End
         End
         Begin VB.OptionButton optCNDetail 
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
            Height          =   240
            Left            =   -73635
            TabIndex        =   21
            Top             =   450
            Width           =   900
         End
         Begin VB.OptionButton optCNBrief 
            Caption         =   "Brief"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Left            =   -72735
            TabIndex        =   20
            Top             =   450
            Width           =   1005
         End
         Begin VB.Frame fra_RDByStation 
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   6060
            Left            =   -74895
            TabIndex        =   98
            Top             =   675
            Width           =   9135
            Begin VB.ListBox lstRDRateType 
               Appearance      =   0  'Flat
               BackColor       =   &H8000000F&
               Height          =   1380
               Left            =   5700
               Style           =   1  'Checkbox
               TabIndex        =   111
               Top             =   405
               Width           =   1695
            End
            Begin VB.PictureBox Picture11 
               Height          =   285
               Left            =   8610
               ScaleHeight     =   225
               ScaleWidth      =   330
               TabIndex        =   109
               Top             =   2640
               Width           =   390
               Begin VB.CommandButton cmdNewRadioMaterial2 
                  Caption         =   "+"
                  BeginProperty Font 
                     Name            =   "Tahoma"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   225
                  Left            =   0
                  TabIndex        =   110
                  Top             =   0
                  Width           =   330
               End
            End
            Begin VB.ComboBox cboRDVersion2 
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
               Left            =   5715
               Style           =   2  'Dropdown List
               TabIndex        =   108
               Top             =   2625
               Width           =   2895
            End
            Begin VB.TextBox txtRDDuration2 
               Alignment       =   1  'Right Justify
               Appearance      =   0  'Flat
               BackColor       =   &H00E0E0E0&
               BeginProperty DataFormat 
                  Type            =   1
                  Format          =   "0"
                  HaveTrueFalseNull=   0
                  FirstDayOfWeek  =   0
                  FirstWeekOfYear =   0
                  LCID            =   1033
                  SubFormatType   =   1
               EndProperty
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   285
               Left            =   5730
               Locked          =   -1  'True
               TabIndex        =   107
               Text            =   "0"
               Top             =   3180
               Width           =   675
            End
            Begin VB.PictureBox Picture12 
               Height          =   360
               Left            =   6375
               ScaleHeight     =   300
               ScaleWidth      =   2580
               TabIndex        =   103
               Top             =   3525
               Visible         =   0   'False
               Width           =   2640
               Begin VB.CommandButton CmdDeleteRD2 
                  Caption         =   "&Delete"
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
                  Left            =   1725
                  TabIndex        =   106
                  ToolTipText     =   "delete print media"
                  Top             =   0
                  Width           =   855
               End
               Begin VB.CommandButton cmdAddRD2 
                  Caption         =   "&Add"
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
                  TabIndex        =   105
                  ToolTipText     =   "add print media"
                  Top             =   0
                  Width           =   855
               End
               Begin VB.CommandButton CmdEditRD2 
                  Caption         =   "&Edit"
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
                  Left            =   855
                  TabIndex        =   104
                  ToolTipText     =   "delete print media"
                  Top             =   0
                  Width           =   855
               End
            End
            Begin VB.ComboBox cboRDDuration2 
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
               Left            =   6360
               Style           =   2  'Dropdown List
               TabIndex        =   102
               Top             =   5430
               Visible         =   0   'False
               Width           =   2355
            End
            Begin VB.ComboBox CboRDSpotType2 
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
               Left            =   5715
               Style           =   2  'Dropdown List
               TabIndex        =   101
               Top             =   2055
               Width           =   2910
            End
            Begin VB.TextBox txtRDRPSGross2 
               Alignment       =   1  'Right Justify
               Appearance      =   0  'Flat
               BackColor       =   &H00E0E0E0&
               BeginProperty DataFormat 
                  Type            =   1
                  Format          =   "0"
                  HaveTrueFalseNull=   0
                  FirstDayOfWeek  =   0
                  FirstWeekOfYear =   0
                  LCID            =   1033
                  SubFormatType   =   1
               EndProperty
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
               Height          =   285
               Left            =   7470
               TabIndex        =   100
               Text            =   "0.00"
               Top             =   405
               Width           =   1110
            End
            Begin VB.TextBox txtRDRPS2 
               Alignment       =   1  'Right Justify
               Appearance      =   0  'Flat
               BackColor       =   &H00E0E0E0&
               BeginProperty DataFormat 
                  Type            =   1
                  Format          =   "0"
                  HaveTrueFalseNull=   0
                  FirstDayOfWeek  =   0
                  FirstWeekOfYear =   0
                  LCID            =   1033
                  SubFormatType   =   1
               EndProperty
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
               Height          =   285
               Left            =   7470
               TabIndex        =   99
               Text            =   "0.00"
               Top             =   765
               Width           =   1110
            End
            Begin ComctlLib.ListView lvRDSelectedStation 
               DragIcon        =   "Frm_MPActivityDetail.frx":02A4
               Height          =   3195
               Left            =   3285
               TabIndex        =   112
               Top             =   405
               Width           =   2325
               _ExtentX        =   4101
               _ExtentY        =   5636
               View            =   3
               Sorted          =   -1  'True
               LabelWrap       =   -1  'True
               HideSelection   =   -1  'True
               HideColumnHeaders=   -1  'True
               _Version        =   327682
               ForeColor       =   -2147483640
               BackColor       =   -2147483643
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
               NumItems        =   1
               BeginProperty ColumnHeader(1) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
                  Key             =   ""
                  Object.Tag             =   ""
                  Text            =   "Station Name"
                  Object.Width           =   5292
               EndProperty
            End
            Begin ComctlLib.TreeView trvRDStationCatalog 
               DragIcon        =   "Frm_MPActivityDetail.frx":05AE
               Height          =   3195
               Left            =   90
               TabIndex        =   113
               Top             =   405
               Width           =   3210
               _ExtentX        =   5662
               _ExtentY        =   5636
               _Version        =   327682
               LabelEdit       =   1
               Style           =   7
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
            Begin MSFlexGridLib.MSFlexGrid FGRDMedium2 
               Height          =   1890
               Left            =   105
               TabIndex        =   114
               Top             =   3960
               Width           =   8895
               _ExtentX        =   15690
               _ExtentY        =   3334
               _Version        =   393216
               Cols            =   8
               FixedCols       =   3
               BackColorFixed  =   12356167
               ForeColorFixed  =   16777215
               FocusRect       =   2
               AllowUserResizing=   1
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
            Begin VB.Label lblRDPathInfo 
               BackColor       =   &H00E0E0E0&
               Caption         =   "Path Info : "
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   210
               Left            =   90
               TabIndex        =   125
               Top             =   3615
               Width           =   5520
            End
            Begin VB.Label lblRDselectedstation 
               Caption         =   "Selected Station :"
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
               Height          =   210
               Left            =   3315
               TabIndex        =   124
               Top             =   135
               Width           =   2085
            End
            Begin VB.Label lblRDavailablestation 
               Caption         =   "Available Station :"
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
               Height          =   210
               Left            =   120
               TabIndex        =   123
               Top             =   135
               Width           =   2070
            End
            Begin VB.Label Label45 
               Caption         =   "Rate Type :"
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
               Height          =   210
               Left            =   5715
               TabIndex        =   122
               Top             =   135
               Width           =   1515
            End
            Begin VB.Label Label46 
               Caption         =   "Version  :"
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
               Height          =   210
               Left            =   5715
               TabIndex        =   121
               Top             =   2400
               Width           =   795
            End
            Begin VB.Label Label47 
               Caption         =   "Duration  :"
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
               Height          =   270
               Left            =   5730
               TabIndex        =   120
               Top             =   2955
               Width           =   900
            End
            Begin VB.Label Label48 
               Caption         =   "sec"
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
               Left            =   6480
               TabIndex        =   119
               Top             =   3255
               Width           =   375
            End
            Begin VB.Label Label49 
               Caption         =   "Spot Type  :"
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
               Height          =   210
               Left            =   5715
               TabIndex        =   118
               Top             =   1815
               Width           =   1245
            End
            Begin VB.Label Label50 
               Caption         =   "Rate/insertion  :"
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
               Height          =   270
               Left            =   7470
               TabIndex        =   117
               Top             =   135
               Width           =   1530
            End
            Begin VB.Label Label51 
               Caption         =   "Nett"
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
               Left            =   8655
               TabIndex        =   116
               Top             =   810
               Width           =   420
            End
            Begin VB.Label Label52 
               Caption         =   "Gross"
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
               Left            =   8655
               TabIndex        =   115
               Top             =   450
               Width           =   420
            End
         End
         Begin VB.Frame fra_RD 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   6060
            Left            =   -74895
            TabIndex        =   126
            Top             =   675
            Width           =   9135
            Begin VB.ComboBox cboRDArea 
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
               Left            =   1515
               Style           =   2  'Dropdown List
               TabIndex        =   143
               Top             =   240
               Width           =   2190
            End
            Begin VB.ComboBox CboRDSpotType 
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
               Left            =   1515
               Style           =   2  'Dropdown List
               TabIndex        =   142
               Top             =   615
               Width           =   2175
            End
            Begin VB.ComboBox cboRDVersion 
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
               Left            =   6045
               Style           =   2  'Dropdown List
               TabIndex        =   141
               Top             =   615
               Width           =   1950
            End
            Begin VB.TextBox txtRDStation 
               Alignment       =   1  'Right Justify
               Appearance      =   0  'Flat
               BackColor       =   &H00E0E0E0&
               BeginProperty DataFormat 
                  Type            =   1
                  Format          =   "0"
                  HaveTrueFalseNull=   0
                  FirstDayOfWeek  =   0
                  FirstWeekOfYear =   0
                  LCID            =   1033
                  SubFormatType   =   1
               EndProperty
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   285
               Left            =   6045
               Locked          =   -1  'True
               TabIndex        =   140
               Top             =   270
               Width           =   735
            End
            Begin VB.PictureBox Picture3 
               Height          =   360
               Left            =   6420
               ScaleHeight     =   300
               ScaleWidth      =   2565
               TabIndex        =   136
               Top             =   1395
               Visible         =   0   'False
               Width           =   2625
               Begin VB.CommandButton cmdAddRD 
                  Caption         =   "&Add"
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
                  TabIndex        =   139
                  ToolTipText     =   "add radio area"
                  Top             =   0
                  Width           =   855
               End
               Begin VB.CommandButton cmdDeleteRD 
                  Caption         =   "&Delete"
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
                  Left            =   1710
                  TabIndex        =   138
                  ToolTipText     =   "delete radio area"
                  Top             =   0
                  Width           =   855
               End
               Begin VB.CommandButton CmdEditRD 
                  Caption         =   "&Edit"
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
                  Left            =   840
                  TabIndex        =   137
                  ToolTipText     =   "delete radio area"
                  Top             =   0
                  Width           =   855
               End
            End
            Begin VB.ComboBox cboRDStation 
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
               Left            =   8010
               Style           =   2  'Dropdown List
               TabIndex        =   135
               Top             =   2625
               Visible         =   0   'False
               Width           =   855
            End
            Begin VB.PictureBox Picture7 
               Height          =   315
               Left            =   3735
               ScaleHeight     =   255
               ScaleWidth      =   330
               TabIndex        =   133
               Top             =   255
               Width           =   390
               Begin VB.CommandButton cmdRDNewArea 
                  Caption         =   "+"
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
                  Left            =   -15
                  TabIndex        =   134
                  Top             =   0
                  Width           =   345
               End
            End
            Begin VB.TextBox txtRDRPS 
               Alignment       =   1  'Right Justify
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
               Height          =   285
               Left            =   1515
               TabIndex        =   132
               Text            =   "0.00"
               Top             =   1395
               Width           =   1605
            End
            Begin VB.TextBox txtRDDuration 
               Alignment       =   1  'Right Justify
               Appearance      =   0  'Flat
               BackColor       =   &H00E0E0E0&
               BeginProperty DataFormat 
                  Type            =   1
                  Format          =   "0"
                  HaveTrueFalseNull=   0
                  FirstDayOfWeek  =   0
                  FirstWeekOfYear =   0
                  LCID            =   1033
                  SubFormatType   =   1
               EndProperty
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   285
               Left            =   6045
               Locked          =   -1  'True
               TabIndex        =   131
               Top             =   1005
               Width           =   735
            End
            Begin VB.PictureBox Picture2 
               Height          =   315
               Left            =   8010
               ScaleHeight     =   255
               ScaleWidth      =   315
               TabIndex        =   129
               Top             =   615
               Width           =   375
               Begin VB.CommandButton cmdNewRadioMaterial 
                  Caption         =   "+"
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
                  Left            =   0
                  TabIndex        =   130
                  Top             =   0
                  Width           =   315
               End
            End
            Begin VB.ComboBox cboRDduration 
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
               Left            =   6375
               Style           =   2  'Dropdown List
               TabIndex        =   128
               Top             =   2715
               Visible         =   0   'False
               Width           =   855
            End
            Begin VB.TextBox txtRDRPSGross 
               Alignment       =   1  'Right Justify
               Appearance      =   0  'Flat
               BackColor       =   &H00E0E0E0&
               BeginProperty DataFormat 
                  Type            =   1
                  Format          =   "0"
                  HaveTrueFalseNull=   0
                  FirstDayOfWeek  =   0
                  FirstWeekOfYear =   0
                  LCID            =   1033
                  SubFormatType   =   1
               EndProperty
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   285
               Left            =   1515
               TabIndex        =   127
               Text            =   "0.00"
               Top             =   1035
               Width           =   1605
            End
            Begin MSFlexGridLib.MSFlexGrid FGRDMedium 
               Height          =   4020
               Left            =   120
               TabIndex        =   144
               Top             =   1815
               Width           =   8895
               _ExtentX        =   15690
               _ExtentY        =   7091
               _Version        =   393216
               Cols            =   7
               FixedCols       =   3
               BackColorFixed  =   12356167
               FocusRect       =   2
               AllowUserResizing=   1
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
            Begin VB.Label Label15 
               Caption         =   "Spot Type  "
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
               Height          =   375
               Left            =   180
               TabIndex        =   153
               Top             =   645
               Width           =   1290
            End
            Begin VB.Label Label16 
               Caption         =   "Version  "
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
               Height          =   315
               Left            =   4710
               TabIndex        =   152
               Top             =   630
               Width           =   1230
            End
            Begin VB.Label Label12 
               Caption         =   "Area  "
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
               Height          =   300
               Left            =   180
               TabIndex        =   151
               Top             =   270
               Width           =   1290
            End
            Begin VB.Label Label14 
               Caption         =   "Station "
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
               Height          =   375
               Left            =   4710
               TabIndex        =   150
               Top             =   270
               Width           =   1230
            End
            Begin VB.Label Label29 
               Caption         =   "Rate/insertion "
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
               Height          =   300
               Left            =   180
               TabIndex        =   149
               Top             =   1050
               Width           =   1290
            End
            Begin VB.Label Label35 
               Caption         =   "Duration "
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
               Height          =   375
               Left            =   4710
               TabIndex        =   148
               Top             =   1020
               Width           =   1230
            End
            Begin VB.Label Label36 
               Caption         =   "sec"
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
               Left            =   6870
               TabIndex        =   147
               Top             =   1035
               Width           =   420
            End
            Begin VB.Label Label38 
               Caption         =   "Gross"
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
               Left            =   3225
               TabIndex        =   146
               Top             =   1065
               Width           =   510
            End
            Begin VB.Label Label39 
               Caption         =   "Nett"
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
               Left            =   3300
               TabIndex        =   145
               Top             =   1440
               Width           =   510
            End
         End
         Begin VB.Frame FramePR 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   6390
            Left            =   -74880
            TabIndex        =   156
            Top             =   345
            Width           =   9135
            Begin VB.TextBox txtPRMediaNameSearch 
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
               Left            =   2400
               MaxLength       =   50
               TabIndex        =   161
               Top             =   210
               Visible         =   0   'False
               Width           =   2445
            End
            Begin VB.PictureBox PicPRAddDelete 
               Height          =   540
               Left            =   6120
               ScaleHeight     =   480
               ScaleWidth      =   2580
               TabIndex        =   176
               Top             =   2025
               Visible         =   0   'False
               Width           =   2640
               Begin VB.CommandButton cmdAddPR 
                  Caption         =   "&Add"
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
                  Left            =   15
                  TabIndex        =   179
                  ToolTipText     =   "add print media"
                  Top             =   0
                  Width           =   855
               End
               Begin VB.CommandButton cmdDeletePR 
                  Caption         =   "&Delete"
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
                  Left            =   1725
                  TabIndex        =   178
                  ToolTipText     =   "delete print media"
                  Top             =   0
                  Width           =   855
               End
               Begin VB.CommandButton cmdEditPR 
                  Caption         =   "&Edit"
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
                  Left            =   870
                  TabIndex        =   177
                  ToolTipText     =   "delete print media"
                  Top             =   0
                  Width           =   855
               End
            End
            Begin VB.TextBox txtPRRate 
               Alignment       =   1  'Right Justify
               Appearance      =   0  'Flat
               BackColor       =   &H00E0E0E0&
               BeginProperty DataFormat 
                  Type            =   1
                  Format          =   "0"
                  HaveTrueFalseNull=   0
                  FirstDayOfWeek  =   0
                  FirstWeekOfYear =   0
                  LCID            =   1033
                  SubFormatType   =   1
               EndProperty
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
               Height          =   285
               Left            =   7815
               TabIndex        =   175
               Text            =   "0.00"
               Top             =   1575
               Width           =   1140
            End
            Begin VB.TextBox txtPRRateGross 
               Alignment       =   1  'Right Justify
               Appearance      =   0  'Flat
               BackColor       =   &H00E0E0E0&
               BeginProperty DataFormat 
                  Type            =   1
                  Format          =   "0"
                  HaveTrueFalseNull=   0
                  FirstDayOfWeek  =   0
                  FirstWeekOfYear =   0
                  LCID            =   1033
                  SubFormatType   =   1
               EndProperty
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
               Height          =   285
               Left            =   7815
               TabIndex        =   174
               Text            =   "0.00"
               Top             =   1245
               Width           =   1140
            End
            Begin VB.TextBox txtPRCol 
               Alignment       =   1  'Right Justify
               Appearance      =   0  'Flat
               BackColor       =   &H00E0E0E0&
               BeginProperty DataFormat 
                  Type            =   1
                  Format          =   "0"
                  HaveTrueFalseNull=   0
                  FirstDayOfWeek  =   0
                  FirstWeekOfYear =   0
                  LCID            =   1033
                  SubFormatType   =   1
               EndProperty
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   285
               Left            =   5580
               Locked          =   -1  'True
               TabIndex        =   173
               Text            =   "0"
               Top             =   1260
               Width           =   1005
            End
            Begin VB.TextBox txtPRMM 
               Alignment       =   1  'Right Justify
               Appearance      =   0  'Flat
               BackColor       =   &H00E0E0E0&
               BeginProperty DataFormat 
                  Type            =   1
                  Format          =   "0"
                  HaveTrueFalseNull=   0
                  FirstDayOfWeek  =   0
                  FirstWeekOfYear =   0
                  LCID            =   1033
                  SubFormatType   =   1
               EndProperty
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   285
               Left            =   5580
               Locked          =   -1  'True
               TabIndex        =   172
               Text            =   "0"
               Top             =   1590
               Width           =   1005
            End
            Begin VB.ComboBox cboPRSpotType 
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
               Left            =   1230
               Style           =   2  'Dropdown List
               TabIndex        =   171
               Top             =   1575
               Width           =   2460
            End
            Begin VB.ComboBox cboPRVersion 
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
               Left            =   1230
               Style           =   2  'Dropdown List
               TabIndex        =   170
               Top             =   1245
               Width           =   2460
            End
            Begin VB.TextBox TxtPRMinSize 
               Alignment       =   1  'Right Justify
               Appearance      =   0  'Flat
               BackColor       =   &H00E0E0E0&
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
               Left            =   5580
               MaxLength       =   50
               TabIndex        =   169
               Text            =   "0"
               Top             =   915
               Width           =   1005
            End
            Begin VB.TextBox TxtPRSatuan 
               Appearance      =   0  'Flat
               BackColor       =   &H00E0E0E0&
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
               MaxLength       =   50
               TabIndex        =   168
               Top             =   900
               Width           =   1830
            End
            Begin VB.TextBox TxtPRPaper 
               Appearance      =   0  'Flat
               BackColor       =   &H00E0E0E0&
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
               Left            =   5580
               MaxLength       =   50
               TabIndex        =   167
               Top             =   570
               Width           =   1830
            End
            Begin VB.TextBox TxtPRColor 
               Appearance      =   0  'Flat
               BackColor       =   &H00E0E0E0&
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
               Left            =   5580
               MaxLength       =   50
               TabIndex        =   166
               Top             =   210
               Width           =   1830
            End
            Begin VB.TextBox TxtPRSize 
               Appearance      =   0  'Flat
               BackColor       =   &H00E0E0E0&
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
               MaxLength       =   50
               TabIndex        =   165
               Top             =   555
               Width           =   1275
            End
            Begin VB.TextBox TxtPRMediaCode 
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
               Left            =   1245
               MaxLength       =   50
               TabIndex        =   164
               Top             =   210
               Width           =   2445
            End
            Begin VB.PictureBox PicPRNewMaterial 
               Height          =   315
               Left            =   3735
               ScaleHeight     =   255
               ScaleWidth      =   300
               TabIndex        =   162
               Top             =   1245
               Width           =   360
               Begin VB.CommandButton cmdNewPRMaterial 
                  Caption         =   "+"
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
                  Left            =   0
                  TabIndex        =   163
                  Top             =   0
                  Width           =   315
               End
            End
            Begin VB.ComboBox cboPRVersionCol 
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
               Left            =   4755
               Style           =   2  'Dropdown List
               TabIndex        =   160
               Top             =   4065
               Visible         =   0   'False
               Width           =   1620
            End
            Begin VB.ComboBox cboPRVersionMM 
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
               Left            =   3090
               Style           =   2  'Dropdown List
               TabIndex        =   159
               Top             =   4065
               Visible         =   0   'False
               Width           =   1620
            End
            Begin VB.TextBox txtPRMediaName 
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
               Left            =   6420
               MaxLength       =   50
               TabIndex        =   158
               Top             =   4065
               Visible         =   0   'False
               Width           =   2445
            End
            Begin VB.TextBox txtPRIsMMC 
               Height          =   315
               Left            =   930
               TabIndex        =   157
               Top             =   4065
               Visible         =   0   'False
               Width           =   1935
            End
            Begin MSFlexGridLib.MSFlexGrid FGPRRate 
               Height          =   5670
               Left            =   135
               TabIndex        =   181
               Top             =   540
               Visible         =   0   'False
               Width           =   8850
               _ExtentX        =   15610
               _ExtentY        =   10001
               _Version        =   393216
               BackColorFixed  =   12356167
               ForeColorFixed  =   16777215
               AllowUserResizing=   3
               Appearance      =   0
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
            Begin MSFlexGridLib.MSFlexGrid FGPRMedium 
               Height          =   4050
               Left            =   135
               TabIndex        =   180
               Top             =   2160
               Width           =   8850
               _ExtentX        =   15610
               _ExtentY        =   7144
               _Version        =   393216
               Cols            =   10
               FixedCols       =   3
               BackColorFixed  =   12356167
               ForeColorFixed  =   16777215
               FocusRect       =   2
               AllowUserResizing=   1
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
            Begin VB.Label Label19 
               Alignment       =   1  'Right Justify
               Caption         =   "Min Size "
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
               Height          =   195
               Left            =   4665
               TabIndex        =   193
               Top             =   945
               Width           =   855
            End
            Begin VB.Label Label18 
               Alignment       =   1  'Right Justify
               Caption         =   "MM "
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
               Height          =   195
               Left            =   4920
               TabIndex        =   192
               Top             =   1590
               Width           =   615
            End
            Begin VB.Label Label11 
               Alignment       =   1  'Right Justify
               Caption         =   "CL "
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
               Height          =   195
               Left            =   5085
               TabIndex        =   191
               Top             =   1305
               Width           =   450
            End
            Begin VB.Label Label20 
               Caption         =   "Satuan "
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
               Height          =   195
               Left            =   240
               TabIndex        =   190
               Top             =   930
               Width           =   945
            End
            Begin VB.Label Label22 
               Alignment       =   1  'Right Justify
               Caption         =   "Paper "
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
               Height          =   195
               Left            =   4680
               TabIndex        =   189
               Top             =   615
               Width           =   855
            End
            Begin VB.Label Label30 
               Alignment       =   1  'Right Justify
               Caption         =   "Gross Rate "
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
               Height          =   195
               Left            =   6660
               TabIndex        =   188
               Top             =   1290
               Width           =   1095
            End
            Begin VB.Label Label33 
               Alignment       =   1  'Right Justify
               Caption         =   "Nett Rate "
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
               Height          =   195
               Left            =   6780
               TabIndex        =   187
               Top             =   1575
               Width           =   990
            End
            Begin VB.Label Label40 
               AutoSize        =   -1  'True
               Caption         =   "Version "
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
               Height          =   195
               Left            =   240
               TabIndex        =   186
               Top             =   1290
               Width           =   570
            End
            Begin VB.Label Label41 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               Caption         =   "Color "
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
               Height          =   195
               Left            =   5145
               TabIndex        =   185
               Top             =   300
               Width           =   420
            End
            Begin VB.Label Label53 
               AutoSize        =   -1  'True
               Caption         =   "Size "
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
               Height          =   195
               Left            =   240
               TabIndex        =   184
               Top             =   585
               Width           =   330
            End
            Begin VB.Label lblPrintCode 
               AutoSize        =   -1  'True
               Caption         =   "Print Code "
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
               Height          =   195
               Left            =   240
               TabIndex        =   183
               Top             =   240
               Width           =   795
            End
            Begin VB.Label Label55 
               AutoSize        =   -1  'True
               Caption         =   "Spot Type "
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
               Height          =   195
               Left            =   240
               TabIndex        =   182
               Top             =   1605
               Width           =   780
            End
         End
         Begin VB.Label Label44 
            Caption         =   "Selection By "
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
            Height          =   285
            Left            =   -74685
            TabIndex        =   155
            Top             =   480
            Width           =   1185
         End
         Begin VB.Label Label9 
            Alignment       =   1  'Right Justify
            Caption         =   "Entry Mode :"
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
            Height          =   285
            Left            =   -74850
            TabIndex        =   154
            Top             =   435
            Width           =   1080
         End
      End
      Begin VB.Frame Frame1 
         Height          =   1320
         Left            =   120
         TabIndex        =   1
         Top             =   45
         Width           =   9360
         Begin VB.TextBox txtMPNumber 
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
            Height          =   285
            Left            =   1680
            Locked          =   -1  'True
            TabIndex        =   6
            Text            =   "txtMPNumber"
            Top             =   195
            Width           =   1575
         End
         Begin VB.TextBox txtMPTask 
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
            Height          =   285
            Left            =   1680
            Locked          =   -1  'True
            TabIndex        =   5
            Text            =   "txtMPTask"
            Top             =   555
            Width           =   3615
         End
         Begin VB.TextBox txtMPActivity 
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
            Height          =   285
            Left            =   6360
            Locked          =   -1  'True
            TabIndex        =   4
            Text            =   "txtMPActivity"
            Top             =   570
            Width           =   2655
         End
         Begin VB.TextBox txtBrandVariant 
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
            Height          =   285
            Left            =   1680
            Locked          =   -1  'True
            TabIndex        =   3
            Text            =   "txtBrandVariant"
            Top             =   915
            Width           =   3615
         End
         Begin VB.TextBox txtMedium 
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
            Height          =   285
            Left            =   6360
            Locked          =   -1  'True
            TabIndex        =   2
            Top             =   930
            Width           =   2655
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
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
            Left            =   240
            TabIndex        =   11
            Top             =   240
            Width           =   855
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Task Desc. "
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
            Left            =   240
            TabIndex        =   10
            Top             =   570
            Width           =   825
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Brand Variant "
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
            Left            =   240
            TabIndex        =   9
            Top             =   900
            Width           =   1020
         End
         Begin VB.Label Label4 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Activity "
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
            Left            =   5670
            TabIndex        =   8
            Top             =   585
            Width           =   585
         End
         Begin VB.Label Label5 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Medium "
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
            Left            =   5670
            TabIndex        =   7
            Top             =   945
            Width           =   585
         End
      End
   End
   Begin VB.Menu MnuPopupOther 
      Caption         =   "MnuPopupOther"
      Visible         =   0   'False
      Begin VB.Menu Mnu_Monthly_Budget 
         Caption         =   "Monthly Budget"
      End
   End
End
Attribute VB_Name = "Frm_MPActivityDetail"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
''*****************************************************************************
' Nama Submodul         :  Frm_MPActivityDetail
' Fungsi Submodul       :  Untuk View / Edit Activity Detail (Medium, Budget, Dimension)
' Nama Programmer       :  Sistyo
' Tgl Pembuatan         :  11 Agustus 2004
' Last Update           :  11 Agustus 2004/Sistyo
'******************************************************************************
Option Explicit
Dim recTemp As New ADODB.Recordset
Dim recTV As New ADODB.Recordset
Dim recOT As New ADODB.Recordset
Dim recCN As New ADODB.Recordset
Dim recCN2 As New ADODB.Recordset
Dim blnOptRDStationFirstClick As Boolean
Dim blnOptCnBriefFirstClick As Boolean
Dim blnSSTabMediumFirstClick() As Boolean

Dim strPRNettRate As String
Dim strPRGrossRate As String
Dim strMode As String
Dim intTabEdit As Integer

Private Sub cboTVMarketName_Click()
'<CSCM>
'********************************************************************************
'Procedure Name     : cboTVMarketName_Click
'Procedure Function : ---
'Input Parameter    : ---
'Output Parameter   : ---
'Date               : 4/4/2016
'LastUpdate/By      : Tedi / Kreatif
'Name Before        : -
'********************************************************************************
'</CSCM>

    cboTVMarketCode.ListIndex = cboTVMarketName.ListIndex

End Sub

Private Sub cmdEditCN_Click()
'<CSCM>
'********************************************************************************
'Procedure Name     : cmdEditCN_Click
'Procedure Function : ---
'Input Parameter    : ---
'Output Parameter   : ---
'Date               : 4/4/2016
'LastUpdate/By      : Tedi / Kreatif
'Name Before        : -
'********************************************************************************
'</CSCM>
    Dim recTemp As New ADODB.Recordset
    Dim strSql As String
    
    If cmdEditCN.Caption = "&Edit" Then
        cmdEditCN.Caption = "&Save"
        cmdDeleteCN.Caption = "&Cancel"
        cmdAddCN.Enabled = False
        tdg_FGCNMedium.Enabled = False
        strSql = "select * from mp_medium_detail where mp_medium_detail_id = '" & tdg_FGCNMedium.Columns(1) & "'"
        recTemp.Open strSql, ConnERP, 1, 3
        If Not recTemp.EOF Then
            cboCNCode.Text = Trim(recTemp("cinema_code"))
            cboCNCode.Enabled = False
            cboCNName.Text = Trim(recTemp("cinema_name"))
            cboCNName.Enabled = False
            cboCNStudio.Text = Trim(recTemp("cinema_studio"))
            cboCNStudio.Enabled = False
        End If
        recTemp.Close
        
        strSql = "select * from mp_plan_dimension where mp_plan_dim_id = '" & tdg_FGCNMedium.Columns(0) & "'"
        recTemp.Open strSql, ConnERP, 1, 3
        If Not recTemp.EOF Then
            cboCNSpotType.Text = recTemp("spot_type")
            Call cboCNSpotType_Click
            cboCNSpotType.Enabled = False
            cboCNVersion.Text = recTemp("version")
            cboCNJenisDurasi = recTemp("cinema_duration")
            txtCNDuration.Text = recTemp("duration")
            txtCNRate.Text = FormatNumber(recTemp("rate_per_spot"), 2)
            txtCNRateGross.Text = FormatNumber(recTemp("gross_rate_per_spot"), 2)
        End If
        recTemp.Close
    Else
        If strMode <> "" Then
            'Save
            strSql = "update mp_plan_dimension "
            strSql = strSql & "set version = '" & Clear_String(cboCNVersion.Text) & "', "
            strSql = strSql & "cinema_duration='" & cboCNJenisDurasi.Text & "', "
            strSql = strSql & "duration=" & RemoveNumberFormat(txtCNDuration.Text) & ","
            strSql = strSql & "rate_per_spot=" & RemoveNumberFormat(txtCNRate.Text) & ","
            strSql = strSql & "gross_rate_per_spot=" & RemoveNumberFormat(txtCNRateGross.Text) & " "
            strSql = strSql & "where mp_plan_dim_id = '" & tdg_FGCNMedium.Columns(0).Value & "'"
            ConnERP.Execute strSql
            'REFRESH TAMPILAN MEDIUM
            Call ViewMediumTrueDB("CN", tdg_FGCNMedium)
            MsgBox "Data Saved!", vbExclamation, strApplication_Name
            'SETTING TOMBOL
        End If
        cmdEditCN.Caption = "&Edit"
        cmdDeleteCN.Caption = "&Delete"
        cmdAddCN.Enabled = True
        tdg_FGCNMedium.Enabled = True
        
        cboCNCode.Enabled = True
        cboCNName.Enabled = True
        cboCNStudio.Enabled = True
        cboCNSpotType.Enabled = True
        initTabCN
    End If
End Sub

Private Sub cmdEditCN2_Click()
'<CSCM>
'********************************************************************************
'Procedure Name     : cmdEditCN2_Click
'Procedure Function : ---
'Input Parameter    : ---
'Output Parameter   : ---
'Date               : 4/4/2016
'LastUpdate/By      : Tedi / Kreatif
'Name Before        : -
'********************************************************************************
'</CSCM>
    Dim recTemp As New ADODB.Recordset
    Dim strSql As String
    
    If cmdEditCN2.Caption = "&Edit" Then
        cmdEditCN2.Caption = "&Save"
        CmdDeleteCN2.Caption = "&Cancel"
        CmdAddCN2.Enabled = False
        tdg_FGCNMedium2.Enabled = False
        txtCNBrief.Text = tdg_FGCNMedium2.Columns(2)
        txtCNRate2.Text = FormatNumber(tdg_FGCNMedium2.Columns(4), 2)
        txtCNRateGross2.Text = FormatNumber(tdg_FGCNMedium2.Columns(3), 2)
    Else
        If strMode <> "" Then
            strSql = "Update mp_plan_dimension set ot_description = '" & Clear_String(txtCNBrief.Text) & "',rate_per_spot=" & RemoveNumberFormat(txtCNRate2.Text) & ",gross_rate_per_spot=" & RemoveNumberFormat(txtCNRateGross2.Text) & " where mp_plan_dim_id = '" & tdg_FGCNMedium2.Columns(0) & "'"
            ConnERP.Execute strSql
            'REFRESH TAMPILAN MEDIUM
            Call ViewMediumTrueDB("CN2", tdg_FGCNMedium2)
            MsgBox "Data Saved!", vbExclamation, strApplication_Name
        End If
        'SETTING TOMBOL
        cmdEditCN2.Caption = "&Edit"
        CmdDeleteCN2.Caption = "&Delete"
        CmdAddCN2.Enabled = True
        tdg_FGCNMedium2.Enabled = True
        txtCNBrief.Text = Empty
        txtCNRate2.Text = Empty
        txtCNRateGross2.Text = Empty
    End If
End Sub

Private Sub CmdEditOT_Click()
'<CSCM>
'********************************************************************************
'Procedure Name     : CmdEditOT_Click
'Procedure Function : ---
'Input Parameter    : ---
'Output Parameter   : ---
'Date               : 4/4/2016
'LastUpdate/By      : Tedi / Kreatif
'Name Before        : -
'********************************************************************************
'</CSCM>
    Dim recTemp As New ADODB.Recordset
    Dim strSql As String
    
    If CmdEditOT.Caption = "&Edit" Then
        CmdEditOT.Caption = "&Save"
        cmdDeleteOT.Caption = "&Cancel"
        cmdAddOT.Enabled = False
        tdg_OTMedium.Enabled = False
        txtOTDescrition.Text = tdg_OTMedium.Columns(2).Text
    Else
        If strMode <> "" Then
            'REFRESH TAMPILAN MEDIUM
            strSql = "Update mp_plan_dimension set ot_description = '" & Clear_String(txtOTDescrition.Text) & "' "
            strSql = strSql & "where mp_plan_dim_id = '" & tdg_OTMedium.Columns(0).Text & "'"
            ConnERP.Execute strSql
            Call ViewMediumTrueDB("OT", tdg_OTMedium)
            MsgBox "Data Saved!", vbExclamation, strApplication_Name
        End If
        'SETTING TOMBOL
        CmdEditOT.Caption = "&Edit"
        cmdDeleteOT.Caption = "&Delete"
        cmdAddOT.Enabled = True
        tdg_OTMedium.Enabled = True
        txtOTDescrition.Text = Empty
    End If
End Sub

Private Sub cmdEditPR_Click()
'<CSCM>
'********************************************************************************
'Procedure Name     : cmdEditPR_Click
'Procedure Function : ---
'Input Parameter    : ---
'Output Parameter   : ---
'Date               : 4/4/2016
'LastUpdate/By      : Tedi / Kreatif
'Name Before        : -
'********************************************************************************
'</CSCM>
    Dim recTemp As New ADODB.Recordset
    Dim strSql As String
    Dim print_size_name As String
    Dim print_color_name As String
    Dim print_paper_name As String
    
    If cmdEditPR.Caption = "&Edit" Then
        cmdEditPR.Caption = "&Save"
        cmdDeletePR.Caption = "&Cancel"
        cmdAddPR.Enabled = False
        FGPRMedium.Enabled = False
        'TAMPLIKAN MEDIA PRINT
        strSql = "select * from mp_medium_detail where mp_medium_detail_id = '" & FGPRMedium.TextMatrix(FGPRMedium.Row, 2) & "'"
        recTemp.Open strSql, ConnERP, 1, 3
        If Not recTemp.EOF Then
            TxtPRMediaCode.Text = recTemp("print_code")
            txtPRMediaNameSearch.Text = recTemp("media_name")
            txtPRMediaName.Text = recTemp("media_name")
        End If
        recTemp.Close
        txtPRMediaNameSearch.Locked = True
        strSql = "select * from mp_plan_dimension where mp_plan_dim_id = '" & FGPRMedium.TextMatrix(FGPRMedium.Row, 1) & "'"
        recTemp.Open strSql, ConnERP, 1, 3
        
        If Not recTemp.EOF Then
            TxtPRSatuan.Text = recTemp("print_size_code")
            TxtPRColor.Text = recTemp("print_color_code")
            TxtPRPaper.Text = recTemp("print_paper_code")
            TxtPRMinSize.Text = recTemp("print_min_size")
            txtPRCol.Text = recTemp("print_mmc_col")
            txtPRMM.Text = recTemp("print_mmc_size")
            txtPRIsMMC.Text = recTemp("print_ismmc")
            cboPRSpotType.Text = recTemp("spot_type")
            txtPRRate.Text = FormatNumber(recTemp("rate_per_spot"), 2)
            txtPRRateGross.Text = FormatNumber(recTemp("gross_rate_per_spot"), 2)
            cboPRVersion.Text = recTemp("version")
            Call cboPRSpotType_Click
            cboPRSpotType.Enabled = False
        End If
        recTemp.Close
    Else
        'Save Data
        'print_color_catalog >> color_code,color_name
        'print_size_catalog >> size_code,size_name
        'print_paper_catalog >> paper_code,paper_name
        If strMode <> "" Then
        print_size_name = ""
        print_color_name = ""
        print_paper_name = ""
        
        recTemp.Open "select size_name from print_size_catalog where size_code='" & TxtPRSatuan.Text & "'", ConnERP, 1, 3
        If Not recTemp.EOF Then print_size_name = recTemp("size_name")
            recTemp.Close
            recTemp.Open "select color_name from print_color_catalog where color_code='" & TxtPRColor.Text & "'", ConnERP, 1, 3
            If Not recTemp.EOF Then print_color_name = recTemp("color_name")
            recTemp.Close
            recTemp.Open "select paper_name from print_paper_catalog where paper_code='" & TxtPRPaper.Text & "'", ConnERP, 1, 3
            If Not recTemp.EOF Then print_paper_name = recTemp("paper_name")
            recTemp.Close
            
            If cboPRSpotType.Text = "Reguler" Then
                txtPRRate.Text = "0"
                txtPRRateGross.Text = "0"
            End If
            
            strSql = "update mp_plan_dimension set "
            strSql = strSql & "print_size_code='" & TxtPRSatuan.Text & "',print_size_name='" & Clear_String(print_size_name) & "',print_color_code='" & Clear_String(TxtPRColor.Text) & "',"
            strSql = strSql & "print_color_name='" & Clear_String(print_color_name) & "',print_paper_code='" & TxtPRPaper.Text & "',print_paper_name='" & Clear_String(print_paper_name) & "',"
            strSql = strSql & "print_min_size=" & TxtPRMinSize.Text & ",print_mmc_col=" & txtPRCol.Text & ",print_mmc_size=" & txtPRMM.Text
            strSql = strSql & ",print_ismmc=" & txtPRIsMMC.Text & ",spot_type='" & cboPRSpotType.Text & "'"
            strSql = strSql & ",rate_per_spot=" & RemoveNumberFormat(txtPRRate.Text) & ",gross_rate_per_spot=" & RemoveNumberFormat(txtPRRateGross.Text)
            strSql = strSql & ",version='" & Clear_String(cboPRVersion.Text) & "' "
            strSql = strSql & "where mp_plan_dim_id = '" & FGPRMedium.TextMatrix(FGPRMedium.Row, 1) & "'"
            
            ConnERP.Execute strSql
            Call ViewMedium("PR", FGPRMedium)
            MsgBox "Data Saved!", vbExclamation, strApplication_Name
        End If
        cmdEditPR.Caption = "&Edit"
        cmdDeletePR.Caption = "&Delete"
        cmdAddPR.Enabled = True
        FGPRMedium.Enabled = True
        txtPRMediaNameSearch.Locked = False
        cboPRSpotType.Enabled = True
        Call DisableControlTabPrint(True)
        Call EmptyTabPrint
        If FGPRMedium.Rows > 0 Then
            If FGPRMedium.Row = 0 Then FGPRMedium.Row = 1
            Call FGPRMedium_Click
        End If
        Call EnableObject(False)
        strMode = ""
    End If
End Sub

Private Sub CmdEditRD_Click()
'<CSCM>
'********************************************************************************
'Procedure Name     : CmdEditRD_Click
'Procedure Function : ---
'Input Parameter    : ---
'Output Parameter   : ---
'Date               : 4/4/2016
'LastUpdate/By      : Tedi / Kreatif
'Name Before        : -
'********************************************************************************
'</CSCM>
    Dim recTemp As New ADODB.Recordset
    Dim strSql As String
    
    If CmdEditRD.Caption = "&Edit" Then
        If strMode = "" Then Exit Sub
        CmdEditRD.Caption = "&Save"
        cmdDeleteRD.Caption = "&Cancel"
        
        cmdAddRD.Enabled = False
        FGRDMedium.Enabled = False
        cmdRDNewArea.Enabled = False
        cboRDArea.Enabled = False
        CboRDSpotType.Enabled = False
        
        'View Data
        strSql = "select * from mp_medium_detail where mp_medium_detail_id='" & FGRDMedium.TextMatrix(FGRDMedium.Row, 2) & "'"
        recTemp.Open strSql, ConnERP, 1, 3
        If Not recTemp.EOF Then
            cboRDArea.Text = Trim(recTemp("area_code"))
        End If
        recTemp.Close
        
        strSql = "select * from mp_plan_dimension where mp_plan_dim_id='" & FGRDMedium.TextMatrix(FGRDMedium.Row, 1) & "'"
        recTemp.Open strSql, ConnERP, 1, 3
        If Not recTemp.EOF Then
            txtRDStation.Text = recTemp("rd_stations")
            CboRDSpotType = recTemp("spot_type")
            cboRDVersion = recTemp("version")
            txtRDDuration = recTemp("duration")
            txtRDRPS.Text = FormatNumber(recTemp("rate_per_spot"))
            txtRDRPSGross.Text = FormatNumber(recTemp("gross_rate_per_spot"))
        End If
        recTemp.Close
    Else
            
        If strMode <> "" Then
            'SAve DAta
            strSql = "update mp_plan_dimension set version='" & Clear_String(cboRDVersion.Text) & "',duration=" & txtRDDuration.Text
            strSql = strSql & ",rate_per_spot=" & RemoveNumberFormat(txtRDRPS.Text) & ",gross_rate_per_spot=" & RemoveNumberFormat(txtRDRPSGross.Text)
            strSql = strSql & " where mp_plan_dim_id = '" & FGRDMedium.TextMatrix(FGRDMedium.Row, 1) & "'"
            ConnERP.Execute strSql
            MsgBox "Data Saved!", vbExclamation, strApplication_Name
        End If
        
        Call ViewMedium("RD", FGRDMedium)
        
        CmdEditRD.Caption = "&Edit"
        cmdDeleteRD.Caption = "&Delete"
        cmdAddRD.Enabled = True
        FGRDMedium.Enabled = True
        cmdRDNewArea.Enabled = True
        cboRDArea.Enabled = True
        CboRDSpotType.Enabled = True
        Call DisableControlTabRadio(True)
        Call EnableObject(False)
        strMode = ""
    End If

End Sub

Private Sub CmdEditRD2_Click()
'<CSCM>
'********************************************************************************
'Procedure Name     : CmdEditRD2_Click
'Procedure Function : ---
'Input Parameter    : ---
'Output Parameter   : ---
'Date               : 4/4/2016
'LastUpdate/By      : Tedi / Kreatif
'Name Before        : -
'********************************************************************************
'</CSCM>
    Dim recTemp As New ADODB.Recordset
    Dim strSql As String
    Dim i As Integer
    Dim isRateSelected As Boolean, isRateAvailable As Boolean
    Dim Station_Code As String, rate_type_code As String, rate_type_name As String
    
    If CmdEditRD2.Caption = "&Edit" Then
        CmdEditRD2.Caption = "&Save"
        CmdDeleteRD2.Caption = "&Cancel"
        cmdAddRD2.Enabled = False
        FGRDMedium2.Enabled = False
        trvRDStationCatalog.Enabled = False
        lvRDSelectedStation.ListItems.Clear
        lvRDSelectedStation.Enabled = False
        
        strSql = "select * from mp_medium_detail where mp_medium_detail_id = '" & FGRDMedium2.TextMatrix(FGRDMedium2.Row, 2) & "'"
        recTemp.Open strSql, ConnERP, 1, 3
        If Not recTemp.EOF Then
            lvRDSelectedStation.ListItems.Add , recTemp("radio_station_code"), recTemp("radio_station_name")
        End If
        recTemp.Close
        
        strSql = "select * from mp_plan_dimension where mp_plan_dim_id = '" & FGRDMedium2.TextMatrix(FGRDMedium2.Row, 1) & "'"
        recTemp.Open strSql, ConnERP, 1, 3
        If Not recTemp.EOF Then
            If recTemp("spot_type") = "Reguler" Then
                For i = 0 To lstRDRateType.ListCount - 1
                    If Mid(lstRDRateType.List(i), 2, InStr(1, lstRDRateType.List(i), "]") - 2) = Trim(recTemp("rd_rate_type_code")) Then
                        lstRDRateType.Selected(i) = True
                        Exit For
                    End If
                Next
            End If
            CboRDSpotType2.Text = recTemp("spot_type")
            cboRDVersion2.Text = recTemp("version")
            cboRDDuration2.Text = recTemp("duration")
            txtRDRPS2.Text = FormatNumber(recTemp("rate_per_spot"))
            txtRDRPSGross2.Text = FormatNumber(recTemp("gross_rate_per_spot"))
            Call CboRDSpotType2_Click
        End If
        recTemp.Close
        CboRDSpotType2.Enabled = False
        
    Else
        If strMode <> "" Then
            isRateSelected = False
            isRateAvailable = False
            If CboRDSpotType2.Text = "Reguler" Then
                For i = 0 To lstRDRateType.ListCount - 1
                    If lstRDRateType.Selected(i) Then
                        isRateSelected = True
                        rate_type_code = Mid(lstRDRateType.List(i), 2, InStr(1, lstRDRateType.List(i), "]") - 2)
                        rate_type_name = Right(lstRDRateType.List(i), Len(lstRDRateType.List(i)) - InStr(1, lstRDRateType.List(i), "]"))
                        Exit For
                    End If
                Next
                If isRateSelected Then
                    'Cek Rate Catalog ada nggak??
                    Station_Code = lvRDSelectedStation.ListItems(1).KEY
                    recTemp.Open "select * from radio_rate where station_code = '" & Station_Code & "' and prime_reg='" & rate_type_code & "'", ConnERP, 1, 3
                    If Not recTemp.EOF Then
                        isRateAvailable = True
                    End If
                    recTemp.Close
                    If isRateAvailable Then
                        strSql = "update mp_plan_dimension set rd_rate_type_code = '" & rate_type_code & "',rd_rate_type_name='" & rate_type_name & "'"
                    Else
                        MsgBox "Selected Rate type not found in radio rate catalog!", vbExclamation, strApplication_Name
                        Exit Sub
                    End If
                Else
                    MsgBox "Please Select Rate Type!", vbExclamation, strApplication_Name
                    Exit Sub
                End If
            Else
                strSql = "update mp_plan_dimension set rate_per_spot=" & RemoveNumberFormat(txtRDRPS2.Text) & ",gross_rate_per_spot=" & RemoveNumberFormat(txtRDRPSGross2.Text)
            End If
            strSql = strSql & ",version = '" & Clear_String(cboRDVersion2.Text) & "',duration=" & txtRDDuration2.Text & " where mp_plan_dim_id = '" & FGRDMedium2.TextMatrix(FGRDMedium2.Row, 1) & "'"
            ConnERP.Execute strSql
            Call ViewMedium("RD2", FGRDMedium2)
            MsgBox "Data Saved!", vbExclamation, strApplication_Name
        End If
        
        CmdEditRD2.Caption = "&Edit"
        CmdDeleteRD2.Caption = "&Delete"
        cmdAddRD2.Enabled = True
        FGRDMedium2.Enabled = True
        trvRDStationCatalog.Enabled = True
        lvRDSelectedStation.ListItems.Clear
        lvRDSelectedStation.Enabled = True
        CboRDSpotType2.Enabled = True
        lstRDRateType.Enabled = True
        Call DisableControlTabRadio(True)
        Call EnableObject(False)
        strMode = ""
    End If
    
End Sub

Private Sub CmdEditTV_Click()
'<CSCM>
'********************************************************************************
'Procedure Name     : CmdEditTV_Click
'Procedure Function : ---
'Input Parameter    : ---
'Output Parameter   : ---
'Date               : 4/4/2016
'LastUpdate/By      : Tedi / Kreatif
'Name Before        : -
'********************************************************************************
'</CSCM>
    Dim recTemp As New ADODB.Recordset
    Dim strSql As String
    
    If CmdEditTV.Caption = "&Edit" Then
        
        'SETTING TOMBOL
        CmdEditTV.Caption = "&Save"
        cmdDeleteTV.Caption = "&Cancel"
        cmdAddTV.Enabled = False
        tdg_TVMedium.Enabled = False
        
        cboTVStationCode.Enabled = False
        cboTVStationName.Enabled = False
        cboTVMarketName.Enabled = False
        CmdNewTVStation.Enabled = False
        'TAMPILKAN DATA YG AKAN DIEDIT
        strSql = "select station_code,station_name,isnull(market_code,0) market_code,isnull(market_name,'NATIONAL') market_name from mp_medium_detail where mp_medium_detail_id = '" & tdg_TVMedium.Columns(1) & "'"
        recTemp.Open strSql, ConnERP, 1, 3
        If Not recTemp.EOF Then
            cboTVStationCode.Text = Trim(recTemp("station_code"))
            cboTVStationName.Text = recTemp("station_name")
            cboTVMarketName.Text = recTemp("market_name")
            cboTVMarketCode.ListIndex = cboTVMarketName.ListIndex
        End If
        recTemp.Close
        strSql = "select spot_type,version,duration,rate_per_spot,gross_rate_per_spot from mp_plan_dimension where mp_plan_dim_id = '" & tdg_TVMedium.Columns(0) & "'"
        recTemp.Open strSql, ConnERP, 1, 3
        If Not recTemp.EOF Then
            cboTVSpotType.Text = recTemp("spot_type")
            Call cboTVSpotType_Click
            If recTemp("spot_type") = "Reguler" Then
                cboTVVersion.Text = recTemp("version")
            Else
                cboTVProgram.Text = recTemp("version")
            End If
            txtTVDuration.Text = recTemp("duration")
            txtTVRate.Text = FormatNumber(recTemp("rate_per_spot"), 2)
            txtTVRateGross.Text = FormatNumber(recTemp("gross_rate_per_spot"), 2)
        End If
        recTemp.Close
    Else
        'SIMPAN DATA
        Dim strVersion As String
        If cboTVVersion.Text = "" Then
            MsgBox "Please Input Version", vbExclamation, strApplication_Name
            Exit Sub
        End If
        If cboTVSpotType.Text = "Reguler" Then
            strVersion = cboTVVersion.Text
        Else
            strVersion = cboTVProgram.Text
        End If
        strSql = "update mp_plan_dimension set spot_type='" & cboTVSpotType.Text & "',version='" & Clear_String(strVersion) & "',duration=" & txtTVDuration.Text & ",rate_per_spot=" & RemoveNumberFormat(txtTVRate.Text) & ",gross_rate_per_spot=" & RemoveNumberFormat(txtTVRateGross.Text) & " where mp_plan_dim_id = '" & tdg_TVMedium.Columns(0) & "'"
        ConnERP.Execute strSql
        'REFRESH TAMPILAN MEDIUM
        Call ViewMediumTrueDB("TV", tdg_TVMedium)
        MsgBox "Data Saved!", vbExclamation, strApplication_Name
        'SETTING TOMBOL
        CmdEditTV.Caption = "&Edit"
        cmdDeleteTV.Caption = "&Delete"
        cmdAddTV.Enabled = True
        tdg_TVMedium.Enabled = True
        With tdg_TVMedium
            .Columns(0).Caption = "Dimension ID"
            .Columns(1).Caption = "Detail ID"
            .Columns(2).Caption = "Station"
            .Columns(3).Caption = "Spot Type"
            .Columns(4).Caption = "Version"
            .Columns(5).Caption = "Duration"
            .Columns(6).Caption = "Gross Rate"
            .Columns(7).Caption = "Nett Rate"
        End With
        cboTVStationCode.Enabled = True
        cboTVStationName.Enabled = True
        cboTVMarketName.Enabled = True
        CmdNewTVStation.Enabled = True
'        FrameTV.Enabled = False
        Call DisableColorTab0(True)
        Call EnableObject(False)
        strMode = ""
        
    End If
End Sub



Private Sub CmdNewTVStation_Click()
'<CSCM>
'********************************************************************************
'Procedure Name     : CmdNewTVStation_Click
'Procedure Function : ---
'Input Parameter    : ---
'Output Parameter   : ---
'Date               : 4/4/2016
'LastUpdate/By      : Tedi / Kreatif
'Name Before        : -
'********************************************************************************
'</CSCM>
    Frm_MPSelectTVStation.show 1
End Sub

Private Sub Form_Load()
'<CSCM>
'********************************************************************************
'Procedure Name     : Form_Load
'Procedure Function : ---
'Input Parameter    : ---
'Output Parameter   : ---
'Date               : 4/4/2016
'LastUpdate/By      : Tedi / Kreatif
'Name Before        : -
'********************************************************************************
'</CSCM>

    Call initform
    Call EnableObject(False)
    FrameTV.Enabled = True
    Call DisableColorTab0(True)
    Call DisableControlTabRadio(True)
    Call DisableControlTabOther(True)
    Call ShowDetailTV
    FGRDMedium.Height = SSTabMedium.Height - (FGRDMedium.Top + 200)
    FGRDMedium2.Height = SSTabMedium.Height - (FGRDMedium2.Top + 200)
    intTabEdit = 9
    If tdg_TVMedium.ApproxCount > 1 Then
        tdg_TVMedium.Row = 0
        Call ShowDetailTV
    End If
    optArea.Value = True
    
End Sub

Private Sub initform()
'<CSCM>
'********************************************************************************
'Procedure Name     : initform
'Procedure Function : ---
'Input Parameter    : ---
'Output Parameter   : ---
'Date               : 4/4/2016
'LastUpdate/By      : Tedi / Kreatif
'Name Before        : -
'********************************************************************************
'</CSCM>
'*****************************************************************************
' Nama Prosedur     :   initform
' Fungsi Prosedur   :   Inisialisasi Form
' Parameter  Input  :
' Parameter Output  :
' Tgl Pembuatan     :   11 Agustus 2004
' Last Update/By    :   12 Agustus 2004/Sistyo
'*****************************************************************************
    Dim strMPACtivityID, strTemp1, strTemp2, strSql As String
    Dim idx, counter As Integer
    Dim JumTab As Integer
    'hide all frame at initiate..
        ReDim blnSSTabMediumFirstClick(SSTabMedium.Tabs) As Boolean
        For counter = 0 To SSTabMedium.Tabs - 1
            SSTabMedium.TabVisible(counter) = False
            blnSSTabMediumFirstClick(counter) = True
        Next
        
    'Init Frame Header
        With frm_MPEdit
            txtMPNumber.Text = .cboMPNum.Text
            txtMPTask.Text = .tdg_Task.Columns(2) '   .FGTask.TextMatrix(.FGTask.Row, 2)
            txtMPActivity.Text = .tdg_Activity.Columns(3) & " (" & .tdg_Activity.Columns(4) & ")" ' .FGActivity.TextMatrix(.FGActivity.Row, 3) & " (" & .FGActivity.TextMatrix(.FGActivity.Row, 4) & ")"
            
            txtBrandVariant.Text = .tdg_Activity.Columns(5) '.FGActivity.TextMatrix(.FGActivity.Row, 5)
            strMPACtivityID = .tdg_Activity.Columns(1) '.FGActivity.TextMatrix(.FGActivity.Row, 1)
        End With
    
    'periksa medium yang dipakai
        recTemp.Open "select medium_name from mp_medium where mp_activity_id='" & strMPACtivityID & "' order by medium_name desc", ConnERP, 1, 3
        While Not recTemp.EOF
            Select Case recTemp(0)
                Case "TV"
                    SSTabMedium.TabVisible(0) = True
                    SSTabMedium.Tab = 0
                Case "Radio"
                    SSTabMedium.TabVisible(1) = True
                    SSTabMedium.Tab = 1
                Case "Print"
                    SSTabMedium.TabVisible(2) = True
                    SSTabMedium.Tab = 2
                Case "Cinema"
                    SSTabMedium.TabVisible(3) = True
                    SSTabMedium.Tab = 3
                Case "Other"
                    SSTabMedium.TabVisible(4) = True
                    SSTabMedium.Tab = 4
            End Select
            txtMedium.Text = txtMedium.Text & recTemp(0) & ", "
            recTemp.MoveNext
        Wend
        recTemp.Close
        txtMedium.Text = Mid(txtMedium.Text, 1, Len(txtMedium.Text) - 2)
        
    'reorder tab view
        JumTab = 0
        For counter = SSTabMedium.Tabs - 1 To 0 Step -1
            If SSTabMedium.TabVisible(counter) Then
                SSTabMedium.Tab = counter
                JumTab = JumTab + 1
            End If
        Next
        If JumTab <> 0 Then
            SSTabMedium.TabsPerRow = JumTab
            blnSSTabMediumFirstClick(SSTabMedium.Tab) = False
            Select Case SSTabMedium.Tab
                Case 0
                    Call LoadTVProperties
                Case 1
                    Call LoadRDProperties
                Case 2
                    Call LoadPRProperties
                Case 3
                    Call LoadCNProperties
                Case 4
                    Call LoadOTProperties
            End Select
        End If
End Sub
Private Sub LoadTVProperties()
'<CSCM>
'********************************************************************************
'Procedure Name     : LoadTVProperties
'Procedure Function : ---
'Input Parameter    : ---
'Output Parameter   : ---
'Date               : 4/4/2016
'LastUpdate/By      : Tedi / Kreatif
'Name Before        : -
'********************************************************************************
'</CSCM>
    'TV PROPERTIES
        Call LoadTVMaterial
        Call LoadTVStation
        Call LoadTVMarket
        'TV Spot Type
            
            cboTVSpotType.AddItem "Reguler"
            cboTVSpotType.AddItem "Sponsorship/Program"
        Call initTabTV
    'END OF TV PROPERTIES
End Sub
Private Sub LoadRDProperties()
'<CSCM>
'********************************************************************************
'Procedure Name     : LoadRDProperties
'Procedure Function : ---
'Input Parameter    : ---
'Output Parameter   : ---
'Date               : 4/4/2016
'LastUpdate/By      : Tedi / Kreatif
'Name Before        : -
'********************************************************************************
'</CSCM>
    'RADIO PROPERTIES
        Call LoadRadioArea
        Call LoadRadioMaterial
        'RD Spot Type
            CboRDSpotType.AddItem "Reguler"
            CboRDSpotType.AddItem "Sponsorship/Program"
        Call initTabRadio
    'END OF RADIO PROPERTIES
End Sub
Private Sub LoadPRProperties()
'<CSCM>
'********************************************************************************
'Procedure Name     : LoadPRProperties
'Procedure Function : ---
'Input Parameter    : ---
'Output Parameter   : ---
'Date               : 4/4/2016
'LastUpdate/By      : Tedi / Kreatif
'Name Before        : -
'********************************************************************************
'</CSCM>
    'PRINT PROPERTIES
        'print spot type
            cboPRSpotType.AddItem "Reguler"
            'cboPRSpotType.AddItem "Sponsorship/Program"
            cboPRSpotType.AddItem "Special Buys"
        Call LoadPRMaterial
        Call initTabPrint
    'END OF PRINT PROPERTIES
End Sub

Private Sub LoadPRMaterial()
'<CSCM>
'********************************************************************************
'Procedure Name     : LoadPRMaterial
'Procedure Function : ---
'Input Parameter    : ---
'Output Parameter   : ---
'Date               : 4/4/2016
'LastUpdate/By      : Tedi / Kreatif
'Name Before        : -
'********************************************************************************
'</CSCM>
    Dim idx As Integer
    idx = 0
    cboPRVersion.Clear
    cboPRVersionCol.Clear
    cboPRVersionMM.Clear
    recTemp.Open "select isnull(material_name,''),isnull([column],0),isnull(size,0) from print_material where brand_code = '" & Mid(txtMPNumber.Text, 1, 4) & "' order by material_name", ConnERP, 1, 3
    While Not recTemp.EOF
        cboPRVersion.AddItem recTemp(0), idx
        cboPRVersionCol.AddItem recTemp(1), idx
        cboPRVersionMM.AddItem recTemp(2), idx
        idx = idx + 1
        recTemp.MoveNext
    Wend
    recTemp.Close
End Sub
    
Private Sub initTabPrint()
'<CSCM>
'********************************************************************************
'Procedure Name     : initTabPrint
'Procedure Function : ---
'Input Parameter    : ---
'Output Parameter   : ---
'Date               : 4/4/2016
'LastUpdate/By      : Tedi / Kreatif
'Name Before        : -
'********************************************************************************
'</CSCM>
    Dim intCol As Integer
    With FGPRRate
        .FixedCols = 0
        .cols = 11
        .Rows = 1
        .TextMatrix(0, 0) = "Media"
        .TextMatrix(0, 1) = "Code"
        .TextMatrix(0, 2) = "Size"
        .TextMatrix(0, 3) = "Paper"
        .TextMatrix(0, 4) = "Color"
        .TextMatrix(0, 5) = "Min Size"
        .TextMatrix(0, 6) = "Gross Rate"
        .TextMatrix(0, 7) = "Nett Rate"
        .TextMatrix(0, 8) = "Valid Date"
        .TextMatrix(0, 9) = "" 'ismmc
        .TextMatrix(0, 10) = "Notes"
        
        .ColWidth(0) = 1600
        .ColWidth(1) = 1000
        .ColWidth(2) = 1200
        .ColWidth(3) = 1000
        .ColWidth(4) = 700
        .ColWidth(5) = 900
        .ColWidth(6) = 1200
        .ColWidth(7) = 1200
        .ColWidth(8) = 1000
        .ColWidth(9) = 0
        .ColWidth(10) = 1700
        
        .Row = 0
        For intCol = 0 To .cols - 1
            .col = intCol
            .CellAlignment = 4
        Next
    End With
    With FGPRMedium
        .cols = 13
        .TextMatrix(0, 0) = "NO."
        .TextMatrix(0, 1) = "" 'MP_Plan_Dim_Id
        .TextMatrix(0, 2) = "" 'MP_Medium_Detail_Id
        .TextMatrix(0, 3) = "Media"
        .TextMatrix(0, 4) = "Spot Type"
        .TextMatrix(0, 5) = "Version"
        .TextMatrix(0, 6) = "Size"
        .TextMatrix(0, 7) = "Satuan"
        .TextMatrix(0, 8) = "Paper"
        .TextMatrix(0, 9) = "Color"
        .TextMatrix(0, 10) = "Min Size"
        .TextMatrix(0, 11) = "Gross rate"
        .TextMatrix(0, 12) = "Nett rate"
        .ColWidth(0) = 350
        .ColWidth(1) = 0
        .ColWidth(2) = 0
        .ColWidth(3) = 3000
        .ColWidth(4) = 1500
        .ColWidth(5) = 3000
        .ColWidth(6) = 700
        .ColWidth(7) = 1500
        .ColWidth(8) = 1500
        .ColWidth(9) = 1500
        .ColWidth(10) = 700
        .ColWidth(11) = 2030
        .ColWidth(12) = 2030
        .Row = 0
        For intCol = 1 To 12
            .col = intCol
            .CellAlignment = 3
        Next
    End With
    Call ViewMedium("PR", FGPRMedium)
    strPRNettRate = "0.00"
    strPRGrossRate = "0.00"
    FGPRRate.ZOrder 0
End Sub
Private Sub LoadCNProperties()
'<CSCM>
'********************************************************************************
'Procedure Name     : LoadCNProperties
'Procedure Function : ---
'Input Parameter    : ---
'Output Parameter   : ---
'Date               : 4/4/2016
'LastUpdate/By      : Tedi / Kreatif
'Name Before        : -
'********************************************************************************
'</CSCM>
    'CINEMA PROPERTIES
        Call LoadCinemaMaterial
        Call LoadCinemaCatalog
        'spottype
            cboCNSpotType.AddItem "Reguler"
            cboCNSpotType.AddItem "Sponsorship/Program"
        'Cinema Jenis Durasi
            recTemp.Open "select distinct jenis from cinema_rate", ConnERP, 1, 3
            cboCNJenisDurasi.Clear
            While Not recTemp.EOF
                cboCNJenisDurasi.AddItem recTemp(0)
                recTemp.MoveNext
            Wend
            recTemp.Close
        Call initTabCN
    'END OF CINEMA PROPERTIES
End Sub
Private Sub LoadOTProperties()
'<CSCM>
'********************************************************************************
'Procedure Name     : LoadOTProperties
'Procedure Function : ---
'Input Parameter    : ---
'Output Parameter   : ---
'Date               : 4/4/2016
'LastUpdate/By      : Tedi / Kreatif
'Name Before        : -
'********************************************************************************
'</CSCM>
    'OTHER PROPERTIES
        'Call LoadOtherSupplierCatalog
        'OT Spot Type
            'cboOTSpotType.AddItem "Reguler"
        Call initTabOther
    'END OF OTHER PROPERTIES
End Sub


Private Sub LoadCinemaCatalog()
'<CSCM>
'********************************************************************************
'Procedure Name     : LoadCinemaCatalog
'Procedure Function : ---
'Input Parameter    : ---
'Output Parameter   : ---
'Date               : 4/4/2016
'LastUpdate/By      : Tedi / Kreatif
'Name Before        : -
'********************************************************************************
'</CSCM>
    Dim idx As Integer
    idx = 0
    cboCNName.Clear
    cboCNCode.Clear
    recTemp.Open "select cinema_name,cinema_id from cinema_catalog order by cinema_name", ConnERP, 1, 3
    While Not recTemp.EOF
        cboCNName.AddItem recTemp(0), idx
        cboCNCode.AddItem recTemp(1), idx
        idx = idx + 1
        recTemp.MoveNext
    Wend
    recTemp.Close
End Sub

Private Sub LoadCinemaMaterial()
'<CSCM>
'********************************************************************************
'Procedure Name     : LoadCinemaMaterial
'Procedure Function : ---
'Input Parameter    : ---
'Output Parameter   : ---
'Date               : 4/4/2016
'LastUpdate/By      : Tedi / Kreatif
'Name Before        : -
'********************************************************************************
'</CSCM>
    Dim idx As Integer
    idx = 0
    cboCNVersion.Clear
    cboCNMaterialJenisDurasi.Clear
    cboCNMaterialDuration.Clear
    recTemp.Open "select material_name,duration,jenis_durasi from cinema_material where brand_code = '" & Mid(txtMPNumber.Text, 1, 4) & "' order by material_name", ConnERP, 1, 3
    While Not recTemp.EOF
        cboCNVersion.AddItem recTemp(0), idx
        cboCNMaterialDuration.AddItem recTemp(1), idx
        cboCNMaterialJenisDurasi.AddItem recTemp(2), idx
        idx = idx + 1
        recTemp.MoveNext
    Wend
    recTemp.Close
End Sub

Private Sub LoadTVStation()
'<CSCM>
'********************************************************************************
'Procedure Name     : LoadTVStation
'Procedure Function : ---
'Input Parameter    : ---
'Output Parameter   : ---
'Date               : 4/4/2016
'LastUpdate/By      : Tedi / Kreatif
'Name Before        : -
'********************************************************************************
'</CSCM>
    Dim idx As Integer
    cboTVStationCode.Clear
    cboTVStationName.Clear
    idx = 0
    recTemp.Open "select station_code,isnull(station_name,'') from tv_station_media_plan where brand_code = 'ALL' or brand_code='" & Left(txtMPNumber.Text, 4) & "' order by station_name", ConnERP, 1, 3
    While Not recTemp.EOF
        cboTVStationCode.AddItem recTemp(0), idx
        If IsNull(recTemp(1)) Then
            cboTVStationName.AddItem "None", idx
        Else
            cboTVStationName.AddItem recTemp(1), idx
        End If
        idx = idx + 1
        recTemp.MoveNext
    Wend
    recTemp.Close
End Sub

Private Sub LoadRadioArea()
'<CSCM>
'********************************************************************************
'Procedure Name     : LoadRadioArea
'Procedure Function : ---
'Input Parameter    : ---
'Output Parameter   : ---
'Date               : 4/4/2016
'LastUpdate/By      : Tedi / Kreatif
'Name Before        : -
'********************************************************************************
'</CSCM>
    Dim strSql As String
    Dim idx As Integer
    strSql = "select a.area_name,count(*) from radio_area_new a "
    strSql = strSql & "inner join radio_area_detail b "
    strSql = strSql & "on a.area_id = b.area_id "
    strSql = strSql & "and a.brand_code = '" & Left(txtMPNumber.Text, 4) & "' "
    strSql = strSql & "group by a.area_name "
    
    idx = 0
    cboRDArea.Clear
    cboRDStation.Clear
    recTemp.Open strSql, ConnERP, 1, 3
    While Not recTemp.EOF
        cboRDArea.AddItem recTemp(0), idx
        cboRDStation.AddItem recTemp(1), idx
        idx = idx + 1
        recTemp.MoveNext
    Wend
    recTemp.Close
End Sub

Private Sub LoadTVMaterial()
'<CSCM>
'********************************************************************************
'Procedure Name     : LoadTVMaterial
'Procedure Function : ---
'Input Parameter    : ---
'Output Parameter   : ---
'Date               : 4/4/2016
'LastUpdate/By      : Tedi / Kreatif
'Name Before        : -
'********************************************************************************
'</CSCM>
    Dim idx As Integer
    idx = 0
    cboTVVersion.Clear
    cboMaterialDuration.Clear
    recTemp.Open "select isnull(material_name,''),isnull(duration,0) from tv_material where brand_code = '" & Mid(txtMPNumber.Text, 1, 4) & "' order by material_name", ConnERP, 1, 3
    While Not recTemp.EOF
        cboTVVersion.AddItem recTemp(0), idx
        cboMaterialDuration.AddItem recTemp(1), idx
        recTemp.MoveNext
        idx = idx + 1
    Wend
    recTemp.Close
End Sub

Private Sub LoadRadioMaterial()
'<CSCM>
'********************************************************************************
'Procedure Name     : LoadRadioMaterial
'Procedure Function : ---
'Input Parameter    : ---
'Output Parameter   : ---
'Date               : 4/4/2016
'LastUpdate/By      : Tedi / Kreatif
'Name Before        : -
'********************************************************************************
'</CSCM>
    Dim idx As Integer
    idx = 0
    cboRDVersion.Clear
    recTemp.Open "select isnull(material_name,''),isnull(duration,0) from radio_material where brand_code = '" & Mid(txtMPNumber.Text, 1, 4) & "' order by material_name", ConnERP, 1, 3
    While Not recTemp.EOF
        cboRDVersion.AddItem recTemp(0), idx
        cboRDduration.AddItem recTemp(1), idx
        recTemp.MoveNext
        idx = idx + 1
    Wend
    recTemp.Close
End Sub

Private Sub LoadRadioMaterial2()
'<CSCM>
'********************************************************************************
'Procedure Name     : LoadRadioMaterial2
'Procedure Function : ---
'Input Parameter    : ---
'Output Parameter   : ---
'Date               : 4/4/2016
'LastUpdate/By      : Tedi / Kreatif
'Name Before        : -
'********************************************************************************
'</CSCM>
    Dim idx As Integer
    idx = 0
    cboRDVersion2.Clear
    recTemp.Open "select isnull(material_name,''),isnull(duration,0) from radio_material where brand_code = '" & Mid(txtMPNumber.Text, 1, 4) & "' order by material_name", ConnERP, 1, 3
    While Not recTemp.EOF
        cboRDVersion2.AddItem recTemp(0), idx
        cboRDDuration2.AddItem recTemp(1), idx
        recTemp.MoveNext
        idx = idx + 1
    Wend
    recTemp.Close
End Sub

Private Sub initTabCN()
'<CSCM>
'********************************************************************************
'Procedure Name     : initTabCN
'Procedure Function : ---
'Input Parameter    : ---
'Output Parameter   : ---
' Tgl Pembuatan     :   19 Agustus 2004
' Last Update/By    :   19 Agustus 2004/Sistyo
'Date               : 4/4/2016
'LastUpdate/By      : Tedi / Kreatif
'Name Before        : -
'********************************************************************************
'</CSCM>
    
    Dim counter As Integer
    frame_CN_brief.Visible = False
    frame_CN_brief.Height = 0
    blnOptCnBriefFirstClick = True
    Call ViewMediumTrueDB("CN", tdg_FGCNMedium)
    Call Give_Head_Name("CN")

End Sub

Private Sub Give_Head_Name(ByVal strNameProcess As String)
'<CSCM>
'********************************************************************************
'Procedure Name     : Give_Head_Name
'Procedure Function : ---
'Input Parameter    : ---
'Output Parameter   : ---
'Date               : 4/4/2016
'LastUpdate/By      : Tedi / Kreatif
'Name Before        : -
'********************************************************************************
'</CSCM>
    Select Case strNameProcess
        Case "CN"
            With tdg_FGCNMedium
            '.Columns(0).Caption = "NO."
            .Columns(0).Caption = "Dimension ID"
            .Columns(1).Caption = "Detail ID"
            .Columns(2).Caption = "Cinema"
            .Columns(3).Caption = "Spot Type"
            .Columns(4).Caption = "Version"
            .Columns(5).Caption = "Duration"
            .Columns(6).Caption = "Gross rate"
            .Columns(7).Caption = "Nett Rate"
            End With
        Case "CN2"
            With tdg_FGCNMedium2
              .Columns(0).Caption = "Dimension ID"
              .Columns(1).Caption = "Detail ID"
              .Columns(2).Caption = "Description"
              .Columns(3).Caption = "Gross Rate"
              .Columns(4).Caption = "Nett Rate"
            End With
    End Select
End Sub


Private Sub initTabTV()
'<CSCM>
'*****************************************************************************
' Nama Prosedur     :   initTabTV
' Fungsi Prosedur   :   Inisialisasi Tab TV
' Parameter  Input  :
' Parameter Output  :
' Tgl Pembuatan     :   16 Agustus 2004
' Last Update/By    :   16 Agustus 2004/Sistyo
'*****************************************************************************
'</CSCM>

    Dim counter As Integer

    Call ViewMediumTrueDB("TV", tdg_TVMedium)
        With tdg_TVMedium
        .Columns(0).Caption = "Dimension ID"
        .Columns(1).Caption = "Detail ID"
        .Columns(2).Caption = "Station"
        .Columns(3).Caption = "Spot Type"
        .Columns(4).Caption = "Version"
        .Columns(5).Caption = "Duration"
        .Columns(6).Caption = "Gross Rate"
        .Columns(7).Caption = "Nett Rate"
    End With

End Sub

Private Sub initTabRadio()
'<CSCM>
'********************************************************************************
'Procedure Name     : initTabRadio
'Procedure Function : ---
'Input Parameter    : ---
'Output Parameter   : ---
'Date               : 4/4/2016
'LastUpdate/By      : Tedi / Kreatif
'Name Before        : -
'********************************************************************************
'</CSCM>
'*****************************************************************************
' Nama Prosedur     :   initTabRadio
' Fungsi Prosedur   :   Inisialisasi Tab Radio
' Parameter  Input  :
' Parameter Output  :
' Tgl Pembuatan     :   16 Agustus 2004
' Last Update/By    :   16 Agustus 2004/Sistyo
'*****************************************************************************
    Dim counter As Integer
    'Default Selection By Area
    fra_RDByStation.Visible = False
    fra_RDByStation.Height = 0
    blnOptRDStationFirstClick = True
    
    With FGRDMedium
        .cols = 9
        .TextMatrix(0, 0) = "NO."
        .TextMatrix(0, 1) = "" 'MP_Plan_Dim_Id
        .TextMatrix(0, 2) = "" 'MP_Medium_Detail_Id
        .TextMatrix(0, 3) = "Area"
        .TextMatrix(0, 4) = "Spot Type"
        .TextMatrix(0, 5) = "Version"
        .TextMatrix(0, 6) = "Station"
        .TextMatrix(0, 7) = "Gross rate"
        .TextMatrix(0, 8) = "Nett rate"
        .ColWidth(0) = 350
        .ColWidth(1) = 0
        .ColWidth(2) = 0
        .ColWidth(3) = 2000
        .ColWidth(4) = 1500
        .ColWidth(5) = 3000
        .ColWidth(6) = 2030
        .ColWidth(7) = 2030
        .ColWidth(8) = 2030
        .Row = 0
        For counter = 1 To 8
            .col = counter
            .CellAlignment = 3
        Next
    End With

    Call ViewMedium("RD", FGRDMedium)

End Sub

Private Sub initTabOther()
'<CSCM>
'********************************************************************************
'Procedure Name     : initTabOther
'Procedure Function : ---
'Input Parameter    : ---
'Output Parameter   : ---
'Date               : 4/4/2016
'LastUpdate/By      : Tedi / Kreatif
'Name Before        : -
'********************************************************************************
'</CSCM>
'*****************************************************************************
' Nama Prosedur     :   initTabOther
' Fungsi Prosedur   :   Inisialisasi Tab Other
' Parameter  Input  :
' Parameter Output  :
' Tgl Pembuatan     :   16 Agustus 2004
' Last Update/By    :   16 Agustus 2004/Sistyo
'*****************************************************************************
  
    Call ViewMediumTrueDB("OT", tdg_OTMedium)

End Sub

Private Sub ViewMedium(medium_code As String, FlexGridName As MSFlexGrid)
'<CSCM>
'********************************************************************************
'Procedure Name     : ViewMedium
'Procedure Function : ---
'Input Parameter    : ---
'Output Parameter   : ---
'Date               : 4/4/2016
'LastUpdate/By      : Tedi / Kreatif
'Name Before        : -
'********************************************************************************
'</CSCM>
'*****************************************************************************
' Nama Prosedur     :   ViewMedium
' Fungsi Prosedur   :   Menampilkan medium detail dan plan dimension
' Parameter  Input  :
' Parameter Output  :
' Tgl Pembuatan     :   13 Agustus 2004
' Last Update/By    :   22 Juni 2005/Sistyo
'*****************************************************************************

    Dim strSql, strMP_Medium_Id As String
    Dim counter, intCol As Integer
    Dim strMP_Medium_code As String
    strMP_Medium_code = medium_code
    If strMP_Medium_code = "RD2" Then strMP_Medium_code = "RD"
    If strMP_Medium_code = "CN2" Then strMP_Medium_code = "CN"
    strMP_Medium_Id = ""
    
    'get MP_Medium_ID
        'strSql = "select mp_medium_id from mp_medium where mp_activity_id='" & Frm_MPEdit.FGActivity.TextMatrix(Frm_MPEdit.FGActivity.Row, 1) & "' and medium_code='" & strMP_Medium_code & "'"
        strSql = "select mp_medium_id from mp_medium where mp_activity_id='" & frm_MPEdit.tdg_Activity.Columns(1) & "' and medium_code='" & strMP_Medium_code & "'"
        
        recTemp.Open strSql, ConnERP, 1, 3
        If Not recTemp.EOF Then
            strMP_Medium_Id = recTemp(0)
        End If
        recTemp.Close
        
    'build SQL Query
        Select Case medium_code
            Case "TV"
                
                strSql = "select b.mp_plan_dim_id,b.mp_medium_detail_id,a.station_name,b.spot_type,isnull(b.version,''),isnull(b.duration,0),b.gross_rate_per_spot,b.rate_per_spot"
                strSql = strSql & " from mp_medium_detail a"
                strSql = strSql & " inner join mp_plan_dimension b"
                strSql = strSql & " on a.mp_medium_detail_id=b.mp_medium_detail_id"
                strSql = strSql & " where a.mp_medium_id='" & strMP_Medium_Id & "'"
            
            Case "RD"
            
                strSql = "select b.mp_plan_dim_id,b.mp_medium_detail_id,' ' + a.area_code,b.spot_type,isnull(b.version,''),' ' + cast(b.rd_stations as varchar) + ' radio stations',b.gross_rate_per_spot,b.rate_per_spot"
                strSql = strSql & " from mp_medium_detail a"
                strSql = strSql & " inner join mp_plan_dimension b"
                strSql = strSql & " on a.mp_medium_detail_id=b.mp_medium_detail_id"
                strSql = strSql & " where a.mp_medium_id='" & strMP_Medium_Id & "' and isnull(b.isRDByStation,0)<>1"
            
            Case "RD2"
            
                strSql = "select b.mp_plan_dim_id,b.mp_medium_detail_id,' ' + a.radio_station_name,b.spot_type,isnull(b.version,''),duration,isnull(rd_rate_type_name,''),b.gross_rate_per_spot,b.rate_per_spot "
                strSql = strSql & " from mp_medium_detail a"
                strSql = strSql & " inner join mp_plan_dimension b"
                strSql = strSql & " on a.mp_medium_detail_id=b.mp_medium_detail_id"
                strSql = strSql & " where a.mp_medium_id='" & strMP_Medium_Id & "' and b.isRDByStation=1"
               
            Case "PR"
                
                strSql = "select b.mp_plan_dim_id,a.mp_medium_detail_id,a.media_name,b.spot_type,b.version,"
                strSql = strSql & " Case b.print_ismmc"
                strSql = strSql & " when 0 then '1'"
                strSql = strSql & " when 1 then cast(b.print_mmc_col as varchar) + ' x ' + cast(b.print_mmc_size as varchar)"
                strSql = strSql & " end [size],b.print_size_name satuan,b.print_paper_name paper, b.print_color_name color, b.print_min_size,b.gross_rate_per_spot,b.rate_per_spot "
                strSql = strSql & " from mp_medium_detail a"
                strSql = strSql & " inner join mp_plan_dimension b on a.mp_medium_detail_id = b.mp_medium_detail_id"
                strSql = strSql & " where a.mp_medium_id='" & strMP_Medium_Id & "'"
                
            Case "CN"
            
                strSql = "select b.mp_plan_dim_id,b.mp_medium_detail_id,a.cinema_name + ', Studio ' + cinema_studio cinema,b.spot_type,isnull(b.version,''),b.Cinema_duration + ' ' + cast(duration as varchar) + ' sec' durasi,b.gross_rate_per_spot,b.rate_per_spot"
                strSql = strSql & " from mp_medium_detail a"
                strSql = strSql & " inner join mp_plan_dimension b"
                strSql = strSql & " on a.mp_medium_detail_id=b.mp_medium_detail_id"
                strSql = strSql & " where a.mp_medium_id='" & strMP_Medium_Id & "' and a.cinema_name <>''"
                
            Case "CN2"
            
                strSql = "select b.mp_plan_dim_id,b.mp_medium_detail_id,b.ot_description,b.gross_rate_per_spot,b.rate_per_spot"
                strSql = strSql & " from mp_medium_detail a"
                strSql = strSql & " inner join mp_plan_dimension b"
                strSql = strSql & " on a.mp_medium_detail_id=b.mp_medium_detail_id"
                strSql = strSql & " where a.mp_medium_id='" & strMP_Medium_Id & "' and a.cinema_name=''"
                
            Case "OT"
                strSql = "select b.mp_plan_dim_id,b.mp_medium_detail_id,isnull(b.OT_Description,'') "
                strSql = strSql & " from mp_medium_detail a"
                strSql = strSql & " inner join mp_plan_dimension b"
                strSql = strSql & " on a.mp_medium_detail_id=b.mp_medium_detail_id"
                strSql = strSql & " where a.mp_medium_id='" & strMP_Medium_Id & "'"
                
        End Select
    
    'view Medium
        recTemp.Open strSql, ConnERP, 1, 3
        With FlexGridName
            counter = 1
            .Rows = counter + 1
            .ColWidth(1) = 1000
            .ColWidth(2) = 1000
            For intCol = 0 To .cols - 1
                .TextMatrix(counter, intCol) = ""
                
            Next
            While Not recTemp.EOF
                .Rows = counter + 1
                .TextMatrix(counter, 0) = counter
                For intCol = 1 To recTemp.Fields.Count
                    .TextMatrix(counter, intCol) = Trim(recTemp(intCol - 1))
                Next
                counter = counter + 1
                recTemp.MoveNext
            Wend
        End With
        recTemp.Close
    
End Sub

Private Sub ViewMediumTrueDB(medium_code As String, TDGridTemp)
'<CSCM>
'********************************************************************************
'Procedure Name     : ViewMedium
'Procedure Function : ---
'Input Parameter    : ---
'Output Parameter   : ---
'Date               : 4/4/2016
'LastUpdate/By      : Tedi / Kreatif
'Name Before        : -
'********************************************************************************
'</CSCM>
'*****************************************************************************
' Nama Prosedur     :   ViewMedium
' Fungsi Prosedur   :   Menampilkan medium detail dan plan dimension
' Parameter  Input  :
' Parameter Output  :
' Tgl Pembuatan     :   13 Agustus 2004
' Last Update/By    :   22 Juni 2005/Sistyo
'*****************************************************************************

    Dim strSql, strMP_Medium_Id As String
    Dim counter, intCol As Integer
    Dim strMP_Medium_code As String
    strMP_Medium_code = medium_code
    If strMP_Medium_code = "RD2" Then strMP_Medium_code = "RD"
    If strMP_Medium_code = "CN2" Then strMP_Medium_code = "CN"
    strMP_Medium_Id = ""
    
    'get MP_Medium_ID
        'strSql = "select mp_medium_id from mp_medium where mp_activity_id='" & Frm_MPEdit.FGActivity.TextMatrix(Frm_MPEdit.FGActivity.Row, 1) & "' and medium_code='" & strMP_Medium_code & "'"
        strSql = "SELECT mp_medium_id "
        strSql = strSql & "FROM mp_medium "
        strSql = strSql & "WHERE mp_activity_id='" & frm_MPEdit.tdg_Activity.Columns(1) & "' "
        strSql = strSql & "AND medium_code='" & strMP_Medium_code & "'"
        
        recTemp.Open strSql, ConnERP, 1, 3
        If Not recTemp.EOF Then
            strMP_Medium_Id = recTemp(0)
        End If
        recTemp.Close
        
    'build SQL Query
        Select Case medium_code
            Case "TV"
                'row_number() over (order by a.mp_medium_id) as rank,
                strSql = "select  b.mp_plan_dim_id,b.mp_medium_detail_id,a.station_name,b.spot_type,isnull(b.version,''),isnull(b.duration,0),b.gross_rate_per_spot,b.rate_per_spot"
                strSql = strSql & " from mp_medium_detail a"
                strSql = strSql & " inner join mp_plan_dimension b"
                strSql = strSql & " on a.mp_medium_detail_id=b.mp_medium_detail_id"
                strSql = strSql & " where a.mp_medium_id='" & strMP_Medium_Id & "' ORDER BY a.mp_medium_id"
                If recTV.State = 1 Then Call CloseRecordset(recTV)
                recTV.Open strSql, ConnERP, adOpenStatic, adLockPessimistic
                TDGridTemp.ClearFields
                Set TDGridTemp.DataSource = recTV
                TDGridTemp.Refresh
            Case "RD"
            
                strSql = "select b.mp_plan_dim_id,b.mp_medium_detail_id,' ' + a.area_code,b.spot_type,isnull(b.version,''),' ' + cast(b.rd_stations as varchar) + ' radio stations',b.gross_rate_per_spot,b.rate_per_spot"
                strSql = strSql & " from mp_medium_detail a"
                strSql = strSql & " inner join mp_plan_dimension b"
                strSql = strSql & " on a.mp_medium_detail_id=b.mp_medium_detail_id"
                strSql = strSql & " where a.mp_medium_id='" & strMP_Medium_Id & "' and isnull(b.isRDByStation,0)<>1"
            
            Case "RD2"
            
                strSql = "select b.mp_plan_dim_id,b.mp_medium_detail_id,' ' + a.radio_station_name,b.spot_type,isnull(b.version,''),duration,isnull(rd_rate_type_name,''),b.gross_rate_per_spot,b.rate_per_spot "
                strSql = strSql & " from mp_medium_detail a"
                strSql = strSql & " inner join mp_plan_dimension b"
                strSql = strSql & " on a.mp_medium_detail_id=b.mp_medium_detail_id"
                strSql = strSql & " where a.mp_medium_id='" & strMP_Medium_Id & "' and b.isRDByStation=1"
               
            Case "PR"
                
                strSql = "select b.mp_plan_dim_id,a.mp_medium_detail_id,a.media_name,b.spot_type,b.version,"
                strSql = strSql & " Case b.print_ismmc"
                strSql = strSql & " when 0 then '1'"
                strSql = strSql & " when 1 then cast(b.print_mmc_col as varchar) + ' x ' + cast(b.print_mmc_size as varchar)"
                strSql = strSql & " end [size],b.print_size_name satuan,b.print_paper_name paper, b.print_color_name color, b.print_min_size,b.gross_rate_per_spot,b.rate_per_spot "
                strSql = strSql & " from mp_medium_detail a"
                strSql = strSql & " inner join mp_plan_dimension b on a.mp_medium_detail_id = b.mp_medium_detail_id"
                strSql = strSql & " where a.mp_medium_id='" & strMP_Medium_Id & "'"
                
            Case "CN"
            
                strSql = "select b.mp_plan_dim_id,b.mp_medium_detail_id,"
                strSql = strSql & "a.cinema_name + ', Studio ' + cinema_studio cinema,"
                strSql = strSql & "b.spot_type,isnull(b.version,''),"
                strSql = strSql & "b.Cinema_duration + ' ' + cast(duration as varchar) + ' sec' durasi,"
                strSql = strSql & "b.gross_rate_per_spot,b.rate_per_spot"
                strSql = strSql & " from mp_medium_detail a"
                strSql = strSql & " inner join mp_plan_dimension b"
                strSql = strSql & " on a.mp_medium_detail_id=b.mp_medium_detail_id"
                strSql = strSql & " where a.mp_medium_id='" & strMP_Medium_Id & "' and a.cinema_name <>''"
                If recCN.State = 1 Then Call CloseRecordset(recCN)
                recCN.Open strSql, ConnERP, adOpenStatic, adLockPessimistic
                TDGridTemp.ClearFields
                Set TDGridTemp.DataSource = recCN
                TDGridTemp.Refresh
                Call Give_Head_Name("CN")
                
            Case "CN2"
            
                strSql = "select b.mp_plan_dim_id,b.mp_medium_detail_id,b.ot_description,b.gross_rate_per_spot,b.rate_per_spot"
                strSql = strSql & " from mp_medium_detail a"
                strSql = strSql & " inner join mp_plan_dimension b"
                strSql = strSql & " on a.mp_medium_detail_id=b.mp_medium_detail_id"
                strSql = strSql & " where a.mp_medium_id='" & strMP_Medium_Id & "' and a.cinema_name=''"
                If recCN2.State = 1 Then Call CloseRecordset(recCN2)
                recCN2.Open strSql, ConnERP, adOpenStatic, adLockPessimistic
                TDGridTemp.ClearFields
                Set TDGridTemp.DataSource = recCN2
                TDGridTemp.Refresh
                Call Give_Head_Name("CN2")
            Case "OT"
                strSql = "Provider=SQLOLEDB.1;"
                strSql = strSql & "Persist Security Info=False;"
                strSql = strSql & "User ID=" & strDBLogin_User & ";"
                strSql = strSql & "password=" & strDBLogin_Password & ";"
                strSql = strSql & "Initial Catalog=" & strDatabase_Name & ";"
                strSql = strSql & "Data Source=" & strServerName
                adoOT.ConnectionString = strSql
                strSql = "select b.mp_plan_dim_id,b.mp_medium_detail_id,isnull(b.OT_Description,'') as Description, '1' as [command] "
                strSql = strSql & " from mp_medium_detail a"
                strSql = strSql & " inner join mp_plan_dimension b"
                strSql = strSql & " on a.mp_medium_detail_id=b.mp_medium_detail_id"
                strSql = strSql & " where a.mp_medium_id='" & strMP_Medium_Id & "'"
                If recOT.State = 1 Then Call CloseRecordset(recOT)
                recOT.Open strSql, ConnERP, adOpenStatic, adLockPessimistic
                adoOT.RecordSource = strSql
                adoOT.Refresh
                TDGridTemp.Refresh
                TDGridTemp.Columns(0).Visible = False
                TDGridTemp.Columns(1).Visible = False
                TDGridTemp.Columns(1).Width = 1000
                If tdg_OTMedium.Columns.Count > 1 Then
                    tdg_OTMedium.Columns(0).Visible = False
                    tdg_OTMedium.Columns(1).Visible = False
                    tdg_OTMedium.Columns(2).Width = 6800
                    tdg_OTMedium.Columns(3).Width = 1750
                    tdg_OTMedium.Columns(3).ButtonAlways = True
                End If
        End Select
End Sub

Private Sub AddPR()
'<CSCM>
'********************************************************************************
'Procedure Name     : AddPR
'Procedure Function : ---
'Input Parameter    : ---
'Output Parameter   : ---
'Date               : 4/4/2016
'LastUpdate/By      : Tedi / Kreatif
'Name Before        : -
'********************************************************************************
'</CSCM>
'*****************************************************************************
' Nama Prosedur     :   AddPR
' Fungsi Prosedur   :   Add Medium Detail and plan Dimension for Print
' Parameter  Input  :
' Parameter Output  :
' Tgl Pembuatan     :   5 Nov 2004
' Last Update/By    :   5 Nov 2004/Sistyo
'*****************************************************************************
    Dim pesan, strMP_Medium_Id As String, strMP_Medium_Detail_ID As String, strSql As String, strMP_Plan_Dim_Id As String
    Dim GrossRate As Double, NettRate As Double
    'field validation
        If TxtPRMediaCode.Text = "" Then
            pesan = MsgBox("Please select media name!", vbCritical + vbOKOnly, strApplication_Name)
            TxtPRMediaCode.SetFocus
            Exit Sub
        End If
        If cboPRVersion.Text = "" Then
            pesan = MsgBox("Please select Material!", vbCritical + vbOKOnly, strApplication_Name)
            cboPRVersion.SetFocus
            Exit Sub
        End If
        If cboPRSpotType.Text = "" Then
            pesan = MsgBox("Please select spot type!", vbCritical + vbOKOnly, strApplication_Name)
            cboPRSpotType.SetFocus
            Exit Sub
        End If
    'end of field validation
    
    strMP_Medium_Id = ""
    strMP_Medium_Detail_ID = ""
    
    'get MP_Medium_ID
        strSql = "select mp_medium_id from mp_medium where mp_activity_id='" & frm_MPEdit.tdg_Activity.Columns(1) & "' and medium_code='PR'"
        recTemp.Open strSql, ConnERP, 1, 3
        If Not recTemp.EOF Then
            strMP_Medium_Id = recTemp(0)
        End If
        recTemp.Close
        
    'get mp_medium_detail_id
        strSql = "select mp_medium_detail_id from mp_medium_detail where mp_medium_id = '" & strMP_Medium_Id & "' and print_code='" & TxtPRMediaCode.Text & "'"
        recTemp.Open strSql, ConnERP, 1, 3
        If Not recTemp.EOF Then
            strMP_Medium_Detail_ID = recTemp(0)
        End If
        recTemp.Close
   
    'create mp_medium_detail_id when mp_medium_detail_id is not defined
        If strMP_Medium_Detail_ID = "" Then
            strMP_Medium_Detail_ID = NextMPMediumDetailID(strMP_Medium_Id)
            'insert mp_medium_detail
                strSql = "insert into mp_medium_detail(mp_medium_detail_id,mp_medium_id,medium_code,print_code,media_name) values ('"
                strSql = strSql & strMP_Medium_Detail_ID & "','" & strMP_Medium_Id & "','PR','"
                strSql = strSql & TxtPRMediaCode.Text & "','" & Clear_String(txtPRMediaName.Text) & "')"
                ConnERP.Execute strSql
        End If
    
    'create MP_Plan_Dim_Id
        
        strMP_Plan_Dim_Id = NextMPPlanDimID(strMP_Medium_Detail_ID)
        
    'insert mp_plan_dimension
        GrossRate = 0: NettRate = 0
        If cboPRSpotType.Text <> "Reguler" Then
            GrossRate = RemoveNumberFormat(txtPRRateGross.Text)
            NettRate = RemoveNumberFormat(txtPRRate.Text)
        End If
        If txtPRIsMMC = 1 Then
            strSql = "insert into mp_plan_dimension(mp_plan_dim_id,mp_medium_detail_id,medium_code,"
            strSql = strSql & "spot_type,version,print_size_code,print_paper_code,print_color_code,print_ismmc,print_mmc_col,print_mmc_size,print_min_size,gross_rate,nett_rate,total_q1,total_q2,total_q3,total_q4,rate_per_spot,gross_rate_per_spot) "
            strSql = strSql & " values ('" & strMP_Plan_Dim_Id & "','" & strMP_Medium_Detail_ID & "','PR','"
            strSql = strSql & cboPRSpotType.Text & "','" & Clear_String(cboPRVersion.Text) & "','" & TxtPRSatuan.Text & "','" & TxtPRPaper.Text & "','" & TxtPRColor.Text & "',1," & txtPRCol.Text & "," & txtPRMM.Text & "," & TxtPRMinSize.Text & ",0,0,0,0,0,0," & NettRate & "," & GrossRate & ")"
        Else
            strSql = "insert into mp_plan_dimension(mp_plan_dim_id,mp_medium_detail_id,medium_code,"
            strSql = strSql & "spot_type,version,print_size_code,print_paper_code,print_color_code,print_ismmc,print_min_size,gross_rate,nett_rate,total_q1,total_q2,total_q3,total_q4,rate_per_spot,gross_rate_per_spot) "
            strSql = strSql & " values ('" & strMP_Plan_Dim_Id & "','" & strMP_Medium_Detail_ID & "','PR','"
            strSql = strSql & cboPRSpotType.Text & "','" & Clear_String(cboPRVersion.Text) & "','" & TxtPRSatuan.Text & "','" & TxtPRPaper.Text & "','" & TxtPRColor.Text & "',0," & TxtPRMinSize.Text & ",0,0,0,0,0,0," & NettRate & "," & GrossRate & ")"
        End If
        ConnERP.Execute strSql
        
    'update mp_master
        ConnERP.Execute "update mp_master set last_update_by='" & strLogin_User & "',last_update_date=getdate() where mp_number='" & txtMPNumber.Text & "'"
        
    pesan = MsgBox("New Medium Added!", vbExclamation, strApplication_Name)
    Call DisableControlTabPrint(True)
    Call EmptyTabPrint
    If FGPRMedium.Rows > 0 Then
        If FGPRMedium.Row = 0 Then FGPRMedium.Row = 1
        Call FGPRMedium_Click
    End If
    Call EnableObject(False)
    strMode = ""
    
End Sub

Private Sub AddTV()
'<CSCM>
'********************************************************************************
'Procedure Name     : AddTV
'Procedure Function : ---
'Input Parameter    : ---
'Output Parameter   : ---
'Date               : 4/4/2016
'LastUpdate/By      : Tedi / Kreatif
'Name Before        : -
'********************************************************************************
'</CSCM>
'*****************************************************************************
' Nama Prosedur     :   AddTV
' Fungsi Prosedur   :   Add Medium Detail and plan Dimension for TV
' Parameter  Input  :
' Parameter Output  :
' Tgl Pembuatan     :   12 Agustus 2004
' Last Update/By    :   12 Agustus 2004/Sistyo
'*****************************************************************************
    Dim strSql As String, strVersion As String, strMP_Medium_Id As String, strMP_Medium_Detail_ID As String, strMP_Plan_Dim_Id As String, flag As String
    Dim counter As Integer
    Dim pesan
    'field validation
        If cboTVStationCode.Text = "" Then
            pesan = MsgBox("Not enough medium information!", vbExclamation, strApplication_Name)
            Exit Sub
        End If
        
        If cboTVSpotType.Text = "" Then
            pesan = MsgBox("Not enough medium information!", vbExclamation, strApplication_Name)
            Exit Sub
        End If
        
        If txtTVDuration.Text = "" Then
            txtTVDuration.Text = 0
        End If
            
        strVersion = cboTVProgram.Text
        If cboTVSpotType.Text = "Reguler" Then
            strVersion = cboTVVersion.Text
        End If
        
        If cboTVMarketName.Text = "" Then
            MsgBox "Please Select TV Market!", vbCritical + vbOKOnly, strApplication_Name
            Exit Sub
        End If
    'end of field validation
    
    strMP_Medium_Id = ""
    strMP_Medium_Detail_ID = ""
    
    'get MP_Medium_ID
        strSql = "select mp_medium_id from mp_medium where mp_activity_id='" & frm_MPEdit.tdg_Activity.Columns(1) & "' and medium_code='TV'"
        recTemp.Open strSql, ConnERP, 1, 3
        If Not recTemp.EOF Then
            strMP_Medium_Id = recTemp(0)
        End If
        recTemp.Close
        
        
    'get mp_medium_detail_id
        strSql = "select mp_medium_detail_id from mp_medium_detail where mp_medium_id = '" & strMP_Medium_Id & "' and station_code='" & cboTVStationCode.Text & "' and market_code=" & cboTVMarketCode.Text
        recTemp.Open strSql, ConnERP, 1, 3
        If Not recTemp.EOF Then
            strMP_Medium_Detail_ID = recTemp(0)
        End If
        recTemp.Close
    
    'create mp_medium_detail_id when mp_medium_detail_id is not defined
        If strMP_Medium_Detail_ID = "" Then
            
            strMP_Medium_Detail_ID = NextMPMediumDetailID(strMP_Medium_Id)
            
            'insert mp_medium_detail
                strSql = "insert into mp_medium_detail(mp_medium_detail_id,mp_medium_id,medium_code,station_code,station_name,market_code,market_name) values ('"
                strSql = strSql & strMP_Medium_Detail_ID & "','" & strMP_Medium_Id & "','TV','"
                strSql = strSql & cboTVStationCode.Text & "','" & Clear_String(cboTVStationName.Text) & "'," & cboTVMarketCode.Text & ",'" & cboTVMarketName.Text & "')"
                ConnERP.Execute strSql
        End If
    
    'create MP_Plan_Dim_Id
        strMP_Plan_Dim_Id = NextMPPlanDimID(strMP_Medium_Detail_ID)
        
    'insert mp_plan_dimension
        strSql = "insert into mp_plan_dimension(mp_plan_dim_id,mp_medium_detail_id,medium_code,"
        strSql = strSql & "spot_type,version,duration,gross_rate,nett_rate,total_q1,total_q2,total_q3,total_q4,rate_per_spot,gross_rate_per_spot) "
        strSql = strSql & " values ('" & strMP_Plan_Dim_Id & "','" & strMP_Medium_Detail_ID & "','TV','"
        strSql = strSql & cboTVSpotType.Text & "','" & Clear_String(CStr(strVersion)) & "'," & Val(txtTVDuration.Text) & ",0,0,0,0,0,0," & RemoveNumberFormat(txtTVRate.Text) & "," & RemoveNumberFormat(txtTVRateGross.Text) & ")"
        ConnERP.Execute strSql
        
    'update mp_master
        ConnERP.Execute "update mp_master set last_update_by='" & strLogin_User & "',last_update_date=getdate() where mp_number='" & txtMPNumber.Text & "'"
        
        pesan = MsgBox("New Medium Added!", vbExclamation, strApplication_Name)
'        FrameTV.Enabled = False
        Call DisableColorTab0(True)
        Call EnableObject(False)
        strMode = ""
End Sub

Private Sub AddRadio()
'<CSCM>
'*****************************************************************************
' Nama Prosedur     :   AddRAdio
' Fungsi Prosedur   :   Add Medium Detail and plan Dimension for Radio
' Parameter  Input  :
' Parameter Output  :
' Tgl Pembuatan     :   16 Agustus 2004
' Update/By         :   16 Agustus 2004/Sistyo
'Date               : 4/4/2016
'LastUpdate/By      : Tedi / Kreatif
'Name Before        : -
'********************************************************************************
'</CSCM>

    Dim strSql As String, strMP_Medium_Id As String, strMP_Medium_Detail_ID As String, strMP_Plan_Dim_Id As String, flag As String
    Dim counter As Integer
    Dim pesan
    
    'field validation
        If cboRDArea.Text = "" Then
            pesan = MsgBox("Not enough medium information!", vbExclamation, strApplication_Name)
            cboRDArea.SetFocus
            Exit Sub
        End If
        
        If CboRDSpotType.Text = "" Then
            pesan = MsgBox("Not enough medium information!", vbExclamation, strApplication_Name)
            CboRDSpotType.SetFocus
            Exit Sub
        End If
        
        If cboRDVersion.Text = "" Then
            pesan = MsgBox("Not enough medium information!", vbExclamation, strApplication_Name)
            cboRDVersion.SetFocus
            Exit Sub
        End If
    
    'end of field validation
    
    strMP_Medium_Id = ""
    strMP_Medium_Detail_ID = ""
    
    'get MP_Medium_ID
        strSql = "select mp_medium_id from mp_medium where mp_activity_id='" & frm_MPEdit.tdg_Activity.Columns(1) & "' and medium_code='RD'"
        recTemp.Open strSql, ConnERP, 1, 3
        If Not recTemp.EOF Then
            strMP_Medium_Id = recTemp(0)
        End If
        recTemp.Close
    
    'get mp_medium_detail_id
        strSql = "select mp_medium_detail_id from mp_medium_detail where mp_medium_id = '" & strMP_Medium_Id & "' and area_code='" & cboRDArea.Text & "'"
        recTemp.Open strSql, ConnERP, 1, 3
        If Not recTemp.EOF Then
            strMP_Medium_Detail_ID = recTemp(0)
        End If
        recTemp.Close
    
    'create mp_medium_detail_id when mp_medium_detail_id is not defined
        If strMP_Medium_Detail_ID = "" Then
            
            strMP_Medium_Detail_ID = NextMPMediumDetailID(strMP_Medium_Id)
            
            'insert mp_medium_detail
                strSql = "insert into mp_medium_detail(mp_medium_detail_id,mp_medium_id,medium_code,area_code) values ('"
                strSql = strSql & strMP_Medium_Detail_ID & "','" & strMP_Medium_Id & "','RD','"
                strSql = strSql & Clear_String(cboRDArea.Text) & "')"
                
                ConnERP.Execute strSql
                
        End If
    
    'Create MP_Plan_Dim_Id
        
        strMP_Plan_Dim_Id = NextMPPlanDimID(strMP_Medium_Detail_ID)
        
    'insert mp_plan_dimension
        strSql = "insert into mp_plan_dimension(mp_plan_dim_id,mp_medium_detail_id,medium_code,"
        strSql = strSql & "spot_type,version,duration,rd_stations,gross_rate,nett_rate,rate_per_spot,gross_rate_per_spot,total_q1,total_q2,total_q3,total_q4) "
        strSql = strSql & " values ('" & strMP_Plan_Dim_Id & "','" & strMP_Medium_Detail_ID & "','RD','"
        strSql = strSql & CboRDSpotType.Text & "','" & Clear_String(cboRDVersion.Text) & "'," & txtRDDuration.Text & "," & txtRDStation.Text & ",0,0," & RemoveNumberFormat(txtRDRPS.Text) & "," & RemoveNumberFormat(txtRDRPSGross.Text) & ",0,0,0,0)"
        ConnERP.Execute strSql
        
    'update mp_master
        ConnERP.Execute "update mp_master set last_update_by='" & strLogin_User & "',last_update_date = getdate() where mp_number='" & txtMPNumber.Text & "'"
    
    pesan = MsgBox("New Medium Added!", vbExclamation, strApplication_Name)
    Call DisableControlTabRadio(True)
    Call EmptyTabRadio
    Call EnableObject(False)
    strMode = ""
End Sub


Private Sub AddCN()
'<CSCM>
'*****************************************************************************
' Nama Prosedur     :   AddCN
' Fungsi Prosedur   :   Add Medium Detail and plan Dimension for Cinema
' Parameter  Input  :
' Parameter Output  :
' Tgl Pembuatan     :   19 Agustus 2004
' Last Update/By    :   19 Agustus 2004/Sistyo
' Date              : 4/4/2016
' LastUpdate/By     : Tedi / Kreatif
' Name Before       : -
'********************************************************************************
'</CSCM>

    Dim strSql As String, strMP_Medium_Id As String, strMP_Medium_Detail_ID As String, strMP_Plan_Dim_Id As String
    Dim counter As Integer
    Dim pesan
    
    'field validation
        If cboCNName.Text = "" Then
            pesan = MsgBox("Not enough medium information!", vbExclamation, strApplication_Name)
            Exit Sub
        End If
        
        If cboCNSpotType.Text = "" Then
            pesan = MsgBox("Not enough medium information!", vbExclamation, strApplication_Name)
            Exit Sub
        End If
        
        If cboCNVersion.Text = "" Then
            pesan = MsgBox("Not enough medium information!", vbExclamation, strApplication_Name)
            Exit Sub
        End If
        
        If cboCNStudio.Text = "" Then
            pesan = MsgBox("Not enough medium information!", vbExclamation, strApplication_Name)
            Exit Sub
        End If
        
        
    'end of field validation
    
    strMP_Medium_Id = ""
    strMP_Medium_Detail_ID = ""
    
    'get MP_Medium_ID
        strSql = "select mp_medium_id from mp_medium where mp_activity_id='" & frm_MPEdit.tdg_Activity.Columns(1) & "' and medium_code='CN'"
        recTemp.Open strSql, ConnERP, 1, 3
        If Not recTemp.EOF Then
            strMP_Medium_Id = recTemp(0)
        End If
        recTemp.Close
        
    'get mp_medium_detail_id
        strSql = "select mp_medium_detail_id from mp_medium_detail where mp_medium_id = '" & strMP_Medium_Id & "' and cinema_code='" & cboCNCode.Text & "' and cinema_studio = '" & cboCNStudio.Text & "'"
        recTemp.Open strSql, ConnERP, 1, 3
        If Not recTemp.EOF Then
            strMP_Medium_Detail_ID = recTemp(0)
        End If
        recTemp.Close
    
    'create mp_medium_detail_id when mp_medium_detail_id is not defined
        If strMP_Medium_Detail_ID = "" Then
            strMP_Medium_Detail_ID = NextMPMediumDetailID(strMP_Medium_Id)
            'insert mp_medium_detail
                strSql = "insert into mp_medium_detail(mp_medium_detail_id,mp_medium_id,medium_code,cinema_code,cinema_name,cinema_studio) values ('"
                strSql = strSql & strMP_Medium_Detail_ID & "','" & strMP_Medium_Id & "','CN','"
                strSql = strSql & cboCNCode.Text & "','" & Clear_String(cboCNName.Text) & "','" & cboCNStudio.Text & "')"
                ConnERP.Execute strSql
        End If
    
    'Create MP_Plan_Dim_Id
        strMP_Plan_Dim_Id = NextMPPlanDimID(strMP_Medium_Detail_ID)
        
    'insert mp_plan_dimension
        strSql = "insert into mp_plan_dimension(mp_plan_dim_id,mp_medium_detail_id,medium_code,"
        strSql = strSql & "spot_type,version,duration,cinema_duration,nett_rate,gross_rate,total_q1,total_q2,total_q3,total_q4,rate_per_spot,gross_rate_per_spot) "
        strSql = strSql & " values ('" & strMP_Plan_Dim_Id & "','" & strMP_Medium_Detail_ID & "','CN','"
        strSql = strSql & cboCNSpotType.Text & "','" & Clear_String(cboCNVersion.Text) & "'," & Val(txtCNDuration.Text) & ",'" & cboCNJenisDurasi.Text & "',0,0,0,0,0,0," & RemoveNumberFormat(txtCNRate.Text) & "," & RemoveNumberFormat(txtCNRateGross.Text) & ")"
        ConnERP.Execute strSql
         
    'update mp_master
        ConnERP.Execute "update mp_master set last_update_by='" & strLogin_User & "',last_update_date = getdate() where mp_number='" & txtMPNumber.Text & "'"
    
    pesan = MsgBox("New Medium Added!", vbExclamation, strApplication_Name)
    Call DisableControlTabCinema(True)
    Call EmptyTabCinema
    Call EnableObject(False)
    strMode = ""
    If tdg_FGCNMedium.ApproxCount > 0 Then
        'If tdg_FGCNMedium.Row = 0 Then tdg_FGCNMedium.Row = 1
            FGCNMedium_Click
    End If
    
End Sub
Private Sub BeforeDeleteMediumTDG(medium_code As String, TrueDBGRID As Object)
'<CSCM>
'********************************************************************************
'Procedure Name     : BeforeDeleteMedium
'Procedure Function : ---
'Input Parameter    : ---
'Output Parameter   : ---
'Date               : 4/4/2016
'LastUpdate/By      : Tedi / Kreatif
'Name Before        : -
'********************************************************************************
'</CSCM>
'*****************************************************************************
' Nama Prosedur     :   BeforeDeleteMedium
' Fungsi Prosedur   :   Confirmation for data deletion
' Parameter  Input  :
' Parameter Output  :
' Tgl Pembuatan     :   13 Agustus 2004
' Last Update/By    :   13 Agustus 2004/Sistyo
'*****************************************************************************
    Dim pesan, konfirmasi
    Dim intCol, intRow As Integer
    Dim str_MP_Plan_Dim_Id As String, str_mp_medium_id As String
    Dim strSql As String
    Dim recTemp As New ADODB.Recordset
    Dim int_datacount As Integer
    Dim isApproved  As Boolean
    Dim IsHaveInsertion As Boolean
    
    isApproved = False
    IsHaveInsertion = False
    
    pesan = MsgBox("Do you want to delete this record?", vbExclamation + vbYesNo, strApplication_Name)
    If pesan = 6 Then  ' user click yes, maka do delete record!!!
        str_MP_Plan_Dim_Id = TrueDBGRID.Columns(0)
        
        str_mp_medium_id = ""
        strSql = "select mp_medium_id from mp_ids where mp_plan_dim_id = '" & str_MP_Plan_Dim_Id & "'"
        recTemp.Open strSql, ConnERP, 1, 3
            str_mp_medium_id = recTemp(0)
        If Not recTemp.EOF Then
        End If
        recTemp.Close
        
        'Check Insertion
        strSql = "select count(*) from mp_Insertion where mp_plan_dim_id = '" & str_MP_Plan_Dim_Id & "'"
        recTemp.Open strSql, ConnERP, 1, 3
        int_datacount = recTemp(0)
        recTemp.Close
        
        If int_datacount <> 0 Then
            IsHaveInsertion = True
        End If
        
        
        strSql = "select count(*) from mp_monthly_activity where mp_medium_id = '" & str_mp_medium_id & "' and approval=1"
        recTemp.Open strSql, ConnERP, 1, 3
        int_datacount = recTemp(0)
        recTemp.Close
        
        If int_datacount <> 0 Then
            isApproved = True
        End If
        
        'If int_datacount <> 0 Then
        
        If IsHaveInsertion And isApproved Then
        
            MsgBox "Cannot delete record because it's contains data that has been approved!", vbExclamation, strApplication_Name
        
        Else
            'strSql = "select count(*) from mp_insertion where mp_plan_dim_id = '" & str_MP_Plan_Dim_Id & "'"
            'recTemp.Open strSql, ConnERP, 1, 3
            'int_datacount = recTemp(0)
            'recTemp.Close
            'If int_datacount <> 0 Then
            If IsHaveInsertion Then
                pesan = MsgBox("The selected record contains insertion data." & vbCrLf & "Continue delete record?", vbYesNo + vbQuestion, strApplication_Name)
                If pesan = 6 Then
                    Call DeleteRecord(tdg_FGCNMedium.Columns(1), tdg_FGCNMedium.Columns(0))
                End If
            Else
                Call DeleteRecord(tdg_FGCNMedium.Columns(1), tdg_FGCNMedium.Columns(0))
            End If
        End If
            
        
        
    End If

End Sub
Private Sub BeforeDeleteMediumTDGrid(medium_code As String, FlexGridName As Object)
'*****************************************************************************
' Nama Prosedur     :   BeforeDeleteMedium
' Fungsi Prosedur   :   Confirmation for data deletion
' Parameter  Input  :
' Parameter Output  :
' Tgl Pembuatan     :   13 Agustus 2004
' Last Update/By    :   13 Agustus 2004/Sistyo
'*****************************************************************************
    Dim pesan, konfirmasi
    Dim intCol, intRow As Integer
    Dim str_MP_Plan_Dim_Id As String, str_mp_medium_id As String
    Dim strSql As String
    Dim rsTemp As New ADODB.Recordset
    Dim int_datacount As Integer
    Dim isApproved  As Boolean
    Dim IsHaveInsertion As Boolean
    
    isApproved = False
    IsHaveInsertion = False
    
    With FlexGridName

        pesan = MsgBox("Do you want to delete this record?", vbExclamation + vbYesNo, strApplication_Name)

        If pesan = 6 Then  ' user click yes, maka do delete record!!!
            str_MP_Plan_Dim_Id = .Columns(0)
            
            str_mp_medium_id = ""
            strSql = "select mp_medium_id from mp_ids where mp_plan_dim_id = '" & str_MP_Plan_Dim_Id & "'"
            rsTemp.Open strSql, ConnERP, 1, 3
                str_mp_medium_id = rsTemp(0)
            If Not rsTemp.EOF Then
            End If
            rsTemp.Close
            
            'Check Insertion
            strSql = "select count(*) from mp_Insertion where mp_plan_dim_id = '" & str_MP_Plan_Dim_Id & "'"
            rsTemp.Open strSql, ConnERP, 1, 3
            int_datacount = rsTemp(0)
            rsTemp.Close
            
            If int_datacount <> 0 Then
                IsHaveInsertion = True
            End If
            
            
            strSql = "select count(*) from mp_monthly_activity where mp_medium_id = '" & str_mp_medium_id & "' and approval=1"
            rsTemp.Open strSql, ConnERP, 1, 3
            int_datacount = rsTemp(0)
            rsTemp.Close
            
            If int_datacount <> 0 Then
                isApproved = True
            End If
            
            'If int_datacount <> 0 Then
            
            If IsHaveInsertion And isApproved Then
            
                MsgBox "Cannot delete record because it's contains data that has been approved!", vbExclamation, strApplication_Name
            
            Else
                If IsHaveInsertion Then
                    pesan = MsgBox("The selected record contains insertion data." & vbCrLf & "Continue delete record?", vbYesNo + vbQuestion, strApplication_Name)
                    If pesan = 6 Then
                        Call DeleteRecord(.Columns(1), .Columns(0))
                    End If
                Else
                    Call DeleteRecord(.Columns(1), .Columns(0))
                End If
            End If
            
            
        End If
    End With
End Sub
Private Sub BeforeDeleteMedium(medium_code As String, FlexGridName As Object)
'<CSCM>
'********************************************************************************
'Procedure Name     : BeforeDeleteMedium
'Procedure Function : ---
'Input Parameter    : ---
'Output Parameter   : ---
'Date               : 4/4/2016
'LastUpdate/By      : Tedi / Kreatif
'Name Before        : -
'********************************************************************************
'</CSCM>
'*****************************************************************************
' Nama Prosedur     :   BeforeDeleteMedium
' Fungsi Prosedur   :   Confirmation for data deletion
' Parameter  Input  :
' Parameter Output  :
' Tgl Pembuatan     :   13 Agustus 2004
' Last Update/By    :   13 Agustus 2004/Sistyo
'*****************************************************************************
    Dim pesan, konfirmasi
    Dim intCol, intRow As Integer
    Dim str_MP_Plan_Dim_Id As String, str_mp_medium_id As String
    Dim strSql As String
    Dim rsTemp As New ADODB.Recordset
    Dim int_datacount As Integer
    Dim isApproved  As Boolean
    Dim IsHaveInsertion As Boolean
    
    isApproved = False
    IsHaveInsertion = False
    
    With FlexGridName
        For intCol = 3 To .cols - 1 'menandai record mana yg akan dihapus, sehingga pada saat msgbox konfirmasi popup, user tahu record yg akan ia hapus
            .col = intCol
            .CellBackColor = vbYellow
        Next
        pesan = MsgBox("Do you want to delete this record?", vbExclamation + vbYesNo, strApplication_Name)
        For intCol = 3 To .cols - 1  'menghilangkan tanda setelah ada konsirmasi dari user
            .col = intCol
            .CellBackColor = vbWhite
        Next
        If pesan = 6 Then  ' user click yes, maka do delete record!!!
            str_MP_Plan_Dim_Id = .TextMatrix(.Row, 1)
            
            str_mp_medium_id = ""
            strSql = "select mp_medium_id from mp_ids where mp_plan_dim_id = '" & str_MP_Plan_Dim_Id & "'"
            rsTemp.Open strSql, ConnERP, 1, 3
                str_mp_medium_id = rsTemp(0)
            If Not rsTemp.EOF Then
            End If
            rsTemp.Close
            
            'Check Insertion
            strSql = "select count(*) from mp_Insertion where mp_plan_dim_id = '" & str_MP_Plan_Dim_Id & "'"
            rsTemp.Open strSql, ConnERP, 1, 3
            int_datacount = rsTemp(0)
            rsTemp.Close
            
            If int_datacount <> 0 Then
                IsHaveInsertion = True
            End If
            
            
            strSql = "select count(*) from mp_monthly_activity where mp_medium_id = '" & str_mp_medium_id & "' and approval=1"
            rsTemp.Open strSql, ConnERP, 1, 3
            int_datacount = rsTemp(0)
            rsTemp.Close
            
            If int_datacount <> 0 Then
                isApproved = True
            End If
            
            'If int_datacount <> 0 Then
            
            If IsHaveInsertion And isApproved Then
            
                MsgBox "Cannot delete record because it's contains data that has been approved!", vbExclamation, strApplication_Name
            
            Else
                'strSql = "select count(*) from mp_insertion where mp_plan_dim_id = '" & str_MP_Plan_Dim_Id & "'"
                'rsTemp.Open strSql, connerp, 1, 3
                'int_datacount = rsTemp(0)
                'rsTemp.Close
                'If int_datacount <> 0 Then
                If IsHaveInsertion Then
                    pesan = MsgBox("The selected record contains insertion data." & vbCrLf & "Continue delete record?", vbYesNo + vbQuestion, strApplication_Name)
                    If pesan = 6 Then
                        Call DeleteRecord(.TextMatrix(.Row, 2), .TextMatrix(.Row, 1))
                    End If
                Else
                    Call DeleteRecord(.TextMatrix(.Row, 2), .TextMatrix(.Row, 1))
                End If
            End If
            
            
        End If
    End With

End Sub
Private Sub cmdDeleteTV_Click()
'<CSCM>
'********************************************************************************
'Procedure Name     : cmdDeleteTV_Click
'Procedure Function : ---
'Input Parameter    : ---
'Output Parameter   : ---
'Date               : 4/4/2016
'LastUpdate/By      : Tedi / Kreatif
'Name Before        : -
'********************************************************************************
'</CSCM>
    If cmdDeleteTV.Caption = "&Delete" Then
        '=====DELETE MEDIUM==============
        Call BeforeDeleteMediumTDG("TV", tdg_TVMedium)
        Call ViewMediumTrueDB("TV", tdg_TVMedium)
        'cmdDeleteTV.Enabled = False
    Else
        '========Cancel Edit========
        cmdDeleteTV.Caption = "&Delete"
        cmdAddTV.Enabled = True
        CmdEditTV.Caption = "&Edit"
        tdg_TVMedium.Enabled = True
        cboTVStationCode.Enabled = True
        cboTVStationName.Enabled = True
        cboTVMarketName.Enabled = True
        CmdNewTVStation.Enabled = True
    End If
End Sub

Private Sub cmdDeleteRD_Click()
'<CSCM>
'********************************************************************************
'Procedure Name     : cmdDeleteRD_Click
'Procedure Function : ---
'Input Parameter    : ---
'Output Parameter   : ---
'Date               : 4/4/2016
'LastUpdate/By      : Tedi / Kreatif
'Name Before        : -
'********************************************************************************
'</CSCM>
    If cmdDeleteRD.Caption = "&Delete" Then
        Call BeforeDeleteMedium("RD", FGRDMedium)
        Call ViewMedium("RD", FGRDMedium)
        'cmdDeleteRD.Enabled = False
    Else
        CmdEditRD.Caption = "&Edit"
        cmdDeleteRD.Caption = "&Delete"
        cmdAddRD.Enabled = True
        FGRDMedium.Enabled = True
        cmdRDNewArea.Enabled = True
        cboRDArea.Enabled = True
        CboRDSpotType.Enabled = True
    End If
End Sub


Private Sub cmdDeleteCN_Click()
'<CSCM>
'********************************************************************************
'Procedure Name     : cmdDeleteCN_Click
'Procedure Function : ---
'Input Parameter    : ---
'Output Parameter   : ---
'Date               : 4/4/2016
'LastUpdate/By      : Tedi / Kreatif
'Name Before        : -
'********************************************************************************
'</CSCM>
    If cmdDeleteCN.Caption = "&Delete" Then
        Call BeforeDeleteMediumTDGrid("CN", tdg_FGCNMedium)
        Call ViewMediumTrueDB("CN", tdg_FGCNMedium)
        'cmdDeleteCN.Enabled = False
    Else
        'SETTING TOMBOL
        cmdEditCN.Caption = "&Edit"
        cmdDeleteCN.Caption = "&Delete"
        cmdAddCN.Enabled = True
        tdg_FGCNMedium.Enabled = True
        
        cboCNCode.Enabled = True
        cboCNName.Enabled = True
        cboCNStudio.Enabled = True
        cboCNSpotType.Enabled = True
    End If
End Sub

Private Sub DeleteRecord(mp_medium_detail_id As String, mp_plan_dim_id As String)
'<CSCM>
'********************************************************************************
'Procedure Name     : DeleteRecord
'Procedure Function : ---
'Input Parameter    : ---
'Output Parameter   : ---
'Tgl Pembuatan      : 13 Agustus 2004
'Last Update/By     : 13 Agustus 2004/Sistyo
'Date               : 4/4/2016
'LastUpdate/By      : Tedi / Kreatif
'Name Before        : -
'********************************************************************************
'</CSCM>

    Dim RcCount As Integer, strSql As String
    'periksa jumlah child, jika cuma 1, parentnya di delete juga
        recTemp.Open "select count(*) from mp_plan_dimension where mp_medium_detail_id = '" & mp_medium_detail_id & "'", ConnERP, 1, 3
        RcCount = recTemp(0)
        recTemp.Close
    
    If RcCount = 1 Then 'directly delete mp_medium_detail
        strSql = "delete from mp_medium_detail where mp_medium_detail_id = '" & mp_medium_detail_id & "'"
        ConnERP.Execute strSql
    Else 'delete mp_plan_dimension only
        strSql = "delete from mp_plan_dimension where mp_plan_dim_id='" & mp_plan_dim_id & "'"
        ConnERP.Execute strSql
    End If
    
    'update mp_master
        ConnERP.Execute "update mp_master set last_update_by='" & strLogin_User & "',last_update_date = getdate() where mp_number='" & txtMPNumber.Text & "'"

        
End Sub


Private Sub cboTVStationCode_Click()
'<CSCM>
'********************************************************************************
'Procedure Name     : cboTVStationCode_Click
'Procedure Function : ---
'Input Parameter    : ---
'Output Parameter   : ---
'Date               : 4/4/2016
'LastUpdate/By      : Tedi / Kreatif
'Name Before        : -
'********************************************************************************
'</CSCM>
    cboTVStationName.ListIndex = cboTVStationCode.ListIndex
    Call loadTVProgram
End Sub

Private Sub loadTVProgram()
'<CSCM>
'********************************************************************************
'Procedure Name     : loadTVProgram
'Procedure Function : ---
'Input Parameter    : ---
'Output Parameter   : ---
'Date               : 4/4/2016
'LastUpdate/By      : Tedi / Kreatif
'Name Before        : -
'********************************************************************************
'</CSCM>
'Loading TV Program of selected tv station
    cboTVProgram.Clear
    recTemp.Open "select distinct programe_name from tv_program_New where station_code = '" & cboTVStationCode.Text & "' order by programe_name desc", ConnERP, 1, 3
    While Not recTemp.EOF
        cboTVProgram.AddItem recTemp(0)
        recTemp.MoveNext
    Wend
    recTemp.Close
End Sub

Private Sub cboTVStationName_Click()
'<CSCM>
'********************************************************************************
'Procedure Name     : cboTVStationName_Click
'Procedure Function : ---
'Input Parameter    : ---
'Output Parameter   : ---
'Date               : 4/4/2016
'LastUpdate/By      : Tedi / Kreatif
'Name Before        : -
'********************************************************************************
'</CSCM>
    cboTVStationCode.ListIndex = cboTVStationName.ListIndex
End Sub

Private Sub cboTVVersion_Click()
'<CSCM>
'********************************************************************************
'Procedure Name     : cboTVVersion_Click
'Procedure Function : ---
'Input Parameter    : ---
'Output Parameter   : ---
'Date               : 4/4/2016
'LastUpdate/By      : Tedi / Kreatif
'Name Before        : -
'********************************************************************************
'</CSCM>
    On Error Resume Next
    txtTVDuration.Text = cboMaterialDuration.List(cboTVVersion.ListIndex)
End Sub

Private Sub cmdAddTV_Click()
'<CSCM>
'********************************************************************************
'Procedure Name     : cmdAddTV_Click
'Procedure Function : ---
'Input Parameter    : ---
'Output Parameter   : ---
'Date               : 4/4/2016
'LastUpdate/By      : Tedi / Kreatif
'Name Before        : -
'********************************************************************************
'</CSCM>

    Call AddTV
    Call ViewMediumTrueDB("TV", tdg_TVMedium)

End Sub


Private Sub cmdAddRD_Click()
'<CSCM>
'********************************************************************************
'Procedure Name     : cmdAddRD_Click
'Procedure Function : ---
'Input Parameter    : ---
'Output Parameter   : ---
'Date               : 4/4/2016
'LastUpdate/By      : Tedi / Kreatif
'Name Before        : -
'********************************************************************************
'</CSCM>
    Call AddRadio
    Call ViewMedium("RD", FGRDMedium)
End Sub


Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    Call CloseRecordset(recTemp)
    Call CloseRecordset(recOT)
End Sub

Private Sub lstRDRateType_Click()
'<CSCM>
'********************************************************************************
'Procedure Name     : lstRDRateType_Click
'Procedure Function : ---
'Input Parameter    : ---
'Output Parameter   : ---
'Date               : 4/4/2016
'LastUpdate/By      : Tedi / Kreatif
'Name Before        : -
'********************************************************************************
'</CSCM>
    'Jika Edit, hanya boleh pilih 1 rate type
    Dim i As Integer
    If cmdAddRD2.Enabled = False Then
        For i = 0 To lstRDRateType.ListCount - 1
            If i <> lstRDRateType.ListIndex Then
                lstRDRateType.Selected(i) = False
            End If
        Next
    End If
End Sub

Private Sub Mnu_Monthly_Budget_Click()
'<CSCM>
'********************************************************************************
'Procedure Name     : Mnu_Monthly_Budget_Click
'Procedure Function : ---
'Input Parameter    : ---
'Output Parameter   : ---
'Date               : 4/4/2016
'LastUpdate/By      : Tedi / Kreatif
'Name Before        : -
'********************************************************************************
'</CSCM>
    Frm_MPOtherMonthlyBudget.show 1
    
End Sub

Private Sub msf_TVMedium_RowColChange()
'<CSCM>
'********************************************************************************
'Procedure Name     : msf_TVMedium_RowColChange
'Procedure Function : Mengosongkan dan mengisi control2 di tab 0
'Input Parameter    : ---
'Output Parameter   : ---
'Date               : 4/4/2016
'LastUpdate/By      : Tedi / Kreatif
'Name Before        : -
'********************************************************************************
'</CSCM>

    Call EmptyTab0
    Call ShowDetailTV

End Sub

Private Sub OptArea_Click()
'<CSCM>
'********************************************************************************
'Procedure Name     : OptArea_Click
'Procedure Function : ---
'Input Parameter    : ---
'Output Parameter   : ---
'Date               : 4/4/2016
'LastUpdate/By      : Tedi / Kreatif
'Name Before        : -
'********************************************************************************
'</CSCM>

    If strMode = "Add" Then
        Call DisableControlTabRadio(True)
        Call EnableObject(False)
        strMode = ""
        'Exit Sub
    End If
    Dim i As Integer
    strMode = ""
    If CmdEditRD.Caption = "&Save" Then
        Call CmdEditRD_Click
    ElseIf CmdEditRD2.Caption = "&Save" Then
        Call CmdEditRD2_Click
    End If
    Call DisableControlTabRadio(True)
    intTabEdit = 9
    If fra_RD.Height = fra_RDByStation.Height Then
        For i = fra_RDByStation.Height To 0 Step -1
            fra_RDByStation.Height = i
            If i Mod 150 = 0 Then
                fra_RDByStation.Refresh
                fra_RD.Refresh
            End If
        Next
    End If
    fra_RDByStation.Visible = False
    If FGRDMedium.Rows > 0 Then
       FGRDMedium.Row = 1
       Call FGRDMedium_Click
    End If
    FGRDMedium.Height = fra_RD.Height - FGRDMedium.Top - 100
End Sub

Private Sub optCNBrief_Click()
'<CSCM>
'********************************************************************************
'Procedure Name     : optCNBrief_Click
'Procedure Function : ---
'Input Parameter    : ---
'Output Parameter   : ---
'Date               : 4/4/2016
'LastUpdate/By      : Tedi / Kreatif
'Name Before        : -
'********************************************************************************
'</CSCM>
    If strMode = "Add" Then
        DisableControlTabCinema True
        EnableObject False
        strMode = ""
        Exit Sub
    End If
    strMode = ""
    If cmdEditCN.Caption = "&Save" Then
        Call cmdEditCN_Click
    ElseIf CmdEditRD2.Caption = "&Save" Then
        Call cmdEditCN2_Click
    End If
    
    Call DisableControlTabCinema(True)
    Call EnableObject(False)
    intTabEdit = 9

    
    Dim i As Integer
    If blnOptCnBriefFirstClick Then
        Call initFrame_CN_brief
        blnOptCnBriefFirstClick = False
    End If
    frame_CN_brief.Visible = True
    frame_CN_brief.Refresh
    If frame_CN_detail.Height <> frame_CN_brief.Height Then
        For i = frame_CN_brief.Height To frame_CN_detail.Height
            frame_CN_brief.Height = i
            If i Mod 150 = 0 Then
                frame_CN_brief.Refresh
            End If
        Next
    End If
    If optCNBrief.Value = True Then
        If tdg_FGCNMedium2.ApproxCount > 0 Then
            'If tdg_FGCNMedium2.Row = 0 Then tdg_FGCNMedium2.Row = 1
            FGCNMedium2_Click
        End If
    Else
        If tdg_FGCNMedium.ApproxCount > 0 Then
            If tdg_FGCNMedium.Row = 0 Then tdg_FGCNMedium.Row = 1
            FGCNMedium_Click
        End If
    End If
End Sub

Private Sub initFrame_CN_brief()
'<CSCM>
'********************************************************************************
'Procedure Name     : initFrame_CN_brief
'Procedure Function : ---
'Input Parameter    : ---
'Output Parameter   : ---
'Date               : 4/4/2016
'LastUpdate/By      : Tedi / Kreatif
'Name Before        : -
'********************************************************************************
'</CSCM>
    Dim counter As Integer
     'Call ViewMedium("CN", FGCNMedium)
    Call ViewMediumTrueDB("CN2", tdg_FGCNMedium2)

End Sub
Private Sub optCNDetail_Click()
'<CSCM>
'********************************************************************************
'Procedure Name     : optCNDetail_Click
'Procedure Function : ---
'Input Parameter    : ---
'Output Parameter   : ---
'Date               : 4/4/2016
'LastUpdate/By      : Tedi / Kreatif
'Name Before        : -
'********************************************************************************
'</CSCM>
    If strMode = "Add" Then
        DisableControlTabCinema True
        EnableObject False
        strMode = ""
        Exit Sub
    End If
    strMode = ""
    If cmdEditCN.Caption = "&Save" Then
        Call cmdEditCN_Click
    ElseIf cmdEditCN2.Caption = "&Save" Then
        Call cmdEditCN2_Click
    End If
    Call DisableControlTabCinema(True)
    Call EnableObject(False)
    intTabEdit = 9
    
    Dim i As Integer
    If frame_CN_detail.Height = frame_CN_brief.Height Then
        For i = frame_CN_brief.Height To 0 Step -1
            frame_CN_brief.Height = i
            If i Mod 150 = 0 Then
                frame_CN_brief.Refresh
                frame_CN_detail.Refresh
            End If
        Next
    End If
    frame_CN_brief.Visible = False
    If optCNDetail.Value = True Then
        If tdg_FGCNMedium2.ApproxCount > 0 Then
            FGCNMedium2_Click
        End If
    Else
        If tdg_FGCNMedium.ApproxCount > 0 Then
            FGCNMedium_Click
        End If
    End If
End Sub

Private Sub OptStation_Click()
'<CSCM>
'********************************************************************************
'Procedure Name     : OptStation_Click
'Procedure Function : ---
'Input Parameter    : ---
'Output Parameter   : ---
'Date               : 4/4/2016
'LastUpdate/By      : Tedi / Kreatif
'Name Before        : -
'********************************************************************************
'</CSCM>
    
    If strMode = "Add" Then
        Call DisableControlTabRadio(True)
        Call EnableObject(False)
        strMode = ""
        Exit Sub
    End If
    strMode = ""
    If CmdEditRD.Caption = "&Save" Then
        Call CmdEditRD_Click
    ElseIf CmdEditRD2.Caption = "&Save" Then
        Call CmdEditRD2_Click
    End If
    intTabEdit = 9
    Dim i As Integer
    If blnOptRDStationFirstClick Then
        Call initFrameRDByStation
        blnOptRDStationFirstClick = False
    End If
    fra_RDByStation.Visible = True
    fra_RDByStation.Refresh
    If fra_RD.Height <> fra_RDByStation.Height Then
        For i = fra_RDByStation.Height To fra_RD.Height
            fra_RDByStation.Height = i
            If i Mod 150 = 0 Then
                fra_RDByStation.Refresh
            End If
        Next
    End If
    If FGRDMedium2.Rows > 0 Then
        FGRDMedium2.Row = 1
        Call FGRDMedium2_Click
    End If
    FGRDMedium2.Height = fra_RDByStation.Height - FGRDMedium2.Top - 100
End Sub

Private Sub initFrameRDByStation()
'<CSCM>
'********************************************************************************
'Procedure Name     : initFrameRDByStation
'Procedure Function : ---
'Input Parameter    : ---
'Output Parameter   : ---
'Date               : 4/4/2016
'LastUpdate/By      : Tedi / Kreatif
'Name Before        : -
'********************************************************************************
'</CSCM>
    Dim counter As Integer
    CboRDSpotType2.AddItem "Reguler"
    CboRDSpotType2.AddItem "Sponsorship/Program"
    With FGRDMedium2
        .cols = 10
        .TextMatrix(0, 0) = "NO."
        .TextMatrix(0, 1) = "" 'MP_Plan_Dim_Id
        .TextMatrix(0, 2) = "" 'MP_Medium_Detail_Id
        .TextMatrix(0, 3) = "Station"
        .TextMatrix(0, 4) = "Spot Type"
        .TextMatrix(0, 5) = "Version"
        .TextMatrix(0, 6) = "Duration"
        .TextMatrix(0, 7) = "Rate Type"
        .TextMatrix(0, 8) = "Gross rate"
        .TextMatrix(0, 9) = "Nett rate"
        .ColWidth(0) = 350
        .ColWidth(1) = 0
        .ColWidth(2) = 0
        .ColWidth(3) = 2000
        .ColWidth(4) = 1500
        .ColWidth(5) = 3000
        .ColWidth(6) = 1000
        .ColWidth(7) = 2400
        .ColWidth(8) = 2030
        .ColWidth(9) = 2030
        .Row = 0
        For counter = 1 To 9
            .col = counter
            .CellAlignment = 3
        Next
    End With

    Call ViewMedium("RD2", FGRDMedium2)
    
    Call LoadRadioStationCatalog
    Call LoadRadioRateType
    Call LoadRadioMaterial2
End Sub

Private Sub LoadRadioRateType()
'<CSCM>
'********************************************************************************
'Procedure Name     : LoadRadioRateType
'Procedure Function : ---
'Input Parameter    : ---
'Output Parameter   : ---
'Date               : 4/4/2016
'LastUpdate/By      : Tedi / Kreatif
'Name Before        : -
'********************************************************************************
'</CSCM>
    lstRDRateType.Clear
    recTemp.Open "select spot_type,spot_name from radio_spot_catalog", ConnERP, 1, 3
    While Not recTemp.EOF
        lstRDRateType.AddItem "[" & Trim(recTemp(0)) & "] " & recTemp(1)
        recTemp.MoveNext
    Wend
    recTemp.Close
End Sub

Private Sub LoadRadioStationCatalog()
'<CSCM>
'********************************************************************************
'Procedure Name     : LoadRadioStationCatalog
'Procedure Function : ---
'Input Parameter    : ---
'Output Parameter   : ---
'Date               : 4/4/2016
'LastUpdate/By      : Tedi / Kreatif
'Name Before        : -
'********************************************************************************
'</CSCM>
    Dim strSql As String
    Dim intAreaId As Integer, intCityId As Integer
    strSql = "select a.station_code,a.station_name,a.city_id,c.city,a.area_id,b.area_name from radio_station a "
    strSql = strSql & "inner join area b on a.area_id = b.area_id "
    strSql = strSql & "inner join city c on a.city_id = c.city_id "
    strSql = strSql & "order by b.area_name,c.city,a.station_name"
    recTemp.Open strSql, ConnERP, 1, 3
    While Not recTemp.EOF
        'load area
        If recTemp("area_id") <> intAreaId Then
            intAreaId = recTemp("area_id")
            trvRDStationCatalog.Nodes.Add , , "Area->" & intAreaId, recTemp("Area_name")
        End If
        'load City
        If recTemp("city_id") <> intCityId Then
            intCityId = recTemp("city_id")
            trvRDStationCatalog.Nodes.Add "Area->" & recTemp("area_id"), tvwChild, "City->" & intCityId, recTemp("City")
        End If
        'load station
        trvRDStationCatalog.Nodes.Add "City->" & recTemp("city_id"), tvwChild, recTemp("station_code"), recTemp("station_name")
        recTemp.MoveNext
    Wend
    recTemp.Close
    'tambah 1 blank area untuk blank space di bottom
    trvRDStationCatalog.Nodes.Add , , "Area->EOF", ""
End Sub

Private Sub SSTabMedium_Click(PreviousTab As Integer)
'<CSCM>
'********************************************************************************
'Procedure Name     : SSTabMedium_Click
'Procedure Function : ---
'Input Parameter    : ---
'Output Parameter   : ---
'Date               : 4/4/2016
'LastUpdate/By      : Tedi / Kreatif
'Name Before        : -
'********************************************************************************
'</CSCM>
    
    If intTabEdit <> 9 Then
        SSTabMedium.Tab = intTabEdit
    End If
    
    Select Case Trim(SSTabMedium.Tab)
        Case 0
            cmdDeleteTV.Enabled = False
            If blnSSTabMediumFirstClick(0) Then
                Call LoadTVProperties
                blnSSTabMediumFirstClick(0) = False
            End If
        Case 1
            cmdDeleteRD.Enabled = False
            If blnSSTabMediumFirstClick(1) Then
                Call LoadRDProperties
                blnSSTabMediumFirstClick(1) = False
            End If
            If optArea.Value = True Then
                If FGRDMedium2.Rows > 0 Then
                    FGRDMedium2.Row = 1
                    Call showDetailRD
                    
                End If
            Else
                

            End If
            
        Case 2
            cmdDeletePR.Enabled = False
            If blnSSTabMediumFirstClick(2) Then
                Call LoadPRProperties
                blnSSTabMediumFirstClick(2) = False
            End If
            DisableControlTabPrint True
            Call EmptyTabPrint
            If FGPRMedium.Rows > 0 Then
                If FGPRMedium.Row = 0 Then FGPRMedium.Row = 1
                Call FGPRMedium_Click
            End If
        Case 3
            optCNDetail.Value = True
            Call DisableControlTabCinema(True)
            cmdDeleteCN.Enabled = False
            If blnSSTabMediumFirstClick(3) Then
                Call LoadCNProperties
                blnSSTabMediumFirstClick(3) = False
            End If
            If tdg_FGCNMedium.ApproxCount > 0 Then
                'If tdg_FGCNMedium.Row = 0 Then tdg_FGCNMedium.Row = 1
                Call showDetailCinema
            End If
 
        Case 4
            cmdDeleteOT.Enabled = False
            If blnSSTabMediumFirstClick(4) Then
                Call LoadOTProperties
                blnSSTabMediumFirstClick(4) = False
            End If
    End Select
    
End Sub

Private Sub CheckCinemaRate()
'<CSCM>
'********************************************************************************
'Procedure Name     : CheckCinemaRate
'Procedure Function : ---
'Input Parameter    : ---
'Output Parameter   : ---
'Date               : 4/4/2016
'LastUpdate/By      : Tedi / Kreatif
'Name Before        : -
'********************************************************************************
'</CSCM>
    recTemp.Open "select * from cinema_rate where cinema_id='" & cboCNCode.Text & "' and studio='" & cboCNStudio.Text & "' and jenis='" & cboCNJenisDurasi.Text & "' and durasi=" & Val(txtCNDuration.Text), ConnERP, 1, 3
    If Not recTemp.EOF Then
        lblCNDuration.ForeColor = vbBlack
    Else
        lblCNDuration.ForeColor = vbRed
    End If
    recTemp.Close
End Sub

Private Sub SSTabMedium_GotFocus()
    If intTabEdit <> 9 Then
        SSTabMedium.Tab = intTabEdit
    End If
End Sub


Private Sub tdg_FGCNMedium_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
    If tdg_FGCNMedium.Columns(1) <> "" Then
        cmdDeleteCN.Enabled = True
        cmdEditCN.Enabled = True
    Else
        cmdDeleteCN.Enabled = False
        cmdEditCN.Enabled = False
    End If
    If tdg_FGCNMedium.ApproxCount > 0 Then
        
        Call showDetailCinema
    End If
End Sub

Private Sub tdg_FGCNMedium2_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
    FGCNMedium2_Click
End Sub



Private Sub tdg_OTMedium_Click()
    If tdg_OTMedium.col = 3 Then
        Frm_MPOtherMonthlyBudget.show 1
    End If
End Sub

Private Sub tdg_OTMedium_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    
'    If Button = 2 Then
'        If tdg_OTMedium.Columns(1).Text <> "" Then
'            PopupMenu MnuPopupOther
'        End If
'    End If

End Sub

Private Sub tdg_OTMedium_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
    If tdg_OTMedium.ApproxCount > 0 Then
        txtOTDescrition.Text = tdg_OTMedium.Columns(2).Text
    End If
End Sub

Private Sub tdg_TVMedium_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
    If tdg_TVMedium.Columns(1) <> "" Then
        cmdDeleteTV.Enabled = True
        CmdEditTV.Enabled = True
    Else
        cmdDeleteTV.Enabled = False
        CmdEditTV.Enabled = False
    End If
    If tdg_TVMedium.ApproxCount > 0 Then
        'If tdg_FGCNMedium.Row = 0 Then tdg_FGCNMedium.Row = 1
        Call ShowDetailTV
    End If
End Sub

Private Sub trvRDStationCatalog_DblClick()
'<CSCM>
'********************************************************************************
'Procedure Name     : trvRDStationCatalog_DblClick
'Procedure Function : ---
'Input Parameter    : ---
'Output Parameter   : ---
'Date               : 4/4/2016
'LastUpdate/By      : Tedi / Kreatif
'Name Before        : -
'********************************************************************************
'</CSCM>
    Dim pesan
    If Mid(trvRDStationCatalog.SelectedItem.KEY, 1, 6) <> "Area->" And Mid(trvRDStationCatalog.SelectedItem.KEY, 1, 6) <> "City->" Then
        On Error GoTo ErrLabel
        lvRDSelectedStation.ListItems.Add , trvRDStationCatalog.SelectedItem.KEY, trvRDStationCatalog.SelectedItem.Text
ErrLabel:
        If Err.Number = 35602 Then
            pesan = MsgBox("This station already added!", vbCritical + vbOKOnly, strApplication_Name)
        End If
    End If
End Sub

Private Sub txtCNDuration_Change()
'<CSCM>
'********************************************************************************
'Procedure Name     : txtCNDuration_Change
'Procedure Function : ---
'Input Parameter    : ---
'Output Parameter   : ---
'Date               : 4/4/2016
'LastUpdate/By      : Tedi / Kreatif
'Name Before        : -
'********************************************************************************
'</CSCM>
    Call CheckCinemaRate
End Sub

Private Sub txtPRCol_Change()
'<CSCM>
'********************************************************************************
'Procedure Name     : txtPRCol_Change
'Procedure Function : ---
'Input Parameter    : ---
'Output Parameter   : ---
'Date               : 4/4/2016
'LastUpdate/By      : Tedi / Kreatif
'Name Before        : -
'********************************************************************************
'</CSCM>
    TxtPRSize.Text = CStr(Val(txtPRCol.Text)) & " x " & CStr(Val(txtPRMM.Text))
End Sub

Private Sub txtPRCol_KeyPress(KeyAscii As Integer)
'<CSCM>
'********************************************************************************
'Procedure Name     : txtPRCol_KeyPress
'Procedure Function : ---
'Input Parameter    : ---
'Output Parameter   : ---
'Date               : 4/4/2016
'LastUpdate/By      : Tedi / Kreatif
'Name Before        : -
'********************************************************************************
'</CSCM>
    If (Chr(KeyAscii) < "0" Or Chr(KeyAscii) > "9") And KeyAscii <> 8 And KeyAscii <> 13 And KeyAscii <> 27 Then
        KeyAscii = 0
        Beep
    End If
End Sub

Private Sub txtPRCol_LostFocus()
'<CSCM>
'********************************************************************************
'Procedure Name     : txtPRCol_LostFocus
'Procedure Function : ---
'Input Parameter    : ---
'Output Parameter   : ---
'Date               : 4/4/2016
'LastUpdate/By      : Tedi / Kreatif
'Name Before        : -
'********************************************************************************
'</CSCM>
    txtPRCol.Text = Val(txtPRCol.Text)
End Sub

Private Sub TxtPRMediaCode_GotFocus()
'<CSCM>
'********************************************************************************
'Procedure Name     : TxtPRMediaCode_GotFocus
'Procedure Function : ---
'Input Parameter    : ---
'Output Parameter   : ---
'Date               : 4/4/2016
'LastUpdate/By      : Tedi / Kreatif
'Name Before        : -
'********************************************************************************
'</CSCM>

    Dim strSql As String, strMediaName As String
    strSql = "select media_name from media_print_catalog where print_code='" & TxtPRMediaCode.Text & "'"
    recTemp.Open strSql, ConnERP, 1, 3
    strMediaName = ""
    If Not recTemp.EOF Then
        strMediaName = recTemp(0)
    End If
    recTemp.Close
    txtPRMediaNameSearch.Text = strMediaName
    txtPRMediaNameSearch.Left = TxtPRMediaCode.Left
    txtPRMediaNameSearch.Visible = True
    If txtPRMediaNameSearch.Enabled Then
        txtPRMediaNameSearch.SetFocus
        txtPRMediaNameSearch.SelStart = Len(txtPRMediaNameSearch.Text)
    End If
End Sub

Private Sub FGPRRate_DblClick()
'<CSCM>
'********************************************************************************
'Procedure Name     : FGPRRate_DblClick
'Procedure Function : ---
'Input Parameter    : ---
'Output Parameter   : ---
'Date               : 4/4/2016
'LastUpdate/By      : Tedi / Kreatif
'Name Before        : -
'********************************************************************************
'</CSCM>
    With FGPRRate
        If .TextMatrix(.Row, 2) <> "" And .Row <> 0 Then
            txtPRMediaName.Text = .TextMatrix(.Row, 0)
            TxtPRMediaCode.Text = .TextMatrix(.Row, 1)
            TxtPRSatuan.Text = .TextMatrix(.Row, 2)
            TxtPRPaper.Text = .TextMatrix(.Row, 3)
            TxtPRColor.Text = .TextMatrix(.Row, 4)
            TxtPRMinSize.Text = .TextMatrix(.Row, 5)
            strPRGrossRate = .TextMatrix(.Row, 6)
            txtPRRateGross.Text = .TextMatrix(.Row, 6)
            strPRNettRate = .TextMatrix(.Row, 7)
            txtPRRate.Text = .TextMatrix(.Row, 7)
            If txtPRMediaNameSearch.Locked = False Then
                txtPRCol.Text = 0
                txtPRMM.Text = 0
            End If
            If .TextMatrix(.Row, 9) = "1" Then
                
                TxtPRSize.Text = txtPRCol.Text & " x " & txtPRMM.Text
                
                
            Else
                TxtPRSize.Text = "1"
                
            End If
            If cboPRSpotType.Enabled Then
                Call cboPRSpotType_Click
            Else
                cboPRSpotType.Enabled = True
                Call cboPRSpotType_Click
                cboPRSpotType.Enabled = False
            End If
            txtPRIsMMC.Text = .TextMatrix(.Row, 9)
        End If
        .Visible = False
    End With
    txtPRMediaNameSearch.Visible = False
    lblPrintCode.Caption = "Print Code : "
End Sub

Private Sub txtPRMediaNameSearch_Change()
'<CSCM>
'********************************************************************************
'Procedure Name     : txtPRMediaNameSearch_Change
'Procedure Function : ---
'Input Parameter    : ---
'Output Parameter   : ---
'Date               : 4/4/2016
'LastUpdate/By      : Tedi / Kreatif
'Name Before        : -
'********************************************************************************
'</CSCM>
    If txtPRMediaNameSearch.Text <> "" And txtPRMediaNameSearch.Visible Then
        Call SearchPRRate
    End If
End Sub

Private Sub txtPRMediaNameSearch_GotFocus()
'<CSCM>
'********************************************************************************
'Procedure Name     : txtPRMediaNameSearch_GotFocus
'Procedure Function : ---
'Input Parameter    : ---
'Output Parameter   : ---
'Date               : 4/4/2016
'LastUpdate/By      : Tedi / Kreatif
'Name Before        : -
'********************************************************************************
'</CSCM>
    
    If txtPRMediaNameSearch.Text <> "" Then
        txtPRMediaNameSearch.SelStart = 0
        txtPRMediaNameSearch.SelLength = Len(txtPRMediaNameSearch.Text)
    End If
'    PicPRAddDelete.Visible = False
    PicPRNewMaterial.Visible = False
    lblPrintCode.Caption = "Search : "
    Call SearchPRRate
End Sub

Private Sub txtPRMediaNameSearch_KeyPress(KeyAscii As Integer)
'<CSCM>
'********************************************************************************
'Procedure Name     : txtPRMediaNameSearch_KeyPress
'Procedure Function : ---
'Input Parameter    : ---
'Output Parameter   : ---
'Date               : 4/4/2016
'LastUpdate/By      : Tedi / Kreatif
'Name Before        : -
'********************************************************************************
'</CSCM>
    If KeyAscii = 27 Or KeyAscii = 13 Then
        txtPRMediaNameSearch.Visible = False
        FGPRRate.Visible = False
        Exit Sub
    End If
End Sub

Private Sub SearchPRRate()
'<CSCM>
'********************************************************************************
'Procedure Name     : SearchPRRate
'Procedure Function : ---
'Input Parameter    : ---
'Output Parameter   : ---
'Date               : 4/4/2016
'LastUpdate/By      : Tedi / Kreatif
'Name Before        : -
'********************************************************************************
'</CSCM>
    Dim strSql As String, baris As Integer, intBackColor As Double
    FGPRRate.Visible = False
    FGPRRate.Rows = 1
    strSql = "SELECT a.Media_Name, b.Print_Code, b.[Size], b.Paper, b.Color, b.Minimum_Size, b.Gross_Rate, b.Netto_Rate, b.Start_Valid_Date, isnull(c.ismmc,0) ismmc,b.notes FROM "
    strSql = strSql & "Media_Print_Rate b inner join print_size_catalog c on b.[size] = c.size_code "
    strSql = strSql & "INNER JOIN Media_Print_Catalog a ON b.Print_Code = a.Print_Code "
    strSql = strSql & "and a.media_name "
    If txtPRMediaNameSearch.Text <> "" Then
        strSql = strSql & "like '" & Clear_String(txtPRMediaNameSearch.Text) & "%'"
    Else
        strSql = strSql & "= '" & Clear_String(txtPRMediaNameSearch.Text) & "'"
    End If
    strSql = strSql & " order by a.media_name"
    baris = 0
    recTemp.Open strSql, ConnERP, 1, 3
    While Not recTemp.EOF
        baris = baris + 1
        If baris Mod 2 = 0 Then
            intBackColor = vbWhite
        Else
            intBackColor = 16761024 'ungu muda
        End If
        With FGPRRate
            .Rows = baris + 1
            .Row = baris
            
            .col = 0
            .CellBackColor = intBackColor
            .Text = recTemp("Media_Name")
            
            .col = 1
            .CellBackColor = intBackColor
            .Text = recTemp("Print_Code")
            
            .col = 2
            .CellBackColor = intBackColor
            .Text = recTemp("Size")
            
            .col = 3
            .CellBackColor = intBackColor
            .Text = recTemp("Paper")
            
            .col = 4
            .CellBackColor = intBackColor
            .Text = recTemp("Color")
            
            .col = 5
            .CellBackColor = intBackColor
            .Text = recTemp("Minimum_Size")
            
            .col = 6
            .CellBackColor = intBackColor
            .Text = IIf(IsNull(recTemp("Gross_rate")), FormatNumber(0, 2), FormatNumber(recTemp("Gross_rate"), 2))
            
            .col = 7
            .CellBackColor = intBackColor
            .Text = IIf(IsNull(recTemp("Netto_rate")), FormatNumber(0, 2), FormatNumber(recTemp("Netto_rate"), 2))
            
            .col = 8
            .CellBackColor = intBackColor
            .Text = recTemp("Start_Valid_Date")
            
            .col = 9
            .CellBackColor = intBackColor
            .Text = recTemp("ismmc")
            
            .col = 10
            .CellBackColor = intBackColor
            .Text = IIf(IsNull(recTemp("notes")), "", recTemp("notes"))
        End With
        recTemp.MoveNext
    Wend
    recTemp.Close
    'If baris > 0 Then
        FGPRRate.Visible = True
    'End If
    
End Sub

Private Sub txtPRMediaNameSearch_LostFocus()
'<CSCM>
'********************************************************************************
'Procedure Name     : txtPRMediaNameSearch_LostFocus
'Procedure Function : ---
'Input Parameter    : ---
'Output Parameter   : ---
'Date               : 4/4/2016
'LastUpdate/By      : Tedi / Kreatif
'Name Before        : -
'********************************************************************************
'</CSCM>
'    PicPRAddDelete.Visible = True
    PicPRNewMaterial.Visible = True
    If FGPRRate.Visible = False Then
        lblPrintCode.Caption = "Print Code : "
    End If
End Sub

Private Sub txtPRMM_Change()
'<CSCM>
'********************************************************************************
'Procedure Name     : txtPRMM_Change
'Procedure Function : ---
'Input Parameter    : ---
'Output Parameter   : ---
'Date               : 4/4/2016
'LastUpdate/By      : Tedi / Kreatif
'Name Before        : -
'********************************************************************************
'</CSCM>
    TxtPRSize.Text = CStr(Val(txtPRCol.Text)) & " x " & CStr(Val(txtPRMM.Text))
End Sub

Private Sub txtPRMM_KeyPress(KeyAscii As Integer)
'<CSCM>
'********************************************************************************
'Procedure Name     : txtPRMM_KeyPress
'Procedure Function : ---
'Input Parameter    : ---
'Output Parameter   : ---
'Date               : 4/4/2016
'LastUpdate/By      : Tedi / Kreatif
'Name Before        : -
'********************************************************************************
'</CSCM>
    If (Chr(KeyAscii) < "0" Or Chr(KeyAscii) > "9") And KeyAscii <> 8 And KeyAscii <> 13 And KeyAscii <> 27 Then
            KeyAscii = 0
            Beep
    End If
End Sub

Private Sub txtPRMM_LostFocus()
'<CSCM>
'********************************************************************************
'Procedure Name     : txtPRMM_LostFocus
'Procedure Function : ---
'Input Parameter    : ---
'Output Parameter   : ---
'Date               : 4/4/2016
'LastUpdate/By      : Tedi / Kreatif
'Name Before        : -
'********************************************************************************
'</CSCM>
    txtPRMM.Text = Val(txtPRMM.Text)
End Sub

Private Sub txtRDDuration_KeyPress(KeyAscii As Integer)
'<CSCM>
'********************************************************************************
'Procedure Name     : txtRDDuration_KeyPress
'Procedure Function : ---
'Input Parameter    : ---
'Output Parameter   : ---
'Date               : 4/4/2016
'LastUpdate/By      : Tedi / Kreatif
'Name Before        : -
'********************************************************************************
'</CSCM>
    If (Chr(KeyAscii) < "0" Or Chr(KeyAscii) > "9") And KeyAscii <> 8 And KeyAscii <> 13 And KeyAscii <> 27 Then
            KeyAscii = 0
            Beep
    End If
End Sub

Private Sub txtRDRPS_GotFocus()
'<CSCM>
'********************************************************************************
'Procedure Name     : txtRDRPS_GotFocus
'Procedure Function : ---
'Input Parameter    : ---
'Output Parameter   : ---
'Date               : 4/4/2016
'LastUpdate/By      : Tedi / Kreatif
'Name Before        : -
'********************************************************************************
'</CSCM>
    txtRDRPS.Text = RemoveNumberFormat(txtRDRPS.Text)
End Sub

Private Sub txtRDRPS_KeyPress(KeyAscii As Integer)
'<CSCM>
'********************************************************************************
'Procedure Name     : txtRDRPS_KeyPress
'Procedure Function : ---
'Input Parameter    : ---
'Output Parameter   : ---
'Date               : 4/4/2016
'LastUpdate/By      : Tedi / Kreatif
'Name Before        : -
'********************************************************************************
'</CSCM>
    If (Chr(KeyAscii) < "0" Or Chr(KeyAscii) > "9") And KeyAscii <> 8 And KeyAscii <> 13 And KeyAscii <> 27 Then
            If Chr(KeyAscii) <> "." Then
                KeyAscii = 0
                Beep
            End If
    End If
End Sub

Private Sub txtRDRPS_LostFocus()
'<CSCM>
'********************************************************************************
'Procedure Name     : txtRDRPS_LostFocus
'Procedure Function : ---
'Input Parameter    : ---
'Output Parameter   : ---
'Date               : 4/4/2016
'LastUpdate/By      : Tedi / Kreatif
'Name Before        : -
'********************************************************************************
'</CSCM>
    txtRDRPS.Text = FormatNumber(Val(txtRDRPS.Text), 2)
End Sub

Private Sub txtRDStation_KeyPress(KeyAscii As Integer)
'<CSCM>
'********************************************************************************
'Procedure Name     : txtRDStation_KeyPress
'Procedure Function : ---
'Input Parameter    : ---
'Output Parameter   : ---
'Date               : 4/4/2016
'LastUpdate/By      : Tedi / Kreatif
'Name Before        : -
'********************************************************************************
'</CSCM>
    If (Chr(KeyAscii) < "0" Or Chr(KeyAscii) > "9") And KeyAscii <> 8 And KeyAscii <> 13 And KeyAscii <> 27 Then
            KeyAscii = 0
            Beep
    End If
End Sub


Private Sub FGPRMedium_Click()
'<CSCM>
'********************************************************************************
'Procedure Name     : FGPRMedium_Click
'Procedure Function : ---
'Input Parameter    : ---
'Output Parameter   : ---
'Date               : 4/4/2016
'LastUpdate/By      : Tedi / Kreatif
'Name Before        : -
'********************************************************************************
'</CSCM>

    If FGPRMedium.TextMatrix(FGPRMedium.Row, 1) <> "" Then
        cmdDeletePR.Enabled = True
        cmdEditPR.Enabled = True
    Else
        cmdDeletePR.Enabled = False
        cmdEditPR.Enabled = False
    End If
    Call showDetailPR

End Sub

Private Sub FGRDMedium_Click()
'<CSCM>
'********************************************************************************
'Procedure Name     : FGRDMedium_Click
'Procedure Function : ---
'Input Parameter    : ---
'Output Parameter   : ---
'Date               : 4/4/2016
'LastUpdate/By      : Tedi / Kreatif
'Name Before        : -
'********************************************************************************
'</CSCM>
    If FGRDMedium.TextMatrix(FGRDMedium.Row, 1) <> "" Then
        cmdDeleteRD.Enabled = True
        CmdEditRD.Enabled = True
    Else
        cmdDeleteRD.Enabled = False
        CmdEditRD.Enabled = False
    End If
    Call showDetailRD
End Sub




Private Sub cboCNName_Click()
'<CSCM>
'********************************************************************************
'Procedure Name     : cboCNName_Click
'Procedure Function : ---
'Input Parameter    : ---
'Output Parameter   : ---
'Date               : 4/4/2016
'LastUpdate/By      : Tedi / Kreatif
'Name Before        : -
'********************************************************************************
'</CSCM>
    cboCNCode.ListIndex = cboCNName.ListIndex
    Call LoadStudioCatalog
    Call CheckCinemaRate
End Sub

Private Sub LoadStudioCatalog()
'<CSCM>
'********************************************************************************
'Procedure Name     : LoadStudioCatalog
'Procedure Function : ---
'Input Parameter    : ---
'Output Parameter   : ---
'Date               : 4/4/2016
'LastUpdate/By      : Tedi / Kreatif
'Name Before        : -
'********************************************************************************
'</CSCM>
    cboCNStudio.Clear
    recTemp.Open "select distinct studio from cinema_rate where cinema_id = '" & cboCNCode.Text & "'"
    While Not recTemp.EOF
        cboCNStudio.AddItem recTemp(0)
        recTemp.MoveNext
    Wend
    recTemp.Close
    
End Sub
Private Sub cmdAddCN_Click()
'<CSCM>
'********************************************************************************
'Procedure Name     : cmdAddCN_Click
'Procedure Function : ---
'Input Parameter    : ---
'Output Parameter   : ---
'Date               : 4/4/2016
'LastUpdate/By      : Tedi / Kreatif
'Name Before        : -
'********************************************************************************
'</CSCM>
    Call AddCN
    Call ViewMediumTrueDB("CN", tdg_FGCNMedium)
    
End Sub

Private Sub cmdAddCN2_Click()
'<CSCM>
'********************************************************************************
'Procedure Name     : cmdAddCN2_Click
'Procedure Function : ---
'Input Parameter    : ---
'Output Parameter   : ---
'Date               : 4/4/2016
'LastUpdate/By      : Tedi / Kreatif
'Name Before        : -
'********************************************************************************
'</CSCM>
    Call AddCN2
    Call ViewMediumTrueDB("CN2", tdg_FGCNMedium2)
End Sub

Private Sub msf_TVMedium_Click()
'<CSCM>
'********************************************************************************
'Procedure Name     : msf_TVMedium_Click
'Procedure Function : ---
'Input Parameter    : ---
'Output Parameter   : ---
'Date               : 4/4/2016
'LastUpdate/By      : Tedi / Kreatif
'Name Before        : -
'********************************************************************************
'</CSCM>
    If msf_TVMedium.TextMatrix(msf_TVMedium.Row, 1) <> "" Then
        cmdDeleteTV.Enabled = True
        CmdEditTV.Enabled = True
    Else
        cmdDeleteTV.Enabled = False
        CmdEditTV.Enabled = False
    End If
End Sub

Private Sub FGCNMedium_Click()
'<CSCM>
'********************************************************************************
'Procedure Name     : FGCNMedium_Click
'Procedure Function : ---
'Input Parameter    : ---
'Output Parameter   : ---
'Date               : 4/4/2016
'LastUpdate/By      : Tedi / Kreatif
'Name Before        : -
'********************************************************************************
'</CSCM>
    If tdg_FGCNMedium.Columns(1) <> "" Then
        cmdDeleteCN.Enabled = True
        cmdEditCN.Enabled = True
    Else
        cmdDeleteCN.Enabled = False
        cmdEditCN.Enabled = False
    End If
    If tdg_FGCNMedium.ApproxCount > 0 Then
        'If tdg_FGCNMedium.Row = 0 Then tdg_FGCNMedium.Row = 1
        Call showDetailCinema
    End If
    
End Sub

Private Sub cboRDArea_Click()
'<CSCM>
'********************************************************************************
'Procedure Name     : cboRDArea_Click
'Procedure Function : ---
'Input Parameter    : ---
'Output Parameter   : ---
'Date               : 4/4/2016
'LastUpdate/By      : Tedi / Kreatif
'Name Before        : -
'********************************************************************************
'</CSCM>

    txtRDStation.Text = cboRDStation.List(cboRDArea.ListIndex)

End Sub

Private Sub cboTVSpotType_Click()
'<CSCM>
'********************************************************************************
'Procedure Name     : cboTVSpotType_Click
'Procedure Function : ---
'Input Parameter    : ---
'Output Parameter   : ---
'Date               : 4/4/2016
'LastUpdate/By      : Tedi / Kreatif
'Name Before        : -
'********************************************************************************
'</CSCM>
    If cboTVSpotType.Text = "Reguler" Then
        cboTVVersion.Visible = True
        cboTVProgram.Visible = False
        'txtTVRate.Enabled = False
        'txtTVRate.Text = "0.00"
        'txtTVRateGross.Enabled = False
        'txtTVRateGross.Text = "0.00"
        Lbl_Version_TV.Caption = "Version :"
        'txtTVDuration.Enabled = True
        lbl_CPRP.Caption = "CPRP : "
        cmdNewTVMaterial.Enabled = True
    Else
        cboTVProgram.Top = cboTVVersion.Top
        cboTVProgram.Left = cboTVVersion.Left
        cboTVVersion.Visible = False
        cboTVProgram.Visible = True
        'txtTVRate.Enabled = True
        'txtTVRateGross.Enabled = True
        Lbl_Version_TV.Caption = "Program :"
        'txtTVDuration.Enabled = False
        lbl_CPRP.Caption = "Rate/insertion : "
        cmdNewTVMaterial.Enabled = False
    End If
End Sub

Private Sub cboPRSpotType_Click()
'<CSCM>
'********************************************************************************
'Procedure Name     : cboPRSpotType_Click
'Procedure Function : ---
'Input Parameter    : ---
'Output Parameter   : ---
'Date               : 4/4/2016
'LastUpdate/By      : Tedi / Kreatif
'Name Before        : -
'********************************************************************************
'</CSCM>
    If cboPRSpotType.Text <> "Reguler" Then
        txtPRRate.Enabled = True
        txtPRRateGross.Enabled = True
        
    Else
        txtPRRate.Enabled = False
        txtPRRate.Text = strPRNettRate
        txtPRRateGross.Enabled = False
        txtPRRateGross.Text = strPRGrossRate
        
    End If
End Sub

Private Sub cboRDVersion_Click()
'<CSCM>
'********************************************************************************
'Procedure Name     : cboRDVersion_Click
'Procedure Function : ---
'Input Parameter    : ---
'Output Parameter   : ---
'Date               : 4/4/2016
'LastUpdate/By      : Tedi / Kreatif
'Name Before        : -
'********************************************************************************
'</CSCM>
    txtRDDuration.Text = cboRDduration.List(cboRDVersion.ListIndex)
End Sub

Private Sub cboRDVersion2_Click()
'<CSCM>
'********************************************************************************
'Procedure Name     : cboRDVersion2_Click
'Procedure Function : ---
'Input Parameter    : ---
'Output Parameter   : ---
'Date               : 4/4/2016
'LastUpdate/By      : Tedi / Kreatif
'Name Before        : -
'********************************************************************************
'</CSCM>
    txtRDDuration2.Text = cboRDDuration2.List(cboRDVersion2.ListIndex)
End Sub



Private Sub cmdAddOT_Click()
'<CSCM>
'********************************************************************************
'Procedure Name     : cmdAddOT_Click
'Procedure Function : ---
'Input Parameter    : ---
'Output Parameter   : ---
'Date               : 4/4/2016
'LastUpdate/By      : Tedi / Kreatif
'Name Before        : -
'********************************************************************************
'</CSCM>
    Call AddOT
    Call ViewMediumTrueDB("OT", tdg_OTMedium)
End Sub

Private Sub AddOT()
'<CSCM>
'********************************************************************************
'Procedure Name     : AddOT
'Procedure Function : ---
'Input Parameter    : ---
'Output Parameter   : ---
'Date               : 4/4/2016
'LastUpdate/By      : Tedi / Kreatif
'Name Before        : -
'********************************************************************************
'</CSCM>
'*****************************************************************************
' Nama Prosedur     :   AddOT
' Fungsi Prosedur   :   Add Medium Detail and plan Dimension for Cinema
' Parameter  Input  :
' Parameter Output  :
' Tgl Pembuatan     :   19 Agustus 2004
' Last Update/By    :   19 Agustus 2004/Sistyo
'*****************************************************************************
    Dim strSql As String, strMP_Medium_Id As String, strMP_Medium_Detail_ID As String, strMP_Plan_Dim_Id As String
    Dim counter As Integer
    Dim pesan
    
    'field validation
        If Trim(txtOTDescrition.Text) = "" Then
            pesan = MsgBox("Please enter description!", vbExclamation, strApplication_Name)
            Exit Sub
        End If
    'end of field validation
    
    strMP_Medium_Id = ""
    strMP_Medium_Detail_ID = ""
    
    'get MP_Medium_ID
        strSql = "select mp_medium_id from mp_medium where mp_activity_id='" & frm_MPEdit.tdg_Activity.Columns(1) & "' and medium_code='OT'"
        recTemp.Open strSql, ConnERP, 1, 3
        If Not recTemp.EOF Then
            strMP_Medium_Id = recTemp(0)
        End If
        recTemp.Close
        
    'get mp_medium_detail_id
        strSql = "select mp_medium_detail_id from mp_medium_detail where mp_medium_id = '" & strMP_Medium_Id & "'"
        recTemp.Open strSql, ConnERP, 1, 3
        If Not recTemp.EOF Then
            strMP_Medium_Detail_ID = recTemp(0)
        End If
        recTemp.Close
    
    'create mp_medium_detail_id when mp_medium_detail_id is not defined
        If strMP_Medium_Detail_ID = "" Then
            strMP_Medium_Detail_ID = NextMPMediumDetailID(strMP_Medium_Id)
            'insert mp_medium_detail
                strSql = "insert into mp_medium_detail(mp_medium_detail_id,mp_medium_id,medium_code,supplier_code,supplier_name) values ('"
                strSql = strSql & strMP_Medium_Detail_ID & "','" & strMP_Medium_Id & "','OT','','')"
                ConnERP.Execute strSql
        End If
    
    'Create MP_Plan_Dim_Id
        
        strMP_Plan_Dim_Id = NextMPPlanDimID(strMP_Medium_Detail_ID)
        
    
    'insert mp_plan_dimension
        'strSQL = "insert into mp_plan_dimension(mp_plan_dim_id,mp_medium_detail_id,medium_code,"
        'strSQL = strSQL & "spot_type,ot_description,nett_rate,gross_rate,total_q1,total_q2,total_q3,total_q4) "
        'strSQL = strSQL & " values ('" & strMP_Plan_Dim_Id & "','" & strMP_Medium_Detail_ID & "','OT','"
        'strSQL = strSQL & cboOTSpotType.Text & "','" & Clear_Enter(Clear_String(txtOTDescrition.Text)) & "',0,0,0,0,0,0)"
        
        strSql = "insert into mp_plan_dimension(mp_plan_dim_id,mp_medium_detail_id,medium_code,"
        strSql = strSql & "ot_description,nett_rate,gross_rate,total_q1,total_q2,total_q3,total_q4) "
        strSql = strSql & " values ('" & strMP_Plan_Dim_Id & "','" & strMP_Medium_Detail_ID & "','OT','"
        strSql = strSql & Clear_String(txtOTDescrition.Text) & "',0,0,0,0,0,0)"
        
        
        ConnERP.Execute strSql
        
    'update mp_master
        ConnERP.Execute "update mp_master set last_update_by='" & strLogin_User & "',last_update_date = getdate() where mp_number='" & txtMPNumber.Text & "'"
    
    pesan = MsgBox("New Medium Added!", vbExclamation, strApplication_Name)
    txtOTDescrition.Text = Empty
End Sub

Private Sub FGOTMedium_Click()
'<CSCM>
'********************************************************************************
'Procedure Name     : FGOTMedium_Click
'Procedure Function : ---
'Input Parameter    : ---
'Output Parameter   : ---
'Date               : 4/4/2016
'LastUpdate/By      : Tedi / Kreatif
'Name Before        : -
'********************************************************************************
'</CSCM>
    If tdg_OTMedium.Columns(1) <> "" Then
        cmdDeleteOT.Enabled = True
        CmdEditOT.Enabled = True
    Else
        cmdDeleteOT.Enabled = False
        CmdEditOT.Enabled = False
    End If
    If tdg_OTMedium.ApproxCount > 0 Then
        txtOTDescrition.Text = tdg_OTMedium.Columns(2)
    End If
End Sub

Private Sub cmdDeleteOT_Click()
'<CSCM>
'********************************************************************************
'Procedure Name     : cmdDeleteOT_Click
'Procedure Function : ---
'Input Parameter    : ---
'Output Parameter   : ---
'Date               : 4/4/2016
'LastUpdate/By      : Tedi / Kreatif
'Name Before        : -
'********************************************************************************
'</CSCM>
     
    If cmdDeleteOT.Caption = "&Delete" Then
        Call BeforeDeleteMediumTDGrid("OT", tdg_OTMedium)
        Call ViewMediumTrueDB("OT", tdg_OTMedium)
        'cmdDeleteOT.Enabled = False
    Else
        'SETTING TOMBOL
        CmdEditOT.Caption = "&Edit"
        cmdDeleteOT.Caption = "&Delete"
        cmdAddOT.Enabled = True
        tdg_OTMedium.Enabled = True
        txtOTDescrition.Text = Empty
    End If
End Sub
Private Sub cmdDeleteOT_OLD_Click()
'<CSCM>
'********************************************************************************
'Procedure Name     : cmdDeleteOT_OLD_Click
'Procedure Function : ---
'Input Parameter    : ---
'Output Parameter   : ---
'Date               : 4/4/2016
'LastUpdate/By      : Tedi / Kreatif
'Name Before        : -
'********************************************************************************
'</CSCM>
'*****************************************************************************
' Nama Prosedur     :   cmdDeleteOT_Click
' Fungsi Prosedur   :   update tampilan flexgrid before / after delete
' Parameter  Input  :
' Parameter Output  :
' Tgl Pembuatan     :   19 Agustus 2004
' Last Update/By    :   19 Agustus 2004/Sistyo
'*****************************************************************************
    Dim pesan
    Dim intCol, intRow As Integer
    
    With tdg_OTMedium
    MsgBox "1"
'        For intCol = 3 To .cols - 1 'menandai record mana yg akan dihapus, sehingga pada saat msgbox konfirmasi popup, user tahu record yg akan ia hapus
'            .col = intCol
'            .CellBackColor = vbYellow
'        Next
'        pesan = MsgBox("Delete Record?", vbYesNo + vbQuestion)
'        For intCol = 3 To .cols - 1  'menghilangkan tanda setelah ada konsirmasi dari user
'            .col = intCol
'            .CellBackColor = vbWhite
'        Next
'        If pesan = 6 Then  ' user click yes, maka do delete record!!!
'            Call DeleteRecord(.TextMatrix(.Row, 2), .TextMatrix(.Row, 1))
'            If .Rows > 2 Then
'                For intRow = .Row To .Rows - 2
'                    .TextMatrix(intRow, 0) = intRow
'                    For intCol = 1 To .cols - 1
'                        .TextMatrix(intRow, intCol) = .TextMatrix(intRow + 1, intCol)
'                    Next
'                Next
'                .Rows = .Rows - 1
'            Else
'                For intCol = 0 To .cols - 1
'                    .TextMatrix(1, intCol) = ""
'                Next
'                cmdDeleteOT.Enabled = False
'            End If
'
'        End If
    End With
End Sub

Private Sub txtTVRate_GotFocus()
'<CSCM>
'********************************************************************************
'Procedure Name     : txtTVRate_GotFocus
'Procedure Function : ---
'Input Parameter    : ---
'Output Parameter   : ---
'Date               : 4/4/2016
'LastUpdate/By      : Tedi / Kreatif
'Name Before        : -
'********************************************************************************
'</CSCM>
    txtTVRate.Text = RemoveNumberFormat(txtTVRate.Text)
End Sub

Private Sub txtTVRate_KeyPress(KeyAscii As Integer)
'<CSCM>
'********************************************************************************
'Procedure Name     : txtTVRate_KeyPress
'Procedure Function : ---
'Input Parameter    : ---
'Output Parameter   : ---
'Date               : 4/4/2016
'LastUpdate/By      : Tedi / Kreatif
'Name Before        : -
'********************************************************************************
'</CSCM>
    If (Chr(KeyAscii) < "0" Or Chr(KeyAscii) > "9") And KeyAscii <> 8 And KeyAscii <> 13 And KeyAscii <> 27 Then
            If Chr(KeyAscii) <> "." Then
                KeyAscii = 0
                Beep
            End If
    End If
End Sub

Private Sub txtTVRate_LostFocus()
'<CSCM>
'********************************************************************************
'Procedure Name     : txtTVRate_LostFocus
'Procedure Function : ---
'Input Parameter    : ---
'Output Parameter   : ---
'Date               : 4/4/2016
'LastUpdate/By      : Tedi / Kreatif
'Name Before        : -
'********************************************************************************
'</CSCM>
    txtTVRate.Text = FormatNumber(Val(txtTVRate.Text), 2)
End Sub

Private Sub txtPRRate_GotFocus()
'<CSCM>
'********************************************************************************
'Procedure Name     : txtPRRate_GotFocus
'Procedure Function : ---
'Input Parameter    : ---
'Output Parameter   : ---
'Date               : 4/4/2016
'LastUpdate/By      : Tedi / Kreatif
'Name Before        : -
'********************************************************************************
'</CSCM>
    txtPRRate.Text = RemoveNumberFormat(txtPRRate.Text)
End Sub

Private Sub txtPRRate_KeyPress(KeyAscii As Integer)
'<CSCM>
'********************************************************************************
'Procedure Name     : txtPRRate_KeyPress
'Procedure Function : ---
'Input Parameter    : ---
'Output Parameter   : ---
'Date               : 4/4/2016
'LastUpdate/By      : Tedi / Kreatif
'Name Before        : -
'********************************************************************************
'</CSCM>
    If (Chr(KeyAscii) < "0" Or Chr(KeyAscii) > "9") And KeyAscii <> 8 And KeyAscii <> 13 And KeyAscii <> 27 Then
            If Chr(KeyAscii) <> "." Then
                KeyAscii = 0
                Beep
            End If
    End If
End Sub

Private Sub txtPRRate_LostFocus()
'<CSCM>
'********************************************************************************
'Procedure Name     : txtPRRate_LostFocus
'Procedure Function : ---
'Input Parameter    : ---
'Output Parameter   : ---
'Date               : 4/4/2016
'LastUpdate/By      : Tedi / Kreatif
'Name Before        : -
'********************************************************************************
'</CSCM>
    txtPRRate.Text = FormatNumber(Val(txtPRRate.Text), 2)
End Sub

Private Sub txtCNRate_GotFocus()
'<CSCM>
'********************************************************************************
'Procedure Name     : txtCNRate_GotFocus
'Procedure Function : ---
'Input Parameter    : ---
'Output Parameter   : ---
'Date               : 4/4/2016
'LastUpdate/By      : Tedi / Kreatif
'Name Before        : -
'********************************************************************************
'</CSCM>
    txtCNRate.Text = RemoveNumberFormat(txtCNRate.Text)
End Sub

Private Sub txtCNRate_KeyPress(KeyAscii As Integer)
'<CSCM>
'********************************************************************************
'Procedure Name     : txtCNRate_KeyPress
'Procedure Function : ---
'Input Parameter    : ---
'Output Parameter   : ---
'Date               : 4/4/2016
'LastUpdate/By      : Tedi / Kreatif
'Name Before        : -
'********************************************************************************
'</CSCM>
    If (Chr(KeyAscii) < "0" Or Chr(KeyAscii) > "9") And KeyAscii <> 8 And KeyAscii <> 13 And KeyAscii <> 27 Then
            If Chr(KeyAscii) <> "." Then
                KeyAscii = 0
                Beep
            End If
    End If
End Sub

Private Sub txtCNRate_LostFocus()
'<CSCM>
'********************************************************************************
'Procedure Name     : txtCNRate_LostFocus
'Procedure Function : ---
'Input Parameter    : ---
'Output Parameter   : ---
'Date               : 4/4/2016
'LastUpdate/By      : Tedi / Kreatif
'Name Before        : -
'********************************************************************************
'</CSCM>
    txtCNRate.Text = FormatNumber(Val(txtCNRate.Text), 2)
End Sub

Private Sub txtTVRateGross_GotFocus()
'<CSCM>
'********************************************************************************
'Procedure Name     : txtTVRateGross_GotFocus
'Procedure Function : ---
'Input Parameter    : ---
'Output Parameter   : ---
'Date               : 4/4/2016
'LastUpdate/By      : Tedi / Kreatif
'Name Before        : -
'********************************************************************************
'</CSCM>
    txtTVRateGross.Text = RemoveNumberFormat(txtTVRateGross.Text)
End Sub

Private Sub txtTVRateGross_KeyPress(KeyAscii As Integer)
'<CSCM>
'********************************************************************************
'Procedure Name     : txtTVRateGross_KeyPress
'Procedure Function : ---
'Input Parameter    : ---
'Output Parameter   : ---
'Date               : 4/4/2016
'LastUpdate/By      : Tedi / Kreatif
'Name Before        : -
'********************************************************************************
'</CSCM>
    If (Chr(KeyAscii) < "0" Or Chr(KeyAscii) > "9") And KeyAscii <> 8 And KeyAscii <> 13 And KeyAscii <> 27 Then
            If Chr(KeyAscii) <> "." Then
                KeyAscii = 0
                Beep
            End If
    End If
End Sub

Private Sub txtTVRateGross_LostFocus()
'<CSCM>
'********************************************************************************
'Procedure Name     : txtTVRateGross_LostFocus
'Procedure Function : ---
'Input Parameter    : ---
'Output Parameter   : ---
'Date               : 4/4/2016
'LastUpdate/By      : Tedi / Kreatif
'Name Before        : -
'********************************************************************************
'</CSCM>
    txtTVRateGross.Text = FormatNumber(Val(txtTVRateGross.Text), 2)
End Sub


Private Sub txtRDRPSGross_GotFocus()
'<CSCM>
'********************************************************************************
'Procedure Name     : txtRDRPSGross_GotFocus
'Procedure Function : ---
'Input Parameter    : ---
'Output Parameter   : ---
'Date               : 4/4/2016
'LastUpdate/By      : Tedi / Kreatif
'Name Before        : -
'********************************************************************************
'</CSCM>
    txtRDRPSGross.Text = RemoveNumberFormat(txtRDRPSGross.Text)
End Sub

Private Sub txtRDRPSGross_KeyPress(KeyAscii As Integer)
'<CSCM>
'********************************************************************************
'Procedure Name     : txtRDRPSGross_KeyPress
'Procedure Function : ---
'Input Parameter    : ---
'Output Parameter   : ---
'Date               : 4/4/2016
'LastUpdate/By      : Tedi / Kreatif
'Name Before        : -
'********************************************************************************
'</CSCM>
    If (Chr(KeyAscii) < "0" Or Chr(KeyAscii) > "9") And KeyAscii <> 8 And KeyAscii <> 13 And KeyAscii <> 27 Then
            If Chr(KeyAscii) <> "." Then
                KeyAscii = 0
                Beep
            End If
    End If
End Sub

Private Sub txtRDRPSGross_LostFocus()
'<CSCM>
'********************************************************************************
'Procedure Name     : txtRDRPSGross_LostFocus
'Procedure Function : ---
'Input Parameter    : ---
'Output Parameter   : ---
'Date               : 4/4/2016
'LastUpdate/By      : Tedi / Kreatif
'Name Before        : -
'********************************************************************************
'</CSCM>
    txtRDRPSGross.Text = FormatNumber(Val(txtRDRPSGross.Text), 2)
End Sub

Private Sub txtPRRateGRoss_GotFocus()
'<CSCM>
'********************************************************************************
'Procedure Name     : txtPRRateGRoss_GotFocus
'Procedure Function : ---
'Input Parameter    : ---
'Output Parameter   : ---
'Date               : 4/4/2016
'LastUpdate/By      : Tedi / Kreatif
'Name Before        : -
'********************************************************************************
'</CSCM>
    txtPRRateGross.Text = RemoveNumberFormat(txtPRRateGross.Text)
End Sub

Private Sub txtPRRateGross_KeyPress(KeyAscii As Integer)
'<CSCM>
'********************************************************************************
'Procedure Name     : txtPRRateGross_KeyPress
'Procedure Function : ---
'Input Parameter    : ---
'Output Parameter   : ---
'Date               : 4/4/2016
'LastUpdate/By      : Tedi / Kreatif
'Name Before        : -
'********************************************************************************
'</CSCM>
    If (Chr(KeyAscii) < "0" Or Chr(KeyAscii) > "9") And KeyAscii <> 8 And KeyAscii <> 13 And KeyAscii <> 27 Then
            If Chr(KeyAscii) <> "." Then
                KeyAscii = 0
                Beep
            End If
    End If
End Sub

Private Sub txtPRRateGross_LostFocus()
'<CSCM>
'********************************************************************************
'Procedure Name     : txtPRRateGross_LostFocus
'Procedure Function : ---
'Input Parameter    : ---
'Output Parameter   : ---
'Date               : 4/4/2016
'LastUpdate/By      : Tedi / Kreatif
'Name Before        : -
'********************************************************************************
'</CSCM>
    txtPRRateGross.Text = FormatNumber(Val(txtPRRateGross.Text), 2)
End Sub

Private Sub txtCNRateGross_GotFocus()
'<CSCM>
'********************************************************************************
'Procedure Name     : txtCNRateGross_GotFocus
'Procedure Function : ---
'Input Parameter    : ---
'Output Parameter   : ---
'Date               : 4/4/2016
'LastUpdate/By      : Tedi / Kreatif
'Name Before        : -
'********************************************************************************
'</CSCM>
    txtCNRateGross.Text = RemoveNumberFormat(txtCNRateGross.Text)
End Sub

Private Sub txtCNRateGross_KeyPress(KeyAscii As Integer)
'<CSCM>
'********************************************************************************
'Procedure Name     : txtCNRateGross_KeyPress
'Procedure Function : ---
'Input Parameter    : ---
'Output Parameter   : ---
'Date               : 4/4/2016
'LastUpdate/By      : Tedi / Kreatif
'Name Before        : -
'********************************************************************************
'</CSCM>
    If (Chr(KeyAscii) < "0" Or Chr(KeyAscii) > "9") And KeyAscii <> 8 And KeyAscii <> 13 And KeyAscii <> 27 Then
            If Chr(KeyAscii) <> "." Then
                KeyAscii = 0
                Beep
            End If
    End If
End Sub

Private Sub txtCNRateGross_LostFocus()
'<CSCM>
'********************************************************************************
'Procedure Name     : txtCNRateGross_LostFocus
'Procedure Function : ---
'Input Parameter    : ---
'Output Parameter   : ---
'Date               : 4/4/2016
'LastUpdate/By      : Tedi / Kreatif
'Name Before        : -
'********************************************************************************
'</CSCM>
    txtCNRateGross.Text = FormatNumber(Val(txtCNRateGross.Text), 2)
End Sub

Private Sub lvRDSelectedStation_DragDrop(Source As Control, X As Single, Y As Single)
'<CSCM>
'********************************************************************************
'Procedure Name     : lvRDSelectedStation_DragDrop
'Procedure Function : ---
'Input Parameter    : ---
'Output Parameter   : ---
'Date               : 4/4/2016
'LastUpdate/By      : Tedi / Kreatif
'Name Before        : -
'********************************************************************************
'</CSCM>
    Dim pesan
    If Source.Name = "trvRDStationCatalog" Then
        On Error GoTo ErrLabel
        lvRDSelectedStation.ListItems.Add , trvRDStationCatalog.SelectedItem.KEY, trvRDStationCatalog.SelectedItem.Text
ErrLabel:
        If Err.Number = 35602 Then
            pesan = MsgBox("This station already added!", vbCritical + vbOKOnly, strApplication_Name)
        End If
    End If
End Sub

Private Sub lvRDSelectedStation_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
'<CSCM>
'********************************************************************************
'Procedure Name     : lvRDSelectedStation_MouseDown
'Procedure Function : ---
'Input Parameter    : ---
'Output Parameter   : ---
'Date               : 4/4/2016
'LastUpdate/By      : Tedi / Kreatif
'Name Before        : -
'********************************************************************************
'</CSCM>
    Dim ListIndex As Integer
    If Button = 1 Then
        ListIndex = GetListIndex(lvRDSelectedStation, Y)
        If ListIndex <> -1 Then
            lvRDSelectedStation.ListItems(ListIndex).Selected = True
            lvRDSelectedStation.Drag
        End If
    End If
End Sub

Private Sub trvRDStationCatalog_DragDrop(Source As Control, X As Single, Y As Single)
'<CSCM>
'********************************************************************************
'Procedure Name     : trvRDStationCatalog_DragDrop
'Procedure Function : ---
'Input Parameter    : ---
'Output Parameter   : ---
'Date               : 4/4/2016
'LastUpdate/By      : Tedi / Kreatif
'Name Before        : -
'********************************************************************************
'</CSCM>
    If Source.Name = "lvRDSelectedStation" Then
        lvRDSelectedStation.ListItems.Remove lvRDSelectedStation.SelectedItem.Index
    End If
End Sub

Private Sub trvRDStationCatalog_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
'<CSCM>
'********************************************************************************
'Procedure Name     : trvRDStationCatalog_MouseDown
'Procedure Function : ---
'Input Parameter    : ---
'Output Parameter   : ---
'Date               : 4/4/2016
'LastUpdate/By      : Tedi / Kreatif
'Name Before        : -
'********************************************************************************
'</CSCM>
    Dim NodeIndex As Integer
    If Button = 1 Then
        If trvRDStationCatalog.Nodes.Count <> 0 Then
            NodeIndex = GetNodeIndex(trvRDStationCatalog, Y)
            trvRDStationCatalog.Nodes(NodeIndex).Selected = True
            lblRDPathInfo.Caption = "Path Info : " & trvRDStationCatalog.Nodes(NodeIndex).FullPath
            If Mid(trvRDStationCatalog.SelectedItem.KEY, 1, 6) <> "Area->" And Mid(trvRDStationCatalog.SelectedItem.KEY, 1, 6) <> "City->" Then
                trvRDStationCatalog.Drag
            End If
        End If
    End If
End Sub

Private Sub trvRDStationCatalog_NodeClick(ByVal Node As ComctlLib.Node)
'<CSCM>
'********************************************************************************
'Procedure Name     : trvRDStationCatalog_NodeClick
'Procedure Function : ---
'Input Parameter    : ---
'Output Parameter   : ---
'Date               : 4/4/2016
'LastUpdate/By      : Tedi / Kreatif
'Name Before        : -
'********************************************************************************
'</CSCM>
    lblRDPathInfo.Caption = "Path Info : " & Node.FullPath
End Sub

Private Sub lblRDPathInfo_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
'<CSCM>
'********************************************************************************
'Procedure Name     : lblRDPathInfo_MouseMove
'Procedure Function : ---
'Input Parameter    : ---
'Output Parameter   : ---
'Date               : 4/4/2016
'LastUpdate/By      : Tedi / Kreatif
'Name Before        : -
'********************************************************************************
'</CSCM>
    lblRDPathInfo.ToolTipText = lblRDPathInfo.Caption
End Sub

Private Sub txtRDRPS2_GotFocus()
'<CSCM>
'********************************************************************************
'Procedure Name     : txtRDRPS2_GotFocus
'Procedure Function : ---
'Input Parameter    : ---
'Output Parameter   : ---
'Date               : 4/4/2016
'LastUpdate/By      : Tedi / Kreatif
'Name Before        : -
'********************************************************************************
'</CSCM>
    txtRDRPS2.Text = RemoveNumberFormat(txtRDRPS2.Text)
End Sub

Private Sub txtRDRPS2_KeyPress(KeyAscii As Integer)
'<CSCM>
'********************************************************************************
'Procedure Name     : txtRDRPS2_KeyPress
'Procedure Function : ---
'Input Parameter    : ---
'Output Parameter   : ---
'Date               : 4/4/2016
'LastUpdate/By      : Tedi / Kreatif
'Name Before        : -
'********************************************************************************
'</CSCM>
    If (Chr(KeyAscii) < "0" Or Chr(KeyAscii) > "9") And KeyAscii <> 8 And KeyAscii <> 13 And KeyAscii <> 27 Then
            If Chr(KeyAscii) <> "." Then
                KeyAscii = 0
                Beep
            End If
    End If
End Sub

Private Sub txtRDRPS2_LostFocus()
'<CSCM>
'********************************************************************************
'Procedure Name     : txtRDRPS2_LostFocus
'Procedure Function : ---
'Input Parameter    : ---
'Output Parameter   : ---
'Date               : 4/4/2016
'LastUpdate/By      : Tedi / Kreatif
'Name Before        : -
'********************************************************************************
'</CSCM>
    txtRDRPS2.Text = FormatNumber(Val(txtRDRPS2.Text), 2)
End Sub

Private Sub txtRDRPSGross2_GotFocus()
'<CSCM>
'********************************************************************************
'Procedure Name     : txtRDRPSGross2_GotFocus
'Procedure Function : ---
'Input Parameter    : ---
'Output Parameter   : ---
'Date               : 4/4/2016
'LastUpdate/By      : Tedi / Kreatif
'Name Before        : -
'********************************************************************************
'</CSCM>
    txtRDRPSGross2.Text = RemoveNumberFormat(txtRDRPSGross2.Text)
End Sub

Private Sub txtRDRPSGross2_KeyPress(KeyAscii As Integer)
'<CSCM>
'********************************************************************************
'Procedure Name     : txtRDRPSGross2_KeyPress
'Procedure Function : ---
'Input Parameter    : ---
'Output Parameter   : ---
'Date               : 4/4/2016
'LastUpdate/By      : Tedi / Kreatif
'Name Before        : -
'********************************************************************************
'</CSCM>
    If (Chr(KeyAscii) < "0" Or Chr(KeyAscii) > "9") And KeyAscii <> 8 And KeyAscii <> 13 And KeyAscii <> 27 Then
            If Chr(KeyAscii) <> "." Then
                KeyAscii = 0
                Beep
            End If
    End If
End Sub

Private Sub txtRDRPSGross2_LostFocus()
'<CSCM>
'********************************************************************************
'Procedure Name     : txtRDRPSGross2_LostFocus
'Procedure Function : ---
'Input Parameter    : ---
'Output Parameter   : ---
'Date               : 4/4/2016
'LastUpdate/By      : Tedi / Kreatif
'Name Before        : -
'********************************************************************************
'</CSCM>
    txtRDRPSGross2.Text = FormatNumber(Val(txtRDRPSGross2.Text), 2)
End Sub

Private Sub AddRadio2()
'<CSCM>
'********************************************************************************
'Procedure Name     : AddRadio2
'Procedure Function : ---
'Input Parameter    : ---
'Output Parameter   : ---
'Date               : 4/4/2016
'LastUpdate/By      : Tedi / Kreatif
'Name Before        : -
'********************************************************************************
'</CSCM>
'*****************************************************************************
' Nama Prosedur     :   AddRAdio2
' Fungsi Prosedur   :   Add Medium Detail and plan Dimension for Radio per station
' Parameter  Input  :
' Parameter Output  :
' Tgl Pembuatan     :   1 Nov 2004
' Last Update/By    :   1 Nov 2004/Sistyo
'*****************************************************************************
    Dim strSql As String, strMP_Medium_Id As String
    Dim strMP_Medium_Detail_ID As String, strMP_Plan_Dim_Id As String
    Dim pesan
    Dim idxRDRateType As Integer, idxRDStation As Integer
    Dim rate_type_code As String, rate_type_name As String, Station_Code As String, Station_name As String
    Dim isRateListed As Boolean
    Dim int_inserted As Integer
    
    strMP_Medium_Id = ""
    
    'Field Validation
    If lvRDSelectedStation.ListItems.Count = 0 Then
        pesan = MsgBox("Please Select Station!", vbCritical + vbOKOnly, strApplication_Name)
        Exit Sub
    End If
    If CboRDSpotType2.Text = "" Then
        pesan = MsgBox("Please Select Spot Type!", vbCritical + vbOKOnly, strApplication_Name)
        Exit Sub
    End If
    If CboRDSpotType2.Text = "Reguler" Then
        If lstRDRateType.SelCount = 0 Then
            pesan = MsgBox("Please Select Rate Type!", vbCritical + vbOKOnly, strApplication_Name)
            Exit Sub
        End If
    End If
    If cboRDVersion2.Text = "" Then
        pesan = MsgBox("Please Select version!", vbCritical + vbOKOnly, strApplication_Name)
        Exit Sub
    End If
    'EO Field Validation
    
    'get MP_Medium_ID
    strSql = "select mp_medium_id from mp_medium where mp_activity_id='" & frm_MPEdit.tdg_Activity.Columns(1) & "' and medium_code='RD'"
    recTemp.Open strSql, ConnERP, 1, 3
    If Not recTemp.EOF Then
        strMP_Medium_Id = recTemp(0)
    End If
    recTemp.Close
 
    'Cek Station
    For idxRDStation = 1 To lvRDSelectedStation.ListItems.Count
        strMP_Medium_Detail_ID = ""
        Station_Code = lvRDSelectedStation.ListItems(idxRDStation).KEY
        Station_name = lvRDSelectedStation.ListItems(idxRDStation).Text
        'get mp_medium_detail_id
        strSql = "select mp_medium_detail_id from mp_medium_detail where mp_medium_id = '" & strMP_Medium_Id & "' and radio_station_code='" & Station_Code & "'"
        recTemp.Open strSql, ConnERP, 1, 3
        If Not recTemp.EOF Then
            strMP_Medium_Detail_ID = recTemp(0)
        End If
        recTemp.Close
        'create mp_medium_detail_id when mp_medium_detail_id is not defined
        If strMP_Medium_Detail_ID = "" Then
            
            strMP_Medium_Detail_ID = NextMPMediumDetailID(strMP_Medium_Id)
            
            'insert mp_medium_detail
            strSql = "insert into mp_medium_detail(mp_medium_detail_id,mp_medium_id,medium_code,radio_station_code,radio_station_name) values ('"
            strSql = strSql & strMP_Medium_Detail_ID & "','" & strMP_Medium_Id & "','RD','"
            strSql = strSql & Station_Code & "','" & Clear_String(Station_name) & "')"
            ConnERP.Execute strSql
        End If
        int_inserted = 0
        If CboRDSpotType2.Text = "Sponsorship/Program" Then
            'Create MP_Plan_Dim_Id
             strMP_Plan_Dim_Id = NextMPPlanDimID(strMP_Medium_Detail_ID)
            'Insert into MP_Plan_Dimension
            strSql = "insert into mp_plan_dimension(mp_plan_dim_id,mp_medium_detail_id,medium_code,"
            strSql = strSql & "spot_type,version,duration,gross_rate,nett_rate,rate_per_spot,gross_rate_per_spot,total_q1,total_q2,total_q3,total_q4,isRDByStation) "
            strSql = strSql & " values ('" & strMP_Plan_Dim_Id & "','" & strMP_Medium_Detail_ID & "','RD','"
            strSql = strSql & CboRDSpotType2.Text & "','" & Clear_String(cboRDVersion2.Text) & "'," & Val(txtRDDuration2.Text) & ",0,0," & RemoveNumberFormat(txtRDRPS2.Text) & "," & RemoveNumberFormat(txtRDRPSGross2.Text) & ",0,0,0,0,1)"
            ConnERP.Execute strSql
            int_inserted = int_inserted + 1
        Else
            For idxRDRateType = 0 To lstRDRateType.ListCount - 1
                If lstRDRateType.Selected(idxRDRateType) Then
                    rate_type_code = Mid(lstRDRateType.List(idxRDRateType), 2, InStr(1, lstRDRateType.List(idxRDRateType), "]") - 2)
                    rate_type_name = Right(lstRDRateType.List(idxRDRateType), Len(lstRDRateType.List(idxRDRateType)) - InStr(1, lstRDRateType.List(idxRDRateType), "]"))
                    'cek Rate Type
                    isRateListed = False
                    recTemp.Open "select * from radio_rate where station_code = '" & Station_Code & "' and prime_reg='" & rate_type_code & "'", ConnERP, 1, 3
                    If Not recTemp.EOF Then
                        isRateListed = True
                    End If
                    recTemp.Close
                    If isRateListed Then
                        'Create MP_Plan_Dim_Id
                        strMP_Plan_Dim_Id = NextMPPlanDimID(strMP_Medium_Detail_ID)
                        'Insert into MP_Plan_Dimension
                        strSql = "insert into mp_plan_dimension(mp_plan_dim_id,mp_medium_detail_id,medium_code,"
                        strSql = strSql & "spot_type,version,duration,rd_rate_type_code,rd_rate_type_name,gross_rate,nett_rate,rate_per_spot,gross_rate_per_spot,total_q1,total_q2,total_q3,total_q4,isRDByStation) "
                        strSql = strSql & " values ('" & strMP_Plan_Dim_Id & "','" & strMP_Medium_Detail_ID & "','RD','"
                        strSql = strSql & CboRDSpotType2.Text & "','" & Clear_String(cboRDVersion2.Text) & "'," & Val(txtRDDuration2.Text) & ",'" & rate_type_code & "','" & rate_type_name & "',0,0,0,0,0,0,0,0,1)"
                        ConnERP.Execute strSql
                        int_inserted = int_inserted + 1
                    End If
                End If
            Next
        End If
    Next
    
    'update mp_master
    ConnERP.Execute "update mp_master set last_update_by='" & strLogin_User & "',last_update_date = getdate() where mp_number='" & txtMPNumber.Text & "'"
    
    pesan = MsgBox(int_inserted & " Medium(s) Added!", vbExclamation, strApplication_Name)
    Call DisableControlTabRadio(True)
    Call EnableObject(False)
    strMode = ""
End Sub


Private Sub cboCNJenisDurasi_Click()
'<CSCM>
'********************************************************************************
'Procedure Name     : cboCNJenisDurasi_Click
'Procedure Function : ---
'Input Parameter    : ---
'Output Parameter   : ---
'Date               : 4/4/2016
'LastUpdate/By      : Tedi / Kreatif
'Name Before        : -
'********************************************************************************
'</CSCM>
    Call CheckCinemaRate
End Sub

Private Sub cboCNSpotType_Click()
'<CSCM>
'********************************************************************************
'Procedure Name     : cboCNSpotType_Click
'Procedure Function : ---
'Input Parameter    : ---
'Output Parameter   : ---
'Date               : 4/4/2016
'LastUpdate/By      : Tedi / Kreatif
'Name Before        : -
'********************************************************************************
'</CSCM>
    If cboCNSpotType.Text = "Reguler" Then
        txtCNRate.Enabled = False
        txtCNRate.Text = "0.00"
        txtCNRateGross.Enabled = False
        txtCNRateGross.Text = "0.00"
        'txtCNDuration.Enabled = True
    Else
        txtCNRate.Enabled = True
        txtCNRateGross.Enabled = True
        'txtCNDuration.Enabled = False
    End If
End Sub

Private Sub cboCNStudio_Click()
'<CSCM>
'********************************************************************************
'Procedure Name     : cboCNStudio_Click
'Procedure Function : ---
'Input Parameter    : ---
'Output Parameter   : ---
'Date               : 4/4/2016
'LastUpdate/By      : Tedi / Kreatif
'Name Before        : -
'********************************************************************************
'</CSCM>
    Call CheckCinemaRate
End Sub

Private Sub cboCNVersion_Click()
'<CSCM>
'********************************************************************************
'Procedure Name     : cboCNVersion_Click
'Procedure Function : ---
'Input Parameter    : ---
'Output Parameter   : ---
'Date               : 4/4/2016
'LastUpdate/By      : Tedi / Kreatif
'Name Before        : -
'********************************************************************************
'</CSCM>
    txtCNDuration.Text = cboCNMaterialDuration.List(cboCNVersion.ListIndex)
    If cboCNVersion.ListIndex = -1 Then Exit Sub
    cboCNJenisDurasi.Text = cboCNMaterialJenisDurasi.List(cboCNVersion.ListIndex)
End Sub

Private Sub CboRDSpotType2_Click()
'<CSCM>
'********************************************************************************
'Procedure Name     : CboRDSpotType2_Click
'Procedure Function : ---
'Input Parameter    : ---
'Output Parameter   : ---
'Date               : 4/4/2016
'LastUpdate/By      : Tedi / Kreatif
'Name Before        : -
'********************************************************************************
'</CSCM>
    If CboRDSpotType2.Text = "Sponsorship/Program" Then
        If strMode <> "" Then
            txtRDRPS2.Enabled = True
            txtRDRPS2.BackColor = &HFFFFFF
            txtRDRPSGross2.Enabled = True
            txtRDRPSGross2.BackColor = &HFFFFFF
            lstRDRateType.Enabled = False
            lstRDRateType.BackColor = &HE0E0E0
        End If
    Else
        If strMode <> "" Then
            txtRDRPS2.Enabled = False
            txtRDRPS2.BackColor = &HE0E0E0
            txtRDRPSGross2.Enabled = False
            txtRDRPSGross2.BackColor = &HE0E0E0
            lstRDRateType.Enabled = True
            lstRDRateType.BackColor = &HFFFFFF
            'txtRDDuration2.Enabled = True
        End If
    End If
End Sub

Private Sub cmdAddRD2_Click()
'<CSCM>
'********************************************************************************
'Procedure Name     : cmdAddRD2_Click
'Procedure Function : ---
'Input Parameter    : ---
'Output Parameter   : ---
'Date               : 4/4/2016
'LastUpdate/By      : Tedi / Kreatif
'Name Before        : -
'********************************************************************************
'</CSCM>
    Call AddRadio2
    Call ViewMedium("RD2", FGRDMedium2)
End Sub

Private Sub CmdDeleteRD2_Click()
'<CSCM>
'********************************************************************************
'Procedure Name     : CmdDeleteRD2_Click
'Procedure Function : ---
'Input Parameter    : ---
'Output Parameter   : ---
'Date               : 4/4/2016
'LastUpdate/By      : Tedi / Kreatif
'Name Before        : -
'********************************************************************************
'</CSCM>
    If CmdDeleteRD2.Caption = "&Delete" Then
        Call BeforeDeleteMedium("RD2", FGRDMedium2)
        Call ViewMedium("RD2", FGRDMedium2)
        'CmdDeleteRD2.Enabled = False
    Else
        CmdEditRD2.Caption = "&Edit"
        CmdDeleteRD2.Caption = "&Delete"
        cmdAddRD2.Enabled = True
        FGRDMedium2.Enabled = True
        trvRDStationCatalog.Enabled = True
        lvRDSelectedStation.ListItems.Clear
        lvRDSelectedStation.Enabled = True
        CboRDSpotType2.Enabled = True
        lstRDRateType.Enabled = True
    End If
End Sub

Private Sub cmdNewCNMaterial_Click()
'<CSCM>
'********************************************************************************
'Procedure Name     : cmdNewCNMaterial_Click
'Procedure Function : ---
'Input Parameter    : ---
'Output Parameter   : ---
'Date               : 4/4/2016
'LastUpdate/By      : Tedi / Kreatif
'Name Before        : -
'********************************************************************************
'</CSCM>

    objOpener = Me.Name
    Frm_Cinema_Material_Catalog.show 1
    Call LoadCinemaMaterial
    
End Sub


Private Sub cmdNewRadioMaterial_Click()
'<CSCM>
'********************************************************************************
'Procedure Name     : cmdNewRadioMaterial_Click
'Procedure Function : ---
'Input Parameter    : ---
'Output Parameter   : ---
'Date               : 4/4/2016
'LastUpdate/By      : Tedi / Kreatif
'Name Before        : -
'********************************************************************************
'</CSCM>
    objOpener = Me.Name
    Frm_Radio_Material_Catalog.show 1
    Call LoadRadioMaterial
    Call LoadRadioMaterial2
End Sub

Private Sub cmdNewRadioMaterial2_Click()
'<CSCM>
'********************************************************************************
'Procedure Name     : cmdNewRadioMaterial2_Click
'Procedure Function : ---
'Input Parameter    : ---
'Output Parameter   : ---
'Date               : 4/4/2016
'LastUpdate/By      : Tedi / Kreatif
'Name Before        : -
'********************************************************************************
'</CSCM>
    objOpener = Me.Name
    Frm_Radio_Material_Catalog.show 1
    Call LoadRadioMaterial
    Call LoadRadioMaterial2
End Sub

Private Sub cmdNewTVMaterial_Click()
'<CSCM>
'********************************************************************************
'Procedure Name     : cmdNewTVMaterial_Click
'Procedure Function : ---
'Input Parameter    : ---
'Output Parameter   : ---
'Date               : 4/4/2016
'LastUpdate/By      : Tedi / Kreatif
'Name Before        : -
'********************************************************************************
'</CSCM>
    objOpener = Me.Name
    Frm_TV_Material_Catalog.show 1
    Call LoadTVMaterial
End Sub

Private Sub cmdRDNewArea_Click()
'<CSCM>
'********************************************************************************
'Procedure Name     : cmdRDNewArea_Click
'Procedure Function : ---
'Input Parameter    : ---
'Output Parameter   : ---
'Date               : 4/4/2016
'LastUpdate/By      : Tedi / Kreatif
'Name Before        : -
'********************************************************************************
'</CSCM>
    
    Frm_MPRadioAreaAdd.show 1

End Sub


Private Sub FGRDMedium2_Click()
'<CSCM>
'********************************************************************************
'Procedure Name     : FGRDMedium2_Click
'Procedure Function : ---
'Input Parameter    : ---
'Output Parameter   : ---
'Date               : 4/4/2016
'LastUpdate/By      : Tedi / Kreatif
'Name Before        : -
'********************************************************************************
'</CSCM>
    If FGRDMedium2.TextMatrix(FGRDMedium2.Row, 1) <> "" Then
        If strMode <> "" Then
            CmdDeleteRD2.Enabled = True
            CmdEditRD2.Enabled = True
        End If
    Else
        If strMode <> "" Then
            CmdDeleteRD2.Enabled = False
            CmdEditRD2.Enabled = False
        End If
    End If
    If FGRDMedium2.Rows > 0 Then
        ShowDetailRDStasion
    End If
End Sub

Private Function NextMPMediumDetailID(strMPMediumID As String) As String
'<CSCM>
'********************************************************************************
'Procedure Name     : NextMPMediumDetailID
'Procedure Function : ---
'Input Parameter    : ---
'Output Parameter   : ---
'Date               : 4/4/2016
'LastUpdate/By      : Tedi / Kreatif
'Name Before        : -
'********************************************************************************
'</CSCM>
    Dim recTemp As New ADODB.Recordset, strSql As String
    strSql = "select isnull(max(cast(substring(mp_medium_detail_id,16,4) as int)),0)+1 from mp_medium_detail where mp_medium_id like '" & Mid(strMPMediumID, 1, 15) & "%'"
    recTemp.Open strSql, ConnERP, 1, 3
        NextMPMediumDetailID = Mid(strMPMediumID, 1, 4) & ".MDUD." & Mid(strMPMediumID, 11, 4) & "." & Right("0000" & CStr(recTemp(0)), 4) 'Create new medium_Detail_Id
    recTemp.Close
End Function

Private Function NextMPPlanDimID(strMPMediumDetailID As String) As String
'<CSCM>
'********************************************************************************
'Procedure Name     : NextMPPlanDimID
'Procedure Function : ---
'Input Parameter    : ---
'Output Parameter   : ---
'Date               : 4/4/2016
'LastUpdate/By      : Tedi / Kreatif
'Name Before        : -
'********************************************************************************
'</CSCM>
    Dim recTemp As New ADODB.Recordset, strSql As String
    strSql = "select isnull(max(cast(substring(mp_plan_dim_id,16,4) as int)),0)+1 from mp_plan_dimension where mp_medium_detail_id like '" & Mid(strMPMediumDetailID, 1, 15) & "%'"
    recTemp.Open strSql, ConnERP, 1, 3
        NextMPPlanDimID = Mid(strMPMediumDetailID, 1, 4) & ".MPDM." & Mid(strMPMediumDetailID, 11, 4) & "." & Right("0000" & CStr(recTemp(0)), 4) 'Create new plan dimension Id
    recTemp.Close
End Function

Private Sub txtCNRateGRoss2_GotFocus()
'<CSCM>
'********************************************************************************
'Procedure Name     : txtCNRateGRoss2_GotFocus
'Procedure Function : ---
'Input Parameter    : ---
'Output Parameter   : ---
'Date               : 4/4/2016
'LastUpdate/By      : Tedi / Kreatif
'Name Before        : -
'********************************************************************************
'</CSCM>
    txtCNRateGross2.Text = RemoveNumberFormat(txtCNRateGross2.Text)
End Sub

Private Sub txtCNRateGross2_KeyPress(KeyAscii As Integer)
'<CSCM>
'********************************************************************************
'Procedure Name     : txtCNRateGross2_KeyPress
'Procedure Function : ---
'Input Parameter    : ---
'Output Parameter   : ---
'Date               : 4/4/2016
'LastUpdate/By      : Tedi / Kreatif
'Name Before        : -
'********************************************************************************
'</CSCM>
    If (Chr(KeyAscii) < "0" Or Chr(KeyAscii) > "9") And KeyAscii <> 8 And KeyAscii <> 13 And KeyAscii <> 27 Then
            If Chr(KeyAscii) <> "." Then
                KeyAscii = 0
                Beep
            End If
    End If
End Sub

Private Sub txtCNRateGross2_LostFocus()
'<CSCM>
'********************************************************************************
'Procedure Name     : txtCNRateGross2_LostFocus
'Procedure Function : ---
'Input Parameter    : ---
'Output Parameter   : ---
'Date               : 4/4/2016
'LastUpdate/By      : Tedi / Kreatif
'Name Before        : -
'********************************************************************************
'</CSCM>
    txtCNRateGross2.Text = FormatNumber(Val(txtCNRateGross2.Text), 2)
End Sub

Private Sub txtCNRate2_GotFocus()
'<CSCM>
'********************************************************************************
'Procedure Name     : txtCNRate2_GotFocus
'Procedure Function : ---
'Input Parameter    : ---
'Output Parameter   : ---
'Date               : 4/4/2016
'LastUpdate/By      : Tedi / Kreatif
'Name Before        : -
'********************************************************************************
'</CSCM>
    txtCNRate2.Text = RemoveNumberFormat(txtCNRate2.Text)
End Sub

Private Sub txtCNRate2_KeyPress(KeyAscii As Integer)
'<CSCM>
'********************************************************************************
'Procedure Name     : txtCNRate2_KeyPress
'Procedure Function : ---
'Input Parameter    : ---
'Output Parameter   : ---
'Date               : 4/4/2016
'LastUpdate/By      : Tedi / Kreatif
'Name Before        : -
'********************************************************************************
'</CSCM>
    If (Chr(KeyAscii) < "0" Or Chr(KeyAscii) > "9") And KeyAscii <> 8 And KeyAscii <> 13 And KeyAscii <> 27 Then
            If Chr(KeyAscii) <> "." Then
                KeyAscii = 0
                Beep
            End If
    End If
End Sub

Private Sub txtCNRate2_LostFocus()
'<CSCM>
'********************************************************************************
'Procedure Name     : txtCNRate2_LostFocus
'Procedure Function : ---
'Input Parameter    : ---
'Output Parameter   : ---
'Date               : 4/4/2016
'LastUpdate/By      : Tedi / Kreatif
'Name Before        : -
'********************************************************************************
'</CSCM>
    txtCNRate2.Text = FormatNumber(Val(txtCNRate2.Text), 2)
End Sub

Private Sub AddCN2()
'<CSCM>
'********************************************************************************
'Procedure Name     : AddCN2
'Procedure Function : ---
'Input Parameter    : ---
'Output Parameter   : ---
'Date               : 4/4/2016
'LastUpdate/By      : Tedi / Kreatif
'Name Before        : -
'********************************************************************************
'</CSCM>
'*****************************************************************************
' Nama Prosedur     :   AddCN2
' Fungsi Prosedur   :   Add Medium Detail and plan Dimension for Cinema brief mode
' Parameter  Input  :
' Parameter Output  :
' Tgl Pembuatan     :   22 Juni 2005
' Last Update/By    :   22 JUni 2005/Sistyo
'*****************************************************************************
    Dim strSql As String, strMP_Medium_Id As String, strMP_Medium_Detail_ID As String, strMP_Plan_Dim_Id As String
    Dim counter As Integer
    Dim pesan
    
    If Trim(txtCNBrief.Text) = "" Then
        MsgBox "Please enter description!", vbExclamation, strApplication_Name

        Exit Sub
    End If
    
    strMP_Medium_Id = ""
    strMP_Medium_Detail_ID = ""
    
    'get MP_Medium_ID
        strSql = "select mp_medium_id from mp_medium where mp_activity_id='" & frm_MPEdit.tdg_Activity.Columns(1) & "' and medium_code='CN'"
        recTemp.Open strSql, ConnERP, 1, 3
        If Not recTemp.EOF Then
            strMP_Medium_Id = recTemp(0)
        End If
        recTemp.Close
        
    'get mp_medium_detail_id
        strSql = "select mp_medium_detail_id from mp_medium_detail where mp_medium_id = '" & strMP_Medium_Id & "' and cinema_code=''"
        recTemp.Open strSql, ConnERP, 1, 3
        If Not recTemp.EOF Then
            strMP_Medium_Detail_ID = recTemp(0)
        End If
        recTemp.Close
    
    'create mp_medium_detail_id when mp_medium_detail_id is not defined
        If strMP_Medium_Detail_ID = "" Then
            strMP_Medium_Detail_ID = NextMPMediumDetailID(strMP_Medium_Id)
            'insert mp_medium_detail
                strSql = "insert into mp_medium_detail(mp_medium_detail_id,mp_medium_id,medium_code,cinema_code,cinema_name,cinema_studio) values ('"
                strSql = strSql & strMP_Medium_Detail_ID & "','" & strMP_Medium_Id & "','CN','','','')"
                ConnERP.Execute strSql
        End If
    
    'Create MP_Plan_Dim_Id
        strMP_Plan_Dim_Id = NextMPPlanDimID(strMP_Medium_Detail_ID)
        
    'insert mp_plan_dimension
        strSql = "insert into mp_plan_dimension(mp_plan_dim_id,mp_medium_detail_id,medium_code,"
        strSql = strSql & "spot_type,version,duration,cinema_duration,nett_rate,gross_rate,total_q1,total_q2,total_q3,total_q4,rate_per_spot,gross_rate_per_spot,ot_description) "
        strSql = strSql & " values ('" & strMP_Plan_Dim_Id & "','" & strMP_Medium_Detail_ID & "','CN','"
        strSql = strSql & "','',0,'',0,0,0,0,0,0," & RemoveNumberFormat(txtCNRate2.Text) & "," & RemoveNumberFormat(txtCNRateGross2.Text) & ",'" & Clear_String(Clear_Enter(txtCNBrief.Text)) & "')"
        ConnERP.Execute strSql
        
    'update mp_master
        ConnERP.Execute "update mp_master set last_update_by='" & strLogin_User & "',last_update_date = getdate() where mp_number='" & txtMPNumber.Text & "'"
    
    pesan = MsgBox("New Medium Added!", vbExclamation, strApplication_Name)
    Call DisableControlTabCinema(True)
    Call EmptyTabCinema
    Call EnableObject(False)
    strMode = ""
    If tdg_FGCNMedium2.ApproxCount > 0 Then
    
            FGCNMedium_Click
    End If
    
End Sub

Private Sub cboPRVersion_Click()
'<CSCM>
'********************************************************************************
'Procedure Name     : cboPRVersion_Click
'Procedure Function : ---
'Input Parameter    : ---
'Output Parameter   : ---
'Date               : 4/4/2016
'LastUpdate/By      : Tedi / Kreatif
'Name Before        : -
'********************************************************************************
'</CSCM>
    If txtPRCol.Enabled Then
        txtPRCol.Text = cboPRVersionCol.List(cboPRVersion.ListIndex)
        txtPRMM.Text = cboPRVersionMM.List(cboPRVersion.ListIndex)
    End If
End Sub



Private Sub cmdAddPR_Click()
'<CSCM>
'********************************************************************************
'Procedure Name     : cmdAddPR_Click
'Procedure Function : ---
'Input Parameter    : ---
'Output Parameter   : ---
'Date               : 4/4/2016
'LastUpdate/By      : Tedi / Kreatif
'Name Before        : -
'********************************************************************************
'</CSCM>
    Call AddPR
    Call ViewMedium("PR", FGPRMedium)
End Sub

Private Sub CmdDeleteCN2_Click()
'<CSCM>
'********************************************************************************
'Procedure Name     : CmdDeleteCN2_Click
'Procedure Function : ---
'Input Parameter    : ---
'Output Parameter   : ---
'Date               : 4/4/2016
'LastUpdate/By      : Tedi / Kreatif
'Name Before        : -
'********************************************************************************
'</CSCM>
    If CmdDeleteCN2.Caption = "&Delete" Then
        Call BeforeDeleteMediumTDGrid("CN2", tdg_FGCNMedium2)
        Call ViewMediumTrueDB("CN2", tdg_FGCNMedium2)
        'CmdDeleteCN2.Enabled = False
    Else
        cmdEditCN2.Caption = "&Edit"
        CmdDeleteCN2.Caption = "&Delete"
        CmdAddCN2.Enabled = True
        tdg_FGCNMedium2.Enabled = True
        txtCNBrief.Text = Empty
        txtCNRate2.Text = Empty
        txtCNRateGross2.Text = Empty
    End If
End Sub

Private Sub cmdDeletePR_Click()
'<CSCM>
'********************************************************************************
'Procedure Name     : cmdDeletePR_Click
'Procedure Function : ---
'Input Parameter    : ---
'Output Parameter   : ---
'Date               : 4/4/2016
'LastUpdate/By      : Tedi / Kreatif
'Name Before        : -
'********************************************************************************
'</CSCM>
    If cmdDeletePR.Caption = "&Delete" Then
        Call BeforeDeleteMedium("PR", FGPRMedium)
        Call ViewMedium("PR", FGPRMedium)
        'cmdDeletePR.Enabled = False
    Else
        cmdEditPR.Caption = "&Edit"
        cmdDeletePR.Caption = "&Delete"
        cmdAddPR.Enabled = True
        FGPRMedium.Enabled = True
        txtPRMediaNameSearch.Locked = False
        cboPRSpotType.Enabled = True
    End If
End Sub

Private Sub cmdNewPRMaterial_Click()
'<CSCM>
'********************************************************************************
'Procedure Name     : cmdNewPRMaterial_Click
'Procedure Function : ---
'Input Parameter    : ---
'Output Parameter   : ---
'Date               : 4/4/2016
'LastUpdate/By      : Tedi / Kreatif
'Name Before        : -
'********************************************************************************
'</CSCM>
    objOpener = Me.Name
    Frm_Print_Material_Catalog.show 1
    Call LoadPRMaterial
End Sub

Private Sub FGCNMedium2_Click()
'<CSCM>
'********************************************************************************
'Procedure Name     : FGCNMedium2_Click
'Procedure Function : ---
'Input Parameter    : ---
'Output Parameter   : ---
'Date               : 4/4/2016
'LastUpdate/By      : Tedi / Kreatif
'Name Before        : -
'********************************************************************************
'</CSCM>
    If tdg_FGCNMedium2.Columns(0) <> "" Then
        CmdDeleteCN2.Enabled = True
        cmdEditCN2.Enabled = True
    Else
        CmdDeleteCN2.Enabled = False
        cmdEditCN2.Enabled = False
    End If
'         txtCNBrief.Text = FGCNMedium2.TextMatrix(FGCNMedium2.Row, 3)
'        txtCNRate2.Text = FormatNumber(FGCNMedium2.TextMatrix(FGCNMedium2.Row, 5), 2)
'        txtCNRateGross2.Text = FormatNumber(FGCNMedium2.TextMatrix(FGCNMedium2.Row, 4), 2)
   
    txtCNBrief.Text = tdg_FGCNMedium2.Columns(2)
    If tdg_FGCNMedium2.Columns(2) = "" Then
        txtCNRate2.Text = FormatNumber(0, 2)
    Else
        txtCNRate2.Text = FormatNumber(tdg_FGCNMedium2.Columns(4), 2)
    End If
    txtCNRateGross2.Text = FormatNumber(IIf(tdg_FGCNMedium2.Columns(3) = "", 0, tdg_FGCNMedium2.Columns(3)), 2)
    
End Sub

Private Sub LoadTVMarket()
'<CSCM>
'********************************************************************************
'Procedure Name     : LoadTVMarket
'Procedure Function : ---
'Input Parameter    : ---
'Output Parameter   : ---
'Date               : 4/4/2016
'LastUpdate/By      : Tedi / Kreatif
'Name Before        : -
'********************************************************************************
'</CSCM>
    If recTemp.State <> adStateClosed Then recTemp.Close
    recTemp.Open "select code,market_name from market_catalog", ConnERP, 1, 3
    While Not recTemp.EOF
        cboTVMarketCode.AddItem recTemp(0)
        cboTVMarketName.AddItem recTemp(1)
        recTemp.MoveNext
    Wend
    recTemp.Close
End Sub

Private Sub EnableObject(ByVal paIsEnable As Boolean)
'<CSCM>
'********************************************************************************
'Procedure Name     : EnableObject
'Procedure Function : ~ Enable/disable control di frame Entry.
'                     ~ Call SetButtonToolbar utk Toolbar/Statusbar AI (artificial intelligence).
'Input Parameter    : paIsEnable: True=Enable, False=Disable.
'Output Parameter   : ---
'Date               : 3/22/2016
'LastUpdate/By      : Abdi / Kreatif
'********************************************************************************
'</CSCM>

    Call SetButtonToolbar(Not paIsEnable, picButton) 'TOOLBAR_AI.
    
End Sub

Sub SetButtonToolbar(ByVal paIsNormalMode As Boolean, picOBJ) 'TOOLBAR_AI.
'<CSCM>
'********************************************************************************
'Procedure Name     : SetButtonToolbar
'Procedure Function : TOOLBAR_AI.
'Input Parameter    : ---
'Output Parameter   : ---
'Date               : 3/22/2016
'LastUpdate/By      : Abdi / Kreatif
'********************************************************************************
'</CSCM>

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
    
    With picButton(enButtonType.bieClose)  'EXIT.
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
        SetPictureTB element.Index, picOBJ
    Next element
                                                                                                                    '
    
'    'Add
'    If Mid(strDummy, 2, 1) = "0" Or Trim(Mid(strDummy, 2, 1)) = "" Then
'        picButton(enButtonType.bieAdd).Enabled = False
'        picButton(enButtonType.bieAdd).Picture = LoadPicture(SetButtonImageEffect(bieAdd, bieDisabled))
'    End If
'
'    'Edit
'    If Mid(strDummy, 3, 1) = "0" Or Trim(Mid(strDummy, 3, 1)) = "" Then
'        picButton(enButtonType.bieEdit).Enabled = False
'        picButton(enButtonType.bieEdit).Picture = LoadPicture(SetButtonImageEffect(bieEdit, bieDisabled))
'    End If
'
'    'Delete
'    If Mid(strDummy, 4, 1) = "0" Or Trim(Mid(strDummy, 4, 1)) = "" Then
'        picButton(enButtonType.bieDelete).Enabled = False
'        picButton(enButtonType.bieDelete).Picture = LoadPicture(SetButtonImageEffect(bieDelete, bieDisabled))
'    End If
    
    
    
End Sub

Sub SetPictureTB(ByVal Index As Integer, picOBJ)
'<CSCM>
'********************************************************************************
'Procedure Name     : SetPictureTB
'Procedure Function : TOOLBAR_AI saat mouse berada di area button.
'Input Parameter    : ---
'Output Parameter   : ---
'Date               : 3/22/2016
'LastUpdate/By      : Abdi / Kreatif
'********************************************************************************
'</CSCM>

    With picOBJ(Index) 'FIRST.
        
        If .Enabled = True Then
            .Picture = LoadPicture(SetButtonImageEffect(Index, bieNormal))
        Else: .Picture = LoadPicture(SetButtonImageEffect(Index, bieDisabled))
        End If
        
    End With
    
End Sub

Sub picButton_Obj(Index As Integer, X As Single, Y As Single, picOBJ)     'TOOLBAR_AI.
'<CSCM>
'********************************************************************************
'Procedure Name     : picButton_Obj
'Procedure Function : TOOLBAR_AI saat mouse berada di area button.
'Input Parameter    : ---
'Output Parameter   : ---
'Date               : 3/22/2016
'LastUpdate/By      : Abdi / Kreatif
'********************************************************************************
'</CSCM>

    If (X < 0) Or (Y < 0) Or (X > picOBJ(Index).Width) Or (Y > picOBJ(Index).Height) Then 'Dua IF ini jangan diubah keluar CASE agar API-nya jalan.
        ReleaseCapture 'The MOUSE_LEAVE pseudo-event.
        picOBJ(Index).Picture = LoadPicture(SetButtonImageEffect(Index, bieNormal)) 'Back to NORMAL.

    ElseIf GetCapture() <> picOBJ(Index).hwnd Then
        SetCapture picOBJ(Index).hwnd 'The MOUSE_ENTER pseudo-event.
        picOBJ(Index).Picture = LoadPicture(SetButtonImageEffect(Index, bieOver)) 'Set to OVER_EFFECT.
    End If
    
End Sub

Sub db_add()
'<CSCM>
'********************************************************************************
'Procedure Name     : db_Add
'Procedure Function : Menambang data baru
'Input Parameter    : ---
'Output Parameter   : ---
'Date               : 4/4/2016
'LastUpdate/By      : Tedi / Kreatif
'Name Before        : -
'********************************************************************************
'</CSCM>
    strMode = "Add"
    intTabEdit = SSTabMedium.Tab
    Select Case SSTabMedium.Tab
    
        Case 0 'Jika sedang diposisi tab =0
            
            Call DisableColorTab0(False)
            Call EmptyTab0
            
        Case 1
            Call DisableControlTabRadio(False)
            Call EmptyTabRadio
            
            If optArea.Value = True Then
                Call cboRDArea_Click
            End If
        Case 2
            Call DisableControlTabPrint(False)
            Call EmptyTabPrint
        
        Case 3
            Call DisableControlTabCinema(False)
            Call EmptyTabCinema
        Case 4
            Call DisableControlTabOther(False)
            Call EmptyTabOther
    End Select
    
    Call EnableObject(True)

End Sub

Sub db_edit()
'<CSCM>
'********************************************************************************
'Procedure Name     : db_Edit
'Procedure Function : Revisi Data
'Input Parameter    : ---
'Output Parameter   : ---
'Date               : 4/4/2016
'LastUpdate/By      : Tedi / Kreatif
'Name Before        : -
'********************************************************************************
'</CSCM>
    
    intTabEdit = SSTabMedium.Tab
    
    strMode = "Edit"
    
    Select Case SSTabMedium.Tab
    
        Case 0 'Jika sedang diposisi tab =0
            If cboTVStationCode.Text = "" Then Exit Sub
            Call DisableColorTab0(False)
            Call CmdEditTV_Click
        Case 1
        
            Call DisableControlTabRadio(False)
            If optArea.Value = True Then
                Call CmdEditRD_Click
            Else
                Call CmdEditRD2_Click
            End If
            
        Case 2
            Call DisableControlTabPrint(False)
            Call cmdEditPR_Click
        Case 3
            Call DisableControlTabCinema(False)
            If optCNDetail.Value = True Then
                Call cmdEditCN_Click
            Else
                Call cmdEditCN2_Click
            End If
        Case 4
            Call DisableControlTabOther(False)
            Call CmdEditOT_Click
       
       End Select
    
    Call EnableObject(True)

End Sub

Sub db_save()
'<CSCM>
'********************************************************************************
'Procedure Name     : db_save
'Procedure Function : Save Data
'Input Parameter    : ---
'Output Parameter   : ---
'Date               : 4/4/2016
'LastUpdate/By      : Tedi / Kreatif
'Name Before        : -
'********************************************************************************
'</CSCM>
    
    Select Case SSTabMedium.Tab
    
        Case 0 'Jika sedang diposisi tab =0
            If strMode = "Add" Then
                Call cmdAddTV_Click 'sesuai permintaan pak yayan
            Else
                Call CmdEditTV_Click
            End If
            
        Case 1
            If strMode = "Add" Then
                If optArea.Value = True Then
                    Call cmdAddRD_Click
                Else
                    Call cmdAddRD2_Click
                End If
            Else
                If optArea.Value = True Then
                    Call CmdEditRD_Click
                Else
                    Call CmdEditRD2_Click
                End If
            End If
            Call ViewMedium("RD", FGRDMedium)
        Case 2
            If strMode = "Add" Then
                Call cmdAddPR_Click
            Else
                Call cmdEditPR_Click
            End If
        Case 3
        
            If strMode = "Add" Then
                If optCNDetail.Value = True Then
                    Call cmdAddCN_Click
                Else
                    Call cmdAddCN2_Click
                End If
            Else
                If optCNDetail.Value = True Then
                    Call cmdEditCN_Click
                    Call FGCNMedium_Click
                Else
                    Call cmdEditCN2_Click
                    Call FGCNMedium2_Click
                End If
            End If
            
            Call DisableControlTabCinema(True)
            Call EmptyTabCinema
            Call EnableObject(False)
            
            strMode = ""
            If optCNDetail.Value = True Then
                If tdg_FGCNMedium.ApproxCount > 0 Then
                    If tdg_FGCNMedium.Row = 0 Then tdg_FGCNMedium.Row = 1
                    FGCNMedium_Click
                End If
            Else
                If tdg_FGCNMedium2.ApproxCount > 0 Then
                    'If FGCNMedium2.Row = 0 Then FGCNMedium2.Row = 1
                    FGCNMedium2_Click
                End If
            End If
        Case 4
            If strMode = "Add" Then
                Call cmdAddOT_Click 'sesuai permintaan pak yayan
            Else
                Call CmdEditOT_Click
            End If
            Call DisableControlTabOther(True)
            Call EmptyTabOther
            Call EnableObject(False)
            
    End Select
    If tdg_OTMedium.ApproxCount > 0 Then
        'If tdg_OTMedium.Row = 0 Then tdg_OTMedium.Row = 1
        FGOTMedium_Click
    End If
    intTabEdit = 9
    If strMode = "" Then
        intTabEdit = 9
    End If
End Sub

Sub db_delete()
'<CSCM>
'********************************************************************************
'Procedure Name     : db_Delete
'Procedure Function : Hapus Data
'Input Parameter    : ---
'Output Parameter   : ---
'Date               : 4/4/2016
'LastUpdate/By      : Tedi / Kreatif
'Name Before        : -
'********************************************************************************
'</CSCM>
    
    Me.MousePointer = vbHourglass
    Select Case SSTabMedium.Tab
    
        Case 0 'Jika sedang diposisi tab =0
            cmdDeleteTV_Click
            Call EmptyTab0
            If tdg_TVMedium.ApproxCount > 0 Then
                tdg_TVMedium.Row = 1
                Call ShowDetailTV
            End If
        Case 1
            If optArea.Value = True Then
                cmdDeleteRD_Click
                Call EmptyTabRadio
                FGRDMedium_Click
            Else
                CmdDeleteRD2_Click
                Call EmptyTabRadio
                FGRDMedium2_Click
            End If
        Case 2
            Call cmdDeletePR_Click
            If FGPRMedium.Rows > 0 Then
                If FGPRMedium.Row = 0 Then FGPRMedium.Row = 1
                Call FGPRMedium_Click
            End If
        Case 3
            If optCNBrief.Value = True Then
                CmdDeleteCN2_Click
            Else
                cmdDeleteCN_Click
            End If
            Call EmptyTabCinema
            If optCNBrief.Value = True Then
                tdg_FGCNMedium2.ApproxCount = 0
                If tdg_FGCNMedium2.ApproxCount > 0 Then
                    FGCNMedium2_Click
                End If
            Else
                tdg_FGCNMedium.Row = 0
                If tdg_FGCNMedium.ApproxCount > 0 Then
                    FGCNMedium_Click
                End If
            End If
        Case 2
            Call cmdDeleteCN_Click
            If tdg_FGCNMedium.ApproxCount > 0 Then
                If tdg_FGCNMedium.Row = 0 Then tdg_FGCNMedium.Row = 1
                Call FGCNMedium_Click
            End If
        Case 4
            cmdDeleteOT_Click
    End Select
    Me.MousePointer = vbNormal

End Sub

Sub db_cancel()
'<CSCM>
'********************************************************************************
'Procedure Name     : db_Cancel
'Procedure Function : Membatalkan Penambahan/Perubahan Data
'Input Parameter    : ---
'Output Parameter   : ---
'Date               : 4/4/2016
'LastUpdate/By      : Tedi / Kreatif
'Name Before        : -
'********************************************************************************
'</CSCM>
    
    Select Case SSTabMedium.Tab
    
        Case 0 'Jika sedang diposisi tab =0
            'FrameTV.Enabled = False
            If strMode <> "Add" Then
                cmdDeleteTV_Click
            End If
            Call DisableColorTab0(True)
            ShowDetailTV
        Case 1
            strMode = ""
            If optArea.Value = True Then
                Call CmdEditRD_Click
            Else
                Call CmdEditRD2_Click
            End If
            Call FGRDMedium_Click
            'Call EmptyTabRadio
            Call DisableControlTabRadio(True)
        Case 2
            If strMode = "Add" Then
                Call DisableControlTabPrint(True)
                Call EmptyTabPrint
                If FGPRMedium.Rows > 0 Then
                    If FGPRMedium.Row = 0 Then FGPRMedium.Row = 1
                    Call FGPRMedium_Click
                End If
            Else
                strMode = ""
                Call cmdEditPR_Click
            End If
        Case 3
            strMode = ""
            If optCNDetail.Value = True Then
                Call cmdEditCN_Click
            Else
                Call cmdEditCN2_Click
            End If
            
            Call DisableControlTabCinema(True)
            Call EmptyTabCinema
            strMode = ""
            If optCNBrief.Value = True Then
                If tdg_FGCNMedium.ApproxCount > 0 Then
                    FGCNMedium_Click
                End If
            Else
                If tdg_FGCNMedium2.ApproxCount > 0 Then
                    FGCNMedium2_Click
                End If
            End If
        Case 4
            Call DisableControlTabOther(True)
            Call EmptyTabOther
            If strMode = "Add" Then
            Else
                strMode = ""
                Call CmdEditOT_Click
            End If
            If tdg_OTMedium.ApproxCount > 0 Then
                If tdg_OTMedium.Row = 0 Then tdg_OTMedium.Row = 1
                Call FGOTMedium_Click
            End If
    End Select
    
    Call EnableObject(False)
    strMode = ""
    intTabEdit = 9
End Sub

Sub DisableColorTab0(ByVal blnEnable As Boolean)
'<CSCM>
'********************************************************************************
'Procedure Name     : DisableColorTab0
'Procedure Function : Disable/Enable  Color Tab 0
'Input Parameter    : blnEnable
'Output Parameter   : ---
'Date               : 4/4/2016
'LastUpdate/By      : Tedi / Kreatif
'Name Before        : -
'********************************************************************************
'</CSCM>
    cboTVStationCode.Enabled = Not blnEnable
    cboTVMarketName.Enabled = Not blnEnable
    cboTVSpotType.Enabled = Not blnEnable
    txtTVRateGross.Enabled = Not blnEnable
    txtTVRate.Enabled = Not blnEnable
    cboTVStationName.Enabled = Not blnEnable
    cboTVVersion.Enabled = Not blnEnable
    txtTVDuration.Enabled = False
    tdg_TVMedium.Enabled = Not blnEnable
        
    If blnEnable = True Then
        cboTVStationCode.BackColor = &HE0E0E0
        cboTVMarketName.BackColor = &HE0E0E0
        cboTVSpotType.BackColor = &HE0E0E0
        txtTVRateGross.BackColor = &HE0E0E0
        txtTVRate.BackColor = &HE0E0E0
        cboTVStationName.BackColor = &HE0E0E0
        cboTVVersion.BackColor = &HE0E0E0
        txtTVDuration.BackColor = &HE0E0E0
        tdg_TVMedium.Enabled = blnEnable
    Else
        cboTVStationCode.BackColor = &HFFFFFF
        cboTVMarketName.BackColor = &HFFFFFF
        cboTVSpotType.BackColor = &HFFFFFF
        txtTVRateGross.BackColor = &HFFFFFF
        txtTVRate.BackColor = &HFFFFFF
        cboTVStationName.BackColor = &HFFFFFF
        cboTVVersion.BackColor = &HFFFFFF
        txtTVDuration.BackColor = &HE0E0E0
        tdg_TVMedium.Enabled = blnEnable
    End If
    cboTVProgram.Visible = False
End Sub

Sub DisableControlTabRadio(ByVal blnValue As Boolean)
'<CSCM>
'********************************************************************************
'Procedure Name     : DisableControlTabRadio
'Procedure Function : Disable/Enable  Control
'Input Parameter    : blnEnable
'Output Parameter   : ---
'Date               : 4/4/2016
'LastUpdate/By      : Tedi / Kreatif
'Name Before        : -
'********************************************************************************
'</CSCM>

    txtRDStation.Locked = True
    CboRDSpotType.Enabled = Not blnValue
    CboRDSpotType2.Enabled = Not blnValue
    cboRDVersion.Enabled = Not blnValue
    cboRDVersion2.Enabled = Not blnValue
    txtRDDuration.Locked = True
    '------------------
    txtRDRPS.Locked = blnValue
    txtRDRPSGross.Locked = blnValue
    '---------------------
    optArea.Enabled = True
    OptStation.Enabled = True
    cboRDArea.Enabled = Not blnValue
    cboRDArea.Enabled = Not blnValue
    '---------------------
    txtRDStation.Locked = Not blnValue
    FGRDMedium2.Enabled = blnValue
    FGRDMedium.Enabled = blnValue
    '-----ByStasion
    txtRDRPS2.Enabled = Not blnValue
    txtRDRPSGross2.Enabled = Not blnValue
    CboRDSpotType2.Enabled = Not blnValue
    trvRDStationCatalog.Enabled = Not blnValue
    lstRDRateType.Enabled = Not blnValue
    cmdNewRadioMaterial2.Enabled = Not blnValue
    cmdRDNewArea.Enabled = Not blnValue
    cmdNewRadioMaterial.Enabled = Not blnValue
 
    If blnValue = True Then
        txtRDDuration.BackColor = &HE0E0E0
        txtRDRPS.BackColor = &HE0E0E0
        txtRDRPSGross.BackColor = &HE0E0E0
        CboRDSpotType.BackColor = &HE0E0E0
        CboRDSpotType2.BackColor = &HE0E0E0
        cboRDVersion.BackColor = &HE0E0E0
        cboRDVersion2.BackColor = &HE0E0E0
        cboRDArea.BackColor = &HE0E0E0
        txtRDStation.BackColor = &HE0E0E0
        '-----ByStasion
        txtRDRPS2.BackColor = &HE0E0E0
        txtRDRPSGross2.BackColor = &HE0E0E0
        CboRDSpotType2.BackColor = &HE0E0E0
    Else
        txtRDDuration.BackColor = &HE0E0E0
        txtRDRPS.BackColor = &HFFFFFF
        txtRDRPSGross.BackColor = &HFFFFFF
        CboRDSpotType.BackColor = &HFFFFFF
        CboRDSpotType2.BackColor = &HFFFFFF
        cboRDVersion.BackColor = &HFFFFFF
        cboRDVersion2.BackColor = &HFFFFFF
        cboRDArea.BackColor = &HFFFFFF
        txtRDStation.BackColor = &HE0E0E0
        '-----ByStasion
        txtRDRPS2.BackColor = &HFFFFFF
        txtRDRPSGross2.BackColor = &HFFFFFF
        CboRDSpotType2.BackColor = &HFFFFFF
    End If

End Sub

Sub EmptyTab0()
'<CSCM>
'********************************************************************************
'Procedure Name     : EmptyTab0
'Procedure Function : Mengosongkan Control  di Tab 0
'Input Parameter    : ---
'Output Parameter   : ---
'Date               : 4/4/2016
'LastUpdate/By      : Tedi / Kreatif
'Name Before        : -
'********************************************************************************
'</CSCM>

        cboTVStationCode.ListIndex = -1
        cboTVMarketName.ListIndex = -1
        cboTVSpotType.ListIndex = -1
        txtTVRateGross.Text = "0"
        txtTVRate.Text = "0"
        cboTVStationName.ListIndex = -1
        cboTVVersion.ListIndex = -1
        txtTVDuration.Text = "0"

End Sub

Sub EmptyTabRadio()
'<CSCM>
'********************************************************************************
'Procedure Name     : EmptyTabRadio
'Procedure Function : Mengosongkan Control  di Tab 1
'Input Parameter    : ---
'Output Parameter   : ---
'Date               : 4/4/2016
'LastUpdate/By      : Tedi / Kreatif
'Name Before        : -
'********************************************************************************
'</CSCM>

    txtRDStation.Text = "0"
    CboRDSpotType.ListIndex = -1
    cboRDVersion.ListIndex = -1
    txtRDDuration.Text = "0"
    txtRDRPS.Text = "0"
    txtRDRPSGross.Text = "0"
    lvRDSelectedStation.ListItems.Clear
    txtRDRPSGross2.Text = 0
    txtRDRPS2.Text = 0
    CboRDSpotType2.ListIndex = -1
    cboRDVersion2.ListIndex = -1
    txtRDDuration2.Text = 0
End Sub

Private Sub picButton_Click(Index As Integer)
'<CSCM>
'********************************************************************************
'Procedure Name     : picButton_Click
'Procedure Function : Action utk Navigation dan CRUD.
'Input Parameter    : ---
'Output Parameter   : ---
'Date               : 3/22/2016
'LastUpdate/By      : Abdi / Kreatif
'********************************************************************************
'</CSCM>
    
  '  Lock_MainForm True
    
    Select Case Index
            
        Case enButtonType.bieAdd  '4 'ADD.
             Call db_add
        Case enButtonType.bieEdit  '5 'EDIT.
             Call db_edit
        Case enButtonType.bieDelete  '6 'DELETE.
             Call db_delete
        Case enButtonType.bieClose  '7 'EXIT.
             Unload Me
        Case enButtonType.bieSave  'SAVE.
             Call db_save
        Case enButtonType.bieCancel 'CANCEL.
             Call db_cancel
    End Select

End Sub

Private Sub picButton_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single) 'TOOLBAR_AI.
'<CSCM>
'********************************************************************************
'Procedure Name     : picButton_MouseDown
'Procedure Function : TOOLBAR_AI saat mouse ditekan.
'Input Parameter    : ---
'Output Parameter   : ---
'Date               : 3/22/2016
'LastUpdate/By      : Abdi / Kreatif
'********************************************************************************
'</CSCM>

    picButton(Index).Picture = LoadPicture(SetButtonImageEffect(Index, bieDown)) 'FIRST.

End Sub

Private Sub picButton_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single) 'TOOLBAR_AI.
'<CSCM>
'********************************************************************************
'Procedure Name     : picButton_MouseMove
'Procedure Function : TOOLBAR_AI saat mouse berada di area button.
'Input Parameter    : ---
'Output Parameter   : ---
'Date               : 3/22/2016
'LastUpdate/By      : Abdi / Kreatif
'********************************************************************************
'</CSCM>

    picButton_Obj Index, X, Y, picButton
    
End Sub

Private Sub ShowDetailTV()
'<CSCM>
'********************************************************************************
'Procedure Name     : ShowDetailTV
'Procedure Function : Show Detail TV
'Input Parameter    : ---
'Output Parameter   : ---
'Date               : 4/4/2016
'LastUpdate/By      : Tedi / Kreatif
'Name Before        : -
'********************************************************************************
'</CSCM>

    Dim recTemp As New ADODB.Recordset
    Dim strSql As String

        'TAMPILKAN DATA DETAIL
        strSql = "select station_code,station_name,isnull(market_code,0) market_code,isnull(market_name,'NATIONAL') market_name from mp_medium_detail where mp_medium_detail_id = '" & tdg_TVMedium.Columns(1) & "'"
        recTemp.Open strSql, ConnERP, 1, 3
        If Not recTemp.EOF Then
            cboTVStationCode.Text = Trim(recTemp("station_code"))
            cboTVStationName.Text = recTemp("station_name")
            cboTVMarketName.Text = recTemp("market_name")
            cboTVMarketCode.ListIndex = cboTVMarketName.ListIndex
        End If
        recTemp.Close
        strSql = "select spot_type,version,duration,rate_per_spot,gross_rate_per_spot from mp_plan_dimension where mp_plan_dim_id = '" & tdg_TVMedium.Columns(0) & "'"
        recTemp.Open strSql, ConnERP, 1, 3
        If Not recTemp.EOF Then
            cboTVSpotType.Text = recTemp("spot_type")
            Call cboTVSpotType_Click
            If recTemp("spot_type") = "Reguler" Then
                cboTVVersion.Text = recTemp("version")
            Else
                cboTVProgram.Text = recTemp("version")
            End If
            txtTVDuration.Text = recTemp("duration")
            txtTVRate.Text = FormatNumber(recTemp("rate_per_spot"), 2)
            txtTVRateGross.Text = FormatNumber(recTemp("gross_rate_per_spot"), 2)
        End If
        Call CloseRecordset(recTemp)
    
End Sub

Sub showDetailRD()
'<CSCM>
'********************************************************************************
'Procedure Name     : showDetailRD
'Procedure Function : Show Detail RD
'Input Parameter    : ---
'Output Parameter   : ---
'Date               : 4/4/2016
'LastUpdate/By      : Tedi / Kreatif
'Name Before        : -
'********************************************************************************
'</CSCM>
        
    Dim strSql As String
    'View Data
    strSql = "select * from mp_medium_detail where mp_medium_detail_id='" & FGRDMedium.TextMatrix(FGRDMedium.Row, 2) & "'"
    recTemp.Open strSql, ConnERP, 1, 3
    
    If Not recTemp.EOF Then
       cboRDArea.Text = Trim(recTemp("area_code"))
    End If
    recTemp.Close
    
    strSql = "select * from mp_plan_dimension where mp_plan_dim_id='" & FGRDMedium.TextMatrix(FGRDMedium.Row, 1) & "'"
    recTemp.Open strSql, ConnERP, 1, 3
    If Not recTemp.EOF Then
       txtRDStation.Text = recTemp("rd_stations")
       CboRDSpotType = recTemp("spot_type")
       cboRDVersion = recTemp("version")
       txtRDDuration = recTemp("duration")
       txtRDRPS.Text = FormatNumber(recTemp("rate_per_spot"))
       txtRDRPSGross.Text = FormatNumber(recTemp("gross_rate_per_spot"))
    End If
    recTemp.Close

End Sub

Sub ShowDetailRDStasion()
'<CSCM>
'********************************************************************************
'Procedure Name     : ShowDetailRDStasion
'Procedure Function : Show Detail RD Stasion
'Input Parameter    : ---
'Output Parameter   : ---
'Date               : 4/4/2016
'LastUpdate/By      : Tedi / Kreatif
'Name Before        : -
'********************************************************************************
'</CSCM>
        
    Dim strSql As String
    Dim intCount As Integer
    strSql = "select * from mp_medium_detail where mp_medium_detail_id = '" & FGRDMedium2.TextMatrix(FGRDMedium2.Row, 2) & "'"
    recTemp.Open strSql, ConnERP, 1, 3
    lvRDSelectedStation.ListItems.Clear
    If Not recTemp.EOF Then
        lvRDSelectedStation.ListItems.Add , recTemp("radio_station_code"), recTemp("radio_station_name")
    End If
    recTemp.Close
    
    strSql = "select * from mp_plan_dimension where mp_plan_dim_id = '" & FGRDMedium2.TextMatrix(FGRDMedium2.Row, 1) & "'"
    recTemp.Open strSql, ConnERP, 1, 3
    If Not recTemp.EOF Then
        If recTemp("spot_type") = "Reguler" Then
            For intCount = 0 To lstRDRateType.ListCount - 1
                If Mid(lstRDRateType.List(intCount), 2, InStr(1, lstRDRateType.List(intCount), "]") - 2) = Trim(recTemp("rd_rate_type_code")) Then
                    lstRDRateType.Selected(intCount) = True
                    Exit For
                End If
            Next
        End If
        CboRDSpotType2.Text = recTemp("spot_type")
        cboRDVersion2.Text = recTemp("version")
        cboRDDuration2.Text = recTemp("duration")
        txtRDRPS2.Text = FormatNumber(recTemp("rate_per_spot"))
        txtRDRPSGross2.Text = FormatNumber(recTemp("gross_rate_per_spot"))
        Call CboRDSpotType2_Click
    End If
    recTemp.Close

End Sub

Sub EmptyTabPrint()
'<CSCM>
'********************************************************************************
'Procedure Name     : EmptyTabRadio
'Procedure Function : Mengosongkan Control  di Tab 1
'Input Parameter    : ---
'Output Parameter   : ---
'Date               : 4/4/2016
'LastUpdate/By      : Tedi / Kreatif
'Name Before        : -
'********************************************************************************
'</CSCM>

    TxtPRMediaCode.Text = ""
    txtPRMediaNameSearch.Text = ""
    TxtPRSize.Text = ""
    TxtPRSatuan.Text = ""
    cboPRVersion.ListIndex = -1
    cboPRSpotType.ListIndex = -1
    TxtPRColor.Text = ""
    TxtPRPaper.Text = ""
    TxtPRMinSize.Text = ""
    txtPRCol.Text = ""
    txtPRMM.Text = ""
    txtPRRateGross.Text = "0"
    txtPRRate.Text = "0"

End Sub

Sub showDetailPR()
'<CSCM>
'********************************************************************************
'Procedure Name     : EmptyTabRadio
'Procedure Function : Mengosongkan Control  di Tab 1
'Input Parameter    : ---
'Output Parameter   : ---
'Date               : 4/4/2016
'LastUpdate/By      : Tedi / Kreatif
'Name Before        : -
'********************************************************************************
'</CSCM>

    Dim strSql As String
    'TAMPLIKAN MEDIA PRINT
    strSql = "select * from mp_medium_detail where mp_medium_detail_id = '" & FGPRMedium.TextMatrix(FGPRMedium.Row, 2) & "'"
    recTemp.Open strSql, ConnERP, 1, 3
    If Not recTemp.EOF Then
        TxtPRMediaCode.Text = recTemp("print_code")
        txtPRMediaNameSearch.Text = recTemp("media_name")
        txtPRMediaName.Text = recTemp("media_name")
    End If
    recTemp.Close
    txtPRMediaNameSearch.Locked = True
    strSql = "select * from mp_plan_dimension where mp_plan_dim_id = '" & FGPRMedium.TextMatrix(FGPRMedium.Row, 1) & "'"
    recTemp.Open strSql, ConnERP, 1, 3
    
    If Not recTemp.EOF Then
        TxtPRSatuan.Text = recTemp("print_size_code")
        TxtPRColor.Text = recTemp("print_color_code")
        TxtPRPaper.Text = recTemp("print_paper_code")
        TxtPRMinSize.Text = recTemp("print_min_size")
        txtPRCol.Text = recTemp("print_mmc_col")
        txtPRMM.Text = recTemp("print_mmc_size")
        txtPRIsMMC.Text = recTemp("print_ismmc")
        cboPRSpotType.Text = recTemp("spot_type")
        txtPRRate.Text = FormatNumber(recTemp("rate_per_spot"), 2)
        txtPRRateGross.Text = FormatNumber(recTemp("gross_rate_per_spot"), 2)
        cboPRVersion.Text = recTemp("version")
        Call cboPRSpotType_Click
        cboPRSpotType.Enabled = False
    End If
    recTemp.Close

End Sub

Sub DisableControlTabPrint(ByVal blnValue As Boolean)
'<CSCM>
'********************************************************************************
'Procedure Name     : DisableControlTabRadio
'Procedure Function : Disable/Enable  Control
'Input Parameter    : blnEnable
'Output Parameter   : ---
'Date               : 4/4/2016
'LastUpdate/By      : Tedi / Kreatif
'Name Before        : -
'********************************************************************************
'</CSCM>

    TxtPRMediaCode.Enabled = Not blnValue
    txtPRMediaNameSearch.Enabled = Not blnValue
    txtPRMediaNameSearch.Locked = blnValue
    TxtPRSize.Enabled = Not blnValue
    TxtPRSatuan.Enabled = Not blnValue
    cboPRVersion.Enabled = Not blnValue
    cboPRSpotType.Enabled = Not blnValue
    TxtPRColor.Enabled = Not blnValue
    TxtPRPaper.Enabled = Not blnValue
    TxtPRMinSize.Enabled = Not blnValue
    txtPRCol.Enabled = Not blnValue
    txtPRMM.Enabled = Not blnValue
    txtPRRateGross.Enabled = Not blnValue
    txtPRRate.Enabled = Not blnValue
    
    If blnValue = True Then
        TxtPRMediaCode.BackColor = &HE0E0E0
        txtPRMediaNameSearch.BackColor = &HE0E0E0
        TxtPRSize.BackColor = &HE0E0E0
        TxtPRSatuan.BackColor = &HE0E0E0
        cboPRVersion.BackColor = &HE0E0E0
        cboPRSpotType.BackColor = &HE0E0E0
        TxtPRColor.BackColor = &HE0E0E0
        TxtPRPaper.BackColor = &HE0E0E0
        TxtPRMinSize.BackColor = &HE0E0E0
        txtPRCol.BackColor = &HE0E0E0
        txtPRMM.BackColor = &HE0E0E0
        txtPRRateGross.BackColor = &HE0E0E0
        txtPRRate.BackColor = &HE0E0E0
    Else
        TxtPRMediaCode.BackColor = &HFFFFFF
        txtPRMediaNameSearch.BackColor = &HFFFFFF
        TxtPRSize.BackColor = &HFFFFFF
        TxtPRSatuan.BackColor = &HFFFFFF
        cboPRVersion.BackColor = &HFFFFFF
        cboPRSpotType.BackColor = &HFFFFFF
        TxtPRColor.BackColor = &HFFFFFF
        TxtPRPaper.BackColor = &HFFFFFF
        TxtPRMinSize.BackColor = &HFFFFFF
        txtPRCol.BackColor = &HFFFFFF
        txtPRMM.BackColor = &HFFFFFF
        txtPRRateGross.BackColor = &HFFFFFF
        txtPRRate.BackColor = &HFFFFFF
    End If

End Sub

Sub DisableControlTabCinema(ByVal blnValue As Boolean)
'<CSCM>
'********************************************************************************
'Procedure Name     : DisableControlTabRadio
'Procedure Function : Disable/Enable  Control
'Input Parameter    : blnEnable
'Output Parameter   : ---
'Date               : 4/4/2016
'LastUpdate/By      : Tedi / Kreatif
'Name Before        : -
'********************************************************************************
'</CSCM>
   
    'Brief
    cboCNName.Enabled = Not blnValue
    cboCNStudio.Enabled = Not blnValue
    txtCNRateGross.Enabled = Not blnValue
    txtCNRate.Enabled = Not blnValue
    cboCNSpotType.Enabled = Not blnValue
    cboCNVersion.Enabled = Not blnValue
    txtCNDuration.Enabled = Not blnValue
    txtCNDuration.Enabled = Not blnValue
    cboCNJenisDurasi.Enabled = Not blnValue
    tdg_FGCNMedium.Enabled = blnValue
    'Detail
    txtCNBrief.Enabled = Not blnValue
    txtCNRateGross2.Enabled = Not blnValue
    txtCNRate2.Enabled = Not blnValue
    tdg_FGCNMedium2.Enabled = blnValue
    
    If blnValue = True Then

        'Brief
        cboCNName.BackColor = &HE0E0E0
        cboCNStudio.BackColor = &HE0E0E0
        txtCNRateGross.BackColor = &HE0E0E0
        txtCNRate.BackColor = &HE0E0E0
        cboCNSpotType.BackColor = &HE0E0E0
        cboCNVersion.BackColor = &HE0E0E0
        txtCNDuration.BackColor = &HE0E0E0
        txtCNDuration.BackColor = &HE0E0E0
        cboCNJenisDurasi.BackColor = &HE0E0E0
        'Detail
        txtCNBrief.BackColor = &HE0E0E0
        txtCNRateGross2.BackColor = &HE0E0E0
        txtCNRate2.BackColor = &HE0E0E0
 
    Else

        'Brief
        cboCNName.BackColor = &HFFFFFF
        cboCNStudio.BackColor = &HFFFFFF
        txtCNRateGross.BackColor = &HFFFFFF
        txtCNRate.BackColor = &HFFFFFF
        cboCNSpotType.BackColor = &HFFFFFF
        cboCNVersion.BackColor = &HFFFFFF
        txtCNDuration.BackColor = &HFFFFFF
        txtCNDuration.BackColor = &HFFFFFF
        cboCNJenisDurasi.BackColor = &HFFFFFF
        'Detail
        txtCNBrief.BackColor = &HFFFFFF
        txtCNRateGross2.BackColor = &HFFFFFF
        txtCNRate2.BackColor = &HFFFFFF

    End If

End Sub

Sub EmptyTabCinema()
'<CSCM>
'********************************************************************************
'Procedure Name     : EmptyTabCinema
'Procedure Function : Mengosongkan Control  di Tab Cinema
'Input Parameter    : ---
'Output Parameter   : ---
'Date               : 4/4/2016
'LastUpdate/By      : Tedi / Kreatif
'Name Before        : -
'********************************************************************************
'</CSCM>
   
    cboCNName.ListIndex = -1
    cboCNStudio.ListIndex = -1
    txtCNRateGross.Text = "0"
    txtCNRate.Text = "0"
    cboCNSpotType.ListIndex = -1
    cboCNVersion.ListIndex = -1
    txtCNDuration.Text = "0"
     
    txtCNBrief.Text = ""
    txtCNRateGross2.Text = "0"
    txtCNRate2.Text = "0"
        
End Sub

Sub showDetailCinema()

    Dim recTemp As New ADODB.Recordset
    Dim strSql As String
    
    cboCNName.Enabled = True
    cboCNStudio.Enabled = True
    cboCNSpotType.Enabled = True
    cboCNVersion.Enabled = True

    
    strSql = "select * from mp_medium_detail where mp_medium_detail_id = '" & tdg_FGCNMedium.Columns(1) & "'"
    recTemp.Open strSql, ConnERP, 1, 3
    If Not recTemp.EOF Then
       cboCNCode.Text = Trim(recTemp("cinema_code"))
       cboCNCode.Enabled = False
       cboCNName.Text = Trim(recTemp("cinema_name"))
       cboCNName.Enabled = False
       cboCNStudio.Text = Trim(recTemp("cinema_studio"))
       cboCNStudio.Enabled = False
    End If
    recTemp.Close
    
    strSql = "select * from mp_plan_dimension where mp_plan_dim_id = '" & tdg_FGCNMedium.Columns(0) & "'"
    recTemp.Open strSql, ConnERP, 1, 3
    If Not recTemp.EOF Then
       cboCNSpotType.Text = recTemp("spot_type")
       Call cboCNSpotType_Click
       cboCNSpotType.Enabled = False
       cboCNVersion.Text = recTemp("version")
       cboCNJenisDurasi = recTemp("cinema_duration")
       txtCNDuration.Text = recTemp("duration")
       txtCNRate.Text = FormatNumber(recTemp("rate_per_spot"), 2)
       txtCNRateGross.Text = FormatNumber(recTemp("gross_rate_per_spot"), 2)
    End If
    Call CloseRecordset(recTemp)

    cboCNName.Enabled = False
    cboCNStudio.Enabled = False
    cboCNSpotType.Enabled = False
    cboCNVersion.Enabled = False
    
End Sub

Sub EmptyTabOther()
'<CSCM>
'********************************************************************************
'Procedure Name     : EmptyTabOther
'Procedure Function : Mengosongkan Control  di Tab Other
'Input Parameter    : ---
'Output Parameter   : ---
'Date               : 4/4/2016
'LastUpdate/By      : Tedi / Kreatif
'Name Before        : -
'********************************************************************************
'</CSCM>

    txtOTDescrition.Text = ""
        
End Sub

Sub DisableControlTabOther(ByVal blnValue As Boolean)
'<CSCM>
'********************************************************************************
'Procedure Name     : DisableControlTabOther
'Procedure Function : Disable/Enable  Control
'Input Parameter    : blnEnable
'Output Parameter   : ---
'Date               : 4/4/2016
'LastUpdate/By      : Tedi / Kreatif
'Name Before        : -
'********************************************************************************
'</CSCM>
   
    'Brief
    txtOTDescrition.Enabled = Not blnValue
    
    
    If blnValue = True Then
        txtOTDescrition.BackColor = &HE0E0E0
    Else
        txtOTDescrition.BackColor = &HFFFFFF
    End If

End Sub

