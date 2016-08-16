VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Object = "{562E3E04-2C31-4ECE-83F4-4017EEE51D40}#8.0#0"; "todg8.ocx"
Begin VB.Form frm_MPEdit 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Edit Media Plan"
   ClientHeight    =   7200
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9900
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7200
   ScaleWidth      =   9900
   StartUpPosition =   1  'CenterOwner
   Begin Threed.SSPanel pnlMain 
      Align           =   1  'Align Top
      Height          =   6555
      Left            =   0
      TabIndex        =   3
      Top             =   750
      Width           =   9900
      _Version        =   65536
      _ExtentX        =   17462
      _ExtentY        =   11562
      _StockProps     =   15
      Caption         =   "SSPanel1"
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
      Begin VB.Frame Frame1 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1305
         Left            =   90
         TabIndex        =   16
         Top             =   30
         Width           =   9705
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
            Height          =   285
            Left            =   1560
            Locked          =   -1  'True
            TabIndex        =   19
            Text            =   "txtClientName"
            Top             =   930
            Width           =   8025
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
            Height          =   285
            Left            =   1560
            Locked          =   -1  'True
            TabIndex        =   18
            Text            =   "txtBrandName"
            Top             =   585
            Width           =   8025
         End
         Begin VB.ComboBox cboMPNum 
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
            Left            =   1560
            Sorted          =   -1  'True
            Style           =   2  'Dropdown List
            TabIndex        =   17
            Top             =   210
            Width           =   2175
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Client Name  "
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
            Left            =   195
            TabIndex        =   22
            Top             =   930
            Width           =   945
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Brand Name  "
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
            Left            =   195
            TabIndex        =   21
            Top             =   600
            Width           =   960
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "MP Number  "
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
            Left            =   195
            TabIndex        =   20
            Top             =   240
            Width           =   900
         End
      End
      Begin VB.Frame Frame2 
         Caption         =   "Task"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2070
         Left            =   105
         TabIndex        =   11
         Top             =   1350
         Width           =   9705
         Begin VB.PictureBox Picture3 
            Height          =   360
            Left            =   6975
            ScaleHeight     =   300
            ScaleWidth      =   2565
            TabIndex        =   12
            TabStop         =   0   'False
            Top             =   1560
            Width           =   2625
            Begin VB.CommandButton cmdDelTask 
               Caption         =   "Delete"
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
               TabIndex        =   15
               ToolTipText     =   "delete task"
               Top             =   15
               Width           =   855
            End
            Begin VB.CommandButton cmdEditTask 
               Caption         =   "Edit"
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
               TabIndex        =   14
               ToolTipText     =   "edit task"
               Top             =   15
               Width           =   855
            End
            Begin VB.CommandButton cmdAddTask 
               Caption         =   "Add"
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
               Left            =   -15
               TabIndex        =   13
               ToolTipText     =   "add task"
               Top             =   15
               Width           =   855
            End
         End
         Begin TrueOleDBGrid80.TDBGrid tdg_Task 
            Height          =   1230
            Left            =   150
            TabIndex        =   24
            Top             =   240
            Width           =   9450
            _ExtentX        =   16669
            _ExtentY        =   2170
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
      End
      Begin VB.Frame Frame3 
         Caption         =   "Activity"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2970
         Left            =   90
         TabIndex        =   4
         Top             =   3420
         Width           =   9705
         Begin VB.PictureBox Picture5 
            Height          =   360
            Left            =   6045
            ScaleHeight     =   300
            ScaleWidth      =   855
            TabIndex        =   9
            TabStop         =   0   'False
            Top             =   2505
            Width           =   915
            Begin VB.CommandButton cmdShowDetail 
               Caption         =   "Detail"
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
               TabIndex        =   10
               ToolTipText     =   "view detail for current activity"
               Top             =   0
               Width           =   855
            End
         End
         Begin VB.PictureBox Picture1 
            Height          =   360
            Left            =   6990
            ScaleHeight     =   300
            ScaleWidth      =   2565
            TabIndex        =   5
            TabStop         =   0   'False
            Top             =   2490
            Width           =   2625
            Begin VB.CommandButton cmdDelActivity 
               Caption         =   "Delete"
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
               TabIndex        =   8
               ToolTipText     =   "delete activity"
               Top             =   0
               Width           =   855
            End
            Begin VB.CommandButton cmdEditActivity 
               Caption         =   "Edit"
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
               TabIndex        =   7
               ToolTipText     =   "edit activity"
               Top             =   0
               Width           =   855
            End
            Begin VB.CommandButton cmdAddActivity 
               Caption         =   "Add"
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
               TabIndex        =   6
               ToolTipText     =   "add activity"
               Top             =   0
               Width           =   855
            End
         End
         Begin TrueOleDBGrid80.TDBGrid tdg_Activity 
            Height          =   2130
            Left            =   105
            TabIndex        =   25
            Top             =   270
            Width           =   9495
            _ExtentX        =   16748
            _ExtentY        =   3757
            _LayoutType     =   4
            _RowHeight      =   -2147483647
            _WasPersistedAsPixels=   0
            Columns(0)._VlistStyle=   0
            Columns(0)._MaxComboItems=   5
            Columns(0).DataField=   ""
            Columns(0)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(1)._VlistStyle=   0
            Columns(1)._MaxComboItems=   5
            Columns(1).DataField=   ""
            Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns.Count   =   2
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
            Splits(0)._ColumnProps(0)=   "Columns.Count=2"
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
            Splits(0)._ColumnProps(12)=   "Column(1).Order=2"
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
            _StyleDefs(40)  =   "Named:id=33:Normal"
            _StyleDefs(41)  =   ":id=33,.parent=0"
            _StyleDefs(42)  =   "Named:id=34:Heading"
            _StyleDefs(43)  =   ":id=34,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
            _StyleDefs(44)  =   ":id=34,.wraptext=-1"
            _StyleDefs(45)  =   "Named:id=35:Footing"
            _StyleDefs(46)  =   ":id=35,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
            _StyleDefs(47)  =   "Named:id=36:Selected"
            _StyleDefs(48)  =   ":id=36,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
            _StyleDefs(49)  =   "Named:id=37:Caption"
            _StyleDefs(50)  =   ":id=37,.parent=34,.alignment=2"
            _StyleDefs(51)  =   "Named:id=38:HighlightRow"
            _StyleDefs(52)  =   ":id=38,.parent=33,.bgcolor=&HFF0000&,.fgcolor=&H8000000E&,.borderColor=&HFF2B2B&"
            _StyleDefs(53)  =   "Named:id=39:EvenRow"
            _StyleDefs(54)  =   ":id=39,.parent=33,.bgcolor=&HFFFF00&"
            _StyleDefs(55)  =   "Named:id=40:OddRow"
            _StyleDefs(56)  =   ":id=40,.parent=33"
            _StyleDefs(57)  =   "Named:id=41:RecordSelector"
            _StyleDefs(58)  =   ":id=41,.parent=34"
            _StyleDefs(59)  =   "Named:id=42:FilterBar"
            _StyleDefs(60)  =   ":id=42,.parent=33,.fgcolor=&H80000005&"
         End
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
      ScaleWidth      =   9900
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   0
      Width           =   9900
      Begin VB.PictureBox picButton 
         Appearance      =   0  'Flat
         BackColor       =   &H00F0F0F0&
         BorderStyle     =   0  'None
         ClipControls    =   0   'False
         ForeColor       =   &H80000008&
         Height          =   750
         Index           =   23
         Left            =   3150
         ScaleHeight     =   750
         ScaleWidth      =   1500
         TabIndex        =   23
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
         Left            =   1620
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
         Index           =   4
         Left            =   90
         ScaleHeight     =   750
         ScaleWidth      =   1500
         TabIndex        =   1
         TabStop         =   0   'False
         Top             =   0
         Width           =   1500
      End
   End
End
Attribute VB_Name = "frm_MPEdit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
''*****************************************************************************
' Nama Submodul         :  Frm_MPEdit
' Nama Programmer       :  Sistyo
' Tgl Pembuatan         :  9 Agustus 2004
' Last Update           :  12 Agustus 2004/Sistyo
'******************************************************************************
Dim rsTemp As New ADODB.Recordset
Dim recTask As New ADODB.Recordset
Dim recActivity As New ADODB.Recordset
Dim rec_TempActivity As New ADODB.Recordset
Dim strSql As String
Option Explicit

Private Sub Form_Load()
    '<CSCM>
    '********************************************************************************
    'Procedure Name     : Form_Load
    'Procedure Function : ---
    'Input Parameter    : ---
    'Output Parameter   : ---
    'Date               : 3/29/2016
    'LastUpdate/By      : Tedi / Kreatif
    'Name Before        : -
    '********************************************************************************
    '</CSCM>
    Call initform
    EnableObject False
    create_table_TempActivity
    tdg_Task_Click
    tdg_Activity_Click
End Sub

Private Sub initform()
    '<CSCM>
    '********************************************************************************
    'Procedure Name     : initform
    'Procedure Function : ---
    'Input Parameter    : ---
    'Output Parameter   : ---
    'Date               : 3/29/2016
    'LastUpdate/By      : Tedi / Kreatif
    'Name Before        : -
    '********************************************************************************
    '</CSCM>
    '*****************************************************************************
    ' Nama Prosedur     :   initform
    ' Fungsi Prosedur   :   Inisialisasi Form
    ' Parameter  Input  :
    ' Parameter Output  :
    ' Tgl Pembuatan     :   09 Agustus 2004
    ' Last Update/By    :   09 Agustus 2004/Sistyo
    '*****************************************************************************
    Dim counter As Integer
    
    txtBrandName.Text = ""
    txtClientName.Text = ""
    
    Call loadMPNumber
    
    If frm_MPInsertion.cboMPNumber.Text <> "" Then
        cboMPNum.Text = frm_MPInsertion.cboMPNumber.Text
    End If

End Sub

Private Sub loadMPNumber()
    '<CSCM>
    '********************************************************************************
    'Procedure Name     : loadMPNumber
    'Procedure Function : ---
    'Input Parameter    : ---
    'Output Parameter   : ---
    'Date               : 3/29/2016
    'LastUpdate/By      : Tedi / Kreatif
    'Name Before        : -
    '********************************************************************************
    '</CSCM>
    '*****************************************************************************
    ' Nama Prosedur     :   loadmpnumber
    ' Fungsi Prosedur   :   Load Media Plan(s) Number
    ' Parameter  Input  :
    ' Parameter Output  :
    ' Tgl Pembuatan     :   09 Agustus 2004
    ' Last Update/By    :   09 Agustus 2004/Sistyo
    '*****************************************************************************
    Dim strBrand_Filter As String
    strBrand_Filter = "select brand_code from media_security_catalog where user_name = '" & strLogin_User & "' and position = 'Planner' and valid_until>=(select getdate())"
    rsTemp.Open "select mp_number from mp_master where approval is null and brand_code in (" & strBrand_Filter & ") and is_latest = 1 order by mp_number", ConnERP, 1, 3
    While Not rsTemp.EOF
        cboMPNum.AddItem rsTemp(0)
        rsTemp.MoveNext
    Wend
    rsTemp.Close

End Sub

Private Sub Get_Description_MP()
    '<CSCM>
    '********************************************************************************
    'Procedure Name     : Get_Description_MP
    'Procedure Function : ---
    'Input Parameter    : ---
    'Output Parameter   : ---
    'Date               : 3/29/2016
    'LastUpdate/By      : Tedi / Kreatif
    'Name Before        : -
    '********************************************************************************
    '</CSCM>
    '********************************************************************************
    'Procedure Name     : GetDescriptionMP
    'Procedure Function : Mendapatkan Nilai [brand_name] dan [client_name]
    'Input Parameter    : ---
    'Output Parameter   : ---
    'Date               : 3/20/2016
    'LastUpdate/By      : Tedi / Kreatif
    'Name Before        : -
    '********************************************************************************
    
    On Error GoTo errGet

    strSql = "SELECT "
    strSql = strSql & "brand_name,"
    strSql = strSql & "client_name "
    strSql = strSql & "FROM [MP_Master] "
    strSql = strSql & "WHERE mp_number = '" & cboMPNum.Text & "'"
    rsTemp.Open strSql, ConnERP, 1, 3
    
    txtBrandName.Text = rsTemp!Brand_Name
    txtClientName.Text = rsTemp!Client_Name
    
    Call CloseRecordset(rsTemp)
    Exit Sub

errGet:
    MsgBox "No Brand_Name or Client Name!", vbExclamation, strApplication_Name
            
End Sub

Private Sub ViewMediaPlan()
    '<CSCM>
    '********************************************************************************
    'Procedure Name     : ViewMediaPlan
    'Procedure Function : ---
    'Input Parameter    : ---
    'Output Parameter   : ---
    'Date               : 3/29/2016
    'LastUpdate/By      : Tedi / Kreatif
    'Name Before        : -
    '********************************************************************************
    '</CSCM>
    '*****************************************************************************
    ' Nama Prosedur     :   ViewMediaPlan
    ' Fungsi Prosedur   :   Tampilkan Media Plan yang terpilih
    ' Parameter  Input  :
    ' Parameter Output  :
    ' Tgl Pembuatan     :   09 Agustus 2004
    ' Last Update/By    :   09 Agustus 2004/Sistyo
    '*****************************************************************************
    
    Dim counter As Integer
    
    CloseRecordset recTask
    recTask.CursorLocation = adUseClient
    strSql = ""
    strSql = strSql & "SELECT ROW_NUMBER() OVER(ORDER BY mp_task_id DESC) AS No,"
    strSql = strSql & "mp_task_id, "
    strSql = strSql & "task_desc "
    strSql = strSql & "FROM mp_task "
    strSql = strSql & "WHERE mp_number = '" & cboMPNum.Text & "'"
    recTask.Open strSql, ConnERP, adOpenDynamic, adLockOptimistic

    Set tdg_Task.DataSource = Nothing
    Set tdg_Task.DataSource = recTask

End Sub

'Private Sub ViewMediaPlan1()
'    '<CSCM>
'    '********************************************************************************
'    'Procedure Name     : ViewMediaPlan1
'    'Procedure Function : ---
'    'Input Parameter    : ---
'    'Output Parameter   : ---
'    'Date               : 3/29/2016
'    'LastUpdate/By      : Tedi / Kreatif
'    'Name Before        : -
'    '********************************************************************************
'    '</CSCM>
'    '*****************************************************************************
'    ' Nama Prosedur     :   ViewMediaPlan
'    ' Fungsi Prosedur   :   Tampilkan Media Plan yang terpilih
'    ' Parameter  Input  :
'    ' Parameter Output  :
'    ' Tgl Pembuatan     :   09 Agustus 2004
'    ' Last Update/By    :   09 Agustus 2004/Sistyo
'    '*****************************************************************************
'    Dim counter As Integer
'    rsTemp.Open "select brand_name, client_name from mp_master where mp_number = '" & cboMPNum.Text & "'", ConnERP, 1, 3
'    If Not rsTemp.EOF Then
'        txtBrandName.Text = rsTemp(0)
'        txtClientName.Text = rsTemp(1)
'    End If
'    rsTemp.Close
'
'    rsTemp.Open "select mp_task_id, task_desc from mp_task where mp_number = '" & cboMPNum.Text & "'", ConnERP, 1, 3
'    FGTask.Rows = 2
'    FGTask.TextMatrix(1, 0) = ""
'    FGTask.TextMatrix(1, 1) = ""
'    FGTask.TextMatrix(1, 2) = ""
'
'    counter = 1
'    While Not rsTemp.EOF
'        FGTask.Rows = counter + 1
'        FGTask.TextMatrix(counter, 0) = counter
'        FGTask.TextMatrix(counter, 1) = rsTemp(0)
'        FGTask.TextMatrix(counter, 2) = rsTemp(1)
'        rsTemp.MoveNext
'        counter = counter + 1
'    Wend
'    rsTemp.Close
'
'    'ROW_NUMBER() OVER(ORDER BY brand_name DESC) AS Row
'End Sub

Public Sub viewActivity()

    '<CSCM>
    '********************************************************************************
    'Procedure Name     : viewActivity
    'Procedure Function : ---
    'Input Parameter    : ---
    'Output Parameter   : ---
    'Date               : 3/29/2016
    'LastUpdate/By      : Tedi / Kreatif
    'Name Before        : -
    '********************************************************************************
    '</CSCM>
    '*****************************************************************************
    ' Nama Prosedur     :   ViewActivity
    ' Fungsi Prosedur   :   Tampilkan Activity dari Task yang terpilih
    ' Parameter  Input  :
    ' Parameter Output  :
    ' Tgl Pembuatan     :   10 Agustus 2004
    ' Last Update/By    :   10 Agustus 2004/Sistyo
    '*****************************************************************************
    Dim strSql As String
    Dim strMedium As String
    Dim counter As Integer
    Dim rsMedium As New ADODB.Recordset
        
    Call create_table_TempActivity
    
    strSql = "select "
    strSql = strSql & "mp_activity_id,"
    strSql = strSql & "mp_task_id,"
    strSql = strSql & "activity_type,"
    strSql = strSql & "activity_desc,"
    strSql = strSql & "brand_variant_name,"
    strSql = strSql & "target_audience,"
    strSql = strSql & "brand_target "
    strSql = strSql & "from mp_activity where mp_task_id = '" & tdg_Task.Columns(1) & "'"
    
    rsTemp.Open strSql, ConnERP, 1, 3
    counter = 0
    While Not rsTemp.EOF
        counter = counter + 1
        rec_TempActivity.AddNew
        rec_TempActivity(0) = counter
        rec_TempActivity(1) = rsTemp(0)
        rec_TempActivity(2) = rsTemp(1)
        rec_TempActivity(3) = rsTemp(2)
        rec_TempActivity(4) = rsTemp(3)
        rec_TempActivity(5) = rsTemp(4)
        rec_TempActivity(6) = rsTemp(5)
        rec_TempActivity(7) = rsTemp(6)
        
        strMedium = ""
        strSql = "SELECT "
        strSql = strSql & "Medium_name "
        strSql = strSql & "FROM mp_medium "
        strSql = strSql & "WHERE mp_activity_id='" & rsTemp(0) & "' "
        strSql = strSql & "ORDER BY medium_name desc"
        
        rsMedium.Open strSql, ConnERP, 1, 3
        
        While Not rsMedium.EOF
            strMedium = strMedium & ", " & rsMedium(0)
            rsMedium.MoveNext
        Wend
        rsMedium.Close
        If Len(strMedium) <> 0 Then
            strMedium = Right(strMedium, Len(strMedium) - 2)
        End If
        rec_TempActivity(8) = strMedium
        rsTemp.MoveNext
    Wend
    rsTemp.Close
    
    Set rsMedium = Nothing
    
    Set tdg_Activity.DataSource = Nothing
    tdg_Activity.ClearFields
    Set tdg_Activity.DataSource = rec_TempActivity
    tdg_Activity.Columns(1).Visible = False
    tdg_Activity.Columns(2).Visible = False
    
    tdg_Activity.Columns(0).Width = 400
    tdg_Activity.Columns(3).Width = 1400
    tdg_Activity.Columns(4).Width = 2500
    tdg_Activity.Columns(5).Width = 1500
    tdg_Activity.Columns(6).Width = 1500
    tdg_Activity.Columns(7).Width = 1500
    tdg_Activity.Columns(8).Width = 1500
    
    tdg_Activity.Columns(3).Caption = "Type"
    tdg_Activity.Columns(4).Caption = "Description"
    tdg_Activity.Columns(5).Caption = "Brand Variant"
    tdg_Activity.Columns(6).Caption = "Target Audence"
    tdg_Activity.Columns(7).Caption = "Brand Target"
    tdg_Activity.Columns(8).Caption = "Medium"

End Sub

Private Sub cboMPNum_click()
    '<CSCM>
    '********************************************************************************
    'Procedure Name     : cboMPNum_click
    'Procedure Function : ---
    'Input Parameter    : ---
    'Output Parameter   : ---
    'Date               : 3/29/2016
    'LastUpdate/By      : Tedi / Kreatif
    'Name Before        : -
    '********************************************************************************
    '</CSCM>

    Call ViewMediaPlan
    
    'Setting Tombol
        
    cmdAddActivity.Enabled = False
    cmdEditActivity.Enabled = False
    cmdDelActivity.Enabled = False
    cmdShowDetail.Enabled = False
    cmdAddTask.Enabled = True
    cmdEditTask.Enabled = False
    cmdDelTask.Enabled = False
    Call Get_Description_MP
    Call tdg_Task_Click

End Sub

Private Sub DelActivity()
    '<CSCM>
    '********************************************************************************
    'Procedure Name     : DelActivity
    'Procedure Function : ---
    'Input Parameter    : ---
    'Output Parameter   : ---
    'Date               : 3/29/2016
    'LastUpdate/By      : Tedi / Kreatif
    'Name Before        : -
    '********************************************************************************
    '</CSCM>
    '*****************************************************************************
    ' Nama Prosedur     :   DelActivity
    ' Fungsi Prosedur   :   Delete Activity
    ' Tgl Pembuatan     :   09 Agustus 2004
    ' Last Update/By    :   09 Agustus 2004/Sistyo
    '*****************************************************************************
    Dim strSql As String
    Dim intRow, intCol As Integer
    Dim pesan
    
    strSql = "delete from mp_activity where mp_activity_id = '" & tdg_Activity.Columns(1) & "'"
    ConnERP.Execute strSql
    'update mp_master
    ConnERP.Execute "update mp_master set last_update_by='" & strLogin_User & "',last_update_date = getdate() where mp_number='" & cboMPNum.Text & "'"
    'Update FlexGrid
'    With tdg_Activity
'        If tdg_Activity.ApproxCount > 2 Then
'            For intRow = tdg_Activity.Row To tdg_Activity.ApproxCount - 2
'                .TextMatrix(intRow, 0) = intRow
'                .TextMatrix(intRow, 1) = .Columns(1)
'                .TextMatrix(intRow, 2) = .Columns(1)
'                .TextMatrix(intRow, 3) = .Columns(1)
'                .TextMatrix(intRow, 4) = .Columns(1)
'                .TextMatrix(intRow, 5) = .Columns(1)
'                .TextMatrix(intRow, 6) = .Columns(1)
'                .TextMatrix(intRow, 7) = .TextMatrix(intRow + 1, 7)
'            Next
'            .Rows = .Rows - 1
'        Else
'            .TextMatrix(1, 0) = ""
'            .TextMatrix(1, 1) = ""
'            .TextMatrix(1, 2) = ""
'            .TextMatrix(1, 3) = ""
'            .TextMatrix(1, 4) = ""
'            .TextMatrix(1, 5) = ""
'            .TextMatrix(1, 6) = ""
'            .TextMatrix(1, 7) = ""
'            cmdEditActivity.Enabled = False
'            cmdDelActivity.Enabled = False
'            cmdShowDetail.Enabled = False
'        End If
'        .SetFocus
'    End With
    pesan = MsgBox("Activity Deleted!", vbExclamation, strApplication_Name)

End Sub

Private Sub fgTask_Click()
    '<CSCM>
    '********************************************************************************
    'Procedure Name     : fgTask_Click
    'Procedure Function : ---
    'Input Parameter    : ---
    'Output Parameter   : ---
    'Date               : 3/29/2016
    'LastUpdate/By      : Tedi / Kreatif
    'Name Before        : -
    '********************************************************************************
    '</CSCM>
    cmdEditActivity.Enabled = False
    cmdDelActivity.Enabled = False
    cmdAddActivity.Enabled = False
    cmdShowDetail.Enabled = False
    cmdEditTask.Enabled = False
    cmdDelTask.Enabled = False
    If tdg_Task.Columns(1) <> "" Then 'Cek Task_Id
        Call viewActivity
        cmdAddActivity.Enabled = True
        cmdEditTask.Enabled = True
        cmdDelTask.Enabled = True
    End If

End Sub


Private Sub cmdAddActivity_Click()
    '<CSCM>
    '********************************************************************************
    'Procedure Name     : cmdAddActivity_Click
    'Procedure Function : ---
    'Input Parameter    : ---
    'Output Parameter   : ---
    'Date               : 3/29/2016
    'LastUpdate/By      : Tedi / Kreatif
    'Name Before        : -
    '********************************************************************************
    '</CSCM>
    frm_MPActivityAdd.show 1

End Sub

Private Sub cmdEditActivity_Click()
    '<CSCM>
    '********************************************************************************
    'Procedure Name     : cmdEditActivity_Click
    'Procedure Function : ---
    'Input Parameter    : ---
    'Output Parameter   : ---
    'Date               : 3/29/2016
    'LastUpdate/By      : Tedi / Kreatif
    'Name Before        : -
    '********************************************************************************
    '</CSCM>
    'Call MarkRow(FGActivity, vbYellow)
    Frm_MPActivityEdit.show 1
    'Call MarkRow(FGActivity, vbWindowBackground)

End Sub

Private Sub cmdDelActivity_Click()
    '<CSCM>
    '********************************************************************************
    'Procedure Name     : cmdDelActivity_Click
    'Procedure Function : ---
    'Input Parameter    : ---
    'Output Parameter   : ---
    'Date               : 3/29/2016
    'LastUpdate/By      : Tedi / Kreatif
    'Name Before        : -
    '********************************************************************************
    '</CSCM>
    Dim pesan
    Call MarkRow(tdg_Activity, vbYellow)
    pesan = MsgBox("Dou you want to Delete Activity?", vbExclamation + vbOKCancel, strApplication_Name)
    Call MarkRow(tdg_Activity, vbWindowBackground)
    If pesan = vbOK Then
        Call BeforeDelActivity
    End If

End Sub

Private Sub cmdShowDetail_Click()
    '<CSCM>
    '********************************************************************************
    'Procedure Name     : cmdShowDetail_Click
    'Procedure Function : ---
    'Input Parameter    : ---
    'Output Parameter   : ---
    'Date               : 3/29/2016
    'LastUpdate/By      : Tedi / Kreatif
    'Name Before        : -
    '********************************************************************************
    '</CSCM>
    Dim pesan
    'Call MarkRow(FGActivity, vbYellow)
    If Trim(tdg_Activity.Columns(1)) = "" Then
        pesan = MsgBox("No Medium Selected, Click Edit Activity Button to Select Medium!", vbExclamation, strApplication_Name)
    Else
        Frm_MPActivityDetail.show 1
        
    End If
    'Call MarkRow(FGActivity, vbWindowBackground)

End Sub

Private Sub AddTask(Task_desc As String)
    '<CSCM>
    '********************************************************************************
    'Procedure Name     : AddTask
    'Procedure Function : ---
    'Input Parameter    : ---
    'Output Parameter   : ---
    'Date               : 3/29/2016
    'LastUpdate/By      : Tedi / Kreatif
    'Name Before        : -
    '********************************************************************************
    '</CSCM>
    '*****************************************************************************
    ' Nama Prosedur     :   AddTask
    ' Fungsi Prosedur   :   Add Task to a Media Plan
    ' Parameter  Input  :   Task_desc as string
    ' Parameter Output  :
    ' Tgl Pembuatan     :   09 Agustus 2004
    ' Last Update/By    :   09 Agustus 2004/Sistyo
    '*****************************************************************************
    Dim strMP_Task_ID, strSql As String
    
    strSql = "select isnull(max(cast(substring(mp_task_id,16,4) as int)),0)+1 from mp_task where mp_number like '" & Mid(cboMPNum.Text, 1, 9) & "%'"
    
    rsTemp.Open strSql, ConnERP, 1, 3
    strMP_Task_ID = Mid(cboMPNum.Text, 1, 4) & ".TASK." & Mid(cboMPNum.Text, 6, 4) & "." & Right("0000" & CStr(rsTemp(0)), 4) 'Create new Task_Id
    rsTemp.Close
    strSql = "insert into mp_task(mp_task_id,mp_number,task_desc) values "
    strSql = strSql & "('" & strMP_Task_ID & "','" & cboMPNum.Text & "','" & Clear_String(Task_desc) & "')"
    ConnERP.Execute strSql
    'View new Task in FlexGrid
    ConnERP.Execute "update mp_master set last_update_by='" & strLogin_User & "',last_update_date = getdate() where mp_number='" & cboMPNum.Text & "'"
    ViewMediaPlan
    tdg_Task_Click

End Sub

Private Sub cmdAddTask_Click()
    '<CSCM>
    '********************************************************************************
    'Procedure Name     : cmdAddTask_Click
    'Procedure Function : ---
    'Input Parameter    : ---
    'Output Parameter   : ---
    'Date               : 3/29/2016
    'LastUpdate/By      : Tedi / Kreatif
    'Name Before        : -
    '********************************************************************************
    '</CSCM>
    Dim strTaskDesc As String
    strTaskDesc = InputBox("Input Task Description..", "Add Task")
    If Trim(strTaskDesc) <> "" Then
        Call AddTask(strTaskDesc)
    End If

End Sub

Private Sub EditTask(Task_desc As String)
    '<CSCM>
    '********************************************************************************
    'Procedure Name     : EditTask
    'Procedure Function : ---
    'Input Parameter    : ---
    'Output Parameter   : ---
    'Date               : 3/29/2016
    'LastUpdate/By      : Tedi / Kreatif
    'Name Before        : -
    '********************************************************************************
    '</CSCM>
    '*****************************************************************************
    ' Nama Prosedur     :   EditTask
    ' Fungsi Prosedur   :   Edit Task description
    ' Parameter  Input  :   Task_desc as string
    ' Parameter Output  :
    ' Tgl Pembuatan     :   09 Agustus 2004
    ' Last Update/By    :   09 Agustus 2004/Sistyo
    '*****************************************************************************
    ConnERP.Execute "update mp_task set task_desc = '" & Clear_String(Task_desc) & "' where mp_task_id='" & tdg_Task.Columns(1) & "'"
    ConnERP.Execute "update mp_master set last_update_by='" & strLogin_User & "',last_update_date= getdate() where mp_number='" & cboMPNum.Text & "'"
    ViewMediaPlan

End Sub

Private Sub cmdEditTask_Click()
    '<CSCM>
    '********************************************************************************
    'Procedure Name     : cmdEditTask_Click
    'Procedure Function : ---
    'Input Parameter    : ---
    'Output Parameter   : ---
    'Date               : 3/29/2016
    'LastUpdate/By      : Tedi / Kreatif
    'Name Before        : -
    '********************************************************************************
    '</CSCM>
    Dim strTaskDesc As String
    'Call MarkRow(FGTask, vbYellow)
    strTaskDesc = InputBox("Input Task Description..", "Edit Task", tdg_Task.Columns(2))
    'Call MarkRow(FGTask, vbWindowBackground)
    If Trim(strTaskDesc) <> "" Then
        Call EditTask(strTaskDesc)
    End If

End Sub

Private Sub DeleteTask()
    '<CSCM>
    '********************************************************************************
    'Procedure Name     : DeleteTask
    'Procedure Function : ---
    'Input Parameter    : ---
    'Output Parameter   : ---
    'Date               : 3/29/2016
    'LastUpdate/By      : Tedi / Kreatif
    'Name Before        : -
    '********************************************************************************
    '</CSCM>
    '*****************************************************************************
    ' Nama Prosedur     :   DeleteTask
    ' Fungsi Prosedur   :   Remove a Task from a Media plan
    ' Parameter  Input  :
    ' Parameter Output  :
    ' Tgl Pembuatan     :   09 Agustus 2004
    ' Last Update/By    :   09 Agustus 2004/Sistyo
    '*****************************************************************************
    Dim intRow As Integer
    Dim pesan
    
    ConnERP.Execute "delete from mp_task where mp_task_id='" & tdg_Task.Columns(1) & "'"
    ConnERP.Execute "update mp_master set last_update_by='" & strLogin_User & "',last_update_date = getdate() where mp_number='" & cboMPNum.Text & "'"
    If tdg_Task.ApproxCount = 0 Then
        cmdEditTask.Enabled = False
        cmdDelTask.Enabled = False
    End If
    Call ViewMediaPlan
    tdg_Task_Click
    pesan = MsgBox("Task Deleted!", vbExclamation, strApplication_Name)
    
End Sub

Private Sub db_Delete_Del_Plan()
    '<CSCM>
    '********************************************************************************
    'Procedure Name     : db_Delete_Del_Plan
    'Procedure Function : ---
    'Input Parameter    : ---
    'Output Parameter   : ---
    'Date               : 3/29/2016
    'LastUpdate/By      : Tedi / Kreatif
    'Name Before        : -
    '********************************************************************************
    '</CSCM>

    Dim strMsg As String
    strMsg = MsgBox("Do you want to delete this Plan?", vbExclamation + vbYesNo, strApplication_Name)
    
    If strMsg = 6 Then 'delete plan
        ConnERP.Execute "delete from mp_master where mp_number = '" & cboMPNum.Text & "'"
        cboMPNum.RemoveItem cboMPNum.ListIndex
        '        cmdDelPlan.Enabled = False
        strMsg = MsgBox("Media Plan Deleted!", vbExclamation, strApplication_Name)
        Call ViewMediaPlan
        Call viewActivity
        
        txtClientName.Text = ""
        txtBrandName.Text = ""
        
        cmdAddTask.Enabled = False
        cmdEditTask.Enabled = False
        cmdDelTask.Enabled = False
        
        cmdAddActivity.Enabled = False
        cmdEditActivity.Enabled = False
        cmdDelActivity.Enabled = False
        cmdShowDetail.Enabled = False
        
    End If
    
End Sub

Private Sub cmdDelTask_Click()
    '<CSCM>
    '********************************************************************************
    'Procedure Name     : cmdDelTask_Click
    'Procedure Function : ---
    'Input Parameter    : ---
    'Output Parameter   : ---
    'Date               : 3/29/2016
    'LastUpdate/By      : Tedi / Kreatif
    'Name Before        : -
    '********************************************************************************
    '</CSCM>

    Dim strMsg As String
    If tdg_Task.Columns(0) <> "" Then
        strMsg = MsgBox("Do you want to delete this task?", vbExclamation + vbYesNo, strApplication_Name)
        If strMsg = 6 Then
            Call BeforeDeleteTask 'cek isi task (medium, insertion)
        End If
    End If

End Sub

Private Sub db_save()
    '<CSCM>
    '********************************************************************************
    'Procedure Name     : db_Save
    'Procedure Function : ---
    'Input Parameter    : ---
    'Output Parameter   : ---
    'Date               : 3/29/2016
    'LastUpdate/By      : Tedi / Kreatif
    'Name Before        : -
    '********************************************************************************
    '</CSCM>
    objOpener = "Frm_MPEdit"
    Frm_MPCreate.show 1

End Sub

Private Sub Form_Unload(Cancel As Integer)
    '<CSCM>
    '********************************************************************************
    'Procedure Name     : Form_Unload
    'Procedure Function : ---
    'Input Parameter    : ---
    'Output Parameter   : ---
    'Date               : 3/29/2016
    'LastUpdate/By      : Tedi / Kreatif
    'Name Before        : -
    '********************************************************************************
    '</CSCM>
    Dim idx As Integer
    Dim strBrand_Filter As String
    
    frm_MPInsertion.cboMPNumber.Clear
    strBrand_Filter = "select brand_code from media_security_catalog where user_name = '" & strLogin_User & "' and position = 'Planner' and valid_until>=(select getdate())"
    rsTemp.Open "select mp_number from mp_master where brand_code in (" & strBrand_Filter & ") and is_latest=1 order by mp_number", ConnERP, 1, 3
    While Not rsTemp.EOF
        frm_MPInsertion.cboMPNumber.AddItem rsTemp(0)
        rsTemp.MoveNext
    Wend
    rsTemp.Close
    
    If cboMPNum.Text <> "" Then
        frm_MPInsertion.cboMPNumber.Text = cboMPNum.Text
    Else
        frm_MPInsertion.msf_MPInsertion.cols = 7
        frm_MPInsertion.msf_MPInsertion.Rows = 5
        frm_MPInsertion.msf_MPInsertion.TextMatrix(0, 6) = ""
        frm_MPInsertion.msf_MPInsertion.TextMatrix(1, 6) = ""
        frm_MPInsertion.msf_MPInsertion.TextMatrix(2, 6) = ""
        frm_MPInsertion.msf_MPInsertion.TextMatrix(3, 6) = ""
        frm_MPInsertion.msf_MPInsertion.TextMatrix(4, 6) = ""
        
        frm_MPInsertion.txtBrandName.Text = ""
        frm_MPInsertion.txtClientName.Text = ""
        frm_MPInsertion.txtCreatedDate.Text = ""
        frm_MPInsertion.txtCreatedBy.Text = ""
        frm_MPInsertion.txtLastUpdateDate.Text = ""
        frm_MPInsertion.txtLastUpdateBy.Text = ""
        
        frm_MPInsertion.lblReleaseDate.Caption = ""
        '        frm_MPInsertion.cmdExportToExcel.Enabled = False
        '        frm_MPInsertion.cmdEdit.Enabled = False
        '        frm_MPInsertion.cmdSummary.Enabled = False
    End If
    
End Sub

'Private Sub MarkRow(FG As Object, ColorIdx As Variant)
'    '<CSCM>
'    '********************************************************************************
'    'Procedure Name     : MarkRow
'    'Procedure Function : ---
'    'Input Parameter    : ---
'    'Output Parameter   : ---
'    'Date               : 3/29/2016
'    'LastUpdate/By      : Tedi / Kreatif
'    'Name Before        : -
'    '********************************************************************************
'    '</CSCM>
'    Dim i As Integer
'    With FG
'        For i = 1 To .cols - 1
'            .col = i
'            .CellBackColor = ColorIdx
'        Next
'    End With
'
'End Sub

Private Sub BeforeDeleteTask()
    '<CSCM>
    '********************************************************************************
    'Procedure Name     : BeforeDeleteTask
    'Procedure Function : ---
    'Input Parameter    : ---
    'Output Parameter   : ---
    'Date               : 3/29/2016
    'LastUpdate/By      : Tedi / Kreatif
    'Name Before        : -
    '********************************************************************************
    '</CSCM>
    Dim strTaskID As String
    Dim strSql As String
    Dim rsTemp As New ADODB.Recordset
    Dim intDataCount As Integer
    Dim pesan
    strTaskID = tdg_Task.Columns(1)
    'cek approval
    strSql = "select count(*) from mp_monthly_activity where approval=1 and mp_medium_id in (select mp_medium_id from mp_ids where mp_task_id='" & strTaskID & "')"
    rsTemp.Open strSql, ConnERP, 1, 3
    intDataCount = rsTemp(0)
    rsTemp.Close
    If intDataCount <> 0 Then
        MsgBox "This Task Contains data that has been approved and can not be deleted!", vbExclamation, strApplication_Name
        Exit Sub
    End If
    
    'cek insertion
    strSql = "select count(*) from mp_insertion where mp_plan_dim_id in (select mp_plan_dim_id from mp_ids where mp_task_id='" & strTaskID & "')"
    rsTemp.Open strSql, ConnERP, 1, 3
    intDataCount = rsTemp(0)
    rsTemp.Close
    If intDataCount <> 0 Then
        pesan = MsgBox("This Task Contains insertion data!" & vbCrLf & "Continue Delete Task?", vbYesNo + vbQuestion, strApplication_Name)
        If pesan = 7 Then
            Exit Sub
        End If
    End If
    
    Call DeleteTask
    
End Sub

Private Sub BeforeDelActivity()
    '<CSCM>
    '********************************************************************************
    'Procedure Name     : BeforeDelActivity
    'Procedure Function : ---
    'Input Parameter    : ---
    'Output Parameter   : ---
    'Date               : 3/29/2016
    'LastUpdate/By      : Tedi / Kreatif
    'Name Before        : -
    '********************************************************************************
    '</CSCM>
    Dim strActivityID As String
    Dim strSql As String
    Dim rsTemp As New ADODB.Recordset
    Dim intDataCount As Integer
    Dim pesan
    strActivityID = tdg_Activity.Columns(1)
    
    'cek approval
    strSql = "select count(*) from mp_monthly_activity where approval=1 and mp_medium_id in (select mp_medium_id from mp_ids where mp_activity_id='" & strActivityID & "')"
    rsTemp.Open strSql, ConnERP, 1, 3
    intDataCount = rsTemp(0)
    rsTemp.Close
    If intDataCount <> 0 Then
        MsgBox "This Activity Contains data that has been approved and can not be deleted!", vbExclamation, strApplication_Name
        Exit Sub
    End If
    
    'cek insertion
    strSql = "select count(*) from mp_insertion where mp_plan_dim_id in (select mp_plan_dim_id from mp_ids where mp_activity_id='" & strActivityID & "')"
    rsTemp.Open strSql, ConnERP, 1, 3
    intDataCount = rsTemp(0)
    rsTemp.Close
    If intDataCount <> 0 Then
        pesan = MsgBox("This Task Contains insertion data!" & vbCrLf & "Continue Delete Task?", vbYesNo + vbQuestion, strApplication_Name)
        If pesan = 7 Then
            Exit Sub
        End If
    End If
    
    Call DelActivity
    tdg_Task_Click
End Sub

Sub SetButtonToolbar(ByVal paIsNormalMode As Boolean, picOBJ) 'TOOLBAR_AI.
    '<CSCM>
    '********************************************************************************
    'Procedure Name     : SetButtonToolbar
    'Procedure Function : ---
    'Input Parameter    : ---
    'Output Parameter   : ---
    'Date               : 3/29/2016
    'LastUpdate/By      : Tedi / Kreatif
    'Name Before        : -
    '********************************************************************************
    '</CSCM>

    Dim objElement As Object
    Dim strDummy As String
    
    With picButton(enButtonType.bieAdd)    'ADD. 4
        .Enabled = paIsNormalMode
        .Visible = paIsNormalMode
    End With
    With picButton(enButtonType.bieDelete)  'EDIT. 5
        .Enabled = paIsNormalMode
        .Visible = paIsNormalMode
    End With
    With picButton(enButtonType.bieClose)     'CLOSE. 23
        .Enabled = paIsNormalMode
        .Visible = paIsNormalMode
    End With

    For Each objElement In picOBJ
        SetPictureTB objElement.Index, paIsNormalMode, picOBJ
    Next objElement
    'Call SetSecurityCRUDStandar("Duration Catalog", picButton, "1")

End Sub

Private Sub picButton_Click(Index As Integer)
    '<CSCM>
    '********************************************************************************
    'Procedure Name     : picButton_Click
    'Procedure Function : ---
    'Input Parameter    : ---
    'Output Parameter   : ---
    'Date               : 3/29/2016
    'LastUpdate/By      : Tedi / Kreatif
    'Name Before        : -
    '********************************************************************************
    '</CSCM>

    '************************************************
    ' Procedure         : picButton_Click
    ' Function          : Action utk Navigation dan CRUD.
    ' Created By        : {73 64 6B}
    ' Date              : 12-Apr-2015/{73 64 6B} --> Semua coding dan query sudah di optimalkan agar faster, readable, safer, standardable.
    '************************************************
    Dim strCode As String, strFileRpt As String
    'Lock_MainForm True
    Select Case Index
        Case enButtonType.bieAdd  'Create New
            db_save
        Case enButtonType.bieDelete
            db_Delete_Del_Plan
        Case enButtonType.bieClose   'Exit.
            Unload Me
    End Select

End Sub

Private Sub picButton_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single) 'TOOLBAR_AI.
    '<CSCM>
    '********************************************************************************
    'Procedure Name     : picButton_MouseDown
    'Procedure Function : ---
    'Input Parameter    : ---
    'Output Parameter   : ---
    'Date               : 3/29/2016
    'LastUpdate/By      : Tedi / Kreatif
    'Name Before        : -
    '********************************************************************************
    '</CSCM>
    '************************************************
    ' Procedure         : picButton_MouseDown
    ' Function          : TOOLBAR_AI saat mouse ditekan.
    ' Created By        : {73 64 6B}
    ' Date              : 12-Apr-2015
    '************************************************
    
    picButton(Index).Picture = LoadPicture(SetButtonImageEffect(Index, bieDown)) 'FIRST.

End Sub

Private Sub picButton_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single) 'TOOLBAR_AI.
    '<CSCM>
    '********************************************************************************
    'Procedure Name     : picButton_MouseMove
    'Procedure Function : ---
    'Input Parameter    : ---
    'Output Parameter   : ---
    'Date               : 3/29/2016
    'LastUpdate/By      : Tedi / Kreatif
    'Name Before        : -
    '********************************************************************************
    '</CSCM>
    '************************************************
    ' Procedure         : picButton_MouseMove
    ' Function          : TOOLBAR_AI saat mouse berada di area button.
    ' Created By        : {73 64 6B}
    ' Date              : 12-Apr-2015
    '************************************************
    
    picButton_Obj Index, Button, Shift, X, Y, picButton

End Sub

Private Sub EnableObject(ByVal paIsEnable As Boolean)
    '<CSCM>
    '********************************************************************************
    'Procedure Name     : EnableObject
    'Procedure Function : ---
    'Input Parameter    : ---
    'Output Parameter   : ---
    'Date               : 3/29/2016
    'LastUpdate/By      : Tedi / Kreatif
    'Name Before        : -
    '********************************************************************************
    '</CSCM>
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

Private Sub tdg_Activity_Click()
    '<CSCM>
    '********************************************************************************
    'Procedure Name     : tdg_Activity_Click
    'Procedure Function : ---
    'Input Parameter    : ---
    'Output Parameter   : ---
    'Date               : 3/29/2016
    'LastUpdate/By      : Tedi / Kreatif
    'Name Before        : FGActivity_Click
    '********************************************************************************
    '</CSCM>
    If tdg_Activity.Columns(1) <> "" Then 'Cek Activity_Id
        cmdEditActivity.Enabled = True
        cmdDelActivity.Enabled = True
        cmdShowDetail.Enabled = True
    End If

End Sub

Private Sub tdg_Activity_DblClick()
'<CSCM>
'********************************************************************************
'Procedure Name     : tdg_Activity_DblClick
'Procedure Function : ---
'Input Parameter    : ---
'Output Parameter   : ---
'Date               : 3/29/2016
'LastUpdate/By      : Tedi / Kreatif
'Name Before        : -
'********************************************************************************
'</CSCM>
    If tdg_Activity.Columns(1) <> "" Then 'Cek Activity_Id
        cmdEditActivity.Enabled = True
        cmdDelActivity.Enabled = True
        cmdShowDetail.Enabled = True
        Call cmdShowDetail_Click
    End If
End Sub

Private Sub tdg_Task_Click()
    '<CSCM>
    '********************************************************************************
    'Procedure Name     : tdg_Task_Click
    'Procedure Function : ---
    'Input Parameter    : ---
    'Output Parameter   : ---
    'Date               : 3/29/2016
    'LastUpdate/By      : Tedi / Kreatif
    'Name Before        : -
    '********************************************************************************
    '</CSCM>
    cmdEditActivity.Enabled = False
    cmdDelActivity.Enabled = False
    cmdAddActivity.Enabled = False
    cmdShowDetail.Enabled = False
    cmdEditTask.Enabled = False
    cmdDelTask.Enabled = False
    If tdg_Task.Columns(1) <> "" Then  'Cek Task_Id
        Call viewActivity
        cmdAddActivity.Enabled = True
        cmdEditTask.Enabled = True
        cmdDelTask.Enabled = True
    End If

End Sub

Public Sub create_table_TempActivity()

    '<CSCM>
    '********************************************************************************
    'Procedure Name     : create_table_TempActivity
    'Procedure Function : ---
    'Input Parameter    : ---
    'Output Parameter   : ---
    'Date               : 3/29/2016
    'LastUpdate/By      : Tedi / Kreatif
    'Name Before        : -
    '********************************************************************************
    '</CSCM>
    '*****************************************
    'Procedure Name     : create_table_TempActivity
    'Procedure Function : create_table_TempActivity
    'Input Parameter    : -
    'Output Parameter   : -
    'Last Update Date   : 18/3/2016
    'Last Update By     : Tedi / Kreatif
    '*********************************

    Set rec_TempActivity = Nothing
    Set rec_TempActivity = New ADODB.Recordset
    
    With rec_TempActivity.Fields
    
        .Append "No", adDouble, 3, adFldIsNullable
        .Append "mp_activity_id", adVarChar, 19, adFldIsNullable
        .Append "mp_task_id", adChar, 19, adFldIsNullable
        .Append "activity_type", adChar, 50, adFldIsNullable
        .Append "activity_desc", adChar, 255, adFldIsNullable
        .Append "brand_variant_name", adChar, 50, adFldIsNullable
        .Append "target_audience", adChar, 50, adFldIsNullable
        .Append "brand_target", adChar, 50, adFldIsNullable
        .Append "Medium_name", adChar, 255, adFldIsNullable
        
    End With
    rec_TempActivity.Open
    
End Sub

Private Sub MarkRow(FG As Object, ColorIdx As Variant)
Dim i As Integer
    With tdg_Activity.Columns
        For i = 1 To .Count - 1
            tdg_Activity.col = i
            tdg_Activity.Columns(i).BackColor = ColorIdx
        Next
    End With
End Sub
