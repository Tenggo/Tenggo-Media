VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.ocx"
Object = "{562E3E04-2C31-4ECE-83F4-4017EEE51D40}#8.0#0"; "todg8.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFlxGrd.ocx"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form Frm_IB_Radio 
   BorderStyle     =   0  'None
   Caption         =   "Implementation Brief Radio"
   ClientHeight    =   8805
   ClientLeft      =   0
   ClientTop       =   -45
   ClientWidth     =   13095
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   177
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   8805
   ScaleWidth      =   13095
   ShowInTaskbar   =   0   'False
   Begin Threed.SSPanel pnlMain 
      Align           =   1  'Align Top
      Height          =   7665
      Left            =   0
      TabIndex        =   17
      Top             =   750
      Width           =   13095
      _Version        =   65536
      _ExtentX        =   23098
      _ExtentY        =   13520
      _StockProps     =   15
      BackColor       =   15790320
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   177
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BevelOuter      =   1
      Begin VB.Frame fraPlanMonth 
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
         Height          =   3480
         Left            =   150
         TabIndex        =   51
         Top             =   1620
         Width           =   12885
         Begin VB.ComboBox cboPlanMonth 
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
            Left            =   1425
            TabIndex        =   55
            Top             =   240
            Width           =   1800
         End
         Begin VB.CommandButton cmdEditPlan 
            Caption         =   "Edit Plan"
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
            Left            =   11505
            TabIndex        =   54
            Top             =   1080
            Width           =   1180
         End
         Begin VB.CommandButton cmdMateri 
            Caption         =   "Material"
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
            Left            =   11505
            TabIndex        =   53
            Top             =   735
            Width           =   1180
         End
         Begin VB.ComboBox cboPlanDetail 
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
            Left            =   5175
            TabIndex        =   52
            Top             =   240
            Width           =   1800
         End
         Begin MSFlexGridLib.MSFlexGrid dbgCity 
            Height          =   2295
            Left            =   2370
            TabIndex        =   56
            Top             =   720
            Width           =   8940
            _ExtentX        =   15769
            _ExtentY        =   4048
            _Version        =   393216
            FixedCols       =   0
            BackColorFixed  =   12356167
            ForeColorFixed  =   16777215
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
         Begin MSFlexGridLib.MSFlexGrid msgMix 
            Height          =   1815
            Left            =   7260
            TabIndex        =   57
            Top             =   720
            Width           =   3030
            _ExtentX        =   5345
            _ExtentY        =   3201
            _Version        =   393216
            FixedCols       =   0
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
         Begin TrueOleDBGrid80.TDBGrid tdgPlan 
            Bindings        =   "Frm_IB_Radio.frx":0000
            Height          =   2250
            Left            =   360
            TabIndex        =   63
            Top             =   750
            Width           =   1860
            _ExtentX        =   3281
            _ExtentY        =   3969
            _LayoutType     =   4
            _RowHeight      =   -2147483647
            _WasPersistedAsPixels=   0
            Columns(0)._VlistStyle=   0
            Columns(0)._MaxComboItems=   5
            Columns(0).Caption=   "Client Brief Id"
            Columns(0).DataField=   "Client_Brief_Id"
            Columns(0)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns.Count   =   1
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
            Splits(0)._ColumnProps(0)=   "Columns.Count=1"
            Splits(0)._ColumnProps(1)=   "Column(0).Width=5186"
            Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
            Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=5106"
            Splits(0)._ColumnProps(4)=   "Column(0)._EditAlways=0"
            Splits(0)._ColumnProps(5)=   "Column(0).Order=1"
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
            _StyleDefs(32)  =   "Splits(0).Columns(0).Style:id=28,.parent=13"
            _StyleDefs(33)  =   "Splits(0).Columns(0).HeadingStyle:id=25,.parent=14"
            _StyleDefs(34)  =   "Splits(0).Columns(0).FooterStyle:id=26,.parent=15"
            _StyleDefs(35)  =   "Splits(0).Columns(0).EditorStyle:id=27,.parent=17"
            _StyleDefs(36)  =   "Named:id=33:Normal"
            _StyleDefs(37)  =   ":id=33,.parent=0"
            _StyleDefs(38)  =   "Named:id=34:Heading"
            _StyleDefs(39)  =   ":id=34,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
            _StyleDefs(40)  =   ":id=34,.wraptext=-1"
            _StyleDefs(41)  =   "Named:id=35:Footing"
            _StyleDefs(42)  =   ":id=35,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
            _StyleDefs(43)  =   "Named:id=36:Selected"
            _StyleDefs(44)  =   ":id=36,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
            _StyleDefs(45)  =   "Named:id=37:Caption"
            _StyleDefs(46)  =   ":id=37,.parent=34,.alignment=2"
            _StyleDefs(47)  =   "Named:id=38:HighlightRow"
            _StyleDefs(48)  =   ":id=38,.parent=33,.bgcolor=&HFF0000&,.fgcolor=&H8000000E&,.borderColor=&HFF2B2B&"
            _StyleDefs(49)  =   "Named:id=39:EvenRow"
            _StyleDefs(50)  =   ":id=39,.parent=33,.bgcolor=&HFFFF00&"
            _StyleDefs(51)  =   "Named:id=40:OddRow"
            _StyleDefs(52)  =   ":id=40,.parent=33"
            _StyleDefs(53)  =   "Named:id=41:RecordSelector"
            _StyleDefs(54)  =   ":id=41,.parent=34"
            _StyleDefs(55)  =   "Named:id=42:FilterBar"
            _StyleDefs(56)  =   ":id=42,.parent=33,.fgcolor=&H80000005&"
         End
         Begin VB.Label lblView 
            BackStyle       =   0  'Transparent
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
            Height          =   255
            Left            =   2565
            TabIndex        =   62
            Top             =   3030
            Width           =   570
         End
         Begin MSForms.OptionButton optCity 
            Height          =   300
            Left            =   4470
            TabIndex        =   61
            Top             =   3000
            Width           =   1875
            BackColor       =   -2147483633
            ForeColor       =   -2147483630
            DisplayStyle    =   5
            Size            =   "3307;529"
            Value           =   "0"
            Caption         =   "by City or Station"
            FontName        =   "Tahoma"
            FontHeight      =   165
            FontCharSet     =   0
            FontPitchAndFamily=   2
         End
         Begin MSForms.OptionButton optArea 
            Height          =   300
            Left            =   3270
            TabIndex        =   60
            Top             =   3000
            Width           =   1035
            BackColor       =   -2147483633
            ForeColor       =   -2147483630
            DisplayStyle    =   5
            Size            =   "1826;529"
            Value           =   "1"
            Caption         =   "by Area"
            FontName        =   "Tahoma"
            FontHeight      =   165
            FontCharSet     =   0
            FontPitchAndFamily=   2
         End
         Begin VB.Label lblPlanMonth 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Plan Month "
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
            Left            =   90
            TabIndex        =   59
            Top             =   255
            Width           =   1245
         End
         Begin VB.Label lblPlanDetail 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Plan Detail by "
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
            Left            =   3855
            TabIndex        =   58
            Top             =   255
            Width           =   1245
         End
      End
      Begin VB.Frame fraIB 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   1545
         Left            =   5730
         TabIndex        =   20
         Top             =   60
         Width           =   7590
         Begin VB.TextBox txtMediaPlanNo 
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
            Left            =   5550
            MaxLength       =   15
            TabIndex        =   25
            Top             =   240
            Width           =   1680
         End
         Begin VB.TextBox txtIBID 
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
            Height          =   325
            Left            =   1575
            TabIndex        =   24
            Top             =   1035
            Width           =   1770
         End
         Begin VB.TextBox txtPrimaryTarget 
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
            Left            =   5565
            MaxLength       =   30
            TabIndex        =   23
            Top             =   630
            Width           =   2070
         End
         Begin VB.ComboBox cboSecondary 
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
            Left            =   5550
            TabIndex        =   22
            Top             =   1035
            Width           =   2100
         End
         Begin VB.ComboBox cboStartingMonth 
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
            TabIndex        =   21
            Top             =   240
            Width           =   1785
         End
         Begin MSComCtl2.DTPicker dtpIBDate 
            Height          =   315
            Left            =   1575
            TabIndex        =   26
            Top             =   630
            Width           =   1800
            _ExtentX        =   3175
            _ExtentY        =   556
            _Version        =   393216
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Format          =   111607809
            CurrentDate     =   36805
         End
         Begin VB.Label lblMediaPlanNo 
            BackStyle       =   0  'Transparent
            Caption         =   "Media &Plan No "
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
            Left            =   3960
            TabIndex        =   32
            Top             =   300
            Width           =   1755
         End
         Begin VB.Label lblIBId 
            AutoSize        =   -1  'True
            Caption         =   "IB Id "
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
            Left            =   150
            TabIndex        =   31
            Top             =   1095
            Width           =   390
         End
         Begin VB.Label lblPrimaryTarget 
            BackStyle       =   0  'Transparent
            Caption         =   "Pri&mary Target "
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
            Left            =   3960
            TabIndex        =   30
            Top             =   675
            Width           =   1755
         End
         Begin VB.Label lblSecondaryTarget 
            BackStyle       =   0  'Transparent
            Caption         =   "Secondar&y Target "
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
            Left            =   3960
            TabIndex        =   29
            Top             =   1065
            Width           =   1755
         End
         Begin VB.Label lblStartingMonth 
            BackStyle       =   0  'Transparent
            Caption         =   "Starting Mont&h "
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
            Left            =   150
            TabIndex        =   28
            Top             =   285
            Width           =   1515
         End
         Begin VB.Label lblIBDate 
            AutoSize        =   -1  'True
            Caption         =   "IB Da&te "
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
            Left            =   150
            TabIndex        =   27
            Top             =   660
            Width           =   585
         End
      End
      Begin VB.Frame Frame1 
         Height          =   1560
         Left            =   150
         TabIndex        =   39
         Top             =   60
         Width           =   5475
         Begin VB.ComboBox cboBrand 
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
            Left            =   1425
            TabIndex        =   45
            Top             =   615
            Width           =   3900
         End
         Begin VB.ComboBox cboYear 
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
            Left            =   1425
            TabIndex        =   44
            Top             =   225
            Width           =   1305
         End
         Begin VB.ComboBox cboBrandVariant 
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
            Left            =   1425
            TabIndex        =   43
            Top             =   1005
            Width           =   3900
         End
         Begin VB.Label lblBrand 
            AutoSize        =   -1  'True
            Caption         =   "&Brand "
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
            Left            =   150
            TabIndex        =   42
            Top             =   675
            Width           =   465
         End
         Begin VB.Label lblYear 
            BackStyle       =   0  'Transparent
            Caption         =   "Y&ear "
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
            Left            =   150
            TabIndex        =   41
            Top             =   285
            Width           =   1245
         End
         Begin VB.Label lblBrandVariant 
            AutoSize        =   -1  'True
            Caption         =   "&Brand Variant "
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
            Left            =   150
            TabIndex        =   40
            Top             =   1065
            Width           =   1020
         End
      End
      Begin VB.Frame fraClientApproval 
         Appearance      =   0  'Flat
         Caption         =   "Client approval"
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
         Height          =   1800
         Left            =   7905
         TabIndex        =   33
         Top             =   5685
         Width           =   3630
         Begin VB.PictureBox picApproved 
            BorderStyle     =   0  'None
            Height          =   1290
            Left            =   465
            ScaleHeight     =   1290
            ScaleWidth      =   2880
            TabIndex        =   46
            Top             =   480
            Width           =   2880
            Begin VB.Label lblTimeApp 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "24:00:00"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H000000C0&
               Height          =   240
               Left            =   1755
               TabIndex        =   49
               Top             =   375
               Width           =   810
            End
            Begin VB.Label lblDateApp 
               Alignment       =   2  'Center
               BackStyle       =   0  'Transparent
               Caption         =   "12/12/2001"
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
               Left            =   165
               TabIndex        =   48
               Top             =   360
               Width           =   2355
            End
            Begin VB.Label lblApprove 
               Alignment       =   2  'Center
               BackStyle       =   0  'Transparent
               Caption         =   "UNAPPROVED"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   12
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H000000C0&
               Height          =   315
               Left            =   -45
               TabIndex        =   47
               Top             =   0
               Width           =   2820
            End
         End
         Begin VB.Label lblStatus 
            Alignment       =   2  'Center
            ForeColor       =   &H000000FF&
            Height          =   270
            Left            =   705
            TabIndex        =   34
            Top             =   1005
            Width           =   2475
         End
      End
      Begin VB.TextBox txtAttachment 
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
         Height          =   690
         Left            =   150
         MaxLength       =   500
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   19
         Top             =   5790
         Width           =   7590
      End
      Begin VB.TextBox txtAnyCons 
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
         Height          =   705
         Left            =   150
         MaxLength       =   500
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   18
         Top             =   6825
         Width           =   7575
      End
      Begin MSMask.MaskEdBox medGrandTotal 
         Height          =   315
         Left            =   9105
         TabIndex        =   35
         Top             =   5250
         Width           =   2430
         _ExtentX        =   4286
         _ExtentY        =   556
         _Version        =   393216
         Appearance      =   0
         Enabled         =   0   'False
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   "#,##0"
         PromptChar      =   "_"
      End
      Begin VB.Label lblAttachment 
         BackStyle       =   0  'Transparent
         Caption         =   "Attachment "
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
         Left            =   180
         TabIndex        =   38
         Top             =   5460
         Width           =   1515
      End
      Begin VB.Label lblAnyCons 
         BackStyle       =   0  'Transparent
         Caption         =   "Any Consideration "
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
         Left            =   150
         TabIndex        =   37
         Top             =   6525
         Width           =   1755
      End
      Begin VB.Label lblGrandTotal 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Grand Total "
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
         Left            =   8145
         TabIndex        =   36
         Top             =   5310
         Width           =   885
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
      ScaleWidth      =   13095
      TabIndex        =   9
      TabStop         =   0   'False
      Top             =   8475
      Width           =   13095
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
         TabIndex        =   14
         TabStop         =   0   'False
         Top             =   15
         Visible         =   0   'False
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
         TabIndex        =   13
         TabStop         =   0   'False
         Top             =   15
         Visible         =   0   'False
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
         TabIndex        =   12
         TabStop         =   0   'False
         Top             =   15
         Visible         =   0   'False
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
         TabIndex        =   11
         TabStop         =   0   'False
         Top             =   15
         Visible         =   0   'False
         Width           =   300
      End
      Begin VB.PictureBox picDescColor 
         BorderStyle     =   0  'None
         Height          =   345
         Left            =   9975
         ScaleHeight     =   345
         ScaleWidth      =   1695
         TabIndex        =   10
         TabStop         =   0   'False
         Top             =   15
         Width           =   1695
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
         TabIndex        =   16
         Tag             =   "Last Modified Date: "
         Top             =   75
         Width           =   2520
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
         TabIndex        =   15
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
      ScaleWidth      =   13095
      TabIndex        =   0
      Top             =   0
      Width           =   13095
      Begin VB.PictureBox picButton 
         Appearance      =   0  'Flat
         BackColor       =   &H00F0F0F0&
         BorderStyle     =   0  'None
         ClipControls    =   0   'False
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
         Height          =   750
         Index           =   66
         Left            =   6210
         ScaleHeight     =   750
         ScaleWidth      =   1500
         TabIndex        =   50
         Top             =   -15
         Width           =   1500
      End
      Begin VB.PictureBox picButton 
         Appearance      =   0  'Flat
         BackColor       =   &H00F0F0F0&
         BorderStyle     =   0  'None
         ClipControls    =   0   'False
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
         Height          =   750
         Index           =   11
         Left            =   10800
         ScaleHeight     =   750
         ScaleWidth      =   1500
         TabIndex        =   8
         Top             =   0
         Width           =   1500
      End
      Begin VB.PictureBox picButton 
         Appearance      =   0  'Flat
         BackColor       =   &H00F0F0F0&
         BorderStyle     =   0  'None
         ClipControls    =   0   'False
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
         Height          =   750
         Index           =   4
         Left            =   90
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
         Height          =   750
         Index           =   8
         Left            =   7740
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
         Height          =   750
         Index           =   10
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
         Height          =   750
         Index           =   7
         Left            =   4680
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
         Height          =   750
         Index           =   6
         Left            =   3150
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
         Height          =   750
         Index           =   5
         Left            =   1620
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
         Height          =   750
         Index           =   23
         Left            =   9270
         ScaleHeight     =   750
         ScaleWidth      =   1500
         TabIndex        =   1
         Top             =   0
         Width           =   1500
      End
   End
   Begin Crystal.CrystalReport crIB 
      Left            =   4275
      Top             =   8340
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      PrintFileLinesPerPage=   60
   End
End
Attribute VB_Name = "Frm_IB_Radio"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
''************************************************
' Function          : entry IB radio'
' Last Update/By    :
'************************************************
Option Explicit

Const Not_Approved = "UNAPPROVED"
Const Approved = "APPROVED"

Dim intIBStartingMonth As Integer

Public strTransProcess As String 'dw

Dim recMateriMix As New ADODB.Recordset
Public recMateriMixTemp As New ADODB.Recordset

Dim recMateri As New ADODB.Recordset
Public recMateriTemp As New ADODB.Recordset

Public recRadioPlan As New ADODB.Recordset
Public recRadioPlanTemp As New ADODB.Recordset

Dim recPlanDetail As New ADODB.Recordset
Public recPlanDetailTemp As New ADODB.Recordset

Dim strButtinCaption As String
Dim recPlanDetailMaterial As New ADODB.Recordset
Public recPlanDetailMaterialTemp As New ADODB.Recordset
Dim strPrivTransac As String

Private Sub Form_Load()

    
    
    Me.AutoRedraw = True
   ' Call VGradient(Me, &HFF8090, &HFF8090, Me.Line2.Y1, Me.Height, 0, Me.Width)
    
    LoadYear cboYear
    
    LoadBrand cboBrand, strLogin_User, "'planner'"
  
    LoadSecondaryTarget cboSecondary
        
    InitialDataMonth
           
    cboPlanDetail.AddItem "City"
    cboPlanDetail.AddItem "Station"
    
    cboPlanDetail.ListIndex = 0
           
    Call setForm(True)
        
    SetButton "NO DATA"
    
    strTransProcess = "NO DATA"
    
    optArea.Enabled = False
    optCity.Enabled = False
    AdjustSizeForm
    Call EnableObject(False)
    
End Sub

Private Sub cboBrand_Click()
    LoadBrandVariant cboBrandVariant, Left(cboBrand.Text, 4)

    If cboBrand.ListIndex <> -1 Then
        Call ClearForm
        
        Call InitGrids
        
        Call setForm(True)
        
        SetButton "NO DATA"
    End If
End Sub

Private Sub cboBrandVariant_click()
    If cboBrandVariant.ListIndex <> -1 Then

        Call ClearForm
        
        Call InitGrids

        Call setForm(True)
    End If
End Sub

Private Sub cboPlanMonth_Click()
    If cboPlanMonth.ListIndex <> -1 Then
        LoadDetail txtIBID.Text, cboYear.Text, Get_Month_Number(cboPlanMonth.Text)
    End If
End Sub

Private Sub cboStartingMonth_Click()
    If cboStartingMonth.ListIndex <> -1 Then
        intIBStartingMonth = Get_Month_Number(cboStartingMonth.Text)
        
        Call monthIB(cboPlanMonth)
    End If
End Sub

Private Sub cboStartingMonth_KeyPress(KeyAscii As Integer)
    KeyAscii = 0
End Sub

Private Sub cboPlanDetail_KeyPress(KeyAscii As Integer)
    KeyAscii = 0
End Sub

Private Sub cboYear_Click()
    If cboYear.ListIndex <> -1 Then
                
        Call ClearForm
        
        Call InitGrids
        
        Call setForm(True)
        
        SetButton "NO DATA"
    End If
End Sub

Private Sub cboYear_KeyPress(KeyAscii As Integer)
    KeyAscii = 0
End Sub

Private Sub db_add()
'cmdAdd_Click
    If Trim(cboYear.Text) = "" Then
        MsgBox "Please Select the Year !", vbExclamation, strApplication_Name
        Exit Sub
    End If
    
    If Trim(cboBrand.Text) = "" Then
        MsgBox "Please Select a Brand !", vbExclamation, strApplication_Name
        Exit Sub
    End If
    
    If Trim(cboBrandVariant.Text) = "" Then
        MsgBox "Please Select a Brand Variant !", vbExclamation, strApplication_Name
        Exit Sub
    End If
        
    'Check apakah Planner Brand
    If Not IsValidAccess(strLogin_User, "Planner", Left(cboBrand.Text, 4)) Then
        MsgBox strMsgAccessDenied, vbCritical, strTitleExclamation
        Exit Sub
    End If
    
    If Not IsFeeAlreadyEntered(Left(cboBrand.Text, 4), cboYear.Text) Then
        MsgBox "Brand Fee information not found for selected year !", vbCritical, strTitleExclamation
        Exit Sub
    End If

    Call ClearForm
    
    Call setForm(False)
    
    SetButton "ADD"
    Call SetButtonToolbar(False, picButton)
    txtIBID.Text = GenerateIBID

    LoadSecondaryTarget cboSecondary

    'Default Value
    recDate.Requery
    'txtEnteredDate.Text = recDate.Fields(0).Value
    dtpIBDate.Value = recDate.Fields(0).Value
    'txtEnteredBy.Text = strLogin_FullName
    
   
    lblApprove.Caption = "UNAPPROVED"
    
    lblLastModifiedDate.Caption = "Last Modified Date: " & recDate.Fields(0).Value & " | "
    lblLastModifiedBy.Caption = "Last Modified By: " & strLogin_FullName
    
    strTransProcess = "ADD"
    
    Call InitGrids
                     
    'Open Temporery Recordset
    PrepareTemp txtIBID.Text
    
    'Get month
    InitialDataMonth
                       
    cboYear.Enabled = False
    cboBrand.Enabled = False
    cboBrandVariant.Enabled = False
    'txtEnteredDate.Enabled = False
    fraClientApproval.Enabled = False
    
    cboStartingMonth.Enabled = True
    cboPlanDetail.Enabled = True
        
    optArea.Enabled = True
    optCity.Enabled = True
    
    dtpIBDate.SetFocus
End Sub

Private Sub db_Cancel()
'cmdCancel_Click
    Dim strTempIBId As String
    Dim prmIBIDIN As New ADODB.Parameter
    Dim cmdIB As New ADODB.Command

    If strTransProcess = "ADD" Then
        If txtIBID.Text <> "" Then
            'Cancel IB
            '=======================
            Rem Set Stored Procedure
            cmdIB.CommandType = adCmdStoredProc
            cmdIB.CommandText = "Cancel_IB_Radio_ID"
            
            Set prmIBIDIN = cmdIB.CreateParameter("New_IB_Id", adChar, adParamInput, 13)
            
            cmdIB.Parameters.Append prmIBIDIN
            
            prmIBIDIN.Value = Trim(txtIBID.Text)
            
            cmdIB.ActiveConnection = ConnERP
            cmdIB.Execute
            
            Call ClearForm
            
            strTransProcess = "NO DATA"
            
            SetButton "NO DATA"
        End If
    End If
    
    If strTransProcess = "EDIT" Then
        strTempIBId = txtIBID.Text
        PrepareTemp strTempIBId
        
        ShowData strTempIBId
        
        strTransProcess = "SHOW"
        
        SetButton "SHOW"
    End If
    
    Call setForm(True)
    Call SetButtonToolbar(True, picButton)
    cboYear.Enabled = True
    cboBrand.Enabled = True
    cboBrandVariant.Enabled = True
    cboPlanDetail.Enabled = False
    
    'Lock Approval Frame
    fraClientApproval.Enabled = True
End Sub

Private Sub db_close()
'cmdClose_Click
    Set recMateriTemp = Nothing
    
    Set recMateriMix = Nothing
    Set recMateriMixTemp = Nothing
        
    Set recRadioPlanTemp = Nothing
       
    Set recPlanDetail = Nothing
    Set recPlanDetailTemp = Nothing
    
    Unload Me
End Sub

Private Sub db_delete()
'before cmdDelete_Click
    Dim prmIN As New ADODB.Parameter
    Dim cmdSP As New ADODB.Command
    
    'Check apakah Planner Brand
    If lblApprove.Caption = "APPROVED" Then
        MsgBox "Can not delete data approved !", vbCritical, strApplication_Name
    Exit Sub
    End If
    If Not IsValidAccess(strLogin_User, "Planner", Left(cboBrand.Text, 4)) Then
        MsgBox strMsgAccessDenied, vbCritical, strTitleExclamation
        Exit Sub
    End If
     
    If Me.txtIBID.Text = "" Then
        MsgBox "Please Select a Brief to delete !", vbInformation, strApplication_Name
        Exit Sub
    End If
        
    If lblApprove.Caption = "APPROVED" Then
        MsgBox "This Implementation Brief Radio has been Approved by Client", vbCritical, strTitleInfo
        Exit Sub
    End If
    
    If MsgBox(strMsgDeleteConfirm, vbInformation + vbYesNo, strTitleConfirm) = vbYes Then
        cmdSP.CommandType = adCmdStoredProc
        cmdSP.CommandText = "Delete_IB_Radio"
        
        Set prmIN = cmdSP.CreateParameter("IB_ID", adChar, adParamInput, 13)
        cmdSP.Parameters.Append prmIN
        prmIN.Value = Me.txtIBID.Text
        
        cmdSP.ActiveConnection = ConnERP
        cmdSP.Execute

        Me.MousePointer = vbDefault
        MsgBox strMsgDeleteDataDone, vbInformation, strTitleInfo
        
        setForm True
        
        ClearForm
        
        SetButton "NO DATA"
        
        InitGrids
    End If
End Sub

Private Sub db_edit()
    If recMateriTemp.State = 0 Then
        MsgBox "IB Radio is Empty!", vbCritical, strApplication_Name
        Exit Sub
    End If
'cmdEdit_Click
    If recMateriTemp.RecordCount > 0 Then
        cmdEditPlan.Enabled = True
    Else
        cmdEditPlan.Enabled = False
    End If
    If lblApprove.Caption = "APPROVED" Then
            MsgBox "Can not edit data approved !", vbCritical, strApplication_Name
        Exit Sub
    End If
    If txtIBID.Text = "" Then
            MsgBox "Please Select a Brief to Edit !", vbInformation, strApplication_Name
        Exit Sub
    End If
    
    'Check apakah Planner Brand
    If Not IsValidAccess(strLogin_User, "Planner", Left(cboBrand.Text, 4)) Then
        MsgBox strMsgAccessDenied, vbCritical, strTitleExclamation
        Exit Sub
    End If
            
    'Lock Approval Frame
    fraClientApproval.Enabled = False
    
    If lblApprove.Caption = Not_Approved Then
        strTransProcess = "EDIT"
        
        SetButton "EDIT"
        
        Call setForm(False)
        Call SetButtonToolbar(False, picButton)
        cboYear.Enabled = False
        cboBrand.Enabled = False
        cboBrandVariant.Enabled = False
        cboStartingMonth.Enabled = False
        cboPlanDetail.Enabled = False
        
        'txtEnteredDate.Enabled = False
        
        optArea.Enabled = True
        optCity.Enabled = True
    Else
        MsgBox "This Implementation Brief Radio has been Approved by Client", vbCritical, strTitleInfo
    End If
End Sub

Private Sub cmdEditPlan_Click()
    If recMateriTemp.RecordCount > 0 Then
        
        cboPlanDetail.Enabled = False
        
        recMateriTemp.Filter = ""
        recRadioPlanTemp.Filter = ""
        recPlanDetailTemp.Filter = ""
        recPlanDetailMaterialTemp.Filter = ""
        
        Frm_IB_Radio_Plan_New.show vbModal
        
        LoadDetail txtIBID.Text, cboYear.Text, Get_Month_Number(cboStartingMonth.Text)
    Else
        MsgBox "Please add new material first !", vbExclamation, strApplication_Name
    End If
End Sub

Private Sub cmdMateri_Click()
    recMateriTemp.Filter = ""
    recRadioPlanTemp.Filter = ""
    recPlanDetailTemp.Filter = ""
    recPlanDetailMaterialTemp.Filter = ""
    
    Frm_IB_Radio_Materi.show vbModal
End Sub

Private Sub db_print()
'cmdPrint_Click
    If Trim(cboBrand.Text) = "" Then
        MsgBox "Please Select a Brand !", vbExclamation, strApplication_Name
        Exit Sub
    End If
   
    crIB.Reset
    crIB.SelectionFormula = "{ib_radio.ib_id}='" & txtIBID.Text & "'"
    
    strQuery = "SELECT IB_Radio.IB_Id, IB_Radio.Target_Primary, IB_Radio.Target_Secondary, IB_Radio.Consideration, IB_Radio.Attachment,IB_Radio_Plan_Detail.Month, IB_Radio_Plan_Detail.Spot, IB_Radio_Plan_Detail.Urban_Flag, IB_Radio_Plan_Detail.Rural_Flag,"
    strQuery = strQuery & "Brand.Brand_Name,IB_Radio_Plan.Budget,Month_Catalog.Month_Name, IB_Radio_Plan_Detail_Material.Schedule, IB_Radio_Plan_Detail_Material.City_Id, IB_Radio_Plan_Detail_Material.Material_Mix,City.City,IB_Radio_Material.Material_Name , IB_Radio_Material.Duration"
    strQuery = strQuery & " FROM erp.dbo.IB_Radio IB_Radio, erp.dbo.IB_Radio_Plan_Detail IB_Radio_Plan_Detail,  erp.dbo.Brand Brand,    erp.dbo.IB_Radio_Plan IB_Radio_Plan, erp.dbo.Month_Catalog Month_Catalog,erp.dbo.IB_Radio_Plan_Detail_Material IB_Radio_Plan_Detail_Material, erp.dbo.City City,erp.dbo.IB_Radio_Material IB_Radio_Material"
    
    strQuery = strQuery & " WHERE IB_Radio.IB_Id = IB_Radio_Plan_Detail.IB_Id AND"
    strQuery = strQuery & " IB_Radio.Brand_code = Brand.Brand_Code AND"
    strQuery = strQuery & " IB_Radio_Plan_Detail.IB_Id = IB_Radio_Plan.IB_Id AND"
    strQuery = strQuery & " IB_Radio_Plan_Detail.Month = IB_Radio_Plan.Month AND"
    strQuery = strQuery & " IB_Radio_Plan_Detail.Year = IB_Radio_Plan.Year AND"
    strQuery = strQuery & " IB_Radio_Plan_Detail.Month = Month_Catalog.Month AND"
    strQuery = strQuery & " IB_Radio_Plan_Detail.IB_Id = IB_Radio_Plan_Detail_Material.IB_Id AND"
    strQuery = strQuery & " IB_Radio_Plan_Detail.Month = IB_Radio_Plan_Detail_Material.Month AND"
    strQuery = strQuery & " IB_Radio_Plan_Detail.Year = IB_Radio_Plan_Detail_Material.Year AND"
    strQuery = strQuery & " IB_Radio_Plan_Detail.Schedule = IB_Radio_Plan_Detail_Material.Schedule AND"
    strQuery = strQuery & " IB_Radio_Plan_Detail.City_Id = IB_Radio_Plan_Detail_Material.City_Id AND"
    strQuery = strQuery & " IB_Radio_Plan_Detail_Material.City_Id = City.City_ID AND"
    strQuery = strQuery & " IB_Radio_Plan_Detail_Material.IB_Id = IB_Radio_Material.IB_ID AND"
    strQuery = strQuery & " IB_Radio_Plan_Detail_Material.material_id = IB_Radio_Material.material_id"
    strQuery = strQuery & " AND IB_Radio.Ib_ID='" & txtIBID.Text & "'"
    strQuery = strQuery & " ORDER BY IB_Radio_Plan_Detail.Month ASC,IB_Radio_Plan_Detail_Material.City_Id ASC,IB_Radio_Plan_Detail_Material.Schedule ASC"
    
    crIB.Connect = "DSN =" & Server_Name & ";UID = " & Login_User & ";DSQ = " & Database_Name & "; PWD =" & Login_Password
    
    If cboPlanDetail.Text = "City" Then
        If optArea.Value = True Then
            crIB.ReportFileName = Report_Dir & "\radio\IB_Radio_area.rpt"
        ElseIf optCity.Value = True Then
            crIB.ReportFileName = Report_Dir & "\radio\IB_Radio.rpt"
        End If
    Else
        'crIB.ReportFileName = App.Path & "\radio\IB_Radio_byStation.rpt"
        crIB.ReportFileName = Report_Dir & "\radio\IB_Radio_byStation.rpt"
    End If
    
    If Is_Special_Brand(Left(cboBrand.Text, 4)) = True Then
        crIB.Formulas(0) = "Marketing ='Marketing'"
    Else
        crIB.Formulas(0) = "Marketing ='Marketing'"
    End If
    
    crIB.WindowShowRefreshBtn = True
    crIB.WindowShowPrintSetupBtn = True
    crIB.WindowState = crptMaximized
    crIB.WindowTitle = " -- IB Radio -- "
    crIB.Action = 1
End Sub

Private Sub db_save()
'sebelumnya cmdSave_Click

    If Trim(txtMediaPlanNo.Text) = "" Then
        MsgBox "You must fill Media Plan Number", vbCritical, strApplication_Name
        txtMediaPlanNo.SetFocus
        Exit Sub
    End If
    
    If Trim(txtPrimaryTarget.Text) = "" Then
        MsgBox "You must fill Primary Target", vbCritical, strApplication_Name
        txtPrimaryTarget.SetFocus
        Exit Sub
    End If
    
    If Trim(cboSecondary.Text) = "" Then
        MsgBox "You must fill Secondary Target", vbCritical, strApplication_Name
        cboSecondary.SetFocus
        Exit Sub
    End If
    
    Call SaveData
    
    Me.MousePointer = vbDefault
    
    MsgBox strMsgSaveDataDone, vbInformation, strTitleInfo
    
    'Refresh Form
    ShowData txtIBID.Text
    
    PrepareTemp txtIBID.Text
    
    SetButton "SHOW"
    
    strTransProcess = "SHOW"
    
    Call setForm(True)
                
    cboYear.Enabled = True
    cboBrand.Enabled = True
    cboBrandVariant.Enabled = True
    cboPlanDetail.Enabled = False

    'Lock Approval Frame
    fraClientApproval.Enabled = True
    
    optArea.Enabled = True
    optCity.Enabled = True
    EnableObject False
    Exit Sub
    
Label_Error:
    MsgBox Err.Number & " " & Err.Description
End Sub

Private Sub db_Find()
'cmdSearch_Click
    If Trim(cboYear.Text) = "" Then
        MsgBox "Please Select Year !", vbInformation, strApplication_Name
        Exit Sub
    End If

    If Trim(cboBrand.Text) = "" Then
        MsgBox "Please Select A Brand !", vbInformation, strApplication_Name
        Exit Sub
    End If
        
    Frm_IB_Radio_Seach.show 1
End Sub

Private Sub SetButton(strTransac As String)
    If UCase(strTransac) = "ADD" Or UCase(strTransac) = "EDIT" Then
        'cmdAdd.Visible = False
        'cmdEdit.Visible = False
        'cmdSave.Visible = True
        'cmdSave.Enabled = True
        'cmdCancel.Visible = True
        'cmdCancel.Enabled = True
        
        'cmdDelete.Enabled = False
        'cmdPrint.Enabled = False
        'cmdClose.Enabled = False
        'cmdSearch.Enabled = False
    ElseIf UCase(strTransac) = "NO DATA" Then
        'cmdAdd.Visible = True
        'cmdAdd.Enabled = True
        'cmdEdit.Visible = True
        'cmdEdit.Enabled = False
        'cmdSave.Visible = False
        'cmdCancel.Visible = False
        
        'cmdDelete.Enabled = False
        'cmdPrint.Enabled = False
        'cmdClose.Enabled = True
        'cmdSearch.Enabled = True
    ElseIf UCase(strTransac) = "SHOW" Then
        'cmdAdd.Visible = True
        'cmdAdd.Enabled = True
        'cmdEdit.Visible = True
        'cmdSave.Visible = False
        'cmdCancel.Visible = False
        
        'cmdPrint.Enabled = True
        'cmdSearch.Enabled = True
        'cmdClose.Enabled = True
        
        If lblApprove.Caption = "UNAPPROVED" Then
            'cmdEdit.Enabled = True
            'cmdDelete.Enabled = True
        Else
            'cmdEdit.Enabled = False
            'cmdDelete.Enabled = False
        End If
    End If
End Sub

Private Sub setForm(blnIsLock As Boolean)
    dtpIBDate.Enabled = Not blnIsLock
    'txtEnteredDate.Enabled = Not blnIsLock
    
    cboStartingMonth.Enabled = Not blnIsLock
    cboSecondary.Enabled = Not blnIsLock
    cboPlanDetail.Enabled = Not blnIsLock 'dw
    
    txtMediaPlanNo.Enabled = Not blnIsLock
    txtPrimaryTarget.Enabled = Not blnIsLock
    txtAttachment.Enabled = Not blnIsLock
    txtAnyCons.Enabled = Not blnIsLock
    
    cmdMateri.Enabled = Not blnIsLock
    cmdEditPlan.Enabled = Not blnIsLock
End Sub

Private Function GenerateIBID() As String
    Rem Set Stored Procedure
    Dim recMediaType As New ADODB.Recordset
    Dim prmYearIN As New ADODB.Parameter
    Dim prmBrandCodeIN As New ADODB.Parameter
    Dim prmNewIBOut As New ADODB.Parameter
    Dim prmMediaTypeIN As New ADODB.Parameter
    Dim cmdIB As New ADODB.Command

    cmdIB.CommandType = adCmdStoredProc
    cmdIB.CommandText = "Get_New_IB_ID_Radio"
    
    Set prmBrandCodeIN = cmdIB.CreateParameter("Brand_Code", adChar, adParamInput, 4)
    Set prmMediaTypeIN = cmdIB.CreateParameter("Media_Type", adChar, adParamInput, 3)
    Set prmYearIN = cmdIB.CreateParameter("Year", adInteger, adParamInput)
    Set prmNewIBOut = cmdIB.CreateParameter("New_IB_Id", adChar, adParamOutput, 13)
    
    cmdIB.Parameters.Append prmBrandCodeIN
    cmdIB.Parameters.Append prmMediaTypeIN
    cmdIB.Parameters.Append prmYearIN
    cmdIB.Parameters.Append prmNewIBOut
    
    prmYearIN.Value = Val(cboYear.Text)
    prmBrandCodeIN.Value = Left(cboBrand.Text, 4)
    
    Rem Get Media Type Code
    strQuery = "SELECT Media_Type_Code FROM Media_Type WHERE Media_Type_Name ='Radio Media Induk'"
    recMediaType.Open strQuery, ConnERP, adOpenStatic, adLockReadOnly
    
    With recMediaType
        If Not .EOF Then
            prmMediaTypeIN.Value = .Fields("Media_Type_Code").Value
        End If
    End With
    
    recMediaType.Close
    Set recMediaType = Nothing
           
    cmdIB.ActiveConnection = ConnERP
    cmdIB.Execute
    
    GenerateIBID = prmNewIBOut.Value
End Function

Private Sub InitialDataMonth()
    Dim intPos As Integer
        
    cboStartingMonth.Clear
    cboPlanMonth.Clear
    
    For intPos = 1 To 12
        cboStartingMonth.AddItem NameMonth(intPos)
    Next intPos
    
    cboStartingMonth.ListIndex = month(Now) - 1
End Sub

Private Function NameMonth(ByVal intMonthStr As Integer) As String
    Select Case intMonthStr
        Case Is = 1
            NameMonth = "January"
        Case Is = 2
            NameMonth = "February"
        Case Is = 3
            NameMonth = "March"
        Case Is = 4
            NameMonth = "April"
        Case Is = 5
            NameMonth = "May"
        Case Is = 6
            NameMonth = "June"
        Case Is = 7
            NameMonth = "July"
        Case Is = 8
            NameMonth = "August"
        Case Is = 9
            NameMonth = "September"
        Case Is = 10
            NameMonth = "October"
        Case Is = 11
            NameMonth = "November"
        Case Is = 12
            NameMonth = "December"
    End Select
End Function

Public Sub ClearForm()
    lblLastModifiedDate.Caption = "Last Modified Date: "
    lblLastModifiedBy.Caption = "Last Modified By: "
    'txtEnteredBy.Text = ""
    txtIBID.Text = Empty
    txtMediaPlanNo.Text = Empty
    txtPrimaryTarget.Text = Empty
    txtAttachment.Text = Empty
    txtAnyCons.Text = Empty
    
    cboStartingMonth.Text = ""
    
    medGrandTotal = 0
    dbgCity.Clear
    dbgCity.Rows = 2
    msgMix.Clear
    msgMix.Rows = 2
    
    lblApprove.Caption = ""
    lblDateApp.Caption = ""
    lblTimeApp.Caption = ""
    lblStatus.Caption = ""
    lblStatus.ToolTipText = ""
End Sub

Public Sub ShowData(ByVal strIBID As String)
    Dim recIB As New ADODB.Recordset

    strQuery = "SELECT * FROM IB_Radio WHERE "
    strQuery = strQuery & "IB_Id='" & strIBID & "'"

    recIB.Open strQuery, ConnERP, adOpenStatic, adLockReadOnly

    With recIB
        Do While .EOF = False
            cboYear.Text = .Fields("year").Value
            cboBrandVariant.Text = IIf(IsNull(.Fields("Brand_Variant_Code").Value), "", .Fields("Brand_Variant_Code").Value & " --> " & .Fields("Brand_Variant_Name").Value)
            cboStartingMonth.Text = Get_Month_Name(.Fields("month").Value)
            cboSecondary.Text = .Fields("Target_Secondary").Value
            
            dtpIBDate.Value = IIf(IsNull(.Fields("Date").Value), "9/9/9999", .Fields("Date").Value)
            
            txtIBID.Text = Trim(.Fields("IB_ID").Value)
            txtMediaPlanNo.Text = .Fields("Media_Plan").Value
            txtPrimaryTarget.Text = .Fields("Target_Primary").Value
            'txtEnteredDate.Text = .Fields("Date_Entered").Value
            'txtEnteredBy.Text = .Fields("Entered_By").Value
            lblLastModifiedDate.Caption = "Last Modified Date: " & .Fields("Date_Entered").Value & " | "
            lblLastModifiedBy.Caption = "Last Modified By: " & .Fields("Entered_By").Value
            
            txtAttachment.Text = .Fields("Attachment").Value
            txtAnyCons.Text = .Fields("Consideration").Value

            intIBStartingMonth = .Fields("month").Value

            If .Fields("Approved_Flag").Value = 0 Then
                lblApprove.ForeColor = vbRed
                lblApprove.Caption = Not_Approved
                lblDateApp.Caption = ""
                lblTimeApp.Caption = ""
            ElseIf .Fields("Approved_Flag").Value = 1 Then
                lblApprove.ForeColor = vbBlack
                lblApprove.Caption = Approved
                lblDateApp.Caption = Format(.Fields("Approved_Date").Value, "dd/mm/yyyy hh:mm:ss AMPM")
                lblTimeApp.Caption = ""
            End If

            'Status
            If .Fields("Status").Value = 1 Then
                lblStatus.Caption = ""
                lblStatus.ToolTipText = ""
            Else
                lblStatus.Caption = "Canceled"
                lblStatus.ToolTipText = .Fields("Cancel_By").Value & " (" & .Fields("Cancel_Date").Value & ")"
            End If

            'dw - Showing Plan Detail Status
            '-------------------------------
            cboPlanDetail.Enabled = False
            If .Fields("IsCity").Value = 0 Then
                cboPlanDetail.Text = "Station"
            Else
                cboPlanDetail.Text = "City"
            End If
            '-------------------------------
            
            .MoveNext
        Loop
    End With

    If recIB.State = adStateOpen Then
        recIB.Close
    End If

    strQuery = "SELECT SUM(Budget) As Grand_Total FROM IB_Radio_Plan WHERE Ib_ID ='" & txtIBID.Text & "'"
    recIB.Open strQuery, ConnERP, adOpenStatic, adLockReadOnly

    With recIB
        If .EOF = False Then
            medGrandTotal.Text = Format(IIf(IsNull(.Fields(0).Value) = True, 0, .Fields(0).Value), "##0,0")
        End If
    End With
    
    Set recIB = Nothing

    monthIB cboPlanMonth
    
    strTransProcess = "SHOW"

    picButton(biePrint).Enabled = True
    'Call SetPictureT(bieEdit, True, picButton)
    Call SetPictureTBEnabled(biePrint, True)
    optArea.Enabled = True
    optCity.Enabled = True
End Sub

Private Sub CreateTabelMateriTemp()
    Set recMateriTemp = Nothing
    Set recMateriTemp = New ADODB.Recordset
    
    With recMateriTemp.Fields
        .Append "Client_Brief_Id", adChar, 10, adFldMayBeNull
        .Append "IB_Id", adChar, 13, adFldMayBeNull
        .Append "Material_ID", adChar, 1, adFldMayBeNull
        .Append "Material_Name", adVarChar, 50, adFldMayBeNull
        .Append "Duration", adVarChar, 5, adFldMayBeNull
    End With
    
    recMateriTemp.Open
End Sub

Public Sub PrepareTemp(ByVal strIBID As String)
    Dim intPos As Integer
    Dim recMaster As New ADODB.Recordset
    
    Rem Load Data FROM IB Radio Table
    If recMaster.State = adStateOpen Then
        recMaster.Close
    End If
    
    strQuery = "SELECT * FROM ib_radio WHERE "
    strQuery = strQuery & "Ib_Id='" & strIBID & "'"
    
    recMaster.Open strQuery, ConnERP, adOpenKeyset, adLockOptimistic
            
    Rem Prepare Tabel Temporary materi
    Call CreateTabelMateriTemp

    If recMateri.State = adStateOpen Then
        recMateri.Close
    End If

    strQuery = "SELECT * FROM IB_Radio_Material "
    strQuery = strQuery & "WHERE IB_ID ='" & strIBID & "'"

    recMateri.CursorLocation = adUseClient
    recMateri.Open strQuery, ConnERP, adOpenDynamic, adLockOptimistic

    'Load Materi Tabel data to temporary tabel
    With recMateri
        Do While .EOF = False
            recMateriTemp.AddNew
            
            For intPos = 0 To 4
                recMateriTemp.Fields(intPos).Value = .Fields(intPos).Value
            Next intPos
            
            recMateriTemp.Update
            .MoveNext
        Loop
    End With
    
    Rem Prepare Tabel Temporary Radio Plan
    Call CreateTabelRadioPlanTemp
    
    If recRadioPlan.State = adStateOpen Then
        recRadioPlan.Close
    End If
    
    strQuery = "SELECT * FROM IB_Radio_Plan "
    strQuery = strQuery & " WHERE IB_ID ='" & strIBID & "'"
   
    recRadioPlan.CursorLocation = adUseClient
    recRadioPlan.Open strQuery, ConnERP, adOpenDynamic, adLockOptimistic
    
    'Load Radio Plan Tabel data to temporary tabel
    With recRadioPlan
        Do While .EOF = False
            recRadioPlanTemp.AddNew
            
            For intPos = 0 To 4
                recRadioPlanTemp.Fields(intPos).Value = .Fields(intPos).Value
            Next intPos
            
            recRadioPlanTemp.Update
            .MoveNext
        Loop
    End With
    
    Rem Prepare Tabel Temporary Radio Plan Detail
    Call CreateTabelPlanDetailTemp
    
    Set recPlanDetail = Nothing
    
    If recPlanDetail.State = adStateOpen Then
        recPlanDetail.Close
    End If
    
    strQuery = "SELECT * FROM IB_Radio_Plan_Detail "
    strQuery = strQuery & " WHERE "
    strQuery = strQuery & " IB_ID ='" & strIBID & "'"
 
    recPlanDetail.CursorLocation = adUseClient
    recPlanDetail.Open strQuery, ConnERP, adOpenDynamic, adLockOptimistic
    
    'Load Radio Plan Detail Tabel data to temporary tabel
    With recPlanDetail
        Do While .EOF = False
            recPlanDetailTemp.AddNew
            
            For intPos = 0 To 9
                recPlanDetailTemp.Fields(intPos).Value = Trim(.Fields(intPos).Value)
            Next intPos
            
            recPlanDetailTemp.Update
            .MoveNext
        Loop
    End With
    
    CreateTabelPlanDetailMaterialTemp
    
    Set recPlanDetailMaterial = Nothing
    
    If recPlanDetailMaterial.State = adStateOpen Then
        recPlanDetailMaterial.Close
    End If
    
    strQuery = "SELECT * FROM IB_Radio_Plan_Detail_Material "
    strQuery = strQuery & " WHERE "
    strQuery = strQuery & " IB_ID ='" & strIBID & "'"
  
    recPlanDetailMaterial.CursorLocation = adUseClient
    recPlanDetailMaterial.Open strQuery, ConnERP, adOpenDynamic, adLockOptimistic
    
    'Load Radio Plan Detail Tabel data to temporary tabel
    With recPlanDetailMaterial
        Do While .EOF = False
            recPlanDetailMaterialTemp.AddNew
            
            For intPos = 0 To 7
                recPlanDetailMaterialTemp.Fields(intPos).Value = .Fields(intPos).Value
            Next intPos
            
            recPlanDetailMaterialTemp.Update
            .MoveNext
        Loop
    End With

    'Show Detail
    If Not recMaster.EOF Then
        LoadDetail strIBID, recMaster.Fields("Year").Value, recMaster.Fields("Month").Value
    End If
    
    recMaster.Close
    Set recMaster = Nothing
End Sub

Private Sub CreateTabelPlanDetailTemp()
    Set recPlanDetailTemp = Nothing
    Set recPlanDetailTemp = New ADODB.Recordset
    
    With recPlanDetailTemp.Fields
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
    
    recPlanDetailTemp.Open
End Sub

Private Sub CreateTabelPlanDetailMaterialTemp()
    Set recPlanDetailMaterialTemp = Nothing
    Set recPlanDetailMaterialTemp = New ADODB.Recordset
    
    With recPlanDetailMaterialTemp.Fields
        .Append "Client_Brief_Id", adChar, 10, adFldMayBeNull
        .Append "IB_Id", adChar, 13, adFldMayBeNull
        .Append "Month", adInteger, , adFldMayBeNull
        .Append "Year", adInteger, , adFldMayBeNull
        .Append "Schedule", adVarChar, 75, adFldMayBeNull
        .Append "City_Code", adInteger, , adFldMayBeNull
        .Append "Material_Id", adChar, 1, adFldMayBeNull
        .Append "Material_Mix", adDouble, , adFldMayBeNull
    End With
    
    recPlanDetailMaterialTemp.Open
End Sub

Private Sub CreateTabelRadioPlanTemp()
    Set recRadioPlanTemp = Nothing
    Set recRadioPlanTemp = New ADODB.Recordset
    
    With recRadioPlanTemp.Fields
        .Append "Client_Brief_Id", adChar, 10, adFldMayBeNull
        .Append "IB_Id", adChar, 13, adFldMayBeNull
        .Append "Month", adInteger, , adFldMayBeNull
        .Append "Year", adInteger, , adFldMayBeNull
        .Append "Budget", adCurrency, , adFldMayBeNull
    End With
    
    recRadioPlanTemp.Open
End Sub

Private Sub LoadDetail(ByVal strIBID As String, ByVal IntYear As Integer, intMonth As Integer)
    Dim intPosRow As Integer
    Dim recIBPlan As New ADODB.Recordset
    Dim StrFilter As String
    Dim strFilterMateri As String
    Dim intRowMaterial As Integer
    Dim strAreaCode As String
    Dim blnWrite As Boolean
    Dim strMaterialId As String
    
    Dim blnIsStation As Boolean 'dw
    Dim recCek As New ADODB.Recordset 'dw

    'dw - cek if ib is by station or not..
    recCek.Open "SELECT IsCity FROM IB_Radio WHERE ib_id='" & strIBID & "'", ConnERP, adOpenStatic, adLockReadOnly
    
    If Not recCek.EOF And Not recCek.BOF Then
        If recCek.Fields("IsCity").Value = 1 Then
            blnIsStation = True
        Else
            blnIsStation = False
        End If
    End If
    '---------------------------------------
    
    If strTransProcess = "ADD" Or strTransProcess = "EDIT" Then
        StrFilter = " month =" & intMonth
        StrFilter = StrFilter & " and Year =" & IntYear
        StrFilter = StrFilter & " and IB_ID ='" & strIBID & "'"

        'Load Plan Detail Data Temp7
        If Frm_IB_Radio.recRadioPlanTemp.RecordCount > 0 Then
           Frm_IB_Radio.recRadioPlanTemp.MoveFirst
        End If
        intPosRow = 0

        
        dbgCity.Rows = 1
        
        If optCity.Value Then
            'view by city(Edit)
                'Load City Temp
            If Frm_IB_Radio.recPlanDetailTemp.RecordCount > 0 Then
                dbgCity.Rows = 1
                dbgCity.cols = 7
                dbgCity.TextMatrix(0, 0) = "SCHEDULE"
                dbgCity.TextMatrix(0, 1) = "SPOT/DAY"
                dbgCity.TextMatrix(0, 2) = "SALES AREA"
                dbgCity.TextMatrix(0, 3) = "URBAN AREA"
                dbgCity.TextMatrix(0, 4) = "RURAL AREA"
                dbgCity.TextMatrix(0, 5) = "MATERIAL ID"
                dbgCity.TextMatrix(0, 6) = "MATERIAL MIX"
                                
                dbgCity.ColWidth(0) = 2000
                dbgCity.ColWidth(1) = 1000
                dbgCity.ColWidth(2) = 1300
                dbgCity.ColWidth(3) = 1000
                dbgCity.ColWidth(4) = 1000
                dbgCity.ColWidth(5) = 1200
                dbgCity.ColWidth(6) = 1300
                
                Frm_IB_Radio.recPlanDetailTemp.MoveFirst
            End If
        
            intPosRow = 0
            With Frm_IB_Radio.recPlanDetailTemp
                .Filter = ""
                .Filter = StrFilter
                Do While .EOF = False
                    intRowMaterial = 0
                    intPosRow = intPosRow + 1
                    dbgCity.Rows = intPosRow + 1
                    
                    If Frm_IB_Radio.cboPlanDetail.Text = "City" Then 'dw
                        dbgCity.TextMatrix(intPosRow, 0) = .Fields("Schedule").Value
                    ElseIf Frm_IB_Radio.cboPlanDetail.Text = "Station" Then
                        dbgCity.TextMatrix(intPosRow, 0) = Trim(Mid(.Fields("Schedule").Value, 1, InStr(.Fields("Schedule").Value, "-") - 1)) 'dw
                    End If
                    
                    dbgCity.TextMatrix(intPosRow, 1) = .Fields("Spot").Value
                    dbgCity.TextMatrix(intPosRow, 2) = GetCityName(Trim(.Fields("City_Code").Value))
                    dbgCity.TextMatrix(intPosRow, 3) = IIf(.Fields("Urban_Flag").Value = 1, "x", "")
                    dbgCity.TextMatrix(intPosRow, 4) = IIf(.Fields("rural_Flag").Value = 1, "x", "")
                    
                    'get filter string
                    strFilterMateri = StrFilter & " AND Schedule ='" & Clear_String(.Fields("Schedule").Value) & "'"
                    strFilterMateri = strFilterMateri & " AND City_Code = " & .Fields("City_Code").Value
                    
                    recPlanDetailMaterialTemp.Filter = ""
                    recPlanDetailMaterialTemp.Filter = strFilterMateri
                    
                    While Not recPlanDetailMaterialTemp.EOF
                        If intRowMaterial = 0 Then
                            dbgCity.TextMatrix(intPosRow, 5) = recPlanDetailMaterialTemp.Fields("Material_Id").Value
                            dbgCity.TextMatrix(intPosRow, 6) = recPlanDetailMaterialTemp.Fields("Material_Mix").Value
                        Else
                            intPosRow = intPosRow + 1
                            dbgCity.Rows = intPosRow + 1
                            
                            'dbgCity.TextMatrix(intPosRow, 0) = .Fields("Schedule").Value
                            dbgCity.TextMatrix(intPosRow, 0) = Trim(Mid(.Fields("Schedule").Value, 1, InStr(.Fields("Schedule").Value, "-") - 1))
                            dbgCity.TextMatrix(intPosRow, 1) = .Fields("Spot").Value
                            
                            dbgCity.TextMatrix(intPosRow, 2) = GetCityName(Trim(.Fields("City_Code").Value))
                            dbgCity.TextMatrix(intPosRow, 3) = IIf(.Fields("Urban_Flag").Value = 1, "x", "")
                            dbgCity.TextMatrix(intPosRow, 4) = IIf(.Fields("rural_Flag").Value = 1, "x", "")
                            
                            dbgCity.TextMatrix(intPosRow, 5) = recPlanDetailMaterialTemp.Fields("Material_Id").Value
                            dbgCity.TextMatrix(intPosRow, 6) = recPlanDetailMaterialTemp.Fields("Material_Mix").Value
                        End If
                        
                        intRowMaterial = intRowMaterial + 1
                        recPlanDetailMaterialTemp.MoveNext
                    Wend
                    
                    .MoveNext
                Loop
            End With
        Else
            'view by area (Edit)
            If Frm_IB_Radio.recPlanDetailTemp.RecordCount > 0 Then
                dbgCity.Rows = 1
                dbgCity.cols = 5
                dbgCity.TextMatrix(0, 0) = "SCHEDULE"
                dbgCity.TextMatrix(0, 1) = "SPOT/DAY"
                dbgCity.TextMatrix(0, 2) = "SALES AREA"
                dbgCity.TextMatrix(0, 3) = "MATERIAL ID"
                dbgCity.TextMatrix(0, 4) = "MATERIAL MIX"
                    
                dbgCity.ColWidth(0) = 2000
                dbgCity.ColWidth(1) = 1000
                dbgCity.ColWidth(2) = 2500
                dbgCity.ColWidth(3) = 1200
                dbgCity.ColWidth(4) = 1700
                
                Frm_IB_Radio.recPlanDetailTemp.MoveFirst
            End If
        
            intPosRow = 0
            With Frm_IB_Radio.recPlanDetailTemp
                .Filter = ""
                .Filter = StrFilter
                Do While .EOF = False
                    intRowMaterial = 0
                    If intPosRow = 0 Then
                        blnWrite = True
                        intPosRow = intPosRow + 1
                        dbgCity.Rows = intPosRow + 1
                        
                        'put to temp
                        strAreaCode = .Fields("Area_Code").Value
                        
                        'untuk rec pertama langsung tulis areanya
                        If Frm_IB_Radio.cboPlanDetail.Text = "City" Then
                            dbgCity.TextMatrix(intPosRow, 0) = .Fields("Schedule").Value
                        ElseIf Frm_IB_Radio.cboPlanDetail.Text = "Station" Then
                            dbgCity.TextMatrix(intPosRow, 0) = Trim(Mid(.Fields("Schedule").Value, 1, InStr(.Fields("Schedule").Value, "-") - 1)) 'dw
                        End If
                        
                        dbgCity.TextMatrix(intPosRow, 1) = .Fields("Spot").Value
                        dbgCity.TextMatrix(intPosRow, 2) = .Fields("Area_Code").Value
                    Else
                        If strAreaCode = .Fields("Area_Code").Value Then
                            blnWrite = False
                            GoTo NextRec
                        Else
                            blnWrite = True
                            intPosRow = intPosRow + 1
                            dbgCity.Rows = intPosRow + 1
                            
                            'put to temp
                            strAreaCode = .Fields("Area_Code").Value
                            
                            'untuk rec pertama langsung tulis areanya
                            'dbgCity.TextMatrix(intPosRow, 0) = .Fields("Schedule").Value
                            dbgCity.TextMatrix(intPosRow, 0) = Trim(Mid(.Fields("Schedule").Value, 1, InStr(.Fields("Schedule").Value, "-") - 1)) 'dw
                            dbgCity.TextMatrix(intPosRow, 1) = .Fields("Spot").Value
                            dbgCity.TextMatrix(intPosRow, 2) = .Fields("Area_Code").Value
                        End If
                    End If
                    
                    If blnWrite Then
                        'get filter string
                        strFilterMateri = StrFilter & " AND Schedule ='" & Clear_String(.Fields("Schedule").Value) & "'"
                        
                        recPlanDetailMaterialTemp.Filter = ""
                        recPlanDetailMaterialTemp.Filter = strFilterMateri
                        
                        While Not recPlanDetailMaterialTemp.EOF
                            If intRowMaterial = 0 Then
                                strMaterialId = recPlanDetailMaterialTemp.Fields("Material_Id").Value & "|" & recPlanDetailMaterialTemp.Fields("Material_Mix").Value
                                
                                dbgCity.TextMatrix(intPosRow, 3) = recPlanDetailMaterialTemp.Fields("Material_Id").Value
                                dbgCity.TextMatrix(intPosRow, 4) = recPlanDetailMaterialTemp.Fields("Material_Mix").Value
                            Else
                                If InStr(1, strMaterialId, recPlanDetailMaterialTemp.Fields("Material_Id").Value & "|" & recPlanDetailMaterialTemp.Fields("Material_Mix").Value) > 0 Then
                                    'do nothing
                                Else
                                    strMaterialId = strMaterialId & "|" & recPlanDetailMaterialTemp.Fields("Material_Id").Value & "|" & recPlanDetailMaterialTemp.Fields("Material_Mix").Value
                                    intPosRow = intPosRow + 1
                                    dbgCity.Rows = intPosRow + 1
                                    
                                    If Me.cboPlanDetail.Text = "City" Then
                                        dbgCity.TextMatrix(intPosRow, 0) = .Fields("Schedule").Value
                                    Else
                                        dbgCity.TextMatrix(intPosRow, 0) = Trim(Mid(.Fields("Schedule").Value, 1, InStr(.Fields("Schedule").Value, "-") - 1))
                                    End If
                                    dbgCity.TextMatrix(intPosRow, 1) = .Fields("Spot").Value
                                    dbgCity.TextMatrix(intPosRow, 2) = .Fields("Area_Code").Value
                                    dbgCity.TextMatrix(intPosRow, 3) = recPlanDetailMaterialTemp.Fields("Material_Id").Value
                                    dbgCity.TextMatrix(intPosRow, 4) = recPlanDetailMaterialTemp.Fields("Material_Mix").Value
                                
                                End If
                            End If
                            
                            intRowMaterial = intRowMaterial + 1
                            recPlanDetailMaterialTemp.MoveNext
                        Wend
                    End If
                    
NextRec:            .MoveNext
                Loop
            End With
        End If
    Else
        Rem Jika hanya lihat data (tidak meng-edit)
        'Load IB Radio _plan detail
        strQuery = "SELECT * FROM Ib_radio_plan WHERE "
        strQuery = strQuery & " ib_id ='" & strIBID & "'"
        strQuery = strQuery & " AND Month =" & intMonth
        strQuery = strQuery & " AND Year =" & IntYear
        
        recIBPlan.CursorLocation = adUseClient
        recIBPlan.Open strQuery, ConnERP, adOpenStatic, adLockReadOnly
        '----TrueDB Grid
        
        Set tdgPlan.DataSource = Nothing
        tdgPlan.ClearFields
        tdgPlan.DataSource = recIBPlan
        tdgPlan.Columns(0).Visible = False
        tdgPlan.Columns(1).Visible = False
        tdgPlan.Columns(2).Visible = False
        tdgPlan.Columns(3).Visible = False
        tdgPlan.Columns(4).NumberFormat = "###,####,####,###"
        tdgPlan.Columns(4).Caption = "BUDGET"
        tdgPlan.Columns(4).Width = 1500
        '----TrueDB Grid
'        msgPlan.Rows = recIBPlan.RecordCount + 1
'        msgPlan.cols = 1
'
'        msgPlan.TextMatrix(0, 0) = "BUDGET"
'        msgPlan.ColWidth(0) = 1500
        
'        intPosRow = 0
'        With recIBPlan
'            Do While .EOF = False
'                intPosRow = intPosRow + 1
'                msgPlan.TextMatrix(intPosRow, 0) = Format(.Fields("Budget").Value, "##,0")
'                .MoveNext
'            Loop
'        End With
        
'        If recIBPlan.State = adStateOpen Then
'            recIBPlan.Close
'        End If
        
        Set recIBPlan = Nothing
        
        If optCity.Value Then
            'view by city
                'Load Sales Area
            strQuery = "SELECT dbo.IB_Radio_Plan_Detail.Schedule, dbo.IB_Radio_Plan_Detail.Area_Code, dbo.IB_Radio_Plan_Detail.Spot, dbo.City.City, dbo.IB_Radio_Plan_Detail.Urban_Flag, "
            strQuery = strQuery & " dbo.IB_Radio_Plan_Detail.Rural_Flag , dbo.IB_Radio_Plan_Detail_Material.material_id, dbo.IB_Radio_Plan_Detail_Material.Material_Mix"
            strQuery = strQuery & " FROM dbo.City INNER JOIN"
            strQuery = strQuery & " dbo.IB_Radio_Plan_Detail ON dbo.City.City_ID = dbo.IB_Radio_Plan_Detail.City_Id INNER JOIN"
            strQuery = strQuery & " dbo.IB_Radio_Plan_Detail_Material ON dbo.IB_Radio_Plan_Detail.Client_Brief_id = dbo.IB_Radio_Plan_Detail_Material.Client_Brief_Id AND"
            strQuery = strQuery & " dbo.IB_Radio_Plan_Detail.IB_Id = dbo.IB_Radio_Plan_Detail_Material.IB_Id AND"
            strQuery = strQuery & " dbo.IB_Radio_Plan_Detail.[Month] = dbo.IB_Radio_Plan_Detail_Material.[Month] AND"
            strQuery = strQuery & " dbo.IB_Radio_Plan_Detail.[Year] = dbo.IB_Radio_Plan_Detail_Material.[Year] AND"
            strQuery = strQuery & " dbo.IB_Radio_Plan_Detail.Schedule = dbo.IB_Radio_Plan_Detail_Material.Schedule AND"
            strQuery = strQuery & " dbo.IB_Radio_Plan_Detail.City_Id = dbo.IB_Radio_Plan_Detail_Material.City_Id"
            strQuery = strQuery & " WHERE "
            strQuery = strQuery & " dbo.IB_Radio_Plan_Detail.ib_id ='" & strIBID & "'"
            strQuery = strQuery & " AND dbo.IB_Radio_Plan_Detail.Month =" & intMonth
            strQuery = strQuery & " AND dbo.IB_Radio_Plan_Detail.Year =" & IntYear
            strQuery = strQuery & " ORDER BY dbo.IB_Radio_Plan_Detail.Schedule, dbo.City.City, dbo.IB_Radio_Plan_Detail_Material.Material_Id"
                   
            recIBPlan.CursorLocation = adUseClient
            recIBPlan.Open strQuery, ConnERP, adOpenStatic, adLockReadOnly
            
            dbgCity.Rows = recIBPlan.RecordCount + 1
            dbgCity.cols = 7
                                   
            dbgCity.TextMatrix(0, 0) = "SCHEDULE"
            dbgCity.TextMatrix(0, 1) = "SPOT/DAY"
            dbgCity.TextMatrix(0, 2) = "SALES AREA"
            dbgCity.TextMatrix(0, 3) = "URBAN AREA"
            dbgCity.TextMatrix(0, 4) = "RURAL AREA"
            dbgCity.TextMatrix(0, 5) = "MATERIAL ID"
            dbgCity.TextMatrix(0, 6) = "MATERIAL MIX"
                                
            dbgCity.ColWidth(0) = 2000
            dbgCity.ColWidth(1) = 1000
            dbgCity.ColWidth(2) = 1300
            dbgCity.ColWidth(3) = 1000
            dbgCity.ColWidth(4) = 1000
            dbgCity.ColWidth(5) = 1200
            dbgCity.ColWidth(6) = 1300
                
            intPosRow = 0
            With recIBPlan
                Do While .EOF = False
                    intPosRow = intPosRow + 1
                    If Frm_IB_Radio.cboPlanDetail.Text = "City" Then 'dw
                        dbgCity.TextMatrix(intPosRow, 0) = IIf(IsNull(.Fields("Schedule").Value), "", .Fields("Schedule").Value)
                    ElseIf Frm_IB_Radio.cboPlanDetail.Text = "Station" Then
                        dbgCity.TextMatrix(intPosRow, 0) = IIf(IsNull(.Fields("Schedule").Value), "", Trim(Mid(.Fields("Schedule").Value, 1, InStr(.Fields("Schedule").Value, "-") - 1))) 'dw
                    End If
                    
                    dbgCity.TextMatrix(intPosRow, 1) = IIf(IsNull(.Fields("Spot").Value), "", .Fields("Spot").Value)
                    dbgCity.TextMatrix(intPosRow, 2) = IIf(IsNull(.Fields("City").Value), "", .Fields("City").Value)
                    dbgCity.TextMatrix(intPosRow, 3) = IIf(.Fields("Urban_Flag").Value = 1, "x", "")
                    dbgCity.TextMatrix(intPosRow, 4) = IIf(.Fields("rural_Flag").Value = 1, "x", "")
                    dbgCity.TextMatrix(intPosRow, 5) = IIf(IsNull(.Fields("Material_Id").Value), "", .Fields("Material_Id").Value)
                    dbgCity.TextMatrix(intPosRow, 6) = FormatNumber(IIf(IsNull(.Fields("Material_Mix").Value), "0", .Fields("Material_Mix").Value), 0)
                    .MoveNext
                Loop
            End With
        Else
            'view by area
            strQuery = "SELECT DISTINCT A.Schedule, A.Spot, A.Area_Code, B.Material_Id, B.Material_Mix"
            strQuery = strQuery & " FROM dbo.IB_Radio_Plan_Detail A LEFT OUTER JOIN"
            strQuery = strQuery & " dbo.IB_Radio_Plan_Detail_Material B ON A.Client_Brief_id = B.Client_Brief_Id AND A.IB_Id = B.IB_Id AND A.[Month] = B.[Month] AND"
            strQuery = strQuery & " a.[Year] = b.[Year] AND a.Schedule = b.Schedule AND a.City_Id = b.City_Id"
            strQuery = strQuery & " WHERE "
            strQuery = strQuery & " A.ib_id ='" & strIBID & "'"
            strQuery = strQuery & " AND A.Month =" & intMonth
            strQuery = strQuery & " AND A.Year =" & IntYear
            
            recIBPlan.CursorLocation = adUseClient
            recIBPlan.Open strQuery, ConnERP, adOpenStatic, adLockReadOnly
            dbgCity.Row = 0
            dbgCity.BackColorFixed = &HBC8A47
            dbgCity.Rows = recIBPlan.RecordCount + 1
            dbgCity.cols = 5
                                    
            dbgCity.TextMatrix(0, 0) = "SCHEDULE"
            dbgCity.TextMatrix(0, 1) = "SPOT/DAY"
            dbgCity.TextMatrix(0, 2) = "SALES AREA"
            dbgCity.TextMatrix(0, 3) = "MATERIAL ID"
            dbgCity.TextMatrix(0, 4) = "MATERIAL MIX"
                   
            dbgCity.ColWidth(0) = 2000
            dbgCity.ColWidth(1) = 1000
            dbgCity.ColWidth(2) = 2500
            dbgCity.ColWidth(3) = 1200
            dbgCity.ColWidth(4) = 1700
                        
            intPosRow = 0
            With recIBPlan
                Do While .EOF = False
                    intPosRow = intPosRow + 1
                    
                    If Frm_IB_Radio.cboPlanDetail.Text = "City" Then 'dw
                        dbgCity.TextMatrix(intPosRow, 0) = IIf(IsNull(.Fields("Schedule").Value), "", .Fields("Schedule").Value)
                    ElseIf Frm_IB_Radio.cboPlanDetail.Text = "Station" Then
                        dbgCity.TextMatrix(intPosRow, 0) = IIf(IsNull(.Fields("Schedule").Value), "", Trim(Mid(.Fields("Schedule").Value, 1, InStr(.Fields("Schedule").Value, "-") - 1))) 'dw
                    End If
                    
                    dbgCity.TextMatrix(intPosRow, 1) = IIf(IsNull(.Fields("Spot").Value), "", .Fields("Spot").Value)
                    dbgCity.TextMatrix(intPosRow, 2) = IIf(IsNull(.Fields("Area_Code").Value), "", .Fields("Area_Code").Value)
                    dbgCity.TextMatrix(intPosRow, 3) = IIf(IsNull(.Fields("Material_Id").Value), "", .Fields("Material_Id").Value)
                    dbgCity.TextMatrix(intPosRow, 4) = FormatNumber(IIf(IsNull(.Fields("Material_Mix").Value), "0", .Fields("Material_Mix").Value), 0)
                    .MoveNext
                Loop
            End With
        End If
        
        If recIBPlan.State = adStateOpen Then
            recIBPlan.Close
        End If
        
        Set recIBPlan = Nothing
    End If
    
    'dw - Close recCek
    recCek.Close
    Set recCek = Nothing
    '--------------------
    
    recMateriTemp.Filter = ""
    recRadioPlanTemp.Filter = ""
    recPlanDetailTemp.Filter = ""
    recPlanDetailMaterialTemp.Filter = ""
End Sub

Private Function GetCityName(strCityCode As String) As String
    Dim recCity As New ADODB.Recordset
    
    strQuery = "SELECT City FROM city WHERE city_id='" & Trim(strCityCode) & "'"
    recCity.Open strQuery, ConnERP, adOpenStatic, adLockPessimistic
    
    With recCity
        Do While .EOF = False
            GetCityName = Trim(.Fields("City").Value)
            
            If recCity.State = adStateOpen Then
                recCity.Close
            End If
            
            Set recCity = Nothing
            Exit Function
            .MoveNext
        Loop
    End With
    
    If recCity.State = adStateOpen Then
        recCity.Close
    End If
    
    Set recCity = Nothing
End Function

Private Sub SaveData()
    Dim recSaveIB As New ADODB.Recordset
    Dim vntBrandVariant As Variant
    Dim strQueryDelete As String
    
    strQuery = "SELECT * FROM IB_Radio WHERE IB_Id ='" & Trim(txtIBID.Text) & "'"
    recSaveIB.Open strQuery, ConnERP, adOpenDynamic, adLockOptimistic
    
    If strTransProcess = "ADD" Then
        Rem IF New Data
        'Save To IB_Radio Tabel
        With recSaveIB
             recDate.Requery
            .AddNew
            .Fields("client_brief_id").Value = ""
            .Fields("Ib_id").Value = Trim(txtIBID.Text)
            .Fields("Revision").Value = 0
            .Fields("month").Value = Get_Month_Number(cboStartingMonth.Text)
            .Fields("Year").Value = Val(cboYear.Text)
            .Fields("Date").Value = dtpIBDate.Value
            .Fields("Date_Entered").Value = recDate.Fields(0).Value
            .Fields("Entered_By").Value = strLogin_FullName
            .Fields("Brand_Code").Value = Left(cboBrand.Text, 4)
            .Fields("Target_Primary").Value = txtPrimaryTarget.Text
            .Fields("Target_Secondary").Value = cboSecondary.Text
            
            'Cluster_Code
            .Fields("Cluster_Code").Value = Mid(cboSecondary.Text, InStr(cboSecondary.Text, " - ") - 1, 1)
            .Fields("Media_Plan").Value = Trim(txtMediaPlanNo.Text)
            
            'Plan_No
            .Fields("Plan_No").Value = Trim(txtMediaPlanNo.Text)
            
            If Trim(cboBrandVariant.Text) <> "" Then
                vntBrandVariant = Split(cboBrandVariant.Text, "-->")
                
                'Brand_Variant_Code
                .Fields("Brand_Variant_Code").Value = Trim(vntBrandVariant(0))
                
                'Brand_Variant_Name
                .Fields("Brand_Variant_Name").Value = Trim(vntBrandVariant(1))
            End If
            
            .Fields("Attachment").Value = txtAttachment.Text
            .Fields("consideration").Value = txtAnyCons.Text
            .Fields("Approved_Flag").Value = 0
            .Fields("IsCity").Value = IIf(cboPlanDetail.Text = "City", 1, 0)
            
            .Update
        End With
                
        If recSaveIB.State = adStateOpen Then
            recSaveIB.Close
        End If
        
        Set recSaveIB = Nothing
                
        'Save To IB_Radio_Plan Tabel
        With recRadioPlan
            If recRadioPlanTemp.RecordCount > 0 Then
                recRadioPlanTemp.Filter = ""
                recRadioPlanTemp.MoveFirst
            End If
            
            Do While recRadioPlanTemp.EOF = False
                .AddNew
                .Fields("Client_Brief_ID").Value = ""
                .Fields("IB_ID").Value = recRadioPlanTemp.Fields("IB_ID").Value
                .Fields("Month").Value = recRadioPlanTemp.Fields("Month").Value
                .Fields("Year").Value = recRadioPlanTemp.Fields("Year").Value
                .Fields("Budget").Value = Format(recRadioPlanTemp.Fields("Budget").Value, "####0")
                .Update
                recRadioPlanTemp.MoveNext
            Loop
        End With
        
        If recRadioPlan.State = adStateOpen Then
            recRadioPlan.Close
        End If
           
        'Save To IB_Radio_Plan_Detail Tabel
        With recPlanDetail
            If recPlanDetailTemp.RecordCount > 0 Then
                recPlanDetailTemp.Filter = ""
                recPlanDetailTemp.MoveFirst
            End If
            
            Do While recPlanDetailTemp.EOF = False
                .AddNew
                .Fields("Client_Brief_ID").Value = ""
                .Fields("IB_ID").Value = recPlanDetailTemp.Fields("IB_ID").Value
                .Fields("Month").Value = recPlanDetailTemp.Fields("Month").Value
                .Fields("Year").Value = recPlanDetailTemp.Fields("Year").Value
                .Fields("City_id").Value = recPlanDetailTemp.Fields("City_code").Value
                .Fields("Schedule").Value = recPlanDetailTemp.Fields("Schedule").Value
                .Fields("Area_Code").Value = recPlanDetailTemp.Fields("Area_Code").Value
                .Fields("Spot").Value = recPlanDetailTemp.Fields("Spot").Value
                .Fields("Urban_Flag").Value = IIf(IsNull(recPlanDetailTemp.Fields("Urban_Flag").Value) = True, 0, recPlanDetailTemp.Fields("Urban_Flag").Value)
                .Fields("Rural_Flag").Value = IIf(IsNull(recPlanDetailTemp.Fields("Rural_Flag").Value) = True, 0, recPlanDetailTemp.Fields("Rural_Flag").Value)
                .Update
                recPlanDetailTemp.MoveNext
            Loop
        End With
        
        If recPlanDetail.State = adStateOpen Then
            recPlanDetail.Close
        End If
        
        Set recPlanDetail = Nothing
        
        'Save To IB_Radio_Plan_Detail_material Tabel
        With recPlanDetailMaterial
            If recPlanDetailMaterialTemp.RecordCount > 0 Then
                recPlanDetailMaterialTemp.Filter = ""
                recPlanDetailMaterialTemp.MoveFirst
            End If
            
            Do While recPlanDetailMaterialTemp.EOF = False
                .AddNew
                .Fields("Client_Brief_ID").Value = ""
                .Fields("IB_ID").Value = recPlanDetailMaterialTemp.Fields("IB_ID").Value
                .Fields("Month").Value = recPlanDetailMaterialTemp.Fields("Month").Value
                .Fields("Year").Value = recPlanDetailMaterialTemp.Fields("Year").Value
                .Fields("City_id").Value = recPlanDetailMaterialTemp.Fields("City_code").Value
                .Fields("Schedule").Value = recPlanDetailMaterialTemp.Fields("Schedule").Value
                .Fields("Material_Id").Value = recPlanDetailMaterialTemp.Fields("Material_Id").Value
                .Fields("Material_Mix").Value = recPlanDetailMaterialTemp.Fields("Material_Mix").Value
                
                .Update
                recPlanDetailMaterialTemp.MoveNext
            Loop
        End With
        
        If recPlanDetailMaterial.State = adStateOpen Then
            recPlanDetailMaterial.Close
        End If
        
        Set recPlanDetailMaterial = Nothing
                       
        'Save To IB_Radio_Material Table
        With recMateri
            If recMateriTemp.RecordCount > 0 Then
                recMateriTemp.MoveFirst
            End If
            
            Do While recMateriTemp.EOF = False
                .AddNew
                .Fields("Client_Brief_ID").Value = ""
                .Fields("IB_ID").Value = recMateriTemp.Fields("IB_ID").Value
                .Fields("Material_ID").Value = recMateriTemp.Fields("Material_ID").Value
                .Fields("material_name").Value = recMateriTemp.Fields("material_name").Value
                .Fields("Duration").Value = recMateriTemp.Fields("Duration").Value
                .Update
                recMateriTemp.MoveNext
            Loop
        End With
        
        If recMateri.State = adStateOpen Then
            recMateri.Close
        End If
        
    ElseIf strTransProcess = "EDIT" Then
        'Save To IB_Radio Tabel for Edit Process
        With recSaveIB
            recDate.Requery
            .Fields("Date_Entered").Value = recDate.Fields(0).Value
            .Fields("Entered_By").Value = strLogin_FullName
            
            'Plan_No
            .Fields("Plan_No").Value = txtMediaPlanNo.Text
            .Fields("Media_Plan").Value = txtMediaPlanNo.Text
            .Fields("Target_Primary").Value = txtPrimaryTarget.Text
            .Fields("Target_Secondary").Value = cboSecondary.Text
            .Fields("Date") = dtpIBDate.Value
            
            'Cluster_Code
            .Fields("Cluster_Code").Value = Mid(cboSecondary.Text, InStr(cboSecondary.Text, " - ") - 1, 1)
            
            'Plan_No
            .Fields("Plan_No").Value = Trim(txtMediaPlanNo.Text)
             
            If Trim(cboBrandVariant.Text) <> "" Then
                vntBrandVariant = Split(cboBrandVariant.Text, "-->")
                
                'Brand_Variant_Code
                .Fields("Brand_Variant_Code").Value = Trim(vntBrandVariant(0))
                
                'Brand_Variant_Name
                .Fields("Brand_Variant_Name").Value = Trim(vntBrandVariant(1))
            End If
                                    
            .Fields("Attachment").Value = txtAttachment.Text
            .Fields("consideration").Value = txtAnyCons.Text
            .Fields("IsCity").Value = IIf(cboPlanDetail.Text = "City", 1, 0)
    
            .Update
        End With
        
        If recSaveIB.State = adStateOpen Then
            recSaveIB.Close
        End If
        
        Set recSaveIB = Nothing
        
        strQueryDelete = "DELETE FROM ib_radio_material WHERE ib_id='" & txtIBID.Text & "'"
        ConnERP.Execute strQueryDelete
        
        strQueryDelete = "DELETE FROM ib_radio_plan_detail_material WHERE ib_id='" & txtIBID.Text & "'"
        ConnERP.Execute strQueryDelete
        
        strQueryDelete = "DELETE FROM ib_radio_plan_detail WHERE ib_id='" & txtIBID.Text & "'"
        ConnERP.Execute strQueryDelete
        
        strQueryDelete = "DELETE FROM ib_radio_plan WHERE ib_id='" & txtIBID.Text & "'"
        ConnERP.Execute strQueryDelete
        
        'Save To IB_Radio_Plan Tabel
        With recRadioPlan
            If recRadioPlanTemp.RecordCount > 0 Then
                recRadioPlanTemp.Filter = ""
                recRadioPlanTemp.MoveFirst
            End If
            
            Do While recRadioPlanTemp.EOF = False
                .AddNew
                .Fields("Client_Brief_ID").Value = ""
                .Fields("IB_ID").Value = recRadioPlanTemp.Fields("IB_ID").Value
                .Fields("Month").Value = recRadioPlanTemp.Fields("Month").Value
                .Fields("Year").Value = recRadioPlanTemp.Fields("Year").Value
                .Fields("Budget").Value = recRadioPlanTemp.Fields("Budget").Value
                .Update
                recRadioPlanTemp.MoveNext
            Loop
        End With
        
        If recRadioPlan.State = adStateOpen Then
            recRadioPlan.Close
        End If
            
        'Save To IB_Radio_Plan_Detail Tabel
        With recPlanDetail
            If recPlanDetailTemp.RecordCount > 0 Then
                recPlanDetailTemp.Filter = ""
                recPlanDetailTemp.MoveFirst
            End If
            
            Do While recPlanDetailTemp.EOF = False
                .AddNew
                .Fields("Client_Brief_ID").Value = ""
                .Fields("IB_ID").Value = recPlanDetailTemp.Fields("IB_ID").Value
                .Fields("Month").Value = recPlanDetailTemp.Fields("Month").Value
                .Fields("Year").Value = recPlanDetailTemp.Fields("Year").Value
                .Fields("Schedule").Value = recPlanDetailTemp.Fields("Schedule").Value
                .Fields("City_id").Value = recPlanDetailTemp.Fields("City_Code").Value
                .Fields("Area_Code").Value = recPlanDetailTemp.Fields("Area_Code").Value
                .Fields("spot").Value = recPlanDetailTemp.Fields("spot").Value
                .Fields("Urban_Flag").Value = recPlanDetailTemp.Fields("Urban_Flag").Value
                .Fields("Rural_Flag").Value = recPlanDetailTemp.Fields("Rural_Flag").Value
                .Update
                recPlanDetailTemp.MoveNext
            Loop
        End With
        
        If recPlanDetail.State = adStateOpen Then
            recPlanDetail.Close
        End If
        
        Set recPlanDetail = Nothing
        
        With recPlanDetailMaterial
            If recPlanDetailMaterialTemp.RecordCount > 0 Then
                recPlanDetailMaterialTemp.Filter = ""
                recPlanDetailMaterialTemp.MoveFirst
            End If
            
            Do While recPlanDetailMaterialTemp.EOF = False
                .AddNew
                .Fields("Client_Brief_ID").Value = ""
                .Fields("IB_ID").Value = recPlanDetailMaterialTemp.Fields("IB_ID").Value
                .Fields("Month").Value = recPlanDetailMaterialTemp.Fields("Month").Value
                .Fields("Year").Value = recPlanDetailMaterialTemp.Fields("Year").Value
                .Fields("Schedule").Value = recPlanDetailMaterialTemp.Fields("Schedule").Value
                .Fields("City_id").Value = recPlanDetailMaterialTemp.Fields("City_Code").Value
                .Fields("Material_Id").Value = recPlanDetailMaterialTemp.Fields("Material_Id").Value
                .Fields("material_Mix").Value = recPlanDetailMaterialTemp.Fields("material_Mix").Value
                
                .Update
                recPlanDetailMaterialTemp.MoveNext
            Loop
        End With
        
        If recPlanDetailMaterial.State = adStateOpen Then
            recPlanDetailMaterial.Close
        End If
        
        Set recPlanDetailMaterial = Nothing
        
        'Save To IB_Radio_Material Table
        With recMateri
            If recMateriTemp.RecordCount > 0 Then
                recMateriTemp.Filter = ""
                recMateriTemp.MoveFirst
            End If
            
            Do While recMateriTemp.EOF = False
                .AddNew
                .Fields("Client_Brief_ID").Value = ""
                .Fields("IB_ID").Value = recMateriTemp.Fields("IB_ID").Value
                .Fields("Material_ID").Value = recMateriTemp.Fields("Material_ID").Value
                .Fields("material_name").Value = recMateriTemp.Fields("material_name").Value
                .Fields("Duration").Value = recMateriTemp.Fields("Duration").Value
                .Update
                recMateriTemp.MoveNext
            Loop
        End With
        
        If recMateri.State = adStateOpen Then
            recMateri.Close
        End If
    End If
End Sub

Public Sub CalTotal()
    Dim dblTotalValue As Double
    
    'If cmdEdit.Visible = False Then
    If lblApprove.Caption = Not_Approved Then
        dblTotalValue = 0
        With recRadioPlanTemp
            .Filter = ""
            .Filter = "IB_ID ='" & txtIBID.Text & "'"
            If .RecordCount > 0 Then
                .MoveFirst
                Do While .EOF = False
                    dblTotalValue = dblTotalValue + .Fields("Budget").Value
                    .MoveNext
                Loop
            End If
            .Filter = ""
        End With
    End If

    medGrandTotal.Text = dblTotalValue
End Sub

Public Sub monthIB(cboName As ComboBox)
    cboName.Clear
    Select Case intIBStartingMonth
        Case 1
            cboName.AddItem "January"
            cboName.AddItem "February"
            cboName.AddItem "March"
        Case 2
            cboName.AddItem "February"
            cboName.AddItem "March"
            cboName.AddItem "April"
        Case 3
            cboName.AddItem "March"
            cboName.AddItem "April"
            cboName.AddItem "May"
        Case 4
            cboName.AddItem "April"
            cboName.AddItem "May"
            cboName.AddItem "June"
        Case 5
            cboName.AddItem "May"
            cboName.AddItem "June"
            cboName.AddItem "July"
        Case 6
            cboName.AddItem "June"
            cboName.AddItem "July"
            cboName.AddItem "August"
        Case 7
            cboName.AddItem "July"
            cboName.AddItem "August"
            cboName.AddItem "September"
        Case 8
            cboName.AddItem "August"
            cboName.AddItem "September"
            cboName.AddItem "October"
        Case 9
            cboName.AddItem "September"
            cboName.AddItem "October"
            cboName.AddItem "November"
        Case 10
            cboName.AddItem "October"
            cboName.AddItem "November"
            cboName.AddItem "December"
        Case 11
            cboName.AddItem "November"
            cboName.AddItem "December"
        Case 12
            cboName.AddItem "December"
    End Select
    
    cboName.ListIndex = 0
End Sub

Private Sub Form_Unload(IntCancel As Integer)
    If recRadioPlanTemp.State = adStateOpen Then
        recRadioPlanTemp.Close
        Set recRadioPlanTemp = Nothing
    End If

    If recPlanDetailTemp.State = adStateOpen Then
        recPlanDetailTemp.Close
        Set recPlanDetailTemp = Nothing
    End If

    If recPlanDetailMaterialTemp.State = adStateOpen Then
        recPlanDetailMaterialTemp.Close
        Set recPlanDetailMaterialTemp = Nothing
    End If

    If recMateriTemp.State = adStateOpen Then
        recMateriTemp.Close
        Set recMateriTemp = Nothing
    End If
End Sub

Private Sub fraClientApproval_DblClick()
    
    If lblApprove.Caption = "APPROVED" Then
            MsgBox "Data Can not be approved.Data has been approved!", vbCritical, strApplication_Name
        Exit Sub
    End If
    If lblApprove.Caption = "" Then
            MsgBox "Data Can not be approved Empy Data!", vbCritical, strApplication_Name
        Exit Sub
    End If
    'Check apakah Planner Brand
    If Not IsValidAccess(strLogin_User, "Planner", Left(cboBrand.Text, 4)) Then
        MsgBox strMsgAccessDenied, vbCritical, strTitleExclamation
        Exit Sub
    End If
    
    If lblApprove.Caption <> Approved Then
        If txtIBID.Text <> "" Then
            If strTransProcess = "SHOW" Then
                'Set frm_Approve_IB.frmWhatApproval = Frm_IB_Radio
                
                frm_Approve_IB.show vbModal
                
                If blnStatusPassword = True Then
                    recDate.Requery
                    
                    strQuery = "UPDATE IB_Radio set Approved_Flag =1, approved_date='" & recDate(0) & "' "
                    strQuery = strQuery & " WHERE IB_ID ='" & txtIBID.Text & "'"
                    ConnERP.Execute strQuery
                    
                    ShowData txtIBID.Text
                    
                    lblApprove.ForeColor = vbBlack
                    lblApprove.Caption = "Approved"
                    lblDateApp.Caption = Format(recDate(0), "dd/mm/yyyy hh:mm:ss AMPM")
                End If
            End If
        End If
    End If
End Sub

Private Sub InitGrids()
    dbgCity.Rows = 1
    dbgCity.cols = 7
    dbgCity.TextMatrix(0, 0) = "SCHEDULE"
    dbgCity.TextMatrix(0, 1) = "SPOT/DAY"
    dbgCity.TextMatrix(0, 2) = "SALES AREA"
    dbgCity.TextMatrix(0, 3) = "URBAN AREA"
    dbgCity.TextMatrix(0, 4) = "RURAL AREA"
    dbgCity.TextMatrix(0, 5) = "MATERIAL ID"
    dbgCity.TextMatrix(0, 6) = "MATERIAL MIX"
    
    dbgCity.ColWidth(0) = 2000
    dbgCity.ColWidth(1) = 1000
    dbgCity.ColWidth(2) = 1300
    dbgCity.ColWidth(3) = 1000
    dbgCity.ColWidth(4) = 1000
    dbgCity.ColWidth(5) = 1200
    dbgCity.ColWidth(6) = 1300

'    msgPlan.Rows = 1
'    msgPlan.cols = 1
'    msgPlan.TextMatrix(0, 0) = "BUDGET"
'    msgPlan.ColWidth(0) = 1500
    
    msgMix.cols = 3
    msgMix.Rows = 1
    msgMix.TextMatrix(0, 0) = "Material"
    msgMix.TextMatrix(0, 1) = "MIX (%)"
    msgMix.TextMatrix(0, 1) = "Duration"
    msgMix.ColWidth(0) = 1000
    msgMix.ColWidth(1) = 700
    msgMix.ColWidth(2) = 700
End Sub


Private Sub OptArea_Click()
    LoadDetail txtIBID.Text, cboYear.Text, Get_Month_Number(cboPlanMonth.Text)
End Sub

Private Sub optCity_Click()
    LoadDetail txtIBID.Text, cboYear.Text, Get_Month_Number(cboPlanMonth.Text)
End Sub

Private Sub cboBrand_KeyPress(KeyAscii As Integer)
    KeyAscii = 0
End Sub

Private Sub cboBrandVariant_KeyPress(KeyAscii As Integer)
    KeyAscii = 0
End Sub

Private Sub cboPlanMonth_KeyPress(KeyAscii As Integer)
    KeyAscii = 0
End Sub

Private Sub cboSecondary_KeyPress(KeyAscii As Integer)
    KeyAscii = 0
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
    With picButton(enButtonType.biePrint)   'FIND.
        .Enabled = paIsNormalMode
        .Visible = paIsNormalMode
        '.Left = picButton(4).Left
    End With
    With picButton(enButtonType.biefind)  'FIND.
        .Enabled = paIsNormalMode
        .Visible = paIsNormalMode
        '.Left = picButton(4).Left
    End With
    
    With picButton(enButtonType.bieApprove)   'APRROVE.
        .Enabled = paIsNormalMode
        .Visible = paIsNormalMode
        '.Left = picButton(4).Left
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
    dbgCity.Width = fraPlanMonth.Width - dbgCity.Left - cmdMateri.Width - 450
    fraClientApproval.Width = pnlMain.Width - fraClientApproval.Left - 150
    picApproved.Left = ((fraClientApproval.Width / 2)) - ((picApproved.Width) / 2) 'fraClientApproval.Width - (picApproved.Left * 2)
    txtMediaPlanNo.Width = fraIB.Width - txtMediaPlanNo.Left - 150
    txtPrimaryTarget.Width = txtMediaPlanNo.Width
'    pnl_Main.Height = Me.ScaleHeight - picToolbar.Height - picStatusBar.Height
'    fra_Deliverable.Height = pnl_Main.Height - (fra_Deliverable.Top + 100)
'    SSTab3.Height = fra_Deliverable.Height - (SSTab3.Top) - 150
'    txtOther_Recomedation.Height = SSTab3.Height - (txtOther_Recomedation.Top) - 150
'    txtAggreed_Channel_shortlist.Height = txtOther_Recomedation.Height
'    fra_DeliverableChannel.Height = pnl_Main.Height - (fra_DeliverableChannel.Top + 100)
'    fraFilter.Width = pnl_Main.Width - (fraFilter.Left * 2)
'    lineFilter.X1 = fraFilter.Width / 2
'    lineFilter.X2 = lineFilter.X1
'    Fra_Approve.Left = lineFilter.X2 + Label7.Left
'    txtYear.Width = lineFilter.X2 - txtYear.Left - 50
'    txtClient_Brief_Id.Width = txtYear.Width
'    txtExtention.Width = txtYear.Width
'    txtStatus.Width = txtYear.Width
'    'left part
'    lbl_dateofPreviousIssue.Left = lineFilter.X1 + Label7.Left
'    dtpDate_Previouse.Left = lbl_dateofPreviousIssue.Left + lbl_dateofPreviousIssue.Width + 50
'    dtpDate_Issue.Left = dtpDate_Previouse.Left
'    lbl_DateIssue.Left = lbl_dateofPreviousIssue.Left
'    lblCountry.Left = lbl_dateofPreviousIssue.Left
'    cboCountry.Left = dtpDate_Previouse.Left
'    Fra_Approve.Left = dtpDate_Previouse.Left
'    fra_DeliverableChannel.Width = pnl_Main.Width - fra_DeliverableChannel.Left - fraFilter.Left
'    lstRec_Channel_Selection.Width = fra_DeliverableChannel.Width - (lstRec_Channel_Selection.Left * 2)
'    lstRec_Channel_Selection.Height = fra_DeliverableChannel.Height - (lstRec_Channel_Selection.Top) - 200
'    chk_All.Top = lstRec_Channel_Selection.Height + lstRec_Channel_Selection.Top + 50
'    lbl_CheckAll.Top = chk_All.Top
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

            Call db_add
        Case enButtonType.bieEdit  '5 'EDIT.
            Call db_edit
        Case enButtonType.bieDelete  '6 'DELETE.
            Call db_delete
        Case enButtonType.biefind   '7 'FIND.
            Call db_Find
        Case enButtonType.biefind   'Find.
            Call db_Find
        Case enButtonType.bieSave  'SAVE.
            Call db_save
        Case enButtonType.biecancel 'CANCEL.
            Call db_Cancel
        Case enButtonType.bieClose  'CANCEL.
            Call db_close
        Case enButtonType.biePrint   'Print.
            Call db_print
        Case enButtonType.bieApprove
            Call fraClientApproval_DblClick

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

