VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.ocx"
Object = "{562E3E04-2C31-4ECE-83F4-4017EEE51D40}#8.0#0"; "todg8.ocx"
Begin VB.Form Frm_IB_Radio_Materi 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Radio Material"
   ClientHeight    =   5415
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6285
   ControlBox      =   0   'False
   Icon            =   "Frm_IB_Radio_Materi.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5415
   ScaleWidth      =   6285
   StartUpPosition =   2  'CenterScreen
   Begin Threed.SSPanel SSPanel1 
      Align           =   1  'Align Top
      Height          =   4545
      Left            =   0
      TabIndex        =   7
      Top             =   750
      Width           =   6285
      _Version        =   65536
      _ExtentX        =   11086
      _ExtentY        =   8017
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
      Begin TrueOleDBGrid80.TDBGrid tdgMateri 
         Bindings        =   "Frm_IB_Radio_Materi.frx":0442
         Height          =   3315
         Left            =   120
         TabIndex        =   19
         Top             =   1080
         Width           =   6060
         _ExtentX        =   10689
         _ExtentY        =   5847
         _LayoutType     =   4
         _RowHeight      =   -2147483647
         _WasPersistedAsPixels=   0
         Columns(0)._VlistStyle=   0
         Columns(0)._MaxComboItems=   5
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
      Begin VB.PictureBox Picture1 
         Height          =   645
         Left            =   240
         ScaleHeight     =   585
         ScaleWidth      =   5625
         TabIndex        =   12
         Top             =   5235
         Visible         =   0   'False
         Width           =   5685
         Begin VB.CommandButton cmdEdit 
            Caption         =   "&Edit"
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
            Left            =   1545
            TabIndex        =   18
            Top             =   -90
            Width           =   1180
         End
         Begin VB.CommandButton cmdAdd 
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
            Height          =   495
            Left            =   315
            TabIndex        =   17
            Top             =   15
            Width           =   1180
         End
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
            Height          =   495
            Left            =   3585
            TabIndex        =   16
            Top             =   0
            Width           =   1180
         End
         Begin VB.CommandButton cmdDelete 
            Caption         =   "&Delete"
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
            Left            =   2865
            TabIndex        =   15
            Top             =   0
            Width           =   1180
         End
         Begin VB.CommandButton cmdSave 
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
            TabIndex        =   14
            Top             =   0
            Width           =   1180
         End
         Begin VB.CommandButton cmdCancel 
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
            Left            =   1200
            TabIndex        =   13
            Top             =   0
            Width           =   1180
         End
      End
      Begin VB.TextBox txtDuration 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
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
         Height          =   325
         Left            =   1470
         MaxLength       =   5
         TabIndex        =   9
         Top             =   600
         Width           =   615
      End
      Begin VB.TextBox txtMateriName 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
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
         Height          =   325
         Left            =   1470
         MaxLength       =   50
         TabIndex        =   8
         Top             =   195
         Width           =   4605
      End
      Begin VB.Label lblDurasi 
         BackStyle       =   0  'Transparent
         Caption         =   "Du&ration "
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
         Left            =   135
         TabIndex        =   11
         Top             =   630
         Width           =   1275
      End
      Begin VB.Label lblMateriName 
         BackStyle       =   0  'Transparent
         Caption         =   "Material &Name "
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
         Left            =   135
         TabIndex        =   10
         Top             =   240
         Width           =   1275
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
      ScaleWidth      =   6285
      TabIndex        =   0
      Top             =   0
      Width           =   6285
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
         Left            =   4680
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
         Index           =   5
         Left            =   1620
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
         Index           =   6
         Left            =   3150
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
         Index           =   10
         Left            =   105
         ScaleHeight     =   750
         ScaleWidth      =   1500
         TabIndex        =   3
         Top             =   -30
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
         Index           =   11
         Left            =   1620
         ScaleHeight     =   750
         ScaleWidth      =   1500
         TabIndex        =   1
         Top             =   0
         Width           =   1500
      End
   End
End
Attribute VB_Name = "Frm_IB_Radio_Materi"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
''************************************************
' Function          : entry material di IB Radio
' Last Update       :.
'************************************************
'
Option Explicit

Dim strMatID(25) As String
Dim recMateri As New ADODB.Recordset
Dim blnEditFlag As Boolean

Private Sub Form_Load()
    
    blnEditFlag = True
    Call InitialData
    Call LoadData
    Call SetButtonToolbar(False, picButton)
    EnableObject False
End Sub

Private Sub db_add()
'cmdAdd_Click
    Call ClearForm
    
    blnEditFlag = False
    
    Call SetButton(True)
    
    txtMateriName.SetFocus
End Sub

Private Sub db_Cancel()
'cmdCancel_Click
    blnEditFlag = True
    
    Call SetButton(False)
    
    Call ClearForm
End Sub



Private Sub db_delete()
'cmdDelete_Click
    
    If Frm_IB_Radio.recPlanDetailMaterialTemp.RecordCount > 0 Then
        Frm_IB_Radio.recPlanDetailMaterialTemp.Filter = ""
        Frm_IB_Radio.recPlanDetailMaterialTemp.Filter = "material_id='" & Frm_IB_Radio.recMateriTemp.Fields("material_id").Value & "'"
    End If
    
    If Not Frm_IB_Radio.recPlanDetailMaterialTemp.EOF And Not Frm_IB_Radio.recPlanDetailMaterialTemp.BOF Then
        MsgBox "You Can't Delete this Material. Check Material Mix", vbCritical, strTitleExclamation
        Exit Sub
    End If
    
    With Frm_IB_Radio
        If .recMateriTemp.RecordCount > 0 Then
            If MsgBox(strMsgDeleteConfirm, vbQuestion + vbYesNo, strTitleConfirm) = vbYes Then
                .recMateriTemp.Delete
            End If
        Else
            MsgBox "There is no data can be deleted.", vbCritical, strTitleMissingInfo
            Exit Sub
        End If
    End With
End Sub

Private Sub db_edit()
'cmdEdit_Click
    If Not Frm_IB_Radio.recMateriTemp.EOF And Not Frm_IB_Radio.recMateriTemp.BOF Then
        blnEditFlag = True
        
        Call SetButton(True)
        
        txtMateriName.Text = Frm_IB_Radio.recMateriTemp.Fields("Material_Name").Value
        txtDuration.Text = Frm_IB_Radio.recMateriTemp.Fields("Duration").Value
        txtMateriName.SetFocus
    Else
        MsgBox "There is no data can be edit", vbCritical, strTitleMissingInfo
    End If
End Sub

Private Sub db_save()
'cmdSave_Click
    Dim strNewMatID As String
    
    If txtMateriName.Text = "" Or txtDuration.Text = "" Then
        MsgBox "Please fill in the data", vbInformation, strTitleMissingInfo
        
        If txtMateriName.Text = "" Then
            txtMateriName.SetFocus
        ElseIf txtDuration.Text = "" Then
            txtDuration.SetFocus
        End If
    Else
        strNewMatID = strGetNewMaterialID
        
        If blnEditFlag = False Then
            Frm_IB_Radio.recMateriTemp.AddNew
            Frm_IB_Radio.recMateriTemp.Fields("Material_ID").Value = strNewMatID
        End If
        
        With Frm_IB_Radio.recMateriTemp
            .Fields("IB_ID").Value = Trim(Frm_IB_Radio.txtIBID.Text)
            .Fields("Material_Name").Value = txtMateriName.Text
            .Fields("Duration").Value = Val(txtDuration.Text)
            .Update
        End With
    
        blnEditFlag = True
        
        Call SetButton(False)
        
        Call ClearForm
        
        Call LoadData
    End If
End Sub

Private Function strGetNewMaterialID() As String
    Dim intPos As Integer
    Dim recMateri As New ADODB.Recordset
    Dim blnHaveMaterialID As Boolean
    
    Set recMateri = Frm_IB_Radio.recMateriTemp.Clone
    
    With recMateri
        intPos = 0
        blnHaveMaterialID = False
        Do While .EOF = False
            If .Fields("Material_ID").Value <> strMatID(intPos) Then
                strGetNewMaterialID = strMatID(intPos)
                
                If recMateri.State = adStateOpen Then
                    recMateri.Close
                End If
                
                Set recMateri = Nothing
                Exit Function
            End If
            
            intPos = intPos + 1
            .MoveNext
        Loop
    End With
    
    strGetNewMaterialID = strMatID(intPos)
    
    If recMateri.State = adStateOpen Then
        recMateri.Close
    End If
    
    Set recMateri = Nothing
End Function

Private Sub InitialData()
    strMatID(0) = "A"
    strMatID(1) = "B"
    strMatID(2) = "C"
    strMatID(3) = "D"
    strMatID(4) = "E"
    strMatID(5) = "F"
    strMatID(6) = "G"
    strMatID(7) = "H"
    strMatID(8) = "I"
    strMatID(9) = "J"
    strMatID(10) = "K"
    strMatID(11) = "L"
    strMatID(12) = "M"
    strMatID(13) = "N"
    strMatID(14) = "O"
    strMatID(15) = "P"
    strMatID(16) = "Q"
    strMatID(17) = "R"
    strMatID(18) = "S"
    strMatID(19) = "T"
    strMatID(20) = "U"
    strMatID(21) = "V"
    strMatID(22) = "W"
    strMatID(23) = "X"
    strMatID(24) = "Y"
    strMatID(25) = "Z"
End Sub

Private Sub SetButton(blnEnable As Boolean)
    cmdAdd.Visible = Not blnEnable
    cmdEdit.Visible = Not blnEnable
    cmdDelete.Enabled = Not blnEnable
    cmdClose.Enabled = Not blnEnable
    
    txtMateriName.Enabled = blnEnable
    txtDuration.Enabled = blnEnable
End Sub

Private Sub ClearForm()
    txtDuration.Text = Empty
    txtMateriName.Text = Empty
End Sub

Private Sub LoadData()

    Set tdgMateri.DataSource = Nothing
    tdgMateri.ClearFields
    tdgMateri.DataSource = Frm_IB_Radio.recMateriTemp
    tdgMateri.Columns(0).Visible = False
    tdgMateri.Columns(1).Visible = False
    tdgMateri.Columns(2).Caption = "Material ID"
    tdgMateri.Columns(3).Caption = "Material Name"
    tdgMateri.Columns(4).Alignment = dbgRight
    tdgMateri.Columns(3).Width = 4050
    
End Sub

Private Sub txtDuration_KeyPress(KeyAscii As Integer)
    If (Chr(KeyAscii) < "0" Or Chr(KeyAscii) > "9") And KeyAscii <> 8 And KeyAscii <> 13 And KeyAscii <> 9 And Chr(KeyAscii) <> "." Then
        KeyAscii = 0
        Beep
        Exit Sub
    End If
End Sub

Private Sub txtMateriName_GotFocus()
    txtMateriName.SelStart = 0
    txtMateriName.SelLength = Len(txtMateriName.Text)
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
    Line2.X1 = 0
    Line2.X2 = mdi_Main.ScaleWidth
    fraIB.Width = Me.Width - (fraIB.Left * 2)
    fraPlanMonth.Width = Me.Width - (fraIB.Left * 2)

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
            Call SetButtonToolbar(False, picButton)
        Case enButtonType.bieEdit  '5 'EDIT.
            Call SetButtonToolbar(False, picButton)
            Call db_edit
        Case enButtonType.bieDelete  '6 'DELETE.
            Call db_delete
        Case enButtonType.bieClose  '9 'EXIT.
            
            Unload Me
        Case enButtonType.bieSave  'SAVE.
            Call db_save
            Call SetButtonToolbar(True, picButton)
        Case enButtonType.biecancel 'CANCEL.
            Call db_Cancel
            Call SetButtonToolbar(True, picButton)
    End Select

End Sub
