VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Object = "{562E3E04-2C31-4ECE-83F4-4017EEE51D40}#8.0#0"; "todg8.ocx"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSAdoDc.ocx"
Begin VB.Form Frm_IB_Radio_Seach 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Select IB Id"
   ClientHeight    =   6195
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3630
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6195
   ScaleWidth      =   3630
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
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
      ScaleWidth      =   3630
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   0
      Width           =   3630
      Begin VB.PictureBox picButton 
         Appearance      =   0  'Flat
         BackColor       =   &H00F0F0F0&
         BorderStyle     =   0  'None
         ClipControls    =   0   'False
         ForeColor       =   &H80000008&
         Height          =   750
         Index           =   11
         Left            =   1620
         Picture         =   "Frm_IB_Radio_Seach.frx":0000
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
         Index           =   17
         Left            =   90
         Picture         =   "Frm_IB_Radio_Seach.frx":1D5A
         ScaleHeight     =   750
         ScaleWidth      =   1500
         TabIndex        =   1
         TabStop         =   0   'False
         Top             =   0
         Width           =   1500
      End
   End
   Begin Threed.SSPanel pnlMain 
      Align           =   1  'Align Top
      Height          =   5340
      Left            =   0
      TabIndex        =   3
      Top             =   750
      Width           =   3630
      _Version        =   65536
      _ExtentX        =   6403
      _ExtentY        =   9419
      _StockProps     =   15
      BackColor       =   15790320
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BevelOuter      =   1
      Begin TrueOleDBGrid80.TDBGrid tdgIBRadio 
         Bindings        =   "Frm_IB_Radio_Seach.frx":396F
         Height          =   5115
         Left            =   120
         TabIndex        =   4
         Top             =   120
         Width           =   3420
         _ExtentX        =   6033
         _ExtentY        =   9022
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
         Splits(0).FilterBar=   -1  'True
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
      Begin MSAdodcLib.Adodc adoData 
         Height          =   495
         Left            =   0
         Top             =   0
         Visible         =   0   'False
         Width           =   1200
         _ExtentX        =   2117
         _ExtentY        =   873
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
         Caption         =   "Adodc1"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         _Version        =   393216
      End
   End
End
Attribute VB_Name = "Frm_IB_Radio_Seach"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim recIbId As New ADODB.Recordset
Dim col As TrueOleDBGrid80.Column
Dim cols As TrueOleDBGrid80.Columns
Dim strSql As String

'
Private Sub Form_Load()
    Dim strSql As String
    Dim intCountUser
    Dim tdbNonActiveStyle As New TrueOleDBGrid80.Style
             
    strSql = ConnERP.ConnectionString & ";password=erp"

    adoData.ConnectionString = strSql
    
    strSql = "SELECT IB_ID FROM IB_Radio WHERE Left(IB_ID,4)='" & Left(Trim(Frm_IB_Radio.cboBrand.Text), 4) & "' AND Year='" & Frm_IB_Radio.cboYear.Text & "'"
    tdgIBRadio.ClearFields
    adoData.RecordSource = strSql
    adoData.Refresh
    tdgIBRadio.Refresh
    tdgIBRadio.Columns(0).Caption = "Implementation Brif ID"
    tdgIBRadio.Columns(0).Width = 3000
    
End Sub

Private Sub db_OK()
    Unload Me
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
        Case enButtonType.bieOK '4 'ADD.
            Call db_OK
        Case Else
            Unload Me
    End Select

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

Private Sub picButton_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single) 'TOOLBAR_AI.
'************************************************
' Procedure         : picButton_MouseDown
' Function          : TOOLBAR_AI saat mouse ditekan.
' Created By        : {73 64 6B}
' Date              : 12-Apr-2015
'************************************************
    
    picButton(Index).Picture = LoadPicture(SetButtonImageEffect(Index, bieDown)) 'FIRST.

End Sub

Private Function getFilter() As String
 '*****************************************
'Procedure Name     : tdgIBRadio_Filter
'Procedure Function :   Creates the SQL statement in adodata.recordset.filter
'                       and only filters text currently. It must be modified to filter other data types.
'Input Parameter    :
'Output Parameter   :
'LastUpdate Date    : 12-02-2016
'LastUpdate By      : Abdi / Creative
'*****************************************

    Dim intCol As Integer
    Dim strTmp As String
    
    strTmp = ""
    
    For Each col In cols
        If Trim(col.FilterText) <> "" Then
            intCol = intCol + 1
            
                strTmp = strTmp & "((" & col.DataField & " LIKE '%" & col.FilterText & "') "
                strTmp = strTmp & "OR (" & col.DataField & " LIKE '" & col.FilterText & "%') "
                strTmp = strTmp & "OR (" & col.DataField & " LIKE '%" & col.FilterText & "%'))"
            
        End If
        
    Next col

    getFilter = strTmp
    
End Function

Private Sub cmdClearFilter_Click()
 '*****************************************
'Procedure Name     : cmdClearFilter
'Procedure Function : Clears filter from grid
'Input Parameter    : -
'Output Parameter   : -
'*****************************************
    '
    For Each col In tdgIBRadio.Columns
        col.FilterText = ""
    Next col
    adoData.Recordset.Filter = adFilterNone

End Sub


Private Sub tdgIBRadio_FilterChange()
'*****************************************
'Procedure Name     : tdgIBRadio_FilterChange
'Procedure Function : Gets called when an action is performed on the filter bar
'Input Parameter    :
'Output Parameter   :
'lastUpdate Date    : 12-02-2016
'LastUpdate By      : Abdi / Creative
'*****************************************

On Error GoTo errHandler
    
    Set cols = tdgIBRadio.Columns
    Dim intCol As Integer
    Dim strDummy As String

    intCol = tdgIBRadio.col
    
    tdgIBRadio.HoldFields
    
    strSql = "SELECT IB_ID FROM IB_Radio WHERE Left(IB_ID,4)='" & Left(Trim(Frm_IB_Radio.cboBrand.Text), 4) & "' AND Year='" & Frm_IB_Radio.cboYear.Text & "'"
  
    If IsNumeric(tdgIBRadio.Columns(intCol).Value) = True Then
        strDummy = getFilter()
        If strDummy <> "" Then
            adoData.RecordSource = strSql & " AND " & strDummy
        End If
        adoData.Refresh
        tdgIBRadio.Refresh
    End If
    
    tdgIBRadio.col = intCol
    tdgIBRadio.EditActive = True
    tdgIBRadio.Columns(0).Caption = "Implementation Brif ID"
    tdgIBRadio.Columns(0).Width = 3000
Exit Sub

errHandler:
    Call cmdClearFilter_Click
End Sub


Private Sub tdgIBRadio_RowColChange(LastRow As Variant, ByVal LastCol As Integer)

    Frm_IB_Radio.txtIBID.Text = tdgIBRadio.Columns(0).Value
    
    If Frm_IB_Radio.txtIBID.Text <> "" Then
        Frm_IB_Radio.ShowData Frm_IB_Radio.txtIBID

        Frm_IB_Radio.PrepareTemp Frm_IB_Radio.txtIBID
    Else
        Frm_IB_Radio.ClearForm
    End If

End Sub
