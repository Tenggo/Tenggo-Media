VERSION 5.00
Object = "{562E3E04-2C31-4ECE-83F4-4017EEE51D40}#8.0#0"; "todg8.ocx"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSAdoDc.ocx"
Begin VB.Form frm_Find_Cient_BriefM 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Find Client Brief Media"
   ClientHeight    =   4905
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3375
   ClipControls    =   0   'False
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
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4905
   ScaleWidth      =   3375
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmd 
      Caption         =   "&Cancel"
      Height          =   285
      Index           =   1
      Left            =   1620
      TabIndex        =   1
      Top             =   4515
      Width           =   1245
   End
   Begin MSAdodcLib.Adodc ado_Data 
      Height          =   495
      Left            =   315
      Top             =   3690
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
   Begin VB.CommandButton cmd 
      Caption         =   "&OK"
      Height          =   285
      Index           =   0
      Left            =   330
      TabIndex        =   0
      Top             =   4515
      Width           =   1245
   End
   Begin TrueOleDBGrid80.TDBGrid tdg_Comp_Name 
      Bindings        =   "frm_Find_Cient_BriefM.frx":0000
      Height          =   4395
      Left            =   15
      TabIndex        =   2
      Top             =   45
      Width           =   3300
      _ExtentX        =   5821
      _ExtentY        =   7752
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
End
Attribute VB_Name = "frm_Find_Cient_BriefM"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'***********************************************************
'Submodul Name      : frm_Find_Cient_BriefM.
'Submodul Function  : Utk mencari data id Cient Brief Media.
'Used Table         : Client_Brief_Media, Brand
'Procedure/Function : cmd_Click, Form_Load
'Programmer Name    : Tedi
'Date               : 25-Apr-2016
'Update             : Tedi
'Update By          : 25-Apr-2016
'***************************************************************

Option Explicit

Dim strSql As String

Public strYear As String    'Parameter input Tahun yang dikirim dari Form frm_client_brieft_media
Public strBrand As String   'Parameter input Brand Code yang dikirim dari Form frm_client_brieft_media
Public strClient_Brief_Id As String 'Parameter input Client_Brief_Id yang dikirim dari Form frm_client_brieft_media, untuk pointer default row

Dim col As TrueOleDBGrid80.Column
Dim cols As TrueOleDBGrid80.Columns

Private Sub cmd_Click(Index As Integer)
'*****************************************
'Procedure Name     : cmd_Click
'Procedure Function : Pilihan OK atau Cancel
'Input Parameter    : ---
'Output Parameter   : ---
'Last Update Date   : 21-02-2016
'Last Update By     : Abdi / Kreatif
'*****************************************
    
    Select Case Index
        Case 0 'OK
            If ado_Data.Recordset.RecordCount < 1 Then Exit Sub
                strClient_Brief_Id = ado_Data.Recordset.Fields("Client_Brief_Id").Value
          
        Case 1 'Cancel
            strClient_Brief_Id = ""
    
    End Select
    
    Unload Me

End Sub

Private Sub Form_Load()
'*****************************************
'Procedure Name     : Form_Load
'Procedure Function : Load data berdasarkan Client_Brief_Id yang di sorot.
'Input Parameter    : ---
'Output Parameter   : ---
'Last Update Date   : 21-02-2016
'Last Update By     : Tedi / Kreatif
'*****************************************
    
    Dim strSql As String
    Dim intCountUser
    Dim tdbNonActiveStyle As New TrueOleDBGrid80.Style
             
    strSql = ConnERP.ConnectionString & ";password=erp"

    ado_Data.ConnectionString = strSql
    
    strSql = "SELECT Year,Client_Brief_Id "
    strSql = strSql & " FROM Client_Brief_Media "
    'strSql = strSql & " WHERE Year ='" & strYear & "' "
    strSql = strSql & " WHERE brand_code ='" & strBrand & "'"
        
    ado_Data.RecordSource = strSql
    ado_Data.Refresh
    tdg_Comp_Name.Refresh
    
    If strClient_Brief_Id <> "" Then
        ado_Data.Recordset.MoveFirst
        ado_Data.Recordset.Find "Client_Brief_Id='" & strClient_Brief_Id & "'"
    End If
    
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
'*****************************************
'Procedure Name     : Form_QueryUnload
'Procedure Function : ~ Lebih fleksibel pake Form_QueryUnload drpd Form_Unload krn ada param Cancel dan UnloadMode yg bisa digunakan utk kondisi tertentu.
'                     ~ Tutup semua recordset
'                     ~ RelaseCapture API utk MouseEvents
'                     ~ Bersihkan semua var string agar memory kembali.
'                     ~ Set form = nothing agar memory kembali.
'Input Parameter    : ~ Cancel: True=Membatalkan unload form, False=Melanjutkan unload form.
'                     ~ UnloadMode:
'                       ~ vbFormControlMenu=Form is being closed by user
'                       ~ vbFormCode=Form is being closed by code
'                       ~ vbAppWindows=The current Windows session is ending
'                       ~ vbAppTaskManager=Task Manager is closing this application
'                       ~ vbFormMDIForm=MDI parent is closing this form
'                       ~ vbFormOwner=The owner form is closing
'Output Parameter   : ---
'Date               : 25-Apr-2015/{73 64 6B}
'*****************************************

    Call CloseRecordset(ado_Data.Recordset)

End Sub

Private Sub tdg_Comp_Name_DblClick()
'*****************************************
'Procedure Name     : tdg_Comp_Name_FilterChange
'Procedure Function : Gets called when an action is performed on the filter bar
'Input Parameter    :
'Output Parameter   :
'lastUpdate Date    : 12-02-2016
'LastUpdate By      : Abdi / Creative
'*****************************************
    
    cmd_Click 0

End Sub

Private Sub tdg_Comp_Name_FilterChange()
'*****************************************
'Procedure Name     : tdg_Comp_Name_FilterChange
'Procedure Function : Gets called when an action is performed on the filter bar
'Input Parameter    :
'Output Parameter   :
'lastUpdate Date    : 12-02-2016
'LastUpdate By      : Abdi / Creative
'*****************************************

On Error GoTo errHandler
    
    Set cols = tdg_Comp_Name.Columns
    Dim intCol As Integer
    Dim strDummy As String

    intCol = tdg_Comp_Name.col
    
    tdg_Comp_Name.HoldFields
    
    strSql = "SELECT Year,Client_Brief_Id "
    strSql = strSql & " FROM Client_Brief_Media "
'    strSql = strSql & " AND Year ='" & strYear & "' "
'    strSql = strSql & " AND brand_code ='" & strBrand & "'"
    
    If IsNumeric(tdg_Comp_Name.Columns(intCol).Value) = True Then
        strDummy = getFilter()
        If strDummy <> "" Then
            ado_Data.RecordSource = strSql & " AND " & strDummy
        Else
            ado_Data.RecordSource = strSql & strDummy
        End If
        ado_Data.Refresh
        tdg_Comp_Name.Refresh
    Else
        If tdg_Comp_Name.Columns(intCol).Value <> "" Then
            strSql = strSql & " AND " & getFilter()
            ado_Data.RecordSource = strSql
        Else
            ado_Data.RecordSource = strSql

        End If
            ado_Data.Refresh
            tdg_Comp_Name.Refresh
    End If
    
    tdg_Comp_Name.col = intCol
    tdg_Comp_Name.EditActive = True

Exit Sub

errHandler:
    Call cmdClearFilter_Click

End Sub

Private Function getFilter() As String
 '*****************************************
'Procedure Name     : tdg_Comp_Name_Filter
'Procedure Function :   Creates the SQL statement in ado_Data.recordset.filter
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
            
            If IsNumeric(tdg_Comp_Name.Columns(intCol).Value) = True Then
                strTmp = strTmp & "((" & col.DataField & " LIKE '%" & Val(col.FilterText) & "') "
                strTmp = strTmp & "OR (" & col.DataField & " LIKE '" & Val(col.FilterText) & "%') "
                strTmp = strTmp & "OR (" & col.DataField & " LIKE '%" & Val(col.FilterText) & "%'))"
            Else
                strTmp = strTmp & "((" & col.DataField & " LIKE '%" & col.FilterText & "') "
                strTmp = strTmp & "OR (" & col.DataField & " LIKE '" & col.FilterText & "%') "
                strTmp = strTmp & "OR (" & col.DataField & " LIKE '%" & col.FilterText & "%'))"
            End If
            
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
    For Each col In tdg_Comp_Name.Columns
        col.FilterText = ""
    Next col
    ado_Data.Recordset.Filter = adFilterNone

End Sub

Private Sub tdg_Comp_Name_KeyPress(KeyAscii As Integer)
'************************************************
' Procedure         : tdg_Comp_Name_KeyPress
' Function          : Exit Browser with Enter
' Parameter Input   : -
' Parameter Output  : -
' Last Update       : Tedi
' Last Update By    : 22 Feb 2016
'************************************************

    If KeyAscii = vbKeyReturn Then
            Unload Me
    End If

End Sub

Private Sub tdg_Comp_Name_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
'************************************************
' Procedure         : tdg_Comp_Name_RowColChange
' Function          : Remove movement recordset Frm_Client_Brief_Media
' Parameter Input   : -
' Parameter Output  : -
' Last Update       : Tedi
' Last Update By    : 22 Feb 2016
'************************************************

    Frm_Client_Brief_Media.RemoteMovement Trim(ado_Data.Recordset!Client_Brief_Id)

End Sub
