VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "mscomctl.OCX"
Begin VB.Form frm_MPSearch 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Search"
   ClientHeight    =   4095
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7080
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4095
   ScaleWidth      =   7080
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmd_set_is_latest 
      Caption         =   "Set As &Latest Version"
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
      Height          =   825
      Left            =   4965
      TabIndex        =   11
      Top             =   180
      Width           =   1995
   End
   Begin MSComctlLib.ListView lvSearchResult 
      Height          =   2610
      Left            =   105
      TabIndex        =   7
      Top             =   1140
      Width           =   6855
      _ExtentX        =   12091
      _ExtentY        =   4604
      View            =   3
      LabelEdit       =   1
      Sorted          =   -1  'True
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FlatScrollBar   =   -1  'True
      FullRowSelect   =   -1  'True
      _Version        =   393217
      ForeColor       =   16744576
      BackColor       =   -2147483643
      BorderStyle     =   1
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
      NumItems        =   5
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "MP Number"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Plan Year"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Brand Code"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "Brand Name"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "Client Name"
         Object.Width           =   2540
      EndProperty
   End
   Begin VB.CommandButton cmdClear 
      Caption         =   "&Clear"
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
      Height          =   360
      Left            =   3615
      TabIndex        =   6
      ToolTipText     =   "close window"
      Top             =   630
      Width           =   1230
   End
   Begin VB.CommandButton cmdSearch 
      Caption         =   "&Search"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   3615
      TabIndex        =   5
      ToolTipText     =   "close window"
      Top             =   210
      Width           =   1230
   End
   Begin VB.Frame Frame1 
      Height          =   990
      Left            =   105
      TabIndex        =   0
      Top             =   45
      Width           =   3390
      Begin VB.ComboBox cboYear 
         Height          =   315
         Left            =   1275
         TabIndex        =   2
         Top             =   570
         Width           =   2025
      End
      Begin VB.ComboBox cboBrandCode 
         Height          =   315
         Left            =   1275
         TabIndex        =   1
         Top             =   195
         Width           =   2025
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         Caption         =   "Year :"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   255
         Left            =   195
         TabIndex        =   4
         Top             =   570
         Width           =   1020
      End
      Begin VB.Label Label1 
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
         ForeColor       =   &H000000FF&
         Height          =   255
         Left            =   90
         TabIndex        =   3
         Top             =   210
         Width           =   1125
      End
   End
   Begin VB.Label Label4 
      Caption         =   "Current Version"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF8080&
      Height          =   210
      Left            =   4515
      TabIndex        =   10
      Top             =   3825
      Width           =   1185
   End
   Begin VB.Shape Shape2 
      BackColor       =   &H00FFC0C0&
      BackStyle       =   1  'Opaque
      FillColor       =   &H00FFC0C0&
      Height          =   180
      Left            =   4275
      Shape           =   4  'Rounded Rectangle
      Top             =   3840
      Width           =   165
   End
   Begin VB.Label Label3 
      Caption         =   "Old Version"
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
      Height          =   210
      Left            =   3345
      TabIndex        =   9
      Top             =   3825
      Width           =   915
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H000000FF&
      BackStyle       =   1  'Opaque
      FillColor       =   &H000000FF&
      Height          =   180
      Left            =   3120
      Shape           =   4  'Rounded Rectangle
      Top             =   3840
      Width           =   165
   End
   Begin VB.Label lblSearchResult 
      Caption         =   "100 item(s) found!"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF8080&
      Height          =   180
      Left            =   135
      TabIndex        =   8
      Top             =   3810
      Width           =   2865
   End
End
Attribute VB_Name = "frm_MPSearch"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'<CSCC>
'********************************************************************************
'Submodul Name      : frm_MPSearch
'Submodul Function  : {MemberName}
'Used Table         : -
'Used SP            : -
'Procedure/Function : -
'Programmer Name    : Yayan
'Date               : -
'Last Update By     : Tedi/Kratif
'Date Update        : 3/9/2016-11:00:57 PM
'Log Update/By      : -
'********************************************************************************
'</CSCC>

Dim recTemp As New ADODB.Recordset
Dim blnBU2 As Boolean
Dim strSql As String

Private Sub cboBrandCode_KeyPress(KeyAscii As Integer)
'<CSCM>
'********************************************************************************
'Procedure Name     : cboBrandCode_KeyPress
'Procedure Function : ---
'Input Parameter    : ---
'Output Parameter   : ---
'Date               : 3/9/2016
'LastUpdate/By      : Tedi / Kreatif
'********************************************************************************
'</CSCM>

    If KeyAscii = 13 Then
        Call cmdSearch_Click
    End If

End Sub

Private Sub cboYear_KeyPress(KeyAscii As Integer)
'<CSCM>
'********************************************************************************
'Procedure Name     : cboYear_KeyPress
'Procedure Function : ---
'Input Parameter    : ---
'Output Parameter   : ---
'Date               : 3/9/2016
'LastUpdate/By      : Tedi / Kreatif
'********************************************************************************
'</CSCM>
    
    If KeyAscii = 13 Then
        Call cmdSearch_Click
    End If

End Sub

Private Sub cmd_set_is_latest_Click()
'<CSCM>
'********************************************************************************
'Procedure Name     : cmd_set_is_latest_Click
'Procedure Function : ---
'Input Parameter    : ---
'Output Parameter   : ---
'Date               : 3/9/2016
'LastUpdate/By      : Tedi / Kreatif
'********************************************************************************
'</CSCM>
    
    Dim strMPNumber As String
    
    Me.MousePointer = vbHourglass
    strMPNumber = lvSearchResult.SelectedItem.Text
    
    'SET CURRENT LATEST VERSION TO HISTORY
    strSql = ""
    strSql = strSql & "UPDATE MP_MASTER "
    strSql = strSql & "SET IS_LATEST = 0 "
    strSql = strSql & "WHERE IS_LATEST=1 "
    strSql = strSql & "AND BRAND_CODE = '" & Left(strMPNumber, 4) & "' "
    strSql = strSql & "AND [YEAR] = " & Mid(strMPNumber, 6, 4)
    ConnERP.Execute strSql
    
    'SET SELECTED MP_NUMBER AS LATEST VERSION
    strSql = ""
    strSql = strSql & "UPDATE MP_MASTER "
    strSql = strSql & "SET IS_LATEST = 1 "
    strSql = strSql & "WHERE MP_NUMBER = '" & strMPNumber & "'"
    ConnERP.Execute strSql
    
    'REFRESH SEARCH RESULT
    Call cmdSearch_Click
    
    MsgBox "Latest version has been changed to " & strMPNumber, vbExclamation, strApplication_Name
    Me.MousePointer = vbDefault

End Sub

Private Sub cmdClear_Click()
'<CSCM>
'********************************************************************************
'Procedure Name     : cmdClear_Click
'Procedure Function : ---
'Input Parameter    : ---
'Output Parameter   : ---
'Date               : 3/9/2016
'LastUpdate/By      : Tedi / Kreatif
'********************************************************************************
'</CSCM>

    lvSearchResult.ListItems.Clear
    lblSearchResult.Caption = ""
    cboBrandCode.Text = ""
    cboYear.Text = ""
    cmdClear.Enabled = False
    cmd_set_is_latest.Enabled = False

End Sub

Private Sub cmdSearch_Click()
'<CSCM>
'********************************************************************************
'Procedure Name     : cmdSearch_Click
'Procedure Function : ---
'Input Parameter    : ---
'Output Parameter   : ---
'Date               : 3/9/2016
'LastUpdate/By      : Tedi / Kreatif
'********************************************************************************
'</CSCM>
    
    Dim strBrand_Filter As String
    Dim blnFound As Boolean
    Dim intPlan As Integer 'Jumlah plan
    Dim lsiSearch As Object
    Dim dblTextColor As Double
    'Dim pesan
    
    If Trim(cboBrandCode) = "" Or Trim(cboYear) = "" Then Exit Sub
    
    lblSearchResult.Caption = ""
    blnFound = False
    cmdClear.Enabled = True
    lvSearchResult.ListItems.Clear
    
    strSql = ""
    strSql = strSql & "SELECT brand_code "
    strSql = strSql & "FROM media_security_catalog "
    strSql = strSql & "WHERE user_name = '" & strLogin_User & "' "
    strSql = strSql & "AND position = 'Planner' "
    strSql = strSql & "AND valid_until>=(select getdate())"
    strBrand_Filter = strSql
    
    If cboBrandCode.ListIndex = -1 Then
        strSql = "select mp_number,"
        strSql = strSql & "[year],"
        strSql = strSql & "brand_code,"
        strSql = strSql & "brand_name,"
        strSql = strSql & "client_name,"
        strSql = strSql & "is_latest "
        strSql = strSql & "from mp_master "
        strSql = strSql & "where brand_name like '" & Replace(cboBrandCode.Text, "*", "%") & "' "
        strSql = strSql & "and [year] like '" & Replace(cboYear.Text, "*", "%") & "' "
        strSql = strSql & "and brand_code in (" & strBrand_Filter & ")"
        recTemp.Open strSql, ConnERP, 1, 3
    Else
        strSql = "select mp_number,[year],brand_code,brand_name,client_name,is_latest from mp_master where brand_code = '" & Mid(cboBrandCode.Text, 1, 4) & "' and [year] like '" & Replace(cboYear.Text, "*", "%") & "' and brand_code in (" & strBrand_Filter & ")"
        recTemp.Open strSql, ConnERP, 1, 3
    End If
    
    intPlan = 0
    While Not recTemp.EOF
        blnFound = True
        If recTemp(5) = 1 Then
            dblTextColor = Shape2.FillColor
        Else
            dblTextColor = Shape1.FillColor
        End If
        Set lsiSearch = lvSearchResult.ListItems.Add(, , recTemp(0))
        lsiSearch.ForeColor = dblTextColor
        lsiSearch.ListSubItems.Add , , recTemp(1)
        lsiSearch.ListSubItems(1).ForeColor = dblTextColor
        lsiSearch.ListSubItems.Add , , recTemp(2)
        lsiSearch.ListSubItems(2).ForeColor = dblTextColor
        lsiSearch.ListSubItems.Add , , recTemp(3)
        lsiSearch.ListSubItems(3).ForeColor = dblTextColor
        lsiSearch.ListSubItems.Add , , recTemp(4)
        lsiSearch.ListSubItems(4).ForeColor = dblTextColor
        intPlan = intPlan + 1
        recTemp.MoveNext
    Wend
    recTemp.Close
    lblSearchResult.Caption = intPlan & " item(s) blnFound!"
    cmd_set_is_latest.Enabled = False

End Sub

Private Sub Form_Load()
'<CSCM>
'********************************************************************************
'Procedure Name     : Form_Load
'Procedure Function : ---
'Input Parameter    : ---
'Output Parameter   : ---
'Date               : 3/9/2016
'LastUpdate/By      : Tedi / Kreatif
'********************************************************************************
'</CSCM>
    
    Dim strBrand_Filter As String
    Dim IntYear As Integer
    Dim strBrandSample As String
    
    '=========Resize Form==============
    resize Me
    '==================================
    
    blnBU2 = False
    strBrandSample = ""
    
    lblSearchResult.Caption = ""
    strBrand_Filter = ""
    strBrand_Filter = strBrand_Filter & "SELECT brand_code,brand_name "
    strBrand_Filter = strBrand_Filter & "FROM brand "
    strBrand_Filter = strBrand_Filter & "WHERE brand_code "
    strBrand_Filter = strBrand_Filter & "IN (select brand_code from media_security_catalog "
    strBrand_Filter = strBrand_Filter & "WHERE user_name = '" & strLogin_User & "' "
    strBrand_Filter = strBrand_Filter & "AND position = 'Planner' "
    strBrand_Filter = strBrand_Filter & "AND valid_until>=(select getdate()))"
    
    recTemp.Open strBrand_Filter, ConnERP, 1, 3
    
    If Not recTemp.EOF Then
        strBrandSample = recTemp("brand_code").Value
    End If
    While Not recTemp.EOF
        cboBrandCode.AddItem recTemp(0) & "->" & recTemp(1)
        recTemp.MoveNext
    Wend
    recTemp.Close
    For IntYear = (Year(Now()) - 5) To (Year(Now()) + 30)
        cboYear.AddItem IntYear
    Next
    cboYear.Text = Year(Now())
    cmd_set_is_latest.Enabled = False
    
    'CEK APAKAH USER BU2 ?
    If strBrandSample <> "" Then
        strSql = "select count(*) from client_special where client_code in (select client_code from brand where brand_code = '" & strBrandSample & "')"
        recTemp.Open strSql, ConnERP, 1, 3
        If recTemp(0) = 0 Then
            blnBU2 = True
        End If
        recTemp.Close
    End If
    
End Sub

Private Sub lvSearchResult_DblClick()
'<CSCM>
'********************************************************************************
'Procedure Name     : lvSearchResult_DblClick
'Procedure Function : ---
'Input Parameter    : ---
'Output Parameter   : ---
'Date               : 3/9/2016
'LastUpdate/By      : Tedi / Kreatif
'********************************************************************************
'</CSCM>
    
    Dim intMPNumber As Integer, blnExist As Boolean
    If lvSearchResult.ListItems.Count > 0 Then
        If lvSearchResult.SelectedItem.ForeColor = vbRed Then
            blnExist = False
            For intMPNumber = 0 To frm_MPInsertion.cboMPNumber.ListCount - 1
                If frm_MPInsertion.cboMPNumber.List(intMPNumber) = lvSearchResult.SelectedItem Then
                    blnExist = True
                End If
            Next
            If Not blnExist Then frm_MPInsertion.cboMPNumber.AddItem lvSearchResult.SelectedItem
        End If
        frm_MPInsertion.cboMPNumber.Text = lvSearchResult.SelectedItem
        Unload Me
    End If

End Sub

Private Sub lvSearchResult_ItemClick(ByVal Item As MSComctlLib.ListItem)
'<CSCM>
'********************************************************************************
'Procedure Name     : lvSearchResult_ItemClick
'Procedure Function : ---
'Input Parameter    : ---
'Output Parameter   : ---
'Date               : 3/9/2016
'LastUpdate/By      : Tedi / Kreatif
'********************************************************************************
'</CSCM>
    
    If Item.ForeColor = Shape1.FillColor And blnBU2 Then
        cmd_set_is_latest.Enabled = True
    Else
        cmd_set_is_latest.Enabled = False
    End If

End Sub
