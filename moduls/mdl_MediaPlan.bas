Attribute VB_Name = "mdl_MediaPlan"
Public objOpener As String
'
Public sRange_Jan As String
Public eRange_Jan As String

Public sRange_Feb As String
Public eRange_Feb As String

Public sRange_Mar As String
Public eRange_Mar As String
'
Public sRange_Apr As String
Public eRange_Apr As String

Public sRange_May As String
Public eRange_May As String

Public sRange_Jun As String
Public eRange_Jun As String

Public sRange_Jul As String
Public eRange_Jul As String

Public sRange_Aug As String
Public eRange_Aug As String

Public sRange_Sep As String
Public eRange_Sep As String

Public sRange_Oct As String
Public eRange_Oct As String

Public sRange_Nov As String
Public eRange_Nov As String

Public sRange_Dec As String
Public eRange_Dec As String

'Public Const intScreenWidthDev As Integer = 12000 'Default 800 x 600
'Public Const intScreenHeightDev As Integer = 9000 'Default 800 x 600
Public Const intScreenWidthDev As Integer = 15360 'Default 1024 x 768
Public Const intScreenHeightDev As Integer = 11520 'Default 1024 x 768

Public intFormWidth As Integer 'Original Form Width
Public intFormHeight As Integer 'Original Form Height

Public Function EngMonthName(month_number As Integer) As String
    Select Case month_number
        Case 1
            EngMonthName = "January"
        Case 2
            EngMonthName = "February"
        Case 3
            EngMonthName = "March"
        Case 4
            EngMonthName = "April"
        Case 5
            EngMonthName = "May"
        Case 6
            EngMonthName = "June"
        Case 7
            EngMonthName = "July"
        Case 8
            EngMonthName = "August"
        Case 9
            EngMonthName = "September"
        Case 10
            EngMonthName = "October"
        Case 11
            EngMonthName = "November"
        Case 12
            EngMonthName = "December"
        Case Else
            EngMonthName = "Invalid Month Number"
    End Select
End Function

Public Function EngMonthIndex(month_name As String) As Integer
    Select Case month_name
        Case "January"
            EngMonthIndex = 1
        Case "February"
            EngMonthIndex = 2
        Case "March"
            EngMonthIndex = 3
        Case "April"
            EngMonthIndex = 4
        Case "May"
            EngMonthIndex = 5
        Case "June"
            EngMonthIndex = 6
        Case "July"
            EngMonthIndex = 7
        Case "August"
            EngMonthIndex = 8
        Case "September"
            EngMonthIndex = 9
        Case "October"
            EngMonthIndex = 10
        Case "November"
            EngMonthIndex = 11
        Case "December"
            EngMonthIndex = 12
        Case Else
            EngMonthIndex = -1
    End Select
End Function

Public Function RemoveNumberFormat(Words As String)
'*****************************************************************************
' Nama Prosedur     :   RemoveNumberFormat
' Fungsi Prosedur   :   Menghilangkan pemisah ribuan pada angka (,)
' Parameter  Input  :   Words As string
' Parameter Output  :
' Tgl Pembuatan     :   09 Agustus 2004
' Last Update/By    :   09 Agustus 2004/Sistyo
'*****************************************************************************
    Dim counter As Integer
    RemoveNumberFormat = ""
    For counter = 1 To Len(Words)
        If Mid(Words, counter, 1) <> "," Then
            RemoveNumberFormat = RemoveNumberFormat & Mid(Words, counter, 1)
        End If
    Next
End Function

Public Function ReplaceNull(Words As Variant)
'*****************************************************************************
' Nama Prosedur     :   ReplaceNull
' Fungsi Prosedur   :   Mengganti Null dgn ""
' Parameter  Input  :   Words As string
' Parameter Output  :
' Tgl Pembuatan     :   08 Sept 2004
' Last Update/By    :   08 Sept 2004/Sistyo
'*****************************************************************************
    If IsNull(Words) Then
        ReplaceNull = ""
    Else
        ReplaceNull = Words
    End If
End Function

'=====================RESIZE FORM====================

Public Sub resize(objForm As Form, Optional strObjException As String)
    Dim Ctl As Control
    intFormWidth = objForm.Width
    intFormHeight = objForm.Height
    
    With objForm
        .Height = .Height / intScreenHeightDev * Screen.Height
        .Width = .Width / intScreenWidthDev * Screen.Width
        .Top = (Screen.Height - .Height) / 2
        .Left = (Screen.Width - .Width) / 2
    End With
    
    On Error Resume Next
    
    For Each Ctl In objForm.Controls
        If Len(strObjException) <> 0 Then
            If InStr(1, strObjException, "[" & Ctl.Name & "]") = 0 Then
                Ctl.Left = Ctl.Left / intFormWidth * objForm.Width
                Ctl.Width = Ctl.Width / intFormWidth * objForm.Width
                Ctl.Top = Ctl.Top / intFormHeight * objForm.Height
                Ctl.Height = Ctl.Height / intFormHeight * objForm.Height
            End If
        Else
            Ctl.Left = Ctl.Left / intFormWidth * objForm.Width
            Ctl.Width = Ctl.Width / intFormWidth * objForm.Width
            Ctl.Top = Ctl.Top / intFormHeight * objForm.Height
            Ctl.Height = Ctl.Height / intFormHeight * objForm.Height
        End If
    Next
    
End Sub

'=====================EO RESIZE FORM====================

'=====================DRAG AND DROP=====================

Public Function GetNodeIndex(ObjTreeView As Object, MouseY As Single) As Integer
'=======================================================================
'Function Name : GetNodeIndex
'Description   : Get Node index under Mouse Position
'Parameter     : -ObjTreeView as TreeView
'                 TreeView Control where the event Raised
'                -MouseY as Single
'                 Coordinate Y of Mouse position relatives to object
'Created / By  : 29 October 2004 / Sistyo
'Apply To      : TreeView
'=======================================================================
    Dim LinePos As Integer, idx As Integer, CurrLine As Integer, NodeIndex As Integer
    LinePos = (MouseY \ (29 * ObjTreeView.Font.Size)) + 1
    CurrLine = 0
    idx = 1
    NodeIndex = -1
    While CurrLine <> LinePos And idx < ObjTreeView.Nodes.Count
        If ObjTreeView.Nodes(idx).Visible Then
            CurrLine = CurrLine + 1
            NodeIndex = idx
        End If
        idx = idx + 1
    Wend
    GetNodeIndex = NodeIndex
End Function

Public Function GetListIndex(ObjListView As Object, MouseY As Single) As Integer
'=======================================================================
'Function Name : GetListIndex
'Description   : Get List index under Mouse Position
'Parameter     : -ObjListView As ListView,
'                 ListView Control where the event Raised
'                -MouseY as Single
'                 Coordinate Y of Mouse position relatives to object
'Created / By  : 29 October 2004 / Sistyo
'Apply To      : ListView
'=======================================================================
    Dim LinePos As Integer, CurrLine As Integer, ListIndex As Integer, idx As Integer
    LinePos = (MouseY \ (26 * ObjListView.Font.Size)) + 1
    CurrLine = 0
    ListIndex = -1
    If ObjListView.ListItems.Count <> 0 Then
        For idx = 1 To ObjListView.ListItems.Count
            If ObjListView.ListItems(idx).Text = ObjListView.GetFirstVisible Then
                ListIndex = idx + LinePos - 1
                Exit For
            End If
        Next
    End If
    If ListIndex > ObjListView.ListItems.Count Then ListIndex = ObjListView.ListItems.Count
    GetListIndex = ListIndex
End Function

Public Function GetListBIndex(ObjListBox As Object, MouseY As Single) As Integer
'=======================================================================
'Function Name : GetListBIndex
'Description   : Get List index under Mouse Position
'Parameter     : -ObjListBox As ListBox
'                 ListBox Control where the event Raised
'                -MouseY as Single
'                 Coordinate Y of Mouse position relatives to object
'Created / By  : 5 Nov 2004 / Sistyo
'Apply To      : ListBox
'=======================================================================
    Dim LinePos As Integer
    LinePos = MouseY \ (24 * ObjListBox.Font.Size)
    On Error GoTo ErrorLabel
    GetListBIndex = LinePos + ObjListBox.TopIndex
    If GetListBIndex > ObjListBox.ListCount - 1 Then GetListBIndex = ObjListBox.ListCount - 1
    Exit Function
ErrorLabel:
    GetListBIndex = -1
End Function
'=====================EO DRAG AND DROP=====================
