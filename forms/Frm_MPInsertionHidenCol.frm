VERSION 5.00
Begin VB.Form Frm_MPInsertionHidenCol 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Hiden Columns"
   ClientHeight    =   2160
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4605
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2160
   ScaleWidth      =   4605
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CheckBox cbk_all 
      Caption         =   "Check All"
      Height          =   285
      Left            =   3465
      TabIndex        =   3
      Top             =   105
      Width           =   1005
   End
   Begin VB.CommandButton CmdOK 
      Caption         =   "OK"
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
      Left            =   3435
      TabIndex        =   2
      ToolTipText     =   "view selected columns"
      Top             =   1680
      Width           =   1005
   End
   Begin VB.ListBox lstColumn 
      Appearance      =   0  'Flat
      Height          =   1155
      Left            =   120
      Style           =   1  'Checkbox
      TabIndex        =   0
      Top             =   405
      Width           =   4335
   End
   Begin VB.Label lblpesan 
      Caption         =   "Select Column :"
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
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   3195
   End
End
Attribute VB_Name = "Frm_MPInsertionHidenCol"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'
'*****************************************************************************
' Nama Submodul         :  Frm_MPInsertionHidenCol
' Fungsi Submodul       :  menampilkan kolom yang di hiden
' Nama Programmer       :  Sistyo
' Tgl Pembuatan         :  19 Agustus 2004
' Last Update           :  19 Agustus 2004/Sistyo
'******************************************************************************
Private Sub Form_Load()
'Menampilkan list kolom yang di hiden
    Dim i As Integer 'column number
    With frm_MPInsertion.LstHidenCol
        lstColumn.Clear
        If .ListCount <> 0 Then
            lblpesan.Caption = "Select column :"
            For i = 1 To .ListCount
                lstColumn.AddItem .List(i - 1)
            Next
        Else
            lblpesan.Caption = "There are no hiden column..."
        End If
    End With
End Sub

Private Sub cmdOK_Click()
'Menampilkan kolom yang dipilih

    Dim i As Integer 'column number in list
    Dim intColNum As Integer 'column number in grid
    Dim intColWidth As Integer ' column width
    Dim j As Integer
    Dim strMonth As String
    With lstColumn
        If .ListCount <> 0 Then
            For i = .ListCount To 1 Step -1
                If .Selected(i - 1) Then
                    intColWidth = Val(Mid(.List(i - 1), InStr(1, .List(i - 1), ":") + 2, InStr(1, .List(i - 1), ">") - 7))
                    If Left(.List(i - 1), 2) = "#_" Then
                        'Month
                        strMonth = Trim(Right(.List(i - 1), Len(.List(i - 1)) - InStr(1, .List(i - 1), ">")))
                        j = 0
                        Do
                            j = j + 1
                        Loop Until frm_MPInsertion.msf_MPInsertion.TextMatrix(1, j) = strMonth
                        
                        While frm_MPInsertion.msf_MPInsertion.TextMatrix(1, j) = strMonth
                            frm_MPInsertion.msf_MPInsertion.ColWidth(j) = intColWidth
                            j = j + 1
                        Wend
                        
                    Else
                        'Bukan Month
                        intColNum = Val(Mid(.List(i - 1), 2, InStr(1, .List(i - 1), ":") - 3))
                        frm_MPInsertion.msf_MPInsertion.ColWidth(intColNum) = intColWidth
                    End If
                    frm_MPInsertion.LstHidenCol.RemoveItem (i - 1)
                End If
            Next
        End If
    End With
    
    Unload Me
    
End Sub

Private Sub cbk_all_Click()
    Dim i As Integer
    For i = 1 To lstColumn.ListCount
        lstColumn.Selected(i - 1) = cbk_all.Value
    Next
End Sub
