VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFlxGrd.ocx"
Begin VB.Form Frm_Date_Commencing 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Print Dates Commencing"
   ClientHeight    =   5010
   ClientLeft      =   150
   ClientTop       =   780
   ClientWidth     =   11430
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5010
   ScaleWidth      =   11430
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox picToolbar 
      Align           =   1  'Align Top
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ClipControls    =   0   'False
      ForeColor       =   &H80000008&
      Height          =   750
      Left            =   0
      ScaleHeight     =   750
      ScaleWidth      =   11430
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   0
      Width           =   11430
      Begin VB.PictureBox picButton 
         Appearance      =   0  'Flat
         BackColor       =   &H00F0F0F0&
         BorderStyle     =   0  'None
         ClipControls    =   0   'False
         ForeColor       =   &H80000008&
         Height          =   750
         Index           =   10
         Left            =   90
         Picture         =   "Frm_Date_Commencing.frx":0000
         ScaleHeight     =   750
         ScaleWidth      =   1500
         TabIndex        =   6
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
         Index           =   11
         Left            =   1620
         Picture         =   "Frm_Date_Commencing.frx":1C02
         ScaleHeight     =   750
         ScaleWidth      =   1500
         TabIndex        =   5
         TabStop         =   0   'False
         Top             =   0
         Width           =   1500
      End
   End
   Begin VB.PictureBox Pc_Tanggal 
      BackColor       =   &H00C0C0C0&
      Height          =   4245
      Left            =   -45
      ScaleHeight     =   4185
      ScaleWidth      =   11385
      TabIndex        =   0
      Top             =   765
      Width           =   11445
      Begin MSFlexGridLib.MSFlexGrid Grd_Tanggal 
         Height          =   3300
         Left            =   210
         TabIndex        =   1
         Top             =   450
         Width           =   10965
         _ExtentX        =   19341
         _ExtentY        =   5821
         _Version        =   393216
         Rows            =   3
         FixedCols       =   0
         BackColor       =   16777215
         ForeColor       =   0
         AllowBigSelection=   0   'False
         HighLight       =   2
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
      Begin VB.Label Label12 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Put detail date"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   270
         Left            =   4980
         TabIndex        =   3
         Top             =   90
         Width           =   1635
      End
      Begin VB.Label Lbl_Ket 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Double Click To Add"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   210
         TabIndex        =   2
         Top             =   3870
         Width           =   1755
      End
   End
   Begin VB.Menu MnuFreezeUnfreeze 
      Caption         =   "MnuFreezeUnfreeze"
      Begin VB.Menu MnuFreeze 
         Caption         =   "Freeze / Unfreeze"
      End
   End
End
Attribute VB_Name = "Frm_Date_Commencing"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'
Option Explicit
Dim cocok As Integer
'''
Private Sub db_Cancel()
    keluar = True
    Unload Me
End Sub

Private Sub db_Save()
    Dim IntSumSpot As Integer
    Dim IntSumSpotMax As Integer

    If cocok = rsWeekCommencing.RecordCount Then
        keluar = False

        Me.Hide
    Else
        If MsgBox("Date selection is not complete yet, Are sure you want to continue.", vbYesNo + vbExclamation, "Confirmation") = vbYes Then
            keluar = False
            Me.Hide
        End If
    End If

End Sub

Private Sub Cmd_save_Click()

End Sub

Private Sub Form_Load()
    Dim int_current_row As Integer, Str_WC As String, int_Header_row As Integer
    Dim i As Integer
       
    cocok = 0
    MnuFreezeUnfreeze.Visible = False
    
    With Grd_Tanggal
        .FixedRows = 1
        .FixedCols = 1
        .Rows = 1
        .cols = 26
        .Clear
        .ColWidth(0) = 2000 'week commencing
        .ColWidth(1) = 1250 'version
        .ColWidth(2) = 1700 'media
        .ColWidth(3) = 3000 'dimension
        .ColWidth(4) = 1500 'spot type
        .ColWidth(5) = 750 'spot
        .ColWidth(6) = 0 'print code
        .ColWidth(7) = 0 'print size code
        .ColWidth(8) = 0 'print color_code
        .ColWidth(9) = 0 'print_paper_code
        .ColWidth(10) = 0 'print mmc col
        .ColWidth(11) = 0 'print mmc size
        .ColWidth(12) = 0 'print min size
        .ColWidth(13) = 0 'nett rate
        .ColWidth(14) = 0 'gross rate
        .ColWidth(15) = 0
        .ColWidth(16) = 500
        .ColWidth(17) = 500
        .ColWidth(18) = 500
        .ColWidth(19) = 500
        .ColWidth(20) = 500
        .ColWidth(21) = 500
        .ColWidth(22) = 500
        .ColWidth(23) = 0
        .ColWidth(24) = 0
        .ColWidth(25) = 0
        .ColAlignment(1) = vbAlignNone

        int_current_row = 0
        int_Header_row = 0
        Str_WC = ""
        
        While Not rsWeekCommencing.EOF
            If rsWeekCommencing("week_commencing") <> Str_WC Then
                If .Rows <> 1 Then .Rows = .Rows + 1
                int_current_row = .Rows - 1
                int_Header_row = int_current_row
                .Row = int_current_row
                
                'Header Row
                For i = 1 To rsWeekCommencing.Fields.Count
                    .col = i - 1
                    .Text = rsWeekCommencing(i - 1).Name
                    .CellBackColor = .BackColorFixed
                    .CellAlignment = 4
                Next
                
                'Header Date
                For i = 1 To 7
                    .col = i + 15
                    '.Text = Day(rsWeekCommencing(0)) + (i - 1) & "/" & month(rsWeekCommencing(0))
                    .Text = Format(DateAdd("d", (i - 1), rsWeekCommencing(0)), "d/m")
                    .CellBackColor = .BackColorFixed
                Next
                Str_WC = rsWeekCommencing("week_commencing")
            End If
            
            .Rows = .Rows + 1
            int_current_row = .Rows - 1
            .Row = int_current_row
            For i = 1 To rsWeekCommencing.Fields.Count
                .col = i - 1
                .Text = " " & rsWeekCommencing(i - 1)
            Next
            .col = .cols - 1
            .Text = int_Header_row
            .col = .cols - 2
            .Text = 0
            .col = .cols - 3
            .Text = ""
'            If .TextMatrix(.Row, 0) = .TextMatrix(.Row - 1, 0) Then .TextMatrix(.Row, 0) = ""
            rsWeekCommencing.MoveNext
        Wend
    End With
End Sub

Private Sub Grd_Tanggal_Click()

    If Trim(Grd_Tanggal.TextMatrix(Grd_Tanggal.Row, Grd_Tanggal.col)) = "" Then
        Lbl_Ket.Caption = "Double Click to Add"
    Else
        Lbl_Ket.Caption = "Double Click to Remove"
    End If

End Sub

Private Sub Grd_Tanggal_DblClick()
    Dim str_date As String
    Dim Str_Spot_Type As String
    Dim strTemp As String
    Dim strTempReg As String
    Dim strTempSpec As String
    Dim IntSpot_Max As Integer
    Dim IntSpot As Integer

    With Grd_Tanggal
        If .col > 15 And .TextMatrix(.Row, .cols - 1) <> "" Then
            str_date = .TextMatrix(.TextMatrix(.Row, .cols - 1), .col)
            Str_Spot_Type = Trim(.TextMatrix(.Row, 4))
            IntSpot_Max = .TextMatrix(.Row, 5)
            IntSpot = .TextMatrix(.Row, .cols - 2)

                If Str_Spot_Type = "Reguler" Then
                    If Trim(.TextMatrix(.Row, .col)) = "" Then
                        If IntSpot < IntSpot_Max Then
                            .CellAlignment = 4
                            .TextMatrix(.Row, .col) = "P"
                            .CellBackColor = vbGreen
                            Lbl_Ket.Caption = "Double Click to Remove"
                            .TextMatrix(.Row, .cols - 2) = .TextMatrix(.Row, .cols - 2) + 1
                            IntSpot = .TextMatrix(.Row, .cols - 2)
                            If IntSpot = IntSpot_Max Then
                                cocok = cocok + 1
                            End If
                            If .TextMatrix(.Row, .cols - 3) = "" Then
                                .TextMatrix(.Row, .cols - 3) = str_date
                            Else
                                .TextMatrix(.Row, .cols - 3) = .TextMatrix(.Row, .cols - 3) & "," & str_date
                            End If
                        Else
                            MsgBox "Sorry, Can't Add Date Commencing", vbExclamation, strCompany_Name
                        End If
                    Else
                        .TextMatrix(.Row, .col) = ""
                        .CellBackColor = vbWhite
                        Lbl_Ket.Caption = "Double Click to Add"
                        .TextMatrix(.Row, .cols - 2) = .TextMatrix(.Row, .cols - 2) - 1
                        IntSpot = .TextMatrix(.Row, .cols - 2)
                        If IntSpot_Max - IntSpot = 1 Then
                            cocok = cocok - 1
                        End If
                        strTemp = Replace("," & .TextMatrix(.Row, .cols - 3) & ",", "," & str_date & ",", ",", 1, 1)
                        strTempReg = Replace("," & .TextMatrix(.Row, .cols - 3) & ",", "," & str_date & ",", ",", 1, 1)
                        If strTemp = "," Then
                            .TextMatrix(.Row, .cols - 3) = ""
                        Else
                            .TextMatrix(.Row, .cols - 3) = Mid(strTemp, 2, Len(strTemp) - 2)
                            If strTempReg = "," Then
                                .TextMatrix(.Row, .cols - 3) = ""
                            Else
                                .TextMatrix(.Row, .cols - 3) = Mid(strTempReg, 2, Len(strTempReg) - 2)
                            End If
                        End If
    
                    End If

            ElseIf Str_Spot_Type = "Special Buys" Then

                If Trim(.TextMatrix(.Row, .col)) = "" Then
                    If IntSpot < IntSpot_Max Then
                        .CellAlignment = 4
                        .TextMatrix(.Row, .col) = "P"
                        .CellBackColor = vbYellow
                        Lbl_Ket.Caption = "Double Click to Remove"
                        .TextMatrix(.Row, .cols - 2) = .TextMatrix(.Row, .cols - 2) + 1
                        IntSpot = .TextMatrix(.Row, .cols - 2)
                            If IntSpot = IntSpot_Max Then
                                cocok = cocok + 1
                            End If
                        If .TextMatrix(.Row, .cols - 3) = "" Then
                            .TextMatrix(.Row, .cols - 3) = str_date
                        Else
                            .TextMatrix(.Row, .cols - 3) = .TextMatrix(.Row, .cols - 3) & "," & str_date
                        End If
                    Else
                        MsgBox "Sorry, Can't Add Date Commencing", vbExclamation, strCompany_Name
                    End If
                Else
                    .TextMatrix(.Row, .col) = ""
                    .CellBackColor = vbWhite
                    Lbl_Ket.Caption = "Double Click to Add"
                    .TextMatrix(.Row, .cols - 2) = .TextMatrix(.Row, .cols - 2) - 1
                    IntSpot = .TextMatrix(.Row, .cols - 2)
                    If IntSpot_Max - IntSpot = 1 Then
                        cocok = cocok - 1
                    End If
                    strTemp = Replace("," & .TextMatrix(.Row, .cols - 3) & ",", "," & str_date & ",", ",", 1, 1)
                    strTempSpec = Replace("," & .TextMatrix(.Row, .cols - 3) & ",", "," & str_date & ",", ",", 1, 1)
                    If strTemp = "," Then
                        .TextMatrix(.Row, .cols - 3) = ""
                    Else
                        .TextMatrix(.Row, .cols - 3) = Mid(strTemp, 2, Len(strTemp) - 2)
                        If strTempSpec = "," Then
                            .TextMatrix(.Row, .cols - 3) = ""
                        Else
                            .TextMatrix(.Row, .cols - 3) = Mid(strTempSpec, 2, Len(strTempSpec) - 2)
                        End If
                    End If

                End If
            End If
        End If

    End With
End Sub

Private Sub Grd_Tanggal_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    If Button = 2 Then
        
        With Grd_Tanggal
            .col = .MouseCol
            .Row = .MouseRow
        End With
        
        Me.PopupMenu MnuFreezeUnfreeze
                     
        
    End If

End Sub

Private Sub MnuFreeze_Click()

    With Grd_Tanggal
        If .col < .FixedCols Then
            .FixedCols = .col
        Else
            .FixedCols = .col + 1
        End If
    End With

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

    Dim strCode As String, strFileRpt As String
    'Lock_MainForm True
    Select Case Index
        Case enButtonType.bieSave   'Save
            db_Save
        Case enButtonType.bieCancel  'Cancel.
            db_Cancel
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



