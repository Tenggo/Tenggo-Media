VERSION 5.00
Begin VB.Form Frm_Radio_Change_Material 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Radio Purchase Order Material Change"
   ClientHeight    =   1665
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   5415
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1665
   ScaleWidth      =   5415
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
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
      Left            =   1305
      TabIndex        =   2
      Top             =   960
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
      Left            =   120
      TabIndex        =   1
      Top             =   960
      Width           =   1180
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "Cl&ose"
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
      Left            =   4110
      TabIndex        =   3
      Top             =   960
      Width           =   1180
   End
   Begin VB.Frame fraMaterial 
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
      Height          =   735
      Left            =   105
      TabIndex        =   4
      Top             =   30
      Width           =   5190
      Begin VB.ComboBox cboMaterial 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   1560
         TabIndex        =   0
         Top             =   240
         Width           =   3495
      End
      Begin VB.Label lblMaterial 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Material Name :"
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
         TabIndex        =   5
         Top             =   270
         Width           =   1335
      End
   End
End
Attribute VB_Name = "Frm_Radio_Change_Material"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
''*************************************************************
'Fungsi Form        : mengubah  materi
'Last Update/By     :
'*************************************************************

Option Explicit

Dim strMatID As String

Private Sub Form_Load()
    Dim recMateri As New ADODB.Recordset
    Dim recDuration As New ADODB.Recordset
    Dim strDuration As String
    
    'dw - Hanya menampilkan materi yang durationnya sesuai
    strQuery = "SELECT DISTINCT duration FROM PO_Media_Radio_Date_Insertion WHERE po_number ='" & Trim(Frm_PO_Media_Radio.txtPONumber.Text) & "'"
    recDuration.Open strQuery, ConnERP, adOpenStatic, adLockReadOnly, adCmdText
    
    Do While Not recDuration.EOF
        strDuration = strDuration & "'" & recDuration.Fields("duration").Value & "',"
        recDuration.MoveNext
    Loop
    
    If strDuration <> "" Then
        strDuration = "(" & Mid(strDuration, 1, Len(strDuration) - 1) & ")"
        
        strQuery = "SELECT * FROM Ib_radio_material WHERE IB_ID='" & Trim(Frm_PO_Media_Radio.strIBID) & "'"
        strQuery = strQuery & "AND duration IN " & strDuration & ""
        
        recMateri.Open strQuery, ConnERP, adOpenStatic, adLockReadOnly, adCmdText
        
        With recMateri
            Do While Not .EOF
                cboMaterial.AddItem .Fields("material_id").Value & " " & .Fields("Material_Name").Value & "/" & .Fields("Duration").Value
                .MoveNext
            Loop
        End With
        
        recMateri.Close
        Set recMateri = Nothing
    Else
        'do nothing
    End If
    
    recDuration.Close
    Set recDuration = Nothing
    
    With Frm_PO_Media_Radio.msgBook
        strMatID = .TextMatrix(.Row, 1)
    End With
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdClose_Click()
    Unload Me
End Sub

Private Sub cmdSave_Click()
    Dim strMaterial As String
    Dim recMateri As New ADODB.Recordset
    
    If cboMaterial.Text = "" Then
        MsgBox "Please Select the Material !", vbCritical, strTitleMissingInfo
        Exit Sub
    End If
    
    strMaterial = "SELECT * FROM ib_radio_material WHERE "
    strMaterial = strMaterial & " ib_id ='" & Trim(Frm_PO_Media_Radio.strIBID) & "'"
    strMaterial = strMaterial & " AND material_id ='" & Left(cboMaterial.Text, 1) & "'"
    recMateri.Open strMaterial, ConnERP, adOpenStatic, adLockReadOnly
    
    With recMateri
        If .EOF = False Then
            strQuery = "UPDATE po_media_radio_date_insertion "
            strQuery = strQuery & "SET material_id ='" & Trim(Left(cboMaterial.Text, 1)) & "'"
            strQuery = strQuery & ", material_NAME ='" & Clear_String(Trim(.Fields("material_name").Value)) & "'"
            strQuery = strQuery & ", Duration =" & .Fields("duration").Value
            strQuery = strQuery & " WHERE po_number ='" & Trim(Frm_PO_Media_Radio.txtPONumber.Text) & "'"
            strQuery = strQuery & " AND material_id ='" & Trim(Left(strMatID, 1)) & "'"
            strQuery = strQuery & " AND Spot_Type ='" & Trim(Left(Frm_PO_Media_Radio.msgBook.TextMatrix(Frm_PO_Media_Radio.msgBook.Row, 3), 2)) & "'"
            ConnERP.Execute strQuery
        End If
    End With
    
    recMateri.Close
    Set recMateri = Nothing
    
    Call Frm_PO_Media_Radio.InitialGrid(Val(Frm_PO_Media_Radio.cboMonth.Text), Val(Frm_PO_Media_Radio.cboYear.Text))
    
    Call Frm_PO_Media_Radio.LoadBookGrid(Frm_PO_Media_Radio.msgBook)
    
    Unload Me
End Sub

Private Sub cboMaterial_KeyPress(KeyAscii As Integer)
    KeyAscii = 0
End Sub
