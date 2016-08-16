VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFlxGrd.ocx"
Begin VB.Form Frm_PO_Media_Radio 
   BorderStyle     =   0  'None
   Caption         =   "Purchase Order Radio Media"
   ClientHeight    =   8985
   ClientLeft      =   45
   ClientTop       =   630
   ClientWidth     =   10695
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "Frm_PO_Media_Radio.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8985
   ScaleWidth      =   10695
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox picToolbar 
      Align           =   1  'Align Top
      Appearance      =   0  'Flat
      BackColor       =   &H000000FF&
      BorderStyle     =   0  'None
      ClipControls    =   0   'False
      ForeColor       =   &H80000008&
      Height          =   750
      Left            =   0
      ScaleHeight     =   750
      ScaleWidth      =   10695
      TabIndex        =   35
      Top             =   0
      Width           =   10695
      Begin VB.PictureBox picButton 
         Appearance      =   0  'Flat
         BackColor       =   &H00F0F0F0&
         BorderStyle     =   0  'None
         ClipControls    =   0   'False
         ForeColor       =   &H80000008&
         Height          =   750
         Index           =   5
         Left            =   135
         ScaleHeight     =   750
         ScaleWidth      =   1500
         TabIndex        =   40
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
         Index           =   23
         Left            =   3195
         ScaleHeight     =   750
         ScaleWidth      =   1500
         TabIndex        =   39
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
         Index           =   10
         Left            =   135
         ScaleHeight     =   750
         ScaleWidth      =   1500
         TabIndex        =   38
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
         Index           =   8
         Left            =   1665
         ScaleHeight     =   750
         ScaleWidth      =   1500
         TabIndex        =   36
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
         Left            =   1665
         ScaleHeight     =   750
         ScaleWidth      =   1500
         TabIndex        =   37
         Top             =   0
         Width           =   1500
      End
   End
   Begin Threed.SSPanel pnl_Main 
      Align           =   2  'Align Bottom
      Height          =   8190
      Left            =   0
      TabIndex        =   0
      Top             =   795
      Width           =   10695
      _Version        =   65536
      _ExtentX        =   18865
      _ExtentY        =   14446
      _StockProps     =   15
      BackColor       =   15790320
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BevelOuter      =   1
      Begin VB.Frame fraRemark 
         Appearance      =   0  'Flat
         Caption         =   "Remark"
         ForeColor       =   &H80000008&
         Height          =   2055
         Left            =   120
         TabIndex        =   33
         Top             =   3990
         Width           =   10455
         Begin VB.TextBox txtDescription 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   1470
            Left            =   150
            MaxLength       =   500
            MultiLine       =   -1  'True
            TabIndex        =   34
            Top             =   345
            Width           =   10140
         End
      End
      Begin VB.Frame fraBooked 
         Appearance      =   0  'Flat
         Caption         =   "Booked Date"
         ForeColor       =   &H80000008&
         Height          =   2055
         Left            =   120
         TabIndex        =   31
         Top             =   3990
         Width           =   10455
         Begin MSFlexGridLib.MSFlexGrid msgBook 
            Height          =   1560
            Left            =   120
            TabIndex        =   32
            Top             =   300
            Width           =   10215
            _ExtentX        =   18018
            _ExtentY        =   2752
            _Version        =   393216
            Cols            =   42
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
      End
      Begin VB.Frame fraTotal 
         Appearance      =   0  'Flat
         Caption         =   "Total"
         ForeColor       =   &H80000008&
         Height          =   1215
         Left            =   5520
         TabIndex        =   26
         Top             =   6150
         Width           =   5055
         Begin VB.TextBox txtGross 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            Height          =   325
            Left            =   1680
            TabIndex        =   28
            Top             =   285
            Width           =   3000
         End
         Begin VB.TextBox txtNett 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            Height          =   325
            Left            =   1680
            TabIndex        =   27
            Top             =   705
            Width           =   3000
         End
         Begin VB.Label lblGross 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Total Gross :"
            Height          =   255
            Left            =   480
            TabIndex        =   30
            Top             =   330
            Width           =   1095
         End
         Begin VB.Label lblNett 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Total Nett :"
            Height          =   255
            Left            =   480
            TabIndex        =   29
            Top             =   750
            Width           =   1095
         End
      End
      Begin VB.Frame fraHeader2 
         Appearance      =   0  'Flat
         ForeColor       =   &H80000008&
         Height          =   1935
         Left            =   120
         TabIndex        =   17
         Top             =   1950
         Width           =   10455
         Begin VB.TextBox txtPONumber 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            Height          =   325
            Left            =   8010
            Locked          =   -1  'True
            TabIndex        =   21
            Top             =   270
            Width           =   2175
         End
         Begin VB.ComboBox cboStation 
            Height          =   315
            Left            =   1515
            Style           =   2  'Dropdown List
            TabIndex        =   20
            Top             =   270
            Width           =   4290
         End
         Begin VB.TextBox txtAddress 
            Appearance      =   0  'Flat
            Height          =   1080
            Left            =   1515
            Locked          =   -1  'True
            MultiLine       =   -1  'True
            ScrollBars      =   1  'Horizontal
            TabIndex        =   19
            Top             =   660
            Width           =   4275
         End
         Begin VB.TextBox txtPODate 
            Appearance      =   0  'Flat
            Height          =   325
            Left            =   8010
            Locked          =   -1  'True
            TabIndex        =   18
            Top             =   660
            Width           =   2175
         End
         Begin VB.Label lblAddress 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Address :"
            Height          =   255
            Left            =   435
            TabIndex        =   25
            Top             =   690
            Width           =   1020
         End
         Begin VB.Label lblOrderNo 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Order No :"
            Height          =   255
            Left            =   6600
            TabIndex        =   24
            Top             =   300
            Width           =   1335
         End
         Begin VB.Label lblTo 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "To :"
            Height          =   255
            Left            =   435
            TabIndex        =   23
            Top             =   300
            Width           =   1020
         End
         Begin VB.Label lblOrderDate 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Order Date :"
            Height          =   255
            Left            =   6600
            TabIndex        =   22
            Top             =   690
            Width           =   1335
         End
      End
      Begin VB.Frame fraNoteCode 
         Appearance      =   0  'Flat
         Caption         =   "Note Code"
         ForeColor       =   &H80000008&
         Height          =   1215
         Left            =   120
         TabIndex        =   14
         Top             =   6150
         Width           =   5295
         Begin VB.ComboBox cboNoteCode 
            Height          =   315
            ItemData        =   "Frm_PO_Media_Radio.frx":0442
            Left            =   225
            List            =   "Frm_PO_Media_Radio.frx":0444
            TabIndex        =   16
            Text            =   "cboNoteCode"
            Top             =   465
            Width           =   3375
         End
         Begin VB.CommandButton cmdNote 
            Caption         =   "Note"
            Height          =   495
            Left            =   3855
            TabIndex        =   15
            Top             =   390
            Width           =   1180
         End
      End
      Begin VB.Frame fraHeader1 
         Appearance      =   0  'Flat
         ForeColor       =   &H80000008&
         Height          =   1815
         Left            =   120
         TabIndex        =   1
         Top             =   45
         Width           =   10455
         Begin VB.TextBox txtBrandVariant 
            Appearance      =   0  'Flat
            ForeColor       =   &H00000000&
            Height          =   325
            Left            =   1530
            Locked          =   -1  'True
            TabIndex        =   7
            Top             =   1350
            Width           =   4275
         End
         Begin VB.ComboBox cboMonth 
            Height          =   315
            Left            =   1530
            Style           =   2  'Dropdown List
            TabIndex        =   6
            Top             =   585
            Width           =   2340
         End
         Begin VB.ComboBox cboYear 
            Height          =   315
            Left            =   1530
            Style           =   2  'Dropdown List
            TabIndex        =   5
            Top             =   210
            Width           =   1065
         End
         Begin VB.ComboBox cboBrand 
            Height          =   315
            Left            =   1530
            Style           =   2  'Dropdown List
            TabIndex        =   4
            Top             =   960
            Width           =   4275
         End
         Begin VB.TextBox txtJobID 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            Height          =   325
            Left            =   7995
            Locked          =   -1  'True
            TabIndex        =   3
            Top             =   210
            Width           =   2175
         End
         Begin VB.ComboBox cboJobNumber 
            Height          =   315
            Left            =   7995
            Style           =   2  'Dropdown List
            TabIndex        =   2
            Top             =   585
            Width           =   2175
         End
         Begin VB.Label lblJobNumber 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Job Number :"
            Height          =   255
            Left            =   6795
            TabIndex        =   13
            Top             =   615
            Width           =   1125
         End
         Begin VB.Label lblBrandVariant 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Brand Variant :"
            ForeColor       =   &H00000000&
            Height          =   255
            Left            =   225
            TabIndex        =   12
            Top             =   1380
            Width           =   1245
         End
         Begin VB.Label lblMonth 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Month :"
            Height          =   255
            Left            =   225
            TabIndex        =   11
            Top             =   615
            Width           =   1245
         End
         Begin VB.Label lblYear 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Year :"
            Height          =   255
            Left            =   225
            TabIndex        =   10
            Top             =   240
            Width           =   1245
         End
         Begin VB.Label lblBrand 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Brand :"
            Height          =   255
            Left            =   225
            TabIndex        =   9
            Top             =   990
            Width           =   1245
         End
         Begin VB.Label lblJobID 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Job ID :"
            Height          =   255
            Left            =   6795
            TabIndex        =   8
            Top             =   240
            Width           =   1125
         End
      End
      Begin Crystal.CrystalReport crPO 
         Left            =   5235
         Top             =   7830
         _ExtentX        =   741
         _ExtentY        =   741
         _Version        =   348160
         PrintFileLinesPerPage=   60
      End
      Begin VB.Line Line2 
         X1              =   0
         X2              =   15210
         Y1              =   7590
         Y2              =   7590
      End
   End
   Begin VB.Menu mnuPopUp 
      Caption         =   "Cancel and Change Material"
      Begin VB.Menu mnuCancelTemplate 
         Caption         =   "Cancel Template"
      End
      Begin VB.Menu mnuEditMateri 
         Caption         =   "Change Materi"
      End
      Begin VB.Menu mnuChangeMateriDescription 
         Caption         =   "Change Materi Description"
         Visible         =   0   'False
      End
   End
End
Attribute VB_Name = "Frm_PO_Media_Radio"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
''dy saat klik edit, masih  bisa pilih radio, tapi button nya jadi salah->ok
'dy untuk meng-edit remark (special job) lewat mana ya?->ok

Option Explicit

Public strIBID As String

Public blnLoadCancelFlag As Boolean

Dim recHeader As New ADODB.Recordset
Dim recAddCostDetail As New ADODB.Recordset
Dim recAddCostDetailOld As New ADODB.Recordset

Dim dblOldNettSpecial As Double
Dim dblOldGrossSpecial As Double

Public dblOldGross As Double
Public dblOldMSC As Double
Public dblOldVAT As Double
Dim dblOldNett As Double

Dim dblNewNett As Double
Dim dblNewGross As Double
Dim dblNewMSC As Double
Dim dblNewVAT As Double
Dim blnEditOrAdd As Boolean

Private Sub cmdSaveNote_Click()

End Sub



Private Sub Form_Load()
    Dim intPos As Integer
    
    Me.AutoRedraw = True
    Call VGradient(Me, &HFF8090, &HFF8090, Me.Line2.Y1, Me.Height, 0, Me.Width)
    
    LoadYear cboYear
    
    Call LoadNote
    
    blnLoadCancelFlag = True
    
    For intPos = 1 To 12
        cboMonth.AddItem Format(intPos, "00") & " - " & Get_Month_Name(intPos)
    Next intPos
            
    LoadBrand cboBrand, strLogin_User, "'Implementor'"

   
    recDate.Requery
    cboMonth.ListIndex = month(recDate(0)) - 1
    
    If cboMonth.ListIndex <> -1 Or cboYear.ListIndex <> -1 Then
        Call InitialGrid(cboMonth.ListIndex + 1, cboYear.ListIndex + 2000)
    End If
    
    Call ButtonLock(False)
    txtGross.Locked = True
    txtNett.Locked = True
    txtDescription.Locked = True
    EnableObject False
    
End Sub

Private Sub CreateTempTable()
    Set recAddCostDetail = Nothing
    Set recAddCostDetail = New ADODB.Recordset
    
    With recAddCostDetail.Fields
        .Append "PO_Number", adVarChar, 14, adFldMayBeNull
        .Append "Station_Code", adVarChar, 5, adFldMayBeNull
        .Append "Cost_Item", adChar, 15, adFldMayBeNull
        .Append "Gross", adCurrency, , adFldMayBeNull
        .Append "Nett", adCurrency, , adFldMayBeNull
        .Append "MSC", adCurrency, , adFldMayBeNull
        .Append "VAT", adCurrency, , adFldMayBeNull
        .Append "Tax_Flag", adSmallInt, , adFldMayBeNull
        .Append "MSC_Flag", adSmallInt, , adFldMayBeNull
        .Append "Cancel_Flag", adSmallInt, , adFldMayBeNull
    End With
    
    recAddCostDetail.Open , , adOpenDynamic, adLockOptimistic
    
    Set recAddCostDetailOld = Nothing
    Set recAddCostDetailOld = New ADODB.Recordset
    
    With recAddCostDetailOld.Fields
        .Append "PO_Number", adVarChar, 14, adFldMayBeNull
        .Append "Station_Code", adVarChar, 5, adFldMayBeNull
        .Append "Cost_Item", adChar, 15, adFldMayBeNull
        .Append "Gross", adCurrency, 15, adFldMayBeNull
        .Append "Nett", adCurrency, , adFldMayBeNull
        .Append "MSC", adCurrency, , adFldMayBeNull
        .Append "VAT", adCurrency, , adFldMayBeNull
        .Append "Tax_Flag", adSmallInt, , adFldMayBeNull
        .Append "MSC_Flag", adSmallInt, , adFldMayBeNull
        .Append "Cancel_Flag", adSmallInt, , adFldMayBeNull
    End With
    
    recAddCostDetailOld.Open , , adOpenDynamic, adLockOptimistic
End Sub

Private Sub PrepareData()
    Dim recPORadio As New ADODB.Recordset
    
    Call CreateTempTable
    
    dblOldNett = 0
    dblOldGross = 0
    dblOldMSC = 0
    dblOldVAT = 0
    
    dblNewNett = 0
    dblNewGross = 0
    dblNewMSC = 0
    dblNewVAT = 0
    
    strQuery = "SELECT * FROM PO_media_radio_addcost_detail WHERE po_Number='" & Me.txtPONumber.Text & "' "
    strQuery = strQuery & " AND (Cancel_Flag=0 OR Cancel_Flag=2)"
    
    recPORadio.CursorLocation = adUseClient
    recPORadio.Open strQuery, ConnERP, adOpenForwardOnly, adLockReadOnly
    
    With recAddCostDetail
        While Not recPORadio.EOF
            .AddNew
            .Fields("po_number").Value = recPORadio.Fields("po_number").Value
            .Fields("station_code").Value = recPORadio.Fields("station_code").Value
            .Fields("cost_item").Value = recPORadio.Fields("cost_item").Value
            .Fields("gross").Value = recPORadio.Fields("gross").Value
            .Fields("nett").Value = recPORadio.Fields("nett").Value
            .Fields("msc").Value = recPORadio.Fields("msc").Value
            .Fields("VAT").Value = recPORadio.Fields("VAT").Value
            .Fields("MSC_Flag").Value = recPORadio.Fields("MSC_Flag").Value
            .Fields("tax_Flag").Value = recPORadio.Fields("Tax_Flag").Value
            .Fields("Cancel_Flag").Value = recPORadio.Fields("Cancel_Flag").Value
            
            dblOldNett = dblOldNett + recPORadio.Fields("nett").Value
            dblOldGross = dblOldGross + recPORadio.Fields("gross").Value
            dblOldMSC = dblOldMSC + recPORadio.Fields("msc").Value
            dblOldVAT = dblOldVAT + recPORadio.Fields("vat").Value
            
            .Update
            
            'old
            recAddCostDetailOld.AddNew
            recAddCostDetailOld.Fields("po_number").Value = recPORadio.Fields("po_number").Value
            recAddCostDetailOld.Fields("station_code").Value = recPORadio.Fields("station_code").Value
            recAddCostDetailOld.Fields("cost_item").Value = recPORadio.Fields("cost_item").Value
            recAddCostDetailOld.Fields("gross").Value = recPORadio.Fields("gross").Value
            recAddCostDetailOld.Fields("nett").Value = recPORadio.Fields("nett").Value
            recAddCostDetailOld.Fields("msc").Value = recPORadio.Fields("msc").Value
            recAddCostDetailOld.Fields("VAT").Value = recPORadio.Fields("VAT").Value
            recAddCostDetailOld.Fields("MSC_Flag").Value = recPORadio.Fields("MSC_Flag").Value
            recAddCostDetailOld.Fields("tax_Flag").Value = recPORadio.Fields("tax_Flag").Value
            recAddCostDetailOld.Fields("Cancel_Flag").Value = recPORadio.Fields("Cancel_Flag").Value
            .Update
            recPORadio.MoveNext
        Wend
        
        recPORadio.Close
        Set recPORadio = Nothing
        
        dblNewNett = dblOldNett
        dblNewGross = dblOldGross
        dblNewMSC = dblOldMSC
        dblNewVAT = dblOldVAT
    End With
End Sub

Public Sub LoadNote()
    Dim recNoteCode As New ADODB.Recordset
    
    strQuery = "SELECT note_code FROM radio_note ORDER BY note_code"
    recNoteCode.Open strQuery, ConnERP, adOpenStatic, adLockReadOnly
    
    cboNoteCode.Clear
    With recNoteCode
        Do While .EOF = False
            cboNoteCode.AddItem Trim(.Fields("note_code").Value)
            .MoveNext
        Loop
    End With
    
    If recNoteCode.State = adStateOpen Then
        recNoteCode.Close
    End If
    
    Set recNoteCode = Nothing
End Sub

Public Sub cboBrand_Click()
    If cboBrand.ListIndex <> -1 Then
        Call LoadJobNumber
    End If
End Sub

Private Sub cboJobNumber_Click()
    Dim recShowDetail As New ADODB.Recordset
    
    If cboJobNumber.ListIndex <> -1 Then
        strQuery = "SELECT theme_flag, Job_Id FROM montly_radio_quotation WHERE job_number ='" & cboJobNumber.Text & "' "
        strQuery = strQuery & " AND Month_Schedule = " & Val(Left(Me.cboMonth.Text, 2)) & " AND "
        strQuery = strQuery & " Year_Schedule =  " & Me.cboYear.Text
        
        recShowDetail.CursorLocation = adUseClient
        recShowDetail.Open strQuery, ConnERP, adOpenStatic, adLockReadOnly
        
        With recShowDetail
            If Not .EOF Then
                fraBooked.Visible = IIf(.Fields("Theme_Flag").Value = 0, True, False)
                fraRemark.Visible = IIf(.Fields("Theme_Flag").Value = 0, False, True)
                Me.txtJobID.Text = .Fields("Job_Id").Value
            End If
        End With
        
        recShowDetail.Close
        Set recShowDetail = Nothing
        
        Call ShowStation
        
        'dw - Untuk yang spesial job (.130.) tidak bisa 'Change Materi'
        If Mid(cboJobNumber.Text, 6, 3) = "130" Then
            mnuEditMateri.Enabled = False
        Else
            mnuEditMateri.Enabled = True
        End If
     End If
End Sub

Private Sub CboMonth_Click()
    If cboMonth.ListIndex > -1 Then
        Call LoadJobNumber
    End If
End Sub

Private Sub cboStation_Click()
    If blnLoadCancelFlag = True Then
        If cboStation.ListIndex <> -1 Then
        
            Call GetBrandInfo
            
            Call ShowDataHeader
            
            'cmdPrint.Enabled = True
            
            Call LoadBookGrid(msgBook)
    
            'cmdEditNote.Enabled = True
        Else
            'cmdEditNote.Enabled = False
            
            Call ClearForm
        End If
    End If
End Sub

Private Sub cboYear_Click()
    If cboYear.ListIndex <> -1 Then
        Call LoadJobNumber
    End If
End Sub

Private Sub dbCancel()
'*****************************************
'Submodul Name      : CheckForClickAll
'Procedure Function : Membatalkan
'Used Object        : -
'Programmer Name    : Tedi
'Date               : 19-11-2015
'Last Update/By     : Tedi
'Date Update        :
'Log Update/By      :
'Previous Name      : cmdEditNotcmdCancel
'***************************************************************'*****************************************

    cboStation.Enabled = True
    
    Call ButtonLock(False)
    
    Call cboStation_Click
    Call EnableObject(False)
    
End Sub



Private Sub dbEdit()
'*****************************************
'Submodul Name      : dbEdit
'Procedure Function : Edit
'Programmer Name    : Tedi
'Date               : 19-11-2015
'Last Update/By     : Tedi
'Date Update        :
'Log Update/By      :
'Previous Name      : cmdEditNote_Click
'***************************************************************'*****************************************
    If txtDescription = "" Then
        MsgBox "PO is Empty", vbCritical, strApplication_Name
        Exit Sub
    End If
    
    If User_Valid("Implementor") Then
        Call ButtonLock(True)
        
        cboStation.Enabled = False
        
        If Mid(cboJobNumber, 6, 3) = "130" Then
            txtDescription.Locked = False
        Else
            txtDescription.Locked = True
        End If
        EnableObject True
    Else
        MsgBox strMsgAccessDenied, vbCritical, strTitleExclamation
        Exit Sub
    End If
End Sub

Private Sub cmdNote_Click()
    If Me.cboStation.ListIndex <> -1 Then
        Set Frm_Radio_Note.frmWhatForm = Frm_PO_Media_Radio
        
        Frm_Radio_Note.strCodeNote = Trim(cboNoteCode.Text)
        Frm_Radio_Note.show vbModal
    Else
        MsgBox "Select Station First", vbExclamation, strTitleMissingInfo
    End If
End Sub

Private Sub dbPrint()
'*****************************************
'Submodul Name      : dbPrint
'Procedure Function : Proses Cetak PO
'Used Object        :
'Programmer Name    : Tedi/Kreatif
'Date               : 10-08-2015
'Last Update/By     : Tedi
'Date Update        :
'Log Update/By      :
'Previous Name      : cmdPrint_Click
'***************************************************************'*****************************************
    MsgBox "Frm_Radio_Media_PrintPOCO tidak ada di ORI"
    Exit Sub
    If txtDescription = "" Then
        MsgBox "PO is Empty", vbCritical, strApplication_Name
        Exit Sub
    End If
    
'    Frm_Radio_Media_PrintPOCO.blnPOPrint = True
'    Frm_Radio_Media_PrintPOCO.blnCOPrint = False
'
'    Set Frm_Radio_Media_PrintPOCO.cboWhatBrand = cboBrand
'    Frm_Radio_Media_PrintPOCO.strNumberField = "Po_Media_radio.Po_Number "
'    Frm_Radio_Media_PrintPOCO.Caption = "Radio Media Purchase Order print"
'    Frm_Radio_Media_PrintPOCO.show vbModal
    
End Sub

Private Sub dbSave()
'*****************************************
'Submodul Name      : dbSave
'Procedure Function : Proses Simpan Note
'Used Object        :
'Programmer Name    : Tedi
'Date               : 10-08-2015
'Last Update/By     : Tedi
'Date Update        :
'Log Update/By      :
'Previous Name      : cmdSaveNote_Click
'***************************************************************'*****************************************
    
    'If cboNoteCode.ListIndex <> -1 Then
    If cboNoteCode.Text <> "" Then
        Call SaveNote
        
        Call ButtonLock(False)
        
        Call cboStation_Click
        EnableObject False
    Else
        MsgBox "Please select a Note Code", vbCritical, strTitleMissingInfo
    End If
End Sub

Private Sub mnuEditMateri_Click()
    If Me.cboStation.ListIndex <> -1 Then
        Frm_Radio_Change_Material.show vbModal
    Else
        MsgBox "Select Station First", vbExclamation, strTitleMissingInfo
    End If
End Sub

Private Sub msgBook_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 2 Then
        If txtDescription.Visible = False And Me.msgBook.Rows > 1 Then
            If Not User_Valid("Implementor") Then
                Exit Sub
            End If
        End If
    End If
End Sub

Private Sub LoadJobNumber()
    Dim recJobNo As New ADODB.Recordset
    
    Call ClearForm
    
'    cmdPrint.Enabled = False
    cboJobNumber.Clear
    cboStation.Clear
    Me.txtJobID.Text = ""
    
    strQuery = "SELECT DISTINCT Job_Number FROM Po_Media_Radio WHERE left(job_number,4)='" & Left(cboBrand.Text, 4) & "' "
    strQuery = strQuery & " AND year = " & Val(cboYear.Text) & " AND month =" & Val(cboMonth.Text)
    strQuery = strQuery & " AND Cancel_Flag<>1"
    recJobNo.Open strQuery, ConnERP, adOpenStatic, adLockReadOnly
    
    With recJobNo
        Do While .EOF = False
            cboJobNumber.AddItem .Fields("Job_Number").Value
            .MoveNext
        Loop
    End With
    
    recJobNo.Close
    Set recJobNo = Nothing
End Sub

Public Sub LoadBookGrid(msgBook As MSFlexGrid)
    Dim recMAterial As New ADODB.Recordset
    Dim recLoad As New ADODB.Recordset
    Dim recSpot As New ADODB.Recordset
    Dim intJumDate As Integer
    Dim intPosRow As Integer
    Dim lngTotalGros As Long
    Dim lngTotalNet As Long
    
    Call InitialGrid(Val(cboMonth.Text), Val(cboYear.Text))
    
    Set recLoad = Nothing
        
    strQuery = "SELECT * FROM PO_Media_Radio_Date_Insertion WHERE po_number ='" & txtPONumber.Text & "' "
    strQuery = strQuery & " AND station_code ='" & strGetStationCode(cboStation.Text, " ") & "'"
    
    recLoad.CursorLocation = adUseClient
    recLoad.Open strQuery, ConnERP, adOpenStatic, adLockReadOnly
    
    strQuery = "SELECT DISTINCT material_id,Material_Name,duration FROM PO_Media_Radio_Date_Insertion "
    strQuery = strQuery & " WHERE po_number ='" & Trim(txtPONumber.Text) & "'"
    
    recMAterial.CursorLocation = adUseClient
    recMAterial.Open strQuery, ConnERP, adOpenStatic, adLockReadOnly
    
    recSpot.CursorLocation = adUseClient
    recSpot.Open "SELECT * FROM radio_Spot_catalog", ConnERP, adOpenStatic, adLockReadOnly
            
    lngTotalGros = 0
    lngTotalNet = 0
    
    msgBook.Rows = msgBook.Rows
    intPosRow = msgBook.Rows
    
    With recMAterial
        Do While .EOF = False
            recSpot.Filter = ""
            
            Do While recSpot.EOF = False
                recLoad.Filter = ""
                
                If recLoad.RecordCount > 0 Then
                    recLoad.MoveFirst
                    intJumDate = 0
                    recLoad.Filter = ""
                    recLoad.Filter = "Spot_type ='" & Trim(recSpot.Fields("Spot_Type").Value) & "' AND Material_ID='" & Trim(.Fields("Material_Id").Value) & "'" & " AND station_code='" & strGetStationCode(cboStation.Text, " ") & "'"
                    
                    If recLoad.RecordCount > 0 Then
                        msgBook.Rows = msgBook.Rows + 1
                        intPosRow = msgBook.Rows - 1
                    End If
                    
                    Do While recLoad.EOF = False
                        If Trim(recLoad.Fields("Spot_Type").Value) = Trim(recSpot.Fields("Spot_Type").Value) Then
                            If .Fields("material_id").Value = recLoad.Fields("Material_id").Value Then
                                intJumDate = intJumDate + 1
                                msgBook.CellForeColor = vbBlack
                                msgBook.TextMatrix(intPosRow, 0) = intJumDate
                                msgBook.TextMatrix(intPosRow, 1) = .Fields("material_id").Value & " " & .Fields("Material_Name").Value & "/" & .Fields("Duration").Value
                                msgBook.TextMatrix(intPosRow, 2) = IIf(msgBook.TextMatrix(intPosRow, 2) = "", recLoad.Fields("Spot").Value, IIf(msgBook.TextMatrix(intPosRow, 2) >= recLoad.Fields("Spot").Value, msgBook.TextMatrix(intPosRow, 2), recLoad.Fields("Spot").Value))
                                msgBook.TextMatrix(intPosRow, 3) = recLoad.Fields("Spot_Type").Value & " (" & recSpot.Fields("Spot_name").Value & ")"
                                msgBook.TextMatrix(intPosRow, 4) = IIf(IsNull(recLoad.Fields("Nett_Rate").Value) = True, 0, recLoad.Fields("Nett_Rate").Value)
                                msgBook.TextMatrix(intPosRow, 5) = IIf(IsNull(recLoad.Fields("Gross_Rate").Value) = True, 0, recLoad.Fields("Gross_Rate").Value)
                                
                                msgBook.col = (recLoad.Fields("Date").Value) + 5
                                msgBook.Row = intPosRow
                                
                                If recLoad.Fields("Cancel_Flag").Value = 1 Then
                                    msgBook.CellBackColor = vbRed
                                    msgBook.TextMatrix(intPosRow, (recLoad.Fields("Date").Value) + 5) = IIf(IsNull(recLoad.Fields("date").Value) = True, "", "C")
                                ElseIf recLoad.Fields("Cancel_Flag").Value = 2 Then
                                    msgBook.CellBackColor = vbGreen
                                    msgBook.TextMatrix(intPosRow, (recLoad.Fields("Date").Value) + 5) = IIf(IsNull(recLoad.Fields("date").Value) = True, "", "S=" & recLoad.Fields("Spot").Value)
                                ElseIf recLoad.Fields("Cancel_Flag").Value = 0 Then
                                    msgBook.CellBackColor = vbGreen
                                    msgBook.TextMatrix(intPosRow, (recLoad.Fields("Date").Value) + 5) = IIf(IsNull(recLoad.Fields("date").Value) = True, "", "B=" & recLoad.Fields("Spot").Value)
                                End If
                            End If
                        End If
                        recLoad.MoveNext
                    Loop
                End If
                recSpot.MoveNext
            Loop
            
            .MoveNext
        Loop
    End With
    
    If recMAterial.State = adStateOpen Then
        recMAterial.Close
    End If
    
    Set recMAterial = Nothing
    
    If recLoad.State = adStateOpen Then
        recLoad.Close
    End If
    
    Set recLoad = Nothing
    
    If recSpot.State = adStateOpen Then
        recSpot.Close
    End If
    
    Set recSpot = Nothing
End Sub

Public Sub InitialGrid(intMonthNow As Integer, intYearNow As Integer)
    Dim intPos As Integer
    Dim intJumMonth As Integer
    Dim strMonth As String
    
    If intMonthNow <= 0 Then Exit Sub
    
    strMonth = intMonthNow & "/01/" & intYearNow
    intJumMonth = Day(DateAdd("m", 1, strMonth) - 1)
    
    msgBook.Clear
    msgBook.Rows = msgBook.Rows + 1
    
    For intPos = 0 To msgBook.Rows - 1
        msgBook.RowHeight(intPos) = 318
    Next intPos
    
    With msgBook
        .ForeColor = vbBlue
        .ColWidth(0) = 500 'Total Days
        .ColWidth(1) = 3000 'Materi
        .ColWidth(2) = 0  'Spots
        .ColWidth(3) = 1500  'Code Rate
        .ColWidth(4) = 1000  'Nett Rate
        .ColWidth(5) = 1000  'Gross Rate
    
        .TextMatrix(0, 0) = "Days"
        .TextMatrix(0, 1) = "Materi/Duration"
        .TextMatrix(0, 2) = "Spots"
        .TextMatrix(0, 3) = "Code Rate"
        .TextMatrix(0, 4) = "Nett Rate"
        .TextMatrix(0, 5) = "Gross Rate"
        .FixedCols = 0
    End With
    
    msgBook.cols = intJumMonth + 6
    msgBook.FixedRows = 1
    
    msgBook.Rows = 1
    For intPos = 6 To intJumMonth + 5
        msgBook.ColWidth(intPos) = 450
        msgBook.TextMatrix(0, intPos) = intPos - 5
    Next intPos
End Sub

Private Function strGetMonth(ByVal strWhatString As String, strWhatSign As String) As String
    Dim intPos As Integer
    Dim intJum As Integer
    
    intJum = Len(strWhatString)
    For intPos = 1 To intJum
        If Mid(strWhatString, intPos, 1) = strWhatSign Then
            strGetMonth = Left(strWhatString, intPos - 1)
            
            If Len(strGetMonth) < 2 Then
                strGetMonth = "0" & strGetMonth
            End If
            
            Exit Function
        End If
    Next intPos
End Function

Private Sub ShowStation()
    Call ClearForm
    
    cboStation.Clear
    If recHeader.State = adStateOpen Then
        recHeader.Close
    End If
    
    strQuery = "SELECT po_media_radio.*, montly_radio_quotation.ib_id FROM po_media_radio "
    strQuery = strQuery & " INNER JOIN montly_radio_quotation ON montly_radio_quotation.job_number = po_media_radio.job_number "
    strQuery = strQuery & " AND montly_radio_quotation.job_id = po_media_radio.job_id "
    strQuery = strQuery & " WHERE po_media_radio.job_number='" & Trim(cboJobNumber.Text) & "' AND po_media_radio.Job_Id='" & txtJobID.Text & "'"
    recHeader.CursorLocation = adUseClient
    recHeader.Open strQuery, ConnERP, adOpenDynamic, adLockReadOnly
    
    With recHeader
        .Filter = ""
        Do While .EOF = False
            cboStation.AddItem .Fields("Station_code").Value & " " & .Fields("Station_Name").Value
            .MoveNext
        Loop
    End With

    'cmdEditNote.Enabled = False
End Sub

Private Sub GetBrandInfo()
    Dim recCek As New ADODB.Recordset
    Dim recULI As New ADODB.Recordset

    strQuery = "SELECT * FROM Montly_Radio_Quotation WHERE Job_Id='" & txtJobID.Text & "' AND Job_Number='" & cboJobNumber.Text & "'"
    recCek.Open strQuery, ConnERP, adOpenStatic, adLockReadOnly

    With recCek
        Do While .EOF = False
            'GEt Apakah Client ULI
            strQuery = "SELECT a.* FROM Client a, Brand b WHERE a.client_code=b.client_code AND b.brand_code='" & Left(Me.cboBrand.Text, 4) & "'"
            recULI.Open strQuery, ConnERP, adOpenStatic, adLockReadOnly

            Brand_Info.ULI = IIf(Trim(recULI.Fields("special_client_flag").Value) = 1, True, False)

            recULI.Close
            Set recULI = Nothing

            'MSC
            Brand_Info.MSC = IIf(IsNull(.Fields("Percent_MSC").Value) = True, 0, .Fields("Percent_MSC").Value) / 100

            Select Case .Fields("MSC_On_Flag").Value
                Case Is = 0
                    Brand_Info.MSC_Nett_Flag = False
                Case Is = 1
                    Brand_Info.MSC_Nett_Flag = True
                Case Is = 2
                    Brand_Info.MSC_Nett_Flag = False
                Case Is = 3
                    Brand_Info.MSC_Nett_Flag = False
                Case Is = 4
                    Brand_Info.MSC_Nett_Flag = True
            End Select

            'MSC (Bonus)
            Brand_Info.Media_Agency_Bonus = IIf(IsNull(.Fields("Percent_MSC").Value) = True, 0, .Fields("Percent_MSC").Value) / 100

            Select Case .Fields("MSC_On_Flag").Value
                Case Is = 0
                    Brand_Info.Media_Agency_Bonus_Nett_Flag = False
                Case Is = 1
                    Brand_Info.Media_Agency_Bonus_Nett_Flag = True
                Case Is = 2
                    Brand_Info.Media_Agency_Bonus_Nett_Flag = False
                Case Is = 3
                    Brand_Info.Media_Agency_Bonus_Nett_Flag = False
                Case Is = 4
                    Brand_Info.Media_Agency_Bonus_Nett_Flag = True
            End Select

            'Brand Agency
            Brand_Info.Club_Agency_Flag = IIf(.Fields("MSC_On_Flag").Value = 1, True, False)
            Brand_Info.Club_Agency_SC = IIf(IsNull(.Fields("Percent_MSC").Value) = True, 0, .Fields("Percent_MSC").Value) / 100

            'VAT
            Brand_Info.Vat = IIf(IsNull(.Fields("VAT_Percent").Value) = True, 0, .Fields("VAT_Percent").Value)
            .MoveNext
        Loop
    End With

    recCek.Close
    Set recCek = Nothing
End Sub

Private Function strGetStationCode(strWhatString As String, strWhatCode As String) As String
    Dim intPos As Integer
    Dim intJum As Integer
    
    intJum = Len(strWhatString)
    For intPos = 1 To intJum
        If Mid(strWhatString, intPos, 1) = strWhatCode Then
            strGetStationCode = Left(strWhatString, intPos - 1)
            Exit Function
        End If
    Next intPos
End Function

Private Sub ClearForm()
    txtPONumber.Text = ""
    txtPODate.Text = ""
    txtAddress.Text = ""
    txtBrandVariant.Text = ""
    txtDescription.Text = ""
    
    txtNett.Text = 0
    txtGross.Text = 0
    cboNoteCode.ListIndex = -1
    Call InitialGrid(cboMonth.ListIndex + 1, cboYear.ListIndex + 2000)
End Sub

Private Sub ShowDataHeader()
    Call ClearForm
    
    With recHeader
        .Requery
        .Filter = ""
        .Filter = "station_code ='" & strGetStationCode(cboStation.Text, " ") & "'"
        
        If .RecordCount > 0 Then
            strIBID = Trim(.Fields("IB_ID").Value)
            txtPONumber.Text = Trim(.Fields("PO_Number").Value)
           
            If IsNull(.Fields("Brand_Variant_Code").Value) = False And IsNull(.Fields("Brand_Variant_name").Value) = False Then
                txtBrandVariant.Text = .Fields("Brand_variant_code").Value & " " & .Fields("Brand_variant_name").Value
            Else
                txtBrandVariant.Text = ""
            End If
            
            txtGross.Text = Format(.Fields("Total_Gross").Value, "#,##0")
            txtNett.Text = Format(.Fields("Total_Nett").Value, "#,##0")
                        
            txtDescription.Text = IIf(IsNull(.Fields("Remark_Special").Value) = True, "", .Fields("remark_special").Value)
            cboNoteCode.Text = Trim(IIf(IsNull(.Fields("Note_Code").Value) = True, "", Trim(.Fields("Note_Code").Value)))
            Me.txtPODate.Text = Format(.Fields("Entered_Date").Value, "dd/mmm/yyyy")
            
            txtAddress.Text = IIf(IsNull(.Fields("address1").Value), "", .Fields("address1").Value) & vbCrLf & IIf(IsNull(.Fields("address2").Value), "", .Fields("address2").Value) & vbCrLf & IIf(IsNull(.Fields("address3").Value), "", .Fields("address3").Value)
        End If
    End With
End Sub

Private Sub mnuCancelTemplate_Click()
    If cboStation.ListIndex <> -1 Then
        If fraRemark.Visible = False Then 'Remark (Job Type : Sponsorship)
            If msgBook.Rows > 1 Then
                Frm_Radio_Booking_Date_Order.blnJobType = True
                Frm_Radio_Booking_Date_Order.blnParsialCancel = False
                Frm_Radio_Booking_Date_Order.intRowPos = msgBook.Row
                Frm_Radio_Booking_Date_Order.show vbModal
            End If
        Else
            Frm_Radio_Booking_Date_Order.blnJobType = False
            Frm_Radio_Booking_Date_Order.blnParsialCancel = False
            Frm_Radio_Booking_Date_Order.intRowPos = msgBook.Row
            Frm_Radio_Booking_Date_Order.show vbModal
        End If
    Else
        MsgBox "Select Station First", vbExclamation, strTitleMissingInfo
    End If
End Sub

Private Sub mnuChangeMateriDescription_Click()
    If cboStation.ListIndex <> -1 Then
        Frm_Radio_Change_Material_Desc.show 1
    Else
        MsgBox "Select Station First", vbExclamation, strTitleMissingInfo
    End If
End Sub

Private Sub txtDescription_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 2 Then
        mnuChangeMateriDescription.Enabled = False
        mnuEditMateri.Enabled = False

        If Not User_Valid("Implementor") Then
            Exit Sub
        End If
    End If
End Sub

Private Sub ButtonLock(blnEnable As Boolean)
    'cmdEditNote.Enabled = Not blnEnable
    'cmdSaveNote.Enabled = blnEnable
    'cmdCancelNote.Enabled = blnEnable
    'cmdClose.Enabled = Not blnEnable
    'cmdPrint.Enabled = Not blnEnable
    fraHeader1.Enabled = Not blnEnable
    cboNoteCode.Enabled = blnEnable
    cmdNote.Enabled = blnEnable
    msgBook.Enabled = Not blnEnable
    txtDescription.Enabled = blnEnable
End Sub

Private Sub SaveNote()
    On Error GoTo chk
    
    ConnERP.BeginTrans
    
    If MsgBox("Save To All Stations ?", vbQuestion + vbYesNo, strTitleConfirm) = vbYes Then
        strQuery = "UPDATE Po_Media_Radio"
        strQuery = strQuery & " SET Note_Code ='" & Trim(cboNoteCode.Text) & "',"
        strQuery = strQuery & " Remark_Special ='" & Clear_String(txtDescription.Text) & "'"
        strQuery = strQuery & " WHERE "
        strQuery = strQuery & " Job_Number ='" & Trim(cboJobNumber.Text) & "' AND Job_Id = '" & Me.txtJobID.Text & "'"
    Else
        strQuery = "UPDATE Po_Media_Radio "
        strQuery = strQuery & " SET Note_Code ='" & Trim(cboNoteCode.Text) & "',"
        strQuery = strQuery & " Remark_Special ='" & Clear_String(txtDescription.Text) & "'"
        strQuery = strQuery & " WHERE "
        strQuery = strQuery & " PO_Number ='" & Trim(txtPONumber.Text) & "'"
        strQuery = strQuery & " AND Station_Code ='" & Trim(strGetStationCode(cboStation.Text, " ")) & "'"
    End If
    
    ConnERP.Execute strQuery
    
    ConnERP.CommitTrans
    
    cboStation.Enabled = True
    Exit Sub
    
chk:
    ConnERP.RollbackTrans
    MsgBox Err.Number & " " & Err.Description, vbCritical, strTitleExclamation
End Sub

Private Sub txtGross_KeyPress(KeyAscii As Integer)
    If (Chr(KeyAscii) < "0" Or Chr(KeyAscii) > "9") And KeyAscii <> 8 And KeyAscii <> 13 And KeyAscii <> 9 And KeyAscii <> 46 Then
        KeyAscii = 0
        Beep
        Exit Sub
    End If
End Sub

Private Sub txtGross_LostFocus()
    txtGross.Text = Format(txtGross.Text, "#,##0")
End Sub

Private Sub txtNett_KeyPress(KeyAscii As Integer)
    If (Chr(KeyAscii) < "0" Or Chr(KeyAscii) > "9") And KeyAscii <> 8 And KeyAscii <> 13 And KeyAscii <> 9 And KeyAscii <> 46 Then
        KeyAscii = 0
        Beep
        Exit Sub
    End If
End Sub

Private Sub txtNett_LostFocus()
    txtNett.Text = Format(txtNett.Text, "#,##0")
End Sub

Private Sub cboJobNumber_KeyPress(KeyAscii As Integer)
    KeyAscii = 0
End Sub

Private Sub cboNoteCode_KeyPress(KeyAscii As Integer)
    KeyAscii = 0
End Sub

Private Sub cboBrand_KeyPress(KeyAscii As Integer)
    KeyAscii = 0
End Sub

Private Sub cboMonth_KeyPress(KeyAscii As Integer)
    KeyAscii = 0
End Sub

Private Sub cboYear_KeyPress(KeyAscii As Integer)
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
    
    With picButton(enButtonType.bieedit) 'EDIT. 5
        .Enabled = paIsNormalMode
        .Visible = paIsNormalMode
    End With
    With picButton(enButtonType.bieSave)  'SAVE.
        .Enabled = Not paIsNormalMode
        .Visible = Not paIsNormalMode
    End With
    With picButton(enButtonType.bieprint)   'FIND.
        .Enabled = paIsNormalMode
        .Visible = paIsNormalMode
    End With
    With picButton(enButtonType.bieClose)      'Quit.
        .Enabled = paIsNormalMode
        .Visible = paIsNormalMode
    End With
    With picButton(enButtonType.bieCancel) 'CANCEL.
        .Enabled = Not paIsNormalMode
        .Visible = Not paIsNormalMode
    End With
    cboBrand.Enabled = paIsNormalMode
    blnEditOrAdd = Not paIsNormalMode
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
    pnl_Main.Height = Me.ScaleHeight - picToolbar.Height - picStatusBar.Height
    fra_Deliverable.Height = pnl_Main.Height - (fra_Deliverable.Top + 100)
    SSTab3.Height = fra_Deliverable.Height - (SSTab3.Top) - 150
    txtOther_Recomedation.Height = SSTab3.Height - (txtOther_Recomedation.Top) - 150
    txtAggreed_Channel_shortlist.Height = txtOther_Recomedation.Height
    fra_DeliverableChannel.Height = pnl_Main.Height - (fra_DeliverableChannel.Top + 100)
    fraFilter.Width = pnl_Main.Width - (fraFilter.Left * 2)
    lineFilter.X1 = fraFilter.Width / 2
    lineFilter.X2 = lineFilter.X1
    Fra_Approve.Left = lineFilter.X2 + Label7.Left
    txtYear.Width = lineFilter.X2 - txtYear.Left - 50
    txtClient_Brief_Id.Width = txtYear.Width
    txtExtention.Width = txtYear.Width
    txtStatus.Width = txtYear.Width
    'left part
    lbl_dateofPreviousIssue.Left = lineFilter.X1 + Label7.Left
    dtpDate_Previouse.Left = lbl_dateofPreviousIssue.Left + lbl_dateofPreviousIssue.Width + 50
    dtpDate_Issue.Left = dtpDate_Previouse.Left
    lbl_DateIssue.Left = lbl_dateofPreviousIssue.Left
    lblCountry.Left = lbl_dateofPreviousIssue.Left
    cboCountry.Left = dtpDate_Previouse.Left
    Fra_Approve.Left = dtpDate_Previouse.Left
    fra_DeliverableChannel.Width = pnl_Main.Width - fra_DeliverableChannel.Left - fraFilter.Left
    lstRec_Channel_Selection.Width = fra_DeliverableChannel.Width - (lstRec_Channel_Selection.Left * 2)
    lstRec_Channel_Selection.Height = fra_DeliverableChannel.Height - (lstRec_Channel_Selection.Top) - 200
    chk_All.Top = lstRec_Channel_Selection.Height + lstRec_Channel_Selection.Top + 50
    lbl_CheckAll.Top = chk_All.Top
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
 
        Case enButtonType.bieedit  '5 'EDIT.
            dbEdit
        Case enButtonType.bieprint    '8 'FIND.
            Call dbPrint
        Case enButtonType.bieClose   '23 'EXIT.
            Unload Me
        Case enButtonType.bieSave  '10 SAVE.
            Call dbSave
        Case enButtonType.bieCancel '11 CANCEL.
            Call dbCancel
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

