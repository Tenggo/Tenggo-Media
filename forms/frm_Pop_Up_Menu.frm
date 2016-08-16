VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "mscomctl.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form frm_Pop_Up_Menu 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Task"
   ClientHeight    =   7050
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9495
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7050
   ScaleWidth      =   9495
   StartUpPosition =   2  'CenterScreen
   Begin Crystal.CrystalReport crpt 
      Left            =   1155
      Top             =   6555
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      PrintFileLinesPerPage=   60
   End
   Begin VB.Frame Frame3 
      Height          =   6015
      Left            =   45
      TabIndex        =   2
      Top             =   0
      Width           =   9384
      Begin VB.TextBox Txt_Notes 
         Height          =   852
         Left            =   810
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   6
         Top             =   4950
         Width           =   8448
      End
      Begin MSComctlLib.ListView List_number 
         Height          =   3990
         Left            =   135
         TabIndex        =   4
         Top             =   495
         Width           =   9150
         _ExtentX        =   16140
         _ExtentY        =   7038
         View            =   3
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   11
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Job ID"
            Object.Width           =   1764
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "From"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Date"
            Object.Width           =   2117
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "Brand"
            Object.Width           =   3176
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Text            =   "IB_ID/MQ_ID/MP_No"
            Object.Width           =   2647
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   5
            Text            =   "Job Number"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   6
            Text            =   "Task"
            Object.Width           =   8820
         EndProperty
         BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   7
            Text            =   "PO"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   8
            Text            =   "Notes"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   9
            Text            =   "Status Message"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(11) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   10
            Text            =   "Activity_id"
            Object.Width           =   2540
         EndProperty
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Notes :"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   165
         TabIndex        =   7
         Top             =   4950
         Width           =   570
      End
      Begin VB.Label Label2 
         Caption         =   "Double Click to Select... !"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   240
         Left            =   150
         TabIndex        =   5
         Top             =   4575
         Width           =   3300
      End
      Begin VB.Label Label1 
         Caption         =   "Job List To Do :"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   132
         TabIndex        =   3
         Top             =   204
         Width           =   1344
      End
   End
   Begin VB.Frame Frame2 
      Height          =   744
      Left            =   8052
      TabIndex        =   0
      Top             =   6165
      Width           =   1365
      Begin VB.CommandButton cmd_Close 
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
         Height          =   444
         Left            =   90
         TabIndex        =   1
         Top             =   204
         Width           =   1185
      End
   End
   Begin VB.Menu Mnu_Open_Job 
      Caption         =   "Open Job"
      Visible         =   0   'False
      Begin VB.Menu Mnu_Delete 
         Caption         =   "Delete"
      End
   End
End
Attribute VB_Name = "frm_Pop_Up_Menu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
''************************************************************
' Unit/Module Name  : POP UP MENU
' Function          : Menu otomatis untuk Approve
'                   :
' Public Variables  : Brief_Id, job_no, po
' Used Tables       : activity_Job_catalog,Current_user_job
' Date              : 04.10.05
' Created By        : Diyah
' Last Update/By    :
' is_done 0 -> blm dibuka
' is_done 1 -> sdh dibuka
' is done 2 -> sdh dikerjakan
'is done 3 -> di delete
'************************************************************
Option Explicit

Dim Brief_Id As String
Dim Job_No As String
Dim PO_Number As String
Dim intActivity As Integer
Dim Job_Id As String
Dim Brand_Name As String
Dim MQ_Number As String
Dim strSql As String
Dim Rs_Pop As New ADODB.Recordset
Dim item_job As ListItem

Private Sub Cmd_Close_Click()
    Unload Me
End Sub

Private Sub Form_Load()
'************************************************************
' Unit/Module Name  : Form_Load
' Function          : Menampilkan data
'                   :
' Public Variables  :
' Used Tables       : activity_Job_catalog,Current_user_job
' Date              : 29.07.2004
' Created By        : Diyah
' Last Update/By    :
'************************************************************
    Dim rsbrand As New ADODB.Recordset
    Dim strSqlBrand As String
    Dim intIDX As Integer
    
    mdi_Main.Tmr_Cek_Task.Enabled = False
    
    Txt_Notes.Text = ""
    'semua task dari user diambil,baik yang sdh dilakukan ato belum, kecuali task yang sdh dihapus(Is_Done=3)
    strSql = "SELECT A.*,B.Is_Done As Is_User_Done,C.Description As Job_Description FROM Current_Job A INNER JOIN Current_User_Job B ON A.Id_Job=B.Id_Job"
    strSql = strSql & " INNER JOIN Activity_Job_Catalog C ON A.Activity_ID=C.Activity_ID"
    strSql = strSql & " WHERE B.User_Name='" & strLogin_User & "' AND B.Is_Done <>3"
    strSql = strSql & " ORDER By A.Date DESC"
    
    If Rs_Pop.State = adStateOpen Then
        Rs_Pop.Close
        Set Rs_Pop = Nothing
        List_number.ListItems.Clear
    End If
    
    Rs_Pop.Open strSql, ConnERP, adOpenStatic, adLockReadOnly, adCmdText

    While Not Rs_Pop.EOF And Not Rs_Pop.BOF
    
        'Mengisi List View
        Set item_job = List_number.ListItems.Add(, , Rs_Pop("Id_Job").Value) 'Job Id
        item_job.ListSubItems.Add , , IIf(IsNull(Rs_Pop("Message_From").Value), "", Rs_Pop("Message_From").Value)
        item_job.ListSubItems.Add , , IIf(IsNull(Rs_Pop("Date").Value), "", Format(Rs_Pop("Date").Value, "dd/mm/yy hh:mm")) 'Date 'Date
            
        '******* Filter Brand berdasarkan Brief_Id
        strSqlBrand = ""
        strSqlBrand = "SELECT Brand_Code, Brand_Name From Brand "
        strSqlBrand = strSqlBrand & "WHERE Brand_Code = '" & Left(Trim(Rs_Pop("Brief_Id").Value), 4) & "'"
        
        If rsbrand.State = adStateOpen Then
           rsbrand.Close
           Set rsbrand = Nothing
        End If
        rsbrand.Open strSqlBrand, ConnERP, adOpenDynamic, adLockReadOnly
       
        item_job.ListSubItems.Add , , rsbrand("Brand_Name").Value 'Brand
        item_job.ListSubItems.Add , , Rs_Pop("Brief_Id").Value 'Brief Id
        item_job.ListSubItems.Add , , IIf(IsNull(Rs_Pop("Job_Number").Value), "", Rs_Pop("Job_Number").Value) 'Job Number
        item_job.ListSubItems.Add , , Rs_Pop("Job_Description").Value  'Task
        item_job.ListSubItems.Add , , IIf(IsNull(Rs_Pop("PO_Number").Value), "", Rs_Pop("PO_Number").Value) 'PO Number
        item_job.ListSubItems.Add , , IIf(IsNull(Rs_Pop("Description").Value), "", Rs_Pop("Description").Value) 'CE Number
        item_job.ListSubItems.Add , , IIf(IsNull(Rs_Pop("Message").Value), "", Rs_Pop("Message").Value) 'Note
        item_job.ListSubItems.Add , , Rs_Pop("Task_Type").Value 'Type task
        item_job.ListSubItems.Add , , Rs_Pop("Activity_Id").Value 'activity_id
        
        'jika belum dibuka(is_done =0) bold-kan tulisan
        If Rs_Pop.Fields("is_user_done").Value = 0 Then
            For intIDX = 1 To 11
                item_job.ListSubItems.Item(intIDX).ForeColor = vbRed
            Next intIDX
        
'            item_job.ListSubItems.item(1).Bold = True
'            item_job.ListSubItems.item(2).Bold = True
'            item_job.ListSubItems.item(3).Bold = True
'            item_job.ListSubItems.item(4).Bold = True
'            item_job.ListSubItems.item(5).Bold = True
'            item_job.ListSubItems.item(6).Bold = True
'            item_job.ListSubItems.item(7).Bold = True
'            item_job.ListSubItems.item(8).Bold = True
'            item_job.ListSubItems.item(9).Bold = True
'            item_job.ListSubItems.item(10).Bold = True
'            item_job.ListSubItems.item(11).Bold = True
        End If
        Rs_Pop.MoveNext
    Wend
    
    List_number.ColumnHeaders(1).Width = 0 'Job Id
    List_number.ColumnHeaders(10).Width = 0 'Task Type
    List_number.ColumnHeaders(11).Width = 0 'activity_id
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If Rs_Pop.State = adStateOpen Then
        Rs_Pop.Close
        Set Rs_Pop = Nothing
    End If
       
    'Update task yg New_Messagenya 1 ke nol,menunjukan bahwa task sudah pernah ditampilkan
    strSql = "UPDATE Current_User_Job SET New_Message=0 WHERE User_Name='" & strLogin_User & "' AND New_Message=1"
    ConnERP.Execute strSql

    'Enable timer untuk cek task
    mdi_Main.Tmr_Cek_Task.Enabled = True
   
End Sub

Private Sub List_number_BeforeLabelEdit(Cancel As Integer)
    Cancel = 1
End Sub

Private Sub List_number_Click()
'************************************************************
' Unit/Module Name  : List_number_Click
' Function          :
'                   :
' Public Variables  :
' Used Tables       :
' Date              : 29.07.2004
' Created By        : Diyah
' Last Update/By    :
'************************************************************
    Dim Item As ListItem
    
    'Cek apakah List item mempunyai data atau tidak
    If Not List_number.ListItems.Count > 0 Then
        'Jika tidak ada data task tidak dilakukan
        Exit Sub
    End If
    
    Set Item = List_number.ListItems(List_number.SelectedItem.Index)
    
    Rs_Pop.MoveFirst
    Rs_Pop.Find "Id_Job=" & Item.Text
    
    If Not Rs_Pop.EOF Then
        Txt_Notes.Text = IIf(IsNull(Rs_Pop.Fields("Message").Value), "", Rs_Pop.Fields("Message").Value)
    Else
        Txt_Notes.Text = ""
    End If
    
End Sub

Private Sub List_number_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    
    With List_number
        .SortKey = ColumnHeader.Index - 1
        .Sorted = True
        If .SortOrder = lvwAscending Then
            .SortOrder = lvwDescending
        Else
            .SortOrder = lvwAscending
        End If
    End With
End Sub

Private Sub List_number_DblClick()
'************************************************************
' Unit/Module Name  : List_number_DblClick
' Function          :
'                   :
' Public Variables  :
' Used Tables       :
' Date              : 27.07.2004
' Created By        : Diyah
' Last Update/By    :
'************************************************************
    Dim Item As ListItem
    Dim pesan As Integer
    Dim intIDX As Integer
    Dim is_MP_Number_In_List As Boolean
    On Error GoTo errHandler

    If Not List_number.ListItems.Count > 0 Then
        Exit Sub
    End If
    
    Set Item = List_number.ListItems(List_number.SelectedItem.Index)
    
    Brand_Name = Item.ListSubItems(3).Text
    Brief_Id = Item.ListSubItems(4).Text
    Job_No = Item.ListSubItems(5).Text
    PO_Number = Item.ListSubItems(7).Text
    intActivity = Item.ListSubItems(11).Text
    
    If Item.ListSubItems.Item(1).ForeColor = vbRed Then
        For intIDX = 1 To 11
            Item.ListSubItems.Item(intIDX).ForeColor = vbBlack
        Next intIDX
'        item.ListSubItems.item(1).Bold = False
'        item.ListSubItems.item(2).Bold = False
'        item.ListSubItems.item(3).Bold = False
'        item.ListSubItems.item(4).Bold = False
'        item.ListSubItems.item(5).Bold = False
'        item.ListSubItems.item(6).Bold = False
'        item.ListSubItems.item(7).Bold = False
'        item.ListSubItems.item(8).Bold = False
'        item.ListSubItems.item(9).Bold = False
'        item.ListSubItems.item(10).Bold = False
'        item.ListSubItems.item(11).Bold = False
        
        'update current_job dan current_user_Job diletakkan di dalam pengecekan bold apa tidak
        'karena jika dia sdh tidak bold maka pasti sdh diupdate data di current job dan currentuserjob
        'jadi pada double klik selanjutnya tidak prlu meng-update data
        
        'Cek apakah status message hanya pemberitahuan ataukah user perlu melakukan task tertentu
        'dari task_type jika 1 -> perintah melakukan sesuatu, 0 ->pemberitahuan
        If Item.ListSubItems(10).Text = 0 Then
            strSql = "UPDATE Current_User_Job SET is_done='2',update_date= getDate()"
            strSql = strSql & " WHERE Id_Job=" & Item.Text & " AND User_Name='" & strLogin_User & "'"
            ConnERP.Execute strSql
        Else
            strSql = "UPDATE Current_User_Job SET is_done='1',update_date= getDate()"
            strSql = strSql & " WHERE Id_Job=" & Item.Text & " AND User_Name='" & strLogin_User & "'"
            ConnERP.Execute strSql
        End If
            
        '---- Double Click saja maka Flag Done = 1 -----
        strSql = "UPDATE  Current_Job SET is_done='1',update_date= getDate() "
        strSql = strSql & " WHERE Activity_Id=" & intActivity & " AND Brief_Id='" & Item.ListSubItems(4).Text & "'"
        strSql = strSql & " AND Job_Number='" & Item.ListSubItems(5).Text & "' AND PO_Number='" & Item.ListSubItems(7).Text & "'"
        strSql = strSql & " AND Is_done=0"
        ConnERP.Execute strSql
    
    End If
        
    'ini untuk menampilkan report dari activity yang dimaksud
    Select Case intActivity
        Case 1, 18
            Show_IB_TV Brief_Id, Brand_Name
        Case 2, 19
            Show_IB_RD Brief_Id, Brand_Name
        Case 3, 20
            Show_IB_PR Brief_Id, Brand_Name
        Case 4, 21
            Show_IB_OT_CN Brief_Id, Brand_Name
        Case 5, 22
            Show_IB_OT_CN Brief_Id, Brand_Name
        Case 6, 12
            Show_Quot_TV Brief_Id, Brand_Name
        Case 7, 13
            Show_Quot_RD Brief_Id
        Case 8, 14
            Show_Quot_PR Brief_Id, Brand_Name
        Case 9, 15
            Show_Quot_OT_CN Brief_Id, Brand_Name
        Case 10, 16
            Show_Quot_OT_CN Brief_Id, Brand_Name
        Case 11
            'Do Nothing
        Case 17
            If Screen.Width < 15360 And Screen.Height < 11520 Then
                pesan = MsgBox("Recomended viewed in 1024 X 768 or Higher Screen Resolution." & vbCrLf & "Click OK to Continue, otherwise Click Cancel", vbOKCancel + vbInformation)
                If pesan = 2 Then
                    Exit Sub
                End If
            End If
            
            Me.MousePointer = vbHourglass
                is_MP_Number_In_List = False
                'Load Form
                Load frm_MPInsertion
                
                With frm_MPInsertion
                    'Cek MP Number ada di list?
                    For intIDX = 1 To .cboMPNumber.ListCount
                        If .cboMPNumber.List(intIDX - 1) = Brief_Id Then
                            is_MP_Number_In_List = True
                            Exit For
                        End If
                    Next
                    'Jika tidak ada, tambahkan MP Number ke dlm list
                    If Not is_MP_Number_In_List Then
                        .cboMPNumber.AddItem Brief_Id
                    End If
                    'pilih MP Number yang akan ditampilkan
                    .cboMPNumber.Text = Brief_Id
                    'Show Form
                    .show 1
                End With
            Me.MousePointer = vbDefault
            
   End Select

    Exit Sub

errHandler:
    MsgBox Err.Number & " : " & Err.Description, vbCritical
End Sub



Private Sub List_number_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 46 Then 'Delete
        mnu_Delete_Click
    End If
End Sub

Private Sub List_number_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If List_number.ListItems.Count Then
        If Button = 2 Then
            PopupMenu Mnu_Open_Job, , , , mnu_Delete
        End If
    End If
End Sub

Private Sub mnu_Delete_Click()
'************************************************************
' Unit/Module Name  : Mnu_Delete_Click
' Function          : Delete Job
'                   :
' Public Variables  :
' Used Tables       :
' Date              : 27.07.2004
' Created By        : Diyah
' Last Update/By    :
'************************************************************
    Dim Item As ListItem
    
    If Not List_number.ListItems.Count > 0 Then
        Exit Sub
    End If
    
    'Set item = Nothing
    Set Item = List_number.ListItems(List_number.SelectedItem.Index)
    
    '--- Is done = 3 untuk status Flag delete ----
    strSql = "UPDATE Current_User_Job "
    strSql = strSql & "SET is_done = '3' , update_date= getdate() "
    strSql = strSql & "WHERE Id_Job=" & Item.Text & " AND User_Name='" & strLogin_User & "'"
    ConnERP.Execute strSql
            
    Call Form_Load
End Sub

Private Sub Show_IB_TV(strIBID As String, Brand As String)
'************************************************************
' Unit/Module Name  : Show_IB_TV
' Function          : show ib tv
'                   :
' Public Variables  :
' Used Tables       :
' Date              : 11.10.05
' Created By        : Diyah
' Last Update/By    :
'************************************************************
    Dim Special_Brand As Boolean
    Dim Rs_Temp As New ADODB.Recordset
    Dim rs_IB_TV As New ADODB.Recordset
    Dim Rs_Montly_Budget As New ADODB.Recordset
    Dim Rs_TO As New ADODB.Recordset
    Dim Index_TO As Integer
    Dim str_week As String
    Dim RS_Materi As New ADODB.Recordset
    Dim Rs_Co As New ADODB.Recordset
    Dim Year_Ib As Integer
    Dim Month_1 As Integer
    Dim Month_2 As Integer
    Dim Month_3 As Integer
    Dim Week_Count_Month1 As Integer
    Dim Week_Count_Month2 As Integer
    Dim Week_Count_Month3 As Integer
    Dim Str_Week_Comm_1 As String
    Dim Str_Week_Comm_2 As String
    Dim Str_Week_Comm_3 As String
    Dim Week_Position As Integer
    Dim RsTempCampaign As New ADODB.Recordset
    Dim StrField As String
    Dim lbl_month1 As String
    Dim lbl_month2 As String
    Dim lbl_month3 As String
    Dim budget_month1 As Currency
    Dim budget_month2 As Currency
    Dim budget_month3 As Currency
        
    Special_Brand = Is_Special_Brand(Left(strIBID, 4))
    
    Set Rs_Temp = Nothing
    Rs_Temp.Open "Select getdate()", ConnERP, , , adCmdText
    
    strSql = "SELECT * FROM IB_TV where IB_ID='" & strIBID & "'"
    rs_IB_TV.Open strSql, ConnERP, adOpenStatic, adLockReadOnly
    
    '============================ Header ============================
    'Brand
    Rpt_TV_IB.Sections("Section4").Controls("Lbl_Brand").Caption = Brand
    'Brand Variant
    If IsNull(rs_IB_TV.Fields("Brand_Variant_Name").Value) Then
        Rpt_TV_IB.Sections("Section4").Controls("Lbl_Brand_Variant").Caption = ""
    Else
        Rpt_TV_IB.Sections("Section4").Controls("Lbl_Brand_Variant").Caption = rs_IB_TV.Fields("Brand_Variant_Name").Value
    End If
    'Date
    Rpt_TV_IB.Sections("Section4").Controls("Lbl_Date").Caption = Format(Date, "mmm dd yyyy")
    'IB ID
    Rpt_TV_IB.Sections("Section4").Controls("Lbl_IB_Id").Caption = strIBID
    
    '====================================== Detail =====================
    'Primary Target
    Rpt_TV_IB.Sections("Section4").Controls("Lbl_Primary_Target").Caption = rs_IB_TV("target_Primary").Value
    'Secontary Taget
    Rpt_TV_IB.Sections("Section4").Controls("Lbl_Secondary_Target").Caption = rs_IB_TV("target_secondary").Value
    If Special_Brand Then
        'Lable Primary
        Rpt_TV_IB.Sections("Section4").Controls("Lbl_Primary").Caption = "Primary/Brand"
        'Lable Secondary
        Rpt_TV_IB.Sections("Section4").Controls("Lbl_Secondary").Caption = "Secondary/ULI"
        'Approval Client
         Rpt_TV_IB.Sections("Section5").Controls("Lbl_Approval_Client").Caption = "Marketing Manager"
    Else
        'Lable Primary
        Rpt_TV_IB.Sections("Section4").Controls("Lbl_Primary").Caption = "Primary"
        'Lable Secondary
        Rpt_TV_IB.Sections("Section4").Controls("Lbl_Secondary").Caption = "Secondary"
        'Approval Client
         Rpt_TV_IB.Sections("Section5").Controls("Lbl_Approval_Client").Caption = "Marketing"
    End If
            
    '=============================== Television Objective =======================================
    strSql = "SELECT * FROM IB_TV_Objective WHERE IB_Id='" & strIBID & "' ORDER BY Objective_Id"
    Rs_TO.Open strSql, ConnERP, , , adCmdText
    
    'Loop TO
    Index_TO = 0
    Do While Not Rs_TO.EOF
        If Rs_TO.Fields("W_C_From").Value = Rs_TO.Fields("W_C_To").Value Then
            str_week = Format(Rs_TO.Fields("W_C_From").Value, "MMM dd")
        Else
            str_week = Format(Rs_TO.Fields("W_C_From").Value, "MMM dd") & " - " & Format(Rs_TO.Fields("W_C_To").Value, "MMM dd")
        End If
        
         Rpt_TV_IB.Sections("Section4").Controls("Lbl_Week_Comm_" & Index_TO).Caption = str_week
        'Camp Type
         Rpt_TV_IB.Sections("Section4").Controls("Lbl_Camp_Type_" & Index_TO).Caption = Rs_TO.Fields("Campaign_Type").Value
        'Freq
         Rpt_TV_IB.Sections("Section4").Controls("Lbl_Frequency_" & Index_TO).Caption = Rs_TO.Fields("Frequency").Value
        'Reach
        Rpt_TV_IB.Sections("Section4").Controls("Lbl_Reach_" & Index_TO).Caption = Rs_TO.Fields("Reach").Value & "%"
        'Tarps
        Rpt_TV_IB.Sections("Section4").Controls("Lbl_Tarps_" & Index_TO).Caption = Rs_TO.Fields("Tarps").Value
        'Budget
        Rpt_TV_IB.Sections("Section4").Controls("Lbl_Budget_Burst_" & Index_TO).Caption = Format(Rs_TO.Fields("Budget_With_MSC").Value, "##,##0")
                        
        'Tampilkan Line
        Rpt_TV_IB.Sections("Section4").Controls("Shape_WC_" & Index_TO + 1).Visible = True
        Rpt_TV_IB.Sections("Section4").Controls("Shape_CT_" & Index_TO + 1).Visible = True
        Rpt_TV_IB.Sections("Section4").Controls("Shape_F_" & Index_TO + 1).Visible = True
        Rpt_TV_IB.Sections("Section4").Controls("Shape_R_" & Index_TO + 1).Visible = True
        Rpt_TV_IB.Sections("Section4").Controls("Shape_T_" & Index_TO + 1).Visible = True
        Rpt_TV_IB.Sections("Section4").Controls("Shape_B_" & Index_TO + 1).Visible = True
        
        Rs_TO.MoveNext
        Index_TO = Index_TO + 1
        
        'Batas TO 10
        If Index_TO > 13 Then
            Exit Do
        End If
    Loop
            
    If Index_TO > 7 Then
        Rpt_TV_IB.Sections("Section4").Controls("Shape_W_2").Visible = True
        Rpt_TV_IB.Sections("Section4").Controls("Shape_W_3").Visible = True
        Rpt_TV_IB.Sections("Section4").Controls("Shape_W_4").Visible = True
        Rpt_TV_IB.Sections("Section4").Controls("Shape_W_5").Visible = True
        Rpt_TV_IB.Sections("Section4").Controls("Shape_W_6").Visible = True
        
        Rpt_TV_IB.Sections("Section4").Controls("Lbl_WC_2").Visible = True
        Rpt_TV_IB.Sections("Section4").Controls("Lbl_TOH_1").Visible = True
        Rpt_TV_IB.Sections("Section4").Controls("Lbl_TOH_2").Visible = True
        Rpt_TV_IB.Sections("Section4").Controls("Lbl_TOH_3").Visible = True
        Rpt_TV_IB.Sections("Section4").Controls("Lbl_TOH_4").Visible = True
        Rpt_TV_IB.Sections("Section4").Controls("Lbl_TOH_5").Visible = True
    End If
    
    Rs_TO.Close
    Set Rs_TO = Nothing
            
    '==================================== Campaign Outline ==============================
    With RsTempCampaign.Fields
        .Append "Objective_id", adVarChar, 100, adFldIsNullable
        
        .Append "Tarps_1", adInteger, , adFldIsNullable
        .Append "Tarps_2", adInteger, , adFldIsNullable
        .Append "Tarps_3", adInteger, , adFldIsNullable
        .Append "Tarps_4", adInteger, , adFldIsNullable
        .Append "Tarps_5", adInteger, , adFldIsNullable
        
        .Append "Tarps_6", adInteger, , adFldIsNullable
        .Append "Tarps_7", adInteger, , adFldIsNullable
        .Append "Tarps_8", adInteger, , adFldIsNullable
        .Append "Tarps_9", adInteger, , adFldIsNullable
        .Append "Tarps_10", adInteger, , adFldIsNullable
        
        .Append "Tarps_11", adInteger, , adFldIsNullable
        .Append "Tarps_12", adInteger, , adFldIsNullable
        .Append "Tarps_13", adInteger, , adFldIsNullable
        .Append "Tarps_14", adInteger, , adFldIsNullable
    End With
    
    RsTempCampaign.Open
        
    Year_Ib = rs_IB_TV.Fields("Year").Value
    Month_1 = rs_IB_TV.Fields("Month").Value
    
    If Month_1 = 12 Then
        Month_2 = 0
        Month_3 = 0
        
        Str_Week_Comm_1 = Get_WeekCommencing(Month_1, Year_Ib)
        Week_Count_Month1 = Len(Str_Week_Comm_1) / 2
        Week_Count_Month2 = 4
        Week_Count_Month3 = 4
        
    ElseIf Month_1 = 11 Then
        Month_2 = Month_1 + 1
        Month_3 = 0
        
        Str_Week_Comm_1 = Get_WeekCommencing(Month_1, Year_Ib)
        Week_Count_Month1 = Len(Str_Week_Comm_1) / 2
        Str_Week_Comm_2 = Get_WeekCommencing(Month_2, Year_Ib)
        Week_Count_Month2 = Len(Str_Week_Comm_2) / 2
        
        Week_Count_Month3 = 4
    Else
        Month_2 = Month_1 + 1
        Month_3 = Month_2 + 1
        
        Str_Week_Comm_1 = Get_WeekCommencing(Month_1, Year_Ib)
        Str_Week_Comm_2 = Get_WeekCommencing(Month_2, Year_Ib)
        Str_Week_Comm_3 = Get_WeekCommencing(Month_3, Year_Ib)
        Week_Count_Month1 = Len(Str_Week_Comm_1) / 2
        Week_Count_Month2 = Len(Str_Week_Comm_2) / 2
        Week_Count_Month3 = Len(Str_Week_Comm_3) / 2
    End If
        
    'Tampil Month1
    Week_Position = 1
    
    If Month_1 <> 0 Then
        Rpt_TV_IB.Sections("Section2").Controls("Lbl_Month_CO_1").Caption = Get_Month_Name(Month_1)
        
        'Width Lable
        Rpt_TV_IB.Sections("Section2").Controls("Lbl_Month_CO_1").Width = Week_Count_Month1 * 335
        'Width Frame
        Rpt_TV_IB.Sections("Section2").Controls("Shape_Month_1").Width = Week_Count_Month1 * 335
         
        'Tampillkan Week Comm
        'Week 1,2,3,4,5
        Do While Str_Week_Comm_1 <> ""
            Rpt_TV_IB.Sections("Section2").Controls("Lbl_Week_" & Week_Position).Caption = Mid(Str_Week_Comm_1, 1, 2)
            Str_Week_Comm_1 = Right(Str_Week_Comm_1, Len(Str_Week_Comm_1) - 2)
            Week_Position = Week_Position + 1
        Loop
    End If
        
    'Tampil Month2
    If Month_2 <> 0 Then
        Rpt_TV_IB.Sections("Section2").Controls("Lbl_Month_CO_2").Caption = Get_Month_Name(Month_2)
        
        'Width Lable
        Rpt_TV_IB.Sections("Section2").Controls("Lbl_Month_CO_2").Width = Week_Count_Month2 * 335
        'Width Frame
        Rpt_TV_IB.Sections("Section2").Controls("Shape_Month_2").Width = Week_Count_Month2 * 335
         
        'Tampillkan Week Comm
        'Week 1,2,3,4,5
        Do While Str_Week_Comm_2 <> ""
            Rpt_TV_IB.Sections("Section2").Controls("Lbl_Week_" & Week_Position).Caption = Mid(Str_Week_Comm_2, 1, 2)
            Str_Week_Comm_2 = Right(Str_Week_Comm_2, Len(Str_Week_Comm_2) - 2)
            Week_Position = Week_Position + 1
        Loop
    Else
        'Width Lable
        Rpt_TV_IB.Sections("Section2").Controls("Lbl_Month_CO_2").Width = 5 * 335
        'Width Frame
        Rpt_TV_IB.Sections("Section2").Controls("Shape_Month_2").Width = 5 * 335
         
    End If
        
    'Tampil Month3
    If Month_3 <> 0 Then
        Rpt_TV_IB.Sections("Section2").Controls("Lbl_Month_CO_3").Caption = Get_Month_Name(Month_3)
       
         'Width Lable
        Rpt_TV_IB.Sections("Section2").Controls("Lbl_Month_CO_3").Width = Week_Count_Month3 * 335
        'Width Frame
        Rpt_TV_IB.Sections("Section2").Controls("Shape_Month_3").Width = Week_Count_Month3 * 335
         
        'Tampillkan Week Comm
        'Week 1,2,3,4,5
        Do While Str_Week_Comm_3 <> ""
            Rpt_TV_IB.Sections("Section2").Controls("Lbl_Week_" & Week_Position).Caption = Mid(Str_Week_Comm_3, 1, 2)
            Str_Week_Comm_3 = Right(Str_Week_Comm_3, Len(Str_Week_Comm_3) - 2)
            Week_Position = Week_Position + 1
        Loop
    Else
        If Month_2 = 0 Then
             'Width Lable
            Rpt_TV_IB.Sections("Section2").Controls("Lbl_Month_CO_3").Width = 4 * 335
            'Width Frame
            Rpt_TV_IB.Sections("Section2").Controls("Shape_Month_3").Width = 4 * 335
        Else
            'Week_Count_Month2
            Rpt_TV_IB.Sections("Section2").Controls("Lbl_Month_CO_3").Width = (14 - (Week_Count_Month2 + Week_Count_Month1)) * 335
            'Width Frame
            Rpt_TV_IB.Sections("Section2").Controls("Shape_Month_3").Width = (14 - (Week_Count_Month2 + Week_Count_Month1)) * 335
        End If
    End If
    
    'Left Dari Frame & Lable
    Rpt_TV_IB.Sections("Section2").Controls("Lbl_Month_CO_2").Left = Rpt_TV_IB.Sections("Section2").Controls("Lbl_Month_CO_1").Left + Rpt_TV_IB.Sections("Section2").Controls("Lbl_Month_CO_1").Width
    Rpt_TV_IB.Sections("Section2").Controls("Lbl_Month_CO_3").Left = Rpt_TV_IB.Sections("Section2").Controls("Lbl_Month_CO_2").Left + Rpt_TV_IB.Sections("Section2").Controls("Lbl_Month_CO_2").Width
    'Fame
    Rpt_TV_IB.Sections("Section2").Controls("Shape_Month_2").Left = Rpt_TV_IB.Sections("Section2").Controls("Shape_Month_1").Left + Rpt_TV_IB.Sections("Section2").Controls("Shape_Month_1").Width
    Rpt_TV_IB.Sections("Section2").Controls("Shape_Month_3").Left = Rpt_TV_IB.Sections("Section2").Controls("Shape_Month_2").Left + Rpt_TV_IB.Sections("Section2").Controls("Shape_Month_2").Width
    
    If (Week_Count_Month1 + Week_Count_Month2 + Week_Count_Month3) = 13 Then
         Rpt_TV_IB.Sections("Section2").Controls("ShapeCO_1").Visible = False
         Rpt_TV_IB.Sections("Section1").Controls("ShapeCO_2").Visible = False
    End If
        
    'Open Material
    strSql = "SELECT * From IB_TV_Objective_Material WHERE "
    strSql = strSql & " Objective_Id IN (SELECT Objective_Id FROM IB_TV_Objective WHERE "
    strSql = strSql & " IB_ID ='" & strIBID & "') ORDER BY Objective_Id"
    
    RS_Materi.Open strSql, ConnERP, , , adCmdText
    
    'Loop Marterial
    Do While Not RS_Materi.EOF
        'assign to recordset temp
        RsTempCampaign.AddNew
        RsTempCampaign.Fields("Objective_id").Value = RS_Materi.Fields("Objective_Id").Value & ":" & RS_Materi.Fields("Material_Name").Value & ", " & RS_Materi.Fields("Material_Duration").Value
        
        For Week_Position = 1 To 14
            If Trim(Rpt_TV_IB.Sections("Section2").Controls("Lbl_Week_" & Week_Position).Caption) <> "" Then
                strSql = "SELECT * FROM IB_TV_Campaign WHERE "
                strSql = strSql & " Objective_Id = " & RS_Materi.Fields("Objective_Id").Value
                strSql = strSql & " AND Material_Name='" & Clear_String(RS_Materi.Fields("Material_Name").Value) & "'"
                strSql = strSql & " AND Material_Duration = " & RS_Materi.Fields("Material_Duration").Value
                
                If Week_Position <= Week_Count_Month1 Then
                    strSql = strSql & " AND Month(Month_Week_commencing)=" & Month_1
                ElseIf Week_Position > Week_Count_Month1 And Week_Position <= (Week_Count_Month1 + Week_Count_Month2) Then
                    strSql = strSql & " AND Month(Month_Week_commencing)=" & Month_2
                Else
                    strSql = strSql & " AND Month(Month_Week_commencing)=" & Month_3
                End If
                
                strSql = strSql & " AND Day(Week_commencing)=" & Trim(Rpt_TV_IB.Sections("Section2").Controls("Lbl_Week_" & Week_Position).Caption)
                Rs_Co.Open strSql, ConnERP, , , adCmdText
                
                StrField = "Tarps_" & Week_Position
                
                'Loop Campaign Outline
                Do While Not Rs_Co.EOF
                    'Tampil Di CO Table
                    RsTempCampaign.Fields(StrField).Value = Rs_Co.Fields("Tarps_Per_Week").Value
                    Rs_Co.MoveNext
                Loop
                
                Rs_Co.Close
            End If
        Next Week_Position
        
        RS_Materi.MoveNext
        RsTempCampaign.Update
        
    Loop
    
    RS_Materi.Close
    Set RS_Materi = Nothing
        
    '=======================End  Rs_Material ==========================
    
    Set Rpt_TV_IB.DataSource = RsTempCampaign
    
    'Assign Recordsource
    Rpt_TV_IB.Sections("Section1").Controls("Txt_Material_0").DataField = "Objective_id"
    
    Rpt_TV_IB.Sections("Section1").Controls("Txt_CO_01").DataField = "Tarps_1"
    Rpt_TV_IB.Sections("Section1").Controls("Txt_CO_02").DataField = "Tarps_2"
    Rpt_TV_IB.Sections("Section1").Controls("Txt_CO_03").DataField = "Tarps_3"
    Rpt_TV_IB.Sections("Section1").Controls("Txt_CO_04").DataField = "Tarps_4"
    
    Rpt_TV_IB.Sections("Section1").Controls("Txt_CO_05").DataField = "Tarps_5"
    Rpt_TV_IB.Sections("Section1").Controls("Txt_CO_06").DataField = "Tarps_6"
    Rpt_TV_IB.Sections("Section1").Controls("Txt_CO_07").DataField = "Tarps_7"
    Rpt_TV_IB.Sections("Section1").Controls("Txt_CO_08").DataField = "Tarps_8"
    
    Rpt_TV_IB.Sections("Section1").Controls("Txt_CO_09").DataField = "Tarps_9"
    Rpt_TV_IB.Sections("Section1").Controls("Txt_CO_10").DataField = "Tarps_10"
    Rpt_TV_IB.Sections("Section1").Controls("Txt_CO_11").DataField = "Tarps_11"
    Rpt_TV_IB.Sections("Section1").Controls("Txt_CO_12").DataField = "Tarps_12"
    
    Rpt_TV_IB.Sections("Section1").Controls("Txt_CO_13").DataField = "Tarps_13"
    Rpt_TV_IB.Sections("Section1").Controls("Txt_CO_14").DataField = "Tarps_14"
        
    '===================================== Budget Split By Month ==========================
    Rs_Montly_Budget.Open "Select Month,budget From IB_TV_Montly_Budget WHERE Client_Brief_ID='" & rs_IB_TV.Fields("Client_brief_id").Value & "' and IB_ID='" & strIBID & "'", ConnERP, adOpenKeyset, adLockReadOnly, adCmdText
    
    If Not Rs_Montly_Budget.BOF And Not Rs_Montly_Budget.EOF Then
        Rs_Montly_Budget.MoveFirst
        
        lbl_month1 = Get_Month_Name(Rs_Montly_Budget.Fields("Month").Value)
        budget_month1 = Format(Rs_Montly_Budget.Fields("budget").Value, "##,##0")
        
        Rs_Montly_Budget.MoveNext
        
        If Not Rs_Montly_Budget.EOF Then
            budget_month2 = Format(Rs_Montly_Budget.Fields("Budget").Value, "##,##0")
            lbl_month2 = Get_Month_Name(Rs_Montly_Budget.Fields("Month").Value)
        Else
            GoTo Close_Recodset
        End If
        
        Rs_Montly_Budget.MoveNext
        
        If Not Rs_Montly_Budget.EOF Then
            budget_month3 = Format(Rs_Montly_Budget.Fields("budget").Value, "##,##0")
            lbl_month3 = Get_Month_Name(Rs_Montly_Budget.Fields("Month").Value)
        Else
            GoTo Close_Recodset
        End If
    End If
    
    Rs_Montly_Budget.Close
        
Close_Recodset:
    If Rs_Montly_Budget.State = adStateOpen Then Rs_Montly_Budget.Close
    
    Me.MousePointer = 0
        
    'Month
    Rpt_TV_IB.Sections("Section4").Controls("Lbl_Month_1").Caption = " " & lbl_month1
    Rpt_TV_IB.Sections("Section4").Controls("Lbl_Month_2").Caption = " " & lbl_month2
    Rpt_TV_IB.Sections("Section4").Controls("Lbl_Month_3").Caption = " " & lbl_month3
    
    'Budget
    Rpt_TV_IB.Sections("Section4").Controls("Lbl_Budget_1").Caption = budget_month1
    Rpt_TV_IB.Sections("Section4").Controls("Lbl_Budget_2").Caption = budget_month2
    Rpt_TV_IB.Sections("Section4").Controls("Lbl_Budget_3").Caption = budget_month3
'======================================== end budget ====================

    '---> Program Type Consideration
    Rpt_TV_IB.Sections("Section5").Controls("Lbl_Program_Consideration").Caption = rs_IB_TV("Consideration").Value
    '---> Other Consideration
    Rpt_TV_IB.Sections("Section5").Controls("Lbl_Other_Consideration").Caption = rs_IB_TV("Other_Consideration").Value
    '---> Attachment
    Rpt_TV_IB.Sections("Section5").Controls("Lbl_Attachment").Caption = rs_IB_TV("Attachment").Value
    
    Rpt_TV_IB.show 1

End Sub

Private Sub Show_IB_RD(strIBID As String, Brand As String)
'************************************************************
' Unit/Module Name  : Show_IB_RD
' Function          : show ib rd
'                   :
' Public Variables  :
' Used Tables       :
' Date              : 11.10.05
' Created By        : Diyah
' Last Update/By    :
'************************************************************
    With Crpt
        .Reset
        .SelectionFormula = "{ib_radio.ib_id}='" & strIBID & "'"
        .Connect = "DSN =" & Server_Name & ";UID = " & Login_User & ";DSQ = " & strDatabase_Name & "; PWD =" & Login_Password
        .ReportFileName = Report_Dir & "\radio\IB_Radio_area.rpt"
        
        If Is_Special_Brand(Left(strIBID, 4)) = True Then
            .Formulas(0) = "Marketing ='Unilever Marketing'"
        Else
            .Formulas(0) = "Marketing ='Marketing'"
        End If
        
        .WindowShowRefreshBtn = True
        .WindowShowPrintSetupBtn = True
        .WindowState = crptMaximized
        .WindowTitle = " -- IB Radio -- "
        .Action = 1
    End With
End Sub

Private Sub Show_IB_PR(strIBID As String, Brand As String)
'************************************************************
' Unit/Module Name  : Show_IB_PR
' Function          : show ib print
'                   :
' Public Variables  :
' Used Tables       :
' Date              : 11.10.05
' Created By        : Diyah
' Last Update/By    :
'************************************************************
    Dim rs_Br_Name As New ADODB.Recordset
    Dim rs_App As New ADODB.Recordset
    
    With Crpt
        .Reset
        
        Me.MousePointer = vbHourglass
        .ReportFileName = Report_Dir + "\Print\ib_print.rpt"
        
        strSql = "SELECT ib_print.brand_variant_name,ib_print.entered_date,brand.brand_name FROM ib_print, brand WHERE substring(ib_print.ib_id,1,4)=brand.brand_Code and ib_id='" & strIBID & "'"
        rs_Br_Name.Open strSql, ConnERP, adOpenStatic, adLockReadOnly
        
        .Formulas(0) = "Brand='" & Brand & "'"
        .Formulas(11) = "Brand_Variant='" & rs_Br_Name("brand_variant_name").Value & "'"
        .Formulas(10) = "Entered_Date='" & Format(rs_Br_Name("entered_date").Value, "mmm dd, yyyy") & "'"
        
        strSql = "SELECT  app_ib_uli, app_ib_non_uli FROM print_information"
        rs_App.Open strSql, ConnERP, adOpenStatic, adLockReadOnly
        
        If Is_Special_Brand(Left(strIBID, 4)) = True Then
            .Formulas(1) = "Primary ='Primary / Brand'"
            .Formulas(2) = "Secondary='Secondary / ULI'"
            .Formulas(3) = "Approval='" & rs_App(0) & "'"
        Else
            .Formulas(1) = "Primary ='Primary'"
            .Formulas(2) = "Secondary='Secondary'"
            .Formulas(3) = "Approval='" & rs_App(1) & "'"
        End If
        
        rs_App.Close
        Set rs_App = Nothing
        
        rs_Br_Name.Close
        Set rs_Br_Name = Nothing
        
        strSql = "SELECT  month, budget FROM ib_print_plan WHERE ib_id='" & strIBID & "' ORDER BY month"
        rs_Br_Name.Open strSql, ConnERP, adOpenStatic, adLockReadOnly
        
        If Not rs_Br_Name.EOF And Not rs_Br_Name.BOF Then
            rs_Br_Name.MoveFirst
        
           .Formulas(4) = "Bulan1='" & Get_Month_Name(rs_Br_Name(0)) & "'"
           .Formulas(5) = "budget1= '" & Format(rs_Br_Name(1), "#,##0") & "'"
           rs_Br_Name.MoveNext
        
            If Not rs_Br_Name.EOF And Not rs_Br_Name.BOF Then
               .Formulas(6) = "Bulan2='" & Get_Month_Name(rs_Br_Name(0)) & "'"
               .Formulas(7) = "budget2= '" & Format(rs_Br_Name(1), "#,##0") & "'"
               rs_Br_Name.MoveNext
            Else
               .Formulas(6) = "Bulan2=''"
               .Formulas(7) = "budget2= ''"
            End If
            
            If Not rs_Br_Name.EOF And Not rs_Br_Name.BOF Then
               .Formulas(8) = "Bulan3='" & Get_Month_Name(rs_Br_Name(0)) & "'"
               .Formulas(9) = "budget3= '" & Format(rs_Br_Name(1), "#,##0") & "'"
            Else
               .Formulas(8) = "Bulan3=''"
               .Formulas(9) = "budget3=''"
            End If
        Else
           .Formulas(4) = "Bulan1=''"
           .Formulas(5) = "budget1= ''"
        End If
                 
        strSql = "{ib_print.ib_id}='" & strIBID & "'"
        
        .SelectionFormula = strSql
        .WindowShowPrintBtn = True
        .WindowShowPrintSetupBtn = True
        .WindowShowRefreshBtn = True
        .WindowState = crptMaximized
        .WindowTitle = "IB Print " & strIBID
        .Connect = "DSN=" & Server_Name & ";UID=" & Login_User & ";PWD=" & Login_Password & ";DSQ=" & strDatabase_Name & ""
        .Action = 1
        
        rs_Br_Name.Close

        Me.MousePointer = vbDefault
    End With

End Sub

Private Sub Show_IB_OT_CN(strIBID As String, Brand As String)
'************************************************************
' Unit/Module Name  : Show_IB_OT_CN
' Function          : show ib other
'                   :
' Public Variables  :
' Used Tables       :
' Date              : 11.10.05
' Created By        : Diyah
' Last Update/By    :
'************************************************************
    Me.MousePointer = vbHourglass
    
    With Crpt
        .Reset
        .ReportFileName = Report_Dir & "\Other\OTHER_IB.RPT"
        .Formulas(0) = "Brand = '" & Brand & "'"
        .SelectionFormula = " {IB_OTHER.IB_ID } = '" & Trim(strIBID) & "'"
        .RetrieveDataFiles
        .WindowShowPrintSetupBtn = True
        .WindowState = crptMaximized
        .WindowShowPrintBtn = True
        .Connect = "DSN=" & Server_Name & ";UID=" & Login_User & ";PWD=" & Login_Password & ";DSQ=" & strDatabase_Name & ""
        .Action = 1
    End With

    Me.MousePointer = vbNormal
End Sub

Private Sub Show_Quot_TV(IB_ID As String, Brand As String)
'************************************************************
' Unit/Module Name  : Show_IB_TV
' Function          : show quotation tv
'                   :
' Public Variables  :
' Used Tables       :
' Date              : 12.10.05
' Created By        : Diyah
' Last Update/By    :
'************************************************************
    Dim Object_Index As Integer
    Dim Rs_Temp As New ADODB.Recordset
    Dim Rs_Quot_TV As New ADODB.Recordset
    Dim Rs_TV_detail As New ADODB.Recordset
    
    Me.MousePointer = vbHourglass
    
    Object_Index = 1
        
    'Temporary Recordset untuk Data Report
    '=====================================
    Set Rs_Temp = Nothing
    Rs_Temp.Open "Select getdate()", ConnERP, , , adCmdText
    
    Set Rpt_TV_Media_Quotation.DataSource = Rs_Temp
       
    '===================================== Header =======================
    strSql = "SELECT * FROM IB_TV_Quotation WHERE IB_ID='" & IB_ID & "'"
    Rs_Quot_TV.Open strSql, ConnERP, adOpenStatic, adLockReadOnly
    
    'IB ID
    Rpt_TV_Media_Quotation.Sections("Section2").Controls("Lbl_IB_id").Caption = IB_ID
    'Client
    Rpt_TV_Media_Quotation.Sections("Section2").Controls("Lbl_Client").Caption = Get_Client_Name(Left(IB_ID, 4))
    'Brand
    Rpt_TV_Media_Quotation.Sections("Section2").Controls("Lbl_Brand").Caption = UCase(Brand)
    'Media Plan No
    Rpt_TV_Media_Quotation.Sections("Section2").Controls("Lbl_Plan_No").Caption = Rs_Quot_TV("Media_Plan_No").Value
    'Dated
    Rpt_TV_Media_Quotation.Sections("Section2").Controls("Lbl_Dated").Caption = Format(Rs_Quot_TV("Date").Value, "mmm dd yyyy")
    'remarks
    Rpt_TV_Media_Quotation.Sections("Section1").Controls("Lbl_Remark").Caption = IIf(IsNull(Rs_Quot_TV("Remarks").Value), "", Rs_Quot_TV("Remarks").Value)
    
    '======================== Detail ==================================
                
    strSql = "SELECT * FROM IB_TV_Quotation_Detail WHERE "
    strSql = strSql & " ib_id='" & IB_ID & "'"
    
    Rs_TV_detail.Open strSql, strApplication_Name, , , adCmdText
    
    While Not Rs_TV_detail.EOF
    'Show To Report
        Rpt_TV_Media_Quotation.Sections("Section1").Controls("Lbl_Month_" & Object_Index).Caption = Get_Month_Name(Rs_TV_detail.Fields("Month").Value) & " " & Rs_TV_detail.Fields("Year").Value
        Rpt_TV_Media_Quotation.Sections("Section1").Controls("Lbl_Job_Id_" & Object_Index).Caption = Rs_TV_detail.Fields("Job_Id").Value
        'Nett Cost
        Rpt_TV_Media_Quotation.Sections("Section1").Controls("Lbl_Nett_Cost_" & Object_Index).Caption = Format(Rs_TV_detail.Fields("Nett_Cost").Value, "##,##0")
        'MSC
        Rpt_TV_Media_Quotation.Sections("Section1").Controls("Lbl_MSC_" & Object_Index).Caption = Format(Rs_TV_detail.Fields("Media_Sptv_Charge").Value + Rs_TV_detail.Fields("Bonus").Value, "##,##0")
        'Others
        Rpt_TV_Media_Quotation.Sections("Section1").Controls("Lbl_Other_" & Object_Index).Caption = Format(Rs_TV_detail.Fields("Other_Charge").Value, "##,##0")
        'total Lintas
        Rpt_TV_Media_Quotation.Sections("Section1").Controls("Lbl_Total_Lintas_" & Object_Index).Caption = Format(Rs_TV_detail.Fields("Total_Lintas").Value, "##,##0")
        'Job_No CA
        Rpt_TV_Media_Quotation.Sections("Section1").Controls("Lbl_Job_No_CA_" & Object_Index).Caption = Rs_TV_detail.Fields("Job_Number_Agency").Value
        'MSC CA
        Rpt_TV_Media_Quotation.Sections("Section1").Controls("Lbl_Msc_CA_" & Object_Index).Caption = Format(Rs_TV_detail.Fields("Agency_Charge").Value, "##,##0")
        'Grand Total
        Rpt_TV_Media_Quotation.Sections("Section1").Controls("Lbl_Grand_Total_" & Object_Index).Caption = Format(Rs_TV_detail.Fields("Grand_Total").Value, "##,##0")
        '1,2,3
        Rpt_TV_Media_Quotation.Sections("Section1").Controls("Line" & Object_Index).Visible = True
        '7,8,9
        Rpt_TV_Media_Quotation.Sections("Section1").Controls("Line" & 6 + Object_Index).Visible = True
        '10,11,12
        Rpt_TV_Media_Quotation.Sections("Section1").Controls("Line" & 9 + Object_Index).Visible = True
        '13,14,15
        Rpt_TV_Media_Quotation.Sections("Section1").Controls("Line" & 12 + Object_Index).Visible = True
        '16,17,18
        Rpt_TV_Media_Quotation.Sections("Section1").Controls("Line" & 15 + Object_Index).Visible = True
        
        Rs_TV_detail.MoveNext
        Object_Index = Object_Index + 1
    Wend
    
'==================================== Show Report =====================
    Rpt_TV_Media_Quotation.show 1
    
    Rs_TV_detail.Close
    Rs_Quot_TV.Close
    
    Me.MousePointer = vbNormal
End Sub

Private Sub Show_Quot_RD(IB_ID As String)
'************************************************************
' Unit/Module Name  : Show_IB_RD
' Function          : show quotation radio
'                   :
' Public Variables  :
' Used Tables       :
' Date              : 11.10.05
' Created By        : Diyah
' Last Update/By    :
'************************************************************
    Dim TxtSQl As String
    Dim rs As New ADODB.Recordset
    
    Me.MousePointer = vbHourglass
    
    TxtSQl = " select   Month_Catalog.Month_Name, Brand.Brand_Name, "
    TxtSQl = TxtSQl & " Client.Client_Name, "
    TxtSQl = TxtSQl & " Company.Company_Name, "
    TxtSQl = TxtSQl & " IB_Radio_Quot.Ib_Id as MediaPlanNo, "
    TxtSQl = TxtSQl & " IB_Radio_Quot.Ib_Id , "
    TxtSQl = TxtSQl & " IB_Radio_Quot.Remarks, "
    TxtSQl = TxtSQl & " IB_Radio_Quot.Date, "
    TxtSQl = TxtSQl & " IB_Radio_Quot.plan_no, "
    TxtSQl = TxtSQl & " IB_Radio_Quotation_Detail.Job_Id, "
    TxtSQl = TxtSQl & " IB_Radio_Quotation_Detail.Month, "
    TxtSQl = TxtSQl & " IB_Radio_Quotation_Detail.Year, "
    TxtSQl = TxtSQl & " IB_Radio_Quotation_Detail.Nett_Cost, "
    TxtSQl = TxtSQl & " IB_Radio_Quotation_Detail.Media_Sptv_Charge, "
    TxtSQl = TxtSQl & " IB_Radio_Quotation_Detail.Other_Charge, "
    TxtSQl = TxtSQl & " IB_Radio_Quotation_Detail.Bonus, "
    TxtSQl = TxtSQl & " IB_Radio_Quotation_Detail.Total_Lintas, "
    TxtSQl = TxtSQl & " IB_Radio_Quotation_Detail.Agency_Charge, "
    TxtSQl = TxtSQl & " IB_Radio_Quotation_Detail.Job_Number_Agency, "
    TxtSQl = TxtSQl & " IB_Radio_Quotation_Detail.Grand_Total "
    TxtSQl = TxtSQl & " from (ib_radio_quot inner join  IB_Radio_Quotation_Detail on IB_Radio_Quotation_Detail.ib_ID = IB_Radio_Quot.ib_ID ) "
    TxtSQl = TxtSQl & " inner join Brand on left(IB_Radio_quot.ib_id,4) = brand.brand_code "
    TxtSQl = TxtSQl & " inner join client on brand.Client_Code = Client.Client_Code "
    TxtSQl = TxtSQl & " inner join company on Brand.Company_Code = Company.Company_Code "
    TxtSQl = TxtSQl & " inner join Month_Catalog on Month_Catalog.Month = ib_radio_quotation_detail.Month "
    TxtSQl = TxtSQl & " where ib_radio_Quot.Ib_ID ='" & IB_ID & "'"
    TxtSQl = TxtSQl & " order by IB_Radio_Quotation_Detail.month asc"
    
    rs.Open TxtSQl, ConnERP, adOpenStatic, adLockReadOnly
    
    Rem Filename .RPT
    Crpt.Reset
    Crpt.ReportFileName = Report_Dir & "\radio\mq_Radio.rpt"
    
    Rem Header Report
    With rs
        If .EOF = False Then
            Crpt.ParameterFields(39) = "Client;" & .Fields("Client_Name") & ";TRUE"
            Crpt.ParameterFields(1) = "Brand;" & .Fields("Brand_Name") & ";TRUE"
            Crpt.ParameterFields(2) = "MediaType;Radio;TRUE"
            Crpt.ParameterFields(3) = "MediaPlanNo;" & IIf(IsNull(.Fields("Plan_No")) = True, "", .Fields("Plan_No")) & ";TRUE"
            Crpt.ParameterFields(4) = "Dated;" & Format(CDate(.Fields("Date")), "mmm/dd/yyyy") & ";TRUE"
            Crpt.ParameterFields(5) = "IBID;" & .Fields("IB_ID") & ";TRUE"
            Crpt.ParameterFields(6) = "Remarks;" & .Fields("Remarks") & ";TRUE"
            Crpt.ParameterFields(7) = "PT;PT. INITIATIF MEDIA INDONESIA;TRUE"
            Crpt.ParameterFields(8) = "Marketing;Marketing Manager;TRUE"
        End If
        
        Rem 1st Month
        If .EOF = False Then
            Crpt.ParameterFields(9) = "MONTH1;" & .Fields("Month_Name") & ";TRUE"
            Crpt.ParameterFields(10) = "Year1;" & .Fields("Year") & ";TRUE"
            Crpt.ParameterFields(11) = "Nett1;" & .Fields("Nett_Cost") & ";TRUE"
            Crpt.ParameterFields(12) = "MSC1;" & .Fields("Media_Sptv_Charge") + .Fields("bonus") & ";TRUE"
            Crpt.ParameterFields(13) = "Other1;" & .Fields("Other_Charge") & ";TRUE"
            Crpt.ParameterFields(14) = "TotalLintas1;" & .Fields("Total_Lintas") & ";TRUE"
            Crpt.ParameterFields(15) = "JobNoAG1;" & .Fields("Job_Number_Agency") & ";TRUE"
            Crpt.ParameterFields(16) = "ClubCharge1;" & .Fields("Agency_Charge") & ";TRUE"
            Crpt.ParameterFields(17) = "GrandTotal1;" & .Fields("Grand_Total") & ";TRUE"
            Crpt.ParameterFields(36) = "JobID1;" & .Fields("Job_ID") & ";TRUE"
            .MoveNext
        Else
            Crpt.ParameterFields(9) = "MONTH1;;TRUE"
            Crpt.ParameterFields(10) = "Year1;;TRUE"
            Crpt.ParameterFields(11) = "Nett1;0;TRUE"
            Crpt.ParameterFields(12) = "MSC1;0;TRUE"
            Crpt.ParameterFields(13) = "Other1;0;TRUE"
            Crpt.ParameterFields(14) = "TotalLintas1;0;TRUE"
            Crpt.ParameterFields(15) = "JobNoAG1;;TRUE"
            Crpt.ParameterFields(16) = "clubcharge1;0;TRUE"
            Crpt.ParameterFields(17) = "GrandTotal1;0;TRUE"
            Crpt.ParameterFields(36) = "JobID1;;TRUE"
    
        End If
       
        Rem 2nd Month
        If .EOF = False Then
            Crpt.ParameterFields(18) = "MONTH2;" & .Fields("Month_Name") & ";TRUE"
            Crpt.ParameterFields(19) = "Year2;" & .Fields("Year") & ";TRUE"
            Crpt.ParameterFields(20) = "Nett2;" & .Fields("Nett_Cost") & ";TRUE"
            Crpt.ParameterFields(21) = "MSC2;" & .Fields("Media_Sptv_Charge") + .Fields("bonus") & ";TRUE"
            Crpt.ParameterFields(22) = "Other2;" & .Fields("Other_Charge") & ";TRUE"
            Crpt.ParameterFields(23) = "TotalLintas2;" & .Fields("Total_Lintas") & ";TRUE"
            Crpt.ParameterFields(24) = "JobNoAG2;" & .Fields("Job_Number_Agency") & ";TRUE"
            Crpt.ParameterFields(25) = "clubcharge2;" & .Fields("Agency_Charge") & ";TRUE"
            Crpt.ParameterFields(26) = "GrandTotal2;" & .Fields("Grand_Total") & ";TRUE"
            Crpt.ParameterFields(37) = "JobID2;" & .Fields("Job_ID") & ";TRUE"
            .MoveNext
        Else
            Crpt.ParameterFields(18) = "MONTH2;;TRUE"
            Crpt.ParameterFields(19) = "Year2;;TRUE"
            Crpt.ParameterFields(20) = "Nett2;0;TRUE"
            Crpt.ParameterFields(21) = "MSC2;0;TRUE"
            Crpt.ParameterFields(22) = "Other2;0;TRUE"
            Crpt.ParameterFields(23) = "TotalLintas2;0;TRUE"
            Crpt.ParameterFields(24) = "JobNoAG2;;TRUE"
            Crpt.ParameterFields(25) = "clubcharge2;0;TRUE"
            Crpt.ParameterFields(26) = "GrandTotal2;0;TRUE"
            Crpt.ParameterFields(37) = "JobID2;;TRUE"
            
        End If
        
        Rem 3rd Month
        If .EOF = False Then
            Crpt.ParameterFields(27) = "MONTH3;" & .Fields("Month_Name") & ";TRUE"
            Crpt.ParameterFields(28) = "Year3;" & .Fields("Year") & ";TRUE"
            Crpt.ParameterFields(29) = "Nett3;" & .Fields("Nett_Cost") & ";TRUE"
            Crpt.ParameterFields(30) = "MSC3;" & .Fields("Media_Sptv_Charge") + .Fields("bonus") & ";TRUE"
            Crpt.ParameterFields(31) = "Other3;" & .Fields("Other_Charge") & ";TRUE"
            Crpt.ParameterFields(32) = "TotalLintas3;" & .Fields("Total_Lintas") & ";TRUE"
            Crpt.ParameterFields(33) = "JobNoAG3;" & .Fields("Job_Number_Agency") & ";TRUE"
            Crpt.ParameterFields(34) = "clubcharge3;" & .Fields("Agency_Charge") & ";TRUE"
            Crpt.ParameterFields(35) = "GrandTotal3;" & .Fields("Grand_Total") & ";TRUE"
            Crpt.ParameterFields(38) = "JobID3;" & .Fields("Job_ID") & ";TRUE"
            .MoveNext
        Else
            Crpt.ParameterFields(27) = "MONTH3;;TRUE"
            Crpt.ParameterFields(28) = "Year3;;TRUE"
            Crpt.ParameterFields(29) = "Nett3;0;TRUE"
            Crpt.ParameterFields(30) = "MSC3;0;TRUE"
            Crpt.ParameterFields(31) = "Other3;0;TRUE"
            Crpt.ParameterFields(32) = "TotalLintas3;0;TRUE"
            Crpt.ParameterFields(33) = "JobNoAG3;;TRUE"
            Crpt.ParameterFields(34) = "clubcharge3;0;TRUE"
            Crpt.ParameterFields(35) = "GrandTotal3;0;TRUE"
            Crpt.ParameterFields(38) = "JobID3;;TRUE"
        End If
    End With
    
    Crpt.WindowState = crptMaximized
    Crpt.WindowShowRefreshBtn = True
    Crpt.WindowShowPrintSetupBtn = True
    Crpt.WindowTitle = " -- Implementation Brief Radio Quotation -- "
    Crpt.Connect = "DSN = " & Server_Name & ";UID = " & Login_User & ";PWD = " & Login_Password & ";DSQ = " & strDatabase_Name
    Crpt.Action = 1
    
    Me.MousePointer = vbNormal
End Sub

Private Sub Show_Quot_PR(IB_ID As String, Brand As String)
'************************************************************
' Unit/Module Name  : Show_IB_PR
' Function          : show quotation print
'                   :
' Public Variables  :
' Used Tables       :
' Date              : 11.10.05
' Created By        : Diyah
' Last Update/By    :
'************************************************************
    Dim Object_Index As Integer
    Dim Rs_Temp As New ADODB.Recordset
    Dim Rs_Quot_PR As New ADODB.Recordset
    Dim Rs_Print_detail As New ADODB.Recordset
    Dim rs_client_name As New ADODB.Recordset
    
    Me.MousePointer = vbHourglass
    
    Object_Index = 1
        
    'Temporary Recordset untuk Data Report
    '=====================================
    Set Rs_Temp = Nothing
    Rs_Temp.Open "Select getdate()", ConnERP, , , adCmdText
    Set Rpt_Print_Media_Quotation.DataSource = Rs_Temp
       
    'Header
    '=====================================
    strSql = "SELECT * FROM IB_Print_Quotation where IB_ID='" & IB_ID & "'"
    Rs_Quot_PR.Open strSql, ConnERP, adOpenStatic, adLockReadOnly
    
    strSql = "select client_name from client where client_code = (select client_code from brand where brand_code ='" & Left(IB_ID, 4) & "')"
    rs_client_name.Open strSql, ConnERP, adOpenStatic, adLockOptimistic

    'IB ID
    Rpt_Print_Media_Quotation.Sections("Section2").Controls("Lbl_IB_id").Caption = IB_ID
    'client
    Rpt_Print_Media_Quotation.Sections("Section2").Controls("Lbl_Client").Caption = rs_client_name(0)
    'Brand
    Rpt_Print_Media_Quotation.Sections("Section2").Controls("Lbl_Brand").Caption = UCase(Brand)
    'Media Plan No
    Rpt_Print_Media_Quotation.Sections("Section2").Controls("Lbl_Plan_No").Caption = Rs_Quot_PR("Plan_No").Value
    'Dated
    Rpt_Print_Media_Quotation.Sections("Section2").Controls("Lbl_Dated").Caption = Format(Rs_Quot_PR("Date").Value, "mmm dd, yyyy")
            
    '=================================== detail ============================
    strSql = "SELECT * FROM IB_Print_Quotation_Detail WHERE "
    strSql = strSql & " ib_id='" & IB_ID & "'"
    'Rs_Print_detail.Close
    Rs_Print_detail.Open strSql, ConnERP, , , adCmdText
    
    While Not Rs_Print_detail.EOF
    'Show To Report
        Rpt_Print_Media_Quotation.Sections("Section1").Controls("Lbl_Month_" & Object_Index).Caption = Get_Month_Name(Rs_Print_detail.Fields("Month").Value) & " " & Rs_Print_detail.Fields("Year").Value
        Rpt_Print_Media_Quotation.Sections("Section1").Controls("Lbl_Job_Id_" & Object_Index).Caption = Rs_Print_detail.Fields("Job_Id").Value
        'Nett Cost
        Rpt_Print_Media_Quotation.Sections("Section1").Controls("Lbl_Nett_Cost_" & Object_Index).Caption = Format(Rs_Print_detail.Fields("Nett_Cost").Value, "##,##0")
        'MSC
        Rpt_Print_Media_Quotation.Sections("Section1").Controls("Lbl_MSC_" & Object_Index).Caption = Format(Rs_Print_detail.Fields("Media_Sptv_Charge").Value + Rs_Print_detail.Fields("Bonus").Value, "##,##0")
        'Others
        Rpt_Print_Media_Quotation.Sections("Section1").Controls("Lbl_Other_" & Object_Index).Caption = Format(Rs_Print_detail.Fields("Other_Charge").Value, "##,##0")
        'total Lintas
        Rpt_Print_Media_Quotation.Sections("Section1").Controls("Lbl_Total_Lintas_" & Object_Index).Caption = Format(Rs_Print_detail.Fields("Total_Lintas").Value, "##,##0")
        'Job_No CA
        Rpt_Print_Media_Quotation.Sections("Section1").Controls("Lbl_Job_No_CA_" & Object_Index).Caption = Rs_Print_detail.Fields("Job_Number_Agency").Value
        'MSC CA
        Rpt_Print_Media_Quotation.Sections("Section1").Controls("Lbl_Msc_CA_" & Object_Index).Caption = Format(Rs_Print_detail.Fields("Agency_Charge").Value, "##,##0")
        'Grand Total
        Rpt_Print_Media_Quotation.Sections("Section1").Controls("Lbl_Grand_Total_" & Object_Index).Caption = Format(Rs_Print_detail.Fields("Grand_Total").Value, "##,##0")
        'line 1,2,3
        Rpt_Print_Media_Quotation.Sections("Section1").Controls("Line" & Object_Index).Visible = True
        '7,8,9
        Rpt_Print_Media_Quotation.Sections("Section1").Controls("Line" & 6 + Object_Index).Visible = True
        '10,11,12
        Rpt_Print_Media_Quotation.Sections("Section1").Controls("Line" & 9 + Object_Index).Visible = True
        '13,14,15
        Rpt_Print_Media_Quotation.Sections("Section1").Controls("Line" & 12 + Object_Index).Visible = True
        '16,17,18
        Rpt_Print_Media_Quotation.Sections("Section1").Controls("Line" & 15 + Object_Index).Visible = True
        
        Object_Index = Object_Index + 1
        
        Rs_Print_detail.MoveNext
    Wend
    
    'Remark
    Rpt_Print_Media_Quotation.Sections("Section1").Controls("Lbl_Remark").Caption = IIf(IsNull(Rs_Quot_PR("Remarks").Value), "", Rs_Quot_PR("Remarks").Value)
        
    'Show Report
    Rpt_Print_Media_Quotation.show 1
    
    Me.MousePointer = vbNormal
    
    Rs_Print_detail.Close
    rs_client_name.Close
    Rs_Quot_PR.Close
End Sub

Private Sub Show_Quot_OT_CN(IB_ID As String, Brand As String)
'************************************************************
' Unit/Module Name  : Show_IB_OT_CN
' Function          : show quotation other
'                   :
' Public Variables  :
' Used Tables       :
' Date              : 11.10.05
' Created By        : Diyah
' Last Update/By    :
'************************************************************
    Dim Rs_Quot_Other As New ADODB.Recordset
    Dim rsTemp As New ADODB.Recordset
    Dim Rs_Quot_Detail As New ADODB.Recordset
    
    Me.MousePointer = vbHourglass
    
    'Header
    '=====================================
    strSql = "SELECT * FROM IB_Other_Quotation where IB_ID='" & IB_ID & "'"
    'Rs_Quot_Other.Close
    Rs_Quot_Other.Open strSql, ConnERP, adOpenStatic, adLockReadOnly
    
    'client
    strSql = " Select C.Client_Name from Client C inner join Brand B on B.Client_Code = C.Client_Code where Brand_Name = '" & Brand & "'"
    rsTemp.Open strSql, ConnERP, adOpenStatic, adLockReadOnly

    Crpt.Reset
    Crpt.ReportFileName = strReport_Dir & "\Other\Other_Quotation.RPT"
    Crpt.Formulas(2) = "Plan_No = '" & Rs_Quot_Other("Media_plan_No").Value & "'"
    Crpt.Formulas(3) = "IB_ID = '" & IB_ID & "'"
    Crpt.Formulas(4) = "Dated = '" & Format(Rs_Quot_Other("Date").Value, "MMMM dd.yy") & "'"
    Crpt.Formulas(0) = "Client = '" & Trim(rsTemp("Client_Name")) & "'"
    Crpt.Formulas(1) = "Brand = '" & Brand & "'"
  
  ' ========================================= Detail =========================
    strSql = "SELECT * FROM IB_other_Quotation_detail  WHERE  ib_id ='" & IB_ID & "' ORDER BY Month"
    Rs_Quot_Detail.Open strSql, ConnERP, adOpenStatic, adLockReadOnly
    
    With Rs_Quot_Detail
        If .EOF = False Then
            Crpt.Formulas(5) = "Month_1 = '" & Get_Month_Name(.Fields("Month").Value) & "'"
            Crpt.Formulas(6) = "Job_No_1 = '" & .Fields("Job_id").Value & "'"
            Crpt.Formulas(7) = "Nett_Cost_1 = '" & .Fields("Nett_Cost") & "'"
            Crpt.Formulas(8) = "MSC_1 = '" & .Fields("Media_Sptv_Charge").Value & "'"
            Crpt.Formulas(9) = "Other_1 = '" & .Fields("Other_Charge").Value & "'"
            Crpt.Formulas(10) = "Total_Lintas_1 = '" & .Fields("Total_Lintas").Value & "'"
            Crpt.Formulas(11) = "Job_No_Club_1 = '" & .Fields("Job_number_Agency").Value & "'"
            Crpt.Formulas(12) = "MSC_Club_1 = '" & .Fields("Agency_Charge").Value & "'"
            Crpt.Formulas(13) = "Grand_Total_1 = '" & .Fields("Grand_Total").Value & "'"
            .MoveNext
        Else
            Crpt.Formulas(5) = "Month_1 = ''"
            Crpt.Formulas(6) = "Job_No_1 = ''"
            Crpt.Formulas(7) = "Nett_Cost_1 = "
            Crpt.Formulas(8) = "MSC_1 = "
            Crpt.Formulas(9) = "Other_1 = "
            Crpt.Formulas(10) = "Total_Lintas_1 = "
            Crpt.Formulas(11) = "Job_No_Club_1 ="
            Crpt.Formulas(12) = "MSC_Club_1 = "
            Crpt.Formulas(13) = "Grand_Total_1 = "
        End If
        
        If .EOF = False Then
            Crpt.Formulas(14) = "Month_2 = '" & Get_Month_Name(.Fields("Month").Value) & "'"
            Crpt.Formulas(15) = "Job_No_2 = '" & .Fields("Job_id").Value & "'"
            Crpt.Formulas(16) = "Nett_Cost_2 = '" & .Fields("Nett_Cost") & "'"
            Crpt.Formulas(17) = "MSC_2 = '" & .Fields("Media_Sptv_Charge").Value & "'"
            Crpt.Formulas(18) = "Other_2 = '" & .Fields("Other_Charge").Value & "'"
            Crpt.Formulas(19) = "Total_Lintas_2 = '" & .Fields("Total_Lintas").Value & "'"
            Crpt.Formulas(20) = "Job_No_Club_2 = '" & .Fields("Job_number_Agency").Value & "'"
            Crpt.Formulas(21) = "MSC_Club_2 = '" & .Fields("Agency_Charge").Value & "'"
            Crpt.Formulas(22) = "Grand_Total_2 = '" & .Fields("Grand_Total").Value & "'"
            .MoveNext
        Else
            Crpt.Formulas(14) = "Month_2 = ''"
            Crpt.Formulas(15) = "Job_No_2 = ''"
            Crpt.Formulas(16) = "Nett_Cost_2 = "
            Crpt.Formulas(17) = "MSC_2 = "
            Crpt.Formulas(18) = "Other_2 ="
            Crpt.Formulas(19) = "Total_Lintas_2 ="
            Crpt.Formulas(20) = "Job_No_Club_2 = "
            Crpt.Formulas(21) = "MSC_Club_2 = "
            Crpt.Formulas(22) = "Grand_Total_2 = "
        End If
        
        If .EOF = False Then
            Crpt.Formulas(23) = "Month_3 = '" & Get_Month_Name(.Fields("Month").Value) & "'"
            Crpt.Formulas(24) = "Job_No_3 = '" & .Fields("Job_id").Value & "'"
            Crpt.Formulas(25) = "Nett_Cost_3 = '" & .Fields("Nett_Cost") & "'"
            Crpt.Formulas(26) = "MSC_3 = '" & .Fields("Media_Sptv_Charge").Value & "'"
            Crpt.Formulas(27) = "Other_3 = '" & .Fields("Other_Charge").Value & "'"
            Crpt.Formulas(28) = "Total_Lintas_3 = '" & .Fields("Total_Lintas").Value & "'"
            Crpt.Formulas(29) = "Job_No_Club_3 = '" & .Fields("Job_number_Agency").Value & "'"
            Crpt.Formulas(30) = "MSC_Club_3 = '" & .Fields("Agency_Charge").Value & "'"
            Crpt.Formulas(31) = "Grand_Total_3 = '" & .Fields("Grand_Total").Value & "'"
            .MoveNext
        Else
            Crpt.Formulas(23) = "Month_3 = ''"
            Crpt.Formulas(24) = "Job_No_3 = ''"
            Crpt.Formulas(25) = "Nett_Cost_3 = "
            Crpt.Formulas(26) = "MSC_3 = "
            Crpt.Formulas(27) = "Other_3 ="
            Crpt.Formulas(28) = "Total_Lintas_3 ="
            Crpt.Formulas(29) = "Job_No_Club_3 = "
            Crpt.Formulas(30) = "MSC_Club_3 = "
            Crpt.Formulas(31) = "Grand_Total_3 = "
        End If
    End With
    
    Crpt.SelectionFormula = "{IB_Other_Quotation.IB_ID} = '" & IB_ID & "'"
    Crpt.WindowShowPrintSetupBtn = True
    Crpt.WindowState = crptMaximized
    Crpt.Connect = "DSN=" & Server_Name & ";UID=" & Login_User & ";PWD=" & Login_Password & ";DSQ=" & strDatabase_Name
    Crpt.Action = 1
    
    Me.MousePointer = vbNormal
    
    Rs_Quot_Detail.Close
    Rs_Quot_Other.Close
    rsTemp.Close
End Sub

