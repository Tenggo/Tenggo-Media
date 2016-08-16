VERSION 5.00
Begin VB.Form Frm_Radio_MQ_IB 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Implementation Brief ID"
   ClientHeight    =   3120
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   2685
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3120
   ScaleWidth      =   2685
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Height          =   2535
      Left            =   120
      TabIndex        =   2
      Top             =   0
      Width           =   2415
      Begin VB.ListBox Lst_IB 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1860
         ItemData        =   "Frm_Radio_MQ_IB.frx":0000
         Left            =   240
         List            =   "Frm_Radio_MQ_IB.frx":0002
         Style           =   1  'Checkbox
         TabIndex        =   3
         Top             =   255
         Width           =   1965
      End
   End
   Begin VB.CommandButton Cmd_Cancel 
      Caption         =   "Cancel"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1440
      TabIndex        =   1
      Top             =   2640
      Width           =   1095
   End
   Begin VB.CommandButton Cmd_OK 
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
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   2640
      Width           =   1095
   End
End
Attribute VB_Name = "Frm_Radio_MQ_IB"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
''*************************************************************
'Nama Form          : Frm_Radio_MQ_IB
'Fungsi Form        : memilih IB
'Programer          : joko
'cerated date       : 4/Nop/01
'Last Update/By     :
'*************************************************************
Private Sub cmd_cancel_Click()
    Frm_Radio_Media_Quot.Cbo_Month_MQ.Visible = False
    Unload Me
End Sub

Private Sub Cmd_Ok_Click()
    Dim Pos_Index As Integer
    Dim Ada_Flag As Boolean
    Dim rs As New ADODB.Recordset
    Dim TxtSQl As String
    Dim Str_IB As String
    Dim Jum_Loop As Integer
    
    Ada_Flag = False
    For Pos_Index = 0 To Lst_IB.ListCount - 1
        Lst_IB.ListIndex = Pos_Index
        If Lst_IB.Selected(Pos_Index) = True Then
            Ada_Flag = True
            Exit For
        End If
    Next Pos_Index
    
    If Ada_Flag = True Then
        If Frm_Radio_Media_Quot.Flex_Quot.col = 1 Then
            Frm_Radio_Media_Quot.IB_ID_1 = ""
        ElseIf Frm_Radio_Media_Quot.Flex_Quot.col = 3 Then
            Frm_Radio_Media_Quot.IB_ID_2 = ""
        ElseIf Frm_Radio_Media_Quot.Flex_Quot.col = 5 Then
            Frm_Radio_Media_Quot.IB_ID_3 = ""
        End If
        
        Rem Get IB ID
        Str_IB = " ( "
        For Pos_Index = 0 To Lst_IB.ListCount - 1
            Lst_IB.ListIndex = Pos_Index
            If Lst_IB.Selected(Pos_Index) = True Then
                Jum_Loop = Jum_Loop + 1
                
                If Frm_Radio_Media_Quot.Flex_Quot.col = 1 Then
                    Frm_Radio_Media_Quot.IB_ID_1 = Frm_Radio_Media_Quot.IB_ID_1 & Trim(Lst_IB.Text)
                End If
                If Frm_Radio_Media_Quot.Flex_Quot.col = 3 Then
                    Frm_Radio_Media_Quot.IB_ID_2 = Frm_Radio_Media_Quot.IB_ID_2 & Trim(Lst_IB.Text)
                End If
                If Frm_Radio_Media_Quot.Flex_Quot.col = 5 Then
                    Frm_Radio_Media_Quot.IB_ID_3 = Frm_Radio_Media_Quot.IB_ID_3 & Trim(Lst_IB.Text)
                End If
                If Jum_Loop = 1 Then
                    Str_IB = Str_IB & " ib_id ='" & Trim(Lst_IB.Text) & "'"
                ElseIf Jum_Loop > 1 Then
                    Str_IB = Str_IB & " or ib_id ='" & Trim(Lst_IB.Text) & "'"
                End If
                
            End If
        Next Pos_Index
        Str_IB = Str_IB & " ) "
        
        If Trim(Frm_Radio_Media_Quot.Cbo_Month_MQ.Text) <> "-None-" Then
            Frm_Radio_Media_Quot.Flex_Quot.TextMatrix(Frm_Radio_Media_Quot.Flex_Quot.Row, Frm_Radio_Media_Quot.Flex_Quot.col) = Trim(Frm_Radio_Media_Quot.Cbo_Month_MQ.Text)
            
             Rem Get Total Budget dari IB yang dipilih
             TxtSQl = "select sum(budget) as Budget from IB_radio_Plan where "
             TxtSQl = TxtSQl & "  Month =" & Get_Month_Number(Frm_Radio_Media_Quot.Cbo_Month_MQ.Text) & " and "
             'TxtSQL = TxtSQL & " where Month =" & Val(Frm_Radio_Media_Quot.Cbo_Month_MQ.Text) & " and Year =" & Val(Frm_Radio_Media_Quot.Cbo_Year.Text)
             TxtSQl = TxtSQl & Str_IB
             
             With rs
                 .Open TxtSQl, ConnERP, adOpenKeyset, adLockReadOnly
                 If Not .EOF Then
                 
                     Frm_Radio_Media_Quot.Flex_Quot.TextMatrix(11, Frm_Radio_Media_Quot.Flex_Quot.col) = Format(.Fields("Budget"), "##,##0")
                 End If
             End With
        End If
       
        Rem Generate Job_ID
            Dim New_Job As String
            Dim New_Month As String
            Dim RS_Code As New ADODB.Recordset
            Dim Radio_Code As String
            
            TxtSQl = "select * from Media_Type where Media_Type_Name ='Radio Media Induk'"
            RS_Code.Open TxtSQl, ConnERP, adOpenStatic, adLockReadOnly
            With RS_Code
                If .EOF = False Then
                    Radio_Code = .Fields(0)
                End If
            End With
            Set RS_Code = Nothing
            New_Month = Get_Month_Number(Frm_Radio_Media_Quot.Cbo_Month_MQ.Text)
            If Len(New_Month) = 1 Then
                New_Month = Trim("0" & Trim(str(Get_Month_Number(Frm_Radio_Media_Quot.Cbo_Month_MQ.Text))))
            Else
                New_Month = Get_Month_Number(Frm_Radio_Media_Quot.Cbo_Month_MQ.Text)
            End If
            
            New_Job = Left(Frm_Radio_Media_Quot.Cbo_Brand.Text, 4) & "." & Trim(Radio_Code) & "." & Right(Trim(Frm_Radio_Media_Quot.Cbo_Year.Text), 2) & New_Month
            Frm_Radio_Media_Quot.Flex_Quot.TextMatrix(1, Frm_Radio_Media_Quot.Flex_Quot.col) = New_Job
        
        'Frm_Radio_Media_Quot.Cbo_Month_MQ.RemoveItem (Get_Month_Number(Frm_Radio_Media_Quot.Cbo_Month_MQ.Text))
        Unload Me
    Else
        MsgBox "Please select Implementation Brief", vbCritical, StrCompany
    End If
    
    Frm_Radio_Media_Quot.Cbo_Month_MQ.Visible = False
End Sub

Private Sub Form_Load()
    Dim rs As New ADODB.Recordset
    Dim TxtSQl As String
    
    TxtSQl = "SELECT IB_ID FROM IB_radio_Plan "
    TxtSQl = TxtSQl & " WHERE Month =" & Get_Month_Number(Frm_Radio_Media_Quot.Cbo_Month_MQ.Text) & " AND Year =" & Val(Frm_Radio_Media_Quot.Cbo_Year.Text)
    TxtSQl = TxtSQl & " AND LEFT(IB_ID,4) ='" & Left(Trim(Frm_Radio_Media_Quot.Cbo_Brand.Text), 4) & "'"
    TxtSQl = TxtSQl & " AND IB_ID IN (SELECT IB_ID FROM Ib_Radio WHERE Approved_Flag=1 AND LEFT(IB_ID,4)='" & Left(Trim(Frm_Radio_Media_Quot.Cbo_Brand.Text), 4) & "' AND YEAR=" & Val(Frm_Radio_Media_Quot.Cbo_Year.Text) & " AND Status=1)"
    
    rs.Open TxtSQl, ConnERP, adOpenKeyset, adLockReadOnly

    With rs
        Do While Not .EOF
            Lst_IB.AddItem Trim(.Fields("IB_ID"))
            .MoveNext
        Loop
    End With
End Sub
