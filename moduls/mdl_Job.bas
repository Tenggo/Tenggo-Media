Attribute VB_Name = "mdl_Job"
Option Explicit
''
Public Sub Add_Current_User_Job(ByVal Activity_Id As Double, ByVal strMessage_From As String, ByVal Brief_Id As String, ByVal Job_No As String, ByVal PO_Number As String, Optional Job_Type As String, Optional ByVal StrDescription As String, Optional ByVal Send_To_Name As String, Optional ByVal strMessage As String, Optional ByRef intTask_Type As Integer)
'************************************************
' Procedure         :
' Function          : Add data job pop up menu
' Date              : 13.10.2005
' Parameter Input   : brief_id,job no,activity_id,po_number,
' Parameter Output  :
' Last Update/By    : Diyah
'************************************************
    Dim Rs_AM As New ADODB.Recordset
    Dim rs_Job_ID As New ADODB.Recordset
    Dim strSql As String
    Dim intJob_Id As Long
    Dim varUser_Name As Variant
    Dim intIndex_User_Name As Integer
   
   '===== Generate New Job ID(Tambahkan ke tabel Current_Job) ========================
   strSql = "INSERT INTO Current_Job(Activity_Id,Brief_Id,Job_Number,Po_Number,Description,Date,Computer_Name,Message_From,Message,Task_Type,Update_Date,User_Name_Message_From)"
   strSql = strSql & "VALUES(" & Activity_Id & ",'" & Brief_Id & "','" & Job_No & "',"
   strSql = strSql & "'" & PO_Number & "','" & StrDescription & "', GetDate(),'" & strComputer_Name & "','" & strMessage_From & "','" & Clear_String(strMessage) & "'," & intTask_Type & ",GetDate(),'" & strLogin_User & "')"
    
   ConnERP.Execute strSql
   '=======================================================================================
   
   'Set task type to default (0 pemberitahuan)
   intTask_Type = 0
   
    If rs_Job_ID.State = adStateOpen Then
        rs_Job_ID.Close
    End If
    
    'Get New Job ID (Ambil  Job_Id yang terakhir dari tabel Current_Job)
    '=======================================================================
    rs_Job_ID.Open "SELECT Max(Id_job) as Id_Job FROM Current_Job WHERE Computer_Name='" & strComputer_Name & "'", ConnERP, adOpenStatic, adLockReadOnly, adCmdText
    
    intJob_Id = IIf(IsNull(rs_Job_ID!Id_Job), 0, rs_Job_ID!Id_Job)
    
    If rs_Job_ID.State = adStateOpen Then
       rs_Job_ID.Close
    End If
    
    Set rs_Job_ID = Nothing
    '===================== End Get New Job ID ===============================
    
    'Jika Subrutin ini dipanggil degan parameter Send_To_Name=""
    '=======================================================================
    If Send_To_Name = "" Then
        Select Case Activity_Id
            Case 1, 3, 4, 5, 6, 8, 9, 10
                strSql = "SELECT user_name from Media_security_catalog WHERE "
                strSql = strSql & " Brand_Code='" & Left(Brief_Id, 4) & "'"
                strSql = strSql & " and Valid_until > Getdate()"
                strSql = strSql & " and  Position ='Implementor'"
            Case 2, 7
                strSql = "SELECT user_name from Media_security_catalog WHERE "
                strSql = strSql & " Brand_Code='" & Left(Brief_Id, 4) & "'"
                strSql = strSql & " and Valid_until > Getdate()"
                strSql = strSql & " and  Position ='Buyer'"
        End Select
    
        If Rs_AM.State = adStateOpen Then
            Rs_AM.Close
            Set Rs_AM = Nothing
        End If
        
        Rs_AM.Open strSql, ConnERP, adOpenStatic, adLockReadOnly
        
        While Not Rs_AM.EOF And Not Rs_AM.BOF
            strSql = "INSERT INTO Current_User_Job (Id_Job,User_Name,Date,Update_date)"
            strSql = strSql & " VALUES (" & intJob_Id & ",'" & Rs_AM.Fields(0).Value & "',GetDate(),GetDate())"
            ConnERP.Execute strSql
            
            Rs_AM.MoveNext
        Wend
            
        Rs_AM.Close
        Set Rs_AM = Nothing
    
    Else '======= Untuk parameter Send_To_Name <>"" ==========
        
        varUser_Name = Split(Send_To_Name, ";")
        
        '==== Cek apakah user name yang akan dikirim lebih dari satu atau tidak ===
        If UBound(varUser_Name) = 0 Then
            'Send_To_Name=hanya 1 user
            'Jika cuma ada satu maka kirim datanya cuma sekali
            strSql = "INSERT INTO Current_User_Job (Id_job,User_Name,Date,Update_Date)"
            strSql = strSql & " VALUES (" & intJob_Id & ",'" & varUser_Name(0) & "',GetDate(),GetDate())"
            ConnERP.Execute strSql
        Else
            ''Send_To_Name > 1 User
            'Looping sebanyak jumlah user yang telah dipilih dan akan diberi pesan
            For intIndex_User_Name = LBound(varUser_Name) To UBound(varUser_Name)
                If varUser_Name(intIndex_User_Name) <> "" Then
                    strSql = "INSERT INTO Current_User_Job (Id_Job,User_Name,Date,Update_Date)"
                    strSql = strSql & " VALUES (" & intJob_Id & ",'" & varUser_Name(intIndex_User_Name) & "',GetDate(),GetDate())"
                    ConnERP.Execute strSql
                End If
            Next
        End If
    End If
End Sub

