Attribute VB_Name = "mod_Media"
Option Explicit
Private Declare Function GetComputerName Lib "kernel32" Alias "GetComputerNameA" (ByVal lpBuffer As String, nSize As Long) As Long

Type Pubic_Variable
    AppID As String
    Server_Name As String
    DBLogin_User As String
    DBLogin_Password As String
    Database_Name As String
    Country As String
    Computer_Name As String
    ConnMDB As New ADODB.Connection
    ConnERP As New ADODB.Connection
    APPLICATION_NAME As String
    Company_Code As String
    Company_Name As String
    Company_Address As String
    Company_Logo As Variant
    Company_BackGround As Variant
    code As String
    Default As String
    Central_Service_User As Boolean
    Login_User As String
    Login_Password As String
    Login_FullName As String
    LogDivisionCode As String
    LogUserDisable As String
    LogMobile_Number As String
    LogPhone_Number As String
    LastLogin As String
    lastLogout As String
    Is_SkipMultiSesion As Boolean
    rs_Date As New ADODB.Recordset
    strPathPO As String
    strPaperType As String
    strPaperTypePOTV As String
    strPaperTypeCOTV As String
    strPaperTypeCO As String
    intPOCOCopyNumber As String
    Report_Dir As String
    '-----------------------------------------
    StrCRMaddress  As String
    SMTP_Server  As String
    Str_Email_From  As String
    Str_Email_From_Name  As String
    Str_MP_Email_Subject  As String
    Str_MP_Email_Template_File  As String
    Str_Quotation_Email_Subject  As String
    Str_Quotation_Email_Template_File  As String
    Str_BC_Email_Subject  As String
    Str_BC_Email_Template_File  As String
    isAdminUser As Boolean
    isGroupHead As Boolean
End Type
'************* VAT *******************
Type Tax
    Vat As Long
End Type

Public Taxes As Tax
Public VarPub As Pubic_Variable
Public Is_SkipMultiSesion As Boolean

Public Sub Main()
'************************************************
' Procedure         : Main
' Function          : Bagian untuk pertama kali Membuka Aplikasi
'                     menjalankan LoadParameter
'                     menjalankan SetConnection
'                     menjalankan SetClient
'                     mengambil nilai strComputerName

' Input Parameter   : -
' Output Parameter  : -
' Programmer By     : -
' Date Update       : Jan-2016
' Update By         : Tedi/ Kreatif
'************************************************

    Dim rec_Temp As New ADODB.Recordset
    Dim Rs_Company As New ADODB.Recordset
    'Set conTemp = New ADODB.Connection
    
    '*******************************************************************************
    '* Load Parameter SQL Server Connection
    '*******************************************************************************
    LoadParameter
    SetERPConnection
    SetVariableCompany
'    strComputerName = conTemp.Properties.Item("workstation id")
'
'    'Get Date server
'    If Not recDate.State = adStateOpen Then
'        recDate.CursorLocation = adUseServer
'        recDate.Open "Server_Date", connERP, adOpenStatic, adLockReadOnly, adCmdStoredProc
'    End If
'
'    'Report Dir
'    str_SQLT = "SELECT * FROM App_Media_Report_Path"
'    recTemplate.Open str_SQLT, connERP, adOpenStatic, adLockReadOnly
'    If recTemplate.EOF Then
'        MsgBox "Parameter Report not found", vbCritical, "Match_IT"
'        End
'    Else
'        strReport_Dir = Trim(recTemplate.Fields("Report_path").Value)
'    End If
'    recTemplate.Close
'    Set recTemplate = Nothing
'
'    blnLogin = False
    Frm_Login.Show 1
'    If blnLogin = True Then mdi_Match_IT.Show Else EndApplication

    Exit Sub
    
errMain:

    If Err.Number = 3021 Then
        MsgBox "Please fill out the Information of your company on menu Configuration on Tenggo Main Launcher", vbExclamation, VarPub.APPLICATION_NAME
    Else
        MsgBox Err.Number & " - " & Err.Description, , vbExclamation, VarPub.APPLICATION_NAME
    End If
    
End Sub

Public Sub LoadParameter()
'************************************************
' Procedure         : LoadParameter
' Function          : Mengisi parameter
' Input Parameter   : --
' Output Parameter  : - strServer_Name      --> untuk nama server
'                     - strLogin_User       --> untuk nama User Login
'                     - strLogin_Password   --> untuk nama Password Login
'                     - strDatabase_Name    --> untuk nama Nama Database ERP
'                     - strDatabase_Name    --> untuk nama server
'                     - strEnvCode          --> untuk nama server
'                     - strPar_Arianna_Base_Dir    --> untuk Arianna_Working_Dir
'                     - strPar_Arianna_Working_Dir --> untuk Arianna_Working_Dir
'                     - strPar_Arianna_Env_Path    --> untuk Arianna_Env_Path
'                     - strDatabase_Name_ERP       --> untuk Database_Name_ERP
'                     - connERP                    --> untuk nama connectin ERP
'                     - connMacthIT                --> untuk nama connectin Matcth IT
' Programmer By     : -
' Date Update       : Apr-2015
' Update By         : Tedi/ Kreatif
'************************************************
    
    Dim rec_Temp As New ADODB.Recordset
    VarPub.ConnMDB.ConnectionTimeout = 0
    VarPub.ConnMDB.CommandTimeout = 0
    VarPub.ConnMDB.Open "Provider=Microsoft.Jet.OLEDB.3.51;Persist Security Info=False;User ID=Admin;Data Source= '" & "..\Tenggo_Parameter.mdb';Mode=Share Deny None;Extended Properties=';COUNTRY=0;CP=1252;LANGID=0x0409';Jet OLEDB:System database='';Jet OLEDB:Registry Path='';Jet OLEDB:Database Password='smart';Jet OLEDB:Global Partial Bulk Ops=2"
    
    '=========== Get Server ===================
    rec_Temp.Open "SELECT * FROM Parameter ", VarPub.ConnMDB, adOpenStatic, adLockReadOnly
    'OurClient.Default = "None"
    While Not rec_Temp.EOF
        Select Case Trim(rec_Temp.Fields("item").Value)
        Case "Server_Name"
            VarPub.Server_Name = IIf(IsNull(rec_Temp("Value").Value), "", rec_Temp("Value").Value)
        Case "User"
            VarPub.DBLogin_User = IIf(IsNull(rec_Temp("Value").Value), "", rec_Temp("Value").Value)
        Case "Password"
            VarPub.DBLogin_Password = IIf(IsNull(rec_Temp("Value").Value), "", rec_Temp("Value").Value)
        Case "Database"
            VarPub.Database_Name = IIf(IsNull(rec_Temp("Value").Value), "", rec_Temp("Value").Value)
        Case "Default"
            VarPub.AppID = IIf(IsNull(rec_Temp("Value").Value), "", rec_Temp("Value").Value)
        End Select
        rec_Temp.MoveNext
    Wend
    
    CloseRecordset rec_Temp
    VarPub.ConnMDB.Close
    Dim dbAccess As Database
    Dim rsDao As DAO.Recordset
    Dim Ero As ADODB.Error
'    Set connMacthIT = New ADODB.Connection
'    Set connERP = New ADODB.Connection
        
    '--------------------- Read Parameter ---------------------------------
    'Set dbAccess = OpenDatabase(App.Path & "\ERP_Parameter2.mdb", False, True, ";pwd=smart")
    
    
   '=========== Get App ID ===================
'    Set rsDao = dbAccess.OpenRecordset("SELECT * FROM parameter WHERE isActive=Yes")
'    If rsDao.EOF Then
'        MsgBox "APP ID Parameter not found", vbCritical, "Error"
'        rsDao.Close
'        End
'    Else
'        strEnvCode = Trim$(rsDao.Fields("Environment_Name").Value)
'        strServer_Name = Trim$(rsDao.Fields("Server").Value)
'        strLogin_Name = Trim$(rsDao.Fields("user").Value)
'        strLogin_Password = Trim$(rsDao.Fields("password").Value)
'        strDatabase_Name = Trim$(rsDao.Fields("Database").Value)
'        strPar_Arianna_Base_Dir = Trim$(rsDao.Fields("Arianna_Base_Dir").Value)
'        strPar_Arianna_Working_Dir = Trim$(rsDao.Fields("Arianna_Working_Dir").Value)
'        strPar_Arianna_Env_Path = Trim$(rsDao.Fields("Arianna_Env_Path").Value)
'        strDatabase_Name_ERP = Trim$(rsDao.Fields("Database_ERP").Value)
'
'        If Trim$(rsDao.Fields("IsWeekbyWeek").Value) = "Yes" Then
'            blnIsWeekbyWeek = True
'        Else
'            blnIsWeekbyWeek = False
'        End If
'    End If
'    rsDao.Close
  
'    strPar_Query_Folder = strPar_Arianna_Base_Dir & "Input\Idn\"
'    strPar_Output_Folder = strPar_Arianna_Base_Dir & "Output2\"

'    blnIsAdminUser = False
'    blnIsWeekCommencing = False
'
'    connMacthIT.ConnectionTimeout = 0
'    connMacthIT.CommandTimeout = 0
'    connMacthIT.Open "Provider=SQLOLEDB.1;Persist Security Info=False;User ID=" & Trim(strLogin_Name) & ";password=" & Trim(strLogin_Password) & ";Initial Catalog=" & Trim(strDatabase_Name) & ";Data Source=" & Trim(strServer_Name)
'
'    'ERP Conn
'    connERP.ConnectionTimeout = 300
'    connERP.CommandTimeout = 300
'    connERP.Open "Provider=SQLOLEDB.1;Persist Security Info=False;User ID=" & Trim(strLogin_Name) & ";password=" & Trim(strLogin_Password) & ";Initial Catalog=" & Trim(strDatabase_Name_ERP) & ";Data Source=" & Trim(strServer_Name)
'    connERP.CursorLocation = adUseClient
'    Dim strBuffer As String
'    Dim lngLen As Long
'
'    'get komputer name
'    strBuffer = Space(255)
'    lngLen = Len(strBuffer)
'
'    If CBool(GetComputerName(strBuffer, lngLen)) Then
'        strComputerName = Left$(strBuffer, lngLen)
'    Else
'        strComputerName = "Local"
'    End If
         
        
    Exit Sub
    
Lable:
'    For Each Ero In connMacthIT.Errors
'        If Ero.NativeError = 17 Then
'            MsgBox "Server Not Found." & vbCrLf & "Please Contact IT Department.", vbExclamation, APPLICATION_NAME
'        End If
'    Next Ero
'    MsgBox Err.Description
'    End

End Sub

Public Sub CloseRecordset(ByRef paRecordset As ADODB.Recordset)
'*****************************************
'Procedure Name     : CloseRecordset
'Procedure Function : Close recordset yg aman dari error.
'Input Parameter    : ---
'Output Parameter   : ---
'Date               : 07-Apr-2015/{73 64 6B}.
'*****************************************
    On Local Error Resume Next

    If Not (paRecordset Is Nothing) Then
        If paRecordset.State = adStateOpen Then paRecordset.Close
    End If

    Set paRecordset = Nothing

    On Local Error GoTo 0
End Sub

Public Sub SetERPConnection()
'************************************************
' Procedure         : SetERPConnection
' Function          : Menyambung Ke Database ERP
' Input Parameter   : -
' Output Parameter  : -
' Programmer By     : -
' Date Update       : Jan-2016
' Update By         : Tedi/ Kreatif
'************************************************
    
    On Error GoTo errERPConnection
    Dim strConn As String
    
    If VarPub.ConnERP.State = 0 Then
        
        VarPub.ConnERP.ConnectionTimeout = 300
        VarPub.ConnERP.CommandTimeout = 300
        
        strConn = "Provider=SQLOLEDB.1;"
        strConn = strConn & "Persist Security Info=False;"
        strConn = strConn & "User ID=" & VarPub.DBLogin_User & ";"
        strConn = strConn & "password=" & VarPub.DBLogin_Password & ";"
        strConn = strConn & "Initial Catalog=" & VarPub.Database_Name & ";"
        strConn = strConn & "Data Source=" & VarPub.Server_Name
        
        VarPub.ConnERP.Open strConn
    
    End If
    
    Exit Sub

errERPConnection:
    
    MsgBox Err.Description, vbExclamation, VarPub.APPLICATION_NAME

End Sub

Public Sub SetVariableCompany()
'************************************************
' Procedure         : SetVariableCompany
' Function          : Set Client Variable
' Input Parameter   : -
' OutPut Parameter  : -
' Programmer By     : Tedi
' Date Update       : Jan-2016
' Update By         : Tedi / Kreatif
'************************************************

    Dim recTemp As New ADODB.Recordset
    
    If VarPub.AppID = "None" Then
        recTemp.Open "SELECT * FROM  company WHERE 1=1", VarPub.ConnERP, adOpenForwardOnly, adLockReadOnly
    Else
        recTemp.Open "SELECT * FROM  company WHERE Company_Code='" & VarPub.AppID & "'", VarPub.ConnERP, adOpenForwardOnly, adLockReadOnly
    End If
    
    VarPub.Company_Code = recTemp!Company_Code
    VarPub.Company_Name = recTemp!Company_Name
    VarPub.Company_Logo = recTemp!Logo
    VarPub.Company_BackGround = recTemp!timesheet_bg

    CloseRecordset recTemp

End Sub

Public Sub GetPictureFromDB(ByRef objImage As Object, ByVal varCode As Variant)
'************************************************
' Procedure         : GetPictureFromDB
' Function          : Tampilkan Gambar image BLOB, untuk ditampilkan ke object Image Box
' Input Parameter   : Memasukan termasuk refernsinya - objImage , varCode - Nilai Gambar dalam variant
' Programmer By     : Tedi
' Date Update       : Jan-2016
' Update By         : Tedi / Kreatif
'************************************************
    
    Dim stStream As ADODB.Stream
    Set stStream = New ADODB.Stream
    
    stStream.Type = adTypeBinary
    stStream.Open
    stStream.Write varCode
    stStream.SaveToFile App.Path & "\Temp.jpg", adSaveCreateOverWrite
    objImage.Picture = LoadPicture(App.Path & "\Temp.jpg")

End Sub

Public Function Clear_String(ByVal sBr As String) As String
'********************************************************************************
'Procedure Name     : Clear_String
'Procedure Function : ---
'Input Parameter    : ---
'Output Parameter   : ---
'Date               : 2/25/2016
'LastUpdate/By      : Tedi / Kreatif
'********************************************************************************
    
    Dim lPos As Long
    Dim sBl As String
    
    If Len(sBr) = 0 Then Exit Function
    lPos = InStr(sBr, Chr$(39))
    While lPos <> 0
      sBl = sBl & Left$(sBr, lPos) & Chr$(39)
      sBr = Right$(sBr, Len(sBr) - lPos)
      lPos = InStr(sBr, Chr$(39))
    Wend
    Clear_String = sBl & sBr

End Function

Public Sub Load_Taxes_Data()
'************************************************
' Procedure         : Load_Taxes_Data()
' Function          : To Load all taxes data
' Date              : 11/13/2000
' Parameter Input   :
' Parameter Output  :
' Last Update/By    :
'************************************************
    Dim RS_Tax As New ADODB.Recordset
    Dim TxtSQl As String

    TxtSQl = "SELECT * FROM tax_catalog"
    RS_Tax.Open TxtSQl, VarPub.ConnERP, adOpenStatic, adLockReadOnly
    With RS_Tax
        Do While .EOF = False
            If .Fields(0).Value = "VAT" Then
                Taxes.Vat = .Fields(2).Value
            End If
            .MoveNext
        Loop
    End With
    Set RS_Tax = Nothing
End Sub

Public Sub EndApplication()
'************************************************
' Procedure         : EndApplication
' Function          : Proses saat Menutup Aplikasi,
'                     menutupsemua Koneksi
' Input Parameter   : -
' Output Parameter  : -
' Programmer By     : -
' Date Update       : Apr-2015
' Update By         : Tedi/ Kreatif
'************************************************
    
    'CloseRecordset recDate
    If Not VarPub.ConnERP Is Nothing Then
        If VarPub.ConnERP.State = 1 Then VarPub.ConnERP.Close: Set VarPub.ConnERP = Nothing
    End If
    End

End Sub

Public Sub RemoveMenus(frm As Form, remove_close As Boolean)
    Dim hMenu As Long
    
    ' Get the form's system menu handle.
    hMenu = GetSystemMenu(frm.hwnd, False)
    If remove_close Then DeleteMenu hMenu, 6, MF_BYPOSITION
End Sub

Public Function Get_Company(Company_Code As String) As String

    Dim recCompany As New ADODB.Recordset
    Dim strQuery As String
    
    strQuery = "SELECT Company_Name from Company WHERE Company_Code='" & Company_Code & "'"
    recCompany.Open strQuery, VarPub.ConnERP, adOpenStatic, adLockReadOnly, adCmdText
    If recCompany.BOF And recCompany.EOF Then
        Get_Company = ""
    Else
        Get_Company = recCompany(0).Value
    End If
    
    recCompany.Close
    Set recCompany = Nothing
End Function

Public Sub Cancel_Brief_Id_Media_Running(Brief_Id As String)

'************************************************
' Procedure         : Cancel_Brief_Id_Media_Running
' Function          : To Cancel Brief Id Media Running
' Date              : 01/10/2001
' Parameter Input   : Brief Id
' Parameter Output  :
' Last Update/By    :
'************************************************
'Cancel Job Number

    On Error Resume Next
    Conn.Execute "cancel_Brief_Id_Media '" & Brief_Id & "'"
        
End Sub
