Attribute VB_Name = "mdl_Media"
'<CSCC>
'********************************************************************************
'Submodul Name      : mdl_Media
'Submodul Function  : {MemberName}
'Used Table         : -
'Used SP            : -
'Procedure/Function : -
'Programmer Name    : Tedi
'Date               : 3/8/2016-2:00:31 AM
'Last Update By     : Tedi
'Date Update        : 3/8/2016-2:00:31 AM
'Log Update/By      : -
'********************************************************************************
'</CSCC>

Option Explicit
Private Declare Function GetComputerName Lib "kernel32" Alias "GetComputerNameA" (ByVal lpBuffer As String, nSize As Long) As Long
Public strQuery As String
'***************** Private API Function *********************
Private Declare Function DeleteMenu Lib "user32" (ByVal hMenu As Long, ByVal nPosition As Long, ByVal wFlags As Long) As Long
Private Declare Function GetSystemMenu Lib "user32" (ByVal hwnd As Long, ByVal bRevert As Long) As Long
Private Const MF_BYPOSITION = &H400&

'****** Declare Variable For TV Schedule *************
Public Rs_Temp_Program As New ADODB.Recordset 'Program
Public Rs_Temp_Program_Rate As New ADODB.Recordset 'Rate
Public Rs_Temp_Program_SRI As New ADODB.Recordset 'Program
Public Rs_Temp_Program_Rate_SRI As New ADODB.Recordset 'Rate
'=============================================================
Public Selected_Quotation_No As String
Public Selected_Quotation_Type As String
Public Selected_Brand_Name As String
Public Selected_Brand_code As String
'=============================================================
Public MP_Mail_Plan_No As String
Public MP_Mail_Approve_By As String
Public MP_Mail_Approve_Date As String
'=============================================================
Public StrCRMaddress As String
Public SMTP_Server As String
Public Str_Email_From As String
Public Str_Email_From_Name As String
Public Str_MP_Email_Subject As String
Public Str_MP_Email_Template_File As String
Public Str_Quotation_Email_Subject As String
Public Str_Quotation_Email_Template_File As String
Public Str_BC_Email_Subject As String
Public Str_BC_Email_Template_File As String

Public Str_Year_TV_Prg_Rate As String
Public Str_Month_TV_Prg_Rate As String
Public Central_Service_User As Boolean

'=================== Global Var Server Date ==========================
Public recDate As New ADODB.Recordset

Public ConnMDB As New ADODB.Connection
Public ConnERP As New ADODB.Connection

Public strAppID As String
Public strServerName As String
Public strDBLogin_User As String
Public strDBLogin_Password As String
Public strDatabase_Name As String
Public strCountry As String
Public strComputer_Name As String
Public strApplication_Name As String
'=================== Global Var Company ==========================
Public strCompany_Code As String
Public strCompany_Name As String
Public strCompany_Address As String
Public strCompany_UsrEmail As String
Public strCompany_PassEmail As String
Public strCompany_isGmail As String
Public strCompany_isPort As String
Public strCompany_Logo_Link As String
Public vntCompany_Logo As Variant
Public vntCompany_BackGround As Variant
'=================== Global Var Company ==========================
Public strCode As String
Public strDefault As String
Public strLogin_User As String
Public strLogin_Password As String
Public strLogin_FullName As String
Public strLogDivisionCode As String
Public strLogUserDisable As String
Public strLogMobile_Number As String
Public strLogPhone_Number As String
Public strLastLogin As String
Public strLastLogout As String
Public strPathPO As String
Public strPaperType As String
Public strPaperTypePOTV As String
Public strPaperTypeCOTV As String
Public strPaperTypeCO As String
Public strPOCOCopyNumber As String
Public strReport_Dir As String

'Public StrCRMaddress As String

Public strSMTP_Server As String
Public strEmail_From As String
Public strEmail_From_Name As String

Public strMP_Email_Subject As String
Public strMP_Email_Template_File As String
Public strQuotation_Email_Subject As String
Public strQuotation_Email_Template_File As String
Public strBC_Email_Subject As String
Public strBC_Email_Template_File As String

'Public StrCRMaddress As String
Public intTotalWeek As Integer

Public blnCentral_Service_User As Boolean
Public blnSkipMultiSesion As Boolean
Public blnIsAdminUser As Boolean
Public blnIsGroupHead As Boolean

Type Brand_Information
    ULI As Boolean
    PSC As Double
    PSC_Nett_Flag As Boolean
    MSC As Double
    MSC_Nett_Flag As Boolean
    Club_Agency_SC As Double
    Club_Agency_Flag As Boolean
    Media_Agency_Bonus As Double
    Media_Agency_Bonus_Nett_Flag As Boolean
    Vat As Double
End Type

Public Brand_Info As Brand_Information

Type Brand_Information_Print
    Brand_Name As String
    Client_Name As String
    ULI As Boolean
    PSC As Double
    PSC_Nett_Flag As Boolean
    MSC As Double
    MSC_Nett_Flag As Boolean
    Club_Agency_SC As Double
    Club_Agency_Flag As Boolean
    Media_Agency_Bonus As Double
    Media_Agency_Bonus_Nett_Flag As Boolean
    DE_Flag As Boolean
    Vat As Double
End Type

Public Brand_InFo_Print As Brand_Information_Print
'************* VAT *******************
Type Tax
    Vat As Long
End Type
Type WeekCommencing
    WeekYear As Integer
    WeekMonth As Integer
    WeekCommencingDate As String
End Type

Type ProgramRatePeriod
    StationCode As String
    TheMonth As Integer
    TheYear As Integer
End Type
Type Media_Tv
    Str_Year_TV_Prg_Rate As String
    Str_Month_TV_Prg_Rate As String
End Type
Public mediaTV As Media_Tv
Public ArrWeekCommencing() As WeekCommencing
Public ArrProgramRatePeriod() As ProgramRatePeriod

Public Taxes As Tax

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
    Frm_Login.show 1
'    If blnLogin = True Then mdi_Match_IT.Show Else EndApplication

    Exit Sub
    
errMain:

    If Err.Number = 3021 Then
        MsgBox "Please fill out the Information of your company on menu Configuration on Tenggo Main Launcher", vbExclamation, strApplication_Name
    Else
        MsgBox Err.Number & " - " & Err.Description, , vbExclamation, strApplication_Name
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
    ConnMDB.ConnectionTimeout = 0
    ConnMDB.CommandTimeout = 0
    ConnMDB.Open "Provider=Microsoft.Jet.OLEDB.3.51;Persist Security Info=False;User ID=Admin;Data Source= '" & "..\Tenggo_Parameter.mdb';Mode=Share Deny None;Extended Properties=';COUNTRY=0;CP=1252;LANGID=0x0409';Jet OLEDB:System database='';Jet OLEDB:Registry Path='';Jet OLEDB:Database Password='smart';Jet OLEDB:Global Partial Bulk Ops=2"
    
    '=========== Get Server ===================
    rec_Temp.Open "SELECT * FROM Parameter ", ConnMDB, adOpenStatic, adLockReadOnly
    'OurClient.Default = "None"
    While Not rec_Temp.EOF
        Select Case Trim(rec_Temp.Fields("item").Value)
        Case "Server_Name"
            strServerName = IIf(IsNull(rec_Temp("Value").Value), "", rec_Temp("Value").Value)
        Case "User"
            strDBLogin_User = IIf(IsNull(rec_Temp("Value").Value), "", rec_Temp("Value").Value)
        Case "Password"
            strDBLogin_Password = IIf(IsNull(rec_Temp("Value").Value), "", rec_Temp("Value").Value)
        Case "Database"
            strDatabase_Name = IIf(IsNull(rec_Temp("Value").Value), "", rec_Temp("Value").Value)
        Case "Default"
            strAppID = IIf(IsNull(rec_Temp("Value").Value), "", rec_Temp("Value").Value)
        End Select
        rec_Temp.MoveNext
    Wend
    
    CloseRecordset rec_Temp
    ConnMDB.Close
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
'    'ERP ConnERP
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
'            MsgBox "Server Not Found." & vbCrLf & "Please Contact IT Department.", vbExclamation, strApplication_Name
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
    
    If ConnERP.State = 0 Then
        
        ConnERP.ConnectionTimeout = 300
        ConnERP.CommandTimeout = 300
        
        strConn = "Provider=SQLOLEDB.1;"
        strConn = strConn & "Persist Security Info=False;"
        strConn = strConn & "User ID=" & strDBLogin_User & ";"
        strConn = strConn & "password=" & strDBLogin_Password & ";"
        strConn = strConn & "Initial Catalog=" & strDatabase_Name & ";"
        strConn = strConn & "Data Source=" & strServerName
        
        ConnERP.Open strConn
    
    End If
    
    Exit Sub

errERPConnection:
    
    MsgBox Err.Description, vbExclamation, strApplication_Name

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
    
    If strAppID = "None" Then
        recTemp.Open "SELECT * FROM  company WHERE 1=1", ConnERP, adOpenForwardOnly, adLockReadOnly
    Else
        recTemp.Open "SELECT * FROM  company WHERE Company_Code='" & strAppID & "'", ConnERP, adOpenForwardOnly, adLockReadOnly
    End If
    
    strCompany_Code = recTemp!Company_Code
    strCompany_Name = recTemp!Company_Name
    vntCompany_Logo = recTemp!Logo
    vntCompany_BackGround = recTemp!timesheet_bg
    strCompany_UsrEmail = Trim(recTemp!Email)
    strCompany_PassEmail = Trim(recTemp!Password)
    strCompany_isPort = recTemp!Port
    strCompany_isGmail = recTemp!isgmail
'    strCompany_Logo_Link = Trim(recTemp!link_image_logo)
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
    RS_Tax.Open TxtSQl, ConnERP, adOpenStatic, adLockReadOnly
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
    If Not ConnERP Is Nothing Then
        If ConnERP.State = 1 Then ConnERP.Close: Set ConnERP = Nothing
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
'<CSCM>
'********************************************************************************
'Procedure Name     : Get_Company
'Procedure Function : ---
'Input Parameter    : ---
'Output Parameter   : ---
'Date               : 3/8/2016
'LastUpdate/By      : Tedi / Kreatif
'********************************************************************************
'</CSCM>

    Dim recCompany As New ADODB.Recordset
    Dim strQuery As String
    
    strQuery = "SELECT Company_Name from Company WHERE Company_Code='" & Company_Code & "'"
    recCompany.Open strQuery, ConnERP, adOpenStatic, adLockReadOnly, adCmdText
    If recCompany.BOF And recCompany.EOF Then
        Get_Company = ""
    Else
        Get_Company = recCompany(0).Value
    End If
    
    recCompany.Close
    Set recCompany = Nothing
End Function

Public Function Get_Brand_Name(Brand_code As String) As String
'************************************************
' Procedure         : Get_Brand_Name
' Function          : To Get Brand Name
' Date              : 10/18/2000
' Parameter Input   : Brand_Code
' Parameter Output  : Get_Brand_Name
' Last Update/By    :
'************************************************
    Dim rs_brand As New ADODB.Recordset
    Dim strSql As String
    
    strSql = "SELECT Brand_Name from Brand WHERE Brand_Code='" & Brand_code & "'"
    rs_brand.Open strSql, ConnERP, adOpenForwardOnly, adLockReadOnly, adCmdText
    If rs_brand.BOF And rs_brand.EOF Then
        Get_Brand_Name = ""
    Else
        Get_Brand_Name = rs_brand(0).Value
    End If
    
    rs_brand.Close
    Set rs_brand = Nothing
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
    ConnERP.Execute "cancel_Brief_Id_Media '" & Brief_Id & "'"
        
End Sub

Public Function IsFeeAlreadyEntered(strBrandCode As String, IntYear As Integer) As Boolean

    Dim recFee As New ADODB.Recordset
    Dim strQuery As String
    Dim blnJanFeeFound As Boolean
    Dim blnFebFeeFound As Boolean
    Dim blnMarFeeFound As Boolean
    Dim blnAprFeeFound As Boolean
    Dim blnMayFeeFound As Boolean
    Dim blnJunFeeFound As Boolean
    Dim blnJulFeeFound As Boolean
    Dim blnAugFeeFound As Boolean
    Dim blnSepFeeFound As Boolean
    Dim blnOctFeeFound As Boolean
    Dim blnNovFeeFound As Boolean
    Dim blnDecFeeFound As Boolean
    
    blnJanFeeFound = False
    blnFebFeeFound = False
    blnMarFeeFound = False
    blnAprFeeFound = False
    blnMayFeeFound = False
    blnJunFeeFound = False
    blnJulFeeFound = False
    blnAugFeeFound = False
    blnSepFeeFound = False
    blnOctFeeFound = False
    blnNovFeeFound = False
    blnDecFeeFound = False
    
    
    strQuery = "SELECT * FROM Brand_Fee WHERE Year=" & IntYear
    strQuery = strQuery & " AND Brand_Code='" & strBrandCode & "'"
    
    recFee.Open strQuery, ConnERP, adOpenForwardOnly, adLockReadOnly, adCmdText
    
    Do While Not recFee.EOF
        Select Case recFee("Month")
            Case 1
                blnJanFeeFound = True
            Case 2
                blnFebFeeFound = True
            Case 3
                blnMarFeeFound = True
            Case 4
                blnAprFeeFound = True
            Case 5
                blnMayFeeFound = True
            Case 6
                blnJunFeeFound = True
            Case 7
                blnJulFeeFound = True
            Case 8
                blnAugFeeFound = True
            Case 9
                blnSepFeeFound = True
            Case 10
                blnOctFeeFound = True
            Case 11
                blnNovFeeFound = True
            Case 12
                blnDecFeeFound = True
        End Select
        
        recFee.MoveNext
    Loop
    
    If blnJanFeeFound And _
        blnFebFeeFound And _
        blnMarFeeFound And _
        blnAprFeeFound And _
        blnMayFeeFound And _
        blnJunFeeFound And _
        blnJulFeeFound And _
        blnAugFeeFound And _
        blnSepFeeFound And _
        blnOctFeeFound And _
        blnNovFeeFound And _
        blnDecFeeFound _
    Then
        IsFeeAlreadyEntered = True
    Else
        IsFeeAlreadyEntered = False
    End If
    
    recFee.Close
    Set recFee = Nothing
    
End Function

Public Sub SleepX()
'************************************************
' Procedure         : SleepX
' Function          : Merefresh Object Visual.
' Created By        : tedi
' Date              : 12-Jan-2016
'************************************************
    
    Dim dt_waitTill As Date
    dt_waitTill = Now() + TimeValue("00:00:01")
    
    While Now() < dt_waitTill
        DoEvents
    Wend

End Sub

Public Function Clear_Enter(ByVal sBr As String) As String
    Dim lPos As Long
    Dim Str_Temp As String
    
    'If Null Exit
    If Len(sBr) = 0 Then Exit Function
    
    'Ger Enter Char
    lPos = InStr(sBr, Chr$(13))
    
    While lPos <> 0
      Str_Temp = Str_Temp & Left$(sBr, lPos - 1) & " "
      sBr = Right$(sBr, Len(sBr) - (lPos + 1))
      lPos = InStr(sBr, Chr$(13))
    Wend
    
    Clear_Enter = Str_Temp & sBr
End Function

Public Function Is_Special_Brand(Brand_code As String) As Boolean
    Dim strSql As String
    Dim rs_brand As New ADODB.Recordset
    
    strSql = "SELECT Client.Special_Client_Flag from client,Brand WHERE Client.Client_Code=Brand.Client_Code AND Brand_Code='" & Brand_code & "'"
    rs_brand.Open strSql, ConnERP, adOpenForwardOnly, adLockReadOnly, adCmdText
    
    If Not rs_brand.BOF And Not rs_brand.EOF Then
        If rs_brand(0).Value = 1 Then
            Is_Special_Brand = True
        Else
            Is_Special_Brand = False
        End If
    End If
    
    rs_brand.Close
    Set rs_brand = Nothing
    
End Function

Public Function Get_Percent_CASC(Brand_code As String) As Double
    
    Dim recCASC As New ADODB.Recordset
    
    recCASC.Open "SELECT Club_Agency_SC FROM Brand WHERE Brand_code='" & Brand_code & "'", ConnERP, adOpenForwardOnly, adLockReadOnly, adCmdText
    
    If Not recCASC.BOF And Not recCASC.EOF Then
        If Not IsNull(recCASC.Fields("Club_Agency_SC").Value) Then
            Get_Percent_CASC = recCASC.Fields("Club_Agency_SC").Value
        Else
            Get_Percent_CASC = 0
        End If
    End If
    
    recCASC.Close
    Set recCASC = Nothing
    
End Function

Public Function Get_Month_Name(month As Integer) As String
    Select Case month
        Case 1
            Get_Month_Name = "January"
        Case 2
            Get_Month_Name = "February"
        Case 3
            Get_Month_Name = "March"
        Case 4
            Get_Month_Name = "April"
        Case 5
            Get_Month_Name = "May"
        Case 6
            Get_Month_Name = "June"
        Case 7
             Get_Month_Name = "July"
        Case 8
            Get_Month_Name = "August"
        Case 9
             Get_Month_Name = "September"
        Case 10
            Get_Month_Name = "October"
        Case 11
            Get_Month_Name = "November"
        Case 12
            Get_Month_Name = "December"
    End Select
End Function

Public Function Get_Month_Number(month As String) As Integer
    Select Case LCase(Trim(month))
        Case "january"
            Get_Month_Number = 1
        Case "february"
            Get_Month_Number = 2
        Case "march"
            Get_Month_Number = 3
        Case "april"
            Get_Month_Number = 4
        Case "may"
            Get_Month_Number = 5
        Case "june"
            Get_Month_Number = 6
        Case "july"
             Get_Month_Number = 7
        Case "august"
            Get_Month_Number = 8
        Case "september"
             Get_Month_Number = 9
        Case "october"
            Get_Month_Number = 10
        Case "november"
            Get_Month_Number = 11
        Case "december"
            Get_Month_Number = 12
        Case Else
            ConnERP.Execute "INSERT INTO Log_Month VALUES('" & Screen.ActiveForm.Name & ":" & month & "')"
    End Select
End Function

Public Function Get_Percent_Netto_Value() As Double
    Dim recPercent_Netto_Value As New ADODB.Recordset
    
    recPercent_Netto_Value.Open "Select Value From Percent_Netto_Value", ConnERP
    If Not recPercent_Netto_Value.EOF Then
        Get_Percent_Netto_Value = recPercent_Netto_Value.Fields("Value").Value
    Else
        Get_Percent_Netto_Value = 0
    End If
    recPercent_Netto_Value.Close
    Set recPercent_Netto_Value = Nothing
    
End Function

Public Function Get_Percent_MSC(Brand_code As String) As Double
    
    Dim rec_Msc As New ADODB.Recordset
    
    rec_Msc.Open "SELECT MSC FROM Brand WHERE Brand_code='" & Brand_code & "'", ConnERP, adOpenForwardOnly, adLockReadOnly, adCmdText
    If Not rec_Msc.BOF And Not rec_Msc.EOF Then
         Get_Percent_MSC = rec_Msc.Fields("MSC").Value
    End If
    rec_Msc.Close
    Set rec_Msc = Nothing
    
End Function

Public Function Get_Percent_Media_Agency_Bonus(Brand_code As String) As Double
    
    Dim recMAB As New ADODB.Recordset
    
    recMAB.Open "SELECT Media_Agency_Bonus FROM Brand WHERE Brand_code='" & Brand_code & "'", ConnERP, adOpenForwardOnly, adLockReadOnly, adCmdText
    If Not recMAB.BOF And Not recMAB.EOF Then
         Get_Percent_Media_Agency_Bonus = recMAB.Fields("Media_Agency_Bonus").Value
    End If
    recMAB.Close
    Set recMAB = Nothing
    
End Function

Public Function Get_MSC_Flag(Brand_code As String) As Integer
'Declare
 Dim recMsc As New ADODB.Recordset
 
    recMsc.CursorLocation = adUseClient
    recMsc.Open "SELECT MSC_Nett_Flag FROM Brand WHERE Brand_code='" & Brand_code & "'", ConnERP, adOpenForwardOnly, adLockReadOnly, adCmdText
    
    
    If Not recMsc.BOF And Not recMsc.EOF Then
        Get_MSC_Flag = recMsc.Fields("MSC_Nett_Flag").Value
    Else
        Get_MSC_Flag = 0
    End If
    
    recMsc.Close
    Set recMsc = Nothing
    
End Function

Public Function Get_Bonus_Fee_Flag(Brand_code As String) As Integer
'Declare

 Dim recBonus_Fee As New ADODB.Recordset

 
    recBonus_Fee.Open "SELECT Media_Agency_Bonus_Nett_Flag FROM Brand WHERE Brand_code='" & Brand_code & "'", ConnERP, adOpenForwardOnly, adLockReadOnly, adCmdText
    If Not recBonus_Fee.BOF And Not recBonus_Fee.EOF Then
        Get_Bonus_Fee_Flag = recBonus_Fee.Fields("Media_Agency_Bonus_Nett_Flag").Value
    Else
        Get_Bonus_Fee_Flag = 0
    End If
    
    
    recBonus_Fee.Close
    Set recBonus_Fee = Nothing
    
End Function

Public Function LoadYear(objYear As Object) As Integer
    Dim IntYear As Integer
    Dim intBegYear As Integer
    Dim intEndYear As Integer
    
    intBegYear = 2002
    intEndYear = 2015
    For IntYear = intBegYear To intEndYear
        objYear.AddItem IntYear
    Next IntYear
    
    objYear.ListIndex = objYear.ListCount - 1
End Function

Public Function IsValidAccess(strUserName As String, strPosition As String, strBrandCode As String) As Boolean
    Dim rsTemp As New ADODB.Recordset
    Dim strSql As String
    
    strSql = ""
    strSql = "SELECT brand_code FROM Media_Security_Catalog "
    strSql = strSql & "WHERE User_name='" & strUserName & "' "
    strSql = strSql & "AND position='" & strPosition & "' "
    strSql = strSql & "AND Brand_Code='" & strBrandCode & "' "
    strSql = strSql & "AND Valid_until > getdate()"
        
    rsTemp.Open strSql, ConnERP, adOpenForwardOnly, adLockReadOnly, adCmdText
    If rsTemp.EOF Then
        IsValidAccess = False
    Else
        IsValidAccess = True
    End If
    
    rsTemp.Close
    Set rsTemp = Nothing
    
End Function

