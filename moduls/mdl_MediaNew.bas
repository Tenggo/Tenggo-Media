Attribute VB_Name = "mdl_MediaNew"
Public Const strTitleConfirm = "Confirmation"
Public Type RGBColor
    'Decimal Value'
    r As Single
    G As Single
    b As Single
    'Hex Value
    HR As String
    HG As String
    HB As String
End Type
'
Public blnStatusPassword As Boolean

Public Const strTitleInfo = "Information"
Public Const strTitleMissingInfo = "Missing Information"
Public Const strTitleExclamation = "Warning"
Public Const strMsgAccessDenied = "Access Denied"
Public Const strMsgSaveDataDone = "Data has been saved"
Public Const strMsgUpdateDataDone = "Data has been updated"
Public Const strMsgDeleteDataDone = "Data has been deleted"
Public Const strMsgDeleteConfirm = "Are you sure want to delete selected record.?"
Public Const strMsgApproved = "Approved.."
Public Const strMsgInvalidPassword = "Invalid Password.!"

'dw - to Find a criteria----------
Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As _
Long, ByVal wMsg As Long, ByVal wParam As Integer, ByVal lParam As Any) As Long

Public Const LB_FINDSTRING = &H18F
'----------------------------------

Public Function ColorConstToRGB(ByVal color As Long) As RGBColor

    Dim ColorRGB As RGBColor
    Dim ColorInHex As String

    ColorInHex = Right$("000000" & Hex$(color), 6)
    
    ColorRGB.HB = Left$(ColorInHex, 2)
    ColorRGB.HG = Mid$(ColorInHex, 3, 2)
    ColorRGB.HR = Right$(ColorInHex, 2)
    
    ColorRGB.b = CDec("&H" & ColorRGB.HB)
    ColorRGB.G = CDec("&H" & ColorRGB.HG)
    ColorRGB.r = CDec("&H" & ColorRGB.HR)
    
    ColorConstToRGB = ColorRGB

End Function

Public Sub VGradient(ByVal pic As Object, ByVal Color1 As OLE_COLOR, Color2 As OLE_COLOR, ByVal start_y, ByVal end_y, ByVal start_x, ByVal end_x)
    Dim hgt As Single
    Dim wid As Single
    
    Dim ClrInRGB As RGBColor
    
    Dim r As Single
    Dim G As Single
    Dim b As Single
    
    Dim dr As Single
    Dim dg As Single
    Dim db As Single
    
    Dim Y As Single
    
    Dim start_r As Integer, end_r As Integer
    Dim start_g As Integer, end_g As Integer
    Dim start_b As Integer, end_b As Integer
    
    ClrInRGB = ColorConstToRGB(Color1)
    
    start_r = ClrInRGB.r
    start_g = ClrInRGB.G
    start_b = ClrInRGB.b
    
    ClrInRGB = ColorConstToRGB(Color2)
    
    end_r = ClrInRGB.r
    end_g = ClrInRGB.G
    end_b = ClrInRGB.b
    
    wid = end_x - start_x
    hgt = end_y - start_y
    dr = (end_r - start_r) / hgt
    dg = (end_g - start_g) / hgt
    db = (end_b - start_b) / hgt
    r = start_r
    G = start_g
    b = start_b
    For Y = start_y To end_y
        pic.Line (start_x, Y)-(end_x, Y), RGB(r, G, b)
        r = r + dr
        G = G + dg
        b = b + db
    Next Y
End Sub

Public Function LoadBrand(objBrand As Object, strUserName As String, strPosition As String) As String
    Dim strQuery As String
    Dim recBrand As New ADODB.Recordset
    
    strQuery = "SELECT * FROM brand WHERE brand_code IN (SELECT brand_code FROM Media_Security_Catalog WHERE User_name='" & strLogin_User & "' AND position IN (" & strPosition & ") and Valid_until > getdate())"
    strQuery = strQuery & " ORDER BY Brand_Code"
    recBrand.Open strQuery, ConnERP, adOpenStatic, adLockReadOnly
    
    With recBrand
        If .EOF = False Then
            Do While .EOF = False
                objBrand.AddItem .Fields("Brand_Code") & " - " & .Fields("Brand_Name")
                .MoveNext
            Loop
        Else
            MsgBox "You don't have access to brand", vbCritical, strTitleExclamation
        End If
    End With

    Set recBrand = Nothing
    
    If objBrand.ListCount <> 0 Then
        objBrand.ListIndex = 0
    End If
End Function

Public Function LoadBrandVariant(objBrandVariant As Object, strBrand As String) As String
    Dim strQuery As String
    Dim recBrandVariant As New ADODB.Recordset

    strQuery = "SELECT * FROM Brand_Variant WHERE Brand_Code='" & strBrand & "'"
    recBrandVariant.Open strQuery, ConnERP, adOpenStatic, adLockReadOnly
    objBrandVariant.Clear

    With recBrandVariant
        If .EOF = False Then
            Do While .EOF = False
                objBrandVariant.AddItem .Fields("Brand_Variant_Code") & " --> " & .Fields("Brand_Variant_Name")
                .MoveNext
            Loop
        End If
    End With
    
    If objBrandVariant.ListCount <> 0 Then
        objBrandVariant.ListIndex = 0
    End If
    
    Set recBrandVariant = Nothing
End Function

Public Function LoadSecondaryTarget(objSecondaryTarget As Object) As String
    Dim strQuery As String
    Dim recSecondaryTarget As New ADODB.Recordset

    strQuery = "SELECT * FROM Cluster"
    recSecondaryTarget.Open strQuery, ConnERP, adOpenStatic, adLockReadOnly
    objSecondaryTarget.Clear

    With recSecondaryTarget
        If .EOF = False Then
            Do While Not .EOF And Not .BOF
                objSecondaryTarget.AddItem recSecondaryTarget(0).Value & " - " & recSecondaryTarget(1).Value
                'objSecondaryTarget.ItemData(objSecondaryTarget.NewIndex) = recSecondaryTarget(0).Value
                .MoveNext
            Loop
        End If
    End With
    
    If objSecondaryTarget.ListCount <> 0 Then
        objSecondaryTarget.ListIndex = 0
    End If
    
    Set recCluster = Nothing
End Function

Public Function blnIsDeFlag(ByVal strBrandCode As String) As Boolean
    Dim strQuery As String
    Dim recBrand As New ADODB.Recordset
    
    strQuery = "select de_flag from brand where brand_code='" & Trim(strBrandCode) & "'"
    recBrand.Open strQuery, ConnERP, adOpenForwardOnly, adLockReadOnly, adCmdText
    
    If Not recBrand.BOF And Not recBrand.EOF Then
        If recBrand(0).Value = 1 Then
            blnIsDeFlag = True
        Else
            blnIsDeFlag = False
        End If
    End If
    
    recBrand.Close
    Set recBrand = Nothing
    
End Function

Public Function strGetPrintMediaQuotationNo(strBrandCode As String, IntYear As Integer) As String
    Dim strQuery As String
    Dim recMax As New ADODB.Recordset
    Dim strMediaCode As String
    Dim dblNo As Double
    Dim strNomor As String
    Dim recCompare As New ADODB.Recordset
    Dim strNewMQNumber As String
    '***************************************
    'prosedur untuk mendapatkan MQ ID Print
    '***************************************
    
    strMediaCode = "015"
    
    'cari dulu di reusable WHERE year and brand
    strQuery = "SELECT * FROM Reuseable_IB_ID_mq_Print WHERE year=" & IntYear & " and brand_code = '" & strBrandCode & "' ORDER BY ib_id "
    recMax.Open strQuery, ConnERP, adOpenStatic, adLockReadOnly, adCmdText
    
    'jika ada StrNewMQNumber = record di reusable
    If recMax.RecordCount > 0 Then
    
        strNewMQNumber = recMax("ib_id")
        
        'Compare Data reusable dengan last
        'cari di last WHERE year and brand
        strQuery = "SELECT last_number FROM Last_IB_ID_mq_Print WHERE year='" & IntYear & "' and brand_code='" & strBrandCode & "'"
        recCompare.Open strQuery, ConnERP, adOpenStatic, adLockReadOnly, adCmdText
        If recCompare.RecordCount = 0 Then
                'jika tidak ada record insert ke last
                strQuery = "INSERT INTO Last_IB_ID_mq_Print(brand_code, year, last_number)"
                strQuery = strQuery & " VALUES ( '" & strBrandCode & "', '" & IntYear & "'," & CInt(Right(Trim(strNewMQNumber), 2)) & ")"
                ConnERP.Execute strQuery
        Else
                'jika ada banding apakah last number-nya yang direusable lebih besar
                If recCompare(0) < CInt(Right(Trim(strNewMQNumber), 2)) Then
                    strQuery = "UPDATE Last_IB_ID_mq_Print set last_number = " & CInt(Right(Trim(strNewMQNumber), 2)) & " WHERE year=" & IntYear & " and brand_code='" & strBrandCode & "'"
                    ConnERP.Execute strQuery
                End If
        End If
        
        ' delete di reusable
        strQuery = "DELETE FROM Reuseable_IB_ID_mq_Print WHERE year=" & recMax("year") & " and Brand_Code= '" & recMax("Brand_Code") & "' and IB_ID= '" & recMax("IB_ID") & "'"
        ConnERP.Execute strQuery
        
        recCompare.Close
        Set recCompare = Nothing
        
        recMax.Close
        Set recMax = Nothing
        
        strGetPrintMediaQuotationNo = strNewMQNumber
        
        Exit Function
    End If
    
    recMax.Close
    Set recMax = Nothing
                    
    'cari nomor WHERE year and brand code
    strQuery = "SELECT last_number FROM Last_IB_ID_mq_Print WHERE year=" & IntYear & " and brand_code='" & strBrandCode & "'"
    recMax.Open strQuery, ConnERP, adOpenStatic, adLockReadOnly
    If recMax.RecordCount > 0 Then
         dblNo = recMax(0).Value + 1
    Else
         dblNo = 1
    End If
    recMax.Close
    Set recMax = Nothing
    
    If Len(Trim(dblNo)) = 1 Then strNomor = "0" + CStr(dblNo) Else strNomor = CStr(dblNo)
        
    strNewMQNumber = strBrandCode & "." & Trim(strMediaCode) & "." & Right(IntYear, 2) & Trim(strNomor)
     
'DELETE di last
    strQuery = "DELETE Last_IB_ID_mq_Print WHERE brand_code='" & strBrandCode & "' and  year=" & IntYear
    ConnERP.Execute strQuery
    
'untuk insert ke last ib print
    strQuery = "INSERT INTO Last_IB_ID_mq_Print(brand_code, year, last_number)"
    strQuery = strQuery & " VALUES ( '" & strBrandCode & "', " & IntYear & "," & CInt(Right(Trim(strNewMQNumber), 2)) & ")"
    ConnERP.Execute strQuery
    
    strGetPrintMediaQuotationNo = strNewMQNumber
       
End Function
