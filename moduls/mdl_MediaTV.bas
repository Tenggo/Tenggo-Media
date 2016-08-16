Attribute VB_Name = "mdl_MediaTV"
Option Explicit
Public StrPrinterType As String
'11092001543PM'
Public Sub Refresh_TV_Schedule(Job_Id As String, Market As String)
    Dim strSql As String
    Dim Rs_TV_Schedule  As New ADODB.Recordset
    Dim Percent_Netto_Value As Double
    Dim Percent_MSC_Paid As Double
    Dim Percent_MSC_Bonus As Double
    '
    Dim Total_Spot_Paid As Integer
    Dim Total_Spot_Bonus As Integer
    Dim Total_Tarps_Paid As Double
    Dim Total_Tarps_Bonus As Double
    
    Dim Total_Spots_perWeek As Integer
    Dim Total_Tarps_Per_Week As Double
    
    Dim Total_Gross_Rate_Paid As Currency
    Dim Total_Gross_Paid_Paid As Currency
    Dim Total_Nett_Rate_Paid As Currency
    
    Dim Total_Gross_Rate_Bonus As Currency
    Dim Total_Gross_Paid_Bonus As Currency
    Dim Total_Nett_Rate_Bonus As Currency
    
    Dim Total_Gross_Value As Currency
    Dim Total_Netto_Value As Currency
    
    Dim MSC_For_Paid As Currency
    Dim MSC_For_Bonus As Currency
    Dim Total_MSC As Currency
    
    Dim DE As Currency
    Dim Total_Cost As Currency
    Dim Vat As Currency
    Dim Grand_Total As Currency
    Dim Local_Tax As Currency
    
    
    
'Initial Data
    Total_Spot_Paid = 0
    Total_Spot_Bonus = 0
    Total_Tarps_Paid = 0
    Total_Tarps_Bonus = 0
    
    Total_Spots_perWeek = 0
    Total_Tarps_Per_Week = 0
    
    Total_Gross_Rate_Paid = 0
    Total_Gross_Paid_Paid = 0
    Total_Nett_Rate_Paid = 0
         
    Total_Gross_Rate_Bonus = 0
    Total_Gross_Paid_Bonus = 0
    Total_Nett_Rate_Bonus = 0
    
    Total_Gross_Value = 0
    Total_Netto_Value = 0
    
    MSC_For_Paid = 0
    MSC_For_Bonus = 0
    Total_MSC = 0
    
    DE = 0
    Total_Cost = 0
    Vat = 0
    Grand_Total = 0
    Local_Tax = 0

'Get Data
    '====================================
    '   Load Percent Netto Value
    '====================================
    Percent_Netto_Value = Get_Percent_Netto_Value()
    
'Get Old Data
    strSql = "SELECT DE,Local_Tax From TV_Schedule WHERE "
    strSql = strSql & " Job_Id='" & Job_Id & "' AND Market='" & Market & "'"
    
    Rs_TV_Schedule.CursorLocation = adUseClient
    Rs_TV_Schedule.Open strSql, ConnERP, , , adCmdText
        Do While Not Rs_TV_Schedule.EOF
            DE = Rs_TV_Schedule.Fields("DE").Value
            Local_Tax = Rs_TV_Schedule.Fields("Local_Tax").Value
            Rs_TV_Schedule.MoveNext
        Loop
    Rs_TV_Schedule.Close
'End Get old Data
    
    strSql = "SELECT Paid_Flag,TVR,Spot,Gross_Rate,Gross_Paid,Nett_Rate FROM TV_Schedule_Program_Insertion "
    strSql = strSql & " WHERE Job_Id='" & Job_Id & "' AND Market='" & Market & "'"
    Rs_TV_Schedule.Open strSql, ConnERP, , , adCmdText
'Loop Per Program Insertion
    Do While Not Rs_TV_Schedule.EOF
        'Get Paid Data
        If Rs_TV_Schedule.Fields("Paid_Flag") = 1 Then
            Total_Spot_Paid = Total_Spot_Paid + Rs_TV_Schedule.Fields("Spot").Value
            Total_Tarps_Paid = Total_Tarps_Paid + (Rs_TV_Schedule.Fields("Spot").Value * Rs_TV_Schedule.Fields("TVR").Value)
            'If Rs_TV_Schedule.Fields("Spot").Value = 0 Then
            '    MsgBox "Sponsorship"
            'End If
            
            If Rs_TV_Schedule.Fields("Spot").Value = 0 Then
                Total_Gross_Rate_Paid = Total_Gross_Rate_Paid + (1 * Rs_TV_Schedule.Fields("Gross_Rate").Value)
                Total_Gross_Paid_Paid = Total_Gross_Paid_Paid + (1 * Rs_TV_Schedule.Fields("Gross_Paid").Value)
                Total_Nett_Rate_Paid = Total_Nett_Rate_Paid + (1 * Rs_TV_Schedule.Fields("Nett_Rate").Value)
            Else
                Total_Gross_Rate_Paid = Total_Gross_Rate_Paid + (Rs_TV_Schedule.Fields("Spot").Value * Rs_TV_Schedule.Fields("Gross_Rate").Value)
                Total_Gross_Paid_Paid = Total_Gross_Paid_Paid + (Rs_TV_Schedule.Fields("Spot").Value * Rs_TV_Schedule.Fields("Gross_Paid").Value)
                Total_Nett_Rate_Paid = Total_Nett_Rate_Paid + (Rs_TV_Schedule.Fields("Spot").Value * Rs_TV_Schedule.Fields("Nett_Rate").Value)
            End If
            
            
            
        'Get Bonus Data
        Else
            Total_Spot_Bonus = Total_Spot_Bonus + Rs_TV_Schedule.Fields("Spot").Value
            Total_Tarps_Bonus = Total_Tarps_Bonus + (Rs_TV_Schedule.Fields("Spot").Value * Rs_TV_Schedule.Fields("TVR").Value)
            
            Total_Gross_Rate_Bonus = Total_Gross_Rate_Bonus + (Rs_TV_Schedule.Fields("Spot").Value * Rs_TV_Schedule.Fields("Gross_Rate").Value)
            Total_Gross_Paid_Bonus = Total_Gross_Paid_Bonus + (Rs_TV_Schedule.Fields("Spot").Value * Rs_TV_Schedule.Fields("Gross_Paid").Value)
            Total_Nett_Rate_Bonus = Total_Nett_Rate_Bonus + (Rs_TV_Schedule.Fields("Spot").Value * Rs_TV_Schedule.Fields("Nett_Rate").Value)
            
        End If
        Rs_TV_Schedule.MoveNext
    Loop
    
    
    'Total
    Total_Gross_Value = Total_Gross_Rate_Paid + Total_Gross_Rate_Bonus
    Total_Netto_Value = Total_Gross_Value - (Total_Gross_Value * (Percent_Netto_Value / 100))
    
    'Get MSC
        'Get Percent MSC
        Percent_MSC_Paid = Get_Percent_MSC(Left(Job_Id, 4))
        Percent_MSC_Bonus = Get_Percent_Media_Agency_Bonus(Left(Job_Id, 4))
        
    'Paid
    Select Case Get_MSC_Flag(Left(Job_Id, 4))
        Case 0
            'On Gross Paid
            MSC_For_Paid = (Percent_MSC_Paid / 100) * Total_Gross_Paid_Paid
        Case 1
            'On Nett
            MSC_For_Paid = (Percent_MSC_Paid / 100) * Total_Nett_Rate_Paid
        Case 2
            'On Gross Rate
            MSC_For_Paid = (Percent_MSC_Paid / 100) * Total_Gross_Rate_Paid
        Case 3
            'On Gross Value
            MSC_For_Paid = (Percent_MSC_Paid / 100) * Total_Gross_Rate_Paid
        Case 4
            'On Nett Value
            MSC_For_Paid = (Percent_MSC_Paid / 100) * (Total_Gross_Rate_Paid - (Total_Gross_Rate_Paid * (Percent_Netto_Value / 100)))
    End Select
    'Bonus
    
    Select Case Get_Bonus_Fee_Flag(Left(Job_Id, 4))
    Case 0
        'On Bonus Gross Paid
        MSC_For_Bonus = (Percent_MSC_Bonus / 100) * Total_Gross_Paid_Bonus
    Case 1
        'On Bonus Nett
        MSC_For_Bonus = (Percent_MSC_Bonus / 100) * Total_Nett_Rate_Bonus
    Case 2
        'On Bonus Gross Rate
        MSC_For_Bonus = (Percent_MSC_Bonus / 100) * Total_Gross_Rate_Bonus
    Case 3
        'On Bonus Gross Value
        MSC_For_Bonus = (Percent_MSC_Bonus / 100) * Total_Gross_Rate_Bonus
    Case 4
        'On Bonus Nett Value
        MSC_For_Bonus = (Percent_MSC_Bonus / 100) * (Total_Gross_Rate_Bonus - (Total_Gross_Rate_Bonus * (Percent_Netto_Value / 100)))
    End Select
    
    'Total
    Total_MSC = MSC_For_Paid + MSC_For_Bonus
    Total_Cost = Total_Nett_Rate_Paid + Total_MSC + DE
    Vat = (Taxes.Vat / 100) * Total_Cost
    Grand_Total = Total_Cost + Vat
    
    'Close Recordset
    Rs_TV_Schedule.Close
'Update TV Schedule Header
    strSql = "UPDATE TV_Schedule SET "
    strSql = strSql & "Total_Spots=" & Total_Spot_Paid
    strSql = strSql & ",Total_TARPs=" & Total_Tarps_Paid
    strSql = strSql & ",Total_Spots_Bonus=" & Total_Spot_Bonus
    strSql = strSql & ",Total_Tarps_Bonus=" & Total_Tarps_Bonus
    strSql = strSql & ",Total_Gross_Value_Bonus=" & Total_Gross_Rate_Bonus
    strSql = strSql & ",Total_Netto_Bonus=" & Total_Nett_Rate_Bonus
    strSql = strSql & ",Total_Gross_Value=" & Total_Gross_Rate_Paid
    strSql = strSql & ",Total_Gross_Paid=" & Total_Gross_Paid_Paid
    strSql = strSql & ",Total_Netto_Value=" & Total_Netto_Value
    strSql = strSql & ",Total_Netto=" & Total_Nett_Rate_Paid
    strSql = strSql & ",Msc_For_Paid=" & MSC_For_Paid
    strSql = strSql & ",Msc_For_Bonus=" & MSC_For_Bonus
    strSql = strSql & ",Total_Msc=" & Total_MSC
    strSql = strSql & ",DE=" & DE
    strSql = strSql & ",Total_Cost=" & Total_Cost
    strSql = strSql & ",VAT=" & Vat
    strSql = strSql & ",Grand_Total=" & Grand_Total
    strSql = strSql & ",Local_Tax=" & Local_Tax
    
    strSql = strSql & " WHERE Job_Id='" & Job_Id & "' AND Market='" & Market & "'"
    
    ConnERP.Execute strSql

'Update TV_Schedule_Weekly_Total
    Dim Rs_Temp As New ADODB.Recordset
    
    strSql = "SELECT Month,Week_commencing,Total_Spots,Total_Tarps FROM TV_Schedule_Weekly_Total "
    strSql = strSql & " WHERE Job_Id='" & Job_Id & "' AND Market='" & Market & "'"
    Rs_TV_Schedule.CursorLocation = adUseServer
    Rs_TV_Schedule.Open strSql, ConnERP, adOpenKeyset, adLockOptimistic, adCmdText
    Do While Not Rs_TV_Schedule.EOF
        'Get Total Spot,Tarps Weekly
        strSql = ""
        strSql = "SELECT SUM(Spot) AS Total_Spots, SUM(Spot*TVR) AS Total_Tarps FROM TV_Schedule_Program_Insertion "
        strSql = strSql & " WHERE Job_Id='" & Job_Id & "' AND Market='" & Market & "'"
        strSql = strSql & " AND Month=" & Rs_TV_Schedule.Fields("Month").Value
        strSql = strSql & " AND Week_Commencing=" & Rs_TV_Schedule.Fields("Week_Commencing").Value
        Rs_Temp.CursorLocation = adUseClient
        Rs_Temp.Open strSql, ConnERP, , , adCmdText
        If Not Rs_Temp.EditMode Then
            If Not IsNull(Rs_Temp.Fields(0).Value) Then
                Rs_TV_Schedule.Fields("Total_Spots").Value = Rs_Temp.Fields("Total_Spots").Value
                Rs_TV_Schedule.Fields("Total_Tarps").Value = Rs_Temp.Fields("Total_Tarps").Value
                Rs_TV_Schedule.Update
            End If
        End If
        Rs_Temp.Close
        'Next Record
        
        Rs_TV_Schedule.MoveNext
    Loop
    
    Set Rs_Temp = Nothing
    Rs_TV_Schedule.Close
    Set Rs_TV_Schedule = Nothing
End Sub

Public Function Get_Month_Week_Commencing_Date(ByVal Var_Date As Date) As Integer
    Dim Begin_date As Date
    Dim End_Date As Date
    Dim Date_Index As Date
    Dim str_week As String
    
    
    str_week = Get_WeekCommencing(month(Var_Date), Year(Var_Date))
    End_Date = DateAdd("d", 6, CDate(Format(month(Var_Date) & "/" & Right(str_week, 2) & "/" & Year(Var_Date), "mm/dd/yyyy")))
    
    If CInt(Left(str_week, 2)) > 10 Then
        If month(Var_Date) = 1 Then
            Begin_date = CDate(Format(IIf(month(Var_Date) - 1 = 0, 12, month(Var_Date) - 1) & "/" & Left(str_week, 2) & "/" & (Year(Var_Date) - 1), "mm/dd/yyyy"))
        Else
            Begin_date = CDate(Format(IIf(month(Var_Date) - 1 = 0, 12, month(Var_Date) - 1) & "/" & Left(str_week, 2) & "/" & Year(Var_Date), "mm/dd/yyyy"))
        End If
    Else
        Begin_date = CDate(Format(month(Var_Date) & "/" & Left(str_week, 2) & "/" & Year(Var_Date), "mm/dd/yyyy"))
    End If
    
    
    If Var_Date > End_Date Then
        If month(Var_Date) = 12 Then
            Get_Month_Week_Commencing_Date = 1
        Else
            Get_Month_Week_Commencing_Date = month(Var_Date) + 1
        End If
    ElseIf Var_Date < Begin_date Then
        If month(Var_Date) = 1 Then
            Get_Month_Week_Commencing_Date = 12
        Else
            Get_Month_Week_Commencing_Date = month(Var_Date) - 1
        End If
    Else
        Get_Month_Week_Commencing_Date = month(Var_Date)
    End If
'    End_Date = DateAdd("d", -6, Var_Date)
'
'    For Date_Index = End_Date To Var_Date
'        'Jika HAri Minggu
'        If Weekday(Date_Index) = 1 Then
'            Get_Month_Week_Commencing_Date = Month(Date_Index)
'            Exit Function
'        End If
'    Next Date_Index
    
 
    
 End Function

 Public Function Get_Week_Commencing_Date(ByVal Var_Date As Date) As Integer
    Dim End_Date As Date
    Dim Date_Index As Date
    
    End_Date = DateAdd("d", -6, Var_Date)
    
    For Date_Index = End_Date To Var_Date
        'Jika HAri Minggu
        If Weekday(Date_Index) = 1 Then
            Get_Week_Commencing_Date = Day(Date_Index)
            Exit Function
        End If
    Next Date_Index
 End Function


Public Function Is_date_In_WeekCommencing(Date_Actual As Integer, Month_Actual As Integer, Month_Week_Comm As Integer, Year_Week_Comm As Integer) As Boolean
    Dim strDate As String
    Dim Actual_Month As Integer
    Dim Date_Array(35) As String
    Dim Array_Index As Integer
    Dim Day_Start As Date
    Dim Day_Stop As Date
    Dim Day_Index As Date
    Dim Str_Date_Actual  As String
    
    Str_Date_Actual = Date_Actual & "/" & Month_Actual
    strDate = Get_WeekCommencing(Month_Week_Comm, Year_Week_Comm)
    '3006132027
    Do While strDate <> ""
        'Ambil Week
        
        'Cari Actual Month
        Actual_Month = Get_Actual_Month(CInt(Mid(strDate, 1, 2)), CInt(Mid(strDate, 1, 2)), Month_Week_Comm, Year_Week_Comm)
        
        'Masukkan Ke Array Tambah 6 HAri
        Day_Start = CDate(Format(Mid(strDate, 1, 2) & "/" & Actual_Month & "/" & Year_Week_Comm, "dd/mm/yyyy"))
        Day_Stop = DateAdd("d", 6, Day_Start)
        For Day_Index = Day_Start To Day_Stop
            Date_Array(Array_Index) = Day(Day_Index) & "/" & month(Day_Index)
            Array_Index = Array_Index + 1
        Next Day_Index
        
        'Strdate - 2 karakter dari depan
        strDate = Right(strDate, Len(strDate) - 2)
    Loop
    'Loop
    For Array_Index = 0 To 35
    'Cari Tanggal apakah dalm Week Commencing
        If Str_Date_Actual = Date_Array(Array_Index) Then
            
            Is_date_In_WeekCommencing = True
            Exit Function
            
        End If
    ' End Loop
        Is_date_In_WeekCommencing = False
    Next Array_Index
        
            
End Function

Public Function Get_Actual_Year(ByVal Date_Act As Integer, ByVal Week_Comm As Integer, ByVal Month_Week_Comm As Integer, ByVal Act_Year As Integer) As Integer
    Dim Str_Week_Commencing As String
    Dim Start_Month As Integer
    Dim Start_Date As Date
    Dim End_Date As Date
    Dim Index_Date As Date
    Dim Str_Temp As String
    Dim Pernah_Flag As Boolean
    Dim act_Year_asli As Integer
    act_Year_asli = Act_Year
    Pernah_Flag = False
    
    Str_Week_Commencing = Get_WeekCommencing(Month_Week_Comm, Act_Year)
    
    Do While Str_Week_Commencing <> ""
        Str_Temp = Mid(Str_Week_Commencing, 1, 2)
        Str_Week_Commencing = Right(Str_Week_Commencing, Len(Str_Week_Commencing) - 2)
        
        If Not Pernah_Flag Then
            If CInt(Str_Temp) > 10 Then
            
                'Old Start_Month = Month_Week_Comm - 1
                'Jika Month=1 minus 1 maka 12
                If Month_Week_Comm = 1 Then
                        Start_Month = 12
                        Act_Year = Act_Year - 1
                Else
                    Start_Month = Month_Week_Comm - 1
                End If
                
            Else
                Start_Month = Month_Week_Comm
            End If
        Else
            ' ?
            Act_Year = act_Year_asli
            Start_Month = Month_Week_Comm
        End If
        
        Pernah_Flag = True
        
        If CInt(Str_Temp) = Week_Comm Then
            
            Start_Date = CDate(Start_Month & "/" & Str_Temp & "/" & Act_Year)
            End_Date = DateAdd("d", 6, Start_Date)
            
            For Index_Date = Start_Date To End_Date
                If Day(Index_Date) = Date_Act Then
                    Get_Actual_Year = Year(Index_Date)
                    Exit Function
                End If
            Next Index_Date
        End If
    Loop
    
    
End Function

Public Function Get_Actual_Month(ByVal Date_Act As Integer, ByVal Week_Comm As Integer, ByVal Month_Week_Comm As Integer, ByVal Act_Year As Integer) As Integer
    Dim Str_Week_Commencing As String
    Dim Start_Month As Integer
    Dim Start_Date As Date
    Dim End_Date As Date
    Dim Index_Date As Date
    Dim Str_Temp As String
    Dim Pernah_Flag As Boolean
    Pernah_Flag = False
    
    Str_Week_Commencing = Get_WeekCommencing(Month_Week_Comm, Act_Year)
    
    Do While Str_Week_Commencing <> ""
        Str_Temp = Mid(Str_Week_Commencing, 1, 2)
        Str_Week_Commencing = Right(Str_Week_Commencing, Len(Str_Week_Commencing) - 2)
        
        If Not Pernah_Flag Then
            If CInt(Str_Temp) > 10 Then
                'Old Start_Month = Month_Week_Comm - 1
                'Jika Month=1 minus 1 maka 12
                If Month_Week_Comm = 1 Then
                        Start_Month = 12
                        Act_Year = Act_Year - 1
                Else
                    Start_Month = Month_Week_Comm - 1
                End If
            Else
                Start_Month = Month_Week_Comm
            End If
        Else
            ' ?
            If Str_Week_Commencing = "" And CInt(Str_Temp) < 10 Then
                Start_Month = Month_Week_Comm + 1
            Else
                Start_Month = Month_Week_Comm
            End If
        End If
        
        Pernah_Flag = True
        
        If CInt(Str_Temp) = Week_Comm Then
            Start_Date = CDate(Start_Month & "/" & Str_Temp & "/" & Act_Year)
            End_Date = DateAdd("d", 6, Start_Date)
            For Index_Date = Start_Date To End_Date
                If Day(Index_Date) = Date_Act Then
                    Get_Actual_Month = month(Index_Date)
                    Exit Function
                End If
            Next Index_Date
        End If
    Loop
    
End Function
Public Function Get_Program_Name(ByVal Actual_Program As String) As String
    Dim Str_Program_Code As String
    Dim strSql As String
    Dim Rs_Program_Code As New ADODB.Recordset
    
    Str_Program_Code = Trim(Left(Actual_Program, 3))
    strSql = "SELECT Program_Name FROM Program_Code WHERE Program_Code='" & Str_Program_Code & "'"
    Rs_Program_Code.Open strSql, ConnERP, , , adCmdText
    If Rs_Program_Code.EOF Then
        Get_Program_Name = Actual_Program
    Else
        Get_Program_Name = Trim(Rs_Program_Code.Fields(0).Value)
    End If
    
    Rs_Program_Code.Close
    Set Rs_Program_Code = Nothing
End Function

Public Function Only_Matrial_Name(ByVal Str_Material As String) As String
    Dim Index_Char As Integer
    Dim Str_Temp As String
    
    For Index_Char = 1 To Len(Str_Material)
        If Mid(Str_Material, Index_Char, 1) = ">" Then
            Str_Temp = Mid(Str_Material, Index_Char + 1, Len(Str_Material) - Index_Char)
            Exit For
        End If
    Next Index_Char
    
    Only_Matrial_Name = Trim(Str_Temp)
    
End Function

Public Sub Create_Program_Rate_Temp()
    'Program
    
    Rs_Temp_Program.Fields.Append "Station_Code", adVarChar, 10, adFldIsNullable
    Rs_Temp_Program.Fields.Append "Market", adVarChar, 30, adFldIsNullable
    Rs_Temp_Program.Fields.Append "Program_Name", adVarChar, 75, adFldIsNullable
    Rs_Temp_Program.Fields.Append "Day", adVarChar, 3, adFldIsNullable
    Rs_Temp_Program.Fields.Append "Start_Time", adVarChar, 8, adFldIsNullable
    Rs_Temp_Program.Fields.Append "End_Time", adVarChar, 8, adFldIsNullable
    Rs_Temp_Program.Fields.Append "TVR", adDouble, , adFldIsNullable
    
    'Rate
    Rs_Temp_Program_Rate.Fields.Append "Station_Code", adVarChar, 10, adFldIsNullable
    Rs_Temp_Program_Rate.Fields.Append "Program_Name", adVarChar, 75, adFldIsNullable
    Rs_Temp_Program_Rate.Fields.Append "Day", adVarChar, 3, adFldIsNullable
    Rs_Temp_Program_Rate.Fields.Append "Start_Time", adVarChar, 8, adFldIsNullable
    Rs_Temp_Program_Rate.Fields.Append "End_Time", adVarChar, 8, adFldIsNullable
    Rs_Temp_Program_Rate.Fields.Append "Duration", adDouble, , adFldIsNullable
    Rs_Temp_Program_Rate.Fields.Append "Gross_Rate", adCurrency, , adFldIsNullable
    Rs_Temp_Program_Rate.Fields.Append "Gross_Paid", adCurrency, , adFldIsNullable
    Rs_Temp_Program_Rate.Fields.Append "Nett_Rate", adCurrency, , adFldIsNullable
    
End Sub
Public Sub Create_Program_Rate_Temp_SRI()
    'Program
    
    Rs_Temp_Program_SRI.Fields.Append "Station_Code", adVarChar, 10, adFldIsNullable
    Rs_Temp_Program_SRI.Fields.Append "Market", adVarChar, 30, adFldIsNullable
    Rs_Temp_Program_SRI.Fields.Append "Program_Name", adVarChar, 75, adFldIsNullable
    Rs_Temp_Program_SRI.Fields.Append "Day", adVarChar, 3, adFldIsNullable
    Rs_Temp_Program_SRI.Fields.Append "Start_Time", adVarChar, 8, adFldIsNullable
     Rs_Temp_Program_SRI.Fields.Append "End_Time", adVarChar, 8, adFldIsNullable
    Rs_Temp_Program_SRI.Fields.Append "TVR", adDouble, , adFldIsNullable
    
    'Rate
    Rs_Temp_Program_Rate_SRI.Fields.Append "Station_Code", adVarChar, 10, adFldIsNullable
    Rs_Temp_Program_Rate_SRI.Fields.Append "Program_Name", adVarChar, 75, adFldIsNullable
    Rs_Temp_Program_Rate_SRI.Fields.Append "Day", adVarChar, 3, adFldIsNullable
    Rs_Temp_Program_Rate_SRI.Fields.Append "Start_Time", adVarChar, 8, adFldIsNullable
    Rs_Temp_Program_Rate_SRI.Fields.Append "End_Time", adVarChar, 8, adFldIsNullable
    Rs_Temp_Program_Rate_SRI.Fields.Append "Duration", adDouble, , adFldIsNullable
    Rs_Temp_Program_Rate_SRI.Fields.Append "Gross_Rate", adCurrency, , adFldIsNullable
    Rs_Temp_Program_Rate_SRI.Fields.Append "Gross_Paid", adCurrency, , adFldIsNullable
    Rs_Temp_Program_Rate_SRI.Fields.Append "Nett_Rate", adCurrency, , adFldIsNullable
    
End Sub

Public Function Get_WeekCommencing(ByVal month As Integer, ByVal Year As Integer) As String
'************************************************
' Procedure         : Get_WeekCommencing
' Function          : To get Week Commencing
' Date              : 01/10/2001
' Parameter Input   : Month, Year
' Parameter Output  :
' Last Update/By    :
'************************************************


'   Dim Left_Date As String
'   Dim a As Date
'   Dim Last_Date
'   Dim week_before As Integer
'   Dim Date_Quee As String
'   Dim Str_Temp As String
'   Dim Month_Before As Integer
'
'   Last_Date = DateAdd("m", 1, CDate(Format(Month & "/01" & "/" & Year, "mm/dd/yyyy"))) - 1
'    For a = CDate(Format(Month & "/01/" & Year, "mm/dd/yyyy")) To Last_Date
'    If Weekday(a) = 1 Then
'            If Day(a) < 10 Then
'                Date_Quee = Date_Quee & "0" & Day(a)
'            Else
'                Date_Quee = Date_Quee & Day(a)
'            End If
'    End If
'   Next a
'   Select Case Month
'    'Week Commencing 5
'    Case 1, 4, 7, 10
'        If Len(Date_Quee) < 10 Then
'            Last_Date = DateAdd("ww", -1, CDate(Format(Month & "/" & Left(Date_Quee, 2) & "/" & Year, "mm/dd/yyyy"))) '- 1
'            For a = Last_Date To CDate(Format(Month & "/01" & "/" & Year, "mm/dd/yyyy"))
'                If Weekday(a) = 1 Then
'                    week_before = Day(a)
'                End If
'            Next a
'            Get_WeekCommencing = week_before & Date_Quee
'        Else
'            Get_WeekCommencing = Date_Quee
'        End If
'        'Get_WeekCommencing = Date_Quee
'     'Week Commencing 4
'    Case 2, 3, 5, 6, 8, 9, 11, 12
'        If Len(Date_Quee) > 8 Then
'
'            Str_Temp = Left(Date_Quee, 8)
'            'Date_Temp = Right(Date_Quee, 2)
'        Else
'            Month_Before = Month - 1
'            Select Case Month_Before
'                Case 2, 3, 5, 6, 8, 9, 11, 12
'                    Left_Date = Get_WeekCommencing((Month - 1), Year)
'                    If Len(Left_Date) > 8 Then
'                        Str_Temp = Right(Left_Date, 2) & Date_Quee
'                    Else
'                        Str_Temp = Date_Quee
'                    End If
'                Case Else
'                     Str_Temp = Date_Quee
'            End Select
'
'            'Date_Temp = ""
'        End If
'            Get_WeekCommencing = Str_Temp
'
'    End Select

'============================== New =====================

'    Dim Rs_Week As New ADODB.Recordset
'    Dim StrSQL As String
'    Dim Str_Temp As String
'
'    'Get Week
'    StrSQL = "SELECT * From Week_Commencing WHERE Month=" & Month & " AND Year=" & Year
'
'    Rs_Week.CursorLocation = adUseClient
'    Rs_Week.Open StrSQL, ConnERP, , , adCmdText
'
'    If Rs_Week.EOF Then
'        Get_WeekCommencing = ""
'    Else
'        If Day(Rs_Week.Fields("Week_1").Value) < 10 Then
'            Str_Temp = "0" & Day(Rs_Week.Fields("Week_1").Value)
'        Else
'            Str_Temp = Day(Rs_Week.Fields("Week_1").Value)
'        End If
'
'        If Day(Rs_Week.Fields("Week_2").Value) < 10 Then
'             Str_Temp = Str_Temp & "0" & Day(Rs_Week.Fields("Week_2").Value)
'        Else
'            Str_Temp = Str_Temp & Day(Rs_Week.Fields("Week_2").Value)
'        End If
'
'        If Day(Rs_Week.Fields("Week_3").Value) < 10 Then
'            Str_Temp = Str_Temp & "0" & Day(Rs_Week.Fields("Week_3").Value)
'        Else
'            Str_Temp = Str_Temp & Day(Rs_Week.Fields("Week_3").Value)
'        End If
'
'        If Day(Rs_Week.Fields("Week_4").Value) < 10 Then
'            Str_Temp = Str_Temp & "0" & Day(Rs_Week.Fields("Week_4").Value)
'        Else
'            Str_Temp = Str_Temp & Day(Rs_Week.Fields("Week_4").Value)
'        End If
'
'        'Str_Temp = Str_Temp & Day(Rs_Week.Fields("Week_5").Value)
'        If Day(Rs_Week.Fields("Week_5").Value) < 10 Then
'            Str_Temp = Str_Temp & "0" & Day(Rs_Week.Fields("Week_5").Value)
'        Else
'            Str_Temp = Str_Temp & Day(Rs_Week.Fields("Week_5").Value)
'        End If
'
'        Get_WeekCommencing = Str_Temp
'
'    End If
'
'    Rs_Week.Close
'    Set Rs_Week = Nothing
    ' =============================== End New ==========================
    
    Dim idx As Integer
    
    For idx = 1 To intTotalWeek
        If month = ArrWeekCommencing(idx).WeekMonth And Year = ArrWeekCommencing(idx).WeekYear Then
            Get_WeekCommencing = ArrWeekCommencing(idx).WeekCommencingDate
            Exit For
        End If
        
        Get_WeekCommencing = "" 'week Not Found
    Next idx
    
End Function
Public Function Only_Program(Str_Program As String) As String
    Dim Index_Char As Integer
    Dim Str_Temp As String
    For Index_Char = 1 To Len(Str_Program)
        If Mid(Str_Program, Index_Char, 1) = "{" Then
            Exit For
        End If
        Str_Temp = Str_Temp & Mid(Str_Program, Index_Char, 1)
    Next Index_Char
    Only_Program = Trim(Str_Temp)
End Function
Public Function Get_Str_Day(Str_Program As String) As String
Dim Index_Char As Integer
    Dim Str_Temp As String
    For Index_Char = 1 To Len(Str_Program)
        If Mid(Str_Program, Index_Char, 1) = "{" Then
            Str_Temp = Str_Temp & Mid(Str_Program, Index_Char + 1, 3)
            Exit For
        End If
    Next Index_Char
    Get_Str_Day = Trim(Str_Temp)
End Function

Public Function Get_Start_Time(Str_Program As String) As String
Dim Index_Char As Integer
Dim Str_Temp As String
Dim Index_Char1 As Integer
Dim Str_Temp1 As String

    For Index_Char = 1 To Len(Str_Program)
        If Mid(Str_Program, Index_Char, 1) = "{" Then
            Str_Temp = Str_Temp & Mid(Str_Program, Index_Char + 5, Len(Str_Program) - (Index_Char + 5))
            Exit For
        End If
    Next Index_Char
    'Get Start
    For Index_Char1 = 1 To Len(Str_Temp)
        If Mid(Str_Temp, Index_Char1, 1) = "-" Then
            Exit For
        End If
        Str_Temp1 = Str_Temp1 & Mid(Str_Temp, Index_Char1, 1)
    Next Index_Char1
    
    
    Get_Start_Time = Trim(Str_Temp1)
End Function
Public Function Get_End_Time(Str_Program As String) As String
Dim Index_Char As Integer
Dim Str_Temp As String
Dim Index_Char1 As Integer
Dim Str_Temp1 As String

    For Index_Char = 1 To Len(Str_Program)
        If Mid(Str_Program, Index_Char, 1) = "{" Then
            Str_Temp = Str_Temp & Mid(Str_Program, Index_Char + 5, Len(Str_Program) - (Index_Char + 5))
            Exit For
        End If
    Next Index_Char
    
     For Index_Char1 = 1 To Len(Str_Temp)
        If Mid(Str_Temp, Index_Char1, 1) = "-" Then
            Str_Temp1 = Mid(Str_Temp, Index_Char1 + 1, Len(Str_Temp) - (Index_Char1))
            Exit For
        End If
    Next Index_Char1
    
    'Get End
    Get_End_Time = Trim(Str_Temp1)
End Function

Public Function Day_Index(ByVal strDay As String) As Integer
    Select Case UCase(strDay)
        Case "SUN"
            Day_Index = 1
        Case "MON"
            Day_Index = 2
        Case "TUE"
            Day_Index = 3
        Case "WED"
            Day_Index = 4
        Case "THU"
            Day_Index = 5
        Case "FRI"
            Day_Index = 6
        Case "SAT"
            Day_Index = 0
    End Select
End Function
Public Function Day_Name(ByVal Day_Index As Integer) As String
    Select Case Day_Index
        Case 1
            Day_Name = "Sun"
        Case 2
            Day_Name = "Mon"
        Case 3
            Day_Name = "Tue"
        Case 4
            Day_Name = "Wed"
        Case 5
            Day_Name = "Thu"
        Case 6
            Day_Name = "Fri"
        Case 7
            Day_Name = "Sat"
    End Select
End Function
Public Function Get_Objective_Id() As Double
    '=========================
    'To GET New Objective ID==
    '=========================
    Dim rsMaxObjId As New ADODB.Recordset
    Dim DblMaxObjId As Double
    Dim strSql As String
    
    strSql = "SELECT * FROM IB_TV_Last_Objective_ID"
    rsMaxObjId.Open strSql, ConnERP, adOpenForwardOnly, adLockReadOnly
    If rsMaxObjId.EOF Then
        DblMaxObjId = 1
        strSql = "INSERT INTO IB_TV_Last_Objective_ID VALUES(" & DblMaxObjId & ")"
        ConnERP.Execute strSql
    Else
        DblMaxObjId = IIf(IsNull(rsMaxObjId(0)), 0, rsMaxObjId(0)) + 1
        'Update untuk Last ID
        strSql = "UPDATE IB_TV_Last_Objective_ID SET Last_id = " & DblMaxObjId
        ConnERP.Execute strSql
    End If
    rsMaxObjId.Close
    Set rsMaxObjId = Nothing
    
    
    
    'assign
    Get_Objective_Id = DblMaxObjId
    
End Function
Public Function Get_Campaign_Id() As Double
    '=========================
    'To GET New Campaign ID==
    '=========================
    Dim rsMaxObjId As New ADODB.Recordset
    Dim DblMaxObjId As Double
    Dim strSql As String
    
    strSql = "SELECT * FROM IB_TV_Last_Campaign_ID"
    rsMaxObjId.Open strSql, ConnERP, adOpenForwardOnly, adLockReadOnly
    If rsMaxObjId.EOF Then
        DblMaxObjId = 1
        strSql = "INSERT INTO IB_TV_Last_Campaign_ID VALUES(" & DblMaxObjId & ")"
        ConnERP.Execute strSql
    Else
        DblMaxObjId = IIf(IsNull(rsMaxObjId(0)), 0, rsMaxObjId(0)) + 1
        'Update untuk Last ID
        strSql = "UPDATE IB_TV_Last_Campaign_ID SET Last_id = " & DblMaxObjId
        ConnERP.Execute strSql
    End If
    rsMaxObjId.Close
    Set rsMaxObjId = Nothing
    
    
    
    'assign
    Get_Campaign_Id = DblMaxObjId
    
End Function

Public Function Get_TV_Media_Quotation_No(strBrandCode As String, intQuotYear As Integer, StrMediaType As String) As String
    Dim Out_Param As ADODB.Parameter
    Dim In_Param1 As ADODB.Parameter
    Dim In_Param2 As ADODB.Parameter
    Dim In_Param3 As ADODB.Parameter
    Dim cmd As New ADODB.Command
    Dim New_MQ_No As String
         
     cmd.CommandType = adCmdStoredProc
     cmd.CommandText = "Get_TV_Media_Quotation_No"
     
     Set In_Param1 = cmd.CreateParameter("Brand_Code", adChar, adParamInput, 4)
     Set In_Param2 = cmd.CreateParameter("Media_Type", adChar, adParamInput, 3)
     Set In_Param3 = cmd.CreateParameter("Year", adInteger, adParamInput)
     Set Out_Param = cmd.CreateParameter("New_MQ_No", adChar, adParamOutput, 13)
        
     cmd.Parameters.Append In_Param1
     cmd.Parameters.Append In_Param2
     cmd.Parameters.Append In_Param3
     cmd.Parameters.Append Out_Param

     In_Param1.Value = strBrandCode
     In_Param2.Value = StrMediaType
     In_Param3.Value = intQuotYear
    
     'Execute
     cmd.ActiveConnection = ConnERP
     cmd.Execute
     
     New_MQ_No = Out_Param.Value
     
     'Put New Job Number
     Get_TV_Media_Quotation_No = New_MQ_No
    End Function
Public Sub Load_TV_Program_Period()
    Dim idx As Integer
    Dim rsTemp As New ADODB.Recordset
    Dim strSql As String
    
    'Exit Sub
      
    'strSQL = "select * from tv_program_period where datediff(day,the_date,getdate())=0"
    strSql = "select * from tv_program_period order by the_date desc"
    
    rsTemp.CursorLocation = adUseClient
    rsTemp.Open strSql, ConnERP, adOpenForwardOnly, adLockReadOnly
    
    If Not rsTemp.EOF Then
        'Ini Nanati di Remark
        mediaTV.Str_Month_TV_Prg_Rate = rsTemp.Fields("Month_Period").Value
        mediaTV.Str_Year_TV_Prg_Rate = rsTemp.Fields("Year_Period").Value
        'End Ini Nanati di remark
        
        'Looping masukkan ke Array
        ReDim ArrProgramRatePeriod(rsTemp.RecordCount)
        idx = 0
        Do While Not rsTemp.EOF
            ArrProgramRatePeriod(idx).StationCode = IIf(IsNull(rsTemp.Fields("Station_code").Value), "", rsTemp.Fields("Station_code").Value)
            ArrProgramRatePeriod(idx).TheMonth = rsTemp.Fields("Month_Period").Value
            ArrProgramRatePeriod(idx).TheYear = rsTemp.Fields("Year_Period").Value
            idx = idx + 1
            rsTemp.MoveNext
        Loop
        
        rsTemp.Close
        Set rsTemp = Nothing
        
        Exit Sub
    End If
    
    rsTemp.Close
    Set rsTemp = Nothing
    
    
'     '===== Year Criteria =========
'    strSQL = "SELECT MAX(Year) AS Tahun FROM TV_Program_Rate_New "
'
'    rsTemp.Open strSQL, ConnERP, adOpenForwardOnly, adLockReadOnly
'    If rsTemp.EOF Then
'        Str_Year_TV_Prg_Rate = Year(Now)
'    Else
'        Str_Year_TV_Prg_Rate = rsTemp.Fields("Tahun").Value
'    End If
'
'    rsTemp.Close
'
'    '====== Month Criteria =========

'    strSQL = "SELECT MAX(Month) AS Bulan FROM TV_Program_Rate_New WHERE "
'    strSQL = strSQL & " Year=" & Str_Year_TV_Prg_Rate
'
'    rsTemp.Open strSQL, ConnERP, adOpenForwardOnly, adLockReadOnly
'
'    If rsTemp.EOF Then
'        Str_Month_TV_Prg_Rate = month(Now)
'    Else
'        Str_Month_TV_Prg_Rate = rsTemp.Fields("Bulan").Value
'    End If
'
'    rsTemp.Close
'    Set rsTemp = Nothing
'
'    ConnERP.BeginTrans
'
'    strSQL = "DELETE FROM tv_program_period "
'    ConnERP.Execute strSQL
'
'    strSQL = "INSERT INTO tv_program_period (The_Date,Month_Period,Year_Period) VALUES(getdate(),'" & Str_Month_TV_Prg_Rate & "','" & Str_Year_TV_Prg_Rate & "')"
'    ConnERP.Execute strSQL
'
'    ConnERP.CommitTrans
    
End Sub

Public Function Get_ProgramRatePeriod(ByVal strStationCode As String, datepart As String) As Integer
    Dim idx As Integer
    
    For idx = 0 To UBound(ArrProgramRatePeriod)
        If (ArrProgramRatePeriod(idx).StationCode) = UCase(strStationCode) Then
            If UCase(datepart) = "M" Then
                Get_ProgramRatePeriod = ArrProgramRatePeriod(idx).TheMonth
            ElseIf UCase(datepart) = "Y" Then
                Get_ProgramRatePeriod = ArrProgramRatePeriod(idx).TheYear
            End If
            Exit For
        End If
        
    Next idx
        
        'Default output
        If UCase(datepart) = "M" Then
            Get_ProgramRatePeriod = month(Date)
        ElseIf UCase(datepart) = "Y" Then
            Get_ProgramRatePeriod = Year(Date)
        End If
        
End Function

Public Sub Load_Week_Commencing()
    Dim Rs_Week As New ADODB.Recordset
    Dim strSql As String
    Dim Str_Temp As String
    Dim idx As Integer
    
    'Get Week
    strSql = "SELECT * From Week_Commencing"
    
    Rs_Week.CursorLocation = adUseClient
    Rs_Week.Open strSql, ConnERP, , , adCmdText
    
    ReDim ArrWeekCommencing(Rs_Week.RecordCount)
    
    intTotalWeek = Rs_Week.RecordCount
    
    idx = 1
    
    Do While Not Rs_Week.EOF
      
        If Day(Rs_Week.Fields("Week_1").Value) < 10 Then
            Str_Temp = "0" & Day(Rs_Week.Fields("Week_1").Value)
        Else
            Str_Temp = Day(Rs_Week.Fields("Week_1").Value)
        End If
        
        If Day(Rs_Week.Fields("Week_2").Value) < 10 Then
             Str_Temp = Str_Temp & "0" & Day(Rs_Week.Fields("Week_2").Value)
        Else
            Str_Temp = Str_Temp & Day(Rs_Week.Fields("Week_2").Value)
        End If
        
        If Day(Rs_Week.Fields("Week_3").Value) < 10 Then
            Str_Temp = Str_Temp & "0" & Day(Rs_Week.Fields("Week_3").Value)
        Else
            Str_Temp = Str_Temp & Day(Rs_Week.Fields("Week_3").Value)
        End If
        
        If Day(Rs_Week.Fields("Week_4").Value) < 10 Then
            Str_Temp = Str_Temp & "0" & Day(Rs_Week.Fields("Week_4").Value)
        Else
            Str_Temp = Str_Temp & Day(Rs_Week.Fields("Week_4").Value)
        End If
        
        'Str_Temp = Str_Temp & Day(Rs_Week.Fields("Week_5").Value)
        If Day(Rs_Week.Fields("Week_5").Value) < 10 Then
            Str_Temp = Str_Temp & "0" & Day(Rs_Week.Fields("Week_5").Value)
        Else
            Str_Temp = Str_Temp & Day(Rs_Week.Fields("Week_5").Value)
        End If
        
        ArrWeekCommencing(idx).WeekYear = Rs_Week.Fields("Year").Value
        ArrWeekCommencing(idx).WeekMonth = Rs_Week.Fields("Month").Value
        ArrWeekCommencing(idx).WeekCommencingDate = Str_Temp
        
        idx = idx + 1
        
        Rs_Week.MoveNext 'Next Week
        
    Loop
       
    Rs_Week.Close
    Set Rs_Week = Nothing
End Sub

Public Function GetStationTVMapping(ByVal StrTVStationCode As String) As String
    Dim recStation As New ADODB.Recordset
    
    strQuery = "SELECT  Nielsen_Station_Code"
    strQuery = strQuery & " FROM TV_Station_Code_Mapping"
    strQuery = strQuery & " WHERE station_code_ERP='" & StrTVStationCode & "'"
    
    recStation.Open strQuery, ConnERP, adOpenStatic, adLockReadOnly
    
    If Not recStation.EOF Then
        If IsNull(recStation.Fields("nielsen_station_code").Value) Then
            GetStationTVMapping = StrTVStationCode
        Else
            GetStationTVMapping = recStation.Fields("nielsen_station_code").Value
        End If
        
    Else
        GetStationTVMapping = StrTVStationCode
    End If
    
    recStation.Close
    Set recStation = Nothing

End Function

