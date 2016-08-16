Attribute VB_Name = "mdl_xpivot"
Public strRefTemplate As String

Public Sub PivotVisualDesign(ByRef paPivot As Pivot)
'************************************************
' Procedure         : PivotVisualDesign
' Function          : Format Tampilan Ketika Pertama kali pivot diload Visual
' Input Parameter   : paPivot ---- isinya Pivot
' Programmer By     : Ryo
' Date Update       : Jan-2016
' Update By         : Ryo/ Kreatif
'************************************************

    Dim strDateFormat As String, strTimeFormat As String
    strDateFormat = "": strTimeFormat = ""

    With paPivot
        
        .BeginUpdate
        .VisualDesign = "gBFLBWIgBAEHhEJAEGg7yCEHBp+A4ABMIAcKAACAkYAIFjACA0cBkYAYNjAEB0YAwPjAICUfBEIhgACIgg6Si4AgRCJ0zjEHmYBPM5QAAYAQGKIYBkAKBQAGaAoDDQMw3QwAAxDGLEEwsACEIrjKCRShyCYZRhGcTALD8EhhECTZKkAYgEiKLoaRzAcwyDAcQRFCKUJxjSapPjaGwCQjQEyxWyJMBOACiaLiabAAimMooUJLUgAXDYGyDEqNKRpChoXhEMg1CTXVgRCKoYTDBKybLnWp4apSA5RS7acB0dRtGBpEi4bhmOY5Aj6Go4SpgNL4FSUNgnNgAArrCJ6apqS6EXRBcjXBDdJwTQ6BaqgPCqDpmXKOQjYVpRFYdURmjIaZuDICoOGAARukSXh1FkCRdDIRR+AAVB1hsG4VlcCYSjmZQjFkax8CcNAAlabpvnGTQskuGAAjSag0HEaoOi8EYvmQch2lCEBECSIBpDGHQOicIwokYQJOgYEIJj4eJQloEgomKQhzhIEJEF0MQRmSCZLA0H5Pm+dZ8H+X5oAWfYAqODAAH8BBxmCcICGCOAqAyA5gmgRgNgSYJHnIBx+gTLRxBARwTjAQQck6QQAHIOphniBgjUKKIGB2CghmGSJtgYTwYFmfASkeM5EEJRRYHya4IiIQZODuCZhlEAgrggBAxlCbYNCMF5qBcNBElGdJyBCBxDjoHw7GQKB6DiFRlgkDJthETppDScoMh+IRAH8PJEEGVJuWIW" & _
            "JmCmGgm2WK4LiQU40EWGYmCOThgDQRQiAIZAdg6eRyDeFplBmfBChqRwzjgRIRgQWZyHmHYkiecBUCCDxBBgdw8AQQZY3CZoaFaF4aGWSRyHMOxEHoToPD0JxxFiBA9GoE5CgkO5mFOOorisaB6BgdQ7kQaZwEiAoqGOVh3BYSBjloMARiCIQUHeLQsBsDhzDubAoEwSQeg8GpwnaHZFjIIJ2BKB4jkgSAeE8ChInCNZthoTBHh8KQjkYMASm0WxGk+IxfiEdIKD2RJBmycgSiMWxynAPBtkEJB4j0IYrFCcQTAeQ5KmoPZnHEfILD4BJBnScwSi+WwsnGQgPCwKp6jKTwhC6fYTE8PAIniHZGCINJ4BOB5jk8BZEEeS5uE6RIPBiUwgDyJZsHCc5GHEe4glEGQInOZBPB+DxMisKZHGWYQ0HwPJxEyTJzh0cwDmYdwZgkE5qGIPREnIAJ0BOZmVEuDERQAx3CXE0vAUAoggAeF0CEeoOxGgEEyPcVY6hiAJHUCgNIQhHDFFkB4TQrhkB6HYMcVApRYhrFONwdwNRJCHGwKRiQ2w8jtFmB4TY1h8g8HaJsCDzRPDCE8PEXQbgshwG0EsD4ghQgHGEPEcQn0tCeG4HcEAexojnFgKoIQHh+BZHzA4cgwR9hTA+BwbI7xph1A4OUJo0x6hEBCO0aQdATjsHcDgSYRx0CuCEJ4cQtw4jdA+B0a4cghD3F6N8QI3wFDkBiPEbQdoijmCcB8bwXxNjmDyN4R" & _
            "4oR0j5H8K8ToOxgC+GYAUdoQBIEPEAIAARCCakHDxK8nQWCeE+J+UEL8tCilHKSUsppTyolTKqVcAwIxIi0K6V8sJYyylnLSWstpby4lzAGDAu5eS9l9L+YEwZhTDmJMWIQZwIw2BGAUPQAIRBRhFFOIABITBlBRCBD8IA0RYBFAMEsDQpAojoFiHgRAyA1AlFcNccQIwoiFEEBMMgswqBgDqAgXoTgDgROGKkSgug0AADKcs5pzzog1CEIEbobhYB5CqMZCQDAhBtECKoXgXQuCpCkIAGI+QNBWAuXUYQYA2iSBSAUXYEQEiqGAPESYWgOhuFmBYWA2g3DWBWD0SwxxXDDDEKEHAmAPB7COHkS4XQZgUF4OQN05qpA3FyHwWo2hPixG0fEPQHB/gKDeBwawTBdAgAW38KoFhbBIBaJEDg9wDidFyNAIQSgDCTAyNAdQcgZCMDIAoNQgQECbHYI0IoYAxCPAcIoJY3gAAcBmHwEoKhShcH8L8d86k0hPckGATAhhQiuDYJ4M8ngDDBHgHQEKfBoA8AmAUYIpxUjODoJ0KA3AdgOCGzoFoBQNgwDiOoOgQxIhBGqBUIokByBHAKJ4Q4rAnB1A6Gd3oGwjCxF4GkC4xLlDsuaGcMggQQhIBkBwQ4HgjjHFiNQOwHRZBcE2CYJMcxIB7A8NIBI2wVhLAkGQVoHglCHF4JgO4Sx54RBgEwOQ4R3gvA6McPI2QaDNDmFTFQzBji8GMHcZwcxWDTByMcSYOR4" & _
            "g0E8BMcg/A7jPFoAQeQfAPgUACOMIYUQDiFAeEUI4lAQjzCKB4KghQSB4GkDQUoGg+CfBoKGNgU3NA3CflMYI8wngeDqCUJgeQqC1CKELMK8QNhXCCPUPoZA9BUBONULYGBhg1GaHEDAxxKixHuBgTw1gCiFA+FgCo6R7B9GSDYJgqQygeGiOURwfRPDTB6KIPckRWDPDOB8Q+JAhBbDOMUXYbxoCRHKLIPwURrgpHgHwbgxxmjFDYNQK4bRuhRC6GccI+5NiXFaO0BYXwIBFHkGIL4dxtACJyTQmwAB1S0J2T0n5CSgjRl4UYpBSimFOKgVIZgYgZgBCEEYcYCAmBUACAgRYRgwgjBQNwEwMQL2eCALEdASBhA/BMJIY4xgwBJAYLYGYHBAjMBIBAHA/BDAZBmAAJQIA5igAAM4J4NRGhQBGDMPIvA/AjAGH8aQQAKCXH4FoP4KRCDPFILQPIbA+ARA0AV8osBNClBkJcJQihmjGyaC0RQyQbgaEYCwWAmwmiuC6B8TYxQ1Dwd2LoVYUBOjTHCLIHdbwygyEeNAUQ1RdjNHEHsPQu+mC1FQEgQQSQBiuGkKsaIXAiDyFKDsVQSA1DQE4FwBwswBjHCaCsSAmh7BPAyDEWYxh6gwB2G3todRGDcBmJATYmABhDD6MQZwcQMh9GqNsAoRBsjuBuDMZgtBPgzCCNoNIwAMjwAMHcQoGhZhbEABYHwlhbDZFeI0KoNgrhOCYA8TozxuDtsdZCyVlBGBjHW" & _
            "GkMBXwmj2CMEbWoRATiZFuIQAwIhkAPDeIITAcQgAFHYAQKsqgWANAGMAIo6RogChaBjg2AOIMgDgCASA5BhBIBWAWgFhhAjgAhLgGhghjg1AJA0AiADgaAirTgjgGAmgJhigIgmAkgFBNBIwLAzAjADgeAyA4hjBIhWA2gNhjgIhIhngMAMhIg+AwgChOgDgSBCgQhkBJAWBGgQAjhIh4hMA6BkhIBeBOgThhBZAwhpgSBQA4hwhqgVhiApBmBagWhlhDgaBXA6glBAACA8A6hmBIhqBmgZhmgIhqBkQ1uJhkgQD2hGBygQAmg5g2AwA7AngJBKB6gehnhJBKB4A7AnBjAyAdgXhoBKAWgGgXhhhaAEhugihmg6A2gOgjxNgshuggBOh6BWB/glhpgKBmB/glRSAEhwgnhhhagGgigoxLB8xYAph6BKgqgqhqhKBKgmp6uTAcgshrBKhWgcgsBPAaAEhyguhqBqh2g+gvxohshyp6g7AWg4gxhsgLAmg4gxBPA6AEh0gzhsArBGhSg0x0A8h0p6hbBmhMg2hthLB2hMrdhbBihZg1guBLgWhmx+gKAEh2g6huALtjA7yDgspUhPB7hWhyg9hvgLhmhyg9BPB6AEh4g/hvhMAHACpFhLh8h4ggBPgcAnAGhChwhMA3AGhCMxAEh6hEhxAMBXAWhFyaBMyZAxgsB3Aek/AMBuAwgSgPg8gXAmhJlrgnApyShMg3AuhLhzAMhHAxgYBhADBLAGA/gzgMh" & _
            "nA6hOhzhMAYhXhPhQgbhIhZnLAFAzBGhRhUrbhDA2BvAmgRBJgUgihTBSA8hUh1BIhzBRgHBaBkAECNiDpYgABpNwJbNxNxigCZBctzpfN1Jgt2gMN3t4t5t6t7t8t9t+t/uAuBgiuCuDuEuFuGhAOIOJOKOLOMONOOOPOQOROSOTOUOVAyBpAJgIhAhfx3gFAph/AWgfhKBhAzhpAtAiBbAfAEBGgBBwgaAshigrBNgmBShigzBxgQhUKrgyhNhGgjALA0AmhVBsgXBHu4BahEAOg5Aug4BUBnB1CkBSBHB3g7hMgmBRh6AUBaBvB7gchHhqB3grg+AQCzAbgIAkgQA6h1Aug/AVhDhPAxBTgagsgGgvAJhHBAhrBQA8gNAbD2gvhlgxhxg7BGguheAAg3BaBtB3AqAOgqAHggAABdAAAaBLh9BahzgIgpBhhzgxgqhkhfBvg8haBih5AzhZA+BzgRg4gthCB0BAAOA3AfALtYBrg2AWhdgUBqgrBxhoBAh3A9BOgvA7ByBQA0gYhOgHBTt/BjAfBbBRBIBEhfgEBkBtAJAUB0BhA7g4BABiAsgABdgcB1A0hpgAhNBYAUhMgBhggIAmAKgBBNBTwBhqBOU4hBgEhNBoBUhbgFhNhIBmAagGghBzBuAVA1gMgIgGARA5AQgShKAmAygigIBuAqgKhhhIBchHgLhLgohFAJAqgjBIAghdwaAIhmL1BggIhQgqgMhjBYAeA7whhIh2A0gQAKgohaBEgS" & _
            "hjgJAaA/BAuGg2BSA3glBJBWBWgVlnhIhNgPgkgJBqBYgXhmV3hcgUAjBJB6BggZhlBZgKBHgZAKgpBohfgbhnAJw9AchmgpA8AAgUBlAphORChnBZB2B5gehoAJhCB9gbAMhqAUhgghhogKAmrwBoBTA6BjgdhjBKA2gLsF2WAPgPgMh6Bkhho2hKB2gegnhpgTA+gXgNAqAaB+gmgphqAagMgqgpBqB6g2gcgqRNhCgrgmxfA6gwgdgpgKBkhiguhrhKxpgvhrB5AighgvhsBLAWhBgbANALAkhjgyhshLA2hOgzhsgTBCgqg0Bsh7BWhWg0JIgqhWg224BigxgPgtBLByhcg3gr1lBRg5hiA7gmhqg63IAchRgq2uAyhyg6huhrhChnA5SJhms+BvtnBfg5gugIBGh8pFyEBngqp9AqBVgIhwBDgWAjgVTnA8gzhEhLh8BUhLhFhuhjBLOTBxg7g7AXAyhxBSh8g7hIhQgThDAmgMBygChog9hKhLhsynhLgygyh3AoAuAzBMgYhwBphUgjBjA6ArBzpyA+hPh0ANAHA8ArBogGglABCPTDNugXzFJbzGzHCgALTIt0pgN2JhhqJiisARgwiuCvCwCxCyCzC0C1BpAEKBqChJBkqEgohCBRAJgWBLgjBAAohYASATARAGBCBBBxh5hcgKAHP5hLhlgihRhFACAbBCAcBMhggDBqBMBIBtA9m1hCAQgagVA+grAXgQACLGrHrIrJrKgyAIBcAMrNg" & _
            "yBFhphQAxBFBOhbA+AAFJBRAag0Alh3B2B5BeAdBXgnBNhyh7B+g/hgAwA3hLhTgqhwhyg7A+hwhNBMBah/Aag8hHhkAyA5AsgeBTAfgNBfhihagYBsAiQygEBCAhBQB6AtAegTBLgmhThogzhoBchqBXAJB9hVAxhYBsAWACh1B0hABeBtg3AbB9hOhzB3g+hdBtB2AsAdBWADBGhEBhAvhWhrBFgPApBlhNhngtA5hdAegHBLgkgxhjgvh0B4AEAWgJgOhEBhgshWAKhvWYhFhkB/h6A9A+BPWCBFheBhh3AQB9g9ATApgOhFBWgqgVAfgPhXhnh3hpB6Bfg4AFh9B+BfBTAVhxB2hdhDAwhKAyAKFrghApAKASB4giAnhJhYBMgcB/hegIJlAnBhACB6ANAoBggjAggCB3aUgIhQSuBSgAgcAmghgIg2AugLBigiBqAnAzBiaHglA1FdBEtNAJBYg+A+gLgjgyBWA4A0BjBYgUhaA2BNBiBogXAlAjhCBOBLAkGRBMgRgLBlAZBAgUAkBkhZAogUgRhNBpAVB/AgBmADgEheA3BkhYhsgWD2gCBaBmgLBmBpgiBqgaBkhZgIhdgYAIAWACB0A5gOBTgQgTAkAIhpAsgSgTa7A2uPBkhiBQgSgehIhphghkgdAXhWASB0A7gOhTgwgNAiLxBAgNAihJAaBel3gpgqB2gcAihpBiAugTAhhpADgyU8gYBmgOglgqhpAagwhsmcAKg2gyA7AaA2AZgI" & _
            "ApArgCANgfNiBKh1gkBnhrgCgRgIJWiQiDiYiDiUiDghgoAdAagQAWgY74AZgQb3gZiJ3+CDgY3/zGTGiggMYCpft1phCpYFCrYGYHJl4IpnYKJo4L4MgS4N4O4P4Q4R4S4T4U4V4W4X4Y4Z4a4b4c4d4e4f4g4h4i4j4k4l4m4nvsYpLLYq4rhN4s4t4u4v4w4x4y4z404142434445464748494+4/5A5B5C5D5E5F5G5H5I5JqJ5LZMZNZOZPZQZRZSZTZUZVZWZXZYZZZaZbZcZdZeZfZgZhZiZjZkZlZmZnZoZpZqZrZsZtZuZvZwZxAPhFhkhzh550BfBfB5gHhcBBh2h0T7A7AVBagqBUB9g/AfB/hPhzh3h+B+ggAHhMhFB7hdg3SVgygMgCgcgjAoBKBCgEgeAosgB2Q8AhA4BUgfgFAhDkgegCgKAYAc1hAdhIAUgCA0Vmg0ggD2gjh+VahihI6bAKhiByBqAkaegCAKAiaggjahgmai6jgPhj6l6m6ngI6paqarasatauGRBOBQgUglAZBOBQAlAJAZA2BKAlAkBzBaBFBfhIAZgEhhA3gNhZA2A7Al2MhggWw4BJgqBnsVsWhNgYhNhZgEgBBgAnBjgYhlA5MKhAgO8IgpA2gCVyg6AOgFghAoAJh0gOgegOBJhZB9BheIA4pIgIhSAkgHbJBSAsgRgmBpg6Bqgc2d7egYbdhSAegSRFhmARgGBo2lg2gTgrgOhKhCgygrhrBDgxgjB" & _
            "hgYgCgSg6AghZh2gixqBaBGB+g6ApAYgEeyCEb0AAb1AAb2AdAb74b5eBb6gQAZiAg"

        .Appearance = Sunken
        
        With .VisualAppearance
            .Add &H1, App.Path & "\Resources\" & "BackColorHeader.ebn"
            .Add &H2, App.Path & "\Resources\" & "SelBackColor.ebn"
        End With
        
        .BackColorHeader = &H1000000
        .SelBackColor = &H2000000
        .SelForeColor = RGB(205, 0, 0)
        .HeaderHeight = 22
        .Background(exBackColorFilter) = RGB(255, 255, 255)
        .Background(exContextMenuAppearance) = &H2000000
        .Background(exSelBackColorFilter) = &H1000000
        .DrawGridLines = exRowLines
        .GridLineStyle = exGridLinesSolid
        .PivotColumnsFloatBarVisible = True
        .PivotBarVisible = PivotBarVisibleEnum.exPivotBarContextSortAscending Or PivotBarVisibleEnum.exPivotBarAllowResizeColumns Or PivotBarVisibleEnum.exPivotBarAllowUndoRedo Or PivotBarVisibleEnum.exPivotBarAutoUpdate Or PivotBarVisibleEnum.exPivotBarAllowFormatContent Or PivotBarVisibleEnum.exPivotBarAllowFormatAppearance Or PivotBarVisibleEnum.exPivotBarAllowValues Or PivotBarVisibleEnum.exPivotBarShowTotals Or PivotBarVisibleEnum.exPivotBarSizable Or PivotBarVisibleEnum.exPivotBarVisible
        .AutoDrag = exAutoDragCopyText
        .FilterBarPromptVisible = True
        .LockRowsColumn = False
        .ShowDataOnDblClick = True
        .SingleSel = False
        .ShowViewCompact = exViewCompact
        .UseVisualTheme = UIVisualThemeEnum.exCalculatorVisualTheme Or UIVisualThemeEnum.exProgressVisualTheme Or UIVisualThemeEnum.exCheckBoxVisualTheme Or UIVisualThemeEnum.exSpinVisualTheme Or UIVisualThemeEnum.exSliderVisualTheme Or UIVisualThemeEnum.exCalendarVisualTheme Or UIVisualThemeEnum.exButtonsVisualTheme Or UIVisualThemeEnum.exFilterBarVisualTheme Or UIVisualThemeEnum.exHeaderVisualTheme

        With .FormatAppearances
            
            .Add("Red").BackColor = RGB(220, 20, 60)
            .Add("Pink").BackColor = RGB(255, 20, 147)
            .Add("Purple").BackColor = RGB(178, 58, 238)
            .Add("Blue").BackColor = RGB(0, 191, 255)
            .Add("Green").BackColor = RGB(144, 238, 144)
            .Add("Yellow").BackColor = RGB(238, 238, 0)
            .Add("Orange").BackColor = RGB(255, 165, 0)
            .Add("Gray").BackColor = RGB(211, 211, 211)
        
        End With
              
        With .FormatContents
            
            ' numeric: NumDigits|DecimalSep|Grouping|ThousandSep|NegativeOrder|LeadingZero
            .Remove "numeric"
            '.Add "numeric", "value format '0|.|3|,|3|0'", "Numeric"
             .Add "numeric", "value format '0|.|3|,|3|0'", "Numeric"
            ' currency: NumDigits|DecimalSep|Grouping|ThousandSep|NegativeOrder|LeadingZero
            .Remove "currency"
            '.Add "currency", "value format '2|.|3|,|1|0'", "Currency"
            .Add "currency", "value format '0|.|3|,|1|0'", "Currency"
            strDateFormat = "(('00' + day(date(value))) right 2)"
            strDateFormat = strDateFormat & " + '-' + (month(date(value)) case (default:'Err'; 1:'Jan'; 2:'Feb'; 3:'Mar'; 4:'Apr'; 5:'May'; 6:'Jun'; 7:'Jul'; 8:'Aug'; 9:'Sep'; 10:'Oct'; 11:'Nov'; 12:'Dec'))"
            strDateFormat = strDateFormat & " + '-' + year(date(value))"

            ' date
            .Remove "date"
            .Add "date", strDateFormat, "Date"

            strTimeFormat = "timeF(date(value)) left 5"

            ' time
            .Remove "time"
            .Add "time", strTimeFormat, "Time"

            ' datetime
            .Remove "datetime"
            .Add "datetime", "(" & strDateFormat & ") + ' ' + (" & strTimeFormat & ")", "Date Time"
        
        End With
        
        .EndUpdate
    
    End With


End Sub

Public Sub PivotPreview(ByRef paPivot As Pivot, ByRef paRecordset As ADODB.Recordset)
'************************************************
' Procedure         : PivotPreview
' Function          : Untuk Preview Pivot
' Input Parameter   : ByRef paPivot As Pivot, ByRef paRecordset As ADODB.Recordset
' Output Parameter  : Preview Pivot
' Programmer By     : -
' Date Update       : Jan-2016
' Update By         : Ryo/ Kreatif
'************************************************

    With paPivot
    
        .BeginUpdate

        .DataSource = paRecordset
        .PivotColumnsFloatBarVisible = True
        .ExpandAll

        .EndUpdate
    
    End With

End Sub

Public Sub PivotUserTemplate(ByRef paPivot As Pivot, ByRef paCommonDialog As CommonDialog, ByVal paFileUserTemplate As String, ByVal paIsUseUserTemplate As Boolean, tipe As String)
'************************************************
' Procedure         : PivotUserTemplate
' Function          : Tampilkan Template Pivot User dari direktori
' Input Parameter   : ByRef paPivot As Pivot, ByRef paCommonDialog As CommonDialog, ByVal paFileUserTemplate As String, ByVal paIsUseUserTemplate As Boolean, tipe As String
' Output Parameter  : Preview Pivot dengan menggunakan Template User
' Programmer By     : -
' Date Update       : Jan-2016
' Update By         : Ryo/ Kreatif
'************************************************

    Dim strMsg As String, strInitDir As String, strFilename As String, strResult As String
    Dim fsoTemplate As New Scripting.FileSystemObject

    On Local Error GoTo TrapErrorHere
    strMsg = "": strInitDir = "": strFilename = "": strResult = ""

    'Set initial directory.
    strInitDir = App.Path & "\TEMP\"

    'Check the file user template is exists or not, or load the custom template.
    strFilename = paFileUserTemplate
    
    If (fsoTemplate.FileExists(strFilename) = False) Or (paIsUseUserTemplate = False) Then
        
        With paCommonDialog
            .CancelError = True
            .DialogTitle = "Pick the template."
            .Flags = cdlOFNFileMustExist Or cdlOFNReadOnly Or cdlOFNHideReadOnly Or cdlOFNExplorer
            .Filter = tipe & " Files|*." & tipe
            .FilterIndex = 1
            .FileName = ""
            .InitDir = strInitDir
            .MaxFileSize = 1024
            .ShowOpen

            strFilename = .FileName
        End With
    
    End If

    'Write to file.
    Open strFilename For Input As #1
        Input #1, strResult
    Close #1

    'Set the layout of pivot.
    paPivot.Layout = strResult
    On Local Error GoTo 0

    GoTo ClearVarsThenExit

TrapErrorHere:
    MsgBox Err.Description, vbExclamation, strApplication_Name
    Err.Clear

ClearVarsThenExit:
    strMsg = "": strInitDir = "": strFilename = "": strResult = ""
    Set fsoTemplate = Nothing

End Sub

Public Sub PivotSaveTemplate(ByRef paPivot As Pivot, strRefTemplate As String, ByRef paCommonDialog As CommonDialog, ByVal paDefaultFileTemplate As String, ByVal paIsUseDefaultTemplate As Boolean, tipe As String)
'************************************************
' Procedure         : PivotSaveTemplate
' Function          : Menyimpan Template
' Input Parameter   : ByRef paPivot As Pivot, strRefTemplate As String, ByRef paCommonDialog As CommonDialog,
'                     ByVal paDefaultFileTemplate As String, ByVal paIsUseDefaultTemplate As Boolean, tipe As String
' Output Parameter  : Preview Pivot dengan menggunakan Template User
' Programmer By     : -
' Date Update       : Jan-2016
' Update By         : Ryo/ Kreatif
'************************************************
    
    Dim strMsg As String, strInitDir As String, strFilename As String

    On Local Error GoTo TrapErrorHere
    
    strMsg = "": strInitDir = "": strFilename = ""
    
    Dim objDialog As Object
 
    'self explanatory
    On Error Resume Next
    '=== Opening Save Dialog
    Set objDialog = CreateObject("MSComDlg.CommonDialog")
    objDialog.Filter = tipe & " Files (." & tipe & ")|*." & tipe
    objDialog.FileName = strFilename
    objDialog.DialogTitle = "Save Template"
    objDialog.InitDir = App.Path & "\TEMP\"
    objDialog.ShowSave
    
    If LenB(objDialog.FileName) = 0 Then Exit Sub
    
    Dim strFilename2 As String
    
    strFilename2 = objDialog.FileName
     
    strFilename = strFilename2
    
    If (paIsUseDefaultTemplate = False) Then
        
        With paCommonDialog
            .CancelError = True
            .DialogTitle = "Save the template as."
            .Flags = cdlOFNFileMustExist Or cdlOFNReadOnly Or cdlOFNHideReadOnly Or cdlOFNExplorer
            .Filter = tipe & " Files (.txt)|*." & tipe
            .FilterIndex = 1
            .FileName = ""
            .InitDir = App.Path & "\TEMP\"
            .MaxFileSize = 1024
            .ShowSave
            strFilename = .FileName
        End With
    
    End If

    'Save the layout of pivot.
    Open strFilename For Output As #1
        Print #1, paPivot.Layout
    Close #1
    
    '=== Direktory Checking not Exist Create One
    If Dir(App.Path & "\TEMP\") = "" Then
        MkDir App.Path & "\TEMP\"
    End If
    
     
    '=== Copy Template Into Default Location
    Dim ofsFilesSys As New FileSystemObject
    Dim ofsFile As File
    Dim strFileArr() As String
    Dim strGetFileName() As String
    
    strFileArr = Split(strFilename2, ":\")
    strGetFileName = Split(strFileArr(1), "\")
    
    Set ofsFile = ofsFilesSys.GetFile(strFilename2)
    
    ofsFile.Copy (App.Path & "\TEMP\" & strGetFileName(UBound(strGetFileName)))
     
    '=== Save User (insert) Template Into Database
    
    strSql = "DELETE FROM TEMP_PIVOT WHERE username='" & strLogin_User & "' "
    strSql = strSql & " AND filename='default.SUMP'"
    ConnERP.Execute strSql
    strSql = "INSERT INTO TEMP_PIVOT(username,report_code,temp_contain,filename) " & _
                       " VALUES('" & strLogin_User & "','" & tipe & "','" & paPivot.Layout & "','default.SUMP')"
    'MsgBox strSQL
    ConnERP.Execute strSql
               
        
    strRefTemplate = strGetFileName(UBound(strGetFileName))
    
    On Local Error GoTo 0

    GoTo ClearVarsThenExit

TrapErrorHere:
    
    MsgBox Err.Description, vbExclamation, strApplication_Name
    Err.Clear

ClearVarsThenExit:
    strMsg = "": strInitDir = "": strFilename = ""

End Sub

