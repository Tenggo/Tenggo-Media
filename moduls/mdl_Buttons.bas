Attribute VB_Name = "mdl_Buttons"
Option Explicit

Public Declare Function SetCapture Lib "user32" (ByVal hwnd As Long) As Long
Public Declare Function ReleaseCapture Lib "user32" () As Long
Public Declare Function GetCapture Lib "user32" () As Long
Public Enum enButtonType
    bieFirst = 0
    biePrev = 1
    bieNext = 2
    bieLast = 3
    bieAdd = 4
    bieEdit = 5
    bieDelete = 6
    biefind = 7
    biePrint = 8
    bieExit = 9
    bieSave = 10
    biecancel = 11
    bieHoliday = 12
    bieGenerate = 13
    bieView = 14
    bieDownload = 15
    bieClosing = 16
    bieOK = 17
    bieUserTemplate = 18
    bieSaveTemplate = 19
    bieDefaultTemplate = 20
    biePreview = 21
    bieExportExcel = 22
    bieClose = 23
    bieCatalog = 24
    bieProjectBrief = 25
    bieReport = 26
    bieExportExcel2 = 27
    bietTBReport = 28
    bieRun = 29
    bieRunDownload = 30
    bieMatchBooking = 31
    bieGetMediaOrder = 32
    bieGetNielsenSpot = 33
    bieLockData = 34
    bieUnlockdata = 35
    bieDemographicProcess = 36
    bieSpotmatch = 37
    biePostBuyAnalysis = 38
    bieReportSpotmatch = 39
    bieGetData = 40
    bieGetDataObject = 41
    bieCustomeDayPart = 42
    bieUpdate = 43
    bieUnMonitoredChannel = 44
    bieSetJob = 45
    bieRemoveJob = 46
    bieInsert = 47
    bieAutomatch = 48
    bieMatchWithOp = 49
    bieMatchFree = 50
    bieManualMatch = 51
    bieRelNielsenSpot = 52
    bieSpotFromOtherChannel = 53
    bieUnMatchAll = 54
    bieUnMatch = 55
    bieUnMatchJobBJob = 56
    bieInserJobToGroup = 57
    bieRemoveJobFromGrp = 58
    bieReset = 59
    bieCopy = 60
    bieTVLayering = 61
    bieSummary = 62
    bieResendEmail = 63
    biePublishToWeb = 64
    bieCreate = 65
    bieApprove = 66
    bieRevisionHistory = 67
    bieApprovedQuotationList = 68
    bieUpdateTVR = 69
    bieClear = 70
    bieSendEmail = 71
    bieUploadSRI = 90
    bieSaveAsOriginal = 91
    bieCheckBudget = 92
    bieAddProgramManualy = 93
    bieLoadProgramFromExcel = 94
    bieUploadValidSpotToTVSchedule = 95
End Enum

Public Enum enButtonEffect
    bieNormal = 0
    bieOver = 1
    bieDown = 2
    bieDisabled = 3
End Enum

Public Function SetButtonImageEffect(ByVal paButtonType As enButtonType, ByVal paButtonEffect As enButtonEffect) As String ' TOOLBAR_AI.
'************************************************
' Procedure         : SetButtonImageEffect
' Function          : TOOLBAR_AI utk menentukan file gambar mana yg akan digunakan pada suatu button.
'Input Parameter    : ~ paButtonType: Refer to enum enButtonType declaration at the top.
'                     ~ paButtonType: Refer to enum enButtonEffect declaration at the top.
'Output Parameter   : String path of image file.
' Created By        : {73 64 6B}
' Date              : 07-Apr-2015
'************************************************
    Dim strButtonEffect As String, strImagePath As String, strImageFullPath As String, strDummyImage As String
    Dim oFSO As Scripting.FileSystemObject
    Dim blnIsExists As Boolean

    strButtonEffect = "": strImagePath = "": strImageFullPath = "": strDummyImage = ""

    Select Case paButtonEffect
        Case bieNormal: strButtonEffect = ""
        Case bieOver: strButtonEffect = "_a"
        Case bieDown: strButtonEffect = "_ab"
        Case bieDisabled: strButtonEffect = "_abc"
    End Select
    strButtonEffect = strButtonEffect & ".jpg"

    strImagePath = App.Path & "\Resources\"

    Select Case paButtonType
        Case bieFirst: strImageFullPath = strImagePath & "first" & strButtonEffect
        Case biePrev: strImageFullPath = strImagePath & "prev" & strButtonEffect
        Case bieNext: strImageFullPath = strImagePath & "next" & strButtonEffect
        Case bieLast: strImageFullPath = strImagePath & "last" & strButtonEffect
        Case bieAdd: strImageFullPath = strImagePath & "add" & strButtonEffect
        Case bieEdit: strImageFullPath = strImagePath & "edit" & strButtonEffect
        Case bieDelete: strImageFullPath = strImagePath & "delete" & strButtonEffect
        Case biefind: strImageFullPath = strImagePath & "find" & strButtonEffect
        Case biePrint: strImageFullPath = strImagePath & "print" & strButtonEffect
        Case bieExit: strImageFullPath = strImagePath & "exit" & strButtonEffect
        Case bieSave: strImageFullPath = strImagePath & "save" & strButtonEffect
        Case biecancel: strImageFullPath = strImagePath & "cancel" & strButtonEffect
        Case bieHoliday: strImageFullPath = strImagePath & "holiday" & strButtonEffect
        Case bieGenerate: strImageFullPath = strImagePath & "generate" & strButtonEffect
        Case bieView: strImageFullPath = strImagePath & "view" & strButtonEffect
        Case bieDownload: strImageFullPath = strImagePath & "download" & strButtonEffect
        Case bieClosing: strImageFullPath = strImagePath & "monthly_closing" & strButtonEffect
        Case bieOK: strImageFullPath = strImagePath & "ok" & strButtonEffect
        Case bieUserTemplate: strImageFullPath = strImagePath & "load_template" & strButtonEffect
        Case bieSaveTemplate: strImageFullPath = strImagePath & "save_template" & strButtonEffect
        Case bieDefaultTemplate: strImageFullPath = strImagePath & "reset_template" & strButtonEffect
        Case biePreview: strImageFullPath = strImagePath & "preview" & strButtonEffect
        Case bieExportExcel: strImageFullPath = strImagePath & "export" & strButtonEffect
        Case bieClose: strImageFullPath = strImagePath & "close" & strButtonEffect
        Case biePreview: strImageFullPath = strImagePath & "preview" & strButtonEffect
        Case bieCatalog: strImageFullPath = strImagePath & "catalogue" & strButtonEffect
        Case bieProjectBrief: strImageFullPath = strImagePath & "project-brief" & strButtonEffect
        Case bieReport: strImageFullPath = strImagePath & "report" & strButtonEffect
        Case bieExportExcel2: strImageFullPath = strImagePath & "export_excel_2" & strButtonEffect
        Case bietTBReport: strImageFullPath = strImagePath & "tbreport" & strButtonEffect
        Case bieRun: strImageFullPath = strImagePath & "run" & strButtonEffect
        Case bieRunDownload: strImageFullPath = strImagePath & "run_process" & strButtonEffect
        Case bieMatchBooking: strImageFullPath = strImagePath & "match_booking" & strButtonEffect
        Case bieGetMediaOrder: strImageFullPath = strImagePath & "get_media_order" & strButtonEffect
        Case bieGetNielsenSpot: strImageFullPath = strImagePath & "get_product_spot" & strButtonEffect
        Case bieLockData: strImageFullPath = strImagePath & "lock_data" & strButtonEffect
        Case bieUnlockdata: strImageFullPath = strImagePath & "unlock_data" & strButtonEffect
        Case bieDemographicProcess: strImageFullPath = strImagePath & "demographic_process" & strButtonEffect
        Case bieSpotmatch: strImageFullPath = strImagePath & "spot_matching" & strButtonEffect
        Case biePostBuyAnalysis: strImageFullPath = strImagePath & "post_buy_analysis" & strButtonEffect
        Case bieReportSpotmatch: strImageFullPath = strImagePath & "report" & strButtonEffect
        Case bieGetData: strImageFullPath = strImagePath & "get_data" & strButtonEffect
        Case bieGetDataObject: strImageFullPath = strImagePath & "getdata_obj" & strButtonEffect
        Case bieUpdate: strImageFullPath = strImagePath & "update" & strButtonEffect
        Case bieSetJob: strImageFullPath = strImagePath & "set_job_number" & strButtonEffect
        Case bieRemoveJob: strImageFullPath = strImagePath & "remove_job_number" & strButtonEffect
        Case bieInsert: strImageFullPath = strImagePath & "insert" & strButtonEffect
        Case bieAutomatch: strImageFullPath = strImagePath & "auto_match" & strButtonEffect
        Case bieMatchWithOp: strImageFullPath = strImagePath & "auto_match_with_option" & strButtonEffect
        Case bieMatchFree: strImageFullPath = strImagePath & "match_free" & strButtonEffect
        Case bieManualMatch: strImageFullPath = strImagePath & "manual_match" & strButtonEffect
        Case bieRelNielsenSpot: strImageFullPath = strImagePath & "release_nielsen_spot" & strButtonEffect
        Case bieSpotFromOtherChannel: strImageFullPath = strImagePath & "add_spot_from_other_channel" & strButtonEffect
        Case bieUnMatchAll: strImageFullPath = strImagePath & "unmatch_all" & strButtonEffect
        Case bieUnMatch: strImageFullPath = strImagePath & "unmatch" & strButtonEffect
        Case bieUnMatchJobBJob: strImageFullPath = strImagePath & "unmatch_job_by_job" & strButtonEffect
        Case bieUnMonitoredChannel: strImageFullPath = strImagePath & "unmonitored_channel" & strButtonEffect
        Case bieCustomeDayPart: strImageFullPath = strImagePath & "customize_daypart" & strButtonEffect
        Case bieInserJobToGroup: strImageFullPath = strImagePath & "insert_job_to_group" & strButtonEffect
        Case bieRemoveJobFromGrp: strImageFullPath = strImagePath & "remove_job_from_group" & strButtonEffect
        Case bieReset: strImageFullPath = strImagePath & "reset" & strButtonEffect
        Case bieCopy: strImageFullPath = strImagePath & "copy" & strButtonEffect
        Case bieTVLayering: strImageFullPath = strImagePath & "tv_layering" & strButtonEffect
        Case bieSummary: strImageFullPath = strImagePath & "budget_summary" & strButtonEffect
        Case bieResendEmail: strImageFullPath = strImagePath & "resend_email" & strButtonEffect
        Case biePublishToWeb: strImageFullPath = strImagePath & "publish_to_web" & strButtonEffect
        Case bieCreate: strImageFullPath = strImagePath & "create" & strButtonEffect
        Case bieClear: strImageFullPath = strImagePath & "clear" & strButtonEffect
        Case bieSendEmail: strImageFullPath = strImagePath & "send" & strButtonEffect
        Case bieApprove: strImageFullPath = strImagePath & "approve" & strButtonEffect
        Case bieRevisionHistory: strImageFullPath = strImagePath & "revision_history" & strButtonEffect
        Case bieApprovedQuotationList: strImageFullPath = strImagePath & "approve_quotation" & strButtonEffect
        Case bieUpdateTVR: strImageFullPath = strImagePath & "update_tvr" & strButtonEffect
        Case bieUploadSRI: strImageFullPath = strImagePath & "upload_sri" & strButtonEffect
        Case bieSaveAsOriginal: strImageFullPath = strImagePath & "save_as_original" & strButtonEffect
        Case bieCheckBudget: strImageFullPath = strImagePath & "check_budget" & strButtonEffect
   End Select

    strDummyImage = strImagePath & "dummy" & strButtonEffect

    Set oFSO = New Scripting.FileSystemObject
    blnIsExists = oFSO.FileExists(strImageFullPath)
    If blnIsExists Then
        SetButtonImageEffect = strImageFullPath
    Else: SetButtonImageEffect = strDummyImage
    End If
    Set oFSO = Nothing

    strButtonEffect = "": strImagePath = "": strImageFullPath = "": strDummyImage = ""

End Function

Public Sub SetPictureTB(ByVal Index As Integer, ByVal paIsNormalMode As Boolean, picOBJ)
'************************************************
' Procedure         : SetPictureTB
' Function          : Untuk set Gambar PicButon Normal atau Disable.
' Parameter Input   : Index paIsNormalMode picOBJ
' Parameter Output  : Object gambar
' Last Update/By    : Tedi
'************************************************

   With picOBJ(Index) 'FIRST.
        
        If .Enabled = True Then
            .Picture = LoadPicture(SetButtonImageEffect(Index, bieNormal))
        Else: .Picture = LoadPicture(SetButtonImageEffect(Index, bieDisabled))
        End If
    End With

End Sub

Public Sub picButton_Obj(ByRef Index As Integer, ByRef Button As Integer, ByRef Shift As Integer, ByRef X As Single, ByRef Y As Single, ByRef picOBJ As Object)  'TOOLBAR_AI.
'************************************************
' Procedure         : picButton_Obj
' Function          : Manipulasi Object Hover dan Normal
' Created By        : {73 64 6B}
' Date              : 12-Apr-2015
' last Update       :
' last Update By    : Tedi / Kreatif
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
