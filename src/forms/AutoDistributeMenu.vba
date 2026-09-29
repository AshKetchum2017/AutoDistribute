Option Explicit

' MacroRunner integration: no reference to the runner project is required.
Private pMRObserver As Object
Private pMRToken As String
Private pMRBehaviorActions As Collection
Private pMRActionIndex As Long
Private pMRBehaviorActive As Boolean
Private pMRBehaviorFailed As Boolean
Private pMRBehaviorAutomatic As Boolean


Private Sub chkCCWRotate90_Click()
    If chkCCWRotate90.Value Then chkCWRotate90.Value = False
End Sub

Private Sub chkCWRotate90_Click()
    If chkCWRotate90.Value Then chkCCWRotate90.Value = False
End Sub

Private Sub chkSequentially_Click()

End Sub

Private Sub cmdClose_Click()
    Unload Me
End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    If pMRBehaviorActive And Not pMRBehaviorFailed Then
        If pMRActionIndex <= pMRBehaviorActions.Count Then
            If pMRBehaviorActions(pMRActionIndex) = "cmdclose" Then
                pMRActionIndex = pMRActionIndex + 1
            Else
                MRReportBehaviorFailure 5, "AutoDistributeMenu.QueryClose", _
                    "Form ditutup sebelum action @" & pMRBehaviorActions(pMRActionIndex) & " dijalankan."
            End If
        End If
    End If
    ADClearObjectDraft Me
End Sub

Private Sub ADApplyWorksheetSize(ByVal widthMM As Double, ByVal heightMM As Double, _
    Optional ByVal sensorMode As String = "", Optional ByVal reportToRunner As Boolean = False, _
    Optional ByVal manageCommandGroup As Boolean = True)
    ADApplyPageSetup widthMM, heightMM, sensorMode, reportToRunner, manageCommandGroup
End Sub

Private Sub cmdDefaultSize_Click()
    ADApplyWorksheetSize 325#, 485#
End Sub

Private Sub cmdExtendedSize_Click()
    ADApplyWorksheetSize 335#, 487#
End Sub

Private Sub cmdMasterPage_Click()
    If Not MRCanRunAction("cmdmasterpage") Then Exit Sub
    On Error GoTo ActionFailed
    ADApplyWorksheetSize 325#, 485#, "Master", pMRBehaviorActive And Not pMRBehaviorFailed
    MRCompleteAction "cmdmasterpage"
    Exit Sub
ActionFailed:
    MRHandleActionFailure "cmdMasterPage", Err.Number, Err.Source, Err.Description
End Sub

Private Sub cmdPerPage_Click()
    If Not MRCanRunAction("cmdperpage") Then Exit Sub
    On Error GoTo ActionFailed
    ADApplyWorksheetSize 325#, 485#, "PerPage", pMRBehaviorActive And Not pMRBehaviorFailed
    MRCompleteAction "cmdperpage"
    Exit Sub
ActionFailed:
    MRHandleActionFailure "cmdPerPage", Err.Number, Err.Source, Err.Description
End Sub

Private Sub cmdProcess_Click()
    Dim doc As Document
    Dim commandGroupOpen As Boolean
    Dim errorNumber As Long
    Dim errorSource As String
    Dim errorDescription As String
    Dim processResult As Variant

    If Not MRCanRunAction("cmdprocess") Then Exit Sub
    On Error GoTo ProcessFailed

    If Not optKissA.Value And Not optDieA.Value Then
        MsgBox "Pilih mode Kiss A atau Die A terlebih dahulu.", vbExclamation, "Auto Distribute"
        Exit Sub
    End If

    Set doc = ActiveDocument
    doc.BeginCommandGroup "Auto Distribute Process"
    commandGroupOpen = True
    If optKissA.Value Then
        ADApplyWorksheetSize 335#, 487#, vbNullString, True, False
    ElseIf optDieA.Value Then
        ADApplyWorksheetSize 325#, 485#, vbNullString, True, False
    End If

    ADProcessStoredObjects optKissA.Value, optDieA.Value, chkCWRotate90.Value, _
        chkCCWRotate90.Value, chkSequentially.Value, False, False, processResult
    doc.Unit = cdrMillimeter
    doc.Rulers.HUnits = cdrMillimeter
    doc.Rulers.VUnits = cdrMillimeter
    doc.EndCommandGroup
    commandGroupOpen = False
    MsgBox CStr(processResult(0)) & " object selesai diproses.", vbInformation, "Auto Distribute"
    If Len(CStr(processResult(1))) > 0 Then MsgBox _
        "Layer berikut masih berisi object dan tidak dihapus:" & vbCrLf & CStr(processResult(1)), _
        vbExclamation, "Auto Distribute"
    MRCompleteAction "cmdprocess"
    Exit Sub

ProcessFailed:
    errorNumber = Err.Number
    errorSource = Err.Source
    errorDescription = Err.Description
    On Error Resume Next
    If commandGroupOpen Then doc.EndCommandGroup
    On Error GoTo 0
    MRHandleActionFailure "cmdProcess", errorNumber, errorSource, errorDescription

End Sub

Private Sub optDieA_Click()
    If optDieA.Value Then optKissA.Value = False

End Sub

Private Sub optKissA_Click()
    If optKissA.Value Then optDieA.Value = False

End Sub

Private Sub cmdSetDesign_Click()
    ADAddSelectedObjectsToContainer "Design", Me
End Sub

Private Sub cmdSetCutLine_Click()
    ADAddSelectedObjectsToContainer "Cut Line", Me
End Sub

Private Sub cmdRemove_Click()
    ADRemoveObjectDraft Me
End Sub

Private Sub cmdClear_Click()
    ADClearObjectDraft Me
End Sub

Private Sub UserForm_Initialize()
    ADRefreshObjectContainerList Me
    If Not optKissA.Value And Not optDieA.Value Then
        optKissA.Value = True
    End If
End Sub

' Set Design/Cut Line stays manual when @cmdProcess is present.
Public Sub MRPrepareBehaviorActions(ByVal actions As Collection, ByVal automatic As Boolean)
    Set pMRBehaviorActions = actions
    pMRActionIndex = 1
    pMRBehaviorFailed = False
    pMRBehaviorActive = True
    pMRBehaviorAutomatic = automatic
End Sub

Public Sub MRExecuteBehaviorAction(ByVal action As String)
    Dim previousIndex As Long
    If Not pMRBehaviorAutomatic Then Err.Raise 5, "AutoDistributeMenu.MRExecuteBehaviorAction", _
        "Action otomatis belum disiapkan."
    previousIndex = pMRActionIndex
    Select Case LCase$(action)
        Case "cmdmasterpage": cmdMasterPage_Click
        Case "cmdperpage": cmdPerPage_Click
        Case "cmdclose": cmdClose_Click
        Case Else: Err.Raise 5, "AutoDistributeMenu.MRExecuteBehaviorAction", _
            "Action otomatis tidak terdaftar: @" & action
    End Select
    If LCase$(action) <> "cmdclose" Then
        If pMRActionIndex = previousIndex Then Err.Raise 5, "AutoDistributeMenu.MRExecuteBehaviorAction", _
            "Action @" & action & " belum selesai."
    End If
End Sub

Public Sub MRFinishAutomaticActions()
    pMRBehaviorAutomatic = False
    Set pMRBehaviorActions = New Collection
    pMRActionIndex = 1
End Sub

Public Sub MRAbortBehavior()
    pMRBehaviorAutomatic = False
    pMRBehaviorActive = False
    pMRBehaviorFailed = True
    Set pMRBehaviorActions = Nothing
End Sub

Public Sub MRBehaviorValue(ByVal target As String, ByVal value As Variant)
    Select Case LCase$(target)
        Case "optkissa"
            optKissA.Value = CBool(value)
            If optKissA.Value Then optDieA.Value = False
        Case "optdiea"
            optDieA.Value = CBool(value)
            If optDieA.Value Then optKissA.Value = False
        Case "chkcwrotate90"
            chkCWRotate90.Value = CBool(value)
            If chkCWRotate90.Value Then chkCCWRotate90.Value = False
        Case "chkccwrotate90"
            chkCCWRotate90.Value = CBool(value)
            If chkCCWRotate90.Value Then chkCWRotate90.Value = False
        Case "chksequentially": chkSequentially.Value = CBool(value)
        Case Else: Err.Raise 5, "AutoDistributeMenu.MRBehaviorValue", "Target tidak terdaftar: " & target
    End Select
End Sub

Public Function MRBehaviorReadValue(ByVal target As String) As Variant
    Select Case LCase$(target)
        Case "optkissa": MRBehaviorReadValue = optKissA.Value
        Case "optdiea": MRBehaviorReadValue = optDieA.Value
        Case "chkcwrotate90": MRBehaviorReadValue = chkCWRotate90.Value
        Case "chkccwrotate90": MRBehaviorReadValue = chkCCWRotate90.Value
        Case "chksequentially": MRBehaviorReadValue = chkSequentially.Value
        Case Else: Err.Raise 5, "AutoDistributeMenu.MRBehaviorReadValue", "Default tidak tersedia: " & target
    End Select
End Function

Private Function MRCanRunAction(ByVal action As String) As Boolean
    If Not pMRBehaviorActive Or pMRBehaviorFailed Then
        MRCanRunAction = True
        Exit Function
    End If
    If pMRBehaviorActions.Count = 0 Then
        MRCanRunAction = True
        Exit Function
    End If
    If pMRActionIndex <= pMRBehaviorActions.Count Then
        If pMRBehaviorActions(pMRActionIndex) = action Then
            MRCanRunAction = True
            Exit Function
        End If
        MsgBox "Action berikutnya dalam behavior: @" & pMRBehaviorActions(pMRActionIndex) & ".", _
            vbExclamation, "Auto Distribute"
    Else
        MsgBox "Seluruh action behavior selesai. Tutup form untuk melanjutkan antrean.", _
            vbInformation, "Auto Distribute"
    End If
End Function

Private Sub MRCompleteAction(ByVal action As String)
    If Not pMRBehaviorActive Or pMRBehaviorFailed Then Exit Sub
    pMRActionIndex = pMRActionIndex + 1
    If pMRBehaviorAutomatic Then Exit Sub
    If action <> "cmdmasterpage" And action <> "cmdperpage" Then Exit Sub
    If pMRActionIndex > pMRBehaviorActions.Count Then Exit Sub
    If pMRBehaviorActions(pMRActionIndex) = "cmdclose" Then Unload Me
End Sub

Private Sub MRHandleActionFailure(ByVal action As String, ByVal number As Long, _
    ByVal source As String, ByVal description As String)
    If pMRBehaviorActive And Not pMRBehaviorFailed Then
        If pMRBehaviorAutomatic Then
            MRAbortBehavior
            If number = 0 Then number = 5
            Err.Raise number, "AutoDistributeMenu." & action, _
                "Source asli: " & source & vbCrLf & description
        Else
            MRReportBehaviorFailure number, source, "Action @" & action & vbCrLf & description
        End If
    Else
        MsgBox "Error " & CStr(number) & vbCrLf & "Description: [" & description & "]", _
            vbCritical, "Auto Distribute"
    End If
End Sub

Private Sub MRReportBehaviorFailure(ByVal number As Long, ByVal source As String, _
    ByVal description As String)
    Dim observer As Object
    Dim token As String
    If number = 0 Then number = 5
    pMRBehaviorFailed = True
    pMRBehaviorActive = False
    Set observer = pMRObserver
    token = pMRToken
    MRDetachRunner
    If observer Is Nothing Then Exit Sub
    On Error GoTo NotifyFailed
    CallByName observer, "BehaviorFailed", VbMethod, token, number, _
        "Source asli: " & source & vbCrLf & description
    Exit Sub
NotifyFailed:
    MsgBox "Gagal melaporkan error ke Macro Runner (" & CStr(Err.Number) & "): " & _
        Err.Description, vbExclamation, "Macro Runner"
End Sub

' Called only by MRTargetBridge; normal menu entry points remain unchanged.
Public Sub MRBindRunner(ByVal observer As Object, ByVal token As String)
    Set pMRObserver = observer
    pMRToken = token
End Sub

Public Sub MRDetachRunner()
    Set pMRObserver = Nothing
    pMRToken = vbNullString
End Sub

Private Sub UserForm_Terminate()
    Dim observer As Object, token As String
    On Error GoTo NotifyFailed
    Set observer = pMRObserver
    token = pMRToken
    MRDetachRunner
    If Not observer Is Nothing Then CallByName observer, "MacroUnloaded", VbMethod, token
    Exit Sub
NotifyFailed:
    MsgBox "Gagal memberitahu Macro Runner bahwa form sudah ditutup (" & CStr(Err.Number) & "): " & _
        Err.Description, vbExclamation, "Macro Runner"
End Sub
