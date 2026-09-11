Option Explicit

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
    ADClearObjectDraft Me
End Sub

Private Sub cmdDefaultSize_Click()
    ADApplyPageSetup 325#, 485#
End Sub

Private Sub cmdExtendedSize_Click()
    ADApplyPageSetup 335#, 487#
End Sub

Private Sub cmdMasterPage_Click()
    ADApplyPageSetup 325#, 485#, "Master"
End Sub

Private Sub cmdPerPage_Click()
    ADApplyPageSetup 325#, 485#, "PerPage"
End Sub

Private Sub cmdProcess_Click()
    On Error GoTo ProcessFailed

    If Not optKissA.Value And Not optDieA.Value Then
        MsgBox "Pilih mode Kiss A atau Die A terlebih dahulu.", vbExclamation, "Auto Distribute"
        Exit Sub
    End If

    ADProcessStoredObjects optKissA.Value, optDieA.Value, chkCWRotate90.Value, _
        chkCCWRotate90.Value, chkSequentially.Value
    Exit Sub

ProcessFailed:
    MsgBox "Error " & Err.Number & vbCrLf & _
        "Description: [" & Err.Description & "]", _
        vbCritical, "Auto Distribute"

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
