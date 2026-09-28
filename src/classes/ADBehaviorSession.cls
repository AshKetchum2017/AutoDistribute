Option Explicit

' Class Module (Name) = ADBehaviorSession. Process selects the manual action path.
Private pMenu As Object
Private pDefaults As Object

Public Sub Start(ByVal script As String, ByVal observer As Object, ByVal token As String)
    Dim contract As New ADBehaviorContract, block As MRBehaviorBlock
    Dim ins As MRBehaviorInstruction, existing As Object, value As Variant
    Dim actions As Collection
    Dim operation As String, number As Long, source As String, description As String
    Dim manualActions As Boolean, closed As Boolean

    On Error GoTo Failed
    operation = "Memvalidasi kontrak AutoDistributeUF"
    Set block = contract.Validate(script)

    operation = "Memeriksa AutoDistributeMenu yang sudah terbuka"
    For Each existing In VBA.UserForms
        If StrComp(TypeName(existing), "AutoDistributeMenu", vbTextCompare) = 0 Then _
            Err.Raise 5, "ADBehaviorSession", "Tutup AutoDistributeMenu yang sudah terbuka sebelum menjalankan behavior."
    Next existing
    Set existing = Nothing

    operation = "Membuka AutoDistributeMenu"
    Set pMenu = New AutoDistributeMenu
    pMenu.MRBindRunner observer, token
    pMenu.Show vbModeless

    operation = "Membaca nilai awal kontrol AutoDistributeMenu"
    SnapshotDefaults block
    Set actions = New Collection
    For Each ins In block.Instructions
        operation = ins.FormName & " / " & ins.Target & " (posisi " & CStr(ins.Position) & ")"
        Select Case ins.Kind
            Case "stage", "skip"
                ' Omitted and skipped controls keep the form's initialized values.
            Case "action"
                actions.Add LCase$(ins.Target)
                If LCase$(ins.Target) = "cmdprocess" Then manualActions = True
            Case Else
                value = ins.Value
                If ins.Kind = "default" Then value = pDefaults(LCase$(ins.Target))
                pMenu.MRBehaviorValue ins.Target, value
        End Select
    Next ins
    operation = "Menyiapkan action AutoDistributeMenu"
    pMenu.MRPrepareBehaviorActions actions, Not manualActions
    If Not manualActions Then
        For Each ins In block.Instructions
            If ins.Kind = "action" Then
                operation = ins.FormName & " / @" & ins.Target & " (posisi " & CStr(ins.Position) & ")"
                pMenu.MRExecuteBehaviorAction ins.Target
                If LCase$(ins.Target) = "cmdclose" Then closed = True
            End If
        Next ins
        If Not closed Then pMenu.MRFinishAutomaticActions
    End If

    Set pMenu = Nothing
    Set pDefaults = Nothing
    Exit Sub

Failed:
    number = Err.Number: source = Err.Source: description = Err.Description
    On Error Resume Next
    If Not pMenu Is Nothing Then
        pMenu.MRAbortBehavior
        pMenu.MRDetachRunner
        If Not pMenu.Visible Then Unload pMenu
    End If
    Set pMenu = Nothing
    Set pDefaults = Nothing
    On Error GoTo 0
    Err.Raise number, "ADBehaviorSession", operation & vbCrLf & _
        "Source asli: " & source & vbCrLf & description
End Sub

Private Sub SnapshotDefaults(ByVal block As MRBehaviorBlock)
    Dim ins As MRBehaviorInstruction, target As String
    Set pDefaults = CreateObject("Scripting.Dictionary")
    For Each ins In block.Instructions
        If ins.Kind = "default" Then
            target = LCase$(ins.Target)
            If Not pDefaults.Exists(target) Then _
                pDefaults.Add target, pMenu.MRBehaviorReadValue(ins.Target)
        End If
    Next ins
End Sub
