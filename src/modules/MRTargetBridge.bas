Option Explicit

' Standard Module (Name) = MRTargetBridge in AutoDistributeUF only.
' This project creates its known form directly; do not copy to other GMS projects.
' Form code must contain MRBindRunner, MRDetachRunner and UserForm_Terminate.
' No reference to the MacroRunner VBA project is needed.
Public Function OpenMacro(ByVal formName As String, ByVal modal As Boolean, _
                          ByVal observer As Object, ByVal token As String) As Boolean
    Dim target As Object, existing As Object
    Dim errorNumber As Long, errorDescription As String, errorSource As String, operation As String
    On Error GoTo Failed
    operation = "Memvalidasi form tujuan AutoDistributeUF"
    If StrComp(formName, "AutoDistributeMenu", vbTextCompare) <> 0 Then _
        Err.Raise 5, "MRTargetBridge", "Form tidak terdaftar dalam AutoDistributeUF: " & formName
    operation = "Memeriksa form yang sudah terbuka"
    For Each existing In VBA.UserForms
        If StrComp(TypeName(existing), formName, vbTextCompare) = 0 Then _
            Err.Raise 5, "MRTargetBridge", "Tutup " & formName & " yang sudah terbuka sebelum menjalankan antrean."
    Next existing
    Set existing = Nothing
    operation = "Membuat instance New AutoDistributeMenu"
    ' Bind the form class in this project without dynamic UserForms.Add lookup.
    Set target = New AutoDistributeMenu
    operation = "Menghubungkan event penutupan " & formName
    CallByName target, "MRBindRunner", VbMethod, observer, token
    operation = "Membuka " & formName
    If modal Then
        target.Show vbModal
    Else
        target.Show vbModeless
    End If
    ' Do not retain a form reference: it would postpone UserForm_Terminate.
    Set target = Nothing
    OpenMacro = True
    Exit Function
Failed:
    errorNumber = Err.Number
    errorSource = Err.Source
    errorDescription = Err.Description
    On Error Resume Next
    If Not target Is Nothing Then
        CallByName target, "MRDetachRunner", VbMethod
        Unload target
    End If
    Set target = Nothing
    On Error GoTo 0
    Err.Raise errorNumber, "MRTargetBridge.OpenMacro", operation & vbCrLf & _
        "Source awal: " & errorSource & vbCrLf & errorDescription
End Function

' Semantic preflight only: no UserForm or document access.
Public Function ValidateBehavior(ByVal script As String, ByVal observer As Object, ByVal token As String) As Boolean
    Dim contract As ADBehaviorContract, block As MRBehaviorBlock
    Dim number As Long, source As String, description As String

    On Error GoTo Failed
    CallByName observer, "BehaviorBridgeEntered", VbMethod, token
    Set contract = New ADBehaviorContract
    Set block = contract.Validate(script)
    CallByName observer, "BehaviorBridgeFinished", VbMethod, token, 0&, vbNullString
    ValidateBehavior = True
    Exit Function
Failed:
    number = Err.Number: source = Err.Source: description = Err.Description
    On Error Resume Next
    CallByName observer, "BehaviorBridgeFinished", VbMethod, token, number, _
        "Source asli: " & source & vbCrLf & description
    On Error GoTo 0
    ValidateBehavior = False
End Function

Public Function RunBehavior(ByVal script As String, ByVal observer As Object, ByVal token As String) As Boolean
    Dim session As ADBehaviorSession
    Dim number As Long, source As String, description As String

    On Error GoTo Failed
    CallByName observer, "BehaviorBridgeEntered", VbMethod, token
    Set session = New ADBehaviorSession
    session.Start script, observer, token
    CallByName observer, "BehaviorBridgeFinished", VbMethod, token, 0&, vbNullString
    RunBehavior = True
    Exit Function
Failed:
    number = Err.Number: source = Err.Source: description = Err.Description
    On Error Resume Next
    CallByName observer, "BehaviorBridgeFinished", VbMethod, token, number, _
        "Source asli: " & source & vbCrLf & description
    On Error GoTo 0
    ' The runner receives the failure through its token callback.
    RunBehavior = False
End Function
