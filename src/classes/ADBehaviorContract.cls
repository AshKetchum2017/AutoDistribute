Option Explicit

' Class Module (Name) = ADBehaviorContract. Pure semantic preflight.
Private pText As String
Private pAt As Long

Public Function Validate(ByVal script As String) As MRBehaviorBlock
    Dim parser As New MRBehaviorParser
    Dim blocks As Collection, block As MRBehaviorBlock
    Dim ins As MRBehaviorInstruction
    Dim stageSeen As Boolean, processed As Boolean, actionSeen As Boolean
    Dim closed As Boolean

    pText = script
    pAt = 1
    Set blocks = parser.Parse(script, Array("AutoDistributeUF"))
    Set block = blocks(1)

    For Each ins In block.Instructions
        pAt = ins.Position
        If closed Then Fail "Tidak boleh ada instruksi setelah @cmdClose."
        If StrComp(ins.FormName, "AutoDistributeMenu", vbTextCompare) <> 0 Then _
            Fail "UserForm tidak terdaftar: " & ins.FormName

        Select Case ins.Kind
            Case "stage"
                If stageSeen Then Fail "AutoDistributeMenu hanya boleh mempunyai satu tahap."
                stageSeen = True
            Case "action"
                Select Case LCase$(ins.Target)
                    Case "cmdprocess"
                        If processed Then Fail "@cmdProcess hanya boleh dicantumkan sekali."
                        processed = True
                        actionSeen = True
                    Case "cmdmasterpage", "cmdperpage"
                        actionSeen = True
                    Case "cmdclose"
                        closed = True
                        actionSeen = True
                    Case Else
                        Fail "Action tidak terdaftar: AutoDistributeMenu.@" & ins.Target
                End Select
            Case Else
                If actionSeen Then Fail "Pengaturan kontrol harus mendahului action."
                ValidateValue ins
        End Select
    Next ins

    pAt = Len(pText) + 1
    Set Validate = block
End Function

Private Sub ValidateValue(ByVal ins As MRBehaviorInstruction)
    Select Case LCase$(ins.Target)
        Case "optkissa", "optdiea", "chkcwrotate90", "chkccwrotate90", "chksequentially"
            If ins.Kind = "skip" Or ins.Kind = "default" Then Exit Sub
            If ins.Kind <> "boolean" Then Fail ins.Target & " harus True atau False."
        Case Else
            Fail "Target tidak terdaftar: AutoDistributeMenu." & ins.Target
    End Select
End Sub

Private Sub Fail(ByVal message As String)
    Dim prefix As String, line As Long, column As Long, lastBreak As Long
    prefix = Left$(pText, pAt - 1)
    line = UBound(Split(prefix, vbLf)) + 1
    lastBreak = InStrRev(prefix, vbLf)
    column = pAt - lastBreak
    Err.Raise 5, "ADBehaviorContract", "Semantic AutoDistributeUF baris " & CStr(line) & _
        ", kolom " & CStr(column) & ": " & message
End Sub
