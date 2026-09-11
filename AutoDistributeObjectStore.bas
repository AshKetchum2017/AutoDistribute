Option Explicit
Option Private Module

Private Type ADObjectEntry
    RoleLabel As String
    StaticID As Long
End Type

Private Const AD_ROLE_DESIGN As String = "Design"
Private Const AD_ROLE_CUT_LINE As String = "Cut Line"

Private mObjectEntries() As ADObjectEntry
Private mObjectCount As Long
Private mDraftDocument As Document

Public Sub ADApplyPageSetup(ByVal widthMM As Double, ByVal heightMM As Double, _
    Optional ByVal sensorMode As String = "")
    Dim doc As Document
    Dim startPage As Page
    Dim destinationPage As Page
    Dim pg As Page
    Dim ly As Layer
    Dim sensorLayers As Collection
    Dim oldUnit As cdrUnit
    Dim unitSaved As Boolean
    Dim commandGroupOpen As Boolean
    Dim i As Long
    Dim operation As String
    Dim layerName As String
    Dim movedLayer As Layer
    Dim sourceShapeCount As Long
    Dim errorNumber As Long
    Dim errorDescription As String

    On Error GoTo SetupFailed
    operation = "Validasi dokumen dan ukuran page"
    If Application.Documents.Count = 0 Then Err.Raise 5, , "Tidak ada dokumen aktif."
    If widthMM <= 0 Or heightMM <= 0 Then Err.Raise 5, , "Ukuran page harus lebih besar dari nol."
    If sensorMode <> "" And sensorMode <> "Master" And sensorMode <> "PerPage" Then _
        Err.Raise 5, , "Mode layer sensor tidak dikenal."
    Set doc = ActiveDocument
    Set startPage = ActivePage
    Set sensorLayers = New Collection
    If sensorMode <> "" Then
        If sensorMode = "Master" Then
            Set destinationPage = doc.MasterPage
        Else
            Set destinationPage = startPage
        End If
        operation = "Membaca layer scpro2 pada page sumber"
        ' Snapshot sebelum perubahan Master memindahkan anggota collection layer.
        ' PerPage harus memakai representasi master layer pada page tujuan,
        ' bukan layer milik doc.MasterPage (page 0).
        For Each ly In startPage.AllLayers
            Select Case ly.Name
                Case "scpro2_printmargin", "scpro2_regmarks", "scpro2_printonly"
                    If sensorMode = "Master" Then
                        If Not ly.Master Then sensorLayers.Add ly
                    Else
                        If ly.Master Then sensorLayers.Add ly
                    End If
            End Select
        Next ly
    End If

    oldUnit = doc.Unit
    unitSaved = True
    operation = "BeginCommandGroup Page Setup"
    doc.BeginCommandGroup "Auto Distribute Page Setup"
    commandGroupOpen = True
    doc.Unit = cdrMillimeter
    operation = "Mengatur ukuran default dokumen"
    doc.MasterPage.SetSize widthMM, heightMM
    For Each pg In doc.Pages
        operation = "Mengatur ukuran Page " & CStr(pg.Index)
        pg.SetSize widthMM, heightMM
    Next pg

    For i = 1 To sensorLayers.Count
        Set ly = sensorLayers(i)
        layerName = ly.Name
        operation = "Mencatat jumlah object pada layer " & layerName
        sourceShapeCount = ly.Shapes.Count
        operation = "Mengaktifkan page tujuan layer " & layerName
        startPage.Activate
        operation = "Mengatur Master = " & CStr(sensorMode = "Master") & ": " & layerName
        ly.Master = (sensorMode = "Master")
        operation = "Memeriksa lokasi layer " & layerName
        ' Ambil ulang dari collection tujuan; referensi sebelum konversi
        ' tidak digunakan untuk menentukan page hasil perpindahan.
        Set movedLayer = ADFindSensorLayer(destinationPage, layerName, sensorMode = "Master")
        If movedLayer Is Nothing Then Err.Raise 5, , _
            "Layer " & layerName & " belum ditemukan pada page tujuan. Gunakan Undo untuk memulihkan operasi."
        operation = "Memeriksa jumlah object setelah perpindahan " & layerName
        If movedLayer.Shapes.Count <> sourceShapeCount Then Err.Raise 5, , _
            "Jumlah object pada layer " & layerName & " berubah setelah perpindahan. Gunakan Undo untuk memulihkan operasi."
    Next i
    If sensorMode <> "" Then ADOrderSensorLayers destinationPage, sensorMode = "Master", operation

    operation = "Memulihkan page dan unit dokumen"
    startPage.Activate
    doc.Unit = oldUnit
    unitSaved = False
    operation = "EndCommandGroup Page Setup"
    doc.EndCommandGroup
    commandGroupOpen = False
    Exit Sub

SetupFailed:
    errorNumber = Err.Number
    errorDescription = Err.Description
    On Error Resume Next
    If Not startPage Is Nothing Then startPage.Activate
    If unitSaved Then doc.Unit = oldUnit
    If commandGroupOpen Then doc.EndCommandGroup
    On Error GoTo 0
    MsgBox "Operasi: " & operation & vbCrLf & "Error " & CStr(errorNumber) & vbCrLf & _
        errorDescription, vbExclamation, "Auto Distribute - Page Setup"
End Sub

Private Function ADFindSensorLayer(ByVal targetPage As Page, ByVal layerName As String, _
    ByVal masterState As Boolean) As Layer
    Dim ly As Layer
    For Each ly In targetPage.Layers
        If ly.Name = layerName And ly.Master = masterState Then
            Set ADFindSensorLayer = ly
            Exit Function
        End If
    Next ly
End Function

Private Sub ADOrderSensorLayers(ByVal targetPage As Page, ByVal masterState As Boolean, _
    ByRef operation As String)
    Dim names As Variant
    Dim name As Variant
    Dim previousLayer As Layer
    Dim ly As Layer
    ' Urutan atas ke bawah sama pada Master Page maupun page lokal.
    names = Array("scpro2_printmargin", "scpro2_regmarks", "scpro2_printonly")
    For Each name In names
        operation = "Membaca layer untuk depth: " & CStr(name)
        Set ly = ADFindSensorLayer(targetPage, CStr(name), masterState)
        If Not ly Is Nothing Then
            If Not previousLayer Is Nothing Then
                operation = "Menempatkan " & ly.Name & " di bawah " & previousLayer.Name
                ly.MoveBelow previousLayer
            End If
            Set previousLayer = ly
        End If
    Next name
End Sub

Public Sub ADClearObjectDraft(ByVal ownerForm As Object)
    Erase mObjectEntries
    mObjectCount = 0
    Set mDraftDocument = Nothing
    ADRefreshObjectContainerList ownerForm
End Sub

Public Sub ADRemoveObjectDraft(ByVal ownerForm As Object)
    Dim list As Object
    Dim i As Long
    Dim j As Long
    Set list = ADGetObjectsListBox(ownerForm)
    If list Is Nothing Then Exit Sub
    For i = mObjectCount - 1 To 0 Step -1
        If list.Selected(i) Then
            For j = i To mObjectCount - 2
                mObjectEntries(j).RoleLabel = mObjectEntries(j + 1).RoleLabel
                mObjectEntries(j).StaticID = mObjectEntries(j + 1).StaticID
            Next j
            mObjectCount = mObjectCount - 1
        End If
    Next i
    If mObjectCount = 0 Then
        Erase mObjectEntries
        Set mDraftDocument = Nothing
    Else
        ReDim Preserve mObjectEntries(0 To mObjectCount - 1)
    End If
    ADRefreshObjectContainerList ownerForm
End Sub

Public Sub ADAddSelectedObjectsToContainer(ByVal roleLabel As String, ByVal ownerForm As Object)
    Dim addedCount As Long

    addedCount = ADAddSelectedObjects(roleLabel)
    ADRefreshObjectContainerList ownerForm

    If addedCount = 0 Then
        MsgBox "Tidak ada object baru yang ditambahkan.", vbInformation, "Object Container"
    End If
End Sub

Public Sub ADRefreshObjectContainerList(ByVal ownerForm As Object)
    Dim lbxObjects As Object
    Dim objectIndex As Long

    Set lbxObjects = ADGetObjectsListBox(ownerForm)
    If lbxObjects Is Nothing Then Exit Sub

    lbxObjects.Clear
    For objectIndex = 0 To mObjectCount - 1
        lbxObjects.AddItem ADGetObjectListText(objectIndex)
    Next objectIndex
End Sub

Public Sub ADProcessStoredObjects(ByVal useKissA As Boolean, ByVal useDieA As Boolean, _
    Optional ByVal rotateCW As Boolean = False, Optional ByVal rotateCCW As Boolean = False, _
    Optional ByVal sequentially As Boolean = False)
    Dim doc As Document
    Dim startPage As Page
    Dim targetPage As Page
    Dim targetLayer As Layer
    Dim resolvedShapes() As Shape
    Dim roles() As String
    Dim executionOrder() As Long
    Dim pageOffsets() As Long
    Dim executionIndex As Long
    Dim roleIndex As Long
    Dim pageIndex As Long
    Dim entryIndex As Long
    Dim targetShape As Shape
    Dim processedCount As Long
    Dim commandGroupOpen As Boolean
    Dim oldUnit As cdrUnit
    Dim unitSaved As Boolean
    Dim operation As String
    Dim warnings As String
    Dim errorNumber As Long
    Dim errorDescription As String

    On Error GoTo ProcessFailed

    operation = "Validasi antrean dan dokumen sumber"
    If mObjectCount = 0 Then Err.Raise 5, , "Tambahkan Design dan Cut Line ke daftar terlebih dahulu."
    If useKissA = useDieA Then Err.Raise 5, , "Pilih salah satu mode Kiss A atau Die A."
    If rotateCW And rotateCCW Then Err.Raise 5, , "Pilih hanya satu arah rotasi."
    If Application.Documents.Count = 0 Then Err.Raise 5, , "Tidak ada dokumen aktif."
    Set doc = ActiveDocument
    If Not doc Is mDraftDocument Then Err.Raise 5, , "Aktifkan dokumen sumber antrean terlebih dahulu."
    Set startPage = ActivePage
    ReDim resolvedShapes(0 To mObjectCount - 1)
    ReDim roles(0 To mObjectCount - 1)
    ' Resolve semua objek sebelum mutasi; pasangan tidak bergeser jika objek hilang.
    For entryIndex = 0 To mObjectCount - 1
        Set targetShape = ADFindShapeByStaticID(mObjectEntries(entryIndex).StaticID)
        If targetShape Is Nothing Then Err.Raise 5, , "Obj. ID: /" & _
            CStr(mObjectEntries(entryIndex).StaticID) & " tidak ditemukan. Perbarui antrean."
        Set resolvedShapes(entryIndex) = targetShape
        roles(entryIndex) = mObjectEntries(entryIndex).RoleLabel
    Next entryIndex
    operation = "Menentukan urutan dan page tujuan antrean"
    ADBuildExecutionPlan roles, useKissA, sequentially, executionOrder, pageOffsets

    oldUnit = doc.Unit
    unitSaved = True
    operation = "BeginCommandGroup"
    doc.BeginCommandGroup "Auto Distribute Object Container"
    commandGroupOpen = True
    doc.Unit = cdrMillimeter
    For executionIndex = 0 To mObjectCount - 1
            entryIndex = executionOrder(executionIndex)
            pageIndex = startPage.Index + pageOffsets(executionIndex)
            operation = "Menyiapkan Page " & CStr(pageIndex)
            If pageIndex > doc.Pages.Count Then
                Set targetPage = doc.InsertPagesEx(1, False, doc.Pages.Count, 335#, 487#)
            Else
                Set targetPage = doc.Pages(pageIndex)
            End If
            targetPage.Activate
            Set targetShape = resolvedShapes(entryIndex)
            If roles(entryIndex) = AD_ROLE_DESIGN Then
                roleIndex = 0
            Else
                roleIndex = 1
            End If
            operation = "Memindahkan Obj. ID: /" & CStr(targetShape.StaticID) & " ke Page " & CStr(pageIndex)
            If useKissA And roleIndex = 1 Then
                Set targetLayer = ADEnsureLayer(targetPage, "Layer 2")
            Else
                Set targetLayer = ADEnsureLayer(targetPage, "Layer 1")
            End If
            targetShape.MoveToLayer targetLayer
            targetLayer.Printable = Not (useKissA And roleIndex = 0)
            operation = "Rotasi dan posisi Obj. ID: /" & CStr(targetShape.StaticID)
            If rotateCW Then targetShape.Rotate -90
            If rotateCCW Then targetShape.Rotate 90
            ADCenterShapeToActivePage targetShape
            If roleIndex = 1 Then ADApplyCutLineOutline targetShape, useKissA, useDieA
            processedCount = processedCount + 1
    Next executionIndex
    ' Cleanup hanya setelah semua objek antrean dipindahkan.
    operation = "Membersihkan layer kosong pada setiap page"
    If useDieA Then ADCleanDieLayers doc, warnings, operation
    operation = "Memulihkan page dan unit"
    startPage.Activate
    doc.Unit = oldUnit
    unitSaved = False
    operation = "EndCommandGroup"
    doc.EndCommandGroup
    commandGroupOpen = False
    MsgBox CStr(processedCount) & " object selesai diproses.", vbInformation, "Auto Distribute"
    If Len(warnings) > 0 Then MsgBox "Layer berikut masih berisi object dan tidak dihapus:" & _
        vbCrLf & warnings, vbExclamation, "Auto Distribute"
    Exit Sub

ProcessFailed:
    errorNumber = Err.Number
    errorDescription = Err.Description

    On Error Resume Next
    If Not startPage Is Nothing Then startPage.Activate
    If unitSaved Then doc.Unit = oldUnit
    If commandGroupOpen Then doc.EndCommandGroup
    On Error GoTo 0

    Err.Raise errorNumber, "ADProcessStoredObjects", "Operasi: " & operation & vbCrLf & errorDescription
End Sub

Private Sub ADBuildExecutionPlan(ByRef roles() As String, ByVal useKissA As Boolean, _
    ByVal sequentially As Boolean, ByRef executionOrder() As Long, ByRef pageOffsets() As Long)
    Dim designs() As Long
    Dim cuts() As Long
    Dim designCount As Long
    Dim cutCount As Long
    Dim segmentFirstPage As Long
    Dim i As Long
    Dim n As Long

    n = UBound(roles) + 1
    ReDim executionOrder(0 To n - 1)
    ReDim pageOffsets(0 To n - 1)
    ReDim designs(0 To n - 1)
    ReDim cuts(0 To n - 1)
    segmentFirstPage = -1
    For i = 0 To n - 1
        executionOrder(i) = i
        If roles(i) = AD_ROLE_DESIGN Then
            designs(designCount) = i
            pageOffsets(i) = designCount
            If segmentFirstPage = -1 Then segmentFirstPage = designCount
            designCount = designCount + 1
        Else
            cuts(cutCount) = i
            cutCount = cutCount + 1
            If sequentially And useKissA Then
                If segmentFirstPage = -1 Then Err.Raise 5, , _
                    "Cut Line pada baris " & CStr(i + 1) & _
                    " tidak memiliki Design sebelumnya sejak Cut Line terakhir. Perbarui antrean."
                pageOffsets(i) = segmentFirstPage
                segmentFirstPage = -1
            End If
        End If
        If sequentially And Not useKissA Then pageOffsets(i) = i
    Next i
    If sequentially Then Exit Sub

    If designCount <> cutCount Then Err.Raise 5, , _
        "Jumlah Design dan Cut Line harus sama untuk membentuk pasangan. Perbarui antrean."
    For i = 0 To designCount - 1
        executionOrder(i * 2) = designs(i)
        executionOrder(i * 2 + 1) = cuts(i)
        If useKissA Then
            pageOffsets(i * 2) = i
            pageOffsets(i * 2 + 1) = i
        Else
            pageOffsets(i * 2) = i * 2
            pageOffsets(i * 2 + 1) = i * 2 + 1
        End If
    Next i
End Sub

Private Function ADEnsureLayer(ByVal targetPage As Page, ByVal layerName As String) As Layer
    Dim ly As Layer
    For Each ly In targetPage.Layers
        If ly.Name = layerName Then
            Set ADEnsureLayer = ly
            Exit Function
        End If
    Next ly
    Set ADEnsureLayer = targetPage.CreateLayer(layerName)
End Function

Private Sub ADCleanDieLayers(ByVal doc As Document, ByRef warnings As String, ByRef operation As String)
    Dim pg As Page
    Dim ly As Layer
    Dim i As Long
    Dim suffix As String
    Dim layerContext As String
    For Each pg In doc.Pages
        operation = "Membaca jumlah layer Page " & CStr(pg.Index)
        For i = pg.Layers.Count To 1 Step -1
            operation = "Mengakses layer index " & CStr(i) & " pada Page " & CStr(pg.Index)
            Set ly = pg.Layers(i)
            operation = "Membaca nama layer index " & CStr(i) & " pada Page " & CStr(pg.Index)
            layerContext = "Page " & CStr(pg.Index) & " - " & ly.Name
            ' Hanya Layer 2, Layer 3, dst.; bukan layer sistem/master.
            operation = "Memeriksa Master dan nama layer: " & layerContext
            If Not ly.Master And Left$(ly.Name, 6) = "Layer " Then
                suffix = Mid$(ly.Name, 7)
                If Len(suffix) > 0 And Not suffix Like "*[!0-9]*" Then
                    If Val(suffix) >= 2 Then
                        operation = "Membaca Shapes.Count: " & layerContext
                        If ly.Shapes.Count > 0 Then
                            warnings = warnings & "Page " & CStr(pg.Index) & " - " & ly.Name & vbCrLf
                        Else
                            operation = "Mengatur Editable = True: " & layerContext
                            ly.Editable = True
                            operation = "Menghapus layer kosong: " & layerContext
                            ly.Delete
                        End If
                    End If
                End If
            End If
        Next i
    Next pg
End Sub

Private Function ADAddSelectedObjects(ByVal roleLabel As String) As Long
    Dim selectedShapes As ShapeRange
    Dim selectedShape As Shape
    Dim selectionIndex As Long
    Dim staticID As Long

    On Error Resume Next
    Set selectedShapes = ActiveSelectionRange
    On Error GoTo 0

    If selectedShapes Is Nothing Then
        MsgBox "Pilih object terlebih dahulu.", vbExclamation, "Object Container"
        Exit Function
    End If

    If selectedShapes.Count = 0 Then
        MsgBox "Pilih object terlebih dahulu.", vbExclamation, "Object Container"
        Exit Function
    End If

    If mObjectCount > 0 Then
        If Not ActiveDocument Is mDraftDocument Then
            MsgBox "Antrean hanya dapat berisi object dari satu dokumen.", vbExclamation, "Object Container"
            Exit Function
        End If
    Else
        Set mDraftDocument = ActiveDocument
    End If

    ' Balik urutan range seleksi untuk urutan penambahan Design maupun Cut Line.
    For selectionIndex = selectedShapes.Count To 1 Step -1
        Set selectedShape = selectedShapes(selectionIndex)
        staticID = CLng(selectedShape.StaticID)

        If Not ADObjectEntryExists(roleLabel, staticID) Then
            ADAppendObjectEntry roleLabel, staticID
            ADAddSelectedObjects = ADAddSelectedObjects + 1
        End If
    Next selectionIndex
End Function

Private Sub ADAppendObjectEntry(ByVal roleLabel As String, ByVal staticID As Long)
    If mObjectCount = 0 Then
        ReDim mObjectEntries(0 To 0)
    Else
        ReDim Preserve mObjectEntries(0 To mObjectCount)
    End If

    mObjectEntries(mObjectCount).RoleLabel = roleLabel
    mObjectEntries(mObjectCount).StaticID = staticID
    mObjectCount = mObjectCount + 1
End Sub

Private Function ADObjectEntryExists(ByVal roleLabel As String, ByVal staticID As Long) As Boolean
    Dim entryIndex As Long

    For entryIndex = 0 To mObjectCount - 1
        If mObjectEntries(entryIndex).StaticID = staticID Then
            ADObjectEntryExists = True
            Exit Function
        End If
    Next entryIndex
End Function

Private Function ADGetObjectListText(ByVal objectIndex As Long) As String
    ADGetObjectListText = mObjectEntries(objectIndex).RoleLabel & _
        " | Obj. ID: /" & CStr(mObjectEntries(objectIndex).StaticID)
End Function

Private Function ADGetObjectsListBox(ByVal ownerForm As Object) As Object
    On Error Resume Next
    Set ADGetObjectsListBox = ownerForm.Controls("lbxObjects")
    On Error GoTo 0
End Function

Private Function ADFindShapeByStaticID(ByVal staticID As Long) As Shape
    Dim pageItem As Page
    Dim layerItem As Layer
    Dim foundShape As Shape

    For Each pageItem In ActiveDocument.Pages
        For Each layerItem In pageItem.Layers
            Set foundShape = ADFindShapeInRange(layerItem.Shapes, staticID)
            If Not foundShape Is Nothing Then
                Set ADFindShapeByStaticID = foundShape
                Exit Function
            End If
        Next layerItem
    Next pageItem
End Function

Private Function ADFindShapeInRange(ByVal shapesToSearch As Shapes, ByVal staticID As Long) As Shape
    Dim shapeItem As Shape
    Dim nestedShape As Shape

    For Each shapeItem In shapesToSearch
        If CLng(shapeItem.StaticID) = staticID Then
            Set ADFindShapeInRange = shapeItem
            Exit Function
        End If

        Set nestedShape = ADFindShapeInGroup(shapeItem, staticID)
        If Not nestedShape Is Nothing Then
            Set ADFindShapeInRange = nestedShape
            Exit Function
        End If
    Next shapeItem
End Function

Private Function ADFindShapeInGroup(ByVal groupShape As Shape, ByVal staticID As Long) As Shape
    Dim childShapes As Shapes
    Dim nestedShape As Shape
    Dim foundShape As Shape

    On Error Resume Next
    Set childShapes = groupShape.Shapes
    On Error GoTo 0

    If childShapes Is Nothing Then Exit Function

    For Each nestedShape In childShapes
        If CLng(nestedShape.StaticID) = staticID Then
            Set ADFindShapeInGroup = nestedShape
            Exit Function
        End If

        Set foundShape = ADFindShapeInGroup(nestedShape, staticID)
        If Not foundShape Is Nothing Then
            Set ADFindShapeInGroup = foundShape
            Exit Function
        End If
    Next nestedShape
End Function

Private Sub ADCenterShapeToActivePage(ByVal targetShape As Shape)
    targetShape.CenterX = ActivePage.SizeWidth / 2
    targetShape.CenterY = ActivePage.SizeHeight / 2
End Sub

Private Sub ADApplyCutLineOutline(ByVal targetShape As Shape, ByVal useKissA As Boolean, ByVal useDieA As Boolean)
    If useKissA Then
        targetShape.Outline.SetProperties Color:=CreateCMYKColor(0, 100, 100, 0)
    ElseIf useDieA Then
        targetShape.Outline.SetProperties Color:=CreateRGBColor(255, 0, 255)
    End If
End Sub
