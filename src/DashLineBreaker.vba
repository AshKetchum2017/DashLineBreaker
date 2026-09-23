Option Explicit
' Merge declarations into the top of the target UserForm code module.
Private pMRObserver As Object
Private pMRToken As String

Private Const DLB_DASH_COUNT As Long = 1
Private Const DLB_DASH_LENGTH_1 As Double = 55#
Private Const DLB_GAP_LENGTH_1 As Double = 5#
Private Const DLB_DASH_DOT_LENGTH As Double = 0#

Private Sub optClockwise_Click()
    If optClockwise.Value Then optCounterClockwise.Value = False
End Sub

Private Sub optCounterClockwise_Click()
    If optCounterClockwise.Value Then optClockwise.Value = False
End Sub

Private Sub UserForm_Initialize()
    If Not optClockwise.Value And Not optCounterClockwise.Value Then
        optClockwise.Value = True
    End If
End Sub

Private Sub cmdClose_Click()
    Unload Me
End Sub

Private Sub cmdProcess_Click()
    Dim targets As New Collection
    Dim sr As ShapeRange
    Dim s As Shape
    Dim i As Long
    Dim dashCreated As Long
    Dim processed As Long
    Dim commandStarted As Boolean
    Dim targetClockwise As Boolean

    If Documents.Count = 0 Then
        MsgBox "Tidak ada document aktif.", vbExclamation
        Exit Sub
    End If

    Set sr = ActiveSelectionRange

    If sr.Count = 0 Then
        MsgBox "Pilih dashed curve terlebih dahulu.", vbExclamation
        Exit Sub
    End If

    If Not optClockwise.Value And Not optCounterClockwise.Value Then
        MsgBox "Pilih arah path terlebih dahulu.", vbExclamation
        Exit Sub
    End If

    targetClockwise = optClockwise.Value

    ' Simpan reference object terlebih dahulu karena selection
    ' akan berubah saat BreakApart dijalankan.
    For Each s In sr.Shapes
        targets.Add s
    Next s

    On Error GoTo ErrHandler
    ActiveDocument.BeginCommandGroup "Explode Dashed Curve"
    commandStarted = True
    Application.Optimization = True

    For i = 1 To targets.Count
        Set s = targets.Item(i)
        dashCreated = ExplodeDashedShape(s, targetClockwise)
        If dashCreated > 0 Then
            processed = processed + 1
        End If
    Next i

SafeExit:
    Application.Optimization = False
    ActiveWindow.Refresh

    If commandStarted Then
        ActiveDocument.EndCommandGroup
    End If

    If processed = 0 Then
        MsgBox "Tidak ada dashed curve yang dapat diproses.", vbInformation
    End If

    Exit Sub

ErrHandler:
    Application.Optimization = False
    If commandStarted Then
        ActiveDocument.EndCommandGroup
    End If
    ActiveWindow.Refresh
    MsgBox "Explode Dash gagal." & vbCrLf & vbCrLf & _
           "Error " & Err.Number & ": " & Err.Description, _
           vbCritical
End Sub

Private Function ExplodeDashedShape(ByVal s As Shape, ByVal targetClockwise As Boolean) As Long
    Dim sp As SubPath
    Dim os As OutlineStyle
    Dim startX As Double
    Dim startY As Double
    Dim directionReversed As Boolean
    Dim outlineWidth As Double
    Dim outlineMiterLimit As Double
    Dim outlineLineCaps As Long
    Dim outlineLineJoin As Long
    Dim dashLen() As Double
    Dim gapLen() As Double
    Dim boundaries() As Double
    Dim dashCount As Long
    Dim boundaryCount As Long
    Dim pairIndex As Long
    Dim i As Long
    Dim pathLength As Double
    Dim patternUnits As Double
    Dim patternScale As Double
    Dim pos As Double
    Dim eps As Double

    On Error GoTo SkipShape

    If s.Type <> cdrCurveShape Then
        s.ConvertToCurves
    End If

    If s.Type <> cdrCurveShape Then GoTo SkipShape

    ' Versi emergency ini memproses satu continuous path
    ' per Shape.
    If s.Curve.SubPaths.Count <> 1 Then GoTo SkipShape
    If s.Outline.Type = cdrNoOutline Then GoTo SkipShape

    ApplyDashPattern s

    Set os = s.Outline.Style
    
    dashCount = os.dashCount
    ' DashCount = 0 berarti solid line
    If dashCount <= 0 Then GoTo SkipShape
    ReDim dashLen(1 To dashCount)
    ReDim gapLen(1 To dashCount)
    patternUnits = 0

    For i = 1 To dashCount
        patternUnits = patternUnits + os.DashLength(i)
        patternUnits = patternUnits + os.GapLength(i)
    Next i

    If patternUnits <= 0 Then GoTo SkipShape

    If s.Outline.DashDotLength > 0 Then
        patternScale = _
            s.Outline.DashDotLength / patternUnits
    Else
        patternScale = s.Outline.Width
    End If

    If patternScale <= 0 Then GoTo SkipShape

    For i = 1 To dashCount
        dashLen(i) = _
            os.DashLength(i) * patternScale
        gapLen(i) = _
            os.GapLength(i) * patternScale
    Next i

    Set sp = s.Curve.SubPaths(1)
    ' Capture the original seam before changing path direction.
    startX = sp.StartNode.PositionX
    startY = sp.StartNode.PositionY

    If sp.IsClockwise <> targetClockwise Then
        sp.ReverseDirection
        directionReversed = True
        Set sp = s.Curve.SubPaths(1)
    End If

    If directionReversed And sp.Closed Then
        RestoreStartNode s.Curve, 1, startX, startY
        Set sp = s.Curve.SubPaths(1)
    End If

    pathLength = sp.Length
    If pathLength <= 0 Then GoTo SkipShape

    eps = pathLength * 0.0000001
    If eps < 0.0000001 Then
        eps = 0.0000001
    End If

    If sp.Closed Then
        sp.StartNode.BreakApart
        Set sp = s.Curve.SubPaths(1)
    End If

    pos = 0
    pairIndex = 1
    Do While pos < pathLength - eps
        pos = pos + dashLen(pairIndex)
        If pos < pathLength - eps Then
            boundaryCount = boundaryCount + 1
            ReDim Preserve boundaries(1 To boundaryCount)
            boundaries(boundaryCount) = pos
        Else
            Exit Do
        End If

        pos = pos + gapLen(pairIndex)
        If pos < pathLength - eps Then
            boundaryCount = boundaryCount + 1
            ReDim Preserve boundaries(1 To boundaryCount)
            boundaries(boundaryCount) = pos
        Else
            Exit Do
        End If

        pairIndex = pairIndex + 1
        If pairIndex > dashCount Then
         pairIndex = 1
        End If
    Loop

    If boundaryCount = 0 Then
        CaptureOutlineProperties s, outlineWidth, outlineMiterLimit, outlineLineCaps, outlineLineJoin
        ApplySolidOutline s, outlineWidth, outlineMiterLimit, outlineLineCaps, outlineLineJoin
        ExplodeDashedShape = 1
        Exit Function
    End If

    For i = boundaryCount To 1 Step -1
        Set sp = s.Curve.SubPaths(1)
        sp.BreakApartAt _
            boundaries(i), _
            cdrAbsoluteSegmentOffset
    Next i

    For i = s.Curve.SubPaths.Count To 2 Step -1
        If (i Mod 2) = 0 Then
            s.Curve.SubPaths(i).Delete
        End If
    Next i

    CaptureOutlineProperties s, outlineWidth, outlineMiterLimit, outlineLineCaps, outlineLineJoin
    ApplySolidOutline s, outlineWidth, outlineMiterLimit, outlineLineCaps, outlineLineJoin
    ExplodeDashedShape = s.Curve.SubPaths.Count

    Exit Function

SkipShape:
    ExplodeDashedShape = 0
End Function

Private Sub ApplyDashPattern(ByVal s As Shape)
    With s.Outline.Style
        .DashCount = DLB_DASH_COUNT
        .DashLength(1) = DLB_DASH_LENGTH_1
        .GapLength(1) = DLB_GAP_LENGTH_1
    End With
    s.Outline.DashDotLength = DLB_DASH_DOT_LENGTH
End Sub

Private Sub CaptureOutlineProperties(ByVal s As Shape, ByRef outlineWidth As Double, _
                                     ByRef outlineMiterLimit As Double, _
                                     ByRef outlineLineCaps As Long, _
                                     ByRef outlineLineJoin As Long)

    outlineWidth = s.Outline.Width
    outlineMiterLimit = s.Outline.MiterLimit
    outlineLineCaps = s.Outline.LineCaps
    outlineLineJoin = s.Outline.LineJoin
End Sub

Private Sub ApplySolidOutline(ByVal s As Shape, ByVal outlineWidth As Double, _
                              ByVal outlineMiterLimit As Double, _
                              ByVal outlineLineCaps As Long, _
                              ByVal outlineLineJoin As Long)
    Dim outlineRange As New ShapeRange

    s.Outline.Style = OutlineStyles.Item(0)
    outlineRange.Add s
    outlineRange.SetOutlineProperties _
        Width:=outlineWidth, _
        MiterLimit:=outlineMiterLimit, _
        LineCaps:=outlineLineCaps, _
        LineJoin:=outlineLineJoin
End Sub

Private Sub RestoreStartNode(ByVal crv As Curve, ByVal subpathIndex As Long, _
                             ByVal startX As Double, ByVal startY As Double)
    Dim n As Node
    Dim targetNode As Node
    Dim distance As Double
    Dim bestDistance As Double
    Dim sp As SubPath

    Set sp = crv.SubPaths(subpathIndex)
    bestDistance = -1#

    For Each n In sp.Nodes
        distance = (n.PositionX - startX) ^ 2 + (n.PositionY - startY) ^ 2
        If bestDistance < 0# Or distance < bestDistance Then
            bestDistance = distance
            Set targetNode = n
        End If
    Next n

    If targetNode Is Nothing Then Exit Sub
    targetNode.BreakApart
    Set sp = crv.SubPaths(subpathIndex)

    If Not sp.Closed Then
        On Error Resume Next
        sp.StartNode.JoinWith sp.EndNode
        If Not sp.Closed Then sp.Closed = True
        On Error GoTo 0
    End If
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
