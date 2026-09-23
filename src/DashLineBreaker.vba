Option Explicit
' Merge declarations into the top of the target UserForm code module.
Private pMRObserver As Object
Private pMRToken As String

Private Const DLB_DASH_COUNT As Long = 1
Private Const DLB_DASH_LENGTH_1 As Double = 55#
Private Const DLB_GAP_LENGTH_1 As Double = 5#
Private Const DLB_DASH_DOT_LENGTH As Double = 0#
Private Const DLB_DEFAULT_OUTLINE_WIDTH As Double = 0.2#

Private Enum DLBMode
    dlbGeneratePattern = 0
    dlbExplodeExisting = 1
End Enum

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
    RunDLB dlbGeneratePattern
End Sub

Private Sub cmdExplode_Click()
    RunDLB dlbExplodeExisting
End Sub

Private Sub RunDLB(ByVal mode As DLBMode)

    Dim doc As Document
    Dim targets As New Collection
    Dim sr As ShapeRange
    Dim s As Shape
    Dim i As Long
    Dim dashCreated As Long
    Dim processed As Long
    Dim commandStarted As Boolean
    Dim targetClockwise As Boolean
    Dim previousOptimization As Boolean
    Dim failed As Boolean
    Dim errorNumber As Long
    Dim errorDescription As String

    On Error GoTo ErrHandler

    If Documents.Count = 0 Then
        MsgBox "Tidak ada document aktif.", vbExclamation
        Exit Sub
    End If

    Set doc = ActiveDocument
    Set sr = ActiveSelectionRange

    If sr Is Nothing Or sr.Count = 0 Then
        MsgBox "Pilih curve terlebih dahulu.", vbExclamation
        Exit Sub
    End If

    If Not optClockwise.Value And Not optCounterClockwise.Value Then
        MsgBox "Pilih arah path terlebih dahulu.", vbExclamation
        Exit Sub
    End If

    targetClockwise = optClockwise.Value

    ' Selection dapat berubah selama proses.
    For Each s In sr.Shapes
        targets.Add s
    Next s

    previousOptimization = Application.Optimization
    doc.BeginCommandGroup "Dash Line Breaker"
    commandStarted = True
    Application.Optimization = True

    For i = 1 To targets.Count
        Set s = targets.Item(i)
        dashCreated = ExplodeDashedShape( _
            s, _
            targetClockwise, _
            mode)

        If dashCreated > 0 Then
            processed = processed + 1
        End If
    Next i

CleanExit:
    On Error Resume Next

    Application.Optimization = previousOptimization
    If commandStarted Then
        doc.EndCommandGroup
    End If
    ActiveWindow.Refresh

    On Error GoTo 0

    If failed Then
        MsgBox "Dash Line Breaker gagal." & vbCrLf & vbCrLf & _
               "Error " & errorNumber & ": " & errorDescription, _
               vbCritical
    ElseIf processed = 0 Then

        If mode = dlbExplodeExisting Then
            MsgBox "Tidak ada dashed curve yang dapat di-explode.", _
                   vbInformation
        Else
            MsgBox "Tidak ada curve yang dapat diproses.", _
                   vbInformation
        End If
    End If

    Exit Sub

ErrHandler:
    failed = True
    errorNumber = Err.Number
    errorDescription = Err.Description

    Resume CleanExit
End Sub

Private Function ExplodeDashedShape( _
    ByVal s As Shape, _
    ByVal targetClockwise As Boolean, _
    ByVal mode As DLBMode) As Long

    Dim sp As SubPath
    Dim os As OutlineStyle
    Dim startX As Double
    Dim startY As Double
    Dim directionReversed As Boolean
    Dim outlineWidth As Double
    Dim outlineMiterLimit As Double
    Dim outlineLineCaps As Long
    Dim outlineLineJoin As Long
    Dim outlineJustification As Long
    Dim outlineBehindFill As Boolean
    Dim outlineScaleWithShape As Boolean
    Dim outlineColor As New Color
    Dim dashLen() As Double
    Dim gapLen() As Double
    Dim boundaries() As Double
    Dim dashCount As Long
    Dim boundaryCount As Long
    Dim boundaryCapacity As Long
    Dim pairIndex As Long
    Dim i As Long
    Dim pathLength As Double
    Dim patternUnits As Double
    Dim patternScale As Double
    Dim pos As Double
    Dim eps As Double
    Dim dashDotLength As Double

    On Error GoTo SkipShape

    If s.Type <> cdrCurveShape Then
        s.ConvertToCurves
    End If

    If s.Type <> cdrCurveShape Then GoTo SkipShape
    If s.Curve.SubPaths.Count <> 1 Then GoTo SkipShape

    Select Case mode

        Case dlbGeneratePattern
            EnsureOutline s

            CaptureOutlineProperties _
                s, _
                outlineWidth, _
                outlineMiterLimit, _
                outlineLineCaps, _
                outlineLineJoin, _
                outlineJustification, _
                outlineBehindFill, _
                outlineScaleWithShape, _
                outlineColor
            If Not ApplyDashPattern(s) Then GoTo SkipShape

        Case dlbExplodeExisting
            If s.Outline.Type = cdrNoOutline Then GoTo SkipShape
            CaptureOutlineProperties _
                s, _
                outlineWidth, _
                outlineMiterLimit, _
                outlineLineCaps, _
                outlineLineJoin, _
                outlineJustification, _
                outlineBehindFill, _
                outlineScaleWithShape, _
                outlineColor
        Case Else
            GoTo SkipShape

    End Select

    Set os = s.Outline.Style

    dashCount = os.DashCount

    If dashCount <= 0 Then GoTo SkipShape

    ReDim dashLen(1 To dashCount)
    ReDim gapLen(1 To dashCount)

    For i = 1 To dashCount
        patternUnits = patternUnits + os.DashLength(i)
        patternUnits = patternUnits + os.GapLength(i)
    Next i

    If patternUnits <= 0 Then GoTo SkipShape

    dashDotLength = s.Outline.DashDotLength
    If dashDotLength > 0 Then
        patternScale = dashDotLength / patternUnits
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

    startX = sp.StartNode.PositionX
    startY = sp.StartNode.PositionY
    If sp.Closed Then

        If sp.IsClockwise <> targetClockwise Then
            sp.ReverseDirection
            directionReversed = True
            Set sp = s.Curve.SubPaths(1)
        End If

        If directionReversed Then
            RestoreStartNode _
                s.Curve, _
                1, _
                startX, _
                startY

            Set sp = s.Curve.SubPaths(1)
        End If

    End If

    pathLength = sp.Length
    If pathLength <= 0 Then GoTo SkipShape

    eps = pathLength * 0.0000001
    If eps < 0.0000001 Then
        eps = 0.0000001
    End If

    pos = 0
    pairIndex = 1
    Do While pos < pathLength - eps

        pos = pos + dashLen(pairIndex)
        If pos < pathLength - eps Then
            AppendBoundary _
                boundaries, _
                boundaryCount, _
                boundaryCapacity, _
                pos
        Else
            Exit Do
        End If

        pos = pos + gapLen(pairIndex)
        If pos < pathLength - eps Then
            AppendBoundary _
                boundaries, _
                boundaryCount, _
                boundaryCapacity, _
                pos
        Else
            Exit Do
        End If

        pairIndex = pairIndex + 1
        If pairIndex > dashCount Then
            pairIndex = 1
        End If
    Loop

    If boundaryCount = 0 Then
        ApplySolidOutline _
            s, _
            outlineWidth, _
            outlineMiterLimit, _
            outlineLineCaps, _
            outlineLineJoin, _
            outlineJustification, _
            outlineBehindFill, _
            outlineScaleWithShape, _
            outlineColor
        ExplodeDashedShape = 1
        Exit Function
    End If

    If sp.Closed Then
        sp.StartNode.BreakApart
        Set sp = s.Curve.SubPaths(1)
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

    ApplySolidOutline _
        s, _
        outlineWidth, _
        outlineMiterLimit, _
        outlineLineCaps, _
        outlineLineJoin, _
        outlineJustification, _
        outlineBehindFill, _
        outlineScaleWithShape, _
        outlineColor

    ExplodeDashedShape = s.Curve.SubPaths.Count

    Exit Function

SkipShape:
    ExplodeDashedShape = 0

End Function

Private Sub EnsureOutline(ByVal s As Shape)
    Dim outlineRange As New ShapeRange

    If s.Outline.Type <> cdrNoOutline Then Exit Sub

    outlineRange.Add s
    outlineRange.SetOutlineProperties Width:=DLB_DEFAULT_OUTLINE_WIDTH
End Sub

Private Function ApplyDashPattern(ByVal s As Shape) As Boolean

    Dim dashLengthText As String
    Dim gapLengthText As String
    Dim dotLengthText As String

    On Error GoTo Failed

    dashLengthText = ToInvariantNumber(DLB_DASH_LENGTH_1)
    gapLengthText = ToInvariantNumber(DLB_GAP_LENGTH_1)
    dotLengthText = ToInvariantNumber(DLB_DASH_DOT_LENGTH)

    s.Style.StringAssign _
        "{""outline"":{" & _
        """dashDotSpec"":""2," & _
        dashLengthText & "," & _
        gapLengthText & """," & _
        """dotLength"":""" & _
        dotLengthText & """}}"

    ApplyDashPattern = _
        (s.Outline.Style.DashCount = DLB_DASH_COUNT)

    Exit Function

Failed:
    ApplyDashPattern = False

End Function

Private Function ToInvariantNumber(ByVal value As Double) As String

    Dim textValue As String

    ' Str$ secara praktis memberi representasi numerik VBA
    ' yang kemudian kita normalisasi lagi untuk berjaga-jaga.
    textValue = Trim$(Str$(value))
    textValue = Replace$(textValue, ",", ".")
    ToInvariantNumber = textValue

End Function

Private Sub CaptureOutlineProperties( _
    ByVal s As Shape, _
    ByRef outlineWidth As Double, _
    ByRef outlineMiterLimit As Double, _
    ByRef outlineLineCaps As Long, _
    ByRef outlineLineJoin As Long, _
    ByRef outlineJustification As Long, _
    ByRef outlineBehindFill As Boolean, _
    ByRef outlineScaleWithShape As Boolean, _
    ByRef outlineColor As Color)

    outlineWidth = s.Outline.Width
    outlineMiterLimit = s.Outline.MiterLimit
    outlineLineCaps = s.Outline.LineCaps
    outlineLineJoin = s.Outline.LineJoin

    outlineJustification = s.Outline.Justification
    outlineBehindFill = s.Outline.BehindFill
    outlineScaleWithShape = s.Outline.ScaleWithShape

    outlineColor.CopyAssign s.Outline.Color

End Sub

Private Sub ApplySolidOutline( _
    ByVal s As Shape, _
    ByVal outlineWidth As Double, _
    ByVal outlineMiterLimit As Double, _
    ByVal outlineLineCaps As Long, _
    ByVal outlineLineJoin As Long, _
    ByVal outlineJustification As Long, _
    ByVal outlineBehindFill As Boolean, _
    ByVal outlineScaleWithShape As Boolean, _
    ByVal outlineColor As Color)

    Dim outlineRange As New ShapeRange

    ' Hanya dash state yang sengaja dinormalisasi menjadi solid.
    s.Outline.Style = OutlineStyles.Item(0)
    outlineRange.Add s

    outlineRange.SetOutlineProperties _
        Width:=outlineWidth, _
        Color:=outlineColor, _
        MiterLimit:=outlineMiterLimit, _
        LineCaps:=outlineLineCaps, _
        LineJoin:=outlineLineJoin

    With s.Outline

        .Justification = outlineJustification
        .BehindFill = outlineBehindFill
        .ScaleWithShape = outlineScaleWithShape

    End With

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

Private Sub AppendBoundary( _
    ByRef boundaries() As Double, _
    ByRef boundaryCount As Long, _
    ByRef boundaryCapacity As Long, _
    ByVal position As Double)

    If boundaryCount >= boundaryCapacity Then
        If boundaryCapacity = 0 Then
            boundaryCapacity = 32
            ReDim boundaries(1 To boundaryCapacity)
        Else
            boundaryCapacity = boundaryCapacity * 2
            ReDim Preserve boundaries(1 To boundaryCapacity)
        End If
    End If

    boundaryCount = boundaryCount + 1
    boundaries(boundaryCount) = position

End Sub
