Attribute VB_Name = "modCamDoc"
Option Explicit

' Адаптер CorelDRAW: выделение -> контуры и отверстия, предпросмотр траекторий.

Private mOldUnit As Long
Private mUnitSaved As Boolean

Public Sub DocBeginUnits()
    On Error Resume Next
    mOldUnit = ActiveDocument.Unit
    mUnitSaved = True
    ActiveDocument.Unit = cdrMillimeter
    On Error GoTo 0
End Sub

Public Sub DocEndUnits()
    On Error Resume Next
    If mUnitSaved Then ActiveDocument.Unit = mOldUnit
    mUnitSaved = False
    On Error GoTo 0
End Sub

Public Function DocCollectContours(ByRef rings As PolySet, ByRef holes As HoleSet) As Boolean
    Dim sh As Shape
    PolySetInit rings
    HoleSetInit holes
    If ActiveDocument Is Nothing Then Exit Function
    For Each sh In ActiveDocument.Selection.Shapes
        DocShapeToRings sh, rings, holes
    Next sh
    DocCollectContours = (rings.Count > 0 Or holes.Count > 0)
End Function

Private Sub DocShapeToRings(ByRef sh As Shape, ByRef rings As PolySet, ByRef holes As HoleSet)
    Dim sub1 As Shape, sp As SubPath, p As Poly
    Dim w As Double, h As Double

    If sh.Type = cdrEllipseShape Then
        w = sh.SizeWidth
        h = sh.SizeHeight
        If w > 0.01 And Abs(w - h) < 0.01 Then
            HoleSetAdd holes, sh.CenterX, sh.CenterY, w
            Exit Sub
        End If
    End If

    If sh.Type <> cdrCurveShape And sh.Type <> cdrGroupShape Then
        On Error Resume Next
        sh.ConvertToCurves
        On Error GoTo 0
    End If

    If sh.Type = cdrGroupShape Then
        For Each sub1 In sh.Shapes
            DocShapeToRings sub1, rings, holes
        Next sub1
        Exit Sub
    End If
    If sh.Type <> cdrCurveShape Then Exit Sub

    For Each sp In sh.Curve.SubPaths
        PolyInit p
        DocSubPathToPoly sp, p, gCfg.FlattenTol
        If p.N >= 3 Then
            p.Closed = sp.Closed
            PolySetAdd rings, p
        End If
    Next sp
End Sub

' Точки сегментов берутся в единицах документа (мм), управляющие точки — абсолютные.
Public Sub DocSubPathToPoly(ByRef sp As SubPath, ByRef p As Poly, ByVal tol As Double)
    Dim seg As Segment
    Dim started As Boolean

    For Each seg In sp.Segments
        If Not started Then
            PolyAdd p, seg.StartNode.PositionX, seg.StartNode.PositionY
            started = True
        End If
        If seg.Type = cdrCurveSegment Then
            GeoFlattenCubic p, _
                seg.StartNode.PositionX, seg.StartNode.PositionY, _
                seg.StartingControlPointX, seg.StartingControlPointY, _
                seg.EndingControlPointX, seg.EndingControlPointY, _
                seg.EndNode.PositionX, seg.EndNode.PositionY, tol
        Else
            PolyAdd p, seg.EndNode.PositionX, seg.EndNode.PositionY
        End If
    Next seg
End Sub

' Рисует траектории в слое CAM_Toolpaths. Не критично для G-кода,
' поэтому ошибки отрисовки не прерывают работу макроса.
Public Sub DocDrawPreview()
    Dim lyr As Layer, crv As Curve, sp As SubPath, s As Shape
    Dim i As Long, have As Boolean

    On Error GoTo done
    Set lyr = ActiveDocument.ActivePage.CreateLayer("CAM_Toolpaths")
    Set crv = Application.CreateCurve(ActiveDocument)
    have = False
    For i = 0 To gMoveCount - 1
        Select Case gMoves(i).Kind
            Case mmCut, mmPlunge
                If have Then
                    sp.AppendLineSegment gMoves(i).X, gMoves(i).Y
                Else
                    Set sp = crv.CreateSubPath(gMoves(i).X, gMoves(i).Y)
                    have = True
                End If
            Case Else
                have = False
        End Select
    Next i
    If crv.SubPaths.Count > 0 Then
        Set s = lyr.CreateCurve(crv)
        s.Outline.Width = 0.1
    End If
done:
    On Error GoTo 0
End Sub
