Attribute VB_Name = "modCamGeo"
Option Explicit

' Геометрия: контейнеры, расстояния, разбор кривых, смещение контура.
' Алгоритмы сверены с tools/cam_reference.py (см. modCamSelfTest).

Private Const PI_ As Double = 3.14159265358979
Private Const MITRE_LIMIT As Double = 4#
Private Const ARC_STEP_DEG As Double = 12#
Private Const JOIN_MIN_FACTOR As Double = 0.6
Private Const EPS_ As Double = 0.000000001

' ---------------------------------------------------------------- контейнеры

Public Sub PolyInit(ByRef p As Poly)
    p.N = 0
    p.Cap = 0
    p.Closed = False
    Erase p.X
    Erase p.Y
End Sub

Public Sub PolyAdd(ByRef p As Poly, ByVal x As Double, ByVal y As Double)
    If p.N > 0 Then
        If Abs(p.X(p.N - 1) - x) < EPS_ And Abs(p.Y(p.N - 1) - y) < EPS_ Then Exit Sub
    End If
    If p.N >= p.Cap Then
        If p.Cap = 0 Then
            p.Cap = 32
        Else
            p.Cap = p.Cap * 2
        End If
        ReDim Preserve p.X(p.Cap - 1)
        ReDim Preserve p.Y(p.Cap - 1)
    End If
    p.X(p.N) = x
    p.Y(p.N) = y
    p.N = p.N + 1
End Sub

Public Sub PolyCopy(ByRef dst As Poly, ByRef src As Poly)
    Dim i As Long
    PolyInit dst
    dst.Closed = src.Closed
    For i = 0 To src.N - 1
        PolyAdd dst, src.X(i), src.Y(i)
    Next i
End Sub

Public Sub PolyReverse(ByRef p As Poly)
    Dim i As Long, j As Long, t As Double
    i = 0
    j = p.N - 1
    Do While i < j
        t = p.X(i): p.X(i) = p.X(j): p.X(j) = t
        t = p.Y(i): p.Y(i) = p.Y(j): p.Y(j) = t
        i = i + 1
        j = j - 1
    Loop
End Sub

Public Sub PolyDropClosing(ByRef p As Poly)
    If p.N > 1 Then
        If Abs(p.X(0) - p.X(p.N - 1)) < EPS_ And Abs(p.Y(0) - p.Y(p.N - 1)) < EPS_ Then
            p.N = p.N - 1
        End If
    End If
End Sub

Public Sub PolySetInit(ByRef s As PolySet)
    s.Count = 0
    Erase s.Items
End Sub

Public Sub PolySetAdd(ByRef s As PolySet, ByRef p As Poly)
    s.Count = s.Count + 1
    ReDim Preserve s.Items(s.Count - 1)
    PolyCopy s.Items(s.Count - 1), p
End Sub

Public Sub HoleSetInit(ByRef s As HoleSet)
    s.Count = 0
    Erase s.Items
End Sub

Public Sub HoleSetAdd(ByRef s As HoleSet, ByVal x As Double, ByVal y As Double, ByVal dia As Double)
    s.Count = s.Count + 1
    ReDim Preserve s.Items(s.Count - 1)
    s.Items(s.Count - 1).X = x
    s.Items(s.Count - 1).Y = y
    s.Items(s.Count - 1).Dia = dia
End Sub

' ---------------------------------------------------------------- базовые величины

Public Function GeoSignedArea(ByRef p As Poly) As Double
    Dim i As Long, n As Long, s As Double
    n = p.N
    If n < 3 Then Exit Function
    For i = 0 To n - 1
        s = s + p.X(i) * p.Y((i + 1) Mod n) - p.X((i + 1) Mod n) * p.Y(i)
    Next i
    GeoSignedArea = s / 2#
End Function

Public Function GeoDistPointSeg(ByVal px As Double, ByVal py As Double, _
        ByVal ax As Double, ByVal ay As Double, ByVal bx As Double, ByVal by As Double) As Double
    Dim dx As Double, dy As Double, ll As Double, t As Double
    dx = bx - ax
    dy = by - ay
    ll = dx * dx + dy * dy
    If ll < EPS_ Then
        GeoDistPointSeg = Sqr((px - ax) ^ 2 + (py - ay) ^ 2)
        Exit Function
    End If
    t = ((px - ax) * dx + (py - ay) * dy) / ll
    If t < 0# Then t = 0#
    If t > 1# Then t = 1#
    GeoDistPointSeg = Sqr((px - (ax + dx * t)) ^ 2 + (py - (ay + dy * t)) ^ 2)
End Function

Public Function GeoDistPointPoly(ByVal px As Double, ByVal py As Double, ByRef ring As Poly) As Double
    Dim i As Long, n As Long, d As Double, best As Double
    n = ring.N
    best = 1E+30
    For i = 0 To n - 1
        d = GeoDistPointSeg(px, py, ring.X(i), ring.Y(i), _
            ring.X((i + 1) Mod n), ring.Y((i + 1) Mod n))
        If d < best Then best = d
    Next i
    GeoDistPointPoly = best
End Function

Public Function GeoPointInPoly(ByVal px As Double, ByVal py As Double, ByRef ring As Poly) As Boolean
    Dim i As Long, n As Long, inside As Boolean
    Dim x1 As Double, y1 As Double, x2 As Double, y2 As Double, xc As Double
    n = ring.N
    For i = 0 To n - 1
        x1 = ring.X(i): y1 = ring.Y(i)
        x2 = ring.X((i + 1) Mod n): y2 = ring.Y((i + 1) Mod n)
        If (y1 > py) <> (y2 > py) Then
            xc = x1 + (py - y1) * (x2 - x1) / (y2 - y1)
            If xc > px Then inside = Not inside
        End If
    Next i
    GeoPointInPoly = inside
End Function

Public Function GeoNestingDepth(ByRef rings As PolySet, ByVal index As Long) As Long
    Dim j As Long, depth As Long
    For j = 0 To rings.Count - 1
        If j <> index Then
            If GeoPointInPoly(rings.Items(index).X(0), rings.Items(index).Y(0), rings.Items(j)) Then
                depth = depth + 1
            End If
        End If
    Next j
    GeoNestingDepth = depth
End Function

Private Function GeoDistPointLine(ByVal px As Double, ByVal py As Double, _
        ByVal ax As Double, ByVal ay As Double, ByVal bx As Double, ByVal by As Double) As Double
    Dim dx As Double, dy As Double, len_ As Double
    dx = bx - ax
    dy = by - ay
    len_ = Sqr(dx * dx + dy * dy)
    If len_ < EPS_ Then
        GeoDistPointLine = Sqr((px - ax) ^ 2 + (py - ay) ^ 2)
        Exit Function
    End If
    GeoDistPointLine = Abs(dx * (py - ay) - dy * (px - ax)) / len_
End Function

Private Sub GeoUnit(ByVal vx As Double, ByVal vy As Double, ByRef ux As Double, ByRef uy As Double)
    Dim len_ As Double
    len_ = Sqr(vx * vx + vy * vy)
    If len_ < EPS_ Then
        ux = 0#
        uy = 0#
    Else
        ux = vx / len_
        uy = vy / len_
    End If
End Sub

Private Function Atn2(ByVal y As Double, ByVal x As Double) As Double
    If x > 0# Then
        Atn2 = Atn(y / x)
    ElseIf x < 0# Then
        If y >= 0# Then
            Atn2 = Atn(y / x) + PI_
        Else
            Atn2 = Atn(y / x) - PI_
        End If
    Else
        If y > 0# Then
            Atn2 = PI_ / 2#
        ElseIf y < 0# Then
            Atn2 = -PI_ / 2#
        Else
            Atn2 = 0#
        End If
    End If
End Function

' ---------------------------------------------------------------- кривые

' Аппроксимирует кубическую кривую Безье ломаной с допуском tol.
' Начальная точка не добавляется: она уже в p.
Public Sub GeoFlattenCubic(ByRef p As Poly, _
        ByVal x0 As Double, ByVal y0 As Double, ByVal x1 As Double, ByVal y1 As Double, _
        ByVal x2 As Double, ByVal y2 As Double, ByVal x3 As Double, ByVal y3 As Double, _
        ByVal tol As Double)
    GeoFlattenRec p, x0, y0, x1, y1, x2, y2, x3, y3, tol, 0
End Sub

Private Sub GeoFlattenRec(ByRef p As Poly, _
        ByVal x0 As Double, ByVal y0 As Double, ByVal x1 As Double, ByVal y1 As Double, _
        ByVal x2 As Double, ByVal y2 As Double, ByVal x3 As Double, ByVal y3 As Double, _
        ByVal tol As Double, ByVal depth As Long)
    Dim d1 As Double, d2 As Double
    Dim ax As Double, ay As Double, bx As Double, by As Double, cx As Double, cy As Double
    Dim dx As Double, dy As Double, ex As Double, ey As Double, mx As Double, my As Double

    If depth > 20 Then
        PolyAdd p, x3, y3
        Exit Sub
    End If
    d1 = GeoDistPointLine(x1, y1, x0, y0, x3, y3)
    d2 = GeoDistPointLine(x2, y2, x0, y0, x3, y3)
    If d1 <= tol And d2 <= tol Then
        PolyAdd p, x3, y3
        Exit Sub
    End If

    ax = (x0 + x1) / 2#: ay = (y0 + y1) / 2#
    bx = (x1 + x2) / 2#: by = (y1 + y2) / 2#
    cx = (x2 + x3) / 2#: cy = (y2 + y3) / 2#
    dx = (ax + bx) / 2#: dy = (ay + by) / 2#
    ex = (bx + cx) / 2#: ey = (by + cy) / 2#
    mx = (dx + ex) / 2#: my = (dy + ey) / 2#

    GeoFlattenRec p, x0, y0, ax, ay, dx, dy, mx, my, tol, depth + 1
    GeoFlattenRec p, mx, my, ex, ey, cx, cy, x3, y3, tol, depth + 1
End Sub

' ---------------------------------------------------------------- пересечения

Private Function GeoLineIntersect(ByVal ax As Double, ByVal ay As Double, _
        ByVal adx As Double, ByVal ady As Double, ByVal bx As Double, ByVal by As Double, _
        ByVal bdx As Double, ByVal bdy As Double, ByRef ix As Double, ByRef iy As Double) As Boolean
    Dim den As Double, t As Double
    den = adx * bdy - ady * bdx
    If Abs(den) < 1E-12 Then Exit Function
    t = ((bx - ax) * bdy - (by - ay) * bdx) / den
    ix = ax + adx * t
    iy = ay + ady * t
    GeoLineIntersect = True
End Function

Private Function GeoSegIntersect(ByVal ax As Double, ByVal ay As Double, _
        ByVal bx As Double, ByVal by As Double, ByVal cx As Double, ByVal cy As Double, _
        ByVal dx As Double, ByVal dy As Double, ByRef px As Double, ByRef py As Double) As Boolean
    Dim rx As Double, ry As Double, sx As Double, sy As Double
    Dim den As Double, t As Double, u As Double
    rx = bx - ax: ry = by - ay
    sx = dx - cx: sy = dy - cy
    den = rx * sy - ry * sx
    If Abs(den) < 1E-12 Then Exit Function
    t = ((cx - ax) * sy - (cy - ay) * sx) / den
    u = ((cx - ax) * ry - (cy - ay) * rx) / den
    If t < -EPS_ Or t > 1# + EPS_ Or u < -EPS_ Or u > 1# + EPS_ Then Exit Function
    px = ax + rx * t
    py = ay + ry * t
    If Sqr((px - ax) ^ 2 + (py - ay) ^ 2) < 1E-07 Then Exit Function
    If Sqr((px - bx) ^ 2 + (py - by) ^ 2) < 1E-07 Then Exit Function
    If Sqr((px - cx) ^ 2 + (py - cy) ^ 2) < 1E-07 Then Exit Function
    If Sqr((px - dx) ^ 2 + (py - dy) ^ 2) < 1E-07 Then Exit Function
    GeoSegIntersect = True
End Function

Private Function GeoFirstSelfIntersection(ByRef ring As Poly, ByRef iOut As Long, _
        ByRef jOut As Long, ByRef px As Double, ByRef py As Double) As Boolean
    Dim n As Long, i As Long, j As Long
    n = ring.N
    For i = 0 To n - 1
        For j = i + 2 To n - 1
            If Not (i = 0 And j = n - 1) Then
                If GeoSegIntersect(ring.X(i), ring.Y(i), _
                        ring.X((i + 1) Mod n), ring.Y((i + 1) Mod n), _
                        ring.X(j), ring.Y(j), _
                        ring.X((j + 1) Mod n), ring.Y((j + 1) Mod n), px, py) Then
                    iOut = i
                    jOut = j
                    GeoFirstSelfIntersection = True
                    Exit Function
                End If
            End If
        Next j
    Next i
End Function

Private Sub GeoSplitLoops(ByRef ring As Poly, ByRef outSet As PolySet, ByVal depth As Long)
    Dim i As Long, j As Long, k As Long
    Dim px As Double, py As Double
    Dim a As Poly, b As Poly

    If depth > 8 Or ring.N < 4 Then
        PolySetAdd outSet, ring
        Exit Sub
    End If
    If Not GeoFirstSelfIntersection(ring, i, j, px, py) Then
        PolySetAdd outSet, ring
        Exit Sub
    End If

    PolyInit a
    PolyAdd a, px, py
    For k = i + 1 To j
        PolyAdd a, ring.X(k), ring.Y(k)
    Next k

    PolyInit b
    PolyAdd b, px, py
    For k = j + 1 To ring.N - 1
        PolyAdd b, ring.X(k), ring.Y(k)
    Next k
    For k = 0 To i
        PolyAdd b, ring.X(k), ring.Y(k)
    Next k

    GeoSplitLoops a, outSet, depth + 1
    GeoSplitLoops b, outSet, depth + 1
End Sub

Private Sub GeoAppendArc(ByRef p As Poly, ByVal cx As Double, ByVal cy As Double, _
        ByVal ax As Double, ByVal ay As Double, ByVal bx As Double, ByVal by As Double, _
        ByVal radius As Double)
    Dim a0 As Double, a1 As Double, sweep As Double
    Dim steps As Long, i As Long
    a0 = Atn2(ay - cy, ax - cx)
    a1 = Atn2(by - cy, bx - cx)
    sweep = a1 - a0
    Do While sweep > PI_
        sweep = sweep - 2# * PI_
    Loop
    Do While sweep < -PI_
        sweep = sweep + 2# * PI_
    Loop
    steps = Int(Abs(sweep) / (ARC_STEP_DEG * PI_ / 180#)) + 1
    If steps < 1 Then steps = 1
    For i = 0 To steps
        PolyAdd p, cx + radius * Cos(a0 + sweep * i / steps), _
                  cy + radius * Sin(a0 + sweep * i / steps)
    Next i
End Sub

' ---------------------------------------------------------------- смещение

' Смещает замкнутый контур: delta>0 наружу, delta<0 внутрь.
' Возвращает петли против часовой стрелки; при узких местах петель может быть несколько.
Public Sub GeoOffsetClosed(ByRef src As Poly, ByVal delta As Double, ByRef outSet As PolySet)
    Dim ring As Poly, raw As Poly, parts As PolySet
    Dim i As Long, n As Long, prev As Long, k As Long, j As Long
    Dim dx() As Double, dy() As Double, nx() As Double, ny() As Double
    Dim ax As Double, ay As Double, bx As Double, by As Double
    Dim ix As Double, iy As Double, ad As Double, closest As Double, d As Double
    Dim ok As Boolean, keep As Boolean

    PolySetInit outSet
    PolyCopy ring, src
    PolyDropClosing ring
    If ring.N < 3 Or Abs(delta) < EPS_ Then
        If ring.N >= 3 Then PolySetAdd outSet, ring
        Exit Sub
    End If
    If GeoSignedArea(ring) < 0# Then PolyReverse ring

    n = ring.N
    ReDim dx(n - 1)
    ReDim dy(n - 1)
    ReDim nx(n - 1)
    ReDim ny(n - 1)
    For i = 0 To n - 1
        GeoUnit ring.X((i + 1) Mod n) - ring.X(i), ring.Y((i + 1) Mod n) - ring.Y(i), dx(i), dy(i)
        nx(i) = dy(i)
        ny(i) = -dx(i)
    Next i

    PolyInit raw
    For i = 0 To n - 1
        prev = (i + n - 1) Mod n
        ax = ring.X(i) + nx(prev) * delta
        ay = ring.Y(i) + ny(prev) * delta
        bx = ring.X(i) + nx(i) * delta
        by = ring.Y(i) + ny(i) * delta
        If Abs(dx(prev) * dy(i) - dy(prev) * dx(i)) < 1E-12 Then
            PolyAdd raw, ax, ay
        Else
            ok = GeoLineIntersect(ax, ay, dx(prev), dy(prev), bx, by, dx(i), dy(i), ix, iy)
            If ok Then
                ad = Sqr((ix - ring.X(i)) ^ 2 + (iy - ring.Y(i)) ^ 2)
            Else
                ad = 1E+30
            End If
            If ok And ad <= Abs(delta) * MITRE_LIMIT Then
                PolyAdd raw, ix, iy
            Else
                GeoAppendArc raw, ring.X(i), ring.Y(i), ax, ay, bx, by, Abs(delta)
            End If
        End If
    Next i

    PolySetInit parts
    GeoSplitLoops raw, parts, 0
    If parts.Count = 1 Then
        PolySetAdd outSet, parts.Items(0)
        Exit Sub
    End If

    For k = 0 To parts.Count - 1
        keep = False
        If parts.Items(k).N >= 3 Then
            If Abs(GeoSignedArea(parts.Items(k))) >= 1E-07 Then
                closest = 1E+30
                For j = 0 To parts.Items(k).N - 1
                    d = GeoDistPointPoly(parts.Items(k).X(j), parts.Items(k).Y(j), ring)
                    If d < closest Then closest = d
                Next j
                If closest >= Abs(delta) * JOIN_MIN_FACTOR Then keep = True
            End If
        End If
        If keep Then PolySetAdd outSet, parts.Items(k)
    Next k
End Sub
