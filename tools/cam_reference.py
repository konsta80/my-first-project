"""Эталонная реализация геометрии и постпроцессора CAM.

Макрос CorelDRAW (src/*.bas) этот файл не использует. Он нужен для двух задач:
  1. проверить алгоритмы (смещение контура, разбор самопересечений, пост)
     до переноса в VBA, где нет возможности прогнать код в этой среде;
  2. получить эталонные значения, с которыми сверяется modCamSelfTest.bas.

Запуск: python3 tools/cam_reference.py
"""

from __future__ import annotations

import math

EPS = 1e-9
MITRE_LIMIT = 4.0          # предел "митры" в долях смещения
ARC_STEP_DEG = 12.0        # шаг скругления на внешних углах
JOIN_MIN_FACTOR = 0.6      # порог отбраковки артефактов при самопересечении

Point = tuple


# ---------------------------------------------------------------- векторы

def sub(a, b):
    return (a[0] - b[0], a[1] - b[1])


def add(a, b):
    return (a[0] + b[0], a[1] + b[1])


def mul(a, s):
    return (a[0] * s, a[1] * s)


def cross(a, b):
    return a[0] * b[1] - a[1] * b[0]


def length(a):
    return math.hypot(a[0], a[1])


def norm(a):
    n = length(a)
    return (0.0, 0.0) if n < EPS else (a[0] / n, a[1] / n)


def signed_area(pts):
    s = 0.0
    n = len(pts)
    for i in range(n):
        x1, y1 = pts[i]
        x2, y2 = pts[(i + 1) % n]
        s += x1 * y2 - x2 * y1
    return s / 2.0


def dedupe(pts, tol=1e-6):
    out = []
    for p in pts:
        if not out or length(sub(p, out[-1])) > tol:
            out.append(p)
    while len(out) > 1 and length(sub(out[0], out[-1])) <= tol:
        out.pop()
    return out


# ---------------------------------------------------------------- дуги/кривые

def _dist_point_line(p, a, b):
    d = sub(b, a)
    if length(d) < EPS:
        return length(sub(p, a))
    return abs(cross(d, sub(p, a))) / length(d)


def flatten_cubic(p0, p1, p2, p3, tol=0.05):
    """Кубическая кривая Безье в ломаную; допуск — отклонение от хорды."""
    out = [p0]
    _flatten(p0, p1, p2, p3, tol, out, 0)
    return out


def _flatten(p0, p1, p2, p3, tol, out, depth):
    if depth > 20 or (max(_dist_point_line(p1, p0, p3),
                          _dist_point_line(p2, p0, p3)) <= tol):
        out.append(p3)
        return
    p01 = mul(add(p0, p1), 0.5)
    p12 = mul(add(p1, p2), 0.5)
    p23 = mul(add(p2, p3), 0.5)
    p012 = mul(add(p01, p12), 0.5)
    p123 = mul(add(p12, p23), 0.5)
    mid = mul(add(p012, p123), 0.5)
    _flatten(p0, p01, p012, mid, tol, out, depth + 1)
    _flatten(mid, p123, p23, p3, tol, out, depth + 1)


# ---------------------------------------------------------------- расстояния/принадлежность

def dist_point_segment(p, a, b):
    d = sub(b, a)
    ll = d[0] * d[0] + d[1] * d[1]
    if ll < EPS:
        return length(sub(p, a))
    t = ((p[0] - a[0]) * d[0] + (p[1] - a[1]) * d[1]) / ll
    t = 0.0 if t < 0.0 else (1.0 if t > 1.0 else t)
    return length(sub(p, add(a, mul(d, t))))


def dist_point_polyline(p, ring):
    n = len(ring)
    return min(dist_point_segment(p, ring[i], ring[(i + 1) % n]) for i in range(n))


def point_in_ring(p, ring):
    inside = False
    n = len(ring)
    for i in range(n):
        x1, y1 = ring[i]
        x2, y2 = ring[(i + 1) % n]
        if (y1 > p[1]) != (y2 > p[1]):
            xc = x1 + (p[1] - y1) * (x2 - x1) / (y2 - y1)
            if xc > p[0]:
                inside = not inside
    return inside


def nesting_depth(rings, index):
    """Сколько других контуров содержат данный: чётное — тело, нечётное — отверстие."""
    probe = rings[index][0]
    depth = 0
    for j, other in enumerate(rings):
        if j != index and point_in_ring(probe, other):
            depth += 1
    return depth


# ---------------------------------------------------------------- смещение контура

def _line_intersection(a, da, b, db):
    den = cross(da, db)
    if abs(den) < 1e-12:
        return None
    t = cross(sub(b, a), db) / den
    return add(a, mul(da, t))


def _arc(center, a, b, radius, step_deg):
    a0 = math.atan2(a[1] - center[1], a[0] - center[0])
    a1 = math.atan2(b[1] - center[1], b[0] - center[0])
    sweep = a1 - a0
    while sweep > math.pi:
        sweep -= 2 * math.pi
    while sweep < -math.pi:
        sweep += 2 * math.pi
    steps = max(1, int(abs(sweep) / math.radians(step_deg)) + 1)
    return [(center[0] + radius * math.cos(a0 + sweep * i / steps),
             center[1] + radius * math.sin(a0 + sweep * i / steps))
            for i in range(steps + 1)]


def _seg_intersection(a, b, c, d):
    r = sub(b, a)
    s = sub(d, c)
    den = cross(r, s)
    if abs(den) < 1e-12:
        return None
    t = cross(sub(c, a), s) / den
    u = cross(sub(c, a), r) / den
    if -1e-9 <= t <= 1 + 1e-9 and -1e-9 <= u <= 1 + 1e-9:
        p = add(a, mul(r, t))
        if min(length(sub(p, a)), length(sub(p, b)),
               length(sub(p, c)), length(sub(p, d))) > 1e-7:
            return p
    return None


def _first_self_intersection(ring):
    n = len(ring)
    for i in range(n):
        a, b = ring[i], ring[(i + 1) % n]
        for j in range(i + 2, n):
            if i == 0 and j == n - 1:
                continue
            c, d = ring[j], ring[(j + 1) % n]
            p = _seg_intersection(a, b, c, d)
            if p is not None:
                return i, j, p
    return None


def split_self_intersections(ring, depth=0):
    """Разрезает кольцо на простые петли по точкам самопересечения."""
    if depth > 8 or len(ring) < 4:
        return [ring]
    hit = _first_self_intersection(ring)
    if hit is None:
        return [ring]
    i, j, p = hit
    loop_a = dedupe([p] + ring[i + 1:j + 1])
    loop_b = dedupe([p] + ring[j + 1:] + ring[:i + 1])
    return (split_self_intersections(loop_a, depth + 1)
            + split_self_intersections(loop_b, depth + 1))


def offset_closed(pts, delta):
    """Смещение замкнутого контура. delta>0 — наружу, delta<0 — внутрь.

    Возвращает список замкнутых петель в порядке против часовой стрелки.
    """
    ring = dedupe(pts)
    if len(ring) < 3 or abs(delta) < EPS:
        return [ring] if len(ring) >= 3 else []
    if signed_area(ring) < 0:
        ring = list(reversed(ring))

    n = len(ring)
    dirs, normals = [], []
    for i in range(n):
        d = norm(sub(ring[(i + 1) % n], ring[i]))
        dirs.append(d)
        normals.append((d[1], -d[0]))     # правая нормаль = наружу для CCW

    raw = []
    for i in range(n):
        prev = (i - 1) % n
        v = ring[i]
        a = add(v, mul(normals[prev], delta))
        b = add(v, mul(normals[i], delta))
        if abs(cross(dirs[prev], dirs[i])) < 1e-12:
            raw.append(a)
            continue
        p = _line_intersection(a, dirs[prev], b, dirs[i])
        if p is not None and length(sub(p, v)) <= abs(delta) * MITRE_LIMIT:
            raw.append(p)
        else:
            raw.extend(_arc(v, a, b, abs(delta), ARC_STEP_DEG))

    raw = dedupe(raw)
    cleaned = split_self_intersections(raw)
    if len(cleaned) == 1 and cleaned[0] is raw:
        return [raw]

    kept = []
    for loop in cleaned:
        if len(loop) < 3 or abs(signed_area(loop)) < 1e-7:
            continue
        closest = min(dist_point_polyline(p, ring) for p in loop)
        if closest >= abs(delta) * JOIN_MIN_FACTOR:
            kept.append(loop)
    return kept


def pocket_loops(pts, radius, stepover, max_loops=200):
    """Концентрические проходы кармана внутрь на радиус + шаг."""
    loops = []
    k = 0
    while k < max_loops:
        d = radius + k * stepover
        current = offset_closed(pts, -d)
        if not current:
            break
        loops.append((d, current))
        k += 1
    return loops


def vcarve_loops(pts, radius, stepover, half_angle_deg, max_depth, max_loops=200):
    """Гравировка V-битой: глубина петли из её отступа от контура."""
    tan_half = math.tan(math.radians(half_angle_deg))
    out = []
    k = 0
    while k < max_loops:
        d = radius + k * stepover
        current = offset_closed(pts, -d)
        if not current:
            break
        depth = min(d / tan_half, max_depth)
        out.append((d, depth, current))
        k += 1
    out.sort(key=lambda item: -item[1])
    return out


# ---------------------------------------------------------------- постпроцессор

def fmt(v, decimals=3):
    """Число в G-коде: точка как разделитель, без хвостовых нулей."""
    s = f"{v:.{decimals}f}"
    if "." in s:
        s = s.rstrip("0").rstrip(".")
    if s in ("-0", ""):
        s = "0"
    return s


class Post:
    """Syntec 60CB (Fanuc-совместимый ISO) с автосменой инструмента."""

    def __init__(self, program=1001, name="PART", tool=1, rpm=18000,
                 safe_z=20.0, coolant=True, planes="G21 G17"):
        self.lines = []
        self.program = program
        self.name = name
        self.tool = tool
        self.rpm = rpm
        self.safe_z = safe_z
        self.coolant = coolant
        self.planes = planes
        self._feed = None

    def header(self):
        self.lines.append("%")
        self.lines.append(f"O{self.program}")
        self.lines.append(f"({self.name})")
        self.lines.append(f"{self.planes} G40 G49 G80 G90")
        self.lines.append("G54")
        self.lines.append(f"M06 T{self.tool}")
        self.lines.append(f"S{self.rpm} M03")
        self.lines.append(f"G43 H{self.tool} Z{fmt(self.safe_z)}")
        if self.coolant:
            self.lines.append("M08")
        self._feed = None

    def footer(self):
        if self.coolant:
            self.lines.append("M09")
        self.lines.append("M05")
        self.lines.append(f"G00 Z{fmt(self.safe_z)}")
        self.lines.append("G91 G28 Z0.")
        self.lines.append("G90")
        self.lines.append("M30")
        self.lines.append("%")

    def rapid(self, x=None, y=None, z=None):
        words = []
        if x is not None:
            words.append("X" + fmt(x))
        if y is not None:
            words.append("Y" + fmt(y))
        if z is not None:
            words.append("Z" + fmt(z))
        if words:
            self.lines.append("G00 " + " ".join(words))

    def cut(self, x=None, y=None, z=None, feed=None):
        words = []
        if x is not None:
            words.append("X" + fmt(x))
        if y is not None:
            words.append("Y" + fmt(y))
        if z is not None:
            words.append("Z" + fmt(z))
        if feed is not None and feed != self._feed:
            words.append("F" + fmt(feed, 1))
            self._feed = feed
        if words:
            self.lines.append("G01 " + " ".join(words))

    def drill(self, x, y, depth, r_plane, feed):
        self.lines.append(
            f"G81 X{fmt(x)} Y{fmt(y)} Z{fmt(depth)} R{fmt(r_plane)} F{fmt(feed, 1)}")

    def end_cycle(self):
        self.lines.append("G80")

    def text(self):
        return "\r\n".join(self.lines) + "\r\n"


def contour_gcode(ring, radius, depth, step_down, feed, plunge, program=1001):
    post = Post(program=program, name="CONTOUR", tool=1)
    post.header()
    loops = offset_closed(ring, radius)
    for loop in loops:
        post.rapid(x=loop[0][0], y=loop[0][1])
        post.rapid(z=3.0)
        z = -abs(step_down)
        while True:
            z = max(z, -abs(depth))
            post.cut(z=z, feed=plunge)
            for p in loop[1:]:
                post.cut(x=p[0], y=p[1], feed=feed)
            post.cut(x=loop[0][0], y=loop[0][1], feed=feed)
            if z <= -abs(depth) + EPS:
                break
            z = z - abs(step_down)
        post.rapid(z=post.safe_z)
    post.footer()
    return post.text()


# ---------------------------------------------------------------- самопроверка

def _rect(w, h):
    return [(0.0, 0.0), (w, 0.0), (w, h), (0.0, h)]


def _circle(r, segments=96):
    return [(r * math.cos(2 * math.pi * i / segments),
             r * math.sin(2 * math.pi * i / segments)) for i in range(segments)]


def _dev_range(ring, ring_ref):
    ds = [dist_point_polyline(p, ring_ref) for p in ring]
    return min(ds), max(ds)


def _edge_mid_dev(ring, ring_ref):
    """Отклонение середин рёбер: у митрованного угла вершина дальше, ребро — нет."""
    n = len(ring)
    mids = [mul(add(ring[i], ring[(i + 1) % n]), 0.5) for i in range(n)]
    ds = [dist_point_polyline(p, ring_ref) for p in mids]
    return min(ds), max(ds)


class Checker:
    def __init__(self):
        self.failed = 0
        self.passed = 0

    def check(self, name, ok, detail=""):
        if ok:
            self.passed += 1
            print(f"  PASS  {name}")
        else:
            self.failed += 1
            print(f"  FAIL  {name}  {detail}")

    def near(self, name, got, want, tol=1e-6):
        self.check(name, abs(got - want) <= tol, f"got {got}, want {want}")


def selftest():
    c = Checker()
    rect = _rect(100, 60)

    print("контур: смещение наружу на 3 мм")
    out = offset_closed(rect, 3.0)
    c.check("одна петля", len(out) == 1, f"петель {len(out)}")
    if out:
        ring = out[0]
        c.check("4 вершины", len(ring) == 4, f"вершин {len(ring)}")
        lo, hi = _edge_mid_dev(ring, rect)
        c.near("кромки отстоят на 3 мм", lo, 3.0, 1e-9)
        c.near("кромки отстоят на 3 мм (макс)", hi, 3.0, 1e-9)
        lo, hi = _dev_range(ring, rect)
        c.near("угол по митре 3*sqrt(2)", hi, 3.0 * 2 ** 0.5, 1e-9)
        c.near("площадь 106x66", abs(signed_area(ring)), 106 * 66, 1e-6)

    print("контур: смещение внутрь на 3 мм")
    out = offset_closed(rect, -3.0)
    c.check("одна петля", len(out) == 1, f"петель {len(out)}")
    if out:
        lo, hi = _dev_range(out[0], rect)
        c.near("все точки на 3 мм", lo, 3.0, 1e-9)
        c.near("углы срезаны по 3 мм", hi, 3.0, 1e-9)
        c.near("площадь 94x54", abs(signed_area(out[0])), 94 * 54, 1e-6)

    print("круг r=25: смещение внутрь на 6 мм")
    circle = _circle(25, 180)
    out = offset_closed(circle, -6.0)
    c.check("одна петля", len(out) == 1, f"петель {len(out)}")
    if out:
        radii = [length(p) for p in out[0]]
        c.check("радиус 19", all(abs(r - 19) < 0.02 for r in radii),
                f"мин {min(radii):.4f} макс {max(radii):.4f}")

    print("вложенность контуров")
    rings = [_rect(100, 60),
             [(20.0, 20.0), (40.0, 20.0), (40.0, 40.0), (20.0, 40.0)]]
    c.check("тело = 0", nesting_depth(rings, 0) == 0)
    c.check("отверстие = 1", nesting_depth(rings, 1) == 1)

    print("карман 40x40, фреза 6, шаг 3")
    loops = pocket_loops(_rect(40, 40), 3.0, 3.0)
    c.check("есть проходы", len(loops) > 3, f"проходов {len(loops)}")
    c.near("первый отступ", loops[0][0], 3.0)

    print("V-бита 90°: глубина = отступ")
    vc = vcarve_loops(_rect(40, 40), 3.0, 3.0, 45.0, 8.0)
    c.check("глубина не больше предела",
            all(depth <= 8.0 + 1e-9 for _, depth, _ in vc))
    c.check("сортировка от глубокого",
            all(vc[i][1] >= vc[i + 1][1] for i in range(len(vc) - 1)))

    print("постпроцессор Syntec 60CB")
    code = contour_gcode(rect, 3.0, 6.0, 2.0, 1200, 300)
    lines = code.splitlines()
    c.check("начинается с %", lines[0] == "%")
    c.check("заканчивается %", lines[-1] == "%")
    c.check("номер программы", lines[1] == "O1001")
    c.check("мм/абсолют", any("G21" in ln and "G90" in ln for ln in lines))
    c.check("смена инструмента", "M06 T1" in lines)
    c.check("корректор длины", any(ln.startswith("G43 H1") for ln in lines))
    c.check("шпиндель", any(ln.startswith("S18000 M03") for ln in lines))
    c.check("конец программы", "M30" in lines)
    c.check("нет запятых", "," not in code.replace("(", "").replace(")", ""))
    z_levels = [ln for ln in lines if "G01" in ln and " Z-" in ln]
    c.check("три прохода по глубине", len(z_levels) == 3, f"проходов {len(z_levels)}")

    print()
    print(f"итог: {c.passed} passed, {c.failed} failed")
    return 0 if c.failed == 0 else 1


if __name__ == "__main__":
    raise SystemExit(selftest())
