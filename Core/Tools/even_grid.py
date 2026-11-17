# -*- coding: utf-8 -*-
"""
Căn đều vị trí đo trên sơ đồ phòng sạch (người dùng yêu cầu 17/11/2026, hồ sơ 120423 đợt 7):
"sơ đồ đo phòng sạch tự động căn đều các điểm đo, và căn đều theo vách tường của khu vực phòng sạch,
ví dụ như ô lưới 3 4 3 cho phòng sạch 10 điểm".

- Mỗi phòng khai bằng tứ giác 4 góc theo **mặt trong vách** trên ảnh mặt bằng, thứ tự
  trên-trái, trên-phải, dưới-phải, dưới-trái (mặt bằng xoay nghiêng hay phòng hình thang vẫn đúng —
  lưới nội suy song tuyến theo 4 góc nên hàng / cột song song với vách).
- `row_counts(n, w, h)`: số hàng và số vị trí từng hàng sao cho ô lưới gần vuông nhất; phần dư chia đối xứng
  (10 vị trí → 3-4-3, 5 → 2-1-2, 16 → 4-4-4-4, 6 → 3-3 hoặc 2-2-2 theo hình dạng phòng).
- Mỗi vị trí nằm ở tâm ô: cách vách nửa ô, cách nhau một ô (hàng ít vị trí hơn chia đều lại theo bề rộng).
- `snap_to_grid(points, quad)`: gán từng vị trí (nhãn giữ nguyên) vào ô lưới gần chỗ đánh dấu gốc nhất
  (tổng quãng dời nhỏ nhất) — số thứ tự vẫn theo phiếu ghi chép / sơ đồ viết tay.

Dùng trong script hồ sơ trước khi truyền `markers` cho `drawio_layout.ensure_layout` (mẫu `build_120423.py`).
Ô nào rơi vào vật cản (toilet, cầu thang…) thì thu tứ giác lại khỏi vật cản, hoặc người dùng kéo trên draw.io.
"""

from __future__ import annotations

import math
from itertools import permutations

Point = tuple[float, float]
Quad = tuple[Point, Point, Point, Point]


def _split(n: int, rows: int) -> list[int]:
    """Chia n vị trí vào `rows` hàng, phần dư đặt đối xứng (hàng giữa trước khi dư lẻ, hai hàng ngoài khi dư chẵn)."""
    q, r = divmod(n, rows)
    counts = [q] * rows
    if r == 0:
        return counts
    order = []
    mid = (rows - 1) / 2
    idx = sorted(range(rows), key=lambda i: (abs(i - mid), i))
    if r % 2 == 1:
        order.append(idx[0])           # hàng giữa
        rest = [i for i in range(rows) if i != idx[0]]
    else:
        rest = list(range(rows))
    # cặp đối xứng từ ngoài vào trong
    pairs = []
    lo, hi = 0, rows - 1
    while lo < hi:
        if lo in rest and hi in rest:
            pairs.append((lo, hi))
        lo, hi = lo + 1, hi - 1
    need = (r - len(order)) // 2
    for a, b in pairs[:need]:
        order += [a, b]
    for i in order:
        counts[i] += 1
    if sum(counts) != n:                # dự phòng: dồn vào hàng giữa
        counts[idx[0]] += n - sum(counts)
    return counts


def row_counts(n: int, w: float, h: float) -> list[int]:
    """Số vị trí từng hàng (trên → dưới) cho phòng rộng w, cao h: ô lưới gần vuông nhất."""
    if n <= 0:
        return []
    best = None
    for rows in range(1, n + 1):
        counts = _split(n, rows)
        if min(counts) == 0:
            break
        cw, ch = w / max(counts), h / rows
        symmetric = counts == counts[::-1]           # ưu tiên lưới đối xứng (3-4-3, 2-1-2)
        score = (not symmetric, abs(math.log(cw / ch)), max(counts) - min(counts))
        if best is None or score < best[0]:
            best = (score, counts)
    return best[1]


def _bilinear(quad: Quad, s: float, t: float) -> Point:
    (x0, y0), (x1, y1), (x2, y2), (x3, y3) = quad
    top = (x0 + (x1 - x0) * s, y0 + (y1 - y0) * s)
    bot = (x3 + (x2 - x3) * s, y3 + (y2 - y3) * s)
    return (top[0] + (bot[0] - top[0]) * t, top[1] + (bot[1] - top[1]) * t)


def _size(quad: Quad) -> tuple[float, float]:
    d = lambda a, b: math.hypot(a[0] - b[0], a[1] - b[1])
    tl, tr, br, bl = quad
    return (d(tl, tr) + d(bl, br)) / 2, (d(tl, bl) + d(tr, br)) / 2


def grid_slots(n: int, quad: Quad, counts: list[int] | None = None) -> list[Point]:
    """Tâm các ô lưới (trên → dưới, trái → phải) của n vị trí trong tứ giác `quad`."""
    w, h = _size(quad)
    counts = counts or row_counts(n, w, h)
    rows = len(counts)
    out = []
    for i, c in enumerate(counts):
        t = (i + 0.5) / rows
        for j in range(c):
            out.append(_bilinear(quad, (j + 0.5) / c, t))
    return out


def _assign(points: list[Point], slots: list[Point]) -> list[int]:
    """Gán điểm -> ô (chỉ số ô cho từng điểm), tổng bình phương quãng dời nhỏ nhất."""
    n = len(points)
    cost = [[(p[0] - s[0]) ** 2 + (p[1] - s[1]) ** 2 for s in slots] for p in points]
    if n <= 7:
        return list(min(permutations(range(n)), key=lambda perm: sum(cost[i][perm[i]] for i in range(n))))
    # tham lam rồi đổi cặp đến khi không giảm được nữa
    free, asg = set(range(n)), [None] * n
    for i in sorted(range(n), key=lambda i: min(cost[i])):
        j = min(free, key=lambda j: cost[i][j])
        asg[i] = j
        free.remove(j)
    improved = True
    while improved:
        improved = False
        for a in range(n):
            for b in range(a + 1, n):
                if cost[a][asg[b]] + cost[b][asg[a]] < cost[a][asg[a]] + cost[b][asg[b]] - 1e-9:
                    asg[a], asg[b] = asg[b], asg[a]
                    improved = True
    return asg


def snap_to_grid(points: list[Point], quad: Quad, counts: list[int] | None = None) -> list[Point]:
    """Vị trí mới (cùng thứ tự `points`) — mỗi điểm vào ô lưới gần chỗ đánh dấu gốc nhất."""
    slots = grid_slots(len(points), quad, counts)
    asg = _assign(points, slots)
    return [slots[j] for j in asg]


def rect_quad(u0: float, u1: float, v0: float, v1: float, angle_deg: float = 0.0) -> Quad:
    """Tứ giác của hình chữ nhật [u0,u1]×[v0,v1] trong hệ trục xoay `angle_deg` (vách nghiêng xuống bên phải > 0)."""
    c, s = math.cos(math.radians(angle_deg)), math.sin(math.radians(angle_deg))
    f = lambda u, v: (u * c - v * s, u * s + v * c)
    return (f(u0, v0), f(u1, v0), f(u1, v1), f(u0, v1))


if __name__ == "__main__":
    for n, w, h in ((10, 800, 700), (5, 210, 480), (16, 1100, 940), (6, 453, 442), (6, 455, 482), (4, 100, 100)):
        print(n, (w, h), row_counts(n, w, h))
