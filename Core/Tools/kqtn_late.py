# -*- coding: utf-8 -*-
"""KQTN có trễ hạn hay không (AGENTS mục 13 — người dùng chốt 12/11/2026).

"Không đánh trễ kết quả thử nghiệm dưới 10 ngày làm việc": KQTN về trong vòng 10 ngày làm việc kể từ
ngày gửi mẫu thì **không** trễ hạn, dù ngày KQTN sau "Ngày yêu cầu cấp KQTN". Chỉ trễ hạn khi cả hai:
  - ngày KQTN sau ngày yêu cầu cấp KQTN (nếu có), và
  - số ngày làm việc từ ngày gửi mẫu đến ngày KQTN > 10.
Ngày làm việc = thứ Hai – thứ Sáu, trừ ngày nghỉ lễ trong NGHI_LE; đếm từ ngày sau ngày gửi mẫu đến hết
ngày KQTN.

    python Core/Tools/kqtn_late.py <ngày gửi mẫu> <ngày KQTN> [ngày yêu cầu]
"""
import datetime as _dt
import sys

GIOI_HAN_NGAY_LAM_VIEC = 10
# Ngày nghỉ lễ (dd/mm/yyyy) không tính là ngày làm việc — thêm khi có hồ sơ rơi vào dịp lễ.
NGHI_LE = {"01/01/2026", "02/09/2026", "01/01/2027"}

CO = "[X] có\t\t[  ] không"
KHONG = "[  ] có\t\t[X] không"


def _d(s):
    if isinstance(s, _dt.date):
        return s
    return _dt.datetime.strptime(s.strip(), "%d/%m/%Y").date()


def working_days(start, end) -> int:
    """Số ngày làm việc sau `start` đến hết `end`."""
    a, b = _d(start), _d(end)
    hol = {_d(x) for x in NGHI_LE}
    n, d = 0, a
    while d < b:
        d += _dt.timedelta(days=1)
        if d.weekday() < 5 and d not in hol:
            n += 1
    return n


def kqtn_late(ngay_kqtn, ngay_gui, ngay_yc=None, limit: int = GIOI_HAN_NGAY_LAM_VIEC) -> bool:
    if ngay_yc and _d(ngay_kqtn) <= _d(ngay_yc):
        return False
    return working_days(ngay_gui, ngay_kqtn) > limit


def tre_han_text(ngay_kqtn, ngay_gui, ngay_yc=None, placeholder="………") -> str:
    """Ô 'KQTN có trễ hạn …' của BCKQ; thiếu ngày KQTN / gửi mẫu → placeholder."""
    if not ngay_kqtn or not ngay_gui:
        return placeholder
    return CO if kqtn_late(ngay_kqtn, ngay_gui, ngay_yc) else KHONG


if __name__ == "__main__":
    sys.stdout.reconfigure(encoding="utf-8")
    gui, kq = sys.argv[1], sys.argv[2]
    yc = sys.argv[3] if len(sys.argv) > 3 else None
    print(f"{working_days(gui, kq)} ngày làm việc -> trễ hạn: {'có' if kqtn_late(kq, gui, yc) else 'không'}")
