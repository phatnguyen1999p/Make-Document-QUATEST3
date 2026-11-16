"""
Utility functions for text normalization, language detection, and date calculations.
"""

from datetime import datetime, timedelta
import re


def remove_accents(input_str: str) -> str:
    """
    Remove Vietnamese accents from a string for case-insensitive matching.
    """
    if not input_str:
        return ""
    s1 = u'ÀÁÂÃÈÉÊÌÍÒÓÔÕÙÚÝàáâãèéêìíòóôõùúýĂăĐđĨĩŨũƠơƯưẠạẢảẤấẦầẨẩẪẫẬậẮắẰằẲẳẴẵẶặẸẹẺẻẼẽẾếỀềỂểỄễỆệỈỉỊịỌọỎỏỐốỒồỔổỖỗỘộỚớỜờỞởỠỡỢợỤụỦủỨứỪừỬửỮữỰựỲỳỴỵỶỷỸỹ'
    s0 = u'AAAAEEEIIOOOOUUYaaaaeeeiioooouuyAaDdIiUuOoUuAaAaAaAaAaAaAaAaAaAaAaAaEeEeEeEeEeEeEeEeIiIiOoOoOoOoOoOoOoOoOoOoOoOoUuUuUuUuUuUuUuYyYyYyYy'
    s = ''
    for c in input_str:
        if c in s1:
            s += s0[s1.index(c)]
        else:
            s += c
    return s


def calculate_target_date(time_rules=None, use_time_rules=False, use_manual_date=False, 
                          manual_date_str=None, current_time=None) -> datetime:
    """
    Calculate target processing date considering manual date overrides, time-based rules, 
    day offsets, and automatic weekend skipping (Saturday & Sunday).
    """
    now = current_time if current_time is not None else datetime.now()

    if use_manual_date and manual_date_str:
        try:
            target = datetime.strptime(str(manual_date_str).strip(), "%d/%m/%Y")
            return target
        except Exception:
            pass

    offset = 0
    if use_time_rules and time_rules:
        try:
            t1_str = str(time_rules.get("t1", "09:00")).strip()
            t2_str = str(time_rules.get("t2", "16:00")).strip()
            t1 = datetime.strptime(t1_str, "%H:%M").time()
            t2 = datetime.strptime(t2_str, "%H:%M").time()
            current_t = now.time()

            x1 = int(time_rules.get("x1", 0))
            x2 = int(time_rules.get("x2", 1))
            x3 = int(time_rules.get("x3", 2))

            if current_t < t1:
                offset = x1
            elif current_t < t2:
                offset = x2
            else:
                offset = x3
        except Exception:
            offset = 0

        target = now + timedelta(days=offset)
        # 5 is Saturday, 6 is Sunday
        if target.weekday() == 5:
            target += timedelta(days=2)
        elif target.weekday() == 6:
            target += timedelta(days=1)
        return target

    return now
