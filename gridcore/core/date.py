"""Date serial numbers.

The 1900 system counts days from 1899-12-31 (serial 1 = 1900-01-01) and, like Lotus 1-2-3 and
Excel, treats 1900 as a leap year: serial 60 is the fictitious 1900-02-29. The 1904 system counts
from 1904-01-01 (serial 0). Times are fractions of a day.
"""

from __future__ import annotations

import enum
import math
import time
from dataclasses import dataclass
from typing import Optional


class DateSystem(enum.Enum):
    D1900 = "1900"
    D1904 = "1904"


@dataclass
class DateTime:
    year: int
    month: int
    day: int
    hour: int = 0
    minute: int = 0
    second: int = 0
    milli: int = 0
    # 0 = Sunday.
    weekday: int = 0


MONTHS = [
    "January", "February", "March", "April", "May", "June", "July", "August", "September", "October",
    "November", "December",
]
WEEKDAYS = ["Sunday", "Monday", "Tuesday", "Wednesday", "Thursday", "Friday", "Saturday"]


def days_from_civil(y: int, m: int, d: int) -> int:
    """Days since 1970-01-01 (Howard Hinnant's algorithm)."""
    y = y - 1 if m <= 2 else y
    era = (y if y >= 0 else y - 399) // 400
    yoe = y - era * 400
    doy = (153 * (m - 3 if m > 2 else m + 9) + 2) // 5 + d - 1
    doe = yoe * 365 + yoe // 4 - yoe // 100 + doy
    return era * 146097 + doe - 719468


def civil_from_days(z: int) -> tuple[int, int, int]:
    z += 719468
    era = (z if z >= 0 else z - 146096) // 146097
    doe = z - era * 146097
    yoe = (doe - doe // 1460 + doe // 36524 - doe // 146096) // 365
    y = yoe + era * 400
    doy = doe - (365 * yoe + yoe // 4 - yoe // 100)
    mp = (5 * doy + 2) // 153
    d = doy - (153 * mp + 2) // 5 + 1
    m = mp + 3 if mp < 10 else mp - 9
    return (y + 1 if m <= 2 else y), m, d


UNIX_1899_12_30 = -25569  # days_from_civil(1899, 12, 30)
UNIX_1904_01_01 = -24107


def is_leap(y: int) -> bool:
    return (y % 4 == 0 and y % 100 != 0) or y % 400 == 0


def days_in_month(y: int, m: int) -> int:
    if m in (1, 3, 5, 7, 8, 10, 12):
        return 31
    if m in (4, 6, 9, 11):
        return 30
    if m == 2:
        return 29 if is_leap(y) else 28
    return 0


def days_in_month_in(sys: DateSystem, y: int, m: int) -> int:
    """Like :func:`days_in_month`, except that February 1900 has 29 days in the 1900 system."""
    if sys is DateSystem.D1900 and y == 1900 and m == 2:
        return 29
    return days_in_month(y, m)


def serial_from_ymd(sys: DateSystem, year: int, month: int, day: int) -> Optional[float]:
    """Serial for a date. Month and day overflow roll over like ``DATE()`` (month 13 = January of
    the next year, day 0 = the last day of the previous month), in the date system's own calendar.
    ``None`` before the epoch or after 9999-12-31."""
    try:
        year, month, day = int(year), int(month), int(day)
    except (OverflowError, ValueError):
        return None
    if abs(year) > 10**9 or abs(month) > 10**9 or abs(day) > 10**12:
        return None
    y = year + 1900 if 0 <= year < 1900 else year
    m0 = month - 1
    y += m0 // 12
    m = m0 % 12 + 1
    if sys is DateSystem.D1900 and not 1900 <= y <= 9999:
        return None
    if sys is DateSystem.D1904 and not 1904 <= y <= 9999:
        return None
    first = days_from_civil(y, m, 1)
    if sys is DateSystem.D1900:
        # Up to February 1900 the serial is one lower: the fictitious 1900-02-29 comes first.
        first = first - UNIX_1899_12_30 - (1 if (y == 1900 and m <= 2) else 0)
    else:
        first = first - UNIX_1904_01_01
    serial = first + day - 1
    if not 0 <= serial <= max_serial(sys):
        return None
    return float(serial)


def serial_in_system(sys: DateSystem, serial_1900: float) -> float:
    """A 1900-system serial (what the clock gives) in the date system ``sys``."""
    return serial_1900 - 1462 if sys is DateSystem.D1904 else serial_1900


def max_serial(sys: DateSystem) -> int:
    """The serial of 9999-12-31, the last date Excel shows."""
    return 2_958_465 if sys is DateSystem.D1900 else 2_957_003


def datetime_from_serial(sys: DateSystem, serial: float) -> Optional[DateTime]:
    """Date/time from a serial. ``None`` for negative or too-large serials."""
    if not isinstance(serial, (int, float)) or not math.isfinite(serial) or not 0.0 <= serial < max_serial(sys) + 1.0:
        return None
    day = math.floor(serial)
    frac_ms = int(math.floor((serial - day) * 86_400_000.0 + 0.5))
    if frac_ms >= 86_400_000:
        frac_ms -= 86_400_000
        day += 1
    if sys is DateSystem.D1900:
        # Excel's weekdays continue the fictitious calendar: serial 1 is a Sunday.
        wd = (day + 6) % 7
        if day == 60:
            y, m, d = 1900, 2, 29
        elif day == 0:
            y, m, d = 1900, 1, 0
        else:
            real = day + 1 if day < 60 else day
            y, m, d = civil_from_days(real + UNIX_1899_12_30)
    else:
        unix = day + UNIX_1904_01_01
        y, m, d = civil_from_days(unix)
        wd = (unix + 4) % 7
    secs = frac_ms // 1000
    return DateTime(y, m, d, secs // 3600, (secs // 60) % 60, secs % 60, frac_ms % 1000, wd)


def time_fraction(h: float, m: float, s: float) -> float:
    return (h * 3600.0 + m * 60.0 + s) / 86400.0


def unix_now_ms() -> float:
    return time.time() * 1000.0


def serial_from_unix_ms(ms: float) -> float:
    """Date serial (1900 system, with fraction) for a Unix time in milliseconds (UTC)."""
    return 25569.0 + ms / 86_400_000.0


def now_serial(local: bool = True) -> float:
    """The current date-time as a serial. ``local`` applies the machine's time-zone offset (Excel's
    NOW() is local time)."""
    t = time.time()
    if local:
        lt = time.localtime(t)
        t += lt.tm_gmtoff or 0
    return serial_from_unix_ms(t * 1000.0)


def current_year() -> int:
    return time.localtime().tm_year
