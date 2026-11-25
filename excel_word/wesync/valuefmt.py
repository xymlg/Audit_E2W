"""数值与文本互转：Excel 显示文本渲染、列宽估算、轻量公式求值与公式缓存读取。"""

import re
from datetime import date, datetime, time, timedelta
from decimal import Decimal, ROUND_HALF_UP
from pathlib import Path
from openpyxl import load_workbook
from openpyxl.styles.numbers import is_date_format
from openpyxl.utils import column_index_from_string
from openpyxl.utils.datetime import from_excel
from .config import ARITH_OK_RE, CELL_REF_RE, NUMBER_RE, RANGE_REF_RE, SHEET_NAME, SUM_RE
from .console import warn


def decimals_of_format(number_format: str | None) -> int:
    if not number_format:
        return 0
    section = str(number_format).split(";")[0]
    if "." not in section:
        return 0
    frac = section.split(".", 1)[1]
    count = 0
    for ch in frac:
        if ch in "0#":
            count += 1
        elif ch in "%":
            continue
        else:
            break
    return count


def to_excel_value(text: str) -> tuple[object, str | None]:
    """Word 单元格文本 → (Excel 值, 数字格式)。数字文本转真数值，其余保持文本。"""
    stripped = text.strip()
    if stripped == "":
        return None, None
    if NUMBER_RE.match(stripped):
        raw = stripped.replace(",", "")
        decimals = len(raw.split(".")[1]) if "." in raw else 0
        try:
            value = float(raw)
        except ValueError:
            return text, None
        number_format = "#,##0" if decimals == 0 else "#,##0." + "0" * decimals
        return value, number_format
    return text, None


def display_text(value: object, number_format: str | None) -> str:
    """Excel 单元格 → 回写 Word 的显示文本（模拟 Excel 的显示效果）。

    覆盖审计常用格式子集：日期/时间（含中文年月日字面量）、千分位有无、小数位、
    百分比、负数括号段、零值段；子集外的格式回退旧口径并按格式警告一次。
    """
    if value is None:
        return ""
    if isinstance(value, bool):
        return "TRUE" if value else "FALSE"
    if isinstance(value, timedelta):
        # openpyxl 把 [h]:mm 等耗时格式读成 timedelta；还原为序列值走下方数字/回退口径
        value = value.days + value.seconds / 86400 + value.microseconds / 86400e6
    fmt = str(number_format) if number_format else "General"
    if isinstance(value, (datetime, date, time)):
        if isinstance(value, datetime):
            dt = value
        elif isinstance(value, date):
            dt = datetime.combine(value, datetime.min.time())
        else:
            dt = datetime.combine(date.min, value)
        rendered = _render_date(fmt, dt)
        if rendered is not None:
            return rendered
        _warn_format(fmt)
        return str(value)
    if isinstance(value, (int, float)):
        if is_date_format(fmt):
            try:
                dt = from_excel(value)
            except (ValueError, OverflowError, TypeError):
                dt = None
            if dt is not None:
                rendered = _render_date(fmt, dt)
                if rendered is not None:
                    return rendered
            _warn_format(fmt)
            return _legacy_number(value, number_format)
        if fmt == "General":
            if isinstance(value, float) and value.is_integer():
                return str(int(value))
            return str(value)
        sections = _split_sections(fmt)
        if value < 0 and len(sections) > 1:
            rendered = _render_number(sections[1], abs(value))
        elif value == 0 and len(sections) > 2:
            rendered = _render_number(sections[2], 0.0)
        else:
            rendered = _render_number(sections[0], value)
        if rendered is not None:
            return rendered
        _warn_format(fmt)
        return _legacy_number(value, number_format)
    return str(value)


_WARNED_FORMATS: set[str] = set()


def _warn_format(fmt: str) -> None:
    if fmt in _WARNED_FORMATS:
        return
    _WARNED_FORMATS.add(fmt)
    warn(f"单元格格式 {fmt} 超出内置显示渲染子集，按千分位+小数位旧口径回写，请核对 Word 显示效果。")


def _legacy_number(value: float, number_format: str | None) -> str:
    decimals = decimals_of_format(number_format)
    if decimals:
        return f"{value:,.{decimals}f}"
    return f"{value:,.0f}"


def _split_sections(fmt: str) -> list[str]:
    sections: list[str] = []
    buf: list[str] = []
    quoted = escaped = False
    for ch in fmt:
        if escaped:
            escaped = False
        elif ch == "\\":
            escaped = True
        elif ch == '"':
            quoted = not quoted
        elif ch == ";" and not quoted:
            sections.append("".join(buf))
            buf = []
            continue
        buf.append(ch)
    sections.append("".join(buf))
    return sections


def _tokens(section: str) -> list[tuple[str, str]] | None:
    """格式段 → 令牌流；遇到子集外记号返回 None（调用方回退+警告）。"""
    toks: list[tuple[str, str]] = []
    i = 0
    while i < len(section):
        ch = section[i]
        if ch == '"':
            j = section.find('"', i + 1)
            if j < 0:
                return None
            toks.append(("lit", section[i + 1 : j]))
            i = j + 1
        elif ch == "\\":
            if i + 1 >= len(section):
                return None
            toks.append(("lit", section[i + 1]))
            i += 2
        elif ch in "0#":
            toks.append(("dig", ch))
            i += 1
        elif ch in ",.%":
            toks.append((ch, ch))
            i += 1
        elif ch in "ymdhs":
            toks.append((ch, ch))
            i += 1
        elif ch == "@":
            toks.append(("at", ch))
            i += 1
        elif ch == "[":
            j = section.find("]", i + 1)
            if j < 0:
                return None
            if not re.fullmatch(r"\$-[0-9A-Fa-f]+", section[i + 1 : j]):
                return None
            i = j + 1  # 区域标记（如 [$-804]），无可见字符
        elif ch in "_*":
            i += 2  # 会计对齐填充（_x 留白、*x 填充符），不产生可见文本
        elif ch == "?":
            toks.append(("pad", ch))
            i += 1
        elif ch in "eE" and i + 1 < len(section) and section[i + 1] in "+-":
            return None
        else:
            toks.append(("lit", ch))
            i += 1
    return toks


def _render_number(section: str, value: float) -> str | None:
    toks = _tokens(section)
    if toks is None:
        return None
    if any(k in ("y", "m", "d", "h", "s", "at") for k, _ in toks):
        return None
    if sum(1 for k, _ in toks if k == "%") > 1:
        return None
    scaled = value * 100 if any(k == "%" for k, _ in toks) else value
    dot_at = next((i for i, (k, _) in enumerate(toks) if k == "."), None)
    int_toks = toks if dot_at is None else toks[:dot_at]
    frac_toks = [] if dot_at is None else toks[dot_at + 1 :]
    if any(k == "." for k, _ in frac_toks):
        return None
    int_digs = [t for k, t in int_toks if k == "dig"]
    frac_digs = [t for k, t in frac_toks if k == "dig"]
    if not int_digs and not frac_digs:
        return "".join(t for k, t in toks if k == "lit")  # 零值段纯字面量（如 —）
    # 整数位数字必须连续（中间夹字面量的怪格式回退）
    kinds = [k for k, _ in int_toks if k in ("dig", "lit", ",")]
    first = kinds.index("dig") if "dig" in kinds else len(kinds)
    last = len(kinds) - 1 - kinds[::-1].index("dig") if "dig" in kinds else -1
    if any(k == "lit" for k in kinds[first : last + 1]):
        return None
    trailing_commas = 0
    for k, _ in reversed(int_toks):
        if k == ",":
            trailing_commas += 1
        elif k == "dig":
            break
    thousands = any(k == "," for k, _ in int_toks[: len(int_toks) - trailing_commas]) and bool(int_digs)
    scaled = scaled / (1000 ** trailing_commas)
    decimals = len(frac_digs)
    quantum = Decimal(1).scaleb(-decimals)
    rounded = Decimal(str(scaled)).quantize(quantum, rounding=ROUND_HALF_UP)
    sign = "-" if rounded < 0 else ""
    rounded = abs(rounded)
    digits = f"{rounded:.{decimals}f}"
    int_part, _, frac_part = digits.partition(".")
    int_zeros = sum(1 for t in int_digs if t == "0")
    int_part = int_part.zfill(max(1, int_zeros))
    if thousands:
        int_part = f"{int(int_part):,}"
    if frac_digs:
        keep = decimals
        while keep > 0 and frac_part[keep - 1] == "0" and frac_digs[keep - 1] == "#":
            keep -= 1
        frac_part = frac_part[:keep]
    out: list[str] = []
    int_done = frac_done = False
    for idx, (k, t) in enumerate(toks):
        if k == "lit":
            out.append(t)
        elif k == "dig":
            if dot_at is None or idx < dot_at:
                if not int_done:
                    out.append(sign + int_part)
                    int_done = True
            elif not frac_done:
                out.append(frac_part)
                frac_done = True
        elif k == ".":
            if frac_part:
                out.append(".")
        elif k == "%":
            out.append("%")
    if not int_done and int_digs:
        out.insert(0, sign + int_part)
    return "".join(out)


def _render_date(fmt: str, dt: datetime) -> str | None:
    if "AM/PM" in fmt or "am/pm" in fmt:
        return None
    sections = _split_sections(fmt)
    toks = _tokens(sections[0])
    if toks is None:
        return None
    if any(k in ("dig", ",", "%", "at") for k, _ in toks):
        return None
    out: list[str] = []
    i = 0
    while i < len(toks):
        k, t = toks[i]
        if k == "lit":
            out.append(t)
            i += 1
            continue
        run = 1
        while i + run < len(toks) and toks[i + run][0] == k:
            run += 1
        prev_k = next((toks[j][0] for j in range(i - 1, -1, -1) if toks[j][0] != "lit"), "")
        next_k = next((toks[j][0] for j in range(i + run, len(toks)) if toks[j][0] != "lit"), "")
        if k == "y":
            if run == 2:
                out.append(f"{dt.year % 100:02d}")
            elif run == 4:
                out.append(f"{dt.year:04d}")
            else:
                return None
        elif k == "m":
            minutes = prev_k == "h" or next_k == "s"
            unit = dt.minute if minutes else dt.month
            if run == 1:
                out.append(str(unit))
            elif run == 2:
                out.append(f"{unit:02d}")
            else:
                return None
        elif k == "d":
            if run == 1:
                out.append(str(dt.day))
            elif run == 2:
                out.append(f"{dt.day:02d}")
            else:
                return None
        elif k == "h":
            if run == 1:
                out.append(str(dt.hour))
            elif run == 2:
                out.append(f"{dt.hour:02d}")
            else:
                return None
        elif k == "s":
            if run == 1:
                out.append(str(dt.second))
            elif run == 2:
                out.append(f"{dt.second:02d}")
            else:
                return None
        else:
            return None
        i += run
    return "".join(out)


def approx_width(text: str) -> float:
    width = 0.0
    for ch in text:
        width += 2.0 if ord(ch) > 0x2E80 else 1.0
    return width


def formula_map(xlsx_path: Path) -> dict[tuple[int, int], object]:
    """读取带缓存值的副本，用于处理用户在 Excel 中写的公式。"""
    cached = load_workbook(str(xlsx_path), data_only=True)
    ws = cached[SHEET_NAME] if SHEET_NAME in cached.sheetnames else cached.worksheets[0]
    result = {}
    for row in ws.iter_rows():
        for cell in row:
            if cell.value is not None:
                result[(cell.row, cell.column)] = cell.value
    return result


def formula_result(ws, row: int, col: int, cached_ws=None):
    """取公式计算结果：优先 Excel 缓存值，否则尝试轻量求值；无结果为 None。"""
    if cached_ws is not None:
        value = cached_ws.cell(row, col).value
        if value is not None:
            return value
    return evaluate_formula(ws, row, col, cached_ws)


def _cached_value(cached, row: int, col: int):
    if cached is None:
        return None
    if hasattr(cached, "cell"):
        return cached.cell(row, col).value
    return cached.get((row, col))


def _cell_numeric(ws, row: int, col: int, cached, depth: int) -> float:
    cell = ws.cell(row, col)
    value = _cached_value(cached, row, col)
    if value is None:
        value = cell.value
    if isinstance(value, str) and value.startswith("="):
        if depth > 5:
            return 0.0
        resolved = evaluate_formula(ws, row, col, cached, depth + 1)
        return float(resolved) if isinstance(resolved, (int, float)) else 0.0
    if isinstance(value, (int, float)) and not isinstance(value, bool):
        return float(value)
    return 0.0


def _range_total(ws, ref1: str, ref2: str, cached, depth: int):
    start = CELL_REF_RE.match(ref1)
    stop = CELL_REF_RE.match(ref2)
    if not start or not stop:
        return None
    c1 = column_index_from_string(start.group(1).upper())
    c2 = column_index_from_string(stop.group(1).upper())
    r1, r2 = int(start.group(2)), int(stop.group(2))
    total = 0.0
    for row in range(min(r1, r2), max(r1, r2) + 1):
        for col in range(min(c1, c2), max(c1, c2) + 1):
            total += _cell_numeric(ws, row, col, cached, depth)
    return total


def evaluate_formula(ws, row: int, col: int, cached=None, depth: int = 0):
    """轻量公式求值：支持 SUM(范围/单元格) 与简单四则运算，算不出时返回 None。

    用于 openpyxl 保存后公式缓存值丢失的情况（常见于“合计”行）。
    """
    cell = ws.cell(row, col)
    value = cell.value
    if not (isinstance(value, str) and value.startswith("=")):
        return None
    expression = value[1:].strip()
    sum_match = SUM_RE.match(expression)
    if sum_match:
        total = 0.0
        args = [item.strip() for item in sum_match.group(1).split(",") if item.strip()]
        for arg in args:
            rng = RANGE_REF_RE.match(arg)
            if rng:
                part = _range_total(ws, f"{rng.group(1)}{rng.group(2)}", f"{rng.group(3)}{rng.group(4)}", cached, depth)
                if part is None:
                    return None
                total += part
                continue
            ref = CELL_REF_RE.match(arg)
            if not ref:
                return None
            total += _cell_numeric(ws, int(ref.group(2)), column_index_from_string(ref.group(1).upper()), cached, depth)
        return total

    substituted = expression
    for ref_text in set(re.findall(r"\$?[A-Za-z]{1,3}\$?\d+", expression)):
        ref = CELL_REF_RE.match(ref_text)
        if not ref:
            continue
        number = _cell_numeric(ws, int(ref.group(2)), column_index_from_string(ref.group(1).upper()), cached, depth)
        substituted = substituted.replace(ref_text, repr(number))
    if not ARITH_OK_RE.match(substituted):
        return None
    try:
        return float(eval(substituted, {"__builtins__": {}}, {}))  # noqa: S307 - 已用白名单正则限制
    except Exception:
        return None
