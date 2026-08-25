"""Small value conversions between Excel COM types and JSON-safe Python types."""


def to_jsonable(value):
    """Recursively convert a COM range value to a JSON-safe Python value.

    Handles the two COM quirks that break json.dumps() directly:
    tuple-of-tuples for multi-cell ranges, and pywintypes.datetime for
    date/time cells.
    """
    if isinstance(value, tuple):
        return [to_jsonable(v) for v in value]
    if hasattr(value, "isoformat"):
        return value.isoformat()
    return value


def hex_to_bgr(hex_color: str) -> int:
    """Convert a '#RRGGBB' or 'RRGGBB' hex string to the BGR integer that
    Excel's Font.Color / Interior.Color properties expect (legacy OLE
    COLORREF byte order, not RGB).
    """
    h = hex_color.lstrip("#")
    if len(h) != 6:
        raise ValueError(f"Invalid hex color: {hex_color!r}. Expected 'RRGGBB' or '#RRGGBB'.")
    r, g, b = int(h[0:2], 16), int(h[2:4], 16), int(h[4:6], 16)
    return (b << 16) | (g << 8) | r


def column_letter(col_num: int) -> str:
    """Convert a 1-based column number to Excel column letters (1 -> 'A', 27 -> 'AA').

    Used to build A1-style range strings by hand (e.g. "B2:D5") instead of
    calling Range.Resize(rows, cols) — Resize has all-optional parameters,
    which early-bound (gencache) pywin32 dispatch can resolve eagerly to the
    unresized range, silently turning a resize call into an unrelated
    Range.Item(row, col) lookup. Building the address as a plain string and
    passing it to Range(...) sidesteps the ambiguity entirely.
    """
    letters = ""
    while col_num > 0:
        col_num, remainder = divmod(col_num - 1, 26)
        letters = chr(65 + remainder) + letters
    return letters


def column_number(letters: str) -> int:
    """Convert Excel column letters to a 1-based column number ('A' -> 1, 'AA' -> 27).

    Inverse of column_letter(). Case-insensitive.
    """
    letters = letters.strip().upper()
    if not letters or not letters.isalpha():
        raise ValueError(f"Invalid column letters: {letters!r}")
    col_num = 0
    for ch in letters:
        col_num = col_num * 26 + (ord(ch) - ord("A") + 1)
    return col_num
