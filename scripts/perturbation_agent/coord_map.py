"""
CoordMap: Unified coordinate transformation tracker.
All perturbations (rename, row shift, col insert) go through this.
"""

import re
from dataclasses import dataclass, field


@dataclass
class CoordMap:
    sheet_map: dict = field(default_factory=dict)      # old_sheet -> new_sheet
    row_shifts: dict = field(default_factory=dict)      # new_sheet -> +N rows
    col_inserts: dict = field(default_factory=dict)     # new_sheet -> [(position, count)]

    def add_sheet_rename(self, old: str, new: str):
        self.sheet_map[old] = new

    def add_row_shift(self, sheet: str, amount: int):
        """Sheet name should be the NEW name (after rename)."""
        self.row_shifts[sheet] = amount

    def add_col_insert(self, sheet: str, position: int, count: int = 1):
        self.col_inserts.setdefault(sheet, []).append((position, count))

    def resolve_sheet(self, sheet: str) -> str:
        return self.sheet_map.get(sheet, sheet)

    def map_cell(self, sheet: str, row: int, col: int):
        new_sheet = self.resolve_sheet(sheet)
        # If sheet is empty, try to infer from available shifts
        effective = new_sheet if new_sheet else self._infer_default_sheet() or ""
        new_row = row + self.row_shifts.get(effective, 0)
        new_col = col
        for pos, cnt in sorted(self.col_inserts.get(effective, [])):
            if col >= pos:
                new_col += cnt
        return new_sheet, new_row, new_col

    def _infer_default_sheet(self):
        """If there's exactly one sheet with shifts/inserts, return it."""
        all_sheets = set(self.row_shifts.keys()) | set(self.col_inserts.keys())
        if len(all_sheets) == 1:
            return next(iter(all_sheets))
        return ""

    def map_cell_ref(self, ref: str, default_sheet: str = "") -> str:
        """Map a cell reference like A1, $B$3, 'Sheet1'!C5."""
        sheet, cell = self._parse_cell_ref(ref, default_sheet)
        col_letter, row_num = self._split_cell(cell)
        col_idx = self._col_to_num(col_letter.replace("$", ""))
        row_idx = int(row_num.replace("$", ""))

        new_sheet, new_row, new_col = self.map_cell(sheet, row_idx, col_idx)

        # Preserve $ signs
        new_col_letter = self._num_to_col(new_col)
        if "$" in col_letter:
            new_col_letter = "$" + new_col_letter
        new_row_str = str(new_row)
        if "$" in row_num:
            new_row_str = "$" + new_row_str

        new_cell = new_col_letter + new_row_str
        if new_sheet and new_sheet != default_sheet:
            return f"'{new_sheet}'!{new_cell}"
        return new_cell

    def map_range(self, range_str: str) -> str:
        """Map a range like 'Sheet1'!A1:B5 or A1:B5,C1:D5.
        Handles sheet names with commas like 'b2b, sez, de'!A5:V10."""
        # Split by comma, but not inside single quotes
        parts = []
        current = ""
        in_quote = False
        for ch in range_str:
            if ch == "'":
                in_quote = not in_quote
                current += ch
            elif ch == "," and not in_quote:
                parts.append(current.strip())
                current = ""
            else:
                current += ch
        if current.strip():
            parts.append(current.strip())

        mapped = []
        for part in parts:
            if ":" in part:
                # Range: map both corners
                sheet, rest = self._extract_sheet(part)
                corners = rest.split(":")
                new_sheet = self.resolve_sheet(sheet) if sheet else ""
                col1, row1 = self._split_cell(corners[0])
                col2, row2 = self._split_cell(corners[1])

                if not row1 or not row2:
                    # Can't parse — return as-is
                    mapped.append(part)
                    continue

                if sheet:
                    _, new_row1, new_col1 = self.map_cell(sheet, int(row1.replace("$", "")), self._col_to_num(col1.replace("$", "")))
                    _, new_row2, new_col2 = self.map_cell(sheet, int(row2.replace("$", "")), self._col_to_num(col2.replace("$", "")))
                else:
                    _, new_row1, new_col1 = self.map_cell("", int(row1.replace("$", "")), self._col_to_num(col1.replace("$", "")))
                    _, new_row2, new_col2 = self.map_cell("", int(row2.replace("$", "")), self._col_to_num(col2.replace("$", "")))

                # Preserve $ signs
                nc1 = ("$" if "$" in col1 else "") + self._num_to_col(new_col1) + ("$" if "$" in row1 else "") + str(new_row1)
                nc2 = ("$" if "$" in col2 else "") + self._num_to_col(new_col2) + ("$" if "$" in row2 else "") + str(new_row2)

                if new_sheet:
                    mapped.append(f"'{new_sheet}'!{nc1}:{nc2}")
                else:
                    mapped.append(f"{nc1}:{nc2}")
            else:
                mapped.append(self.map_cell_ref(part))
        return ",".join(mapped)

    def map_answer_position(self, ap: str) -> str:
        """Map answer position, handling multiple ranges and quoted sheets.
        Also normalizes malformed formats like Analyse'!D5:K12."""
        # Normalize malformed formats:
        #   Analyse'!D5:K12 -> 'Analyse'!D5:K12  (missing leading quote)
        #   'Sheet1!'A1     -> 'Sheet1'!A1        (quote before ! instead of after)
        # Split by comma first to handle each part independently
        import re as _re
        parts = _re.split(r",(?='|[A-Z])", ap)
        normalized = []
        for part in parts:
            part = part.strip()
            # Case 1: word'! with no leading quote -> add leading quote
            # Match: start-of-string followed by word chars then '!
            part = _re.sub(r"^(\w+)'!", r"'\1'!", part)
            # Case 2: 'Sheet1!'A1 -> 'Sheet1'!A1
            part = _re.sub(r"'([^'!]+)!'", r"'\1'!", part)
            normalized.append(part)
        ap = ",".join(normalized)
        # Remove trailing stray quotes
        ap = ap.rstrip("'").rstrip('"')
        return self.map_range(ap)

    def to_dict(self) -> dict:
        return {
            "sheet_map": self.sheet_map,
            "row_shifts": self.row_shifts,
            "col_inserts": {k: v for k, v in self.col_inserts.items()},
        }

    # --- helpers ---

    @staticmethod
    def _parse_cell_ref(ref, default_sheet=""):
        sheet, cell = CoordMap._extract_sheet(ref)
        return sheet or default_sheet, cell

    @staticmethod
    def _extract_sheet(ref):
        # 'Sheet1'!A1 or Sheet1!A1
        m = re.match(r"'([^']+)'!(.*)", ref)
        if m:
            return m.group(1), m.group(2)
        m = re.match(r"(\w+)!(.*)", ref)
        if m:
            return m.group(1), m.group(2)
        return "", ref

    @staticmethod
    def _split_cell(cell):
        m = re.match(r"(\$?[A-Z]+)(\$?\d+)", cell)
        if m:
            return m.group(1), m.group(2)
        return cell, ""

    @staticmethod
    def _col_to_num(col):
        result = 0
        for c in col.upper():
            result = result * 26 + (ord(c) - ord("A") + 1)
        return result

    @staticmethod
    def _num_to_col(n):
        result = ""
        while n > 0:
            n, remainder = divmod(n - 1, 26)
            result = chr(65 + remainder) + result
        return result
