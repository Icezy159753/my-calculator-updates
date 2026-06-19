from __future__ import annotations

import argparse
import re
import struct
from os.path import commonprefix
from pathlib import Path
from typing import Callable

import pandas as pd
import pyreadstat


SPSS_STRING_RE = re.compile(r'"((?:[^"]|"")*)"')
VAR_LABEL_PAIR_RE = re.compile(r'([A-Za-z_@$#][\w@$#]*)\s+"((?:[^"]|"")*)"', re.S)
VALUE_PAIR_RE = re.compile(r'(?m)^\s*([-+]?\d+(?:\.\d+)?)\s+"((?:[^"]|"")*)"')
IGNORED_TEXT_FILES = {"requirements.txt", "readme.txt"}
LYCHEE_DATA_RE = re.compile(r"^SPSS_[af].*", re.I)
MR_VARIABLE_RE = re.compile(r"^(.+)\$(\d+)$")
MRSET_GROUP_RE = re.compile(r"/(MCGROUP|MDGROUP)\b(.*?)(?=/MCGROUP\b|/MDGROUP\b|\Z)", re.S | re.I)

# SPSS multiple-response set label limit (bytes); keep well within it.
MAX_MRSET_LABEL_BYTES = 255

# A line that begins with one of these keywords starts a brand-new SPSS command.
# Lychee sometimes emits labels with an unescaped " inside a "-quoted string
# (it should be doubled as ""), which flips naive quote tracking on and makes the
# splitter swallow every following command into one giant statement — silently
# dropping hundreds of VALUE LABELS. Forcing a statement break on these keywords
# bounds the damage to the single malformed label instead of the rest of the file.
COMMAND_START_RE = re.compile(
    r"(?i)^\s*(?:"
    r"GET\s+DATA|GET\s+FILE|SAVE|EXECUTE|DATASET\b|SET\b|"
    r"ADD\s+VALUE\s+LABELS|VALUE\s+LABELS|VARIABLE\s+LABELS|"
    r"VARIABLE\s+LEVEL|VARIABLE\s+WIDTH|VARIABLE\s+ALIGNMENT|VARIABLE\s+ROLE|"
    r"FORMATS|MISSING\s+VALUES|MRSETS|RENAME\s+VARIABLES|FORMAT\b"
    r")\b"
)


def clean_sps_string(value: str) -> str:
    return value.replace('""', '"').replace("\r", " ").replace("\n", " ").strip()


# SPSS hard limits (in BYTES, not characters): variable labels 255, value labels 120.
# Thai characters are 3 bytes each in UTF-8, so long Thai labels exceed these
# limits easily. If the writer truncates mid-character, the label ends with an
# invalid UTF-8 sequence and IBM SPSS renders the whole label as "□".
MAX_VARIABLE_LABEL_BYTES = 255
MAX_VALUE_LABEL_BYTES = 120


def truncate_utf8(value: str, max_bytes: int) -> str:
    """Truncate ``value`` to at most ``max_bytes`` UTF-8 bytes on a character boundary."""
    raw = value.encode("utf-8")
    if len(raw) <= max_bytes:
        return value
    return raw[:max_bytes].decode("utf-8", errors="ignore")


def smart_truncate_label(value: str, max_bytes: int = MAX_VARIABLE_LABEL_BYTES) -> str:
    """Truncate an over-long label in the middle so the distinguishing tail survives.

    Lychee labels look like ``C3 <question>-<choice>``: the part that tells the
    items of a multiple-response set apart is the choice name at the END, so a
    plain end-truncation would collapse all items onto an identical label.
    """
    raw = value.encode("utf-8")
    if len(raw) <= max_bytes:
        return value

    ellipsis = "..."  # 3 bytes
    sep_index = value.rfind("-")
    if sep_index != -1:
        tail = value[sep_index:]
        tail_bytes = len(tail.encode("utf-8"))
        if tail_bytes <= max_bytes // 2:
            head = truncate_utf8(value, max_bytes - tail_bytes - len(ellipsis))
            return head + ellipsis + tail

    # No usable "-choice" suffix: keep ~2/3 of the head and ~1/3 of the tail.
    head_budget = (max_bytes - len(ellipsis)) * 2 // 3
    tail_budget = max_bytes - len(ellipsis) - head_budget
    head = truncate_utf8(value, head_budget)
    tail = raw[-tail_budget:].decode("utf-8", errors="ignore")
    return head + ellipsis + tail


def truncate_labels_preserving_distinction(
    labels: dict[str, str], max_bytes: int = MAX_VARIABLE_LABEL_BYTES
) -> dict[str, str]:
    """Truncate labels, then repair groups where distinct sources collapsed together.

    Some Lychee labels (e.g. "other-specify" loop variables) carry the
    distinguishing brand name in the MIDDLE and end with repeated question
    text, so any head/tail truncation can collapse them. For each collision
    group, strip the prefix and suffix common to the colliding sources and
    splice the differing middle segment back into the truncated label.
    """
    truncated = {name: smart_truncate_label(label, max_bytes) for name, label in labels.items()}
    ellipsis = "..."

    groups: dict[str, list[str]] = {}
    for name, short in truncated.items():
        groups.setdefault(short, []).append(name)

    for names in groups.values():
        sources = sorted({labels[name] for name in names})
        if len(sources) < 2:
            continue
        prefix_len = len(commonprefix(sources))
        suffix_len = len(commonprefix([source[::-1] for source in sources]))
        for name in names:
            source = labels[name]
            mid_end = max(prefix_len, len(source) - suffix_len)
            mid = source[prefix_len:mid_end].strip()
            if not mid:
                continue
            mid = truncate_utf8(mid, (max_bytes - 2 * len(ellipsis)) // 2)
            head_budget = max_bytes - 2 * len(ellipsis) - len(mid.encode("utf-8"))
            head = truncate_utf8(source, head_budget)
            truncated[name] = head + ellipsis + mid + ellipsis

    return truncated


def _line_toggles_quote(line: str, in_quote: bool) -> bool:
    """Return the quote state at the end of ``line`` given the state before it.

    A doubled ``""`` is an escaped quote inside a string, not a delimiter.
    """
    index = 0
    length = len(line)
    while index < length:
        if line[index] == '"':
            if in_quote and index + 1 < length and line[index + 1] == '"':
                index += 2
                continue
            in_quote = not in_quote
        index += 1
    return in_quote


def strip_sps_comments(text: str) -> str:
    lines: list[str] = []
    in_quote = False
    for line in text.splitlines():
        stripped = line.strip()
        # Never treat a line inside a quoted label as blank or as a comment,
        # otherwise multi-line labels containing "*" or "." lines get dropped.
        if not in_quote:
            if not stripped:
                continue
            if stripped.startswith("*") and stripped.endswith("."):
                continue
        in_quote = _line_toggles_quote(line, in_quote)
        lines.append(line.rstrip())
    return "\n".join(lines)


def split_sps_statements(text: str) -> list[str]:
    statements: list[str] = []
    current: list[str] = []
    in_quote = False

    for line in text.splitlines():
        # A new command keyword at the start of a line always begins a new
        # statement. Flushing here recovers from a previous statement whose quote
        # state got stuck "open" because of an unescaped " inside a label, so one
        # malformed label can no longer swallow the commands that follow it.
        if current and COMMAND_START_RE.match(line):
            statements.append("\n".join(current).strip())
            current = []
            in_quote = False

        current.append(line)

        # Track quote state across physical lines so that a period inside a
        # quoted label (e.g. "...ทั่วไปนี้." or a wrapped multi-line label) is
        # never mistaken for a command terminator.
        in_quote = _line_toggles_quote(line, in_quote)
        if in_quote:
            # The statement (label text) continues on the next physical line.
            continue

        stripped = line.strip()
        if stripped == "." or (stripped.endswith(".") and not re.search(r"\b[AF]\d+\.\d+$", stripped, re.I)):
            statements.append("\n".join(current).strip())
            current = []

    tail = "\n".join(current).strip()
    if tail:
        statements.append(tail)
    return statements


def parse_get_data_variables(statements: list[str]) -> tuple[list[str], dict[str, str], dict[str, str]]:
    for statement in statements:
        if not statement.upper().startswith("GET DATA"):
            continue

        match = re.search(r"/VARIABLES\s*=\s*(.*)\.\s*$", statement, re.I | re.S)
        if not match:
            continue

        variables: list[str] = []
        variable_formats: dict[str, str] = {}
        string_widths: dict[str, str] = {}

        for raw_line in match.group(1).splitlines():
            line = raw_line.strip()
            if not line or line.startswith("/"):
                continue

            parts = line.split()
            if len(parts) < 2:
                continue

            name, fmt = parts[0], parts[1].upper()
            variables.append(name)
            variable_formats[name] = fmt
            if fmt.startswith("A"):
                width = re.sub(r"\D", "", fmt) or "255"
                string_widths[name] = f"A{width}"

        return variables, variable_formats, string_widths

    raise ValueError("Could not find a GET DATA /VARIABLES block in the syntax files.")


def parse_variable_labels(statements: list[str]) -> dict[str, str]:
    labels: dict[str, str] = {}
    for statement in statements:
        if not statement.upper().lstrip().startswith("VARIABLE LABELS"):
            continue
        body = re.sub(r"^\s*VARIABLE\s+LABELS\s+", "", statement, flags=re.I).strip().rstrip(".").strip()
        pairs = VAR_LABEL_PAIR_RE.findall(body)

        # A clean command has exactly two double-quotes per variable. Lychee
        # sometimes leaves an unescaped " inside a label (it should be doubled as
        # ""), so the strict pattern stops mid-text and the label is truncated.
        # When the quote count doesn't match, recover the single-variable label
        # as everything between the first and last quote so the inner quote (and
        # the rest of the sentence) survive as literal text.
        if body.count('"') != 2 * len(pairs):
            single = re.match(r'([A-Za-z_@$#][\w@$#]*)\s+"(.*)"\s*$', body, re.S)
            if single:
                labels[single.group(1)] = clean_sps_string(single.group(2))
                continue

        for name, label in pairs:
            labels[name] = clean_sps_string(label)
    return labels


def parse_value_labels(statements: list[str]) -> dict[str, dict[float | int, str]]:
    labels: dict[str, dict[float | int, str]] = {}

    for statement in statements:
        if not statement.upper().startswith("VALUE LABELS"):
            continue

        body = re.sub(r"^VALUE\s+LABELS\s+", "", statement, flags=re.I).rstrip(".")
        first_value = re.search(r"(?m)^\s*[-+]?\d+(?:\.\d+)?\s+\"", body)
        if not first_value:
            continue

        variable_block = body[: first_value.start()]
        variable_names = [
            token
            for token in re.split(r"\s+", variable_block.strip())
            if token and not token.startswith("/")
        ]

        value_map: dict[float | int, str] = {}
        for raw_value, label in VALUE_PAIR_RE.findall(body[first_value.start() :]):
            cleaned_label = clean_sps_string(label)
            if cleaned_label == "":
                continue
            numeric_value = float(raw_value)
            value: float | int = int(numeric_value) if numeric_value.is_integer() else numeric_value
            value_map[value] = cleaned_label

        for variable_name in variable_names:
            labels.setdefault(variable_name, {}).update(value_map)

    return labels


def parse_variable_measure(statements: list[str]) -> dict[str, str]:
    measure_map = {
        "NOMINAL": "nominal",
        "ORDINAL": "ordinal",
        "SCALE": "scale",
    }
    measures: dict[str, str] = {}

    for statement in statements:
        if not statement.upper().startswith("VARIABLE LEVEL"):
            continue
        body = re.sub(r"^VARIABLE\s+LEVEL\s+", "", statement, flags=re.I).rstrip(".")
        for chunk in body.split("/"):
            match = re.search(r"\((NOMINAL|ORDINAL|SCALE)\)", chunk, re.I)
            if not match:
                continue
            measure = measure_map[match.group(1).upper()]
            vars_part = chunk[: match.start()]
            for variable_name in re.split(r"\s+", vars_part.strip()):
                if variable_name:
                    measures[variable_name] = measure

    return measures


def mrset_sort_key(variable_name: str) -> tuple[str, int]:
    match = MR_VARIABLE_RE.match(variable_name)
    if not match:
        return variable_name.lower(), 0
    return match.group(1).lower(), int(match.group(2))


def parse_mrsets(statements: list[str], valid_names: list[str]) -> list[dict]:
    """Parse explicit ``MRSETS /MCGROUP`` / ``/MDGROUP`` definitions from the syntax.

    Lychee exports the real multiple-response sets — with their proper set names
    and descriptive LABELs — in an ``MRSETS`` command. Reading those is far better
    than guessing sets from the ``$n`` naming, because we keep the labels SPSS
    shows in "Define Multiple Response Sets" and the exact intended grouping.
    Variables not present in the data file are dropped, and sets left with fewer
    than two variables are skipped.
    """
    allowed = set(valid_names)
    used_names: set[str] = set()
    sets: list[dict] = []
    for statement in statements:
        if not statement.upper().lstrip().startswith("MRSETS"):
            continue
        for kind, body in MRSET_GROUP_RE.findall(statement):
            name_match = re.search(r"(?i)\bNAME\s*=\s*(\$?\w+)", body)
            vars_match = re.search(r"(?i)\bVARIABLES\s*=\s*(.*)", body, re.S)
            if not name_match or not vars_match:
                continue
            label_match = re.search(r'(?i)\bLABEL\s*=\s*"((?:[^"]|"")*)"', body)
            value_match = re.search(r"(?i)\bVALUE\s*=\s*(\S+)", body)

            variables = [token for token in vars_match.group(1).split() if token in allowed]
            if len(variables) < 2:
                continue

            # Lychee reuses the SAME NAME (e.g. $NQ22) for every grid row, so the
            # 29 distinct sets collide and SPSS keeps only one. Derive the set name
            # from the variables' common base (e.g. NQ22#1 -> $NQ22_1), which is
            # unique per set, and only fall back to the declared NAME otherwise.
            source_name = name_match.group(1).lstrip("$")
            base = _common_mr_base(variables)
            candidate = "$" + re.sub(r"\W+", "_", base or source_name).strip("_")
            unique = candidate
            suffix = 2
            while unique.lower() in used_names:
                unique = f"{candidate}_{suffix}"
                suffix += 1
            used_names.add(unique.lower())

            label = truncate_utf8(clean_sps_string(label_match.group(1)) if label_match else "",
                                  MAX_MRSET_LABEL_BYTES)
            sets.append(
                {
                    "name": unique,
                    "is_dichotomy": kind.upper() == "MDGROUP",
                    "counted_value": value_match.group(1) if value_match else "1",
                    "label": label,
                    "variables": variables,
                }
            )
    return sets


def _common_mr_base(variables: list[str]) -> str | None:
    """Return the shared ``base`` of ``base$n`` variable names, or None if mixed."""
    bases = {(MR_VARIABLE_RE.match(name).group(1) if MR_VARIABLE_RE.match(name) else name)
             for name in variables}
    return bases.pop() if len(bases) == 1 else None


def auto_mrsets(column_names: list[str]) -> list[dict]:
    """Fallback when the syntax has no MRSETS command: infer category sets from
    the ``base$n`` naming. Labels are unknown, so they are left empty."""
    return [
        {"name": f"${base_name}", "is_dichotomy": False, "counted_value": "1",
         "label": "", "variables": variables}
        for base_name, variables in build_mrsets(column_names).items()
    ]


def build_mrsets(column_names: list[str]) -> dict[str, list[str]]:
    grouped: dict[str, list[str]] = {}
    for column_name in column_names:
        match = MR_VARIABLE_RE.match(column_name)
        if not match:
            continue
        grouped.setdefault(match.group(1), []).append(column_name)

    return {
        base_name: sorted(variables, key=mrset_sort_key)
        for base_name, variables in sorted(grouped.items())
        if len(variables) > 1
    }


def encode_mr_string(value: str) -> bytes:
    raw = value.encode("utf-8", errors="ignore")
    return str(len(raw)).encode("ascii") + b" " + raw


def find_dict_terminator(data: bytes) -> int:
    """Return the byte offset of the genuine type-999 dictionary terminator.

    The terminator is an ``int32 999`` followed by an ``int32 0`` filler that
    separates the dictionary from the case data. A naive ``data.find()`` for that
    8-byte pattern is unsafe: the same bytes occur by chance inside value-label
    values, print/write formats and string data, so on large dictionaries the
    first match lands far before the real terminator and any record spliced there
    corrupts the file ("Invalid SPSS Statistics data file"). We instead walk the
    dictionary record-by-record so the offset is always the real boundary.
    """
    if data[:4] not in (b"$FL2", b"$FL3"):
        raise ValueError("Not a recognizable SPSS .sav file (missing $FL2/$FL3 header).")

    pos = 176  # fixed-size file header
    length = len(data)
    while pos < length:
        rec_type = struct.unpack_from("<i", data, pos)[0]
        if rec_type == 999:
            return pos
        if rec_type == 2:  # variable record
            has_label = struct.unpack_from("<i", data, pos + 8)[0]
            n_missing = struct.unpack_from("<i", data, pos + 12)[0]
            pos += 32
            if has_label == 1:
                label_len = struct.unpack_from("<i", data, pos)[0]
                pos += 4 + ((label_len + 3) // 4) * 4
            pos += abs(n_missing) * 8
        elif rec_type == 3:  # value labels, always followed by a type-4 index record
            count = struct.unpack_from("<i", data, pos + 4)[0]
            pos += 8
            for _ in range(count):
                label_len = data[pos + 8]
                pos += 8 + ((1 + label_len + 7) // 8) * 8
            pos += 8 + struct.unpack_from("<i", data, pos + 4)[0] * 4
        elif rec_type == 6:  # document record
            n_lines = struct.unpack_from("<i", data, pos + 4)[0]
            pos += 8 + n_lines * 80
        elif rec_type == 7:  # extension record (subtype/size/count header)
            size = struct.unpack_from("<i", data, pos + 8)[0]
            count = struct.unpack_from("<i", data, pos + 12)[0]
            pos += 16 + size * count
        else:
            raise ValueError(f"Unexpected SPSS record type {rec_type} at offset {pos}.")

    raise ValueError("Could not find the SPSS dictionary terminator.")


def patch_mrsets_record(sav_path: Path, mr_sets: list[dict]) -> int:
    if not mr_sets:
        return 0

    lines: list[bytes] = []
    for mr_set in mr_sets:
        set_name = mr_set["name"].lower().encode("utf-8", errors="ignore")
        # Variable names are kept verbatim so they match the dictionary exactly.
        variable_list = " ".join(mr_set["variables"]).encode("utf-8", errors="ignore")
        if mr_set["is_dichotomy"]:
            # $name=D<counted-value> <label> vars  — IBM writes the counted-value
            # length immediately after D with no separating space (e.g. =D1 1 7 ...).
            header = b"=D" + encode_mr_string(mr_set["counted_value"]) + b" " + encode_mr_string(mr_set["label"])
        else:
            # $name=C <label> vars
            header = b"=C " + encode_mr_string(mr_set["label"])
        lines.append(set_name + header + b" " + variable_list)

    payload = b"\n".join(lines) + b"\n"
    record = (
        (7).to_bytes(4, "little")
        + (7).to_bytes(4, "little")
        + (1).to_bytes(4, "little")
        + len(payload).to_bytes(4, "little")
        + payload
    )

    data = sav_path.read_bytes()
    insert_at = find_dict_terminator(data)
    sav_path.write_bytes(data[:insert_at] + record + data[insert_at:])
    return len(mr_sets)


def patch_variable_attributes_record(sav_path: Path, full_labels: dict[str, str]) -> int:
    """Store full (untruncated) labels as a custom variable attribute.

    Writes a variable-attributes extension record (type 7, subtype 18) so the
    complete question text survives the 255-byte variable label limit. The
    attribute shows up in SPSS Variable View as a "FullLabel" column.
    """
    if not full_labels:
        return 0

    portions: list[str] = []
    for name, text in full_labels.items():
        value = text.replace("'", "''")
        portions.append(f"{name}:FullLabel('{value}'\n)")

    payload = "/".join(portions).encode("utf-8")
    record = (
        (7).to_bytes(4, "little")
        + (18).to_bytes(4, "little")
        + (1).to_bytes(4, "little")
        + len(payload).to_bytes(4, "little")
        + payload
    )

    data = sav_path.read_bytes()
    insert_at = find_dict_terminator(data)
    sav_path.write_bytes(data[:insert_at] + record + data[insert_at:])
    return len(full_labels)


def read_sps_files(paths: list[Path]) -> tuple[list[str], dict[str, str], dict[str, str]]:
    all_statements: list[str] = []
    get_data_statements: list[str] = []

    for path in paths:
        text = strip_sps_comments(path.read_text(encoding="utf-8-sig", errors="replace"))
        statements = split_sps_statements(text)
        all_statements.extend(statements)
        if any(statement.upper().startswith("GET DATA") for statement in statements):
            get_data_statements.extend(statements)

    variables, formats, string_widths = parse_get_data_variables(get_data_statements or all_statements)
    return variables, formats, string_widths, all_statements


def convert(
    input_txt: Path,
    syntax_files: list[Path],
    output_sav: Path,
    logger: Callable[[str], None] | None = None,
    row_compress: bool = True,
    compress: bool = False,
    add_mrsets: bool = True,
) -> dict[str, int | str]:
    def log(message: str) -> None:
        if logger:
            logger(message)

    log("Reading SPSS syntax files...")
    variables, formats, string_widths, statements = read_sps_files(syntax_files)

    log("Reading tab-delimited text data...")
    df = pd.read_csv(
        input_txt,
        sep="\t",
        header=0,
        names=variables,
        usecols=range(len(variables)),
        dtype=str,
        keep_default_na=False,
        encoding="utf-8-sig",
        engine="python",
    )

    for column_name, fmt in formats.items():
        if column_name not in df.columns or fmt.startswith("A"):
            continue
        df[column_name] = pd.to_numeric(df[column_name].replace("", pd.NA), errors="coerce")

    log("Parsing labels and variable metadata...")
    column_labels = parse_variable_labels(statements)
    value_labels = parse_value_labels(statements)
    variable_measure = parse_variable_measure(statements)

    # Clamp labels to the SPSS byte limits so truncation never lands mid-character.
    # Over-long labels are middle-truncated (keeping the distinguishing choice
    # suffix) and the full text is preserved in a "FullLabel" variable attribute.
    full_labels = {
        name: label
        for name, label in column_labels.items()
        if len(label.encode("utf-8")) > MAX_VARIABLE_LABEL_BYTES
    }
    if full_labels:
        log(
            f"Truncating {len(full_labels)} variable labels longer than "
            f"{MAX_VARIABLE_LABEL_BYTES} bytes (full text kept in FullLabel attribute)..."
        )
    column_labels = truncate_labels_preserving_distinction(column_labels)
    value_labels = {
        name: {
            value: truncate_utf8(label, MAX_VALUE_LABEL_BYTES) for value, label in value_map.items()
        }
        for name, value_map in value_labels.items()
    }

    log("Writing SPSS .sav file...")
    pyreadstat.write_sav(
        df,
        str(output_sav),
        column_labels=column_labels,
        variable_value_labels=value_labels,
        variable_measure=variable_measure or None,
        variable_format=string_widths or None,
        compress=compress,
        row_compress=row_compress,
    )

    mrset_count = 0
    if add_mrsets:
        # Prefer the real MRSETS command from the syntax (keeps set labels); only
        # fall back to inferring sets from the $n naming when none are defined.
        mr_sets = parse_mrsets(statements, list(df.columns)) or auto_mrsets(list(df.columns))
        if mr_sets:
            labeled = sum(1 for s in mr_sets if s["label"])
            log(f"Adding {len(mr_sets)} MRSETS ({labeled} with labels)...")
            mrset_count = patch_mrsets_record(output_sav, mr_sets)

    full_label_count = 0
    if full_labels:
        log("Storing full labels as variable attributes...")
        full_label_count = patch_variable_attributes_record(output_sav, full_labels)

    return {
        "output": str(output_sav),
        "rows": len(df),
        "columns": len(df.columns),
        "variable_labels": len(column_labels),
        "value_label_sets": len(value_labels),
        "mrsets": mrset_count,
        "full_label_attributes": full_label_count,
        "file_size_kb": output_sav.stat().st_size // 1024,
    }


def resolve_input_txt(base: Path, txt_arg: str | None) -> Path:
    if txt_arg:
        return (base / txt_arg).resolve()

    all_candidates = [path for path in base.glob("*.txt") if path.name.lower() not in IGNORED_TEXT_FILES]
    lychee_candidates = [path for path in all_candidates if LYCHEE_DATA_RE.match(path.stem)]
    candidates = sorted(
        lychee_candidates or all_candidates,
        key=lambda path: path.stat().st_size,
        reverse=True,
    )
    if not candidates:
        raise FileNotFoundError("No text data file found. Use --txt to specify one.")
    return candidates[0].resolve()


def resolve_sps_files(base: Path, sps_args: list[str] | None) -> list[Path]:
    if sps_args:
        return [(base / path).resolve() for path in sps_args]

    all_candidates = sorted(base.glob("*.sps"))
    candidates = [path for path in all_candidates if LYCHEE_DATA_RE.match(path.stem)] or all_candidates
    if not candidates:
        raise FileNotFoundError("No .sps syntax files found. Use --sps to specify one or more files.")
    return [path.resolve() for path in candidates]


def main() -> None:
    parser = argparse.ArgumentParser(
        description="Run the supported SPSS GET DATA syntax subset and write an SPSS .sav file without IBM SPSS."
    )
    parser.add_argument(
        "--txt",
        default=None,
        help="Input tab-delimited TXT file. If omitted, the largest .txt data file in the folder is used.",
    )
    parser.add_argument(
        "--sps",
        nargs="+",
        default=None,
        help="SPSS syntax files to read metadata from. If omitted, all .sps files in the folder are used.",
    )
    parser.add_argument("--out", default="SPSS.sav", help="Output SAV file.")
    args = parser.parse_args()

    base = Path.cwd()
    input_txt = resolve_input_txt(base, args.txt)
    syntax_files = resolve_sps_files(base, args.sps)
    output_sav = (base / args.out).resolve()

    print(f"Text data: {input_txt}")
    print(f"SPS files: {len(syntax_files)}")
    for syntax_file in syntax_files:
        print(f"  - {syntax_file.name}")
    print("Compression: Compatible")

    result = convert(
        input_txt,
        syntax_files,
        output_sav,
        row_compress=True,
        compress=False,
    )
    print(f"Wrote {result['output']}")
    print(f"Rows: {result['rows']:,}")
    print(f"Columns: {result['columns']:,}")
    print(f"Variable labels: {result['variable_labels']:,}")
    print(f"Value label sets applied: {result['value_label_sets']:,}")
    print(f"MRSETS added: {result['mrsets']:,}")
    print(f"Full labels kept as attributes: {result['full_label_attributes']:,}")
    print(f"File size: {result['file_size_kb']:,} KB")


if __name__ == "__main__":
    main()
