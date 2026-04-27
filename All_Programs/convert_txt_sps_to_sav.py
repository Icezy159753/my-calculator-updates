from __future__ import annotations

import argparse
import re
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


def clean_sps_string(value: str) -> str:
    return value.replace('""', '"').replace("\r", " ").replace("\n", " ").strip()


def strip_sps_comments(text: str) -> str:
    lines: list[str] = []
    for line in text.splitlines():
        stripped = line.strip()
        if not stripped:
            continue
        if stripped.startswith("*") and stripped.endswith("."):
            continue
        lines.append(line.rstrip())
    return "\n".join(lines)


def split_sps_statements(text: str) -> list[str]:
    statements: list[str] = []
    current: list[str] = []

    for line in text.splitlines():
        current.append(line)
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
        if not statement.upper().startswith("VARIABLE LABELS"):
            continue
        body = re.sub(r"^VARIABLE\s+LABELS\s+", "", statement, flags=re.I).rstrip(".")
        for name, label in VAR_LABEL_PAIR_RE.findall(body):
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


def patch_mrsets_record(sav_path: Path, mrsets: dict[str, list[str]], labels: dict[str, str]) -> int:
    if not mrsets:
        return 0

    lines: list[bytes] = []
    for base_name, variables in mrsets.items():
        set_name = f"${base_name}".lower()
        variable_list = " ".join(variable.lower() for variable in variables)
        lines.append(
            set_name.encode("utf-8", errors="ignore")
            + b"=C "
            + b"0  "
            + variable_list.encode("utf-8", errors="ignore")
        )

    payload = b"\n".join(lines) + b"\n"
    record = (
        (7).to_bytes(4, "little")
        + (7).to_bytes(4, "little")
        + (1).to_bytes(4, "little")
        + len(payload).to_bytes(4, "little")
        + payload
    )

    data = sav_path.read_bytes()
    marker = (999).to_bytes(4, "little") + (0).to_bytes(4, "little")
    insert_at = data.find(marker)
    if insert_at < 0:
        raise ValueError("Could not find the SPSS dictionary terminator while adding MRSETS.")

    sav_path.write_bytes(data[:insert_at] + record + data[insert_at:])
    return len(mrsets)


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
        log("Adding MRSETS metadata...")
        mrset_count = patch_mrsets_record(output_sav, build_mrsets(list(df.columns)), column_labels)

    return {
        "output": str(output_sav),
        "rows": len(df),
        "columns": len(df.columns),
        "variable_labels": len(column_labels),
        "value_label_sets": len(value_labels),
        "mrsets": mrset_count,
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
    print(f"File size: {result['file_size_kb']:,} KB")


if __name__ == "__main__":
    main()
