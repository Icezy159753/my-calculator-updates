from __future__ import annotations

from dataclasses import dataclass
import re
from typing import Mapping

import numpy as np
import pandas as pd
from scipy.stats import ttest_ind

RESULT_COLUMNS = [
    "Filter Condition",
    "Filter Query",
    "Attribute",
    "Base N",
    "Base (%) 1+2",
    "Base (%) JAR",
    "Base (%) 4+5",
    "Mean 1+2",
    "Mean JAR",
    "Mean 4+5",
    "Weighted Penalty 1+2",
    "Impact Level 1+2",
    "Weighted Penalty 4+5",
    "Impact Level 4+5",
    "p-value 1+2",
    "Sig 1+2",
    "p-value 4+5",
    "Sig 4+5",
]
NUMERIC_RESULT_COLUMNS = [
    "Base N",
    "Base (%) 1+2",
    "Base (%) JAR",
    "Base (%) 4+5",
    "Mean 1+2",
    "Mean JAR",
    "Mean 4+5",
    "Weighted Penalty 1+2",
    "Weighted Penalty 4+5",
]
DEFAULT_JAR_SCALE = 5
FIVE_POINT_JAR_SCORES = {1, 2, 3, 4, 5}
SEVEN_POINT_JAR_SCORES = {1, 2, 3, 4, 5, 6, 7}
NO_FILTER_LABEL = "รวมทั้งหมด"
NO_FILTER_QUERY_LABEL = "(No Filter)"
HIGH_INFLUENCE_THRESHOLD = 0.50
MEDIUM_INFLUENCE_THRESHOLD = 0.20
HIGH_INFLUENCE_LABEL = "High influence - Requires adjustment"
MEDIUM_INFLUENCE_LABEL = "Influences overall liking - May need adjustment"


class FilterQueryError(ValueError):
    """Raised when a filter query cannot be parsed by pandas."""


@dataclass(frozen=True)
class FilterSpec:
    label: str
    query: str = ""

    @property
    def display_query(self) -> str:
        return self.query or NO_FILTER_QUERY_LABEL


def ttest_significance(group_scores: pd.Series, jar_scores: pd.Series) -> tuple[float, str]:
    """Run Welch's t-test between a non-JAR group and the JAR group."""
    group = group_scores.dropna()
    jar = jar_scores.dropna()
    if len(group) < 2 or len(jar) < 2:
        return np.nan, ""
    _, p_value = ttest_ind(group, jar, equal_var=False)
    if p_value < 0.01:
        return round(p_value, 4), "**"
    if p_value < 0.05:
        return round(p_value, 4), "*"
    return round(p_value, 4), "ns"


def normalize_filter_query(raw_query: str, columns: pd.Index) -> str:
    """Normalize a user-entered pandas query by fixing '=' and column casing."""
    safe_query = re.sub(r"(?<![<>!=])=(?!=)", "==", raw_query)
    col_map = {str(column).lower(): column for column in columns}

    def replace_case(match: re.Match[str]) -> str:
        word = match.group(0)
        return str(col_map.get(word.lower(), word))

    return re.sub(r"\b[A-Za-z_]\w*\b", replace_case, safe_query)


def apply_filter(df: pd.DataFrame, raw_query: str) -> pd.DataFrame:
    """Apply a user-entered filter query to a DataFrame."""
    if not raw_query:
        return df.copy()

    # ยุบ newline/tab/ช่องว่างซ้อนให้เหลือ space เดียว กัน pandas มองเป็น multi-line expression
    raw_query = " ".join(raw_query.split())
    if not raw_query:
        return df.copy()

    safe_query = normalize_filter_query(raw_query, df.columns)
    try:
        return df.query(safe_query)
    except Exception as exc:  # pandas raises different parser/runtime errors
        raise FilterQueryError(str(exc)) from exc


def analyze_penalty(
    df: pd.DataFrame,
    dep_var: str,
    jar_vars: list[str],
    filters: list[FilterSpec],
    label_map: Mapping[str, str] | None = None,
    jar_scale_map: Mapping[str, int] | None = None,
) -> pd.DataFrame:
    """Compute penalty analysis results for all selected filters and attributes."""
    if label_map is None:
        label_map = {}
    if jar_scale_map is None:
        jar_scale_map = {}

    results: list[dict[str, object]] = []

    for filter_spec in filters:
        filtered_df = apply_filter(df, filter_spec.query)
        if filtered_df.empty:
            results.extend(
                _build_empty_rows(
                    filter_spec=filter_spec,
                    jar_vars=jar_vars,
                    label_map=label_map,
                )
            )
            continue

        dep_clean = filtered_df.copy()
        dep_clean[dep_var] = pd.to_numeric(dep_clean[dep_var], errors="coerce")
        dep_clean = dep_clean.dropna(subset=[dep_var])

        for jar_var in jar_vars:
            results.append(
                _analyze_single_attribute(
                    df=dep_clean,
                    dep_var=dep_var,
                    jar_var=jar_var,
                    filter_spec=filter_spec,
                    label_map=label_map,
                    jar_scale=jar_scale_map.get(jar_var),
                )
            )

    result_df = pd.DataFrame(results, columns=RESULT_COLUMNS)
    for column in NUMERIC_RESULT_COLUMNS:
        result_df[column] = result_df[column].fillna(0)
    return result_df


def _analyze_single_attribute(
    df: pd.DataFrame,
    dep_var: str,
    jar_var: str,
    filter_spec: FilterSpec,
    label_map: Mapping[str, str],
    jar_scale: int | None,
) -> dict[str, object]:
    attribute_label = label_map.get(jar_var) or jar_var
    temp = df.dropna(subset=[jar_var]).copy()
    temp[jar_var] = pd.to_numeric(temp[jar_var], errors="coerce")
    valid_scores, jar_group_score, low_group_scores, high_group_scores = _resolve_jar_groups(
        temp[jar_var],
        jar_scale,
    )
    temp = temp[temp[jar_var].isin(valid_scores)]

    if temp.empty:
        return _build_empty_row(filter_spec, attribute_label)

    jar_data = temp[temp[jar_var] == jar_group_score]
    group_12 = temp[temp[jar_var].isin(low_group_scores)]
    group_45 = temp[temp[jar_var].isin(high_group_scores)]

    total = len(temp)
    base_jar = _to_percent(len(jar_data), total)
    base_12 = _to_percent(len(group_12), total)
    base_45 = _to_percent(len(group_45), total)

    mean_jar = _rounded_mean(jar_data[dep_var])
    mean_12 = _rounded_mean(group_12[dep_var])
    mean_45 = _rounded_mean(group_45[dep_var])

    penalty_12 = _weighted_penalty(mean_12, mean_jar, base_12)
    penalty_45 = _weighted_penalty(mean_45, mean_jar, base_45)

    p_12, sig_12 = ttest_significance(group_12[dep_var], jar_data[dep_var])
    p_45, sig_45 = ttest_significance(group_45[dep_var], jar_data[dep_var])

    return {
        "Filter Condition": filter_spec.label,
        "Filter Query": filter_spec.display_query,
        "Attribute": attribute_label,
        "Base N": total,
        "Base (%) 1+2": base_12,
        "Base (%) JAR": base_jar,
        "Base (%) 4+5": base_45,
        "Mean 1+2": mean_12,
        "Mean JAR": mean_jar,
        "Mean 4+5": mean_45,
        "Weighted Penalty 1+2": penalty_12,
        "Impact Level 1+2": classify_penalty_impact(penalty_12),
        "Weighted Penalty 4+5": penalty_45,
        "Impact Level 4+5": classify_penalty_impact(penalty_45),
        "p-value 1+2": p_12,
        "Sig 1+2": sig_12,
        "p-value 4+5": p_45,
        "Sig 4+5": sig_45,
    }


def _build_empty_rows(
    filter_spec: FilterSpec,
    jar_vars: list[str],
    label_map: Mapping[str, str],
) -> list[dict[str, object]]:
    return [
        _build_empty_row(filter_spec, label_map.get(jar_var) or jar_var)
        for jar_var in jar_vars
    ]


def _build_empty_row(filter_spec: FilterSpec, attribute_label: str) -> dict[str, object]:
    return {
        "Filter Condition": filter_spec.label,
        "Filter Query": filter_spec.display_query,
        "Attribute": attribute_label,
        "Base N": 0,
        "Base (%) 1+2": 0,
        "Base (%) JAR": 0,
        "Base (%) 4+5": 0,
        "Mean 1+2": 0,
        "Mean JAR": 0,
        "Mean 4+5": 0,
        "Weighted Penalty 1+2": 0,
        "Impact Level 1+2": "",
        "Weighted Penalty 4+5": 0,
        "Impact Level 4+5": "",
        "p-value 1+2": np.nan,
        "Sig 1+2": "",
        "p-value 4+5": np.nan,
        "Sig 4+5": "",
    }


def _to_percent(numerator: int, denominator: int) -> int:
    if denominator <= 0:
        return 0
    return int(round((numerator / denominator) * 100 + 1e-9))


def _rounded_mean(series: pd.Series) -> float:
    if series.empty:
        return np.nan
    mean = series.mean()
    return round(mean, 2) if pd.notna(mean) else np.nan


def _weighted_penalty(group_mean: float, jar_mean: float, base_pct: int) -> float:
    if pd.isna(group_mean) or pd.isna(jar_mean):
        return np.nan
    drop = round(group_mean - jar_mean, 2)
    return round(drop * (base_pct / 100), 4)


def classify_penalty_impact(weighted_penalty: float) -> str:
    """Classify the impact strength based on the absolute weighted penalty."""
    if pd.isna(weighted_penalty):
        return ""
    impact = abs(float(weighted_penalty))
    if impact >= HIGH_INFLUENCE_THRESHOLD:
        return HIGH_INFLUENCE_LABEL
    if impact >= MEDIUM_INFLUENCE_THRESHOLD:
        return MEDIUM_INFLUENCE_LABEL
    return ""


def _resolve_jar_groups(series: pd.Series, jar_scale: int | None) -> tuple[set[int], int, set[int], set[int]]:
    resolved_scale = _detect_jar_scale(series, jar_scale)
    if resolved_scale == 7:
        return SEVEN_POINT_JAR_SCORES, 4, {1, 2, 3}, {5, 6, 7}
    return FIVE_POINT_JAR_SCORES, 3, {1, 2}, {4, 5}


def _detect_jar_scale(series: pd.Series, jar_scale: int | None) -> int:
    if jar_scale in {5, 7}:
        return jar_scale

    clean_values = pd.to_numeric(series, errors="coerce").dropna().astype(int)
    if clean_values.empty:
        return DEFAULT_JAR_SCALE
    max_value = int(clean_values.max())
    if max_value >= 7:
        return 7
    return DEFAULT_JAR_SCALE


def build_summary_output(result_df: pd.DataFrame) -> pd.DataFrame:
    """Build a compact, reader-friendly summary from the result table."""
    summaries: list[dict[str, object]] = []

    for _, row in result_df.iterrows():
        too_little_summary = _describe_penalty_side(
            side_name="1+2",
            weighted_penalty=row.get("Weighted Penalty 1+2"),
            p_value=row.get("p-value 1+2"),
            sig=row.get("Sig 1+2"),
            impact_label=row.get("Impact Level 1+2", ""),
        )
        too_much_summary = _describe_penalty_side(
            side_name="4+5",
            weighted_penalty=row.get("Weighted Penalty 4+5"),
            p_value=row.get("p-value 4+5"),
            sig=row.get("Sig 4+5"),
            impact_label=row.get("Impact Level 4+5", ""),
        )
        key_message = _build_key_message(
            sig_12=row.get("Sig 1+2"),
            sig_45=row.get("Sig 4+5"),
            impact_12=row.get("Impact Level 1+2", ""),
            impact_45=row.get("Impact Level 4+5", ""),
        )
        summaries.append(
            {
                "Filter": row.get("Filter Condition", ""),
                "Attribute": row.get("Attribute", ""),
                "ประเด็นหลัก": key_message,
                "ฝั่ง 1+2": too_little_summary,
                "ฝั่ง 4+5": too_much_summary,
                "สรุปแนะนำ": _build_recommendation(
                    key_message=key_message,
                    impact_12=row.get("Impact Level 1+2", ""),
                    impact_45=row.get("Impact Level 4+5", ""),
                ),
            }
        )

    return pd.DataFrame(
        summaries,
        columns=[
            "Filter",
            "Attribute",
            "ประเด็นหลัก",
            "ฝั่ง 1+2",
            "ฝั่ง 4+5",
            "สรุปแนะนำ",
        ],
    )


def _describe_penalty_side(
    side_name: str,
    weighted_penalty: object,
    p_value: object,
    sig: object,
    impact_label: object,
) -> str:
    if pd.isna(weighted_penalty):
        return "ไม่มีข้อมูล"

    penalty_value = float(weighted_penalty)
    significance = "" if pd.isna(sig) else str(sig).strip()
    impact_text = str(impact_label).strip() if not pd.isna(impact_label) else ""
    impact_short = _to_thai_impact(impact_text)

    if pd.isna(p_value):
        if penalty_value == 0:
            return "ไม่มีฐานข้อมูลพอสำหรับเทียบกับ JAR"
        return f"Penalty {penalty_value:.2f} แต่ข้อมูลไม่พอทดสอบนัยสำคัญ"

    p_text = f"{float(p_value):.4f}"
    if significance in {"*", "**"}:
        impact_phrase = f", {impact_short}" if impact_short else ""
        return f"ต่างจาก JAR (p={p_text}, {significance}, penalty={penalty_value:.2f}{impact_phrase})"
    if significance == "ns":
        return f"ไม่ต่างจาก JAR (p={p_text}, penalty={penalty_value:.2f})"
    return f"p={p_text}, penalty={penalty_value:.2f}"


def _build_key_message(sig_12: object, sig_45: object, impact_12: object, impact_45: object) -> str:
    is_sig_12 = str(sig_12).strip() in {"*", "**"}
    is_sig_45 = str(sig_45).strip() in {"*", "**"}
    has_impact_12 = bool(str(impact_12).strip())
    has_impact_45 = bool(str(impact_45).strip())

    if is_sig_12 and is_sig_45:
        return "มีประเด็นทั้งฝั่ง Too little และ Too much"
    if is_sig_12:
        return "มีประเด็นฝั่ง Too little"
    if is_sig_45:
        return "มีประเด็นฝั่ง Too much"
    if has_impact_12 or has_impact_45:
        return "มี penalty แต่ยังไม่ชัดเรื่องนัยสำคัญ"
    return "ยังไม่พบประเด็นเด่น"


def _build_recommendation(key_message: str, impact_12: object, impact_45: object) -> str:
    impact_12_text = _to_thai_impact(str(impact_12).strip())
    impact_45_text = _to_thai_impact(str(impact_45).strip())

    if key_message == "มีประเด็นทั้งฝั่ง Too little และ Too much":
        return "ควรดูทั้งสองฝั่งร่วมกันและพิจารณาปรับ attribute นี้"
    if key_message == "มีประเด็นฝั่ง Too little":
        return f"ควรตรวจฝั่ง Too little เพิ่มเติม{f' ({impact_12_text})' if impact_12_text else ''}"
    if key_message == "มีประเด็นฝั่ง Too much":
        return f"ควรตรวจฝั่ง Too much เพิ่มเติม{f' ({impact_45_text})' if impact_45_text else ''}"
    if key_message == "มี penalty แต่ยังไม่ชัดเรื่องนัยสำคัญ":
        return "มีแนวโน้ม แต่ควรดูฐานข้อมูลและผลนัยสำคัญร่วมด้วย"
    return "ยังไม่จำเป็นต้องสรุปว่าต้องปรับจากผลชุดนี้"


def _to_thai_impact(impact_label: str) -> str:
    if impact_label == HIGH_INFLUENCE_LABEL:
        return "ผลกระทบสูง"
    if impact_label == MEDIUM_INFLUENCE_LABEL:
        return "มีผลต่อ overall liking"
    return ""
