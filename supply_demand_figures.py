"""Generate the final TRB supply-and-demand figures with and without paper numbers.

The script is self-contained.  It reads ``Supply summary.xlsx`` and
``VOT&WTP.xlsx`` from the same folder as this file, creates one clean SVG and
one numbered SVG for every requested attribute, and writes two CSV reports.

Numbering rule
--------------
One paper is identified by Author + Year.  Papers are ordered by publication
year (newest first) and then by their original row order.  If one paper has
several rows or several plotted values, all of its marks use the same number.
"""
from __future__ import annotations

# ---------------------------------------------------------------------------
# 1. Imports and paths
# ---------------------------------------------------------------------------

import argparse
import math
import re
from collections import Counter, defaultdict
from dataclasses import dataclass
from pathlib import Path

import matplotlib as mpl
import matplotlib.pyplot as plt
import numpy as np
import pandas as pd
from matplotlib.lines import Line2D
from matplotlib.patches import Patch


SCRIPT_DIR = Path(__file__).resolve().parent
SUPPLY_FILE = SCRIPT_DIR / "Supply summary.xlsx"
VOT_FILE = SCRIPT_DIR / "VOT&WTP.xlsx"
DEFAULT_OUTPUT_DIR = SCRIPT_DIR / "final_figures"


# ---------------------------------------------------------------------------
# 2. Shared visual settings
# ---------------------------------------------------------------------------

mpl.rcParams.update(
    {
        "font.family": "sans-serif",
        "font.sans-serif": ["Arial", "DejaVu Sans", "Liberation Sans"],
        "font.size": 12,
        "axes.unicode_minus": False,
        "axes.linewidth": 0.8,
        "axes.spines.right": True,
        "axes.spines.top": True,
        # Keep SVG text editable instead of converting it to paths.
        "svg.fonttype": "none",
        "savefig.facecolor": "white",
    }
)

FIGURE_WIDTH_IN = 13.5
FIGURE_HEIGHT_IN = 8.5
AXES_LEFT = 0.062
AXES_RIGHT = 0.975

# All figures use the same physical bar width and point size.
GLOBAL_BAR_WIDTH_PX = 7.0
GLOBAL_POINT_AREA_PT2 = 55.5
NUMBER_FONT_SIZE_PT = 7.0

USE_ORDER = ("Intracity", "Intercity", "Airport shuttle", "Not specified")
BAR_COLORS = {
    "Intracity": "#F28E8E",
    "Intercity": "#F2C14E",
    "Airport shuttle": "#74B9E8",
    "Not specified": "#A7DCA0",
}
POINT_COLORS = {
    "Intracity": "#8B1A1A",
    "Intercity": "#D97706",
    "Airport shuttle": "#168AAD",
    "Not specified": "#4F8A4B",
}

BROAD_REGIONS = {
    "asia",
    "europe",
    "africa",
    "oceania",
    "north america",
    "south america",
    "latin america",
    "middle east",
    "not specified",
    "unknown",
    "global",
    "worldwide",
}


# ---------------------------------------------------------------------------
# 3. Figure definitions
# ---------------------------------------------------------------------------

@dataclass(frozen=True)
class PlotSpec:
    """Describe the source columns, axis limits, and output name of one figure."""

    workbook: str
    sheet: str
    parameter: str
    output_name: str
    y_label: str
    y_segments: tuple[tuple[float, float], ...]
    lower_candidates: tuple[str, ...]
    upper_candidates: tuple[str, ...]
    scale: float = 1.0
    lock_y_limits: bool = False


PLOT_SPECS = (
    PlotSpec(
        "Supply summary.xlsx",
        "Access time",
        "Access+Egress",
        "Access_time",
        "Access and egress time (min)",
        ((0, 60),),
        ("MIN", "Min"),
        ("MAX", "Max"),
    ),
    PlotSpec(
        "Supply summary.xlsx",
        "Cost",
        "Cost (USD/km)",
        "Cost",
        "Cost (USD/km)",
        ((0, 12),),
        ("Min", "MIN"),
        ("Max", "MAX"),
    ),
    PlotSpec(
        "Supply summary.xlsx",
        "Speed",
        "Speed",
        "Speed",
        "Speed (km/h)",
        ((0, 500),),
        ("Min", "MIN"),
        ("Max", "MAX"),
    ),
    PlotSpec(
        "Supply summary.xlsx",
        "Capacity",
        "Capacity",
        "Capacity",
        "Capacity (seats/vehicle)",
        ((0, 15), (35, 45)),
        ("MIN", "Min"),
        ("MAX", "Max"),
    ),
    PlotSpec(
        "Supply summary.xlsx",
        "Altitude",
        "Altitude",
        "Altitude",
        "Altitude (m)",
        ((0, 1700), (2900, 3300)),
        ("MIN", "Min"),
        ("MAX", "Max"),
    ),
    PlotSpec(
        "Supply summary.xlsx",
        "Range",
        "Range",
        "Range",
        "Range (km)",
        ((0, 450),),
        ("MIN", "Min"),
        ("MAX", "Max"),
    ),
    PlotSpec(
        "Supply summary.xlsx",
        "Time spent at vertiports",
        "Waiting/Boarding/Deboading time (min)",
        "Process",
        "Time spent at the vertiports (min)",
        ((0, 70),),
        ("MIN", "Min"),
        ("MAX", "Max"),
    ),
    PlotSpec(
        "VOT&WTP.xlsx",
        "VOT_processed",
        "VOT (USD/h)",
        "VOT",
        "VOT (USD/h)",
        ((0, 180), (480, 540)),
        ("Min", "VOT_min"),
        ("Max", "VOT_max"),
        lock_y_limits=True,
    ),
    PlotSpec(
        "Supply summary.xlsx",
        "Vertiport number (area size)",
        "Vertiport density (vertiport/km2)",
        "Vertiport_density_area",
        "Vertiport density (vertiport per thousand km²)",
        ((0, 90),),
        ("Min", "MIN"),
        ("Max", "MAX"),
        scale=1000.0,
        lock_y_limits=True,
    ),
)


# ---------------------------------------------------------------------------
# 4. Text normalization and country lookup
# ---------------------------------------------------------------------------

# Purpose: turn mixed Excel cell contents into consistent text for matching.
def clean_text(value: object) -> str:
    """Convert an Excel value to normalized one-line text."""
    if pd.isna(value):
        return ""
    return re.sub(r"\s+", " ", str(value)).strip()


# Purpose: standardize author names while preserving paper suffixes such as (a).
def normalize_author(value: object) -> str:
    """Normalize author strings while retaining suffixes such as (a) or (b)."""
    text = clean_text(value)
    text = re.sub(r"\s*\((?:19|20)\d{2}\)\s*$", "", text)
    text = re.sub(r"\s*&\s*", " and ", text)
    return re.sub(r"\s+", " ", text).strip()


# Purpose: create a fallback author key when a country lookup lacks an a/b suffix.
def author_without_suffix(author: str) -> str:
    """Remove a paper suffix only for fallback country matching."""
    return re.sub(r"\s*\([a-z]\)\s*$", "", author, flags=re.IGNORECASE).strip()


# Purpose: collapse common country-name variants into one display value.
def clean_country(value: object) -> str:
    """Standardize common country-name variants."""
    country = clean_text(value)
    replacements = {
        "usa": "US",
        "u.s.": "US",
        "u.s.a.": "US",
        "united states": "US",
        "uk": "UK",
        "united kingdom": "UK",
    }
    return replacements.get(country.casefold(), country)


# Purpose: distinguish an actual country from a broad region or missing value.
def is_specific_country(value: object) -> bool:
    """Return True when a value names a specific country rather than a region."""
    country = clean_country(value)
    return bool(country) and country.casefold() not in BROAD_REGIONS


# Purpose: assign every source row to one of the four plotting categories.
def normalize_use(value: object) -> str:
    """Map source descriptions to the four colors used in the figures."""
    text = clean_text(value).casefold()
    if "intra" in text:
        return "Intracity"
    if "inter" in text:
        return "Intercity"
    if "airport" in text or "railway" in text or "station access" in text:
        return "Airport shuttle"
    return "Not specified"


# Purpose: locate a required input workbook beside this script and fail clearly.
def workbook_path(filename: str) -> Path:
    """Resolve an input workbook inside the Submit folder."""
    path = SCRIPT_DIR / filename
    if not path.exists():
        raise FileNotFoundError(f"Required input workbook not found: {path}")
    return path


# Purpose: build reusable Author + Year lookups for country information.
def build_country_lookup() -> tuple[
    dict[tuple[str, int], Counter], dict[tuple[str, int], Counter]
]:
    """Collect country information from all usable sheets in both workbooks."""
    exact: dict[tuple[str, int], Counter] = defaultdict(Counter)
    base: dict[tuple[str, int], Counter] = defaultdict(Counter)
    for filename in ("Supply summary.xlsx", "VOT&WTP.xlsx"):
        path = workbook_path(filename)
        for sheet_name in pd.ExcelFile(path).sheet_names:
            frame = pd.read_excel(path, sheet_name=sheet_name)
            if "Author" not in frame.columns or "Year" not in frame.columns:
                continue
            for _, row in frame.iterrows():
                author = normalize_author(row.get("Author"))
                year = pd.to_numeric(row.get("Year"), errors="coerce")
                if not author or pd.isna(year):
                    continue
                country = ""
                if "Country" in frame.columns and is_specific_country(row.get("Country")):
                    country = clean_country(row.get("Country"))
                elif "Region" in frame.columns and is_specific_country(row.get("Region")):
                    country = clean_country(row.get("Region"))
                if not country:
                    continue
                exact[(author.casefold(), int(year))][country] += 1
                base[(author_without_suffix(author).casefold(), int(year))][country] += 1
    return dict(exact), dict(base)


# Purpose: select the dominant country only when the lookup result is unambiguous.
def choose_lookup_country(counter: Counter | None) -> str:
    """Use a country only when the most frequent match is unambiguous."""
    if not counter:
        return ""
    ranked = counter.most_common()
    if len(ranked) > 1 and ranked[0][1] == ranked[1][1]:
        return ""
    return ranked[0][0]


# Purpose: obtain the best available country for one plotting row.
def resolve_country(
    row: pd.Series,
    exact_lookup: dict[tuple[str, int], Counter],
    base_lookup: dict[tuple[str, int], Counter],
) -> str:
    """Resolve a specific country from the row or another matching sheet."""
    if "Country" in row.index and is_specific_country(row.get("Country")):
        return clean_country(row.get("Country"))
    if "Region" in row.index and is_specific_country(row.get("Region")):
        return clean_country(row.get("Region"))

    author = normalize_author(row.get("Author"))
    year = pd.to_numeric(row.get("Year"), errors="coerce")
    if author and not pd.isna(year):
        country = choose_lookup_country(exact_lookup.get((author.casefold(), int(year))))
        if not country:
            country = choose_lookup_country(
                base_lookup.get((author_without_suffix(author).casefold(), int(year)))
            )
        if country:
            return country

    region = clean_country(row.get("Region")) if "Region" in row.index else ""
    return region or "Country not specified"


# ---------------------------------------------------------------------------
# 5. Read and prepare plotting data
# ---------------------------------------------------------------------------

# Purpose: support minor capitalization differences in Min/Max column names.
def first_existing(frame: pd.DataFrame, candidates: tuple[str, ...]) -> str | None:
    """Return the first candidate column that is present in the worksheet."""
    return next((column for column in candidates if column in frame.columns), None)


# Purpose: convert a numeric cell or a textual range into lower and upper values.
def parse_numeric_range(value: object) -> tuple[float, float] | None:
    """Parse a single value or the first two values in a textual range."""
    if pd.isna(value) or isinstance(value, pd.Timestamp):
        return None
    if isinstance(value, (int, float, np.integer, np.floating)):
        numeric = float(value)
        return numeric, numeric

    text = clean_text(value).replace("−", "-").replace("–", "-")
    matches = re.findall(r"(?<!\d)-?\d+(?:\.\d+)?(?:[eE][+-]?\d+)?", text)
    if not matches:
        return None
    values = [float(match) for match in matches]
    if len(values) == 1:
        return values[0], values[0]
    return min(values[0], values[1]), max(values[0], values[1])


# Purpose: convert one source worksheet into normalized plotting records.
def prepare_data(
    spec: PlotSpec,
    exact_lookup: dict[tuple[str, int], Counter],
    base_lookup: dict[tuple[str, int], Counter],
) -> pd.DataFrame:
    """Read one worksheet and convert every valid row into a plotting record."""
    path = workbook_path(spec.workbook)
    frame = pd.read_excel(path, sheet_name=spec.sheet).copy()
    lower_column = first_existing(frame, spec.lower_candidates)
    upper_column = first_existing(frame, spec.upper_candidates)
    records: list[dict[str, object]] = []

    for source_index, row in frame.iterrows():
        author = normalize_author(row.get("Author"))
        year = pd.to_numeric(row.get("Year"), errors="coerce")
        if not author or pd.isna(year):
            continue

        lower = (
            pd.to_numeric(row.get(lower_column), errors="coerce")
            if lower_column
            else np.nan
        )
        upper = (
            pd.to_numeric(row.get(upper_column), errors="coerce")
            if upper_column
            else np.nan
        )
        parsed = parse_numeric_range(row.get(spec.parameter))
        if pd.isna(lower) or pd.isna(upper):
            if parsed is None:
                continue
            lower, upper = parsed

        # Excel occasionally turns text such as 5-10 into a date serial.
        if spec.sheet == "Access time" and max(float(lower), float(upper)) > 1440:
            recovered = None
            for candidate in ("Access time", "Access", "Egress"):
                recovered = parse_numeric_range(row.get(candidate))
                if recovered is not None:
                    break
            if recovered is None:
                continue
            lower, upper = recovered

        lower = float(lower) * spec.scale
        upper = float(upper) * spec.scale
        if not np.isfinite(lower) or not np.isfinite(upper):
            continue
        if lower > upper:
            lower, upper = upper, lower

        records.append(
            {
                "source_index": int(source_index),
                "author": author,
                "year": int(year),
                "country": resolve_country(row, exact_lookup, base_lookup),
                "use": normalize_use(row.get("Use")),
                "lower": lower,
                "upper": upper,
                "is_range": not math.isclose(
                    lower, upper, rel_tol=1e-10, abs_tol=1e-12
                ),
            }
        )

    data = pd.DataFrame(records)
    if data.empty:
        return data

    # First establish stable positions inside each year.  Bars are later
    # repositioned to a common physical width by prepare_uniform_data().
    data["x"] = 0.0
    data["bar_width"] = 0.055
    use_rank = {use: rank for rank, use in enumerate(USE_ORDER)}
    cluster_width = 0.84
    range_rows = data.loc[data["is_range"]]
    max_bars = (
        int(range_rows.groupby("year").size().max()) if not range_rows.empty else 1
    )
    initial_bar_width = cluster_width / max_bars

    for year, indices in data.groupby("year", sort=True).groups.items():
        ordered = (
            data.loc[list(indices)]
            .assign(_use_rank=lambda item: item["use"].map(use_rank))
            .sort_values(
                ["_use_rank", "author", "lower", "upper", "source_index"],
                kind="stable",
            )
        )
        bars = ordered.loc[ordered["is_range"]]
        points = ordered.loc[~ordered["is_range"]]
        if not bars.empty:
            centers = year + (
                np.arange(len(bars)) - (len(bars) - 1) / 2
            ) * initial_bar_width
            data.loc[bars.index, "x"] = centers
        if not points.empty:
            point_width = cluster_width / max(len(points), 1)
            centers = year + (
                np.arange(len(points)) - (len(points) - 1) / 2
            ) * point_width
            data.loc[points.index, "x"] = centers

    return data


# Purpose: enforce the common physical bar width used by all nine figures.
def prepare_uniform_data(
    spec: PlotSpec,
    exact_lookup: dict[tuple[str, int], Counter],
    base_lookup: dict[tuple[str, int], Counter],
) -> pd.DataFrame:
    """Apply the same physical bar width to every range in every figure."""
    data = prepare_data(spec, exact_lookup, base_lookup)
    if data.empty:
        return data

    # Recover month-day pairs when process-time ranges were saved as dates.
    if spec.sheet == "Time spent at vertiports":
        source = pd.read_excel(workbook_path(spec.workbook), sheet_name=spec.sheet)
        for index, row in data.iterrows():
            raw = source.at[int(row["source_index"]), spec.parameter]
            if isinstance(raw, pd.Timestamp):
                lower, upper = sorted((float(raw.month), float(raw.day)))
                data.at[index, "lower"] = lower
                data.at[index, "upper"] = upper
                data.at[index, "is_range"] = not np.isclose(lower, upper)

    year_span = int(data["year"].max()) - int(data["year"].min()) + 1.1
    axes_width_px = (AXES_RIGHT - AXES_LEFT) * FIGURE_WIDTH_IN * 96.0
    fixed_width = GLOBAL_BAR_WIDTH_PX * year_span / axes_width_px

    for year, group in data.loc[data["is_range"]].groupby("year", sort=True):
        ordered = group.sort_values(["x", "source_index"], kind="stable")
        if len(ordered) * fixed_width >= 0.94:
            raise ValueError(
                f"{spec.output_name}: too many bars in {year} for a safe slot"
            )
        centers = year + (
            np.arange(len(ordered)) - (len(ordered) - 1) / 2
        ) * fixed_width
        data.loc[ordered.index, "x"] = centers

    data["bar_width"] = fixed_width
    return data


# ---------------------------------------------------------------------------
# 6. Geometry checks and axis helpers
# ---------------------------------------------------------------------------

# Purpose: provide a quality-control check for bars occupying the same year.
def count_bar_overlaps(data: pd.DataFrame) -> int:
    """Count horizontal overlaps between range bars in the same year."""
    overlaps = 0
    for _, group in data.loc[data["is_range"]].groupby("year"):
        intervals = sorted(
            (
                float(row["x"] - row["bar_width"] / 2),
                float(row["x"] + row["bar_width"] / 2),
            )
            for _, row in group.iterrows()
        )
        for previous, current in zip(intervals[:-1], intervals[1:]):
            if current[0] < previous[1] - 1e-10:
                overlaps += 1
    return overlaps


# Purpose: verify that every range bar in a figure uses one common width.
def count_distinct_bar_widths(data: pd.DataFrame) -> int:
    """Return the number of distinct range-bar widths in a figure."""
    widths = data.loc[data["is_range"], "bar_width"].round(12).unique()
    return int(len(widths))


# Purpose: tighten ordinary y-axes while preserving deliberate broken-axis limits.
def adjusted_y_segments(
    data: pd.DataFrame,
    configured_segments: tuple[tuple[float, float], ...],
) -> tuple[tuple[float, float], ...]:
    """Keep broken axes fixed and tighten a single y-axis above its data."""
    if len(configured_segments) > 1:
        return configured_segments

    lower_limit, upper_limit = configured_segments[0]
    visible_values: list[float] = []
    for _, row in data.iterrows():
        lower = float(row["lower"])
        upper = float(row["upper"])
        if upper < lower_limit or lower > upper_limit:
            continue
        visible_values.extend([max(lower, lower_limit), min(upper, upper_limit)])
    if not visible_values:
        return configured_segments

    visible_max = max(visible_values)
    span = upper_limit - lower_limit
    raw_upper = visible_max + max(span * 0.08, abs(visible_max) * 0.025, 1e-9)
    magnitude = abs(raw_upper)
    if magnitude < 100:
        quantum = 5.0
    elif magnitude < 500:
        quantum = 10.0
    elif magnitude < 2000:
        quantum = 50.0
    else:
        quantum = 100.0
    return ((lower_limit, math.ceil(raw_upper / quantum) * quantum),)


# Purpose: calculate readable panel heights for figures containing axis breaks.
def segment_height_ratios(
    segments_top_to_bottom: list[tuple[float, float]],
) -> list[float]:
    """Give broken-axis panels readable heights without extreme imbalance."""
    spans = np.array(
        [upper - lower for lower, upper in segments_top_to_bottom], dtype=float
    )
    if len(spans) == 1:
        return [1.0]
    spans = np.sqrt(np.maximum(spans, 1e-9))
    return np.clip(spans / spans.min(), 1.0, 3.0).tolist()


# Purpose: visually mark the discontinuity between broken-axis panels.
def add_break_marks(axes: list[plt.Axes]) -> None:
    """Draw diagonal marks between adjacent broken-axis panels."""
    if len(axes) < 2:
        return
    size = 0.008
    for upper_axis, lower_axis in zip(axes[:-1], axes[1:]):
        style = dict(color="black", clip_on=False, linewidth=0.8)
        upper_axis.plot(
            (-size, +size), (-size, +size), transform=upper_axis.transAxes, **style
        )
        upper_axis.plot(
            (1 - size, 1 + size),
            (-size, +size),
            transform=upper_axis.transAxes,
            **style,
        )
        lower_axis.plot(
            (-size, +size),
            (1 - size, 1 + size),
            transform=lower_axis.transAxes,
            **style,
        )
        lower_axis.plot(
            (1 - size, 1 + size),
            (1 - size, 1 + size),
            transform=lower_axis.transAxes,
            **style,
        )


# Purpose: find the visible top of a bar or point where its number should sit.
def choose_anchor(
    axes_with_limits: list[tuple[plt.Axes, tuple[float, float]]],
    lower: float,
    upper: float,
) -> tuple[plt.Axes, float]:
    """Choose the highest visible part of a bar or point for its number."""
    for axis, (segment_lower, segment_upper) in axes_with_limits:
        overlap_lower = max(lower, segment_lower)
        overlap_upper = min(upper, segment_upper)
        if overlap_lower <= overlap_upper:
            return axis, overlap_upper

    midpoint = (lower + upper) / 2
    axis, limits = min(
        axes_with_limits,
        key=lambda item: min(
            abs(midpoint - item[1][0]), abs(midpoint - item[1][1])
        ),
    )
    return axis, float(np.clip(midpoint, limits[0], limits[1]))


# ---------------------------------------------------------------------------
# 7. Legend and paper-number functions
# ---------------------------------------------------------------------------

# Purpose: construct only the legend entries represented in the current figure.
def build_legend(data: pd.DataFrame) -> tuple[list[object], list[str]]:
    """Build legend entries only for mark types actually present."""
    handles: list[object] = []
    labels: list[str] = []
    for use in USE_ORDER:
        use_rows = data.loc[data["use"] == use]
        if use_rows.empty:
            continue
        if bool(use_rows["is_range"].any()):
            handles.append(
                Patch(
                    facecolor=BAR_COLORS[use],
                    edgecolor="black",
                    linewidth=0.6,
                    alpha=0.55,
                )
            )
            labels.append(f"{use} (range)")
        if bool((~use_rows["is_range"]).any()):
            handles.append(
                Line2D(
                    [0],
                    [0],
                    marker="o",
                    linestyle="none",
                    markersize=7,
                    color=POINT_COLORS[use],
                )
            )
            labels.append(f"{use} (single value)")
    return handles, labels


# Purpose: assign one stable number to each unique Author + Year paper.
def paper_number_map(data: pd.DataFrame) -> dict[tuple[str, int], int]:
    """Assign one number to each unique Author + Year paper."""
    papers = (
        data.groupby(["author", "year"], sort=False)["source_index"]
        .min()
        .reset_index()
        .sort_values(
            ["year", "source_index"],
            ascending=[False, True],
            kind="stable",
        )
    )
    return {
        (str(row.author), int(row.year)): number
        for number, row in enumerate(papers.itertuples(index=False), start=1)
    }


# Purpose: add the assigned paper number above every corresponding mark.
def add_number_marks(
    data: pd.DataFrame,
    axes_with_limits: list[tuple[plt.Axes, tuple[float, float]]],
    numbers: dict[tuple[str, int], int],
) -> None:
    """Place the paper number immediately above each corresponding mark."""
    for _, row in data.iterrows():
        axis, y_value = choose_anchor(
            axes_with_limits, float(row["lower"]), float(row["upper"])
        )
        axis.annotate(
            str(numbers[(str(row["author"]), int(row["year"]))]),
            xy=(float(row["x"]), y_value),
            xytext=(0, 3.2),
            textcoords="offset points",
            ha="center",
            va="bottom",
            fontsize=NUMBER_FONT_SIZE_PT,
            fontweight="normal",
            color="#111111",
            annotation_clip=True,
            clip_on=True,
            zorder=7,
        )


# ---------------------------------------------------------------------------
# 8. Plot one attribute and save both variants
# ---------------------------------------------------------------------------

# Purpose: draw one attribute and save both clean and numbered SVG versions.
def plot_one(
    spec: PlotSpec,
    exact_lookup: dict[tuple[str, int], Counter],
    base_lookup: dict[tuple[str, int], Counter],
    output_dir: Path,
) -> tuple[dict[str, object], list[dict[str, object]]]:
    """Draw one attribute, save clean/numbered SVGs, and return QA records."""
    data = prepare_uniform_data(spec, exact_lookup, base_lookup)
    if data.empty:
        raise ValueError(
            f"No plottable data found for {spec.workbook} / {spec.sheet}"
        )
    if count_bar_overlaps(data):
        raise ValueError(f"Overlapping bars detected in {spec.output_name}")
    if count_distinct_bar_widths(data) > 1:
        raise ValueError(f"Non-uniform bar widths detected in {spec.output_name}")

    year_min = int(data["year"].min())
    year_max = int(data["year"].max())
    years = list(range(year_min, year_max + 1))
    x_limits = (year_min - 0.55, year_max + 0.55)
    plot_segments = (
        spec.y_segments
        if spec.lock_y_limits
        else adjusted_y_segments(data, spec.y_segments)
    )
    segments_top_to_bottom = list(reversed(plot_segments))

    figure = plt.figure(figsize=(FIGURE_WIDTH_IN, FIGURE_HEIGHT_IN))
    grid = figure.add_gridspec(
        len(segments_top_to_bottom),
        1,
        left=AXES_LEFT,
        right=AXES_RIGHT,
        bottom=0.10,
        top=0.975,
        hspace=0.055,
        height_ratios=segment_height_ratios(segments_top_to_bottom),
    )
    axes = [
        figure.add_subplot(grid[index, 0])
        for index in range(len(segments_top_to_bottom))
    ]
    axes_with_limits = list(zip(axes, segments_top_to_bottom))

    for axis, limits in axes_with_limits:
        for _, row in data.iterrows():
            use = str(row["use"])
            if bool(row["is_range"]):
                axis.bar(
                    float(row["x"]),
                    float(row["upper"] - row["lower"]),
                    bottom=float(row["lower"]),
                    width=float(row["bar_width"]),
                    color=BAR_COLORS[use],
                    edgecolor="black",
                    linewidth=0.8,
                    alpha=0.55,
                    zorder=3,
                )
            else:
                axis.scatter(
                    float(row["x"]),
                    float(row["lower"]),
                    s=GLOBAL_POINT_AREA_PT2,
                    color=POINT_COLORS[use],
                    edgecolor="white",
                    linewidth=0.55,
                    zorder=4,
                )

        axis.set_ylim(*limits)
        axis.set_xlim(*x_limits)
        axis.set_xticks(years)
        axis.grid(axis="y", color="#E5E5E5", linewidth=0.45, zorder=0)
        axis.tick_params(axis="both", labelsize=12, length=5, width=1.0)

        # Explicitly include both ends of every y-axis segment.
        lower_limit, upper_limit = limits
        ticks = [
            float(value)
            for value in axis.get_yticks()
            if lower_limit <= value <= upper_limit
        ]
        ticks.extend([float(lower_limit), float(upper_limit)])
        axis.set_yticks(sorted(set(round(value, 10) for value in ticks)))

    # Hide only the internal borders of a broken axis; keep the outer frame.
    for axis in axes[:-1]:
        axis.spines["bottom"].set_visible(False)
        axis.tick_params(axis="x", which="both", bottom=False, labelbottom=False)
    for axis in axes[1:]:
        axis.spines["top"].set_visible(False)

    axes[-1].set_xlabel("Year", fontsize=14, labelpad=8)
    figure.text(
        0.018,
        0.535,
        spec.y_label,
        rotation=90,
        va="center",
        ha="center",
        fontsize=14,
    )
    add_break_marks(axes)

    handles, labels = build_legend(data)
    if handles:
        midpoint = (year_min + year_max) / 2
        left_rows = data.loc[data["year"] <= midpoint, "upper"]
        right_rows = data.loc[data["year"] > midpoint, "upper"]
        left_peak = float(left_rows.max()) if not left_rows.empty else -np.inf
        right_peak = float(right_rows.max()) if not right_rows.empty else -np.inf
        location = "upper left" if left_peak <= right_peak else "upper right"
        legend = axes[0].legend(
            handles,
            labels,
            loc=location,
            ncol=2,
            frameon=True,
            framealpha=0.92,
            facecolor="white",
            edgecolor="#BDBDBD",
            fontsize=7.0,
            handletextpad=0.55,
            columnspacing=1.0,
            borderpad=0.6,
            labelspacing=0.45,
        )
        legend.set_zorder(8)

    clean_dir = output_dir / "without_numbers"
    numbered_dir = output_dir / "with_numbers"
    clean_dir.mkdir(parents=True, exist_ok=True)
    numbered_dir.mkdir(parents=True, exist_ok=True)

    # Save the clean figure before adding any numbers.
    clean_svg = clean_dir / f"{spec.output_name}.svg"
    figure.savefig(clean_svg, format="svg")

    # Add the paper numbers to the existing geometry and save the second copy.
    numbers = paper_number_map(data)
    add_number_marks(data, axes_with_limits, numbers)
    numbered_svg = numbered_dir / f"{spec.output_name}.svg"
    figure.savefig(numbered_svg, format="svg")
    plt.close(figure)

    mapping_rows: list[dict[str, object]] = []
    for (author, year), number in numbers.items():
        rows = data.loc[(data["author"] == author) & (data["year"] == year)]
        mapping_rows.append(
            {
                "Figure": spec.output_name,
                "Number": number,
                "Author": author,
                "Year": year,
                "Countries/regions": ", ".join(
                    sorted(str(value) for value in rows["country"].unique())
                ),
                "Marks": len(rows),
            }
        )

    report = {
        "Figure": spec.output_name,
        "Workbook": spec.workbook,
        "Sheet": spec.sheet,
        "Rows plotted": len(data),
        "Papers numbered": len(numbers),
        "Bar overlap count": count_bar_overlaps(data),
        "Distinct bar widths": count_distinct_bar_widths(data),
        "Bar width (px)": GLOBAL_BAR_WIDTH_PX,
        "Point area (pt^2)": GLOBAL_POINT_AREA_PT2,
        "Number font (pt)": NUMBER_FONT_SIZE_PT,
        "Y segments": str(plot_segments),
    }
    return report, mapping_rows


# ---------------------------------------------------------------------------
# 9. Program entry point
# ---------------------------------------------------------------------------

# Purpose: read the optional command-line output directory.
def parse_args() -> argparse.Namespace:
    """Parse command-line options and return the selected output directory."""
    parser = argparse.ArgumentParser(
        description="Generate clean and numbered TRB SVG figures."
    )
    parser.add_argument(
        "--output-dir",
        type=Path,
        default=DEFAULT_OUTPUT_DIR,
        help="Output folder. Default: Submit/final_figures",
    )
    return parser.parse_args()


# Purpose: coordinate input validation, plotting, and CSV report generation.
def main() -> None:
    """Validate inputs, draw all figures, and save the QA/mapping CSV files."""
    workbook_path(SUPPLY_FILE.name)
    workbook_path(VOT_FILE.name)
    args = parse_args()
    output_dir = args.output_dir.resolve()

    exact_lookup, base_lookup = build_country_lookup()
    reports: list[dict[str, object]] = []
    mappings: list[dict[str, object]] = []

    for spec in PLOT_SPECS:
        print(
            f"Plotting {spec.output_name} from {spec.workbook} / {spec.sheet} ...",
            flush=True,
        )
        report, mapping = plot_one(
            spec, exact_lookup, base_lookup, output_dir=output_dir
        )
        reports.append(report)
        mappings.extend(mapping)

    pd.DataFrame(reports).to_csv(
        output_dir / "figure_manifest.csv", index=False, encoding="utf-8-sig"
    )
    pd.DataFrame(mappings).to_csv(
        output_dir / "author_number_mapping.csv",
        index=False,
        encoding="utf-8-sig",
    )

    print(pd.DataFrame(reports).to_string(index=False), flush=True)
    print(f"Saved clean SVGs to: {output_dir / 'without_numbers'}", flush=True)
    print(f"Saved numbered SVGs to: {output_dir / 'with_numbers'}", flush=True)


if __name__ == "__main__":
    main()
