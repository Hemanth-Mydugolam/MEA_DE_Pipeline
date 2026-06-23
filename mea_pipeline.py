"""
Combine MEA neuralMetrics CSV exports into per-metric Excel workbooks.

For every CSV file found under the experiment path, this script reads:
  - Analysis Start (line 4, value in column B)
  - Analysis End   (line 5, value in column B)
  - Electrode names (line 97, "Measurement" header row)
  - Per-electrode values for each metric:
        Number of Spikes        -> line 99
        Mean Firing Rate (Hz)   -> line 100
        Number of Bursts        -> line 103

Each metric gets its own output workbook with:
  - an "All" sheet containing every file as a row, all electrodes as columns
  - one sheet per well (A1, A2, ... D6) containing only that well's electrodes

Line numbers above are 1-indexed, matching what you see opening the CSV in Excel.
"""

import csv
import re
import sys
from pathlib import Path

import pandas as pd
import yaml

ANALYSIS_START_LINE = 4
ANALYSIS_END_LINE = 5
HEADER_LINE = 97

METRICS = {
    "Number_of_Spikes": 99,
    "Mean_Firing_Rate": 100,
    "Number_of_Bursts": 103,
}

WELL_ROWS = ["A", "B", "C", "D"]
WELL_COLS = range(1, 7)
WELLS = [f"{r}{c}" for r in WELL_ROWS for c in WELL_COLS]


def natural_sort_key(path: Path):
    return [int(part) if part.isdigit() else part.lower() for part in re.split(r"(\d+)", path.stem)]


def load_config(config_path: Path) -> dict:
    with open(config_path, "r", encoding="utf-8") as f:
        config = yaml.safe_load(f)
    required = ["experiment_path", "experiment_name"]
    missing = [k for k in required if k not in config]
    if missing:
        raise ValueError(f"Config is missing required keys: {missing}")
    config.setdefault("output_dir", "Output")
    return config


def split_csv_line(line: str) -> list:
    return next(csv.reader([line]))


def to_float(value: str):
    value = value.strip() if value else ""
    if value == "":
        return None
    try:
        return float(value)
    except ValueError:
        return None


def parse_csv_file(file_path: Path) -> dict:
    with open(file_path, "r", encoding="utf-8-sig") as f:
        lines = f.readlines()

    analysis_start = to_float(split_csv_line(lines[ANALYSIS_START_LINE - 1])[1])
    analysis_end = to_float(split_csv_line(lines[ANALYSIS_END_LINE - 1])[1])

    header_fields = split_csv_line(lines[HEADER_LINE - 1])
    electrodes = [e.strip() for e in header_fields[1:] if e.strip() != ""]

    metric_values = {}
    for metric_name, line_no in METRICS.items():
        fields = split_csv_line(lines[line_no - 1])
        values = fields[1:]
        metric_values[metric_name] = [
            to_float(values[i]) if i < len(values) else None
            for i in range(len(electrodes))
        ]

    return {
        "file": file_path.stem,
        "analysis_start": analysis_start,
        "analysis_end": analysis_end,
        "electrodes": electrodes,
        "metric_values": metric_values,
    }


def build_all_dataframe(records: list, metric_name: str) -> pd.DataFrame:
    electrode_order = records[0]["electrodes"]

    rows = []
    for record in records:
        row = {
            "File": record["file"],
            "Analysis Start": record["analysis_start"],
            "Analysis End": record["analysis_end"],
        }
        electrode_to_value = dict(zip(record["electrodes"], record["metric_values"][metric_name]))
        for electrode in electrode_order:
            row[electrode] = electrode_to_value.get(electrode)
        rows.append(row)

    columns = ["File", "Analysis Start", "Analysis End"] + electrode_order
    return pd.DataFrame(rows, columns=columns)


def well_sheet(all_df: pd.DataFrame, well: str) -> pd.DataFrame:
    well_columns = [c for c in all_df.columns if c.startswith(f"{well}_")]
    base_columns = ["File", "Analysis Start", "Analysis End"]
    return all_df[base_columns + well_columns]


def write_metric_workbook(all_df: pd.DataFrame, output_path: Path) -> None:
    with pd.ExcelWriter(output_path, engine="openpyxl") as writer:
        all_df.to_excel(writer, sheet_name="All", index=False)
        for well in WELLS:
            df = well_sheet(all_df, well)
            if df.shape[1] <= 3:
                continue
            df.to_excel(writer, sheet_name=well, index=False)


def main(config_path: str = "config.yaml"):
    config = load_config(Path(config_path))
    experiment_path = Path(config["experiment_path"])
    experiment_name = config["experiment_name"]
    output_dir = Path(config["output_dir"]) / experiment_name

    csv_files = sorted(experiment_path.glob("*.csv"), key=natural_sort_key)
    if not csv_files:
        raise FileNotFoundError(f"No CSV files found in {experiment_path}")

    print(f"Found {len(csv_files)} CSV file(s) in {experiment_path}")
    records = [parse_csv_file(f) for f in csv_files]

    output_dir.mkdir(parents=True, exist_ok=True)

    for metric_name in METRICS:
        all_df = build_all_dataframe(records, metric_name)
        output_path = output_dir / f"{metric_name}.xlsx"
        write_metric_workbook(all_df, output_path)
        print(f"Wrote {output_path}")

    print("Pipeline run successfully completed!")


if __name__ == "__main__":
    config_arg = sys.argv[1] if len(sys.argv) > 1 else "config.yaml"
    main(config_arg)
