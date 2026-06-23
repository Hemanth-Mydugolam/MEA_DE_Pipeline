# MEA Neural Metrics Pipeline

Combines a folder of AxIS `neuralMetrics` CSV exports into three easy-to-read
Excel workbooks — one each for **Mean Firing Rate**, **Number of Spikes**, and
**Number of Bursts** — with an `All` summary sheet plus one sheet per well
(A1, A2, ... D6).

---

## 1. First-time setup

You only need to do this once per computer.

### 1.1 Install Python

1. Download Python 3.10+ from https://www.python.org/downloads/
2. During install, check the box **"Add Python to PATH"**.
3. Verify it installed correctly. Open PowerShell and run:
   ```
   python --version
   ```
   You should see something like `Python 3.12.x`.

### 1.2 Get the project files

Make sure you have this folder with these files in it:
```
DE_Pipeline/
├── config.yaml
├── mea_pipeline.py
├── requirements.txt
└── README.md
```

### 1.3 Install the required Python packages

Open PowerShell, navigate to the project folder, and run:

```powershell
cd "C:\path\to\DE_Pipeline"
python -m pip install -r requirements.txt
```

This installs `pandas`, `openpyxl`, and `PyYAML`. You only need to do this
once (or again if `requirements.txt` changes).

---

## 2. Assumptions

The pipeline relies on the following assumptions about the input data. If any
of these don't hold, the script will either error out or silently produce
wrong output — check this list first when something looks off.

- **All input files are `.csv`** — the script reads every `.csv` file directly
  inside `experiment_path` (no subfolders, no `.xlsx`/`.txt` exports).
- **Every CSV has the identical AxIS export layout** — same row/column
  structure across all files in the experiment:
  - Line 4, column B → Analysis Start
  - Line 5, column B → Analysis End
  - Line 97 → header row with electrode names (`Measurement` row)
  - Line 99 → Number of Spikes values
  - Line 100 → Mean Firing Rate (Hz) values
  - Line 103 → Number of Bursts values

  These are fixed line numbers (`ANALYSIS_START_LINE`, `ANALYSIS_END_LINE`,
  `HEADER_LINE`, `METRICS` in `mea_pipeline.py`) — the script does **not**
  search for these rows by label, it reads them positionally.
- **The first file's electrode list defines the column order** for the `All`
  sheet of every metric. All other files are expected to report the same set
  of electrodes (in any order); a file with extra/missing electrodes will
  have those columns left blank rather than aligned/flagged.
- **Electrode names follow the `<Well><Number>` pattern** (e.g. `A1_11`,
  `D6_44`) — well-sheet splitting (`well_sheet()`) groups columns by matching
  the `A1_`, `A2_`, ... `D6_` prefix, over a fixed 4x6 well grid (rows A-D,
  columns 1-6).
- **Files are named so that natural numeric sort reflects experiment order**
  (e.g. `FileSegment1`, `FileSegment2`, ... `FileSegment30`) — this is the row
  order used in every output sheet.
- **CSV encoding is UTF-8 (with or without BOM)** — files are opened as
  `utf-8-sig`.

---

## 3. Running the pipeline

### 2.1 Edit `config.yaml`

Open `config.yaml` in any text editor and set three values:

```yaml
# Path to the folder containing the experiment's neuralMetrics CSV files
experiment_path: "C:\\Users\\you\\Box\\B6_30s bin_Treated at 10.82s"

# Name used for the output folder (and as a prefix in logging)
experiment_name: "ANE2267_hDRG_1482908_061026_DIV05"

# Folder where the experiment_name output folder will be created
output_dir: "Output"
```

Notes:
- `experiment_path` must point to the folder containing the `.csv` files
  (it will read **every** `.csv` file in that folder — no subfolders).
- Use double backslashes (`\\`) in Windows paths, or switch to forward
  slashes (`C:/Users/you/Box/...`), either works.
- `experiment_name` becomes the output subfolder name — pick something that
  identifies this experiment/run.
- `output_dir` is where that subfolder gets created. `"Output"` creates it
  next to the script; you can also give a full path.

### 2.2 Run the script

From PowerShell, in the project folder:

```powershell
python mea_pipeline.py
```

You should see output like:
```
Found 30 CSV file(s) in C:\Users\you\Box\B6_30s bin_Treated at 10.82s
Wrote Output\ANE2267_hDRG_1482908_061026_DIV05\Number_of_Spikes.xlsx
Wrote Output\ANE2267_hDRG_1482908_061026_DIV05\Mean_Firing_Rate.xlsx
Wrote Output\ANE2267_hDRG_1482908_061026_DIV05\Number_of_Bursts.xlsx
Pipeline run successfully completed!
```

#### Running with a different config file

If you keep multiple config files (e.g. for different experiments), pass the
path as an argument:

```powershell
python mea_pipeline.py path\to\other_config.yaml
```

---

## 4. Understanding the output

Inside `output_dir/<experiment_name>/` you'll find three Excel files:

- `Mean_Firing_Rate.xlsx`
- `Number_of_Spikes.xlsx`
- `Number_of_Bursts.xlsx`

Each workbook has the same layout:

- **`All` sheet** — one row per CSV file, in natural numeric order
  (FileSegment1, FileSegment2, ... FileSegment30, not alphabetical
  1,10,11,...2). Columns are: `File`, `Analysis Start`, `Analysis End`,
  followed by every electrode (e.g. `A1_11`, `A1_12`, ... `D6_44`).
- **One sheet per well** (`A1`, `A2`, ... `D6`) — same rows/order as `All`,
  but only that well's electrode columns plus `Analysis Start`/`Analysis End`.

---

## 5. Troubleshooting

| Problem | Fix |
|---|---|
| `python: command not found` | Python isn't on PATH — reinstall and check "Add Python to PATH". |
| `ModuleNotFoundError: No module named 'pandas'` (or yaml/openpyxl) | Run `python -m pip install -r requirements.txt` again. |
| `FileNotFoundError: No CSV files found in ...` | Check `experiment_path` in `config.yaml` is correct and contains `.csv` files directly (not in a subfolder). |
| Output looks empty / wrong electrode columns | Confirm the CSV layout still matches: line 97 = `Measurement` header row, line 99/100/103 = Number of Spikes / Mean Firing Rate / Number of Bursts. If AxIS changes its export format, the line numbers in `mea_pipeline.py` (`ANALYSIS_START_LINE`, `ANALYSIS_END_LINE`, `HEADER_LINE`, `METRICS`) need updating to match. |
| Want different metrics or rows | Edit the `METRICS` dictionary near the top of `mea_pipeline.py` — keys become output filenames, values are the 1-indexed CSV line numbers. |
