
#!/usr/bin/env python3
import argparse
import datetime as dt
import hashlib
import importlib
import json
import os
from pathlib import Path
import subprocess
import sys
from typing import Any, Dict, Optional


REQUIRED_PACKAGES = {
    "pandas": "pandas",
    "plotly": "plotly",
    "openpyxl": "openpyxl",
}


def ensure_packages() -> None:
    missing = []
    for import_name, package_name in REQUIRED_PACKAGES.items():
        try:
            importlib.import_module(import_name)
        except ImportError:
            missing.append(package_name)

    if missing:
        print(f"Installing missing dependencies: {', '.join(missing)}")
        subprocess.check_call([sys.executable, "-m", "pip", "install", *missing])


ensure_packages()

import pandas as pd  # noqa: E402
import plotly.express as px  # noqa: E402
import plotly.io as pio  # noqa: E402


CONFIG_DIR = Path.home() / ".gantt_generator"
CONFIG_DIR.mkdir(parents=True, exist_ok=True)

STATUS_COLORS = {
    "Complete": "green",
    "In Progress": "orange",
    "Pending": "lightgray",
    "Upcoming": "blue",
    "Overdue": "red",
}
REQUIRED_COLUMNS = {"Task", "Start Date", "End Date", "Status"}


def try_import_tk():
    try:
        import tkinter as tk
        from tkinter import filedialog, messagebox, ttk
        return tk, ttk, filedialog, messagebox
    except Exception:
        return None, None, None, None


def sanitize_filename(value: str) -> str:
    cleaned = "".join(ch if ch.isalnum() or ch in (" ", "-", "_") else "_" for ch in value).strip()
    return cleaned or "Training_Project_Timeline"


def normalize_bool(value: Any) -> bool:
    if isinstance(value, bool):
        return value
    if pd.isna(value):
        return False
    text = str(value).strip().lower()
    return text in {"true", "1", "yes", "y"}


def get_project_id(title: str, file_path: str) -> str:
    base = f"{title}_{Path(file_path).resolve()}"
    return hashlib.md5(base.encode("utf-8")).hexdigest()


def config_path_for(project_id: str) -> Path:
    return CONFIG_DIR / f"{project_id}.json"


def list_saved_configs():
    return sorted(
        [p for p in CONFIG_DIR.glob("*.json") if p.name != "last_used.json"],
        key=lambda p: p.stat().st_mtime,
        reverse=True,
    )


def load_project_config(project_id: str) -> Optional[Dict[str, Any]]:
    path = config_path_for(project_id)
    if path.exists():
        return json.loads(path.read_text(encoding="utf-8"))
    return None


def save_project_config(project_id: str, config: Dict[str, Any]) -> None:
    config_path_for(project_id).write_text(json.dumps(config, indent=2), encoding="utf-8")
    (CONFIG_DIR / "last_used.json").write_text(
        json.dumps({"last_project": project_id, "file_path": config["file_path"]}, indent=2),
        encoding="utf-8",
    )


def load_last_used() -> Optional[Dict[str, Any]]:
    last_path = CONFIG_DIR / "last_used.json"
    if last_path.exists():
        return json.loads(last_path.read_text(encoding="utf-8"))
    return None


def open_file(path: Path) -> None:
    try:
        if sys.platform.startswith("darwin"):
            subprocess.run(["open", str(path)], check=False)
        elif os.name == "nt":
            os.startfile(str(path))  # type: ignore[attr-defined]
        else:
            subprocess.run(["xdg-open", str(path)], check=False)
    except Exception as exc:
        print(f"Warning: Could not open file automatically: {exc}")


def choose_file_gui() -> Optional[str]:
    tk, _, filedialog, _, = try_import_tk()
    if not tk:
        return None

    root = tk.Tk()
    root.withdraw()
    root.update_idletasks()
    selected = filedialog.askopenfilename(
        title="Select Excel Tracker File",
        filetypes=[("Excel Files", "*.xlsx *.xlsm *.xls"), ("All Files", "*.*")],
    )
    root.destroy()
    return selected or None


def unified_project_prompt(last_config: Dict[str, Any]) -> Optional[str]:
    tk, ttk, _, _ = try_import_tk()
    if not tk:
        return None

    root = tk.Tk()
    root.title("Load Previous Project")
    root.geometry("500x240")
    try:
        root.eval("tk::PlaceWindow . center")
    except Exception:
        pass

    result = {"choice": None}
    project_name = last_config.get("title", "Unknown")
    last_date = last_config.get("last_updated", "?")

    label = tk.Label(
        root,
        text=(
            "Would you like to load the most recently used project?\n\n"
            f"Project: {project_name}\n"
            f"Last Accessed: {last_date}"
        ),
        padx=20,
        pady=20,
        justify="left",
    )
    label.pack()

    def choose(choice: str) -> None:
        result["choice"] = choice
        root.quit()

    frame = ttk.Frame(root)
    frame.pack(pady=10)

    ttk.Button(frame, text="Yes", width=15, command=lambda: choose("yes")).grid(row=0, column=0, padx=5)
    ttk.Button(frame, text="Choose Another", width=15, command=lambda: choose("choose")).grid(row=0, column=1, padx=5)
    ttk.Button(frame, text="Create New", width=15, command=lambda: choose("new")).grid(row=0, column=2, padx=5)
    ttk.Button(frame, text="Cancel", width=15, command=lambda: choose("cancel")).grid(row=1, column=1, pady=10)

    root.mainloop()
    root.destroy()
    return result["choice"]


def choose_config_gui() -> Optional[str]:
    tk, ttk, _, messagebox = try_import_tk()
    if not tk:
        return None

    configs = list_saved_configs()
    if not configs:
        if messagebox:
            messagebox.showinfo("No Saved Projects", "No saved project configurations were found.")
        return None

    root = tk.Tk()
    root.title("Choose a Saved Project")

    tk.Label(root, text="Choose a saved project configuration:").pack(padx=10, pady=10)
    listbox = tk.Listbox(root, width=110, height=10)
    display_map = []

    for path in configs:
        conf = json.loads(path.read_text(encoding="utf-8"))
        title = conf.get("title", "Unknown Title")
        start = conf.get("start_date", "?")
        weeks = conf.get("num_weeks", "?")
        updated = conf.get("last_updated", "?")
        display = f"{title} | Start: {start} | Weeks: {weeks} | Updated: {updated}"
        display_map.append((path.stem, display))
        listbox.insert(tk.END, display)

    listbox.pack(padx=10, pady=5)

    selected = {"project_id": None}

    def select() -> None:
        selection = listbox.curselection()
        if selection:
            selected["project_id"] = display_map[selection[0]][0]
            root.quit()

    ttk.Button(root, text="Load Selected Project", command=select).pack(pady=10)
    root.mainloop()
    root.destroy()

    return selected["project_id"]


def get_project_params_gui(existing: Optional[Dict[str, Any]], excel_path: str) -> Dict[str, Any]:
    tk, ttk, _, _ = try_import_tk()
    if not tk:
        raise RuntimeError("Tkinter is not available, so the GUI parameter prompt cannot be shown.")

    result = {
        "submitted": False,
        "title": "",
        "start_date": "",
        "num_weeks": 8,
    }

    param_root = tk.Tk()
    param_root.title("Training Project Parameters")

    tk.Label(param_root, text="Project Title:").grid(row=0, column=0, sticky="w", padx=8, pady=5)
    tk.Label(param_root, text="Start Date (YYYY-MM-DD):").grid(row=1, column=0, sticky="w", padx=8, pady=5)
    tk.Label(param_root, text="Number of Weeks:").grid(row=2, column=0, sticky="w", padx=8, pady=5)

    title_entry = ttk.Entry(param_root, width=40)
    date_entry = ttk.Entry(param_root, width=20)
    weeks_entry = ttk.Entry(param_root, width=20)

    if existing:
        title_entry.insert(0, existing.get("title", ""))
        date_entry.insert(0, existing.get("start_date", ""))
        weeks_entry.insert(0, str(existing.get("num_weeks", 8)))
    else:
        title_entry.insert(0, Path(excel_path).stem.replace("_", " "))
        date_entry.insert(0, str(dt.date.today()))
        weeks_entry.insert(0, "8")

    title_entry.grid(row=0, column=1, padx=8, pady=5)
    date_entry.grid(row=1, column=1, padx=8, pady=5)
    weeks_entry.grid(row=2, column=1, padx=8, pady=5)

    def submit() -> None:
        result["submitted"] = True
        result["title"] = title_entry.get().strip() or "Training Project Timeline"
        result["start_date"] = date_entry.get().strip() or str(dt.date.today())

        try:
            result["num_weeks"] = int(weeks_entry.get().strip() or "8")
        except ValueError as exc:
            raise ValueError("Number of Weeks must be a whole number.") from exc

        param_root.quit()

    ttk.Button(param_root, text="Generate Timeline", command=submit).grid(row=3, column=0, columnspan=2, pady=12)

    param_root.mainloop()
    param_root.destroy()

    if not result["submitted"]:
        raise SystemExit("Cancelled by user.")

    return {
        "title": result["title"],
        "start_date": result["start_date"],
        "num_weeks": result["num_weeks"],
        "file_path": excel_path,
        "last_updated": str(dt.date.today()),
    }


def resolve_runtime_config(args) -> Dict[str, Any]:
    excel_path: Optional[str] = None
    config: Optional[Dict[str, Any]] = None

    if args.headless:
        if not args.excel:
            raise SystemExit("--headless requires --excel /path/to/tracker.xlsx")
        excel_path = str(Path(args.excel).expanduser().resolve())

        config_source = None
        if args.config:
            config_source = Path(args.config).expanduser().resolve()
        else:
            sibling_config = Path(excel_path).with_name("project_config.json")
            if sibling_config.exists():
                config_source = sibling_config

        if config_source and config_source.exists():
            config = json.loads(config_source.read_text(encoding="utf-8"))

        config = config or {}
        config["title"] = args.title or config.get("title") or Path(excel_path).stem.replace("_", " ")
        config["start_date"] = args.start_date or config.get("start_date") or str(dt.date.today())
        config["num_weeks"] = int(args.num_weeks or config.get("num_weeks") or 8)
        config["file_path"] = excel_path
        config["last_updated"] = str(dt.date.today())

        project_id = get_project_id(config["title"], excel_path)
        save_project_config(project_id, config)
        return config

    last_used = load_last_used()
    if last_used and Path(last_used.get("file_path", "")).exists():
        last_project_config = load_project_config(last_used["last_project"])
        if last_project_config:
            user_choice = unified_project_prompt(last_project_config)

            if user_choice == "yes":
                config = last_project_config
                excel_path = config["file_path"]

            elif user_choice == "choose":
                selected_project_id = choose_config_gui()
                if selected_project_id:
                    config = load_project_config(selected_project_id)
                    if config:
                        excel_path = config["file_path"]

            elif user_choice == "new":
                pass
            else:
                raise SystemExit("Cancelled by user.")

    if not excel_path:
        excel_path = choose_file_gui()
        if not excel_path:
            if args.excel:
                excel_path = str(Path(args.excel).expanduser().resolve())
            else:
                raise SystemExit("No Excel file selected.")

    excel_path = str(Path(excel_path).expanduser().resolve())

    if config is None:
        config = get_project_params_gui(existing=None, excel_path=excel_path)

    config["file_path"] = excel_path
    config["last_updated"] = str(dt.date.today())

    project_id = get_project_id(config["title"], excel_path)
    save_project_config(project_id, config)
    return config


def load_tracker(excel_path: str) -> pd.DataFrame:
    try:
        df = pd.read_excel(excel_path, engine="openpyxl")
    except Exception as exc:
        raise RuntimeError(f"Failed to load Excel file: {exc}") from exc

    df.columns = [str(col).strip() for col in df.columns]
    missing = REQUIRED_COLUMNS - set(df.columns)
    if missing:
        raise ValueError(f"Missing required columns: {sorted(missing)}")

    for col in ("Start Date", "End Date"):
        df[col] = pd.to_datetime(df[col], errors="coerce")

    invalid_dates = df[df["Start Date"].isna() | df["End Date"].isna()]
    if not invalid_dates.empty:
        sample_tasks = invalid_dates["Task"].head(5).astype(str).tolist()
        raise ValueError(
            "One or more rows have invalid Start Date or End Date values. "
            f"Sample tasks: {sample_tasks}"
        )

    df["Task"] = df["Task"].astype(str).fillna("").str.strip()
    df = df[df["Task"] != ""].copy()

    if "Task Order" in df.columns:
        df["Task Order"] = pd.to_numeric(df["Task Order"], errors="coerce").fillna(999999).astype(int)
    else:
        df["Task Order"] = range(len(df))

    if "Sort Key" in df.columns:
        df["Sort Key"] = pd.to_numeric(df["Sort Key"], errors="coerce").fillna(df["Task Order"]).astype(int)
    else:
        df["Sort Key"] = df["Task Order"]

    if "Is Parent" in df.columns:
        df["Is Parent"] = df["Is Parent"].apply(normalize_bool)
    else:
        df["Is Parent"] = False

    today = pd.Timestamp(dt.date.today()).normalize()
    df["Status"] = df["Status"].astype(str).str.strip()
    overdue_mask = (df["Status"].str.lower() != "complete") & (df["End Date"] < today)
    df.loc[overdue_mask, "Status"] = "Overdue"

    return df


def build_display_labels(df: pd.DataFrame) -> pd.DataFrame:
    df = df.copy()
    df = df.sort_values(by=["Sort Key", "Task Order", "Start Date", "End Date", "Task"], kind="stable")

    if "DisplayLabel" in df.columns and df["DisplayLabel"].notna().any():
        base_labels = df["DisplayLabel"].astype(str).str.strip()
        use_existing = base_labels != ""
    else:
        base_labels = pd.Series([""] * len(df), index=df.index)
        use_existing = pd.Series([False] * len(df), index=df.index)

    generated = df.apply(
        lambda row: f"{int(row['Sort Key']):04d} - {row['Task']}",
        axis=1,
    )

    labels = base_labels.where(use_existing, generated)
    labels = labels.where(~df["Is Parent"], "◆ " + labels)
    df["DisplayLabel"] = labels

    return df


def build_figure(df: pd.DataFrame, config: Dict[str, Any]):
    display_title = str(config["title"]).replace("_", " ").replace("-", " ")

    df = build_display_labels(df)

    fig = px.timeline(
        df,
        x_start="Start Date",
        x_end="End Date",
        y="DisplayLabel",
        color="Status",
        color_discrete_map=STATUS_COLORS,
        title=display_title,
        hover_data={
            "Task": True,
            "Start Date": True,
            "End Date": True,
            "Status": True,
            "Parent Slide Deck": True if "Parent Slide Deck" in df.columns else False,
            "Task Order": False,
            "Sort Key": False,
            "Is Parent": False,
            "DisplayLabel": False,
        },
    )

    task_order = df["DisplayLabel"].tolist()
    fig.update_yaxes(
        categoryorder="array",
        categoryarray=task_order,
        autorange="reversed",
        tickfont=dict(size=11),
        title=None,
    )

    fig.update_layout(
        height=max(900, 90 + (len(df) * 28)),
        margin=dict(l=80, r=20, t=60, b=120),
        showlegend=True,
        plot_bgcolor="white",
        paper_bgcolor="white",
        xaxis_title=None,
        legend_title_text="Status",
    )

    start_date = pd.to_datetime(config["start_date"])
    num_weeks = int(config["num_weeks"])
    for i in range(num_weeks):
        week_date = start_date + pd.Timedelta(weeks=i)
        fig.add_vline(x=week_date, line=dict(color="lightgray", dash="dash"))
        fig.add_annotation(
            x=week_date,
            y=1,
            yref="paper",
            showarrow=False,
            text=f"Week {i + 1}",
            font=dict(size=10),
            xanchor="left",
            yanchor="bottom",
        )

    today = pd.Timestamp(dt.date.today()).normalize()
    fig.add_vline(x=today, line=dict(color="darkred", width=2))
    fig.add_annotation(
        x=today,
        y=1,
        yref="paper",
        showarrow=False,
        text=f"Last Updated: {config['last_updated']}",
        font=dict(size=16, color="darkred"),
        xanchor="left",
        yanchor="top",
    )

    return fig


def legend_html(last_updated: str) -> str:
    return f"""
<p style='margin-top: 30px; font-family: sans-serif; font-size: 16px;'><em>Last Updated: {last_updated}</em></p>
<div style='margin-top: 30px; font-family: sans-serif; font-size: 14px;'>
  <h3>Status Legend</h3>
  <table style='border-collapse: collapse; width: 80%;'>
    <thead>
      <tr>
        <th style='text-align: left; padding: 6px;'>Status</th>
        <th style='text-align: left; padding: 6px;'>Definition</th>
        <th style='text-align: left; padding: 6px;'>Color</th>
      </tr>
    </thead>
    <tbody>
      <tr>
        <td style='padding: 6px;'>✅ <strong>Complete</strong></td>
        <td style='padding: 6px;'>Task is fully done and signed off</td>
        <td style='padding: 6px; color: green;'>Green</td>
      </tr>
      <tr>
        <td style='padding: 6px;'>🔠 <strong>In Progress</strong></td>
        <td style='padding: 6px;'>Task is actively being worked on</td>
        <td style='padding: 6px; color: orange;'>Orange</td>
      </tr>
      <tr>
        <td style='padding: 6px;'>⏳ <strong>Pending</strong></td>
        <td style='padding: 6px;'>Task is ready to start as soon as time is available</td>
        <td style='padding: 6px; color: lightgray;'>Light Gray</td>
      </tr>
      <tr>
        <td style='padding: 6px;'>🔵 <strong>Upcoming</strong></td>
        <td style='padding: 6px;'>Task is scheduled for a later date</td>
        <td style='padding: 6px; color: blue;'>Blue</td>
      </tr>
      <tr>
        <td style='padding: 6px;'>⚠️ <strong>Overdue</strong></td>
        <td style='padding: 6px;'>Task should be complete but is past its due date</td>
        <td style='padding: 6px; color: red;'>Red</td>
      </tr>
    </tbody>
  </table>
</div>
"""


def write_output(fig, config: Dict[str, Any]) -> Path:
    excel_path = Path(config["file_path"])
    safe_title = sanitize_filename(config["title"])
    filename = f"{safe_title}_Project-Tracker.html"
    output_path = excel_path.parent / filename

    html_body = pio.to_html(fig, full_html=False, include_plotlyjs="cdn", config={"displaylogo": False})
    full_html = f"""<!DOCTYPE html>
<html lang="en">
<head>
  <meta charset="utf-8">
  <title>{config['title']}</title>
  <meta name="viewport" content="width=device-width, initial-scale=1">
</head>
<body>
  {html_body}
  {legend_html(config['last_updated'])}
</body>
</html>
"""

    output_path.write_text(full_html, encoding="utf-8")
    return output_path


def parse_args():
    parser = argparse.ArgumentParser(description="Generate a stakeholder timeline HTML view from an Excel tracker.")
    parser.add_argument("--headless", action="store_true", help="Run without GUI prompts.")
    parser.add_argument("--excel", help="Path to the Excel tracker file.")
    parser.add_argument("--config", help="Optional path to a JSON config file.")
    parser.add_argument("--title", help="Project title override.")
    parser.add_argument("--start-date", help="Start date override in YYYY-MM-DD format.")
    parser.add_argument("--num-weeks", type=int, help="Number of weeks to display.")
    parser.add_argument("--no-open", action="store_true", help="Do not auto-open the generated HTML file.")
    return parser.parse_args()


def main():
    args = parse_args()

    try:
        config = resolve_runtime_config(args)
        df = load_tracker(config["file_path"])
        config["last_updated"] = str(dt.date.today())

        project_id = get_project_id(config["title"], config["file_path"])
        save_project_config(project_id, config)

        fig = build_figure(df, config)
        output_path = write_output(fig, config)

        print(f"\n✅ Timeline generated successfully: {output_path}")
        if not args.no_open:
            open_file(output_path)

        _, _, _, messagebox = try_import_tk()
        if not args.headless and messagebox:
            messagebox.showinfo("Success", f"Timeline chart has been updated successfully!\n\n{output_path}")

    except subprocess.CalledProcessError as exc:
        print(f"❌ Failed to install one or more dependencies: {exc}")
        raise SystemExit(1) from exc
    except Exception as exc:
        print(f"❌ {exc}")
        raise SystemExit(1) from exc


if __name__ == "__main__":
    main()
