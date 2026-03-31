import json
import os
import re
import shutil
import threading
import time
from pathlib import Path
from datetime import datetime
from urllib import request, error
import tkinter as tk
from tkinter import filedialog, scrolledtext
from tkinter import ttk

SEASON_PATTERN = re.compile(r"^season\s*(\d+)$", re.IGNORECASE)
DEFAULT_PATH = r"H:\media\tv\anime"
DEFAULT_SONARR_URL = "http://localhost:8989"

APP_NAME = "Season Folder Normalizer"
SETTINGS_FILE_NAME = "season_folder_normalizer_settings.json"


def get_script_directory():
    if "__file__" in globals():
        return Path(__file__).resolve().parent
    return Path.cwd()


def get_settings_file_path():
    return get_script_directory() / SETTINGS_FILE_NAME


def padded_name(season_num):
    return f"Season {season_num:02d}"


def is_padded(name, season_num):
    return name.lower() == padded_name(season_num).lower()


def get_downloads_folder():
    return str(Path.home() / "Downloads")


def timestamp():
    return datetime.now().strftime("%Y-%m-%d_%H-%M-%S")


def normalize_fs_path(path):
    return os.path.normcase(os.path.normpath(path or ""))


def load_settings_file():
    settings_path = get_settings_file_path()
    try:
        if settings_path.exists():
            with open(settings_path, "r", encoding="utf-8") as f:
                return json.load(f)
    except Exception:
        pass
    return {}


def save_settings_file(data):
    settings_path = get_settings_file_path()
    with open(settings_path, "w", encoding="utf-8") as f:
        json.dump(data, f, indent=2)


class Logger:
    def __init__(self, output_callback, log_path):
        self.output_callback = output_callback
        self.log_path = log_path
        self.lines = []

    def log(self, message="", color="default"):
        self.lines.append(message)
        self.output_callback(message, color)

    def info(self, message=""):
        self.log(message, "default")

    def header(self, message=""):
        self.log(message, "header")

    def success(self, message=""):
        self.log(message, "success")

    def dryrun(self, message=""):
        self.log(message, "dryrun")

    def warning(self, message=""):
        self.log(message, "warning")

    def summary(self, message=""):
        self.log(message, "summary")

    def save(self):
        with open(self.log_path, "w", encoding="utf-8") as f:
            f.write("\n".join(self.lines))


def scan_season_folders(show_path):
    seasons = {}

    try:
        for item in os.listdir(show_path):
            item_path = os.path.join(show_path, item)
            if not os.path.isdir(item_path):
                continue

            match = SEASON_PATTERN.match(item.strip())
            if not match:
                continue

            season_num = int(match.group(1))
            seasons.setdefault(season_num, []).append(item)
    except Exception as e:
        raise RuntimeError(f"Failed to scan '{show_path}': {e}")

    return seasons


def is_show_folder(path):
    try:
        return bool(scan_season_folders(path))
    except Exception:
        return False


def find_show_folders(root_path, recursive=False):
    if not recursive:
        return [root_path] if is_show_folder(root_path) else []

    results = []
    for current_root, dirnames, _ in os.walk(root_path):
        if is_show_folder(current_root):
            results.append(current_root)
            dirnames[:] = []
    return sorted(results, key=str.lower)


def analyze_show(show_path):
    seasons = scan_season_folders(show_path)

    duplicates = []
    naming_issues = []
    already_correct = []

    for season_num in sorted(seasons):
        folders = sorted(seasons[season_num], key=str.lower)
        correct_folder = None
        non_matching_folders = []

        for folder in folders:
            if is_padded(folder, season_num):
                correct_folder = folder
            else:
                non_matching_folders.append(folder)

        if correct_folder and non_matching_folders:
            duplicates.append({
                "season_num": season_num,
                "correct_folder": correct_folder,
                "non_matching_folders": non_matching_folders,
                "all_folders": folders,
            })
        elif not correct_folder and len(non_matching_folders) == 1:
            naming_issues.append({
                "season_num": season_num,
                "current_folder": non_matching_folders[0],
                "desired_folder": padded_name(season_num),
            })
        elif correct_folder and not non_matching_folders:
            already_correct.append({
                "season_num": season_num,
                "folder": correct_folder,
            })

    return {
        "show_path": show_path,
        "show_name": os.path.basename(show_path.rstrip("\\/")),
        "duplicates": duplicates,
        "naming_issues": naming_issues,
        "already_correct": already_correct,
    }


def move_contents(source_path, target_path, logger, dry_run):
    for item in sorted(os.listdir(source_path), key=str.lower):
        src_item = os.path.join(source_path, item)
        dst_item = os.path.join(target_path, item)

        if os.path.exists(dst_item):
            logger.warning(f"    Skipping (already exists): {item}")
            continue

        if dry_run:
            logger.dryrun(f"    [PREVIEW] Would move: {item}")
        else:
            shutil.move(src_item, dst_item)
            logger.success(f"    Moved: {item}")


def remove_if_empty(folder_path, folder_name, logger, dry_run):
    if dry_run:
        logger.dryrun(f"  [PREVIEW] Would remove folder after moving its contents: {folder_name}")
        return

    remaining = os.listdir(folder_path)
    if not remaining:
        os.rmdir(folder_path)
        logger.success(f"  Removed empty folder: {folder_name}")
    else:
        logger.warning(f"  Folder not empty after move, so it was not deleted: {folder_name}")


def process_show(show_info, logger, dry_run, fix_duplicates, rename_to_scheme):
    show_path = show_info["show_path"]
    show_name = show_info["show_name"]

    stats = {
        "shows_processed": 1,
        "duplicates_found": len(show_info["duplicates"]),
        "naming_issues_found": len(show_info["naming_issues"]),
        "renamed": 0,
        "merged": 0,
        "skipped": 0,
        "errors": 0,
    }

    logger.info()
    logger.header("=" * 80)
    logger.header(f"Show: {show_name}")
    logger.info(f"Path: {show_path}")

    if not show_info["duplicates"] and not show_info["naming_issues"]:
        logger.success("  No duplicate season folders or naming issues found.")
        return stats

    if show_info["duplicates"]:
        logger.header("Duplicate season folders found:")
        for dup in show_info["duplicates"]:
            logger.warning(f"  Season {dup['season_num']}:")
            for folder in dup["all_folders"]:
                logger.info(f"    Found: {folder}")

    if show_info["naming_issues"]:
        logger.header("Season folders that do not match the naming format:")
        for item in show_info["naming_issues"]:
            logger.info(f"  {item['current_folder']} -> {item['desired_folder']}")

    if fix_duplicates:
        for dup in show_info["duplicates"]:
            season_num = dup["season_num"]
            target = dup["correct_folder"]
            target_path = os.path.join(show_path, target)

            logger.info()
            logger.header(f"Fixing duplicate folders for Season {season_num}:")
            for source in dup["non_matching_folders"]:
                source_path = os.path.join(show_path, source)
                logger.info(f"  Moving from '{source}' -> '{target}'")
                move_contents(source_path, target_path, logger, dry_run)
                remove_if_empty(source_path, source, logger, dry_run)
                stats["merged"] += 1

    if rename_to_scheme:
        for item in show_info["naming_issues"]:
            source = item["current_folder"]
            desired = item["desired_folder"]
            source_path = os.path.join(show_path, source)
            target_path = os.path.join(show_path, desired)

            logger.info()
            logger.header(f"Updating Season {item['season_num']} to match the naming format:")
            if os.path.exists(target_path):
                logger.warning(
                    f"  Cannot rename '{source}' -> '{desired}' because the target folder already exists"
                )
                stats["skipped"] += 1
            else:
                if dry_run:
                    logger.dryrun(f"  [PREVIEW] Would rename: '{source}' -> '{desired}'")
                else:
                    os.rename(source_path, target_path)
                    logger.success(f"  Renamed: '{source}' -> '{desired}'")
                stats["renamed"] += 1

    return stats


def combine_stats(all_stats):
    total = {
        "shows_processed": 0,
        "duplicates_found": 0,
        "naming_issues_found": 0,
        "renamed": 0,
        "merged": 0,
        "skipped": 0,
        "errors": 0,
    }

    for stats in all_stats:
        for key in total:
            total[key] += stats.get(key, 0)

    return total


class SonarrClient:
    TERMINAL_COMMAND_STATES = {"completed", "failed", "aborted", "cancelled", "orphaned"}

    def __init__(self, base_url, api_key, timeout=20):
        self.base_url = base_url.rstrip("/")
        self.api_key = api_key.strip()
        self.timeout = timeout

    def _request(self, method, endpoint, payload=None):
        url = f"{self.base_url}/api/v3/{endpoint.lstrip('/')}"
        headers = {
            "X-Api-Key": self.api_key,
            "Accept": "application/json",
        }

        data = None
        if payload is not None:
            headers["Content-Type"] = "application/json"
            data = json.dumps(payload).encode("utf-8")

        req = request.Request(url, data=data, headers=headers, method=method.upper())

        try:
            with request.urlopen(req, timeout=self.timeout) as response:
                raw = response.read()
                if not raw:
                    return None
                return json.loads(raw.decode("utf-8"))
        except error.HTTPError as e:
            detail = ""
            try:
                body = e.read().decode("utf-8", errors="replace")
                detail = f" | {body[:500]}"
            except Exception:
                pass
            raise RuntimeError(f"Sonarr HTTP {e.code}: {e.reason}{detail}")
        except error.URLError as e:
            raise RuntimeError(f"Could not reach Sonarr: {e.reason}")
        except TimeoutError:
            raise RuntimeError("Timed out while contacting Sonarr")
        except Exception as e:
            raise RuntimeError(f"Unexpected Sonarr error: {e}")

    def get_system_status(self):
        return self._request("GET", "system/status")

    def get_series(self):
        return self._request("GET", "series") or []

    def get_command(self, command_id):
        return self._request("GET", f"command/{command_id}")

    def run_command(self, name, series_id=None):
        payload = {"name": name}
        if series_id is not None:
            payload["seriesId"] = series_id
        return self._request("POST", "command", payload)

    def wait_for_command(self, command_id, timeout_seconds=20, poll_interval=1.0):
        deadline = time.time() + timeout_seconds
        last_status = None
        last_payload = None

        while time.time() < deadline:
            payload = self.get_command(command_id)
            if not isinstance(payload, dict):
                return {"status": "unknown", "payload": payload}

            status = str(payload.get("status", "unknown")).lower()
            last_status = status
            last_payload = payload

            if status in self.TERMINAL_COMMAND_STATES:
                return {"status": status, "payload": payload}

            time.sleep(poll_interval)

        return {"status": last_status or "timeout", "payload": last_payload}

    def build_series_path_map(self):
        result = {}
        for item in self.get_series():
            path = item.get("path") or ""
            if path:
                result[normalize_fs_path(path)] = item
        return result


class ScrollablePanel(ttk.Frame):
    def __init__(self, parent, canvas_bg):
        super().__init__(parent)

        self.canvas = tk.Canvas(self, highlightthickness=0, bd=0, bg=canvas_bg)
        self.scrollbar = ttk.Scrollbar(self, orient="vertical", command=self.canvas.yview)
        self.content = ttk.Frame(self.canvas)

        self.content.bind(
            "<Configure>",
            lambda e: self.canvas.configure(scrollregion=self.canvas.bbox("all"))
        )

        self.window_id = self.canvas.create_window((0, 0), window=self.content, anchor="nw")
        self.canvas.configure(yscrollcommand=self.scrollbar.set)

        self.canvas.pack(side="left", fill="both", expand=True)
        self.scrollbar.pack(side="right", fill="y")

        self.canvas.bind("<Configure>", self._on_canvas_configure)
        self.canvas.bind_all("<MouseWheel>", self._on_mousewheel)

    def _on_canvas_configure(self, event):
        self.canvas.itemconfig(self.window_id, width=event.width)

    def _on_mousewheel(self, event):
        try:
            widget = self.winfo_containing(event.x_root, event.y_root)
            if widget and (widget == self.canvas or str(widget).startswith(str(self.canvas))):
                self.canvas.yview_scroll(int(-1 * (event.delta / 120)), "units")
        except Exception:
            pass

    def set_canvas_colors(self, bg):
        self.canvas.configure(bg=bg)


class App:
    def __init__(self, root):
        self.root = root
        self.root.title(APP_NAME)
        self.root.geometry("1480x960")
        self.root.minsize(1240, 820)

        self.path_var = tk.StringVar(value=DEFAULT_PATH)
        self.recursive_var = tk.BooleanVar(value=True)
        self.dry_run_var = tk.BooleanVar(value=True)
        self.fix_duplicates_var = tk.BooleanVar(value=True)
        self.rename_to_scheme_var = tk.BooleanVar(value=True)
        self.dark_mode_var = tk.BooleanVar(value=True)

        self.sonarr_enabled_var = tk.BooleanVar(value=False)
        self.sonarr_url_var = tk.StringVar(value=DEFAULT_SONARR_URL)
        self.sonarr_api_key_var = tk.StringVar(value="")
        self.sonarr_path_local_var = tk.StringVar(value=r"H:\media\tv\anime")
        self.sonarr_path_remote_var = tk.StringVar(value="/data/media/tv/anime")

        self.settings_state_var = tk.StringVar(value="Settings loaded")
        self.sonarr_last_verified_var = tk.StringVar(value="")
        self.activity_var = tk.StringVar(value="Waiting to start")
        self.summary_shows_var = tk.StringVar(value="0")
        self.summary_dupes_var = tk.StringVar(value="0")
        self.summary_names_var = tk.StringVar(value="0")
        self.summary_actions_var = tk.StringVar(value="0")

        self.prompt_event = None
        self.prompt_result = None
        self.run_in_progress = False
        self.banner_kind = "info"
        self.settings_dirty = False
        self._suspend_dirty_tracking = True

        self.style = ttk.Style()
        self.configure_theme()
        self.load_saved_settings()
        self.build_ui()
        self.bind_setting_watchers()
        self._suspend_dirty_tracking = False
        self.refresh_settings_state()

        self.root.protocol("WM_DELETE_WINDOW", self.on_close)

    def get_settings_payload(self):
        return {
            "app": {
                "path": self.path_var.get().strip(),
                "recursive": self.recursive_var.get(),
                "dry_run": self.dry_run_var.get(),
                "fix_duplicates": self.fix_duplicates_var.get(),
                "rename_to_scheme": self.rename_to_scheme_var.get(),
                "dark_mode": self.dark_mode_var.get(),
            },
            "sonarr": {
                "enabled": self.sonarr_enabled_var.get(),
                "url": self.sonarr_url_var.get().strip(),
                "api_key": self.sonarr_api_key_var.get().strip(),
                "path_local": self.sonarr_path_local_var.get().strip(),
                "path_remote": self.sonarr_path_remote_var.get().strip(),
                "last_verified": self.sonarr_last_verified_var.get().strip(),
            },
        }

    def apply_settings_payload(self, data):
        app = data.get("app", {})
        sonarr = data.get("sonarr", {})

        self.path_var.set(app.get("path", DEFAULT_PATH))
        self.recursive_var.set(app.get("recursive", True))
        self.dry_run_var.set(app.get("dry_run", True))
        self.fix_duplicates_var.set(app.get("fix_duplicates", True))
        self.rename_to_scheme_var.set(app.get("rename_to_scheme", True))
        self.dark_mode_var.set(app.get("dark_mode", True))

        self.sonarr_enabled_var.set(sonarr.get("enabled", False))
        self.sonarr_url_var.set(sonarr.get("url", DEFAULT_SONARR_URL))
        self.sonarr_api_key_var.set(sonarr.get("api_key", ""))
        self.sonarr_path_local_var.set(sonarr.get("path_local", r"H:\media\tv\anime"))
        self.sonarr_path_remote_var.set(sonarr.get("path_remote", "/data/media/tv/anime"))
        self.sonarr_last_verified_var.set(sonarr.get("last_verified", ""))

    def load_saved_settings(self):
        data = load_settings_file()
        self.apply_settings_payload(data)

    def save_all_settings(self, quiet=False):
        save_settings_file(self.get_settings_payload())
        self.settings_dirty = False
        self.refresh_settings_state()
        if not quiet:
            self.write_output(f"Saved settings to: {get_settings_file_path()}", "success")

    def reload_saved_settings(self):
        self._suspend_dirty_tracking = True
        self.load_saved_settings()
        self._suspend_dirty_tracking = False
        self.configure_theme()
        self.settings_dirty = False
        self.refresh_settings_state()
        self.write_output(f"Reloaded settings from: {get_settings_file_path()}", "default")

    def clear_sonarr_api_key(self):
        self.sonarr_api_key_var.set("")
        self.write_output("Cleared Sonarr API key from the form.", "default")

    def bind_setting_watchers(self):
        watched_vars = [
            self.path_var,
            self.recursive_var,
            self.dry_run_var,
            self.fix_duplicates_var,
            self.rename_to_scheme_var,
            self.dark_mode_var,
            self.sonarr_enabled_var,
            self.sonarr_url_var,
            self.sonarr_api_key_var,
            self.sonarr_path_local_var,
            self.sonarr_path_remote_var,
        ]
        for var in watched_vars:
            var.trace_add("write", self.on_setting_changed)

    def on_setting_changed(self, *_):
        if self._suspend_dirty_tracking:
            return
        self.settings_dirty = True
        self.refresh_settings_state()

    def refresh_settings_state(self):
        if self.settings_dirty:
            state = "Unsaved changes"
        elif self.sonarr_last_verified_var.get().strip():
            state = f"Saved • Sonarr verified {self.sonarr_last_verified_var.get().strip()}"
        else:
            state = "Saved locally"
        self.settings_state_var.set(state)

    def configure_theme(self):
        dark = self.dark_mode_var.get()

        if dark:
            self.colors = {
                "bg": "#111315",
                "panel": "#1a1d21",
                "text": "#f2f5f8",
                "muted": "#9ca8b5",
                "accent": "#4f8cff",
                "accent_text": "#ffffff",
                "border": "#303743",
                "success": "#67d58a",
                "warning": "#ff8d7a",
                "dryrun": "#f0c36a",
                "summary": "#c597ff",
                "header": "#74d8d5",
                "button": "#242a31",
                "button_active": "#2d353e",
                "entry": "#15191d",
                "output_bg": "#0d0f12",
                "activity_bg": "#20242a",
                "prompt_info_bg": "#1f314e",
                "prompt_warn_bg": "#4b2f21",
                "prompt_success_bg": "#21392a",
            }
        else:
            self.colors = {
                "bg": "#eef1f5",
                "panel": "#ffffff",
                "text": "#1f2933",
                "muted": "#66778a",
                "accent": "#2563eb",
                "accent_text": "#ffffff",
                "border": "#d8e0ea",
                "success": "#1f8a46",
                "warning": "#c8562f",
                "dryrun": "#b37400",
                "summary": "#7e49c6",
                "header": "#0f8b8d",
                "button": "#edf2f8",
                "button_active": "#dde6f1",
                "entry": "#ffffff",
                "output_bg": "#ffffff",
                "activity_bg": "#f4f7fb",
                "prompt_info_bg": "#eaf2ff",
                "prompt_warn_bg": "#fff1ea",
                "prompt_success_bg": "#ecf9f0",
            }

        self.style.theme_use("clam")
        self.root.configure(bg=self.colors["bg"])

        self.style.configure("TFrame", background=self.colors["bg"])
        self.style.configure("Card.TFrame", background=self.colors["panel"], relief="flat")
        self.style.configure("InnerCard.TFrame", background=self.colors["panel"], relief="flat")

        self.style.configure(
            "TLabel",
            background=self.colors["bg"],
            foreground=self.colors["text"],
            font=("Segoe UI", 10),
        )
        self.style.configure(
            "Hero.TLabel",
            background=self.colors["bg"],
            foreground=self.colors["text"],
            font=("Segoe UI Semibold", 20),
        )
        self.style.configure(
            "Subtle.TLabel",
            background=self.colors["bg"],
            foreground=self.colors["muted"],
            font=("Segoe UI", 9),
        )
        self.style.configure(
            "CardTitle.TLabel",
            background=self.colors["panel"],
            foreground=self.colors["text"],
            font=("Segoe UI Semibold", 11),
        )
        self.style.configure(
            "CardBody.TLabel",
            background=self.colors["panel"],
            foreground=self.colors["muted"],
            font=("Segoe UI", 9),
        )
        self.style.configure(
            "MetricValue.TLabel",
            background=self.colors["panel"],
            foreground=self.colors["text"],
            font=("Segoe UI Semibold", 20),
        )
        self.style.configure(
            "MetricLabel.TLabel",
            background=self.colors["panel"],
            foreground=self.colors["muted"],
            font=("Segoe UI", 9),
        )
        self.style.configure(
            "Activity.TLabel",
            background=self.colors["panel"],
            foreground=self.colors["muted"],
            font=("Segoe UI", 9),
        )
        self.style.configure(
            "FieldLabel.TLabel",
            background=self.colors["panel"],
            foreground=self.colors["muted"],
            font=("Segoe UI", 9),
        )

        self.style.configure(
            "TButton",
            font=("Segoe UI", 10),
            padding=(10, 7),
            relief="flat",
            background=self.colors["button"],
            foreground=self.colors["text"],
            borderwidth=0,
        )
        self.style.map(
            "TButton",
            background=[("active", self.colors["button_active"])],
            foreground=[("active", self.colors["text"])],
        )

        self.style.configure(
            "Accent.TButton",
            background=self.colors["accent"],
            foreground=self.colors["accent_text"],
            borderwidth=0,
        )
        self.style.map(
            "Accent.TButton",
            background=[("active", self.colors["accent"])],
            foreground=[("active", self.colors["accent_text"])],
        )

        self.style.configure(
            "TCheckbutton",
            background=self.colors["panel"],
            foreground=self.colors["text"],
            font=("Segoe UI", 10),
        )
        self.style.map(
            "TCheckbutton",
            background=[("active", self.colors["panel"])],
            foreground=[("active", self.colors["text"])],
        )

        self.style.configure(
            "TEntry",
            fieldbackground=self.colors["entry"],
            foreground=self.colors["text"],
            bordercolor=self.colors["border"],
            lightcolor=self.colors["border"],
            darkcolor=self.colors["border"],
            insertcolor=self.colors["text"],
            padding=7,
        )

        if hasattr(self, "controls_panel"):
            self.controls_panel.set_canvas_colors(self.colors["bg"])

        if hasattr(self, "output"):
            self.output.configure(
                bg=self.colors["output_bg"],
                fg=self.colors["text"],
                insertbackground=self.colors["text"],
                selectbackground=self.colors["accent"],
                highlightbackground=self.colors["border"],
                highlightcolor=self.colors["accent"],
            )
            self.output.tag_config("default", foreground=self.colors["text"])
            self.output.tag_config("header", foreground=self.colors["header"])
            self.output.tag_config("success", foreground=self.colors["success"])
            self.output.tag_config("dryrun", foreground=self.colors["dryrun"])
            self.output.tag_config("warning", foreground=self.colors["warning"])
            self.output.tag_config("summary", foreground=self.colors["summary"])

        if hasattr(self, "banner_frame"):
            self._apply_banner_style(self.banner_kind)

        if hasattr(self, "activity_chip"):
            self.activity_chip.configure(
                bg=self.colors["activity_bg"],
                fg=self.colors["text"],
                highlightbackground=self.colors["border"],
                highlightcolor=self.colors["border"],
            )

    def toggle_theme(self):
        self.configure_theme()

    def build_metric_card(self, parent, label_text, variable):
        card = ttk.Frame(parent, style="Card.TFrame", padding=12)
        card.pack(side="left", fill="both", expand=True, padx=4)
        ttk.Label(card, textvariable=variable, style="MetricValue.TLabel").pack(anchor="w")
        ttk.Label(card, text=label_text, style="MetricLabel.TLabel").pack(anchor="w", pady=(2, 0))

    def build_compact_option(self, parent, row, col, text, variable, description):
        cell = ttk.Frame(parent, style="InnerCard.TFrame")
        cell.grid(row=row, column=col, sticky="nsew", padx=(0 if col == 0 else 8, 0), pady=(0, 8))
        ttk.Checkbutton(cell, text=text, variable=variable).pack(anchor="w")
        ttk.Label(
            cell,
            text=description,
            style="CardBody.TLabel",
            wraplength=250,
            justify="left",
        ).pack(anchor="w", padx=(24, 0), pady=(2, 0))

    def build_ui(self):
        outer = ttk.Frame(self.root, padding=16)
        outer.pack(fill="both", expand=True)

        header = ttk.Frame(outer)
        header.pack(fill="x", pady=(0, 12))

        header_left = ttk.Frame(header)
        header_left.pack(side="left", fill="x", expand=True)

        ttk.Label(header_left, text=APP_NAME, style="Hero.TLabel").pack(anchor="w")
        ttk.Label(
            header_left,
            text="Scan TV and anime folders, preview changes, merge duplicate season folders, standardise season naming, and optionally refresh Sonarr afterwards.",
            style="Subtle.TLabel",
        ).pack(anchor="w", pady=(3, 0))

        header_right = ttk.Frame(header)
        header_right.pack(side="right", anchor="n")

        ttk.Checkbutton(
            header_right,
            text="Dark mode",
            variable=self.dark_mode_var,
            command=self.toggle_theme,
        ).pack(anchor="e")

        summary_row = ttk.Frame(outer)
        summary_row.pack(fill="x", pady=(0, 12))

        self.build_metric_card(summary_row, "Shows with issues", self.summary_shows_var)
        self.build_metric_card(summary_row, "Duplicate groups", self.summary_dupes_var)
        self.build_metric_card(summary_row, "Naming issues", self.summary_names_var)
        self.build_metric_card(summary_row, "Changes applied", self.summary_actions_var)

        content = ttk.Frame(outer)
        content.pack(fill="both", expand=True)

        content.columnconfigure(0, weight=2)
        content.columnconfigure(1, weight=3)
        content.rowconfigure(0, weight=1)

        left_panel_wrap = ttk.Frame(content)
        left_panel_wrap.grid(row=0, column=0, sticky="nsew", padx=(0, 10))

        right_panel = ttk.Frame(content)
        right_panel.grid(row=0, column=1, sticky="nsew")

        self.controls_panel = ScrollablePanel(left_panel_wrap, self.colors["bg"])
        self.controls_panel.pack(fill="both", expand=True)

        self.build_controls_panel(self.controls_panel.content)
        self.build_output_panel(right_panel)

        self.configure_theme()

    def build_controls_panel(self, parent):
        path_card = ttk.Frame(parent, style="Card.TFrame", padding=14)
        path_card.pack(fill="x", pady=(0, 10))

        ttk.Label(path_card, text="Media folder", style="CardTitle.TLabel").pack(anchor="w")
        ttk.Label(
            path_card,
            text="Choose a single show folder or a higher-level library folder.",
            style="CardBody.TLabel",
        ).pack(anchor="w", pady=(2, 8))

        path_row = ttk.Frame(path_card, style="InnerCard.TFrame")
        path_row.pack(fill="x")

        self.path_entry = ttk.Entry(path_row, textvariable=self.path_var)
        self.path_entry.pack(side="left", fill="x", expand=True)

        ttk.Button(path_row, text="Browse...", command=self.browse_folder).pack(side="left", padx=(8, 0))

        actions_card = ttk.Frame(parent, style="Card.TFrame", padding=14)
        actions_card.pack(fill="x", pady=(0, 10))

        ttk.Label(actions_card, text="Actions", style="CardTitle.TLabel").pack(anchor="w")
        ttk.Label(
            actions_card,
            text="Run a scan to review issues, or build a preview and apply changes using the selected options.",
            style="CardBody.TLabel",
        ).pack(anchor="w", pady=(2, 8))

        actions_row = ttk.Frame(actions_card, style="InnerCard.TFrame")
        actions_row.pack(fill="x")

        self.scan_button = ttk.Button(actions_row, text="Run Scan", command=self.start_scan)
        self.scan_button.pack(side="left")

        self.fix_button = ttk.Button(actions_row, text="Run Fix", command=self.start_fix, style="Accent.TButton")
        self.fix_button.pack(side="left", padx=(8, 0))

        activity_row = ttk.Frame(actions_card, style="InnerCard.TFrame")
        activity_row.pack(fill="x", pady=(10, 0))

        ttk.Label(activity_row, text="Current activity", style="Activity.TLabel").pack(anchor="w")

        self.activity_chip = tk.Label(
            activity_row,
            textvariable=self.activity_var,
            font=("Segoe UI Semibold", 9),
            anchor="w",
            justify="left",
            padx=10,
            pady=6,
            bd=1,
            relief="solid",
        )
        self.activity_chip.pack(fill="x", pady=(6, 0))

        self.banner_frame = tk.Frame(actions_card, bd=1, relief="solid", padx=12, pady=10)
        self.banner_frame.pack(fill="x", pady=(12, 0))
        self.banner_frame.pack_forget()

        self.banner_message = tk.Label(
            self.banner_frame,
            text="",
            anchor="w",
            justify="left",
            wraplength=360,
            font=("Segoe UI", 10),
        )
        self.banner_message.pack(side="left", fill="x", expand=True)

        self.banner_button_frame = ttk.Frame(self.banner_frame)
        self.banner_button_frame.pack(side="right", padx=(12, 0))

        self.banner_primary_btn = ttk.Button(
            self.banner_button_frame,
            text="Yes",
            command=lambda: self.respond_prompt(True)
        )
        self.banner_primary_btn.pack(side="left")

        self.banner_secondary_btn = ttk.Button(
            self.banner_button_frame,
            text="No",
            command=lambda: self.respond_prompt(False)
        )
        self.banner_secondary_btn.pack(side="left", padx=(8, 0))

        recent_card = ttk.Frame(parent, style="Card.TFrame", padding=14)
        recent_card.pack(fill="x", pady=(0, 10))

        ttk.Label(recent_card, text="Recent changes", style="CardTitle.TLabel").pack(anchor="w")
        ttk.Label(
            recent_card,
            text=(
                "• Moved Actions directly under Media folder\n"
                "• Renamed the right-side panel to Output log\n"
                "• Simplified Sonarr post-fix workflow to RefreshSeries then RescanSeries\n"
                "• Cleaned up settings persistence, dirty-state tracking, and save/reload UX"
            ),
            style="CardBody.TLabel",
            justify="left",
        ).pack(anchor="w", pady=(2, 0))

        scan_card = ttk.Frame(parent, style="Card.TFrame", padding=14)
        scan_card.pack(fill="x", pady=(0, 10))

        ttk.Label(scan_card, text="Scan options", style="CardTitle.TLabel").pack(anchor="w")
        ttk.Label(
            scan_card,
            text="These settings control where the app looks and how the scan behaves.",
            style="CardBody.TLabel",
        ).pack(anchor="w", pady=(2, 8))

        scan_grid = ttk.Frame(scan_card, style="InnerCard.TFrame")
        scan_grid.pack(fill="x")
        scan_grid.columnconfigure(0, weight=1)
        scan_grid.columnconfigure(1, weight=1)

        self.build_compact_option(
            scan_grid, 0, 0,
            "Recursive scanning",
            self.recursive_var,
            "Search subfolders and stop when a show folder is found."
        )
        self.build_compact_option(
            scan_grid, 0, 1,
            "Preview changes before applying them",
            self.dry_run_var,
            "Show the planned moves and renames first, then confirm inside the app."
        )

        fix_card = ttk.Frame(parent, style="Card.TFrame", padding=14)
        fix_card.pack(fill="x", pady=(0, 10))

        ttk.Label(fix_card, text="Cleanup options", style="CardTitle.TLabel").pack(anchor="w")
        ttk.Label(
            fix_card,
            text="These settings control what the app is allowed to change.",
            style="CardBody.TLabel",
        ).pack(anchor="w", pady=(2, 8))

        fix_grid = ttk.Frame(fix_card, style="InnerCard.TFrame")
        fix_grid.pack(fill="x")
        fix_grid.columnconfigure(0, weight=1)
        fix_grid.columnconfigure(1, weight=1)

        self.build_compact_option(
            fix_grid, 0, 0,
            "Merge duplicate season folders",
            self.fix_duplicates_var,
            "Move files from folders like 'Season 1' into the correctly padded folder."
        )
        self.build_compact_option(
            fix_grid, 0, 1,
            "Rename folders to match the season format",
            self.rename_to_scheme_var,
            "Rename folders like 'Season 4' to 'Season 04' when there is no conflict."
        )

        sonarr_card = ttk.Frame(parent, style="Card.TFrame", padding=14)
        sonarr_card.pack(fill="x", pady=(0, 10))

        ttk.Label(sonarr_card, text="Sonarr settings", style="CardTitle.TLabel").pack(anchor="w")
        ttk.Label(
            sonarr_card,
            text="Connect to Sonarr v4, save settings beside this script, and automatically run RefreshSeries then RescanSeries after a successful fix.",
            style="CardBody.TLabel",
        ).pack(anchor="w", pady=(2, 8))

        sonarr_toggle_row = ttk.Frame(sonarr_card, style="InnerCard.TFrame")
        sonarr_toggle_row.pack(fill="x", pady=(0, 8))

        ttk.Checkbutton(
            sonarr_toggle_row,
            text="Enable Sonarr integration",
            variable=self.sonarr_enabled_var,
        ).pack(anchor="w")

        sonarr_fields = ttk.Frame(sonarr_card, style="InnerCard.TFrame")
        sonarr_fields.pack(fill="x")
        sonarr_fields.columnconfigure(0, weight=1)
        sonarr_fields.columnconfigure(1, weight=1)

        url_cell = ttk.Frame(sonarr_fields, style="InnerCard.TFrame")
        url_cell.grid(row=0, column=0, sticky="ew", padx=(0, 8))
        ttk.Label(url_cell, text="Sonarr URL", style="FieldLabel.TLabel").pack(anchor="w", pady=(0, 4))
        self.sonarr_url_entry = ttk.Entry(url_cell, textvariable=self.sonarr_url_var)
        self.sonarr_url_entry.pack(fill="x")

        api_cell = ttk.Frame(sonarr_fields, style="InnerCard.TFrame")
        api_cell.grid(row=0, column=1, sticky="ew")
        ttk.Label(api_cell, text="API key", style="FieldLabel.TLabel").pack(anchor="w", pady=(0, 4))
        self.sonarr_api_entry = ttk.Entry(api_cell, textvariable=self.sonarr_api_key_var, show="*")
        self.sonarr_api_entry.pack(fill="x")

        mapping_fields = ttk.Frame(sonarr_card, style="InnerCard.TFrame")
        mapping_fields.pack(fill="x", pady=(8, 0))
        mapping_fields.columnconfigure(0, weight=1)
        mapping_fields.columnconfigure(1, weight=1)

        local_map_cell = ttk.Frame(mapping_fields, style="InnerCard.TFrame")
        local_map_cell.grid(row=0, column=0, sticky="ew", padx=(0, 8))
        ttk.Label(local_map_cell, text="Local library root", style="FieldLabel.TLabel").pack(anchor="w", pady=(0, 4))
        self.sonarr_local_map_entry = ttk.Entry(local_map_cell, textvariable=self.sonarr_path_local_var)
        self.sonarr_local_map_entry.pack(fill="x")

        remote_map_cell = ttk.Frame(mapping_fields, style="InnerCard.TFrame")
        remote_map_cell.grid(row=0, column=1, sticky="ew")
        ttk.Label(remote_map_cell, text="Sonarr library root", style="FieldLabel.TLabel").pack(anchor="w", pady=(0, 4))
        self.sonarr_remote_map_entry = ttk.Entry(remote_map_cell, textvariable=self.sonarr_path_remote_var)
        self.sonarr_remote_map_entry.pack(fill="x")

        ttk.Label(
            sonarr_card,
            text="The app targets Sonarr v4, which still uses the /api/v3 endpoint path.",
            style="CardBody.TLabel",
        ).pack(anchor="w", pady=(8, 6))

        status_row = ttk.Frame(sonarr_card, style="InnerCard.TFrame")
        status_row.pack(fill="x", pady=(8, 0))

        ttk.Label(status_row, text="Settings status", style="FieldLabel.TLabel").pack(anchor="w")
        ttk.Label(status_row, textvariable=self.settings_state_var, style="CardBody.TLabel").pack(anchor="w", pady=(2, 0))

        sonarr_actions = ttk.Frame(sonarr_card, style="InnerCard.TFrame")
        sonarr_actions.pack(fill="x", pady=(8, 0))

        self.test_sonarr_button = ttk.Button(
            sonarr_actions,
            text="Test and Save",
            command=self.start_test_sonarr
        )
        self.test_sonarr_button.pack(side="left")

        self.save_settings_button = ttk.Button(
            sonarr_actions,
            text="Save Settings",
            command=self.save_settings_button_clicked
        )
        self.save_settings_button.pack(side="left", padx=(8, 0))

        self.reload_sonarr_button = ttk.Button(
            sonarr_actions,
            text="Reload Saved",
            command=self.reload_saved_settings
        )
        self.reload_sonarr_button.pack(side="left", padx=(8, 0))

        self.clear_sonarr_api_button = ttk.Button(
            sonarr_actions,
            text="Clear API Key",
            command=self.clear_sonarr_api_key
        )
        self.clear_sonarr_api_button.pack(side="left", padx=(8, 0))

    def build_output_panel(self, parent):
        output_card = ttk.Frame(parent, style="Card.TFrame", padding=14)
        output_card.pack(fill="both", expand=True)

        top_row = ttk.Frame(output_card, style="InnerCard.TFrame")
        top_row.pack(fill="x")

        ttk.Label(top_row, text="Output log", style="CardTitle.TLabel").pack(side="left")
        ttk.Button(top_row, text="Clear Output", command=self.clear_output).pack(side="right")

        ttk.Label(
            output_card,
            text="Scan results, previews, file moves, renames, Sonarr activity, warnings, and the saved log path appear here.",
            style="CardBody.TLabel",
        ).pack(anchor="w", pady=(4, 10))

        self.output = scrolledtext.ScrolledText(
            output_card,
            wrap="word",
            font=("Consolas", 10),
            relief="flat",
            borderwidth=0,
            padx=12,
            pady=12,
        )
        self.output.pack(fill="both", expand=True)

    def _apply_banner_style(self, kind):
        color_map = {
            "info": self.colors["prompt_info_bg"],
            "warning": self.colors["prompt_warn_bg"],
            "success": self.colors["prompt_success_bg"],
        }
        bg = color_map.get(kind, self.colors["prompt_info_bg"])
        self.banner_frame.configure(
            bg=bg,
            highlightbackground=self.colors["border"],
            highlightcolor=self.colors["border"],
        )
        self.banner_message.configure(bg=bg, fg=self.colors["text"])

    def show_banner(self, message, kind="info", show_buttons=False, primary_text="Yes", secondary_text="No"):
        def _show():
            self.banner_kind = kind
            self._apply_banner_style(kind)
            self.banner_message.config(text=message)
            self.banner_primary_btn.config(text=primary_text)
            self.banner_secondary_btn.config(text=secondary_text)

            if show_buttons:
                self.banner_button_frame.pack(side="right", padx=(12, 0))
            else:
                self.banner_button_frame.pack_forget()

            self.banner_frame.pack(fill="x", pady=(12, 0))

        self.root.after(0, _show)

    def hide_banner(self):
        self.root.after(0, lambda: self.banner_frame.pack_forget())

    def browse_folder(self):
        selected = filedialog.askdirectory(initialdir=self.path_var.get() or DEFAULT_PATH)
        if selected:
            self.path_var.set(selected)

    def clear_output(self):
        self.output.delete("1.0", tk.END)

    def write_output(self, message, tag="default"):
        def _write():
            self.output.insert(tk.END, message + "\n", tag)
            self.output.see(tk.END)

        self.root.after(0, _write)

    def set_activity(self, text):
        self.root.after(0, lambda: self.activity_var.set(text))

    def set_summary(self, shows=0, dupes=0, names=0, actions=0):
        def _set():
            self.summary_shows_var.set(str(shows))
            self.summary_dupes_var.set(str(dupes))
            self.summary_names_var.set(str(names))
            self.summary_actions_var.set(str(actions))

        self.root.after(0, _set)

    def set_buttons_enabled(self, enabled):
        def _set():
            state = "normal" if enabled else "disabled"
            self.scan_button.config(state=state)
            self.fix_button.config(state=state)
            self.test_sonarr_button.config(state=state)
            self.save_settings_button.config(state=state)
            self.reload_sonarr_button.config(state=state)
            self.clear_sonarr_api_button.config(state=state)

        self.root.after(0, _set)

    def respond_prompt(self, result):
        self.prompt_result = result
        if self.prompt_event:
            self.prompt_event.set()
        self.hide_banner()

    def ask_inline_question(self, message, kind="info", primary_text="Yes", secondary_text="No"):
        event = threading.Event()
        self.prompt_event = event
        self.prompt_result = None
        self.show_banner(
            message,
            kind=kind,
            show_buttons=True,
            primary_text=primary_text,
            secondary_text=secondary_text,
        )
        event.wait()
        result = self.prompt_result
        self.prompt_event = None
        self.prompt_result = None
        return bool(result)

    def save_settings_button_clicked(self):
        try:
            self.save_all_settings()
        except Exception as e:
            self.write_output(f"Could not save settings: {e}", "warning")
            self.show_banner(f"Could not save settings: {e}", kind="warning")

    def validate_path(self):
        path = os.path.normpath(self.path_var.get().strip().strip('"'))
        if not path:
            self.write_output("ERROR: Please select a folder path.", "warning")
            self.show_banner("Choose a folder before starting a scan or fix run.", kind="warning")
            self.set_activity("Waiting for a valid folder")
            return None

        if not os.path.exists(path):
            self.write_output(f"ERROR: Path does not exist: {path}", "warning")
            self.show_banner("The selected folder does not exist. Please choose a valid path.", kind="warning")
            self.set_activity("Waiting for a valid folder")
            return None

        return path

    def map_local_path_to_sonarr_path(self, local_path):
        local_root = self.sonarr_path_local_var.get().strip().rstrip("\\/")
        remote_root = self.sonarr_path_remote_var.get().strip().rstrip("\\/")

        if not local_root or not remote_root:
            return local_path

        norm_local_root = normalize_fs_path(local_root)
        norm_local_path = normalize_fs_path(local_path)

        if norm_local_path == norm_local_root:
            return remote_root

        if norm_local_path.startswith(norm_local_root + os.sep):
            suffix = local_path[len(local_root):].lstrip("\\/")
            if "/" in remote_root:
                suffix = suffix.replace("\\", "/")
                return remote_root + "/" + suffix
            return remote_root + os.sep + suffix

        return local_path

    def get_sonarr_client(self):
        if not self.sonarr_enabled_var.get():
            return None

        base_url = self.sonarr_url_var.get().strip().rstrip("/")
        api_key = self.sonarr_api_key_var.get().strip()

        if not base_url or not api_key:
            raise RuntimeError("Sonarr integration is enabled, but the URL or API key is missing")

        return SonarrClient(base_url=base_url, api_key=api_key)

    def maybe_log_sonarr_config(self, logger):
        if not self.sonarr_enabled_var.get():
            logger.info("Sonarr integration: Disabled")
            return

        logger.info("Sonarr integration: Enabled")
        logger.info(f"Sonarr URL: {self.sonarr_url_var.get().strip()}")
        logger.info("Post-fix workflow: RefreshSeries -> RescanSeries")
        logger.info(f"Local path root: {self.sonarr_path_local_var.get().strip() or '(not set)'}")
        logger.info(f"Sonarr path root: {self.sonarr_path_remote_var.get().strip() or '(not set)'}")
        logger.info(f"Settings file: {get_settings_file_path()}")

    def prepare_run(self):
        path = self.validate_path()
        if not path:
            return None, None

        self.hide_banner()

        try:
            self.save_all_settings(quiet=True)
        except Exception as e:
            self.write_output(f"Warning: could not save settings before run: {e}", "warning")

        log_path = os.path.join(
            get_downloads_folder(),
            f"season_folder_normalizer_{timestamp()}.log"
        )

        logger = Logger(self.write_output, log_path)
        logger.header(APP_NAME)
        logger.header("-" * 80)
        logger.info(f"Selected path: {path}")
        logger.info(f"Recursive scanning: {'Yes' if self.recursive_var.get() else 'No'}")
        logger.info(f"Preview before applying changes: {'Yes' if self.dry_run_var.get() else 'No'}")
        logger.info(f"Merge duplicate season folders: {'Yes' if self.fix_duplicates_var.get() else 'No'}")
        logger.info(f"Rename folders to match format: {'Yes' if self.rename_to_scheme_var.get() else 'No'}")
        self.maybe_log_sonarr_config(logger)
        logger.info(f"Log file: {log_path}")
        logger.info()

        return path, logger

    def collect_show_data(self, path, logger):
        show_folders = find_show_folders(path, recursive=self.recursive_var.get())

        if not show_folders:
            logger.warning("No show folders found.")
            logger.save()
            self.set_summary(0, 0, 0, 0)
            self.show_banner("No show folders were found in the selected location.", kind="warning")
            self.set_activity("No shows found")
            return None

        logger.success(f"Found {len(show_folders)} show folder(s).")
        logger.info()

        analyzed = [analyze_show(show_path) for show_path in show_folders]
        shows_with_issues = [show for show in analyzed if show["duplicates"] or show["naming_issues"]]

        duplicate_count = sum(len(show["duplicates"]) for show in analyzed)
        naming_count = sum(len(show["naming_issues"]) for show in analyzed)

        self.set_summary(
            shows=len(shows_with_issues),
            dupes=duplicate_count,
            names=naming_count,
            actions=0,
        )

        return analyzed, shows_with_issues

    def start_scan(self):
        if self.run_in_progress:
            return
        threading.Thread(target=self.run_scan, daemon=True).start()

    def start_fix(self):
        if self.run_in_progress:
            return
        threading.Thread(target=self.run_fix, daemon=True).start()

    def start_test_sonarr(self):
        if self.run_in_progress:
            return
        threading.Thread(target=self.run_test_sonarr, daemon=True).start()

    def run_test_sonarr(self):
        self.run_in_progress = True
        self.set_buttons_enabled(False)
        self.set_activity("Testing Sonarr connection")

        try:
            client = self.get_sonarr_client()
            if client is None:
                self.show_banner("Enable Sonarr integration first, then enter the URL and API key.", kind="warning")
                self.set_activity("Waiting for Sonarr settings")
                return

            status = client.get_system_status()
            version = status.get("version", "Unknown")
            app_name = status.get("appName", "Sonarr")

            self.sonarr_last_verified_var.set(datetime.now().strftime("%Y-%m-%d %H:%M:%S"))
            self.save_all_settings(quiet=True)

            self.write_output(f"Sonarr connection successful: {app_name} {version}", "success")
            self.write_output(f"Saved settings to: {get_settings_file_path()}", "success")
            self.show_banner(f"Connected successfully to {app_name} {version}. Settings saved.", kind="success")
            self.set_activity("Sonarr connection succeeded")

        except Exception as e:
            self.write_output(f"Sonarr test failed: {e}", "warning")
            self.show_banner(f"Sonarr connection failed: {e}", kind="warning")
            self.set_activity("Sonarr connection failed")
        finally:
            self.run_in_progress = False
            self.set_buttons_enabled(True)

    def run_scan(self):
        self.run_in_progress = True
        self.set_buttons_enabled(False)
        self.set_activity("Scanning folders")

        try:
            path, logger = self.prepare_run()
            if not path:
                return

            result = self.collect_show_data(path, logger)
            if not result:
                return

            _, shows_with_issues = result

            logger.header("=" * 80)
            logger.header("Scan Results")

            if not shows_with_issues:
                logger.success("No duplicate season folders or naming issues found.")
                logger.save()
                self.show_banner("Scan finished. No season folder issues were found.", kind="success")
                self.set_activity("Scan finished")
                return

            logger.summary(f"Found {len(shows_with_issues)} show(s) with season folder issues.")
            logger.info()

            for show in sorted(shows_with_issues, key=lambda x: x["show_name"].lower()):
                logger.header(f"Show: {show['show_name']}")
                logger.info(f"Path: {show['show_path']}")

                for dup in show["duplicates"]:
                    logger.warning(f"  Duplicate season folders for Season {dup['season_num']}:")
                    for folder in dup["all_folders"]:
                        logger.info(f"    - {folder}")

                for item in show["naming_issues"]:
                    logger.info(f"  Naming issue: {item['current_folder']} -> {item['desired_folder']}")

                logger.info()

            logger.save()
            self.show_banner(
                "Scan finished. Review the results on the right, then run Fix when you are ready.",
                kind="success",
            )
            self.set_activity("Scan finished")

        except Exception as e:
            self.write_output(f"ERROR: {e}", "warning")
            self.show_banner(f"Scan failed: {e}", kind="warning")
            self.set_activity("Scan failed")
        finally:
            self.run_in_progress = False
            self.set_buttons_enabled(True)

    def queue_series_command(self, client, logger, title, series_id, command_name):
        result = client.run_command(command_name, series_id=series_id)
        command_id = result.get("id") if isinstance(result, dict) else None
        initial_status = result.get("status") if isinstance(result, dict) else None

        if command_id:
            logger.success(
                f"Sonarr queued {command_name} for '{title}'"
                f" (command id: {command_id}, status: {initial_status})"
            )

            waited = client.wait_for_command(command_id, timeout_seconds=12, poll_interval=1.0)
            final_status = waited.get("status", "unknown")
            payload = waited.get("payload") or {}
            message = payload.get("message") or ""

            if final_status == "completed":
                logger.success(f"Sonarr command {command_name} completed for '{title}'")
            elif final_status in {"queued", "started", "timeout"}:
                logger.info(f"Sonarr command {command_name} is still {final_status} for '{title}'")
            else:
                logger.warning(
                    f"Sonarr command {command_name} ended with status '{final_status}' for '{title}'"
                    + (f": {message}" if message else "")
                )
        else:
            logger.warning(f"Sonarr returned no command id for {command_name} on '{title}'")

        return result

    def notify_sonarr_after_fix(self, affected_show_paths, logger):
        if not self.sonarr_enabled_var.get():
            logger.info("Sonarr refresh skipped because Sonarr integration is disabled.")
            return None

        client = self.get_sonarr_client()
        if client is None:
            return None

        self.set_activity("Refreshing Sonarr")
        logger.info()
        logger.header("=" * 80)
        logger.header("Sonarr Refresh")
        logger.info("Using Sonarr path mapping:")
        logger.info(f"  Local root : {self.sonarr_path_local_var.get().strip() or '(not set)'}")
        logger.info(f"  Sonarr root: {self.sonarr_path_remote_var.get().strip() or '(not set)'}")
        logger.info("Workflow: RefreshSeries -> RescanSeries")

        series_map = client.build_series_path_map()

        matched = []
        unmatched = []

        for show_path in sorted(set(affected_show_paths), key=str.lower):
            mapped_path = self.map_local_path_to_sonarr_path(show_path)
            match = series_map.get(normalize_fs_path(mapped_path))

            if match:
                matched.append((show_path, mapped_path, match))
            else:
                unmatched.append((show_path, mapped_path))

        refreshed = 0
        rescanned = 0

        for local_path, sonarr_path, series in matched:
            series_id = series.get("id")
            title = series.get("title", f"Series {series_id}")

            logger.info()
            logger.info(f"Matched Sonarr series: {title}")
            logger.info(f"  Local path : {local_path}")
            logger.info(f"  Sonarr path: {sonarr_path}")

            self.queue_series_command(client, logger, title, series_id, "RefreshSeries")
            refreshed += 1

            self.queue_series_command(client, logger, title, series_id, "RescanSeries")
            rescanned += 1

        logger.summary(f"Sonarr series matched by path: {len(matched)}")
        logger.summary(f"RefreshSeries queued: {refreshed}")
        logger.summary(f"RescanSeries queued: {rescanned}")

        if unmatched:
            logger.warning("These folders did not match a Sonarr series path:")
            for local_path, sonarr_path in unmatched:
                logger.warning(f"  Local path : {local_path}")
                logger.warning(f"  Sonarr path: {sonarr_path}")

        return {
            "matched": matched,
            "unmatched": unmatched,
            "refreshed": refreshed,
            "rescanned": rescanned,
        }

    def run_fix(self):
        self.run_in_progress = True
        self.set_buttons_enabled(False)
        self.set_activity("Building preview")

        try:
            if not self.fix_duplicates_var.get() and not self.rename_to_scheme_var.get():
                self.write_output("ERROR: Enable at least one cleanup option before running Fix.", "warning")
                self.show_banner("Turn on at least one cleanup option before running Fix.", kind="warning")
                self.set_activity("Waiting for cleanup options")
                return

            path, logger = self.prepare_run()
            if not path:
                return

            result = self.collect_show_data(path, logger)
            if not result:
                return

            analyzed, shows_with_issues = result

            if not shows_with_issues:
                logger.success("Nothing needs changing.")
                logger.save()
                self.show_banner("Nothing needs changing. All detected season folders already match the current rules.", kind="success")
                self.set_activity("Nothing needs changing")
                return

            total_duplicates = sum(len(show["duplicates"]) for show in analyzed)
            total_naming = sum(len(show["naming_issues"]) for show in analyzed)

            logger.header("=" * 80)
            logger.header("Planned Changes")
            logger.summary(f"Shows with issues: {len(shows_with_issues)}")
            logger.summary(f"Duplicate groups found: {total_duplicates}")
            logger.summary(f"Naming issues found: {total_naming}")
            logger.info()

            preview_stats = [
                process_show(
                    show,
                    logger,
                    dry_run=True,
                    fix_duplicates=self.fix_duplicates_var.get(),
                    rename_to_scheme=self.rename_to_scheme_var.get(),
                )
                for show in shows_with_issues
            ]

            preview_total = combine_stats(preview_stats)
            estimated_actions = preview_total["merged"] + preview_total["renamed"]

            self.set_summary(
                shows=len(shows_with_issues),
                dupes=total_duplicates,
                names=total_naming,
                actions=estimated_actions,
            )

            logger.info()
            logger.header("=" * 80)
            logger.header("Preview Summary")
            logger.summary(f"Shows affected: {preview_total['shows_processed']}")
            logger.summary(f"Duplicate groups found: {preview_total['duplicates_found']}")
            logger.summary(f"Naming issues found: {preview_total['naming_issues_found']}")
            logger.summary(f"Folders to merge: {preview_total['merged']}")
            logger.summary(f"Folders to rename: {preview_total['renamed']}")
            logger.summary(f"Possible skips: {preview_total['skipped']}")

            proceed = self.ask_inline_question(
                "Preview finished. Do you want to apply these changes now?",
                kind="warning",
                primary_text="Apply Changes",
                secondary_text="Cancel",
            )

            if not proceed:
                logger.info()
                logger.warning("Fix run cancelled after preview.")
                logger.save()
                self.show_banner("No changes were made. The fix run was cancelled after the preview.", kind="warning")
                self.set_activity("Preview cancelled")
                return

            self.set_activity("Applying changes")

            logger.info()
            logger.header("=" * 80)
            logger.header("Applying Changes")

            actual_stats = [
                process_show(
                    show,
                    logger,
                    dry_run=False,
                    fix_duplicates=self.fix_duplicates_var.get(),
                    rename_to_scheme=self.rename_to_scheme_var.get(),
                )
                for show in shows_with_issues
            ]

            total = combine_stats(actual_stats)
            actions_taken = total["merged"] + total["renamed"]

            sonarr_result = None
            if actions_taken > 0:
                affected_show_paths = [show["show_path"] for show in shows_with_issues]
                try:
                    sonarr_result = self.notify_sonarr_after_fix(affected_show_paths, logger)
                except Exception as sonarr_error:
                    logger.warning(f"Sonarr update failed: {sonarr_error}")

            logger.info()
            logger.header("=" * 80)
            logger.header("Final Summary")
            logger.summary(f"Shows processed: {total['shows_processed']}")
            logger.summary(f"Duplicate groups found: {total['duplicates_found']}")
            logger.summary(f"Naming issues found: {total['naming_issues_found']}")
            logger.summary(f"Folders merged: {total['merged']}")
            logger.summary(f"Folders renamed: {total['renamed']}")
            logger.summary(f"Skipped: {total['skipped']}")
            logger.summary(f"Errors: {total['errors']}")

            if sonarr_result is not None:
                logger.summary(f"Sonarr series updated: {len(sonarr_result['matched'])}")
                logger.summary(f"Sonarr series not matched: {len(sonarr_result['unmatched'])}")

            logger.info()
            logger.success("Fix run finished.")
            logger.save()

            self.set_summary(
                shows=len(shows_with_issues),
                dupes=total["duplicates_found"],
                names=total["naming_issues_found"],
                actions=actions_taken,
            )

            if sonarr_result is not None:
                self.show_banner(
                    f"Fix run finished. Applied {actions_taken} change(s), then queued Sonarr updates for {len(sonarr_result['matched'])} series.",
                    kind="success",
                )
            else:
                self.show_banner(
                    f"Fix run finished. Applied {actions_taken} change(s). A log file was saved to Downloads.",
                    kind="success",
                )

            self.set_activity("Fix run finished")

        except Exception as e:
            self.write_output(f"ERROR: {e}", "warning")
            self.show_banner(f"Fix failed: {e}", kind="warning")
            self.set_activity("Fix failed")
        finally:
            self.run_in_progress = False
            self.set_buttons_enabled(True)

    def on_close(self):
        if self.settings_dirty:
            try:
                self.save_all_settings(quiet=True)
            except Exception:
                pass
        self.root.destroy()


def main():
    root = tk.Tk()
    App(root)
    root.mainloop()


if __name__ == "__main__":
    main()
