"""
cli_menu.py — ThreatMap Infra Post-Scan Interactive Menu

Provides:
  - PostScanMenu: arrow-key navigation using questionary
  - open_in_browser(): opens HTML in best available browser
  - export_libreoffice(): converts TXT → ODT if LibreOffice available,
                          otherwise copies TXT with instructions
  - show_logs(): prints path and first N lines of scan log

Designed to loop until user selects Exit.
Never crashes — all actions wrapped in try/except.
"""

import os
import shutil
import subprocess
import sys
from pathlib import Path
from typing import Optional

import questionary
from rich.console import Console
from rich.rule import Rule
from rich.syntax import Syntax

from scan_logger import get_logger

log     = get_logger("menu")
console = Console()

Q = questionary.Style([
    ("qmark",       "fg:red bold"),
    ("question",    "fg:white bold"),
    ("answer",      "fg:cyan bold"),
    ("pointer",     "fg:red bold"),
    ("highlighted", "fg:cyan bold"),
    ("selected",    "fg:cyan"),
])

# ── Helpers ───────────────────────────────────────────────────────────────────

def _ok(msg):   console.print(f"  [bold green][[+]][/bold green]  {msg}")
def _w(msg):    console.print(f"  [bold yellow][[!]][/bold yellow]  {msg}")
def _e(msg):    console.print(f"  [bold red][[-]][/bold red]  {msg}")
def _i(msg):    console.print(f"  [bold blue][[*]][/bold blue]  {msg}")


def open_in_browser(html_path: str) -> bool:
    """Open HTML in best available browser. Returns True if opened."""
    if not html_path or not Path(html_path).is_file():
        _e(f"HTML report not found: {html_path}")
        return False

    browsers = ["firefox","chromium","chromium-browser","google-chrome",
                "brave-browser","opera","xdg-open"]
    for browser in browsers:
        if shutil.which(browser):
            try:
                subprocess.Popen(
                    [browser, html_path],
                    stdout=subprocess.DEVNULL,
                    stderr=subprocess.DEVNULL,
                )
                _ok(f"Opened in [bold]{browser}[/bold]")
                _i(f"Path: [cyan]{html_path}[/cyan]")
                return True
            except Exception as exc:
                log.warning("[menu] browser %s failed: %s", browser, exc)
                continue

    _w("No browser found. Open this file manually:")
    console.print(f"  [cyan]{html_path}[/cyan]")
    return False


def export_libreoffice(txt_path: str, output_dir: str) -> Optional[str]:
    """
    Convert TXT report → ODT using LibreOffice headless.
    If LibreOffice not available, saves a copy as .txt with clear instructions.
    Returns path to created file.
    """
    if not txt_path or not Path(txt_path).is_file():
        _e(f"TXT report not found: {txt_path}")
        return None

    out_dir = Path(output_dir)
    out_dir.mkdir(parents=True, exist_ok=True)

    # Try LibreOffice headless conversion
    lo_bin = shutil.which("libreoffice") or shutil.which("soffice")
    if lo_bin:
        try:
            _i("Converting to ODT via LibreOffice...")
            result = subprocess.run(
                [lo_bin, "--headless", "--convert-to", "odt",
                 "--outdir", str(out_dir), txt_path],
                capture_output=True, text=True, timeout=30,
            )
            if result.returncode == 0:
                # LibreOffice names it <original_stem>.odt
                stem   = Path(txt_path).stem
                odt    = out_dir / f"{stem}.odt"
                if odt.exists():
                    _ok(f"ODT saved: [cyan]{odt}[/cyan]")
                    return str(odt)
        except subprocess.TimeoutExpired:
            _w("LibreOffice conversion timed out")
        except Exception as exc:
            log.warning("[menu] LibreOffice conversion failed: %s", exc)

    # Fallback: save plain text copy, explain how to open
    stem    = Path(txt_path).stem
    txt_out = out_dir / f"{stem}_export.txt"
    shutil.copy2(txt_path, txt_out)
    _w("LibreOffice not found — saved plain text instead:")
    console.print(f"  [cyan]{txt_out}[/cyan]")
    console.print()
    console.print("  [dim]To install LibreOffice:[/dim]")
    console.print("  [white]sudo apt install libreoffice[/white]")
    console.print()
    console.print("  [dim]Or open the .txt file in any word processor and save as .odt[/dim]")
    return str(txt_out)


def show_logs(log_path: str, tail_lines: int = 30) -> None:
    """Print log file path and tail of scan log."""
    if not log_path or not Path(log_path).is_file():
        _w(f"Log file not found: {log_path}")
        return

    _ok(f"Log file: [cyan]{log_path}[/cyan]")
    console.print()

    try:
        lines = Path(log_path).read_text(encoding="utf-8", errors="replace").splitlines()
        tail  = lines[-tail_lines:] if len(lines) > tail_lines else lines

        if len(lines) > tail_lines:
            _i(f"Showing last {tail_lines} of {len(lines)} lines")
        console.print()
        console.print(Rule(style="dim"))
        for line in tail:
            # Colour-code log levels for readability
            if " ERROR " in line or " CRITICAL " in line:
                console.print(f"  [red]{line}[/red]")
            elif " WARNING " in line:
                console.print(f"  [yellow]{line}[/yellow]")
            elif " INFO " in line:
                console.print(f"  [dim]{line}[/dim]")
            else:
                console.print(f"  [dim]{line}[/dim]")
        console.print(Rule(style="dim"))
    except Exception as exc:
        _e(f"Could not read log: {exc}")


def show_report_paths(paths: dict[str, str]) -> None:
    """Print all generated report paths in a clean table."""
    console.print()
    console.print(Rule("[dim]Generated Reports[/dim]", style="dim green"))
    console.print()
    labels = {"html":"HTML Report ","txt":"Text Report ","json":"JSON Data   "}
    for fmt, path in paths.items():
        label = labels.get(fmt, fmt.upper())
        exists = Path(path).is_file() if path else False
        icon   = "[green]✔[/green]" if exists else "[red]✗[/red]"
        console.print(f"  {icon}  {label} → [cyan]{path}[/cyan]")
    console.print()


# ── Post-scan menu ────────────────────────────────────────────────────────────

class PostScanMenu:
    """
    Interactive post-scan menu. Loops until user selects Exit.

    Args:
        report_paths: dict with keys 'html', 'txt', 'json'
        log_path:     path to scan.log
        output_dir:   directory for any additional exports
    """

    def __init__(
        self,
        report_paths: dict[str, str],
        log_path: str,
        output_dir: str,
    ) -> None:
        self.paths      = report_paths
        self.log_path   = log_path
        self.output_dir = output_dir

    def run(self) -> None:
        """Enter the menu loop. Returns only when user chooses Exit."""
        show_report_paths(self.paths)

        # Auto-open HTML immediately on first entry
        html = self.paths.get("html","")
        if html and Path(html).is_file():
            _i("Opening HTML report in browser...")
            open_in_browser(html)
            console.print()

        console.print(Rule("[dim]What would you like to do?[/dim]", style="dim"))
        console.print()

        while True:
            try:
                choice = questionary.select(
                    "  Select action:",
                    choices=self._build_choices(),
                    style=Q,
                ).ask()
            except (KeyboardInterrupt, EOFError):
                console.print()
                _i("Exiting.")
                break

            if choice is None or choice == "exit":
                console.print()
                _ok("Session complete. Stay safe.")
                console.print()
                break

            console.print()
            self._handle(choice)
            console.print()
            console.print(Rule(style="dim"))
            console.print()

    def _build_choices(self) -> list[dict]:
        choices = []

        html = self.paths.get("html","")
        if html and Path(html).is_file():
            choices.append({"name": "📄  View HTML Report (browser)", "value": "html"})

        txt = self.paths.get("txt","")
        if txt and Path(txt).is_file():
            choices.append({"name": "📝  Export as LibreOffice (.odt)", "value": "odt"})

        choices.append({"name": "📋  Show Report File Paths", "value": "paths"})
        choices.append({"name": "🔍  View Raw Scan Log",      "value": "logs"})
        choices.append({"name": "──────────────────────",    "value": "sep",
                        "disabled": "─"})
        choices.append({"name": "🚪  Exit",                   "value": "exit"})

        return choices

    def _handle(self, choice: str) -> None:
        try:
            if choice == "html":
                html = self.paths.get("html","")
                if html:
                    open_in_browser(html)
                else:
                    _e("HTML report was not generated.")

            elif choice == "odt":
                txt = self.paths.get("txt","")
                if txt:
                    export_libreoffice(txt, self.output_dir)
                else:
                    _e("Text report was not generated.")

            elif choice == "paths":
                show_report_paths(self.paths)

            elif choice == "logs":
                show_logs(self.log_path)

        except Exception as exc:
            _e(f"Action failed: {exc}")
            log.error("[menu] action %s failed: %s", choice, exc)
