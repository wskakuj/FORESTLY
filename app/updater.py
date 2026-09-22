"""
Forestly — Aktualizator GitHub (Mixin)
===================================================
Zależności: config.py (CURRENT_VERSION, GITHUB_USER, GITHUB_REPO)
Odpowiada za: sprawdzanie nowej wersji na GitHubie, pobieranie i
              instalowanie aktualizacji .exe przez skrypt PowerShell.
"""

import os
import sys
import json
import time
import base64
import subprocess
import tempfile
import webbrowser
import urllib.request
import threading
import re
from pathlib import Path

from tkinter import messagebox

from app.config import CURRENT_VERSION, GITHUB_USER, GITHUB_REPO


class UpdaterMixin:
    """Mixin dla ModernApp — metody aktualizacji z GitHub."""

    def check_github_update(self, manual=False):
        api_url = (
            f"https://api.github.com/repos/{GITHUB_USER}/{GITHUB_REPO}/releases/latest"
        )

        def parse_version(v_str):
            try:
                return tuple(map(int, re.findall(r"\d+", str(v_str))))
            except Exception:
                return (0,)

        def _check():
            try:
                if manual:
                    self.update_status("Sprawdzanie aktualizacji...", "#0078D7")
                req = urllib.request.Request(
                    api_url, headers={"User-Agent": "Forestly-Updater"}
                )
                with urllib.request.urlopen(req, timeout=5) as response:
                    data = json.loads(response.read().decode("utf-8"))
                    latest_version = data.get("tag_name")
                    if latest_version and parse_version(latest_version) > parse_version(
                            CURRENT_VERSION
                    ):
                        download_url = None
                        # W Release jest kilka wariantow EXE (Forestly.exe,
                        # Forestly_OLD.exe) — wybieramy wlasciwy.
                        exe_assets = [a for a in data.get("assets", []) if a["name"].endswith(".exe")]
                        chosen = next((a for a in exe_assets if a["name"] == "Forestly.exe"), None)
                        if chosen is None and exe_assets:
                            chosen = exe_assets[0]
                        if chosen:
                            download_url = chosen["browser_download_url"]
                        msg = f"Dostępna jest nowa wersja programu: {latest_version}\n(Obecnie używasz: {CURRENT_VERSION})\nCzy chcesz automatycznie pobrać i zainstalować aktualizację?"
                        changelog_body = data.get("body", "")
                        if messagebox.askyesno("Dostępna aktualizacja!", msg):
                            if download_url:
                                self.download_and_update(download_url, latest_version, changelog_body)
                            else:
                                self.log(
                                    "[UPDATE] Znaleziono wydanie, ale brak pliku .exe w załącznikach. Otwieram stronę..."
                                )
                                webbrowser.open(data.get("html_url"))
                    else:
                        if manual:
                            messagebox.showinfo(
                                "Aktualizacja",
                                f"Posiadasz najnowszą wersję programu ({CURRENT_VERSION}).",
                            )
                        self.update_status("Gotowy", "#0078D7", animate=False)
            except Exception as e:
                if manual:
                    # Pokazujemy błąd TYLKO wtedy, gdy użytkownik sam kliknął "Sprawdź update"
                    messagebox.showerror(
                        "Błąd połączenia", f"Nie udało się połączyć z GitHubem:\n{e}"
                    )
                    self.log(
                        f"[UPDATE BŁĄD] Nie można pobrać informacji o aktualizacji: {e}"
                    )
                    self.update_status("Gotowy", "#0078D7", animate=False)
                # Przy automatycznym sprawdzaniu (manual=False) ignorujemy błędy po cichu

        # WAŻNE: Ta linijka musi być na tym samym poziomie wcięcia co 'def _check():'
        threading.Thread(target=_check, daemon=True).start()

    def download_and_update(self, url, new_version, changelog_text=""):
            if not getattr(sys, "frozen", False):
                messagebox.showwarning(
                    "Wersja deweloperska",
                    "Automatyczna podmiana pliku działa tylko po skompilowaniu programu do .exe!",
                )
                return

            try:
                self.log("[UPDATE] Przygotowywanie graficznego instalatora Windows...")
                self.update_status("Uruchamianie aktualizatora...", "#0078D7")

                current_exe_path = Path(sys.executable).resolve()
                target_dir_path = current_exe_path.parent
                pid = os.getpid()

                def ps_literal(value: str) -> str:
                    return "'" + str(value).replace("'", "''") + "'"

                exe_path_ps = ps_literal(current_exe_path)
                target_dir_ps = ps_literal(target_dir_path)
                url_ps = ps_literal(url)

                changelog_data = json.dumps(
                    {
                        "version": new_version,
                        "changelog": changelog_text,
                    },
                    ensure_ascii=False,
                )

                import base64
                b64_changelog = base64.b64encode(changelog_data.encode("utf-8")).decode("utf-8")

                ps_script = f"""
    Add-Type -AssemblyName System.Windows.Forms
    Add-Type -AssemblyName System.Drawing

    function Clear-PyInstallerEnv {{
        $names = @(
            '_MEIPASS',
            '_MEIPASS2',
            'PYTHONHOME',
            'PYTHONPATH',
            'TCL_LIBRARY',
            'TK_LIBRARY',
            '_PYVENV_LAUNCHER_',
            '__PYVENV_LAUNCHER__'
        )

        foreach ($n in $names) {{
            Remove-Item -Path "Env:$n" -ErrorAction SilentlyContinue
        }}

        Get-ChildItem Env: -ErrorAction SilentlyContinue |
            Where-Object {{ $_.Name -like '_MEI*' -or $_.Name -like '_PYI*' }} |
            ForEach-Object {{ Remove-Item -Path "Env:$($_.Name)" -ErrorAction SilentlyContinue }}

        if ($env:PATH) {{
            $clean = $env:PATH -split ';' | Where-Object {{ $_ -and ($_ -notmatch '_MEI') }}
            $env:PATH = ($clean -join ';')
        }}
    }}

    function Test-FileLocked {{
        param([string]$Path)

        if (-not (Test-Path -Path $Path)) {{
            return $false
        }}

        try {{
            $fs = [System.IO.File]::Open($Path, 'Open', 'ReadWrite', 'None')
            $fs.Close()
            return $false
        }} catch {{
            return $true
        }}
    }}

    Clear-PyInstallerEnv

    $form = New-Object System.Windows.Forms.Form
    $form.Text = "Aktualizator Forestly"
    $form.Size = New-Object System.Drawing.Size(480, 160)
    $form.StartPosition = "CenterScreen"
    $form.FormBorderStyle = "FixedToolWindow"
    $form.BackColor = [System.Drawing.Color]::FromArgb(37, 37, 38)
    $form.ForeColor = [System.Drawing.Color]::White
    $form.TopMost = $true

    $label = New-Object System.Windows.Forms.Label
    $label.Location = New-Object System.Drawing.Point(20, 20)
    $label.Size = New-Object System.Drawing.Size(440, 30)
    $label.Font = New-Object System.Drawing.Font("Segoe UI", 11)
    $label.Text = "Czekam na zamknięcie starej wersji programu..."
    $form.Controls.Add($label)

    $progressBar = New-Object System.Windows.Forms.ProgressBar
    $progressBar.Location = New-Object System.Drawing.Point(20, 60)
    $progressBar.Size = New-Object System.Drawing.Size(420, 20)
    $progressBar.Style = "Marquee"
    $progressBar.MarqueeAnimationSpeed = 30
    $form.Controls.Add($progressBar)

    $form.Add_Shown({{
        $form.Refresh()

        $pidToWait = {pid}
        $exePath = {exe_path_ps}
        $targetDir = {target_dir_ps}
        $url = {url_ps}
        $tempExe = Join-Path $env:TEMP "Forestly_Najnowszy.exe"

        $waitStopwatch = [System.Diagnostics.Stopwatch]::StartNew()

        while ((Get-Process -Id $pidToWait -ErrorAction SilentlyContinue) -or (Test-FileLocked $exePath)) {{
            [System.Windows.Forms.Application]::DoEvents()
            Start-Sleep -Milliseconds 200

            if ($waitStopwatch.Elapsed.TotalSeconds -gt 30) {{
                break
            }}
        }}

        Start-Sleep -Milliseconds 500

        $label.Text = "Pobieranie nowej wersji. To może chwilę potrwać..."
        $form.Refresh()

        try {{
            [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
            $webClient = New-Object System.Net.WebClient
            $webClient.DownloadFileAsync([uri]$url, $tempExe)

            while ($webClient.IsBusy) {{
                [System.Windows.Forms.Application]::DoEvents()
                Start-Sleep -Milliseconds 50
            }}

            $file = Get-Item $tempExe -ErrorAction SilentlyContinue
            if ($null -eq $file -or ($file.Length / 1MB) -lt 10) {{
                $label.Text = "BŁĄD: Pobrany plik jest uszkodzony."
                $label.ForeColor = [System.Drawing.Color]::Red
                $progressBar.Style = "Blocks"
                $form.Refresh()
                Start-Sleep -Seconds 5
                $form.Close()
                exit 1
            }}

            $label.Text = "Pobrano poprawnie. Podmiana plików..."
            $form.Refresh()
            Start-Sleep -Milliseconds 500

            if (Test-FileLocked $exePath) {{
                Start-Sleep -Seconds 2
            }}

            $backupName = [System.IO.Path]::GetFileName($exePath) + ".old_" + (Get-Date -Format yyyyMMddHHmmss)
            $backupPath = Join-Path $targetDir $backupName

            Remove-Item -Path $backupPath -Force -ErrorAction SilentlyContinue

            if (Test-Path -Path $exePath) {{
                try {{
                    Rename-Item -Path $exePath -NewName $backupName -Force -ErrorAction Stop
                }} catch {{
                    Remove-Item -Path $exePath -Force -ErrorAction SilentlyContinue
                }}
            }}

            Move-Item -Path $tempExe -Destination $exePath -Force
            Remove-Item -Path $backupPath -Force -ErrorAction SilentlyContinue

            $changelogFile = Join-Path $targetDir "pending_changelog.json"
            $b64Data = "{b64_changelog}"
            $jsonBytes = [System.Convert]::FromBase64String($b64Data)
            [System.IO.File]::WriteAllBytes($changelogFile, $jsonBytes)

            $label.Text = "Zakończono! Uruchamianie nowej wersji..."
            $label.ForeColor = [System.Drawing.Color]::LightGreen
            $progressBar.Style = "Blocks"
            $progressBar.Value = 100
            $form.Refresh()

            Start-Sleep -Seconds 1
            Clear-PyInstallerEnv

            $psi = New-Object System.Diagnostics.ProcessStartInfo
            $psi.FileName = $exePath
            $psi.WorkingDirectory = $targetDir
            $psi.UseShellExecute = $false
            $psi.CreateNoWindow = $true

            $removeNames = @(
                '_MEIPASS',
                '_MEIPASS2',
                'PYTHONHOME',
                'PYTHONPATH',
                'TCL_LIBRARY',
                'TK_LIBRARY',
                '_PYVENV_LAUNCHER_',
                '__PYVENV_LAUNCHER__'
            )

            foreach ($n in $removeNames) {{
                if ($psi.EnvironmentVariables.ContainsKey($n)) {{
                    $psi.EnvironmentVariables.Remove($n)
                }}
            }}

            $envKeys = @($psi.EnvironmentVariables.Keys)
            foreach ($key in $envKeys) {{
                if ($key -like '_MEI*' -or $key -like '_PYI*') {{
                    $psi.EnvironmentVariables.Remove($key)
                }}
            }}

            $pathKey = $null
            foreach ($key in @($psi.EnvironmentVariables.Keys)) {{
                if ($key -eq 'PATH') {{
                    $pathKey = $key
                }}
            }}

            if ($pathKey) {{
                $psi.EnvironmentVariables[$pathKey] = $env:PATH
            }} else {{
                $psi.EnvironmentVariables['PATH'] = $env:PATH
            }}

            [System.Diagnostics.Process]::Start($psi) | Out-Null
        }} catch {{
            $label.Text = "Wystąpił błąd podczas aktualizacji."
            $label.ForeColor = [System.Drawing.Color]::Red
            $progressBar.Style = "Blocks"
            $form.Refresh()
            Start-Sleep -Seconds 5
        }}

        $form.Close()
    }})

    $form.ShowDialog()
    """

                clean_env = {}
                skip_exact = {
                    "_MEIPASS",
                    "_MEIPASS2",
                    "PYTHONHOME",
                    "PYTHONPATH",
                    "TCL_LIBRARY",
                    "TK_LIBRARY",
                    "_PYVENV_LAUNCHER_",
                    "__PYVENV_LAUNCHER__",
                }

                meipass_dir = None
                if getattr(sys, "_MEIPASS", None):
                    try:
                        meipass_dir = Path(sys._MEIPASS).resolve(strict=False)
                    except Exception:
                        meipass_dir = Path(sys._MEIPASS)

                for key, value in os.environ.items():
                    upper_key = key.upper()

                    if upper_key in skip_exact:
                        continue

                    if upper_key.startswith("_MEI") or upper_key.startswith("_PYI"):
                        continue

                    out_key = key

                    if upper_key == "PATH":
                        out_key = "PATH"
                        parts = []

                        for p in str(value).split(os.pathsep):
                            if not p:
                                continue

                            if "_MEI" in p.upper():
                                continue

                            try:
                                pp = Path(p).resolve(strict=False)
                                if pp.name.upper().startswith("_MEI"):
                                    continue
                                if meipass_dir and pp == meipass_dir:
                                    continue
                            except Exception:
                                pass

                            parts.append(p)

                        value = os.pathsep.join(parts)

                    clean_env[out_key] = value

                subprocess.Popen(
                    [
                        "powershell",
                        "-NoProfile",
                        "-STA",
                        "-ExecutionPolicy",
                        "Bypass",
                        "-Command",
                        ps_script,
                    ],
                    env=clean_env,
                    creationflags=subprocess.CREATE_NO_WINDOW,
                )

                self.destroy()
                os._exit(0)

            except Exception as e:
                self.log(f"[UPDATE BŁĄD] {e}")
                messagebox.showerror("Błąd", str(e))
                self.update_status("Gotowy", "#0078D7", animate=False)

