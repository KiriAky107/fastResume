"""Local DOCX to PDF conversion using an isolated Office automation session.

No document content is sent to a server.  The caller owns the final output;
all conversion profiles, manifests and intermediate files are temporary.
"""

from __future__ import annotations

import json
import math
import os
from pathlib import Path
import shutil
import signal
import subprocess
import tempfile


# The command is fixed; document paths travel through a JSON manifest rather
# than being interpolated into PowerShell source text.
_WORD_SCRIPT = r"""
$ErrorActionPreference = 'Stop'
[Console]::OutputEncoding = [Text.UTF8Encoding]::new($false)
$manifest = Get-Content -LiteralPath $env:FASTRESUME_WORD_MANIFEST -Raw -Encoding UTF8 | ConvertFrom-Json
$word = $null
$document = $null
$documents = $null
$scratch = $null
$windows = $null
$window = $null
$owned = $false
try {
    Add-Type -TypeDefinition @'
using System;
using System.Runtime.InteropServices;
public static class FastResumeWordWindow {
    [DllImport("user32.dll")]
    public static extern uint GetWindowThreadProcessId(IntPtr hwnd, out uint processId);
}
'@
    $previous = @(Get-Process -Name WINWORD -ErrorAction SilentlyContinue | ForEach-Object { $_.Id })
    $started = [DateTime]::UtcNow
    $word = New-Object -ComObject Word.Application
    $candidates = @(Get-CimInstance Win32_Process -Filter "Name='WINWORD.EXE'" | Where-Object {
        $previous -notcontains [int]$_.ProcessId -and
        $_.CreationDate.ToUniversalTime() -ge $started.AddSeconds(-1) -and
        $_.CommandLine -match '/Automation\s+-Embedding'
    })
    if ($candidates.Count -ne 1) {
        throw 'A private Word automation process could not be uniquely identified; close Word or select LibreOffice. Existing documents were not changed.'
    }
    # Word exposes HWND on Window, not Application. Create an unsaved scratch
    # document only after finding a new private automation process, then bind
    # that process to this exact COM instance before changing its settings.
    $documents = $word.Documents
    $scratch = $documents.Add()
    $windows = $scratch.Windows
    $window = $windows.Item(1)
    [uint32]$wordProcessId = 0
    [void][FastResumeWordWindow]::GetWindowThreadProcessId([IntPtr]$window.Hwnd, [ref]$wordProcessId)
    if ($wordProcessId -eq 0 -or $previous -contains [int]$wordProcessId -or $wordProcessId -ne [uint32]$candidates[0].ProcessId) {
        throw 'Word attached to an existing session; close Word or select LibreOffice. Existing documents were not changed.'
    }
    $process = Get-Process -Id $wordProcessId -ErrorAction Stop
    $owned = $true
    $owner = @{ pid = [int]$wordProcessId; start_ticks = $process.StartTime.ToUniversalTime().Ticks.ToString() }
    [IO.File]::WriteAllText($manifest.owner, ($owner | ConvertTo-Json -Compress), [Text.UTF8Encoding]::new($false))
    $word.Visible = $false
    $word.DisplayAlerts = 0
    $word.AutomationSecurity = 3
    $word.ScreenUpdating = $false
    $scratch.Close(0)
    [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($window)
    $window = $null
    [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($windows)
    $windows = $null
    [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($scratch)
    $scratch = $null
    $document = $documents.Open([string]$manifest.docx, $false, $true, $false)
    $document.Repaginate()
    $document.ExportAsFixedFormat([string]$manifest.pdf, 17, $false, 0, 0, 1, 1, 0, $true, $true, 1, $true, $true, $false)
} catch {
    [Console]::Error.WriteLine($_.Exception.Message)
    exit 1
} finally {
    if ($null -ne $document) {
        try { $document.Close(0) } catch {}
        try { [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($document) } catch {}
    }
    if ($null -ne $scratch) {
        try { $scratch.Close(0) } catch {}
        try { [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($scratch) } catch {}
    }
    foreach ($item in @($window, $windows, $documents)) {
        if ($null -ne $item) {
            try { [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($item) } catch {}
        }
    }
    if ($null -ne $word) {
        if ($owned) { try { $word.Quit(0) } catch {} }
        try { [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($word) } catch {}
    }
}
"""


_WORD_CLEANUP_SCRIPT = r"""
$ErrorActionPreference = 'Stop'
if (Test-Path -LiteralPath $env:FASTRESUME_WORD_OWNER) {
    $owner = Get-Content -LiteralPath $env:FASTRESUME_WORD_OWNER -Raw -Encoding UTF8 | ConvertFrom-Json
    $process = Get-Process -Id ([int]$owner.pid) -ErrorAction SilentlyContinue
    if ($null -ne $process -and $process.ProcessName -eq 'WINWORD' -and
        $process.StartTime.ToUniversalTime().Ticks.ToString() -eq [string]$owner.start_ticks) {
        Stop-Process -Id $process.Id -Force
    }
}
"""


def _powershell() -> str | None:
    found = shutil.which("pwsh") or shutil.which("powershell")
    if found:
        return found
    candidate = Path(os.environ.get("SystemRoot", "C:/Windows")) / "System32/WindowsPowerShell/v1.0/powershell.exe"
    return str(candidate) if candidate.is_file() else None


def _hidden_options() -> dict:
    if os.name == "nt":
        return {"creationflags": subprocess.CREATE_NO_WINDOW}
    return {}


def _message(completed: subprocess.CompletedProcess) -> str:
    # Do not repeat the entire Office log or echo document content in errors.
    value = (completed.stderr or completed.stdout or "").strip()
    return value[:1800] or f"process exited with status {completed.returncode}"


def _cleanup_word(powershell: str, owner: Path) -> None:
    """Only terminate the Word PID that this job positively identified as new."""
    if not owner.is_file():
        return
    environment = os.environ.copy()
    environment["FASTRESUME_WORD_OWNER"] = str(owner)
    try:
        subprocess.run(
            [powershell, "-NoProfile", "-NonInteractive", "-Command", _WORD_CLEANUP_SCRIPT],
            env=environment, capture_output=True, timeout=15, check=False,
            **_hidden_options(),
        )
    except (OSError, subprocess.TimeoutExpired):
        # A failure to verify ownership must never turn into a broad WINWORD kill.
        pass


def _convert_word(docx: Path, pdf: Path, temporary: Path, timeout: float) -> None:
    if os.name != "nt":
        raise RuntimeError("The Word PDF backend requires Windows and Microsoft Word.")
    powershell = _powershell()
    if not powershell:
        raise RuntimeError("PowerShell is required for the Word PDF backend.")
    manifest = temporary / "word-manifest.json"
    owner = temporary / "word-owner.json"
    manifest.write_text(
        json.dumps({"docx": str(docx), "pdf": str(pdf), "owner": str(owner)}, ensure_ascii=False),
        encoding="utf-8",
    )
    environment = os.environ.copy()
    environment["FASTRESUME_WORD_MANIFEST"] = str(manifest)
    try:
        try:
            completed = subprocess.run(
                [powershell, "-NoProfile", "-NonInteractive", "-Command", _WORD_SCRIPT],
                env=environment, capture_output=True, text=True,
                encoding="utf-8", errors="replace", timeout=timeout, check=False,
                **_hidden_options(),
            )
        except subprocess.TimeoutExpired as exc:
            raise RuntimeError(f"Word PDF conversion exceeded {timeout:g} seconds.") from exc
        except OSError as exc:
            raise RuntimeError(f"Cannot start the Word PDF backend: {exc}") from exc
        if completed.returncode != 0:
            raise RuntimeError(f"Word PDF conversion failed: {_message(completed)}")
    finally:
        # Quit() normally handles this. Also covers timeout and interruption.
        _cleanup_word(powershell, owner)
    if not pdf.is_file() or pdf.stat().st_size == 0:
        raise RuntimeError("Word did not create a PDF file.")


def _libreoffice(office_path: str | Path | None) -> str | None:
    if office_path:
        candidate = Path(office_path).expanduser()
        if candidate.is_file():
            # The .com wrapper on Windows waits for the conversion to finish.
            wrapper = candidate.with_suffix(".com")
            return str(wrapper.resolve() if os.name == "nt" and wrapper.is_file() else candidate.resolve())
        found = shutil.which(str(office_path))
        if found:
            return found
        raise RuntimeError(f"LibreOffice executable does not exist: {office_path}")
    for name in (("soffice.com", "soffice", "libreoffice") if os.name == "nt" else ("soffice", "libreoffice")):
        found = shutil.which(name)
        if found:
            return found
    candidates = [Path("/Applications/LibreOffice.app/Contents/MacOS/soffice")]
    if os.name == "nt":
        candidates = [
            Path(base) / "LibreOffice/program" / executable
            for base in (os.environ.get("ProgramFiles"), os.environ.get("ProgramFiles(x86)"))
            if base
            for executable in ("soffice.com", "soffice.exe")
        ]
    return next((str(path) for path in candidates if path.is_file()), None)


def _stop_process_tree(process: subprocess.Popen) -> None:
    if os.name == "nt" and process.poll() is None:
        taskkill = Path(os.environ.get("SystemRoot", "C:/Windows")) / "System32/taskkill.exe"
        try:
            subprocess.run(
                [str(taskkill), "/PID", str(process.pid), "/T", "/F"],
                capture_output=True, timeout=10, check=False, **_hidden_options(),
            )
        except (OSError, subprocess.TimeoutExpired):
            process.kill()
    elif os.name != "nt":
        try:
            os.killpg(process.pid, signal.SIGKILL)
        except ProcessLookupError:
            pass
    try:
        process.wait(timeout=5)
    except subprocess.TimeoutExpired:
        process.kill()
        try:
            process.wait(timeout=5)
        except subprocess.TimeoutExpired as exc:
            raise RuntimeError("LibreOffice did not exit after termination; its temporary profile could not be safely released.") from exc
    finally:
        for stream in (process.stdout, process.stderr):
            if stream is not None:
                stream.close()


def _convert_libreoffice(
    docx: Path, pdf: Path, temporary: Path, timeout: float, office_path: str | Path | None,
) -> None:
    executable = _libreoffice(office_path)
    if not executable:
        raise RuntimeError("LibreOffice was not found. Install it or pass --office-path /path/to/soffice.")
    profile = temporary / "libreoffice-profile"
    export = temporary / "libreoffice-export"
    export.mkdir()
    command = [
        executable, f"-env:UserInstallation={profile.as_uri()}",
        "--headless", "--convert-to", "pdf:writer_pdf_Export",
        "--outdir", str(export), str(docx),
    ]
    options = _hidden_options()
    if os.name != "nt":
        options["start_new_session"] = True
    try:
        process = subprocess.Popen(command, stdout=subprocess.PIPE, stderr=subprocess.PIPE, **options)
    except OSError as exc:
        raise RuntimeError(f"Cannot start LibreOffice: {exc}") from exc
    try:
        stdout, stderr = process.communicate(timeout=timeout)
    except subprocess.TimeoutExpired as exc:
        _stop_process_tree(process)
        raise RuntimeError(f"LibreOffice PDF conversion exceeded {timeout:g} seconds.") from exc
    except BaseException:
        _stop_process_tree(process)
        raise
    if process.returncode != 0:
        result = subprocess.CompletedProcess(command, process.returncode, stdout.decode("utf-8", "replace"), stderr.decode("utf-8", "replace"))
        raise RuntimeError(f"LibreOffice PDF conversion failed: {_message(result)}")
    converted = export / f"{docx.stem}.pdf"
    if not converted.is_file() or converted.stat().st_size == 0:
        raise RuntimeError("LibreOffice returned without creating a PDF. Check that the source DOCX can be opened.")
    shutil.move(str(converted), str(pdf))


def convert_pdf(
    docx: Path, pdf: Path, backend: str = "auto",
    office_path: str | Path | None = None, timeout: float = 120,
    workdir: Path | None = None,
) -> None:
    """Export one DOCX locally; raise RuntimeError with an actionable error.

    ``auto`` tries Word first on Windows, then LibreOffice. Supplying an explicit
    ``office_path`` selects LibreOffice in auto mode. Existing outputs are never
    overwritten by this low-level converter.
    """
    docx, pdf = Path(docx).resolve(), Path(pdf).resolve()
    if not docx.is_file():
        raise RuntimeError(f"Source DOCX does not exist: {docx}")
    if pdf.exists():
        raise RuntimeError(f"PDF output already exists: {pdf}")
    if not pdf.parent.is_dir():
        raise RuntimeError(f"PDF output directory does not exist: {pdf.parent}")
    if backend not in {"auto", "word", "libreoffice"}:
        raise RuntimeError(f"Unknown PDF backend: {backend}")
    if isinstance(timeout, bool) or not isinstance(timeout, (int, float)) or not math.isfinite(timeout) or timeout <= 0:
        raise RuntimeError("PDF conversion timeout must be a finite positive number.")
    choices = [backend]
    if backend == "auto":
        choices = ["libreoffice"] if office_path or os.name != "nt" else ["word", "libreoffice"]
    errors = []
    with tempfile.TemporaryDirectory(prefix="fastresume-pdf-", dir=workdir) as temporary:
        root = Path(temporary).resolve()
        for choice in choices:
            try:
                if choice == "word":
                    _convert_word(docx, pdf, root, timeout)
                else:
                    _convert_libreoffice(docx, pdf, root, timeout, office_path)
                return
            except RuntimeError as exc:
                errors.append(str(exc))
                # A failed backend may leave an incomplete PDF in caller staging.
                if pdf.is_file():
                    pdf.unlink()
    raise RuntimeError("\n".join(errors))
