# CMD Piper - Standalone Accessible Terminal App
# Version: 1.0
# Author: Vatsal Gautam/Hades
#
# Requirements:
#   pip install pywinpty wxPython
#
# Run:
#   python cmd_piper_app.py

import wx
import os
import sys
import re
import queue
import threading
import datetime
import subprocess
import time
import ctypes

# ── pywinpty (optional) ────────────────────────────────────────────────────────
try:
    import winpty
    HAS_WINPTY = True
except ImportError:
    HAS_WINPTY = False


# ─────────────────────────────────────────────────────────────────────────────
#  ANSI / VT100 escape-sequence stripping
# ─────────────────────────────────────────────────────────────────────────────
_ANSI_RE = re.compile(
    r'\x1b'
    r'(?:'
    r'\[[0-9;?]*[A-Za-z]'
    r'|'
    r'\][^\x07\x1b]*(?:\x07|\x1b\\)?'
    r'|'
    r'[@-_][^@-_]?'
    r')'
    r'|\r'
)

def _strip_ansi(text: str) -> str:
    return _ANSI_RE.sub('', text)


# ─────────────────────────────────────────────────────────────────────────────
#  NVDA speech via nvdaControllerClient.dll
# ─────────────────────────────────────────────────────────────────────────────
#
#  nvdaControllerClient is the official IPC DLL that NVDA ships.
#  nvdaController_speakText(wchar_p)  ->  0 on success.
#  nvdaController_cancelSpeech()      ->  cancels current speech.
#
#  If NVDA is not running or the DLL is not found, nvda_speak() returns False
#  silently — the output box still updates for sighted/other-SR users.

_nvda_dll   = None
_nvda_tried = False   # only search once per session

def _load_nvda_dll() -> None:
    global _nvda_dll, _nvda_tried
    if _nvda_tried:
        return
    _nvda_tried = True

    candidates = []
    for base in [
        os.environ.get("PROGRAMFILES",      r"C:\Program Files"),
        os.environ.get("PROGRAMFILES(X86)", r"C:\Program Files (x86)"),
    ]:
        if base:
            candidates += [
                os.path.join(base, "NVDA", "nvdaControllerClient64.dll"),
                os.path.join(base, "NVDA", "nvdaControllerClient32.dll"),
            ]
    # also look next to the script (user can copy the DLL here)
    here = os.path.dirname(os.path.abspath(sys.argv[0]))
    candidates += [
        os.path.join(here, "nvdaControllerClient64.dll"),
        os.path.join(here, "nvdaControllerClient32.dll"),
    ]

    for path in candidates:
        if not os.path.isfile(path):
            continue
        try:
            dll = ctypes.CDLL(path)
            dll.nvdaController_speakText.restype  = ctypes.c_int
            dll.nvdaController_speakText.argtypes = [ctypes.c_wchar_p]
            dll.nvdaController_cancelSpeech.restype  = ctypes.c_int
            dll.nvdaController_cancelSpeech.argtypes = []
            _nvda_dll = dll
            return
        except Exception:
            continue


def nvda_speak(text: str, interrupt: bool = False) -> bool:
    """Speak *text* through NVDA. Returns True if NVDA accepted it."""
    _load_nvda_dll()
    if _nvda_dll is None:
        return False
    try:
        if interrupt:
            _nvda_dll.nvdaController_cancelSpeech()
        return _nvda_dll.nvdaController_speakText(text) == 0
    except Exception:
        return False


# ─────────────────────────────────────────────────────────────────────────────
#  Constants
# ─────────────────────────────────────────────────────────────────────────────
APP_NAME    = "CMD Piper"
APP_VERSION = "1.0"

POLL_INTERVAL_MS = 50    # how often we drain the output queue (ms)
ANNOUNCE_CHUNK   = 250   # max chars sent to NVDA per announcement
MAX_OUTPUT_LINES = 2000  # trim output box when it exceeds this many lines

INTERACTIVE_PROGRAMS = {
    "python", "python3", "py",
    "node", "nodejs",
    "irb", "lua", "fsi",
    "powershell", "pwsh",
    "cmd", "bash", "sh", "zsh",
}

INSTRUCTIONS = f"""\
{APP_NAME} v{APP_VERSION} - Standalone Accessible Terminal
Author: Vatsal Gautam

HOW TO USE:
- Type any command in the Command box and press Enter or click Run.
- Output appears in the Output box and is announced by NVDA automatically.
- Use Find (Ctrl+F) to search through output text.
- Use Save Output (Ctrl+S) to save output to a .txt file.
- Use Clear Output (Ctrl+L) to wipe the output box.
- Type 'exit' or press Escape / Alt+F4 to close the app.

KEYBOARD SHORTCUTS:
  Enter       Run the typed command
  Escape      Close the app
  Alt+F4      Close the app
  Ctrl+C      Interrupt / stop the running command
  Up arrow    Previous command in history
  Down arrow  Next command in history
  Ctrl+F      Open Find dialog
  Ctrl+S      Save output to file
  Ctrl+L      Clear output
  F1          Show these instructions

BUILT-IN COMMANDS:
  cd <path>   Change working directory
  cd          Show current working directory
  cls / clear Clear the output box
  exit        Close the app

INTERACTIVE PROGRAMS (Python REPL, Node, etc.):
  Type the program name ALONE, e.g. 'python' or 'node'.
  'python script.py' runs a script (not interactive).
  Type 'exit()' (Python) or 'exit' to end the session.

SCREEN READER / NVDA:
  NVDA speech is via nvdaControllerClient.dll (NVDA's own IPC library).
  If the DLL is found, every new output chunk is spoken directly through
  NVDA without any focus change.  If NVDA is not running the output box
  still updates normally.

TIPS:
  Pipes and redirects work: e.g. 'dir | find "txt"'
  Working directory starts at your home folder.
  Saved files go to your Documents folder by default.
"""


# ─────────────────────────────────────────────────────────────────────────────
#  Helpers
# ─────────────────────────────────────────────────────────────────────────────

def _is_interactive_command(command: str) -> bool:
    command = command.strip().strip('"').strip("'").strip()
    parts   = command.split()
    if not parts:
        return False
    name = parts[0].lower()
    if name.endswith(".exe"):
        name = name[:-4]
    if name not in INTERACTIVE_PROGRAMS:
        return False
    args = parts[1:]
    if not args:
        return True
    if name in ("python", "python3", "py") and args == ["-i"]:
        return True
    return False


def _make_env() -> dict:
    env = os.environ.copy()
    env["PYTHONUNBUFFERED"]  = "1"
    env["PYTHONIOENCODING"]  = "utf-8"
    env["PYTHONUTF8"]        = "1"
    env["TERM"]              = "dumb"
    env["NO_COLOR"]          = "1"
    return env


def _kill_tree(process) -> None:
    try:
        subprocess.run(
            ["taskkill", "/F", "/T", "/PID", str(process.pid)],
            creationflags=subprocess.CREATE_NO_WINDOW,
            stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL,
        )
    except Exception:
        try:
            process.kill()
        except Exception:
            pass


# ─────────────────────────────────────────────────────────────────────────────
#  Backend sessions
# ─────────────────────────────────────────────────────────────────────────────

class _PtySession:
    """pywinpty-backed session. Used for both interactive and regular commands.

    Non-interactive commands are wrapped in 'cmd /c ...' so that shell
    builtins (echo, dir, set, type, if, for, ...) work correctly.
    Interactive programs (python, node, etc.) are spawned directly so their
    own REPL prompt handling isn't broken by an extra cmd shell.
    """

    def __init__(self, command: str, cwd: str, out_q: queue.Queue,
                 is_interactive: bool = False):
        self._q    = out_q
        self._dead = False
        cols, rows = 220, 50
        # Wrap non-interactive commands in cmd /c so shell builtins work.
        # Do NOT add extra quotes around the whole command — cmd /c passes
        # everything after /c directly to the shell, and double-quoting turns
        # "git init" into a single token that cmd tries to run as an executable.
        spawn_cmd = command if is_interactive else f"cmd /c {command}"
        self._pty  = winpty.PtyProcess.spawn(
            spawn_cmd, cwd=cwd, env=_make_env(), dimensions=(rows, cols),
        )
        threading.Thread(target=self._reader, daemon=True).start()

    def _reader(self):
        while not self._dead:
            try:
                data = self._pty.read(4096)
                if not data:
                    break
                self._q.put(_strip_ansi(data))
            except EOFError:
                break
            except Exception:
                break
        self._dead = True

    def write(self, text: str):
        if not self._dead:
            try:
                self._pty.write(text + "\r\n")
            except Exception:
                pass

    def interrupt(self):
        """Send Ctrl+C (ETX) to the running process via the PTY."""
        if not self._dead:
            try:
                self._pty.write("\x03")   # ASCII ETX = Ctrl+C
            except Exception:
                pass

    def poll(self):
        if self._dead:
            return 0
        try:
            return None if self._pty.isalive() else 0
        except Exception:
            return 0

    def kill(self):
        self._dead = True
        try:
            self._pty.terminate(force=True)
        except Exception:
            pass

    @property
    def pid(self):
        return getattr(self._pty, "pid", None)


class _PlainSession:
    """Plain subprocess fallback (no PTY)."""

    def __init__(self, command: str, cwd: str, out_q: queue.Queue,
                 is_interactive: bool):
        self._q           = out_q
        self._interactive = is_interactive
        self._process = subprocess.Popen(
            command, shell=True,
            stdout=subprocess.PIPE, stderr=subprocess.STDOUT,
            stdin=subprocess.PIPE,
            cwd=cwd,
            creationflags=subprocess.CREATE_NO_WINDOW,
            text=True, encoding="utf-8", errors="replace",
            bufsize=0, env=_make_env(),
        )
        threading.Thread(target=self._reader, daemon=True).start()

    def _reader(self):
        try:
            if self._interactive:
                buf = ""
                while True:
                    ch = self._process.stdout.read(1)
                    if not ch:
                        break
                    buf += ch
                    if ch == "\n" or buf.endswith(">>> ") or buf.endswith("... "):
                        self._q.put(buf)
                        buf = ""
                if buf:
                    self._q.put(buf)
            else:
                for line in self._process.stdout:
                    self._q.put(line)
        except Exception:
            pass
        self._process.wait()
        rc = self._process.returncode
        if not self._interactive and rc != 0:
            self._q.put(f"[Process exited with code {rc}]\n")

    def write(self, text: str):
        try:
            self._process.stdin.write(text + "\n")
            self._process.stdin.flush()
        except (OSError, BrokenPipeError):
            pass

    def interrupt(self):
        """Send Ctrl+C to the subprocess process group."""
        if self._process and self._process.poll() is None:
            try:
                ctypes.windll.kernel32.GenerateConsoleCtrlEvent(0, self._process.pid)
            except Exception:
                try:
                    self._process.send_signal(__import__('signal').CTRL_C_EVENT)
                except Exception:
                    pass

    def poll(self):
        return self._process.poll() if self._process else 0

    def kill(self):
        _kill_tree(self._process)

    @property
    def pid(self):
        return self._process.pid if self._process else None


# ─────────────────────────────────────────────────────────────────────────────
#  Find dialog
# ─────────────────────────────────────────────────────────────────────────────

class FindDialog(wx.Dialog):

    def __init__(self, parent, target: wx.TextCtrl):
        super().__init__(parent, title="Find in Output",
                         style=wx.DEFAULT_DIALOG_STYLE)
        self.target   = target
        self.last_pos = 0

        sizer = wx.BoxSizer(wx.VERTICAL)
        row   = wx.BoxSizer(wx.HORIZONTAL)
        row.Add(wx.StaticText(self, label="Find:"),
                flag=wx.ALIGN_CENTER_VERTICAL | wx.RIGHT, border=5)
        self.search_box = wx.TextCtrl(self, style=wx.TE_PROCESS_ENTER)
        self.search_box.SetName("Search term")
        self.search_box.Bind(wx.EVT_TEXT_ENTER, self._find)
        row.Add(self.search_box, proportion=1, flag=wx.EXPAND)
        sizer.Add(row, flag=wx.EXPAND | wx.ALL, border=8)

        self.case_check = wx.CheckBox(self, label="&Case sensitive")
        sizer.Add(self.case_check, flag=wx.LEFT | wx.BOTTOM, border=8)

        btns = wx.BoxSizer(wx.HORIZONTAL)
        b_find  = wx.Button(self, label="&Find Next")
        b_close = wx.Button(self, label="&Close", id=wx.ID_CLOSE)
        b_find.Bind(wx.EVT_BUTTON, self._find)
        b_close.Bind(wx.EVT_BUTTON, lambda e: self.Close())
        btns.Add(b_find,  flag=wx.RIGHT, border=5)
        btns.Add(b_close)
        sizer.Add(btns, flag=wx.ALIGN_RIGHT | wx.ALL, border=8)

        self.SetSizer(sizer)
        self.Fit()
        self.Centre()
        self.search_box.SetFocus()
        self.Bind(wx.EVT_CHAR_HOOK, lambda e:
                  self.Close() if e.GetKeyCode() == wx.WXK_ESCAPE else e.Skip())

    def _find(self, _event):
        term = self.search_box.GetValue()
        if not term:
            return
        full      = self.target.GetValue()
        cs        = self.case_check.IsChecked()
        haystack  = full if cs else full.lower()
        needle    = term if cs else term.lower()
        pos       = haystack.find(needle, self.last_pos)
        if pos == -1 and self.last_pos > 0:
            pos = haystack.find(needle, 0)
            if pos != -1:
                wx.MessageBox("Search wrapped to beginning.", "Find",
                              wx.OK | wx.ICON_INFORMATION, self)
        if pos == -1:
            wx.MessageBox(f'"{term}" not found.', "Find",
                          wx.OK | wx.ICON_INFORMATION, self)
            self.last_pos = 0
        else:
            self.target.SetSelection(pos, pos + len(term))
            self.target.ShowPosition(pos)
            self.target.SetFocus()
            self.last_pos = pos + len(term)


# ─────────────────────────────────────────────────────────────────────────────
#  Main frame
# ─────────────────────────────────────────────────────────────────────────────

class CMDPiperFrame(wx.Frame):

    def __init__(self):
        super().__init__(None, title=f"{APP_NAME} v{APP_VERSION}",
                         style=wx.DEFAULT_FRAME_STYLE)

        self.cwd          = os.path.expanduser("~")
        self.session      = None
        self.interactive  = False
        self.out_q        = queue.Queue()
        self._prompt_pending = False   # True while _wait_finish is running

        self.history         = []
        self.history_index   = -1
        self.history_draft   = ""

        self._announce_buf   = ""

        self._build_ui()
        self._build_menu()

        self.poll_timer = wx.Timer(self)
        self.Bind(wx.EVT_TIMER, self._on_poll, self.poll_timer)
        self.poll_timer.Start(POLL_INTERVAL_MS)

        self.Bind(wx.EVT_CLOSE,     self._on_close)
        self.Bind(wx.EVT_CHAR_HOOK, self._on_char_hook)

        self._append(f"CMD Piper ready.  Working directory: {self.cwd}\n{self.cwd}> ")
        self.input_box.SetFocus()

        self.SetSize((820, 600))
        self.SetMinSize((500, 360))
        self.Centre()

    # ── menu ──────────────────────────────────────────────────────────────────

    def _build_menu(self):
        mb = wx.MenuBar()

        fm = wx.Menu()
        fm.Append(wx.ID_SAVE, "&Save Output\tCtrl+S")
        fm.AppendSeparator()
        fm.Append(wx.ID_EXIT, "E&xit\tAlt+F4")
        mb.Append(fm, "&File")

        em = wx.Menu()
        em.Append(wx.ID_FIND, "&Find\tCtrl+F")
        self._id_clear = wx.NewIdRef()
        em.Append(self._id_clear, "C&lear Output\tCtrl+L")
        mb.Append(em, "&Edit")

        hm = wx.Menu()
        self._id_instr = wx.NewIdRef()
        hm.Append(self._id_instr, "&Instructions\tF1")
        mb.Append(hm, "&Help")

        self.SetMenuBar(mb)
        self.Bind(wx.EVT_MENU, self._on_save,         id=wx.ID_SAVE)
        self.Bind(wx.EVT_MENU, lambda e: self.Close(), id=wx.ID_EXIT)
        self.Bind(wx.EVT_MENU, self._on_find,         id=wx.ID_FIND)
        self.Bind(wx.EVT_MENU, self._on_clear,        id=self._id_clear)
        self.Bind(wx.EVT_MENU, self._on_instructions, id=self._id_instr)

    # ── UI ────────────────────────────────────────────────────────────────────

    def _build_ui(self):
        panel = wx.Panel(self)
        sz    = wx.BoxSizer(wx.VERTICAL)

        sz.Add(wx.StaticText(panel, label="Output:"),
               flag=wx.LEFT | wx.TOP, border=8)

        self.output_box = wx.TextCtrl(
            panel,
            style=wx.TE_MULTILINE | wx.TE_READONLY | wx.TE_RICH2 | wx.HSCROLL,
        )
        self.output_box.SetBackgroundColour(wx.Colour(12, 12, 12))
        self.output_box.SetForegroundColour(wx.Colour(204, 204, 204))
        self.output_box.SetName("Output")
        sz.Add(self.output_box, proportion=1,
               flag=wx.EXPAND | wx.LEFT | wx.RIGHT | wx.BOTTOM, border=8)

        row = wx.BoxSizer(wx.HORIZONTAL)
        row.Add(wx.StaticText(panel, label="Command:"),
                flag=wx.ALIGN_CENTER_VERTICAL | wx.RIGHT, border=6)
        self.input_box = wx.TextCtrl(panel, style=wx.TE_PROCESS_ENTER)
        self.input_box.SetName("Command input")
        self.input_box.Bind(wx.EVT_TEXT_ENTER, self._on_run)
        self.input_box.Bind(wx.EVT_KEY_DOWN,   self._on_key_down)
        row.Add(self.input_box, proportion=1, flag=wx.EXPAND)
        sz.Add(row, flag=wx.EXPAND | wx.LEFT | wx.RIGHT | wx.BOTTOM, border=8)

        btns = wx.BoxSizer(wx.HORIZONTAL)
        for label, handler in [
            ("&Run",          self._on_run),
            ("&Find",         self._on_find),
            ("&Save Output",  self._on_save),
            ("C&lear Output", self._on_clear),
            ("&Instructions", self._on_instructions),
            ("C&lose",        lambda e: self.Close()),
        ]:
            b = wx.Button(panel, label=label)
            b.Bind(wx.EVT_BUTTON, handler)
            btns.Add(b, flag=wx.RIGHT, border=5)
        sz.Add(btns, flag=wx.ALIGN_RIGHT | wx.LEFT | wx.RIGHT | wx.BOTTOM, border=8)

        panel.SetSizer(sz)

    # ── keyboard ──────────────────────────────────────────────────────────────

    def _on_char_hook(self, event):
        key  = event.GetKeyCode()
        ctrl = event.ControlDown()

        # Ctrl+C — interrupt running session
        if key == ord('C') and ctrl:
            if self.session and self.session.poll() is None:
                if self.interactive:
                    # In a REPL just send ETX — let the program handle it
                    self.session.interrupt()
                    self._append("^C\n")
                else:
                    # Non-interactive: kill the whole process tree immediately.
                    # interrupt() alone doesn't reach children of cmd /c.
                    self._prompt_pending = False   # claim the prompt slot before kill
                    self.session.kill()
                    self._append("^C\n")
                    self._append(f"{self.cwd}> ")
                return   # consume — never copy to clipboard while a command runs

        if key == wx.WXK_ESCAPE:
            focused = self.FindFocus()
            if focused and isinstance(focused.GetTopLevelParent(), wx.Dialog):
                event.Skip()
                return
            self.Close()
            return
        event.Skip()

    def _on_key_down(self, event):
        key = event.GetKeyCode()
        if key == wx.WXK_UP:
            if not self.history:
                event.Skip(); return
            if self.history_index == -1:
                self.history_draft = self.input_box.GetValue()
                self.history_index = len(self.history) - 1
            elif self.history_index > 0:
                self.history_index -= 1
            self.input_box.SetValue(self.history[self.history_index])
            self.input_box.SetInsertionPointEnd()
        elif key == wx.WXK_DOWN:
            if self.history_index == -1:
                event.Skip(); return
            if self.history_index < len(self.history) - 1:
                self.history_index += 1
                self.input_box.SetValue(self.history[self.history_index])
            else:
                self.history_index = -1
                self.input_box.SetValue(self.history_draft)
            self.input_box.SetInsertionPointEnd()
        else:
            event.Skip()

    # ── output / announce ─────────────────────────────────────────────────────

    def _append(self, text: str):
        text = _strip_ansi(text)
        self.output_box.AppendText(text)
        self.output_box.ShowPosition(self.output_box.GetLastPosition())

        # Trim top of output box when it grows too large.
        # wx.TE_RICH2 becomes very slow past a few thousand lines.
        if self.output_box.GetNumberOfLines() > MAX_OUTPUT_LINES:
            # Remove the top ~20% of lines to avoid trimming on every append
            trim_to   = MAX_OUTPUT_LINES // 5
            trim_end  = self.output_box.GetLineLength(trim_to)
            # Sum up character positions for all lines to trim
            pos = sum(
                self.output_box.GetLineLength(i) + 1
                for i in range(trim_to)
            )
            self.output_box.Remove(0, pos)

        self._announce_buf += text
        if "\n" in self._announce_buf or len(self._announce_buf) >= ANNOUNCE_CHUNK:
            self._flush_announce()

    def _flush_announce(self):
        chunk = self._announce_buf[:ANNOUNCE_CHUNK]
        self._announce_buf = self._announce_buf[ANNOUNCE_CHUNK:]
        chunk = chunk.strip()
        if not chunk:
            return
        # Speak through NVDA directly — no accessibility-tree tricks needed
        nvda_speak(chunk)

    def _on_poll(self, _event):
        chunks = []
        try:
            while True:
                chunks.append(self.out_q.get_nowait())
        except queue.Empty:
            pass
        if chunks:
            self._append("".join(chunks))
        if self._announce_buf:
            self._flush_announce()

    # ── history ───────────────────────────────────────────────────────────────

    def _push_history(self, cmd: str):
        if not cmd:
            return
        if not self.history or self.history[-1] != cmd:
            self.history.append(cmd)
        self.history_index = -1
        self.history_draft = ""

    # ── run ───────────────────────────────────────────────────────────────────

    def _on_run(self, _event):
        cmd = self.input_box.GetValue().strip()
        if not cmd:
            return
        self.input_box.Clear()
        self._push_history(cmd)

        if self.interactive and self.session and self.session.poll() is None:
            self._send(cmd)
            return

        if self._builtin(cmd):
            return

        self._run_command(cmd)

    def _send(self, text: str):
        try:
            self.session.write(text)
            if not HAS_WINPTY:
                self._append(f"{text}\n")
        except Exception:
            self._append(f"{self.cwd}> ")
            self.interactive = False
            self.session     = None

    def _builtin(self, command: str) -> bool:
        parts = command.strip().split(None, 1)
        cmd   = parts[0].lower()
        if cmd == "cd":
            if len(parts) < 2:
                self._append(f"{self.cwd}\n{self.cwd}> ")
            else:
                target = parts[1].strip()
                new    = os.path.normpath(os.path.join(self.cwd, target))
                if os.path.isdir(new):
                    self.cwd = new
                    self._append(f"Changed directory to: {self.cwd}\n{self.cwd}> ")
                else:
                    self._append(f"cd: Not found: {new}\n{self.cwd}> ")
            return True
        if cmd in ("cls", "clear"):
            self._on_clear(None); return True
        if cmd == "exit":
            self.Close(); return True
        return False

    def _run_command(self, command: str):
        self._append(f"\n> {command}\n")
        is_interactive = _is_interactive_command(command)
        if is_interactive:
            self._append("[Interactive session started — type input below]\n")

        if self.session and self.session.poll() is None:
            self.session.kill()
            self.session = None

        try:
            if HAS_WINPTY:
                try:
                    self.session = _PtySession(command, self.cwd, self.out_q, is_interactive)
                except Exception:
                    # pywinpty failed (old system, ConPTY not available, etc.)
                    # — fall back to plain subprocess silently
                    self.session = _PlainSession(command, self.cwd, self.out_q, is_interactive)
            else:
                self.session = _PlainSession(command, self.cwd, self.out_q, is_interactive)
        except Exception as e:
            self._append(f"[Error: {e}]\n{self.cwd}> ")
            return

        self.interactive = is_interactive
        if not is_interactive:
            self._prompt_pending = True
            threading.Thread(target=self._wait_finish,
                             args=(self.session,), daemon=True).start()

    def _wait_finish(self, session):
        while session.poll() is None:
            time.sleep(0.05)
        time.sleep(0.08)   # let reader thread flush its last lines
        def _done():
            # Only emit the prompt if Ctrl+C hasn't already done so
            if self._prompt_pending:
                self._prompt_pending = False
                self._append(f"{self.cwd}> ")
            self.interactive = False
        wx.CallAfter(_done)

    # ── buttons ───────────────────────────────────────────────────────────────

    def _on_find(self, _event):
        dlg = FindDialog(self, self.output_box)
        dlg.ShowModal()
        dlg.Destroy()
        self.input_box.SetFocus()

    def _on_clear(self, _event):
        self.output_box.Clear()
        self._append(f"{self.cwd}> ")

    def _on_save(self, _event):
        ts   = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
        docs = os.path.join(os.path.expanduser("~"), "Documents")
        if not os.path.isdir(docs):
            docs = os.path.expanduser("~")
        dlg = wx.FileDialog(self, "Save Output As", defaultDir=docs,
                            defaultFile=f"CMDPiper_{ts}.txt",
                            wildcard="Text files (*.txt)|*.txt|All files|*.*",
                            style=wx.FD_SAVE | wx.FD_OVERWRITE_PROMPT)
        if dlg.ShowModal() == wx.ID_OK:
            path = dlg.GetPath()
            try:
                with open(path, "w", encoding="utf-8") as f:
                    f.write(self.output_box.GetValue())
                wx.MessageBox(f"Saved to:\n{path}", "Saved",
                              wx.OK | wx.ICON_INFORMATION, self)
            except Exception as e:
                wx.MessageBox(f"Could not save:\n{e}", "Error",
                              wx.OK | wx.ICON_ERROR, self)
        dlg.Destroy()

    def _on_instructions(self, _event):
        dlg = wx.Dialog(self, title=f"{APP_NAME} — Instructions",
                        style=wx.DEFAULT_DIALOG_STYLE | wx.RESIZE_BORDER)
        s   = wx.BoxSizer(wx.VERTICAL)
        tc  = wx.TextCtrl(dlg, value=INSTRUCTIONS,
                          style=wx.TE_MULTILINE | wx.TE_READONLY | wx.TE_RICH2,
                          size=(580, 460))
        tc.SetName("Instructions text")
        s.Add(tc, proportion=1, flag=wx.EXPAND | wx.ALL, border=8)
        b = wx.Button(dlg, label="&Close", id=wx.ID_OK)
        b.Bind(wx.EVT_BUTTON, lambda e: dlg.EndModal(wx.ID_OK))
        s.Add(b, flag=wx.ALIGN_RIGHT | wx.ALL, border=8)
        dlg.SetSizer(s)
        dlg.Centre()
        tc.SetFocus()
        dlg.ShowModal()
        dlg.Destroy()
        self.input_box.SetFocus()

    def _on_close(self, _event):
        self.poll_timer.Stop()
        if self.session and self.session.poll() is None:
            self.session.kill()
        self.Destroy()


# ─────────────────────────────────────────────────────────────────────────────
#  Entry point
# ─────────────────────────────────────────────────────────────────────────────

class CMDPiperApp(wx.App):
    def OnInit(self):
        frame = CMDPiperFrame()
        frame.Show()
        self.SetTopWindow(frame)
        return True


if __name__ == "__main__":
    CMDPiperApp(redirect=False).MainLoop()