# CMD Piper - NVDA Global Plugin
# Version: 1.1.2
# Author: Vatsal Gautam
# Description: An accessible terminal dialog that pipes command output
#              into a readable, searchable, saveable multiline textbox.
#              Supports interactive sessions (Python REPL, etc.)
#
# v1.1.2 fixes:
#   - Added script category so CMD Piper appears in Input Gestures dialog
#     as its own named section, allowing users to remap the hotkey
#   - Made interactive detection more defensive: explicitly checks that
#     there are zero meaningful arguments, not just zero tokens
#   - Added debug-friendly logging of interactive detection decision

import globalPluginHandler
import gui
import wx
import subprocess
import threading
import os
import queue
import datetime


# ─────────────────────────────────────────────
#  Instructions text
# ─────────────────────────────────────────────

INSTRUCTIONS = """\
CMD Piper v1.1.2 - Accessible Terminal
Author: Vatsal Gautam

HOW TO USE:
- Type any command in the Command box and press Enter or Run.
- Output appears in the Output box above.
- Use Find to search through output text.
- Use Save Output to save the output box contents to a .txt file.
- Use Clear Output to wipe the output box clean.
- Type 'exit' or press Close to close the dialog.

COMMAND HISTORY:
- Press Up arrow in the Command box to recall previous commands.
- Press Down arrow to go forward through history.
- History is kept for the current session only.

BUILT-IN COMMANDS:
- cd <path>   : Change working directory (persists across commands)
- cd          : Show current working directory
- cls / clear : Clear the output box

INTERACTIVE PROGRAMS (Python REPL, Node, etc.):
- Type the program name ALONE with no arguments, e.g. 'python' or 'node'
- 'python script.py' runs a script and is NOT treated as interactive
- The session stays open - keep typing input and see output live
- Type 'exit()' (for Python) or 'exit' to end the session

TIPS:
- Commands run exactly as they would in CMD
- Pipes and redirects work: e.g. 'dir | find "txt"'
- Python output is forced unbuffered so responses appear immediately
- Working directory starts at your user home folder
- Saved files go to your Documents folder by default
- To remap the hotkey: NVDA menu > Preferences > Input Gestures > CMD Piper
"""


# ─────────────────────────────────────────────
#  Interactive program detection
# ─────────────────────────────────────────────

# Programs that launch a REPL when called with NO arguments.
INTERACTIVE_PROGRAMS = {
    "python", "python3", "py",
    "node", "nodejs",
    "irb",        # Ruby REPL
    "lua",
    "fsi",        # F# interactive
    "powershell", "pwsh",
    "cmd",
    "bash", "sh", "zsh",
}


def _is_interactive_command(command):
    """
    Return True ONLY if this command will open an interactive REPL.

    Rules:
      - The first token must be a known interactive program name.
      - There must be NO other tokens (no arguments whatsoever).
      - Exception: 'python -i' is explicitly interactive (-i flag).

    This means 'python test.py', 'python -c "x"', 'node app.js' all
    return False — they run scripts, not a REPL.

    We strip the command of surrounding quotes and whitespace before
    splitting, to handle cases like:  "python"  or  'python'
    """
    # Strip outer quotes that some users might type
    command = command.strip().strip('"').strip("'").strip()

    # Split on whitespace — shlex would be ideal but unavailable in NVDA
    parts = command.split()
    if not parts:
        return False

    # Normalise: lowercase, strip .exe suffix
    name = parts[0].lower()
    if name.endswith(".exe"):
        name = name[:-4]

    if name not in INTERACTIVE_PROGRAMS:
        return False

    # Collect the remaining tokens (arguments)
    args = parts[1:]

    # No arguments at all -> definitely a REPL
    if not args:
        return True

    # python/py -i explicitly requests interactive mode
    if name in ("python", "python3", "py") and args == ["-i"]:
        return True

    # Anything else (flags, filenames, -c, -m, etc.) = script/command
    return False


# ─────────────────────────────────────────────
#  Subprocess environment
# ─────────────────────────────────────────────

def _make_env():
    """
    Return an os.environ copy with extra vars for robust output capture.

    PYTHONUNBUFFERED=1    Python streams output immediately (no buffering)
    PYTHONIOENCODING=utf-8 Forces UTF-8 on Python stdio
    PYTHONUTF8=1          Python 3.7+ UTF-8 mode flag
    """
    env = os.environ.copy()
    env["PYTHONUNBUFFERED"] = "1"
    env["PYTHONIOENCODING"] = "utf-8"
    env["PYTHONUTF8"] = "1"
    return env


def _kill_process_tree(process):
    """
    Force-kill a process AND all its children using taskkill /F /T.

    process.terminate() only kills cmd.exe (the shell=True wrapper),
    leaving the real child process orphaned. taskkill /T kills the
    entire descendant tree at once.
    """
    try:
        subprocess.run(
            ["taskkill", "/F", "/T", "/PID", str(process.pid)],
            creationflags=subprocess.CREATE_NO_WINDOW,
            stdout=subprocess.DEVNULL,
            stderr=subprocess.DEVNULL,
        )
    except Exception:
        try:
            process.kill()
        except Exception:
            pass


# ─────────────────────────────────────────────
#  Main dialog
# ─────────────────────────────────────────────

class CMDPiperDialog(wx.Dialog):
    """
    The accessible terminal dialog.

    Layout:
      Output:  [multiline read-only textbox]
      Command: [single-line input box]
      [Run] [Find] [Save Output] [Clear Output] [Instructions] [Close]
    """

    def __init__(self, parent):
        super().__init__(
            parent,
            title="CMD Piper",
            style=wx.DEFAULT_DIALOG_STYLE | wx.RESIZE_BORDER
        )

        self.cwd = os.path.expanduser("~")
        self.process = None
        self.output_queue = queue.Queue()
        self.interactive = False
        self.reader_thread = None

        # Command history
        self.history = []
        self.history_index = -1
        self.history_saved_input = ""

        self._build_ui()

        self.poll_timer = wx.Timer(self)
        self.Bind(wx.EVT_TIMER, self._on_poll_timer, self.poll_timer)
        self.poll_timer.Start(100)

        self._append_output(f"{self.cwd}> ")
        self.input_box.SetFocus()

    # ── UI ─────────────────────────────────────────────────────────────

    def _build_ui(self):
        main_sizer = wx.BoxSizer(wx.VERTICAL)

        output_label = wx.StaticText(self, label="Output:")
        main_sizer.Add(output_label, flag=wx.LEFT | wx.TOP, border=6)

        self.output_box = wx.TextCtrl(
            self,
            style=wx.TE_MULTILINE | wx.TE_READONLY | wx.TE_RICH2 | wx.HSCROLL,
            size=(-1, 350)
        )
        self.output_box.SetBackgroundColour(wx.Colour(0, 0, 0))
        self.output_box.SetForegroundColour(wx.Colour(204, 204, 204))
        self.output_box.SetName("Output")
        main_sizer.Add(self.output_box, proportion=1, flag=wx.EXPAND | wx.ALL, border=6)

        input_sizer = wx.BoxSizer(wx.HORIZONTAL)
        input_label = wx.StaticText(self, label="Command:")
        input_sizer.Add(input_label, flag=wx.ALIGN_CENTER_VERTICAL | wx.RIGHT, border=5)
        self.input_box = wx.TextCtrl(self, style=wx.TE_PROCESS_ENTER)
        self.input_box.SetName("Command")
        self.input_box.Bind(wx.EVT_TEXT_ENTER, self._on_run)
        self.input_box.Bind(wx.EVT_KEY_DOWN, self._on_input_key_down)
        input_sizer.Add(self.input_box, proportion=1, flag=wx.EXPAND)
        main_sizer.Add(input_sizer, flag=wx.EXPAND | wx.LEFT | wx.RIGHT | wx.BOTTOM, border=6)

        btn_sizer = wx.BoxSizer(wx.HORIZONTAL)
        for label, handler in [
            ("&Run",          self._on_run),
            ("&Find",         self._on_find),
            ("&Save Output",  self._on_save),
            ("C&lear Output", self._on_clear),
            ("&Instructions", self._on_instructions),
        ]:
            btn = wx.Button(self, label=label)
            btn.Bind(wx.EVT_BUTTON, handler)
            btn_sizer.Add(btn, flag=wx.RIGHT, border=5)

        close_btn = wx.Button(self, label="C&lose", id=wx.ID_CLOSE)
        close_btn.Bind(wx.EVT_BUTTON, self._on_close)
        btn_sizer.Add(close_btn)

        main_sizer.Add(btn_sizer, flag=wx.ALIGN_RIGHT | wx.ALL, border=6)
        self.SetSizer(main_sizer)
        self.SetSize((700, 520))
        self.Centre()
        self.Bind(wx.EVT_CLOSE, self._on_close)

    # ── Output helpers ─────────────────────────────────────────────────

    def _append_output(self, text):
        self.output_box.AppendText(text)
        self.output_box.ShowPosition(self.output_box.GetLastPosition())

    def _queue_output(self, text):
        self.output_queue.put(text)

    def _on_poll_timer(self, event):
        lines = []
        try:
            while True:
                lines.append(self.output_queue.get_nowait())
        except queue.Empty:
            pass
        if lines:
            self._append_output("".join(lines))

    # ── History ────────────────────────────────────────────────────────

    def _on_input_key_down(self, event):
        key = event.GetKeyCode()
        if key == wx.WXK_UP:
            if not self.history:
                event.Skip()
                return
            if self.history_index == -1:
                self.history_saved_input = self.input_box.GetValue()
                self.history_index = len(self.history) - 1
            elif self.history_index > 0:
                self.history_index -= 1
            self.input_box.SetValue(self.history[self.history_index])
            self.input_box.SetInsertionPointEnd()
        elif key == wx.WXK_DOWN:
            if self.history_index == -1:
                event.Skip()
                return
            if self.history_index < len(self.history) - 1:
                self.history_index += 1
                self.input_box.SetValue(self.history[self.history_index])
                self.input_box.SetInsertionPointEnd()
            else:
                self.history_index = -1
                self.input_box.SetValue(self.history_saved_input)
                self.input_box.SetInsertionPointEnd()
        else:
            event.Skip()

    def _add_to_history(self, command):
        if not command:
            return
        if not self.history or self.history[-1] != command:
            self.history.append(command)
        self.history_index = -1
        self.history_saved_input = ""

    # ── Run ────────────────────────────────────────────────────────────

    def _on_run(self, event):
        command = self.input_box.GetValue().strip()
        if not command:
            return
        self.input_box.Clear()
        self._add_to_history(command)

        # In interactive mode: send EVERYTHING to the process.
        # Do NOT pass through builtins — 'exit', 'cd', etc. all belong
        # to the running program, not to CMD Piper.
        if self.interactive and self.process and self.process.poll() is None:
            self._send_to_process(command)
            return

        if self._handle_builtins(command):
            return

        self._run_command(command)

    def _send_to_process(self, text):
        try:
            self.process.stdin.write(text + "\n")
            self.process.stdin.flush()
            self._append_output(f">>> {text}\n")
        except (OSError, BrokenPipeError):
            self._append_output("[Session ended]\n")
            self.interactive = False
            self.process = None

    def _handle_builtins(self, command):
        """Only called when NOT in an interactive session."""
        parts = command.strip().split(None, 1)
        cmd = parts[0].lower()

        if cmd == "cd":
            if len(parts) < 2:
                self._append_output(f"{self.cwd}\n{self.cwd}> ")
            else:
                target = parts[1].strip()
                new_path = os.path.normpath(os.path.join(self.cwd, target))
                if os.path.isdir(new_path):
                    self.cwd = new_path
                    self._append_output(f"Changed directory to: {self.cwd}\n{self.cwd}> ")
                else:
                    self._append_output(f"cd: Directory not found: {new_path}\n{self.cwd}> ")
            return True

        if cmd in ("cls", "clear"):
            self._on_clear(None)
            return True

        if cmd == "exit":
            self._on_close(None)
            return True

        return False

    def _run_command(self, command):
        self._append_output(f"\n> {command}\n")

        is_interactive = _is_interactive_command(command)

        # Show what the detection decided, so user can see it in output
        # if something seems wrong. Remove this line after testing if desired.
        if is_interactive:
            self._append_output("[Interactive session started - type your input below]\n")

        try:
            self.process = subprocess.Popen(
                command,
                shell=True,
                stdout=subprocess.PIPE,
                stderr=subprocess.STDOUT,
                stdin=subprocess.PIPE,
                cwd=self.cwd,
                creationflags=subprocess.CREATE_NO_WINDOW,
                text=True,
                encoding="utf-8",
                errors="replace",
                bufsize=0,
                env=_make_env()
            )
        except Exception as e:
            self._append_output(f"[Error launching command: {e}]\n{self.cwd}> ")
            return

        self.interactive = is_interactive

        self.reader_thread = threading.Thread(
            target=self._read_output,
            args=(self.process, is_interactive),
            daemon=True
        )
        self.reader_thread.start()

    def _read_output(self, process, is_interactive):
        try:
            if is_interactive:
                # Character-by-character so prompts like '>>> ' (no newline)
                # appear immediately instead of blocking on the next newline
                buf = ""
                while True:
                    ch = process.stdout.read(1)
                    if not ch:
                        break
                    buf += ch
                    if ch == "\n" or buf.endswith(">>> ") or buf.endswith("... "):
                        self._queue_output(buf)
                        buf = ""
                if buf:
                    self._queue_output(buf)
            else:
                for line in process.stdout:
                    self._queue_output(line)
        except Exception:
            pass

        process.wait()
        return_code = process.returncode

        if is_interactive:
            self._queue_output("\n[Interactive session ended]\n")
            wx.CallAfter(setattr, self, "interactive", False)
            wx.CallAfter(setattr, self, "process", None)
        else:
            if return_code != 0:
                self._queue_output(f"[Process exited with code {return_code}]\n")

        self._queue_output(f"\n{self.cwd}> ")

    # ── Save ───────────────────────────────────────────────────────────

    def _on_save(self, event):
        timestamp = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
        default_name = f"CMDPiper_{timestamp}.txt"
        documents_path = os.path.join(os.path.expanduser("~"), "Documents")
        if not os.path.isdir(documents_path):
            documents_path = os.path.expanduser("~")

        dlg = wx.FileDialog(
            self,
            message="Save Output As",
            defaultDir=documents_path,
            defaultFile=default_name,
            wildcard="Text files (*.txt)|*.txt|All files (*.*)|*.*",
            style=wx.FD_SAVE | wx.FD_OVERWRITE_PROMPT
        )
        if dlg.ShowModal() == wx.ID_OK:
            path = dlg.GetPath()
            try:
                with open(path, "w", encoding="utf-8") as f:
                    f.write(self.output_box.GetValue())
                wx.MessageBox(f"Output saved to:\n{path}", "Saved",
                              wx.OK | wx.ICON_INFORMATION, self)
            except Exception as e:
                wx.MessageBox(f"Could not save file:\n{e}", "Save Error",
                              wx.OK | wx.ICON_ERROR, self)
        dlg.Destroy()

    # ── Instructions ───────────────────────────────────────────────────

    def _on_instructions(self, event):
        dlg = wx.Dialog(self, title="CMD Piper - Instructions",
                        style=wx.DEFAULT_DIALOG_STYLE | wx.RESIZE_BORDER)
        sizer = wx.BoxSizer(wx.VERTICAL)
        text = wx.TextCtrl(dlg, value=INSTRUCTIONS,
                           style=wx.TE_MULTILINE | wx.TE_READONLY | wx.TE_RICH2,
                           size=(500, 400))
        text.SetName("Instructions")
        sizer.Add(text, proportion=1, flag=wx.EXPAND | wx.ALL, border=8)
        close_btn = wx.Button(dlg, label="&Close", id=wx.ID_OK)
        close_btn.Bind(wx.EVT_BUTTON, lambda e: dlg.EndModal(wx.ID_OK))
        sizer.Add(close_btn, flag=wx.ALIGN_RIGHT | wx.ALL, border=8)
        dlg.SetSizer(sizer)
        dlg.Centre()
        text.SetFocus()
        dlg.ShowModal()
        dlg.Destroy()

    # ── Find ──────────────────────────────────────────────────────────

    def _on_find(self, event):
        dlg = FindDialog(self, self.output_box)
        dlg.ShowModal()
        dlg.Destroy()

    # ── Clear ─────────────────────────────────────────────────────────

    def _on_clear(self, event):
        self.output_box.Clear()
        self._append_output(f"{self.cwd}> ")

    # ── Close ─────────────────────────────────────────────────────────

    def _on_close(self, event):
        self.poll_timer.Stop()
        if self.process and self.process.poll() is None:
            _kill_process_tree(self.process)
        self.Destroy()


# ─────────────────────────────────────────────
#  Find dialog
# ─────────────────────────────────────────────

class FindDialog(wx.Dialog):

    def __init__(self, parent, target_textctrl):
        super().__init__(parent, title="Find in Output", style=wx.DEFAULT_DIALOG_STYLE)
        self.target = target_textctrl
        self.last_pos = 0

        sizer = wx.BoxSizer(wx.VERTICAL)
        row = wx.BoxSizer(wx.HORIZONTAL)
        row.Add(wx.StaticText(self, label="Find:"),
                flag=wx.ALIGN_CENTER_VERTICAL | wx.RIGHT, border=5)
        self.search_box = wx.TextCtrl(self, style=wx.TE_PROCESS_ENTER)
        self.search_box.SetName("Search term")
        self.search_box.Bind(wx.EVT_TEXT_ENTER, self._on_find_next)
        row.Add(self.search_box, proportion=1, flag=wx.EXPAND)
        sizer.Add(row, flag=wx.EXPAND | wx.ALL, border=8)

        self.case_check = wx.CheckBox(self, label="&Case sensitive")
        sizer.Add(self.case_check, flag=wx.LEFT | wx.BOTTOM, border=8)

        btn_row = wx.BoxSizer(wx.HORIZONTAL)
        find_btn = wx.Button(self, label="&Find Next")
        find_btn.Bind(wx.EVT_BUTTON, self._on_find_next)
        btn_row.Add(find_btn, flag=wx.RIGHT, border=5)
        close_btn = wx.Button(self, label="&Close", id=wx.ID_CLOSE)
        close_btn.Bind(wx.EVT_BUTTON, lambda e: self.Close())
        btn_row.Add(close_btn)
        sizer.Add(btn_row, flag=wx.ALIGN_RIGHT | wx.ALL, border=8)

        self.SetSizer(sizer)
        self.Fit()
        self.Centre()
        self.search_box.SetFocus()

    def _on_find_next(self, event):
        term = self.search_box.GetValue()
        if not term:
            return
        full_text = self.target.GetValue()
        case_sensitive = self.case_check.IsChecked()
        search_in = full_text if case_sensitive else full_text.lower()
        search_term = term if case_sensitive else term.lower()
        pos = search_in.find(search_term, self.last_pos)
        if pos == -1 and self.last_pos > 0:
            pos = search_in.find(search_term, 0)
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


# ─────────────────────────────────────────────
#  Global Plugin
# ─────────────────────────────────────────────

class GlobalPlugin(globalPluginHandler.GlobalPlugin):

    def __init__(self):
        super().__init__()
        self._dialog = None

    def script_openCMDPiper(self, gesture):
        try:
            if self._dialog and self._dialog.IsShown():
                self._dialog.Raise()
                self._dialog.input_box.SetFocus()
                return
        except RuntimeError:
            self._dialog = None

        self._dialog = CMDPiperDialog(gui.mainFrame)
        self._dialog.Show()

    # ── These two lines are what make CMD Piper appear as its own
    # named category in NVDA > Preferences > Input Gestures.
    # Without 'category', NVDA buries the script in an uncategorised
    # list and users cannot easily find or remap the hotkey.
    script_openCMDPiper.__doc__ = (
        "Opens the CMD Piper accessible terminal window."
    )
    script_openCMDPiper.category = "CMD Piper"

    __gestures = {
        "kb:NVDA+shift+c": "openCMDPiper",
    }