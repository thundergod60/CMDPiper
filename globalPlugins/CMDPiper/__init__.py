# CMD Piper - NVDA Global Plugin
# Version: 1.1
# Author: Vatsal Gautam
# Description: An accessible terminal dialog that pipes command output
#              into a readable, searchable, saveable multiline textbox.
#              Supports interactive sessions (Python REPL, etc.)
#
# v1.1 changes:
#   - Save Output button (saves to Documents with timestamped filename)
#   - Up/Down arrow command history in the input box
#   - PTY robustness: forces unbuffered output so Python REPL and other
#     programs stream output immediately instead of buffering it silently
#   - Cleaner prompt display after each command

import globalPluginHandler  # Base class for all global plugins
import gui                  # NVDA's GUI module (built on wx)
import wx                   # wxPython - the GUI toolkit NVDA uses
import subprocess           # To run commands and capture output
import threading            # So commands don't freeze the UI
import os                   # For paths, environment variables
import queue                # Thread-safe output passing between threads
import datetime             # For timestamped save filenames


# ─────────────────────────────────────────────
#  Instructions text
# ─────────────────────────────────────────────

INSTRUCTIONS = """\
CMD Piper v1.1 - Accessible Terminal
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
- Just type the program name, e.g. 'python' or 'node'
- The session stays open - keep typing input and see output live
- Type 'exit()' (for Python) or 'exit' to end the session

TIPS:
- Commands run exactly as they would in CMD
- Pipes and redirects work: e.g. 'dir | find "txt"'
- Python output is forced to be unbuffered, so REPL responses
  appear immediately rather than waiting for the process to end
- Working directory starts at your user home folder
- Saved files go to your Documents folder by default
"""


# ─────────────────────────────────────────────
#  Build a robust environment for subprocesses
# ─────────────────────────────────────────────

def _make_env():
    """
    Build an environment dictionary for subprocess calls.

    The key additions over a plain os.environ.copy():

    PYTHONUNBUFFERED=1
        Tells Python NOT to buffer stdout/stderr. Without this, when you
        run 'python' interactively, Python buffers its output and you see
        nothing until the buffer fills up or the process ends. With this
        set, every print() and every REPL response appears immediately.

    PYTHONIOENCODING=utf-8
        Forces Python's stdin/stdout/stderr to use UTF-8. Without this,
        on some Windows systems Python defaults to cp1252 or cp850, which
        causes UnicodeDecodeError when the output contains non-ASCII text.

    PYTHONUTF8=1
        Python 3.7+ UTF-8 mode flag - same effect as PYTHONIOENCODING
        but also affects file I/O. Belt-and-suspenders approach.

    These variables are harmless to non-Python programs - they are simply
    ignored by cmd.exe, node, etc.
    """
    env = os.environ.copy()
    env["PYTHONUNBUFFERED"] = "1"
    env["PYTHONIOENCODING"] = "utf-8"
    env["PYTHONUTF8"] = "1"
    return env


# ─────────────────────────────────────────────
#  The Dialog Window
# ─────────────────────────────────────────────

class CMDPiperDialog(wx.Dialog):
    """
    The main accessible terminal dialog.

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

        # ── Terminal state ─────────────────────────────────────────────
        self.cwd = os.path.expanduser("~")
        self.process = None
        self.output_queue = queue.Queue()
        self.interactive = False
        self.reader_thread = None

        # ── Command history ────────────────────────────────────────────
        # A list of previously run commands, oldest first.
        # history_index tracks where we are when the user presses Up/Down.
        # -1 means "not browsing history" (showing current input).
        self.history = []
        self.history_index = -1
        # When the user starts browsing history, we save whatever they
        # had typed in the input box so we can restore it if they press
        # Down past the end of history.
        self.history_saved_input = ""

        # ── Build UI ───────────────────────────────────────────────────
        self._build_ui()

        # Timer polls the output queue every 100ms
        self.poll_timer = wx.Timer(self)
        self.Bind(wx.EVT_TIMER, self._on_poll_timer, self.poll_timer)
        self.poll_timer.Start(100)

        # Show working directory prompt as the first line
        self._append_output(f"{self.cwd}> ")

        self.input_box.SetFocus()

    # ── UI Construction ────────────────────────────────────────────────

    def _build_ui(self):
        main_sizer = wx.BoxSizer(wx.VERTICAL)

        # Output label
        output_label = wx.StaticText(self, label="Output:")
        main_sizer.Add(output_label, flag=wx.LEFT | wx.TOP, border=6)

        # Output textbox
        self.output_box = wx.TextCtrl(
            self,
            style=wx.TE_MULTILINE | wx.TE_READONLY | wx.TE_RICH2 | wx.HSCROLL,
            size=(-1, 350)
        )
        self.output_box.SetBackgroundColour(wx.Colour(0, 0, 0))
        self.output_box.SetForegroundColour(wx.Colour(204, 204, 204))
        self.output_box.SetName("Output")
        main_sizer.Add(self.output_box, proportion=1, flag=wx.EXPAND | wx.ALL, border=6)

        # Input row
        input_sizer = wx.BoxSizer(wx.HORIZONTAL)
        input_label = wx.StaticText(self, label="Command:")
        input_sizer.Add(input_label, flag=wx.ALIGN_CENTER_VERTICAL | wx.RIGHT, border=5)

        self.input_box = wx.TextCtrl(self, style=wx.TE_PROCESS_ENTER)
        self.input_box.SetName("Command")
        self.input_box.Bind(wx.EVT_TEXT_ENTER, self._on_run)

        # Up/Down arrow keys for command history
        # We catch them here before wx tries to move focus
        self.input_box.Bind(wx.EVT_KEY_DOWN, self._on_input_key_down)

        input_sizer.Add(self.input_box, proportion=1, flag=wx.EXPAND)
        main_sizer.Add(input_sizer, flag=wx.EXPAND | wx.LEFT | wx.RIGHT | wx.BOTTOM, border=6)

        # Button row
        btn_sizer = wx.BoxSizer(wx.HORIZONTAL)

        self.run_btn = wx.Button(self, label="&Run")
        self.run_btn.Bind(wx.EVT_BUTTON, self._on_run)
        btn_sizer.Add(self.run_btn, flag=wx.RIGHT, border=5)

        self.find_btn = wx.Button(self, label="&Find")
        self.find_btn.Bind(wx.EVT_BUTTON, self._on_find)
        btn_sizer.Add(self.find_btn, flag=wx.RIGHT, border=5)

        self.save_btn = wx.Button(self, label="&Save Output")
        self.save_btn.Bind(wx.EVT_BUTTON, self._on_save)
        btn_sizer.Add(self.save_btn, flag=wx.RIGHT, border=5)

        self.clear_btn = wx.Button(self, label="C&lear Output")
        self.clear_btn.Bind(wx.EVT_BUTTON, self._on_clear)
        btn_sizer.Add(self.clear_btn, flag=wx.RIGHT, border=5)

        self.instructions_btn = wx.Button(self, label="&Instructions")
        self.instructions_btn.Bind(wx.EVT_BUTTON, self._on_instructions)
        btn_sizer.Add(self.instructions_btn, flag=wx.RIGHT, border=5)

        self.close_btn = wx.Button(self, label="C&lose", id=wx.ID_CLOSE)
        self.close_btn.Bind(wx.EVT_BUTTON, self._on_close)
        btn_sizer.Add(self.close_btn)

        main_sizer.Add(btn_sizer, flag=wx.ALIGN_RIGHT | wx.ALL, border=6)

        self.SetSizer(main_sizer)
        self.SetSize((700, 520))
        self.Centre()
        self.Bind(wx.EVT_CLOSE, self._on_close)

    # ── Output helpers ─────────────────────────────────────────────────

    def _append_output(self, text):
        """Append text to the output box. Must run on the main thread."""
        self.output_box.AppendText(text)
        self.output_box.ShowPosition(self.output_box.GetLastPosition())

    def _queue_output(self, text):
        """Thread-safe: put text in the queue for the main thread to display."""
        self.output_queue.put(text)

    def _on_poll_timer(self, event):
        """Runs every 100ms - drains the output queue into the textbox."""
        lines = []
        try:
            while True:
                lines.append(self.output_queue.get_nowait())
        except queue.Empty:
            pass
        if lines:
            self._append_output("".join(lines))

    # ── Command history ────────────────────────────────────────────────

    def _on_input_key_down(self, event):
        """
        Handle Up/Down arrow keys in the input box for command history.

        Up arrow   -> go back in history (older commands)
        Down arrow -> go forward in history (newer commands)

        Any other key is passed through normally via event.Skip().
        """
        key = event.GetKeyCode()

        if key == wx.WXK_UP:
            # Nothing in history? Do nothing.
            if not self.history:
                event.Skip()
                return

            # If we are not currently browsing, save what the user typed
            if self.history_index == -1:
                self.history_saved_input = self.input_box.GetValue()
                # Start from the most recent command
                self.history_index = len(self.history) - 1
            elif self.history_index > 0:
                self.history_index -= 1
            # else: already at oldest command, stay there

            self.input_box.SetValue(self.history[self.history_index])
            # Move cursor to end of text
            self.input_box.SetInsertionPointEnd()

        elif key == wx.WXK_DOWN:
            if self.history_index == -1:
                # Not browsing - nothing to do
                event.Skip()
                return

            if self.history_index < len(self.history) - 1:
                self.history_index += 1
                self.input_box.SetValue(self.history[self.history_index])
                self.input_box.SetInsertionPointEnd()
            else:
                # Went past the end of history - restore saved input
                self.history_index = -1
                self.input_box.SetValue(self.history_saved_input)
                self.input_box.SetInsertionPointEnd()

        else:
            # Not a history key - let wx handle it normally
            event.Skip()

    def _add_to_history(self, command):
        """
        Add a command to history.
        Avoids consecutive duplicates (same as bash behaviour).
        Resets the history browsing index.
        """
        if not command:
            return
        # Don't add if it's the same as the last command
        if self.history and self.history[-1] == command:
            pass
        else:
            self.history.append(command)
        # Always reset browsing position after running a command
        self.history_index = -1
        self.history_saved_input = ""

    # ── Command execution ──────────────────────────────────────────────

    def _on_run(self, event):
        """Called when user presses Enter or clicks Run."""
        command = self.input_box.GetValue().strip()
        if not command:
            return

        self.input_box.Clear()

        # If in an interactive session, send directly to the running process
        if self.interactive and self.process and self.process.poll() is None:
            # Still add to history so the user can recall what they typed
            self._add_to_history(command)
            self._send_to_process(command)
            return

        # Add to history before running
        self._add_to_history(command)

        if self._handle_builtins(command):
            return

        self._run_command(command)

    def _send_to_process(self, text):
        """Send a line of input to a running interactive process."""
        try:
            self.process.stdin.write(text + "\n")
            self.process.stdin.flush()
            self._append_output(f">>> {text}\n")
        except (OSError, BrokenPipeError):
            self._append_output("[Session ended]\n")
            self.interactive = False
            self.process = None

    def _handle_builtins(self, command):
        """
        Handle built-in commands that can't go through subprocess.
        Returns True if handled, False to pass on to subprocess.
        """
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

        if cmd == "exit" and not self.interactive:
            self._on_close(None)
            return True

        return False

    def _run_command(self, command):
        """
        Run a command in a subprocess.

        Interactive program detection: if the command starts with a known
        interactive program name (python, node, etc.) we keep stdin open
        and mark self.interactive = True so subsequent input is piped in.

        For ALL processes we use _make_env() which sets PYTHONUNBUFFERED,
        PYTHONIOENCODING, and PYTHONUTF8. This forces Python (and other
        programs that respect these vars) to stream output line-by-line
        immediately rather than buffering it until the process ends.
        """
        self._append_output(f"\n> {command}\n")

        interactive_programs = [
            "python", "python3", "py",
            "node", "nodejs",
            "irb",        # Ruby REPL
            "lua",
            "fsi",        # F# interactive
            "powershell", "pwsh",
            "cmd",
            "bash", "sh", "zsh",
        ]
        first_word = command.strip().split()[0].lower().replace(".exe", "")
        is_interactive = first_word in interactive_programs

        try:
            self.process = subprocess.Popen(
                command,
                shell=True,
                stdout=subprocess.PIPE,
                stderr=subprocess.STDOUT,   # Merge stderr into stdout
                stdin=subprocess.PIPE,
                cwd=self.cwd,
                creationflags=subprocess.CREATE_NO_WINDOW,
                text=True,
                encoding="utf-8",
                errors="replace",
                bufsize=0,                  # 0 = fully unbuffered (was 1 before)
                env=_make_env()             # Includes PYTHONUNBUFFERED etc.
            )
        except Exception as e:
            self._append_output(f"[Error launching command: {e}]\n{self.cwd}> ")
            return

        if is_interactive:
            self.interactive = True
            self._append_output("[Interactive session started - type your input below]\n")
        else:
            self.interactive = False

        self.reader_thread = threading.Thread(
            target=self._read_output,
            args=(self.process, is_interactive),
            daemon=True
        )
        self.reader_thread.start()

    def _read_output(self, process, is_interactive):
        """
        Background thread: reads stdout from the process line by line
        and puts each line into the queue for the main thread to display.

        We read character-by-character for interactive sessions so that
        partial lines (like a REPL prompt ">>> " with no newline) appear
        immediately rather than waiting for a newline that never comes.
        """
        try:
            if is_interactive:
                # Character-by-character read so prompts (>>>) appear instantly
                buf = ""
                while True:
                    ch = process.stdout.read(1)
                    if not ch:
                        break
                    buf += ch
                    # Flush the buffer on newline OR when we see a REPL prompt
                    if ch == "\n" or buf.endswith(">>> ") or buf.endswith("... "):
                        self._queue_output(buf)
                        buf = ""
                # Flush anything remaining
                if buf:
                    self._queue_output(buf)
            else:
                # Line-by-line is fine for non-interactive commands
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

    # ── Save Output ────────────────────────────────────────────────────

    def _on_save(self, event):
        """
        Save the contents of the output box to a .txt file.

        Default location: user's Documents folder.
        Default filename: CMDPiper_YYYYMMDD_HHMMSS.txt
        The file dialog lets the user rename or pick a different location.
        """
        # Build a timestamped default filename
        timestamp = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
        default_name = f"CMDPiper_{timestamp}.txt"

        # Documents folder path
        documents_path = os.path.join(os.path.expanduser("~"), "Documents")
        # Fallback if Documents doesn't exist for some reason
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
                content = self.output_box.GetValue()
                with open(path, "w", encoding="utf-8") as f:
                    f.write(content)
                wx.MessageBox(
                    f"Output saved to:\n{path}",
                    "Saved",
                    wx.OK | wx.ICON_INFORMATION,
                    self
                )
            except Exception as e:
                wx.MessageBox(
                    f"Could not save file:\n{e}",
                    "Save Error",
                    wx.OK | wx.ICON_ERROR,
                    self
                )

        dlg.Destroy()

    # ── Instructions ───────────────────────────────────────────────────

    def _on_instructions(self, event):
        """Show a scrollable instructions dialog."""
        dlg = wx.Dialog(
            self,
            title="CMD Piper - Instructions",
            style=wx.DEFAULT_DIALOG_STYLE | wx.RESIZE_BORDER
        )
        sizer = wx.BoxSizer(wx.VERTICAL)

        text = wx.TextCtrl(
            dlg,
            value=INSTRUCTIONS,
            style=wx.TE_MULTILINE | wx.TE_READONLY | wx.TE_RICH2,
            size=(500, 400)
        )
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
        """Open the Find dialog."""
        dlg = FindDialog(self, self.output_box)
        dlg.ShowModal()
        dlg.Destroy()

    # ── Clear ─────────────────────────────────────────────────────────

    def _on_clear(self, event):
        """Clear the output textbox."""
        self.output_box.Clear()
        self._append_output(f"{self.cwd}> ")

    # ── Close ─────────────────────────────────────────────────────────

    def _on_close(self, event):
        """Cleanly shut down any running process before closing."""
        self.poll_timer.Stop()

        if self.process and self.process.poll() is None:
            try:
                self.process.terminate()
                self.process.wait(timeout=2)
            except Exception:
                try:
                    self.process.kill()
                except Exception:
                    pass

        self.Destroy()


# ─────────────────────────────────────────────
#  Find Dialog
# ─────────────────────────────────────────────

class FindDialog(wx.Dialog):
    """Simple Find dialog with wrap-around and case-sensitive support."""

    def __init__(self, parent, target_textctrl):
        super().__init__(parent, title="Find in Output", style=wx.DEFAULT_DIALOG_STYLE)
        self.target = target_textctrl
        self.last_pos = 0

        sizer = wx.BoxSizer(wx.VERTICAL)

        row = wx.BoxSizer(wx.HORIZONTAL)
        row.Add(
            wx.StaticText(self, label="Find:"),
            flag=wx.ALIGN_CENTER_VERTICAL | wx.RIGHT, border=5
        )
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

        # Wrap around if not found from current position
        if pos == -1 and self.last_pos > 0:
            pos = search_in.find(search_term, 0)
            if pos != -1:
                wx.MessageBox(
                    "Search wrapped to beginning.", "Find",
                    wx.OK | wx.ICON_INFORMATION, self
                )

        if pos == -1:
            wx.MessageBox(
                f'"{term}" not found.', "Find",
                wx.OK | wx.ICON_INFORMATION, self
            )
            self.last_pos = 0
        else:
            self.target.SetSelection(pos, pos + len(term))
            self.target.ShowPosition(pos)
            self.target.SetFocus()
            self.last_pos = pos + len(term)


# ─────────────────────────────────────────────
#  Global Plugin Entry Point
# ─────────────────────────────────────────────

class GlobalPlugin(globalPluginHandler.GlobalPlugin):
    """NVDA Global Plugin - registers hotkey and opens CMD Piper dialog."""

    def __init__(self):
        super().__init__()
        self._dialog = None

    def script_openCMDPiper(self, gesture):
        """Open (or focus) the CMD Piper dialog."""
        try:
            if self._dialog and self._dialog.IsShown():
                self._dialog.Raise()
                self._dialog.input_box.SetFocus()
                return
        except RuntimeError:
            # Dialog widget was destroyed - clear reference and open fresh
            self._dialog = None

        self._dialog = CMDPiperDialog(gui.mainFrame)
        self._dialog.Show()

    script_openCMDPiper.__doc__ = "Open CMD Piper accessible terminal"

    __gestures = {
        "kb:NVDA+shift+c": "openCMDPiper",
    }