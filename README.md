# CMD Piper

**Author:** Vatsal Gautam

## Overview

CMD Piper is an NVDA add-on that allows you to run command-line (CMD) commands in a fully accessible interface. Instead of interacting directly with the command prompt, all input and output are handled through accessible text boxes.

## Features

- Run CMD commands from an accessible interface
- Output displayed in a readable textbox
- Built-in search (find text in output)
- Save output for logging or later use
- Command history (similar to a terminal)

## Installation

1. Download the add-on from:
   https://www.github.com/thundergod60/CMDPiper

2. Install the `.nvda-addon` file using NVDA:
   - NVDA + N → Tools → Manage Add-ons → Install

3. Restart NVDA

## Usage

### Opening the interface

Press:

NVDA + Shift + C

### Interface components

- **Command input box**  
  Type your command and press Enter or use the Run button.

- **Run button**  
  Executes the command (Alt + R)

- **Output box**  
  Displays command output and command history

- **Find**  
  Search for specific text in the output

- **Clear Output**  
  Clears all output

- **Instructions**  
  Displays usage instructions

- **Close**  
  Closes the window

## Keyboard Shortcuts

- **NVDA + Shift + C** — Open CMD Piper
- **Alt + R** — Run command
- **Alt + F** — Open find dialog
- **Alt + N** — Find next occurrence
- **Alt + L** — Close window *(or clarify if this is different)*

## Limitations

- Interactive programs (e.g., Python REPL) may not work correctly
- TUI-based applications (e.g., Vim, Nano) are not supported
- Some programs that redraw the screen may not display output properly

## Credits

Developed by Vatsal Gautam

Contact: vatshalgamer@gmail.com
