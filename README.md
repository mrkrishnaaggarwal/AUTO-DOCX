# Auto-DOCX

A professional Python tool that executes Python scripts and Jupyter notebooks, automatically generating formatted Word documents with output, images, and markdown content.

## Features

- ✅ Execute Python scripts (`.py`) and Jupyter notebooks (`.ipynb`)
- ✅ Capture stdout output and matplotlib images
- ✅ Render markdown cells from notebooks
- ✅ Generate professional Word documents with monospaced formatting
- ✅ Include source code by default (use `--no-source` to exclude)
- ✅ Save and reuse Python environment and roll number settings
- ✅ Command-line interface with flexible options

## Installation

### From PyPI

```bash
pip install auto-docx
```

### From source

```bash
# Clone the repository
cd AUTO_DOCX

# Install in development mode
pip install -e .
```

## Usage

### Basic Usage

```bash
# Execute a Python script
auto-docx my_script.py

# Execute a Jupyter notebook
auto-docx my_notebook.ipynb
```

### Command-Line Options

```bash
auto-docx <script_path> [options]

Arguments:
  script_path              Path to the Python file (.py) or notebook (.ipynb)

Options:
  -o, --output PATH        Custom output path for the Word document
  --no-source              Exclude source code from the document
  --timeout SECONDS        Timeout for execution (default: 300)
  -v, --verbose            Enable verbose output
  -r, --roll-no ROLL       Roll number to include in document header
  --save-roll              Save roll number as default for future runs
  --list-envs              List available Python environments
  --env ENV                Select environment by name or index
  --save-env               Save environment as default for future runs
  --python PATH            Path to Python executable
  -h, --help               Show help message
```

### Examples

```bash
# Basic execution
auto-docx my_script.py

# Execute notebook with custom output
auto-docx notebook.ipynb -o results/output.docx

# Set roll number and save for future use
auto-docx script.py --roll-no 123456789 --save-roll

# Use specific Python environment and save it
auto-docx script.py --env 3 --save-env

# List available environments
auto-docx --list-envs

# Exclude source code, set timeout
auto-docx my_script.py --no-source --timeout 60
```

## Output Format

The generated Word document includes:

1. **Header** - Document title (filename) and roll number
2. **Source Code** - The script/notebook source code (included by default)
3. **Output** - Captured stdout, images, and rendered markdown cells

All code sections use **Courier New** monospaced font for proper formatting.

## Configuration

Settings are saved to `~/.auto_docx_config.json`:

```json
{
  "env": "3",
  "roll_no": "123456789"
}
```

Once saved with `--save-env` or `--save-roll`, these values are automatically used in future runs.

## Project Structure

```
AUTO_DOCX/
├── pyproject.toml          # Project configuration
├── requirements.txt        # Dependencies
├── README.md               # This file
└── src/
    └── auto_docx/
        ├── __init__.py         # Package initialization
        ├── main.py             # Entry point and CLI
        ├── executor.py         # Script execution logic
        ├── document.py         # Word document generation
        └── notebook_runner.py  # Jupyter notebook execution
```

## License

MIT License - See LICENSE file for details.
