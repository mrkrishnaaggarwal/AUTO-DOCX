# #!/usr/bin/env python3
# """
# Auto-DOCX: Main entry point and command-line interface.

# Execute Python scripts and generate Word documents with their output.
# """

# import argparse
# import json
# import sys
# from pathlib import Path

# from .executor import ScriptExecutor

# CONFIG_FILE = Path.home() / ".auto_docx_config.json"


# def load_config() -> dict:
#     """Load saved configuration."""
#     if CONFIG_FILE.exists():
#         try:
#             return json.loads(CONFIG_FILE.read_text())
#         except Exception:
#             return {}
#     return {}


# def save_config(config: dict) -> None:
#     """Save configuration to file."""
#     CONFIG_FILE.write_text(json.dumps(config, indent=2))
# from .document import DocumentGenerator


# def create_parser() -> argparse.ArgumentParser:
#     """Create and configure the argument parser."""
#     parser = argparse.ArgumentParser(
#         prog="auto-docx",
#         description="Execute a Python script and log its output to a Word document.",
#         epilog="Example: auto-docx my_script.py -o output.docx --include-source",
#         formatter_class=argparse.RawDescriptionHelpFormatter,
#     )
    
#     parser.add_argument(
#         "script_path",
#         type=str,
#         nargs="?",
#         help="Path to the Python file (.py) or notebook (.ipynb) to execute",
#     )
    
#     parser.add_argument(
#         "-o", "--output",
#         type=str,
#         default=None,
#         metavar="PATH",
#         help="Custom output path for the Word document (default: <script>_output.docx)",
#     )
    
#     parser.add_argument(
#         "--no-source",
#         action="store_true",
#         help="Exclude the script's source code from the document (source is included by default)",
#     )

#     parser.add_argument(
#         "--list-envs",
#         action="store_true",
#         help="List available Python environments and exit",
#     )

#     parser.add_argument(
#         "--env",
#         type=str,
#         default=None,
#         metavar="ENV",
#         help="Select an environment by name or index from --list-envs",
#     )

#     parser.add_argument(
#         "--save-env",
#         action="store_true",
#         help="Save the current --env as the default for future runs",
#     )

#     parser.add_argument(
#         "--python",
#         type=str,
#         default=None,
#         metavar="PATH",
#         help="Path to the Python executable to use for execution",
#     )
    
#     parser.add_argument(
#         "--timeout",
#         type=int,
#         default=300,
#         metavar="SECONDS",
#         help="Timeout for script execution in seconds (default: 300)",
#     )
    
#     parser.add_argument(
#         "-v", "--verbose",
#         action="store_true",
#         help="Enable verbose output",
#     )
    
#     parser.add_argument(
#         "-r", "--roll-no",
#         type=str,
#         default=None,
#         metavar="ROLL",
#         help="Roll number to include in the document header",
#     )

#     parser.add_argument(
#         "--save-roll",
#         action="store_true",
#         help="Save the current --roll-no as the default for future runs",
#     )
    
#     parser.add_argument(
#         "--version",
#         action="version",
#         version="%(prog)s 1.0.0",
#     )
    
#     return parser


# def main(args: list[str] | None = None) -> int:
#     """
#     Main entry point for the auto-docx tool.
    
#     Args:
#         args: Command-line arguments (uses sys.argv if None)
        
#     Returns:
#         Exit code (0 for success, 1 for errors)
#     """
#     parser = create_parser()
#     parsed_args = parser.parse_args(args)
    
#     try:
#         # List environments if requested
#         if parsed_args.list_envs:
#             envs = ScriptExecutor.discover_envs()
#             if not envs:
#                 print("No Python environments found.")
#             else:
#                 print("Available Python environments:")
#                 for idx, env in enumerate(envs):
#                     print(f"  [{idx}] {env['name']} ({env['source']}): {env['python']}")
#             return 0

#         # Validate input file
#         if not parsed_args.script_path:
#             parser.print_usage()
#             print("Error: script_path is required unless using --list-envs", file=sys.stderr)
#             return 1

#         script_path = Path(parsed_args.script_path)
        
#         if not script_path.exists():
#             print(f"Error: File not found: {script_path}", file=sys.stderr)
#             return 1
        
#         if script_path.suffix.lower() not in {".py", ".ipynb"}:
#             print("Error: File must be a Python script (.py) or notebook (.ipynb)", file=sys.stderr)
#             return 1
        
#         if parsed_args.verbose:
#             print("=" * 60)
#             print("AUTO-DOCX: Script Execution Logger")
#             print("=" * 60)

#         # Resolve python executable
#         python_executable = parsed_args.python
#         env_to_use = parsed_args.env
        
#         # Load saved env if none specified
#         if not env_to_use and not python_executable:
#             config = load_config()
#             env_to_use = config.get("env")
        
#         # Save env if requested
#         if parsed_args.save_env and parsed_args.env:
#             config = load_config()
#             config["env"] = parsed_args.env
#             save_config(config)
#             print(f"Saved default environment: {parsed_args.env}")
        
#         if env_to_use and not python_executable:
#             envs = ScriptExecutor.discover_envs()
#             python_executable = ScriptExecutor.select_python(env_to_use, envs)
#             if not python_executable:
#                 print(f"Error: Environment not found: {env_to_use}", file=sys.stderr)
#                 return 1
        
#         # Resolve roll number
#         roll_no = parsed_args.roll_no
#         if not roll_no:
#             config = load_config()
#             roll_no = config.get("roll_no")
        
#         # Save roll number if requested
#         if parsed_args.save_roll and parsed_args.roll_no:
#             config = load_config()
#             config["roll_no"] = parsed_args.roll_no
#             save_config(config)
#             print(f"Saved default roll number: {parsed_args.roll_no}")
        
#         # Execute the script
#         executor = ScriptExecutor(
#             timeout=parsed_args.timeout,
#             verbose=parsed_args.verbose,
#             python_executable=python_executable,
#         )
#         result = executor.execute(script_path)
        
#         # Generate the document (source included by default, unless --no-source is used)
#         generator = DocumentGenerator(
#             include_source=not parsed_args.no_source,
#             verbose=parsed_args.verbose,
#             roll_no=roll_no,
#         )
#         output_path = generator.generate(result, parsed_args.output)
        
#         # Print results
#         print(f"\n{'=' * 60}")
#         print(f"Document generated: {output_path}")
#         print(f"Script return code: {result.return_code}")
#         print(f"Status: {'SUCCESS' if result.return_code == 0 else 'COMPLETED WITH ERRORS'}")
#         print(f"{'=' * 60}")
        
#         # Return the script's exit code
#         return 0 if result.return_code == 0 else result.return_code
        
#     except FileNotFoundError as e:
#         print(f"Error: {e}", file=sys.stderr)
#         return 1
#     except ValueError as e:
#         print(f"Error: {e}", file=sys.stderr)
#         return 1
#     except Exception as e:
#         import traceback
#         print(f"Error: {e}", file=sys.stderr)
#         traceback.print_exc()
#         return 1


# if __name__ == "__main__":
#     sys.exit(main())


#!/usr/bin/env python3
"""
Auto-DOCX: Main entry point and command-line interface.

Execute Python scripts and generate Word documents with their output.
"""

import argparse
import json
import shutil
import sys
from pathlib import Path

from .executor import ScriptExecutor
from .document import DocumentGenerator

CONFIG_FILE = Path.home() / ".auto_docx_config.json"


def load_config() -> dict:
    """Load saved configuration."""
    if CONFIG_FILE.exists():
        try:
            return json.loads(CONFIG_FILE.read_text())
        except Exception:
            return {}
    return {}


def save_config(config: dict) -> None:
    """Save configuration to file."""
    CONFIG_FILE.write_text(json.dumps(config, indent=2))


def create_parser() -> argparse.ArgumentParser:
    """Create and configure the argument parser."""
    parser = argparse.ArgumentParser(
        prog="auto-docx",
        description="Execute a Python script and log its output to a Word document.",
        epilog="Example: auto-docx my_script.py -o output.docx --include-source",
        formatter_class=argparse.RawDescriptionHelpFormatter,
    )
    
    parser.add_argument(
        "script_path",
        type=str,
        nargs="?",
        help="Path to the Python file (.py) or notebook (.ipynb) to execute",
    )
    
    parser.add_argument(
        "-o", "--output",
        type=str,
        default=None,
        metavar="PATH",
        help="Custom output path for the Word document (default: <script>_output.docx)",
    )
    
    parser.add_argument(
        "--no-source",
        action="store_true",
        help="Exclude the script's source code from the document",
    )

    parser.add_argument(
        "--list-envs",
        action="store_true",
        help="List available Python environments and exit",
    )

    parser.add_argument(
        "--env",
        type=str,
        default=None,
        metavar="ENV",
        help="Select an environment by name or index from --list-envs",
    )

    parser.add_argument(
        "--save-env",
        action="store_true",
        help="Save the current --env as the default for future runs",
    )

    parser.add_argument(
        "--python",
        type=str,
        default=None,
        metavar="PATH",
        help="Path to the Python executable to use for execution",
    )
    
    parser.add_argument(
        "--timeout",
        type=int,
        default=300,
        metavar="SECONDS",
        help="Timeout for script execution in seconds (default: 300)",
    )
    
    parser.add_argument(
        "-v", "--verbose",
        action="store_true",
        help="Enable verbose output",
    )
    
    parser.add_argument(
        "-r", "--roll-no",
        type=str,
        default=None,
        metavar="ROLL",
        help="Roll number to include in the document header",
    )

    parser.add_argument(
        "--save-roll",
        action="store_true",
        help="Save the current --roll-no as the default for future runs",
    )
    
    parser.add_argument(
        "--version",
        action="version",
        version="%(prog)s 1.0.1",
    )
    
    return parser


def main(args: list[str] | None = None) -> int:
    """Main entry point for the auto-docx tool."""
    parser = create_parser()
    parsed_args = parser.parse_args(args)
    
    try:
        # List environments if requested
        if parsed_args.list_envs:
            envs = ScriptExecutor.discover_envs()
            if not envs:
                print("No Python environments found.")
            else:
                print("Available Python environments:")
                for idx, env in enumerate(envs):
                    print(f"  [{idx}] {env['name']} ({env['source']}): {env['python']}")
            return 0

        # Validate input file
        if not parsed_args.script_path:
            parser.print_usage()
            print("Error: script_path is required unless using --list-envs", file=sys.stderr)
            return 1

        script_path = Path(parsed_args.script_path)
        
        if not script_path.exists():
            print(f"Error: File not found: {script_path}", file=sys.stderr)
            return 1
        
        if script_path.suffix.lower() not in {".py", ".ipynb"}:
            print("Error: File must be a Python script (.py) or notebook (.ipynb)", file=sys.stderr)
            return 1
        
        if parsed_args.verbose:
            print("=" * 60)
            print("AUTO-DOCX: Script Execution Logger")
            print("=" * 60)

        # Resolve python executable
        python_executable = parsed_args.python
        env_to_use = parsed_args.env
        
        if not env_to_use and not python_executable:
            config = load_config()
            env_to_use = config.get("env")
        
        if parsed_args.save_env and parsed_args.env:
            config = load_config()
            config["env"] = parsed_args.env
            save_config(config)
            print(f"Saved default environment: {parsed_args.env}")
        
        if env_to_use and not python_executable:
            envs = ScriptExecutor.discover_envs()
            python_executable = ScriptExecutor.select_python(env_to_use, envs)
            if not python_executable:
                print(f"Error: Environment not found: {env_to_use}", file=sys.stderr)
                return 1
        
        # Resolve roll number
        roll_no = parsed_args.roll_no
        if not roll_no:
            config = load_config()
            roll_no = config.get("roll_no")
        
        if parsed_args.save_roll and parsed_args.roll_no:
            config = load_config()
            config["roll_no"] = parsed_args.roll_no
            save_config(config)
            print(f"Saved default roll number: {parsed_args.roll_no}")
        
        # Execute the script
        executor = ScriptExecutor(
            timeout=parsed_args.timeout,
            verbose=parsed_args.verbose,
            python_executable=python_executable,
        )
        result = executor.execute(script_path)
        
        # Generate the document
        generator = DocumentGenerator(
            include_source=not parsed_args.no_source,
            verbose=parsed_args.verbose,
            roll_no=roll_no,
        )
        output_path = generator.generate(result, parsed_args.output)
        
        # Clean up images folder
        images_dir = script_path.parent / f".{script_path.stem}_images"
        if images_dir.exists():
            shutil.rmtree(images_dir)
        
        # Print results
        print(f"\n{'=' * 60}")
        print(f"Document generated: {output_path}")
        print(f"Script return code: {result.return_code}")
        print(f"Status: {'SUCCESS' if result.return_code == 0 else 'COMPLETED WITH ERRORS'}")
        print(f"{'=' * 60}")
        
        return 0 if result.return_code == 0 else result.return_code
        
    except FileNotFoundError as e:
        print(f"Error: {e}", file=sys.stderr)
        return 1
    except ValueError as e:
        print(f"Error: {e}", file=sys.stderr)
        return 1
    except Exception as e:
        import traceback
        print(f"Error: {e}", file=sys.stderr)
        traceback.print_exc()
        return 1


if __name__ == "__main__":
    sys.exit(main())