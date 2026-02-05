"""
Script execution module.

Handles the execution of Python scripts and notebooks using subprocess and captures their output.
"""

import json
import os
import re
import shutil
import subprocess
import sys
import tempfile
from dataclasses import dataclass
from pathlib import Path
from typing import Iterable, Optional

IMAGE_MARKER_PREFIX = "__AUTO_DOCX_IMAGE__"


@dataclass
class OutputItem:
    """Represents an output item in order (text or image)."""

    kind: str  # "text" or "image"
    content: str


@dataclass
class ExecutionResult:
    """Container for script execution results."""
    
    stdout: str
    stderr: str
    return_code: int
    script_path: Path
    source_code: str
    success: bool
    output_items: Optional[list[OutputItem]] = None
    error_message: Optional[str] = None
    
    @property
    def has_errors(self) -> bool:
        """Check if execution produced any errors."""
        return bool(self.stderr) or self.return_code != 0


class ScriptExecutor:
    """Executes Python scripts and notebooks and captures their output."""
    
    def __init__(self, timeout: int = 300, verbose: bool = False, python_executable: Optional[str] = None, inputs: Optional[list[str]] = None):
        """
        Initialize the executor.
        
        Args:
            timeout: Maximum execution time in seconds (default: 300)
            verbose: Enable verbose output to console
            python_executable: Path to Python interpreter to use
            inputs: List of input values for input() prompts in notebooks
        """
        self.timeout = timeout
        self.verbose = verbose
        self.python_executable = python_executable
        self.inputs = inputs or []

    @staticmethod
    def discover_envs() -> list[dict[str, str]]:
        """Discover available Python environments on the system."""
        envs: list[dict[str, str]] = []
        seen: set[str] = set()

        def add_env(name: str, python_path: str, source: str) -> None:
            if not python_path:
                return
            python_path = os.path.abspath(python_path)
            key = python_path.lower()
            if key in seen:
                return
            seen.add(key)
            envs.append({"name": name, "python": python_path, "source": source})

        # Current Python
        add_env("current", sys.executable, "system")

        # Conda environments
        conda = shutil.which("conda")
        if conda:
            try:
                result = subprocess.run(
                    [conda, "env", "list", "--json"],
                    capture_output=True,
                    text=True,
                    timeout=5,
                )
                if result.returncode == 0:
                    data = json.loads(result.stdout or "{}")
                    for env_path in data.get("envs", []):
                        env_name = Path(env_path).name
                        python_path = Path(env_path) / "bin" / "python"
                        if python_path.exists():
                            add_env(env_name, str(python_path), "conda")
            except Exception:
                pass

        # Common virtualenv locations
        for base in [Path.home() / ".virtualenvs", Path.home() / "venvs"]:
            if base.exists() and base.is_dir():
                for env_dir in base.iterdir():
                    python_path = env_dir / "bin" / "python"
                    if python_path.exists():
                        add_env(env_dir.name, str(python_path), "venv")

        return envs

    @staticmethod
    def select_python(env_identifier: str, envs: Iterable[dict[str, str]]) -> Optional[str]:
        """Select a Python path from discovered environments by name or index."""
        envs_list = list(envs)
        if env_identifier.isdigit():
            idx = int(env_identifier)
            if 0 <= idx < len(envs_list):
                return envs_list[idx]["python"]
        for env in envs_list:
            if env_identifier == env.get("name"):
                return env.get("python")
        return None
    
    def execute(self, script_path: str | Path) -> ExecutionResult:
        """
        Execute a Python script or notebook and capture its output.
        
        Args:
            script_path: Path to the Python script or notebook to execute
            
        Returns:
            ExecutionResult containing stdout, stderr, return code, and metadata
            
        Raises:
            FileNotFoundError: If the script file doesn't exist
            ValueError: If the file is not a Python script or notebook
        """
        script_path = Path(script_path).resolve()
        
        # Validate the script path
        if not script_path.exists():
            raise FileNotFoundError(f"Script not found: {script_path}")
        
        if not script_path.is_file():
            raise ValueError(f"Path is not a file: {script_path}")
        
        if script_path.suffix.lower() not in {".py", ".ipynb"}:
            raise ValueError(f"File is not a Python script or notebook: {script_path}")
        
        # Read the source code
        try:
            source_code = script_path.read_text(encoding="utf-8")
        except UnicodeDecodeError:
            source_code = script_path.read_text(encoding="latin-1")
        
        if self.verbose:
            print(f"[INFO] Executing script: {script_path}")
            print(f"[INFO] Timeout: {self.timeout} seconds")
        
        python_exec = self.python_executable or sys.executable

        # Execute the script or notebook
        try:
            if script_path.suffix.lower() == ".ipynb":
                return self._execute_notebook(script_path, source_code, python_exec, self.inputs)
            return self._execute_script(script_path, source_code, python_exec)

        except subprocess.TimeoutExpired as e:
            error_msg = f"Execution timed out after {self.timeout} seconds"
            if self.verbose:
                print(f"[ERROR] {error_msg}")
            
            return ExecutionResult(
                stdout=e.stdout or "" if hasattr(e, "stdout") else "",
                stderr=e.stderr or "" if hasattr(e, "stderr") else "",
                return_code=-1,
                script_path=script_path,
                source_code=source_code,
                success=False,
                error_message=error_msg,
            )
            
        except Exception as e:
            error_msg = f"Failed to execute: {str(e)}"
            if self.verbose:
                print(f"[ERROR] {error_msg}")
            
            return ExecutionResult(
                stdout="",
                stderr=str(e),
                return_code=-1,
                script_path=script_path,
                source_code=source_code,
                success=False,
                error_message=error_msg,
            )

    def _execute_script(self, script_path: Path, source_code: str, python_exec: str) -> ExecutionResult:
        """Execute a .py script using the image-aware runner."""
        # Create persistent images directory next to the script
        persistent_images_dir = script_path.parent / f".{script_path.stem}_images"
        persistent_images_dir.mkdir(parents=True, exist_ok=True)

        with tempfile.TemporaryDirectory() as tmpdir:
            images_dir = Path(tmpdir) / "images"
            images_dir.mkdir(parents=True, exist_ok=True)

            env = os.environ.copy()
            project_src = Path(__file__).resolve().parents[1]
            existing = env.get("PYTHONPATH", "")
            env["PYTHONPATH"] = f"{project_src}{os.pathsep}{existing}" if existing else str(project_src)

            result = subprocess.run(
                [python_exec, "-m", "auto_docx.runner", str(script_path), str(images_dir)],
                capture_output=True,
                text=True,
                timeout=self.timeout,
                cwd=script_path.parent,
                env=env,
            )

            if self.verbose:
                print(f"[INFO] Execution completed with return code: {result.returncode}")

            # Copy images to persistent location
            output_items = self._parse_output_with_images(result.stdout, images_dir)
            output_items = self._copy_images_to_persistent(output_items, persistent_images_dir)

            return ExecutionResult(
                stdout=result.stdout,
                stderr=result.stderr,
                return_code=result.returncode,
                script_path=script_path,
                source_code=source_code,
                success=result.returncode == 0,
                output_items=output_items,
            )

    def _execute_notebook(self, notebook_path: Path, source_code: str, python_exec: str, inputs: list[str]) -> ExecutionResult:
        """Execute a .ipynb notebook using the notebook runner."""
        # Create persistent images directory next to the notebook
        persistent_images_dir = notebook_path.parent / f".{notebook_path.stem}_images"
        persistent_images_dir.mkdir(parents=True, exist_ok=True)

        with tempfile.TemporaryDirectory() as tmpdir:
            images_dir = Path(tmpdir) / "images"
            images_dir.mkdir(parents=True, exist_ok=True)
            output_json = Path(tmpdir) / "output.json"

            env = os.environ.copy()
            project_src = Path(__file__).resolve().parents[1]
            existing = env.get("PYTHONPATH", "")
            env["PYTHONPATH"] = f"{project_src}{os.pathsep}{existing}" if existing else str(project_src)

            # Pass inputs as additional arguments
            cmd = [
                python_exec,
                "-m",
                "auto_docx.notebook_runner",
                str(notebook_path),
                str(output_json),
                str(images_dir),
                str(self.timeout),
            ] + inputs

            result = subprocess.run(
                cmd,
                capture_output=True,
                text=True,
                timeout=self.timeout,
                cwd=notebook_path.parent,
                env=env,
            )

            if self.verbose:
                print(f"[INFO] Notebook execution return code: {result.returncode}")

            if output_json.exists():
                data = json.loads(output_json.read_text(encoding="utf-8"))
                output_items = [OutputItem(**item) for item in data.get("output_items", [])]
                # Copy images to persistent location
                output_items = self._copy_images_to_persistent(output_items, persistent_images_dir)
                notebook_source = data.get("source_code", source_code)
                return ExecutionResult(
                    stdout=data.get("stdout", ""),
                    stderr=data.get("stderr", ""),
                    return_code=data.get("return_code", result.returncode),
                    script_path=notebook_path,
                    source_code=notebook_source,
                    success=result.returncode == 0,
                    output_items=output_items,
                    error_message=data.get("error_message"),
                )

            return ExecutionResult(
                stdout=result.stdout,
                stderr=result.stderr,
                return_code=result.returncode,
                script_path=notebook_path,
                source_code=source_code,
                success=result.returncode == 0,
                error_message="Notebook runner did not produce output.",
            )

    def _parse_output_with_images(self, stdout: str, images_dir: Path) -> list[OutputItem]:
        """Parse stdout and split into text and image items based on markers."""
        items: list[OutputItem] = []
        buffer: list[str] = []
        pattern = re.compile(rf"^{re.escape(IMAGE_MARKER_PREFIX)}:(.+)$")

        for line in stdout.splitlines():
            match = pattern.match(line.strip())
            if match:
                if buffer:
                    items.append(OutputItem(kind="text", content="\n".join(buffer)))
                    buffer = []
                image_path = match.group(1).strip()
                if not os.path.isabs(image_path):
                    image_path = str(images_dir / image_path)
                items.append(OutputItem(kind="image", content=image_path))
            else:
                buffer.append(line)

        if buffer:
            items.append(OutputItem(kind="text", content="\n".join(buffer)))

        return items

    def _copy_images_to_persistent(self, items: list[OutputItem], persistent_dir: Path) -> list[OutputItem]:
        """Copy images from temp directory to persistent location and update paths."""
        new_items: list[OutputItem] = []
        img_counter = 0

        for item in items:
            if item.kind == "image":
                src_path = Path(item.content)
                if src_path.exists():
                    img_counter += 1
                    dst_path = persistent_dir / f"img_{img_counter:04d}{src_path.suffix}"
                    shutil.copy2(src_path, dst_path)
                    new_items.append(OutputItem(kind="image", content=str(dst_path)))
                else:
                    # Image doesn't exist, skip or keep original
                    new_items.append(item)
            else:
                new_items.append(item)

        return new_items
