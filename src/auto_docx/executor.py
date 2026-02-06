"""
Script execution module.

Handles the execution of Python scripts and notebooks using subprocess
with interactive support (stdin) and race-condition-free output capturing.
"""

import base64
import json
import os
import re
import shutil
import subprocess
import sys
import tempfile
import threading
import queue
from dataclasses import dataclass
from pathlib import Path
from typing import Iterable, Optional, List, Dict, Union

# Protocol Markers
IMAGE_MARKER_PREFIX = "__AUTO_DOCX_IMAGE__"
MD_START_MARKER = "__AUTO_DOCX_MD_START__"
MD_END_MARKER = "__AUTO_DOCX_MD_END__"

@dataclass
class OutputItem:
    """Represents an output item in order (text, image, or markdown)."""
    kind: str  # "text", "image", "markdown"
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
    output_items: Optional[List[OutputItem]] = None
    error_message: Optional[str] = None
    
    @property
    def has_errors(self) -> bool:
        return bool(self.stderr) or self.return_code != 0

class ScriptExecutor:
    """Executes Python scripts and notebooks interactively."""
    
    def __init__(self, timeout: int = 300, verbose: bool = False, python_executable: Optional[str] = None):
        self.timeout = timeout
        self.verbose = verbose
        self.python_executable = python_executable

    @staticmethod
    def discover_envs() -> List[Dict[str, str]]:
        """Discover available Python environments on the system."""
        envs: List[Dict[str, str]] = []
        seen: set[str] = set()

        def add_env(name: str, python_path: str, source: str) -> None:
            if not python_path: return
            python_path = os.path.abspath(python_path)
            key = python_path.lower()
            if key in seen: return
            seen.add(key)
            envs.append({"name": name, "python": python_path, "source": source})

        add_env("current", sys.executable, "system")

        conda = shutil.which("conda")
        if conda:
            try:
                result = subprocess.run([conda, "env", "list", "--json"], capture_output=True, text=True, timeout=5)
                if result.returncode == 0:
                    data = json.loads(result.stdout or "{}")
                    for env_path in data.get("envs", []):
                        env_name = Path(env_path).name
                        python_path = Path(env_path) / "bin" / "python"
                        if python_path.exists():
                            add_env(env_name, str(python_path), "conda")
            except Exception: pass

        for base in [Path.home() / ".virtualenvs", Path.home() / "venvs"]:
            if base.exists() and base.is_dir():
                for env_dir in base.iterdir():
                    python_path = env_dir / "bin" / "python"
                    if python_path.exists():
                        add_env(env_dir.name, str(python_path), "venv")
        return envs

    @staticmethod
    def select_python(env_identifier: str, envs: Iterable[Dict[str, str]]) -> Optional[str]:
        envs_list = list(envs)
        if env_identifier.isdigit():
            idx = int(env_identifier)
            if 0 <= idx < len(envs_list):
                return envs_list[idx]["python"]
        for env in envs_list:
            if env_identifier == env.get("name"):
                return env.get("python")
        return None
    
    def execute(self, script_path: Union[str, Path]) -> ExecutionResult:
        script_path = Path(script_path).resolve()
        
        if not script_path.exists():
            raise FileNotFoundError(f"Script not found: {script_path}")
        
        # Read the source code
        try:
            source_code = script_path.read_text(encoding="utf-8")
        except UnicodeDecodeError:
            source_code = script_path.read_text(encoding="latin-1")
        
        if self.verbose:
            print(f"[INFO] Executing script: {script_path}")
        
        python_exec = self.python_executable or sys.executable

        try:
            # Convert Notebooks to Scripts, then run interactively
            if script_path.suffix.lower() == ".ipynb":
                return self._execute_notebook_as_script(script_path, python_exec)
            else:
                return self._execute_script_interactive(script_path, source_code, python_exec)

        except Exception as e:
            import traceback
            traceback.print_exc()
            return ExecutionResult(
                stdout="", stderr=str(e), return_code=-1,
                script_path=script_path, source_code=source_code,
                success=False, error_message=f"Execution Failed: {str(e)}"
            )

    def _execute_notebook_as_script(self, notebook_path: Path, python_exec: str) -> ExecutionResult:
        """Convert notebook to script (including Markdown cells) and execute interactively."""
        persistent_images_dir = notebook_path.parent / f".{notebook_path.stem}_images"
        persistent_images_dir.mkdir(parents=True, exist_ok=True)

        with tempfile.TemporaryDirectory() as tmpdir:
            temp_script_path = Path(tmpdir) / "temp_nb_exec.py"
            
            # Convert .ipynb to .py
            try:
                import nbformat
                nb = nbformat.read(notebook_path, as_version=4)
                
                converted_lines = ["import base64", ""]
                source_code_display = [] 

                for cell in nb.cells:
                    if cell.cell_type == 'code':
                        converted_lines.append(f"# %% [Code]\n{cell.source}\n")
                        source_code_display.append(cell.source)
                    elif cell.cell_type == 'markdown':
                        md_content = cell.source
                        b64_content = base64.b64encode(md_content.encode('utf-8')).decode('utf-8')
                        
                        injector_code = (
                            f"print('{MD_START_MARKER}')\n"
                            f"print(base64.b64decode('{b64_content}').decode('utf-8'))\n"
                            f"print('{MD_END_MARKER}')\n"
                        )
                        converted_lines.append(f"# %% [Markdown]\n{injector_code}")

                converted_source = "\n".join(converted_lines)
                temp_script_path.write_text(converted_source, encoding='utf-8')
                
                clean_source_code = "\n\n".join(source_code_display)
                
            except Exception as e:
                return ExecutionResult(
                    stdout="", stderr=str(e), return_code=1,
                    script_path=notebook_path, source_code="", success=False,
                    error_message=f"Failed to parse notebook: {e}"
                )

            result = self._execute_script_interactive(temp_script_path, converted_source, python_exec)

            result.script_path = notebook_path
            result.source_code = clean_source_code
            
            final_items = []
            if result.output_items:
                for item in result.output_items:
                    if item.kind == 'image':
                        src = Path(item.content)
                        if src.exists():
                            dst = persistent_images_dir / src.name
                            shutil.copy2(src, dst)
                            final_items.append(OutputItem('image', str(dst)))
                    else:
                        final_items.append(item)
            
            result.output_items = final_items
            return result

    def _execute_script_interactive(self, script_path: Path, source_code: str, python_exec: str) -> ExecutionResult:
        """
        Execute a .py script using Popen + Threading for interactivity.
        """
        temp_dir_obj = tempfile.TemporaryDirectory()
        temp_dir = Path(temp_dir_obj.name)
        images_dir = temp_dir / "images"
        images_dir.mkdir(parents=True, exist_ok=True)

        env = os.environ.copy()
        project_src = Path(__file__).resolve().parents[1]
        env["PYTHONPATH"] = f"{project_src}{os.pathsep}{env.get('PYTHONPATH', '')}"

        cmd = [python_exec, "-u", "-m", "auto_docx.runner", str(script_path), str(images_dir)]

        process = subprocess.Popen(
            cmd,
            stdin=sys.stdin,
            stdout=subprocess.PIPE,
            stderr=subprocess.STDOUT, 
            text=True,
            cwd=script_path.parent,
            env=env,
            bufsize=0
        )

        out_queue = queue.Queue()

        def stream_reader(pipe):
            """
            Reads output char-by-char.
            Filters protocol markers (Markdown/Images) from the Terminal output,
            but captures everything for the Document generation.
            """
            buffer = ""
            hiding_markdown = False
            markers = [MD_START_MARKER, MD_END_MARKER, IMAGE_MARKER_PREFIX]

            while True:
                char = pipe.read(1)
                if not char:
                    break
                
                # 1. Capture for Document (Always)
                out_queue.put(char)
                
                # 2. Terminal Output Logic (Filtering)
                buffer += char
                
                # Handling inside Markdown Block
                if hiding_markdown:
                    if MD_END_MARKER in buffer:
                        hiding_markdown = False
                        buffer = "" # Consumed marker
                    elif char == '\n':
                        buffer = "" # Clear line buffer
                    continue

                # Detect Start of Markdown
                if MD_START_MARKER in buffer:
                    hiding_markdown = True
                    buffer = "" # Consumed marker
                    continue

                # Detect Image Marker (Single line suppress)
                if IMAGE_MARKER_PREFIX in buffer:
                    if char == '\n':
                        buffer = "" # Consumed line
                    continue
                
                # Intelligent Flushing
                # If buffer starts with part of a marker (e.g. "_", "__", "__A"), hold it.
                # Otherwise, print it.
                is_partial_match = False
                for m in markers:
                    if m.startswith(buffer):
                        is_partial_match = True
                        break
                
                if is_partial_match:
                    continue # Hold
                else:
                    # Safe to print
                    sys.stdout.write(buffer)
                    sys.stdout.flush()
                    buffer = ""

        t_out = threading.Thread(target=stream_reader, args=(process.stdout,))
        t_out.start()

        try:
            process.wait(timeout=self.timeout)
        except subprocess.TimeoutExpired:
            process.kill()
            if self.verbose:
                print(f"\n[ERROR] Timeout after {self.timeout}s")
        
        t_out.join()

        full_output = ""
        while not out_queue.empty():
            full_output += out_queue.get()

        output_items = self._parse_output_stream(full_output, images_dir)

        persistent_images_dir = script_path.parent / f".{script_path.stem}_images"
        persistent_images_dir.mkdir(parents=True, exist_ok=True)
        
        final_items = self._copy_images_to_persistent(output_items, persistent_images_dir)
        
        temp_dir_obj.cleanup()

        return ExecutionResult(
            stdout=full_output, 
            stderr="",          
            return_code=process.returncode,
            script_path=script_path,
            source_code=source_code,
            success=process.returncode == 0,
            output_items=final_items
        )

    def _parse_output_stream(self, stdout: str, images_dir: Path) -> List[OutputItem]:
        """Parses interleaved text, images, and markdown markers."""
        items: List[OutputItem] = []
        buffer: List[str] = []
        
        lines = stdout.splitlines()
        state = "NORMAL" 
        
        for line in lines:
            if MD_START_MARKER in line:
                if buffer:
                    items.append(OutputItem(kind="text", content="\n".join(buffer)))
                    buffer = []
                state = "IN_MARKDOWN"
                continue
            
            if MD_END_MARKER in line:
                if buffer:
                    items.append(OutputItem(kind="markdown", content="\n".join(buffer)))
                    buffer = []
                state = "NORMAL"
                continue
            
            if state == "NORMAL" and IMAGE_MARKER_PREFIX in line:
                parts = line.split(f"{IMAGE_MARKER_PREFIX}:")
                if parts[0].strip():
                    buffer.append(parts[0])
                
                if buffer:
                    items.append(OutputItem(kind="text", content="\n".join(buffer)))
                    buffer = []
                
                image_path = parts[1].strip()
                if not os.path.isabs(image_path):
                    image_path = str(images_dir / image_path)
                items.append(OutputItem(kind="image", content=image_path))
                continue
            
            buffer.append(line)

        if buffer:
            kind = "markdown" if state == "IN_MARKDOWN" else "text"
            items.append(OutputItem(kind=kind, content="\n".join(buffer)))

        return items

    def _copy_images_to_persistent(self, items: List[OutputItem], persistent_dir: Path) -> List[OutputItem]:
        new_items: List[OutputItem] = []
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
                    new_items.append(item)
            else:
                new_items.append(item)
        return new_items