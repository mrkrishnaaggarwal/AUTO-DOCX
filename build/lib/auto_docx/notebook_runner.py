# """
# Notebook runner for executing .ipynb files and extracting output items.
# """

# from __future__ import annotations

# import base64
# import json
# import sys
# from pathlib import Path

# import nbformat
# from nbclient import NotebookClient


# def _save_image(data_b64: str, images_dir: Path, index: int) -> str:
#     images_dir.mkdir(parents=True, exist_ok=True)
#     image_path = images_dir / f"nb_{index:04d}.png"
#     image_path.write_bytes(base64.b64decode(data_b64))
#     return str(image_path)


# def main() -> int:
#     if len(sys.argv) < 5:
#         print("Usage: python -m auto_docx.notebook_runner <notebook> <output_json> <images_dir> <timeout>", file=sys.stderr)
#         return 1

#     notebook_path = Path(sys.argv[1]).resolve()
#     output_json = Path(sys.argv[2]).resolve()
#     images_dir = Path(sys.argv[3]).resolve()
#     timeout = int(sys.argv[4])

#     output_items = []
#     stderr = ""
#     stdout = ""
#     return_code = 0
#     error_message = None

#     try:
#         nb = nbformat.read(str(notebook_path), as_version=4)
#         client = NotebookClient(nb, timeout=timeout, allow_errors=True)
#         client.execute()

#         # Source code (code cells only)
#         source_cells = []
#         for idx, cell in enumerate(nb.cells):
#             if cell.cell_type == "code":
#                 source_cells.append(f"# Cell {idx + 1}\n{cell.source}")
#         source_code = "\n\n".join(source_cells)

#         img_index = 0
#         for cell in nb.cells:
#             if cell.cell_type == "markdown":
#                 # Include markdown cells as markdown output items
#                 md_content = cell.source.strip()
#                 if md_content:
#                     output_items.append({"kind": "markdown", "content": md_content})
#                 continue
#             if cell.cell_type != "code":
#                 continue
#             for output in cell.get("outputs", []):
#                 output_type = output.get("output_type")
#                 if output_type == "stream":
#                     text = output.get("text", "")
#                     if text:
#                         output_items.append({"kind": "text", "content": text.rstrip("\n")})
#                         stdout += text
#                 elif output_type in {"execute_result", "display_data"}:
#                     data = output.get("data", {})
#                     if "image/png" in data:
#                         img_index += 1
#                         path = _save_image(data["image/png"], images_dir, img_index)
#                         output_items.append({"kind": "image", "content": path})
#                     elif "text/plain" in data:
#                         text = data["text/plain"]
#                         output_items.append({"kind": "text", "content": str(text).rstrip("\n")})
#                 elif output_type == "error":
#                     traceback_lines = output.get("traceback", [])
#                     err_text = "\n".join(traceback_lines)
#                     if err_text:
#                         stderr += err_text + "\n"
#                         output_items.append({"kind": "text", "content": err_text})
#                         return_code = 1
#                         error_message = "Notebook execution error"

#         output_json.write_text(
#             json.dumps(
#                 {
#                     "stdout": stdout,
#                     "stderr": stderr,
#                     "return_code": return_code,
#                     "error_message": error_message,
#                     "source_code": source_code,
#                     "output_items": output_items,
#                 },
#                 ensure_ascii=False,
#             ),
#             encoding="utf-8",
#         )
#         return return_code

#     except Exception as exc:
#         output_json.write_text(
#             json.dumps(
#                 {
#                     "stdout": "",
#                     "stderr": str(exc),
#                     "return_code": 1,
#                     "error_message": "Notebook runner failed",
#                     "source_code": "",
#                     "output_items": [],
#                 },
#                 ensure_ascii=False,
#             ),
#             encoding="utf-8",
#         )
#         return 1


# if __name__ == "__main__":
#     sys.exit(main())


"""
Notebook conversion utility.
Maintained for backward compatibility or direct usage.
"""

from __future__ import annotations
import sys
from pathlib import Path
import nbformat

def main() -> int:
    """
    Simulates execution by converting to script and printing source.
    Note: Real execution is now handled by executor.py's interactive runner.
    """
    if len(sys.argv) < 2:
        print("Usage: python -m auto_docx.notebook_runner <notebook>", file=sys.stderr)
        return 1

    notebook_path = Path(sys.argv[1]).resolve()
    try:
        nb = nbformat.read(str(notebook_path), as_version=4)
        code_cells = [c.source for c in nb.cells if c.cell_type == 'code']
        print("\n".join(code_cells))
        return 0
    except Exception as e:
        print(f"Error reading notebook: {e}", file=sys.stderr)
        return 1

if __name__ == "__main__":
    sys.exit(main())