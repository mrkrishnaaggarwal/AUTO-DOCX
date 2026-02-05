"""
Auto-DOCX: Execute Python scripts and log output to Word documents.

This package provides functionality to:
- Execute Python scripts in a subprocess
- Capture stdout and stderr
- Generate formatted Word documents with the results
"""

__version__ = "1.0.0"
__author__ = "Developer"

from .executor import ScriptExecutor, ExecutionResult
from .document import DocumentGenerator

__all__ = [
    "ScriptExecutor",
    "ExecutionResult", 
    "DocumentGenerator",
    "__version__",
]
