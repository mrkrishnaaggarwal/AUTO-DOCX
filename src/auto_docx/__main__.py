"""
Allow running the package as a module: python -m auto_docx
"""

from .main import main
import sys

if __name__ == "__main__":
    sys.exit(main())
