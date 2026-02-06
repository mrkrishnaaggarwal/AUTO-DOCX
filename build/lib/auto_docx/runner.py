"""
Image-aware runner for executing Python scripts.

This module patches matplotlib and cv2 to capture images in order and emit markers.
"""

from __future__ import annotations

import os
import runpy
import sys
import traceback
from pathlib import Path

IMAGE_MARKER_PREFIX = "__AUTO_DOCX_IMAGE__"


def _print_image_marker(image_path: Path) -> None:
    print(f"{IMAGE_MARKER_PREFIX}:{image_path}", flush=True)


def _patch_matplotlib(images_dir: Path, counter: list[int]) -> None:
    try:
        import matplotlib

        matplotlib.use("Agg")
        import matplotlib.pyplot as plt  # noqa: WPS433

        original_show = plt.show

        def save_figures():
            for fig_num in plt.get_fignums():
                fig = plt.figure(fig_num)
                counter[0] += 1
                image_path = images_dir / f"mpl_{counter[0]:04d}.png"
                fig.savefig(image_path, bbox_inches="tight")
                _print_image_marker(image_path)

        def patched_show(*args, **kwargs):  # noqa: ANN001
            save_figures()
            return original_show(*args, **kwargs)

        plt.show = patched_show

        def save_remaining():
            save_figures()

        # Save any figures created without show()
        import atexit  # noqa: WPS433

        atexit.register(save_remaining)
    except Exception:
        # Matplotlib not available or patching failed; ignore
        return


def _patch_cv2(images_dir: Path, counter: list[int]) -> None:
    try:
        import cv2  # noqa: WPS433

        def patched_imshow(winname, mat):  # noqa: ANN001
            counter[0] += 1
            image_path = images_dir / f"cv2_{counter[0]:04d}.png"
            try:
                cv2.imwrite(str(image_path), mat)
                _print_image_marker(image_path)
            except Exception:
                pass

        def patched_wait_key(delay=0):  # noqa: ANN001
            return 0

        def patched_destroy_all():
            return None

        cv2.imshow = patched_imshow
        cv2.waitKey = patched_wait_key
        cv2.destroyAllWindows = patched_destroy_all
    except Exception:
        # OpenCV not available; ignore
        return


def main() -> int:
    if len(sys.argv) < 3:
        print("Usage: python -m auto_docx.runner <script_path> <images_dir>", file=sys.stderr)
        return 1

    script_path = Path(sys.argv[1]).resolve()
    images_dir = Path(sys.argv[2]).resolve()

    os.environ.setdefault("MPLBACKEND", "Agg")
    images_dir.mkdir(parents=True, exist_ok=True)

    counter = [0]
    _patch_matplotlib(images_dir, counter)
    _patch_cv2(images_dir, counter)

    try:
        runpy.run_path(str(script_path), run_name="__main__")
        return 0
    except SystemExit as exc:
        return int(getattr(exc, "code", 0) or 0)
    except Exception:
        traceback.print_exc()
        return 1


if __name__ == "__main__":
    sys.exit(main())
