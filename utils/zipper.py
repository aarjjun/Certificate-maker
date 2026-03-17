"""
zipper.py
---------
Utility to compress a folder of generated certificate images into a
single ZIP archive for download.
"""

import os
import zipfile


def zip_certificates(source_dir: str, zip_path: str) -> str:
    """
    Add all PNG files in `source_dir` to a ZIP archive at `zip_path`.

    Args:
        source_dir: Directory containing the generated .png certificate files.
        zip_path:   Full path (including filename) for the output .zip file.

    Returns:
        The absolute path of the created ZIP file.
    """
    os.makedirs(os.path.dirname(zip_path), exist_ok=True)

    with zipfile.ZipFile(zip_path, "w", zipfile.ZIP_DEFLATED) as zf:
        for filename in os.listdir(source_dir):
            if filename.lower().endswith(".png"):
                full_path = os.path.join(source_dir, filename)
                # Store only the bare filename inside the ZIP (no folder path)
                zf.write(full_path, arcname=filename)

    return zip_path
