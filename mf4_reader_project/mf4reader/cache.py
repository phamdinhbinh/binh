"""Caching and file fingerprint helpers."""

from __future__ import annotations

import hashlib
from pathlib import Path
from typing import Union

PathLike = Union[str, bytes, Path]


def file_hash(path: PathLike, chunk_size: int = 1024 * 1024) -> str:
    """Return the SHA-256 hex digest for *path*.

    The file is streamed in chunks so large MF4/BLF files do not need to be
    loaded into memory at once.
    """
    target = Path(path)
    digest = hashlib.sha256()
    with target.open("rb") as handle:
        while True:
            chunk = handle.read(chunk_size)
            if not chunk:
                break
            digest.update(chunk)
    return digest.hexdigest()
