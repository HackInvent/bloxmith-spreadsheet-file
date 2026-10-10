"""Atomic, descriptor-relative publication of a completed file; no implicit overwrite."""

from __future__ import annotations

import os
from pathlib import Path
import stat
import uuid

from .logic import fail

SUBDIRECTORY = "spreadsheets"


def destination(config, storage, root_dir):
    """Open an explicitly configured existing directory, or the node's own storage folder."""
    if config["output_directory"]:
        try:
            path = Path(config["output_directory"]).expanduser()
            path = (path if path.is_absolute() else Path(root_dir) / path).resolve(strict=True)
        except (OSError, RuntimeError):
            fail("storage_unavailable", "output_directory does not exist; create it before writing.")
        if not path.is_dir():
            fail("storage_unavailable", "output_directory must be a directory.")
    else:
        if not callable(storage):
            fail("storage_unavailable", "Persistent block storage is unavailable for this run; set output_directory.")
        try:
            path = Path(storage()).resolve(strict=True) / SUBDIRECTORY
        except Exception:  # noqa: BLE001 - framework storage errors cross a process proxy as several types.
            fail("storage_unavailable", "Persistent block storage is unavailable for this run; set output_directory.")
        try:
            path.mkdir(mode=0o700)
        except FileExistsError:
            pass
        except OSError:
            fail("write_failed", "The block storage folder cannot be created.")
    try:
        return path, os.open(path, os.O_RDONLY | os.O_DIRECTORY | os.O_NOFOLLOW | getattr(os, "O_CLOEXEC", 0))
    except OSError:
        fail("storage_unavailable", "The output directory cannot be opened (symbolic links are refused).")


def publish(data, name, config, storage, root_dir, check):
    """Write a private temporary file, fsync it, then link (or, if allowed, replace) the final name."""
    path, directory = destination(config, storage, root_dir)
    temporary = ".spreadsheet-" + uuid.uuid4().hex + ".tmp"
    created = False
    try:
        try:
            descriptor = os.open(temporary, os.O_WRONLY | os.O_CREAT | os.O_EXCL | os.O_NOFOLLOW | getattr(os, "O_CLOEXEC", 0),
                                 0o600, dir_fd=directory)
            created = True
            with os.fdopen(descriptor, "wb") as handle:
                handle.write(data)
                handle.flush()
                os.fsync(handle.fileno())
        except OSError:
            fail("write_failed", "The temporary file cannot be written in the output directory.")
        check()
        try:
            current, opened = path.stat(), os.fstat(directory)
        except OSError:
            fail("write_failed", "The output directory disappeared during the write.")
        if (current.st_dev, current.st_ino) != (opened.st_dev, opened.st_ino):
            fail("write_failed", "The output directory changed during the write.")
        overwritten = False
        if config["overwrite"]:
            try:
                existing = os.stat(name, dir_fd=directory, follow_symlinks=False)
            except FileNotFoundError:
                existing = None
            except OSError:
                fail("write_failed", "The existing target cannot be inspected.")
            if existing is not None and not stat.S_ISREG(existing.st_mode):
                fail("target_exists", "The target exists and is not a regular file; it is never replaced.")
            try:
                os.replace(temporary, name, src_dir_fd=directory, dst_dir_fd=directory)
            except OSError:
                fail("write_failed", "The completed file cannot replace the target.")
            created, overwritten = False, existing is not None
        else:
            try:
                os.link(temporary, name, src_dir_fd=directory, dst_dir_fd=directory, follow_symlinks=False)
            except FileExistsError:
                fail("target_exists", "A file with this name already exists; enable overwrite or add {unique} to file_name.")
            except OSError:
                fail("write_failed", "The completed file cannot be published; the filesystem must support hard links.")
            try:
                os.unlink(temporary, dir_fd=directory)
                created = False
            except OSError:
                pass  # the final name is complete; the finally clause retries this cleanup
        try:
            os.fsync(directory)
        except OSError:
            pass
        return str(path / name), overwritten
    finally:
        if created:
            try:
                os.unlink(temporary, dir_fd=directory)
            except OSError:
                pass
        os.close(directory)
