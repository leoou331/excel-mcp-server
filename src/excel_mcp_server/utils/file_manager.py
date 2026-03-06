"""Safe file management with locking and context managers"""

import logging
import time
from contextlib import contextmanager
from pathlib import Path
from typing import Generator, Optional

from filelock import FileLock, Timeout
from openpyxl import Workbook, load_workbook
from openpyxl.workbook import Workbook as WorkbookType

from ..config import settings
from .security import SecurityError

logger = logging.getLogger(__name__)


class FileManager:
    """Manages Excel file operations with safety guarantees."""
    
    @contextmanager
    def safe_open(
        self,
        file_path: str | Path,
        read_only: bool = False,
        data_only: bool = False,
        timeout: Optional[float] = None,
    ) -> Generator[WorkbookType, None, None]:
        """Safely open an Excel workbook with file locking."""
        path = Path(file_path).resolve()
        timeout = timeout or settings.security.lock_timeout_seconds
        
        # Create lock file path
        lock_dir = path.parent / ".excel_locks"
        lock_dir.mkdir(exist_ok=True)
        lock_path = lock_dir / f"{path.name}.lock"
        lock = FileLock(str(lock_path), timeout=timeout)
        
        wb: Optional[WorkbookType] = None
        lock_acquired = False
        
        try:
            lock.acquire(blocking=True)
            lock_acquired = True
            logger.debug(f"Lock acquired for {path}")
            
            if path.exists():
                wb = load_workbook(
                    str(path),
                    read_only=read_only,
                    data_only=data_only,
                    keep_links=False,
                )
            elif not read_only:
                wb = Workbook()
            else:
                raise SecurityError(f"File not found: {path}")
            
            yield wb
            
        except Timeout:
            logger.error(f"Failed to acquire lock for {path}")
            raise SecurityError(f"File is locked by another process: {path}")
        except Exception as e:
            logger.exception(f"Error opening workbook {path}")
            raise
        finally:
            if wb is not None:
                try:
                    wb.close()
                    logger.debug(f"Workbook closed: {path}")
                except Exception as e:
                    logger.warning(f"Error closing workbook: {e}")
            
            if lock_acquired:
                try:
                    lock.release()
                    logger.debug(f"Lock released for {path}")
                except Exception as e:
                    logger.warning(f"Error releasing lock: {e}")
    
    @contextmanager
    def safe_write(
        self,
        file_path: str | Path,
        timeout: Optional[float] = None,
    ) -> Generator[WorkbookType, None, None]:
        """Safely open an Excel workbook for writing with auto-save."""
        path = Path(file_path).resolve()
        
        with self.safe_open(path, read_only=False, data_only=False, timeout=timeout) as wb:
            try:
                yield wb
                wb.save(str(path))
                logger.info(f"Workbook saved: {path}")
            except Exception as e:
                logger.error(f"Error during write, changes not saved: {e}")
                raise
    
    def safe_create(
        self,
        file_path: str | Path,
        timeout: Optional[float] = None,
    ) -> bool:
        """Create a new Excel workbook. Returns True if created."""
        path = Path(file_path).resolve()
        
        if path.exists():
            return False
        
        with self.safe_write(path, timeout=timeout) as wb:
            pass  # Workbook created automatically
        
        return True


file_manager = FileManager()
