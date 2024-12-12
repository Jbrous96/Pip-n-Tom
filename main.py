import typing
from dataclasses import dataclass
from pathlib import Path
import shelve
import win32api
import win32process
import win32con
from win32com.client import GetObject

class RemotePythonError(Exception):
    """Base exception class for RemotePython errors"""
    pass

class InjectionError(RemotePythonError):
    """Raised when DLL injection fails"""
    pass

class InitializationError(RemotePythonError):
    """Raised when Python initialization fails"""
    pass

@dataclass
class ProcessInfo:
    pid: int
    handle: int
    executable: str

class RemotePython:
    """
    A class for managing Python execution in remote processes.
    
    This class provides capabilities to:
    - Inject Python DLL into remote processes
    - Execute Python code in the context of remote processes
    - Manage remote Python interpreter state
    - Cache procedure addresses for performance
    
    Args:
        pid (int): Process ID of the target process
        python_dll_path (str, optional): Path to Python DLL. Defaults to current Python version
        shelf_name (str, optional): Cache file for procedure addresses. Defaults to None
        
    Example:
        python = RemotePython(pid=1234)
        python.run("import sys; print(sys.version)")
    """
    
    def __init__(
        self,
        pid: int,
        python_dll_path: str = None,
        shelf_name: str = None
    ):
        self._process = self._initialize_process(pid)
        self._dll_path = python_dll_path or self._get_default_python_dll()
        self._proc_cache = self._initialize_cache(shelf_name)
        
        if not self._inject_and_initialize():
            raise InitializationError("Failed to initialize remote Python environment")
            
    def _initialize_process(self, pid: int) -> ProcessInfo:
        """Initialize process handle and information"""
        try:
            handle = win32api.OpenProcess(
                win32con.PROCESS_ALL_ACCESS,
                False,
                pid
            )
            exe_path = win32process.GetModuleFileNameEx(handle, 0)
            return ProcessInfo(pid=pid, handle=handle, executable=exe_path)
        except Exception as e:
            raise RemotePythonError(f"Failed to initialize process: {e}")

    def _inject_and_initialize(self) -> bool:
        """Inject Python DLL and initialize interpreter"""
        try:
            self._remote_dll = self._inject_dll(self._dll_path)
            if not self.is_python_injected():
                raise InjectionError("Python DLL injection failed")
                
            self._initialize_python()
            if not self.is_python_initialized():
                raise InitializationError("Python initialization failed")
                
            # Cache Py_Finalize for cleanup
            self._get_proc_address("Py_Finalize")
            return True
            
        except Exception as e:
            self.cleanup()
            raise RemotePythonError(f"Injection/initialization failed: {e}")

    def run(self, code: str) -> int:
        """
        Execute Python code in the remote process.
        
        Args:
            code (str): Python code to execute
            
        Returns:
            int: Exit code from execution
            
        Raises:
            RemotePythonError: If execution fails
        """
        if not isinstance(code, str):
            raise TypeError("Code must be a string")
            
        try:
            return self._execute_remote("PyRun_SimpleString", code)
        except Exception as e:
            raise RemotePythonError(f"Code execution failed: {e}")

    def set_python_path(self, paths: typing.List[str]):
        """
        Set the Python path in remote process.
        
        Args:
            paths: List of paths to add to sys.path
        """
        path_str = ";".join(str(Path(p)) for p in paths)
        self._execute_remote("PySys_SetPath", path_str)

    def cleanup(self):
        """Clean up resources and remove Python from remote process"""
        if hasattr(self, '_remote_dll'):
            try:
                self._execute_remote("Py_Finalize")
                self._unload_dll(self._remote_dll)
            except Exception as e:
                if hasattr(self, '_debug') and self._debug:
                    print(f"Cleanup error: {e}")
                    
        self._close_process()
        self._close_cache()

    def __enter__(self):
        return self
        
    def __exit__(self, exc_type, exc_val, exc_tb):
        self.cleanup()

    def __call__(self, code: str) -> int:
        """Shorthand for run()"""
        return self.run(code)
