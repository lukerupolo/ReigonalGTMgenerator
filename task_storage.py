"""
Simple task storage for async processing to avoid 60-second timeout issues
"""
import uuid
import json
import os
import tempfile
from datetime import datetime
from typing import Dict, Any

# Global storage for task status
task_storage: Dict[str, Dict[str, Any]] = {}

class TaskManager:
    @staticmethod
    def create_task() -> str:
        """Create a new task with unique ID"""
        task_id = str(uuid.uuid4())
        task_storage[task_id] = {
            "id": task_id,
            "status": "started",
            "progress": 0,
            "message": "Task initialized",
            "created_at": datetime.now().isoformat(),
            "updated_at": datetime.now().isoformat(),
            "result_file": None,
            "error": None
        }
        return task_id
    
    @staticmethod
    def update_task(task_id: str, status: str = None, progress: int = None, 
                   message: str = None, result_file: str = None, error: str = None):
        """Update task status"""
        if task_id not in task_storage:
            return False
        
        task = task_storage[task_id]
        if status:
            task["status"] = status
        if progress is not None:
            task["progress"] = progress
        if message:
            task["message"] = message
        if result_file:
            task["result_file"] = result_file
        if error:
            task["error"] = error
        
        task["updated_at"] = datetime.now().isoformat()
        return True
    
    @staticmethod
    def get_task(task_id: str) -> Dict[str, Any]:
        """Get task status"""
        return task_storage.get(task_id, {"error": "Task not found"})
    
    @staticmethod
    def cleanup_task(task_id: str):
        """Remove task from storage"""
        if task_id in task_storage:
            task = task_storage[task_id]
            # Clean up result file if it exists
            if task.get("result_file") and os.path.exists(task["result_file"]):
                try:
                    os.remove(task["result_file"])
                except:
                    pass
            del task_storage[task_id]

# Create a temporary directory for storing result files
TEMP_DIR = tempfile.mkdtemp(prefix="presemulator_")
print(f"[TASK_STORAGE] Temporary directory created: {TEMP_DIR}")
