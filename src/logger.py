"""
Logging utilities for the CV Research Experience Manager.
Stores logs locally under user's logs directory.
"""

import json
import csv
from pathlib import Path
from datetime import datetime
from typing import List, Dict, Any, Optional
from dataclasses import asdict

from models import LogEntry, OperationResult
from config import get_config, AppConfig


class OperationLogger:
    """Logger for CV modification operations."""
    
    def __init__(self, user_id: Optional[str] = None, config: Optional[AppConfig] = None):
        self.config = config or get_config()
        self.user_id = user_id or self.config.get_user_id()
        self.logs_path = self.config.get_user_logs_path(self.user_id)
        self.logs_path.mkdir(parents=True, exist_ok=True)
        
        self.entries: List[LogEntry] = []
        self.operation_name: str = ""
        self.start_time: datetime = datetime.now()
    
    def start_operation(self, operation_name: str) -> None:
        """Start a new operation log."""
        self.operation_name = operation_name
        self.start_time = datetime.now()
        self.entries = []
    
    def log(
        self,
        operation: str,
        phase: str = "",
        subcategory: str = "",
        year: int = 0,
        sponsor: str = "",
        protocol: str = "",
        details: str = ""
    ) -> LogEntry:
        """Log a single operation entry."""
        entry = LogEntry(
            timestamp=datetime.now(),
            operation=operation,
            phase=phase,
            subcategory=subcategory,
            year=year,
            sponsor=sponsor,
            protocol=protocol,
            details=details,
        )
        self.entries.append(entry)
        return entry
    
    def log_inserted(self, phase: str, subcategory: str, year: int, sponsor: str, protocol: str = "", details: str = "") -> LogEntry:
        return self.log("inserted", phase, subcategory, year, sponsor, protocol, details)
    
    def log_matched_existing(self, phase: str, subcategory: str, year: int, sponsor: str, protocol: str = "", details: str = "") -> LogEntry:
        return self.log("matched-existing", phase, subcategory, year, sponsor, protocol, details)
    
    def log_skipped_duplicate(self, phase: str, subcategory: str, year: int, sponsor: str, protocol: str = "", details: str = "") -> LogEntry:
        return self.log("skipped-duplicate", phase, subcategory, year, sponsor, protocol, details)
    
    def log_replaced(self, phase: str, subcategory: str, year: int, sponsor: str, protocol: str = "", details: str = "") -> LogEntry:
        return self.log("replaced", phase, subcategory, year, sponsor, protocol, details)
    
    def log_skipped_no_match(self, phase: str, subcategory: str, year: int, sponsor: str, protocol: str = "", details: str = "") -> LogEntry:
        return self.log("skipped-no-match", phase, subcategory, year, sponsor, protocol, details)
    
    def log_ambiguous(self, phase: str, subcategory: str, year: int, sponsor: str, protocol: str = "", details: str = "") -> LogEntry:
        return self.log("ambiguous-below-threshold", phase, subcategory, year, sponsor, protocol, details)
    
    def log_no_changes(self, details: str = "") -> LogEntry:
        return self.log("no-changes", details=details)
    
    def get_summary(self) -> Dict[str, Any]:
        """Get summary of logged operations."""
        counts: Dict[str, int] = {}
        by_phase: Dict[str, Dict[str, int]] = {}
        by_year: Dict[int, Dict[str, int]] = {}
        
        for entry in self.entries:
            op = entry.operation
            counts[op] = counts.get(op, 0) + 1
            
            if entry.phase:
                if entry.phase not in by_phase:
                    by_phase[entry.phase] = {}
                by_phase[entry.phase][op] = by_phase[entry.phase].get(op, 0) + 1
            
            if entry.year:
                if entry.year not in by_year:
                    by_year[entry.year] = {}
                by_year[entry.year][op] = by_year[entry.year].get(op, 0) + 1
        
        return {
            "operation": self.operation_name,
            "start_time": self.start_time.isoformat(),
            "end_time": datetime.now().isoformat(),
            "total_entries": len(self.entries),
            "counts": counts,
            "by_phase": by_phase,
            "by_year": by_year,
        }
    
    def save_json(self, filename: Optional[str] = None) -> Path:
        """Save log entries as JSON."""
        if filename is None:
            timestamp = self.start_time.strftime("%Y%m%d_%H%M%S")
            safe_op = "".join(c if c.isalnum() or c == '_' else '_' for c in self.operation_name.lower())
            filename = f"{safe_op}_{timestamp}.json"
                
        filepath = self.logs_path / filename
        
        data = {
            "summary": self.get_summary(),
            "entries": [entry.to_dict() for entry in self.entries],
        }
        
        with open(filepath, 'w', encoding='utf-8') as f:
            json.dump(data, f, indent=2)
        
        return filepath
    
    def save_csv(self, filename: Optional[str] = None) -> Path:
        """Save log entries as CSV."""
        if filename is None:
            timestamp = self.start_time.strftime("%Y%m%d_%H%M%S")
            safe_op = "".join(c if c.isalnum() or c == '_' else '_' for c in self.operation_name.lower())
            filename = f"{safe_op}_{timestamp}.csv"
        
        filepath = self.logs_path / filename
        
        fieldnames = ['timestamp', 'operation', 'phase', 'subcategory', 'year', 'sponsor', 'protocol', 'details']
        
        with open(filepath, 'w', newline='', encoding='utf-8') as f:
            writer = csv.DictWriter(f, fieldnames=fieldnames)
            writer.writeheader()
            for entry in self.entries:
                writer.writerow(entry.to_dict())
        
        return filepath
    
    def to_result(self, success: bool, output_path: Optional[str] = None, error_message: str = "") -> OperationResult:
        """Convert log to OperationResult."""
        return OperationResult(
            success=success,
            output_path=output_path,
            log_entries=self.entries.copy(),
            summary=self.get_summary(),
            error_message=error_message,
        )


def log_access_denied(user_id: str, attempted_resource: str, config: Optional[AppConfig] = None) -> None:
    """Log an access denied event."""
    cfg = config or get_config()
    logs_path = cfg.get_user_logs_path(user_id)
    logs_path.mkdir(parents=True, exist_ok=True)
    
    log_file = logs_path / "access_denied.log"
    
    entry = {
        "timestamp": datetime.now().isoformat(),
        "user_id": user_id,
        "attempted_resource": attempted_resource,
        "action": "blocked",
    }
    
    with open(log_file, 'a', encoding='utf-8') as f:
        f.write(json.dumps(entry) + "\n")
