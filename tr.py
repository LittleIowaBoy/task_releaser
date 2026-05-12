import pandas as pd
from typing import List, Callable, Dict, Any
from pathlib import Path


class ExcelParser:
    """
    A general-purpose Excel parser that reads data and filters based on conditions.
    """
    
    def __init__(self, filepath: str):
        """
        Initialize the parser with an Excel file.
        
        Args:
            filepath: Path to the Excel file
        """
        self.filepath = filepath
        self.df = None

    def list_sheets(self) -> list[str]:
        """Return Excel sheet names. Returns an empty list for CSV files
        or if the workbook cannot be opened."""
        if self.filepath.lower().endswith(".csv"):
            return []
        try:
            xls = pd.ExcelFile(self.filepath)
            return list(xls.sheet_names)
        except Exception:
            return []
        
    def read_excel(self, sheet_name: int | str = 0) -> pd.DataFrame:
        """
        Read an Excel or CSV file into a DataFrame.
        
        Args:
            sheet_name: Sheet index (0) or name (default: 0). Ignored for CSV files.
            
        Returns:
            DataFrame containing the data
        """
        try:
            if self.filepath.endswith('.csv'):
                self.df = pd.read_csv(self.filepath)
            else:
                self.df = pd.read_excel(self.filepath, sheet_name=sheet_name)
            return self.df
        except FileNotFoundError:
            return None
        except Exception:
            return None
    
    def filter_by_condition(self, column: str, condition: Callable[[Any], bool]) -> pd.DataFrame:
        """
        Filter rows where a condition is met on a specific column.
        
        Args:
            column: Column name to apply condition to
            condition: Function that returns True/False for each value
            
        Returns:
            Filtered DataFrame
        """
        if self.df is None:
            return None
        
        if column not in self.df.columns:
            return None
        
        try:
            return self.df[self.df[column].apply(condition)]
        except Exception:
            return None
    
    def filter_by_value(self, column: str, value: Any) -> pd.DataFrame:
        """
        Filter rows where a column equals a specific value.
        
        Args:
            column: Column name
            value: Value to match
            
        Returns:
            Filtered DataFrame
        """
        return self.filter_by_condition(column, lambda x: x == value)
    
    def filter_by_range(self, column: str, min_val: float = None, max_val: float = None) -> pd.DataFrame:
        """
        Filter rows where a column value is within a range.
        
        Args:
            column: Column name
            min_val: Minimum value (inclusive), None for no minimum
            max_val: Maximum value (inclusive), None for no maximum
            
        Returns:
            Filtered DataFrame
        """
        def range_condition(x):
            if pd.isna(x):
                return False
            if min_val is not None and x < min_val:
                return False
            if max_val is not None and x > max_val:
                return False
            return True
        
        return self.filter_by_condition(column, range_condition)
    
    def filter_by_contains(self, column: str, substring: str, case_sensitive: bool = False) -> pd.DataFrame:
        """
        Filter rows where a column contains a substring.
        
        Args:
            column: Column name
            substring: Substring to search for
            case_sensitive: Whether matching is case-sensitive
            
        Returns:
            Filtered DataFrame
        """
        def contains_condition(x):
            if pd.isna(x):
                return False
            x_str = str(x)
            if case_sensitive:
                return substring in x_str
            else:
                return substring.lower() in x_str.lower()
        
        return self.filter_by_condition(column, contains_condition)
    
    def get_values(self, dataframe: pd.DataFrame, column: str) -> List[Any]:
        """
        Extract values from a specific column in a DataFrame.
        
        Args:
            dataframe: Source DataFrame
            column: Column name to extract
            
        Returns:
            List of values from the column
        """
        if dataframe is None or dataframe.empty:
            return []
        
        if column not in dataframe.columns:
            return []
        
        return dataframe[column].tolist()
    
    def get_multiple_columns(self, dataframe: pd.DataFrame, columns: List[str]) -> Dict[str, List[Any]]:
        """
        Extract values from multiple columns in a DataFrame.
        Performs case-insensitive and flexible column matching.
        
        Args:
            dataframe: Source DataFrame
            columns: List of column names to extract
            
        Returns:
            Dictionary mapping column names to lists of values
        """
        if dataframe is None or dataframe.empty:
            return {}
        
        result = {}
        for column in columns:
            if column in dataframe.columns:
                result[column] = dataframe[column].tolist()
            else:
                matching_cols = [col for col in dataframe.columns if col.lower() == column.lower()]
                if matching_cols:
                    result[column] = dataframe[matching_cols[0]].tolist()
        
        return result

    def get_column_names(self) -> List[str]:
        """Get list of column names in the loaded data."""
        if self.df is None:
            return []
        return list(self.df.columns)

    def get_unique_numeric_values(self, column: str) -> List[int]:
        """Return all unique numeric values from ``column`` as ints.

        Used by templates that just want every row's ID without a comparison
        (e.g. the Locked Full Container Chase Tasks export).
        """
        if self.df is None or column not in self.df.columns:
            return []
        out: List[int] = []
        seen = set()
        for value in self.df[column].dropna():
            if not self._is_numeric(value):
                continue
            iv = int(float(value))
            if iv not in seen:
                seen.add(iv)
                out.append(iv)
        return out
    
    def _is_numeric(self, value: Any) -> bool:
        """
        Check if a value is numeric (int or float).
        
        Args:
            value: Value to check
            
        Returns:
            True if value is numeric, False otherwise
        """
        if pd.isna(value):
            return False
        try:
            float(value)
            return True
        except (ValueError, TypeError):
            return False
    
    def get_task_ids_where_condition(self, task_id_col: str, condition_col1: str, condition_col2: str, 
                                      comparison: str = ">=", item_col: str = "Item") -> tuple:
        """
        Get Task IDs where condition is met, and items from Task IDs that don't meet the condition.
        For Task IDs that appear multiple times, ALL instances must meet the condition to be returned.
        
        Args:
            task_id_col: Column name to extract values from (e.g., 'Task ID')
            condition_col1: First column for comparison (e.g., 'Active OHB')
            condition_col2: Second column for comparison (e.g., 'Allocated')
            comparison: Comparison operator as string: '>', '<', '>=', '<=', '==', '!='
            item_col: Column name for items to track (e.g., 'Item')
            
        Returns:
            Tuple of two objects:
                - task_ids: List of unique Task IDs where ALL instances meet the condition
                - items_not_met: Dict of items from Task IDs that don't meet condition, with count of affected Task IDs
        """
        if self.df is None or self.df.empty:
            return [], {}

        missing_cols = [
            col for col in [task_id_col, condition_col1, condition_col2, item_col]
            if col not in self.df.columns
        ]
        if missing_cols:
            return [], {}
        
        try:
            # Convert comparison columns to numeric where possible to avoid string/formatting issues
            s1 = pd.to_numeric(self.df[condition_col1], errors='coerce')
            s2 = pd.to_numeric(self.df[condition_col2], errors='coerce')

            # Create boolean mask based on comparison using numeric series (NaNs will be False)
            if comparison == ">":
                mask = s1 > s2
            elif comparison == "<":
                mask = s1 < s2
            elif comparison == ">=":
                mask = s1 >= s2
            elif comparison == "<=":
                mask = s1 <= s2
            elif comparison == "==":
                mask = s1 == s2
            elif comparison == "!=":
                mask = s1 != s2
            else:
                return [], {}
            
            # Add a temporary column to track which rows meet the condition
            self.df['__condition_met__'] = mask.fillna(False)
            
            # Group by Task ID and check if ALL instances meet the condition
            valid_task_ids = []
            invalid_task_ids = []
            
            for task_id in self.df[task_id_col].unique():
                task_id_rows = self.df[self.df[task_id_col] == task_id]
                # Check if all instances of this Task ID meet the condition
                if task_id_rows['__condition_met__'].all():
                    # Only add if task_id is numeric
                    if self._is_numeric(task_id):
                        valid_task_ids.append(int(task_id))
                else:
                    # Task ID does not meet condition
                    if self._is_numeric(task_id):
                        invalid_task_ids.append(int(task_id))
            
            # For items in Task IDs that don't meet the condition, count how many Task IDs each item affects
            items_not_met_dict = {}
            if invalid_task_ids:
                # Only consider rows that do NOT meet the comparison (i.e., mask is False)
                invalid_rows_df = self.df[~mask & self.df[task_id_col].isin(invalid_task_ids)]
                for item in invalid_rows_df[item_col].unique():
                    # Count how many unique Task IDs this item appears in among the failing rows
                    affected_task_ids = invalid_rows_df[invalid_rows_df[item_col] == item][task_id_col].unique()
                    items_not_met_dict[item] = len(affected_task_ids)
            
            # Remove temporary column
            self.df.drop('__condition_met__', axis=1, inplace=True)
            
            return valid_task_ids, items_not_met_dict
        
        except Exception:
            return [], {}


# Example usage
if __name__ == "__main__":
    # Find the most recently downloaded file
    downloads_path = Path.home() / "Downloads"
    
    # Get all CSV and Excel files from Downloads folder, sorted by modification time
    files = []
    for ext in ['*.csv', '*.xlsx', '*.xls']:
        files.extend(downloads_path.glob(ext))
    
    if files:
        # Sort by modification time (most recent first)
        most_recent_file = max(files, key=lambda f: f.stat().st_mtime)
        filepath = str(most_recent_file)
        print(f"Using most recently downloaded file: {most_recent_file.name}")
    else:
        print("Error: No CSV or Excel files found in Downloads folder")
        filepath = None
    
    if filepath:
        # Initialize parser with the most recent file
        parser = ExcelParser(filepath)
    
    # Read the file
    try:
        parser.read_excel()
    except:
        print("No file found")
    
    # Get Task IDs where Active OHB >= Allocated
    if parser.df is None:
        print("No data loaded")
    elif parser.df is not None:
        task_ids = parser.get_task_ids_where_condition(
            task_id_col="Task ID",
            condition_col1="Active OHB",
            condition_col2="Allocated",
            comparison=">="
        )
        
        print("\n--- Task IDs where Active OHB >= Allocated ---")
        print(task_ids)
    else:
        print("Failed to load data")
