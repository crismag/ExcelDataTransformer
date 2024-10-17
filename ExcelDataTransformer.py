import pandas as pd
import json
import yaml
import argparse
import os
import glob
from typing import Tuple, List, Dict, Union, Any


class ExcelDataTransformer:
    def __init__(self, input_file: str = None):
        """
        Initialize parser with input Excel file.
        Parsing information like 'HEADER_KEYWORD_TEXT' and data structure will be configured later.
        """
        self.input_file: Union[str, None] = input_file
        self.df: Union[pd.DataFrame, None] = None
        self.header_keyword: Union[str, None] = None
        self.start_table: int = 0
        self.end_table: Union[int, None] = None
        self.data_structure: Dict[str, Any] = {'DATA_GROUP_COLLECTION': {}}  # Default structure
        self.base_report_path: Union[str, None] = None
        self.filename_pattern: Union[str, None] = None

    def configure(self, **config: Dict[str, Any]) -> None:
        """
        Configure parsing information before performing actions.
        Configurations are passed as a dictionary of key-value pairs.
        """
        self.header_keyword = config.get('header_keyword', 'HEADER_KEYWORD_TEXT')
        self.start_table = config.get('start_table', 0)
        self.end_table = config.get('end_table', None)
        self.base_report_path = config.get('base_report_path', None)
        self.filename_pattern = config.get('filename_pattern', None)
        self.data_structure = config.get('data_structure', {'DATA_GROUP_COLLECTION': {}})

    def find_xlsx_file(self, data_group: str, category: str) -> str:
        """Find XLSX file using a configurable path pattern."""
        if not self.base_report_path or not self.filename_pattern:
            raise ValueError("Base report path or filename pattern is not configured.")

        pattern = os.path.join(self.base_report_path, self.filename_pattern.format(
            data_group=data_group, category=category))
        matches = glob.glob(pattern)

        if not matches:
            raise FileNotFoundError(f"No files found matching {pattern}")
        if len(matches) > 1:
            raise ValueError(f"Multiple files found for {pattern}")
        return matches[0]

    def _load_excel(self) -> pd.DataFrame:
        """Load Excel file and validate structure based on configuration."""
        try:
            df = pd.read_excel(self.input_file, sheet_name=0)
        except Exception as e:
            raise ValueError(f"Failed to load Excel file: {e}")

        process_rows = df[df.iloc[:, 0] == self.header_keyword].index.tolist()
        if not process_rows:
            raise ValueError(f"No '{self.header_keyword}' headers found in {self.input_file}")

        tables = [
            df.iloc[process_rows[i] + 1: process_rows[i + 1]
                    if i + 1 < len(process_rows) else len(df)]
            .reset_index(drop=True)
            .assign(**{df.columns[0]: df.iloc[process_rows[i]]})
            for i in range(len(process_rows))
        ]

        self.df = pd.concat(tables, ignore_index=True).iloc[self.start_table:self.end_table]
        return self.df

    def filter_data(self, select_columns: Union[str, None], 
                    where_clause: Union[str, None]) -> pd.DataFrame:
        """Filter data by selected columns and optional where clause."""
        filtered_df = self.df if not where_clause else self.df.query(where_clause)
        return filtered_df[select_columns.split(",")] if select_columns else filtered_df

    def output_data(self, data: pd.DataFrame, output_format: str) -> Union[str, List[Dict[str, Any]]]:
        """Format filtered data as JSON, YAML, or CSV."""
        if output_format == 'json':
            return json.dumps(data.to_dict(orient='records'), indent=2)
        elif output_format == 'yaml':
            return yaml.dump(data.to_dict(orient='records'), default_flow_style=False)
        elif output_format == 'csv':
            return data.to_csv(index=False)
        else:
            raise ValueError(f"Unsupported format: {output_format}")

    def show_headers(self) -> None:
        """Print all headers from the loaded DataFrame."""
        if self.df is None:
            print("DataFrame not loaded. Run the configure and _load_excel method first.")
        else:
            print("Available Headers:")
            for header in self.df.columns:
                print(header)


def detect_file_format(file_path: str) -> str:
    """Detect the format of the existing file (JSON, YAML, or CSV)."""
    with open(file_path, 'r') as f:
        first_line = f.readline().strip()

    if not first_line:
        raise ValueError(f"File {file_path} is empty or corrupted")

    if first_line.startswith('{'):
        return 'json'
    elif first_line.startswith('---'):
        return 'yaml'
    elif ',' in first_line or first_line.lower().startswith('sep='):
        return 'csv'
    else:
        raise ValueError("Unknown file format")


def update_output_file(output_file: str, data: Any, parser: ExcelDataTransformer, 
                       data_group: str, category: str) -> None:
    """Update or insert data in output file (JSON/YAML)."""
    if os.path.exists(output_file):
        try:
            output_format = detect_file_format(output_file)
        except ValueError as e:
            print(f"Error detecting file format: {e}")
            return

        if output_format == 'csv':
            raise ValueError("Cannot update CSV files incrementally")

        try:
            with open(output_file, 'r') as f:
                content = json.loads(f.read()) if output_format == 'json' else yaml.safe_load(f)

            content.setdefault(parser.data_structure['DATA_GROUP_COLLECTION'], {}).setdefault(
                data_group, {})[category] = data
        except Exception as e:
            raise ValueError(f"Error reading or parsing {output_file}: {e}")
    else:
        content = {parser.data_structure: {data_group: {category: data}}}
        output_format = 'json'

    with open(output_file, 'w') as f:
        if output_format == 'json':
            json.dump(content, f, indent=2)
        elif output_format == 'yaml':
            yaml.dump(content, f, default_flow_style=False)


def create_argparser() -> argparse.ArgumentParser:
    """Create argument parser."""
    parser = argparse.ArgumentParser(description="Parse and filter Excel data.")
    group = parser.add_mutually_exclusive_group(required=True)
    group.add_argument('-i', '--input', help="Input Excel file path")
    group.add_argument('--base_report_path', help="Base path to construct file location")

    parser.add_argument('--data_group', help="Data group name")
    parser.add_argument('--category', help="Category name")
    parser.add_argument('--select', help="Columns to select, e.g., 'HENq,pTYLE' (optional)")
    parser.add_argument('--where', help="Filter condition, e.g., 'DOG == \"DOG_t\" and SIZE == 25' (optional)")
    parser.add_argument('--output', help="Output file to update/insert content")
    parser.add_argument('--show_headers', action='store_true', help="Print headers only")

    return parser


def main() -> None:
    """Main function to process Excel data."""
    args = create_argparser().parse_args()

    config = {
        'header_keyword': 'HEADER_KEYWORD_TEXT',
        'start_table': 0,
        'end_table': None,
        'base_report_path': args.base_report_path,
        'filename_pattern': "project/xml_data/{data_group}/{category}/report/{category}_*_meas.xlsx"
    }

    parser = ExcelDataTransformer()
    parser.configure(**config)

    try:
        input_file = args.input or parser.find_xlsx_file(args.data_group, args.category)
    except (FileNotFoundError, ValueError) as e:
        print(f"Error: {e}")
        return

    parser.input_file = input_file

    try:
        parser._load_excel()
    except Exception as e:
        print(f"Error loading Excel file: {e}")
        return

    if args.show_headers:
        parser.show_headers()
        return

    if not args.output:
        print("Error: --output is required unless --show_headers is used.")
        return

    try:
        filtered_data = parser.filter_data(select_columns=args.select, where_clause=args.where)
        update_output_file(args.output, filtered_data, parser, args.data_group, args.category)

        output_format = detect_file_format(args.output) if os.path.exists(args.output) else 'json'
        print(parser.output_data(filtered_data, output_format))

    except Exception as e:
        print(f"Error: {e}")


if __name__ == "__main__":
    main()
