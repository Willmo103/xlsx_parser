import argparse


def setup_cli_args():
    """
    Set up the command line interface arguments using argparse.

    Returns:
    - argparse.ArgumentParser: Argument parser with defined arguments.
    """
    # Initialize the argument parser
    parser = argparse.ArgumentParser(
        description="Manipulate Excel data based on command line arguments."
    )
    # Define the mandatory file path argument with validation

    def valid_file(path):
        if not path.endswith(".xlsx"):
            raise argparse.ArgumentTypeError(f"File {path} is not a valid .xlsx file.")
        return path

    parser.add_argument("file_path", type=valid_file, help="Path to the .xlsx file.")

    # Define the optional output file path argument with default value
    parser.add_argument(
        "-o",
        "--output",
        type=str,
        default="parsed_output.xlsx",
        help="Output file path. Default is 'parsed_output.xlsx'.",
    )

    # Define the filter argument with sub-options
    filter_group = parser.add_argument_group(
        "filter", "Filter data based on specific criteria."
    )
    filter_group.add_argument(
        "-f",
        "--filter",
        choices=["rows", "columns"],
        help="Filter data based on rows or columns.",
    )
    filter_group.add_argument("--name", type=str, help="Specify name for filtering.")
    filter_group.add_argument("--value", type=str, help="Specify value for filtering.")
    filter_group.add_argument("--color", type=str, help="Specify color for filtering.")
    filter_group.add_argument(
        "--range",
        type=str,
        help="Specify range for filtering in the format 'start:end'.",
    )

    # Define the list argument with sub-options
    list_group = parser.add_argument_group(
        "list", "List specific data from the Excel file."
    )
    list_group.add_argument(
        "-l",
        "--list",
        choices=["rows", "columns", "all", "column_headers"],
        help="List specific data.",
    )

    # Define the sort_by argument with sub-options
    sort_group = parser.add_argument_group(
        "sort_by", "Sort data based on specific criteria."
    )
    sort_group.add_argument(
        "-s",
        "--sort_by",
        choices=["column_value", "column_header"],
        help="Sort data based on column values or headers.",
    )
    sort_group.add_argument(
        "--order",
        choices=["a", "d"],
        default="a",
        help="Specify sort order. 'a' for ascending and 'd' for descending.",
    )

    return parser


# If the script is run directly, parse the arguments and print them for demonstration purposes
if __name__ == "__main__":
    args_parser = setup_cli_args()
    args = args_parser.parse_args()
    print(args)
