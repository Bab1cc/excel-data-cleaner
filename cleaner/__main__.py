"""Command line entry point: python -m cleaner [input_folder] [-o output.xlsx]"""
import argparse
import sys

from cleaner.output import OUTPUT_PATH, write_workbook
from cleaner.pipeline import run
from cleaner.reader import RAW_DIR


def main(argv=None):
    """Runs the whole cleaning process. Returns 0 on success, 1 on an error.
    argv=None means: take the arguments from the command line."""
    parser = argparse.ArgumentParser(prog="python -m cleaner",
                                     description="Clean and merge order exports from Excel files.")
    parser.add_argument("input", nargs="?", default=RAW_DIR,
                        help="folder with the .xlsx exports (default: %(default)s)")
    parser.add_argument("-o", "--output", default=OUTPUT_PATH,
                        help="where to save the result (default: %(default)s)")
    args = parser.parse_args(argv)

    # Expected problems (unknown column, no files, missing folder...) get a short
    # message instead of a traceback. Anything else is a bug and should crash loudly.
    try:
        orders, rejected, report = run(args.input)
    except (ValueError, FileNotFoundError) as error:
        print(f"Error: {error}", file=sys.stderr)
        return 1

    path = write_workbook(orders, rejected, report, args.output)
    print(f"Read {report['rows_read']} rows from {args.input}")
    print(f"  {report['rows_clean']} clean orders")
    print(f"  {report['rows_rejected']} rejected rows (see the Rejected sheet)")
    print(f"Saved {path}")
    return 0


if __name__ == "__main__":
    sys.exit(main())