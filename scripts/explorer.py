"""Prints a quick profile of every raw Excel file."""
from pathlib import Path

import pandas as pd

RAW_DIR = Path("data/raw")

pd.set_option("display.width", 200)
pd.set_option("display.max_columns", None)


def profile(path):
    print("=" * 70)
    print(path.name)
    print("=" * 70)
    df = pd.read_excel(path)

    rows, cols = df.shape
    print(f"\nRows: {rows}, columns: {cols}")

    print("\nColumns:", list(df.columns))

    print("\nFirst 5 rows:")
    print(df.head(5))

    print("\nEmpty values per column:")
    print(df.isna().sum())

    for col in df.columns:
        print(f"\nTop 10 values in {col!r}:")
        print(df[col].value_counts().head(10))
    print()


def main():
    for path in sorted(RAW_DIR.glob("*.xlsx")):
        profile(path)


if __name__ == "__main__":
    main()