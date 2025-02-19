import os
from datetime import datetime

import matplotlib.pyplot as plt
import numpy as np
import pandas as pd
from matplotlib.backends.backend_pdf import PdfPages


def read_and_clean_data(file_path, sheet_name="Performance vs Spend"):
    """Read and clean the Excel data."""
    # Read the Excel file
    df = pd.read_excel(file_path, sheet_name=sheet_name)

    # Clean up the DataFrame
    clean_df = df.copy()

    # Set the first column as the index and give it a name
    clean_df.set_index(clean_df.columns[0], inplace=True)
    clean_df.index.name = "Category"

    # Remove any columns that contain 'Unnamed'
    clean_df = clean_df.loc[:, ~clean_df.columns.astype(str).str.contains("Unnamed")]

    # Convert string dates to datetime
    date_columns = []
    valid_columns = []

    for col in clean_df.columns:
        if isinstance(col, (pd.Timestamp, datetime)):
            date_columns.append(col)
            valid_columns.append(col)
        elif isinstance(col, str) and any(
            month in col.upper()
            for month in [
                "JAN",
                "FEB",
                "MAR",
                "APR",
                "MAY",
                "JUN",
                "JUL",
                "AUG",
                "SEP",
                "OCT",
                "NOV",
                "DEC",
            ]
        ):
            try:
                date = pd.to_datetime(col)
                date_columns.append(date)
                valid_columns.append(col)
            except:
                continue

    # Keep only valid date columns
    clean_df = clean_df[valid_columns]
    clean_df.columns = date_columns

    # Convert all values to numeric, replacing non-numeric values with NaN
    for col in clean_df.columns:
        clean_df[col] = pd.to_numeric(clean_df[col], errors="coerce")

    # Drop rows where all values are NaN
    clean_df = clean_df.dropna(how="all")

    # Sort columns by date
    clean_df = clean_df.reindex(sorted(clean_df.columns), axis=1)

    return clean_df


def create_visualization(df, output_path, rows_to_track):
    """Create visualization PDF with multiple pages."""
    with PdfPages(output_path) as pdf:
        # Create plots two at a time
        for i in range(0, len(rows_to_track), 2):
            fig, (ax1, ax2) = plt.subplots(2, 1, figsize=(12, 16))
            fig.suptitle("Marketing Performance Analysis", fontsize=16)

            # Plot data for both charts
            for idx, ax in enumerate([ax1, ax2]):
                if i + idx < len(rows_to_track):
                    category = rows_to_track[i + idx]
                    if category in df.index:
                        # Plot category data
                        ax.plot(
                            df.columns,
                            df.loc[category],
                            label=f"{category} Performance",
                            marker="o",
                        )

                        # Plot spend data
                        if "ESP Spend" in df.index:
                            ax.plot(
                                df.columns,
                                df.loc["ESP Spend"],
                                label="ESP Spend",
                                linestyle="--",
                                alpha=0.7,
                            )
                        if "Sage Spend" in df.index:
                            ax.plot(
                                df.columns,
                                df.loc["Sage Spend"],
                                label="Sage Spend",
                                linestyle="--",
                                alpha=0.7,
                            )

                        ax.set_title(f"{category} vs Marketing Spend")
                        ax.set_xlabel("Date")
                        ax.set_ylabel("Amount")
                        ax.legend()
                        ax.grid(True)
                        plt.setp(ax.xaxis.get_majorticklabels(), rotation=45)

            plt.tight_layout()
            if any(cat in df.index for cat in rows_to_track[i : i + 2]):
                pdf.savefig(fig)
            plt.close()


def main():
    # File paths
    input_file_path = "/Users/alexharrod/Library/CloudStorage/Dropbox/LogoIncluded/Marketing/2025 Marketing Digital Marketing Tracker.xlsx"
    output_dir = "/Users/alexharrod/Library/CloudStorage/Dropbox/LogoIncluded/Marketing/Report outputs for Python"
    os.makedirs(output_dir, exist_ok=True)

    # Read and clean data
    print("Reading and cleaning data...")
    df = read_and_clean_data(input_file_path)

    # Print debug information
    print("\nDataFrame Info:")
    print(df.info())
    print("\nDataFrame Head:")
    print(df.head())
    print("\nAvailable Categories:")
    print(df.index.tolist())

    # List of rows to track
    rows_to_track = ["SP", "MG", "FN", "AT", "JB", "WC", "EB", "HW", "CB", "CM"]

    # Create output filename
    output_filename = f"adspend_vs_category_sales_01-2025.pdf"
    output_path = os.path.join(output_dir, output_filename)

    # Create visualization
    print(f"\nCreating visualization at: {output_path}")
    create_visualization(df, output_path, rows_to_track)
    print("Visualization completed!")


if __name__ == "__main__":
    main()
