import os
from datetime import datetime

import matplotlib.pyplot as plt
import numpy as np
import pandas as pd

# Full file path for input and output
input_file_path = "/Users/alexharrod/Library/CloudStorage/Dropbox/LogoIncluded/Marketing/2025 Marketing Digital Marketing Tracker.xlsx"
output_dir = "/Users/alexharrod/Library/CloudStorage/Dropbox/LogoIncluded/Marketing/Report outputs for Python"

# Read the Excel file
df = pd.read_excel(input_file_path, sheet_name="Performance vs Spend")

# List of rows to track
rows_to_track = ["SP", "MG", "FN", "AT", "JB", "WC", "EB", "HW", "CB", "CM"]


# Function to extract data for a specific row
def extract_row_data(df, row_name):
    row = df[df.iloc[:, 0] == row_name]
    if not row.empty:
        # Convert row to series and remove the first column (row name)
        series = row.iloc[0, 1:].replace("", 0).astype(float)
        return series
    return None


# Extract data for specified rows and spend
data_dict = {}
for row in rows_to_track + ["ESP Spend", "Sage Spend"]:
    row_data = extract_row_data(df, row)
    if row_data is not None:
        data_dict[row] = row_data

# Convert to DataFrame
data_df = pd.DataFrame(data_dict)

# Remove columns with all zeros
data_df = data_df.loc[:, (data_df != 0).any(axis=0)]

# Identify the first date with data
month_columns = data_df.index
first_date = str(month_columns[0])

# Debug print to see what date we're starting with
print("Original date from DataFrame:", first_date)

# Create base output filename with forced 2025 year
try:
    date_obj = pd.to_datetime(first_date)
    month = date_obj.month
    base_filename = f"adspend_vs_category_sales_{month:02d}-2025"
except Exception as e:
    print(f"Date parsing error: {e}")
    base_filename = "adspend_vs_category_sales"

# Prepare months for x-axis
months = data_df.index

# Calculate number of pages needed (2 charts per page)
charts_per_page = 2
total_pages = (len(rows_to_track) + charts_per_page - 1) // charts_per_page

# Create separate pages with 2 charts each
for page in range(total_pages):
    plt.figure(figsize=(20, 24))  # Adjusted for 2 charts per page
    plt.subplots_adjust(hspace=0.3)

    # Process 2 charts for current page
    for i in range(charts_per_page):
        chart_index = page * charts_per_page + i
        if chart_index < len(rows_to_track):
            channel = rows_to_track[chart_index]

            plt.subplot(2, 1, i + 1)

            # Plot the specific channel's performance
            plt.plot(
                months,
                data_df[channel],
                marker="o",
                label=channel,
                color="red",
                linewidth=3,
            )

            # Plot ESP and Sage Spend
            plt.plot(
                months,
                data_df["ESP Spend"],
                marker="o",
                label="ESP Spend",
                color="blue",
                linestyle="--",
                linewidth=3,
            )
            plt.plot(
                months,
                data_df["Sage Spend"],
                marker="o",
                label="Sage Spend",
                color="green",
                linestyle="--",
                linewidth=3,
            )

            plt.title(
                f"{channel} Performance with Spend", fontsize=18, fontweight="bold"
            )
            plt.xlabel("Months", fontsize=16)
            plt.ylabel("Value", fontsize=16)
            plt.xticks(rotation=45, fontsize=14)
            plt.yticks(fontsize=14)
            plt.grid(True, linestyle=":", alpha=0.7)
            plt.legend(fontsize=14)

    plt.suptitle(
        "Marketing Channel Performance Alongside ESP and Sage Spend",
        fontsize=22,
        fontweight="bold",
        y=0.95,
    )

    # Save each page as a separate file
    output_filename = f"{base_filename}_page{page+1}.png"
    output_path = os.path.join(output_dir, output_filename)
    plt.savefig(output_path, bbox_inches="tight", dpi=400)
    plt.close()  # Close the figure to free memory
    print(f"Saved page {page+1} to: {output_path}")

# Print summary statistics for each channel
print("\nPerformance Summary by Channel:")
for channel in rows_to_track:
    print(f"\n{channel} Channel:")
    print(data_df[channel].describe())

# Calculate correlation between each channel and spend
print("\nCorrelation between Channels and Spend:")
correlations = pd.DataFrame(
    {
        "Channel": rows_to_track,
        "ESP Spend Correlation": [
            data_df[channel].corr(data_df["ESP Spend"]) for channel in rows_to_track
        ],
        "Sage Spend Correlation": [
            data_df[channel].corr(data_df["Sage Spend"]) for channel in rows_to_track
        ],
    }
)
print(correlations)
