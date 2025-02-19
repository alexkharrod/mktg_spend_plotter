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

# Create output filename with forced 2025 year
try:
    # Parse the datetime
    date_obj = pd.to_datetime(first_date)
    # Force the year to 2025 while keeping the original month
    month = date_obj.month
    output_filename = f"adspend_vs_category_sales_{month:02d}-2025.png"
except Exception as e:
    print(f"Date parsing error: {e}")
    output_filename = "adspend_vs_category_sales.png"

# Debug print to verify filename
print(f"Generated filename: {output_filename}")

# Full output path
output_path = os.path.join(output_dir, output_filename)

# Prepare months for x-axis
months = data_df.index

# Create a multi-panel figure with larger size
plt.figure(figsize=(30, 35))  # Increased from (25, 30)
plt.subplots_adjust(hspace=0.5, wspace=0.3)  # Increased vertical spacing

# Color palette for consistent coloring
colors = plt.cm.Set1(np.linspace(0, 1, len(rows_to_track)))

# Create a subplot for each marketing channel
for i, channel in enumerate(rows_to_track, 1):
    plt.subplot(5, 2, i)

    # Plot the specific channel's performance with thicker lines
    plt.plot(
        months,
        data_df[channel],
        marker="o",
        label=channel,
        color="red",
        linewidth=3,  # Increased linewidth
    )

    # Plot ESP and Sage Spend on the same chart with thicker lines
    plt.plot(
        months,
        data_df["ESP Spend"],
        marker="o",
        label="ESP Spend",
        color="blue",
        linestyle="--",
        linewidth=3,  # Increased linewidth
    )
    plt.plot(
        months,
        data_df["Sage Spend"],
        marker="o",
        label="Sage Spend",
        color="green",
        linestyle="--",
        linewidth=3,  # Increased linewidth
    )

    plt.title(
        f"{channel} Performance with Spend", fontsize=16, fontweight="bold"
    )  # Increased from 12
    plt.xlabel("Months", fontsize=14)  # Increased from 10
    plt.ylabel("Value", fontsize=14)  # Increased from 10
    plt.xticks(rotation=45, fontsize=12)  # Added fontsize
    plt.yticks(fontsize=12)  # Added fontsize
    plt.grid(True, linestyle=":", alpha=0.7)
    plt.legend(fontsize=12)  # Added fontsize to legend

plt.suptitle(
    "Marketing Channel Performance Alongside ESP and Sage Spend",
    fontsize=20,  # Increased from 16
    fontweight="bold",
    y=0.95,  # Adjusted position of main title
)

# Ensure output directory exists
os.makedirs(output_dir, exist_ok=True)

# Save the figure with higher DPI
plt.savefig(output_path, bbox_inches="tight", dpi=400)  # Increased DPI from 300
print(f"File saved to: {output_path}")

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
