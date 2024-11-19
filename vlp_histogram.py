import pandas as pd
import matplotlib.pyplot as plt
import numpy as np
from scipy.stats import gaussian_kde
from statistics import mode, StatisticsError
import os

# Prompt user for file path
file_path = input("Enter the file path to the Excel file: ")

# Validate file path
if not os.path.exists(file_path):
    print("Error: File does not exist. Please check the path and try again.")
    exit()

try:
    # Load the Excel file
    data = pd.read_excel(file_path)

    # Drop columns with all NaN values
    data = data.dropna(axis=1, how="all")

    # Normalize data for histograms and KDE
    fig, ax = plt.subplots(figsize=(10, 6))
    colors = plt.cm.plasma(np.linspace(0, 1, len(data.columns)))  # Use Plasma colormap for styling

    for i, column in enumerate(data.columns):
        # Filter valid data within the desired range
        valid_data = data[column].dropna()
        valid_data = valid_data[(valid_data >= 0) & (valid_data <= 515)]

        if valid_data.empty:
            continue  # Skip empty columns after filtering

        # Normalize the valid data
        max_val = valid_data.max() if valid_data.max() != 0 else 1  # Avoid division by zero
        normalized_data = valid_data / max_val

        # Plot normalized histogram
        ax.hist(
            normalized_data,
            bins=20,
            alpha=0.5,
            color=colors[i],
            edgecolor="black",
            label=f"{column} (Normalized Histogram)",
            density=True,
        )

        # Plot KDE for normalized data
        if len(normalized_data) > 1:  # Ensure KDE can be calculated
            density = gaussian_kde(normalized_data)
            x_vals = np.linspace(0, 1, 1000)
            ax.plot(
                x_vals,
                density(x_vals),
                color=colors[i],
                lw=2.5,
                label=f"{column} (KDE)",
            )

    # Set x-axis limits and ticks
    ax.set_xlim(0, 1)  # Normalized range
    ax.set_xticks(np.linspace(0, 1, 11))  # Tick marks at 0.1 intervals

    ax.set_title("Normalized Histograms and KDE for Filtered Samples (0-515 nm)", fontsize=14, fontweight="bold", color="#4B4B4B")
    ax.set_xlabel("Normalized Size", fontsize=12, fontweight="bold", color="#4B4B4B")
    ax.set_ylabel("Density", fontsize=12, fontweight="bold", color="#4B4B4B")
    ax.legend(loc="upper right", fontsize=10)
    ax.grid(True, linestyle="--", alpha=0.7)

    plt.tight_layout()

    # Calculate summary statistics for original (non-normalized) data
    stats = pd.DataFrame({
        "Mean": [round(data[column].mean(), 2) for column in data.columns],
        "Median": [round(data[column].median(), 2) for column in data.columns],
        "Mode": [
            round(mode(data[column].dropna()), 2) if len(data[column].dropna()) > 0 else "N/A"
            for column in data.columns
        ],
        "Minimum": [round(data[column].min(), 2) for column in data.columns],
        "Maximum": [round(data[column].max(), 2) for column in data.columns],
        "N (Count)": [data[column].count() for column in data.columns],
        "Std Dev": [round(data[column].std(), 2) for column in data.columns]
    }, index=data.columns)

    # Display the table below the graph
    fig, ax_table = plt.subplots(figsize=(12, 5))  # Create a new figure for the table
    ax_table.axis('tight')  # Adjust table size to fit
    ax_table.axis('off')  # Hide the axis

    # Create a styled table
    table = ax_table.table(
        cellText=stats.values,
        colLabels=stats.columns,
        rowLabels=stats.index,
        cellLoc="center",
        loc="center",
    )
    table.auto_set_font_size(False)
    table.set_fontsize(10)
    table.auto_set_column_width(col=list(range(len(stats.columns) + 1)))

    # Apply elegant styling
    header_color = "#1f77b480"  # Academic blue with transparency
    row_label_color = "#6d6d6d80"  # Neutral gray with transparency
    cell_background = "#f9f9f9"  # Slightly lighter gray for data cells
    text_color = "#333333"  # Dark gray for text

    for (row, col), cell in table.get_celld().items():
        if row == 0:
            cell.set_text_props(weight="bold", color="white")
            cell.set_facecolor(header_color)
        elif col == -1:
            cell.set_text_props(weight="bold", color="white")
            cell.set_facecolor(row_label_color)
        else:
            cell.set_facecolor(cell_background)

        cell.set_edgecolor("#dddddd")
        cell.set_text_props(color=text_color)
        cell.PAD = 0.5  # Adjust padding inside cells

    plt.show()

except StatisticsError:
    print("Error calculating mode: One or more columns have no unique mode.")
except Exception as e:
    print(f"An error occurred while processing the file: {e}")
