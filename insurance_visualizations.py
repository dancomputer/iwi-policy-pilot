import matplotlib.pyplot as plt
import pandas as pd
import numpy as np

def create_area_payout_chart(df_wide_numeric: pd.DataFrame, save_path: str = None):
    """
    Create a stacked bar chart showing "Area Payout (US Dollars) per Year"
    
    Parameters:
    - df_wide_numeric: The numeric regional statistics dataframe from build_regional_statistics
    - save_path: Optional path to save the chart (e.g., "chart.png")
    """
    
    # Color palette matching the Excel sheets
    area_colors_hex = {
        "Northern Zone": "#1F77B4",
        "Central Zone": "#2CA02C", 
        "Lake Zone": "#FF7F0E",
        "Western Zone": "#9467BD",
        "Southern Highlands Zone": "#8C564B",
        "Coastal Zone": "#17BECF",
        "Zanzibar (Islands)": "#7F7F7F",
    }
    
    # Extract year rows (those that are numeric strings representing years)
    year_rows = []
    for idx in df_wide_numeric.index:
        if str(idx).isdigit() and 1900 <= int(idx) <= 2100:
            year_rows.append(idx)
    
    year_rows = sorted([int(y) for y in year_rows])
    
    if not year_rows:
        print("No year data found in the dataframe")
        return
    
    # Find area total columns (those ending with "Total")
    area_columns = []
    for col in df_wide_numeric.columns:
        if str(col).endswith(" Total") and not str(col).startswith("Overall"):
            area_name = str(col).replace(" Total", "")
            if area_name in area_colors_hex:
                area_columns.append((col, area_name))
    
    if not area_columns:
        print("No area total columns found")
        return
    
    # Prepare data for stacking
    years = year_rows
    area_data = {}
    
    for col, area_name in area_columns:
        values = []
        for year in years:
            val = df_wide_numeric.at[str(year), col]
            # Handle NaN/None values
            if pd.isna(val) or val is None:
                values.append(0)
            else:
                values.append(float(val))
        area_data[area_name] = values
    
    # Create the stacked bar chart
    fig, ax = plt.subplots(figsize=(12, 8))
    
    # Bottom values for stacking
    bottom = np.zeros(len(years))
    
    # Plot each area
    for area_name in area_colors_hex.keys():
        if area_name in area_data:
            values = area_data[area_name]
            color = area_colors_hex[area_name]
            
            ax.bar(years, values, bottom=bottom, label=area_name, 
                   color=color, alpha=0.8, edgecolor='white', linewidth=0.5)
            bottom += values
    
    # Customize the chart
    ax.set_title("Area Payout (US Dollars) per Year", fontsize=16, fontweight='bold', pad=20)
    
    # Format y-axis to show values in a readable format
    ax.yaxis.set_major_formatter(plt.FuncFormatter(lambda x, p: f'${x:,.0f}'))
    
    # Set x-axis ticks to show all years and rotate labels vertically with bigger font
    ax.set_xticks(years)
    ax.set_xticklabels(years, rotation=90, fontsize=12)
    
    # Make y-axis tick labels bigger
    ax.tick_params(axis='y', labelsize=12)
    
    # Add legend at bottom center with no box and square markers
    ax.legend(bbox_to_anchor=(0.5, -0.10), loc='upper center', fontsize=12, ncol=3, 
              frameon=False, handlelength=1, handletextpad=0.5, 
              markerscale=0.8, markerfirst=True)
    
    # Update legend markers to be squares
    legend = ax.get_legend()
    for handle in legend.legend_handles:
        handle.set_width(10)
        handle.set_height(10)
    
    # Add grid for better readability
    ax.grid(True, alpha=0.3, axis='y')
    
    # Tight layout to prevent label cutoff
    plt.tight_layout()
    
    # Save if path provided
    if save_path:
        plt.savefig(save_path, dpi=300, bbox_inches='tight')
        print(f"Chart saved to: {save_path}")
    
    plt.show()
    return fig

def create_area_payout_percentage_chart(df_wide_numeric: pd.DataFrame, df_final: pd.DataFrame, save_path: str = None):
    """
    Create a stacked bar chart showing "Area Payout Percentage by Year"
    
    Parameters:
    - df_wide_numeric: The numeric regional statistics dataframe from build_regional_statistics
    - df_final: The original dataframe containing all farmer/pixel data
    - save_path: Optional path to save the chart (e.g., "chart.png")
    """
    
    # Color palette matching the Excel sheets
    area_colors_hex = {
        "Northern Zone": "#1F77B4",
        "Central Zone": "#2CA02C", 
        "Lake Zone": "#FF7F0E",
        "Western Zone": "#9467BD",
        "Southern Highlands Zone": "#8C564B",
        "Coastal Zone": "#17BECF",
        "Zanzibar (Islands)": "#7F7F7F",
    }
    
    # Extract year rows (those that are numeric strings representing years)
    year_rows = []
    for idx in df_wide_numeric.index:
        if str(idx).isdigit() and 1900 <= int(idx) <= 2100:
            year_rows.append(idx)
    
    year_rows = sorted([int(y) for y in year_rows])
    
    if not year_rows:
        print("No year data found in the dataframe")
        return
    
    # Find area total columns (those ending with "Total")
    area_columns = []
    for col in df_wide_numeric.columns:
        if str(col).endswith(" Total") and not str(col).startswith("Overall"):
            area_name = str(col).replace(" Total", "")
            if area_name in area_colors_hex:
                area_columns.append((col, area_name))
    
    if not area_columns:
        print("No area total columns found")
        return
    
    # Prepare data for stacking
    years = year_rows
    area_data = {}
    
    for col, area_name in area_columns:
        values = []
        for year in years:
            val = df_wide_numeric.at[str(year), col]
            # Handle NaN/None values
            if pd.isna(val) or val is None:
                values.append(0)
            else:
                values.append(float(val))
        area_data[area_name] = values
    
    # Calculate percentages for each year
    area_percentages = {}
    for i, year in enumerate(years):
        year_total = sum(area_data[area][i] for area in area_data.keys())
        for area in area_data.keys():
            if area not in area_percentages:
                area_percentages[area] = []
            if year_total > 0:
                percentage = (area_data[area][i] / year_total) * 100
            else:
                percentage = 0
            area_percentages[area].append(percentage)
    
    # Get counts for title
    num_regions = len(area_columns)
    
    # Get village and farmer counts from df_final
    num_villages = len(df_final['Village'].unique()) if 'Village' in df_final.columns else "N/A"
    num_farmers = len(df_final['Pixel_ID'].unique()) if 'Pixel_ID' in df_final.columns else "N/A"
    
    # Create the stacked bar chart
    fig, ax = plt.subplots(figsize=(12, 8))
    
    # Bottom values for stacking
    bottom = np.zeros(len(years))
    
    # Plot each area
    for area_name in area_colors_hex.keys():
        if area_name in area_percentages:
            values = area_percentages[area_name]
            color = area_colors_hex[area_name]
            
            ax.bar(years, values, bottom=bottom, label=area_name, 
                   color=color, alpha=0.8, edgecolor='white', linewidth=0.5)
            bottom += values
    
    # Customize the chart
    title = f"Area Pay Out Percentage by Year - {num_regions} Regions, {num_villages} Villages, {num_farmers} Farmers"
    ax.set_title(title, fontsize=16, fontweight='bold', pad=20)
    
    # Format y-axis to show percentages
    ax.set_ylim(0, 100)
    ax.yaxis.set_major_formatter(plt.FuncFormatter(lambda x, p: f'{x:.0f}%'))
    
    # Set x-axis ticks to show all years and rotate labels vertically with bigger font
    ax.set_xticks(years)
    ax.set_xticklabels(years, rotation=90, fontsize=12)
    
    # Make y-axis tick labels bigger
    ax.tick_params(axis='y', labelsize=12)
    
    # Add legend at bottom center with no box and square markers
    ax.legend(bbox_to_anchor=(0.5, -0.10), loc='upper center', fontsize=12, ncol=3, 
              frameon=False, handlelength=1, handletextpad=0.5, 
              markerscale=0.8, markerfirst=True)
    
    # Update legend markers to be squares
    legend = ax.get_legend()
    for handle in legend.legend_handles:
        handle.set_width(10)
        handle.set_height(10)
    
    # Add grid for better readability
    ax.grid(True, alpha=0.3, axis='y')
    
    # Tight layout to prevent label cutoff
    plt.tight_layout()
    
    # Save if path provided
    if save_path:
        plt.savefig(save_path, dpi=300, bbox_inches='tight')
        print(f"Chart saved to: {save_path}")
    
    plt.show()
    return fig

def create_diversification_benefit_chart(df_wide_numeric: pd.DataFrame, save_path: str = None):
    """
    Create a chart showing "Diversification Benefit: CoV Reduction due to Portfolio Effect"
    
    Parameters:
    - df_wide_numeric: The numeric regional statistics dataframe from build_regional_statistics
    - save_path: Optional path to save the chart (e.g., "chart.png")
    """
    
    # Look for CoV rows in the dataframe
    avg_pixel_cov_row = None
    area_cov_row = None
    
    for idx in df_wide_numeric.index:
        if 'average non-zero/blank pixel cov' in str(idx).lower():
            avg_pixel_cov_row = idx
        elif 'area cov' in str(idx).lower():
            area_cov_row = idx
    
    if avg_pixel_cov_row is None or area_cov_row is None:
        print("Required CoV data not found in the dataframe")
        print(f"Available rows: {list(df_wide_numeric.index)}")
        return
    
    # Extract area and regional CoV data
    areas = []
    avg_cov_regions = []  # Average CoV of pixels within each area
    area_cov = []         # CoV at area level (from area totals)
    
    # Extract data for each area
    for col in df_wide_numeric.columns:
        if str(col).endswith(" Total") and not str(col).startswith("Overall"):
            area_name = str(col).replace(" Total", "")
            areas.append(area_name)
            
            # Get average pixel CoV for this area
            avg_pixel_cov_val = df_wide_numeric.at[avg_pixel_cov_row, col]
            if pd.isna(avg_pixel_cov_val) or avg_pixel_cov_val is None:
                avg_pixel_cov_val = 0
            else:
                avg_pixel_cov_val = float(avg_pixel_cov_val)
            avg_cov_regions.append(avg_pixel_cov_val)
            
            # Get area-level CoV for this area
            area_cov_val = df_wide_numeric.at[area_cov_row, col]
            if pd.isna(area_cov_val) or area_cov_val is None:
                area_cov_val = 0
            else:
                area_cov_val = float(area_cov_val)
            area_cov.append(area_cov_val)
    
    # Get national CoV from overall total column
    if "Overall Total" in df_wide_numeric.columns:
        national_cov_val = df_wide_numeric.at[area_cov_row, "Overall Total"]
        if pd.isna(national_cov_val) or national_cov_val is None:
            national_cov = 0
        else:
            national_cov = float(national_cov_val)
    else:
        # If no overall column, calculate weighted average of area CoVs
        national_cov = np.mean(area_cov) if area_cov else 0
    
    # Calculate diversification benefit (reduction from regional average to national)
    diversification_benefit = []
    for i in range(len(areas)):
        if avg_cov_regions[i] != 0:
            benefit = (1 - (national_cov / avg_cov_regions[i])) * 100
        else:
            benefit = 0
        diversification_benefit.append(benefit)
    
    # Create the chart
    fig, ax1 = plt.subplots(figsize=(12, 8))
    
    # Set up the bar positions
    x = np.arange(len(areas))
    width = 0.25
    
    # Create bars
    bars1 = ax1.bar(x - width, avg_cov_regions, width, label='Average CoV of pixels within Area', 
                    color='#5B9BD5', alpha=0.8)
    bars2 = ax1.bar(x, area_cov, width, label='Area CoV', 
                    color='#FF7F0E', alpha=0.8)
    bars3 = ax1.bar(x + width, [national_cov] * len(areas), width, label='National CoV', 
                    color='#A5A5A5', alpha=0.8)
    
    # Create second y-axis for diversification benefit
    ax2 = ax1.twinx()
    
    # Add diversification benefit as yellow dots
    ax2.scatter(x, diversification_benefit, color='#FFD700', s=300, zorder=5,
               label='Diversification Benefit: National CoV Compared to Regional Average')
    
    # Customize primary y-axis (left)
    ax1.set_xlabel('')
    ax1.set_ylabel('Coefficient of Variation', fontsize=12, fontweight='bold')
    ax1.set_title('Diversification Benefit : CoV Reduction due to Portfolio Effect', 
                 fontsize=16, fontweight='bold', pad=20)
    ax1.set_xticks(x)
    ax1.set_xticklabels(areas, fontsize=12)
    ax1.tick_params(axis='y', labelsize=12)
    ax1.set_ylim(0, 6)
    ax1.grid(True, alpha=0.3, axis='y')
    
    # Customize secondary y-axis (right)
    ax2.set_ylabel('Reduction of Regional Average CoV due to Diversification Effects', 
                  fontsize=10, fontweight='bold')
    ax2.tick_params(axis='y', labelsize=10)
    ax2.set_ylim(0, 100)
    
    # Invert right y-axis to show percentages decreasing
    ax2.invert_yaxis()
    
    # Format right y-axis as percentages
    ax2.yaxis.set_major_formatter(plt.FuncFormatter(lambda x, p: f'{x:.0f}%'))
    
    # Combine legends
    lines1, labels1 = ax1.get_legend_handles_labels()
    lines2, labels2 = ax2.get_legend_handles_labels()
    ax1.legend(lines1 + lines2, labels1 + labels2, 
              bbox_to_anchor=(0.5, -0.10), loc='upper center', fontsize=10, ncol=2,
              frameon=False)
    
    plt.tight_layout()
    
    # Save if path provided
    if save_path:
        plt.savefig(save_path, dpi=300, bbox_inches='tight')
        print(f"Chart saved to: {save_path}")
    
    plt.show()
    return fig
# Usage example:
# df_regional, df_regional_fmt = build_regional_statistics(df_final)
# fig = create_area_payout_chart(df_regional, "area_payout_chart.png")
# fig = create_area_payout_percentage_chart(df_regional, df_final, "area_payout_percentage_chart.png")
# fig = create_diversification_benefit_chart(df_regional, "diversification_benefit_chart.png")