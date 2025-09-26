from pathlib import Path
import random
import re
from typing import Optional
import pandas as pd
 
# file: merge_pixel_timeseries.py
"""
Merge three data sources into a single pandas DataFrame:
- village_pixel_matches_maize-nkasi.csv : has Region, District, Latitude, Longitude, Pixel
    (adds FarmerID sequential index)
- ThreeVariableContiguous-SyntheticYield-Conservative-metadata.csv : has pixel, Threshold_Yield
- ThreeVariableContiguous-SyntheticYield-timeseries-metadata.csv : columns like "pixel 0", "pixel 1", ...
    rows correspond to years 1981-2022 (or include a Year column)

Output: df_final (pandas DataFrame) with metadata + long timeseries (Year, Yield) per Pixel.
"""


# --- Configure input paths (adjust if needed) ---
PATH_VILL = Path(r"C:\Users\danie\NecessaryM1InternshipCode\ProjectRice\PolicyPilot\iwi-policy-pilot\data\village_pixel_matches_maize-nkasi.csv")
PATH_THRESH = Path(r"C:\Users\danie\NecessaryM1InternshipCode\ProjectRice\OutputCalendarDays180_Maize_1981_2022_SPARSE\ThreeVariableContiguous-SyntheticYield-Optimistic-metadata.csv")
PATH_TS = Path(r"C:\Users\danie\NecessaryM1InternshipCode\ProjectRice\OutputCalendarDays180_Maize_1981_2022_SPARSE\ThreeVariableContiguous-SyntheticYield-Conservative_timeseries.csv")

def _ensure_pixel_col(df: pd.DataFrame) -> pd.DataFrame:
        # Normalize pixel column name to 'Pixel' if possible
        cols = {c: c for c in df.columns}
        for c in df.columns:
                if re.fullmatch(r'(?i)pixel$', c.strip()):
                        cols[c] = 'Pixel'
                        break
        df = df.rename(columns=cols)
        return df

def _find_year_column(df: pd.DataFrame) -> Optional[str]:
        # Detect a column that contains year values between 1981 and 2022
        for c in df.columns:
                try:
                        vals = pd.to_numeric(df[c], errors='coerce').dropna().astype(int)
                        pct_in_range = ((vals >= 1981) & (vals <= 2022)).mean()
                        if pct_in_range > 0.8:
                                return c
                except Exception:
                        continue
        return None

def load_and_merge() -> pd.DataFrame:
        # Read village / pixel matches
        df_vill = pd.read_csv(PATH_VILL, low_memory=False)
        df_vill = _ensure_pixel_col(df_vill)
        if 'Pixel' not in df_vill.columns:
                # try lowercase
                if 'pixel' in df_vill.columns:
                        df_vill = df_vill.rename(columns={'pixel': 'Pixel'})
        # Add FarmerID as simple sequential index starting at 1
        df_vill = df_vill.reset_index(drop=True)
        df_vill['FarmerID'] = range(1, len(df_vill) + 1)

        # Read threshold metadata and normalize pixel column name
        df_thresh = pd.read_csv(PATH_THRESH, low_memory=False)
        df_thresh = _ensure_pixel_col(df_thresh)
        if 'Pixel' not in df_thresh.columns and 'pixel' in df_thresh.columns:
                df_thresh = df_thresh.rename(columns={'pixel': 'Pixel'})

        # Convert Pixel types to numeric where possible
        for df in (df_vill, df_thresh):
                if 'Pixel' in df.columns:
                        df['Pixel'] = pd.to_numeric(df['Pixel'], errors='coerce').astype('Int64')

        # Merge village with threshold metadata on Pixel
        df_meta = pd.merge(df_vill, df_thresh, on='Pixel', how='left', suffixes=('', '_thresh'))

        # Read timeseries metadata
        df_ts = pd.read_csv(PATH_TS, low_memory=False)

        # Detect year column (if present). If none, assume rows correspond to 1981..2022 in order.
        year_col = _find_year_column(df_ts)
        if year_col is None:
                # create Year column from 1981..2022 assuming rows are in order
                nrows = len(df_ts)
                start_year = 1981
                years = list(range(start_year, start_year + nrows))
                df_ts = df_ts.copy()
                df_ts.insert(0, 'Year', years)
                year_col = 'Year'
        else:
                # ensure year column is integer
                df_ts[year_col] = pd.to_numeric(df_ts[year_col], errors='coerce').astype('Int64')

        # Identify pixel columns: those whose name contains digits and the word 'pixel' (case-insensitive)
        pixel_cols = []
        for c in df_ts.columns:
                if c == year_col:
                        continue
                if re.search(r'(?i)pixel', c) or re.search(r'^\d+$', c.strip()):
                        pixel_cols.append(c)
        # Fallback: if no pixel-like columns found, assume all non-year cols are pixel columns
        if not pixel_cols:
                pixel_cols = [c for c in df_ts.columns if c != year_col]

        # Melt to long format
        df_long = df_ts.melt(id_vars=[year_col], value_vars=pixel_cols,
                                                 var_name='pixel_col', value_name='Yield')
        df_long = df_long.rename(columns={year_col: 'Year'})

        # Extract numeric Pixel id from pixel_col names
        def extract_pixel_id(s: str) -> Optional[int]:
                if pd.isna(s):
                        return None
                # try to find an integer in the column name
                m = re.search(r'(\d+)', str(s))
                if m:
                        return int(m.group(1))
                # if the whole column name is numeric
                try:
                        return int(str(s).strip())
                except Exception:
                        return None

        df_long['Pixel'] = df_long['pixel_col'].apply(extract_pixel_id).astype('Int64')

        # Convert Yield to numeric
        df_long['Yield'] = pd.to_numeric(df_long['Yield'], errors='coerce')

        # Merge long timeseries with metadata on Pixel
        df_final = pd.merge(df_meta, df_long.drop(columns=['pixel_col']), on='Pixel', how='left')

        # Optional: reorder columns (Pixel, Year, Yield, Threshold_Yield, metadata...)
        cols_front = ['Pixel', 'Year', 'Yield']
        if 'Threshold_Yield' in df_final.columns:
                cols_front.append('Threshold_Yield')
        remaining = [c for c in df_final.columns if c not in cols_front]
        df_final = df_final[cols_front + remaining]
        
        #ADD additional values
        df_final["Yield_Abs"] = df_final["Yield"]*df_final["Threshold_Yield"]
        #1. Define "Attach" and "Detach" columns as 95th and 5th percentiles of absolute yield by pixel 
        df_final["Attach"] = df_final.groupby("Pixel")["Yield_Abs"].transform(lambda x: x.quantile(0.5))
        df_final["Detach"] = df_final.groupby("Pixel")["Yield_Abs"].transform(lambda x: x.quantile(0.15))
        #2. Define Loan Amount as a random amount between 1000 and 1500. Each region has same loan amount.
        # dictionary mapping each unique region to a random number
        region_map = {region: random.randint(1000, 1500) for region in df_final["Region"].unique()}
        # Map the values from the dictionary to the 'Region' column
        df_final["Loan_Amount"] = df_final["Region"].map(region_map)
        print("Assigned random Loan Amount between 1000 and 1500 per Region")
        #3. Rename FarmerID as "Pixel ID"
        df_final = df_final.rename(columns={"FarmerID": "Pixel_ID"})
        #4. rename Pixel as Index_ID
        df_final = df_final.rename(columns={"Pixel": "Index_ID"})

        #5. Add Area: Dictionary based on mapping Region to North/South etc.
        tanzania_zones = {
            "Northern Zone": [
                "Arusha",
                "Kilimanjaro",
                "Manyara",
                "Tanga"
            ],
            "Central Zone": [
                "Dodoma",
                "Singida",
                "Tabora"
            ],
            "Lake Zone": [
                "Geita",
                "Kagera",
                "Mara",
                "Mwanza",
                "Shinyanga",
                "Simiyu"
            ],
            "Western Zone": [
                "Kigoma",
                "Katavi"
            ],
            "Southern Highlands Zone": [
                "Iringa",
                "Mbeya",
                "Njombe",
                "Rukwa",
                "Ruvuma",
                "Songwe"
            ],
            "Coastal Zone": [
                "Dar es Salaam",
                "Lindi",
                "Morogoro",
                "Mtwara",
                "Pwani"  # 'Pwani' means 'Coast' in Swahili
            ],
            "Zanzibar (Islands)": [
                "Pemba North",
                "Pemba South",
                "Unguja North",  # Also known as Zanzibar North
                "Unguja South",  # Also known as Zanzibar South & Central
                "Mjini Magharibi" # Also known as Zanzibar Urban West
            ]
        }
        def map_region_to_area(region: str) -> str:
            # Map a region name to its corresponding area using the tanzania_zones dictionary
            # Do not be case sensitive, use fuzzy matching too
            region = region.lower()
            for area, regions in tanzania_zones.items():
                for reg in regions:
                    if region == reg.lower() or region in reg.lower() or reg.lower() in region:
                        return area
            return "Unknown"
        df_final["Area"] = df_final["Region"].apply(map_region_to_area)

        #II. Do some basic processing
        #1. Payout Base fraction: 0 if Yield_Abs > Attach, 1 if Yield_Abs < Detach, linear in between
        df_final["PayoutsPercent"] = df_final.apply(lambda x: 0 if x["Yield_Abs"] > x["Attach"] else 1 if x["Yield_Abs"] < x["Detach"] else (x["Attach"] - x["Yield_Abs"]) / (x["Attach"] - x["Detach"]), axis=1)
        #2. Payout amount base:
        df_final["Sum_Insured"] = df_final["Loan_Amount"] * 0.4 
        print("using Sum_Insured as 40% of Loan Amount")
        df_final["PayoutAmountBase"] = df_final["PayoutsPercent"] * df_final["Sum_Insured"]
        #3. Payout stats: Average, SD, Coefficient of Variation (CoV), Min, Max, 90th percentile, 95th percentile per Pixel
        # compute per-pixel (Pixel_ID) statistics for PayoutAmountBase and attach them back to df_final
        stat_aggs = {
            'PayoutAvg': 'mean',
            'PayoutSD': 'std',
            'PayoutMin': 'min',
            'PayoutMax': 'max',
            'Payout90': lambda x: x.quantile(0.90),
            'Payout95': lambda x: x.quantile(0.95),
        }
        stats = df_final.groupby('Pixel_ID')['PayoutAmountBase'].agg(**stat_aggs)

        # Coefficient of variation: SD / mean (guard against division by zero)
        stats['PayoutCoV'] = stats['PayoutSD'] / stats['PayoutAvg'].replace({0: pd.NA})
        stats = stats.fillna(0)

        # Merge the statistics back into the main dataframe (one row per original row, stats repeated per Pixel_ID)
        df_final = df_final.merge(stats, how='left', left_on='Pixel_ID', right_index=True)

        print("Computed per-pixel payout statistics (avg, sd, cov, min, max, 90th, 95th).")


        return df_final

####################################################
def build_regional_statistics(df_final: pd.DataFrame, verbose: bool = False):
    """
    Build wide dataframe of regional + area + overall statistics for sheet 5.
    Adds region name normalization so region columns appear (was only Overall Total).
    Set verbose=True to print debugging info.
    """
    # --- Normalize / canonicalize region names ---
    def normalize_region_name(s):
        if pd.isna(s):
            return s
        s = str(s).strip()
        # remove word 'region'
        s = re.sub(r'\bregion\b', '', s, flags=re.IGNORECASE)
        # collapse spaces
        s = re.sub(r'\s+', ' ', s)
        s = s.title()
        corrections = {
            'Rukva': 'Rukwa',   # common typo seen
            'Mbeya Rural': 'Mbeya',
            'Kilimanjaro Region': 'Kilimanjaro',
        }
        return corrections.get(s, s)

    df_final = df_final.copy()
    if 'Region' not in df_final.columns:
        raise ValueError("df_final missing 'Region' column")
    df_final['Region'] = df_final['Region'].astype(str).apply(normalize_region_name)

    if verbose:
        print("Normalized Region list:", sorted(df_final['Region'].dropna().unique()))

    REGION_GROUPS = {
        "Northern Zone": ["Arusha", "Kilimanjaro", "Manyara", "Tanga"],
        "Central Zone": ["Dodoma", "Singida", "Tabora"],
        "Lake Zone": ["Geita", "Kagera", "Mara", "Mwanza", "Shinyanga", "Simiyu"],
        "Western Zone": ["Kigoma", "Katavi"],
        "Southern Highlands Zone": ["Iringa", "Mbeya", "Njombe", "Rukwa", "Ruvuma", "Songwe"],
        "Coastal Zone": ["Dar Es Salaam", "Lindi", "Morogoro", "Mtwara", "Pwani"],
        "Zanzibar (Islands)": ["Pemba North", "Pemba South", "Unguja North",
                               "Unguja South", "Mjini Magharibi"]
    }

    # Build display columns (case-insensitive match)
    region_available_lower = {r.lower(): r for r in df_final['Region'].dropna().unique()}
    display_columns = []
    column_meta = []  # (label, type, area, member_regions)

    for area, regions in REGION_GROUPS.items():
        existing = []
        for r in regions:
            key = r.lower()
            if key in region_available_lower:
                # use canonical group name (r) as column label
                existing.append(r)
        if verbose:
            print(f"[{area}] matched regions: {existing}")
        for r in existing:
            display_columns.append(r)
            column_meta.append((r, "region", area, [r]))
        if existing:
            gname = f"{area} Total"
            display_columns.append(gname)
            column_meta.append((gname, "group_total", area, existing))

    overall_name = "Overall Total"
    display_columns.append(overall_name)
    all_regions_in_data = sorted(df_final['Region'].dropna().unique().tolist())
    column_meta.append((overall_name, "overall_total", None, all_regions_in_data))

    # Required pixel-level columns
    pixel_cols_needed = [
        'Pixel_ID', 'Region', 'Loan_Amount', 'Sum_Insured',
        'PayoutAvg', 'PayoutSD', 'PayoutMin', 'PayoutMax', 'Payout90', 'Payout95', 'PayoutCoV'
    ]
    for c in pixel_cols_needed:
        if c not in df_final.columns:
            raise ValueError(f"Required column '{c}' not found in df_final")

    df_pixels = df_final[pixel_cols_needed].drop_duplicates(subset='Pixel_ID')
    df_payouts = df_final[['Pixel_ID', 'Region', 'Year', 'PayoutAmountBase']].copy()

    # Pixel status
    pixel_groups = df_payouts.groupby('Pixel_ID')
    pixel_is_blank = pixel_groups['PayoutAmountBase'].apply(lambda s: s.isna().all())
    pixel_is_zero = pixel_groups['PayoutAmountBase'].apply(
        lambda s: (not s.isna().all()) and (s.fillna(0).sum() == 0)
    )
    pixel_region_map = df_pixels.set_index('Pixel_ID')['Region']
    pixel_status = pd.DataFrame({
        'Region': pixel_region_map,
        'is_blank': pixel_is_blank,
        'is_zero': pixel_is_zero
    })

    annual_region_sums = {
        region: grp.groupby('Year')['PayoutAmountBase'].sum(min_count=1)
        for region, grp in df_payouts.groupby('Region')
    }
    years = sorted(df_final['Year'].dropna().unique().tolist())

    def compute_loans(regions):
        return df_pixels[df_pixels['Region'].isin(regions)]['Loan_Amount'].sum()

    def compute_sum_insured(regions):
        return df_pixels[df_pixels['Region'].isin(regions)]['Sum_Insured'].sum()

    def compute_year_value(regions, year):
        total = 0.0
        any_data = False
        for r in regions:
            series = annual_region_sums.get(r)
            if series is not None and year in series.index:
                val = series.loc[year]
                if pd.notna(val):
                    total += val
                    any_data = True
        return total if any_data else None

    def compute_pixel_counts(regions):
        pix = df_pixels[df_pixels['Region'].isin(regions)]['Pixel_ID'].unique()
        if len(pix) == 0:
            return 0, 0, 0, 0
        sub_status = pixel_status.loc[pix]
        n_blank = int(sub_status['is_blank'].sum())
        n_zero = int(sub_status['is_zero'].sum())
        n_zero_blank = n_blank + n_zero
        return len(pix), n_zero_blank, n_blank, n_zero

    def compute_avg_cov_non_zero(regions):
        subset = df_pixels[df_pixels['Region'].isin(regions)]
        if subset.empty:
            return None
        statuses = pixel_status.loc[subset['Pixel_ID']]
        valid_ids = statuses.index[~(statuses['is_blank'] | statuses['is_zero'])]
        if len(valid_ids) == 0:
            return None
        return df_pixels[df_pixels['Pixel_ID'].isin(valid_ids)]['PayoutCoV'].mean()

    # REPLACED: pixel-level distribution => area-level annual totals distribution
    def compute_area_level_distribution(regions):
        """
        Compute distribution of annual total payouts for the given regions:
        - For each year, sum payouts across all pixels in the regions.
        - Return avg, sd, min, max, p90, p95 over those annual totals.
        Include years with zero totals; skip years with no data (None).
        """
        totals = []
        for year in years:
            v = compute_year_value(regions, year)
            if v is not None:
                totals.append(float(v))
        if not totals:
            return dict(avg=None, sd=None, min=None, max=None, p90=None, p95=None)

        n = len(totals)
        avg = sum(totals) / n
        if n > 1:
            mean = avg
            variance = sum((x - mean) ** 2 for x in totals) / (n - 1)
            sd = variance ** 0.5
        else:
            sd = 0.0

        s = pd.Series(totals)
        return dict(
            avg=avg,
            sd=sd,
            min=min(totals),
            max=max(totals),
            p90=s.quantile(0.90),
            p95=s.quantile(0.95)
        )

    def compute_area_cov(regions):
        """Calculate CoV of area-level annual payout totals as SD/Average using the same totals as distribution."""
        dist = compute_area_level_distribution(regions)
        avg, sd = dist.get('avg'), dist.get('sd')
        if avg is None or sd is None or avg == 0:
            return None
        return sd / avg

    base_rows = (["Loan amounts (USD)", "Sum insured", "Area", "Region"] +
                 [str(y) for y in years] +
                 ["Average Payout", "SD", "Min", "Max",
                  "90th percentile", "95th percentile",
                  "Number of Pixels", "Number of Zero and Blank Pixels",
                  "Number of Blank Pixels", "Number of Zero Pixel",
                  "Average non-zero/blank pixel CoV", "Area CoV"])

    rows = {}
    def init_row(label):
        rows[label] = {col: None for col in display_columns}

    for r in base_rows:
        init_row(r)

    for (col_label, col_type, area, member_regions) in column_meta:
        if col_type == "region":
            rows["Area"][col_label] = area
            rows["Region"][col_label] = col_label
        elif col_type == "group_total":
            rows["Area"][col_label] = area
            rows["Region"][col_label] = "Total"
        else:
            rows["Area"][col_label] = "Overall Total"
            rows["Region"][col_label] = "Overall Total"

    for (col_label, col_type, area, member_regions) in column_meta:
        rows["Loan amounts (USD)"][col_label] = compute_loans(member_regions)
        rows["Sum insured"][col_label] = compute_sum_insured(member_regions)
        for y in years:
            rows[str(y)][col_label] = compute_year_value(member_regions, y)
        # UPDATED: use area-level distribution instead of pixel-level
        dist = compute_area_level_distribution(member_regions)
        rows["Average Payout"][col_label] = dist['avg']
        rows["SD"][col_label] = dist['sd']
        rows["Min"][col_label] = dist['min']
        rows["Max"][col_label] = dist['max']
        rows["90th percentile"][col_label] = dist['p90']
        rows["95th percentile"][col_label] = dist['p95']
        n_pixels, n_zero_blank, n_blank, n_zero = compute_pixel_counts(member_regions)
        rows["Number of Pixels"][col_label] = n_pixels
        rows["Number of Zero and Blank Pixels"][col_label] = n_zero_blank or None
        rows["Number of Blank Pixels"][col_label] = n_blank or None
        rows["Number of Zero Pixel"][col_label] = n_zero or None
        rows["Average non-zero/blank pixel CoV"][col_label] = compute_avg_cov_non_zero(member_regions)
        rows["Area CoV"][col_label] = compute_area_cov(member_regions)

    order_of_rows = [
        "Loan amounts (USD)",
        "Sum insured",
        "Area",
        "Region",
        *[str(y) for y in years],
        "Average Payout",
        "SD",
        "Min",
        "Max",
        "90th percentile",
        "95th percentile",
        "Number of Pixels",
        "Number of Zero and Blank Pixels",
        "Number of Blank Pixels",
        "Number of Zero Pixel",
        "Average non-zero/blank pixel CoV",
        "Area CoV"
    ]
    import pandas as _pd
    df_wide_numeric = _pd.DataFrame({r: rows[r] for r in order_of_rows}).T

    def fmt_int(x):
        if x is None or (_pd.isna(x)):
            return "-"
        try:
            return f"{int(round(float(x))):,}".replace(",", " ")
        except Exception:
            return "-"

    def fmt_float(x, decimals=2):
        if x is None or _pd.isna(x):
            return "-"
        return f"{float(x):.{decimals}f}"

    df_wide_formatted = df_wide_numeric.copy()
    for rlab in df_wide_formatted.index:
        for col in df_wide_formatted.columns:
            val = df_wide_formatted.at[rlab, col]
            if rlab in ("Area", "Region"):
                df_wide_formatted.at[rlab, col] = "-" if val is None else val
            elif rlab in ("Average Payout", "SD", "Min", "Max",
                          "90th percentile", "95th percentile",
                          "Average non-zero/blank pixel CoV", "Area CoV"):
                df_wide_formatted.at[rlab, col] = fmt_float(val, 2)
            elif rlab in ("Loan amounts (USD)", "Sum insured") or rlab.isdigit():
                df_wide_formatted.at[rlab, col] = fmt_int(val)
            else:
                df_wide_formatted.at[rlab, col] = fmt_int(val)

    if verbose:
        print("Final columns:", df_wide_formatted.columns.tolist())

    return df_wide_numeric, df_wide_formatted
if __name__ == "__main__":
        df_final = load_and_merge()
        print("Merged dataframe shape:", df_final.shape)
        #print(df_final.head())
        df_regional, df_regional_fmt = build_regional_statistics(df_final)
        print("Regional statistics dataframe shape:", df_regional.shape)
        # Optionally save to disk:
        # df_final.to_csv("merged_pixel_timeseries_long.csv", index=False)
        #wb = build_modelled_yields_sheet(df_final)
        #wb.save("policy_pilot_output3.xlsx")