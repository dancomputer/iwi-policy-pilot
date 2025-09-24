from pathlib import Path
import random
import re
from typing import Optional
import pandas as pd

from build_excel_sheet1 import build_modelled_yields_sheet

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
        df_final["Loan_Amount"] = df_final.groupby("Region").apply(lambda x: random.randint(1000, 1500))
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

        return df_final

if __name__ == "__main__":
        df_final = load_and_merge()
        print("Merged dataframe shape:", df_final.shape)
        print(df_final.head())
        # Optionally save to disk:
        # df_final.to_csv("merged_pixel_timeseries_long.csv", index=False)
        wb = build_modelled_yields_sheet(df_final)
        wb.save("policy_pilot_output3.xlsx")