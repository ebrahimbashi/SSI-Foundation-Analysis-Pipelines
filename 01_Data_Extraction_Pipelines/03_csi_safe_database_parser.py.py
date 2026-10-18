import os
import pandas as pd
import numpy as np
from scipy.interpolate import griddata

# Source Combined Workbook Export
EXCEL_SOURCE = "Exported-file.xlsx"
OUTPUT_EXCEL = "SAFE_Native_Export_Cutlines.xlsx"

# Standard Verification Cut-Line Positions
Y_STRIPS = [0.0, 5.0, 10.0, 15.0, 20.0]
X_STRIPS = [0.0, 5.0, 10.0, 15.0, 20.0, 25.0, 30.0]
TOL = 0.15  # 15 cm spatial tolerance for mesh node grouping


def run_combined_workbook_extraction():
    print("==================================================================")
    print(" CSI SAFE COMBINED WORKBOOK AUTOMATION PIPELINE                  ")
    print("==================================================================")

    xls = pd.ExcelFile(EXCEL_SOURCE)
    print(f"[1/5] Loaded source workbook. Sheets available: {xls.sheet_names}")

    # 1. Read Soil Pressure Data from Sheet
    print("[2/5] Extracting Soil Pressures & Mesh Coordinates...")
    df_soil_raw = pd.read_excel(
        EXCEL_SOURCE, sheet_name='Soil Pressures', header=1).iloc[1:].reset_index(drop=True)
    df_soil_raw['Joint_clean'] = df_soil_raw['Joint'].astype(
        str).str.strip().str.replace('.0', '', regex=False)
    df_soil_raw['Soil_Pressure_kPa'] = pd.to_numeric(
        df_soil_raw['Soil Pressure'], errors='coerce')
    df_soil_raw['X_m'] = (pd.to_numeric(
        df_soil_raw['Global X'], errors='coerce') / 1000.0).round(4)
    df_soil_raw['Y_m'] = (pd.to_numeric(
        df_soil_raw['Global Y'], errors='coerce') / 1000.0).round(4)

    df_soil_grouped = df_soil_raw.groupby(
        ['Joint_clean', 'X_m', 'Y_m'], as_index=False)['Soil_Pressure_kPa'].mean()

    # 2. Read Element Bending Forces (M11, M22) from Sheet
    print("[3/5] Extracting Area Shell Bending Moments...")
    df_forces_raw = pd.read_excel(
        EXCEL_SOURCE, sheet_name='Element Forces - Area Shells', header=1).iloc[1:].reset_index(drop=True)
    df_forces_raw['Joint_clean'] = df_forces_raw['Joint'].astype(
        str).str.strip().str.replace('.0', '', regex=False)
    df_forces_raw['M11'] = pd.to_numeric(df_forces_raw['M11'], errors='coerce')
    df_forces_raw['M22'] = pd.to_numeric(df_forces_raw['M22'], errors='coerce')

    df_forces_grouped = df_forces_raw.groupby('Joint_clean', as_index=False).agg({
        'M11': 'mean',
        'M22': 'mean'
    }).rename(columns={'M11': 'M11_kNm_m', 'M22': 'M22_kNm_m'})

    # Merge Soil & Forces
    df_mesh = pd.merge(df_soil_grouped, df_forces_grouped,
                       on='Joint_clean', how='left')

    # 3. Read Displacement & Point Connectivity from Sheets
    print("[4/5] Extracting Joint Displacements and Mapping Spatial Surface...")
    df_disp_raw = pd.read_excel(
        EXCEL_SOURCE, sheet_name='Joint Displacements', header=1).iloc[1:].reset_index(drop=True)
    df_disp_raw['Joint_clean'] = df_disp_raw['Unique Name'].astype(
        str).str.strip().str.replace('.0', '', regex=False)
    df_disp_raw['Uz_mm'] = pd.to_numeric(df_disp_raw['Uz'], errors='coerce')

    df_conn_raw = pd.read_excel(
        EXCEL_SOURCE, sheet_name='Point Object Connectivity', header=1).iloc[1:].reset_index(drop=True)
    df_conn_raw['Joint_clean'] = df_conn_raw['UniqueName'].astype(
        str).str.strip().str.replace('.0', '', regex=False)
    df_conn_raw['X_m'] = pd.to_numeric(
        df_conn_raw['X'], errors='coerce').round(4)
    df_conn_raw['Y_m'] = pd.to_numeric(
        df_conn_raw['Y'], errors='coerce').round(4)

    df_disp_coords = pd.merge(
        df_disp_raw[['Joint_clean', 'Uz_mm']],
        df_conn_raw[['Joint_clean', 'X_m', 'Y_m']],
        on='Joint_clean',
        how='inner'
    )

    # Map Spatial Displacements to FE Mesh Nodes via Interpolation
    pts_known = df_disp_coords[['X_m', 'Y_m']].values
    val_uz = df_disp_coords['Uz_mm'].values
    pts_target = df_mesh[['X_m', 'Y_m']].values

    uz_cubic = griddata(pts_known, val_uz, pts_target, method='cubic')
    uz_linear = griddata(pts_known, val_uz, pts_target, method='linear')
    df_mesh['Uz_mm'] = np.where(np.isnan(uz_cubic), uz_linear, uz_cubic)

    df_final = df_mesh[['Joint_clean', 'X_m', 'Y_m', 'Uz_mm',
                        'Soil_Pressure_kPa', 'M11_kNm_m', 'M22_kNm_m']]

    # 4. Output Multi-Tab Excel Workbook with Cut-Lines
    print("[5/5] Generating Master Excel Cut-Line Workbook...")
    output_path = os.path.abspath(OUTPUT_EXCEL)
    with pd.ExcelWriter(output_path, engine='openpyxl') as writer:
        df_final.sort_values(by=['Y_m', 'X_m']).to_excel(
            writer, sheet_name='Full_Mesh_Data', index=False)

        for y_val in Y_STRIPS:
            sub = df_final[(df_final['Y_m'] >= y_val - TOL) &
                           (df_final['Y_m'] <= y_val + TOL)].sort_values(by='X_m')
            sub.to_excel(writer, sheet_name=f"Cut_Y={y_val:g}m", index=False)

        for x_val in X_STRIPS:
            sub = df_final[(df_final['X_m'] >= x_val - TOL) &
                           (df_final['X_m'] <= x_val + TOL)].sort_values(by='Y_m')
            sub.to_excel(writer, sheet_name=f"Cut_X={x_val:g}m", index=False)

    print(f"\n[SUCCESS] Master workbook generated successfully: {output_path}")


if __name__ == "__main__":
    run_combined_workbook_extraction()
