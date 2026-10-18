import os
import sys
import numpy as np
import pandas as pd
from scipy.spatial import Voronoi
from shapely.geometry import Polygon, box

# Logger class to write output simultaneously to console and text file


class TeeLogger(object):
    def __init__(self, filename="voronoi_equilibrium_output.txt"):
        self.terminal = sys.stdout
        self.log = open(filename, "w", encoding="utf-8")

    def write(self, message):
        self.terminal.write(message)
        self.log.write(message)
        self.log.flush()

    def flush(self):
        self.terminal.flush()
        self.log.flush()


sys.stdout = TeeLogger("voronoi_equilibrium_output.txt")

EXCEL_SOURCE = "Exported-file-W1.xlsx"


def run_excel_bounded_voronoi_integration():
    print("Executing complete boundary-aware Voronoi integration for CSI SAFE Excel workbook...")

    # 1. Extract Soil Pressures & Mesh Coordinates from Sheet
    df_soil_raw = pd.read_excel(
        EXCEL_SOURCE, sheet_name='Soil Pressures', header=1).iloc[1:].reset_index(drop=True)

    # Clean and convert numeric fields
    df_soil_raw['Soil_Pressure_kPa'] = pd.to_numeric(
        df_soil_raw['Soil Pressure'], errors='coerce')
    df_soil_raw['X_m'] = (pd.to_numeric(
        df_soil_raw['Global X'], errors='coerce') / 1000.0).round(4)
    df_soil_raw['Y_m'] = (pd.to_numeric(
        df_soil_raw['Global Y'], errors='coerce') / 1000.0).round(4)

    df_clean = df_soil_raw.dropna(
        subset=['Soil_Pressure_kPa', 'X_m', 'Y_m']).copy()
    df_clean['Stress_abs'] = df_clean['Soil_Pressure_kPa'].abs()

    # Mat Footprint (32m x 22m)
    x_min, x_max = -1.0, 31.0
    y_min, y_max = -1.0, 21.0
    mat_boundary = box(x_min, y_min, x_max, y_max)

    # 2. Resolve Coincident Nodes by grouping exact coordinates
    node_df = df_clean.groupby(['X_m', 'Y_m'])[
        'Stress_abs'].max().reset_index()

    points = node_df[['X_m', 'Y_m']].values
    stresses = node_df['Stress_abs'].values

    # 3. Add Outer Boundary Ghost Frame to Force Bounded Voronoi Cells on Edge Nodes
    buffer_dist = 10.0
    ghost_frame = np.array([
        [x_min - buffer_dist, y_min - buffer_dist],
        [x_max + buffer_dist, y_min - buffer_dist],
        [x_max + buffer_dist, y_max + buffer_dist],
        [x_min - buffer_dist, y_max + buffer_dist]
    ])

    all_points = np.vstack([points, ghost_frame])

    # Run Voronoi diagram on node coordinates including ghost boundary
    vor = Voronoi(all_points)

    integrated_force = 0.0
    total_computed_area = 0.0

    # 4. Integrate Area * Stress over real mat nodes
    for i in range(len(points)):
        region_idx = vor.point_region[i]
        region = vor.regions[region_idx]

        if len(region) > 0 and -1 not in region:
            polygon_vertices = [vor.vertices[v] for v in region]
            poly = Polygon(polygon_vertices)

            # Clip cell strictly to the mat footprint
            clipped_poly = poly.intersection(mat_boundary)

            if not clipped_poly.is_empty:
                area = clipped_poly.area
                integrated_force += stresses[i] * area
                total_computed_area += area

    # 5. Verification Output
    target_load = 157715.52
    print("\n================ BOUNDED VORONOI EQUILIBRIUM ================")
    print(f"Matched Mat Nodes       : {len(points)}")
    print(
        f"Integrated Surface Area : {total_computed_area:.2f} m² (Target: 704.00 m²)")
    print(f"Exact Integrated Force  : {integrated_force:.2f} kN")
    print(f"Applied Load Target     : {target_load:.2f} kN")
    print(
        f"Equilibrium Compliance  : {(integrated_force / target_load) * 100:.2f}%")
    print("==============================================================")


if __name__ == "__main__":
    run_excel_bounded_voronoi_integration()
