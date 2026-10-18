import os
import pandas as pd
import numpy as np
from scipy.interpolate import griddata, Akima1DInterpolator, interp1d
from scipy.spatial import cKDTree
from plxscripting.easy import *

# ==============================================================================
# 1. USER CONFIGURATION (ACTIVE OPEN MODEL)
# ==============================================================================
PORT = 10001
PASSWORD = "GvYr8!m#>5u?Lt!W"

PHASE_NAME = "Phase_3"  # Exact or partial phase name in open model

# Mat Bounding Box (m)
X_MIN, X_MAX = -1.0, 31.0
Y_MIN, Y_MAX = -1.0, 21.0
Z_BASE = -3.0          # Bottom elevation of mat (m)
THICKNESS = 1.2        # Mat thickness (m)

GRID_STEP = 0.5        # Resampling resolution (m)

# Axis Cutlines for Excel Sheets
Y_CUTLINES = [0.0, 5.0, 10.0, 15.0, 20.0]
X_CUTLINES = [0.0, 5.0, 10.0, 15.0, 20.0, 25.0, 30.0]

OUTPUT_EXCEL = "PLAXIS_Active_Model_Akima_Results.xlsx"

# ==============================================================================
# 2. CONNECT TO ACTIVE PLAXIS OUTPUT SESSION
# ==============================================================================
try:
    s_o, g_o = new_server("localhost", PORT, password=PASSWORD)
    print(f"[INFO] Connected to active PLAXIS Output session on port {PORT}.")
except Exception as e:
    raise ConnectionError(f"[ERROR] Connection failed: {e}")

phase = next((p for p in g_o.Phases if PHASE_NAME.lower()
             in p.Identification[:].lower()), None)
if phase is None:
    raise ValueError(f"[ERROR] Phase '{PHASE_NAME}' not found in open model.")
print(f"[INFO] Active Phase Locked: {phase.Identification[:]}")


def get_plaxis_attribute(result_group, target_keywords):
    """Dynamically matches result attributes across different PLAXIS 3D releases."""
    available_attrs = [a for a in dir(result_group) if not a.startswith('_')]
    for kw in target_keywords:
        for attr in available_attrs:
            if kw.lower() == attr.lower():
                return getattr(result_group, attr)
    for kw in target_keywords:
        for attr in available_attrs:
            if kw.lower() in attr.lower():
                return getattr(result_group, attr)
    return None


# ==============================================================================
# 3. EXTRACT PLATE RESULTS (Uz, M11, M22)
# ==============================================================================
print("[INFO] Querying Plate structural results...")
px = np.array(g_o.getresults(phase, g_o.ResultTypes.Plate.X, 'node'))
py = np.array(g_o.getresults(phase, g_o.ResultTypes.Plate.Y, 'node'))
pz = np.array(g_o.getresults(phase, g_o.ResultTypes.Plate.Z, 'node'))

uz = np.array(g_o.getresults(
    phase, g_o.ResultTypes.Plate.Uz, 'node')) * 1000.0  # m -> mm

m11_attr = get_plaxis_attribute(g_o.ResultTypes.Plate, ['M11', 'M1'])
m22_attr = get_plaxis_attribute(g_o.ResultTypes.Plate, ['M22', 'M2'])

m11 = np.array(g_o.getresults(phase, m11_attr, 'node'))
m22 = np.array(g_o.getresults(phase, m22_attr, 'node'))

z_top = Z_BASE + THICKNESS
plate_mask = (px >= X_MIN) & (px <= X_MAX) & \
             (py >= Y_MIN) & (py <= Y_MAX) & \
             (pz >= Z_BASE - 0.1) & (pz <= z_top + 0.1)

px_mat, py_mat = px[plate_mask], py[plate_mask]
uz_mat, m11_mat, m22_mat = uz[plate_mask], m11[plate_mask], m22[plate_mask]

# ==============================================================================
# 4. EXTRACT INTERFACE EFFECTIVE SOIL STRESS
# ==============================================================================
print("[INFO] Querying Interface Effective Normal Stress...")
sig_eff_mat = np.full(len(px_mat), np.nan)

try:
    eff_attr = g_o.ResultTypes.Interface.InterfaceEffectiveNormalStress
    ix = np.array(g_o.getresults(phase, g_o.ResultTypes.Interface.X, 'node'))
    iy = np.array(g_o.getresults(phase, g_o.ResultTypes.Interface.Y, 'node'))
    iz = np.array(g_o.getresults(phase, g_o.ResultTypes.Interface.Z, 'node'))
    sp_eff = np.array(g_o.getresults(phase, eff_attr, 'node'))

    # Extended boundary mask to capture interface nodes right up to cantilever edges
    if_mask = (ix >= X_MIN - 0.5) & (ix <= X_MAX + 0.5) & \
              (iy >= Y_MIN - 0.5) & (iy <= Y_MAX + 0.5) & \
              (np.abs(iz - Z_BASE) <= 0.5)

    ix_m, iy_m, eff_m = ix[if_mask], iy[if_mask], sp_eff[if_mask]

    if len(ix_m) > 0:
        tree = cKDTree(np.column_stack((ix_m, iy_m)))
        dists, idxs = tree.query(np.column_stack((px_mat, py_mat)))
        sig_eff_mat = np.where(dists < 1.20, np.abs(eff_m[idxs]), np.nan)
        print(f"[SUCCESS] Mapped interface stress for {len(ix_m)} nodes.")
except Exception as e:
    print(f"[ERROR] Failed extracting interface data: {e}")

# ==============================================================================
# 5. LOCALIZED BIN-BY-BIN ADAPTIVE AKIMA 1D INTERPOLATION ENGINE
# ==============================================================================


def akima_1d_slice(x_raw, y_raw, values_raw, cut_val, is_y_cut, target_coords, bin_size=1.0):
    """
    Slices raw FE mesh nodes along a cut plane using localized spatial binning:
      - Interior bins: Enforces strict original tiers [0.12, 0.22, 0.35] to preserve pure 1D path.
      - Overhang bins: Applies targeted tiers [0.12, 0.22, 0.35, 0.50].
      - Overhang stencil clamping: Restricts node borrowing so search windows do not cross 
        interior column lines, protecting short (0.55m) cantilevers from span contamination.
      - Boundary tail fallback: Queries 2D FE mesh triangulation for target points 
        beyond raw 1D node extent, preventing flat horizontal lines at Y > 20m.
    """
    slice_axis = x_raw if is_y_cut else y_raw
    cross_axis = y_raw if is_y_cut else x_raw

    axis_min, axis_max = target_coords.min(), target_coords.max()
    bin_edges = np.arange(axis_min, axis_max + bin_size, bin_size)

    selected_node_indices = []

    # 1. Localized Bin-by-Bin Search
    for i in range(len(bin_edges) - 1):
        b_start, b_end = bin_edges[i], bin_edges[i + 1]
        bin_mask = (slice_axis >= b_start) & (slice_axis <= b_end)

        is_overhang = (b_start < 0.0) or (b_end > 30.0) if is_y_cut else (
            b_start < 0.0) or (b_end > 20.0)
        col_bound_far = 30.0 if is_y_cut else 20.0
        tiers = [0.12, 0.22, 0.35, 0.50] if is_overhang else [0.12, 0.22, 0.35]

        bin_captured = []
        for tol in tiers:
            cross_mask = np.abs(cross_axis - cut_val) <= tol
            combined_mask = bin_mask & cross_mask & ~np.isnan(values_raw)

            captured_indices = np.where(combined_mask)[0]
            if len(captured_indices) > 0:
                bin_captured = captured_indices.tolist()
                break

        # Stencil expansion fallback (Cross-axis up to 0.80m for overhangs)
        if len(bin_captured) < 6:
            max_tol = 0.80 if is_overhang else 0.35

            # Clamp backward search for overhang bins at the interior column line
            if is_overhang:
                if b_start < 0.0:  # Bottom/Left cantilever (< 0m)
                    s_min = b_start - 0.75
                    s_max = min(b_end + 0.75, 0.0)
                else:             # Top/Right cantilever (> 30m X / > 20m Y)
                    s_min = max(b_start - 0.75, col_bound_far)
                    s_max = b_end + 0.75
            else:
                s_min = b_start - 0.75
                s_max = b_end + 0.75

            exp_mask = (slice_axis >= s_min) & (slice_axis <= s_max) & \
                       (np.abs(cross_axis - cut_val) <=
                        max_tol) & ~np.isnan(values_raw)
            bin_captured = np.where(exp_mask)[0].tolist()

        selected_node_indices.extend(bin_captured)

    if len(selected_node_indices) == 0:
        return np.full(len(target_coords), np.nan)

    selected_node_indices = np.unique(selected_node_indices)

    sc = slice_axis[selected_node_indices]
    sv = values_raw[selected_node_indices]

    valid = ~np.isnan(sc) & ~np.isnan(sv)
    sc, sv = sc[valid], sv[valid]

    sort_idx = np.argsort(sc)
    sc_sorted, sv_sorted = sc[sort_idx], sv[sort_idx]

    sc_unique, inverse_idx = np.unique(sc_sorted, return_inverse=True)
    sv_unique = np.bincount(
        inverse_idx, weights=sv_sorted) / np.bincount(inverse_idx)

    # 2D FE Mesh Triangulation Fallback for Boundary Coordinates
    raw_coords_2d = np.column_stack((x_raw, y_raw))
    valid_2d = ~np.isnan(values_raw)

    def fill_missing_with_2d_mesh(coords_1d, computed_values):
        out_mask = np.isnan(computed_values) | (
            coords_1d < sc_unique.min()) | (coords_1d > sc_unique.max())
        if np.any(out_mask):
            target_x = np.full(
                np.sum(out_mask), cut_val) if not is_y_cut else coords_1d[out_mask]
            target_y = coords_1d[out_mask] if not is_y_cut else np.full(
                np.sum(out_mask), cut_val)

            pts_out = np.column_stack((target_x, target_y))
            vals_2d_lin = griddata(
                raw_coords_2d[valid_2d], values_raw[valid_2d], pts_out, method='linear')
            vals_2d_near = griddata(
                raw_coords_2d[valid_2d], values_raw[valid_2d], pts_out, method='nearest')

            computed_values[out_mask] = np.where(
                np.isnan(vals_2d_lin), vals_2d_near, vals_2d_lin)
        return computed_values

    if len(sc_unique) < 5:
        if len(sc_unique) >= 2:
            f_lin = interp1d(sc_unique, sv_unique, kind='linear',
                             fill_value=np.nan, bounds_error=False)
            res = f_lin(target_coords)
            return fill_missing_with_2d_mesh(target_coords, res)
        return fill_missing_with_2d_mesh(target_coords, np.full(len(target_coords), np.nan))

    # 2. Fit Continuous Akima 1D Spline
    try:
        akima_model = Akima1DInterpolator(sc_unique, sv_unique)
        interp_res = akima_model(target_coords)
        return fill_missing_with_2d_mesh(target_coords, interp_res)
    except Exception:
        f_fallback = interp1d(
            sc_unique, sv_unique, kind='linear', fill_value=np.nan, bounds_error=False)
        res = f_fallback(target_coords)
        return fill_missing_with_2d_mesh(target_coords, res)


# ==============================================================================
# 6. RESAMPLE FULL 2D GRID & GENERATE AKIMA CUTLINES
# ==============================================================================
print(f"[INFO] Resampling 2D grid and fitting Localized Adaptive Akima 1D splines...")

x_grid_vec = np.arange(X_MIN, X_MAX + GRID_STEP, GRID_STEP)
y_grid_vec = np.arange(Y_MIN, Y_MAX + GRID_STEP, GRID_STEP)
grid_x, grid_y = np.meshgrid(x_grid_vec, y_grid_vec)

# 2D Full Grid Surface (Bivariate linear/nearest interpolation)
raw_points = np.column_stack((px_mat, py_mat))


def interpolate_2d_field(values):
    if np.all(np.isnan(values)):
        return np.full(grid_x.shape, np.nan)
    lin = griddata(raw_points, values, (grid_x, grid_y), method='linear')
    near = griddata(raw_points, values, (grid_x, grid_y), method='nearest')
    return np.where(np.isnan(lin), near, lin)


df_full_grid = pd.DataFrame({
    'X_m': grid_x.flatten(),
    'Y_m': grid_y.flatten(),
    'Uz_mm': interpolate_2d_field(uz_mat).flatten(),
    'Effective_Soil_Stress_kPa': interpolate_2d_field(sig_eff_mat).flatten(),
    'M11_kNm_m': interpolate_2d_field(m11_mat).flatten(),
    'M22_kNm_m': interpolate_2d_field(m22_mat).flatten()
})

# ==============================================================================
# 7. EXPORT TO MULTI-SHEET EXCEL REPORT WITH LOCALIZED AKIMA CUTLINES
# ==============================================================================
print(
    f"[INFO] Writing Excel workbook with Localized Adaptive Akima cutline profiles: {OUTPUT_EXCEL}")
cols = ['Joint_ID', 'X_m', 'Y_m', 'Uz_mm',
        'Effective_Soil_Stress_kPa', 'M11_kNm_m', 'M22_kNm_m']

with pd.ExcelWriter(OUTPUT_EXCEL, engine='openpyxl') as writer:
    # 1. Full 2D Grid Sheet
    df_full_grid.to_excel(writer, sheet_name='Full_Mat_Grid', index=False)

    # 2. Y-Cutlines using Localized Adaptive Akima Splines
    for yc in Y_CUTLINES:
        uz_akima = akima_1d_slice(
            px_mat, py_mat, uz_mat, yc, is_y_cut=True, target_coords=x_grid_vec)
        sig_akima = akima_1d_slice(
            px_mat, py_mat, sig_eff_mat, yc, is_y_cut=True, target_coords=x_grid_vec)
        m11_akima = akima_1d_slice(
            px_mat, py_mat, m11_mat, yc, is_y_cut=True, target_coords=x_grid_vec)
        m22_akima = akima_1d_slice(
            px_mat, py_mat, m22_mat, yc, is_y_cut=True, target_coords=x_grid_vec)

        print(f"  -> Extracted Y-Cutline at Y = {yc:5.1f} m")

        df_y = pd.DataFrame({
            'Joint_ID': [f"GRID_Y{yc}_{i}" for i in range(len(x_grid_vec))],
            'X_m': x_grid_vec,
            'Y_m': np.full(len(x_grid_vec), yc),
            'Uz_mm': uz_akima,
            'Effective_Soil_Stress_kPa': sig_akima,
            'M11_kNm_m': m11_akima,
            'M22_kNm_m': m22_akima
        })
        sheet_name = f"Cut_Y={int(yc) if yc == int(yc) else yc}m"
        df_y[cols].to_excel(writer, sheet_name=sheet_name, index=False)

    # 3. X-Cutlines using Localized Adaptive Akima Splines
    for xc in X_CUTLINES:
        uz_akima = akima_1d_slice(
            px_mat, py_mat, uz_mat, xc, is_y_cut=False, target_coords=y_grid_vec)
        sig_akima = akima_1d_slice(
            px_mat, py_mat, sig_eff_mat, xc, is_y_cut=False, target_coords=y_grid_vec)
        m11_akima = akima_1d_slice(
            px_mat, py_mat, m11_mat, xc, is_y_cut=False, target_coords=y_grid_vec)
        m22_akima = akima_1d_slice(
            px_mat, py_mat, m22_mat, xc, is_y_cut=False, target_coords=y_grid_vec)

        print(f"  -> Extracted X-Cutline at X = {xc:5.1f} m")

        df_x = pd.DataFrame({
            'Joint_ID': [f"GRID_X{xc}_{i}" for i in range(len(y_grid_vec))],
            'X_m': np.full(len(y_grid_vec), xc),
            'Y_m': y_grid_vec,
            'Uz_mm': uz_akima,
            'Effective_Soil_Stress_kPa': sig_akima,
            'M11_kNm_m': m11_akima,
            'M22_kNm_m': m22_akima
        })
        sheet_name = f"Cut_X={int(xc) if xc == int(xc) else xc}m"
        df_x[cols].to_excel(writer, sheet_name=sheet_name, index=False)

print("==========================================================================")
print("[SUCCESS] Processing Complete!")
print("Full 2D Mat surface and Localized Adaptive Akima 1D cutlines exported.")
print("==========================================================================")
