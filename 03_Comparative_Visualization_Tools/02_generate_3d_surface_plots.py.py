import os
import matplotlib.pyplot as plt
import numpy as np
import pandas as pd

# ==============================================================================
# 1. FILE & SHEET MAPPING CONFIGURATION
# ==============================================================================
MODEL_CONFIG = {
    "M1": {
        "file": "M1_Uniform_Winkler.xlsx",
        "sheet": "Full_Mesh_Data",
        "label": "M1: Uniform Winkler Model (SAFE Baseline)",
    },
    "M4": {
        "file": "M4_Fully_Coupled_SSI_Baseline.xlsx",
        "sheet": "Full_Mat_Grid",
        "label": "M4: Fully Coupled SSI Baseline",
    },
}

OUTPUT_DIR = "./Surface_3D_Figures"
os.makedirs(OUTPUT_DIR, exist_ok=True)


# ==============================================================================
# 2. HELPER DATA EXTRACTION FUNCTION
# ==============================================================================
def load_grid_data(model_key):
    """Loads 3D grid data directly from Excel, resolving coordinates, settlement,

    soil pressure, and bending moments M11 & M22.
    """
    cfg = MODEL_CONFIG[model_key]
    file_path = cfg["file"]
    sheet_name = cfg["sheet"]

    if not os.path.exists(file_path):
        raise FileNotFoundError(f"Required file not found: {file_path}")

    xl = pd.ExcelFile(file_path)
    if sheet_name not in xl.sheet_names:
        raise ValueError(
            f"Expected sheet '{sheet_name}' not found in {file_path}. "
            f"Available sheets: {xl.sheet_names}"
        )

    df = pd.read_excel(xl, sheet_name=sheet_name)
    cols = df.columns.tolist()

    # Resolve Coordinates
    x_col = next(
        (c for c in ["X_m", "X", "X_Coordinate_m"] if c in cols), None
    )
    y_col = next(
        (c for c in ["Y_m", "Y", "Y_Coordinate_m"] if c in cols), None
    )

    # Resolve Geotechnical Outputs
    sett_col = next(
        (
            c
            for c in [
                "Uz_mm",
                "Uz",
                "PLAXIS_Settlement_mm",
                "Settlement",
                "Uz.1",
            ]
            if c in cols
        ),
        None,
    )
    pres_col = next(
        (
            c
            for c in [
                "Effective_Soil_Stress_kPa",
                "Soil_Pressure_kPa",
                "soil pressure",
                "PLAXIS_SoilPressure_kPa",
                "Pressure_kPa",
            ]
            if c in cols
        ),
        None,
    )

    # Resolve Flexural Bending Moments
    m11_col = next(
        (
            c
            for c in [
                "M11_kNm_m",
                "M11",
                "Moment(M-1)",
                "M1(comb1)",
                "PLAXIS_BendingMoment_M11_kNm_m",
                "M3(comb1)",
                "M3",
            ]
            if c in cols
        ),
        None,
    )
    m22_col = next(
        (
            c
            for c in [
                "M22_kNm_m",
                "M22",
                "Moment(M-2)",
                "M2(comb1)",
                "PLAXIS_BendingMoment_M22_kNm_m",
            ]
            if c in cols
        ),
        None,
    )

    required_cols = [x_col, y_col, sett_col, pres_col, m11_col]
    if not all(required_cols):
        raise KeyError(
            f"Could not resolve required columns in {file_path} (Sheet: {sheet_name}). "
            f"Available columns: {cols}"
        )

    # Clean missing values
    active_cols = [x_col, y_col, sett_col, pres_col, m11_col] + (
        [m22_col] if m22_col else []
    )
    df_clean = df.dropna(subset=active_cols).copy()

    x = pd.to_numeric(df_clean[x_col], errors="coerce").to_numpy()
    y = pd.to_numeric(df_clean[y_col], errors="coerce").to_numpy()
    uz = -np.abs(
        pd.to_numeric(df_clean[sett_col], errors="coerce").to_numpy()
    )  # Enforce negative settlement
    p = np.abs(
        pd.to_numeric(df_clean[pres_col], errors="coerce").to_numpy()
    )  # Enforce positive pressure

    m11 = pd.to_numeric(df_clean[m11_col], errors="coerce").to_numpy()
    m22 = (
        pd.to_numeric(df_clean[m22_col], errors="coerce").to_numpy()
        if m22_col
        else np.copy(m11)
    )

    # Flip sign for M4 bending moments to match M1 convention
    if model_key == "M4":
        m11 = -m11
        m22 = -m22

    return x, y, uz, p, m11, m22


# ==============================================================================
# 3. 3D SURFACE PLOTTING ENGINE
# ==============================================================================
def plot_3d_surface_comparison(
    x1, y1, z1, x4, y4, z4, z_label, title, filename_prefix, cmap_name="viridis_r"
):
    """Generates side-by-side 3D surface comparative plots for M1 vs M4."""
    fig = plt.figure(figsize=(18, 8), dpi=300)
    plt.style.use(
        "seaborn-v0_8-whitegrid"
        if "seaborn-v0_8-whitegrid" in plt.style.available
        else "default"
    )
    plt.rcParams["font.family"] = "serif"

    # Shared Z-axis scale across both subplots for direct visual parity
    z_min = min(np.min(z1), np.min(z4))
    z_max = max(np.max(z1), np.max(z4))

    # --- SUBPLOT 1: M1 Uniform Winkler ---
    ax1 = fig.add_subplot(1, 2, 1, projection="3d")
    surf1 = ax1.plot_trisurf(
        x1, y1, z1, cmap=cmap_name, edgecolor="none", alpha=0.9, vmin=z_min, vmax=z_max
    )
    ax1.set_title(
        MODEL_CONFIG["M1"]["label"], fontsize=13, fontweight="bold", pad=12
    )
    ax1.set_xlabel("X Coordinate (m)", fontsize=10,
                   fontweight="bold", labelpad=8)
    ax1.set_ylabel("Y Coordinate (m)", fontsize=10,
                   fontweight="bold", labelpad=8)
    ax1.set_zlabel(z_label, fontsize=10, fontweight="bold", labelpad=8)
    ax1.set_zlim(z_min, z_max)
    ax1.view_init(elev=28, azim=-125)

    # --- SUBPLOT 2: M4 Fully Coupled SSI Baseline ---
    ax2 = fig.add_subplot(1, 2, 2, projection="3d")
    surf2 = ax2.plot_trisurf(
        x4, y4, z4, cmap=cmap_name, edgecolor="none", alpha=0.9, vmin=z_min, vmax=z_max
    )
    ax2.set_title(
        MODEL_CONFIG["M4"]["label"], fontsize=13, fontweight="bold", pad=12
    )
    ax2.set_xlabel("X Coordinate (m)", fontsize=10,
                   fontweight="bold", labelpad=8)
    ax2.set_ylabel("Y Coordinate (m)", fontsize=10,
                   fontweight="bold", labelpad=8)
    ax2.set_zlabel(z_label, fontsize=10, fontweight="bold", labelpad=8)
    ax2.set_zlim(z_min, z_max)
    ax2.view_init(elev=28, azim=-125)

    # Colorbar
    cbar_ax = fig.add_axes([0.92, 0.20, 0.018, 0.60])
    cbar = fig.colorbar(surf2, cax=cbar_ax)
    cbar.set_label(z_label, fontsize=11, fontweight="bold")

    fig.suptitle(title, fontsize=16, fontweight="bold", y=0.96)
    plt.subplots_adjust(wspace=0.15)

    # Export vector PDF and high-res PNG
    pdf_path = os.path.join(OUTPUT_DIR, f"{filename_prefix}.pdf")
    png_path = os.path.join(OUTPUT_DIR, f"{filename_prefix}.png")

    plt.savefig(pdf_path, format="pdf", bbox_inches="tight")
    plt.savefig(png_path, format="png", dpi=300, bbox_inches="tight")
    plt.close(fig)

    print(f"Generated 3D figure: {pdf_path}")


# ==============================================================================
# 4. MAIN EXECUTION
# ==============================================================================
def main():
    print("Extracting 3D full mesh data for M1 and M4...")

    x1, y1, uz1, p1, m11_1, m22_1 = load_grid_data("M1")
    x4, y4, uz4, p4, m11_4, m22_4 = load_grid_data("M4")

    # 1. Settlement Surface (Uz)
    plot_3d_surface_comparison(
        x1,
        y1,
        uz1,
        x4,
        y4,
        uz4,
        z_label="Vertical Settlement $U_z$ (mm)",
        title="3D Mat Settlement Surface: M1 (Uniform Winkler) vs. M4 (Fully Coupled SSI)",
        filename_prefix="3D_Surface_Settlement_M1_vs_M4",
        cmap_name="viridis_r",
    )

    # 2. Soil Contact Pressure Surface (p)
    plot_3d_surface_comparison(
        x1,
        y1,
        p1,
        x4,
        y4,
        p4,
        z_label="Contact Soil Pressure (kPa)",
        title="3D Contact Soil Pressure Surface: M1 (Uniform Winkler) vs. M4 (Fully Coupled SSI)",
        filename_prefix="3D_Surface_SoilPressure_M1_vs_M4",
        cmap_name="plasma",
    )

    # 3. Bending Moment Surface M11
    plot_3d_surface_comparison(
        x1,
        y1,
        m11_1,
        x4,
        y4,
        m11_4,
        z_label="Bending Moment $M_{11}$ (kNm/m)",
        title="3D Bending Moment $M_{11}$ Surface: M1 (Uniform Winkler) vs. M4 (Fully Coupled SSI)",
        filename_prefix="3D_Surface_Moment_M11_M1_vs_M4",
        cmap_name="coolwarm",
    )

    # 4. Bending Moment Surface M22
    plot_3d_surface_comparison(
        x1,
        y1,
        m22_1,
        x4,
        y4,
        m22_4,
        z_label="Bending Moment $M_{22}$ (kNm/m)",
        title="3D Bending Moment $M_{22}$ Surface: M1 (Uniform Winkler) vs. M4 (Fully Coupled SSI)",
        filename_prefix="3D_Surface_Moment_M22_M1_vs_M4",
        cmap_name="coolwarm",
    )

    print(
        "\nAll 4 comparative 3D surface figures (Uz, Pressure, M11, M22) successfully generated!"
    )


if __name__ == "__main__":
    main()
