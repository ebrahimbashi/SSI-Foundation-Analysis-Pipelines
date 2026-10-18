import math
import os
import matplotlib.pyplot as plt
import numpy as np
import pandas as pd

# ==============================================================================
# 1. FILE MAPPING & MASTER VISUAL STYLING
# ==============================================================================
MODEL_FILES = {
    "M4": "M4_Fully_Coupled_SSI_Baseline.xlsx",
    "M3": "M3_Semi_Coupled_SSI.xlsx",
    "M1": "M1_Uniform_Winkler.xlsx",
    "M2a": "M2a_Pseudo_Coduto.xlsx",
    "M2b": "M2b_Pseudo_Bowles.xlsx",
    "M2c": "M2c_Pseudo_Hanse_George.xlsx",
}

UNIFIED_STYLES = {
    "M4": {
        "color": "#FCD116",
        "ls": "-",
        "lw": 3.0,
        "label": "Fully Coupled SSI Baseline (M4)",
    },
    "M3": {
        "color": "#D95F02",
        "ls": "-",
        "lw": 2.5,
        "label": "Semi-Coupled SSI (PLAXIS 3D) (M3)",
    },
    "M1": {
        "color": "#7570B3",
        "ls": "--",
        "lw": 1.8,
        "label": "Winkler Model (SAFE Baseline) (M1)",
    },
    "M2a": {
        "color": "#1B9E77",
        "ls": ":",
        "lw": 1.8,
        "label": "Pseudo-Coupled Coduto (M2a)",
    },
    "M2b": {
        "color": "#E7298A",
        "ls": "-.",
        "lw": 1.8,
        "label": "Pseudo-Coupled Bowles (M2b)",
    },
    "M2c": {
        "color": "#333333",
        "ls": (0, (6, 4)),
        "lw": 2.2,
        "label": "Pseudo-Coupled Hansen-George (M2c)",
    },
}

CUTLINES_Y = [f"Cut_Y={v}m" for v in [0, 5, 10, 15, 20]]
CUTLINES_X = [f"Cut_X={v}m" for v in [0, 5, 10, 15, 20, 25, 30]]
ALL_CUTLINES = [(sheet, "Y") for sheet in CUTLINES_Y] + [
    (sheet, "X") for sheet in CUTLINES_X
]

OUTPUT_DIR = "./Cutline_Comparative_Figures"
os.makedirs(OUTPUT_DIR, exist_ok=True)


# ==============================================================================
# 2. COLUMN RESOLUTION HELPER
# ==============================================================================
def resolve_columns(df, cutline_type):
    cols = df.columns.tolist()

    coord_candidates = (
        ["X_m", "X", "X_Coordinate_m", "(X,10)"]
        if cutline_type == "Y"
        else ["Y_m", "Y", "Y_Coordinate_m", "(Y,10)"]
    )
    coord_col = next((c for c in coord_candidates if c in cols), None)

    sett_candidates = [
        "Uz_mm",
        "Uz",
        "PLAXIS_Settlement_mm",
        "Settlement",
        "Uz.1",
        "Uz.2",
    ]
    sett_col = next((c for c in sett_candidates if c in cols), None)

    pres_candidates = [
        "Effective_Soil_Stress_kPa",
        "Soil_Pressure_kPa",
        "soil pressure",
        "PLAXIS_SoilPressure_kPa",
        "Pressure_kPa",
        "soil pressure.1",
        "soil pressure.2",
    ]
    pres_col = next((c for c in pres_candidates if c in cols), None)

    mom_candidates = (
        [
            "M11_kNm_m",
            "M11",
            "Moment(M-1)",
            "M1(comb1)",
            "PLAXIS_BendingMoment_M11_kNm_m",
            "M3(comb1)",
            "M3",
        ]
        if cutline_type == "Y"
        else [
            "M22_kNm_m",
            "M22",
            "Moment(M-2)",
            "M2(comb1)",
            "PLAXIS_BendingMoment_M22_kNm_m",
            "M3(comb1)",
            "M3",
        ]
    )
    mom_col = next((c for c in mom_candidates if c in cols), None)

    return coord_col, sett_col, pres_col, mom_col


# ==============================================================================
# 3. BATCH GENERATION ENGINE
# ==============================================================================
def generate_all_cutlines():
    plt.style.use(
        "seaborn-v0_8-whitegrid"
        if "seaborn-v0_8-whitegrid" in plt.style.available
        else "default"
    )
    plt.rcParams["font.family"] = "serif"

    print(
        f"Starting direct Excel extraction for {len(ALL_CUTLINES)} cutlines...")

    for sheet_name, cutline_type in ALL_CUTLINES:
        fig, axs = plt.subplots(
            nrows=3, ncols=1, figsize=(11, 16), sharex=True, dpi=300
        )

        all_moments = []
        all_pressures = []

        for model_key, file_path in MODEL_FILES.items():
            if not os.path.exists(file_path):
                continue

            try:
                df = pd.read_excel(file_path, sheet_name=sheet_name)
                c_coord, c_sett, c_pres, c_mom = resolve_columns(
                    df, cutline_type
                )

                if not all([c_coord, c_sett, c_pres, c_mom]):
                    continue

                df_clean = df.dropna(
                    subset=[c_coord, c_sett, c_pres, c_mom]
                ).copy()
                df_clean[c_coord] = pd.to_numeric(
                    df_clean[c_coord], errors="coerce"
                )
                df_clean = df_clean.sort_values(by=c_coord).reset_index(
                    drop=True
                )

                x_val = df_clean[c_coord].to_numpy().astype(float)
                sett_val = -np.abs(
                    df_clean[c_sett].to_numpy().astype(float)
                )  # Enforce negative settlement
                pres_val = np.abs(
                    df_clean[c_pres].to_numpy().astype(float)
                )  # Enforce positive pressure

                # Extract raw bending moments
                mom_val = df_clean[c_mom].to_numpy().astype(float)

                # FLIP SIGN FOR M3 AND M4 ONLY
                if model_key in ["M3", "M4"]:
                    mom_val = -mom_val

                all_pressures.extend(pres_val)
                all_moments.extend(mom_val)

                style = UNIFIED_STYLES[model_key]

                axs[0].plot(x_val, sett_val, **style)
                axs[1].plot(x_val, pres_val, **style)
                axs[2].plot(x_val, mom_val, **style)

            except Exception as e:
                print(
                    f"Warning: Could not process {sheet_name} from {file_path}: {e}"
                )

        # Panel 0: Settlement (Uz)
        coord_label = (
            "Cross-Section Position X (m)"
            if cutline_type == "Y"
            else "Cross-Section Position Y (m)"
        )
        axs[0].set_ylabel(
            "Vertical Settlement (mm)", fontsize=11, fontweight="bold"
        )
        axs[0].set_title(
            f"Combined Soil-Structure Interaction Analysis — {sheet_name}",
            fontsize=14,
            fontweight="bold",
            pad=15,
        )
        axs[0].set_ylim(-170, 5)
        axs[0].grid(True, linestyle=":", alpha=0.6)
        axs[0].legend(
            loc="upper right",
            frameon=True,
            framealpha=0.95,
            edgecolor="#cccccc",
            fontsize=8.5,
        )

        # Panel 1: Soil Pressure
        max_p = max(all_pressures) if all_pressures else 400
        axs[1].set_ylabel(
            "Contact Soil Pressure (kPa)", fontsize=11, fontweight="bold"
        )
        axs[1].set_ylim(-5, max_p * 1.25)
        axs[1].grid(True, linestyle=":", alpha=0.6)
        axs[1].legend(
            loc="upper right",
            frameon=True,
            framealpha=0.95,
            edgecolor="#cccccc",
            fontsize=8.5,
        )

        # Panel 2: Bending Moment
        min_m = min(all_moments) if all_moments else -2000
        max_m = max(all_moments) if all_moments else 500

        fixed_lower_limit = min_m - 150 if min_m < 0 else -400
        fixed_upper_limit = max_m + (max_m * 0.25)
        total_span = fixed_upper_limit - fixed_lower_limit

        tick_interval = (
            500 if total_span > 2500 else (200 if total_span > 1200 else 100)
        )
        start_tick = (
            math.floor(fixed_lower_limit / tick_interval) * tick_interval
        )
        end_tick = math.ceil(fixed_upper_limit / tick_interval) * tick_interval
        explicit_ticks = np.arange(
            start_tick, end_tick + tick_interval, tick_interval
        )

        mom_label = (
            "Bending Moment M11 (kNm/m)"
            if cutline_type == "Y"
            else "Bending Moment M22 (kNm/m)"
        )
        axs[2].set_xlabel(coord_label, fontsize=11, fontweight="bold")
        axs[2].set_ylabel(mom_label, fontsize=11, fontweight="bold")
        axs[2].set_ylim(fixed_lower_limit, fixed_upper_limit)
        axs[2].set_yticks(explicit_ticks)
        axs[2].axhline(
            0, color="#000000", linestyle="-", linewidth=1.3, alpha=0.8
        )
        axs[2].grid(True, linestyle=":", alpha=0.6)
        axs[2].legend(
            loc="upper right",
            frameon=True,
            framealpha=0.95,
            edgecolor="#cccccc",
            fontsize=8.5,
        )

        domain_max = 31.0 if cutline_type == "Y" else 21.0
        for ax in axs:
            ax.set_xlim(-1.0, domain_max)

        plt.tight_layout()

        pdf_path = os.path.join(OUTPUT_DIR, f"Profiles_{sheet_name}.pdf")
        png_path = os.path.join(OUTPUT_DIR, f"Profiles_{sheet_name}.png")

        plt.savefig(pdf_path, format="pdf", bbox_inches="tight")
        plt.savefig(png_path, format="png", dpi=300, bbox_inches="tight")
        plt.close(fig)

        print(f"Generated: {pdf_path}")

    print("\nBatch generation complete! All 12 cutline figures updated.")


if __name__ == "__main__":
    generate_all_cutlines()
