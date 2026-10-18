# SSI-Foundation-Analysis-Pipelines

Production Python automation and batch scripting codebase supporting the MSc Thesis research:
> **Evaluation of Soil-Structure Interaction (SSI) Modeling Frameworks for High-Rise Mat Foundations**

## 🛠 Repository Architecture & Execution Sequence

### 00_PLAXIS_Command_Runners/
* **`01_plaxis_geometry_surface_loads.txt`**: Constructs the 32m x 22m mat foundation geometry, generates 35 discrete column surface patches (0.9m x 0.9m), and assigns linearly varying vertical stress fields with zero out-of-plane gradients for service load combinations (Phase_3).
* **`02_plaxis_sloped_excavation.txt`**: Automates 3D soil excavation cuts and 2H:1V sloped boundary cutters around the foundation perimeter.

### 01_Data_Extraction_Pipelines/
* **`01_plaxis_3d_akima_extraction.py` & `02_plaxis_3d_phase3_extraction.py`**: Connects via PLAXIS 3D remote scripting API, queries plate/interface effective normal stress fields via cKDTree, and extracts multi-axis profiles using a Localized Adaptive Akima 1D Spline engine.
* **`03_csi_safe_database_parser.py`**: Parses CSI SAFE database export workbooks, mapping point object connectivity and joint displacements onto area shell mesh centroids via bivariate cubic surface interpolation.

### 02_Static_Equilibrium_Verification/
* **`01_voronoi_equilibrium_plaxis.py` & `02_voronoi_equilibrium_safe.py`**: Performs boundary-aware Bounded Voronoi Tessellation Integration to verify global vertical static force equilibrium across all numerical frameworks.

### 03_Comparative_Visualization_Tools/
* **`01_generate_2d_cutline_plots.py` & `02_generate_3d_surface_plots.py`**: Renders publication-quality 2D cross-sectional cutline curves and side-by-side 3D surface comparative visualizers.
* **`03_generate_master_docx_tables.py`**: Automates the generation of formatted Word (.docx) comparative data tables across 7 evaluation axes.

## 📦 Python Dependencies
```bash
pip install numpy pandas scipy matplotlib shapely python-docx openpyxl plxscripting
---
---
## 👤 Author
**Ebrahim Bashi**  
MSc Student in Structural Engineering  
School of Civil and Environmental Engineering, College of Technology and Built Environment, Addis Ababa University.
