import sys
from plxscripting.easy import *
import numpy as np
from scipy.spatial import Voronoi
from shapely.geometry import Polygon, box

# Logger class to write output simultaneously to terminal and a text file


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

# 1. Connect to PLAXIS Output Server
s_out, g_out = new_server(port=10001, password="GvYr8!m#>5u?Lt!W")
phase = g_out.Phases[-1]

print("Executing complete boundary-aware Voronoi integration...")

# 2. Extract Soil volume total vertical stress (SigzzT)
sigzz_raw = g_out.getresults(phase, g_out.ResultTypes.Soil.SigzzT, "node")
x_raw = g_out.getresults(phase, g_out.ResultTypes.Soil.X, "node")
y_raw = g_out.getresults(phase, g_out.ResultTypes.Soil.Y, "node")
z_raw = g_out.getresults(phase, g_out.ResultTypes.Soil.Z, "node")

# Mat Footprint (32m x 22m)
x_min, x_max = -1.0, 31.0
y_min, y_max = -1.0, 21.0
target_z = -3.0
tol = 0.05

mat_boundary = box(x_min, y_min, x_max, y_max)

# 3. Resolve Coincident Nodes
node_dict = {}
for x, y, z, sig in zip(x_raw, y_raw, z_raw, sigzz_raw):
    if abs(float(z) - target_z) < tol:
        xv, yv = float(x), float(y)
        if (x_min - tol) <= xv <= (x_max + tol) and (y_min - tol) <= yv <= (y_max + tol):
            key = (round(xv, 3), round(yv, 3))
            sig_val = abs(float(sig))
            if key not in node_dict or sig_val > node_dict[key]:
                node_dict[key] = sig_val

points = np.array(list(node_dict.keys()))
stresses = np.array(list(node_dict.values()))

# 4. Add Outer Boundary Ghost Frame to Force Bounded Voronoi Cells on Edge Nodes
buffer_dist = 10.0
ghost_frame = np.array([
    [x_min - buffer_dist, y_min - buffer_dist],
    [x_max + buffer_dist, y_min - buffer_dist],
    [x_max + buffer_dist, y_max + buffer_dist],
    [x_min - buffer_dist, y_max + buffer_dist]
])

all_points = np.vstack([points, ghost_frame])

# Run Voronoi on point set including ghost boundary
vor = Voronoi(all_points)

integrated_force = 0.0
total_computed_area = 0.0

# Only iterate over real mat nodes (exclude ghost frame)
for i in range(len(points)):
    region_idx = vor.point_region[i]
    region = vor.regions[region_idx]

    if len(region) > 0 and -1 not in region:
        polygon_vertices = [vor.vertices[v] for v in region]
        poly = Polygon(polygon_vertices)

        # Clip cell strictly to the 32m x 22m footprint
        clipped_poly = poly.intersection(mat_boundary)

        if not clipped_poly.is_empty:
            area = clipped_poly.area
            integrated_force += stresses[i] * area
            total_computed_area += area

# 5. Output Verification
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
