import os
import pandas as pd

# ------------------------------------------------------------------ #
#           s0 file generator_MADE BY NICOLAS NAVARRO                #
# ------------------------------------------------------------------ #

# ------------------------------------------------------------------ #
# 1.  EDIT THESE PATHS
# ------------------------------------------------------------------ #
EXCEL_PATH = r"C:\Users\nnavarrosimancas\Documents\Post-processing scripts\s0_template_cases.xlsx"
OUTPUT_DIR = r"C:\Users\nnavarrosimancas\Documents\Post-processing scripts"   # <– no trailing “\”
MODEL_PATH = r"E:\BR Team\P58\3DModel_P58.kfx"          # geometry path (fixed)

os.makedirs(OUTPUT_DIR, exist_ok=True)

# ------------------------------------------------------------------ #
# 2.  Read Excel (must contain the 3 columns below)
# ------------------------------------------------------------------ #
df = pd.read_excel(EXCEL_PATH)
REQUIRED = ["Case ID", "Directory bmg", "Directory CO2", "Directory CO"]
if any(col not in df.columns for col in REQUIRED):
    raise ValueError("Excel is missing one of the required columns")

# ------------------------------------------------------------------ #
# 3.  Fixed template strings (only variable parts are {placeholders})
# ------------------------------------------------------------------ #
HEAD = f"""# Created by kfxview
ISOCOL: 1
GRATING: 0.2
INCLUDE: {MODEL_PATH}
SIZEPX: 1280 649
LOADFIELD: {{bmg}}
LOADFIELD: {{co2}}
LOADFIELD: {{co}}
"""

# ---- individual picture macros (camera, iso, contour etc.) -------- #
CO_Y = r"""SAVEIMAGE: {case}_CO_Y.jpg
> key + p f d O s 
> centerpoint 285.334 -1.10014 48.0573
> width 226.165
> radius 226.165
> v-ang 0 90
> XM -18.852
> YM -62.959
> ZM -1.91996
> XP 1045.5
> YP 57.501
> ZP 299.501
> bg 1 1 1
> fg 0 0 0
> frame 1
LEGFORMAT: %g
CONTOURPROJ: 0 0 0
ISOPARAM: 255 0 0 0.4 1 1
CONTOURMODE: 0 0 1 0 
CONTOUR: s 0.03 1 1 3 
"""

CO2_X = r"""SAVEIMAGE: {case}_CO2_X.jpg
> key + p f d O s 
> centerpoint 147.652 1.71812 36.9331
> width 325.678
> radius 325.678
> v-ang -0.176566 0
> XM -18.852
> YM -62.959
> ZM -1.91996
> XP 1045.5
> YP 57.501
> ZP 299.501
> bg 1 1 1
> fg 0 0 0
> frame 1
LEGFORMAT: %g
CONTOURPROJ: 0 0 0
ISOPARAM: 255 0 0 0.4 1 1
CONTOURMODE: 0 0 1 0 
CONTOUR: s 0.03 1 1 3 
"""

CO2_Y = r"""SAVEIMAGE: {case}_CO2_Y.jpg
> key + p f d O s 
> centerpoint 285.334 -1.10014 48.0573
> width 226.165
> radius 226.165
> v-ang 89.955 270
> XM -18.852
> YM -62.959
> ZM -1.91996
> XP 1045.5
> YP 57.501
> ZP 299.501
> bg 1 1 1
> fg 0 0 0
> frame 1
LEGFORMAT: %g
CONTOURPROJ: 0 0 0
ISOPARAM: 255 0 0 0.4 1 1
CONTOURMODE: 0 0 1 0 
CONTOUR: s 0.03 1 1 3 
"""

RAD_Y = r"""SAVEIMAGE: {case}_rad_Y.jpg
> key + p f d O s 
> centerpoint 275.856 20.578 40.485
> width 271.398
> radius 271.398
> v-ang 0 90
> XM -18.852
> YM -62.959
> ZM -1.91996
> XP 1045.5
> YP 57.501
> ZP 299.501
> bg 1 1 1
> fg 0 0 0
> frame 1
LEGFORMAT: %g
CONTOURPROJ: 0 0 0
ISOPARAM: 255 255 0 0.4 1 1
CONTOURMODE: 0 0 1 0 
CONTOUR: s 1580 0 1 1 
ISOPARAM: 255 0 0 0.4 1 1
CONTOURMODE: 0 0 0 0 
CONTOUR: s 4780 0 1 2 
"""

RAD_Z = r"""SAVEIMAGE: {case}_rad_Z.jpg
> key + p f O s 
> centerpoint 265.561 -1.44085 47.4999
> width 271.398
> radius 271.398
> v-ang 89.9943 270
> XM -18.852
> YM -62.959
> ZM -1.91996
> XP 1045.5
> YP 57.501
> ZP 299.501
> bg 1 1 1
> fg 0 0 0
> frame 1
LEGFORMAT: %g
CONTOURPROJ: 0 0 0
ISOPARAM: 255 255 0 0.4 1 1
CONTOURMODE: 0 0 1 0 
CONTOUR: s 1580 0 1 1 
ISOPARAM: 255 0 0 0.4 1 1
CONTOURMODE: 0 0 0 0 
CONTOUR: s 4780 0 1 2 
"""

RAD_XYZ = r"""SAVEIMAGE: {case}_rad_xyz.jpg
> key + p f d O s 
> centerpoint 275.856 20.578 40.485
> width 271.398
> radius 271.398
> v-ang 18.8926 108.174
> XM -18.852
> YM -62.959
> ZM -1.91996
> XP 1045.5
> YP 57.501
> ZP 299.501
> bg 1 1 1
> fg 0 0 0
> frame 1
LEGFORMAT: %g
CONTOURPROJ: 0 0 0
ISOPARAM: 255 255 0 0.4 1 1
CONTOURMODE: 0 0 1 0 
CONTOUR: s 1580 0 1 1 
ISOPARAM: 255 0 0 0.4 1 1
CONTOURMODE: 0 0 0 0 
CONTOUR: s 4780 0 1 2 
"""

# list of all picture templates to write
TEMPLATES = [
    ("CO_Y",   CO_Y),
    ("CO2_X",  CO2_X),
    ("CO2_Y",  CO2_Y),
    ("rad_Y",  RAD_Y),
    ("rad_Z",  RAD_Z),
    ("rad_xyz", RAD_XYZ)            # <-- new entry
]

# ------------------------------------------------------------------ #
# 4.  Generate .s0 files per case
# ------------------------------------------------------------------ #
for _, row in df.iterrows():
    case = row["Case ID"]
    paths = {
        "case": case,
        "bmg": row["Directory bmg"].strip(),
        "co2": row["Directory CO2"].strip(),
        "co":  row["Directory CO"].strip(),
    }
    head = HEAD.format(**paths)

    for suffix, body in TEMPLATES:
        s0_text = head + body.format(**paths)
        script_name = f"{case}_{suffix}.jpg.s0"
        with open(os.path.join(OUTPUT_DIR, script_name), "w", encoding="utf-8", newline="\n") as f:
            f.write(s0_text)
        print("Wrote", script_name)

print("\nAll .s0 scripts generated.")

