README_s0_generator
===================

# ------------------------------------------------------------------ #
#           s0 file generator_MADE BY NICOLAS NAVARRO                #
# ------------------------------------------------------------------ #

Purpose
-------
This Python script creates KFX-View batch picture scripts (*.s0) for a set of
simulation cases listed in an Excel file.  
The generated .s0 files can be executed with **KFX-View version 3.8** to
produce JPEG images for CO, CO₂, and radiation at predefined camera angles.

Files produced per case
-----------------------
For every row in the Excel sheet the script writes six files:

  <CaseID>_CO_Y.jpg.s0        – CO plume, Y-axis view  
  <CaseID>_CO2_X.jpg.s0       – CO₂ plume, X-axis view  
  <CaseID>_CO2_Y.jpg.s0       – CO₂ plume, Y-axis view  
  <CaseID>_rad_Y.jpg.s0       – Radiation contour, Y-axis view  
  <CaseID>_rad_Z.jpg.s0       – Radiation contour, Z-axis view  
  <CaseID>_rad_xyz.jpg.s0     – Radiation contour, oblique xyz view  

Each script
1. Loads the fixed P58 geometry model (3DModel_P58.kfx).  
2. Loads three result fields (…_bmg_exit.r3d, …_CO2_exit.r3d, …_CO_exit.r3d)
   whose full paths come from the Excel sheet.  
3. Applies camera / contour / iso-surface settings stored in the template.  
4. Saves a JPEG picture with the name shown above.

Prerequisites
-------------
• Python 3 with the **pandas** package installed (`pip install pandas`).  
• **KFX-View 3.8.x** (this version renders correctly with the exported scripts).

Excel layout
------------
Workbook: `s0_template_cases.xlsx`  
Required columns (exact spelling):

  Case ID          – e.g. CUC5412001A_downstream_2_negx_NE  
  Directory bmg    – full path to *_bmg_exit.r3d  
  Directory CO2    – full path to *_CO2_exit.r3d  
  Directory CO     – full path to *_CO_exit.r3d  

Running the script
------------------
1. Open `generate_s0_files.py` and edit the three paths at the top:

   • EXCEL_PATH  – location of the Excel file  
   • OUTPUT_DIR  – folder where the .s0 scripts will be written  
   • MODEL_PATH  – full path to 3DModel_P58.kfx

2. In a command prompt run

       python generate_s0_files.py

3. Check OUTPUT_DIR – six .s0 files per case are created.

Batch execution in KFX-View 3.8
-------------------------------
Run each script with the –b (batch) switch, e.g.

    kfxview.exe  -b  CUC5412001A_downstream_2_negx_NE_CO_Y.jpg.s0

Repeat (or loop) for every script.  
Each run loads, grabs one frame, saves the image, and exits.

Notes / troubleshooting
-----------------------
• Camera numbers and contour settings are hard-coded.  
  To add new views, copy a block in the script, edit the numbers, and add it to
  the TEMPLATES list.  
• Image resolution is 1280 × 649.  
  In some remote-desktop sessions only part of the image may appear; reduce
  SIZEPX or run KFX-View locally.  
• All three result-field files must exist.  If one is missing KFX-View issues a
  warning but still saves the picture.  
• Use **KFX-View 3.8.x** – newer GUI versions may treat the “> key …” macros
  differently and the batch render may fail or give black areas.

Enjoy your automated figure generation!