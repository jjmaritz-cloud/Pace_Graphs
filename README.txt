Amino Eggsactly Streamlit Graph App v12
======================================

Purpose
-------
Standalone Excel-backed graphing app for stitched rearing-to-layer flock performance.

How to run
----------
1. Unzip the package.
2. Double-click run_amino_graphs.bat.
3. The first run may install Python packages.
4. Streamlit should open in your browser.

Refreshing data from the app
----------------------------
1. Refresh the Excel query in Amino_Eggsactly_Data_V1.xlsx.
2. Open the Streamlit app.
3. Use the sidebar upload box: Upload Amino_Eggsactly_Data_V1.xlsx.
4. Click "Save uploaded data workbook and refresh".
5. The app saves the uploaded workbook into the data folder and refreshes the graphs.

The previous data workbook is backed up automatically in:
data/backups/

Manual refresh option
---------------------
You can still manually replace files in the data folder if preferred:
- data/Amino_Eggsactly_Data_V1.xlsx
- data/Amino_Eggsactly_Rearing_Layer_Match.xlsx

Then click Clear cache / refresh data in the sidebar.

Backend files
-------------
data/Amino_Eggsactly_Data_V1.xlsx
- DATA
- Standards ISA Floor

data/Amino_Eggsactly_Rearing_Layer_Match.xlsx
- Import_Bridge
- Layer_Flock_Summary
- Rearing_Flock_Summary

Important graph rules
---------------------
- DATA sheet is filtered to Reporting_Period = Weekly only.
- Standards are not matched by breed; the app uses the single default standards curve and aligns it by age in weeks.
- Latest incomplete weekly layer row is excluded by default when the 7_Days column indicates it is incomplete.
- Zero metric values are treated as missing so lines connect rather than dropping to zero.
- Decimal values are rounded to a maximum of 2 decimal places.
- Cumulative eggs per bird axis is fixed at 0-450 eggs.

v12 update
----------
- Added sidebar upload feature for Amino_Eggsactly_Data_V1.xlsx.
- Uploaded workbook is validated before saving.
- Previous backend workbook is backed up automatically before replacement.
- Cache is cleared and the app refreshes after upload.
