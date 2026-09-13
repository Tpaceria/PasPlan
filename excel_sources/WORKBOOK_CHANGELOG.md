# Workbook change log

Updated file: `NEW_FINAL_LOGO_EXACT_5_05x1_00cm.xlsm` and copy `PasPlan_NORDLAND.xlsm`.

## DATA sheet updates

- Headers A1:G1 set to: Remarks, Latitude, Longitude, Distance, Course, Distance to go, Min Depth.
- Empty remarks cells in column A now use `=IF(Bn="","",CHAR(160))` so VBA row-count logic (column A based) tracks active route rows.
- Column F now uses cumulative formula `=IF(Dn="","",SUM($D$2:Dn))` for distance-to-go recalculation.
- Existing route rows where D looked like Course and E looked like Distance were normalized (D=Distance, E=Course).

## PLAN sheet template updates

- Header metadata added in template: Vessel=NORDLAND, Master=S.Merkov, Prepared=04.09.2026, Voyage No=09-26.
- Min Depth column formulas added in X12:X46: `=IF(DATA!Gx="","",DATA!Gx)`.

## Verification snapshot

- DATA max row: 600
- PLAN max row: 230
- Sample formulas: DATA!A3==IF(B3="","",CHAR(160)), DATA!F3==IF(D3="","",SUM($D$2:D3)), PLAN!X12==IF(DATA!G3="","",DATA!G3)
