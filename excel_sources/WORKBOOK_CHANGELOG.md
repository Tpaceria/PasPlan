# Workbook change log

Updated file: `NEW_FINAL_LOGO_EXACT_5_05x1_00cm.xlsm` and copy `PasPlan_NORDLAND.xlsm`.

## DATA sheet updates

- Headers A1:G1 set to: Remarks, Latitude, Longitude, Distance, Course, Distance to go, Min Depth.
- Empty remarks cells in column A now use `=IF(Bn="","",CHAR(160))` for compatibility with the embedded workbook macro.
- Column F now uses cumulative formula `=IF(Dn="","",SUM($D$2:Dn))` for distance-to-go recalculation.
- Existing route rows where D looked like Course and E looked like Distance were normalized (D=Distance, E=Course).

## PLAN sheet template updates

- Header metadata added in template: Vessel=NORDLAND, Master=S.Merkov, Prepared=04.09.2026, Voyage No=09-26.
- Min Depth column formulas added in X12:X46: `=IF(DATA!Gx="","",DATA!Gx)`.

## Verification snapshot

- DATA max row: 600
- PLAN max row: 230
- Sample formulas: DATA!A3==IF(B3="","",CHAR(160)), DATA!F3==IF(D3="","",SUM($D$2:D3)), PLAN!X12==IF(DATA!G3="","",DATA!G3)

## Extracted VBA review notes

- `Лист2_cls.txt` reflects the event trigger range `A:G` so Min Depth edits also trigger refresh.
- `PassagePlanModule_bas.txt` includes `DATA!G -> PLAN!X` mapping and clears `X` in `ClearMappedCells`.
- `PassagePlanModule_bas.txt` message boxes use readable English strings (`Done/Error`) to avoid mojibake.
- `PassagePlanModule_bas.txt` row-counting logic is updated to use route column `B` and normalize `CHAR(160)` as empty.
- `PassagePlanModule_bas.txt` transfer formulas are wrapped with `IFERROR` to prevent DATA formula errors from surfacing in PLAN output rows.
- `Лист2_cls.txt` limits auto-refresh scope to only populated route rows and skips very large edits for better data-entry performance.
- `PassagePlanModule_bas.txt` recalculates only PLAN scope (`wsP.Calculate`) instead of full Excel recalculation, and restores `EnableEvents` to its prior state after execution.
- `Лист2_cls.txt` invokes the refresh macro via `Application.Run` scoped to `ThisWorkbook` to avoid cross-workbook dispatch ambiguity.
- `Лист2_cls.txt` escapes apostrophes in workbook names before building the `Application.Run` target.
- `PassagePlanModule_bas.txt` derives print-area last column from template print settings instead of hard-coding `BN`.
