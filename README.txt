NOVEMBER 2 UPDATE — HORIZONTAL / ELEVATION HIGHLIGHTING
  Northing/Easting outside +/-0.250 ft highlight all protected report data
  on the pile row; elevation outside +/-0.250 ft highlights only Delta Elevation.
  Drill Date, Rig Number and Comments stay light green and editable.
  Comments stay empty unless the user entered a note.
  See SOUTHERN_SURVEY_CONDITIONAL_FORMAT_UPDATE.txt for details.

NOVEMBER 2, 2026 UPDATE — NEW STANDARDS + PER-AXIS TOLERANCE
==========================================================
- New multi-sheet standards spreadsheets may contain conflicting copies of
  one pile's coordinates. PIVOT now shows a conflict warning and requires
  confirmation to prioritize the main visible All Areas design sheet.
  Confirm all disputed design coordinates with the survey/design team.
- Printable Excel Sheet 1: Design, As-Built, individual N/E/Z deviations,
  and editable Comments. Radial Offset is no longer on this sheet.
- Match Verification Sheet 2 now contains Radial Offset.
- The entire Sheet 1 row highlights red if ABS(Delta Northing),
  ABS(Delta Easting), OR ABS(Delta Elevation) exceeds 0.250 ft (3 in).
  The exact endpoints -0.250 and +0.250 remain in tolerance.
- Single-shot status screening now uses the same per-axis N/E/Z thresholds.
- Drill Date / Rig Number / Comments remain editable on a protected report.

IMPORTANT UPDATE: NO MORE SILENTLY MISSING PILES
==============================================
This build accounts for each recognized pile ID from imported survey CSVs.
Incomplete center groups, two-shot groups, and incomplete C/HS pairs appear
as PENDING REVIEW with no fabricated as-built coordinates or deviations.
Import the prior report AND all original survey CSV files for a complete
running-total recovery. See COVERAGE_FIX_README.txt for detailed instructions.

PIVOT | A Faulk Product
=======================

OCTOBER 29 UPDATE — SINGLE SHOT TOLERANCE + SIMPLE C/HS PROMPT
- For an isolated pile-coded shot, Match Points now matches it directly against
  design with NO extra single-shot confirmation window. If its radial offset
  is <= 0.250 feet AND its elevation difference is between -0.250 and +0.250
  feet (inclusive), the result is marked SINGLE SHOT - ASSUMED CENTER.
- An assumed center keeps the actual measured N/E/Z; it is NOT a computed
  polygon centroid and does NOT silently replace survey coordinates with
  standard/design coordinates. Piles outside tolerance or with missing
  elevation / uncertain identification are flagged for review.
- A complete C + HS/H pair triggers ONE simple Yes/No question at Match Points:
  Yes uses the HS/H elevation; No keeps the C-shot elevation. Both always use
  the C shot's Northing and Easting. No separate per-pair dialogs appear.
- Importing a CSV does not show a redundant tagged-shot information popup;
  the regular shot-date confirmation stays in place.
- Re-clicking Match Points does not reopen the C/HS question for already matched
  pairs. Original raw survey observations are preserved.

OCTOBER 29 FIX — SINGLE-SHOT AS-BUILT FILES
- The center screen now shows ONLY groups with 3 or more distinct perimeter locations by default.
  One- and two-shot piles no longer fill the table with invalid center calculations.
- If you need to inspect a one- or two-shot group, turn on 'Show incomplete groups (1–2 shots)'.
- CSVs imported from the center window use the SAME importer as the main PIVOT window.
  Single pile-coded points are retained for Match Points; they do not require a centroid.
- Click Match Points on the main window to review/match single-shot observations.
  Two-shot pile groups are not automatically converted to centers; obtain further
  perimeter shots or review their original measurements before using them.
- The standalone program still requires no Python on recipient computers.

BUILD THE STANDALONE WINDOWS APPLICATION
1. Extract this folder somewhere outside OneDrive, such as C:\PIVOT.
2. Double-click BUILD_STANDALONE_EXE.bat on a Windows PC that has Python.
3. The standalone dist\PIVOT.exe does NOT require Python on recipient computers.

USING PIVOT -- MAIN BUTTON ORDER
1. Load Standards: load design standards (.xlsx).
2. Import CSV: load a dated field survey such as 260170_102126_DM.csv.
   The program recognizes the shot date, including pile-coded single shots.
   Importing a CSV does NOT force-open the center workspace.
3. Calculate Centers (OPTIONAL): use this only for pile groups with at least
   three perimeter shots. A single exterior shot cannot define a polygon
   center, and a two-shot group is also insufficient. You can skip this step.
4. Match Points: match ordinary as-built points after importing the CSV.
   For complete C + HS/H pairs, one PRE-MATCH popup asks whether to use the
   high-spot elevation (YES) or keep the center-shot elevation (NO).
   The C shot ALWAYS provides Northing and Easting. No pair-elevation popup
   appears after successful matching. A changed selection replaces the prior
   derived pair record rather than duplicating it in history.
   Single pile-coded shots now match directly, without another popup.
   Each of Northing, Easting and elevation deviation within +/-0.250 ft
   gives SINGLE SHOT - ASSUMED CENTER; other measurements are flagged.
   The measured coordinates and source observations are preserved.
   Calculated-and-accepted centers enter the main PIVOT table directly.
5. Import Previous Report: bring forward earlier running-total records from
   a previously exported PIVOT Excel report (.xlsx). It can be done before
   or after the current survey.
6. Export Excel: create the Southern Survey running-total Excel report with
   two visible sheets, regardless of the current on-screen filter:
       Sheet 1: Pile Deviation Report (printable; Comments are editable)
       Sheet 2: Match Verification (includes Date Shot and Source Job)
   It also contains hidden machine-readable history for later reimports.

Use Show: Current piles / All shot history / Most recent import and Date Shot
to filter the main PIVOT table. Current piles is the default and displays the
latest dated survey for each pile ID. Previous shots are preserved in the hidden
history for running totals but do not create duplicate rows in the printable report. Before Excel export, unresolved center candidates trigger a
reminder that you can review, ignore, or explicitly bypass.

Previous PIVOT reports generated before dates were tracked can be imported, but
those records are marked Unknown (legacy) because the original reports did not
record the date the survey was shot.

The report workbook is worksheet-protected (password PIVOT2026) to prevent
accidental editing. This is not cryptographic protection.

A source-code comment easter egg is intentionally preserved.

SOUTHERN SURVEY EXCEL REPORT UPDATE
- Sheet 1 includes a Comments column after Radial Offset. Comments cells are editable.
- Report calculations, layout and branding remain protected.
- Filter drop-downs on row 6 enable filtering by Pile # prefix and other values.
- Excel generally does not permit sorting locked cells in a protected worksheet,
  even when sorting is allowed by the protection options. Use PIVOT's Pile ID
  sorting before exporting if you need a different row order.
- Save the annotated XLSX and import it via Import Previous Report to preserve
  its comments in the next cumulative report.


AUTOMATIC PILE CENTER CALCULATION (NEW)
- Click Calculate Pile Centers to open the separate Carlson-style center workspace.
- Import one or several CSVs containing perimeter observations such as B053.
  You can also use the main Import Shot button: perimeter-only CSVs route here
  instead of being incorrectly treated as a single as-built center shot.
- Pile descriptions default to regex ^[A-Z]{1,4}\d{1,6}$. Edit the pattern
  in the center window for different survey coding conventions.
- The default method is closed polygon (shoelace area-weighted centroid).
- View the actual perimeter and point labels in plan view with equal axis scales.
- Review warnings, reorder a vertex with Up/Down, exclude or restore a point.
- Optional Best-Fit Circle provides radius and RMS residual for comparisons.
- Accept each result: its calculated center is immediately added as a DERIVED
  row to the main PIVOT deviation table (with field point CENTER-B053, etc.).
  The center is paired with the EXACT corresponding pile ID in loaded standards.
  If standards have not yet been loaded, the center row remains visible with
  CENTER - STANDARD NOT FOUND until standards are loaded.
- Use Return to Main PIVOT to view accepted centers and export the usual
  Southern Survey running-total workbook alongside ordinary field shots.
- When accepting a center, PIVOT asks which exterior-shot elevation to use:
  Average, Highest (Southern Survey preference), Lowest, or Leave blank.
  Only included perimeter observations are used for the chosen height.
  The chosen height is a derived value, not a direct observation at the center.
  Both the main deviation report and the default accepted-centers CSV export
  preserve each pile's own elevation selection.
- You can still export accepted centers as Point,N,E,Z,Description CSV separately.
- If perimeter edits invalidate a center's acceptance, its current-session
  derived row is removed until the revised geometry is reviewed and reaccepted.
- Export Calculation Audit (.csv) for original shots, polygon order, included
  observations, fitted circle and warnings. Original survey files are never edited.
- Southern Survey Excel formatting remains unchanged, and the center result
  can now appear in its running-total deviation and Match Verification sheets.
- Run TEST_PILE_CENTERS.bat on the build computer to run automated geometry tests.
- The geometry module does not require additional third-party packages.
- PIVOT's original deviation matching and Southern Survey reports are retained.

QUALITY CONTROL: A centroid is not a survey-certified center. Partial circumferences,
interior shots, spatial clusters and badly shaped boundaries require manual review.

CENTER WORKSPACE LAYOUT UPDATE
- Accept Selected Center is now the first and largest button in the Perimeter Review control area.
- Review controls and CSV export controls remain visible at the bottom while the plan view resizes.
- The window opens within the available screen size to accommodate smaller displays.
- Export settings and export buttons occupy separate rows to prevent clipping.
- No calculation, deviation report, or Southern Survey branding changes were made.

PRE-EXPORT PILE CENTER CHECK (NEW)
- Before Export Running Total, PIVOT checks the perimeter observations imported
  during the current session for repeated pile descriptions (2 or more shots).
- If any such pile has not had its center accepted and added to the main table,
  the Pile Center Review dialog appears before the Save As window.
- The review list shows Pile ID, perimeter shot count, source-file count and
  status: Needs Review / Accepted / Ignored.
- Calculate / Review Centers opens the existing polygon/circle center workspace.
  Review the perimeter, select an elevation, then Accept Selected Center to add
  the derived center to PIVOT's main running-total table.
- Trash / Ignore Selected marks unneeded pile groups as Ignored for THIS SESSION.
  It does not remove or edit the original CSV survey observations or historical
  records imported from previous reports. If a center was accepted during this
  session, ignoring it removes that derived center from the current main table.
- Restore Ignored reverses the decision. You can open the list any time from
  the application's Pile Centers > Review Center Queue menu.
- If new source observations arrive for an ignored pile, PIVOT automatically
  returns that pile to Needs Review, so new field evidence cannot be missed.
- Export Anyway explicitly continues without calculating outstanding centers.
  Cancel Export closes the reminder without writing a report.
- Single pile-labeled observations are not considered multi-shot perimeter
  groups and do not trigger the reminder on their own.
- The check only knows about raw perimeter CSV observations imported into the
  current PIVOT session; old Excel reports do not contain perimeter-shot history.
- All existing Southern Survey report branding, worksheet protection, Comments,
  filters, running-total data and highest/average/lowest elevation choices remain.

CENTER WORKSPACE -- UPDATED LAYOUT AND TRASH CONTROLS
- The Calculate Centers button now comes before Match Points on the main toolbar.
- The center calculator opens as a standard resizable desktop window with the
  native Windows Maximize control and an explicit Maximize / Restore button.
- The top toolbar imports perimeter CSV(s), calculates all pile centers and
  sets the pile-description recognition pattern.
- The left workspace provides two tabs: Pile Results and Perimeter Shots / Edit.
  Select the pile in Pile Results, then use Perimeter Shots / Edit if needed.
  The actual measured polygon and computed center remain visible in the plan
  view on the right, alongside the selected method and quality warnings.
- Accept Center, Trash / Ignore and Restore Pile are persistent buttons above
  the tabs. These actions apply to the SELECTED PILE ID, not a single vertex.
- Trash / Ignore removes an unnecessary pile from current-session reminders.
  If its center was already accepted in this session, that derived main-table
  row is removed. It does not delete original survey observations or older
  imported reports. Restore Pile brings that group back for review.
- Exclude / Include Shot in the editor is DIFFERENT: it affects only the
  selected observation's participation in the center calculation. It does
  not trash the pile or delete the survey observation.
- Ignored groups automatically return to Needs Review if additional perimeter
  observations for that pile are imported later.
- Approved elevation selection (Highest / Average / Lowest / Blank), separate
  center-point CSV export, detailed audit and Southern Survey Excel formatting
  are unchanged.

SINGLE-SHOT PILE CENTER IGNORE FIX
- A pile-coded survey observation with only one or two perimeter shots is not
  enough for a valid polygon center, but can now be selected and ignored/trash-marked
  from the Pile Center Calculator and from the pre-export review queue.
- Example: 260170_102826.csv point DC_4 has pile description B067 and only
  one observation; select pile B067 and click Trash / Ignore. It can be restored.
- The pre-export review list also shows the source field point number(s).
- Non-pile check point descriptions such as DC_3 / CHK are ignored by the center
  calculator automatically; no trash action is required for those.
- Trash / Ignore never deletes the original CSV point or any prior report record.

SOUTHERN SURVEY RADIAL-OFFSET ROW HIGHLIGHTING (CURRENT)
- On Sheet 1 (Pile Deviation Report), the ONLY field that triggers red
  highlighting is Radial Offset in column R.
- If a pile's numeric radial offset is > 0.250 (or < -0.250), the entire
  data row A:S turns light red with dark red text, including the Comments cell.
- Radial offsets of 0.250 or less remain normal. A blank radial offset
  stays normal. Northing, Easting and Elevation deviations do NOT independently
  trigger red highlighting, even if they exceed 0.250.
- These are live Excel conditional-formatting rules. The comment column
  remains editable, the other report cells remain protected, and filtering,
  running totals, logo/branding, and the Match Verification sheet are retained.


SURVEYED CENTER + HIGH-SPOT PAIRING (SOUTHERN SURVEY)
- Some CSV files include two separate field observations for a pile:
    A218_C (or A218C): directly surveyed center Northing and Easting
    A218_HS / A218_H (or A218HS): designated high-spot Elevation
- The pairing is automatic by exact pile ID + shot date, NOT field point number.
  Center and high-spot rows may have different field point numbers.
- Import your dated CSV with button 2 and click 4 Match Points. Button 3
  Calculate Centers is only necessary for perimeter polygon observations.
- PIVOT will create ONE derived row per complete C + HS/H pair. Its field point
  number is the center shot number, with both source point numbers recorded in
  Match Verification. The center's own elevation is not used for this row.
- A C or HS observation without its mate is flagged; PIVOT never treats the
  separate exterior/high-spot point as the pile's horizontal center.
- If multiple competing C or HS shots exist for the same pile on one date,
  PIVOT flags the ambiguous pile rather than silently selecting one.
- C/HS pairs from separate files are allowed when pile ID and shot date match.
  Re-importing the same raw observations does not duplicate derived results.
- If the designated high-spot elevation is below the center-shot elevation,
  PIVOT displays an explicit review warning and STILL uses the labeled HS/H
  elevation, rather than silently choosing the higher numeric value.
- A direct C/HS center uses an EXACT pile-ID lookup in the standards. The app
  never assigns a nearby but different design pile to an explicitly labeled C.
- Data and source provenance survive Southern Survey Excel export/re-import.
  The original CSVs and the perimeter-centroid workflow are unchanged.
- Run TEST_PILE_CENTERS.bat to run all geometry and C/HS pairing tests.


SOUTHERN SURVEY LATEST-PILE + ELEVATION TOLERANCE UPDATE
- Entire printed pile row turns red when Radial Offset exceeds 0.250 ft OR
  absolute Elevation Deviation exceeds 0.250 ft. Exact +/-0.250 remains normal.
- Repeated pile IDs from an imported previous report and newer CSV now show
  one latest result per pile by default and in the printable Excel sheets.
- Original shots (and their dates/notes) remain in the hidden report history.
- Switch Show to All shot history to review old shots in the main workspace.
- The active report's editable Comments link back to the right stored record.

UPDATE — OCTOBER 29, 2026: OPTIONAL CENTER WORKFLOW
- Pile codes with only one shot no longer trigger/require a center calculation.
- Multiple exterior shots still support polygon-center review, but the window
  does not open automatically during CSV import.
- Two-shot groups appear in the optional queue as insufficient, not overdue.
- Single-shot locations are matched only after confirmation and are clearly
  marked for center verification; they are never silently labeled a center.
- Excel export still prompts for any genuine unresolved multi-shot centers
  but lets the user choose Export Anyway without data loss.
- Southern Survey export formatting, history deduplication, C/HS pairing,
  protection, Comments, filters and conditional formatting are unchanged.

SOUTHERN SURVEY — DRILL DATE AND RIG NUMBER UPDATE (2026-10-29)
--------------------------------------------------------------
The printable Pile Deviation Report now shows, under AS-BUILT LOCATION:
    Pile # | Drill Date | Rig Number | Northing | Easting | Elevation | Date Shot

Drill Date (column G), Rig Number (column H) and Comments (column U)
are editable directly in Excel without unprotecting the worksheet.
All design/as-built measurements, deviation calculations, and branding remain
locked with the existing worksheet protection settings.

The drill date is not inferred from the Date Shot. Southern Survey users fill
in Drill Date and Rig Number when that construction information is available.
After editing, save the Excel workbook, then import it using PIVOT's
"Import Previous Report" control to keep these values in the running total.
The new metadata is also stored in the hidden _PIVOT_HISTORY sheet. Older
PIVOT reports without these fields remain compatible.
