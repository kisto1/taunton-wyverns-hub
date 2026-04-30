/**
 * onEdit: when a checkbox is checked (TRUE), replace the block of pasted names
 * directly under the checkbox with the current names from column A that match "Yes"/TRUE
 *
 * Behavior:
 * - Only reacts to the edited cell becoming TRUE (checked).
 * - Clears the contiguous block of cells below the checkbox in the same column,
 *   then writes the fresh list starting on the row immediately below the checkbox.
 * - If no matches, it will still clear the previous pasted block (so the area is empty).
 *
 * Adjust dataStartRow if your names start on a different row. (In your sheet it's row 7.)
 */
function onEdit(e) {
  try {
    const sh = e.range.getSheet();
    const r = e.range;
    const editedRow = r.getRow();
    const editedCol = r.getColumn();
    const newValue = e.value; // for simple onEdit events, checkboxes show "TRUE" (string) or boolean true

    // Only respond when the cell was changed to TRUE (a checked checkbox)
    if (newValue !== 'TRUE' && newValue !== true) return;

    // ---------- CONFIG ----------
    // Row where the names (column A) start in your sheet. Change if different.
    const dataStartRow = 7;
    // ----------------------------

    // Ensure there's at least some data area
    const lastRow = Math.max(sh.getLastRow(), dataStartRow);
    if (lastRow < dataStartRow) return;

    // Read names (col A) and corresponding responses in the edited column
    const numRows = lastRow - dataStartRow + 1;
    const names = sh.getRange(dataStartRow, 1, numRows, 1).getValues();
    const responses = sh.getRange(dataStartRow, editedCol, numRows, 1).getValues();

    // Build the list of names to write (matches "Yes" or TRUE)
    const out = [];
    for (let i = 0; i < responses.length; i++) {
      const resp = responses[i][0];
      if ((typeof resp === 'string' && resp.trim().toLowerCase() === 'yes') ||
          resp === true ||
          String(resp).toLowerCase() === 'true') {
        const nm = names[i][0];
        if (nm !== "" && nm !== null) out.push([String(nm)]);
      }
    }

    // Destination start row: one row below the checkbox
    const pasteRow = editedRow + 1;

    // Clear existing contiguous block under the checkbox (so re-checking replaces it)
    // Find the first blank cell starting at pasteRow
    let scanRow = pasteRow;
    const sheetLastRow = sh.getLastRow();
    while (scanRow <= sheetLastRow) {
      const val = sh.getRange(scanRow, editedCol).getValue();
      if (val === "" || val === null) break;
      scanRow++;
    }
    const blockLen = Math.max(0, scanRow - pasteRow);
    if (blockLen > 0) {
      sh.getRange(pasteRow, editedCol, blockLen, 1).clearContent();
    }

    // If there's something to paste, write it starting at pasteRow
    if (out.length > 0) {
      sh.getRange(pasteRow, editedCol, out.length, 1).setValues(out);
    }

  } catch (err) {
    // For a simple onEdit trigger we don't surface logs to the user.
    // If you want to debug, remove the try/catch or use Logger.log(err);
  }
}
