# eventRegister.gs — Code Review

## Issues

**1. Checkbox detection mismatch (line 18)**

`e.value` for a checkbox returns the string `"TRUE"` or `"FALSE"`. The check on line 18 handles both `'TRUE'` and `true`, but the `false`/unchecking case is not handled — unchecking a box triggers `onEdit` but returns early cleanly, which is fine. However, this also means **any cell edited to the literal string `"TRUE"`** in any column would pass this check. Low risk given the sheet structure, but worth noting.

**2. Column match on date+description is fragile (lines 49–52)**

If two different events share the same date and description (e.g., two sessions on the same day with the same name), the delete step will wipe both. The description is `"${row 1 value} ${row 4 value}"` — if row 4 is blank or generic, collisions become more likely.

**3. `ts.clearContents()` then re-write is risky (lines 55–56)**

If `setValues` throws after `clearContents()`, you've lost all transaction data with no recovery path. The catch block will show a toast but data is already gone.

**4. `filtered[0].length` could throw if `filtered` is empty (line 56)**

If the Transactions sheet is completely empty (header only and all rows matched), `filtered` would contain just the header row — but if somehow even that's gone, `filtered[0]` would throw. Low risk but unguarded.

**5. No guard against editing a header/metadata row (lines 27–28)**

If someone edits a checkbox in rows 1–6 (event metadata rows), `editedCol` is valid but `responses` would be read from rows 7 onwards against column metadata — probably not a real checkbox, but no explicit row guard exists.

---

**6. MPK addition - we need something that specifies whether an event uses member credit or not, and only add a transaction if it's member credit.

## Improvement Ideas

**A. Use `deleteRow` instead of clear+rewrite**

Rather than `clearContents` + `setValues`, filter in-place: find the rows to delete by row index and delete them with `deleteRow()` (iterating bottom-up). Much safer — no risk of data loss on error.

**B. Confirm the checkbox is in row 6 (the "attended?" row)**

Add a row check: `if (e.range.getRow() !== 6) return;` (or whatever row the checkboxes live in). This makes the trigger far more precise and prevents accidental fires.

**C. Improve description uniqueness**

Consider including the column index or a unique event ID in the description used for matching deletions, so same-day same-name events don't collide.

**D. Wrap the delete+insert with a rollback snapshot**

Before `clearContents`, capture `tsData` as a backup. If the `setValues` fails, write `tsData` back. This protects against data loss in the current clear+rewrite approach.

**E. Toast for "unchecked" / no-op case**

Currently unchecking a box silently does nothing. A brief toast like "Registration cleared" when a checkbox is unchecked (and rows were deleted) could be helpful for users.

**F. `parseMoney` doesn't handle zero**

If a fee is `£0` (free event), `parseMoney` returns `0`, then `transactionValue` becomes `-0`. Not harmful but slightly odd. A free event probably shouldn't add a transaction row at all, or should add one with value `0`.
