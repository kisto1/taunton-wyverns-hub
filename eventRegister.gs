function onEdit(e) {
  const ss = e.source;

  const result = {
    successCount: 0,
    hasPeople: false,
    invalidFee: false,
    error: null
  };

  try {
    const sh = e.range.getSheet();
    const editedCol = e.range.getColumn();
    const newValue = e.value;

    if (sh.getName() !== 'EVENT REGISTER') return;
    if (newValue !== 'TRUE' && newValue !== true) return;

    const dataStartRow = 7;
    const ts = ss.getSheetByName('Transactions');
    if (!ts) throw new Error('Transactions sheet not found');

    const lastRow = Math.max(sh.getLastRow(), dataStartRow);
    const numRows = lastRow - dataStartRow + 1;

    const names = sh.getRange(dataStartRow, 1, numRows, 1).getValues();
    const responses = sh.getRange(dataStartRow, editedCol, numRows, 1).getValues();

    const date = sh.getRange(2, editedCol).getValue();
    const description = `${sh.getRange(1, editedCol).getValue()} ${sh.getRange(4, editedCol).getValue()}`;

    const feeValue = sh.getRange(5, editedCol).getValue();
    const feeNumberValue = parseMoney(feeValue);

    const transactionValue =
      typeof feeNumberValue === 'number' && !isNaN(feeNumberValue)
        ? feeNumberValue * -1
        : '';

    if (transactionValue === '') result.invalidFee = true;

    // -------- DELETE EXISTING TRANSACTIONS --------
    const tsData = ts.getDataRange().getValues();

    const filtered = tsData.filter((row, i) => {
      if (i === 0) return true;

      const rowDate = row[0];
      const rowDesc = row[3];

      return !(isSameDate(rowDate, date) && rowDesc === description);
    });

    ts.clearContents();
    ts.getRange(1, 1, filtered.length, filtered[0].length).setValues(filtered);

    // -------- BUILD NEW TRANSACTIONS --------
    const out = [];

    for (let i = 0; i < responses.length; i++) {
      const resp = responses[i][0];

      const isYes =
        (typeof resp === 'string' && resp.trim().toLowerCase() === 'yes') ||
        resp === true ||
        String(resp).toLowerCase() === 'true';

      if (isYes) {
        result.hasPeople = true;

        const person = names[i][0];
        if (!person) continue;

        out.push([
          date,
          "Non-cash",
          "-",
          description,
          transactionValue,
          "Balance transaction - session",
          person
        ]);

        result.successCount++;
      }
    }

    // IMPORTANT: if no people, we still want deletion only
    if (out.length > 0) {
      ts.getRange(ts.getLastRow() + 1, 1, out.length, out[0].length).setValues(out);
    }

    // -------- TOAST MESSAGES --------
    if (!result.hasPeople) {
      showToast(
        ss,
        'There were no people for this event. Ensure the people have been registered for this event correctly.'
      );
      return;
    }

    if (result.successCount > 0 && result.invalidFee) {
      showToast(
        ss,
        'Transactions updated successfully, but the fee could not be extracted. Ask the treasurer to update the transaction values manually.'
      );
      return;
    }

    showToast(ss, 'Transactions updated successfully');

  } catch (err) {
    result.error = err.message || String(err);

    showToast(ss, `An error occurred - ${result.error}`);
    Logger.log(err);
  }
}

function showToast(ss, msg, title = 'Transactions') {
  ss.toast(msg, title, 6); 
}

function parseMoney(v) {
  if (typeof v === 'number' && !isNaN(v)) return v;

  if (typeof v === 'string') {
    const trimmed = v.trim();
    const match = trimmed.match(/^£?\s*(\d+(\.\d+)?)$/);

    if (match) {
      const num = Number(match[1]);
      return isNaN(num) ? '' : num;
    }
  }

  return '';
}

// Helper: compares dates ignoring time component
function isSameDate(d1, d2) {
  if (!(d1 instanceof Date) || !(d2 instanceof Date)) return false;

  return d1.getFullYear() === d2.getFullYear() &&
         d1.getMonth() === d2.getMonth() &&
         d1.getDate() === d2.getDate();
}
