function getColumnIndex(columnLabel) {
  columnLabel = columnLabel.toUpperCase();
  let index = 0;
  let multiplier = 1;

  for (let i = columnLabel.length - 1; i >= 0; i--) {
    const charCode = columnLabel.charCodeAt(i) - 65;
    index += (charCode + 1) * multiplier;
    multiplier *= 26;
  }

  return index - 1;
}
function numberToColumnLetter(n) {
  let letter = '';
  while (n > 0) {
    let remainder = (n - 1) % 26;
    letter = String.fromCharCode(65 + remainder) + letter;
    n = Math.floor((n - 1) / 26);
  }
  return letter;
}
function notNull(string) {
  return string && string !== "" && string !== "X" && string !== "Ｘ";
}

function transpose(matrix) {
  if (!Array.isArray(matrix) || matrix.length === 0) return [];
  return matrix[0].map((_, colIndex) => matrix.map(row => row[colIndex]));
}

function formatDate(date) {
  date = handleDate(date)
  return Utilities.formatDate(date, Session.getScriptTimeZone(), "yyyy-MM-dd");
}

function handleDate(date) {
  if (date instanceof Date) return date;
  if (typeof date === "string") {
    if (date.split("/")[0].length === 2) {
      date = "20" + date;
    }
    date = new Date(date);
    return date;
  }
  const startDate = new Date("1899-12-30");
  date = new Date(date * (1000 * 60 * 60 * 24) + startDate.valueOf());
  return date;
}

function dateDifferenceInDays(date1, date2) {
  // Convert inputs to Date objects if they're not already
  const d1 = new Date(date1);
  const d2 = new Date(date2);

  // Get the time difference in milliseconds
  const diffTime = Math.abs(d2 - d1);

  // Convert milliseconds to days
  const diffDays = Math.ceil(diffTime / (1000 * 60 * 60 * 24));

  return diffDays;
}
