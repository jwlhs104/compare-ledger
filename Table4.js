function getAdvancePayments() {
  const ledgers = [
    {
      resort: '苗王',
      table4Id: '1WbArDTJhPtLzyS9DBBYQKrnrhzrQUtzmsnePUF9UVt0'
    },
    {
      resort: '雫石',
      table4Id: '1zLXq6-8HuLa8r8W5MwnPDO9_n0g8g013OO2xoMXOTfA'
    },
        {
      resort: '萬座',
      table4Id: '1UlbiimWii0Cp24rXG5Q5ur8tatb5sNr665qJsxNGVYg'
    },
  ]
  ledgers.forEach(ledger => getAdvancePayment(ledger))
  function getAdvancePayment (ledger) {
    const sheet = SpreadsheetApp.openById(ledger.table4Id).getSheetByName('日記帳');
    const data = sheet.getRange("A10:EW3000").getValues();
    
    const startCol = getColumnIndex("DB");
    const maxDays = 6;
    const interval = 8;

    const hotelOffset = 5; // 6th column in each 8-column block
    const advancePaymentCol = getColumnIndex("AU");
    const customerRowGroupCol = getColumnIndex("CA"); // 1: first row, 2: second row, etc.
    const ageCol = getColumnIndex("CU");
    const bookingMap = {};

    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      // Find the first hotel booking for this customer
      let firstBookingDate = null;
      for (let day = 0; day < maxDays; day++) {
        const baseCol = startCol + (day * interval);
        const hotelCol = baseCol + hotelOffset;
        const dateCol = baseCol;

        const hotel = row[hotelCol];
        const date = row[dateCol];

        if (hotel && notNull(date) && !firstBookingDate) {
          firstBookingDate = new Date(date);
          break;
        }
      }
      if (!firstBookingDate) continue;
      const customerRowGroup = Number(row[customerRowGroupCol]) || 1;
      const offsetDays = (customerRowGroup - 1) * 6;
      const trueStartDate = new Date(firstBookingDate);
      trueStartDate.setDate(firstBookingDate.getDate() - offsetDays);
      const dateStr = trueStartDate.toISOString().split('T')[0]; // "YYYY-MM-DD"
        
      // Determine age group
      const ageGroup = row[ageCol] === "成人" ? "成人" : "孩童";

      // Sum advance payment
      const advancePayment = Number(row[advancePaymentCol]) || 0;
      if (!bookingMap[dateStr]) {
        bookingMap[dateStr] = { "成人": 0, "孩童": 0 };
      }
      bookingMap[dateStr][ageGroup] += advancePayment;
    }

    const targetSheet = SpreadsheetApp.getActive().getSheetByName(ledger.resort)
    const targetValues = targetSheet.getRange("A3:B").getValues();
    const writeCol = 3;
    targetValues.forEach((row, index) => {
      if (!notNull(row[0])) return
      const date = new Date(row[0]).toISOString().split('T')[0];
      const ageGroup = row[1]
      if (bookingMap[date] && bookingMap[date][ageGroup]) {
        targetSheet.getRange(3+index, writeCol).setValue(bookingMap[date][ageGroup])
      }
    })

  }
}
function getColumnIndex(columnLabel) {
  // Convert the column label to uppercase for consistency
  columnLabel = columnLabel.toUpperCase();

  let index = 0;
  let multiplier = 1;

  for (let i = columnLabel.length - 1; i >= 0; i--) {
    const charCode = columnLabel.charCodeAt(i) - 65; // ASCII code of 'A' is 65
    index += (charCode + 1) * multiplier; // Add 1 to charCode to account for 'A' being 1, 'B' being 2, and so on
    multiplier *= 26;
  }

  return index - 1; // Adjust for 0-based index (optional, depends on your use case)
}
function notNull(string) {
  return string && string !== "" && string !== "X" && string !== "Ｘ" && string !=="#VALUE!"
}