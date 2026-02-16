class KarizawaHotelInvoiceParser {
  constructor(seasonStartYear) {
    // seasonStartYear: e.g. 2025 for the 2025-26 season
    this.seasonStartYear = seasonStartYear;
  }

  parse(sheet) {
    // Data starts at row 6 in the sheet (index 5 in 0-based array)
    const data = sheet.getRange("A6:R1000").getValues();
    const entries = [];

    for (let i = 0; i < data.length; i++) {
      const row = data[i];
      const checkInRaw = row[3];   // D: 到着日
      const checkOutRaw = row[4];  // E: 出発日
      const name = row[5];         // F: 氏名/グループ名
      const roomFee = row[8];      // I: 宿泊代（R対象）
      const taxFee = row[9];       // J: 入湯税
      const otherFee = row[10];    // K: その他（Rなし）
      const total = row[11];       // L: 合計
      const commission = row[13];  // N: 手数料

      // Skip rows with no check-in date
      if (!checkInRaw && !checkOutRaw) continue;

      // Skip marker/memo rows
      const nameStr = typeof name === "string" ? name : "";
      if (
        nameStr.includes("この行使用済") ||
        nameStr.includes("リフト券代") ||
        (!nameStr && !checkInRaw)
      )
        continue;

      const checkIn = this.parseDate(checkInRaw);
      const checkOut = this.parseDate(checkOutRaw);

      if (!checkIn || !checkOut) continue;

      const nights =
        (checkOut.getTime() - checkIn.getTime()) / (1000 * 60 * 60 * 24);
      if (nights <= 0) continue;

      const roomFeeNum = this.parseAmount(roomFee);
      const taxFeeNum = this.parseAmount(taxFee);
      const otherFeeNum = this.parseAmount(otherFee);
      const totalNum = this.parseAmount(total);
      const commissionNum = this.parseAmount(commission);

      // Skip rows with all-zero amounts (likely empty rows)
      if (roomFeeNum === 0 && taxFeeNum === 0 && otherFeeNum === 0 && totalNum === 0) continue;

      entries.push({
        name: nameStr,
        checkIn,
        checkOut,
        nights,
        roomFee: roomFeeNum,
        taxFee: taxFeeNum,
        otherFee: otherFeeNum,
        total: totalNum,
        commission: commissionNum,
      });
    }

    return entries;
  }

  parseDate(value) {
    if (value instanceof Date && !isNaN(value.getTime())) return value;
    if (typeof value === "string" && value.trim()) {
      // Handle "12月6日" style strings
      const match = value.match(/(\d+)月(\d+)日/);
      if (match) {
        const month = parseInt(match[1]);
        const day = parseInt(match[2]);
        // Oct-Dec → seasonStartYear, Jan-Mar → seasonStartYear+1
        const year = month >= 10 ? this.seasonStartYear : this.seasonStartYear + 1;
        return new Date(year, month - 1, day);
      }
    }
    return null;
  }

  parseAmount(value) {
    if (typeof value === "number") return value;
    if (typeof value === "string") {
      // Remove whitespace, commas; handle parentheses as negative
      const cleaned = value
        .replace(/\s/g, "")
        .replace(/,/g, "")
        .replace(/^\((.+)\)$/, "-$1");
      return parseFloat(cleaned) || 0;
    }
    return 0;
  }
}
