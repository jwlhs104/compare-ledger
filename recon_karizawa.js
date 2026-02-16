class KarizawaHotelInvoiceReconcileLogic {
  // Aggregates parsed invoice entries into per-night totals keyed by date string.
  // Each entry spans multiple nights; amounts are distributed evenly per night.
  reconcile(entries) {
    // result[dateStr]["成人"] = { 房費, 入湯稅, 手数料, 其他, 總價 }
    const result = {};

    for (const entry of entries) {
      const { checkIn, nights, roomFee, taxFee, otherFee, total, commission } = entry;

      const nightlyRoomFee = roomFee / nights;
      const nightlyTax = taxFee / nights;
      const nightlyOther = otherFee / nights;
      const nightlyTotal = total / nights;
      const nightlyCommission = commission / nights;

      for (let n = 0; n < nights; n++) {
        const nightDate = new Date(checkIn);
        nightDate.setDate(checkIn.getDate() + n);
        const dateStr = formatDate(nightDate);

        if (!result[dateStr]) {
          result[dateStr] = {
            成人: { 房費: 0, 入湯稅: 0, 手数料: 0, 其他: 0, 總價: 0 },
          };
        }

        result[dateStr].成人.房費 += nightlyRoomFee;
        result[dateStr].成人.入湯稅 += nightlyTax;
        result[dateStr].成人.手数料 += nightlyCommission;
        result[dateStr].成人.其他 += nightlyOther;
        result[dateStr].成人.總價 += nightlyTotal;
      }
    }

    // Round to avoid floating-point artifacts
    for (const dateStr in result) {
      const row = result[dateStr].成人;
      for (const key in row) {
        row[key] = Math.round(row[key]);
      }
    }

    return result;
  }
}
