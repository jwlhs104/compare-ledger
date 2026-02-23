class KarizawaHotelInvoiceReconcileLogic {
  // Aggregates parsed invoice entries keyed by check-in date.
  // All amounts for a stay are written to the check-in date only.
  reconcile(entries) {
    // result[checkInDateStr]["成人"] = { 房費, 入湯稅, 手数料, 其他, 總價 }
    const result = {};

    for (const entry of entries) {
      const { checkIn, roomFee, taxFee, otherFee, total, commission } = entry;
      const dateStr = formatDate(checkIn);

      if (!result[dateStr]) {
        result[dateStr] = {
          成人: { 房費: 0, 入湯稅: 0, 手数料: 0, 其他: 0, 總價: 0 },
        };
      }

      result[dateStr].成人.房費 += roomFee;
      result[dateStr].成人.入湯稅 += taxFee;
      result[dateStr].成人.手数料 += commission;
      result[dateStr].成人.其他 += otherFee;
      result[dateStr].成人.總價 += total;
    }

    return result;
  }
}
