class NaebaInvoiceParser extends BaseInvoiceParser {
  parse(sheet) {
    const data = sheet.getRange("A13:FG3000").getValues();
    const rooms = [];
    const roomIndexs = {};
    let firstDate = null;
    // 行程日同天(皆為第一天)且日期相同參考分房代碼決定入住人數
    for (let i = 0; i < data.length; i++) {
      const row = data[i];
      const continuation = row[getColumnIndex("CA")];
      const bonusMeal = row[getColumnIndex("FG")];
      const orderNumber = row[getColumnIndex("D")];
      const auPrice = row[getColumnIndex("AU")];

      let roomRecord = null;
      firstDate = continuation === 1 ? null : firstDate;

      for (let day = 0; day < 6; day++) {
        const baseCol = getColumnIndex("DB") + day * 8;
        const date = row[baseCol];
        const roomType = row[baseCol + 5];
        const roomNumber = row[baseCol + 6];
        const mealType = row[baseCol + 7];
        const formattedDate = date instanceof Date ? formatDate(date) : null;

        if (notNull(roomType)) {
          firstDate = firstDate === null ? formattedDate : firstDate;
          const key = [
            formattedDate,
            day,
            roomType,
            roomNumber,
            continuation,
          ].join("/");
          if (key in roomIndexs) {
            roomRecord = rooms[roomIndexs[key]];
          } else {
            roomRecord = new RoomRecord(
              roomNumber,
              formattedDate,
              roomType,
              firstDate
            );
            roomIndexs[key] = rooms.length;
            rooms.push(roomRecord);
          }
          const ageType = this.detectAge(row);
          const personRecord = new PersonRecord(
            row[4],
            mealType,
            ageType,
            bonusMeal,
            orderNumber,
            auPrice,
            continuation
          );
          roomRecord.addPerson(personRecord);
        } else {
          firstDate = formattedDate;
        }
      }
    }
    return rooms;
  }

  detectAge(row) {
    const rawAge = row[getColumnIndex("BW")];
    if (rawAge === "成人") return rawAge;
    return "孩童";
  }
}
