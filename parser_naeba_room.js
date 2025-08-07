class NaebaRoomInvoiceParser extends BaseInvoiceParser {
  parse(sheet) {
    const sizeMap = this.createRoomSize(sheet, 10);
    const roomRanges = this.createRoomRanges(sheet, 10);
    const data = sheet.getRange("E7:BCZ997").getValues();
    const rooms = [];
    const roomIndexes = {};
    for (let i = 0; i < data[0].length; i += 12) {
      const firstDate = formatDate(data[0][i + 4]);
      let roomRecord = null;
      let roomNumber = 1;
      for (let j = 3; j < data.length; j++) {
        if (data[j][i] === "IN") {
          let orderNumberIndex = i + 2;
          let nameIndex = i + 3;
          let dateOfBirthIndex = i + 6;
          let ageGroupIndex = i + 8;
          let roomType = this.getRoomByMin(roomRanges, j - 3);
          let firstDayRoomRecord = new RoomRecord(
            roomNumber,
            firstDate,
            roomType,
            firstDate
          );
          rooms.push(firstDayRoomRecord);
          const roomSize = this.getRoomSize(sizeMap, j + 7);
          for (let peopleCount = 0; peopleCount < roomSize; peopleCount++) {
            const rowNum = j + peopleCount;
            if (data[rowNum][i].startsWith("ホテル")) {
              roomType += "(官網)";
              firstDayRoomRecord.roomType = roomType;
            }
            if (data[rowNum][i].startsWith("WEB")) {
              roomType += "(網路訂房)";
              firstDayRoomRecord.roomType = roomType;
            }
            const orderNumber = data[rowNum][orderNumberIndex];
            const name = data[rowNum][nameIndex];
            const ageGroupText = data[rowNum][ageGroupIndex];
            const dateOfBirth = formatDate(data[rowNum][dateOfBirthIndex]);
            const ageGroup = this.determineAgeGroup(
              ageGroupText,
              dateOfBirth,
              firstDate
            );
            if (name) {
              firstDayRoomRecord.addPerson(
                new PersonRecord(
                  name,
                  null,
                  ageGroup,
                  null,
                  orderNumber,
                  0,
                  null
                )
              );
            }
          }
          nameIndex += 12;
          let adjacentName = data[j][nameIndex];
          let isIn = data[j][nameIndex - 3] === "IN";
          let dayCount = 1;
          let midStaySpecialPricingStartDate = null; // Track when mid-stay special pricing begins
          let currentRoomType = roomType;
          let hasMidStaySpecialPricing = false;
          while (notNull(adjacentName) && !isIn) {
            let adjacentDate = new Date(firstDate);
            adjacentDate.setDate(adjacentDate.getDate() + dayCount);
            adjacentDate = formatDate(adjacentDate);
            
            // Skip special pricing detection if room type is already special
            if (!roomType.endsWith("(官網)") && !roomType.endsWith("(網路訂房)")) {
              for (let peopleCount = 0; peopleCount < roomSize; peopleCount++) {
                const rowNum = j + peopleCount;
                if (data[rowNum][nameIndex - 3]) { // Check the correct position for special pricing markers
                  if (data[rowNum][nameIndex - 3].startsWith("ホテル")) {
                    currentRoomType = roomType + "(官網)";
                    hasMidStaySpecialPricing = true;
                    break;
                  }
                  if (data[rowNum][nameIndex - 3].startsWith("WEB")) {
                    currentRoomType = roomType + "(網路訂房)";
                    hasMidStaySpecialPricing = true;
                    break;
                  }
                }
              }
            }
            
            // Determine the appropriate first date for mid-stay special pricing
            let roomFirstDate = firstDate;
            if (hasMidStaySpecialPricing && midStaySpecialPricingStartDate === null) {
              // This is the start of mid-stay special pricing
              midStaySpecialPricingStartDate = adjacentDate;
              roomFirstDate = midStaySpecialPricingStartDate;
            } else if (hasMidStaySpecialPricing && midStaySpecialPricingStartDate !== null) {
              // Continuing mid-stay special pricing
              roomFirstDate = midStaySpecialPricingStartDate;
            }
            
            roomRecord = new RoomRecord(
              roomNumber,
              adjacentDate,
              currentRoomType,
              roomFirstDate
            );
            roomRecord.people = firstDayRoomRecord.people;
            // roomIndexes[key] = rooms.length
            rooms.push(roomRecord);
            nameIndex += 12;
            dayCount++;
            adjacentName = data[j][nameIndex];
            isIn = data[j][nameIndex - 3] === "IN";
          }
          roomNumber++;
        }
      }
    }
    return rooms;
  }
  determineAgeGroup(ageGroup, dateOfBirth, checkInDate) {
    if (ageGroup) {
      return ageGroup === "CH" ? "孩童" : "成人";
    }
    if (!dateOfBirth) {
      return "成人";
    }
    const today = new Date(checkInDate);
    const birthDate = new Date(dateOfBirth);

    // Calculate age in years
    let age = today.getFullYear() - birthDate.getFullYear();

    // Adjust if the birthday hasn't occurred yet this year
    const isBeforeBirthday =
      today.getMonth() < birthDate.getMonth() ||
      (today.getMonth() === birthDate.getMonth() &&
        today.getDate() < birthDate.getDate());
    if (isBeforeBirthday) {
      age--;
    }

    return age >= 12 ? "成人" : "孩童";
  }
  getRoomSize(sizeMap, index) {
    if (!sizeMap[index]) {
      throw new Error(`Index ${index} not in sizeMap`);
    }
    return sizeMap[index];
  }
  createRoomSize(roomSheet, startRow) {
    const mergedRanges = roomSheet
      .getRange(`C${startRow}:C997`)
      .getMergedRanges();
    const mergedMap = {};
    mergedRanges.forEach((range) => {
      const startRow = range.getRow();
      const numRows = range.getNumRows();
      for (let r = startRow; r < startRow + numRows; r++) {
        mergedMap[r] = numRows;
      }
    });
    return mergedMap;
  }
  createRoomRanges(roomSheet, startRow) {
    const searchValues = roomSheet.getRange(`C${startRow}:C997`).getValues();
    const roomRanges = [];
    searchValues.forEach(([room], index) => {
      room = room.trim ? room.trim() : room;
      if (room.startsWith && room.startsWith("苗王")) {
        roomRanges.push({
          min: index,
          room,
        });
      }
    });
    return roomRanges;
  }
  getRoomByMin(roomRanges, row) {
    for (let i = roomRanges.length - 1; i >= 0; i--) {
      if (row >= roomRanges[i].min) {
        return roomRanges[i].room;
      }
    }
    return null; // Return null for invalid rows
  }
}
