class BaseInvoiceParser {
  parse(sheet) {
    const data = sheet.getRange("A13:FG3000").getValues();
    const rooms = [];
    const roomIndexs = {};
    let stayContext = this.initializeStayContext();

    for (let i = 0; i < data.length; i++) {
      const row = data[i];
      const guestInfo = this.extractGuestInfo(row);

      stayContext = this.resetContextIfNewGuest(
        stayContext,
        guestInfo.continuation
      );

      for (let day = 0; day < 6; day++) {
        const dayInfo = this.extractDayInfo(row, day);

        if (notNull(dayInfo.roomType)) {
          const scenario = this.detectScenario(
            dayInfo.roomType,
            stayContext.previousRoomType
          );
          stayContext = this.handleScenario(
            scenario,
            dayInfo,
            stayContext,
            rooms,
            roomIndexs,
            guestInfo
          );
        } else {
          stayContext = this.handleEmptyRoomType(stayContext);
        }
      }
    }
    return rooms;
  }

  initializeStayContext() {
    return {
      firstDate: null,
      currentStayFirstDate: null,
      previousRoomType: null,
      firstDateIndexMap: {},
    };
  }

  extractGuestInfo(row) {
    const rawAge = row[getColumnIndex("BW")];
    const ageType = rawAge === "成人" ? rawAge : "孩童";

    return {
      continuation: row[getColumnIndex("CA")],
      bonusMeal: row[getColumnIndex("FG")],
      orderNumber: row[getColumnIndex("D")],
      auPrice: row[getColumnIndex("AU")],
      guestName: row[4],
      ageType: ageType,
    };
  }

  extractDayInfo(row, day) {
    const baseCol = getColumnIndex("DB") + day * 8;
    const date = row[baseCol];
    return {
      date,
      roomType: row[baseCol + 5],
      roomNumber: row[baseCol + 6],
      mealType: row[baseCol + 7],
      formattedDate: date instanceof Date ? formatDate(date) : null,
      day,
    };
  }

  resetContextIfNewGuest(stayContext, continuation) {
    if (continuation === 1) {
      return this.initializeStayContext();
    }
    return stayContext;
  }

  detectScenario(roomType, previousRoomType) {
    const isSpecialPricing =
      roomType.endsWith("(官網)") || roomType.endsWith("(網路訂房)");
    const isRoomTypeChange = previousRoomType && previousRoomType !== roomType;

    if (isSpecialPricing) {
      return "specialPricing";
    } else if (isRoomTypeChange) {
      return "changeRoomType";
    } else {
      return "regular";
    }
  }

  handleScenario(scenario, dayInfo, stayContext, rooms, roomIndexs, guestInfo) {
    const updatedContext = { ...stayContext };

    if (!updatedContext.previousRoomType) {
      updatedContext.previousRoomType = dayInfo.roomType;
    }

    let roomFirstDate;

    switch (scenario) {
      case "specialPricing":
        roomFirstDate = this.handleSpecialPricing(dayInfo, updatedContext);
        break;
      case "changeRoomType":
        roomFirstDate = this.handleRoomTypeChange(dayInfo, updatedContext);
        break;
      case "regular":
      default:
        roomFirstDate = this.handleRegularScenario(dayInfo, updatedContext);
        break;
    }

    this.createOrUpdateRoomRecord(
      dayInfo,
      roomFirstDate,
      rooms,
      roomIndexs,
      guestInfo,
      updatedContext.firstDateIndexMap
    );
    updatedContext.previousRoomType = dayInfo.roomType;

    return updatedContext;
  }

  handleSpecialPricing(dayInfo, stayContext) {
    if (stayContext.currentStayFirstDate === null) {
      stayContext.currentStayFirstDate = dayInfo.formattedDate;
    }
    return stayContext.currentStayFirstDate;
  }

  handleRoomTypeChange(dayInfo, stayContext) {
    stayContext.firstDate = dayInfo.formattedDate;
    return stayContext.firstDate;
  }

  handleRegularScenario(dayInfo, stayContext) {
    if (stayContext.firstDate === null) {
      stayContext.firstDate = dayInfo.formattedDate;
    }
    return stayContext.firstDate;
  }

  handleEmptyRoomType(stayContext) {
    return {
      firstDate: null,
      currentStayFirstDate: null,
      previousRoomType: stayContext.previousRoomType,
      firstDateIndexMap: stayContext.firstDateIndexMap,
    };
  }

  createOrUpdateRoomRecord(
    dayInfo,
    roomFirstDate,
    rooms,
    roomIndexs,
    guestInfo,
    firstDateIndexMap
  ) {
    const key = [
      dayInfo.formattedDate,
      dayInfo.day,
      dayInfo.roomType,
      dayInfo.roomNumber,
      guestInfo.continuation,
    ].join("/");

    let firstDateIndex;
    if (roomFirstDate === dayInfo.formattedDate) {
      if (firstDateIndexMap[roomFirstDate] !== undefined) {
        firstDateIndex = firstDateIndexMap[roomFirstDate];
      } else {
        firstDateIndex = dayInfo.day;
        firstDateIndexMap[roomFirstDate] = firstDateIndex;
      }
    } else {
      if (firstDateIndexMap[roomFirstDate] !== undefined) {
        firstDateIndex = firstDateIndexMap[roomFirstDate];
      } else {
        firstDateIndex = null;
      }
    }

    let roomRecord;
    if (key in roomIndexs) {
      roomRecord = rooms[roomIndexs[key]];
    } else {
      roomRecord = new RoomRecord(
        dayInfo.roomNumber,
        dayInfo.formattedDate,
        dayInfo.roomType,
        roomFirstDate,
        firstDateIndex
      );
      roomIndexs[key] = rooms.length;
      rooms.push(roomRecord);
    }

    const ageType = guestInfo.ageType;
    const personRecord = new PersonRecord(
      guestInfo.guestName,
      dayInfo.mealType,
      ageType,
      guestInfo.bonusMeal,
      guestInfo.orderNumber,
      guestInfo.auPrice,
      guestInfo.continuation
    );
    roomRecord.addPerson(personRecord);
  }
}
class BaseRoomInvoiceParser {
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
          let foodIndex = i + 1;
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
            const food = foodTextMappingData[data[rowNum][foodIndex]];
            const ageGroup = this.determineAgeGroup(
              ageGroupText,
              dateOfBirth,
              firstDate
            );
            if (name) {
              firstDayRoomRecord.addPerson(
                new PersonRecord(
                  name,
                  food,
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
            if (
              !roomType.endsWith("(官網)") &&
              !roomType.endsWith("(網路訂房)")
            ) {
              for (let peopleCount = 0; peopleCount < roomSize; peopleCount++) {
                const rowNum = j + peopleCount;
                if (data[rowNum][nameIndex - 3]) {
                  // Check the correct position for special pricing markers
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
            if (
              hasMidStaySpecialPricing &&
              midStaySpecialPricingStartDate === null
            ) {
              // This is the start of mid-stay special pricing
              midStaySpecialPricingStartDate = adjacentDate;
              roomFirstDate = midStaySpecialPricingStartDate;
            } else if (
              hasMidStaySpecialPricing &&
              midStaySpecialPricingStartDate !== null
            ) {
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

class BaseSnowTicketParser {
  parse(sheet) {
    throw new Error("parse() 必須由子類別實作");
  }
}
