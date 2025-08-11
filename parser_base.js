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

class BaseSnowTicketParser {
  parse(sheet) {
    throw new Error("parse() 必須由子類別實作");
  }
}
