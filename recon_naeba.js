class NaebaReconciliationLogic extends BaseReconciliationLogic {
  constructor(priceSheet, roomTypeMap) {
    super();
    this.priceSheet = priceSheet;
    this.roomTypeMap = roomTypeMap;
    this.specialPriceData = null;
  }

  createPriceTable() {
    const priceTable = {};
    const priceValues = this.priceSheet.getRange("F4:DE127").getValues();
    const dateValues = this.priceSheet.getRange("A7:A127").getValues();
    transpose(priceValues).forEach((col) => {
      const roomType = col[0];
      const people = col[1];
      const ageGroup = col[2];
      priceTable[roomType] = priceTable[roomType] || {};
      priceTable[roomType][people] = priceTable[roomType][people] || {};
      priceTable[roomType][people][ageGroup] =
        priceTable[roomType][people][ageGroup] || {};

      dateValues.forEach(([date], index) => {
        const formattedDate = formatDate(date);
        priceTable[roomType][people][ageGroup][formattedDate] = col[index + 3];
      });
    });
    this.priceTable = priceTable;
  }

  loadSpecialPriceData() {
    if (this.specialPriceData === null) {
      try {
        const specialPriceSheet = SpreadsheetApp.openById(
          "1OUafnXTtXwwKgkWlRKJn9Om34_egwvieQfMhdRVENjE"
        );
        const sheet = specialPriceSheet.getSheetByName("訂房預約狀況(使用中)");
        this.specialPriceData = sheet.getDataRange().getValues();
      } catch (error) {
        console.error("Error loading special price data:", error);
        this.specialPriceData = [];
      }
    }
  }

  getRoomPrice(roomTypeText, peopleAmount, ageGroup, date, orderNumber = null, name = null) {
    // Check if roomType ends with (官網) or (網路訂房)
    if (
      roomTypeText.endsWith("(官網)") ||
      roomTypeText.endsWith("(網路訂房)")
    ) {
      return this.getPriceFromGoogleSheet(
        roomTypeText,
        peopleAmount,
        name,
        date,
        orderNumber
      );
    }

    let maxPeople = Math.max(...Object.keys(this.priceTable[roomTypeText]));
    peopleAmount = peopleAmount > maxPeople ? maxPeople : peopleAmount;
    const rawPrice =
      this.priceTable?.[roomTypeText]?.[peopleAmount]?.["成人"]?.[date];
    return rawPrice;
  }

  getPriceFromGoogleSheet(
    roomTypeText,
    peopleAmount,
    name,
    date,
    orderNumber
  ) {
    try {
      // Use cached data
      const data = this.specialPriceData;

      // Get column indices using helper function
      const orderNumberCol = getColumnIndex("AG");
      const roomTypeCol = getColumnIndex("E");
      const dateCol = getColumnIndex("B");
      const priceCol = getColumnIndex("L");
      const nameCol = getColumnIndex("AH"); // Column H for name
      const peopleCol = getColumnIndex("J"); // Column J for people count

      // Search for matching orderNumber first (most specific match)
      if (orderNumber && name) {
        for (let i = 1; i < data.length; i++) {
          // Skip header row
          const row = data[i];
          const sheetOrderNumber = row[orderNumberCol];
          const sheetPrice = row[priceCol];
          const sheetName = row[nameCol];
          const sheetPeople = row[peopleCol];

          if (
            sheetOrderNumber &&
            sheetOrderNumber.toString() === orderNumber.toString() &&
            sheetName === name
          ) {
            if (
              !isNaN(sheetPrice) &&
              sheetPrice !== null &&
              sheetPrice !== ""
            ) {
              // Return object with price and additional data for room type formatting
              return {
                price: Number(sheetPrice),
                orderNumber: sheetOrderNumber,
                name: sheetName || "",
                people: sheetPeople || peopleAmount,
                roomType: row[roomTypeCol] || "",
              };
            }
          }
        }
      }

      if (orderNumber) {
        for (let i = 1; i < data.length; i++) {
          // Skip header row
          const row = data[i];
          const sheetOrderNumber = row[orderNumberCol];
          const sheetPrice = row[priceCol];
          const sheetName = row[nameCol];
          const sheetPeople = row[peopleCol];

          if (
            sheetOrderNumber &&
            sheetOrderNumber.toString() === orderNumber.toString() 
          ) {
            if (
              !isNaN(sheetPrice) &&
              sheetPrice !== null &&
              sheetPrice !== ""
            ) {
              // Return object with price and additional data for room type formatting
              return {
                price: Number(sheetPrice),
                orderNumber: sheetOrderNumber,
                name: sheetName || "",
                people: sheetPeople || peopleAmount,
                roomType: row[roomTypeCol] || "",
              };
            }
          }
        }
      }

      // If no matching record found
      return `${roomTypeText}-${peopleAmount}入住-Google Sheet無報價`;
    } catch (error) {
      console.error("Error accessing Google Sheet:", error);
      return `${roomTypeText}-${peopleAmount}入住-Google Sheet查詢錯誤`;
    }
  }

  createStayRecords(rooms) {
    const roomMap = new Map(); // key = roomNumber + firstDate

    for (let room of rooms) {
      let key = room.roomType + "_" + room.roomNumber + "_" + room.firstDate + "_" + room.firstDateIndex;
      if (!roomMap.has(key)) {
        roomMap.set(key, new StayRecord(room.firstDate, room.roomNumber));
      }
      roomMap.get(key).addRoomRecord(room);
    }

    const stayRecords = Array.from(roomMap.values());
    return stayRecords;
  }

  getSpecialPricingData(roomTypeText, roomPeopleCount, orderNumber, date, name) {
    try {
      const rawPrice = this.getRoomPrice(roomTypeText, roomPeopleCount, "成人", date, orderNumber, name);
      
      if (typeof rawPrice === "object" && rawPrice.price) {
        return { price: rawPrice.price, data: rawPrice, error: null };
      } else if (typeof rawPrice === "number") {
        return { price: rawPrice, data: null, error: null };
      } else if (typeof rawPrice === "string") {
        return { price: null, data: null, error: `${roomTypeText}-${roomPeopleCount}入住-無報價` };
      } else {
        return { price: null, data: null, error: `${roomTypeText}-${roomPeopleCount}入住-查無房型報價` };
      }
    } catch (e) {
      return { price: null, data: null, error: `錯誤: 價格查詢例外 - ${e.message}` };
    }
  }

  getRegularPricingData(roomTypeText, roomPeopleCount, person, date) {
    try {
      const rawPrice = this.getRoomPrice(roomTypeText, roomPeopleCount, person.ageGroup, date, person.orderNumber);
      
      if (typeof rawPrice === "number") {
        return { price: Math.round(rawPrice), error: null };
      } else if (typeof rawPrice === "string") {
        return { price: null, error: `${roomTypeText}-${roomPeopleCount}入住-無報價` };
      } else {
        return { price: null, error: `${roomTypeText}-${roomPeopleCount}入住-查無房型報價` };
      }
    } catch (e) {
      return { price: null, error: `錯誤: 價格查詢例外 - ${e.message}` };
    }
  }

  formatRoomType(isSpecialPricing, specialPriceData, room, totalNights, index) {
    if (isSpecialPricing && specialPriceData) {
      // For special pricing, only add room type for the first person on the first day
      if (index === 0 && room.date === room.firstDate) {
        return `${specialPriceData.orderNumber}${specialPriceData.name || "無名稱"}${specialPriceData.people}人$${specialPriceData.price}|${room.roomNumber}_${room.date}_${index}`;
      }
      return null;
    } else {
      return `${room.roomType}-${room.people.length}入住${totalNights}晚`;
    }
  }

  getGroupCategory(roomTypeText, person) {
    if (roomTypeText.endsWith("(官網)")) {
      return "官網";
    } else if (roomTypeText.endsWith("(網路訂房)")) {
      return "網路訂房";
    } else {
      return person.ageGroup;
    }
  }

  processRoomType(ageGroupData, roomType, price, room, totalNights) {
    if (!roomType) return;

    ageGroupData.房型[roomType] = ageGroupData.房型[roomType] || {
      people: 0,
      price: 0,
      totalNights,
    };

    ageGroupData.房型[roomType].people++;
    const dayDiff = dateDifferenceInDays(room.date, room.firstDate) + 1;
    ageGroupData.房型[roomType].day = ageGroupData.房型[roomType].day > dayDiff
      ? ageGroupData.房型[roomType].day
      : dayDiff;

    if (typeof price === "number") {
      ageGroupData.房型[roomType].price += price;
      ageGroupData.房費 += price;
    }
  }

  processMealAndTax(ageGroupData, person) {
    const mealType = person.mealType;
    
    if (mealType === "早餐" || mealType === "早晚餐") {
      ageGroupData.餐食 += foodCostData.苗王[person.ageGroup].早餐;
    }
    if (mealType === "晚餐" || mealType === "早晚餐") {
      ageGroupData.餐食 += foodCostData.苗王[person.ageGroup].晚餐;
    }
    if (person.ageGroup === "成人") {
      ageGroupData.入湯稅 += 150;
    }
  }

  reconcile(rooms) {
    this.createPriceTable();
    this.loadSpecialPriceData();
    const result = {};

    const stays = this.createStayRecords(rooms);
    stays.forEach((stay) => {
      const stayRooms = stay.roomRecords;
      const totalNights = stay.getTotalNights();
      const calDate = stay.firstDate;
      
      result[calDate] = result[calDate] || {
        成人: { 房費: 0, 餐食: 0, 入湯稅: 0, 房型: {} },
        孩童: { 房費: 0, 餐食: 0, 入湯稅: 0, 房型: {} },
        官網: { 房費: 0, 餐食: 0, 入湯稅: 0, 房型: {} },
        網路訂房: { 房費: 0, 餐食: 0, 入湯稅: 0, 房型: {} },
      };

      stayRooms.forEach((room) => {
        const roomTypeText = room.roomType;
        if (!roomTypeText || roomTypeText.includes("東品") || roomTypeText.includes("不佔床")) {
          return;
        }

        const isSpecialPricing = roomTypeText.endsWith("(官網)") || roomTypeText.endsWith("(網路訂房)");
        let specialPricingResult = null;

        // Get special pricing data once per room if needed
        if (isSpecialPricing) {
          specialPricingResult = this.getSpecialPricingData(
            roomTypeText, 
            room.people.length, 
            room.people[0]?.orderNumber, 
            room.date,
            room.people[0]?.name
          );
        }

        room.people.forEach((person, index) => {
          let price = null;
          let priceError = null;

          // Calculate price based on pricing type
          if (isSpecialPricing) {
            if (index === 0) {
              price = specialPricingResult.price;
              priceError = specialPricingResult.error;
            } else {
              price = 0; // Other people don't get room price
            }
          } else {
            const regularResult = this.getRegularPricingData(roomTypeText, room.people.length, person, room.date);
            price = regularResult.price;
            priceError = regularResult.error;
          }

          const groupCategory = this.getGroupCategory(roomTypeText, person);
          const ageGroupData = result[calDate][groupCategory];
          
          if (!ageGroupData) return;

          // Process room type if there's a valid price
          if (price) {
            const roomType = this.formatRoomType(isSpecialPricing, specialPricingResult?.data, room, totalNights, index);
            this.processRoomType(ageGroupData, roomType, price, room, totalNights);
          }

          // Log errors
          if (priceError) {
            ageGroupData.房型.錯誤列表 = ageGroupData.房型.錯誤列表 || [];
            if (!ageGroupData.房型.錯誤列表.includes(priceError)) {
              ageGroupData.房型.錯誤列表.push(priceError);
            }
          }

          // Process meals and tax
          this.processMealAndTax(ageGroupData, person);
        });
      });
    });

    return result;
  }

  checkAU(rooms) {
    this.createPriceTable();
    const result = {};
    const stays = this.createStayRecords(rooms);
    stays.forEach((stay) => {
      const stayRooms = stay.roomRecords;
      const totalNights = stay.getTotalNights();
      stayRooms.forEach((room) => {
        const roomTypeText = this.roomTypeMap[room.roomType];
        if (!roomTypeText) return;
        room.people.forEach((person, index) => {
          let price;
          let key = person.orderNumber + "_" + person.name;
          if (person.continuation !== 1) key += `_${person.continuation}`;
          result[key] = result[key] || {
            AU價格: person.auPrice,
            計算價格: 0,
          };

          const rawPrice = this.getRoomPrice(
            roomTypeText,
            room.people.length,
            person.ageGroup,
            room.date,
            person.orderNumber,
            person.name
          );
          if (
            typeof rawPrice === "number" &&
            result[key]["計算價格"] !== "無報價"
          ) {
            price = Math.round(rawPrice);
            result[key]["計算價格"] += price;
          } else {
            result[key]["計算價格"] = "無報價";
          }
        });
      });
    });
    const tableOutput = [];
    Object.keys(result).forEach((key) => {
      const auPrice = result[key]["AU價格"];
      const calPrice = result[key]["計算價格"];
      if (auPrice + calPrice !== 0) {
        tableOutput.push([key, auPrice, calPrice, auPrice + calPrice]);
      }
    });
    return tableOutput;
  }
}
