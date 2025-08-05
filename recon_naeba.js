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

  getRoomPrice(roomTypeText, peopleAmount, ageGroup, date, orderNumber = null) {
    // Check if roomType ends with (官網) or (網路訂房)
    if (
      roomTypeText.endsWith("(官網)") ||
      roomTypeText.endsWith("(網路訂房)")
    ) {
      return this.getPriceFromGoogleSheet(
        roomTypeText,
        peopleAmount,
        ageGroup,
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
    ageGroup,
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
      const nameCol = getColumnIndex("H"); // Column H for name
      const peopleCol = getColumnIndex("J"); // Column J for people count

      // Search for matching orderNumber first (most specific match)
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

      // If no orderNumber match found, try matching by room type and date
      for (let i = 1; i < data.length; i++) {
        // Skip header row
        const row = data[i];
        const sheetRoomType = row[roomTypeCol];
        const sheetDate = row[dateCol];
        const sheetPrice = row[priceCol];
        const sheetName = row[nameCol];
        const sheetPeople = row[peopleCol];

        // Match room type (remove the suffix for comparison)
        const baseRoomType = roomTypeText.replace(/\((官網|網路訂房)\)$/, "");

        if (sheetRoomType && sheetRoomType.includes(baseRoomType)) {
          // Format date for comparison
          let formattedSheetDate;
          if (sheetDate instanceof Date) {
            formattedSheetDate = formatDate(sheetDate);
          } else if (typeof sheetDate === "string") {
            formattedSheetDate = formatDate(new Date(sheetDate));
          }

          if (
            formattedSheetDate === date &&
            !isNaN(sheetPrice) &&
            sheetPrice !== null &&
            sheetPrice !== ""
          ) {
            // Return object with price and additional data for room type formatting
            return {
              price: Number(sheetPrice),
              orderNumber: orderNumber || "",
              name: sheetName || "",
              people: sheetPeople || peopleAmount,
              roomType: sheetRoomType || "",
            };
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
      const key = room.roomType + "_" + room.roomNumber + "_" + room.firstDate;
      if (!roomMap.has(key)) {
        roomMap.set(key, new StayRecord(room.firstDate, room.roomNumber));
      }
      roomMap.get(key).addRoomRecord(room);
    }

    const stayRecords = Array.from(roomMap.values());
    return stayRecords;
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
        // const roomTypeText = this.roomTypeMap[room.roomType];
        const roomTypeText = room.roomType;
        if (
          !roomTypeText ||
          roomTypeText.includes("東品") ||
          roomTypeText.includes("不佔床")
        ) {
          return;
        }

        // Check if this is a special pricing room (官網 or 網路訂房)
        const isSpecialPricing =
          roomTypeText.endsWith("(官網)") ||
          roomTypeText.endsWith("(網路訂房)");

        // For special pricing rooms, get the total room price once and distribute across stay days
        let roomPrice = null;
        let roomPriceError = null;
        let specialPriceData = null;

        if (isSpecialPricing) {
          try {
            const rawPrice = this.getRoomPrice(
              roomTypeText,
              room.people.length,
              "成人",
              room.date,
              room.people[0]?.orderNumber
            );
            if (typeof rawPrice === "object" && rawPrice.price) {
              // Divide the total stay price by total nights to get daily price
              roomPrice = rawPrice.price;
              specialPriceData = rawPrice;
            } else if (typeof rawPrice === "number") {
              // Fallback for backward compatibility
              roomPrice = rawPrice;
            } else if (typeof rawPrice === "string") {
              roomPriceError = `${roomTypeText}-${room.people.length}入住-無報價`;
            } else {
              roomPriceError = `${roomTypeText}-${room.people.length}入住-查無房型報價`;
            }
          } catch (e) {
            roomPriceError = `錯誤: 價格查詢例外 - ${e.message}`;
          }
        }

        room.people.forEach((person, index) => {
          const mealType = person.mealType;
          let price;
          let priceError = null;

          if (!roomTypeText) {
            // priceError = `錯誤: 找不到對應房型 ${room.roomType}`;
            return;
          } else {
            if (isSpecialPricing) {
              // For special pricing, only assign the daily room price to the first person
              if (index === 0) {
                price = roomPrice;
                priceError = roomPriceError;
              } else {
                // Other people in the same room don't get room price (already counted)
                price = 0;
              }
            } else {
              // For regular pricing, get price per person
              try {
                const rawPrice = this.getRoomPrice(
                  roomTypeText,
                  room.people.length,
                  person.ageGroup,
                  room.date,
                  person.orderNumber
                );
                if (typeof rawPrice === "number") {
                  price = Math.round(rawPrice);
                } else if (typeof rawPrice === "string") {
                  priceError = `${roomTypeText}-${room.people.length}入住-無報價`;
                } else {
                  priceError = `${roomTypeText}-${room.people.length}入住-查無房型報價`;
                }
              } catch (e) {
                priceError = `錯誤: 價格查詢例外 - ${e.message}`;
              }
            }
          }

          const dateResult = result[calDate];

          // Determine the grouping category based on room type
          let groupCategory;
          if (roomTypeText.endsWith("(官網)")) {
            groupCategory = "官網";
          } else if (roomTypeText.endsWith("(網路訂房)")) {
            groupCategory = "網路訂房";
          } else {
            groupCategory = person.ageGroup; // Use original age group for regular rooms
          }

          const ageGroupData = dateResult[groupCategory];
          if (!ageGroupData) {
            return;
          }

          // Format room type based on whether it's special pricing or not
          if (price) {
            let roomType;
            if (isSpecialPricing && specialPriceData) {
              // For special pricing, only add room type for the first person on the first day (same logic as price)
              if (index === 0 && room.date === room.firstDate) {
                // Format: {orderNumber}{name in specialPriceSheet}{room people(specialPriceSheet column J)}{price}
                // Add unique identifier to prevent aggregation, will be stripped in result_writer
                roomType = `${specialPriceData.orderNumber}${
                  specialPriceData.name || "無名稱"
                }${specialPriceData.people}人$${specialPriceData.price}|${
                  room.roomNumber
                }_${room.date}_${index}`;
              }
            } else {
              roomType = `${room.roomType}-${room.people.length}入住${totalNights}晚`;
            }

            // Only process room type if it's defined
            if (roomType) {
              ageGroupData.房型[roomType] = ageGroupData.房型[roomType] || {
                people: 0,
                price: 0,
                totalNights,
              };

              ageGroupData.房型[roomType].people++;
              const dayDiff =
                dateDifferenceInDays(room.date, room.firstDate) + 1;
              ageGroupData.房型[roomType].day =
                ageGroupData.房型[roomType].day > dayDiff
                  ? ageGroupData.房型[roomType].day
                  : dayDiff;
              if (typeof price === "number") {
                ageGroupData.房型[roomType].price += price;
                ageGroupData.房費 += price;
              }
            }
          }

          // Log error if price lookup failed
          if (priceError) {
            ageGroupData.房型.錯誤列表 = ageGroupData.房型.錯誤列表 || [];
            if (!ageGroupData.房型.錯誤列表.includes(priceError)) {
              ageGroupData.房型.錯誤列表.push(priceError);
            }
          }

          if (mealType === "早餐" || mealType === "早晚餐") {
            ageGroupData.餐食 += foodCostData.雫石[person.ageGroup].早餐;
          }
          if (mealType === "晚餐" || mealType === "早晚餐") {
            ageGroupData.餐食 += foodCostData.雫石[person.ageGroup].晚餐;
          }

          if (person.ageGroup === "成人") {
            ageGroupData.入湯稅 += 150;
          }
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
            person.orderNumber
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
