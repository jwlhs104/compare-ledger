function compareLedgers() {
  const ledgers = [
    {
      resort: "苗王",
      roomSheetId: "1fkUnjDDIv-dN6HJUCpFcPhanCmbvGO2Amdn3LQdFCC0",
      roomSheetName: "Naeba25-26",
      priceSheetId: "1S6og0-YWx_tW8S52DHl3PwyWhs3J8OGSAjNoR22iHv8",
      priceSheetName: "苗王",
      finalRoomSheets: getFinalRoomSheets("1185z1drXgKZ9B13gI2hkwW22ZOa0-A4j"),
      snowTicketStr: "I4YLE-",
      snowTicketSinglePrice: 4600,
    },
    // {
    //   resort: "雫石",
    //   roomSheetId: "1y9OoKqukIhNMTuP9tyNTn_SuHltCz9jWhoszVrmjtgY",
    //   roomSheetName: "Shizukuishi24-25",
    //   priceSheetId: "1zgoWCVg4voK9rjos2NhbVXlZf8LXSrQRerAep5U3kAA",
    //   priceSheetName: "雫石",
    //   snowTicketStr: "I2YEL-",
    //   snowTicketSinglePrice: 4500,
    //   finalRoomSheets: getFinalRoomSheets("1oRFJzErn0sz-4DZzcRhWnTO8g3pF9HW8"),
    //   // snowTicketCountMethod: "byDay"
    // },
    // {
    //   resort: "萬座",
    //   roomSheetId: "1ephb8kx0upJzf5oL4HkBFMktPF8_j99G9349NqiCoSQ",
    //   roomSheetName: "Manza24-25",
    //   priceSheetId: "1zgoWCVg4voK9rjos2NhbVXlZf8LXSrQRerAep5U3kAA",
    //   priceSheetName: "萬座",
    //   snowTicketStr: "I3YLE-",
    //   snowTicketSinglePrice: 2500,
    //   finalRoomSheets: getFinalRoomSheets("16zSasV5sN6m2BsxYu_6Lg9O_K5YwW1_4"),
    //   // snowTicketCountMethod: "byDay"
    // },
    // {
    //   resort: "輕井澤",
    //   roomSheetId: "1nnd9-CHbRKw75XfLvR1jDpzIlS4LV4cFJx-AsVhu-u0",
    //   roomSheetName: "Karuizawa24-25",
    //   priceSheetId: "1zgoWCVg4voK9rjos2NhbVXlZf8LXSrQRerAep5U3kAA",
    //   priceSheetName: "輕井澤",
    //   snowTicketStr: "I3YLE-",
    //   snowTicketSinglePrice: 5200,
    //   finalRoomSheets: getFinalRoomSheets("1qCL2VZSXqZEyt0jeD81IGEk9z7i3oKVx"),
    //   // snowTicketCountMethod: "byDay"
    // },
    // {
    //   resort: "志賀",
    //   roomSheetId: "1AwLsT205kIEe75rFPalckQdqDIb96IzWPbIER6rOChU",
    //   roomSheetName: "Shigakogen24-25",
    //   priceSheetId: "1zgoWCVg4voK9rjos2NhbVXlZf8LXSrQRerAep5U3kAA",
    //   priceSheetName: "志賀",
    //   snowTicketStr: "I3YLE-",
    //   snowTicketSinglePrice: 6000,
    //   finalRoomSheets: getFinalRoomSheets("1gH5qVJwLpYoolP5AkM4CW8BL8bPZyOmx"),
    //   // snowTicketCountMethod: "byDay"
    // },
    // {
    //   resort: "品川",
    //   roomSheetId: "1MiQVzJUfjxXknjt09RtOrRZ2Cn5REb-4Na8AlKMMz94",
    //   roomSheetName: "Shinagawa24-25",
    //   priceSheetId: "1zgoWCVg4voK9rjos2NhbVXlZf8LXSrQRerAep5U3kAA",
    //   priceSheetName: "品川",
    //   snowTicketStr: "I3YLE-",
    //   snowTicketSinglePrice: 6000,
    //   finalRoomSheets: getFinalRoomSheets("1aVI-FyCkZwxp9y-gS3r-zXvJb2122fcb"),
    //   // snowTicketCountMethod: "byDay"
    // },
    // {
    //   resort: "音羽屋",
    //   roomSheetId: "1gi8vWDpiklVjWcit711c8QLGHqAhI-iTP8trpLGJRTE",
    //   roomSheetName: "Otowaya24-25",
    //   priceSheetId: "1zgoWCVg4voK9rjos2NhbVXlZf8LXSrQRerAep5U3kAA",
    //   priceSheetName: "音羽屋",
    //   snowTicketStr: "I3YLE-",
    //   snowTicketSinglePrice: 6000,
    //   finalRoomSheets: getFinalRoomSheets("1DJ-CA1Rw3qwnNP42xdgKu7fPN_NWE3mu"),
    //   // snowTicketCountMethod: "byDay"
    // },
  ];
  ledgers.forEach((ledger) => {
    compareLedger(ledger);
  });
  function getFinalRoomSheets(folderId) {
    const folder = DriveApp.getFolderById(folderId);
    const files = folder.getFiles();
    const subfolders = folder.getFolders();
    const fileList = [];

    while (files.hasNext()) {
      const file = files.next();
      fileList.push({
        name: file.getName(),
        id: file.getId(),
      });
    }

    while (subfolders.hasNext()) {
      const subfolder = subfolders.next();
      const subfolderName = subfolder.getName();
      if (subfolderName.includes("回覆") || subfolderName.includes("NG")) {
        continue;
      }
      fileList.push(...getFinalRoomSheets(subfolder.getId()));
    }

    return fileList;
  }
}

function compareLedger({
  roomSheetId,
  roomSheetName,
  resort,
  priceSheetId,
  priceSheetName,
  finalRoomSheets,
  snowTicketStr,
  snowTicketSinglePrice,
  snowTicketCountMethod,
}) {
  const startRow = 10;
  const dateSize = 12;
  const dateOffset = 4;
  const targetRowOffset = 11;

  const roomSheet =
    SpreadsheetApp.openById(roomSheetId).getSheetByName(roomSheetName);
  const targetSheet = SpreadsheetApp.getActive().getSheetByName("房費核對");
  const priceSheet =
    SpreadsheetApp.openById(priceSheetId).getSheetByName(priceSheetName);
  const timeColumn = priceSheet.getRange("A1:A").getDisplayValues();
  const roomRows = transpose(priceSheet.getRange("4:6").getValues());
  const roomValues = roomSheet.getRange("Q10:BCZ997").getValues();
  const roomRanges = createRoomRanges(roomSheet, startRow);

  let snowTicketCountingDict = {};
  const startDate = new Date(2025, 11, 2);
  const lastDay = new Date(2026, 2, 31);
  let searchIndex = 0;
  for (
    var searchDate = startDate;
    searchDate <= lastDay;
    searchDate.setDate(searchDate.getDate() + 1)
  ) {
    Logger.log(searchDate);
    targetSheet.getRange(targetRowOffset + searchIndex, 1).setValue(searchDate);
    targetSheet.getRange(targetRowOffset + searchIndex, 2).setValue("成人");
    targetSheet
      .getRange(targetRowOffset + searchIndex + 1, 1)
      .setValue(searchDate);
    targetSheet.getRange(targetRowOffset + searchIndex + 1, 2).setValue("孩童");
    targetSheet
      .getRange(targetRowOffset + searchIndex + 2, 1)
      .setValue(searchDate);
    targetSheet.getRange(targetRowOffset + searchIndex + 2, 2).setValue("其他");
    targetSheet
      .getRange(targetRowOffset + searchIndex + 3, 1)
      .setValue(searchDate);
    targetSheet
      .getRange(targetRowOffset + searchIndex + 3, 2)
      .setValue("當日總計");
    targetSheet
      .getRange(targetRowOffset + searchIndex + 3, 1, 1, 38)
      .setBackground("#cfe2f3")
      .setFontColor("#0000ff");

    if (!targetSheet.getRange(targetRowOffset + searchIndex, 3).getValue()) {
      searchIndex += 4;
      continue;
    }
    if (
      targetSheet.getRange(targetRowOffset + searchIndex, 11).getValue() ||
      targetSheet.getRange(targetRowOffset + searchIndex, 21).getValue()
    ) {
      searchIndex += 4;
      continue;
    }
    // 有FNL房表才計算
    const searchValues = roomValues.map((row) =>
      row.slice(searchIndex * 3, searchIndex * 3 + 12 * 8)
    );
    _searchAndCalculate(roomSheet, startRow, 11, searchValues);
    if (finalRoomSheets) {
      const finalRoomSheet = getFinalRoomSheet(finalRoomSheets, searchDate);
      if (finalRoomSheet) {
        // _searchAndCalculate(roomSheet, startRow, 11)
        _searchAndCalculate(finalRoomSheet, 4, 21);
      } else {
        // targetSheet.getRange(targetRowOffset+searchIndex, 12, 3, 6).setValue("無FNL")
        targetSheet
          .getRange(targetRowOffset + searchIndex, 22, 3, 6)
          .setValue("無FNL");
      }
    }

    searchIndex += 4;
  }
  function _searchAndCalculateShigakogen(roomSheet) {
    const searchValues = roomSheet.getRange("A3:I").getValues();
    let rooms = [];
    let room = {};
    searchValues.forEach((row, index) => {
      if (index % 4 === 0) {
        if (Object.keys(room).length > 0) {
          if (room.people.length > 0) {
            room.type = formatRoomText(room.type, room.people.length);
            room.prices = getRoomPrice(
              room.checkInDate,
              room.stayingDays,
              room.type,
              room.people.length
            );
            rooms.push(room);
          }
          room = {};
        }
        room.size = 4;
        room.checkInDate = row[1];
        room.stayingDays = row[2];
        room.checkOutDate = row[3];
        room.type = "志西四床${people}21.7";
      }
      const name = row[4];
      const id = row[5];
      const dateOfBirth = row[6];
      room.people = room.people || [];
      if (name) {
        room.people.push({
          name,
          id,
          dateOfBirth,
          foods: [],
        });
      }
    });
    return rooms;
  }
  function _searchAndCalculate(
    roomSheet,
    startRow,
    targetColOffset,
    searchValues
  ) {
    if (roomSheet.getRange("A1").getValue() === "到着日") {
      const rooms = _searchAndCalculateShigakogen(roomSheet);
      createResult(rooms);
      return;
    }
    if (!searchValues) {
      const dateRow = roomSheet
        .getRange(`${startRow - 3}:${startRow - 3}`)
        .getDisplayValues();
      roomRanges = createRoomRanges(roomSheet, startRow);
      const searchDateString = formatDate(searchDate);
      const dateIndex =
        dateRow[0].findIndex((date) => date === searchDateString) + 1;
      if (dateIndex === 0) return;
      searchValues = roomSheet
        .getRange(
          startRow,
          dateIndex - dateOffset,
          roomSheet.getLastRow() - startRow + 1,
          dateSize * 8
        )
        .getValues();
    }

    const rooms = [];
    let room = {};
    let roomCounter = 0;
    searchValues.forEach((row, index) => {
      if (row[0] === "IN") {
        if (Object.keys(room).length !== 0) {
          Logger.log([room, roomCounter]);
          throw new Error("room is not empty");
        }

        let nameIndex = 3 + 12;
        let stayingDays = 1;
        let adjacentName = searchValues[index][nameIndex];
        let isIn = searchValues[index][nameIndex - 3] === "IN";
        while (notNull(adjacentName) && !isIn) {
          nameIndex += 12;
          stayingDays++;
          adjacentName = searchValues[index][nameIndex];
          isIn = searchValues[index][nameIndex - 3] === "IN";
        }
        let checkOutDate = new Date(searchDate);
        checkOutDate.setDate(searchDate.getDate() + stayingDays - 1);

        room.checkInDate = searchDate;
        room.size = getRoomSize(roomSheet.getRange(startRow + index + 1, 3));
        room.stayingDays = stayingDays;
        room.checkOutDate = checkOutDate;
        room.type = getRoomByMin(roomRanges, index);
        if (!isNaN(row[8])) {
          room.manualPrice = row[8];
        }
        roomCounter = 0;
      }
      if (roomCounter <= room.size) {
        const name = row[4];
        const id = row[5];
        const dateOfBirth = row[6];
        const snowTicket = row[7];
        const ageGroup = row[8];
        const foods = [];
        let foodIndex = 1;
        for (let i = 0; i < room.stayingDays; i++) {
          foods.push(searchValues[index][foodIndex]);
          foodIndex += 12;
        }
        if (name) {
          room.people = room.people || [];
          room.people.push({
            name,
            id,
            dateOfBirth,
            snowTicket,
            foods,
            ageGroup,
          });
        }
        roomCounter++;
      }
      if (roomCounter === room.size) {
        room.type = formatRoomText(room.type, room.people.length);
        room.prices = getRoomPrice(
          room.checkInDate,
          room.stayingDays,
          room.type,
          room.people.length
        );
        rooms.push(room);
        roomCounter = 0;
        room = {};
      }
    });
    createResult(rooms);
    function createResult(rooms) {
      let result = {
        成人: {
          roomPrice: 0,
          snowTicketPrice: 0,
          hotSpringFee: 0,
          foodPrice: 0,
        },
        孩童: {
          roomPrice: 0,
          snowTicketPrice: 0,
          hotSpringFee: 0,
          foodPrice: 0,
        },
        其他: {
          roomPrice: 0,
          snowTicketPrice: 0,
          hotSpringFee: 0,
          foodPrice: 0,
        },
      };
      // Logger.log(JSON.stringify(rooms, null, 2))
      let roomErrors = {
        成人: {},
        孩童: {},
        其他: {},
      };
      let roomDetails = {
        成人: {},
        孩童: {},
        其他: {},
      };
      let roomManuals = {
        成人: {},
        孩童: {},
        其他: {},
      };
      rooms.forEach((room) => {
        let currentRoomErrors = room.prices.filter((price) => isNaN(price));
        const useManual = currentRoomErrors.length > 0 && room.manualPrice;
        if (useManual) {
          result["成人"]["roomPrice"] += room.manualPrice;
        }

        room.people.forEach((person, personIndex) => {
          const age = determineAgeGroup(person, searchDate);

          if (useManual) {
            const key = room.people.length + "-" + room.stayingDays;
            roomManuals[age][room.type] = roomManuals[age][room.type] || {};
            roomManuals[age][room.type][key] = roomManuals[age][room.type][
              key
            ] || { price: 0, count: 0 };
            roomManuals[age][room.type][key]["count"]++;
            if (personIndex === 0) {
              roomManuals[age][room.type][key]["price"] += room.manualPrice;
            }
          } else {
            if (currentRoomErrors.length === 0) {
              let partialRoomPrice = room.prices.reduce(
                (partialSum, a) => partialSum + a,
                0
              );
              if (
                resort in roomDiscountData &&
                age in roomDiscountData[resort]
              ) {
                partialRoomPrice *= roomDiscountData[resort][age];
              }
              result[age]["roomPrice"] += partialRoomPrice;
              roomDetails[age][room.type] = roomDetails[age][room.type] || {};
              const key = room.people.length + "-" + room.stayingDays;
              roomDetails[age][room.type][key] = roomDetails[age][room.type][
                key
              ] || { price: 0, count: 0 };
              roomDetails[age][room.type][key]["price"] += partialRoomPrice;
              roomDetails[age][room.type][key]["count"]++;
            } else {
              const error = currentRoomErrors[0];
              const roomType = error.split("(")[0];
              const errorType = error.split(")")[error.split(")").length - 1];
              const key = room.people.length + "-" + room.stayingDays;
              roomErrors[age][roomType] = roomErrors[age][roomType] || {};
              roomErrors[age][roomType][key] = roomErrors[age][roomType][
                key
              ] || { price: errorType, count: 0 };
              roomErrors[age][roomType][key]["count"]++;
            }
          }

          // Logger.log([age, person.snowTicket, person.snowTicket.includes(snowTicketStr)])
          if (
            age === "成人" &&
            person.snowTicket &&
            person.snowTicket.includes(snowTicketStr)
          ) {
            if (snowTicketCountMethod === "byDay") {
              snowTicketCountingDict = countSnowTicketByDay(
                snowTicketCountingDict,
                person.snowTicket,
                searchDate
              );
            } else {
              result[age]["snowTicketPrice"] +=
                Number(person.snowTicket.split(snowTicketStr)[1]) *
                snowTicketSinglePrice;
            }
          }
          person.foods.forEach((food) => {
            if (food in foodTextMappingData) {
              const foodType = foodTextMappingData[food];
              const costOfFood = foodCostData[resort][age][foodType];
              result[age]["foodPrice"] += costOfFood;
            }
          });
          if (age === "成人" && resort !== "志賀") {
            result[age]["hotSpringFee"] += 150 * room.stayingDays;
          }
        });
      });

      if (snowTicketCountMethod === "byDay") {
        snowTicketCountingDict[searchDate] =
          snowTicketCountingDict[searchDate] || 0;
        snowTicketPrice =
          snowTicketCountingDict[searchDate] * snowTicketSinglePrice;
      }
      Object.keys(result).forEach((age, ageIndex) => {
        const rowIndex = targetRowOffset + searchIndex + ageIndex;
        const detailStr = createDetailStr(roomDetails[age]);
        const manualStr = createDetailStr(roomManuals[age]);
        const errorStr = createDetailStr(roomErrors[age]);
        const finalStr = [detailStr, manualStr, errorStr]
          .filter((str) => str)
          .join("\n");
        const green = SpreadsheetApp.newTextStyle()
          .setForegroundColor("green")
          .build();
        const red = SpreadsheetApp.newTextStyle()
          .setForegroundColor("red")
          .build();
        let value = SpreadsheetApp.newRichTextValue().setText(finalStr);
        let cursor = 0;
        if (detailStr.length > 0) {
          cursor = cursor + detailStr.length + 1;
          targetSheet
            .getRange(rowIndex, targetColOffset + 1)
            .setBackground("white")
            .setFontColor("black");
        }
        if (manualStr.length > 0) {
          value = value.setTextStyle(cursor, cursor + manualStr.length, green);
          cursor = cursor + manualStr.length + 1;
          targetSheet
            .getRange(rowIndex, targetColOffset + 1)
            .setFontColor("green");
        }
        if (errorStr.length > 0) {
          value = value.setTextStyle(cursor, cursor + errorStr.length, red);
          targetSheet
            .getRange(rowIndex, targetColOffset + 1)
            .setBackground("red")
            .setFontColor("white");
        }

        value = value.build();
        const detailRange = targetSheet.getRange(rowIndex, targetColOffset);
        const roomPriceRange = targetSheet.getRange(
          rowIndex,
          targetColOffset + 1
        );
        const foodPriceRange = targetSheet.getRange(
          rowIndex,
          targetColOffset + 2
        );
        const hotSpringFeeRange = targetSheet.getRange(
          rowIndex,
          targetColOffset + 3
        );
        const manualFeeRange = targetSheet.getRange(
          rowIndex,
          targetColOffset + 4
        );
        const snowTicketPriceRange = targetSheet.getRange(
          rowIndex,
          targetColOffset + 5
        );
        const otherRange = targetSheet.getRange(rowIndex, targetColOffset + 6);
        detailRange.setRichTextValue(value);
        roomPriceRange.setValue(result[age]["roomPrice"]);
        hotSpringFeeRange.setValue(result[age]["hotSpringFee"]);
        resort === "志賀"
          ? manualFeeRange.setFormula(
              `=(-0.1)*(${roomPriceRange.getA1Notation()} + ${foodPriceRange.getA1Notation()})`
            )
          : manualFeeRange.setFormula(
              `=(-0.1)*(${roomPriceRange.getA1Notation()} + ${snowTicketPriceRange.getA1Notation()} + ${foodPriceRange.getA1Notation()})`
            );
        // snowTicketPriceRange.setValue(result[age]["snowTicketPrice"])
        snowTicketPriceRange.setValue(0);
        foodPriceRange.setValue(result[age]["foodPrice"]);
        otherRange.setValue(0);
      });

      // targetSheet.getRange(targetRowOffset+searchIndex+2, targetColOffset+1).setFormula(`=SUM(${targetSheet.getRange(targetRowOffset+searchIndex, targetColOffset+1, 2, 1).getA1Notation()})`)
      // targetSheet.getRange(targetRowOffset+searchIndex+2, targetColOffset+2).setValue(`=SUM(${targetSheet.getRange(targetRowOffset+searchIndex, targetColOffset+2, 2, 1).getA1Notation()})`)
      // targetSheet.getRange(targetRowOffset+searchIndex+2, targetColOffset+3).setValue(`=SUM(${targetSheet.getRange(targetRowOffset+searchIndex, targetColOffset+3, 2, 1).getA1Notation()})`)
      // targetSheet.getRange(targetRowOffset+searchIndex+2, targetColOffset+4).setFormula(`=SUM(${targetSheet.getRange(targetRowOffset+searchIndex+2, targetColOffset+1, 1, 3).getA1Notation()})`)
    }
  }

  // Helper Function
  function createDetailStr(obj) {
    let detailArr = [];
    Object.keys(obj)
      .sort((a, b) => a.localeCompare(b))
      .forEach((roomType) => {
        Object.keys(obj[roomType])
          .sort((a, b) => {
            const checkInPeopleA = Number(a.split("-")[0]);
            const checkInPeopleB = Number(b.split("-")[0]);
            return checkInPeopleA - checkInPeopleB;
          })
          .forEach((key) => {
            const checkInPeople = key.split("-")[0];
            const stayingDays = key.split("-")[1];
            let { price, count } = obj[roomType][key];
            price = isNaN(price) ? price : Math.floor(price);
            detailArr.push(
              `${roomType}-${checkInPeople}入住${stayingDays}晚(${count}人)$${price}`
            );
          });
      });
    return detailArr.join("\n");
  }
  function getRoomSize(cell) {
    if (cell.isPartOfMerge()) {
      const roomSize = cell.getMergedRanges()[0].getNumRows();
      return roomSize;
    }
    throw new Error("cell is not part of merge");
  }
  function getMergeCellValue(cell) {
    if (cell.isPartOfMerge()) {
      const roomNum = cell.getMergedRanges()[0].getValue();
      return roomNum;
    }
    throw new Error("cell is not part of merge");
  }
  function countSnowTicketByDay(
    previousDict,
    personSnowTicketStr,
    checkInDate
  ) {
    snowTicketNum = Number(personSnowTicketStr.split(snowTicketStr)[1]);
    for (let i = 0; i < snowTicketNum; i++) {
      let countingDate = new Date(checkInDate);
      countingDate.setDate(countingDate.getDate() + i);
      previousDict[countingDate] = previousDict[countingDate] || 0;
      previousDict[countingDate]++;
    }
    return previousDict;
  }

  function formatDate(date) {
    const year = date.getFullYear().toString().slice(-2); // Get last 2 digits of the year
    const month = String(date.getMonth() + 1).padStart(2, "0"); // Add 1 to month and pad with 0
    const day = String(date.getDate()).padStart(2, "0"); // Pad day with 0

    return `${year}/${month}/${day}`;
  }

  function formatRoomText(roomText, people) {
    let peopleText = "";
    if (people === 1) peopleText = "單";
    if (people === 2) peopleText = "雙";
    if (people === 3) peopleText = "三";
    if (people === 4) peopleText = "四";
    return roomText.replace(/\$\{[^}]+\}/g, peopleText);
  }

  function createRoomRanges(roomSheet, startRow) {
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
  function getRoomByMin(roomRanges, row) {
    for (let i = roomRanges.length - 1; i >= 0; i--) {
      if (row >= roomRanges[i].min) {
        return roomRanges[i].room;
      }
    }
    return null; // Return null for invalid rows
  }

  function getRoomPrice(startDate, stayingDays, roomType, people) {
    const startDateStr = formatDate(startDate);
    const roomTypePriceText =
      roomType in roomTypePriceShiftMap
        ? roomTypePriceShiftMap[roomType]
        : roomType;
    const rowIndex = timeColumn.findIndex((row) => row[0] === startDateStr) + 1;
    const colIndex =
      roomRows.findIndex(
        (col) => col[0] === roomTypePriceText && col[1] === people
      ) + 1;
    if (rowIndex === 0 || colIndex === 0) {
      return [`${roomType}(${people}人)無報價`];
    }
    return priceSheet
      .getRange(rowIndex, colIndex, stayingDays, 1)
      .getValues()
      .map((row, i) => {
        if (!isNaN(row[0])) return row[0];
        const date = new Date(startDate);
        date.setDate(startDate.getDate() + i);
        return `${roomType}(${people}人)(${formatDate(date)})${row[0]}`;
      });
  }
  function transpose(matrix) {
    if (!Array.isArray(matrix) || matrix.length === 0) return [];
    return matrix[0].map((_, colIndex) => matrix.map((row) => row[colIndex]));
  }
  function determineAgeGroup(person, checkInDate) {
    if (person.ageGroup) {
      return person.ageGroup === "CH" ? "孩童" : "成人";
    }
    if (!person.dateOfBirth) {
      return "成人";
    }
    const dateOfBirth = person.dateOfBirth;
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
  function getFinalRoomSheet(finalRoomSheets, searchDate) {
    const searchDateStr = formatDate(searchDate).split("/").join("");
    const finalSheet = finalRoomSheets.find(({ name }) =>
      name.startsWith(searchDateStr)
    );
    if (!finalSheet) return;
    return SpreadsheetApp.openById(finalSheet.id).getSheets()[0];
  }
  function notNull(string) {
    return string && string !== "" && string !== "X" && string !== "Ｘ";
  }
}
