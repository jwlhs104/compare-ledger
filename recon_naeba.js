class NaebaReconciliationLogic extends BaseReconciliationLogic {
  constructor(priceSheet, roomTypeMap) {
    super();
    this.priceSheet = priceSheet;
    this.roomTypeMap = roomTypeMap
  }

  createPriceTable() {
    const priceTable = {}
    const priceValues = this.priceSheet.getRange("F4:DE127").getValues();
    const dateValues= this.priceSheet.getRange("A7:A127").getValues();
    transpose(priceValues).forEach(col => {
      const roomType = col[0];
      const people = col[1];
      const ageGroup = col[2];
      priceTable[roomType] = priceTable[roomType] || {}
      priceTable[roomType][people] = priceTable[roomType][people] || {}
      priceTable[roomType][people][ageGroup] = priceTable[roomType][people][ageGroup] || {}

      dateValues.forEach(([date], index) => {
        const formattedDate = formatDate(date);
        priceTable[roomType][people][ageGroup][formattedDate] = col[index+3];
      })
    })
    this.priceTable = priceTable;
  }

  getRoomPrice(roomTypeText, peopleAmount, ageGroup, date) {
    let maxPeople = Math.max(...Object.keys(this.priceTable[roomTypeText]))
    peopleAmount = peopleAmount > maxPeople ? maxPeople : peopleAmount;
    const rawPrice = this.priceTable?.[roomTypeText]?.[peopleAmount]?.["成人"]?.[date];
    return rawPrice;
  }

  createStayRecords(rooms) {
    const roomMap = new Map();  // key = roomNumber + firstDate

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
    const result = {};

    const stays = this.createStayRecords(rooms);
    stays.forEach(stay => {
      const stayRooms = stay.roomRecords;
      const totalNights = stay.getTotalNights();
      const calDate = stay.firstDate;
      result[calDate] = result[calDate] || {
        成人: { 房費: 0, 餐食: 0, 入湯稅: 0, 房型: {}},
        孩童: { 房費: 0, 餐食: 0, 入湯稅: 0, 房型: {}},
        官網: { 房費: 0, 餐食: 0, 入湯稅: 0, 房型: {}},
        網路訂房: { 房費: 0, 餐食: 0, 入湯稅: 0, 房型: {}},
      }
      stayRooms.forEach(room => {
        // const roomTypeText = this.roomTypeMap[room.roomType];
        const roomTypeText = room.roomType
        if (!roomTypeText || roomTypeText.includes("東品") || roomTypeText.includes("不佔床")) {
          return
        }
        room.people.forEach((person, index) => {
          const mealType = person.mealType;
          let price;
          let priceError = null;

          if (!roomTypeText) {
            // priceError = `錯誤: 找不到對應房型 ${room.roomType}`;
            return
          } else {
            try {
              const rawPrice = this.getRoomPrice(roomTypeText, room.people.length, person.ageGroup, room.date)
              if (typeof rawPrice === 'number') {
                price = Math.round(rawPrice);
              } else if (typeof rawPrice === 'string') {
                priceError = `${roomTypeText}-${room.people.length}入住-無報價`;
              } else {
                priceError = `${roomTypeText}-${room.people.length}入住-查無房型報價`;
              }
            } catch (e) {
              priceError = `錯誤: 價格查詢例外 - ${e.message}`;
            }
          }

          const dateResult = result[calDate];
          const ageGroupData = dateResult[person.ageGroup];
          if (!ageGroupData) {
            return
          }
          const roomType = `${room.roomType}-${room.people.length}入住${totalNights}晚`
          ageGroupData.房型[roomType] = ageGroupData.房型[roomType] || {
            people: 0,
            price: 0,
            totalNights
          };

          ageGroupData.房型[roomType].people++;
          const dayDiff = dateDifferenceInDays(room.date, room.firstDate) + 1
          ageGroupData.房型[roomType].day = ageGroupData.房型[roomType].day > dayDiff ? ageGroupData.房型[roomType].day : dayDiff;


          if (typeof price === 'number') {
            ageGroupData.房型[roomType].price += price;
            ageGroupData.房費 += price;
          }

          // Log error if price lookup failed
          if (priceError) {
            
            ageGroupData.房型.錯誤列表 = ageGroupData.房型.錯誤列表 || [];
            if (!ageGroupData.房型.錯誤列表.includes(priceError)) {
              ageGroupData.房型.錯誤列表.push(priceError);
            }
          }

          if (mealType === '早餐' || mealType === '早晚餐') {
            ageGroupData.餐食 += foodCostData.雫石[person.ageGroup].早餐;
          }
          if (mealType === '晚餐' || mealType === '早晚餐') {
            ageGroupData.餐食 += foodCostData.雫石[person.ageGroup].晚餐;
          }

          if (person.ageGroup === "成人") {
            ageGroupData.入湯稅 += 150;
          }

        });
        
      })
    })

    return result;
  }

  checkAU(rooms) {
    this.createPriceTable();
    const result = {};
    const stays = this.createStayRecords(rooms);
    stays.forEach(stay => {
      const stayRooms = stay.roomRecords;
      const totalNights = stay.getTotalNights();
      stayRooms.forEach(room => {
        const roomTypeText = this.roomTypeMap[room.roomType];
        if (!roomTypeText) return
        room.people.forEach((person, index) => {
          let price;
          let key = person.orderNumber + "_" + person.name;
          if (person.continuation !== 1) key += `_${person.continuation}`
          result[key] = result[key] || {
            "AU價格": person.auPrice,
            "計算價格": 0
          }

          const rawPrice = this.getRoomPrice(roomTypeText, room.people.length, person.ageGroup, room.date)
          if (typeof rawPrice === 'number' && result[key]["計算價格"] !== "無報價") {
            price = Math.round(rawPrice);
            result[key]["計算價格"] += price
          } else {
            result[key]["計算價格"] = "無報價"
          }
        })
      })
    })
    const tableOutput = []
    Object.keys(result).forEach(key => {
      const auPrice = result[key]["AU價格"]
      const calPrice = result[key]["計算價格"]
      if (auPrice+calPrice !== 0) {
        tableOutput.push([key, auPrice, calPrice, auPrice+calPrice])
      }
    })
    return tableOutput;
  }
}