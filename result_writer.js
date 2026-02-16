class ResultWriter {

  constructor(columnMap) {
    this.columnMap = columnMap
  }

  writeToSheet(result, targetSheet) {

    const targetValues = targetSheet.getRange("A11:B").getValues();
    targetValues.forEach(([date, item], index) => {
      const formattedDate = formatDate(date);
      const rowIndex = 11 + index;

      if (formattedDate in result && item in result[formattedDate]) {
        const dateResult = result[formattedDate];

        // Write main value
        Object.keys(this.columnMap).forEach(key => {
          if (item !== "其他" && key === "手續費") {
            targetSheet.getRange(rowIndex, this.columnMap[key]).setFormula(`=(-0.1)*(${numberToColumnLetter(this.columnMap["房費"])}${rowIndex})`)
            return
          }
          if (!(key in dateResult[item])) return
          if (key === "房型") {
            const roomText = this.formatRoomTypes(dateResult[item].房型);
            const errorList = dateResult[item].房型.錯誤列表 || [];
            const errorText = errorList.length ? '\n' + errorList.join('\n') : '';

            if (errorText) {
              const errorColor = errorText.includes("無報價") ? "#0000FF" : "#FF0000";
              const richText = SpreadsheetApp.newRichTextValue()
                .setText(roomText + errorText)
                .setTextStyle(roomText.length, roomText.length + errorText.length,
                  SpreadsheetApp.newTextStyle().setForegroundColor(errorColor).build()
                )
                .build();
              targetSheet.getRange(rowIndex, this.columnMap.房型).setRichTextValue(richText);
            } else {
              targetSheet.getRange(rowIndex, this.columnMap.房型).setValue(roomText);
            }
            return
          }

          targetSheet.getRange(rowIndex, this.columnMap[key]).setValue(dateResult[item][key]);
        })


        
      }
      else if (item !== "當日總計" && item !== "系統問題") {
        targetSheet.getRange(rowIndex, this.columnMap.房型, 1, 7).clearContent();
      }

    });
  }

  formatRoomTypes(roomTypeObj) {
    if (!roomTypeObj || typeof roomTypeObj !== "object") return "";
    return Object.entries(roomTypeObj)
      .sort(([a_roomType, a] , [b_roomType, b]) => a.people - b.people)
      .sort(([a_roomType] , [b_roomType]) => a_roomType.localeCompare(b_roomType))
      .map(([roomType, { people, price, totalNights }]) => {
        if (roomType === "錯誤列表") return "";
        
        // Check if this is a special price room type (already formatted)
        // Special price format: {orderNumber}{name}{people}人${price}|uniqueId
        if (roomType.includes("人$")) {
          // Strip the unique identifier (everything after |) for display
          const displayRoomType = roomType.split("|")[0];
          return displayRoomType; // Already formatted, return without unique ID
        }
        
        // Regular room type formatting
        return `${roomType}(${people/totalNights}人)$${price}`;
      })
      .join('\n');
  }
}

class ResultWriterHotelInvoice {
  constructor(columnMap) {
    this.columnMap = columnMap;
  }

  // result: { dateStr: { 成人: { 房費, 入湯稅, 手数料, 其他, 總價 } } }
  writeToSheet(result, targetSheet) {
    const targetValues = targetSheet.getRange("A11:B").getValues();

    targetValues.forEach(([date, item], index) => {
      const formattedDate = formatDate(date);
      const rowIndex = 11 + index;

      if (formattedDate in result && item in result[formattedDate]) {
        const dayRow = result[formattedDate][item];

        if ("房費" in this.columnMap && "房費" in dayRow) {
          targetSheet.getRange(rowIndex, this.columnMap["房費"]).setValue(dayRow["房費"]);
        }
        if ("入湯稅" in this.columnMap && "入湯稅" in dayRow) {
          targetSheet.getRange(rowIndex, this.columnMap["入湯稅"]).setValue(dayRow["入湯稅"]);
        }
        if ("手続費" in this.columnMap && "手数料" in dayRow) {
          targetSheet.getRange(rowIndex, this.columnMap["手続費"]).setValue(dayRow["手数料"]);
        }
        if ("其他" in this.columnMap && "其他" in dayRow) {
          targetSheet.getRange(rowIndex, this.columnMap["其他"]).setValue(dayRow["其他"]);
        }
        if ("總計" in this.columnMap && "總價" in dayRow) {
          targetSheet.getRange(rowIndex, this.columnMap["總計"]).setValue(dayRow["總價"]);
        }
      }
    });
  }
}

class ResultWriterSnowTicket {

  writeToSheet(result, targetSheet, column, orderNumberColumn) {
    const targetValues = targetSheet.getRange("A4:B").getValues();
    targetValues.forEach(([date, item], index) => {
      const formattedDate = formatDate(date);
      const rowIndex = 4 + index;

      if (formattedDate in result && item in result[formattedDate]) {
        const dateResult = result[formattedDate];

        // Write main value
        targetSheet.getRange(rowIndex, column).setValue(dateResult[item]);  
        if (orderNumberColumn) {
          targetSheet.getRange(rowIndex, orderNumberColumn).setValue(Array.from(dateResult.票券編號[item]).join("\n"))
        }      
      }
      else if (item !=="當日總計") {
        targetSheet.getRange(rowIndex, column).setValue(0);        
      }


    });
  }

}