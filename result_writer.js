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
      .map(([roomType, { people, price, totalNights }]) => roomType === "錯誤列表" ? "" : `${roomType}(${people/totalNights}人)$${price}`)
      .join('\n');
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