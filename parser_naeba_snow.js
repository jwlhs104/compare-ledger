class NaebaSnowTicketParser extends BaseSnowTicketParser {
  parse(sheet, filterResort) {
    const data = sheet.getRange("A13:FG3000").getValues();
    const snowTickets = [];

    for (let i = 0; i < data.length; i++) {
      const row = data[i];
      const orderNumber = row[getColumnIndex("D")];
      const rawAge = row[getColumnIndex("BW")];
      const ageType = rawAge === "成人" ? "成人" : "孩童";

      // Parse each day's snow ticket data (6 days max)
      for (let day = 0; day < 6; day++) {
        const baseCol = getColumnIndex("DB") + day * 8;
        const date = row[baseCol];
        const snowTicketData = row[baseCol + 1]; // Snow ticket is at offset 4 in each 8-column day block

        if (date instanceof Date && snowTicketData && typeof snowTicketData === 'string') {
          // Check if snow ticket string matches the pattern
          if (notNull(snowTicketData)) {
            const formattedDate = formatDate(date);
            snowTickets.push(new SnowTicketRecord(
              formattedDate,
              ageType,
              orderNumber
            )); 
          }
        }
      }
    }

    return snowTickets;
  }
}

class NaebaResortSnowTicketParser extends BaseSnowTicketParser {
  parse(sheet) {
    // Resort sheet parsing - snow ticket rentals from the resort
    const data = sheet.getRange("A2:Z1000").getValues();
    const snowTickets = [];

    for (let i = 0; i < data.length; i++) {
      const row = data[i];
      const orderNumber = row[1]; // Order number at column B (index 1)
      const firstDate = row[2]; // First date at column C (index 2)
      const snowTicketCount = row[5]; // Snow ticket count at column F (index 5)
      const snowTicketStr = row[6]; // Snow ticket string at column G (index 6)

      // Only process if snowTicketCount > 0
      if (!isNaN(snowTicketCount) && snowTicketCount > 0 && firstDate instanceof Date) {
        // Determine age type from snowTicketStr
        let ageType;
        if (snowTicketStr && typeof snowTicketStr === 'string') {
          if (snowTicketStr.includes("成")) {
            ageType = "成人";
          } else if (snowTicketStr.includes("孩")) {
            ageType = "孩童";
          } else {
            throw new Error(`無法判斷年齡類型: ${snowTicketStr} at row ${i + 2}`);
          }
        } else {
          throw new Error(`雪票字串為空 at row ${i + 2}`);
        }

        // Create one record per day for snowTicketCount days
        for (let dayOffset = 0; dayOffset < snowTicketCount; dayOffset++) {
          const currentDate = new Date(firstDate);
          currentDate.setDate(firstDate.getDate() + dayOffset);
          const formattedDate = formatDate(currentDate);

          snowTickets.push(new SnowTicketRecord(
            formattedDate,
            ageType,
            orderNumber
          ));
        }
      }
    }

    return snowTickets;
  }
}
