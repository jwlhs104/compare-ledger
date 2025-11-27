class NaebaSnowTicketParser extends BaseSnowTicketParser {
  parse(sheet, filterResort) {
    const data = sheet.getRange("A13:FG3000").getValues();
    const snowTickets = [];
    const snowTicketStr = "I4YLE-";

    for (let i = 0; i < data.length; i++) {
      const row = data[i];
      const orderNumber = row[getColumnIndex("D")];
      const rawAge = row[getColumnIndex("BW")];
      const ageType = rawAge === "成人" ? "成人" : "孩童";

      // Parse each day's snow ticket data (6 days max)
      for (let day = 0; day < 6; day++) {
        const baseCol = getColumnIndex("DB") + day * 8;
        const date = row[baseCol];
        const snowTicketData = row[baseCol + 4]; // Snow ticket is at offset 4 in each 8-column day block

        if (date instanceof Date && snowTicketData && typeof snowTicketData === 'string') {
          // Check if snow ticket string matches the pattern
          if (snowTicketData.includes(snowTicketStr)) {
            const formattedDate = formatDate(date);
            const ticketCountStr = snowTicketData.split(snowTicketStr)[1];
            const ticketCount = Number(ticketCountStr);

            if (!isNaN(ticketCount) && ticketCount > 0) {
              // Create one record per ticket
              for (let t = 0; t < ticketCount; t++) {
                snowTickets.push(new SnowTicketRecord(
                  formattedDate,
                  ageType,
                  orderNumber
                ));
              }
            }
          }
        }
      }
    }

    return snowTickets;
  }
}

class NaebaResortSnowTicketParser extends BaseSnowTicketParser {
  parse(sheet) {
    // Resort sheet parsing - typically snow ticket rentals from the resort
    const data = sheet.getRange("A2:Z1000").getValues();
    const snowTickets = [];
    const snowTicketStr = "I4YLE-";

    for (let i = 0; i < data.length; i++) {
      const row = data[i];
      const date = row[0]; // Date column
      const orderNumber = row[1]; // Order number column
      const snowTicketData = row[7]; // Assuming snow ticket data is in column H (index 7)
      const ageType = "成人"; // Resort tickets typically for adults

      if (date instanceof Date && snowTicketData && typeof snowTicketData === 'string') {
        if (snowTicketData.includes(snowTicketStr)) {
          const formattedDate = formatDate(date);
          const ticketCountStr = snowTicketData.split(snowTicketStr)[1];
          const ticketCount = Number(ticketCountStr);

          if (!isNaN(ticketCount) && ticketCount > 0) {
            // Create one record per ticket
            for (let t = 0; t < ticketCount; t++) {
              snowTickets.push(new SnowTicketRecord(
                formattedDate,
                ageType,
                orderNumber
              ));
            }
          }
        }
      }
    }

    return snowTickets;
  }
}
