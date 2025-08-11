class HotelReconciliationFactory {
  static create(hotelName, priceSheet, roomTypeMap, roomSpecifier, searchRanges) {
    switch (hotelName) {
      case '苗王':
        return {
          table4Parser: new NaebaInvoiceParser(),
          roomParser: new NaebaRoomInvoiceParser(roomSpecifier, searchRanges),
          logic: new NaebaReconciliationLogic(hotelName, priceSheet, roomTypeMap)
        };
      default:
        throw new Error(`尚未支援飯店: ${hotelName}`);
    }
  }
}

class SnowTicketReconciliationFactory {
  static create(hotelName) {
    switch (hotelName) {
      case '苗王':
        return {
          table4Parser: new NaebaSnowTicketParser(),
          resortParser: new NaebaResortSnowTicketParser(),
          logic: new NaebaSnowTicketReconciliationLogic()
        };
      default:
        throw new Error(`尚未支援飯店: ${hotelName}`);
    }
  }
}