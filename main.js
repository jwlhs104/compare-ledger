const ledger =
  {
    hotelName: "苗王",
    table4SheetId: "1kHRIDsHUecxT0cod_8HTFMw-Yw0enwZlmXmha0K0L6s",
    table4SheetName: "日記帳",
    roomSheetId: "1fkUnjDDIv-dN6HJUCpFcPhanCmbvGO2Amdn3LQdFCC0",
    roomSheetName: "Naeba25-26",
    resortSheetId: "1tZmH4KPmjirHk6HJpJma4MtsYJ48SIgTrRzdthpfwrk",
    resortSheetName: "Rental表自動",
    priceSheetId: "1S6og0-YWx_tW8S52DHl3PwyWhs3J8OGSAjNoR22iHv8",
    priceSheetName: "苗王",
    snowTicketColumn: 5,
    orderNumberColumn: 14,
    resortSnowTicketColumn: 16
  }

/**
 * HotelReconciliationFactory(hotelName, priceSheet) -> parser, logic
 * parser.parse(sheet, targetDateStr) -> RoomRecord[]
 * logic.reconsile(rooms) -> result
 */

// main.gs
function reconcileHotel() {
  const { hotelName, table4SheetId, table4SheetName, roomSheetId, roomSheetName, priceSheetId, priceSheetName } = ledger
  const table4Sheet = SpreadsheetApp.openById(table4SheetId).getSheetByName(table4SheetName);
  const roomSheet = SpreadsheetApp.openById(roomSheetId).getSheetByName(roomSheetName);
  const priceSheet = SpreadsheetApp.openById(priceSheetId).getSheetByName(priceSheetName);
  const targetSheet = SpreadsheetApp.getActive().getSheetByName("房費核對");
  const { table4Parser, roomParser, logic } = HotelReconciliationFactory.create(hotelName, priceSheet, roomTypeMap);
  const table4Rooms = table4Parser.parse(table4Sheet);
  const table4Result = logic.reconcile(table4Rooms);
  const table4Writer = new ResultWriter(columnMap["表4"]);
  table4Writer.writeToSheet(table4Result, targetSheet);
  const rooms = roomParser.parse(roomSheet);
  const result = logic.reconcile(rooms);
  const roomWriter = new ResultWriter(columnMap["房表"]);
  roomWriter.writeToSheet(result, targetSheet);
}

function reconcileSnowTicket() {
  const { hotelName, table4SheetId, table4SheetName, snowTicketColumn, resortSheetId, resortSheetName, resortSnowTicketColumn, orderNumberColumn, filterResort } = ledger
  const table4Sheet = SpreadsheetApp.openById(table4SheetId).getSheetByName(table4SheetName);
  const resortSheet = SpreadsheetApp.openById(resortSheetId).getSheetByName(resortSheetName);
  const targetSheet = SpreadsheetApp.getActive().getSheetByName("苗王雪票核對");
  const { table4Parser, resortParser, logic } = SnowTicketReconciliationFactory.create(hotelName);
  const table4Tickets = table4Parser.parse(table4Sheet, filterResort);
  const table4Result = logic.reconcile(table4Tickets);
  const resultWriter = new ResultWriterSnowTicket();
  resultWriter.writeToSheet(table4Result, targetSheet, snowTicketColumn, orderNumberColumn);

  const resortTickets = resortParser.parse(resortSheet);
  const resortResult = logic.reconcile(resortTickets);
  resultWriter.writeToSheet(resortResult, targetSheet, resortSnowTicketColumn);

}

function au_check() {
  const { hotelName, table4SheetId, table4SheetName, roomSheetId, roomSheetName, priceSheetId, priceSheetName } = ledger
  const table4Sheet = SpreadsheetApp.openById(table4SheetId).getSheetByName(table4SheetName);
  const priceSheet = SpreadsheetApp.openById(priceSheetId).getSheetByName(priceSheetName);
  const { table4Parser, logic } = HotelReconciliationFactory.create(hotelName, priceSheet, roomTypeMap);
  const table4Rooms = table4Parser.parse(table4Sheet);
  const table4Result = logic.checkAU(table4Rooms);
  const targetSheet = SpreadsheetApp.getActive().getSheetByName("出團帳vs房表帳單比對")
  targetSheet.getRange(3, 1, targetSheet.getLastRow(), targetSheet.getLastColumn()).clearContent()
  if (table4Result.length > 0) {
    targetSheet.getRange(3, 1, table4Result.length, table4Result[0].length).setValues(table4Result)
  }
}

function onOpen(){
  SpreadsheetApp.getUi()
   .createMenu('櫻雪專用功能')
   .addSeparator()
   .addItem('更新房費核對','reconcileHotel')  
   .addSeparator()
   .addToUi();
}