const karizawaInvoice = {
  hotelInvoiceSheetId: "",       // TODO: fill in the Google Sheets ID of 01.飯店帳單-25-26 輕井澤
  hotelInvoiceSheetName: "Sheet1", // TODO: fill in the correct sheet tab name
  seasonStartYear: 2025,
  targetSheetName: "核對",        // TODO: confirm the target sheet tab name
}

const ledger =
  {
    hotelName: "苗王",
    table4SheetId: "1kHRIDsHUecxT0cod_8HTFMw-Yw0enwZlmXmha0K0L6s",
    table4SheetName: "日記帳",
    roomSheetId: "1fkUnjDDIv-dN6HJUCpFcPhanCmbvGO2Amdn3LQdFCC0",
    roomSheetName: "Naeba25-26",
    resortSheetId: "1d9vUaoT6NkwXoWLEOlSZe6xV1dI-T_Hp0J9iVBmLOm0",
    resortSheetName: "Rental表自動",
    priceSheetId: "1S6og0-YWx_tW8S52DHl3PwyWhs3J8OGSAjNoR22iHv8",
    priceSheetName: "苗王",
    snowTicketColumn: 5,
    orderNumberColumn: 14,
    resortSnowTicketColumn: 16,
    resortOrderNumberColumn: 25,
    filterResort: null,
    roomSpecifier: "苗王",
    roomNameRanges: "C10:C997",
    roomSizeRanges: "C10:C997"
  }

/**
 * HotelReconciliationFactory(hotelName, priceSheet) -> parser, logic
 * parser.parse(sheet, targetDateStr) -> RoomRecord[]
 * logic.reconsile(rooms) -> result
 */

// main.gs
function reconcileHotel() {
  const { hotelName, table4SheetId, table4SheetName, roomSheetId, roomSheetName, priceSheetId, priceSheetName, roomSpecifier, roomNameRanges, roomSizeRanges } = ledger
  const table4Sheet = SpreadsheetApp.openById(table4SheetId).getSheetByName(table4SheetName);
  const roomSheet = SpreadsheetApp.openById(roomSheetId).getSheetByName(roomSheetName);
  const priceSheet = SpreadsheetApp.openById(priceSheetId).getSheetByName(priceSheetName);
  const targetSheet = SpreadsheetApp.getActive().getSheetByName("房費核對");
  const { table4Parser, roomParser, logic } = HotelReconciliationFactory.create({hotelName, priceSheet, roomTypeMap, roomSpecifier, roomNameRanges, roomSizeRanges});
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
  const { hotelName, table4SheetId, table4SheetName, snowTicketColumn, resortSheetId, resortSheetName, resortSnowTicketColumn, orderNumberColumn, filterResort, resortOrderNumberColumn } = ledger
  const table4Sheet = SpreadsheetApp.openById(table4SheetId).getSheetByName(table4SheetName);
  const resortSheet = SpreadsheetApp.openById(resortSheetId).getSheetByName(resortSheetName);
  const targetSheet = SpreadsheetApp.getActive().getSheetByName("雪票核對");
  const { table4Parser, resortParser, logic } = SnowTicketReconciliationFactory.create(hotelName);
  const table4Tickets = table4Parser.parse(table4Sheet, filterResort);
  const table4Result = logic.reconcile(table4Tickets);
  const resultWriter = new ResultWriterSnowTicket();
  resultWriter.writeToSheet(table4Result, targetSheet, snowTicketColumn, orderNumberColumn);

  const resortTickets = resortParser.parse(resortSheet);
  const resortResult = logic.reconcile(resortTickets);
  resultWriter.writeToSheet(resortResult, targetSheet, resortSnowTicketColumn, resortOrderNumberColumn);

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

function reconcileKarizawaHotelInvoice() {
  const { hotelInvoiceSheetId, hotelInvoiceSheetName, seasonStartYear, targetSheetName } = karizawaInvoice;
  const invoiceSheet = SpreadsheetApp.openById(hotelInvoiceSheetId).getSheetByName(hotelInvoiceSheetName);
  const targetSheet = SpreadsheetApp.getActive().getSheetByName(targetSheetName);

  const parser = new KarizawaHotelInvoiceParser(seasonStartYear);
  const entries = parser.parse(invoiceSheet);

  const logic = new KarizawaHotelInvoiceReconcileLogic();
  const result = logic.reconcile(entries);

  const writer = new ResultWriterHotelInvoice(columnMap["飯店帳單"]);
  writer.writeToSheet(result, targetSheet);
}

function onOpen(){
  SpreadsheetApp.getUi()
   .createMenu('櫻雪專用功能')
   .addSeparator()
   .addItem('更新房費核對','reconcileHotel')
   .addSeparator()
   .addItem('更新雪票核對', 'reconcileSnowTicket')
   .addSeparator()
   .addItem('更新飯店帳單 (輕井澤)', 'reconcileKarizawaHotelInvoice')
   .addSeparator()
   .addToUi();
}