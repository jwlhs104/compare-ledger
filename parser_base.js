class BaseInvoiceParser {
  parse(sheet) {
    throw new Error("parse() 必須由子類別實作");
  }
}

class BaseSnowTicketParser {
  parse(sheet) {
    throw new Error("parse() 必須由子類別實作");   
  }
}