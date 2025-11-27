class NaebaSnowTicketReconciliationLogic {
  constructor() {
    this.snowTicketSinglePrice = 4600; // Naeba snow ticket price per ticket
  }

  reconcile(snowTickets) {
    const result = {};

    snowTickets.forEach((ticket) => {
      const date = ticket.issueDate;
      const ageType = ticket.type;
      const orderNumber = ticket.orderNumber;

      // Initialize date entry if not exists
      if (!result[date]) {
        result[date] = {
          "成人": 0,
          "孩童": 0,
          票券編號: {
            "成人": new Set(),
            "孩童": new Set()
          }
        };
      }

      // Add price for this ticket (each ticket is ¥4600)
      result[date][ageType] += this.snowTicketSinglePrice;

      // Add order number to the set for tracking
      if (orderNumber) {
        result[date].票券編號[ageType].add(orderNumber);
      }
    });

    return result;
  }
}
