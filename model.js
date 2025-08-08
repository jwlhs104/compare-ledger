class StayRecord {
  constructor(firstDate, roomNumber) {
    this.firstDate = firstDate;
    this.roomNumber = roomNumber;
    this.roomRecords = [];  // 存放所有這次入住的 RoomRecord
  }

  addRoomRecord(roomRecord) {
    this.roomRecords.push(roomRecord);
  }

  getTotalNights() {
    return this.roomRecords.length;
  }

  getTotalCost() {
    return this.roomRecords.reduce((sum, room) => {
      return sum + room.getCost(); // 你可在 RoomRecord 裡實作 getCost()
    }, 0);
  }

  getPeopleSignature() {
    // 可用於比對是否同一組人，作為分群條件之一
    return this.roomRecords[0].people.map(p => p.name).sort().join(',');
  }
}


class RoomRecord {
  constructor(roomNumber, date, roomType, firstDate, firstDateIndex) {
    this.roomNumber = roomNumber;
    this.firstDateIndex = firstDateIndex;
    this.date = date;
    this.roomType = roomType;
    this.firstDate = firstDate;
    this.people = [];
  }

  addPerson({name, mealType, ageGroup, bonusMeal, orderNumber, auPrice, continuation}) {
    const person = new PersonRecord(name, mealType, ageGroup, bonusMeal, orderNumber, auPrice, continuation)
    this.people.push(person)
  }

}

class PersonRecord {
  constructor(name, mealType, ageGroup, bonusMeal, orderNumber, auPrice, continuation) {
    this.name = name
    this.mealType = mealType
    this.ageGroup = ageGroup
    this.bonusMeal = bonusMeal
    this.orderNumber = orderNumber
    this.auPrice = auPrice || 0
    this.continuation = continuation
  }
}

class SnowTicketRecord {
  constructor(issueDate, type, orderNumber) {
    this.issueDate = issueDate;
    this.type = type;
    this.orderNumber = orderNumber;
  }
  updateType() {
    if (this.type >= 3) return true;
    this.type +=1;
    return false;
  }
}