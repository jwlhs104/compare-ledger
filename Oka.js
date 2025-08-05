function scanRoomRoot() {
  const okaSheet = SpreadsheetApp.getActive().getSheetByName("oka 24-25房表")
  const workingValues = okaSheet.getRange("A5:A122").getValues();
  const offset = 5;
  const result = []
  workingValues.forEach((row, index) => {
    if (row[0]) {
      result.push(offset+index)
    }
  })
  result.sort((a, b) => a-b)

  const cache = CacheService.getScriptCache();
  cache.put("room_scan_result", JSON.stringify(result), 600);
}

function testCountCheckIn() {
  const testValues = SpreadsheetApp.getActive().getSheetByName("oka 24-25房表").getRange("LR5:LS122").getValues()
  countCheckIn(testValues)
}

 /**
* @param {calculateRange} range to calculate
* @return {number}
* @customfunction
*/
function countCheckIn(calculateRange) {
  const cache = CacheService.getScriptCache();
  let roomRoots = cache.get("room_scan_result");
  const offset = 5;
  if (!roomRoots) {
    scanRoomRoot();
    roomRoots = cache.get("room_scan_result");
  }
  
  roomRoots = roomRoots ? JSON.parse(roomRoots) : [];
  let count = 0;
  calculateRange.forEach((row, index) => {
    if (safeStartsWitch(row[1], "毒")) {
      if (safeStartsWitch(row[0], "入")) {
        count ++;
      }
      else if (!row[0]) {
        const root = searchRoot(offset+index, roomRoots)
        const rootValue = calculateRange[root-offset][0]
        if (safeStartsWitch(rootValue, "入")) {
          count++;
        }
      }
    }
  })

  return count;

  function searchRoot(number, arr) {
    for (let i = arr.length-1; i>=0; i--) {
      const rootNumber = arr[i]
      if (number >= rootNumber) {
        return rootNumber
      }
    }
    throw new Error("Can't find root")
  }

  function safeStartsWitch(value, key) {
    return value && typeof value === "string" && value.startsWith(key)
  }
}