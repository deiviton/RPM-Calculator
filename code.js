function getData() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('Página1');
  const range = sheet.getDataRange();
  const values = range.getValues();
  
  var idealFinalDrive = idealFinalDriveCalc(values, sheet);
  var newRange = sheet.getRange('D11');
  newRange.setValue(idealFinalDrive);

  var fAutomagikerFinalDriveCalc = automagikerFinalDriveCalc(values, sheet);
  var newRange = sheet.getRange('J12');
  newRange.setValue(fAutomagikerFinalDriveCalc);
}

function idealFinalDriveCalc(values, sheet) {
  var isAutomagiker = false;
  var idealTopSpeed = values[1][1];
  var pi = values[9][1];
  var tireDiameterIn = values[8][1];
  var tireBandLength = ((pi * 2) * (tireDiameterIn / 2));
  var maxRPM = values[2][1];

  var lastGearRatio = lastGearSetter(values, isAutomagiker, sheet);
  var finalDrive = 6.11;
  var tireRotations = 0;
  var speed = 0;

  while (speed < idealTopSpeed) {
    finalDrive = finalDrive - 0.01;
    tireRotations = ((maxRPM / lastGearRatio) / finalDrive);
    speed = (((tireBandLength * (tireRotations / 60)) * 3600) * 2.54) / 100000;
  }
  return finalDrive;
}

function automagikerFinalDriveCalc(values, sheet) {
  var isAutomagiker = true;
  var idealTopSpeed = values[1][1];
  var pi = values[9][1];
  var tireDiameterIn = values[8][1];
  var tireBandLength = ((pi * 2) * (tireDiameterIn / 2));
  var maxRPM = values[2][1];
  var lastGearRatio = lastGearSetter(values, isAutomagiker, sheet);
  var finalDrive = 6.11;
  var tireRotations = 0;
  var speed = 0;

  while (speed < idealTopSpeed) {
    finalDrive = finalDrive - 0.01;
    tireRotations = ((maxRPM / lastGearRatio) / finalDrive);
    speed = (((tireBandLength * (tireRotations / 60)) * 3600) * 2.54) / 100000;
  }
  return finalDrive;
}

function lastGearSetter(values, isAutomagiker, sheet){
  let gearsNumber = values[11][1];
  var finalGear;
  
  if (isAutomagiker == false){
    switch (gearsNumber){
      case 3:
        finalGear = values[4][3];
        sheet.getRange("G11").setFormula("=($B$3/$D5)/$D$11");
        break;
      case 4:
        finalGear = values[5][3];
        sheet.getRange("G11").setFormula("=($B$3/$D6)/$D$11");
        break;
      case 5:
        finalGear = values[6][3];
        sheet.getRange("G11").setFormula("=($B$3/$D7)/$D$11");
        break;
      case 6:
        finalGear = values[7][3];
        sheet.getRange("G11").setFormula("=($B$3/$D8)/$D$11");
        break;
      case 7:
        finalGear = values[8][3];
        sheet.getRange("G11").setFormula("=($B$3/$D9)/$D$11");
        break;
      case 8:
        finalGear = values[9][3];
        sheet.getRange("G11").setFormula("=($B$3/$D10)/$D$11");
        break;
      default:
        break;
    }
  }
  else {
     switch (gearsNumber){
      case 3:
        finalGear = values[5][9];
        sheet.getRange("L12").setFormula("=($B$3/$J6)/$J$12");
        break;
      case 4:
        finalGear = values[6][9];
        sheet.getRange("L12").setFormula("=($B$3/$J7)/$J$12");
        break;
      case 5:
        finalGear = values[7][9];
        sheet.getRange("L12").setFormula("=($B$3/$J8)/$J$12");
        break;
      case 6:
        finalGear = values[8][9];
        sheet.getRange("L12").setFormula("=($B$3/$J9)/$J$12");
        break;
      case 7:
        finalGear = values[9][9];
        sheet.getRange("L12").setFormula("=($B$3/$J10)/$J$12");
        break;
      case 8:
        finalGear = values[10][9];
        sheet.getRange("L12").setFormula("=($B$3/$J11)/$J$12");
        break;
      default:
        break;
    }
  }

  return finalGear;
}