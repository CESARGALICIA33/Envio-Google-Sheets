function onEdit(e) {
    const sheet = e.source.getActiveSheet();
    const sheetName = sheet.getName();
    const targetSheetName = "Folios"; 
  
    // Verificar si la edición está en la hoja correcta
    if (sheetName !== targetSheetName) return;
  
    const range = e.range;
    const column = range.getColumn();
    const row = range.getRow();
  
    // Verificar si la edición está en las columnas A, B o C
    if (column < 1 || column > 3) return;
  
    // Verificar si la celda editada ya contiene un folio y si no fue eliminado
    const currentValue = range.getValue();
    if (currentValue && currentValue.startsWith("IMX") && e.value !== "") {
      return;
    }
  
    // Obtener el prefijo y el siguiente folio no repetido
    const prefix = "IMX";
    let nextFolioNumber = getNextFolioNumber(sheet);
    let nextFolio = prefix + nextFolioNumber.toString().padStart(4, '0');
  
    // Verificar si el folio ya existe y encontrar uno que no esté repetido
    while (isFolioUsed(sheet, nextFolio)) {
      nextFolioNumber++;
      nextFolio = prefix + nextFolioNumber.toString().padStart(4, '0');
    }
  
    // Asignar el nuevo folio a la celda editada
    sheet.getRange(row, column).setValue(nextFolio);
  }
  
  function getNextFolioNumber(sheet) {
    const prefix = "IMX";
    let lastFolio = 8000; // Número inicial, cambiar si es necesario
    const dataA = sheet.getRange("A:A").getValues();
    const dataB = sheet.getRange("B:B").getValues();
    const dataC = sheet.getRange("C:C").getValues();
  
    // Buscar el folio más alto en cada columna
    [dataA, dataB, dataC].forEach(data => {
      data.forEach(cell => {
        if (cell[0] && cell[0].startsWith(prefix)) {
          const number = parseInt(cell[0].substring(prefix.length), 10);
          if (!isNaN(number) && number > lastFolio) {
            lastFolio = number;
          }
        }
      });
    });
  
    return lastFolio + 1;
  }
  
  function isFolioUsed(sheet, folio) {
    const dataA = sheet.getRange("A:A").getValues();
    const dataB = sheet.getRange("B:B").getValues();
    const dataC = sheet.getRange("C:C").getValues();
  
    // Verificar si el folio ya está usado en alguna de las columnas
    const allData = dataA.concat(dataB).concat(dataC);
    for (let i = 0; i < allData.length; i++) {
      if (allData[i][0] === folio) {
        return true;
      }
    }
  
    return false;
  }
  