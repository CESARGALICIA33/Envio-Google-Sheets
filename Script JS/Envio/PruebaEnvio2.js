function copiarDatos2() {
  // IDs de los documentos
  var origenId = '1w5a6Rh_aH3Ab1-Rm6vdyn9ulaEVGBNZL0qt8k598O0o';  // ID del documento de origen
  var destinoId = '15F2E6Mof7aL2nhYy8R5XnSmzfxElmH2bBiZoC24mjTA'; // ID del documento de destino
  
  // Nombre de la hoja de origen
  var nombreHojaOrigen = 'Prueba';  // Cambia esto por el nombre real de la hoja
  
  // Abrir el documento Origen y Destino
  var origen = SpreadsheetApp.openById(origenId);
  var destino = SpreadsheetApp.openById(destinoId);
  
  // Obtener la hoja de origen por su nombre
  var hojaOrigen = origen.getSheetByName(nombreHojaOrigen);
  if (!hojaOrigen) {
    SpreadsheetApp.getUi().alert('No se encontró la hoja "' + nombreHojaOrigen + '" en el documento de origen.');
    return;
  }
  
  // Obtener el valor ingresado en la celda B1 de la hoja de origen
  var numero = hojaOrigen.getRange('BC7:BF10').getValue(); // Obtener el número ingresado
  SpreadsheetApp.getUi().alert('El valor de numero es: ' + numero);
  
  // Validar que el número esté en el rango válido (del 2 al 50)
  if (numero < 2 || numero > 1000) {
    SpreadsheetApp.getUi().alert('Ingrese un número válido en la celda BC7:BF10');
    return;
  }
  
  // Definir el rango de celdas según el número ingresado
  var rangoCeldas = 'AL' + numero + ':BR' + numero;
  
  // Obtener los datos del rango de celdas especificado en la hoja de origen
  var datos = hojaOrigen.getRange(rangoCeldas).getValues();
  
  // Obtener la hoja de destino
  var hojaDestino = destino.getSheetByName('Ventas'); 
  if (!hojaDestino) {
    SpreadsheetApp.getUi().alert('No se encontró la hoja "Prueba" en el documento de destino.');
    return;
  }
  
  // Obtener la última fila con datos en la hoja de destino
  var ultimaFila = hojaDestino.getLastRow();
  
  // Definir el rango de celdas donde se pegarán los nuevos datos de manera incremental
  var rangoPegado = hojaDestino.getRange(ultimaFila + 1, 1, datos.length, datos[0].length);
  
  // Pegar los datos en la hoja de destino de manera incremental
  rangoPegado.setValues(datos);

  // Limpiar el contenido de la celda B1 en la hoja de origen
  hojaOrigen.getRange('BC7:BF10').clearContent();
  
  // Mensaje de confirmación
  SpreadsheetApp.getUi().alert('Datos enviados exitosamente.');
}