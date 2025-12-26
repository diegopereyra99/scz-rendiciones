function showCompletionAlert(result) {
  const ui = SpreadsheetApp.getUi();

  const count =
    result && typeof result.count === 'number'
      ? result.count
      : 0;

  ui.alert(
    'Proceso finalizado',
    'Los archivos fueron procesados correctamente.\n\n' +
    'Comprobantes detectados: ' + count + '\n\n' +
    'Revisá las filas resaltadas y marcá el tick ✅ cuando cada fila esté pronta.',
    ui.ButtonSet.OK
  );
}

function showErrorAlert(message) {
  const ui = SpreadsheetApp.getUi();

  const msg =
    message && String(message).trim()
      ? String(message)
      : 'Ocurrió un error inesperado durante el procesamiento.';

  ui.alert(
    'Error',
    msg,
    ui.ButtonSet.OK
  );
}