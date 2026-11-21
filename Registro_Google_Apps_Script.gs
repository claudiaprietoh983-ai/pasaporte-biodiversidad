// Pegar en Extensiones > Apps Script de una hoja de cálculo nueva.
// Configurar e implementar como aplicación web: ejecutar como YO, acceso CUALQUIERA.
const DESTINATARIO = 'claudiaprietoh983@gmail.com';
const NOMBRE_HOJA = 'Certificados';
function doPost(e) {
  try {
    const d = JSON.parse(e.postData.contents);
    if (!d.nombre || !d.empresa || d.estado !== 'Certificado') throw new Error('Registro incompleto');
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(NOMBRE_HOJA) || ss.insertSheet(NOMBRE_HOJA);
    if (sheet.getLastRow() === 0) sheet.appendRow(['Fecha recepción','Nombre','Empresa','Fecha juego','Estado','Sellos']);
    const safe = v => "'" + String(v || '').replace(/[\r\n]/g,' ').slice(0,250);
    sheet.appendRow([new Date(),safe(d.nombre),safe(d.empresa),safe(d.fecha),safe(d.estado),safe(d.sellos)]);
    // Notificación individual al completar el juego.
    MailApp.sendEmail({to:DESTINATARIO,subject:'Pasaporte Biodiversidad: nuevo certificado',body:'Se ha registrado un participante certificado.\nNombre: '+d.nombre+'\nEmpresa: '+d.empresa+'\nFecha: '+d.fecha+'\nRevisa la hoja de cálculo para el listado completo.'});
    return ContentService.createTextOutput('OK');
  } catch(err) {return ContentService.createTextOutput('ERROR: '+err.message)}
}
// Para enviar el listado acumulado manualmente desde Apps Script.
function enviarArchivoCSV() {
  const ss=SpreadsheetApp.getActiveSpreadsheet();const sheet=ss.getSheetByName(NOMBRE_HOJA);
  if(!sheet) throw new Error('No hay certificados todavía');
  const csv=sheet.getDataRange().getDisplayValues().map(row=>row.map(v=>'"'+String(v).replace(/"/g,'""')+'"').join(',')).join('\r\n');
  MailApp.sendEmail({to:DESTINATARIO,subject:'Listado de certificados - Pasaporte de la Biodiversidad',body:'Adjunto el listado de participantes certificados.',attachments:[Utilities.newBlob('\ufeff'+csv,'text/csv','participantes_certificados.csv')]});
}
