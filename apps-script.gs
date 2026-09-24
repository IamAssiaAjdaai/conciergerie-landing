function doPost(e) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheets()[0];
  const data = JSON.parse(e.postData.contents);

  if (sheet.getLastRow() === 0) {
    sheet.appendRow([
      'submittedAt',
      'city',
      'district',
      'propertyType',
      'bedrooms',
      'listed',
      'name',
      'phone',
      'source'
    ]);
  }

  sheet.appendRow([
    data.submittedAt || '',
    data.city || '',
    data.district || '',
    data.propertyType || '',
    data.bedrooms || '',
    data.listed || '',
    data.name || '',
    data.phone || '',
    data.source || 'Landing page'
  ]);

  return ContentService
    .createTextOutput(JSON.stringify({ ok: true }))
    .setMimeType(ContentService.MimeType.JSON);
}
