function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('CNC')
    .addItem('Pobierz pliki dla zestawu (z kolorami)...', 'promptAndDownloadWithColors')
    .addItem('Pobierz pliki dla modułu...', 'promptAndDownloadModule')
    .addItem('Pobierz listę elementów dla zestawu...', 'promptAndCreateElementList')
    .addItem('Pobierz listę elementów dla modułu...', 'promptAndCreateElementListModule')
    .addToUi();

  ui.createMenu('Sync')
    .addItem('Ustaw / edytuj link dla elementu...', 'promptAndSyncLink')
    .addItem('Porównaj linki (SyncLinks)', 'promptAndCompareLinks')
    .addItem('Masowe sprawdzenie linków', 'massCheckAndFixLinks')
    .addToUi();
}

function promptAndDownloadWithColors() {
  const ui = SpreadsheetApp.getUi();
  const resp = ui.prompt('Pobierz pliki DXF z kolorami', 'Podaj numer zestawu (np. P1608):', ui.ButtonSet.OK_CANCEL);
  if (resp.getSelectedButton() !== ui.Button.OK) return;
  const setId = resp.getResponseText().trim();
  if (!setId) {
    ui.alert('Nie podano numeru zestawu.');
    return;
  }

  try {
    downloadSetFilesWithColors(setId);
  } catch (e) {
    ui.alert('Błąd', 'Wystąpił błąd: ' + e.message, ui.ButtonSet.OK);
  }
}

function promptAndDownload() {
  const ui = SpreadsheetApp.getUi();
  const resp = ui.prompt('Pobierz pliki DXF', 'Podaj numer zestawu (np. P1608):', ui.ButtonSet.OK_CANCEL);
  if (resp.getSelectedButton() !== ui.Button.OK) return;
  const setId = resp.getResponseText().trim();
  if (!setId) {
    ui.alert('Nie podano numeru zestawu.');
    return;
  }
  try {
    // Deprecated: single-set download removed. Use "Pobierz pliki dla zestawu (z kolorami)" and leave colors unset.
    ui.alert('Opcja usunięta', 'Użyj menu "Pobierz pliki dla zestawu (z kolorami)..." i nie wybieraj kolorów, aby pobrać bez podziału.', ui.ButtonSet.OK);
  } catch (e) {
    ui.alert('Błąd', 'Wystąpił błąd: ' + e.message, ui.ButtonSet.OK);
  }
}

function promptAndCreateElementList() {
  const ui = SpreadsheetApp.getUi();
  const resp = ui.prompt('Pobierz listę elementów', 'Podaj numer zestawu (np. P1608):', ui.ButtonSet.OK_CANCEL);
  if (resp.getSelectedButton() !== ui.Button.OK) return;
  const setId = resp.getResponseText().trim();
  if (!setId) {
    ui.alert('Nie podano numeru zestawu.');
    return;
  }
  try {
    createElementListForSet(setId);
  } catch (e) {
    ui.alert('Błąd', 'Wystąpił błąd: ' + e.message, ui.ButtonSet.OK);
  }
}

function promptAndCreateElementListModule() {
  const ui = SpreadsheetApp.getUi();
  const resp = ui.prompt('Pobierz listę elementów (moduł)', 'Podaj numer modułu (np. M1594):', ui.ButtonSet.OK_CANCEL);
  if (resp.getSelectedButton() !== ui.Button.OK) return;
  const modId = resp.getResponseText().trim();
  if (!modId) {
    ui.alert('Nie podano numeru modułu.');
    return;
  }
  try {
    createElementListForModule(modId);
  } catch (e) {
    ui.alert('Błąd', 'Wystąpił błąd: ' + e.message, ui.ButtonSet.OK);
  }
}

function promptAndDownloadModule() {
  const ui = SpreadsheetApp.getUi();
  const resp = ui.prompt('Pobierz pliki DXF', 'Podaj numer modułu (np. M1594):', ui.ButtonSet.OK_CANCEL);
  if (resp.getSelectedButton() !== ui.Button.OK) return;
  const modId = resp.getResponseText().trim();
  if (!modId) {
    ui.alert('Nie podano numeru modułu.');
    return;
  }
  try {
    downloadModuleFiles(modId);
  } catch (e) {
    ui.alert('Błąd', 'Wystąpił błąd: ' + e.message, ui.ButtonSet.OK);
  }
}
