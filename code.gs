function lance_script(){
  onOpen()
  SpreadsheetApp.getActive().toast('Mise à jour en cours...', 'Status', -1);
  ajout_lignes_debut()
  var sheet = SpreadsheetApp.getActiveSheet();
  getALLcal()
  SpreadsheetApp.getActive().toast('Mise à jour terminée !', 'Status', 3);
}

function onOpen() {
  SpreadsheetApp.getUi()
  .createMenu('IUT Planning')
  .addItem('Rafraichir', 'lance_script')
    .addSeparator()
    .addSubMenu(SpreadsheetApp.getUi().createMenu('Affichage')
      .addItem('Étendre tous les groupes', 'expandAllGroups')
      .addItem('Regrouper tous les groupes', 'collapseAllGroups'))
    .addToUi();
}

function expandAllGroups() {
  var sheet = SpreadsheetApp.getActiveSheet();
  sheet.expandAllRowGroups()
}

function collapseAllGroups() {
  var sheet = SpreadsheetApp.getActiveSheet();
  sheet.collapseAllRowGroups()
}

function ajout_lignes_debut(){
  var aujourdhui = new Date();
  var datedujour = aujourdhui.toISOString();
  datedujour = gooddate(datedujour)
  reset();
  var sheet = SpreadsheetApp.getActiveSheet();
  sheet.appendRow(["Dernière actualisation : " + datedujour])
  sheet.appendRow(["--------------------------------------"])
  // recuperation des lignes de dates
  sheet.appendRow(getLigneDate("1"))
  sheet.appendRow(getLigneDate("2"))
  sheet.appendRow(getLigneDate("3"))

  //Alignement dans les cellules
  var range = sheet.getDataRange(); 
  // Set horizontal alignment to LEFT for all cells in that range
  range.setHorizontalAlignment("left"); 

  // Bordure titres des evenements et alignement au centre.
  var range = sheet.getRange("A3:E5"); 
  range.setBorder(true, true, true, true, false, false, '#000000', SpreadsheetApp.BorderStyle.SOLID);
  range.setHorizontalAlignment("center")
  range.setBackgroundColor('#c3c9d0');
  setColumnWidths()
  sheet.setFrozenRows(5);
  sheet.setFrozenColumns(5);
}

function sheetId(){
  const sheet = SpreadsheetApp.getActiveSheet();
}

function getColumnLetter(columnNumber) {
  let columnLetter = '';
  while (columnNumber > 0) {
    const remainder = (columnNumber - 1) % 26;
    columnLetter = String.fromCharCode(65 + remainder) + columnLetter;
    columnNumber = Math.floor((columnNumber - 1) / 26);
  }
  return columnLetter;
}

function translateToFrench(shortName) {
  const translations = {
    'Mon': 'Lun', 'Tue': 'Mar', 'Wed': 'Mer', 'Thu': 'Jeu', 
    'Fri': 'Ven', 'Sat': 'Sam', 'Sun': 'Dim',
    'Jan': 'Jan', 'Feb': 'Fév', 'Mar': 'Mar', 'Apr': 'Avr',
    'May': 'Mai', 'Jun': 'Jui', 'Jul': 'Jui', 'Aug': 'Aoû',
    'Sep': 'Sep', 'Oct': 'Oct', 'Nov': 'Nov', 'Dec': 'Déc'
  };
  return translations[shortName] || shortName;
}

function generateDateLines() {
  const today = new Date();
  const fourMonthsLater = new Date(today);
  fourMonthsLater.setMonth(today.getMonth() + 4);

  const ligne1 = [], ligne2 = [], ligne3 = [];
  let currentDate = new Date(today);
  let currentMonth = currentDate.toLocaleString('en-US', { month: 'short' });

  // Start from column F (6)
  let colIndex = 6;
  
  while (currentDate <= fourMonthsLater) {
    const month = currentDate.toLocaleString('en-US', { month: 'short' });
    const day = currentDate.getDate();
    const weekday = currentDate.toLocaleString('en-US', { weekday: 'short' });

    // Ligne 1 (mois)
    ligne1.push(currentMonth === month ? '' : 
      translateToFrench(month) + ' ' + currentDate.getFullYear());
    
    // Ligne 2 (jour)
    ligne2.push(day.toString());
    
    // Ligne 3 (jour de la semaine)
    ligne3.push(translateToFrench(weekday));

    currentMonth = month;
    currentDate.setDate(currentDate.getDate() + 1);
    colIndex++;
  }

  return {
    ligne1: ligne1,
    ligne2: ligne2,
    ligne3: ligne3
  };
}

function getLigneDate(numligne) {
  const dateLines = generateDateLines();
  const emptyColumns = ["", "", "", "", ""];
  
  if (numligne === "1") return emptyColumns.concat(dateLines.ligne1);
  if (numligne === "2") {
    return ["Calendrier", "Description", "Jour", "Lieu", "Heure"].concat(dateLines.ligne2);
  }
  if (numligne === "3") return emptyColumns.concat(dateLines.ligne3);
  
  return [];
}

function calculateCellRange(startDate, endDate, rowNum) {
  const today = new Date();
  const baseCol = 6; // Column F
  
  const startDays = Math.floor((startDate - today) / (1000 * 60 * 60 * 24));
  const endDays = Math.floor((endDate - today) / (1000 * 60 * 60 * 24));
  
  const startCol = getColumnLetter(baseCol + startDays);
  const endCol = getColumnLetter(baseCol + endDays);
  
  return `${startCol}${rowNum}:${endCol}${rowNum}`;
}

function listCalendarEventsById(calendarId, num, coul) {
  if (!calendarId) {
    Logger.log("Erreur : Aucun ID de calendrier n'a été fourni.");
    return;
  }
  var sheet = SpreadsheetApp.getActiveSheet();
  var today = new Date();
  var fourMonthsLater = new Date();
  fourMonthsLater.setMonth(today.getMonth() + 4);
  
  var calendar = CalendarApp.getCalendarById(calendarId);
  if (!calendar) {
    Logger.log("Erreur : Calendrier non trouvé");
    return;
  }

  var events = calendar.getEvents(today, fourMonthsLater);
  if (events.length > 0) {
    // Préparer les données en mémoire
    const rangesToColor = [];
    const notesToAdd = [];
    const rowsToAdd = [];
    
    events.forEach(function(event) {
      const range = calculateCellRange(event.getStartTime(), event.getEndTime(), num);
      rangesToColor.push({
        range: range,
        note: event.getTitle() + "\n" + formatDateTime(event.getStartTime())
      });
      
      rowsToAdd.push({
        data: ["", event.getTitle(), formatDate(event.getStartTime()), 
               event.getLocation(), formatTime(event.getStartTime())],
        startDate: event.getStartTime(),
        endDate: event.getEndTime()
      });
    });

    // Appliquer les couleurs et notes en batch
    rangesToColor.forEach(({range, note}) => {
      sheet.getRange(range).setBackgroundColor(coul);
      addNotesToRange(range, note);
    });

    // Ajouter toutes les lignes en une seule fois
    const ligne_debut = sheet.getLastRow() + 1;
    const rowsData = rowsToAdd.map(row => row.data);
    sheet.getRange(ligne_debut, 1, rowsData.length, 5).setValues(rowsData);
    
    // Colorer toutes les nouvelles lignes en une fois
    sheet.getRange(ligne_debut, 1, rowsData.length, 5).setBackgroundColor(coul);

    // Colorer les cellules de dates pour chaque ligne de détail
    rowsToAdd.forEach((row, index) => {
      const rowNum = ligne_debut + index;
      const dateRange = calculateCellRange(row.startDate, row.endDate, rowNum);
      sheet.getRange(dateRange).setBackgroundColor(coul);
    });

    // Grouper les lignes
    const ligne_fin = ligne_debut + rowsData.length - 1;
    groupeLigne(ligne_debut, ligne_fin);
    sheet.getRange(`A${ligne_debut}:E${ligne_fin}`).setHorizontalAlignment("left");
  }
}

function addNotesToRange(rangeA1Notation, note) {
  const sheet = SpreadsheetApp.getActiveSheet();
  const range = sheet.getRange(rangeA1Notation);
  const values = range.getValues();
  const notes = range.getNotes();
  const newValues = [];
  const newNotes = [];

  for (let i = 0; i < values.length; i++) {
    newValues[i] = [];
    newNotes[i] = [];
    for (let j = 0; j < values[i].length; j++) {
      const existingNote = notes[i][j];
      newValues[i][j] = existingNote ? String(Number(values[i][j] || 0) + 1) : "1";
      newNotes[i][j] = existingNote ? `${existingNote}\n${note}\n` : note;
    }
  }

  range.setValues(newValues);
  range.setNotes(newNotes);
}

function getcaldata(calID, num) {
  const calendar = CalendarApp.getCalendarById(calID);
  
  // Vérifie si le calendrier est sélectionné
  if (!calendar.isSelected()) {
  Logger.log(`Calendrier ${calID} non sélectionné - ignoré`);
  return; // Sort de la fonction si le calendrier n'est pas sélectionné
  }

  const calendarName = calendar.getName();
  const calendarDesc = calendar.getDescription();
  const calendarColor = calendar.getColor();
  
  Logger.log(num + " " + calID + ' : ' + calendarName + " : " + calendarDesc + ' -> ' + calendarColor);
  
  var sheet = SpreadsheetApp.getActiveSheet();
  num = sheet.getLastRow() + 1;
  range = "A" + num + ":E" + num;
  Logger.log(range);
  
  var range = sheet.getRange("A" + num + ":E" + num);
  range.setBackgroundColor(calendarColor);
  sheet.appendRow([calendarName, calendarDesc]);
  
  Logger.log(listCalendarEventsById(calID, num, calendarColor));
}

function getALLcal() {
  const calendars = CalendarApp.getAllOwnedCalendars().filter(cal => cal.isSelected());
  const sheet = SpreadsheetApp.getActiveSheet();
  
  calendars.forEach((calendar, index) => {
    const calendarName = calendar.getName();
    const calendarDesc = calendar.getDescription();
    const calendarColor = calendar.getColor();
    const rowNum = sheet.getLastRow() + 1;
    
    sheet.getRange(`A${rowNum}:E${rowNum}`)
         .setBackgroundColor(calendarColor);
    sheet.appendRow([calendarName, calendarDesc]);
    
    listCalendarEventsById(calendar.getId(), rowNum, calendarColor);
  });
}

function formatDateTime(date) {
  const dateObj = new Date(date);
  const isStartOfDay = dateObj.getHours() === 0 && dateObj.getMinutes() === 0;

  const options = {
    timeZone: 'Europe/Paris',
    year: 'numeric',
    month: 'long',
    day: 'numeric',
    ...(isStartOfDay ? {} : {
      hour: '2-digit',
      minute: '2-digit'
    })
  };

  return dateObj.toLocaleString('fr-FR', options);
}

function formatDate(dateString) {
  const date = new Date(dateString);
  
  const month = String(date.getMonth() + 1).padStart(2, '0');
  const day = String(date.getDate()).padStart(2, '0');
  const year = date.getFullYear();

  return `${month}/${day}/${year}`;
}

function formatTime(dateString) {
  const date = new Date(dateString);
  
  const hours = String(date.getHours()).padStart(2, '0');
  const minutes = String(date.getMinutes()).padStart(2, '0');

  return `${hours}:${minutes}`;
}
function groupeLigne(debut,fin){
  var sheet = SpreadsheetApp.getActiveSheet();
  
  // Specify the rows to group (for example, rows 2 to 5)
  var startRow = debut;
  var endRow = fin;
  
  // Create the group
  sheet.getRange(startRow, 1, endRow - startRow + 1, 1).shiftRowGroupDepth(1);
  sheet.collapseAllRowGroups()
}


function gooddate(datepasgood){
  //Date vs Heure
  var tabDXH = datepasgood.split("T");
  var date = tabDXH[0];
  var heure = tabDXH[1];
  
  var tabDJMA = date.split("-");
  var jour = tabDJMA[2];
  var mois = tabDJMA[1];
  var annee = tabDJMA[0];

  var tabHMS = heure.split(":");
  var heure = tabHMS[0];
  var minute = tabHMS[1];
  var seconde = tabHMS[2];
  var rab = tabHMS[3];

  var dategood = jour + "-" + mois + "-" + annee + " " + heure + ":" + minute;
  Logger.log(dategood)
  return(dategood);
}

function setColumnWidths() {
  const sheet = SpreadsheetApp.getActiveSheet();
  const lastColumn = sheet.getLastColumn();
  
  // Commencer à partir de la colonne F (6ème colonne)
  const startColumn = 6;
  
  // Parcourir toutes les colonnes à partir de F
  for (let col = startColumn; col <= lastColumn; col++) {
  sheet.setColumnWidth(col, 28);
  }
  sheet.setColumnWidth(2,300)

  sheet.setColumnWidth(3,78)
  sheet.setColumnWidth(5,50)
}

function reset() {
  const sheet = SpreadsheetApp.getActiveSheet();
  const lastRow = sheet.getLastRow();
  
  // Si il y a des lignes à supprimer
  if (lastRow > 0) {
    // Supprime toutes les lignes en une seule opération
    sheet.deleteRows(1, lastRow);
  }
}