function lance_script(startDate = new Date(), endDate = null){
  const ss = SpreadsheetApp.getActive();
  ss.toast('Initialisation...', 'Status', -1);
  
  if (!endDate) {
    endDate = new Date(startDate);
    endDate.setMonth(startDate.getMonth() + 4);
  }
  
  // Stocker les dates dans les propriétés du document
  PropertiesService.getDocumentProperties().setProperties({
    'startDate': startDate.toISOString(),
    'endDate': endDate.toISOString()
  });
  
  ss.toast('Préparation de la feuille...', 'Status', -1);
  ajout_lignes_debut();
  
  ss.toast('Chargement des calendriers...', 'Status', -1);
  getALLcal();
  
  ss.toast('Mise à jour terminée !', 'Status', 3);
}

function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('IUT Planning')
    .addItem('Rafraichir (4 mois)', 'lance_script')
    .addItem('Afficher année scolaire...', 'showYearPickerDialog')
    .addItem('Sélectionner une plage de dates...', 'showDateRangePickerDialog')
    .addSeparator()
    .addSubMenu(ui.createMenu('Affichage')
      .addItem('Étendre tous les groupes', 'expandAllGroups')
      .addItem('Regrouper tous les groupes', 'collapseAllGroups')
      .addSeparator()
      .addItem('Aller à aujourd\'hui', 'detectAndScrollToToday'))
    .addToUi();
}

function detectAndScrollToToday() {
  const html = HtmlService.createHtmlOutput(`
    <script>
      const width = window.innerWidth;
      google.script.run
        .withSuccessHandler(() => google.script.host.close())
        .processWindowWidthAndScroll(width);
    </script>
  `)
  .setWidth(1)
  .setHeight(1);
  
  SpreadsheetApp.getUi().showModalDialog(html, 'Chargement...');
}

function processWindowWidthAndScroll(width) {
  const sheet = SpreadsheetApp.getActiveSheet();
  const props = PropertiesService.getDocumentProperties();
  const startDate = new Date(props.getProperty('startDate'));
  const today = new Date();
  
  // S'assurer que les deux dates sont au début de la journée pour une comparaison correcte
  startDate.setHours(0, 0, 0, 0);
  today.setHours(0, 0, 0, 0);
  
  // Calculer le nombre de jours depuis la date de début
  const diffDays = Math.round((today - startDate) / (1000 * 60 * 60 * 24));
  
  // Obtenir la largeur totale des colonnes fixes (A à E)
  const fixedWidth = [
    sheet.getColumnWidth(1),
    sheet.getColumnWidth(2),
    sheet.getColumnWidth(3),
    sheet.getColumnWidth(4),
    sheet.getColumnWidth(5)
  ].reduce((sum, w) => sum + w, 0);
  
  // Largeur standard d'une colonne de date (28px)
  const dateColumnWidth = 28;
  
  // Calculer le nombre de colonnes visibles
  const visibleWidth = width - fixedWidth;
  const visibleDateColumns = Math.floor(visibleWidth / dateColumnWidth);
  
  // Calculer la colonne cible
  const offset = Math.floor(visibleDateColumns);
  const todayColumn = Math.max(6, 6 + diffDays - offset);
  
  // Faire défiler jusqu'à la colonne calculée
  sheet.setActiveRange(sheet.getRange(1, todayColumn));
}

function showYearPickerDialog() {
  const ui = SpreadsheetApp.getUi();
  const result = ui.prompt(
    'Année scolaire',
    'Entrez l\'année de début (ex: 2024 pour 2024-2025)',
    ui.ButtonSet.OK_CANCEL
  );

  if (result.getSelectedButton() == ui.Button.OK) {
    const year = parseInt(result.getResponseText());
    if (!isNaN(year)) {
      const startDate = new Date(year, 8, 1); // 1er septembre
      const endDate = new Date(year + 1, 7, 31); // 31 août
      lance_script(startDate, endDate);
    }
  }
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
  const props = PropertiesService.getDocumentProperties();
  reset();
  var sheet = SpreadsheetApp.getActiveSheet();
  
  // Déterminer le type de planning et le titre
  const startDateStr = props.getProperty('startDate');
  const endDateStr = props.getProperty('endDate');
  const startDate = new Date(startDateStr);
  const endDate = new Date(endDateStr);
  
  // Formater les dates pour l'affichage
  const formatDateFr = (date) => date.toLocaleDateString('fr-FR', { 
    day: '2-digit', 
    month: '2-digit',
    year: 'numeric'
  });

  // Déterminer si on est en mode année scolaire ou plage personnalisée
  const isYearMode = endDate.getMonth() === 7; // Vérifie si la date de fin est en août
  const titre = isYearMode 
    ? "Planning annuel"
    : `Planning du ${formatDateFr(startDate)} au ${formatDateFr(endDate)}`;
  
  sheet.appendRow(["Dernière actualisation : " + new Date().toLocaleString('fr-FR')]);
  sheet.appendRow([titre]);
  sheet.appendRow(getLigneDate("1"));
  sheet.appendRow(getLigneDate("2"));
  sheet.appendRow(getLigneDate("3"));

  // Griser les weekends
  const dateLines = generateDateLines();
  dateLines.weekendColumns.forEach(index => {
    const column = 6 + index; // 6 est la colonne de base (F)
    const range = sheet.getRange(3, column, sheet.getMaxRows(), 1);
    range.setBackgroundColor('#f3f3f3');
  });

  // Ajouter les bordures verticales pour les débuts de mois (par dessus le gris)
  dateLines.monthStartColumns.forEach(index => {
    const column = 6 + index;
    const range = sheet.getRange(1, column, sheet.getMaxRows(), 1);
    range.setBorder(null, true, null, null, null, null, '#000000', SpreadsheetApp.BorderStyle.SOLID);
  });

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
    'Mon': 'Lu', 'Tue': 'Ma', 'Wed': 'Me', 'Thu': 'Je', 
    'Fri': 'Ve', 'Sat': 'Sa', 'Sun': 'Di',
    'Jan': 'Janvier', 'Feb': 'Février', 'Mar': 'Mars', 'Apr': 'Avril',
    'May': 'Mai', 'Jun': 'Juin', 'Jul': 'Juillet', 'Aug': 'Août',
    'Sep': 'Septembre', 'Oct': 'Octobre', 'Nov': 'Novembre', 'Dec': 'Décembre'
  };
  return translations[shortName] || shortName;
}

function generateDateLines() {
  const props = PropertiesService.getDocumentProperties();
  const startDate = props.getProperty('startDate') 
    ? new Date(props.getProperty('startDate')) 
    : new Date(new Date().setHours(0,0,0,0));
  const endDate = props.getProperty('endDate') 
    ? new Date(props.getProperty('endDate')) 
    : new Date(startDate.getTime() + (4 * 30 * 24 * 60 * 60 * 1000));

  const ligne1 = [], ligne2 = [], ligne3 = [];
  const monthStartColumns = []; // Indices des débuts de mois
  const weekendColumns = []; // Ajouter un tableau pour les colonnes de weekend
  let currentDate = new Date(startDate);
  let currentMonth = currentDate.getMonth();
  let dayIndex = 0;

  while (currentDate <= endDate) {
    // Détecter le weekend (6 = samedi, 0 = dimanche)
    if (currentDate.getDay() === 6 || currentDate.getDay() === 0) {
      weekendColumns.push(dayIndex);
    }

    // Détecter le premier jour du mois
    if (currentDate.getDate() === 1 || dayIndex === 0) {
      monthStartColumns.push(dayIndex);
    }

    const month = currentDate.toLocaleString('en-US', { month: 'short' });
    const day = currentDate.getDate();
    const weekday = currentDate.toLocaleString('en-US', { weekday: 'short' });

    ligne1.push(day === 1 || currentMonth !== currentDate.getMonth() 
      ? translateToFrench(month) + ' ' + currentDate.getFullYear()
      : '');
    ligne2.push(day.toString());
    ligne3.push(translateToFrench(weekday));

    currentMonth = currentDate.getMonth();
    currentDate.setDate(currentDate.getDate() + 1);
    dayIndex++;
  }

  return {
    ligne1, ligne2, ligne3,
    startDate, endDate,
    monthStartColumns,
    weekendColumns // Ajouter les colonnes de weekend au retour
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

function calculateCellRange(eventStartDate, eventEndDate, rowNum) {
  const props = PropertiesService.getDocumentProperties();
  const referenceDate = props.getProperty('startDate') 
    ? new Date(props.getProperty('startDate')) 
    : new Date(new Date().setHours(0,0,0,0));

  const baseCol = 6;
  
  // Créer des dates en début de journée pour la comparaison
  const eventStart = new Date(eventStartDate);
  eventStart.setHours(0,0,0,0);
  
  // Si l'événement se termine à minuit, reculer d'un jour
  const eventEnd = new Date(eventEndDate);
  if (eventEnd.getHours() === 0 && eventEnd.getMinutes() === 0) {
    eventEnd.setDate(eventEnd.getDate() - 1);
  }
  eventEnd.setHours(0,0,0,0);
  
  const refDate = new Date(referenceDate);
  refDate.setHours(0,0,0,0);
  
  // Calculer la différence en jours sans tenir compte des heures
  const startDays = Math.floor((eventStart - refDate) / (1000 * 60 * 60 * 24));
  const endDays = Math.floor((eventEnd - refDate) / (1000 * 60 * 60 * 24)); // Correction: refRefDate -> refDate
  
  // S'assurer que l'événement est dans la période affichée
  if (startDays < -1) return null;
  
  // Ajuster les colonnes
  const startCol = getColumnLetter(baseCol + Math.max(0, startDays));
  const endCol = getColumnLetter(baseCol + Math.max(0, endDays));
  
  return `${startCol}${rowNum}:${endCol}${rowNum}`;
}

function getALLcal() {
  const sheet = SpreadsheetApp.getActiveSheet();
  const ss = SpreadsheetApp.getActive();
  const props = PropertiesService.getDocumentProperties();
  const startDate = new Date(props.getProperty('startDate') || new Date());
  const endDate = new Date(props.getProperty('endDate') || new Date(startDate.getTime() + (4 * 30 * 24 * 60 * 60 * 1000)));
  
  ss.toast('Récupération des calendriers...', 'Status', -1);
  // Filtrer les agendas cochés et exclure ceux dont le nom contient "semaine" (insensible à la casse)
  const calendars = CalendarApp.getAllCalendars()
    .filter(cal => cal.isSelected() && !/semaine/i.test(cal.getName()));

  ss.toast('Chargement des événements...', 'Status', -1);
  const allData = calendars.map(calendar => {
    const events = calendar.getEvents(startDate, endDate)
      .map(event => ({
        title: event.getTitle(),
        startTime: event.getStartTime(),
        endTime: event.getEndTime(),
        location: event.getLocation() || ''
      }));

    return {
      name: calendar.getName(),
      description: calendar.getDescription() || '',
      color: calendar.getColor(),
      events: events
    };
  });

  ss.toast('Préparation des données...', 'Status', -1);
  let currentRow = sheet.getLastRow() + 1;
  const updates = [];
  const colorUpdates = [];
  const dateRanges = [];
  // Modifier la structure de eventSynthesis pour stocker les événements
  const eventSynthesis = {};

  // Préparer toutes les mises à jour pour tous les calendriers
  allData.forEach(calData => {
    // Ajouter l'en-tête du calendrier dans tous les cas
    updates.push({
      range: sheet.getRange(currentRow, 1, 1, 5),
      values: [[calData.name, calData.description, '', '', '']]
    });
    colorUpdates.push({
      range: sheet.getRange(currentRow, 1, 1, 5),
      color: calData.color
    });

    if (calData.events.length > 0) {
      // Traitement des événements si le calendrier en a
      const eventRows = calData.events.map(event => [
        '',
        event.title,
        formatDate(event.startTime),
        event.location,
        formatTime(event.startTime)
      ]);

      if (eventRows.length > 0) {
        updates.push({
          range: sheet.getRange(currentRow + 1, 1, eventRows.length, 5),
          values: eventRows
        });
        colorUpdates.push({
          range: sheet.getRange(currentRow + 1, 1, eventRows.length, 5),
          color: calData.color
        });

        // Préparer les plages de dates et la synthèse
        calData.events.forEach((event, index) => {
          const rowNum = currentRow + 1 + index;
          const dateRange = calculateCellRange(event.startTime, event.endTime, rowNum);
          if (dateRange) {
            const [startCol, endCol] = dateRange.split(':').map(ref => ref.match(/[A-Z]+/)[0]);
            const startIdx = columnLetterToNumber(startCol);
            const endIdx = columnLetterToNumber(endCol);
            
            // Mise à jour du compteur de synthèse et stockage des événements
            for (let col = startIdx; col <= endIdx; col++) {
              const cellRef = `${getColumnLetter(col)}${currentRow}`;
              if (!eventSynthesis[cellRef]) {
                eventSynthesis[cellRef] = {
                  count: 0,
                  color: calData.color,
                  events: []
                };
              }
              eventSynthesis[cellRef].count += 1;
              eventSynthesis[cellRef].events.push({
                title: event.title,
                time: formatTime(event.startTime)
              });
            }

            dateRanges.push({
              range: dateRange,
              color: calData.color,
              note: event.title + "\n" + formatDateTime(event.startTime)
            });
          }
        });
        currentRow += eventRows.length + 1;
      } else {
        currentRow++;
      }
    } else {
      // Incrémenter currentRow même sans événements
      currentRow++;
    }
  });

  // Appliquer les mises à jour standard
  updates.forEach(update => update.range.setValues(update.values));
  colorUpdates.forEach(update => update.range.setBackground(update.color));

  // Appliquer la synthèse avec les couleurs et les notes
  Object.entries(eventSynthesis).forEach(([cellRef, data]) => {
    const cell = sheet.getRange(cellRef);
    if (data.count > 1) {
      cell.setValue(String(data.count));
    }
    cell.setBackground(data.color);
    
    // Créer la note de synthèse
    const noteText = data.events
      .sort((a, b) => a.time.localeCompare(b.time))
      .map(event => `${event.time} - ${event.title}`)
      .join('\n');
    cell.setNote(noteText);
  });

  // Appliquer les plages de dates
  dateRanges.forEach(update => {
    const range = sheet.getRange(update.range);
    range.setBackground(update.color);
    range.setNote(update.note);
  });
  correctGroup()
  sheet.collapseAllRowGroups();
}

// Fonction utilitaire pour convertir une lettre de colonne en nombre
function columnLetterToNumber(letter) {
  let column = 0;
  const length = letter.length;
  for (let i = 0; i < length; i++) {
    column += (letter.charCodeAt(i) - 64) * Math.pow(26, length - i - 1);
  }
  return column;
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
  
  const day = String(date.getDate()).padStart(2, '0');
  const month = String(date.getMonth() + 1).padStart(2, '0');
  const year = date.getFullYear();

  return `${day}/${month}/${year}`;
}

function formatTime(dateString) {
  const date = new Date(dateString);
  
  const hours = String(date.getHours()).padStart(2, '0');
  const minutes = String(date.getMinutes()).padStart(2, '0');

  return `${hours}:${minutes}`;
}

function setColumnWidths() {
  const sheet = SpreadsheetApp.getActiveSheet();
  const lastColumn = sheet.getLastColumn();
  
  // Commencer à partir de la colonne F (6ème colonne)
  const startColumn = 6;
  
  // Parcourir toutes les colonnes à partir de F
  for (let col = startColumn; col <= lastColumn; col++) {
  sheet.setColumnWidth(col, 20);
  }
  sheet.setColumnWidth(2,300)

  sheet.setColumnWidth(3,78)
  sheet.setColumnWidth(5,50)
}

function reset() {
  const ss = SpreadsheetApp.getActive();
  const oldSheet = ss.getActiveSheet();
  const props = PropertiesService.getDocumentProperties();
  const startDate = props.getProperty('startDate') ? new Date(props.getProperty('startDate')) : new Date();
  const endDate = props.getProperty('endDate') ? new Date(props.getProperty('endDate')) : new Date(startDate.getTime() + (4 * 30 * 24 * 60 * 60 * 1000));
  const formatDateFr = date => date.toLocaleDateString('fr-FR', { day: '2-digit', month: '2-digit', year: 'numeric' });

  let periodLabel;
  if (endDate.getMonth() === 7) {
    periodLabel = `${startDate.getFullYear()}-${endDate.getFullYear()}`;
  } else {
    periodLabel = `${formatDateFr(startDate)} au ${formatDateFr(endDate)}`;
  }
  let newSheetName = "Planning " + periodLabel;

  // Si la feuille existe déjà, ajouter un suffixe numérique
  let suffix = 1;
  while (ss.getSheetByName(newSheetName)) {
    newSheetName = "Planning " + periodLabel + " (" + suffix + ")";
    suffix++;
  }

  const newSheet = ss.insertSheet(newSheetName);
  newSheet.activate();
  ss.deleteSheet(oldSheet);
  newSheet.setName(newSheetName);
}

function showDateRangePickerDialog() {
  const html = HtmlService.createHtmlOutput(`
    <style>
      .date-container {
        display: flex;
        gap: 15px;
        margin-bottom: 20px;
        min-width: 0; /* Empêche le débordement des flex items */
      }
      .date-field {
        flex: 1;
        min-width: 0; /* Permet aux champs de rétrécir si nécessaire */
        overflow: hidden; /* Contient les débordements */
      }
      label {
        display: block;
        margin-bottom: 5px;
        white-space: nowrap; /* Empêche le retour à la ligne du texte */
      }
      input[type="date"] {
        width: 100%;
        padding: 5px;
        box-sizing: border-box; /* Inclut padding dans la largeur totale */
      }
      input[type="button"] {
        width: 100%;
        padding: 8px;
        cursor: pointer;
      }
    </style>
    <form>
      <div class="date-container">
        <div class="date-field">
          <label>Date de début:</label>
          <input type="date" id="startDate" required>
        </div>
        <div class="date-field">
          <label>Date de fin:</label>
          <input type="date" id="endDate" required>
        </div>
      </div>
      <input type="button" value="Valider" onclick="submitDates()">
    </form>
    <script>
      function submitDates() {
        var startDate = document.getElementById('startDate').value;
        var endDate = document.getElementById('endDate').value;
        if (startDate && endDate) {
          google.script.run
            .withSuccessHandler(function() {
              google.script.host.close();
            })
            .withFailureHandler(function(error) {
              console.error('Erreur:', error);
              google.script.host.close();
            })
            .processCustomDateRange(startDate, endDate);
        }
      }
    </script>
  `)
  .setWidth(400)
  .setHeight(150)
  .setTitle('Sélectionner une plage de dates');
  SpreadsheetApp.getUi().showModalDialog(html, 'Sélectionner une plage de dates');
}

function correctGroup() {
  const sheet = SpreadsheetApp.getActiveSheet();
  const firstDataRow = 6;
  const lastRow = sheet.getLastRow();
  if (lastRow < firstDataRow) return;
  const colors = sheet.getRange(firstDataRow, 1, lastRow - firstDataRow + 1, 1).getBackgrounds();
  let prevColor = colors[0][0];
  let start = firstDataRow;
  for (let i = 1; i < colors.length; i++) {
    const color = colors[i][0];
    if (color !== prevColor) {
      // Groupe uniquement si plus d'une ligne
      Logger.log(i + "  " + start + " -> " + (firstDataRow + i - 1))
      if ((firstDataRow + i - 1) > start) {
        sheet.getRange(start+1, 1, (firstDataRow + i - 1) - start, 1).shiftRowGroupDepth(1);
      }
      start = firstDataRow + i;
      prevColor = color;
    }
  }
  // Dernier groupe
  if ((firstDataRow + colors.length - 1) > start) {
    sheet.getRange(start+1, 1, (firstDataRow + colors.length - 1) - start, 1).shiftRowGroupDepth(1);
  }
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

function processCustomDateRange(startDateStr, endDateStr) {
  const startDate = new Date(startDateStr);
  const endDate = new Date(endDateStr);
  lance_script(startDate, endDate);
}