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
	.addToUi();
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

function listCalendarEventsById(calendarId,num,coul) {
  if (!calendarId) {
	Logger.log("Erreur : Aucun ID de calendrier n'a été fourni.");
	return;
  }
  var sheet = SpreadsheetApp.getActiveSheet();
  var today = new Date();
  var fourMonthsLater = new Date();
  fourMonthsLater.setMonth(today.getMonth() + 4); // Ajoute 4 mois à la date actuelle
  var calendar;
  try {
	calendar = CalendarApp.getCalendarById(calendarId);
  } catch (e) {
	Logger.log("Erreur lors de la récupération du calendrier par ID: " + e.message);
	Logger.log("Assurez-vous que l'ID est correct et que le script a les permissions d'accès à ce calendrier.");
	return;
  }

  if (!calendar) {
	Logger.log("Erreur : Le calendrier avec l'ID '" + calendarId + "' n'a pas été trouvé ou le script n'a pas les permissions nécessaires.");
	return;
  }
  var events = calendar.getEvents(today, fourMonthsLater);
  if (events.length > 0) {
	//Colorise la première ligne
	Logger.log("Colorise la première ligne du calendrier")
	events.forEach(function(event) {
	  Logger.log(event.getStartTime())
	  Logger.log("https://winlog.iut-rodez.fr/admin/planing/ljour.php?l=s&d=" + formatDate(event.getStartTime()) + "&f=" + formatDate(event.getEndTime()) + "&r="+num)
	  respcolor  = UrlFetchApp.fetch("https://winlog.iut-rodez.fr/admin/planing/ljour.php?l=s&d=" + formatDate(event.getStartTime()) + "&f=" + formatDate(event.getEndTime()) + "&r="+num)
	  rangecolor = respcolor.getContentText()
	  Logger.log(JSON.parse(rangecolor))
	  sheet.getRange(JSON.parse(rangecolor)).setBackgroundColor(coul);
	  Logger.log(rangecolor)
	  var note = event.getTitle() + "\n" + formatDateTime(event.getStartTime())
	  addNotesToRange(JSON.parse(rangecolor),note)
	  Utilities.sleep(1000);
	});
	//
	
	// Repasse pour faire une ligne par évenement :
	var ligne_debut = sheet.getLastRow() +1
	events.forEach(function(event) {
	  //Ajustement de la date : 
	  sheet.appendRow(["",event.getTitle(),formatDate(event.getStartTime()),event.getLocation(),formatTime(event.getStartTime())])
	  var range = sheet.getRange("A"+sheet.getLastRow()+":E"+sheet.getLastRow());
	range.setBackgroundColor(coul);
	  respcolor  = UrlFetchApp.fetch("https://winlog.iut-rodez.fr/admin/planing/ljour.php?l=s&d=" + formatDate(event.getStartTime()) + "&f=" + formatDate(event.getEndTime()) + "&r="+num)
	  rangecolor = respcolor.getContentText()
	  sheet.getRange(JSON.parse(rangecolor)).setBackgroundColor(coul);
	  Utilities.sleep(1000);
	});
	var ligne_fin = sheet.getLastRow()
	groupeLigne(ligne_debut,ligne_fin)
	var range = sheet.getRange("A"+ ligne_debut+":E"+ligne_fin); 
	range.setHorizontalAlignment("left")
  
  } else {
	Logger.log("Aucun événement trouvé pour les 4 prochains mois dans le calendrier '" + calendar.getName() + "' (ID: " + calendar.getId() + ").");
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

function addNotesToRange(rangeA1Notation, note) {
  Logger.log(rangeA1Notation)
  var sheet = SpreadsheetApp.getActiveSheet();
  
  // Récupère la plage de cellules à partir de la notation A1
  const range = sheet.getRange(rangeA1Notation);
  
  // Récupère toutes les cellules de la plage
  const numRows = range.getNumRows();
  const numCols = range.getNumColumns();

  // Parcourt chaque cellule de la plage
  for (let row = 1; row <= numRows; row++) {
	let nbevent = 0;
	let notesSet = new Set(); // Utiliser un Set pour garder trace des notes uniques

	for (let col = 1; col <= numCols; col++) {
	  const cell = range.getCell(row, col);
	  const existingNote = cell.getNote();
	  
	  if (existingNote) {
		// Ajouter la note au Set
		notesSet.add(existingNote);
		// Mettre à jour le compteur avec le nombre total de notes uniques
		Logger.log(cell.getValue() + 1)
		cell.setValue(String(cell.getValue() + 1));
	  }
	  else{
		cell.setValue(String(1));
	  }
	  
	  const combinedNote = existingNote 
		? `${existingNote}\n${note}\n`
		: note;
	  
	  cell.setNote(combinedNote);
	  Logger.log(`Nombre de notes: ${nbevent}`);
	}
  }
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

function getALLcal(){
  const calendars = CalendarApp.getAllOwnedCalendars();
  for (var i = 0; i < calendars.length; i++) {
	var calendar = calendars[i];
	var calID = calendar.getId()
	getcaldata(calID,i)
  }
  Logger.log('This user owns %s calendars.', calendars.length); 
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

function getLigneDate(numligne){

  var response = UrlFetchApp.fetch("https://winlog.iut-rodez.fr/admin/planing/ljour.php?l="+numligne);

  // Get the content of the response as a string
  var jsonString = response.getContentText();

  var ligne2_data;

  // Parse the JSON string into a JavaScript array
  // The PHP script directly outputs the array, so no ".result" is needed
  ligne2_data = JSON.parse(jsonString);
  var emptyColumns = ["", "", "", "", ""];
  if(numligne == "2"){
	emptyColumns = ["Calendrier", "Description", "Jour", "Lieu", "Heure"]
  }
  ligne2_data = emptyColumns.concat(ligne2_data); 

  // Log the parsed data to see what you received
  Logger.log("Parsed data: " + JSON.stringify(ligne2_data));

  // Check if the array is empty before appending
  if (ligne2_data && ligne2_data.length > 0) {
	return(ligne2_data);
  } else {
	Logger.log("Data fetched from URL is empty. Cannot append row.");
	// You might want to throw an error or handle this case differently
  }
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

function reset(){
  var sh = SpreadsheetApp.getActiveSheet();
  var values = sh.getDataRange().getValues();
  //Logger.log(values.length);
  for(var i=0, iLen=values.length; i<iLen; i++) {
	sh.deleteRow(1);
  }
}