<?php

header('Content-Type: application/json'); // Set the content type to JSON

function getColumnLetterFromNumber($columnNumber){
	if ($columnNumber <= 0) {
		throw new InvalidArgumentException("Le numéro de colonne doit être un entier positif.");
	}

	$columnLetter = '';
	while ($columnNumber > 0) {
		// Décrémente le nombre de 1 pour s'adapter à la base 0 (A=0, B=1, ...)
		// car le modulo est de 0 à 25.
		$remainder = ($columnNumber - 1) % 26;
		
		// Convertit le reste en lettre ASCII (A=65)
		$columnLetter = chr(ord('A') + $remainder) . $columnLetter;
		
		// Divise le nombre par 26 pour passer à la "position" suivante
		$columnNumber = floor(($columnNumber - 1) / 26);
	}

	return $columnLetter;
}

function translateToFrench($shortName) {
	// Tableaux de correspondance
	$days = [
		'mon' => 'Lun',
		'tue' => 'Mar',
		'wed' => 'Mer',
		'thu' => 'Jeu',
		'fri' => 'Ven',
		'sat' => 'Sam',
		'sun' => 'Dim'
	];
	
	$months = [
		'jan' => 'Jan',
		'feb' => 'Fév',
		'mar' => 'Mar',
		'apr' => 'Avr',
		'may' => 'Mai',
		'jun' => 'Jui',
		'jul' => 'Jui',
		'aug' => 'Aoû',
		'sep' => 'Sep',
		'oct' => 'Oct',
		'nov' => 'Nov',
		'dec' => 'Déc'
	];

	// Convertir en minuscules pour la recherche
	$search = strtolower($shortName);
	
	// Chercher dans les jours
	if (array_key_exists($search, $days)) {
		return $days[$search];
	}
	
	// Chercher dans les mois
	if (array_key_exists($search, $months)) {
		return $months[$search];
	}
	
	// Si aucune correspondance trouvée
	return 'Non trouvé';
}

$date = new DateTime();
$sixMonthsLater = new DateTime();
$sixMonthsLater->modify('+4 months'); // Modified to +4 months as in your original script

$mois = $date->format('M');

$ligne1 = []; // Initialize as an empty array
$ligne2 = [];
$ligne3 = [];
$ligne4 = [];
$letter = [];

// Add initial values for the current day
$ligne1[] = translateToFrench($mois) . " " .$date->format('Y');
$ligne2[] = $date->format('d');
$ligne3[] = translateToFrench($date->format('D'));
$ligne4[] = strtotime($date->format('d-M-Y'));
$letter[] = getColumnLetterFromNumber(6);
$numlet = 7;

if ($_GET["l"] == "s") {
	$debut = explode("/",$_GET["d"]);
	$debut = strtotime($debut[1]."-".$debut[0]."-".$debut["2"]);
	$fin = strtotime($_GET["f"]);
	$r = $_GET["r"];
}

while ($date < $sixMonthsLater) {
	$date->modify('+1 day'); // Move to the next day first

	if ($mois == $date->format('M')) {
		$ligne1[] = ""; // Add empty string for the same month
	} else {
		$ligne1[] = translateToFrench($date->format('M')) . " " .$date->format('Y'); // Add new month
		$mois = $date->format("M");
	}
	$ligne2[] = $date->format("d");
	$ligne3[] = translateToFrench($date->format("D"));
	$ligne4[] = strtotime($date->format("d-M-Y"));
	$letter[] = getColumnLetterFromNumber($numlet);
	
	if($_GET["l"]=="s"){
		//echo strtotime($date->format("d-m-Y")) . " " . $fin . "\n";
		if(strtotime($date->format("d-m-Y")) >= $debut){
			$range=getColumnLetterFromNumber($numlet);
			$debut = 9999999999999999;
		}
		if(strtotime($date->format("d-m-Y")) >= $fin){
			$range=$range .$r. ":".getColumnLetterFromNumber($numlet-1).$r;
			$fin= 9999999999999999;
		}
	}

	
	$numlet++;
}

$result = $_GET["l"];

if ($result == "1") {
	echo json_encode($ligne1);
} elseif ($result == "2") {
	echo json_encode($ligne2);
} elseif ($result == "3") {
	echo json_encode($ligne3);
} elseif ($result == "l") {
	echo json_encode($letter);
} elseif ($result == "s") {
	echo json_encode($range);
} else {
	// Optional: Handle invalid 'l' parameter
	http_response_code(400); // Bad Request
	echo json_encode(["error" => "Invalid 'l' parameter. Use 1, 2, or 3."]);
}

?>