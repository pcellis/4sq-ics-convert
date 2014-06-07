<?php

$scriptStart = microtime(true);
$scriptEnd = "";

date_default_timezone_set("America/Los_Angeles");

require_once(dirname(__FILE__) . "/classes/PHPExcel.php");

$excel = new PHPExcel();

// Initialize worksheet skeleton
$worksheet = new PHPExcel_Worksheet($excel, "Foursquare Checkin History");

// Remove initial worksheet created
$excel->removeSheetByIndex(0);

// Read ICS data
$files = glob("*.ics");
$checkins = array();

foreach($files as $file) {
	$eventCount = 0;
	echo "Reading file: " . $file . "... ";

	$content = file_get_contents($file);
	$lines = explode("\n", $content);
	$lines = array_map("trim", $lines);

	$checkin = array();

	$eventStart = false;

	foreach($lines as $line) {
		if($line == "BEGIN:VEVENT") {
			$eventStart = true;
			$checkin = array();
			continue;
		}
		else if($line == "END:VEVENT") {
			$eventCount++;
			$eventStart = false;
			ksort($checkin);

			$index = $checkin["__INDEX__"];
			unset($checkin["__INDEX__"]);

			$checkins[$index] = $checkin;
			unset($checkin);

			continue;
		}

		// Check to ensure we're reading an event. If we are, great.
		if($eventStart) {
			$firstColon = strpos($line, ":");
			$key = substr($line, 0, $firstColon);
			$value = substr($line, strlen($key) + 1);

			if($key === "DTSTAMP" || $key === "DTSTART" || $key === "DTEND") {
				if($key == "DTSTART") {
					$checkin["__INDEX__"] = strtotime($value);
				}

				$value = date("Y-m-d H:i:sA", strtotime($value));
			}

			// Replace escaped values
			$value = str_replace("\,", ",", $value);

			$checkin[$key] = $value;
		}
	}

	echo "done. " . $eventCount . " event(s) processed.\n";
}

ksort($checkins);
reset($checkins);

$firstKey = key($checkins);

// Add header row
$array = array();

foreach($checkins[$firstKey] as $key => $value) {
	$array[] = $key;
}

array_unshift($checkins, $array);
unset($array);

// Write to Excel
$worksheet->fromArray($checkins, EMPTY_CELL_STRING);
$excel->addSheet($worksheet, 0);

// Freeze header row and set bold text for header
echo "Formatting worksheet data...\n";
$excel->setActiveSheetIndex(0);
$excel->getActiveSheet()->freezePane('A2');

$styleArray = array(
	'font' => array(
		'bold' => true
	)
);

$excel->getActiveSheet()->getStyle('A1:I1')->applyFromArray($styleArray);

// Autofit to content for all columns
$array = array("A", "B", "C", "D", "E", "F", "G", "H", "I");

foreach($array as $value) {
	$excel->getActiveSheet()->getColumnDimension($value)->setAutoSize(true);
}

// Apply background coloring
echo "Applying row coloring to checkin dates...\n";

$styleArray = array(
	'fill' => array(
		'type' => PHPExcel_Style_Fill::FILL_SOLID,
		'color' => array('rgb' => 'CCCCFF')
	)
);

$highestRow = $excel->getActiveSheet()->getHighestRow();

$colorToggle = false;
$highestDate = "";

for($i = 0; $i < $highestRow; $i++) {
	$rowData = $excel->getActiveSheet()->getCellByColumnAndRow(1, $i);
	$rowValue = $rowData->getValue();
	$rowValue = substr($rowValue, 0, strpos($rowValue, " "));

	if($highestDate != $rowValue) {
		$colorToggle = !$colorToggle;
		$highestDate = $rowValue;
	}

	if($colorToggle === true) {
		$excel->getActiveSheet()->getStyle('A' . $i . ':I' . $i)->applyFromArray($styleArray);
	}
}

// Write to disk
echo "Writing checkin spreadsheet to disk...\n";
$excelWriter = new PHPExcel_Writer_Excel2007($excel);
$excelWriter->save("output.xlsx");

$scriptEnd = microtime(true);

echo "Done! Run time: " . round($scriptEnd - $scriptStart, 2) . "sec\n";
