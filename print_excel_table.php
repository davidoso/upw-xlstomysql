<?php
date_default_timezone_set('UTC');
require('XLSXReader.php');
$xlsx = new XLSXReader('BALL_POS.xlsx');
$sheetNames = $xlsx->getSheetNames();
?>

<!DOCTYPE html>
<html>
<head>
	<title>XLSXReader Sample</title>
	<style>
		body {
			font-family: Helvetica, sans-serif;
			font-size: 12px;
		}
		table, td {
			border: 1px solid #000;
			border-collapse: collapse;
			padding: 2px 4px;
		}
	</style>
</head>
<body>
<h1>XLSXReader Sample</h1>
<h2>Sheet data</h2>

<?php
foreach($sheetNames as $sheetName) {
	$sheet = $xlsx->getSheet($sheetName);
	$data = $sheet->getData();
	echo '<table>';
	foreach($data as $row) {
		echo '<tr>';
		foreach($row as $cell) {
			echo '<td>' . $cell . '</td>';
		}
		echo '</tr>';
	}
	echo '</table>';
	echo '<br><br>';
}
?>

</body>
</html>