<?php
require('XLSXReader.php');					// PHP library to fetch data from a spreadsheet
require('credentials.php');					// NOTE: Modify this: Database credentials
$xlsx = new XLSXReader('BALL_POS.xlsx');	// NOTE: Modify this: Excel file
$data = $xlsx->getSheetData('Feuil1');		// NOTE: Modify this: Sheet name
$prefix = DB_PREFIX;

// print_r($data)

// Create connection
$conn = new mysqli(DB_HOSTNAME, DB_USERNAME, DB_PASSWORD, DB_DATABASE, DB_PORT);
$conn->set_charset("utf8");

// Check connection
if($conn->connect_error) {
	die("CONNECTION FAILED: " . $conn->connect_error);
}

// String to concatenate values and thus insert bulk data, instead of a single insert statement for each row
$values = '';

for($i = 1; $i < count($data); $i++) { // Row 0 on sheet contains column names
	// Column name (from db table) = Cell value (from spreadsheet)
	// The following variables are sorted as they appear on the db table
	$name 			= mysqli_real_escape_string($conn, trim($data[$i][0]));
	$address		= mysqli_real_escape_string($conn, trim($data[$i][2]));
	$city			= mysqli_real_escape_string($conn, trim($data[$i][4]));
	$postcode		= mysqli_real_escape_string($conn, trim($data[$i][3]));
	$country		= mysqli_real_escape_string($conn, trim($data[$i][5]));
	$zone			= mysqli_real_escape_string($conn, trim($data[$i][6]));
	$phone			= mysqli_real_escape_string($conn, trim($data[$i][10]));
	$image			= mysqli_real_escape_string($conn, trim($data[$i][1]));
	$is_service		= strtolower(substr($data[$i][17], 0, 1)) == 'y' ? 1 : 0;
	$service_phone	= mysqli_real_escape_string($conn, trim($data[$i][18]));
	$service_email	= mysqli_real_escape_string($conn, trim($data[$i][19]));
	$service_hours	= mysqli_real_escape_string($conn, trim($data[$i][20]));
	$is_store		= strtolower(substr($data[$i][21], 0, 1)) == 'y' ? 1 : 0;
	$store_phone	= mysqli_real_escape_string($conn, trim($data[$i][22]));
	$store_hours	= mysqli_real_escape_string($conn, trim($data[$i][23]));
	$lat			= $data[$i][8];
	$lng			= $data[$i][9];
	// Concatenate two websites
	$website		= mysqli_real_escape_string($conn, trim($data[$i][24] . $data[$i][25]));
	// Get country_id. If country name doesn't match then country_id = 0
	$res = $conn->query("SELECT country_id FROM oc_country WHERE LOWER(name) = LOWER('$country')")->fetch_assoc();
	$country_id = empty($res) ? 0 : $res['country_id'];
	// Get zone_id. If name doesn't match or country_id is not found then zone_id = 0
	$res = $conn->query("SELECT zone_id FROM oc_zone WHERE LOWER(name) = LOWER('$zone') AND country_id = $country_id")->fetch_assoc();
	$zone_id = empty($res) ? 0 : $res['zone_id'];
	// Concatenate values for SQL query
	$values = $values . "('$name', '$address', '$city', '$postcode', $country_id, $zone_id, '$phone', '$image', $is_service, '$service_phone', '$service_email', '$service_hours', $is_store, '$store_phone', '$store_hours', '$website', $lat, $lng),";
}

// Remove last comma
$values = substr($values, 0, -1);
// echo $values;

// Run insert statement
$sql = "INSERT INTO " . $prefix . "store_locator
	(name, address, city, postcode, country_id, zone_id, phone, image, is_service, service_phone, service_email, service_hours, is_store, store_phone, store_hours, website, lat, lng)
	VALUES " . $values;

if($conn->query($sql) === TRUE) {
	echo "New record created successfully";
} else {
	echo "Error while running SQL query: " . $conn->error;
}

// Close connection
$conn->close();

// References:
// https://www.igorkromin.net/index.php/2017/12/07/how-to-pass-parameters-to-your-php-script-via-the-command-line
// https://stackoverflow.com/questions/13869170/how-to-work-with-if-definedb-host-localhost
// https://stackoverflow.com/questions/2528213/php-read-xlsx-excel-2007-file
?>