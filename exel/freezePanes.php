
<?php 
require_once 'Spreadsheet/Excel/Writer.php';

// Creating a workbook
$workbook = new Spreadsheet_Excel_Writer();

// sending HTTP headers
$workbook->send('Anibal-exel.xls');



$worksheet =& $workbook->addWorksheet("Some Worksheet");

/* ... */

/* This freezes the first six rows of the worksheet: */
$worksheet->freezePanes(array(6, 0));

/* To freeze the first column, one must use the following syntax: */
$worksheet->freezePanes(array(0, 1));

/* Freeze the first six rows and start the scrollable region at row nine: */
$worksheet->freezePanes(array(6, 0, 9, 0));

$workbook->close();

?>