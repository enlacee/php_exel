<?php 
require_once 'Spreadsheet/Excel/Writer.php';

// Creating a workbook
$workbook = new Spreadsheet_Excel_Writer();

// sending HTTP headers
$workbook->send('Anibal-exel.xls');
//OPCIONAL
// $workbook->setTempDir('/home/xnoguer/temp');

//ESTILOS EXEL
$negrita =& $workbook->addFormat();
$negrita->setBold();


// Creating a worksheet
$worksheet =& $workbook->addWorksheet('My first worksheet');




$worksheet->write(10, 2, 12286.26, $format_total_box);


$worksheet->mergeCells ( 10 , 2 , 13 , 5 );




$workbook->close();

?>

