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


// The actual data
$worksheet->write(0, 0, 'Name',$negrita);
$worksheet->write(0, 1, 'Age',$negrita);
$worksheet->write(0, 2, 'precio', $negrita);

$worksheet->write(1, 0, 'Copitan Norabuena ANIBAL');
$worksheet->write(1, 1, 30);
$worksheet->write(1, 2, 2.2);

$worksheet->write(2, 0, 'Johann Schmidt');
$worksheet->write(2, 1, 31);

//ARRAY 
$array = array();

for ($i=0; $i < 10 ;$i++){
	$array[]= "hola".$i;
}


$worksheet->writeRow( 3, 1, $array , $negrita );

//$workbook->setVersion(8);
// Let's send the file
$workbook->close();





?>

