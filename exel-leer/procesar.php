<?php 
require_once 'Spreadsheet/Excel/reader.php';

// Creating a workbook
//$workbook = new Spreadsheet_Excel_Writer();

	$data = new Spreadsheet_Excel_Reader();	
	$data->read("exel.xls"); 
	
	
	//Obtenemos un libro del Excel  
	$sheet = $data->sheets[0];
	
	//Obtenemos las CELDAS  
	$cells = $data->sheets[0]['cells'];
	
	
	//Obtenemos una FILA determinada  
	$row = $data->sheets[0]['cells'][0]; 
	
	//Obtenemos una CELDA CONCRETA  
	$data->sheets[0]['cells'][0][0];  
	
	
	//Obtenemos el número de filas y columnas del Excel  
	$nrows = $data->sheets[0]['numRows'];  
	$ncols = $data->sheets[0]['numCols'];
	
	
	echo "<pre>";
	
	echo "hola...";
	
	echo "fila 		= $nrows";
	
	echo "columna 	= $ncols";
	
	echo  "Obtenemos una FILA determinada  = $row";
	
	echo "Obtenemos un libro del Excel shee";
	echo "<hr><br>";
	
	print_r($data);
	echo "</pre>"; 
	