<?php 
require_once 'Spreadsheet/Excel/Writer.php';

// Creating a workbook
$workbook = new Spreadsheet_Excel_Writer();

// sending HTTP headers
$workbook->send('Plantilla-exel.xls');

//------------------------------------- Estilos Inicio -------------------------------------//
$negrita =& $workbook->addFormat();
$negrita->setBold();

//-- Colores RGB
$workbook->setCustomColor(11, 0,0,150);
$workbook->setCustomColor(12, 192,192,192);
$workbook->setCustomColor(13, 221,60,16);
$workbook->setCustomColor(14, 255,255,0);




//--- format_titulo
$format_titulo = $workbook->addFormat();
$format_titulo->setSize(14);
$format_titulo->setColor("black");
$format_titulo->setBold();

//--format_rojo
$format_rojo = $workbook->addFormat();
$format_rojo->setSize(8);
$format_rojo->setColor("red");
$format_rojo->setBold();

//--format_simple
$format_simple = $workbook->addFormat();
$format_simple->setSize(8);
$format_simple->setColor("black");


//--format_decimal
$format_decimal =& $workbook->addFormat();
$format_decimal->setNumFormat('$#,##0.00;$-#,##0.00');
$format_decimal->setSize(8);

//--format_decimal_total_azul
$format_decimal_total_azul =& $workbook->addFormat();
$format_decimal_total_azul->setNumFormat('$#,##0.00;$-#,##0.00');
$format_decimal_total_azul->setColor(11);
$format_decimal_total_azul->setBold();
$format_decimal_total_azul->setSize(8);
$format_decimal_total_azul->setTop(1);
$format_decimal_total_azul->setBottom(2);
$format_decimal_total_azul->setBorderColor('black');
//$format_decimal_total_azul->setPattern(1);
//$format_decimal_total_azul->setBorder(1);
//$format_decimal_total_azul->setFgColor(12);

//--format_decimal_total_rojo
$format_decimal_total_rojo =& $workbook->addFormat();
$format_decimal_total_rojo->setNumFormat('$#,##0.00;$-#,##0.00');
$format_decimal_total_rojo->setColor("red");
$format_decimal_total_rojo->setBold();
$format_decimal_total_rojo->setSize(8);
$format_decimal_total_rojo->setTop(1);
$format_decimal_total_rojo->setBottom(2);
$format_decimal_total_rojo->setBorderColor('black');



//-- format_porcent
$format_porcent =& $workbook->addFormat();
$format_porcent->setNumFormat('#,##0.00%;-#,##0.00%');
$format_decimal->setSize(8);


//--format_rojo_decimal
$format_rojo_decimal =& $workbook->addFormat();
$format_rojo_decimal->setNumFormat('$#,##0.00;$-#,##0.00');
$format_rojo_decimal->setSize(8);
$format_rojo_decimal->setColor("red");
$format_rojo_decimal->setBold();

//--format_rojo_decimal_2
$format_rojo_decimal_2 =& $workbook->addFormat();
$format_rojo_decimal_2->setNumFormat('($#,##0.00);($-#,##0.00)');
$format_rojo_decimal_2->setSize(8);
$format_rojo_decimal_2->setColor("red");
$format_rojo_decimal_2->setBold();





//--format_fecha
$format_fecha =& $workbook->addFormat();
$format_fecha->setNumFormat('dd/mm/yyyy');
$format_fecha->setAlign('right');

//--format_fecha_2
$format_fecha_2 = & $workbook->addFormat( array( 'NumFormat' => 'dd/mm/yy' ));

//--format_fecha_3
$format_fecha_3 =& $workbook->addFormat();
$format_fecha_3->setNumFormat('ddd,dd-mmm-yy');
$format_fecha_3->setAlign('center');



//--format_total_box
$format_total_box =& $workbook->addFormat();
$format_total_box->setSize(8);
$format_total_box->setBold();
//$format_total_box->setBorder(2);
$format_total_box->setTop(1);
$format_total_box->setBottom(2);
$format_total_box->setBorderColor('black');

//--format_tabla_head
$format_tabla_head =& $workbook->addFormat();
$format_tabla_head->setBold();
$format_tabla_head->setColor(11);
$format_tabla_head->setFgColor(12);
$format_tabla_head->setAlign('center'); //ALINEACION



//-- Format_txt_Centrado
$format_tabla_head_centrado =& $workbook->addFormat();
$format_tabla_head_centrado->setBold();
$format_tabla_head_centrado->setSize(8);
$format_tabla_head_centrado->setTextWrap(1);
$format_tabla_head_centrado->setBorder(1);
$format_tabla_head_centrado->setColor(11);
$format_tabla_head_centrado->setFgColor(12);
$format_tabla_head_centrado->setBgColor (12);
$format_tabla_head_centrado->setVAlign('vequal_space');
//$format_tabla_head_centrado->setHAlign('equal_space');
$format_tabla_head_centrado->setAlign('center');


//--format_line_separador
$format_line_separador =& $workbook->addFormat();
$format_line_separador->setBottom(1);
$format_line_separador->setBorderColor('black');


											  
//-- format_semana
$format_semana =& $workbook->addFormat();
$format_semana->setBold();
$format_semana->setSize(8);
$format_semana->setTextWrap(1);
$format_semana->setColor(11);
$format_semana->setFgColor(12);
//$format_semana->setBgColor ("red");
$format_semana->setVAlign('vequal_space');
//$Format_semana->setHAlign('equal_space');
$format_semana->setAlign('center');

//-- format_fecha_amarillo
$format_fecha_amarillo =& $workbook->addFormat();
$format_fecha_amarillo->setSize(8);
$format_fecha_amarillo->setNumFormat('ddd,dd-mmm-yy');
$format_fecha_amarillo->setAlign('center');
$format_fecha_amarillo->setFgColor(14); //14 = amarillo rgb

//-- format_decimal_amarillo
$format_decimal_amarillo =& $workbook->addFormat();
$format_decimal_amarillo->setNumFormat('$#,##0.00;$-#,##0.00');
$format_decimal_amarillo->setSize(8);
$format_decimal_amarillo->setFgColor(14);


//-- format_simple_amarillo
$format_simple_amarillo =& $workbook->addFormat();
$format_simple_amarillo->setSize(8);
$format_simple_amarillo->setFgColor(14);














//------------------------------------- Estilos Final -------------------------------------//

// Creating a worksheet
$worksheet =& $workbook->addWorksheet('Cuponeate 01');





//------------------------------------- Plantilla -------------------------------------//

$imagen_bmp ="img/estado-cuenta.bmp";
$x=0;
$y=0;
$scale_x=1;
$scale_y=1;
$worksheet->insertBitmap ( 0, 0, $imagen_bmp, $x, $y, $scale_x, $scale_y );


$worksheet->write(5, 3, 'ESTADO DE CUENTA XXXXX',$format_titulo);


$worksheet->write(6, 1, 'Garantia', $format_rojo);
$worksheet->write(6, 2, 10000, $format_rojo_decimal);

$worksheet->write(7, 1, 'Visa', $format_simple);
$worksheet->write(8, 1, 'Mastercard', $format_simple);
$worksheet->write(9, 1, 'American Express', $format_simple);

$worksheet->write(7, 2, 12224.64,$format_decimal);
$worksheet->write(8, 2, 61.62, $format_decimal);
$worksheet->write(9, 2, 0.00, $format_decimal);

$worksheet->write(10, 2, 12286.26, $format_total_box);


// tabla Principall Lote
// -- Estableciendo formato en Columnas
$worksheet->setColumn(0,0,1);
$worksheet->setColumn(1,1,15);
$worksheet->setColumn(2,2,15);
$worksheet->setColumn(3,3,15);
$worksheet->setColumn(4,4,15);
$worksheet->setColumn(5,5,15);
$worksheet->setColumn(6,6,15);
$worksheet->setColumn(7,7,15);
$worksheet->setColumn(8,8,15);
$worksheet->setColumn(9,9,15);
$worksheet->setColumn(10,10,15);
$worksheet->setColumn(11,11,15);
$worksheet->setColumn(12,12,15);
$worksheet->setColumn(13,13,15);
$worksheet->setColumn(14,14,15);
$worksheet->setColumn(15,15,15);
$worksheet->setColumn(16,16,18);

//-- Estableciendo formato fila 
$worksheet->setRow( 12, 25 );
//$worksheet->setRow ( integer $row , integer $height , mixed $format=0 )


$array = array( "LOTES PROCESADOS",
				"FECHA DE PROCESO",
				"MARCA",
				"NUMERO DE PEDIDOS",
				"(+) VENTAS DEL DIA",
				"% TASA PROM COMISION",
				"(-) COSTO POR TRX",
				"(-)% DSCTO COMISION",
				"(-)IGV",
				"(-)DEVOLUCIONES Y CONTRACARGOS",
				"(+)DSCTO COMISION DEV.",
				"(+)IGV COMISION DEV.",
				"NETO A PAGAR",
				"FECHA DE ABONO",
				"VENTAS ACUMULADAS",
				" xx xx xxxx "
				);
$worksheet->writeRow( 12, 1, $array , $format_tabla_head_centrado );
//-- visa
$worksheet->write( 13, 1, 500 , 			$format_simple );
$worksheet->write( 13, 2, '04/09/2011', 	$format_fecha );
$worksheet->write( 13, 3, 'visa', 			$format_simple );
$worksheet->write( 13, 4, 1291.62 , 		$format_decimal );
$worksheet->write( 13, 5, 0.08756, 			$format_porcent );

$worksheet->write( 13, 6, 500 , 			$format_rojo_decimal_2 );
$worksheet->write( 13, 7, 430 , 			$format_rojo_decimal_2 );
$worksheet->write( 13, 8, 500 , 			$format_rojo_decimal_2 );
$worksheet->write( 13, 9, 500 , 			$format_rojo_decimal_2 );
$worksheet->write( 13, 10, 500 , 			$format_rojo_decimal_2 );
$worksheet->write( 13, 11, 14.23 ,			$format_decimal );
$worksheet->write( 13, 12, 13.58 ,			$format_decimal );
$worksheet->write( 13, 13, 128.253 , 		$format_decimal );
$worksheet->write( 13, 14, '14/09/2011' ,	$format_fecha_3 );
$worksheet->write( 13, 15, 1291.62 , 		$format_decimal );
$worksheet->write( 13, 16, '04/09/2011' ,	$format_fecha );

//-- Mastercard
$worksheet->write( 14, 1, 500 , 			$format_simple );
$worksheet->write( 14, 2, '04/09/2011', 	$format_fecha );
$worksheet->write( 14, 3, 'Mastercard', 	$format_simple );
$worksheet->write( 14, 4, 1291.62 , 		$format_decimal );
$worksheet->write( 14, 5, 0.08756, 			$format_porcent );

$worksheet->write( 14, 6, 500 , 			$format_rojo_decimal_2 );
$worksheet->write( 14, 7, 430 , 			$format_rojo_decimal_2 );
$worksheet->write( 14, 8, 500 , 			$format_rojo_decimal_2 );
$worksheet->write( 14, 9, 500 , 			$format_rojo_decimal_2 );
$worksheet->write( 14, 10, 500 , 			$format_rojo_decimal_2 );
$worksheet->write( 14, 11, 14.23 ,			$format_decimal );
$worksheet->write( 14, 12, 13.58 ,			$format_decimal );
$worksheet->write( 14, 13, 128.253 , 		$format_decimal );
$worksheet->write( 14, 14, '04/09/2011' ,	$format_fecha_3 );
$worksheet->write( 14, 15, 1291.62 , 		$format_decimal );
$worksheet->write( 14, 16, '04/09/2011' ,	$format_fecha );

//-- American Express
$worksheet->write( 15, 1, 500 , 			$format_simple );
$worksheet->write( 15, 2, '04/09/2011', 	$format_fecha );
$worksheet->write( 15, 3, 'American Express',$format_simple );
$worksheet->write( 15, 4, 1291.62 , 		$format_decimal );
$worksheet->write( 15, 5, 0.08756, 			$format_porcent );

$worksheet->write( 15, 6, 500 , 			$format_rojo_decimal_2 );
$worksheet->write( 15, 7, 430 , 			$format_rojo_decimal_2 );
$worksheet->write( 15, 8, 500 , 			$format_rojo_decimal_2 );
$worksheet->write( 15, 9, 500 , 			$format_rojo_decimal_2 );
$worksheet->write( 15, 10, 500 , 			$format_rojo_decimal_2 );
$worksheet->write( 15, 11, 14.23 ,			$format_decimal );
$worksheet->write( 15, 12, 13.58 ,			$format_decimal );
$worksheet->write( 15, 13, 128.253 , 		$format_decimal );
$worksheet->write( 15, 14, '04/09/2011' ,	$format_fecha_3 );
$worksheet->write( 15, 15, 1291.62 , 		$format_decimal );
$worksheet->write( 15, 16, '04/09/2011' ,	$format_fecha );

//--Total
$worksheet->write( 16, 1, "",	 			$format_simple );
$worksheet->write( 16, 2, '', 				$format_simple );
$worksheet->write( 16, 3, '',				$format_simple );
$worksheet->write( 16, 4, 2300.02 , 		$format_decimal_total_azul );
$worksheet->write( 16, 5, 0.08756, 			$format_porcent );

$worksheet->write( 16, 6, 500 , 			$format_decimal_total_rojo );
$worksheet->write( 16, 7, 430 , 			$format_decimal_total_rojo );
$worksheet->write( 16, 8, 500 , 			$format_decimal_total_rojo );
$worksheet->write( 16, 9, 500 , 			$format_decimal_total_rojo );
$worksheet->write( 16, 10, 500 , 			$format_decimal_total_azul );
$worksheet->write( 16, 11, "",				$format_decimal_total_azul );
$worksheet->write( 16, 12, "",				$format_decimal_total_azul );
$worksheet->write( 16, 13, "", 				$format_decimal );
$worksheet->write( 16, 14, "",				$format_fecha_3 );
$worksheet->write( 16, 15, "", 				$format_decimal );
$worksheet->write( 16, 16, "",				$format_fecha );

//--linea Separacion Lote
$worksheet->write( 17, 1, "",	 	$format_line_separador );
$worksheet->write( 17, 2, "", 		$format_line_separador );
$worksheet->write( 17, 3, "",		$format_line_separador );
$worksheet->write( 17, 4, "", 		$format_line_separador );
$worksheet->write( 17, 5, "", 		$format_line_separador );

$worksheet->write( 17, 6, "", 		$format_line_separador );
$worksheet->write( 17, 7, "", 		$format_line_separador );
$worksheet->write( 17, 8, "", 		$format_line_separador );
$worksheet->write( 17, 9, "", 		$format_line_separador );
$worksheet->write( 17, 10, "", 		$format_line_separador );
$worksheet->write( 17, 11, "",		$format_line_separador );
$worksheet->write( 17, 12, "",		$format_line_separador );
$worksheet->write( 17, 13, "", 		$format_line_separador );
$worksheet->write( 17, 14, "",		$format_line_separador );
$worksheet->write( 17, 15, "", 		$format_line_separador );
$worksheet->write( 17, 16, "",		$format_line_separador );

//--celda separador semana
$worksheet->write(20, 1, 'SEMANA XXXX XXXX XXXX ', $format_semana);
// Merge cells from row 0, col 0 to row 2, col 2
$worksheet->setMerge(20, 1, 20, 16);




//-- texto amarillo Depositado
$worksheet->write(23, 4, 'DEPOSITADO', $format_simple_amarillo);
$worksheet->write(23, 5, 2500000, $format_decimal_amarillo);



/*
$worksheet->write(0, 0, 'Name',$format_title);
$worksheet->write(0, 1, 'Age',$format_title2);
$worksheet->write(0, 2, 'precio', $format_title3);

$worksheet->write(1, 0, 'Copitan Norabuena ANIBAL');
$worksheet->write(1, 1, 30,$format_decimal);
$worksheet->write(1, 2, 2.2,$format_decimal);

$worksheet->write(2, 0, 'Johann Schmidt');
$worksheet->write(2, 1, 31,$format_decimal);
$worksheet->write(2, 2, 10/12/2011);

$worksheet->write(3, 0, 'Juan Herrera');
$worksheet->write(3, 1, 32,$format_decimal);
*/


$workbook->close();

?>


