<html>
  <head>
    <style type="text/css">
    table {
    	border-collapse: collapse;
    }        
    td {
    	border: 1px solid black;
    	padding: 0 0.5em;
    }        
    </style>
  </head>
  <body>
    <table>
    <?php
require_once 'Spreadsheet/Excel/reader.php';    
    // initialize reader object
    $excel = new Spreadsheet_Excel_Reader();
    
    // read spreadsheet data
    $excel->read('exel.xls');    
    
    // iterate over spreadsheet cells and print as HTML table
    $x=1;
    while($x<=$excel->sheets[0]['numRows']) { //Cuenta el # de filas = 6
      echo "\t<tr>\n";
      $y=1;
      while($y<=$excel->sheets[0]['numCols']) { //Cuenta el # de Columna = 2
        $cell = isset($excel->sheets[0]['cells'][$x][$y]) ? $excel->sheets[0]['cells'][$x][$y] : '';
        echo "\t\t<td>$cell</td>\n";  
        $y++;
      }  
      echo "\t</tr>\n";
      $x++;
    }
    ?>    
    </table>
  </body>
</html>