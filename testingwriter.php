
<!-- <?php
	// require_once ('PHP_XLSXWriter/xlsxwriter.class.php');
    // $headers = array('id' => 'string', 'name' => 'string', 'description' => 'string', 'n1' => 'string', 'n2' => 'string', 'n3' => 'string', 'n4' => 'string', 'n5' => 'string', 'n6' => 'string', 'n7' => 'string');
    // $sheet_names = array('january', 'february', 'march', 'april', 'may', 'june');
    // $start = microtime(true);
    // $writer = new XLSXWriter();
    // foreach ($sheet_names as $sheet_name) {
    //     $writer->writeSheetHeader($sheet_name, $headers);
    //     for ($i = 0; $i < 10000; $i++) {
    //         $writer->writeSheetRow($sheet_name, random_row());
    //     }
    // }
    // $writer->writeToFile('test.xlsx');
    // file_put_contents("php://stderr", '#' . floor(memory_get_peak_usage() / 1024 / 1024) . "MB" . "\n");
    // file_put_contents("php://stderr", '#' . sprintf("%1.2f", microtime(true) - $start) . "s" . "\n");
    // function random_row()
    // {
    //     return $row = array(rand() % 10000, chr(rand(97, 122)) . chr(rand(97, 122)) . chr(rand(97, 122)) . chr(rand(97, 122)) . chr(rand(97, 122)) . chr(rand(97, 122)) . chr(rand(97, 122)), md5(uniqid()), rand() % 10000, rand() % 10000, rand() % 10000, rand() % 10000, rand() % 10000, rand() % 10000, rand() % 10000);
    // }

?> -->

<?php
// ini_set('memory_limit', '-1');
// ini_set('max_execution_time', 6000); 


require "vendor/autoload.php";
use PhpSpreadsheet\Spreadsheet;
use PhpSpreadsheet\Writer\Xlsx;



$fileName = 'F18.xlsx';
$reader = new \PhpOffice\PhpSpreadsheet\Reader\Xlsx();
//$reader->setReadDataOnly(true);

// $spreadsheet = \PhpOffice\PhpSpreadsheet\IOFactory::load('F18.xlsx');
$spreadsheet = $reader->load($fileName);

// print_r($spreadsheet->getSheetNames());    //Worksheets in the xlsx files


// echo json_encode($reader->listWorksheetNames($fileName));




$worksheet = $spreadsheet->getSheet(0);

$highestCol = $worksheet->getHighestColumn();
$row = $worksheet->getRowIterator(1)->current();




$cellIterator = $row->getCellIterator();
$cellIterator->setIterateOnlyExistingCells(true);


$columns = array();
foreach ($cellIterator as $cell) {
    $columns[] =  $cell->getValue();
}

print_r($columns);

$format = array("Subject","CourseNum","Section","Credits","Schedule_Type","Instructional_Method","CAP","SWS","Pmt","Linked_Crs","Cross-Listed","Part_of_Term","Notes to Registrar","Order","Days","Begin_Time","End_Time","Location","StartDate","EndDate","PrimaryInstructor","Inst_workload","Inst_%_Responsibility","Instructor_2","Inst_2_workload","Inst_2_%_Resp","Instructor_3","Inst_3_workload","Inst_3_%_Resp","CRN","TERM","TERM_DESC","COLL_CODE","DEPT","Long_title","Subtitle","Expr1036","START_DATE","END_DATE","Section_Text","Prerequisite","CAMPUS","ProvostMessage","RegistrarMessage",null,null,null,null,null,null,null,null,"FACULTY");
echo "\r\n";
print_r($format);

var_dump($columns==$format);
// $dataArray = $worksheet->toArray();

// var_dump($dataArray);

$highestRow = $worksheet->getHighestRow(); // e.g. 10
$highestColumn = $worksheet->getHighestColumn(); // e.g 'F'
$highestColumnIndex = \PhpOffice\PhpSpreadsheet\Cell\Coordinate::columnIndexFromString($highestColumn); // e.g. 5



print $highestColumn;





// print $highestColumnIndex;

// $worksheet -> removeRow(4);     //Seriously delete row using old library



// echo "<table border='1'>";
// foreach ($worksheet->getRowIterator() as $row) {
//      echo '<tr>' . PHP_EOL;
     
//     $cellIterator = $row->getCellIterator();
//     echo "<td>".$row->getRowIndex()."</td>";
//     @$cellIterator->setIterateOnlyExistingCells(TRUE); // This loops through all cells,
//                                                        //    even if a cell value is not set.
//                                                        // By default, only cells that have a value
//                                                        //    set will be iterated.
                                    
//     foreach ($cellIterator as $cell) {
//         echo '<td>' .
//              $cell->getValue()
//              .'</td>' . PHP_EOL;
//     }
//     echo '</tr>' . PHP_EOL;
// }
// echo "</table>";

// date_default_timezone_set('EST');



$rows = [];
  $skipfirst = true;
foreach ($worksheet->getRowIterator() AS $row) {
   if($skipfirst){$skipfirst=false; continue;}  
  
    $cellIterator = $row->getCellIterator();
	try {
        $cellIterator->setIterateOnlyExistingCells(true);
    } catch (Exception $e) {
        break;
    }
   
    // var_dump($cellIterator->getIterateOnlyExistingCells( ));
 
    // $cellIterator->setIterateOnlyExistingCells(TRUE);
    
    //$cellIterator->setIterateOnlyExistingCells(TRUE); // This loops through all cells,
    $cells = [];
    foreach ($cellIterator as $cell) {
        $cells[] = $cell->getValue();

        if($cell->getCoordinate()[0] == 'S' or $cell->getCoordinate()[0] == 'T' ){
            $worksheet->getStyle($cell->getCoordinate())->getNumberFormat()->setFormatCode('m/dd/yyyy');    //Convert back to Excel standard of date
        }
        // if($cell->getCoordinate()[0] == 'S' or $cell->getCoordinate()[0] == 'T' ){

        //     $worksheet->getStyle($cell->getCoordinate())->getNumberFormat()->setFormatCode('m/dd/yyyy');
        //    // echo $cell->getValue();
        // }
        
    }
    // print(json_encode($cells[0]));
        
    if(!is_null($cells[0])){
    $cells[15]=(\PhpOffice\PhpSpreadsheet\Shared\Date::excelToDateTimeObject($cells[15]))->format("H:i:s");
    $cells[16]=(\PhpOffice\PhpSpreadsheet\Shared\Date::excelToDateTimeObject($cells[16]))->format("H:i:s");
    $cells[18]=((\PhpOffice\PhpSpreadsheet\Shared\Date::excelToDateTimeObject($cells[18]))->format("n/j/Y"));
    $cells[19]=(\PhpOffice\PhpSpreadsheet\Shared\Date::excelToDateTimeObject($cells[19]))->format("D M d Y h:i:s");
    $rows[$row->getRowIndex()] = $cells;
    }
    
    // $filtered = array_filter($cells, function($var){return !is_null($var);} );

    //  break;


}

// var_dump($row);


// $skipfirst = true;
// foreach ($rows as $row){
//     if($skipfirst){$skipfirst=false; continue;}          
//            $row[15]=(\PhpOffice\PhpSpreadsheet\Shared\Date::excelToDateTimeObject($row[15]))->format("H:i:s");
//            $row[16]=(\PhpOffice\PhpSpreadsheet\Shared\Date::excelToDateTimeObject($row[16]))->format("H:i:s");

      

// }

// echo datetime()->format();
// echo date('D M t Y h:i:s');
// var_dump($worksheet)
// $spreadsheet->disconnectWorksheets();
// unset($spreadsheet);
// print_r($rows);
$jsonREST = json_encode($rows);
// echo $jsonREST;
// foreach($rows AS $row){
//     foreach ($row as $cell){
//         print $cell;
//     }
// }

//For writing


// $returnDateType = \PhpOffice\PhpSpreadsheet\Calculation\Functions::getReturnDateType();
// var_dump($returnDateType);
//\PhpOffice\PhpSpreadsheet\Calculation\Functions::setReturnDateType('E');

//date_default_timezone_set('America/New_York');
// PhpOffice\PhpSpreadsheet\Shared\TimeZone::setTimeZone('America/New_York');

//  print \PhpOffice\PhpSpreadsheet\Shared\TimeZone::getTimeZone();

//  $baseDate = \PhpOffice\PhpSpreadsheet\Shared\Date::getExcelCalendar();

//  print $baseDate;

// $getexceltime =  $worksheet->getCell('P2')->getValue();
// // $phptime = \PhpOffice\PhpSpreadsheet\Shared\Date::ExcelToPHP($getexceltime);
// // $InvDate = date($format='h:mm:ss AM/PM',  $phptime); 

// //  print $InvDate;
// echo $getexceltime;

//  if (abs(($getexceltime-0.60416666666667)/0.60416666666667) < 0.00001){
     
//    // $getexceltime==0.60416666666667) {
//      echo 'True';
//  }


 $nutime = '2:30 PM';
// // // $now = time();
$exceltime = \PhpOffice\PhpSpreadsheet\Shared\Date::PHPTOExcel(strtotime($nutime)+3600);   //PHP/Unix (UST time) to Excel timestamp

// print $exceltime;
// print PHPExcel_Calculation_DateTime::TIMEVALUE($nutime);
//  printexceltime);   
// // var_dump(\PhpOffice\PhpSpreadsheet\Shared\Date::getDefaultTimezone());

// $worksheet->getCell('P2')->setValue($exceltime);
// $worksheet->getStyle('P2')->getNumberFormat()->setFormatCode(\PhpOffice\PhpSpreadsheet\Style\NumberFormat::FORMAT_DATE_TIME1);




// // $phptime = \PhpOffice\PhpSpreadsheet\Shared\Date::excelToTimestamp('0.60416666666667');  //Excel timestamp to PHP/Unix timestamp

// // // print date('g:i A',$phptime-3600);

// // $worksheet->getCell('A1')->setValue('John');
// // $worksheet->getCell('A2')->setValue('Smith');

$writer = \PhpOffice\PhpSpreadsheet\IOFactory::createWriter($spreadsheet, 'Xlsx');
$writer->save('test.xlsx');

// chmod("test.xlsx",0766);








// print_r($worksheet);


?>

<script type="text/javascript" src="http://code.jquery.com/jquery.min.js"></script>
<script type="text/javascript" src="schedulejs/jquery.calendar.js"></script>
<link rel="stylesheet" media="screen" href="schedulejs/jquery.calendar.css"/>


<script>
window.onload = ()=>{
var jsonin = <?php echo $jsonREST; ?>;

//var jjson = JSON.parse(jsonin);
console.log(Object.keys(jsonin).length);
var events=[];
console.log(jsonin);

$.each(jsonin, (key, array)=>{

       //use return instead of continue because of the function
     if(array[14]){

            var days = array[14];         //convert the day into array of characters           
            
            for(index=0; index<days.length; index++){
             var day = days.charAt(index); 
            
			var dayinNo = (day=='M') ? 0 : (day=='T' ? 1 : (day=='W' ? 2 : (day=='R' ? 3 : (day=='F' ? 4 : 5))));

			if(array[0] && dayinNo<5){       //only row with Subject field and valid day
				events[events.length] = {
					uid		: key+array[0]+array[1]+array[2]+day,
					begins	:  $.cal.date().addDays(dayinNo).format('Y-m-d')+ array[15],
					ends	: $.cal.date().addDays(dayinNo).format('Y-m-d')+ array[16],
					resource : array[17],
					notes	: array[0]+array[1]+array[2]+"\n"+array[17]+"\n"+array[20],
					color	: '#990066'
					};
			    }


            }
        }
        //console.log(events);
       // console.log(array[17].split(" ")[1]);
     // if(array[17] != )
 });


   var startdate;
						for (var key in jsonin) {
						 startdate = new Date(Date.parse(jsonin[key][18]));  //set the start date from excel through PHPSpreadsheet
						 break;
						} 

		//console.log(events);
console.log(startdate);
//console.log( $.cal.date());

// console.log( $.cal.date().addDays(2-$.cal.date().format('N')));
// console.log($.cal.date().addDays(0).format('Y-m-d')+' 12:30:00');
};
</script>
