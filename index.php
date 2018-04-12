<?php
require "vendor/autoload.php";
use \PhpOffice\PhpSpreadsheet\Spreadsheet;
use \PhpOffice\PhpSpreadsheet\Writer\Xlsx;
// use Symfony\Component\Cache\Simple\FilesystemCache;

// $cache = new Symfony\Component\Cache\Simple\FilesystemCache();



$cacheMethod = \PHPExcel_CachedObjectStorageFactory::cache_in_memory_gzip;                //compress memory usage
\PHPExcel_Settings::setCacheStorageMethod($cacheMethod);
ini_set('memory_limit', '1024M');

// \PhpOffice\PhpSpreadsheet\Settings::setCache($cache);

session_start();

// ini_set('memory_limit', '-1');
// ini_set('max_execution_time', 6000); 

if ($_SERVER["REQUEST_METHOD"] == "POST") {         //To catch the POST Request
	 if(isset($_POST["Action"])){
		$clientRequest = $_POST;

		switch($clientRequest["Action"] ) { 
			case "ADD": 
			Add($clientRequest['CourseNum'],$clientRequest['Section'],$clientRequest['PrimaryInstructor'],$clientRequest['Location'], $clientRequest['Days'] ,$clientRequest['BeginTime'],$clientRequest['EndTime']); 
			// echo 'Got here'; exit;
			break;
			case 'UPDATE':
			Update($clientRequest['LineNumber'],$clientRequest['CourseNum'], $clientRequest['Section'],$clientRequest['PrimaryInstructor'],$clientRequest['Location'],$clientRequest['BeginTime'],$clientRequest['EndTime'], $clientRequest['Days'] );
			break;
			case 'DELETE':
			Delete($clientRequest['LineNum']);
			break;
			case 'SETSHEET':
					$_SESSION['worksheetName'] = $clientRequest['RequestedSheet'];

			break;



			
			 
		
		}
	}
}


$reader = new \PhpOffice\PhpSpreadsheet\Reader\Xlsx();

 

	if(isset($_SESSION['xlsxFile']))
	{  if(!isset($_SESSION['fileSource'])){$_SESSION['fileSource'] = '';}

	if($_SESSION['fileSource'] !== $_SESSION['xlsxFile']){
		$_SESSION['fileSource'] = $_SESSION['xlsxFile'];
			unset($_SESSION['worksheetName']);
		}
		// $worksheets = $reader->listWorksheetNames($_SESSION['xlsx']);
		$_SESSION['worksheets'] = json_encode($reader->listWorksheetNames($_SESSION['fileSource']));

	}
	else{
		header("Location: init.php");

	}


// $reader->setReadDataOnly(true);

// $spreadsheet = \PhpOffice\PhpSpreadsheet\IOFactory::load('F18.xlsx');
//  $_SESSION['spreadsheet'] = $reader->load($fileName);


// 
// $_SESSION['spreadsheet'] = serialize($reader->load($fileName));
//  echo json_encode($spreadsheet->getSheetNames());    //Worksheets in the xlsx files
// print_r($reader->listWorksheetNames($fileName));

if(isset($_SESSION['worksheetName'])){

	unset($_SESSION['worksheet']);     //reset the current working worksheet

$spreadsheet = $reader->load($_SESSION['fileSource']);

 $worksheet = $spreadsheet->getSheetByName($_SESSION['worksheetName']);   //$_SESSION['worksheetName']
 unset($spreadsheet);
				$firstrow = $worksheet->getRowIterator(1)->current();    //prepare for verification
				$cellIterator = $firstrow->getCellIterator();
				try {
					$cellIterator->setIterateOnlyExistingCells(true);
				} catch (Exception $e) { exit;}	
				$columns = array();  //to check with format
				foreach ($cellIterator as $cell) {
					$columns[] =  $cell->getValue();
				}
 
	$format = array("Subject","CourseNum","Section","Credits","Schedule_Type","Instructional_Method","CAP","SWS","Pmt","Linked_Crs","Cross-Listed","Part_of_Term","Notes to Registrar","Order","Days","Begin_Time","End_Time","Location","StartDate","EndDate","PrimaryInstructor","Inst_workload","Inst_%_Responsibility","Instructor_2","Inst_2_workload","Inst_2_%_Resp","Instructor_3","Inst_3_workload","Inst_3_%_Resp","CRN","TERM","TERM_DESC","COLL_CODE","DEPT","Long_title","Subtitle","Expr1036","START_DATE","END_DATE","Section_Text","Prerequisite","CAMPUS","ProvostMessage","RegistrarMessage",null,null,null,null,null,null,null,null,"FACULTY");

	if($columns != $format) {
		echo "<script>alert(\"Sorry ".$_SESSION['worksheetName']." sheet is not compatible with this scheduler!\")</script>"; 
		
	}else{

		//  $_SESSION['worksheet']= serialize((unserialize($_SESSION['spreadsheet']))->getSheetByName('Fall'));
		//   $worksheet = unserialize($_SESSION['worksheet']);
		// $worksheet -> removeRow(83);
						$rows = [];
						$skipfirstline = true;
						$foundEmpty = true;
						foreach ($worksheet->getRowIterator() AS $row) {
							if($skipfirstline){$skipfirstline=false; continue;}  
							
									$cellIterator = $row->getCellIterator();
											try {
												$cellIterator->setIterateOnlyExistingCells(true);
											} catch (Exception $e) {
												break;
											}
									// $cellIterator->setIterateOnlyExistingCells(FALSE); // This loops through all cells,
									$cells = [];
									foreach ($cellIterator as $cell) {
										$cells[] = $cell->getValue();

										if($cell->getCoordinate()[0] == 'S' or $cell->getCoordinate()[0] == 'T' ){
											$worksheet->getStyle($cell->getCoordinate())->getNumberFormat()->setFormatCode('m/dd/yyyy');    //Convert back to Excel standard of date
										}
									}
									if(isset($cells) && !is_null($cells[0])){

										$cells[15]=(\PhpOffice\PhpSpreadsheet\Shared\Date::excelToDateTimeObject($cells[15]))->format("H:i:s");  //starttime
										$cells[16]=(\PhpOffice\PhpSpreadsheet\Shared\Date::excelToDateTimeObject($cells[16]))->format("H:i:s");	 //enddtime
										$cells[18]=(\PhpOffice\PhpSpreadsheet\Shared\Date::excelToDateTimeObject($cells[18]))->format("n/j/Y");  //startdate
										$cells[19]=(\PhpOffice\PhpSpreadsheet\Shared\Date::excelToDateTimeObject($cells[19]))->format("n/j/Y");	 //enddate
										$rows[$row->getRowIndex()] = $cells;
										
									}
									else{
										if ($foundEmpty){
										$_SESSION['blank_row']=$row->getRowIndex();
											$foundEmpty=false;
										}
									}
						}
					$jsonREST = json_encode($rows);

				
					$_SESSION['worksheet']= serialize($worksheet);
					unset($worksheet);
					// $spreadsheet->disconnectWorksheets();            //Disconnect the worksheets adapter
					// unset($spreadsheet);                             //Destroy from the memory
					
	}

}

function Add($courseNum, $section, $instrctor, $location, $days, $begintime, $endtime){   //$subj, $courseNum, $section, $instrctor, $location, $begintime, $endtime
	// global $worksheet, $spreadsheet;
	// global $blank_row;
	// require_once '/vendor/phpoffice/phpspreadsheet/src/Boostrap.php';

	\PhpOffice\PhpSpreadsheet\Cell\Cell::setValueBinder( new \PhpOffice\PhpSpreadsheet\Cell\AdvancedValueBinder() );

	$worksheet = unserialize($_SESSION['worksheet']);
	$blank_row = $_SESSION['blank_row'];
	// echo $worksheet->getCell('A2')->getValue();
	$worksheet->insertNewRowBefore($blank_row, 1);
//    $begin_excelstandard = \PhpOffice\PhpSpreadsheet\Shared\Date::PHPTOExcel(strtotime($begintime)+3600);
//    $end_excelstandard = \PhpOffice\PhpSpreadsheet\Shared\Date::PHPTOExcel(strtotime($endtime)+3600);

	$worksheet->getCell('A'.$blank_row)->setValue(substr($courseNum,0,3));
	 $worksheet->getCell('B'.$blank_row)->setValue(substr($courseNum,3));
	 $worksheet->getCell('C'.$blank_row)->setValue($section);
	 $worksheet->getCell('U'.$blank_row)->setValue($instrctor);
	 $worksheet->getCell('R'.$blank_row)->setValue($location);
	 $worksheet->getCell('O'.$blank_row)->setValue($days);
	 $worksheet->getCell('P'.$blank_row)->setValue(\PhpOffice\PhpSpreadsheet\Shared\Date::PHPTOExcel(strtotime($begintime)+7200)); 
	 $worksheet->getStyle('P'.$blank_row)->getNumberFormat()->setFormatCode(\PhpOffice\PhpSpreadsheet\Style\NumberFormat::FORMAT_DATE_TIME1);
	//  $worksheet->getCell('P'.$blank_row)->setValue($begintime);
	 $worksheet->getCell('Q'.$blank_row)->setValue(\PhpOffice\PhpSpreadsheet\Shared\Date::PHPTOExcel(strtotime($endtime)+7200));
	 $worksheet->getStyle('Q'.$blank_row)->getNumberFormat()->setFormatCode(\PhpOffice\PhpSpreadsheet\Style\NumberFormat::FORMAT_DATE_TIME1);

  updateFile($worksheet);
}

function Delete($row_no){  //linenumber
	$worksheet = unserialize($_SESSION['worksheet']);
	
	$worksheet->removeRow($row_no);
	updateFile($worksheet);
}


function Update($row, $courseNum, $section, $instructor, $location, $begintime, $endtime, $days){ //linenumber, $subj, $courseNum, $section, $instrctor, $location, $begintime, $endtime
	$worksheet = unserialize($_SESSION['worksheet']);

	$worksheet->getCell('A'.$row)->setValue(substr($courseNum,0,3));
	$worksheet->getCell('B'.$row)->setValue(substr($courseNum,3));
	$worksheet->getCell('C'.$row)->setValue($section);
	$worksheet->getCell('U'.$row)->setValue($instructor);
	$worksheet->getCell('R'.$row)->setValue($location);
	$worksheet->getCell('O'.$row)->setValue($days);
	$worksheet->getCell('P'.$row)->setValue(\PhpOffice\PhpSpreadsheet\Shared\Date::PHPTOExcel(strtotime($begintime)+7200)); 
	$worksheet->getStyle('P'.$row)->getNumberFormat()->setFormatCode(\PhpOffice\PhpSpreadsheet\Style\NumberFormat::FORMAT_DATE_TIME1);
   //  $worksheet->getCell('P'.$row)->setValue($begintime);
	$worksheet->getCell('Q'.$row)->setValue(\PhpOffice\PhpSpreadsheet\Shared\Date::PHPTOExcel(strtotime($endtime)+7200));
	$worksheet->getStyle('Q'.$row)->getNumberFormat()->setFormatCode(\PhpOffice\PhpSpreadsheet\Style\NumberFormat::FORMAT_DATE_TIME1);

  updateFile($worksheet);

}

function updateFile($worksheet){
	$spreadsheet = \PhpOffice\PhpSpreadsheet\IOFactory::load($_SESSION['fileSource']);
	// $worksheet = unserialize($_SESSION['worksheet']);

	$index = $spreadsheet->getIndex($spreadsheet->getSheetByName($_SESSION['worksheetName']));
	$spreadsheet->removeSheetByIndex($index);
	
	$spreadsheet->addExternalSheet($worksheet, $index);

	$spreadsheet->setActiveSheetIndexByName($_SESSION['worksheetName']);

	$writer = \PhpOffice\PhpSpreadsheet\IOFactory::createWriter($spreadsheet, 'Xlsx');
	$writer->save($_SESSION['fileSource']);

	$spreadsheet->disconnectWorksheets();            //Disconnect the worksheets adapter
		unset($spreadsheet);
		unset($worksheet);


	echo chmod($_SESSION['fileSource'],0766);

	// var_dump(\PhpOffice\PhpSpreadsheet\Shared\Date::getDefaultTimezone());
	

	//  var_dump($_SESSION['worksheet']);
	 exit;
}

////////////////////////////////////////////////////////////////////////////////////////////////////////PHP END HERE///////////////////////////////////////////////////////////////////////////////////////////////////
?>         

<!doctype html>
<html>
<head>
<meta charset="utf-8">
<meta http-equiv="X-UA-Compatible" content="IE=edge,chrome=1">
<meta name="viewport" content="width=device-width, initial-scale=1">
<title>Scheduler</title>


<script type="text/javascript" src="schedulejs/jquery-3.3.1.min.js"></script>
<script type="text/javascript" src="schedulejs/jquery.calendar.js"></script>
<link rel="stylesheet" media="screen" href="schedulejs/jquery.calendar.css"/>

<!-- TIME PICKER 1 LOCAL SCRIPT -->
<link rel="stylesheet" href="timepicker/jquery.timepicker.css">
<script src="timepicker/jquery.timepicker.js"></script>


<script type="text/javascript" src="schedulejs/jquery-ui.min.js"></script>
<link rel="stylesheet" media="screen" href="schedulejs/jquery-ui.css"/>

<script>
		$('.contentcontainer').hide();
	// This prevents Flash of Unstyled Content by hiding all HTML before everyting can load.  Show is called in function() below.
</script>
<script src="schedulejs/loading/js/jquery.loadingModal.js"></script>


<script type="text/javascript">

var jsonin = <?php if(isset($jsonREST)){ echo $jsonREST; } else { echo '{}';}  ?>;
console.log(jsonin);

	$(function(){

		// console.log($.isEmptyObject(jsonin));
		 if(!$.isEmptyObject(jsonin)){


						//prepare for plugin

						var resources = {};
						var events =[];
						var startdate;
						for (var key in jsonin) {
							var pasretime = Date.parse(jsonin[key][18]);
						 startdate = new Date(Date.parse(jsonin[key][18]));  //set the start date from excel through PHPSpreadsheet
						 if(pasretime > 18000000){
						 break;
						 }
						}
						//  console.log(startdate);
						// console.log($.cal.date(0).addDays(0).format('Y-m-d'));


						function eventNresourceGenerator() {	
							//var totalLine = Object.keys(jsonin).length; //This include the blank lines at the end, required valdidation to put intjs
						// console.log(jsonin);
						var classInsession = {};
						var profInclass = {};

						$.each(jsonin, (key, array)=>{	

								if(array[17] != null){
								var roomKey = (array[17].split(" "))[1];
								var uniquer = (array[17].split(" "))[0];
								var resourceKey;
								if(typeof roomKey != 'undefined' && 3>= roomKey.length <= 4){

									resourceKey = (roomKey.length == 3) ? ((uniquer[0]=='H') ? 'E2'+uniquer[0]+roomKey : uniquer[0]+roomKey) :  ((roomKey[0]=='A') ?  roomKey.replace('A','E') : roomKey);
									resources[resourceKey]= array[17];  
						
									}
								}




						//use return instead of continue because of the function
								if(array[14] && array[0]){

									var days = array[14];         //convert the day into array of characters			
									
								for(index=0; index<days.length; index++){
									var day = days.charAt(index); 
									
									var dayinNo = (day=='M') ? 0 : (day=='T' ? 1 : (day=='W' ? 2 : (day=='R' ? 3 : (day=='F' ? 4 : 5))));
									
									
									if(dayinNo<5 ){       					//only row with Subject field and valid day
										events[events.length] = {
											uid		: key+array[0]+array[1]+array[2]+day,
											begins	:  $.cal.date(startdate).addDays(dayinNo).format('Y-m-d')+' '+ array[15],
											ends	: $.cal.date(startdate).addDays(dayinNo).format('Y-m-d')+' '+ array[16],
											resource : resourceKey,
											notes	: "<label class='coursesec'>"+array[0]+'&nbsp;'+array[1]+'</label>&#09;<label class=section>'+array[2]+"</label>\n"+"<label class = 'room'>" + array[17] + "</label>" +"\n"+"<label class='prof'>"+array[20]+'</label>',
											//title : array[15].substring(0,5) +"-"+ array[16].substring(0,5)
											color :  (function(){
														if((day+array[17]+array[15]) in classInsession && classInsession != undefined ){
															events[classInsession[day+array[17]+array[15]]]['color'] = '#CD0000';     //Redden existing 
															return '#CD0000';  //Redden the conflct
														}
														else{
																if((day+array[20]+array[15]) in profInclass && classInsession != undefined){
																	events[profInclass[day+array[20]+array[15]]]['color'] = 'orange';
																	return   'orange';
																}
																
																	return '#255BA1';
																
														}
													})()
											};

											classInsession[day+array[17]+array[15]] = events.length - 1;  //dayRoomStartTime 
											profInclass[day+array[20]+array[15]] = events.length - 1;//dayprofstarttime
											

										}



									}

								}
							});
								// console.log(classInsession);
								//  console.log(classInsession);
								// return events;
						}

						eventNresourceGenerator();
						// console.log(resources);


						var sortedResources = function(){                        //sorted the Json object to start with MAK A1105
									var keys = [];
									var sorted_obj = {};

										for (var key in resources){
											if(resources.hasOwnProperty(key)){
												keys.unshift(key);
											}
										}

										keys.sort();
										//  keys.push(keys.shift());
										//  keys.reverse();


										$.each(keys, (i, key)=>{

												sorted_obj[key] = resources[key];
												

										});

										return sorted_obj;
						}();
						console.log(sortedResources);

			/////preparation ended here


	var red=[], orange=[], blue=[];
	var trackprof = {};	var prof = false; var selected = false;

		
	$('#calendar').cal({           				//plugin-call
		
		resources : sortedResources,

		startdate   :    startdate,  

		daytimestart  : '08:00:00',
		daytimeend    : '21:30:00',
		maskdatelabel : 'l',

		allowcreation : 'both',  
		allowresize		: false,
		allowmove		: false,
		allowselect		: true,
		allowremove		: false,
		allownotesedit	: false,
		allowhtml   : true, 


		daystodisplay : 5,

		maskeventlabelend : ' - '+'g:i A',

		 minwidth : 140,

		minheight   : 20,

		eventcreate : function( ){
			console.log('creation event hit');
		},
		
		eventmove : function( uid ){
			console.log( 'Moved event: '+uid, arguments );
		},

		eventselect : function(){
									// $("label.coursesec, label.prof").on('click',
								
								selected = true;
								if(red.length>0){
																red.forEach(function(obj){
																	$(obj).css('background-color', '#CD0000');          //'rgb(249, 6, 6)'
																	$(obj).parent().css('background-color','rgb(194, 10, 10)'); 
																});
															
														}
								 if(orange.length>0){
															orange.forEach(function(item){
																	$(item).css('background-color', 'orange');          //'rgb(249, 6, 6)'
																	$(item).parent().css('background-color', 'rgb(242, 162, 13)'); 
																});
																
														}
								 if(blue.length>0){
															blue.forEach(function(item){
																	$(item).css({'background-color': 'rgb(40, 114, 210)', 'color':'white'});          //'rgb(249, 6, 6)'
																	$(item).parent().css({'background-color': 'rgb(37, 91, 162)', 'color':'white'}); 
																});
																
														}

		},
		
		eventremove : function( uid ){
			console.log( 'Removed event: '+uid );
		},
		
		
		eventnotesedit : function( uid ){
			console.log( 'Edited Notes for event: '+uid );
		},

		eventdraw   : function (){
			console.log('Draw event');
		},


		
		
		// Load events as .ics
		events : events
	
		});
	
										//  $("label.coursesec, label.prof").css( 'pointer-events', 'auto' );
									
									$("label.coursesec, label.prof").on('click' ,function(e){
										if(selected){										
											e.stopPropagation();
											selected=false;

										}
										else{
											//  var orange = ($(this).parent().css('background-color') === 'rgb(250, 182, 56)') ? true : false;
	 
											 var whichclass = $(this).attr('class');
											  prof = (whichclass == 'prof') ? true : false;
	 
											 // console.log(Object.keys(trackprof).length);
											 if(Object.keys(trackprof).length >0){  //to retore the color in the next click after name click
	 
												 document.querySelectorAll('[data-id]').forEach((mainblock)=>{
														 if($(mainblock).attr('data-id') in trackprof){
													 $(mainblock).attr('style',(trackprof[$(mainblock).attr('data-id')]).title);
													 $(mainblock).children('.details').attr('style',(trackprof[$(mainblock).attr('data-id')]).note);
													 // console.log($(mainblock).attr('style'));
														 }
	 
												 });
												 
												 }
										if(prof){
											if(red.length>0){
																red.forEach(function(obj){
																	$(obj).css('background-color', '#CD0000');          //'rgb(249, 6, 6)'
																	$(obj).parent().css('background-color','rgb(194, 10, 10)'); 
																});
																
														}
										 if(orange.length>0){
															orange.forEach(function(item){
																	$(item).css('background-color', 'orange');          //'rgb(249, 6, 6)'
																	$(item).parent().css('background-color', 'rgb(242, 162, 13)'); 
																});
																
														}
										  if(blue.length>0){
															blue.forEach(function(item){
																	$(item).css({'background-color': 'rgb(40, 114, 210)', 'color':'white'});          //'rgb(249, 6, 6)'
																	$(item).parent().css({'background-color': 'rgb(37, 91, 162)', 'color':'white'}); 
																});
																
														}

												}

												 
												 
									document.querySelectorAll("label."+whichclass).forEach((session)=>{
									if($(this).html()===$(session).html()){		
												 // console.log($(session).html());
									if($(session).parent().css('background-color') === 'rgb(249, 6, 6)') red.push($(session).parent());
									if($(session).parent().css('background-color') === 'rgb(250, 182, 56)') orange.push($(session).parent());
									if($(session).parent().css('background-color') === 'rgb(40, 114, 210)') blue.push($(session).parent());

									 if(prof){trackprof[$(session).parent().parent().attr('data-id')]={'title' : $(session).parent().parent().attr('style'), 'note': $(session).parent().attr('style')} ;}

											  var title = prof ? 'limegreen' : 'gold';
											  var note = prof ? 'lime' : 'yellow';
												 $(session).parent().css({'background-color': note, 'color':'black'});
												 $(session).parent().parent().css({'background-color': title, 'color':'black'});
												 }
											
										});  //for each loop end here
	 
											
												 if(!prof){
														 trackprof={};
												 }	
											}

									});
		
		toalternate=true;
		document.querySelectorAll("div.ui-cal-label-date").forEach((dayCol)=>{
			// console.log($(el).html().substring(0,3));
			// $(dayCol).children('p').html($(dayCol).children('p').html().substring(0,3));

			if(toalternate){
				$(dayCol).css({"background-color" : 'light grey', 'color': '#255BA1', 'font-weight' : 'bold'}); toalternate=false;
			}
			else{
			 $(dayCol).css({"background" : "linear-gradient(to bottom right, #0575E6, #0575E6)", 'color': 'white', 'font-weight' : 'bold'}); toalternate=true;
			}
		
			// $(el).html($(el).html().substring(0,3));
		});

				document.querySelectorAll("span.button-remove").forEach((btn)=>{  
				
				$(btn).before().after('<div class="button-remove"></div>');
				$(btn ).remove();
				

				});
						

//-----------------------------------------------------------------------------------------------------------------------------------------		
//-----------------------------------------------------------------------------------------------------------------------------------------			
//-----------------------------------------------------------------------------------------------------------------------------------------		
			// Tyler Section --------------------------------------------------------------------------------------------------------------
//-----------------------------------------------------------------------------------------------------------------------------------------		
//-----------------------------------------------------------------------------------------------------------------------------------------		
//-----------------------------------------------------------------------------------------------------------------------------------------	
$('.timepicker').timepicker({
				timeFormat: 'h:mm p',
				interval: 10,
				minTime: '8:00am',
				maxTime: '10:00pm',
				defaultTime: '8:00am',
				startTime: '8:00am',
				dynamic: true,
				dropdown: false,
    			scrollbar: false
			});
$('.timepickerend').timepicker({
				timeFormat: 'h:mm p',
				interval: 10,
				minTime: '8:00am',
				maxTime: '10:00pm',
				defaultTime: '8:50am',
				startTime: '8:50am',
				dynamic: true,
				dropdown: false,
    			scrollbar: false
});

			// Event handlers for Calendar
			// Title click to allow editing of the calendar events
			$("p.title").on('click', function(){
				

				var parent = $(this).parent();
				var details = parent.children( ".details" );
				var dataid = parent.attr("data-id")
				//alert(this.innerHTML);	
				var lineNum = dataid.trim().substring(0, 3).match(/\d+/);
				var key = dataid.substring(0, ((dataid.length)-1));
				//$('#editCourse').dialog('open');
				//alert($(this).parent().childNodes[3].innerHTML);


				var queryall = document.querySelectorAll("div[data-id]");
				var matches = [];

				queryall.forEach( 
  					function(currentValue, currentIndex, listObj) { 
						//console.log(currentValue.getAttribute("data-id"));
						var dataid = currentValue.getAttribute("data-id");

						if(dataid.substring(0,(dataid.length-1)) == key){
							matches.push(currentValue.getAttribute("data-id"));
						}

  					},
  						
				);

				var days = '';
				matches.forEach( 
  					function(currentValue, currentIndex, listObj) { 
						//console.log(currentValue.getAttribute("data-id"));
						//console.log(currentValue);
						days += (currentValue.substr(-1));
  					},
  						
				);

				var coursesecdetails = details.children( ".coursesec").text();
				var roomdetails = details.children( ".room" ).text();
				var profdetails = details.children( ".prof" ).text();
				var sectiondetails = details.children( ".section" ).text();
				var time = parent.children( ".title").text();
				var begin = (time.substr(0, (time.indexOf("-")-1))).trim();
				var end = (time.substr((time.indexOf("-")+1), time.length)).trim();
				$("#editLineNum").val(lineNum);

				$("#editsNum").val(sectiondetails);
				$("#editcNum").val(coursesecdetails);
				$("#editprof").val(profdetails);
				$('#editrNum option:contains("roomdetails")');
				$("#editdays").val(days);
				//$("#editdays").val(parent.attr("data-id").substr(-1));
				$("#editbeginTime").val(begin);
				$("#editendTime").val(end);
				//$("select[name='editRoom']").find("option[value='" + roomdetails + "']").attr("selected",true);
				//$('select').val(roomdetails);
				$("#editRoomAuto").val(roomdetails);
				$('#editCourse').dialog('open');

			});

			// Delete button click to allow deletion of courses.  
			$(".button-remove").on('click', function(e){

				$("#deleteLine").val($(this).parent().attr("data-id").trim().substring(0, 3).match(/\d+/))
				$('#deleteDialog').dialog('open');
			});
				
				
			// Initializes the add course Dialog
			$("#addCourse").dialog({
        			autoOpen: false,
					modal: true,
					width: 400,
					zIndex: 9999,
			});


			$("#editCourse").dialog({
        			autoOpen: false,
					modal: true,
					width: 400,
					zIndex: 9999,
        	
			});
		
			// Initializes the delete course Dialog
		    $("#deleteDialog").dialog({
        		modal: true,
        		bgiframe: true,
        		width: 500,
        		height: 200,
        		autoOpen: false,
				zIndex: 9999,
    		});

			// Adds the buttons for the add course dialog
			$("#addCourse").dialog('option', 'buttons', {
            	"Add Course" : function() {
					//var ddl = document.getElementById("roomNum");
					//var room = ddl.options[ddl.selectedIndex].text;

					var beginTime = $('#beginTime').val();
					var endTime = $('#endTime').val();
					if(beginTime == endTime && beginTime != '' && endTime != ''){
						alert("Begin time cannot equal end time!");
						return;
					}
					//console.log(room);
					// alert("done");
					modaltext = 'Adding course to \"'+selection+'\" spreadsheet';
					modalbkcolor = '#255BA1';
					spinner = spinners[1];
				$.ajax({
						type: "POST",
						data: {
						Action: "ADD",
						CourseNum: $('#cNum').val(),
						Section: $('#sNum').val(),
						PrimaryInstructor: $('#prof').val(),
						Location: $('#addroomnumauto').val(),
						Days: $("#days").val(),
						BeginTime: beginTime,
						EndTime: endTime,
						}, 
						
					}).done(function(data){
						// console.log(data);
						// $("#addCourse").dialog('close');
						refresh();
					});

            },
            	"Cancel" : function() {
                $(this).dialog("close");
            	}
        	});


			// Adds the buttons for the add course dialog
			$("#editCourse").dialog('option', 'buttons', {
            	"Confirm Edit" : function() {
					//var ddl = document.getElementById("editrNum");
					//var room = ddl.options[ddl.selectedIndex].text;
			modaltext = "Updating changes to "+selection;
			modalbkcolor = '#255BA1';
			spinner = spinners[2];
			$.ajax({type: "POST",
						data: {
						Action: "UPDATE",
						LineNumber: $("#editLineNum").val(),
						CourseNum: $('#editcNum').val(),
						Section: $('#editsNum').val(),
						PrimaryInstructor: $('#editprof').val(),
						Location: $("#editRoomAuto").val(),
						BeginTime: $('#editbeginTime').val(),
						EndTime: $('#editendTime').val(),
						Days: $("#editdays").val()
						}, 
						success: function(data){
							if(data){
							refresh();
							}							
					},
			});

			},
				
            	"Cancel" : function() {
                $(this).dialog("close");
            	}
        	});

			
			// Adds the button for confirm or cancel to the deleteDialog
			// Handles action of confirm.
			$("#deleteDialog").dialog('option', 'buttons', {
            	"Yes" : function() {
					// Confirm yes, will not close modal window until server says action is completed.
					modaltext = 'Deleting course from '+selection ;
					modalbkcolor = '#800000';
					spinner = spinners[9];
					$.ajax({type: "POST",
						data: {
						Action: "DELETE",
						LineNum: $("#deleteLine").val()
					}, 
					success: function(data){
						if(data){
						refresh();
						}
					},
					error: function(data){
						alert("Action could not be completed.");
						$(this).dialog("close");

					}
					});
				
            },
            	"No" : function() {
                $(this).dialog("close");
            	}
        	});


			// Loops through resources JSON and adds each room with key to the ddls
			// var addRooms = document.getElementById("roomNum");
			// var editRooms = document.getElementById("editrNum");
			// for (var key in resources){
			// 	var option1 = document.createElement("option");
			// 	option1.value = resources[key];
			// 	option1.text = resources[key];
			// 	var option2 = document.createElement("option");
			// 	option2.value = resources[key];
			// 	option2.text = resources[key];

			// 	addRooms.add(option1);
			// 	editRooms.add(option2);
			// }
	// This prevents FOUC 
	 $('.contentcontainer').css("visibility", "visible");
	
		}  //if  statement for json obj validation
	
			/*****************************************************************************************************
	******************************************************************************************************
	******************************************************************************************************
	*
	*        Section to Dynamically add buttons for sheets selection
	*			Dear Chit, this section requires that the buttonSection div be added wherever you
				wish in the HTML section.  
	******************************************************************************************************
	******************************************************************************************************
	*****************************************************************************************************/
	//NOTE this is the array that the button names come from 
	var sheets = <?php echo $_SESSION['worksheets']; ?>;
	var selection = "<?php if(isset($_SESSION['worksheetName'])){echo $_SESSION['worksheetName'];}else{echo "none";} ?>";

	$('#buttonSection').append('File Picker: <input type="button" onclick="location.href=\'init.php\';" value=<?php echo basename($_SESSION["fileSource"], ".xlsx")?> (current) /> SpreadSheets: ');
	sheets.forEach((sheet)=>{
		if(selection == sheet){
		$('#buttonSection').append('<button type="button" class="sheetName" disabled>'+sheet+'</button>');
		}
		else{
			$('#buttonSection').append('<button type="button" class="sheetName">'+sheet+'</button>');

		}
		});

	$('button.sheetName').on('click', function(){
			var name = ($(this).html());
			modaltext = "Loading spreadsheet : "+name;
			modalbkcolor = '#006600';
			spinner = spinners[7];
	$.ajax({type: "POST",
						data: {
					Action: "SETSHEET",
					RequestedSheet: name
					}, 
					success: function(data){
				 refresh();
					}
		});
	
	});

	


	// for(var i = 0; i < sheets.length; i++){
	// var expandClient = '<button type="button"  id="sheetButton'+ i +'"><i class="fa fa-plus-square-o"></i>' + sheets[i] + '</button>';

	// $('#buttonSection').append(expandClient);

	// $('#sheetButton'+i).click(function() {
	// 	var requestedSheet = this.id;
	// 		//Sends the id of the button.
	// 		//Requested sheet corresponds with index of sheet name
	// 		//from the sheets array
	// 	$.ajax({type: "POST",
	// 				data: {
	// 				Action: "SETSHEET",
	// 				RequestedSheet: requestedSheet
	// 				}, 
	// 				success: function(data){

	// 				}
	// 	});
	// 	console.log(this.id);
 
	// });

	// }
	/***************************************************************
	****************************************************************
	****************************************************************
	*
	*		END OF DYNAMICALLY ADDED BUTTON SECTION
	*
	*
	****************************************************************
	****************************************************************
	***************************************************************/
	//AutoComplete Section
	var queryprofs = document.querySelectorAll("label.prof");
				var profs = [];

				queryprofs.forEach(
					function(currentValue, currentIndex, listObj) { 
						//console.log(currentValue.getAttribute("data-id"));
						var prof = currentValue.innerHTML;
						// if(dataid.substring(0,(dataid.length-1)) == key){
						// 	matches.push(currentValue.getAttribute("data-id"));
						// }
						if(!profs.includes(prof)){
							profs.push(prof);
							console.log(prof);
						}

  					}
				)

			    $( "#prof" ).autocomplete({
      			source: profs
    			});	

				$("#editprof").autocomplete({
					source: profs
				});

		var queryRooms = document.querySelectorAll("label.room");
				var rooms = [];

				queryRooms.forEach(
					function(currentValue, currentIndex, listObj) { 
						//console.log(currentValue.getAttribute("data-id"));
						var room = currentValue.innerHTML;
						// if(dataid.substring(0,(dataid.length-1)) == key){
						// 	matches.push(currentValue.getAttribute("data-id"));
						// }
						if(!rooms.includes(room)){
							rooms.push(room);
							//console.log(prof);
						}

  					}
				)


		$("#addroomnumauto").autocomplete({
			source: rooms
		});

		$("#editRoomAuto").autocomplete({
			source: rooms
		});
				







			$('#days, #editdays, #cNum, #editcNum').keyup(function() {
				this.value = this.value.toUpperCase();	
			});

			var spinners = ['rotatingPlane','wave', 'wanderingCubes','spinner','chasingDots','threeBounce','circle', 'cubeGrid','fadingCircle','foldingCube'];
			var modaltext, modalbkcolor, spinner;
			$(document).ajaxStart(function(){
				$( document ).blur();
				//spinIndex = (spinIndex < spinners.length) ? (spinIndex++) : (spinIndex=0);
				$('body').loadingModal({
						position: 'auto',
						text: modaltext,
						color: '#fff',
						opacity: '0.99',
						backgroundColor: modalbkcolor,
						animation: spinner
						});
						
				
			});




//-----------------------------------------------------------------------------------------------------------------------------------------		
//-----------------------------------------------------------------------------------------------------------------------------------------		
//-----------------------------------------------------------------------------------------------------------------------------------------		
//-----------------------------------------------------------------------------------------------------------------------------------------		
//-----------------------------------------------------------------------------------------------------------------------------------------		
//-------------------------------------------//The end of jQuery Document Ready scope ----------------------------------------------------------------------------------------------		
// ----------------------------------------------------------------------------------------------------------------------------------------	

   }); 		  




	

		

function refresh(){
	location.replace("index.php");
}


//the root script scope in html</script>   
<link rel="stylesheet" href="schedulejs/loading/css/jquery.loadingModal.css">

<style type="text/css">
html,body{
font-family: 'Roboto';
font-size: 10px;
overflow: hidden;

}
#calendar{
position: absolute;
top: 70px;
left: 15px;
right: 15px;
bottom: 20px;
border: 1px solid #bbb;
}

.ui-cal .ui-cal-event .button-remove{
position: absolute;
top: 3px; right: 3px;
width: 10px;
height: 10px;
background: transparent url(data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAAAkAAAAKCAYAAABmBXS+AAAC7mlDQ1BJQ0MgUHJvZmlsZQAAeAGFVM9rE0EU/jZuqdAiCFprDrJ4kCJJWatoRdQ2/RFiawzbH7ZFkGQzSdZuNuvuJrWliOTi0SreRe2hB/+AHnrwZC9KhVpFKN6rKGKhFy3xzW5MtqXqwM5+8943731vdt8ADXLSNPWABOQNx1KiEWlsfEJq/IgAjqIJQTQlVdvsTiQGQYNz+Xvn2HoPgVtWw3v7d7J3rZrStpoHhP1A4Eea2Sqw7xdxClkSAog836Epx3QI3+PY8uyPOU55eMG1Dys9xFkifEA1Lc5/TbhTzSXTQINIOJT1cVI+nNeLlNcdB2luZsbIEL1PkKa7zO6rYqGcTvYOkL2d9H5Os94+wiHCCxmtP0a4jZ71jNU/4mHhpObEhj0cGDX0+GAVtxqp+DXCFF8QTSeiVHHZLg3xmK79VvJKgnCQOMpkYYBzWkhP10xu+LqHBX0m1xOv4ndWUeF5jxNn3tTd70XaAq8wDh0MGgyaDUhQEEUEYZiwUECGPBoxNLJyPyOrBhuTezJ1JGq7dGJEsUF7Ntw9t1Gk3Tz+KCJxlEO1CJL8Qf4qr8lP5Xn5y1yw2Fb3lK2bmrry4DvF5Zm5Gh7X08jjc01efJXUdpNXR5aseXq8muwaP+xXlzHmgjWPxHOw+/EtX5XMlymMFMXjVfPqS4R1WjE3359sfzs94i7PLrXWc62JizdWm5dn/WpI++6qvJPmVflPXvXx/GfNxGPiKTEmdornIYmXxS7xkthLqwviYG3HCJ2VhinSbZH6JNVgYJq89S9dP1t4vUZ/DPVRlBnM0lSJ93/CKmQ0nbkOb/qP28f8F+T3iuefKAIvbODImbptU3HvEKFlpW5zrgIXv9F98LZua6N+OPwEWDyrFq1SNZ8gvAEcdod6HugpmNOWls05Uocsn5O66cpiUsxQ20NSUtcl12VLFrOZVWLpdtiZ0x1uHKE5QvfEp0plk/qv8RGw/bBS+fmsUtl+ThrWgZf6b8C8/UXAeIuJAAAACXBIWXMAAAsTAAALEwEAmpwYAAAAvElEQVQYGU2QLQvCUBiF793CitH/IGhRMK34V0wu+Iu0idUoCAaLzWAYjAVtwyIDh01Qr8+Zd7ADz85533v2wYxBzrkIZnCEq3fNkc5VCGEBjb5NwJcQqpT45QPPfb7gpc+JSqkfNngMa5jACqRUpXcd/5e5/4Rpa/cKWFY6QBnoNT28gDNIlUrbOhpzx0dwghg0Szu9rg9F6/HteGMY1FXCEPbwAUl+gLEKlhBYa/VvOsy6qwsl5OyfOv8B52rwq5FQN5gAAAAASUVORK5CYII=) no-repeat 50% 50%;
cursor: pointer;
opacity: 0.6;

}

.ui-cal .ui-cal-event .button-remove:hover,
.ui-cal .ui-cal-event.selected .button-remove{
opacity: 1;

}




</style>

</head>

<body>


<h1 style="margin:0px auto 0 auto; text-align:center;">CIS Department Scheduler</h1>

<div id="buttonSection" >	<button class="contentcontainer" style="border-radius: 10px; color: blue; background-color: white; height: 3em; width: 125px;" onclick="$('#addCourse').dialog('open');">Add Course</button>
	<button class="contentcontainer" style="border-radius: 10px; color: blue; background-color: white; height: 3em; width: 125px;margin-right: 200px;" onclick="refresh()">Refresh</button></div>

<div class="contentcontainer" style="visibility: hidden">

	<div style="margin:10px auto 100px auto; text-align:center;">

	</div>



	<div id="calendar"></div>

<!-- ****************************************************************************************************** -->
<!-- ****************************************************************************************************** -->
<!-- ****************************************************************************************************** -->
<!--  Dialogs -->
<!-- ****************************************************************************************************** -->
<!-- ****************************************************************************************************** -->
<!-- ****************************************************************************************************** -->

  <div id="addCourse" title="Add A Course">
  	<form>
	  		<div style="margin-left: auto; margin-right:auto; width: 66%;">
	  		<div style="float: right;">
			<label for="beginTime">Begin Time:</label>
			<input align="right" class="timepicker" type="text" id="beginTime" pattern="\d?\d:\d\d\s(AM|PM)" required>
			</div>
			</br>
			</br>
			<div style="float: right;">
			<label for="endTime">End Time:</label>
			<input class="timepickerend" type="text" id="endTime" pattern="\d?\d:\d\d\s(AM|PM)" required>
			</div>
			</br>
			</br>
			<div style="float: right;">
			<label for="day">Days:</label>
			<input type="text" id="days" pattern="((M)|(T)|(W)|(R)|(F))+" placeholder="MTWRF" required>
			</div>
			</br>
			</br>
			<div style="float: right;">
			<label for="cNum">Course Number:</label>
			<input type="text" id="cNum" pattern="[A-Z][A-Z][A-Z]\s\d\d\d" placeholder="CIS 173" required>
			</div>
			</br>
			</br>
			<div style="float: right;">
			<label for="sNum">Section Number:</label>
			<input type="text" id="sNum" pattern="\d\d" placeholder="01" required>
			</div>
			</br>
			</br>
			<div style="float: right;">
			<div class="ui-widget">
			<label for="prof">Professor:</label>
			<input type="text" id="prof" required>
			</div>
			</div>
			</br>
			</br>
			<div style="float: right;">
			<label for="roomNum">Room Number:</label>
			<!--<select name="addroom" id="roomNum"></select>-->
			<input type="text" id="addroomnumauto" pattern="[A-Z]{2,3}\s[A-Z|1-9]{3,6}">
			</div>
			</div>
		
	</form>
</div>

<div id="deleteDialog" title="Confirm Delete">
<input type="hidden" id="deleteLine">
<h2>This action will delete all instances of this course!</h2>
<p>Are you sure you want to perform this action?</p>
<p>To delete a single day, edit the course days.</p>
</div>





<div id="editCourse" title="Edit A Course">
  	<form>
	  <input id="editLineNum" type="hidden">
	  <div style="margin-left: auto; margin-right:auto; width: 66%;">
	  <div style="float: right;">

			<label for="editbeginTime">Begin Time</label>
			<input class="timepicker" type="text" id="editbeginTime" pattern="\d?\d:\d\d\s(AM|PM)" required>
			</div>
			</br>
			</br>
			<div style="float: right;">
			<label for="editendTime">End Time:</label>
			<input class="timepickerend" type="text" id="editendTime" pattern="\d?\d:\d\d\s(AM|PM)" required>
			</div>
			</br>
			</br>
			<div style="float: right;">
			<label for="editdays">Days:</label>
			<input type="text" id="editdays" pattern="((M)|(T)|(W)|(R)|(F))+" required>
			</div>
			</br>
			</br>
			<div style="float: right;">
			<label for="editcNum">Course Number:</label>
			<input type="text" id="editcNum" pattern="[A-Z][A-Z][A-Z]\s\d\d\d">
			</div>
			</br>
			</br>	
			<div style="float: right;">
			<label for="editsNum">Section Number:</label>
			<input type="text" id="editsNum" pattern="\d\d" required>
			</div>
			</br>
			</br>
			<div style="float: right;">
			<div class="ui-widget">
			<label for="editprof">Professor:</label>
			<input type="text" id="editprof" required>
			</div>
			</div>
			</br>
			</br>
			<div style="float: right;">
			<label for="editrNum">Room Number:</label>
			<!--<select name="editRoom" id="editrNum"></select>-->
			<input type="text" id="editRoomAuto" pattern="[A-Z]{2,3}\s[A-Z|1-9]{3,6}">
			</div>
			</div>
	</form>
</div>
<!-- ****************************************************************************************************** -->
<!-- ****************************************************************************************************** -->
<!-- ****************************************************************************************************** -->
<!-- ****************************************************************************************************** -->
<!-- ****************************************************************************************************** -->
<!-- ****************************************************************************************************** -->


</div>

</body>

</html>
