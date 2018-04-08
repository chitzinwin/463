<?
ini_set('session.gc_maxlifetime', 3600);
session_set_cookie_params(3600);

session_start();
ini_set('memory_limit', '512M');
if ($_SERVER["REQUEST_METHOD"] == "POST") {
    // $entityBody = file_get_contents('php://input');
    // echo $entityBody;
    if(isset($_POST["filename"]) && isset($_POST["value"]) ){
         $_SESSION["xlsxFile"] = $_POST["value"];
         echo  file_exists($_SESSION["xlsxFile"]);    //final confirmation the file exsitence
        exit;
    }
    
    echo false; exit;
    
}


$files = glob('spreadsheets/*.xlsx');
//  print(json_encode($files));
$json = new stdClass();
 foreach($files as $file){
    // chown($file, 'daemon');
    // chmod($file,0766);
   $name = basename($file, '.xlsx');
    $json->$name =$file;
 }

// print json_encode($json);
// print $json->F18;
 $jsonfromPHP = json_encode($json);


?>

<!doctype html>
<html lang="en">
<head>
<script type="text/javascript" src="schedulejs/jquery-3.3.1.min.js"></script>
<script type="text/javascript" src="schedulejs/jquery-ui.min.js"></script>
<link rel="stylesheet" media="screen" href="schedulejs/jquery-ui.css"/>


<script>
var files = <?php echo $jsonfromPHP; ?>;   //array of filename with path

var buttons = new Object();
    $.each(files, (key, value)=>{	
        buttons[key]= function () {

                            $.ajax({
                            type: 'post',
                            url: 'init.php',
                            data: {filename: key, value : value},
                            success: function(data) {
                                console.log(data);
                            if(data){
                            location.replace("index.php");      //only file selection is successful switch to next page
                             }
                                 }                     
                            

                        });
              }
}
);

// buttons[buttonsength]= {
//       text: "Ok",
//       icons: {
//         primary: "ui-icon-heart"
//       },
//       click: function() {
//         $( this ).dialog( "close" );
//       }
 
//       // Uncommenting the following line would hide the text,
//       // resulting in the label being used as a tooltip
//       //showText: false
//     }

console.log(files);

$( function() {
    var dialog =  $( "#dialog" ).dialog({
        autoOpen: true,
        width: 500,
        modal: true,
        buttons : buttons
    });
  

    
 } );


 

        
</script>
<style>

    .ui-dialog .ui-dialog-buttonpane .ui-dialog-buttonset
    {
        float:none; 

    }
    .ui-dialog .ui-dialog-buttonpane
    {
        text-align:center;
         vertical-align:bottom;
      line-height: 3.5em;
      display: flex;
  justify-content: center;
  align-items: center;
    }
 
 .ui-button{
     display: block;
 }


.ui-dialog-content::-webkit-scrollbar {
  -webkit-appearance: none;
  width: 22px;
  height: 22px;
}

.ui-dialog-content::-webkit-scrollbar-thumb {
  border-radius: 8px;
  border: 2px solid white; /* should match background, can't be transparent */
  background-color: rgba(0, 0, 0, .5);
}
</style>
</head>
<body>
<div id="dialog" title="File(.xlsx) found in spreadsheet directory"></div>
</body>

</html>