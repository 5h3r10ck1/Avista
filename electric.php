<?php

include_once 'Classes/RDA.php';
include_once 'config.php';
include_once 'Classes/PHPExcel.php';
include_once 'functions/functions.php';

set_time_limit(0);
$time_start = microtime(true);


try{
$eType = PHPExcel_IOFactory::identify($eFile);
$eReader = PHPExcel_IOFactory::createReader($eType);
$eReader->setLoadSheetsOnly($eSheetName);
$eData = $eReader->load($eFile)->getActiveSheet();
}
catch(PHPExcel_Reader_Exception $e){
	var_dump($e->getMessage());
	$log = fopen('errorLog.txt','a+');
	$message = date('Y-m-d H:i:s') . ' Error loading Electric File: '.$e->getMessage()."\r\n";
	fwrite($log, $message);
	fclose($log);
	die();
}

$eRDA = new RDA($user, $password, $server);
$eRDA1 = new RDA($user, $password, $server);
$eRDA2 = new RDA($user, $password, $server);
$eRDA3 = new RDA($user, $password, $server);







$eZones = $eRDA->GetZoneList();
foreach($eZones as $zone){
	if($zone['ZoneName'] == 'Avista-Electric-Chart') $eZoneID = $zone['ZoneID'];
}



$eRDA->setZoneIDs($eZoneID);
$eRDA1->setZoneIDs($eZoneID);
$eRDA2->setZoneIDs($eZoneID);
$eRDA3->setZoneIDs($eZoneID);


$eB1 = 'Electric Crew Schedule';
$eB2 = 'Electric Servicemen';
$eB3 = 'Electric General Information';

$eBulletins1 = array();
$eBulletins2 = array();
$eBulletins3 = array();

$eBulletins = $eRDA->GetBulletinList();


if (isAssoc($eBulletins[$eZoneID])){
	$eBulletins[$eZoneID] = array($eBulletins[$eZoneID]);
}

foreach ($eBulletins[$eZoneID] as $bulletin) {
	
	if($bulletin['Description'] == $eB1) array_push( $eBulletins1, $bulletin['GUID'] );
	if($bulletin['Description'] == $eB2) array_push( $eBulletins2, $bulletin['GUID'] );
	if($bulletin['Description'] == $eB3) array_push( $eBulletins3, $bulletin['GUID'] );
	
}



$i = 0;
$images = array();
foreach ($eData->getDrawingCollection() as $drawing) {
	$imagePath = imageConvert($drawing,'Electric');
	array_push($images, $imagePath);
}








$eRDA1->setTemplateName('Schedule');

$eRDA1->Description = $eB1;

$eRDA1->setBlock('Crew Schedule Title', $eB1);

$eRDA1->setBlock( 'Month 1',$eData->getCell('B5')->getCalculatedValue() );
$eRDA1->setBlock( 'Month 2',$eData->getCell('C5')->getCalculatedValue() );
$eRDA1->setBlock( 'Month 3',$eData->getCell('D5')->getCalculatedValue() );
$eRDA1->setBlock( 'Month 4',$eData->getCell('E5')->getCalculatedValue() );
$eRDA1->setBlock( 'Month 5',$eData->getCell('F5')->getCalculatedValue() );

$eRDA1->setBlock( 'Date 1',dateConvert($eData->getCell('B6')->getCalculatedValue()) );
$eRDA1->setBlock( 'Date 2',dateConvert($eData->getCell('C6')->getCalculatedValue()) );
$eRDA1->setBlock( 'Date 3',dateConvert($eData->getCell('D6')->getCalculatedValue()) );
$eRDA1->setBlock( 'Date 4',dateConvert($eData->getCell('E6')->getCalculatedValue()) );
$eRDA1->setBlock( 'Date 5',dateConvert($eData->getCell('F6')->getCalculatedValue()) );

$eRDA1->setBlock( 'Crew 1', blockConcat($eData,0,7,5) );
$eRDA1->setBlock( 'Crew Names 1', blockConcat($eData,1,7,5) );
$eRDA1->setBlock( 'Crew Names 2', blockConcat($eData,2,7,5) );
$eRDA1->setBlock( 'Crew Names 3', blockConcat($eData,3,7,5) );
$eRDA1->setBlock( 'Crew Names 4', blockConcat($eData,4,7,5) );
$eRDA1->setBlock( 'Crew Names 5', blockConcat($eData,5,7,5) );
$eRDA1->setBlock( 'Notes 1', blockConcat($eData,6,7,5) );

$eRDA1->setBlock( 'Crew 2', blockConcat($eData,0,13,5) );
$eRDA1->setBlock( 'Crew Names 6', blockConcat($eData,1,13,5) );
$eRDA1->setBlock( 'Crew Names 7', blockConcat($eData,2,13,5) );
$eRDA1->setBlock( 'Crew Names 8', blockConcat($eData,3,13,5) );
$eRDA1->setBlock( 'Crew Names 9', blockConcat($eData,4,13,5) );
$eRDA1->setBlock( 'Crew Names 10', blockConcat($eData,5,13,5) );
$eRDA1->setBlock( 'Notes 2', blockConcat($eData,6,13,5) );

$eRDA1->setBlock( 'Crew 3', blockConcat($eData,0,19,5) );
$eRDA1->setBlock( 'Crew Names 11', blockConcat($eData,1,19,5) );
$eRDA1->setBlock( 'Crew Names 12', blockConcat($eData,2,19,5) );
$eRDA1->setBlock( 'Crew Names 13', blockConcat($eData,3,19,5) );
$eRDA1->setBlock( 'Crew Names 14', blockConcat($eData,4,19,5) );
$eRDA1->setBlock( 'Crew Names 15', blockConcat($eData,5,19,5) );
$eRDA1->setBlock( 'Notes 3', blockConcat($eData,6,19,5) );

$eRDA1->setBlock( 'Crew 4', blockConcat($eData,0,25,5) );
$eRDA1->setBlock( 'Crew Names 16', blockConcat($eData,1,25,5) );
$eRDA1->setBlock( 'Crew Names 17', blockConcat($eData,2,25,5) );
$eRDA1->setBlock( 'Crew Names 18', blockConcat($eData,3,25,5) );
$eRDA1->setBlock( 'Crew Names 19', blockConcat($eData,4,25,5) );
$eRDA1->setBlock( 'Crew Names 20', blockConcat($eData,5,25,5) );
$eRDA1->setBlock( 'Notes 4', blockConcat($eData,6,25,5) );

$eRDA1->setBlock( 'Crew 5', blockConcat($eData,0,31,5) );
$eRDA1->setBlock( 'Crew Names 21', blockConcat($eData,1,31,5));
$eRDA1->setBlock( 'Crew Names 22', blockConcat($eData,2,31,5) );
$eRDA1->setBlock( 'Crew Names 23', blockConcat($eData,3,31,5) );
$eRDA1->setBlock( 'Crew Names 24', blockConcat($eData,4,31,5) );
$eRDA1->setBlock( 'Crew Names 25', blockConcat($eData,5,31,5) );
$eRDA1->setBlock( 'Notes 5', blockConcat($eData,6,31,5) );











$eRDA2->setTemplateName('Schedule 2');

$eRDA2->Description = $eB2;

$eRDA2->setBlock('Crew Schedule Title', $eB2);

$eRDA2->setBlock( 'Month 1',$eData->getCell('B43')->getCalculatedValue() );
$eRDA2->setBlock( 'Month 2',$eData->getCell('C43')->getCalculatedValue() );
$eRDA2->setBlock( 'Month 3',$eData->getCell('D43')->getCalculatedValue() );
$eRDA2->setBlock( 'Month 4',$eData->getCell('E43')->getCalculatedValue() );
$eRDA2->setBlock( 'Month 5',$eData->getCell('F43')->getCalculatedValue() );

$eRDA2->setBlock( 'Date 1',dateConvert($eData->getCell('B44')->getCalculatedValue()) );
$eRDA2->setBlock( 'Date 2',dateConvert($eData->getCell('C44')->getCalculatedValue()) );
$eRDA2->setBlock( 'Date 3',dateConvert($eData->getCell('D44')->getCalculatedValue()) );
$eRDA2->setBlock( 'Date 4',dateConvert($eData->getCell('E44')->getCalculatedValue()) );
$eRDA2->setBlock( 'Date 5',dateConvert($eData->getCell('F44')->getCalculatedValue()) );

$eRDA2->setBlock( 'Crew 1', blockConcat($eData,0,45,1) );
$eRDA2->setBlock( 'Crew Names 1', blockConcat($eData,1,45,1) );
$eRDA2->setBlock( 'Crew Names 2', blockConcat($eData,2,45,1) );
$eRDA2->setBlock( 'Crew Names 3', blockConcat($eData,3,45,1) );
$eRDA2->setBlock( 'Crew Names 4', blockConcat($eData,4,45,1) );
$eRDA2->setBlock( 'Crew Names 5', blockConcat($eData,5,45,1) );
$eRDA2->setBlock( 'Notes 1', blockConcat($eData,6,45,1) );

$eRDA2->setBlock( 'Crew 2', blockConcat($eData,0,46,1) );
$eRDA2->setBlock( 'Crew Names 6', blockConcat($eData,1,46,1) );
$eRDA2->setBlock( 'Crew Names 7', blockConcat($eData,2,46,1) );
$eRDA2->setBlock( 'Crew Names 8', blockConcat($eData,3,46,1) );
$eRDA2->setBlock( 'Crew Names 9', blockConcat($eData,4,46,1) );
$eRDA2->setBlock( 'Crew Names 10', blockConcat($eData,5,46,1) );
$eRDA2->setBlock( 'Notes 2', blockConcat($eData,6,46,1) );

$eRDA2->setBlock( 'Crew 3', blockConcat($eData,0,47,1) );
$eRDA2->setBlock( 'Crew Names 11', blockConcat($eData,1,47,1) );
$eRDA2->setBlock( 'Crew Names 12', blockConcat($eData,2,47,1) );
$eRDA2->setBlock( 'Crew Names 13', blockConcat($eData,3,47,1) );
$eRDA2->setBlock( 'Crew Names 14', blockConcat($eData,4,47,1) );
$eRDA2->setBlock( 'Crew Names 15', blockConcat($eData,5,47,1) );
$eRDA2->setBlock( 'Notes 3', blockConcat($eData,6,47,1) );

$eRDA2->setBlock( 'Crew 4', blockConcat($eData,0,48,4) );
$eRDA2->setBlock( 'Crew Names 16', blockConcat($eData,1,48,4) );
$eRDA2->setBlock( 'Crew Names 17', blockConcat($eData,2,48,4) );
$eRDA2->setBlock( 'Crew Names 18', blockConcat($eData,3,48,4) );
$eRDA2->setBlock( 'Crew Names 19', blockConcat($eData,4,48,4) );
$eRDA2->setBlock( 'Crew Names 20', blockConcat($eData,5,48,4) );
$eRDA2->setBlock( 'Notes 4', blockConcat($eData,6,48,4) );


foreach ($images as $key => $image) {
	if ($key <=3){
		// echo "setting: Web Picture " . ($key+1) . " as: " . $imagesServerPath.$image;
		$eRDA2->setBlock( "Web Picture ".($key+1), $imagesServerPath.$image );
	}
}









$eRDA3->setTemplateName('Schedule 3');

$eRDA3->Description = $eB3;

$eRDA3->setBlock('Crew Schedule Title', $eB3);

$eRDA3->setBlock( 'Month 1',$eData->getCell('B43')->getCalculatedValue() );
$eRDA3->setBlock( 'Month 2',$eData->getCell('C43')->getCalculatedValue() );
$eRDA3->setBlock( 'Month 3',$eData->getCell('D43')->getCalculatedValue() );
$eRDA3->setBlock( 'Month 4',$eData->getCell('E43')->getCalculatedValue() );
$eRDA3->setBlock( 'Month 5',$eData->getCell('F43')->getCalculatedValue() );

$eRDA3->setBlock( 'Date 1',dateConvert($eData->getCell('B44')->getCalculatedValue()) );
$eRDA3->setBlock( 'Date 2',dateConvert($eData->getCell('C44')->getCalculatedValue()) );
$eRDA3->setBlock( 'Date 3',dateConvert($eData->getCell('D44')->getCalculatedValue()) );
$eRDA3->setBlock( 'Date 4',dateConvert($eData->getCell('E44')->getCalculatedValue()) );
$eRDA3->setBlock( 'Date 5',dateConvert($eData->getCell('F44')->getCalculatedValue()) );

$eRDA3->setBlock( 'Line One',$eData->getCell('A73')->getCalculatedValue() );
$eRDA3->setBlock( 'Crew Names 1',$eData->getCell('B73')->getCalculatedValue() );
$eRDA3->setBlock( 'Crew Names 2',$eData->getCell('C73')->getCalculatedValue() );
$eRDA3->setBlock( 'Crew Names 3',$eData->getCell('D73')->getCalculatedValue() );
$eRDA3->setBlock( 'Crew Names 4',$eData->getCell('E73')->getCalculatedValue() );
$eRDA3->setBlock( 'Crew Names 5',$eData->getCell('F73')->getCalculatedValue() );
$eRDA3->setBlock( 'Notes 1',$eData->getCell('G73')->getCalculatedValue() );

$eRDA3->setBlock( 'Line Two',$eData->getCell('A74')->getCalculatedValue() );
$eRDA3->setBlock( 'Crew Names 6',$eData->getCell('B74')->getCalculatedValue() );
$eRDA3->setBlock( 'Crew Names 7',$eData->getCell('C74')->getCalculatedValue() );
$eRDA3->setBlock( 'Crew Names 8',$eData->getCell('D74')->getCalculatedValue() );
$eRDA3->setBlock( 'Crew Names 9',$eData->getCell('E74')->getCalculatedValue() );
$eRDA3->setBlock( 'Crew Names 10',$eData->getCell('F74')->getCalculatedValue() );
$eRDA3->setBlock( 'Notes 2',$eData->getCell('G74')->getCalculatedValue() );

$eRDA3->setBlock( 'Line 3',$eData->getCell('A75')->getCalculatedValue() );
$eRDA3->setBlock( 'Crew Names 11',$eData->getCell('B75')->getCalculatedValue() );
$eRDA3->setBlock( 'Crew Names 12',$eData->getCell('C75')->getCalculatedValue() );
$eRDA3->setBlock( 'Crew Names 13',$eData->getCell('D75')->getCalculatedValue() );
$eRDA3->setBlock( 'Crew Names 14',$eData->getCell('E75')->getCalculatedValue() );
$eRDA3->setBlock( 'Crew Names 15',$eData->getCell('F75')->getCalculatedValue() );
$eRDA3->setBlock( 'Notes 3',$eData->getCell('G75')->getCalculatedValue() );

$eRDA3->setBlock( 'Line Four',$eData->getCell('A76')->getCalculatedValue() );
$eRDA3->setBlock( 'New Block 2',$eData->getCell('B76')->getCalculatedValue() );
$eRDA3->setBlock( 'New Block 3',$eData->getCell('C76')->getCalculatedValue() );
$eRDA3->setBlock( 'New Block 4',$eData->getCell('D76')->getCalculatedValue() );
$eRDA3->setBlock( 'New Block 5',$eData->getCell('E76')->getCalculatedValue() );
$eRDA3->setBlock( 'New Block 6',$eData->getCell('F76')->getCalculatedValue() );
$eRDA3->setBlock( 'New Block 7',$eData->getCell('G76')->getCalculatedValue() );

$eRDA3->setBlock( 'Line Five',$eData->getCell('A77')->getCalculatedValue() );
$eRDA3->setBlock( 'New Block 8',$eData->getCell('B77')->getCalculatedValue() );
$eRDA3->setBlock( 'New Block 9',$eData->getCell('C77')->getCalculatedValue() );
$eRDA3->setBlock( 'New Block 10',$eData->getCell('D77')->getCalculatedValue() );
$eRDA3->setBlock( 'New Block 11',$eData->getCell('E77')->getCalculatedValue() );
$eRDA3->setBlock( 'New Block 12',$eData->getCell('F77')->getCalculatedValue() );
$eRDA3->setBlock( 'New Block 13',$eData->getCell('G77')->getCalculatedValue() );

$eRDA3->setBlock( 'Line Six',$eData->getCell('A78')->getCalculatedValue() );
$eRDA3->setBlock( 'New Block 14',$eData->getCell('B78')->getCalculatedValue() );
$eRDA3->setBlock( 'New Block 15',$eData->getCell('C78')->getCalculatedValue() );
$eRDA3->setBlock( 'New Block 16',$eData->getCell('D78')->getCalculatedValue() );
$eRDA3->setBlock( 'New Block 17',$eData->getCell('E78')->getCalculatedValue() );
$eRDA3->setBlock( 'New Block 18',$eData->getCell('F78')->getCalculatedValue() );
$eRDA3->setBlock( 'New Block 19',$eData->getCell('G78')->getCalculatedValue() );

$eRDA3->setBlock( 'Line Seven',$eData->getCell('A79')->getCalculatedValue() );
$eRDA3->setBlock( 'New Block 20',$eData->getCell('B79')->getCalculatedValue() );
$eRDA3->setBlock( 'New Block 21',$eData->getCell('C79')->getCalculatedValue() );
$eRDA3->setBlock( 'New Block 22',$eData->getCell('D79')->getCalculatedValue() );
$eRDA3->setBlock( 'New Block 23',$eData->getCell('E79')->getCalculatedValue() );
$eRDA3->setBlock( 'New Block 24',$eData->getCell('F79')->getCalculatedValue() );
$eRDA3->setBlock( 'New Block 25',$eData->getCell('G79')->getCalculatedValue() );

$eRDA3->setBlock( 'Line Eight',$eData->getCell('A80')->getCalculatedValue() );
$eRDA3->setBlock( 'New Block',$eData->getCell('B80')->getCalculatedValue() );
$eRDA3->setBlock( 'New Block 26',$eData->getCell('C80')->getCalculatedValue() );
$eRDA3->setBlock( 'New Block 28',$eData->getCell('D80')->getCalculatedValue() );
$eRDA3->setBlock( 'New Block 29',$eData->getCell('E80')->getCalculatedValue() );
$eRDA3->setBlock( 'New Block 30',$eData->getCell('F80')->getCalculatedValue() );
$eRDA3->setBlock( 'New Block 31',$eData->getCell('G80')->getCalculatedValue() );

$eRDA3->setBlock( 'Line Nine',$eData->getCell('A81')->getCalculatedValue() );
$eRDA3->setBlock( 'New Block 33',$eData->getCell('B81')->getCalculatedValue() );
$eRDA3->setBlock( 'New Block 34',$eData->getCell('C81')->getCalculatedValue() );
$eRDA3->setBlock( 'New Block 35',$eData->getCell('D81')->getCalculatedValue() );
$eRDA3->setBlock( 'New Block 36',$eData->getCell('E81')->getCalculatedValue() );
$eRDA3->setBlock( 'New Block 37',$eData->getCell('F81')->getCalculatedValue() );
$eRDA3->setBlock( 'New Block 38',$eData->getCell('G81')->getCalculatedValue() );

$eRDA3->setBlock( 'Line Ten',$eData->getCell('A82')->getCalculatedValue() );
$eRDA3->setBlock( 'New Block 39',$eData->getCell('B82')->getCalculatedValue() );
$eRDA3->setBlock( 'New Block 40',$eData->getCell('C82')->getCalculatedValue() );
$eRDA3->setBlock( 'New Block 41',$eData->getCell('D82')->getCalculatedValue() );
$eRDA3->setBlock( 'New Block 42',$eData->getCell('E82')->getCalculatedValue() );
$eRDA3->setBlock( 'New Block 43',$eData->getCell('F82')->getCalculatedValue() );
$eRDA3->setBlock( 'New Block 44',$eData->getCell('G82')->getCalculatedValue() );

$eRDA3->setBlock( 'Line Eleven',$eData->getCell('A83')->getCalculatedValue() );
$eRDA3->setBlock( 'New Block 45',$eData->getCell('B83')->getCalculatedValue() );
$eRDA3->setBlock( 'New Block 46',$eData->getCell('C83')->getCalculatedValue() );
$eRDA3->setBlock( 'New Block 47',$eData->getCell('D83')->getCalculatedValue() );
$eRDA3->setBlock( 'New Block 48',$eData->getCell('E83')->getCalculatedValue() );
$eRDA3->setBlock( 'New Block 49',$eData->getCell('F83')->getCalculatedValue() );
$eRDA3->setBlock( 'New Block 50',$eData->getCell('G83')->getCalculatedValue() );

$eRDA3->setBlock( 'Crew 4', blockConcat($eData,0,84,12) );
$eRDA3->setBlock( 'Crew Names 16', blockConcatInline($eData,1,84,12) );
$eRDA3->setBlock( 'Crew Names 17', blockConcatInline($eData,2,84,12) );
$eRDA3->setBlock( 'Crew Names 18', blockConcatInline($eData,3,84,12) );
$eRDA3->setBlock( 'Crew Names 19', blockConcatInline($eData,4,84,12) );
$eRDA3->setBlock( 'Crew Names 20', blockConcatInline($eData,5,84,12) );
$eRDA3->setBlock( 'Notes 4', blockConcat($eData,6,84,12) );

$eRDA3->setBlock( 'Message',$eData->getCell('A97')->getCalculatedValue() );








foreach ($eBulletins1 as $ID) {
	$eRDA1->DeletePage($ID);
}

foreach ($eBulletins2 as $ID) {
	$eRDA2->DeletePage($ID);
}
foreach ($eBulletins3 as $ID) {
	$eRDA3->DeletePage($ID);
}

if($eRDA1->getLastError() != '')echo $eRDA1->getLastError().'<br>';
if($eRDA2->getLastError() != '')echo $eRDA2->getLastError()."<br>";
if($eRDA3->getLastError() != '')echo $eRDA3->getLastError()."<br>";




$eRDA1->CreatePage();
$eRDA2->CreatePage();
$eRDA3->CreatePage();

if($eRDA1->getLastError() != '')echo $eRDA1->getLastError().'<br>';
if($eRDA2->getLastError() != '')echo $eRDA2->getLastError()."<br>";
if($eRDA3->getLastError() != '')echo $eRDA3->getLastError()."<br>";





$time_end = microtime(true);
$time = $time_end - $time_start;

echo "Finished in " . intval($time / 60). " minutes and ". ($time % 60) . " seconds.";



