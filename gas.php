<?php

include_once 'Classes/RDA.php';
include_once 'config.php';
include_once 'Classes/PHPExcel.php';
include_once 'functions/functions.php';

set_time_limit(0);
$time_start = microtime(true);


try{
$gType = PHPExcel_IOFactory::identify($gFile);
$gReader = PHPExcel_IOFactory::createReader($gType);
$gReader->setLoadSheetsOnly($gSheetName);
$gData = $gReader->load($gFile)->getActiveSheet();
}
catch(PHPExcel_Reader_Exception $e){
	var_dump($e->getMessage());
	$log = fopen('errorLog.txt','a+');
	$message = date('Y-m-d H:i:s') . ' Error loading File: '.$e->getMessage()."\r\n";
	fwrite($log, $message);
	fclose($log);
	die();
}

$gRDA = new RDA($user, $password, $server);
$gRDA1 = new RDA($user, $password, $server);
$gRDA2 = new RDA($user, $password, $server);
$gRDA3 = new RDA($user, $password, $server);







$gZones = $gRDA->GetZoneList();
foreach($gZones as $zone){
	if($zone['ZoneName'] == 'Avista-Gas-Chart') $gZoneID = $zone['ZoneID'];
}



$gRDA->setZoneIDs($gZoneID);
$gRDA1->setZoneIDs($gZoneID);
$gRDA2->setZoneIDs($gZoneID);
$gRDA3->setZoneIDs($gZoneID);



$gB1 = 'Gas Crew Schedule';
$gB2 = 'Inspectors/Information';
$gB3 = 'CDA GAS Servicemen';


$gBulletins1 = array();
$gBulletins2 = array();
$gBulletins3 = array();


$gBulletins = $gRDA->GetBulletinList();


if (isAssoc($gBulletins[$gZoneID])){
	$gBulletins[$gZoneID] = array($gBulletins[$gZoneID]);
}

foreach ($gBulletins[$gZoneID] as $bulletin) {
	if($bulletin['Description'] == $gB1) array_push( $gBulletins1, $bulletin['GUID'] );
	if($bulletin['Description'] == $gB2) array_push( $gBulletins2, $bulletin['GUID'] );
	if($bulletin['Description'] == $gB3) array_push( $gBulletins3, $bulletin['GUID'] );
}



$i = 0;
$images = array();
foreach ($gData->getDrawingCollection() as $drawing) {
	$imagePath = imageConvert($drawing,'Gas');
	array_push($images, $imagePath);
}






$gRDA1->setTemplateName('Schedule');

$gRDA1->Description = $gB1;

$gRDA1->setBlock('Crew Schedule Title', $gB1);

$gRDA1->setBlock( 'Month 1',$gData->getCell('D5')->getCalculatedValue() );
$gRDA1->setBlock( 'Month 2',$gData->getCell('E5')->getCalculatedValue() );
$gRDA1->setBlock( 'Month 3',$gData->getCell('F5')->getCalculatedValue() );
$gRDA1->setBlock( 'Month 4',$gData->getCell('G5')->getCalculatedValue() );
$gRDA1->setBlock( 'Month 5',$gData->getCell('H5')->getCalculatedValue() );

$gRDA1->setBlock( 'Date 1',dateConvert($gData->getCell('D6')->getCalculatedValue()) );
$gRDA1->setBlock( 'Date 2',dateConvert($gData->getCell('E6')->getCalculatedValue()) );
$gRDA1->setBlock( 'Date 3',dateConvert($gData->getCell('F6')->getCalculatedValue()) );
$gRDA1->setBlock( 'Date 4',dateConvert($gData->getCell('G6')->getCalculatedValue()) );
$gRDA1->setBlock( 'Date 5',dateConvert($gData->getCell('H6')->getCalculatedValue()) );

$gRDA1->setBlock( 'Crew 1', blockConcat($gData,2,7,6) );
$gRDA1->setBlock( 'Crew Names 1', blockConcat($gData,3,7,6) );
$gRDA1->setBlock( 'Crew Names 2', blockConcat($gData,4,7,6) );
$gRDA1->setBlock( 'Crew Names 3', blockConcat($gData,5,7,6) );
$gRDA1->setBlock( 'Crew Names 4', blockConcat($gData,6,7,6) );
$gRDA1->setBlock( 'Crew Names 5', blockConcat($gData,7,7,6) );
$gRDA1->setBlock( 'Notes 1', blockConcat($gData,8,7,6) );

$gRDA1->setBlock( 'Crew 2', blockConcat($gData,2,13,6) );
$gRDA1->setBlock( 'Crew Names 6', blockConcat($gData,3,13,6) );
$gRDA1->setBlock( 'Crew Names 7', blockConcat($gData,4,13,6) );
$gRDA1->setBlock( 'Crew Names 8', blockConcat($gData,5,13,6) );
$gRDA1->setBlock( 'Crew Names 9', blockConcat($gData,6,13,6) );
$gRDA1->setBlock( 'Crew Names 10', blockConcat($gData,7,13,6) );
$gRDA1->setBlock( 'Notes 2', blockConcat($gData,8,13,6) );

$gRDA1->setBlock( 'Crew 3', blockConcat($gData,2,19,6) );
$gRDA1->setBlock( 'Crew Names 11', blockConcat($gData,3,19,6) );
$gRDA1->setBlock( 'Crew Names 12', blockConcat($gData,4,19,6) );
$gRDA1->setBlock( 'Crew Names 13', blockConcat($gData,5,19,6) );
$gRDA1->setBlock( 'Crew Names 14', blockConcat($gData,6,19,6) );
$gRDA1->setBlock( 'Crew Names 15', blockConcat($gData,7,19,6) );
$gRDA1->setBlock( 'Notes 3', blockConcat($gData,8,19,6) );

$gRDA1->setBlock( 'Crew 4', blockConcat($gData,2,25,6) );
$gRDA1->setBlock( 'Crew Names 16', blockConcat($gData,3,25,6) );
$gRDA1->setBlock( 'Crew Names 17', blockConcat($gData,4,25,6) );
$gRDA1->setBlock( 'Crew Names 18', blockConcat($gData,5,25,6) );
$gRDA1->setBlock( 'Crew Names 19', blockConcat($gData,6,25,6) );
$gRDA1->setBlock( 'Crew Names 20', blockConcat($gData,7,25,6) );
$gRDA1->setBlock( 'Notes 4', blockConcat($gData,8,25,6) );

$gRDA1->setBlock( 'Crew 5', blockConcatNWS($gData,2,31,2).blockConcatNWS($gData,2,33,2) );
$gRDA1->setBlock( 'Crew Names 21', blockConcatNWS($gData,3,31,2).blockConcatNWS($gData,3,33,2));
$gRDA1->setBlock( 'Crew Names 22', blockConcatNWS($gData,4,31,2).blockConcatNWS($gData,4,33,2) );
$gRDA1->setBlock( 'Crew Names 23', blockConcatNWS($gData,5,31,2).blockConcatNWS($gData,5,33,2) );
$gRDA1->setBlock( 'Crew Names 24', blockConcatNWS($gData,6,31,2).blockConcatNWS($gData,6,33,2) );
$gRDA1->setBlock( 'Crew Names 25', blockConcatNWS($gData,7,31,2).blockConcatNWS($gData,7,33,2) );
$gRDA1->setBlock( 'Notes 5', blockConcatNWS($gData,8,31,2).blockConcatNWS($gData,8,33,2) );







$gRDA2->setTemplateName('Schedule 2');

$gRDA2->Description = $gB2;

$gRDA2->setBlock('Crew Schedule Title', $gB2);

$gRDA2->setBlock( 'Month 1',$gData->getCell('D42')->getCalculatedValue() );
$gRDA2->setBlock( 'Month 2',$gData->getCell('E42')->getCalculatedValue() );
$gRDA2->setBlock( 'Month 3',$gData->getCell('F42')->getCalculatedValue() );
$gRDA2->setBlock( 'Month 4',$gData->getCell('G42')->getCalculatedValue() );
$gRDA2->setBlock( 'Month 5',$gData->getCell('H42')->getCalculatedValue() );


$gRDA2->setBlock( 'Date 1',dateConvert($gData->getCell('D43')->getCalculatedValue()) );
$gRDA2->setBlock( 'Date 2',dateConvert($gData->getCell('E43')->getCalculatedValue()) );
$gRDA2->setBlock( 'Date 3',dateConvert($gData->getCell('F43')->getCalculatedValue()) );
$gRDA2->setBlock( 'Date 4',dateConvert($gData->getCell('G43')->getCalculatedValue()) );
$gRDA2->setBlock( 'Date 5',dateConvert($gData->getCell('H43')->getCalculatedValue()) );


$gRDA2->setBlock( 'Line 1',$gData->getCell('C44')->getCalculatedValue() );
$gRDA2->setBlock( 'Crew Names 1',$gData->getCell('D44')->getCalculatedValue() );
$gRDA2->setBlock( 'Crew Names 2',$gData->getCell('E44')->getCalculatedValue() );
$gRDA2->setBlock( 'Crew Names 3',$gData->getCell('F44')->getCalculatedValue() );
$gRDA2->setBlock( 'Crew Names 4',$gData->getCell('G44')->getCalculatedValue() );
$gRDA2->setBlock( 'Crew Names 5',$gData->getCell('H44')->getCalculatedValue() );
$gRDA2->setBlock( 'Notes 1',$gData->getCell('I44')->getCalculatedValue() );


$gRDA2->setBlock( 'Line 2',$gData->getCell('C45')->getCalculatedValue() );
$gRDA2->setBlock( 'Crew Names 6',$gData->getCell('D45')->getCalculatedValue() );
$gRDA2->setBlock( 'Crew Names 7',$gData->getCell('E45')->getCalculatedValue() );
$gRDA2->setBlock( 'Crew Names 8',$gData->getCell('F45')->getCalculatedValue() );
$gRDA2->setBlock( 'Crew Names 9',$gData->getCell('G45')->getCalculatedValue() );
$gRDA2->setBlock( 'Crew Names 10',$gData->getCell('H45')->getCalculatedValue() );
$gRDA2->setBlock( 'Notes 2',$gData->getCell('I45')->getCalculatedValue() );


$gRDA2->setBlock( 'Line 3',$gData->getCell('C46')->getCalculatedValue() );
$gRDA2->setBlock( 'Crew Names 11',$gData->getCell('D46')->getCalculatedValue() );
$gRDA2->setBlock( 'Crew Names 12',$gData->getCell('E46')->getCalculatedValue() );
$gRDA2->setBlock( 'Crew Names 13',$gData->getCell('F46')->getCalculatedValue() );
$gRDA2->setBlock( 'Crew Names 14',$gData->getCell('G46')->getCalculatedValue() );
$gRDA2->setBlock( 'Crew Names 15',$gData->getCell('H46')->getCalculatedValue() );
$gRDA2->setBlock( 'Notes 3',$gData->getCell('I46')->getCalculatedValue() );


$gRDA2->setBlock( 'Line Four',$gData->getCell('C47')->getCalculatedValue() );
$gRDA2->setBlock( 'New Block 2',$gData->getCell('D47')->getCalculatedValue() );
$gRDA2->setBlock( 'New Block 3',$gData->getCell('E47')->getCalculatedValue() );
$gRDA2->setBlock( 'New Block 4',$gData->getCell('F47')->getCalculatedValue() );
$gRDA2->setBlock( 'New Block 5',$gData->getCell('G47')->getCalculatedValue() );
$gRDA2->setBlock( 'New Block 6',$gData->getCell('H47')->getCalculatedValue() );
$gRDA2->setBlock( 'New Block 7',$gData->getCell('I47')->getCalculatedValue() );


$gRDA2->setBlock( 'Line Five',$gData->getCell('C48')->getCalculatedValue() );
$gRDA2->setBlock( 'New Block 8',$gData->getCell('D48')->getCalculatedValue() );
$gRDA2->setBlock( 'New Block 9',$gData->getCell('E48')->getCalculatedValue() );
$gRDA2->setBlock( 'New Block 10',$gData->getCell('F48')->getCalculatedValue() );
$gRDA2->setBlock( 'New Block 11',$gData->getCell('G48')->getCalculatedValue() );
$gRDA2->setBlock( 'New Block 12',$gData->getCell('H48')->getCalculatedValue() );
$gRDA2->setBlock( 'New Block 13',$gData->getCell('I48')->getCalculatedValue() );


$gRDA2->setBlock( 'Line Six',$gData->getCell('C49')->getCalculatedValue() );
$gRDA2->setBlock( 'New Block 14',$gData->getCell('D49')->getCalculatedValue() );
$gRDA2->setBlock( 'New Block 15',$gData->getCell('E49')->getCalculatedValue() );
$gRDA2->setBlock( 'New Block 16',$gData->getCell('F49')->getCalculatedValue() );
$gRDA2->setBlock( 'New Block 17',$gData->getCell('G49')->getCalculatedValue() );
$gRDA2->setBlock( 'New Block 18',$gData->getCell('H49')->getCalculatedValue() );
$gRDA2->setBlock( 'New Block 19',$gData->getCell('I49')->getCalculatedValue() );


$gRDA2->setBlock( 'Line Seven', blockConcat($gData,2,50,7) );
$gRDA2->setBlock( 'Crew Names 16', blockConcat($gData,3,50,7) );
$gRDA2->setBlock( 'Crew Names 17', blockConcat($gData,4,50,7) );
$gRDA2->setBlock( 'Crew Names 18', blockConcat($gData,5,50,7) );
$gRDA2->setBlock( 'Crew Names 19', blockConcat($gData,6,50,7) );
$gRDA2->setBlock( 'Crew Names 20', blockConcat($gData,7,50,7) );
$gRDA2->setBlock( 'Notes 7', blockConcat($gData,8,50,7) );


$gRDA2->setBlock( 'Line Eight', blockConcat($gData,2,57,9) );
$gRDA2->setBlock( 'Crew Names 21', blockConcat($gData,3,57,9) );
$gRDA2->setBlock( 'Crew Names 22', blockConcat($gData,4,57,9) );
$gRDA2->setBlock( 'Crew Names 23', blockConcat($gData,5,57,9) );
$gRDA2->setBlock( 'Crew Names 24', blockConcat($gData,6,57,9) );
$gRDA2->setBlock( 'Crew Names 25', blockConcat($gData,7,57,9) );
$gRDA2->setBlock( 'Notes 8', blockConcat($gData,8,57,9) );














$gRDA3->setTemplateName('Schedule 3');

$gRDA3->Description = $gB3;

$gRDA3->setBlock('Crew Schedule Title', $gB3);

$gRDA3->setBlock( 'Month 1',$gData->getCell('D79')->getCalculatedValue() );
$gRDA3->setBlock( 'Month 2',$gData->getCell('E79')->getCalculatedValue() );
$gRDA3->setBlock( 'Month 3',$gData->getCell('F79')->getCalculatedValue() );
$gRDA3->setBlock( 'Month 4',$gData->getCell('G79')->getCalculatedValue() );
$gRDA3->setBlock( 'Month 5',$gData->getCell('H79')->getCalculatedValue() );

$gRDA3->setBlock( 'Date 1',dateConvert($gData->getCell('D80')->getCalculatedValue()) );
$gRDA3->setBlock( 'Date 2',dateConvert($gData->getCell('E80')->getCalculatedValue()) );
$gRDA3->setBlock( 'Date 3',dateConvert($gData->getCell('F80')->getCalculatedValue()) );
$gRDA3->setBlock( 'Date 4',dateConvert($gData->getCell('G80')->getCalculatedValue()) );
$gRDA3->setBlock( 'Date 5',dateConvert($gData->getCell('H80')->getCalculatedValue()) );

$gRDA3->setBlock( 'Crew 1', blockConcatInlineNC($gData,2,81,2) );
$gRDA3->setBlock( 'Crew Names 1', $gData->getCell('D81')->getCalculatedValue() );
$gRDA3->setBlock( 'Crew Names 2', $gData->getCell('E81')->getCalculatedValue() );
$gRDA3->setBlock( 'Crew Names 3', $gData->getCell('F81')->getCalculatedValue() );
$gRDA3->setBlock( 'Crew Names 4', $gData->getCell('G81')->getCalculatedValue() );
$gRDA3->setBlock( 'Crew Names 5', $gData->getCell('H81')->getCalculatedValue() );
$gRDA3->setBlock( 'Notes 1', $gData->getCell('I81')->getCalculatedValue() );

$gRDA3->setBlock( 'Crew 2', blockConcatInlineNC($gData,2,83,2) );
$gRDA3->setBlock( 'Crew Names 6', $gData->getCell('D83')->getCalculatedValue() );
$gRDA3->setBlock( 'Crew Names 7', $gData->getCell('E83')->getCalculatedValue() );
$gRDA3->setBlock( 'Crew Names 8', $gData->getCell('F83')->getCalculatedValue() );
$gRDA3->setBlock( 'Crew Names 9', $gData->getCell('G83')->getCalculatedValue() );
$gRDA3->setBlock( 'Crew Names 10', $gData->getCell('H83')->getCalculatedValue() );
$gRDA3->setBlock( 'Notes 2', $gData->getCell('I83')->getCalculatedValue() );

$gRDA3->setBlock( 'Crew 3', blockConcatInlineNC($gData,2,85,2) );
$gRDA3->setBlock( 'Crew Names 11', $gData->getCell('D85')->getCalculatedValue() );
$gRDA3->setBlock( 'Crew Names 12', $gData->getCell('E85')->getCalculatedValue() );
$gRDA3->setBlock( 'Crew Names 13', $gData->getCell('F85')->getCalculatedValue() );
$gRDA3->setBlock( 'Crew Names 14', $gData->getCell('G85')->getCalculatedValue() );
$gRDA3->setBlock( 'Crew Names 15', $gData->getCell('H85')->getCalculatedValue() );
$gRDA3->setBlock( 'Notes 3', $gData->getCell('I85')->getCalculatedValue() );

$gRDA3->setBlock( 'Crew 4', blockConcatInlineNC($gData,2,87,2) );
$gRDA3->setBlock( 'Crew Names 16', $gData->getCell('D87')->getCalculatedValue() );
$gRDA3->setBlock( 'Crew Names 17', $gData->getCell('E87')->getCalculatedValue() );
$gRDA3->setBlock( 'Crew Names 18', $gData->getCell('F87')->getCalculatedValue() );
$gRDA3->setBlock( 'Crew Names 19', $gData->getCell('G87')->getCalculatedValue() );
$gRDA3->setBlock( 'Crew Names 20', $gData->getCell('H87')->getCalculatedValue() );
$gRDA3->setBlock( 'Notes 4', $gData->getCell('I87')->getCalculatedValue() );

$gRDA3->setBlock( 'Crew 5', blockConcatInlineNC($gData,2,89,2) );
$gRDA3->setBlock( 'Crew Names 21', $gData->getCell('D89')->getCalculatedValue() );
$gRDA3->setBlock( 'Crew Names 22', $gData->getCell('E89')->getCalculatedValue() );
$gRDA3->setBlock( 'Crew Names 23', $gData->getCell('F89')->getCalculatedValue() );
$gRDA3->setBlock( 'Crew Names 24', $gData->getCell('G89')->getCalculatedValue() );
$gRDA3->setBlock( 'Crew Names 25', $gData->getCell('H89')->getCalculatedValue() );
$gRDA3->setBlock( 'Notes 5', $gData->getCell('I89')->getCalculatedValue() );


$gRDA3->setBlock( 'Message',$gData->getCell('C91')->getCalculatedValue() );


foreach ($images as $key => $image) {
	if ($key <=3){
		// echo "setting: Web Picture " . ($key+1) . " as: " . $imagesServerPath.$image;
		$gRDA3->setBlock( "Web Picture ".($key+1), $imagesServerPath.$image );
	}
}




foreach ($gBulletins1 as $ID) {
	$gRDA1->DeletePage($ID);
}

foreach ($gBulletins2 as $ID) {
	$gRDA2->DeletePage($ID);
}
foreach ($gBulletins3 as $ID) {
	$gRDA3->DeletePage($ID);
}


if($gRDA1->getLastError() != '')echo $gRDA1->getLastError()."<br>";
if($gRDA2->getLastError() != '')echo $gRDA2->getLastError()."<br";
if($gRDA3->getLastError() != '')echo $gRDA3->getLastError()."<br>";



$gRDA1->CreatePage();
$gRDA2->CreatePage();
$gRDA3->CreatePage();

if($gRDA1->getLastError() != '')echo $gRDA1->getLastError()."<br>";
if($gRDA2->getLastError() != '')echo $gRDA2->getLastError()."<br";
if($gRDA3->getLastError() != '')echo $gRDA3->getLastError()."<br>";





$time_end = microtime(true);
$time = $time_end - $time_start;

echo "Finished in " . intval($time / 60). " minutes and ". ($time % 60) . " seconds.";