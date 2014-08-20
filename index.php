<?php

include_once 'RDA.php';
include_once 'config.php';
include_once 'Classes/PHPExcel.php';
include_once 'Classes/chunk_read_filter.php';

set_time_limit(0);




///////////////////////////////////////////////
////////////////////////////////////////////////////
//////////////////////////////////////////////////
///////////////////////////////////////////////// Electric XLS Document


try{
$eType = PHPExcel_IOFactory::identify($eFile);
$eReader = PHPExcel_IOFactory::createReader($eType);
$eReader->setLoadSheetsOnly($eSheetName);
$eData = $eReader->load($eFile)->getActiveSheet();
}
catch(PHPExcel_Reader_Exception $e){
	var_dump($e->getMessage());
	$log = fopen('errorLog.txt','a+');
	$message = date('Y-m-d H:i:s') . ' Error loading File: '.$e->getMessage()."\r\n";
	fwrite($log, $message);
	fclose($log);
	die();
}

$eRDA1 = new RDA($user, $password, $server);







$eZones = $eRDA1->GetZoneList();
foreach($eZones as $zone){
	if($zone['ZoneName'] == 'Avista-Electric-Chart') $eZoneID = $zone['ZoneID'];
	if($zone['ZoneName'] == 'Avista-Gas-Chart') $gZoneID = $zone['ZoneID'];
}



$eRDA1->setZoneIDs($eZoneID);



$eB1 = 'Electric Crew Schedule';
$eB2 = 'Electric Servicemen';
$eB3 = 'Electric General Information';


$eBulletins = $eRDA1->GetBulletinList();

foreach ($eBulletins[$eZoneID] as $bulletin) {
	if($bulletin['Description'] == $eB1) $eB1ID = $bulletin['GUID'];
	if($bulletin['Description'] == $eB2) $eB2ID = $bulletin['GUID'];
	if($bulletin['Description'] == $eB3) $eB3ID = $bulletin['GUID'];
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

echo "Deleted Bulletin: " . $eRDA1->DeletePage($eB1ID) . "<br>";

echo "Created Bulletin: " . $eRDA1->CreatePage() . "<br>";

if($eRDA1->getLastError() != '')echo $eRDA1->getLastError().'<br>';

$eRDA1->clearBlocks();







$eRDA1->Description = $eB2;

$eRDA1->setBlock('Crew Schedule Title', $eB2);

$eRDA1->setBlock( 'Month 1',$eData->getCell('B43')->getCalculatedValue() );
$eRDA1->setBlock( 'Month 2',$eData->getCell('C43')->getCalculatedValue() );
$eRDA1->setBlock( 'Month 3',$eData->getCell('D43')->getCalculatedValue() );
$eRDA1->setBlock( 'Month 4',$eData->getCell('E43')->getCalculatedValue() );
$eRDA1->setBlock( 'Month 5',$eData->getCell('F43')->getCalculatedValue() );

$eRDA1->setBlock( 'Date 1',dateConvert($eData->getCell('B44')->getCalculatedValue()) );
$eRDA1->setBlock( 'Date 2',dateConvert($eData->getCell('C44')->getCalculatedValue()) );
$eRDA1->setBlock( 'Date 3',dateConvert($eData->getCell('D44')->getCalculatedValue()) );
$eRDA1->setBlock( 'Date 4',dateConvert($eData->getCell('E44')->getCalculatedValue()) );
$eRDA1->setBlock( 'Date 5',dateConvert($eData->getCell('F44')->getCalculatedValue()) );

$eRDA1->setBlock( 'Crew 1', blockConcat($eData,0,45,1) );
$eRDA1->setBlock( 'Crew Names 1', blockConcat($eData,1,45,1) );
$eRDA1->setBlock( 'Crew Names 2', blockConcat($eData,2,45,1) );
$eRDA1->setBlock( 'Crew Names 3', blockConcat($eData,3,45,1) );
$eRDA1->setBlock( 'Crew Names 4', blockConcat($eData,4,45,1) );
$eRDA1->setBlock( 'Crew Names 5', blockConcat($eData,5,45,1) );
$eRDA1->setBlock( 'Notes 1', blockConcat($eData,6,45,1) );

$eRDA1->setBlock( 'Crew 2', blockConcat($eData,0,46,1) );
$eRDA1->setBlock( 'Crew Names 6', blockConcat($eData,1,46,1) );
$eRDA1->setBlock( 'Crew Names 7', blockConcat($eData,2,46,1) );
$eRDA1->setBlock( 'Crew Names 8', blockConcat($eData,3,46,1) );
$eRDA1->setBlock( 'Crew Names 9', blockConcat($eData,4,46,1) );
$eRDA1->setBlock( 'Crew Names 10', blockConcat($eData,5,46,1) );
$eRDA1->setBlock( 'Notes 2', blockConcat($eData,6,46,1) );

$eRDA1->setBlock( 'Crew 3', blockConcat($eData,0,47,1) );
$eRDA1->setBlock( 'Crew Names 11', blockConcat($eData,1,47,1) );
$eRDA1->setBlock( 'Crew Names 12', blockConcat($eData,2,47,1) );
$eRDA1->setBlock( 'Crew Names 13', blockConcat($eData,3,47,1) );
$eRDA1->setBlock( 'Crew Names 14', blockConcat($eData,4,47,1) );
$eRDA1->setBlock( 'Crew Names 15', blockConcat($eData,5,47,1) );
$eRDA1->setBlock( 'Notes 3', blockConcat($eData,6,47,1) );

$eRDA1->setBlock( 'Crew 4', blockConcat($eData,0,48,4) );
$eRDA1->setBlock( 'Crew Names 16', blockConcat($eData,1,48,4) );
$eRDA1->setBlock( 'Crew Names 17', blockConcat($eData,2,48,4) );
$eRDA1->setBlock( 'Crew Names 18', blockConcat($eData,3,48,4) );
$eRDA1->setBlock( 'Crew Names 19', blockConcat($eData,4,48,4) );
$eRDA1->setBlock( 'Crew Names 20', blockConcat($eData,5,48,4) );
$eRDA1->setBlock( 'Notes 4', blockConcat($eData,6,48,4) );

$eRDA1->setBlock( 'Crew 5', '' );
$eRDA1->setBlock( 'Crew Names 21', '' );
$eRDA1->setBlock( 'Crew Names 22', '' );
$eRDA1->setBlock( 'Crew Names 23', '' );
$eRDA1->setBlock( 'Crew Names 24', '' );
$eRDA1->setBlock( 'Crew Names 25', '' );
$eRDA1->setBlock( 'Notes 5', '' );

echo "Deleted Bulletin: " . $eRDA1->DeletePage($eB2ID) . "<br>";

 echo "Created Bulletin: " . $eRDA1->CreatePage() . "<br>";

if($eRDA1->getLastError() != '')echo $eRDA1->getLastError()."<br>";

$eRDA1->clearBlocks();







$eRDA1->setTemplateName('Schedule static text size');

$eRDA1->Description = $eB3;

$eRDA1->setBlock('Crew Schedule Title', $eB3);

$eRDA1->setBlock( 'Month 1','' );
$eRDA1->setBlock( 'Month 2','' );
$eRDA1->setBlock( 'Month 3','' );
$eRDA1->setBlock( 'Month 4','' );
$eRDA1->setBlock( 'Month 5','' );

$eRDA1->setBlock( 'Date 1','' );
$eRDA1->setBlock( 'Date 2','' );
$eRDA1->setBlock( 'Date 3','' );
$eRDA1->setBlock( 'Date 4','' );
$eRDA1->setBlock( 'Date 5','' );

$eRDA1->setBlock( 'Crew 1', blockConcatNWS($eData,0,73,4) );
$eRDA1->setBlock( 'Crew Names 1', blockConcatNWS($eData,1,73,4) );
$eRDA1->setBlock( 'Crew Names 2', blockConcatNWS($eData,2,73,4) );
$eRDA1->setBlock( 'Crew Names 3', blockConcatNWS($eData,3,73,4) );
$eRDA1->setBlock( 'Crew Names 4', blockConcatNWS($eData,4,73,4) );
$eRDA1->setBlock( 'Crew Names 5', blockConcatNWS($eData,5,73,4) );
$eRDA1->setBlock( 'Notes 1', blockConcatNWS($eData,6,73,4) );

$eRDA1->setBlock( 'Crew 2', blockConcatNWS($eData,0,77,4) );
$eRDA1->setBlock( 'Crew Names 6', blockConcatNWS($eData,1,77,4) );
$eRDA1->setBlock( 'Crew Names 7', blockConcatNWS($eData,2,77,4) );
$eRDA1->setBlock( 'Crew Names 8', blockConcatNWS($eData,3,77,4) );
$eRDA1->setBlock( 'Crew Names 9', blockConcatNWS($eData,4,77,4) );
$eRDA1->setBlock( 'Crew Names 10', blockConcatNWS($eData,5,77,4) );
$eRDA1->setBlock( 'Notes 2', blockConcatNWS($eData,6,77,4) );

$eRDA1->setBlock( 'Crew 3', blockConcatNWS($eData,0,81,3) );
$eRDA1->setBlock( 'Crew Names 11', blockConcatNWS($eData,1,81,3) );
$eRDA1->setBlock( 'Crew Names 12', blockConcatNWS($eData,2,81,3) );
$eRDA1->setBlock( 'Crew Names 13', blockConcatNWS($eData,3,81,3) );
$eRDA1->setBlock( 'Crew Names 14', blockConcatNWS($eData,4,81,3) );
$eRDA1->setBlock( 'Crew Names 15', blockConcatNWS($eData,5,81,3) );
$eRDA1->setBlock( 'Notes 3', blockConcatNWS($eData,6,81,3) );

$eRDA1->setBlock( 'Crew 4', blockConcat($eData,0,84,12) );
$eRDA1->setBlock( 'Crew Names 16', blockConcat($eData,1,84,12) );
$eRDA1->setBlock( 'Crew Names 17', blockConcat($eData,2,84,12) );
$eRDA1->setBlock( 'Crew Names 18', blockConcat($eData,3,84,12) );
$eRDA1->setBlock( 'Crew Names 19', blockConcat($eData,4,84,12) );
$eRDA1->setBlock( 'Crew Names 20', blockConcat($eData,5,84,12) );
$eRDA1->setBlock( 'Notes 4', blockConcat($eData,6,84,12) );

$eRDA1->setBlock( 'Crew 5', '' );
$eRDA1->setBlock( 'Crew Names 21', '' );
$eRDA1->setBlock( 'Crew Names 22', '' );
$eRDA1->setBlock( 'Crew Names 23', '' );
$eRDA1->setBlock( 'Crew Names 24', '' );
$eRDA1->setBlock( 'Crew Names 25', '' );
$eRDA1->setBlock( 'Notes 5', '' );

echo "Deleted Bulletin: " . $eRDA1->DeletePage($eB3ID) . "<br>";

 echo "Created Bulletin: " . $eRDA1->CreatePage() . "<br>";

if($eRDA1->getLastError() != '')echo $eRDA1->getLastError()."<br>";

$eRDA1->clearBlocks();






///////////////////////////////////////////////
////////////////////////////////////////////////////
//////////////////////////////////////////////////
///////////////////////////////////////////////// GAS XLS Document


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



$gRDA1 = new RDA($user, $password, $server);




$gRDA1->setZoneIDs($gZoneID);






$gRDA1->setTemplateName('Schedule');

$gRDA1->Description = 'Gas Crew Schedule';

$gRDA1->setBlock('Crew Schedule Title', 'Gas Crew Schedule');

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

 echo "Created Bulletin: " . $gRDA1->CreatePage() . "<br>";

if($gRDA1->getLastError() != '')echo $gRDA1->getLastError()."<br>";

$gRDA1->clearBlocks();




$gRDA1->setTemplateName('Schedule static text size');

$gRDA1->Description = 'Inspectors/Information';

$gRDA1->setBlock('Crew Schedule Title', 'Inspectors/Information');

$gRDA1->setBlock( 'Month 1',$gData->getCell('D42')->getCalculatedValue() );
$gRDA1->setBlock( 'Month 2',$gData->getCell('E42')->getCalculatedValue() );
$gRDA1->setBlock( 'Month 3',$gData->getCell('F42')->getCalculatedValue() );
$gRDA1->setBlock( 'Month 4',$gData->getCell('G42')->getCalculatedValue() );
$gRDA1->setBlock( 'Month 5',$gData->getCell('H42')->getCalculatedValue() );


$gRDA1->setBlock( 'Date 1',dateConvert($gData->getCell('D43')->getCalculatedValue()) );
$gRDA1->setBlock( 'Date 2',dateConvert($gData->getCell('E43')->getCalculatedValue()) );
$gRDA1->setBlock( 'Date 3',dateConvert($gData->getCell('F43')->getCalculatedValue()) );
$gRDA1->setBlock( 'Date 4',dateConvert($gData->getCell('G43')->getCalculatedValue()) );
$gRDA1->setBlock( 'Date 5',dateConvert($gData->getCell('H43')->getCalculatedValue()) );

$gRDA1->setBlock( 'Crew 1', blockConcatNWS($gData,2,44,4) );
$gRDA1->setBlock( 'Crew Names 1', blockConcatNWS($gData,3,44,4) );
$gRDA1->setBlock( 'Crew Names 2', blockConcatNWS($gData,4,44,4) );
$gRDA1->setBlock( 'Crew Names 3', blockConcatNWS($gData,5,44,4) );
$gRDA1->setBlock( 'Crew Names 4', blockConcatNWS($gData,6,44,4) );
$gRDA1->setBlock( 'Crew Names 5', blockConcatNWS($gData,7,44,4) );
$gRDA1->setBlock( 'Notes 1', blockConcatNWS($gData,8,44,4) );

$gRDA1->setBlock( 'Crew 2', blockConcatNWS($gData,2,48,2) );
$gRDA1->setBlock( 'Crew Names 6', blockConcatNWS($gData,3,48,2) );
$gRDA1->setBlock( 'Crew Names 7', blockConcatNWS($gData,4,48,2) );
$gRDA1->setBlock( 'Crew Names 8', blockConcatNWS($gData,5,48,2) );
$gRDA1->setBlock( 'Crew Names 9', blockConcatNWS($gData,6,48,2) );
$gRDA1->setBlock( 'Crew Names 10', blockConcatNWS($gData,7,48,2) );
$gRDA1->setBlock( 'Notes 2', blockConcatNWS($gData,8,48,2) );

$gRDA1->setBlock( 'Crew 3', blockConcat($gData,2,50,7) );
$gRDA1->setBlock( 'Crew Names 11', blockConcat($gData,3,50,7) );
$gRDA1->setBlock( 'Crew Names 12', blockConcat($gData,4,50,7) );
$gRDA1->setBlock( 'Crew Names 13', blockConcat($gData,5,50,7) );
$gRDA1->setBlock( 'Crew Names 14', blockConcat($gData,6,50,7) );
$gRDA1->setBlock( 'Crew Names 15', blockConcat($gData,7,50,7) );
$gRDA1->setBlock( 'Notes 3', blockConcat($gData,8,50,7) );

$gRDA1->setBlock( 'Crew 4', blockConcat($gData,2,57,5) );
$gRDA1->setBlock( 'Crew Names 16', blockConcat($gData,3,57,5) );
$gRDA1->setBlock( 'Crew Names 17', blockConcat($gData,4,57,5) );
$gRDA1->setBlock( 'Crew Names 18', blockConcat($gData,5,57,5) );
$gRDA1->setBlock( 'Crew Names 19', blockConcat($gData,6,57,5) );
$gRDA1->setBlock( 'Crew Names 20', blockConcat($gData,7,57,5) );
$gRDA1->setBlock( 'Notes 4', blockConcat($gData,8,57,6) );

$gRDA1->setBlock( 'Crew 5', blockConcatNWS($gData,2,57,2) );
$gRDA1->setBlock( 'Crew Names 21', blockConcatNWS($gData,3,62,4) );
$gRDA1->setBlock( 'Crew Names 22', blockConcatNWS($gData,4,62,4) );
$gRDA1->setBlock( 'Crew Names 23', blockConcatNWS($gData,5,62,4) );
$gRDA1->setBlock( 'Crew Names 24', blockConcatNWS($gData,6,62,4) );
$gRDA1->setBlock( 'Crew Names 25', blockConcatNWS($gData,7,62,4) );
$gRDA1->setBlock( 'Notes 5', blockConcatNWS($gData,8,62,4) );

 echo "Created Bulletin: " . $gRDA1->CreatePage() . "<br>";

if($gRDA1->getLastError() != '')echo $gRDA1->getLastError()."<br";

$gRDA1->clearBlocks();







$gRDA1->setTemplateName('Schedule');

$gRDA1->Description = 'CDA GAS Servicemen';

$gRDA1->setBlock('Crew Schedule Title', 'CDA GAS Servicemen');

$gRDA1->setBlock( 'Month 1',$gData->getCell('D79')->getCalculatedValue() );
$gRDA1->setBlock( 'Month 2',$gData->getCell('E79')->getCalculatedValue() );
$gRDA1->setBlock( 'Month 3',$gData->getCell('F79')->getCalculatedValue() );
$gRDA1->setBlock( 'Month 4',$gData->getCell('G79')->getCalculatedValue() );
$gRDA1->setBlock( 'Month 5',$gData->getCell('H79')->getCalculatedValue() );

$gRDA1->setBlock( 'Date 1',dateConvert($gData->getCell('D80')->getCalculatedValue()) );
$gRDA1->setBlock( 'Date 2',dateConvert($gData->getCell('E80')->getCalculatedValue()) );
$gRDA1->setBlock( 'Date 3',dateConvert($gData->getCell('F80')->getCalculatedValue()) );
$gRDA1->setBlock( 'Date 4',dateConvert($gData->getCell('G80')->getCalculatedValue()) );
$gRDA1->setBlock( 'Date 5',dateConvert($gData->getCell('H80')->getCalculatedValue()) );

$gRDA1->setBlock( 'Crew 1', blockConcat($gData,2,81,2) );
$gRDA1->setBlock( 'Crew Names 1', $gData->getCell('D81')->getCalculatedValue() );
$gRDA1->setBlock( 'Crew Names 2', $gData->getCell('E81')->getCalculatedValue() );
$gRDA1->setBlock( 'Crew Names 3', $gData->getCell('F81')->getCalculatedValue() );
$gRDA1->setBlock( 'Crew Names 4', $gData->getCell('G81')->getCalculatedValue() );
$gRDA1->setBlock( 'Crew Names 5', $gData->getCell('H81')->getCalculatedValue() );
$gRDA1->setBlock( 'Notes 1', $gData->getCell('I81')->getCalculatedValue() );

$gRDA1->setBlock( 'Crew 2', blockConcat($gData,2,83,2) );
$gRDA1->setBlock( 'Crew Names 6', $gData->getCell('D83')->getCalculatedValue() );
$gRDA1->setBlock( 'Crew Names 7', $gData->getCell('E83')->getCalculatedValue() );
$gRDA1->setBlock( 'Crew Names 8', $gData->getCell('F83')->getCalculatedValue() );
$gRDA1->setBlock( 'Crew Names 9', $gData->getCell('G83')->getCalculatedValue() );
$gRDA1->setBlock( 'Crew Names 10', $gData->getCell('H83')->getCalculatedValue() );
$gRDA1->setBlock( 'Notes 2', $gData->getCell('I83')->getCalculatedValue() );

$gRDA1->setBlock( 'Crew 3', blockConcat($gData,2,85,2) );
$gRDA1->setBlock( 'Crew Names 11', $gData->getCell('D85')->getCalculatedValue() );
$gRDA1->setBlock( 'Crew Names 12', $gData->getCell('E85')->getCalculatedValue() );
$gRDA1->setBlock( 'Crew Names 13', $gData->getCell('F85')->getCalculatedValue() );
$gRDA1->setBlock( 'Crew Names 14', $gData->getCell('G85')->getCalculatedValue() );
$gRDA1->setBlock( 'Crew Names 15', $gData->getCell('H85')->getCalculatedValue() );
$gRDA1->setBlock( 'Notes 3', $gData->getCell('I85')->getCalculatedValue() );

$gRDA1->setBlock( 'Crew 4', blockConcat($gData,2,87,2) );
$gRDA1->setBlock( 'Crew Names 16', $gData->getCell('D87')->getCalculatedValue() );
$gRDA1->setBlock( 'Crew Names 17', $gData->getCell('E87')->getCalculatedValue() );
$gRDA1->setBlock( 'Crew Names 18', $gData->getCell('F87')->getCalculatedValue() );
$gRDA1->setBlock( 'Crew Names 19', $gData->getCell('G87')->getCalculatedValue() );
$gRDA1->setBlock( 'Crew Names 20', $gData->getCell('H87')->getCalculatedValue() );
$gRDA1->setBlock( 'Notes 4', $gData->getCell('I87')->getCalculatedValue() );

$gRDA1->setBlock( 'Crew 5', blockConcat($gData,2,89,2) );
$gRDA1->setBlock( 'Crew Names 21', $gData->getCell('D89')->getCalculatedValue() );
$gRDA1->setBlock( 'Crew Names 22', $gData->getCell('E89')->getCalculatedValue() );
$gRDA1->setBlock( 'Crew Names 23', $gData->getCell('F89')->getCalculatedValue() );
$gRDA1->setBlock( 'Crew Names 24', $gData->getCell('G89')->getCalculatedValue() );
$gRDA1->setBlock( 'Crew Names 25', $gData->getCell('H89')->getCalculatedValue() );
$gRDA1->setBlock( 'Notes 5', $gData->getCell('I89')->getCalculatedValue() );

 echo "Created Bulletin: " . $gRDA1->CreatePage() . "<br>";

if($gRDA1->getLastError() != '')echo $gRDA1->getLastError()."<br>";

$gRDA1->clearBlocks();







///////////////////////////////////////////////
////////////////////////////////////////////////////
//////////////////////////////////////////////////
///////////////////////////////////////////////// Generic Functions



function blockConcat($data,$col,$row,$rowCount){
	
	$result = '';
	
	for($x = 0;$x<$rowCount;$x++){
		
		$cell = $data->getCellByColumnAndRow($col,$row+$x)->getCalculatedValue();
		if(trim($cell) != ''){
			$result .= $cell;
			$result .= "<br>";
		}
	}

	return $result;

}

function blockConcatNWS($data,$col,$row,$rowCount){
	
	$result = '';
	
	for($x = 0;$x<$rowCount;$x++){
		
		$cell = $data->getCellByColumnAndRow($col,$row+$x)->getCalculatedValue();
		
		$result .= $cell;
		$result .= "<br>";
		
	}

	return $result;

}

function dateConvert($float){

	$date = date_create('1899-12-30');
	$date->add(new DateInterval('P'.$float.'D'));
	return $date->format('jS');

}




