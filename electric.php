<?php

include_once 'RDA.php';
include_once 'config.php';
include_once 'Classes/PHPExcel.php';
include_once 'functions.php';

set_time_limit(0);

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

if(isset($eB1ID)) $eRDA1->DeletePage($eB1ID);

$eRDA1->CreatePage();

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

if(isset($eB2ID))$eRDA1->DeletePage($eB2ID);

$eRDA1->CreatePage();

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

if(isset($eB3ID))$eRDA1->DeletePage($eB3ID);

$eRDA1->CreatePage();

if($eRDA1->getLastError() != '')echo $eRDA1->getLastError()."<br>";

$eRDA1->clearBlocks();