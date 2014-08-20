<?php


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