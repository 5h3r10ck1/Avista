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

	return substr($result,0,-4);

}

function blockConcatNWS($data,$col,$row,$rowCount){
	
	$result = '';
	
	for($x = 0;$x<$rowCount;$x++){
		
		$cell = $data->getCellByColumnAndRow($col,$row+$x)->getCalculatedValue();
		
		$result .= $cell;
		$result .= "<br>";
		
	}

	return substr($result,0,-4);

}

function blockConcatInline($data,$col,$row,$rowCount){
	
	$result = '';
	
	for($x = 0;$x<$rowCount;$x++){
		
		$cell = $data->getCellByColumnAndRow($col,$row+$x)->getCalculatedValue();
		if(trim($cell) != ''){
			$result .= $cell;
			$result .= ", ";
		}

	}

	return substr($result,0,-2);

}

function blockConcatInlineNC($data,$col,$row,$rowCount){
	
	$result = '';
	
	for($x = 0;$x<$rowCount;$x++){
		
		$cell = $data->getCellByColumnAndRow($col,$row+$x)->getCalculatedValue();
		if(trim($cell) != ''){
			$result .= $cell;
			$result .= " ";
		}

	}

	return substr($result,0,-1);

}

function dateConvert($float){
	
	if(is_float($float)){
		$date = date_create('1899-12-30');
		$date->add(new DateInterval('P'.$float.'D'));
		return $date->format('j');
	}
	else{
		var_dump($float . ' is not a valid Excel Date');
		$log = fopen('errorLog.txt','a+');
		$message = date('Y-m-d H:i:s').' ' . $float . 'is not a valid Excel Date'."\r\n";
		fwrite($log, $message);
		fclose($log);
		return $float;
	} 

}

function imageConvert($drawing,$type){
	
	global $i;

	global $imagesPath;

	if ($drawing instanceof PHPExcel_Worksheet_MemoryDrawing) {
		ob_start();
		call_user_func(
			$drawing->getRenderingFunction(),
			$drawing->getImageResource()
		);
		$imageContents = ob_get_contents();
		ob_end_clean();
		switch ($drawing->getMimeType()) {
			case PHPExcel_Worksheet_MemoryDrawing::MIMETYPE_PNG :
					$extension = 'png'; break;
			case PHPExcel_Worksheet_MemoryDrawing::MIMETYPE_GIF:
					$extension = 'gif'; break;
			case PHPExcel_Worksheet_MemoryDrawing::MIMETYPE_JPEG :
					$extension = 'jpg'; break;
		}
	} else {
		$zipReader = fopen($drawing->getPath(),'r');
		$imageContents = '';
		while (!feof($zipReader)) {
			$imageContents .= fread($zipReader,1024);
		}
		fclose($zipReader);
		$extension = $drawing->getExtension();
	}
	
	$myFileName = $imagesPath.$type.++$i.'.'.$extension;
	
	file_put_contents($myFileName,$imageContents);
	
	if($remoteImages)ftp_file($myFileName);

	return $myFileName;

}




function isAssoc($arr)
{
    return array_keys($arr) !== range(0, count($arr) - 1);
}







function ftp_file($file){
	$ftp_server = "ftp.sphillipssite.com";
	$ftp_conn = ftp_connect($ftp_server) or die("Could not connect to $ftp_server");
	$login = ftp_login($ftp_conn, "sphillips1976", 'Seth7384!');
	ftp_pasv($ftp_conn, true);
	ftp_chdir($ftp_conn, '/images/');


	// upload file
	if (ftp_put($ftp_conn, substr($file, strpos($file, '/')+1), $file, FTP_BINARY))
	  {
	  echo "Successfully uploaded $file<br>";
	  }
	else
	  {
	  	$log = fopen('errorLog.txt','a+');
		$message = date('Y-m-d H:i:s').' ' .'error uploading file'."\r\n";
		fwrite($log, $message);
		fclose($log);
	  }

	// close connection
	ftp_close($ftp_conn);

}

function getSMBFiles($file){

	global $remoteFolder;
	global $SMBUser;
	global $SMBPass;
	global $imagesPath;

	$smbc = new smbclient ($remoteFolder, $SMBUser, $SMBPass);

	if (!$smbc->get ($file, $imagesPath.$file))
	{
	    $log = fopen('errorLog.txt','a+');
		$message = date('Y-m-d H:i:s').' ' .'Failed to retrieve file '.$file."\r\n";
		fwrite($log, $message);
		fclose($log);
	    print "Failed to retrieve file:\n";
	    
	}
	else
	{
	    echo "Transferred file successfully.<br>";
	}
}