<?php

class chunkReadFilter implements PHPExcel_Reader_IReadFilter{

	private $_startRow = 0;
	private $_endRow = 0;

	public function __construct($startRow, $chunkSize){

		$this->_startRow = $startRow;
		$this->_endRow = $startRow + $chunkSize;
	}

	public function readCell($column, $row, $worksheetname = ''){

		if (($row == 1) || ($row >= $this->_startRow && $row < $this->_endRow)) {
			return true;
		}
		return false;
		
	}

}