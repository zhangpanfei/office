<?php

namespace zpfei\office;
use PHPExcel;
use PHPExcel_Writer_Excel2007;
use Exception;

class Excel
{
	public $excel;
	public $properties = ['creator','lastmodifiedby','title','subject','description'];
	public $data = [];
	public $index = false;
	public function __construct()
	{
		$this->excel = new PHPExcel();
	}

	public function __call($action,$params)
	{
		$action = 'set'.$action;
		if (in_array(strtolower($action), $this->properties)) {
			throw new Exception("None propertie {$action}");
		}
		if (!method_exists($this->excel->getProperties(), $action)) {
			throw new Exception("None action {$action}");
		}

		call_user_func([$this->excel->getProperties(),$action],$params[0]);
		return $this;
	}

	public function sheet($index=0)
	{
		$this->excel->setActiveSheetIndex(0);
		return $this;
	}

	public function header(Array $data)
	{
		array_unshift($this->data, $data);
		return $this;
	}
	public function data(Array $data)
	{
		$this->data = array_merge($this->data, $data);
		return $this;
	}

	private function setData()
	{
		if ($this->index===false) {
			$this->sheet();
		}

		$flag = range('A','Z');
		foreach ($this->data as $key=>$data) {
			$line = $key+1;
			foreach ($data as $k=>$val) {
				$this->excel->getActiveSheet()->SetCellValue($flag[$k].$line, $val);
			}
		}
	}

	public function save($filename)
	{
		$this->setData();
		$objWriter = new PHPExcel_Writer_Excel2007($this->excel);
		$objWriter->save($filename);
	}
	public function output($filename)
	{
		$this->setData();
		$objWriter = new PHPExcel_Writer_Excel2007($this->excel);
		$objWriter->save($filename);
		header("Pragma: public");
		header("Expires: 0");
		header("Cache-Control:must-revalidate, post-check=0, pre-check=0");
		header("Content-Type:application/force-download");
		header("Content-Type:application/vnd.ms-execl");
		header("Content-Type:application/octet-stream");
		header("Content-Type:application/download");
		header('Content-Disposition:attachment;filename='.$filename);
		header("Content-Transfer-Encoding:binary");
		$objWriter->save('php://output');
		$res = @unlink($filename);
		return $res;
	}
}