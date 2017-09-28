<?php
include '../vendor/autoload.php';

use zpfei\office\Excel;

$header = ['用户名','密码','年龄'];
$data = [
	['李雷',123456,18],
	['韩梅梅',789456,16],
	['小白',898566,19],
];
$excel = new Excel();
$excel->title('Example')
	  ->sheet(0)->header($header)->data($data)
	  ->sheet(1)->header($header)->data($data)
	  ->sheet(2)->header($header)->data($data)
	  ->sheet(3)->header($header)->data($data)
	  ->sheet(4)->header($header)->data($data)

	  ->save('emamlpe.xls');