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
$excel->title('Example')->header($header)->data($data)->output('examlpe.xls');