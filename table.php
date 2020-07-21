<?php
//call the autoload
require 'vendor/autoload.php';
//load phpspreadsheet class using namespaces
use PhpOffice\PhpSpreadsheet\Spreadsheet;
//call iofactory instead of xlsx writer
use PhpOffice\PhpSpreadsheet\IOFactory;
use PhpOffice\PhpSpreadsheet\Style\Alignment;
use PhpOffice\PhpSpreadsheet\Style\Fill;

//styling arrays
//table head style
$tableHead = [
	'font'=>[
		'color'=>[
			'rgb'=>'FFFFFF'
		],
		'bold'=>true,
		'size'=>11
	],
	'fill'=>[
		'fillType' => Fill::FILL_SOLID,
		'startColor' => [
			'rgb' => '538ED5'
		]
	],
];
//even row
$evenRow = [
	'fill'=>[
		'fillType' => Fill::FILL_SOLID,
		'startColor' => [
			'rgb' => '00BDFF'
		]
	]
];
//odd row
$oddRow = [
	'fill'=>[
		'fillType' => Fill::FILL_SOLID,
		'startColor' => [
			'rgb' => '00EAFF'
		]
	]
];

//styling arrays end

//make a new spreadsheet object
$spreadsheet = new Spreadsheet();
//get current active sheet (first sheet)
$sheet = $spreadsheet->getActiveSheet();

//set default font
$spreadsheet->getDefaultStyle()
	->getFont()
	->setName('Arial')
	->setSize(10);

//heading
$spreadsheet->getActiveSheet()
	->setCellValue('A1',"Participant Students");

//merge heading
$spreadsheet->getActiveSheet()->mergeCells("A1:F1");

// set font style
$spreadsheet->getActiveSheet()->getStyle('A1')->getFont()->setSize(20);

// set cell alignment
$spreadsheet->getActiveSheet()->getStyle('A1')->getAlignment()->setHorizontal(Alignment::HORIZONTAL_CENTER);

//setting column width
$spreadsheet->getActiveSheet()->getColumnDimension('A')->setWidth(5);
$spreadsheet->getActiveSheet()->getColumnDimension('B')->setWidth(20);
$spreadsheet->getActiveSheet()->getColumnDimension('C')->setWidth(20);
$spreadsheet->getActiveSheet()->getColumnDimension('D')->setWidth(30);
$spreadsheet->getActiveSheet()->getColumnDimension('E')->setWidth(12);
$spreadsheet->getActiveSheet()->getColumnDimension('F')->setWidth(10);

//header text
$spreadsheet->getActiveSheet()
	->setCellValue('A2',"ID")
	->setCellValue('B2',"First Name")
	->setCellValue('C2',"Last Name")
	->setCellValue('D2',"Email")
	->setCellValue('E2',"Gender")
	->setCellValue('F2',"Class");

//set font style and background color
$spreadsheet->getActiveSheet()->getStyle('A2:F2')->applyFromArray($tableHead);

//the content
//read the json file
$file = file_get_contents('student-data.json');
$studentData = json_decode($file,true);

//loop through the data
//current row
$row=3;
foreach($studentData as $student){
	$spreadsheet->getActiveSheet()
		->setCellValue('A'.$row , $student['id'])
		->setCellValue('B'.$row , $student['first_name'])
		->setCellValue('C'.$row , $student['last_name'])
		->setCellValue('D'.$row , $student['email'])
		->setCellValue('E'.$row , $student['gender'])
		->setCellValue('F'.$row , $student['class']);
	
	//set row style
	if( $row % 2 == 0 ){
		//even row
		$spreadsheet->getActiveSheet()->getStyle('A'.$row.':F'.$row)->applyFromArray($evenRow);
	}else{
		//odd row
		$spreadsheet->getActiveSheet()->getStyle('A'.$row.':F'.$row)->applyFromArray($oddRow);
	}
	//increment row
	$row++;
}

//autofilter
//define first row and last row
$firstRow=2;
$lastRow=$row-1;
//set the autofilter
$spreadsheet->getActiveSheet()->setAutoFilter("A".$firstRow.":F".$lastRow);














//set the header first, so the result will be treated as an xlsx file.
header('Content-Type: application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');

//make it an attachment so we can define filename
header('Content-Disposition: attachment;filename="result.xlsx"');

//create IOFactory object
$writer = IOFactory::createWriter($spreadsheet, 'Xlsx');
//save into php output
$writer->save('php://output');
