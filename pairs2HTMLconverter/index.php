<?php
header('Content-Type: text/html; charset=utf-8');

// для великих таблиць Excel може знадобиться
// збільшення об'єму пам'яті, а також часу виконання скрипта
// (в такому разі розкоментуйте один або два рядки нижче)

ini_set("memory_limit", "64M");
ini_set("max_execution_time", 60*5); // 60*5 сек = 5 хв.

// під'єднюємо необхідний для читання таблиці скрипт
require_once "Classes/PHPExcel/Reader/Excel5.php";


$excelFileName = "tablelist.xls";		// Excel filename

$objReader = new PHPExcel_Reader_Excel5();
$objPHPExcel = $objReader->load( $excelFileName );
$objWorksheet = $objPHPExcel->getActiveSheet();

$pairs = array();								//Array of pairs
$pairs_number = 4;								//Number of pairs per day
$mergeCells = $objWorksheet->getMergeCells();	//Array of merge cells

//echo "<pre>";print_r($mergeCells);

foreach ($objWorksheet->getRowIterator() as $row) {
	
	$cellIterator = $row->getCellIterator();
	$cellIterator->setIterateOnlyExistingCells(false);
	
	//Pairs counter
	$pair = 0;
	
	//Traversal of the each pairs
	for($curPair = 1; $curPair <= $pairs_number; $curPair++){
		//0-monday, 1-thuesday ...
		$dayOfWeek = 0;
		
		foreach ($cellIterator as $cell) {
		
			//Current col index
			$col_indx = $cell->columnIndexFromString($cell->getColumn());
			
			//echo "<pre>"; print_r($cell->getValue());
			
			//First collumn is the name of lecturer
			if($col_indx == 1){
				// Print name of lecturer
				if($curPair == 1) echo $cell->getValue();
				continue;
			}
			
			//Pair counter
			$pair += .5;
			
			if(round($pair) == $curPair){
				//echo "<pre> - "; print_r($cell->getValue());
				//Week type
				$weekType = $col_indx % 2; // 0 - NUMERATOR ... - DENOMINATOR
				
				if($weekType == 0){
					//First line - audience, second - pair and groups in EXCEL file
					if($row->getRowIndex() % 2){
						
						//Detect merge cells
						foreach($mergeCells as $mergeCell){
							if($cell->isInRange($mergeCell)){
								$pairs[$curPair]['num'][$dayOfWeek]['merge'] = 1;
							}
						}
						
						$pairs[$curPair]['num'][$dayOfWeek]['room'] = $cell->getValue();
					}else{
						//preg_match("/^(.*)\n(.*)$/",$cell->getValue(), $out);
						
						//PAIR FOR THE GROUP OR ANOTHER TEXT
						preg_match("/^(?:(.*)\n(.*)|(.*))$/",$cell->getValue(), $out);
						$pairs[$curPair]['num'][$dayOfWeek]['subject'] = !$out[1] ? $out[3] : $out[1];
						$pairs[$curPair]['num'][$dayOfWeek]['groups'] = $out[2];
						
					}
				}else{
					//First line - audience, second - pair and groups in EXCEL file
					if($row->getRowIndex() % 2){ 
						$pairs[$curPair]['den'][$dayOfWeek]['room'] = $cell->getValue();
						//Detect merge cells
						foreach($mergeCells as $mergeCell){
							if($cell->isInRange($mergeCell)){
								$pairs[$curPair]['den'][$dayOfWeek]['merge'] = 1;
								//$pairs[$curPair]['den'][$dayOfWeek]['merge'] = $cell->getCoordinate()." = ".$mergeCell;
							}
						}
					}else{
						//preg_match("/^(.*)\n(.*)$/",$cell->getValue(), $out);
						
						//PAIR FOR THE GROUP OR ANOTHER TEXT
						preg_match("/^(?:(.*)\n(.*)|(.*))$/",$cell->getValue(), $out);
						$pairs[$curPair]['den'][$dayOfWeek]['subject'] = !$out[1] ? $out[3] : $out[1];
						$pairs[$curPair]['den'][$dayOfWeek]['groups'] = $out[2];
					}
				}
			}
			
			//Constrains
			if($pair == $pairs_number) $pair = 0;
			if( ($col_indx - 1) % ($pairs_number * 2) == 0) $dayOfWeek++;
			//echo " - ".$col_indx."<br />";
		}
	}
	//Printing results
	if($row->getRowIndex() % 2 == 0){
		//echo "<pre>"; print_r($pairs);
		echo '<table border="1">' . tablelist2HTML($pairs) . "</table>";
		echo "<textarea cols='50' rows='20'>" . tablelist2HTML($pairs) . "</textarea><br /><br /><br /><br />";
		
		//Empty array $pairs
		$pairs = array();
	}
}

/*** Function for converting tablelist in HTML ***/
function tablelist2HTML($pairs){
	$pairTimes = array( 
		1=>'I<br />8:00<br />9:35', 
		2=>'II<br />9:50<br />11:25',
		3=>'III<br />11:55<br />13:30', 
		4=>'IV<br />13:45<br />15:20',
		5=>'V<br />15:35<br />17:10' 
	);
	
	$html = '';
	
	foreach($pairs as $pairID=>$pair){
		foreach($pair as $typeWeek => $days){
			if($typeWeek == 'num'){
				$html .= "\r\n" . "<!-- PAIR $pairID -->" . "\r\n";
				$html .= '<tr class="num">' . "\r\n";
				$html .= "\t" . '<td rowspan="2" class="pair">'.$pairTimes[$pairID].'</td>' . "\r\n";
				foreach($days as $day){
					$html .= createCell($day['subject'], $day['room'], $day['groups'], !empty($day['merge']));
				}
				$html .= "</tr>\r\n";
			}elseif($typeWeek == 'den'){
				$html .= '<tr class="den">' . "\r\n";
				foreach($days as $day){
					if(!empty($day['merge'])){continue;}
					
					$html .= createCell($day['subject'], $day['room'], $day['groups']);
				}
				$html .= "</tr>\r\n";
			}
		}
	}
	
	return $html;
}

/* Function for generating cells */ 
function createCell($subject, $room = '', $groups = '', $merge = false){
	$merge ? $merge = ' class="merge" rowspan="2"' : '';
	$groups == '' ? $separator = '' : $separator = ' / ';
	$groups == '' ? $groups = '' : $groups = "<div>$groups</div>";
	
	return $subject == '' ? "\t<td>-</td>\r\n" : "\t<td$merge>$subject$separator$room$groups</td>\r\n";
}

//echo "<pre>";
//print_r($pairs);

/*$pairs = array(
	1=>array(
		'num'=>array(
			'Пн'=>array(
				'room'=>'317м',
				'subject'=>'ТММ',
				'groups'=>'110опс,130'
			),'Вт'=>array(
				'room'=>'318м',
				'subject'=>'ТехМ',
				'groups'=>'110опс,130'
			),'Ср'=>array(
				'room'=>'',
				'subject'=>'',
				'groups'=>''
			),'Чт'=>array(
				'room'=>'',
				'subject'=>'',
				'groups'=>''
			),'Пт'=>array(
				'room'=>'',
				'subject'=>'',
				'groups'=>''
			)
		),'den'=>array(
			'Пн'=>array(
				'room'=>'',
				'subject'=>'',
				'groups'=>''
			),'Вт'=>array(
				'room'=>'318м',
				'subject'=>'ТехМ',
				'groups'=>'110опс,130'
			),'Ср'=>array(
					'room'=>'',
					'subject'=>'',
					'groups'=>''
			),'Чт'=>array(
					'room'=>'',
					'subject'=>'',
					'groups'=>''
			),'Пт'=>array(
					'room'=>'',
					'subject'=>'',
					'groups'=>''
			)
		)
	),2=>array(
		'num'=>array(
			'Пн'=>array(
				'room'=>'317м',
				'subject'=>'ТММ',
				'groups'=>'110опс,130'
			),'Вт'=>array(
				'room'=>'318м',
				'subject'=>'ТехМ',
				'groups'=>'110опс,130'
			),'Ср'=>array(
				'room'=>'',
				'subject'=>'',
				'groups'=>''
			),'Чт'=>array(
				'room'=>'',
				'subject'=>'',
				'groups'=>''
			),'Пт'=>array(
				'room'=>'',
				'subject'=>'',
				'groups'=>''
			)
		),'den'=>array(
			'Пн'=>array(
				'room'=>'',
				'subject'=>'',
				'groups'=>''
			),'Вт'=>array(
				'room'=>'318м',
				'subject'=>'ТехМ',
				'groups'=>'110опс,130'
			),'Ср'=>array(
					'room'=>'',
					'subject'=>'',
					'groups'=>''
			),'Чт'=>array(
					'room'=>'',
					'subject'=>'',
					'groups'=>''
			),'Пт'=>array(
					'room'=>'',
					'subject'=>'',
					'groups'=>''
			)
		)
	)
);*/
?>
