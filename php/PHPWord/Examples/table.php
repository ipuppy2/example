<?php
header('Content-Type:text/html;charset=utf-8');
require_once '../PHPWord.php';

// New Word Document
$PHPWord = new PHPWord();

// New portrait section
$section = $PHPWord->createSection();

// Define table style arrays
/*$styleTable = array('borderSize'=>6, 'borderColor'=>'006699', 'cellMargin'=>80);
$styleFirstRow = array('borderBottomSize'=>18, 'borderBottomColor'=>'0000FF', 'bgColor'=>'66BBFF');
*/


// Define cell style arrays
$styleCell = array('valign'=>'center');
$styleCellBTLR = array('valign'=>'center', 'textDirection'=>PHPWord_Style_Cell::TEXT_DIR_BTLR);

// Define font style for first row
// label 的样式
$titleFontStyle = array('bold'=>true, 'align'=>'center');
$valueFontStyle=array();


// 
$cellValueMergeStartStyle=array_merge($valueFontStyle,array(
	'cellMerge'=>'restart',
));

$cellTitleMergeStartStyle=array_merge($titleFontStyle,array(
	'cellMerge'=>'restart',
));

$rowTitleMergeStartStyle=array_merge($titleFontStyle,array(
	'rowMerge'=>'restart',
));

define('CELL_LABEL_TITLE_WIDTH',3500);



// 标题
$PHPWord->addFontStyle('headTitleFontStyle', array('bold'=>true/*, 'italic'=>true*/, 'size'=>16));
$PHPWord->addParagraphStyle('headTitlePStyle', array('align'=>'center', 'spaceAfter'=>100));
$section->addText('澳 通 人 才 网 登 记 表','headTitleFontStyle','headTitlePStyle');

// 备注
$PHPWord->addFontStyle('markFontStyle', array('size'=>12));
$PHPWord->addParagraphStyle('markPStyle', array('align'=>'right', 'spaceAfter'=>100));
$section->addText('填表日期：　　　年　　月　　日','markFontStyle','markPStyle');



// 添加表格的样式
$PHPWord->addTableStyle('myOwnTableStyle', array(
	'borderSize'=>6,
	// 'borderColor'=>'006699',
	'cellMargin'=>80
), array(
	// 'borderBottomSize'=>18,
	// 'borderBottomColor'=>'0000FF',
	// 'bgColor'=>'66BBFF'
	)
);

// 添加一个表格
$table = $section->addTable('myOwnTableStyle');


// 第一行
$table->addRow(/*900*/);

// Add cells
$table->addCell(CELL_LABEL_TITLE_WIDTH, $styleCell)->addText('姓名', $titleFontStyle);
$table->addCell(2000, $styleCell)->addText('曾繁斌', $valueFontStyle);


$table->addCell(CELL_LABEL_TITLE_WIDTH, $styleCell)->addText('性别', $titleFontStyle);
$table->addCell(2000, $styleCell)->addText('男', $valueFontStyle);


$table->addCell(CELL_LABEL_TITLE_WIDTH, $styleCell)->addText('出生年月', $titleFontStyle);
$table->addCell(2000, $styleCell)->addText('1991-11', $valueFontStyle);

$table->addCell(CELL_LABEL_TITLE_WIDTH, $styleCell)->addText('民族', $titleFontStyle);
$table->addCell(2000, $styleCell)->addText('汉族', $valueFontStyle);

$table->addCell(2000, array(
	'rowMerge'=>'restart',
))->addImage('_earth.JPG', array('width'=>100, 'height'=>100, 'align'=>'center'));


// 第二行
$table->addRow();
$table->addCell(CELL_LABEL_TITLE_WIDTH, $styleCell)->addText('身高', $titleFontStyle);
$table->addCell(2000, $styleCell)->addText('曾繁斌', $valueFontStyle);


$table->addCell(CELL_LABEL_TITLE_WIDTH, $styleCell)->addText('体重', $titleFontStyle);
$table->addCell(2000, $styleCell)->addText('男', $valueFontStyle);


$table->addCell(CELL_LABEL_TITLE_WIDTH, $styleCell)->addText('出生地', $titleFontStyle);
$table->addCell(2000, $styleCell)->addText('1991-11', $valueFontStyle);

$table->addCell(CELL_LABEL_TITLE_WIDTH, $styleCell)->addText('婚姻状况', $titleFontStyle);
$table->addCell(2000, $styleCell)->addText('汉族', $valueFontStyle);

$table->addCell(2000, array(
	'rowMerge'=>'continue',
));


// 第三行
$table->addRow();
$table->addCell(CELL_LABEL_TITLE_WIDTH, $styleCell)->addText('身份证号', $titleFontStyle);
$table->addCell(2000, array(
	'cellMerge'=>'restart',
))->addText('ffffffffff');

for($i=0;$i<6;$i++){
	$table->addCell(2000, array(
		'cellMerge'=>'continue',
	));
}

$table->addCell(2000, array(
	'rowMerge'=>'continue',
));

// 第四行
$table->addRow();
$table->addCell(CELL_LABEL_TITLE_WIDTH, $styleCell)->addText('户口所在地', $titleFontStyle);
$table->addCell(2000, array(
	'cellMerge'=>'restart',
))->addText('户口');

mergeCell(6);

$table->addCell(2000, array(
	'rowMerge'=>'continue',
));

// 第五行
$table->addRow();
$table->addCell(CELL_LABEL_TITLE_WIDTH, $styleCell)->addText('QQ号码', $titleFontStyle);
$table->addCell(2000, $cellValueMergeStartStyle)->addText('1991-11', $valueFontStyle);
mergeCell(2);

$table->addCell(CELL_LABEL_TITLE_WIDTH, $cellTitleMergeStartStyle)->addText('微信号码', $titleFontStyle);
mergeCell(1);
$table->addCell(2000, $cellValueMergeStartStyle)->addText('1991-11', $valueFontStyle);
mergeCell(2);

// 第六行
$table->addRow();
$table->addCell(CELL_LABEL_TITLE_WIDTH, $styleCell)->addText('联系电话', $titleFontStyle);
$table->addCell(2000, $cellValueMergeStartStyle)->addText('1991-11', $valueFontStyle);
mergeCell(2);

$table->addCell(CELL_LABEL_TITLE_WIDTH, $cellTitleMergeStartStyle)->addText('朋友号码', $titleFontStyle);
mergeCell(1);
$table->addCell(2000, $cellValueMergeStartStyle)->addText('1991-11', $valueFontStyle);
mergeCell(2);


// 教育
outputMutilRows(array(
	'labelTitle'=>'受教育情况',

	'titles'=>array(
		array(
			'name'=>'就读时间',
			'cell'=>3,
		),

		array(
			'name'=>'学校名称',
			'cell'=>2,
		),

		array(
			'name'=>'所学专业',
			'cell'=>2,
		),

		array(
			'name'=>'学历',
			'cell'=>1,
		),
	),

	'values'=>array(
		array(
			'2010年9月至2013年6月',
			'广西水利电力职业技术学院',
			'计算机应用与技术',
			'大专'
		),

		array(
			'2010年9月至2013年6月',
			'广西水利电力职业技术学院',
			'计算机应用与技术',
			'大专'
		),
	),
));

// 工作经历
outputMutilRows(array(
	'labelTitle'=>'工作经历',

	'titles'=>array(
		array(
			'name'=>'起止时间',
			'cell'=>3,
		),

		array(
			'name'=>'公司单位名称',
			'cell'=>2,
		),

		array(
			'name'=>'担任职务',
			'cell'=>2,
		),

		array(
			'name'=>'工资',
			'cell'=>1,
		),
	),

	'values'=>array(
		array(
			'2010年9月至2013年6月',
			'珠海科速互联网络技术有限公司',
			'网页设计师',
			'8500'
		),

		array(
			'2010年9月至2013年6月',
			'珠海科速互联网络技术有限公司',
			'网页设计师',
			'8500'
		),
	),
));

// 语言
outputMutilRows(array(
	'labelTitle'=>'语言水平',

	'titles'=>array(
		array(
			'name'=>'语种',
			'cell'=>2,
		),

		array(
			'name'=>'阅读',
			'cell'=>2,
		),

		array(
			'name'=>'听说',
			'cell'=>2,
		),

		array(
			'name'=>'写作',
			'cell'=>1,
		),

		array(
			'name'=>'其他语言',
			'cell'=>1,
		),
	),

	'values'=>array(
		array(
			'粤语',
			'一般',
			'一般',
			'一般',
			'其它语言',

		),

		array(
			'英语',
			'一般',
			'一般',
			'一般',
			'其它语言',

		),
	),
));

// 求职意向
outputMutilRows(array(
	'labelTitle'=>'求职意向',

	'titles'=>array(
		array(
			'name'=>'意向职位',
			'cell'=>3,
		),

		array(
			'name'=>'期待工资',
			'cell'=>2,
		),

		array(
			'name'=>'住的问题',
			'cell'=>2,
		),

		array(
			'name'=>'吃的问题',
			'cell'=>1,
		),
	),

	'values'=>array(
		array(
			'php工程师',
			'9500',

			'都可',
			'都可',
		)
	),
));

// 其它备注
$table->addRow();
$table->addCell(CELL_LABEL_TITLE_WIDTH, $cellTitleMergeStartStyle)->addText('备注', $titleFontStyle);
$table->addCell(2000, $cellValueMergeStartStyle)->addText('1991-11', $valueFontStyle);
mergeCell(7);



/**
 * 输出多行
 * @param  [type] $configs [description]
 * @return [type]          [description]
 */
function outputMutilRows($configs){

	global $rowTitleMergeStartStyle,
	$titleFontStyle,
	$cellValueMergeStartStyle,
	$valueFontStyle,
	$styleCell ,
	$table;


	$titleList=$configs['titles'];
	$valueList=$configs['values'];
	$table->addRow();

	$table->addCell(CELL_LABEL_TITLE_WIDTH, array_merge(
			$rowTitleMergeStartStyle,
			array(
				// 'valign'=>'middle',
			)
		)
	)->addText($configs['labelTitle'], $titleFontStyle);

	$cellMerges=array();
	foreach($titleList as $titles){
		$table->addCell(CELL_LABEL_TITLE_WIDTH, $titles['cell']>1?$cellValueMergeStartStyle:$styleCell)->addText($titles['name'], $titleFontStyle);

		// $table->addCell();
		
		mergeCell(--$titles['cell']);
		$cellMerges[]=$titles['cell'];
	}

	foreach($valueList as $values){

		$table->addRow();
		$table->addCell(2000, array(
			'rowMerge'=>'continue',
		));

		foreach($values as $index=>$value){

			$table->addCell(CELL_LABEL_TITLE_WIDTH, $cellMerges[$index]?$cellValueMergeStartStyle:$styleCell)->addText($value, $valueFontStyle);
			mergeCell($cellMerges[$index]);
		}
	}

}


// Save File
$objWriter = PHPWord_IOFactory::createWriter($PHPWord, 'Word2007');
$objWriter->save('table.docx');



/**
 * 单元格横向合并
 * @param  [type] $count [description]
 * @return [type]        [description]
 */
function mergeCell($count){
	global $table;
	// echo $count.'<br />';
	for($i=0;$i<$count;$i++){
		$table->addCell(2000, array(
			'cellMerge'=>'continue',
		));
	}
}
