<?php

require "../php-export-data.class.php";

$excel = new ExportDataExcel('browser');
$excel->filename = "test.xml";

$data = array(
	array(1,2,3),
	array(0.1, 2.15),
	array('0,1', '2,15', '.5', '2.', '1,2,3'),
	array("asdf","jkl","semi"), 
	array("1273623874628374634876","=asdf","10-10"),
	array(
		// iso dates: https://en.wikipedia.org/wiki/ISO_8601
		"2010-01-02", "2010-01-02 10:00",
		"2010-01-02T10:00",
		"2010-01-02T10:00:00",
		"2010-01-02T10:00:00,000",
		"2010-01-02T10:00:00,000",
		// US dates
		"02/01/2010 00:00", "02/01/2010", "2/1/2010",
	),
);

$excel->initialize();
foreach($data as $row) {
	$excel->addRow($row);
}
$excel->finalize();
