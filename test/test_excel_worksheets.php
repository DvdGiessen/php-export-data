<?php
/**
  * Test file for the multiple Excel worksheet feature
  *
  * @author Christian Lescuyer <christian@goelette.net>
  * @copyright Goélette
  */

require '../../simpletest/autorun.php';
require '../php-export-data.class.php';

class TestOfPedExportDataExcel extends UnitTestCase {

  function test_Ped_File_CreatesFile()
  {
    $filename = 'one.xml';
    @unlink($filename);
    $this->assertFalse(file_exists($filename));
    $excel = new ExportDataExcel('file');
    $excel->filename = $filename;
    $excel->initialize();
    $excel->finalize();
    $this->assertTrue(file_exists($filename));
  }

  function test_Ped_String_CreatesValidSpreadsheetML()
  {
    $excel = new ExportDataExcel('string');
    $excel->initialize();
    $excel->finalize();
    $xml = $excel->getString();
    $dom = new DOMDocument();
    libxml_use_internal_errors(true);
    $dom->loadXML($xml);
    $dom->schemaValidate('SpreadsheetML/excelss.xsd');
    $this->assertTrue($this->check_libxml_errors());
  }

  function test_Ped_String_CreatesOneWorksheet()
  {
    $excel = new ExportDataExcel('string');
    $excel->initialize();
    $excel->finalize();
    $xml = $excel->getString();
    $dom = new DOMDocument();
    libxml_use_internal_errors(true);
    $dom->loadXML($xml);
    $sheets = $dom->getElementsByTagName('Worksheet');
    $this->assertEqual(1, $sheets->length);
  }

  function test_Ped_String_CreatesMultipleWorksheets()
  {
    $excel = new ExportDataExcel('string');
    $excel->initialize();
    $excel->newSheet('Sheet 2');
    $excel->newSheet('Sheet 3');
    $excel->finalize();
    $xml = $excel->getString();
    $dom = new DOMDocument();
    libxml_use_internal_errors(true);
    $dom->loadXML($xml);
    $sheets = $dom->getElementsByTagName('Worksheet');
    $this->assertEqual(3, $sheets->length);
  }

  function test_Ped_String_CreatesNamedWorksheets()
  {
    $excel = new ExportDataExcel('string');
    $excel->title = 'Test sheet 1';
    $excel->initialize();
    $excel->newSheet('Test sheet 2');
    $excel->newSheet('Test sheet 3');
    $excel->finalize();
    $xml = $excel->getString();
    $dom = new DOMDocument();
    libxml_use_internal_errors(true);
    $dom->loadXML($xml);
    $sheets = $dom->getElementsByTagName('Worksheet');
    $this->assertEqual('Test sheet 1', $sheets->item(0)->attributes->item(0)->nodeValue);
    $this->assertEqual('Test sheet 2', $sheets->item(1)->attributes->item(0)->nodeValue);
    $this->assertEqual('Test sheet 3', $sheets->item(2)->attributes->item(0)->nodeValue);
  }

  function test_Ped_String_WritesInEachWorksheet()
  {
    $excel = new ExportDataExcel('string');
    $excel->title = 'Test sheet 1';
    $excel->initialize();
    $excel->addRow(array('data 1'));
    $excel->newSheet('Test sheet 2');
    $excel->addRow(array('data 2'));
    $excel->newSheet('Test sheet 3');
    $excel->addRow(array('data 3'));
    $excel->finalize();
    $xml = $excel->getString();
    $dom = new DOMDocument();
    libxml_use_internal_errors(true);
    $dom->loadXML($xml);
    $sheets = $dom->getElementsByTagName('Worksheet');
    $this->assertEqual('data 1', trim($sheets->item(0)->nodeValue));
    $this->assertEqual('data 2', trim($sheets->item(1)->nodeValue));
    $this->assertEqual('data 3', trim($sheets->item(2)->nodeValue));
    // Check 'data 3' does NOT appear in the other sheets
    $this->assertNotEqual('data 3', trim($sheets->item(0)->nodeValue));
    $this->assertNotEqual('data 3', trim($sheets->item(1)->nodeValue));
  }

  function test_Ped_String_SetsMultilineStyle()
  {
    $excel = new ExportDataExcel('string');
    $excel->initialize();
    $excel->addRow(array("line 1\nline 2"));
    $excel->finalize();
    $xml = $excel->getString();
    $dom = new DOMDocument();
    libxml_use_internal_errors(true);
    $dom->loadXML($xml);
    $cells = $dom->getElementsByTagName('Cell');
    $this->assertTrue($cells->item(0)->hasAttributes());
    $this->assertEqual('ss:StyleID', $cells->item(0)->attributes->item(0)->nodeName);
    $this->assertEqual('sMultiLine', $cells->item(0)->attributes->item(0)->nodeValue);
    $data = $dom->getElementsByTagName('Data');
    $this->assertEqual("line 1\nline 2", trim($cells->item(0)->nodeValue));
  }

  private function check_libxml_errors()
  {
    $errors = libxml_get_errors();
    $ok = true;
    foreach ($errors as $error) {
      switch ($error->level) {
        case LIBXML_ERR_WARNING:
          break;
        case LIBXML_ERR_ERROR:
          print "Error $error->code: $error->message";
          $ok = false;
          break;
        case LIBXML_ERR_FATAL:
          print "Fatal Error $error->code: $error->message";
          $ok = false;
          break;
      }
    }
    libxml_clear_errors();
    return $ok;
  }
}

