<?php
/**
  * Test file for the multiple Excel worksheet feature
  *
  * @author Christian Lescuyer <christian@goelette.net>
  * @copyright Goélette
  */

require '../../simpletest/autorun.php';
require '../php-export-data.class.php';

class TestOfPedExcel extends UnitTestCase {

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

