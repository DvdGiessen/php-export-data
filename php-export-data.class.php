<?php
// php-export-data by Eli Dickinson, http://github.com/elidickinson/php-export-data

/**
 * ExportData is the base class for exporters to specific file formats. See other
 * classes below.
 */
abstract class ExportData {
	protected $exportTo; // Set in constructor to one of 'browser', 'file', 'string'
	protected $stringData; // stringData so far, used if export string mode
	protected $tempFile; // handle to temp file (for export file mode)
	protected $tempFilename; // temp file name and path (for export file mode)
    protected $thousandsSeparator; // handle thousands separator in excel export

	public $filename; // file mode: the output file name; browser mode: file name for download; string mode: not used

	public function __construct($exportTo = "browser", $filename = "exportdata", $thousandsSeparator = false) {
		if(!in_array($exportTo, array('browser','file','string') )) {
			throw new Exception("$exportTo is not a valid ExportData export type");
		}
		$this->exportTo = $exportTo;
		$this->filename = $filename;
        $this->thousandsSeparator = $thousandsSeparator;
	}
	
	public function initialize() {
		
		switch($this->exportTo) {
			case 'browser':
				$this->sendHttpHeaders();
				break;
			case 'string':
				$this->stringData = '';
				break;
			case 'file':
				$this->tempFilename = tempnam(sys_get_temp_dir(), 'exportdata');
				$this->tempFile = fopen($this->tempFilename, "w");
				break;
		}
		
		$this->write($this->generateHeader());
	}
	
	public function addRow($row) {
		$this->write($this->generateRow($row));
	}
	
	public function finalize() {
		
		$this->write($this->generateFooter());
		
		switch($this->exportTo) {
			case 'browser':
				flush();
				break;
			case 'string':
				// do nothing
				break;
			case 'file':
				// close temp file and move it to correct location
				fclose($this->tempFile);
				rename($this->tempFilename, $this->filename);
				break;
		}
	}
	
	public function getString() {
		return $this->stringData;
	}
	
	abstract public function sendHttpHeaders();
	
	protected function write($data) {
		switch($this->exportTo) {
			case 'browser':
				echo $data;
				break;
			case 'string':
				$this->stringData .= $data;
				break;
			case 'file':
				fwrite($this->tempFile, $data);
				break;
		}
	}
	
	protected function generateHeader() {
		// can be overridden by subclass to return any data that goes at the top of the exported file
	}
	
	protected function generateFooter() {
		// can be overridden by subclass to return any data that goes at the bottom of the exported file		
	}
	
	// In subclasses generateRow will take $row array and return string of it formatted for export type
	abstract protected function generateRow($row);
	
}

/**
 * ExportDataTSV - Exports to TSV (tab separated value) format.
 */
class ExportDataTSV extends ExportData {
	
	function generateRow($row) {
		foreach ($row as $key => $value) {
			// Escape inner quotes and wrap all contents in new quotes.
			// Note that we are using \" to escape double quote not ""
			$row[$key] = '"'. str_replace('"', '\"', $value) .'"';
		}
		return implode("\t", $row) . "\n";
	}
	
	function sendHttpHeaders() {
		header("Content-type: text/tab-separated-values");
    header("Content-Disposition: attachment; filename=".basename($this->filename));
	}
}

/**
 * ExportDataCSV - Exports to CSV (comma separated value) format.
 */
class ExportDataCSV extends ExportData {
	
	function generateRow($row) {
		foreach ($row as $key => $value) {
			// Escape inner quotes and wrap all contents in new quotes.
			// Note that we are using \" to escape double quote not ""
			$row[$key] = '"'. str_replace('"', '\"', $value) .'"';
		}
		return implode(",", $row) . "\n";
	}
	
	function sendHttpHeaders() {
		header("Content-type: text/csv");
		header("Content-Disposition: attachment; filename=".basename($this->filename));
	}
}


/**
 * ExportDataExcel exports data into an XML format  (spreadsheetML) that can be 
 * read by MS Excel 2003 and newer as well as OpenOffice
 * 
 * Creates a workbook with a single worksheet (title specified by
 * $title).
 * 
 * Note that using .XML is the "correct" file extension for these files, but it
 * generally isn't associated with Excel. Using .XLS is tempting, but Excel 2007 will
 * throw a scary warning that the extension doesn't match the file type.
 * 
 * Based on Excel XML code from Excel_XML (http://github.com/oliverschwarz/php-excel)
 *  by Oliver Schwarz
 */
class ExportDataExcel extends ExportData {
	
	public $encoding = 'UTF-8'; // encoding type to specify in file. 
	// Note that you're on your own for making sure your data is actually encoded to this encoding
	
	public $title = 'Sheet1'; // title for Worksheet 
	
	function generateHeader() {
		
		// Workbook header
    $output = '<?xml version="1.0" encoding="' . $this->encoding . '"?>' . "\n";
    $output.= '<?mso-application progid="Excel.Sheet"?>' . "\n"; // Get the .xml file to open in Excel as a default
		$output.= '<Workbook xmlns="urn:schemas-microsoft-com:office:spreadsheet" xmlns:x="urn:schemas-microsoft-com:office:excel" xmlns:ss="urn:schemas-microsoft-com:office:spreadsheet" xmlns:html="http://www.w3.org/TR/REC-html40">' . "\n";
		
		// Set up styles
		$output .= '<Styles>' . "\n";
		$output .= '<Style ss:ID="sDT"><NumberFormat ss:Format="Short Date"/></Style>' . "\n";
        $output .= '<Style ss:ID="s63"><NumberFormat ss:Format="Standard"/></Style>' . "\n";
		$output .= '<Style ss:ID="sMultiLine"><Alignment ss:Vertical="Bottom" ss:WrapText="1"/></Style>' . "\n";
		$output .= '</Styles>' . "\n";
		
		// worksheet header
                $output .= $this->generateWorksheetHeader($this->title);
		
		return $output;
        }

        // 2014-01-02 added by GDX
        function generateWorksheetHeader($title)
        {
            return sprintf("<Worksheet ss:Name=\"%s\">\n    <Table>\n", htmlentities($title));
        }

        function generateWorksheetFooter()
        {
            return "    </Table>\n</Worksheet>\n";
        }

        public function addWorksheet($title)
        {
            $output = '';
            $output .= $this->generateWorksheetFooter();
            $output .= $this->generateWorksheetHeader($title);
            $this->write($output);
        }
	
	function generateFooter() {
		$output = '';
		
		// worksheet footer
		$output .= $this->generateWorksheetFooter();
		
		// workbook footer
		$output .= '</Workbook>';
		
		return $output;
	}
  function newSheet($name)
  {
    $this->write("    </Table>\n</Worksheet>\n");
    $this->write(sprintf("<Worksheet ss:Name=\"%s\">\n    <Table>\n", htmlentities($name)));
  }
	function generateRow($row) {
		$output = '';
		$output .= "        <Row>\n";
		foreach ($row as $k => $v) {
			$output .= $this->generateCell($v);
		}
		$output .= "        </Row>\n";
		return $output;
	}
	
	protected function generateCell($item) {
		$output = '';
		$style = '';
		
		// Tell Excel to treat as a number. Note that Excel only stores roughly 15 digits, so keep 
		// as text if number is longer than that.
		if(preg_match("/^-?\d+(?:[.,]\d+)?$/",$item) && (strlen($item) < 15)) {
			$item = str_replace(',', '.', $item);
			$type = 'Number';
            $style = $this->thousandsSeparator ? 's63' : null; // defined in header; tells excel to format number with thousands separator
		}
		// Sniff for valid dates; should look something like 2010-07-14 or 7/14/2010 etc. Can
		// also have an optional time after the date.
		//
		// Note we want to be very strict in what we consider a date. There is the possibility
		// of really screwing up the data if we try to reformat a string that was not actually 
		// intended to represent a date.
		elseif (preg_match('_^
				# dates: 2010-07-14 or 7/14/2010
				(?:\d{1,2}|\d{4})[/-]\d{1,2}[/-](?:\d{1,2}|\d{4})
				# optional time
				(?:[T\s]+ # separated by space or T
				[\d:,]+ # hours:minutes
				)?
				$_x', $item) &&
					($timestamp = strtotime($item)) &&
					($timestamp > 0) &&
					($timestamp < strtotime('+500 years'))) {
			$type = 'DateTime';
			$item = strftime("%Y-%m-%dT%H:%M:%S",$timestamp);
			$style = 'sDT'; // defined in header; tells excel to format date for display
		}
		else {
			$type = 'String';
			if (strpos($item, "\n") !== false) {
			  $style = 'sMultiLine';
		}
		}
				
		$item = str_replace('&#039;', '&apos;', htmlspecialchars($item, ENT_QUOTES));
		$item = str_replace("\n", '&#10;', $item);
		$output .= "            ";
		$output .= $style ? "<Cell ss:StyleID=\"$style\">" : "<Cell>";
		$output .= sprintf("<Data ss:Type=\"%s\">%s</Data>", $type, $item);
		$output .= "</Cell>\n";
		
		return $output;
	}
	
	function sendHttpHeaders() {
		header("Content-Type: application/vnd.ms-excel; charset=" . $this->encoding);
		header("Content-Disposition: inline; filename=\"" . basename($this->filename) . "\"");
	}
	
}
