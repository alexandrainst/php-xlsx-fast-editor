<?php

/**
 * PHP library to make basic but fast read & write operations on existing Excel workbooks.
 * Originally written by [Alexandre Alapetite](https://github.com/Alkarex) for the [Alexandra Institute](https://alexandra.dk/), 2023.
 *
 * @author Alexandre Alapetite <alexandre.alapetite@alexandra.dk>
 * @category PHP
 * @license https://gnu.org/licenses/lgpl.html GNU LGPL
 * @link https://github.com/alexandrainst/php-xlsx-fast-editor
 * @package XlsxFastEditor
 */

declare(strict_types=1);

namespace alexandrainst\XlsxFastEditor;

/**
 * Main class to fast edit an existing XLSX/XLSM document (Microsoft Excel 2007+, Office Open XML Workbook)
 * using low-level ZIP and XML manipulation.
 */
final class XlsxFastEditor
{
	/** @internal */
	public const _OXML_NAMESPACE = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main';

	private const CALC_CHAIN_CACHE_PATH = 'xl/calcChain.xml';
	private const SHARED_STRINGS_PATH = 'xl/sharedStrings.xml';
	private const WORKBOOK_PATH = 'xl/workbook.xml';
	private const WORKBOOK_RELS_PATH = 'xl/_rels/workbook.xml.rels';

	private \ZipArchive $zip;

	/**
	 * Cache of the XPath instances associated to the DOM of the XML documents.
	 * @var array<string,\DOMXPath>
	 */
	private array $documents = [];

	/**
	 * Track which documents have pending changes.
	 * @var array<string,bool>
	 */
	private array $modified = [];

	/**
	 * Whether the calcChain must be cleared on save.
	 */
	private bool $mustClearCalcChain = false;

	/**
	 * @throws XlsxFastEditorZipException
	 */
	public function __construct(string $filename)
	{
		$this->zip = new \ZipArchive();
		$zipCode = $this->zip->open($filename, \ZipArchive::CREATE);
		if ($zipCode !== true) {
			throw new XlsxFastEditorZipException("Cannot open workbook {$filename}!", $zipCode);
		}
	}

	/**
	 * Mark a document fragment as modified.
	 * @param string $path The path of the document inside the ZIP document.
	 */
	private function touchPath(string $path): void
	{
		$this->modified[$path] = true;
	}

	/**
	 * Mark a document fragment as modified.
	 * @internal
	 * @param int $sheetNumber Worksheet number (base 1)
	 */
	public function _touchWorksheet(int $sheetNumber): void
	{
		$path = self::getWorksheetPath($sheetNumber);
		$this->touchPath($path);
	}

	/**
	 * Will clear the calcChain on save.
	 * @internal
	 */
	public function _clearCalcChain(): void
	{
		$this->mustClearCalcChain = true;
	}

	/**
	 * Close the underlying document archive.
	 * Note: changes need to be explicitly saved before (see `XlsxFastEditor::save()`)
	 * Note: the object should not be used anymore afterwards.
	 * @throws XlsxFastEditorZipException
	 */
	public function close(): void
	{
		$this->documents = [];
		if (!$this->zip->close()) {
			throw new XlsxFastEditorZipException("Error closing the underlying document!");
		}
	}

	/**
	 * Saves the modified document fragments.
	 * @param bool $close Automatically close the underlying document archive (see `XlsxFastEditor::close()`)
	 * @throws XlsxFastEditorZipException
	 * @throws XlsxFastEditorXmlException
	 */
	public function save(bool $close = true): void
	{
		if ($this->mustClearCalcChain) {
			// Removes calcChain.xml, which contains some cache for formulas,
			// as the cache might become invalid after writing in cells containing formulas.
			$this->zip->deleteName(self::CALC_CHAIN_CACHE_PATH);
			$this->mustClearCalcChain = false;
		}
		foreach ($this->modified as $name => $pending) {
			if (!$pending || !isset($this->documents[$name])) {
				continue;
			}
			$xpath = $this->documents[$name];
			if (!$this->zip->deleteName($name)) {
				throw new XlsxFastEditorZipException("Error deleting old fragment {$name}!");
			}
			$dom = $xpath->document;
			$xml = $dom->saveXML();
			if ($xml === false) {
				throw new XlsxFastEditorXmlException("Error saving changes {$name}!");
			}
			if (!$this->zip->addFromString($name, $xml)) {
				throw new XlsxFastEditorZipException("Error saving new fragment {$name}!");
			}
		}
		$this->modified = [];

		if ($close) {
			$this->close();
		}
	}

	/**
	 * Extracts a worksheet from the internal ZIP document,
	 * parse the XML, open the DOM, and
	 * returns an XPath instance associated to the DOM at the given XML path.
	 * The XPath instance is then cached.
	 * @param string $path The path of the document inside the ZIP document.
	 * @throws XlsxFastEditorFileFormatException
	 * @throws XlsxFastEditorXmlException
	 */
	private function getXPathFromPath(string $path): \DOMXPath
	{
		if (isset($this->documents[$path])) {
			return $this->documents[$path];
		}

		$xml = $this->zip->getFromName($path);
		if ($xml === false) {
			throw new XlsxFastEditorFileFormatException("Missing XML fragment {$path}!");
		}

		$dom = new \DOMDocument();
		if ($dom->loadXML($xml, LIBXML_NOERROR | LIBXML_NONET | LIBXML_NOWARNING) === false) {
			throw new XlsxFastEditorXmlException("Error reading XML fragment {$path}!");
		}

		$xpath = new \DOMXPath($dom);
		$xpath->registerNamespace('o', self::_OXML_NAMESPACE);

		$this->documents[$path] = $xpath;
		return $xpath;
	}

	/**
	 * Returns a DOM document of the given XML path.
	 * @param string $path The path of the document inside the ZIP document.
	 * @throws XlsxFastEditorFileFormatException
	 * @throws XlsxFastEditorXmlException
	 */
	private function getDomFromPath(string $path): \DOMDocument
	{
		return $this->getXPathFromPath($path)->document;
	}

	/**
	 * Excel can either use a base date from year 1900 (Microsoft Windows) or from year 1904 (old Apple MacOS).
	 * https://support.microsoft.com/en-us/office/date-systems-in-excel-e7fe7167-48a9-4b96-bb53-5612a800b487
	 * @phpstan-return 1900|1904
	 * @return int `1900` or `1904`
	 * @throws XlsxFastEditorFileFormatException
	 * @throws XlsxFastEditorXmlException
	 */
	public function getWorkbookDateSystem(): int
	{
		static $baseYear = 0;
		if ($baseYear == 0) {
			$xpath = $this->getXPathFromPath(self::WORKBOOK_PATH);
			$date1904 = $xpath->evaluate('normalize-space(/o:workbook/o:workbookPr/@date1904)');
			if (is_string($date1904) && in_array(strtolower(trim($date1904)), ['true', '1'], true)) {
				$baseYear = 1904;
			} else {
				$baseYear = 1900;
			}
		}
		return $baseYear;
	}

	/**
	 * Convert an internal Excel float representation of a date to a standard `DateTime`.
	 * @param int $workbookDateSystem {@see XlsxFastEditor::getWorkbookDateSystem()}
	 * @phpstan-param 1900|1904 $workbookDateSystem
	 * @internal
	 * @throws \InvalidArgumentException
	 */
	public static function excelDateToDateTime(float $excelDateTime, int $workbookDateSystem = 1900): \DateTimeImmutable
	{
		static $baseDate1900 = null;
		static $baseDate1904 = null;
		if ($workbookDateSystem === 1900) {
			if ($excelDateTime < 1) {
				// Make cells with only time (no date) to start on 1900-01-01
				$excelDateTime++;
			}
			if ($excelDateTime < 60) {
				// https://learn.microsoft.com/en-us/office/troubleshoot/excel/wrongly-assumes-1900-is-leap-year
				$excelDateTime++;
			}
			// 1 January 1900 as serial number 1 in the 1900 Date System, accounting for leap year problem
			if ($baseDate1900 === null) {
				$baseDate1900 = new \DateTimeImmutable('1899-12-30');
			}
			$excelBaseDate = $baseDate1900;
		} elseif ($workbookDateSystem === 1904) {
			// 1 January 1904 as serial number 0 in the 1904 Date System
			if ($baseDate1904 === null) {
				$baseDate1904 = new \DateTimeImmutable('1904-01-01');
			}
			$excelBaseDate = $baseDate1904;
		} else {
			throw new \InvalidArgumentException('Invalid Excel workbook date system! Supported values: 1900, 1904');
		}

		$daysOffset = floor($excelDateTime);
		$iso8601 = "P{$daysOffset}D";

		$timeFraction = $excelDateTime - $daysOffset;
		if ($timeFraction > 0) {
			// Convert days to seconds with no more than milliseconds precision
			$seconds = floor($timeFraction * 86400000) / 1000;
			$iso8601 .= "T{$seconds}S";
		}

		try {
			return $excelBaseDate->add(new \DateInterval($iso8601));
		} catch (\Exception $ex) {
			throw new \InvalidArgumentException('Invalid date!', $ex->getCode(), $ex);
		}
	}

	/**
	 * Count the number of worksheets in the workbook.
	 * @throws XlsxFastEditorFileFormatException
	 * @throws XlsxFastEditorXmlException
	 */
	public function getWorksheetCount(): int
	{
		$xpath = $this->getXPathFromPath(self::WORKBOOK_PATH);
		$count = $xpath->evaluate('count(/o:workbook/o:sheets/o:sheet)');
		return is_numeric($count) ? (int)$count : 0;
	}

	/**
	 * Get a worksheet number (ID) from its name (base 1).
	 * @param string $sheetName The name of the worksheet to look up.
	 * @return int The worksheet ID, or `-1` if not found.
	 * @throws XlsxFastEditorFileFormatException
	 * @throws XlsxFastEditorXmlException
	 */
	public function getWorksheetNumber(string $sheetName): int
	{
		$xpath = $this->getXPathFromPath(self::WORKBOOK_PATH);
		$xpath->registerNamespace('r', 'http://schemas.openxmlformats.org/officeDocument/2006/relationships');
		$rId = $xpath->evaluate("normalize-space(/o:workbook/o:sheets/o:sheet[@name='$sheetName'][1]/@r:id)");
		if (!is_string($rId) || $rId === '') {
			return -1;
		}

		try {
			$xpath = $this->getXPathFromPath(self::WORKBOOK_RELS_PATH);
			$xpath->registerNamespace('pr', 'http://schemas.openxmlformats.org/package/2006/relationships');
			$target = $xpath->evaluate("normalize-space(/pr:Relationships/pr:Relationship[@Id='$rId'][1]/@Target)");
			if (is_string($target) && preg_match('/(\d+)/i', $target, $matches)) {
				return (int)$matches[1];
			}
		} catch (XlsxFastEditorFileFormatException $ex) {	// phpcs:ignore Generic.CodeAnalysis.EmptyStatement.DetectedCatch
		}

		if (preg_match('/(\d+)/i', $rId, $matches)) {
			return (int)$matches[1];
		}
		return -1;
	}

	/**
	 * Get a worksheet name from its number (ID).
	 * @param int $sheetNumber The number of the worksheet to look up.
	 * @return string|null The worksheet name, or `null` if not found.
	 * @throws XlsxFastEditorFileFormatException
	 * @throws XlsxFastEditorXmlException
	 */
	public function getWorksheetName(int $sheetNumber): ?string
	{
		$xpath = $this->getXPathFromPath(self::WORKBOOK_PATH);
		$sheetName = $xpath->evaluate("normalize-space(/o:workbook/o:sheets/o:sheet[$sheetNumber]/@name)");
		return is_string($sheetName) ? $sheetName : null;
	}

	private static function getWorksheetPath(int $sheetNumber): string
	{
		return "xl/worksheets/sheet{$sheetNumber}.xml";
	}

	/**
	 * Defines the *Full calculation on load* policy for the specified worksheet.
	 * @param int $sheetNumber Worksheet number (base 1)
	 * @throws XlsxFastEditorFileFormatException
	 * @throws XlsxFastEditorXmlException
	 */
	public function setFullCalcOnLoad(int $sheetNumber, bool $value): void
	{
		$this->mustClearCalcChain = true;
		$dom = $this->getDomFromPath(self::getWorksheetPath($sheetNumber));
		$sheetCalcPr = null;
		$sheetCalcPrs = $dom->getElementsByTagName('sheetCalcPr');
		if ($sheetCalcPrs->length > 0) {
			$sheetCalcPr = $sheetCalcPrs[0];
		} else {
			$sheetDatas = $dom->getElementsByTagName('sheetData');
			if ($sheetDatas->length > 0) {
				$sheetData = $sheetDatas[0];
				if ($sheetData instanceof \DOMElement) {
					try {
						$sheetCalcPr = $dom->createElement('sheetCalcPr');
					} catch (\DOMException $dex) {
						throw new XlsxFastEditorXmlException("Error creating XML fragment for setFullCalcOnLoad!", $dex->code, $dex);
					}
					if ($sheetCalcPr !== false && $sheetData->parentNode !== null) {
						$sheetData->parentNode->insertBefore($sheetCalcPr, $sheetData->nextSibling);
					}
				}
			}
		}
		if ($sheetCalcPr instanceof \DOMElement) {
			$sheetCalcPr->setAttribute('fullCalcOnLoad', $value ? 'true' : 'false');
			$this->_touchWorksheet($sheetNumber);
		}
	}

	/**
	 * Get the *Full calculation on load* policy for the specified worksheet.
	 * @param int $sheetNumber Worksheet number (base 1)
	 * @throws XlsxFastEditorFileFormatException
	 * @throws XlsxFastEditorXmlException
	 */
	public function getFullCalcOnLoad(int $sheetNumber): ?bool
	{
		$dom = $this->getDomFromPath(self::getWorksheetPath($sheetNumber));
		$sheetCalcPrs = $dom->getElementsByTagName('sheetCalcPr');
		if ($sheetCalcPrs->length > 0) {
			$sheetCalcPr = $sheetCalcPrs[0];
			if ($sheetCalcPr instanceof \DOMElement) {
				$fullCalcOnLoad = $sheetCalcPr->getAttribute('fullCalcOnLoad');
				if ($fullCalcOnLoad !== '') {
					$fullCalcOnLoad = strtolower(trim($fullCalcOnLoad));
					return in_array($fullCalcOnLoad, ['true', '1'], true);
				}
			}
		}
		return null;
	}

	/**
	 * Get the row of the given number in the given worksheet.
	 * @param int $sheetNumber Worksheet number (base 1)
	 * @param int $rowNumber Number (ID) of the row (base 1). Warning: this is not an index (not all rows necessarily exist in a sequence)
	 * @param int $accessMode To control the behaviour when the cell does not exist:
	 * set to `XlsxFastEditor::ACCESS_MODE_NULL` to return `null` (default),
	 * set to `XlsxFastEditor::ACCESS_MODE_EXCEPTION` to raise an `XlsxFastEditorInputException` exception,
	 * set to `XlsxFastEditor::ACCESS_MODE_AUTOCREATE` to auto-create the cell.
	 * @return XlsxFastEditorRow|null A row, potentially `null` if the row does not exist and `$accessMode` is set to `XlsxFastEditor::ACCESS_MODE_NULL`
	 * @phpstan-return ($accessMode is XlsxFastEditor::ACCESS_MODE_NULL ? XlsxFastEditorRow|null : XlsxFastEditorRow)
	 * @throws XlsxFastEditorFileFormatException
	 * @throws XlsxFastEditorInputException optionally if the corresponding cell does not exist, depending on choice of `$accessMode`
	 * @throws XlsxFastEditorXmlException
	 */
	public function getRow(int $sheetNumber, int $rowNumber, int $accessMode = XlsxFastEditor::ACCESS_MODE_NULL): ?XlsxFastEditorRow
	{
		$xpath = $this->getXPathFromPath(self::getWorksheetPath($sheetNumber));
		$rows = $xpath->query("/o:worksheet/o:sheetData/o:row[@r='{$rowNumber}'][1]");
		$row = null;
		if ($rows !== false && $rows->length > 0) {
			$row = $rows[0];
			if (!($row instanceof \DOMElement)) {
				throw new XlsxFastEditorXmlException("Error querying XML fragment for row {$sheetNumber} of worksheet {$sheetNumber}!");
			}
		}

		if ($row === null) {
			// The <row> was not found

			switch ($accessMode) {
				case XlsxFastEditor::ACCESS_MODE_EXCEPTION:
					throw new XlsxFastEditorInputException("Row {$sheetNumber}/{$rowNumber} not found!");
				case XlsxFastEditor::ACCESS_MODE_AUTOCREATE:
					$sheetDatas = $xpath->document->getElementsByTagName('sheetData');
					if ($sheetDatas->length === 0) {
						throw new XlsxFastEditorXmlException("Cannot find sheetData for worksheet {$sheetNumber}!");
					}
					$sheetData = $sheetDatas[0];
					if (!($sheetData instanceof \DOMElement)) {
						throw new XlsxFastEditorXmlException("Error querying XML fragment for worksheet {$sheetNumber}!");
					}
					try {
						$row = $xpath->document->createElement('row');
					} catch (\DOMException $dex) {
						throw new XlsxFastEditorXmlException("Error creating row {$sheetNumber}/{$rowNumber}!", $dex->code, $dex);
					}
					if ($row === false) {
						throw new XlsxFastEditorXmlException("Error creating row {$sheetNumber}/{$rowNumber}!");
					}
					$row->setAttribute('r', (string)$rowNumber);

					// Excel expects the lines to be sorted
					$sibling = $sheetData->firstElementChild;
					while ($sibling !== null && (int)$sibling->getAttribute('r') < $rowNumber) {
						$sibling = $sibling->nextElementSibling;
					}
					$sheetData->insertBefore($row, $sibling);
					break;
				default:
				case XlsxFastEditor::ACCESS_MODE_NULL:
					return null;
			}
		}

		return new XlsxFastEditorRow($this, $sheetNumber, $row);
	}

	/**
	 * Get the first existing row of the worksheet.
	 * @param int $sheetNumber Worksheet number (base 1)
	 * @return XlsxFastEditorRow|null The first row of the worksheet if there is any row, `null` otherwise.
	 * @throws XlsxFastEditorFileFormatException
	 * @throws XlsxFastEditorXmlException
	 */
	public function getFirstRow(int $sheetNumber): ?XlsxFastEditorRow
	{
		$xpath = $this->getXPathFromPath(self::getWorksheetPath($sheetNumber));
		$rs = $xpath->query("/o:worksheet/o:sheetData/o:row[position() = 1]");
		if ($rs !== false && $rs->length > 0) {
			$r = $rs[0];
			if (!($r instanceof \DOMElement)) {
				throw new XlsxFastEditorXmlException("Error querying XML fragment for row {$sheetNumber} of worksheet {$sheetNumber}!");
			}
			return new XlsxFastEditorRow($this, $sheetNumber, $r);
		}
		return null;
	}

	/**
	 * Get the last existing row of the worksheet.
	 * @param int $sheetNumber Worksheet number (base 1)
	 * @return XlsxFastEditorRow|null The last row of the worksheet if there is any row, `null` otherwise.
	 * @throws XlsxFastEditorFileFormatException
	 * @throws XlsxFastEditorXmlException
	 */
	public function getLastRow(int $sheetNumber): ?XlsxFastEditorRow
	{
		$xpath = $this->getXPathFromPath(self::getWorksheetPath($sheetNumber));
		$rs = $xpath->query("/o:worksheet/o:sheetData/o:row[position() = last()]");
		if ($rs !== false && $rs->length > 0) {
			$r = $rs[0];
			if (!($r instanceof \DOMElement)) {
				throw new XlsxFastEditorXmlException("Error querying XML fragment for row {$sheetNumber} of worksheet {$sheetNumber}!");
			}
			return new XlsxFastEditorRow($this, $sheetNumber, $r);
		}
		return null;
	}

	/**
	 * Delete the specified row of the specified worksheet.
	 * @param int $sheetNumber Worksheet number (base 1)
	 * @return bool `true` if the deletion succeeds, `false` otherwise.
	 * @throws XlsxFastEditorFileFormatException
	 * @throws XlsxFastEditorXmlException
	 */
	public function deleteRow(int $sheetNumber, int $rowNumber): bool
	{
		$xpath = $this->getXPathFromPath(self::getWorksheetPath($sheetNumber));
		$rs = $xpath->query("/o:worksheet/o:sheetData/o:row[@r='{$rowNumber}'][1]");
		if ($rs !== false && $rs->length > 0) {
			$r = $rs[0];
			if (!($r instanceof \DOMElement) || $r->parentNode === null) {
				throw new XlsxFastEditorXmlException("Error querying XML fragment for row {$sheetNumber} of worksheet {$sheetNumber}!");
			}
			return $r->parentNode->removeChild($r) != false;
		}
		return false;
	}

	/**
	 * To iterate over the all the rows of a given worksheet.
	 * @return \Traversable<XlsxFastEditorRow>
	 * @throws XlsxFastEditorFileFormatException
	 * @throws XlsxFastEditorXmlException
	 */
	public function rowsIterator(int $sheetNumber): \Traversable
	{
		$xpath = $this->getXPathFromPath(self::getWorksheetPath($sheetNumber));
		$rs = $xpath->query("/o:worksheet/o:sheetData/o:row");
		if ($rs !== false) {
			for ($i = 0; $i < $rs->length; $i++) {
				$r = $rs[$i];
				if (!($r instanceof \DOMElement)) {
					throw new XlsxFastEditorXmlException("Error querying XML fragment for row {$sheetNumber}!");
				}
				yield new XlsxFastEditorRow($this, $sheetNumber, $r);
			}
		}
	}

	/**
	 * Produce an array from a worksheet, indexed by column name (like `'AB'`) first, then line (like `12`).
	 * Only the existing lines and cells are included.
	 * @return array<string,array<int,null|string>> An array that can be accessed like `$array['AB'][12]`
	 * @throws XlsxFastEditorFileFormatException
	 * @throws XlsxFastEditorXmlException
	 */
	public function readArray(int $sheetNumber): array
	{
		$table = [];
		foreach ($this->rowsIterator($sheetNumber) as $row) {
			foreach ($row->cellsIterator() as $cell) {
				$table[$cell->column()][$row->number()] = $cell->readString();
			}
		}
		return $table;
	}

	/**
	 * Produce an array from a worksheet, indexed by column header (like `'columnName'`) first, then line (like `12`),
	 * having the column header defined in the first existing line of the spreadsheet.
	 * Only the existing lines and cells are included.
	 * @return array<string,array<int,null|string>> An array that can be accessed like `$array['columnName'][12]`
	 * @throws XlsxFastEditorFileFormatException
	 * @throws XlsxFastEditorXmlException
	 */
	public function readArrayWithHeaders(int $sheetNumber): array
	{
		$table = [];
		$headers = [];
		$firstRow = true;
		foreach ($this->rowsIterator($sheetNumber) as $row) {
			if ($firstRow) {
				foreach ($row->cellsIterator() as $cell) {
					$headers[$cell->column()] = $cell->readString();
				}
				$firstRow = false;
				continue;
			}
			foreach ($row->cellsIterator() as $cell) {
				$header = $headers[$cell->column()] ?? $cell->column();
				$table[$header][$row->number()] = $cell->readString();
			}
		}
		return $table;
	}

	/**
	 * Sort cells (such as `'B3'`, `'AA23'`) on column first (such as `'B'`, `'AA'`) and then line (such as `3`, `23`).
	 * @param $ref1 A cell reference such as `'B3'`
	 * @param $ref1 A cell reference such as `'AA23'`
	 * @return int `<0` if $ref1 is before $ref2; `>0` if $ref1 is greater than $ref2, and `0` if they are equal.
	 */
	public static function cellOrderCompare(string $ref1, string $ref2): int
	{
		if (preg_match('/^([A-Z]+)(\d+)$/', $ref1, $matches1) === 1 && preg_match('/^([A-Z]+)(\d+)$/', $ref2, $matches2) === 1) {
			$column1 = $matches1[1];
			$column2 = $matches2[1];
			$length1 = strlen($column1);
			$length2 = strlen($column2);
			if ($length1 !== $length2) {
				return $length1 <=> $length2;
			}
			$cmp = strcmp($column1, $column2);
			if ($cmp !== 0) {
				return $cmp;
			}
			$line1 = (int)$matches1[2];
			$line2 = (int)$matches2[2];
			return $line1 <=> $line2;
		}
		return strcmp($ref1, $ref2);
	}

	/**
	 * Gives the name of the highest column used in the spreadsheet (e.g. `BA`),
	 * or null if there is none.
	 * @throws XlsxFastEditorFileFormatException
	 * @throws XlsxFastEditorXmlException
	 */
	public function getHighestColumnName(int $sheetNumber): ?string
	{
		$rightMostCell = '';
		$xpath = $this->getXPathFromPath(self::getWorksheetPath($sheetNumber));
		$cs = $xpath->query("/o:worksheet/o:sheetData/o:row/o:c[last()]");
		if ($cs !== false) {
			for ($i = 0; $i < $cs->length; $i++) {
				$c = $cs[$i];
				if (!($c instanceof \DOMElement)) {
					throw new XlsxFastEditorXmlException("Error querying XML fragment for row {$sheetNumber}!");
				}
				$cellName = $c->getAttribute('r');
				if ($cellName !== '' && ($rightMostCell === '' || self::cellOrderCompare($rightMostCell, $cellName) < 0)) {
					$rightMostCell = $cellName;
				}
			}
		}
		return $rightMostCell === '' ? null : XlsxFastEditorCell::nameToColumn($rightMostCell);
	}

	/** To return `null` when accessing a row or cell that does not exist, e.g. via {@see XlsxFastEditor::getCell()} */
	public const ACCESS_MODE_NULL = 0;
	/** To throw an exception when accessing a row or cell that does not exist, e.g. via {@see XlsxFastEditor::getCell()} */
	public const ACCESS_MODE_EXCEPTION = 1;
	/** To auto-create the cell when accessing a row or cell that does not exist, e.g. via {@see XlsxFastEditor::getCell()} */
	public const ACCESS_MODE_AUTOCREATE = 2;

	/**
	 * Access the specified cell in the specified worksheet. Can create it automatically if asked to.
	 * The corresponding row can also be automatically created if it does not exist already, but the worksheet cannot be automatically created.
	 *
	 * ℹ️ Instead of calling multiple times this function, consider the faster navigation methods
	 * `XlsxFastEditor::rowsIterator()`, `XlsxFastEditor::getFirstRow()`, `XlsxFastEditorRow::cellsIterator()`,
	 * `XlsxFastEditorRow::getNextRow()`, `XlsxFastEditorRow::getFirstCell()`, `XlsxFastEditorCell::getNextCell()`, etc.
	 *
	 * @param int $sheetNumber Worksheet number (base 1)
	 * @param string $cellName Cell name such as `'B4'`
	 * @param int $accessMode To control the behaviour when the cell does not exist:
	 * set to `XlsxFastEditor::ACCESS_MODE_NULL` to return `null` (default),
	 * set to `XlsxFastEditor::ACCESS_MODE_EXCEPTION` to raise an `XlsxFastEditorInputException` exception,
	 * set to `XlsxFastEditor::ACCESS_MODE_AUTOCREATE` to auto-create the cell.
	 * @return XlsxFastEditorCell|null A cell, potentially `null` if the cell does not exist and `$accessMode` is set to `XlsxFastEditor::ACCESS_MODE_NULL`
	 * @phpstan-return ($accessMode is XlsxFastEditor::ACCESS_MODE_NULL ? XlsxFastEditorCell|null : XlsxFastEditorCell)
	 * @internal
	 * @throws XlsxFastEditorFileFormatException
	 * @throws \InvalidArgumentException if `$cellName` has an invalid format
	 * @throws XlsxFastEditorInputException optionally if the corresponding cell does not exist, depending on choice of `$accessMode`
	 * @throws XlsxFastEditorXmlException
	 */
	public function getCell(int $sheetNumber, string $cellName, int $accessMode = XlsxFastEditor::ACCESS_MODE_NULL): ?XlsxFastEditorCell
	{
		if (!ctype_alnum($cellName)) {
			throw new \InvalidArgumentException("Invalid cell reference {$cellName}!");
		}
		$cellName = strtoupper($cellName);

		$xpath = $this->getXPathFromPath(self::getWorksheetPath($sheetNumber));
		$cs = $xpath->query("/o:worksheet/o:sheetData/o:row/o:c[@r='{$cellName}'][1]");
		$c = null;
		if ($cs !== false && $cs->length > 0) {
			$c = $cs[0];
			if (!($c instanceof \DOMElement)) {
				throw new XlsxFastEditorXmlException("Error querying XML fragment for cell {$sheetNumber}/{$cellName}!");
			}
		}

		if ($c === null) {
			// The cell <c> was not found

			switch ($accessMode) {
				case XlsxFastEditor::ACCESS_MODE_EXCEPTION:
					throw new XlsxFastEditorInputException("Internal error accessing cell {$sheetNumber}/{$cellName}!");
				case XlsxFastEditor::ACCESS_MODE_AUTOCREATE:
					$rowNumber = (int)preg_replace('/[^\d]+/', '', $cellName);
					if ($rowNumber === 0) {
						throw new \InvalidArgumentException("Invalid line in cell reference {$cellName}!");
					}
					$row = $this->getRow($sheetNumber, $rowNumber, $accessMode);
					return $row->getCell($cellName, $accessMode);
				default:
				case XlsxFastEditor::ACCESS_MODE_NULL:
					return null;
			}
		}

		return new XlsxFastEditorCell($this, $sheetNumber, $c);
	}

	/**
	 * Access the specified cell in the specified worksheet, or `null` if if does not exist.
	 *
	 * ℹ️ Instead of calling multiple times this function, consider the faster navigation methods
	 * `XlsxFastEditor::rowsIterator()`, `XlsxFastEditor::getFirstRow()`, `XlsxFastEditorRow::cellsIterator()`,
	 * `XlsxFastEditorRow::getNextRow()`, `XlsxFastEditorRow::getFirstCell()`, `XlsxFastEditorCell::getNextCell()`, etc.
	 *
	 * @param int $sheetNumber Worksheet number (base 1)
	 * @param string $cellName Cell name such as `'B4'`
	 * @return XlsxFastEditorCell|null A cell, potentially `null` if the cell does not exist
	 * @throws \InvalidArgumentException if `$cellName` has an invalid format
	 * @throws XlsxFastEditorFileFormatException
	 * @throws XlsxFastEditorXmlException
	 */
	public function getCellOrNull(int $sheetNumber, string $cellName): ?XlsxFastEditorCell
	{
		try {
			return $this->getCell($sheetNumber, $cellName, XlsxFastEditor::ACCESS_MODE_NULL);
		} catch (XlsxFastEditorInputException $iex) {
			// Will not happen
			return null;
		}
	}

	/**
	 * Access the specified cell in the specified worksheet, or autocreate it if it does not already exist.
	 * The corresponding row can also be automatically created if it does not exist already, but the worksheet cannot be automatically created.
	 * @param int $sheetNumber Worksheet number (base 1)
	 * @param string $cellName Cell name such as `'B4'`
	 * @return XlsxFastEditorCell A cell
	 * @throws \InvalidArgumentException if `$cellName` has an invalid format
	 * @throws XlsxFastEditorFileFormatException
	 * @throws XlsxFastEditorXmlException
	 */
	public function getCellAutocreate(int $sheetNumber, string $cellName): XlsxFastEditorCell
	{
		try {
			return $this->getCell($sheetNumber, $cellName, XlsxFastEditor::ACCESS_MODE_AUTOCREATE);
		} catch (XlsxFastEditorInputException $iex) {
			// Will not happen
			throw new XlsxFastEditorXmlException('Internal error with getCell!', $iex->getCode(), $iex);
		}
	}

	/**
	 * Read a formula in the given worksheet at the given cell location.
	 *
	 * @param int $sheetNumber Worksheet number (base 1)
	 * @param string $cellName Cell name such as `'B4'`
	 * @return string|null an integer if the cell exists and contains a formula, `null` otherwise.
	 * @throws \InvalidArgumentException if `$cellName` has an invalid format
	 * @throws XlsxFastEditorFileFormatException
	 * @throws XlsxFastEditorXmlException
	 */
	public function readFormula(int $sheetNumber, string $cellName): ?string
	{
		$cell = $this->getCellOrNull($sheetNumber, $cellName);
		return $cell === null ? null : $cell->readFormula();
	}

	/**
	 * Read a float in the given worksheet at the given cell location.
	 *
	 * @param int $sheetNumber Worksheet number (base 1)
	 * @param string $cellName Cell name such as `'B4'`
	 * @return float|null a float if the cell exists and contains a number, `null` otherwise.
	 * @throws \InvalidArgumentException if `$cellName` has an invalid format
	 * @throws XlsxFastEditorFileFormatException
	 * @throws XlsxFastEditorXmlException
	 */
	public function readFloat(int $sheetNumber, string $cellName): ?float
	{
		$cell = $this->getCellOrNull($sheetNumber, $cellName);
		return $cell === null ? null : $cell->readFloat();
	}

	/**
	 * Read a date/time in the given worksheet at the given cell location.
	 *
	 * @param int $sheetNumber Worksheet number (base 1)
	 * @param string $cellName Cell name such as `'B4'`
	 * @return \DateTimeImmutable|null a date if the cell exists and contains a number, `null` otherwise.
	 * @throws \InvalidArgumentException if `$cellName` has an invalid format
	 * @throws XlsxFastEditorFileFormatException
	 * @throws XlsxFastEditorXmlException
	 */
	public function readDateTime(int $sheetNumber, string $cellName): ?\DateTimeImmutable
	{
		$cell = $this->getCellOrNull($sheetNumber, $cellName);
		return $cell === null ? null : $cell->readDateTime();
	}

	/**
	 * Read an integer in the given worksheet at the given cell location.
	 *
	 * @param int $sheetNumber Worksheet number (base 1)
	 * @param string $cellName Cell name such as `'B4'`
	 * @return int|null an integer if the cell exists and contains a number, `null` otherwise.
	 * @throws \InvalidArgumentException if `$cellName` has an invalid format
	 * @throws XlsxFastEditorFileFormatException
	 * @throws XlsxFastEditorXmlException
	 */
	public function readInt(int $sheetNumber, string $cellName): ?int
	{
		$cell = $this->getCellOrNull($sheetNumber, $cellName);
		return $cell === null ? null : $cell->readInt();
	}

	/**
	 * Access a string stored in the shared strings list.
	 * @param int $stringNumber String number (ID), base 0.
	 * @internal
	 * @throws XlsxFastEditorFileFormatException
	 * @throws XlsxFastEditorXmlException
	 */
	public function _getSharedString(int $stringNumber): ?string
	{
		$stringNumber++;	// Base 1

		$xpath = $this->getXPathFromPath(self::SHARED_STRINGS_PATH);
		$ts = $xpath->query("/o:sst/o:si[$stringNumber]//o:t");
		if ($ts !== false && $ts->length > 0) {
			$text = '';
			foreach ($ts as $t) {
				if (!($t instanceof \DOMElement)) {
					throw new XlsxFastEditorXmlException("Error querying XML shared string {$stringNumber}!");
				}
				$text .= $t->nodeValue;
			}
			return $text;
		}
		return null;
	}

	/**
	 * Read a string in the given worksheet at the given cell location,
	 * compatible with the shared string approach.
	 *
	 * @param int $sheetNumber Worksheet number (base 1)
	 * @param string $cellName Cell name such as `'B4'`
	 * @return string|null a string if the cell exists and contains a value, `null` otherwise.
	 * @throws \InvalidArgumentException if `$cellName` has an invalid format
	 * @throws XlsxFastEditorFileFormatException
	 * @throws XlsxFastEditorXmlException
	 */
	public function readString(int $sheetNumber, string $cellName): ?string
	{
		$cell = $this->getCellOrNull($sheetNumber, $cellName);
		return $cell === null ? null : $cell->readString();
	}

	private static function getWorksheetRelPath(int $sheetNumber): string
	{
		return "xl/worksheets/_rels/sheet{$sheetNumber}.xml.rels";
	}

	/**
	 * Access an hyperlink referenced from a cell of the specified sheet.
	 * @param string $rId Hyperlink reference.
	 * @internal
	 * @throws \InvalidArgumentException if `$rId` has an invalid format
	 * @throws XlsxFastEditorFileFormatException
	 * @throws XlsxFastEditorXmlException
	 */
	public function _getHyperlink(int $sheetNumber, string $rId): ?string
	{
		if (!ctype_alnum($rId)) {
			throw new \InvalidArgumentException("Invalid internal hyperlink reference {$sheetNumber}/{$rId}!");
		}
		$xpath = $this->getXPathFromPath(self::getWorksheetRelPath($sheetNumber));
		$xpath->registerNamespace('pr', 'http://schemas.openxmlformats.org/package/2006/relationships');
		$target = $xpath->evaluate(<<<xpath
			normalize-space(/pr:Relationships/pr:Relationship[@Id='{$rId}'
			and @Type='http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink'][1]/@Target)
		xpath);
		return is_string($target) ? $target : null;
	}

	/**
	 * Read a hyperlink in the given worksheet at the given cell location.
	 *
	 * @param int $sheetNumber Worksheet number (base 1)
	 * @param string $cellName Cell name such as `'B4'`
	 * @return string|null a string if the cell exists and contains a hyperlink, `null` otherwise.
	 * @throws \InvalidArgumentException if `$cellName` has an invalid format
	 * @throws XlsxFastEditorFileFormatException
	 * @throws XlsxFastEditorXmlException
	 */
	public function readHyperlink(int $sheetNumber, string $cellName): ?string
	{
		$cell = $this->getCellOrNull($sheetNumber, $cellName);
		return $cell === null ? null : $cell->readHyperlink();
	}

	/**
	 * Change an hyperlink associated to the given cell of the given worksheet.
	 * @return bool True if any hyperlink was cleared, false otherwise.
	 * @internal
	 * @throws \InvalidArgumentException if `$cellName` has an invalid format
	 * @throws XlsxFastEditorFileFormatException
	 * @throws XlsxFastEditorXmlException
	 */
	public function _setHyperlink(int $sheetNumber, string $rId, string $value): bool
	{
		if (!ctype_alnum($rId)) {
			throw new \InvalidArgumentException("Invalid internal hyperlink reference {$sheetNumber}/{$rId}!");
		}
		$xmlPath = self::getWorksheetRelPath($sheetNumber);
		$xpath = $this->getXPathFromPath($xmlPath);
		$xpath->registerNamespace('pr', 'http://schemas.openxmlformats.org/package/2006/relationships');
		$hyperlinks = $xpath->query(<<<xpath
			/pr:Relationships/pr:Relationship[@Id='{$rId}'
			and @Type='http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink'][1]
		xpath);
		if ($hyperlinks !== false && $hyperlinks->length > 0) {
			$hyperlink = $hyperlinks[0];
			if (!($hyperlink instanceof \DOMElement)) {
				throw new XlsxFastEditorXmlException("Error querying XML fragment for hyperlink {$sheetNumber}/{$rId}!");
			}
			$this->touchPath($xmlPath);
			return $hyperlink->setAttribute('Target', $value) !== false;
		}
		return false;
	}

	/**
	 * Replace the hyperlink of the cell, if that cell already has an hyperlink.
	 * Warning: does not support the creation of a new hyperlink.
	 *
	 * @param int $sheetNumber Worksheet number (base 1)
	 * @param string $cellName Cell name such as `'B4'`
	 * @return bool True if the hyperlink could be replaced, false otherwise.
	 * @throws \InvalidArgumentException if `$cellName` has an invalid format
	 * @throws XlsxFastEditorFileFormatException
	 * @throws XlsxFastEditorXmlException
	 */
	public function writeHyperlink(int $sheetNumber, string $cellName, string $value): bool
	{
		$cell = $this->getCellOrNull($sheetNumber, $cellName);
		return $cell === null ? false : $cell->writeHyperlink($value);
	}

	/**
	 * Write a formulat in the given worksheet at the given cell location, without changing the type/style of the cell.
	 * Auto-creates the cell if it does not already exists.
	 * Removes the formulas of the cell, if any.
	 *
	 * @param int $sheetNumber Worksheet number (base 1)
	 * @param string $cellName Cell name such as `'B4'`
	 * @throws \InvalidArgumentException if `$cellName` has an invalid format
	 * @throws XlsxFastEditorFileFormatException
	 * @throws XlsxFastEditorXmlException
	 */
	public function writeFormula(int $sheetNumber, string $cellName, string $value): void
	{
		$cell = $this->getCellAutocreate($sheetNumber, $cellName);
		$cell->writeFormula($value);
	}

	/**
	 * Write a number in the given worksheet at the given cell location, without changing the type/style of the cell.
	 * Auto-creates the cell if it does not already exists.
	 * Removes the formulas of the cell, if any.
	 *
	 * @param int $sheetNumber Worksheet number (base 1)
	 * @param string $cellName Cell name such as `'B4'`
	 * @param float $value
	 * @throws \InvalidArgumentException if `$cellName` has an invalid format
	 * @throws XlsxFastEditorFileFormatException
	 * @throws XlsxFastEditorXmlException
	 */
	public function writeFloat(int $sheetNumber, string $cellName, float $value): void
	{
		$cell = $this->getCellAutocreate($sheetNumber, $cellName);
		$cell->writeFloat($value);
	}

	/**
	 * Write a number in the given worksheet at the given cell location, without changing the type/style of the cell.
	 * Auto-creates the cell if it does not already exists.
	 * Removes the formulas of the cell, if any.
	 *
	 * @param int $sheetNumber Worksheet number (base 1)
	 * @param string $cellName Cell name such as `'B4'`
	 * @param int $value
	 * @throws \InvalidArgumentException if `$cellName` has an invalid format
	 * @throws XlsxFastEditorFileFormatException
	 * @throws XlsxFastEditorXmlException
	 */
	public function writeInt(int $sheetNumber, string $cellName, int $value): void
	{
		$cell = $this->getCellAutocreate($sheetNumber, $cellName);
		$cell->writeInt($value);
	}

	/**
	 * Adds a new shared string and returns its ID.
	 * @internal
	 * @param string $value Value of the new shared string.
	 * @return int the ID of the new shared string.
	 * @throws XlsxFastEditorFileFormatException
	 * @throws XlsxFastEditorXmlException
	 */
	public function _makeNewSharedString(string $value): int
	{
		$dom = $this->getDomFromPath(self::SHARED_STRINGS_PATH);
		if ($dom->firstElementChild === null) {
			throw new XlsxFastEditorXmlException('Invalid shared strings!');
		}

		try {
			$si = $dom->createElement('si');
		} catch (\DOMException $dex) {
			throw new XlsxFastEditorXmlException('Error creating <si> in shared strings!', $dex->code, $dex);
		}
		if ($si === false) {
			throw new XlsxFastEditorXmlException('Error creating <si> in shared strings!');
		}

		try {
			$t = $dom->createElement('t', $value);
		} catch (\DOMException $dex) {
			throw new XlsxFastEditorXmlException('Error creating <t> in shared strings!', $dex->code, $dex);
		}
		if ($t === false) {
			throw new XlsxFastEditorXmlException('Error creating <t> in shared strings!');
		}
		$si->appendChild($t);
		$dom->firstElementChild->appendChild($si);

		$count = (int)$dom->firstElementChild->getAttribute('count');
		$dom->firstElementChild->setAttribute('count', (string)($count + 1));

		$uniqueCount = $dom->getElementsByTagName('si')->length;
		$dom->firstElementChild->setAttribute('uniqueCount', (string)$uniqueCount);

		$this->touchPath(self::SHARED_STRINGS_PATH);
		return $uniqueCount - 1;	// Base 0
	}

	/**
	 * Write a string in the given worksheet at the given cell location, without changing the type/style of the cell.
	 * Auto-creates the cell if it does not already exists.
	 * Removes the formulas of the cell, if any.
	 *
	 * @param int $sheetNumber Worksheet number (base 1)
	 * @param string $cellName Cell name such as `'B4'`
	 * @throws \InvalidArgumentException if `$cellName` has an invalid format
	 * @throws XlsxFastEditorFileFormatException
	 * @throws XlsxFastEditorXmlException
	 */
	public function writeString(int $sheetNumber, string $cellName, string $value): void
	{
		$cell = $this->getCellAutocreate($sheetNumber, $cellName);
		$cell->writeString($value);
	}

	/**
	 * Regex search & replace text strings in all worksheets using [`preg_replace()`](https://php.net/function.preg-replace)
	 *
	 * @param string|array<string> $pattern The pattern to search for.
	 * @param string|array<string> $replacement The string or an array with strings to replace.
	 * @return int The number of replacements done.
	 * @throws XlsxFastEditorFileFormatException
	 * @throws XlsxFastEditorXmlException
	 */
	public function textReplace($pattern, $replacement): int
	{
		$dom = $this->getDomFromPath(self::SHARED_STRINGS_PATH);
		$elements = $dom->getElementsByTagName('t');
		$nb = 0;

		if ($elements->length > 0) {
			foreach ($elements as $element) {
				if ($element instanceof \DOMElement) {
					$text = preg_replace($pattern, $replacement, $element->textContent);
					if (is_string($text) && $element->textContent !== $text) {
						$element->textContent = $text;
						$nb++;
					}
				}
			}
		}

		if ($nb > 0) {
			$this->touchPath(self::SHARED_STRINGS_PATH);
		}
		return $nb;
	}
}
