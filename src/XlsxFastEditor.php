<?php

namespace alexandrainst\XlsxFastEditor;

/**
 * Fast edit an existing XLSX/XLSM document (Microsoft Excel 2007+, Office Open XML Workbook)
 * using low-level ZIP and XML manipulation.
 */
final class XlsxFastEditor
{
	private const OXML_NAMESPACE = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main';
	private const SHARED_STRINGS_PATH = 'xl/sharedStrings.xml';
	private const WORKBOOK_PATH = 'xl/workbook.xml';
	private const CALC_CHAIN_CACHE_PATH = 'xl/calcChain.xml';

	private \ZipArchive $zip;

	/**
	 * Cache of the XML documents.
	 * @var array<string,\DOMDocument>
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
	 * @param int $sheetNumber Worksheet number (base 1)
	 */
	private function touchWorksheet(int $sheetNumber): void
	{
		$path = self::getWorksheetPath($sheetNumber);
		$this->touchPath($path);
	}

	/**
	 * Close the underlying document archive.
	 * Note: changes need to be explicitly saved before (see `XlsxFastEditor::save()`)
	 * Note: the object should not be used anymore afterwards.
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
			$dom = $this->documents[$name];
			if (!$this->zip->deleteName($name)) {
				throw new XlsxFastEditorZipException("Error deleting old fragment {$name}!");
			}
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
	 * Get a worksheet number (ID) from its name.
	 * @param string $sheetName The name of the worksheet to look up.
	 * @return int The worksheet ID, or -1 if not found.
	 */
	public function getWorksheetNumber(string $sheetName): int
	{
		$dom = $this->getDomFromPath(self::WORKBOOK_PATH);
		$xpath = new \DOMXPath($dom);
		$xpath->registerNamespace('o', self::OXML_NAMESPACE);
		$sheetId = $xpath->evaluate("normalize-space(//o:sheet[@name='$sheetName'][1]/@sheetId)");
		if (is_string($sheetId)) {
			return (int)$sheetId;
		}
		return -1;
	}

	private static function getWorksheetPath(int $sheetNumber): string
	{
		return "xl/worksheets/sheet{$sheetNumber}.xml";
	}

	/**
	 * Extracts a worksheet from the internal ZIP document,
	 * parse the XML, and returns a DOM document.
	 * The DOM document is then cached.
	 * @param string $path The path of the document inside the ZIP document.
	 */
	private function getDomFromPath(string $path): \DOMDocument
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

		$this->documents[$path] = $dom;
		return $dom;
	}

	/**
	 * Access the DOMElement representing a cell formula `<f>` in the worksheet.
	 *
	 * @param int $sheetNumber Worksheet number (base 1)
	 * @param $cellName Cell name such as `B4`
	 */
	private function getF(int $sheetNumber, string $cellName): ?\DOMElement
	{
		if (!ctype_alnum($cellName)) {
			throw new XlsxFastEditorInputException("Invalid cell reference {$cellName}! ");
		}
		$cellName = strtoupper($cellName);

		$dom = $this->getDomFromPath(self::getWorksheetPath($sheetNumber));
		$xpath = new \DOMXPath($dom);
		$xpath->registerNamespace('o', self::OXML_NAMESPACE);

		$f = null;
		$fs = $xpath->query("(//o:c[@r='$cellName'])[1]/o:f");
		if ($fs !== false && $fs->length > 0) {
			$f = $fs[0];
			if (!($f instanceof \DOMElement)) {
				throw new XlsxFastEditorXmlException("Error querying XML fragment for cell formula {$sheetNumber}/{$cellName}!");
			}
		}
		return $f;
	}

	/**
	 * Read a formula in the given worksheet at the given cell location.
	 *
	 * @param int $sheetNumber Worksheet number (base 1)
	 * @param $cellName Cell name such as `B4`
	 */
	public function readFormula(int $sheetNumber, string $cellName): ?string
	{
		$f = $this->getF($sheetNumber, $cellName);
		if ($f === null || !is_string($f->nodeValue) || $f->nodeValue === '') {
			return null;
		}
		return '=' . $f->nodeValue;
	}

	/**
	 * Access the DOMElement representing a cell value `<v>` in the worksheet.
	 *
	 * @param int $sheetNumber Worksheet number (base 1)
	 * @param $cellName Cell name such as `B4`
	 */
	private function getV(int $sheetNumber, string $cellName): ?\DOMElement
	{
		if (!ctype_alnum($cellName)) {
			throw new XlsxFastEditorInputException("Invalid cell reference {$cellName}! ");
		}
		$cellName = strtoupper($cellName);

		$dom = $this->getDomFromPath(self::getWorksheetPath($sheetNumber));
		$xpath = new \DOMXPath($dom);
		$xpath->registerNamespace('o', self::OXML_NAMESPACE);

		$v = null;
		$vs = $xpath->query("(//o:c[@r='$cellName'])[1]/o:v");
		if ($vs !== false && $vs->length > 0) {
			$v = $vs[0];
			if (!($v instanceof \DOMElement)) {
				throw new XlsxFastEditorXmlException("Error querying XML fragment for cell value {$sheetNumber}/{$cellName}!");
			}
		}
		return $v;
	}

	/**
	 * Read a float in the given worksheet at the given cell location.
	 *
	 * @param int $sheetNumber Worksheet number (base 1)
	 * @param $cellName Cell name such as `B4`
	 */
	public function readFloat(int $sheetNumber, string $cellName): ?float
	{
		$v = $this->getV($sheetNumber, $cellName);
		if ($v === null || !is_string($v->nodeValue) || !is_numeric($v->nodeValue)) {
			return null;
		}
		return (float)$v->nodeValue;
	}

	/**
	 * Read an integer in the given worksheet at the given cell location.
	 *
	 * @param int $sheetNumber Worksheet number (base 1)
	 * @param $cellName Cell name such as `B4`
	 */
	public function readInt(int $sheetNumber, string $cellName): ?int
	{
		$v = $this->getV($sheetNumber, $cellName);
		if ($v === null || !is_string($v->nodeValue) || !is_numeric($v->nodeValue)) {
			return null;
		}
		return (int)$v->nodeValue;
	}

	/**
	 * Read a string in the given worksheet at the given cell location,
	 * compatible with the shared string approach.
	 *
	 * @param int $sheetNumber Worksheet number (base 1)
	 * @param $cellName Cell name such as `B4`
	 */
	public function readString(int $sheetNumber, string $cellName): ?string
	{
		$v = $this->getV($sheetNumber, $cellName);
		if ($v === null || !is_string($v->nodeValue)) {
			return null;
		}
		$c = $v->parentNode;
		if ($c === null || !($c instanceof \DOMElement)) {
			throw new XlsxFastEditorXmlException("Error querying XML fragment for cell {$sheetNumber}/{$cellName}!");
		}

		if ($c->getAttribute('t') === 's') {
			// Shared string

			if (!ctype_digit($v->nodeValue)) {
				throw new XlsxFastEditorXmlException("Error querying XML fragment for shared string {$sheetNumber}/{$cellName}!");
			}

			$sharedStringId = 1 + (int)$v->nodeValue;

			$dom = $this->getDomFromPath(self::SHARED_STRINGS_PATH);
			$xpath = new \DOMXPath($dom);
			$xpath->registerNamespace('o', self::OXML_NAMESPACE);

			$ts = $xpath->query("/o:sst/o:si[$sharedStringId]/o:t[1]");
			if ($ts !== false && $ts->length > 0) {
				$t = $ts[0];
				if (!($t instanceof \DOMElement)) {
					throw new XlsxFastEditorXmlException("Error querying XML shared string for {$sheetNumber}/{$cellName}!");
				}
				return $t->nodeValue;
			}
		} else {
			// Local value
			return $v->nodeValue;
		}

		return null;
	}

	/**
	 * Defines the *Full calculation on load* policy for the specified worksheet.
	 * @param int $sheetNumber Worksheet number (base 1)
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
					$sheetCalcPr = $dom->createElement('sheetCalcPr');
					if ($sheetCalcPr !== false && $sheetData->parentNode !== null) {
						$sheetData->parentNode->insertBefore($sheetCalcPr, $sheetData->nextSibling);
					}
				}
			}
		}
		if ($sheetCalcPr instanceof \DOMElement) {
			$sheetCalcPr->setAttribute('fullCalcOnLoad', $value ? 'true' : 'false');
			$this->touchWorksheet($sheetNumber);
		}
	}

	/**
	 * Sort cells within the same line, such as B3, AA3. Compare only the column part.
	 * @param $ref1 A cell reference such as B3
	 * @param $ref1 A cell reference such as AA3
	 * @return int -1 if $ref1 is before $ref2; 1 if $ref1 is greater than $ref2, and 0 if they are equal.
	 */
	private static function columnOrderCompare(string $ref1, string $ref2): int
	{
		$pattern = '/[^A-Z]+/';
		$column1 = preg_replace($pattern, '', $ref1) ?? '';
		$column2 = preg_replace($pattern, '', $ref2) ?? '';
		$length1 = strlen($column1);
		$length2 = strlen($column2);
		if ($length1 !== $length2) {
			return $length1 <=> $length2;
		}
		return strcmp($ref1, $ref2);
	}

	/**
	 * Access the DOMElement representing a cell in the worksheet.
	 * Creates it if necessary.
	 *
	 * @param int $sheetNumber Worksheet number (base 1)
	 * @param $cellName Cell name such as `B4`
	 * @param bool $autoCreate Set to true to automatically create a cell if it does not exist already, false to make no change.
	 */
	private function getCell(int $sheetNumber, string $cellName, bool $autoCreate): ?\DOMElement
	{
		$dom = $this->getDomFromPath(self::getWorksheetPath($sheetNumber));
		$xpath = new \DOMXPath($dom);
		$xpath->registerNamespace('o', self::OXML_NAMESPACE);

		if (!ctype_alnum($cellName)) {
			throw new XlsxFastEditorInputException("Invalid cell reference {$cellName}! ");
		}
		$cellName = strtoupper($cellName);

		$c = null;
		$cs = $xpath->query("(//o:c[@r='$cellName'])[1]");
		if ($cs !== false && $cs->length > 0) {
			$c = $cs[0];
			if (!($c instanceof \DOMElement)) {
				throw new XlsxFastEditorXmlException("Error querying XML fragment for cell {$sheetNumber}/{$cellName}!");
			}
		}

		if ($c === null && $autoCreate) {
			// The cell <c> was not found

			$rowNumber = (int)preg_replace('/[^\d]+/', '', $cellName);
			if ($rowNumber === 0) {
				throw new XlsxFastEditorInputException("Invalid cell reference {$cellName}!");
			}

			$row = null;
			$rows = $xpath->query("(//o:row[@r='$rowNumber'])[1]");
			if ($rows !== false && $rows->length > 0) {
				$row = $rows[0];
				if (!($row instanceof \DOMElement)) {
					throw new XlsxFastEditorXmlException("Error querying XML fragment for cell {$sheetNumber}/{$cellName}!");
				}
			}

			if ($row === null) {
				// The <row> was not found

				$sheetDatas = $dom->getElementsByTagName('sheetData');
				if ($sheetDatas->length === 0) {
					throw new XlsxFastEditorXmlException("Cannot find sheetData for worksheet {$sheetNumber}!");
				}
				$sheetData = $sheetDatas[0];
				if (!($sheetData instanceof \DOMElement)) {
					throw new XlsxFastEditorXmlException("Error querying XML fragment for worksheet {$sheetNumber}!");
				}
				$row = $dom->createElement('row');
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
			}

			$c = $dom->createElement('c');
			if ($c === false) {
				throw new XlsxFastEditorXmlException("Error creating cell {$sheetNumber}/{$cellName}!");
			}
			$c->setAttribute('r', $cellName);

			// Excel expects the cells to be sorted
			$sibling = $row->firstElementChild;
			while ($sibling !== null && self::columnOrderCompare($sibling->getAttribute('r'), $cellName) < 0) {
				$sibling = $sibling->nextElementSibling;
			}
			$row->insertBefore($c, $sibling);
		}

		return $c;
	}

	/**
	 * Clean a cell to have its value written.
	 * @param \DOMElement $c A `<c>` cell element.
	 * @return \DOMElement The `<v>` value element of the provided cell, or null in case of error.
	 */
	private function initCellValue(\DOMElement $c): \DOMElement
	{
		$v = null;
		$c->removeAttribute('t');	// Remove type, if it exists
		for ($i = $c->childNodes->length - 1; $i >= 0; $i--) {
			// Remove all childs except <v>
			$child = $c->childNodes[$i];
			if ($child instanceof \DOMElement) {
				if ($child->localName === 'v') {
					$v = $child;
				} else {
					if ($child->localName === 'f') {
						// This cell had a formula. Must clear calcChain:
						$this->mustClearCalcChain = true;
					}
					$c->removeChild($child);
				}
			}
		}
		if ($v === null) {
			// There was no existing <v>
			$v = $c->ownerDocument === null ? null : $c->ownerDocument->createElement('v');
			if ($v == false) {
				throw new XlsxFastEditorXmlException('Error creating value for cell!');
			}
			$c->appendChild($v);
		}

		return $v;
	}

	/**
	 * Write a formulat in the given worksheet at the given cell location, without changing the type/style of the cell.
	 * Removes the formulas of the cell, if any.
	 *
	 * @param int $sheetNumber Worksheet number (base 1)
	 * @param $cellName Cell name such as `B4`
	 */
	public function writeFormula(int $sheetNumber, string $cellName, string $value): void
	{
		$c = $this->getCell($sheetNumber, $cellName, true);
		if ($c === null) {
			throw new XlsxFastEditorInputException("Internal error accessing cell {$sheetNumber}/{$cellName}!");
		}

		$value = ltrim($value, '=');

		$vs = $c->getElementsByTagName('v');
		for ($i = $vs->length - 1; $i >= 0; $i--) {
			$v = $vs[$i];
			if ($v instanceof \DOMElement) {
				$c->removeChild($v);
			}
		}

		$fs = $c->getElementsByTagName('f');
		for ($i = $fs->length - 1; $i >= 0; $i--) {
			$f = $fs[$i];
			if ($f instanceof \DOMElement) {
				$c->removeChild($f);
			}
		}

		$dom = $c->ownerDocument;
		if ($dom === null) {
			throw new XlsxFastEditorInputException("Internal error accessing cell {$sheetNumber}/{$cellName}!");
		}
		$f = $dom->createElement('f', $value);
		$c->appendChild($f);

		$this->mustClearCalcChain = true;
		$this->touchWorksheet($sheetNumber);
	}

	/**
	 * Write a number in the given worksheet at the given cell location, without changing the type/style of the cell.
	 * Removes the formulas of the cell, if any.
	 *
	 * @param int $sheetNumber Worksheet number (base 1)
	 * @param $cellName Cell name such as `B4`
	 * @param int|float $value
	 */
	private function writeNumber(int $sheetNumber, string $cellName, $value): void
	{
		$c = $this->getCell($sheetNumber, $cellName, true);
		if ($c === null) {
			throw new XlsxFastEditorInputException("Internal error accessing cell {$sheetNumber}/{$cellName}!");
		}
		$v = $this->initCellValue($c);
		$v->nodeValue = (string)$value;
		$this->touchWorksheet($sheetNumber);
	}

	/**
	 * Write a number in the given worksheet at the given cell location, without changing the type/style of the cell.
	 * Removes the formulas of the cell, if any.
	 *
	 * @param int $sheetNumber Worksheet number (base 1)
	 * @param $cellName Cell name such as `B4`
	 * @param float $value
	 */
	public function writeFloat(int $sheetNumber, string $cellName, float $value): void
	{
		$this->writeNumber($sheetNumber, $cellName, $value);
	}

	/**
	 * Write a number in the given worksheet at the given cell location, without changing the type/style of the cell.
	 * Removes the formulas of the cell, if any.
	 *
	 * @param int $sheetNumber Worksheet number (base 1)
	 * @param $cellName Cell name such as `B4`
	 * @param int $value
	 */
	public function writeInt(int $sheetNumber, string $cellName, int $value): void
	{
		$this->writeNumber($sheetNumber, $cellName, $value);
	}

	/**
	 * Adds a new shared string and returns its ID.
	 * @param string $value Value of the new shared string.
	 * @return int the ID of the new shared string.
	 */
	private function makeNewSharedString(string $value): int
	{
		$dom = $this->getDomFromPath(self::SHARED_STRINGS_PATH);
		if ($dom->firstElementChild === null) {
			throw new XlsxFastEditorXmlException('Invalid shared strings!');
		}

		$si = $dom->createElement('si');
		if ($si === false) {
			throw new XlsxFastEditorXmlException('Error creating <si> in shared strings!');
		}
		$t = $dom->createElement('t', $value);
		if ($t === false) {
			throw new XlsxFastEditorXmlException('Error creating <t> in shared strings!');
		}
		$si->appendChild($t);
		$dom->firstElementChild->appendChild($si);

		$count = (int)$dom->firstElementChild->getAttribute('count');
		$dom->firstElementChild->setAttribute('count', (string)($count + 1));

		$uniqueCount = $dom->getElementsByTagName('si')->length;
		$dom->firstElementChild->setAttribute('uniqueCount', (string)$uniqueCount);
		return $uniqueCount - 1;	// Base 0
	}

	/**
	 * Write a string in the given worksheet at the given cell location, without changing the type/style of the cell.
	 * Removes the formulas of the cell, if any.
	 *
	 * @param int $sheetNumber Worksheet number (base 1)
	 * @param $cellName Cell name such as `B4`
	 */
	public function writeString(int $sheetNumber, string $cellName, string $value): void
	{
		$c = $this->getCell($sheetNumber, $cellName, true);
		if ($c === null) {
			throw new XlsxFastEditorInputException("Internal error accessing cell {$sheetNumber}/{$cellName}!");
		}
		$v = self::initCellValue($c);
		$c->setAttribute('t', 's');	// Type shared string
		$sharedStringId = self::makeNewSharedString($value);
		$v->nodeValue = (string)$sharedStringId;
		$this->touchWorksheet($sheetNumber);
	}

	/**
	 * Regex search & replace text strings in all worksheets using [`preg_replace()`](https://php.net/function.preg-replace)
	 *
	 * @param string|array<string> $pattern The pattern to search for.
	 * @param string|array<string> $replacement The string or an array with strings to replace.
	 * @return int The number of replacements done.
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
