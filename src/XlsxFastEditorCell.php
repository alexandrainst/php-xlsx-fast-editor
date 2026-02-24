<?php

declare(strict_types=1);

namespace alexandrainst\XlsxFastEditor;

/**
 * Class used to return some read-only raw information about a cell.
 */
final class XlsxFastEditorCell
{
	private XlsxFastEditor $editor;
	private int $sheetNumber;
	private \DOMElement $c;
	private ?\DOMXPath $xpath = null;

	/**
	 * @internal
	 */
	public function __construct(XlsxFastEditor $editor, int $sheetNumber, \DOMElement $c)
	{
		$this->editor = $editor;
		$this->sheetNumber = $sheetNumber;
		$this->c = $c;
	}

	/**
	 * @throws XlsxFastEditorXmlException
	 */
	private function getXPath(): \DOMXPath
	{
		if ($this->xpath === null) {
			$dom = $this->c->ownerDocument;
			if ($dom === null) {
				throw new XlsxFastEditorXmlException("Internal error accessing cell {$this->name()}!");
			}
			$xpath = new \DOMXPath($dom);
			$xpath->registerNamespace('o', XlsxFastEditor::_OXML_NAMESPACE);
			$xpath->registerNamespace('or', 'http://schemas.openxmlformats.org/officeDocument/2006/relationships');
			$this->xpath = $xpath;
		}
		return $this->xpath;
	}

	/**
	 * Cell name (e.g., `'D4'`).
	 */
	public function name(): string
	{
		return $this->c->getAttribute('r');
	}

	/**
	 * $return column name (e.g., `'D'`).
	 * @throws XlsxFastEditorXmlException
	 */
	public static function nameToColumn(string $name): string
	{
		if (preg_match('/^([A-Z]+)/', $name, $matches) == 0 || empty($matches[1])) {
			throw new XlsxFastEditorXmlException("Error querying column name for cell name {$name}!");
		}
		return $matches[1];
	}

	/**
	 * Column name (e.g., `'D'`).
	 * @throws XlsxFastEditorXmlException
	 */
	public function column(): string
	{
		return self::nameToColumn($this->name());
	}

	/**
	 * Access the previous existing cell, if any, `null` otherwise.
	 * ℹ️ This is a faster method than `XlsxFastEditorRow::getCellOrNull()`
	 */
	public function getPreviousCell(): ?XlsxFastEditorCell
	{
		$c = $this->c->previousElementSibling;
		while ($c instanceof \DOMElement) {
			if ($c->localName === 'c') {
				return new XlsxFastEditorCell($this->editor, $this->sheetNumber, $c);
			}
			$c = $c->previousElementSibling;
		}
		return null;
	}

	/**
	 * Access the next existing cell, if any, `null` otherwise.
	 * ℹ️ This is a faster method than `XlsxFastEditorRow::getCellOrNull()`
	 */
	public function getNextCell(): ?XlsxFastEditorCell
	{
		$c = $this->c->nextElementSibling;
		while ($c instanceof \DOMElement) {
			if ($c->localName === 'c') {
				return new XlsxFastEditorCell($this->editor, $this->sheetNumber, $c);
			}
			$c = $c->nextElementSibling;
		}
		return null;
	}

	/**
	 * Access the parent row of the cell.
	 * @throws XlsxFastEditorXmlException
	 */
	public function getRow(): XlsxFastEditorRow
	{
		$r = $this->c->parentNode;
		if (!($r instanceof \DOMElement)) {
			throw new XlsxFastEditorXmlException("Error querying XML row for cell {$this->name()}!");
		}
		return new XlsxFastEditorRow($this->editor, $this->sheetNumber, $r);
	}

	/**
	 * Read a formula in the given worksheet at the given cell location.
	 * @return string|null an integer if the cell exists and contains a formula, `null` otherwise.
	 * @throws XlsxFastEditorXmlException
	 */
	public function readFormula(): ?string
	{
		$fs = $this->c->getElementsByTagName('f');
		if ($fs->length > 0) {
			$v = $fs[0];
			if (!($v instanceof \DOMElement)) {
				throw new XlsxFastEditorXmlException("Error querying XML formula for cell {$this->name()}!");
			}
			return '=' . $v->nodeValue;
		}
		return null;
	}

	private function value(): ?string
	{
		$vs = $this->c->getElementsByTagName('v');
		if ($vs->length > 0) {
			$v = $vs[0];
			if ($v instanceof \DOMElement) {
				return $v->nodeValue;
			}
		}
		return null;
	}

	/**
	 * Read the float value of the cell.
	 * @return float|null a float if the cell exists and contains a number, `null` otherwise.
	 */
	public function readFloat(): ?float
	{
		$value = $this->value();
		if ($value === null || !is_numeric($value)) {
			return null;
		}
		return (float)$value;
	}

	/**
	 * Read the date/time value of the cell, if any.
	 * @return \DateTimeImmutable|null a date if the cell exists and contains a number, `null` otherwise.
	 * @throws XlsxFastEditorFileFormatException
	 * @throws XlsxFastEditorXmlException
	 */
	public function readDateTime(): ?\DateTimeImmutable
	{
		$value = $this->readFloat();
		if ($value === null) {
			return null;
		}
		try {
			return XlsxFastEditor::excelDateToDateTime($value, $this->editor->getWorkbookDateSystem());
		} catch (\InvalidArgumentException $iaex) {
			// Never happens
			return null;
		}
	}

	/**
	 * Read the integer value of the cell.
	 * @return int|null an integer if the cell exists and contains a number, `null` otherwise.
	 */
	public function readInt(): ?int
	{
		$value = $this->value();
		if ($value === null || !is_numeric($value)) {
			return null;
		}
		return (int)$value;
	}

	/**
	 * Read the string value of the cell, compatible with the shared string approach.
	 * @return string|null a string if the cell exists and contains a value, `null` otherwise.
	 * @throws XlsxFastEditorFileFormatException
	 * @throws XlsxFastEditorXmlException
	 */
	public function readString(): ?string
	{
		$value = $this->value();
		if ($value === null) {
			return null;
		}

		if ($this->c->getAttribute('t') === 's') {
			// Shared string
			if (!ctype_digit($value)) {
				throw new XlsxFastEditorXmlException("Error querying XML fragment for shared string in cell {$this->name()}!");
			}
			return $this->editor->_getSharedString((int)$value);
		} else {
			// Local value
			return $value;
		}
	}

	/**
	 * Read the hyperlink value of the cell, if any.
	 * @return string|null a string if the cell exists and contains a hyperlink, `null` otherwise.
	 * @throws XlsxFastEditorFileFormatException
	 * @throws XlsxFastEditorXmlException
	 */
	public function readHyperlink(): ?string
	{
		$xpath = $this->getXPath();
		$rid = $xpath->evaluate("normalize-space(/o:worksheet/o:hyperlinks/o:hyperlink[@ref='{$this->name()}'][1]/@or:id)");
		if (!is_string($rid) || $rid === '') {
			return null;
		}
		try {
			return $this->editor->_getHyperlink($this->sheetNumber, $rid);
		} catch (\InvalidArgumentException $iax) {
			throw new XlsxFastEditorXmlException("Error querying XML fragment for hyperlink in cell {$this->name()}!", $iax->getCode(), $iax);
		}
	}

	/**
	 * Clean the cell to have its value written.
	 * @return \DOMElement The `<v>` value element of the provided cell, or `null` in case of error.
	 * @throws XlsxFastEditorXmlException
	 */
	private function initCellValue(): \DOMElement
	{
		$v = null;
		$this->c->removeAttribute('t');	// Remove type, if it exists
		for ($i = $this->c->childNodes->length - 1; $i >= 0; $i--) {
			// Remove all children except <v>
			$child = $this->c->childNodes[$i];
			if ($child instanceof \DOMElement) {
				if ($child->localName === 'v') {
					$v = $child;
				} else {
					if ($child->localName === 'f') {
						// This cell had a formula. Must clear calcChain:
						$this->editor->_clearCalcChain();
					}
					$this->c->removeChild($child);
				}
			}
		}
		if ($v === null) {
			// There was no existing <v>
			try {
				$v = $this->c->ownerDocument === null ? null : $this->c->ownerDocument->createElement('v');
			} catch (\DOMException $dex) {
				throw new XlsxFastEditorXmlException("Error creating value for cell {$this->name()}!", $dex->code, $dex);
			}
			if ($v == false) {
				throw new XlsxFastEditorXmlException("Error creating value for cell {$this->name()}!");
			}
			$this->c->appendChild($v);
		}

		return $v;
	}

	/**
	 * Write a formula, without changing the type/style of the cell.
	 * Removes the formulas of the cell, if any.
	 * @throws XlsxFastEditorXmlException
	 */
	public function writeFormula(string $value): void
	{
		$value = ltrim($value, '=');

		$vs = $this->c->getElementsByTagName('v');
		for ($i = $vs->length - 1; $i >= 0; $i--) {
			$v = $vs[$i];
			if ($v instanceof \DOMElement) {
				$this->c->removeChild($v);
			}
		}

		$fs = $this->c->getElementsByTagName('f');
		for ($i = $fs->length - 1; $i >= 0; $i--) {
			$f = $fs[$i];
			if ($f instanceof \DOMElement) {
				$this->c->removeChild($f);
			}
		}

		$dom = $this->c->ownerDocument;
		if ($dom === null) {
			throw new XlsxFastEditorXmlException("Internal error accessing cell {$this->name()}!");
		}
		try {
			// First, we create an empty element t
			$f = $dom->createElement('f');
			if ($f === false) {
				throw new XlsxFastEditorXmlException("Error creating DOMElement of formula for cell {$this->name()}!");
			}
			// Add content as a text node
			$textNode = $dom->createTextNode($value);
			if ($textNode === false) {
				throw new XlsxFastEditorXmlException("Error creating text node of formula for cell {$this->name()}!");
			}
			$f->appendChild($textNode);
		} catch (\DOMException $dex) {
			throw new XlsxFastEditorXmlException("Error creating formula for cell {$this->name()}!", $dex->code, $dex);
		}

		$this->c->appendChild($f);

		$this->editor->_clearCalcChain();
		$this->editor->_touchWorksheet($this->sheetNumber);
	}

	/**
	 * Write a number, without changing the type/style of the cell.
	 * Removes the formulas of the cell, if any.
	 * @param int|float $value
	 * @throws XlsxFastEditorXmlException
	 */
	private function writeNumber($value): void
	{
		$v = $this->initCellValue();
		$v->nodeValue = (string)$value;
		$this->editor->_touchWorksheet($this->sheetNumber);
	}

	/**
	 * Write a float, without changing the type/style of the cell.
	 * Removes the formulas of the cell, if any.
	 * @param float $value
	 * @throws XlsxFastEditorXmlException
	 */
	public function writeFloat(float $value): void
	{
		$this->writeNumber($value);
	}

	/**
	 * Write the date/time value of the cell, without changing the type/style of the cell.
	 * Removes the formulas of the cell, if any.
	 * @param \DateTimeInterface $value
	 * @throws \InvalidArgumentException
	 * @throws XlsxFastEditorFileFormatException
	 * @throws XlsxFastEditorXmlException
	 */
	public function writeDateTime(\DateTimeInterface $value): void
	{
		$floatValue = XlsxFastEditor::dateTimeToExcelDate($value, $this->editor->getWorkbookDateSystem());
		$this->writeNumber($floatValue);
	}

	/**
	 * Write an integer, without changing the type/style of the cell.
	 * Removes the formulas of the cell, if any.
	 * @param int $value
	 * @throws XlsxFastEditorXmlException
	 */
	public function writeInt(int $value): void
	{
		$this->writeNumber($value);
	}

	/**
	 * Write a string, without changing the type/style of the cell.
	 * Removes the formulas of the cell, if any.
	 * @throws XlsxFastEditorFileFormatException
	 * @throws XlsxFastEditorXmlException
	 */
	public function writeString(string $value): void
	{
		$v = self::initCellValue();
		$this->c->setAttribute('t', 's');	// Type shared string
		$sharedStringId = $this->editor->_makeNewSharedString($value);
		$v->nodeValue = (string)$sharedStringId;
		$this->editor->_touchWorksheet($this->sheetNumber);
	}

	/**
	 * Replace the hyperlink of the cell, if that cell already has an hyperlink.
	 * Warning: does not support the creation of a new hyperlink.
	 * @return bool True if the hyperlink could be replaced, false otherwise.
	 * @throws XlsxFastEditorFileFormatException
	 * @throws XlsxFastEditorXmlException
	 */
	public function writeHyperlink(string $value): bool
	{
		$xpath = $this->getXPath();
		$rId = $xpath->evaluate("normalize-space(/o:worksheet/o:hyperlinks/o:hyperlink[@ref='{$this->name()}'][1]/@or:id)");
		if (!is_string($rId) || $rId === '') {
			return false;
		}
		try {
			return $this->editor->_setHyperlink($this->sheetNumber, $rId, $value);
		} catch (\InvalidArgumentException $iax) {
			throw new XlsxFastEditorXmlException("Error querying XML fragment for hyperlink in cell {$this->name()}!", $iax->getCode(), $iax);
		}
	}
}
