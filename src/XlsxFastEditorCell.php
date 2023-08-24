<?php

namespace alexandrainst\XlsxFastEditor;

/**
 * Class used to return some read-only raw information about a cell.
 */
final class XlsxFastEditorCell
{
	private XlsxFastEditor $editor;
	private int $sheetNumber;
	private \DOMElement $c;

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
	 * Cell name (e.g., `'D4'`).
	 */
	public function name(): string
	{
		return $this->c->getAttribute('r');
	}

	/**
	 * Access the previous existing cell, if any, null otherwise.
	 */
	public function getPreviousCell(): ?XlsxFastEditorCell
	{
		$c = $this->c->previousElementSibling;
		while ($c !== null) {
			if ($c->localName === 'c') {
				return new XlsxFastEditorCell($this->editor, $this->sheetNumber, $c);
			}
			$c = $this->c->previousElementSibling;
		}
		return null;
	}

	/**
	 * Access the next existing cell, if any, null otherwise.
	 */
	public function getNextCell(): ?XlsxFastEditorCell
	{
		$c = $this->c->nextElementSibling;
		while ($c !== null) {
			if ($c->localName === 'r') {
				return new XlsxFastEditorCell($this->editor, $this->sheetNumber, $c);
			}
			$c = $this->c->nextElementSibling;
		}
		return null;
	}

	/**
	 * Access the parent row of the cell.
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
			if (!($v instanceof \DOMElement)) {
				throw new XlsxFastEditorXmlException("Error querying XML value for cell {$this->name()}!");
			}
			return $v->nodeValue;
		}
		return null;
	}

	/**
	 * Read the float value of the cell.
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
	 * Read the integer value of the cell.
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
	 * Read the string value of the cell,
	 * compatible with the shared string approach.
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
	 * Clean the cell to have its value written.
	 * @return \DOMElement The `<v>` value element of the provided cell, or null in case of error.
	 */
	private function initCellValue(): \DOMElement
	{
		$v = null;
		$this->c->removeAttribute('t');	// Remove type, if it exists
		for ($i = $this->c->childNodes->length - 1; $i >= 0; $i--) {
			// Remove all childs except <v>
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
			$v = $this->c->ownerDocument === null ? null : $this->c->ownerDocument->createElement('v');
			if ($v == false) {
				throw new XlsxFastEditorXmlException('Error creating value for cell!');
			}
			$this->c->appendChild($v);
		}

		return $v;
	}

	/**
	 * Write a formulat, without changing the type/style of the cell.
	 * Removes the formulas of the cell, if any.
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
			throw new XlsxFastEditorInputException("Internal error accessing cell {$this->name()}!");
		}
		$f = $dom->createElement('f', $value);
		$this->c->appendChild($f);

		$this->editor->_clearCalcChain();
		$this->editor->_touchWorksheet($this->sheetNumber);
	}

	/**
	 * Write a number, without changing the type/style of the cell.
	 * Removes the formulas of the cell, if any.
	 *
	 * @param int|float $value
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
	 */
	public function writeFloat(float $value): void
	{
		$this->writeNumber($value);
	}

	/**
	 * Write an integer, without changing the type/style of the cell.
	 * Removes the formulas of the cell, if any.
	 * @param int $value
	 */
	public function writeInt(int $value): void
	{
		$this->writeNumber($value);
	}

	/**
	 * Write a string, without changing the type/style of the cell.
	 * Removes the formulas of the cell, if any.
	 */
	public function writeString(string $value): void
	{
		$v = self::initCellValue();
		$this->c->setAttribute('t', 's');	// Type shared string
		$sharedStringId = $this->editor->_makeNewSharedString($value);
		$v->nodeValue = (string)$sharedStringId;
		$this->editor->_touchWorksheet($this->sheetNumber);
	}
}