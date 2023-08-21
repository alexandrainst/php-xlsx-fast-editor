<?php

namespace alexandrainst\XlsxFastEditor;

/**
 * Class used to return some read-only raw information about a cell.
 */
final class XlsxFastEditorCell
{
	private XlsxFastEditor $editor;
	private \DOMElement $c;

	public function __construct(XlsxFastEditor $editor, \DOMElement $c)
	{
		$this->editor = $editor;
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
				return new XlsxFastEditorCell($this->editor, $c);
			}
			$r = $this->c->previousElementSibling;
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
				return new XlsxFastEditorCell($this->editor, $c);
			}
			$c = $this->c->nextElementSibling;
		}
		return null;
	}

	/**
	 * Access the parent row of the cell.
	 */
	public function row(): XlsxFastEditorRow
	{
		$r = $this->c->parentNode;
		if (!($r instanceof \DOMElement)) {
			throw new XlsxFastEditorXmlException("Error querying XML row for cell {$this->name()}!");
		}
		return new XlsxFastEditorRow($this->editor, $r);
	}

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
			return $this->editor->getSharedString((int)$value);
		} else {
			// Local value
			return $value;
		}
	}
}
