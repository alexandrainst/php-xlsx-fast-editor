<?php

namespace alexandrainst\XlsxFastEditor;

/**
 * Class used to return some read-only raw information about a row.
 */
final class XlsxFastEditorRow
{
	private XlsxFastEditor $editor;
	private int $sheetNumber;
	private \DOMElement $r;
	private ?\DOMXPath $xpath = null;

	/**
	 * @internal
	 */
	public function __construct(XlsxFastEditor $editor, int $sheetNumber, \DOMElement $r)
	{
		$this->editor = $editor;
		$this->sheetNumber = $sheetNumber;
		$this->r = $r;
	}

	/**
	 * Row number (ID).
	 * Warning: this is not an index.
	 */
	public function number(): int
	{
		return (int)$this->r->getAttribute('r');
	}

	private function getXPath(): \DOMXPath
	{
		if ($this->xpath === null) {
			$dom = $this->r->ownerDocument;
			if ($dom === null) {
				throw new XlsxFastEditorInputException("Internal error accessing row {$this->number()}!");
			}
			$xpath = new \DOMXPath($dom);
			$xpath->registerNamespace('o', XlsxFastEditor::_OXML_NAMESPACE);
			$this->xpath = $xpath;
		}
		return $this->xpath;
	}

	/**
	 * Access the previous existing row, if any, null otherwise.
	 */
	public function getPreviousRow(): ?XlsxFastEditorRow
	{
		$r = $this->r->previousElementSibling;
		while ($r !== null) {
			if ($r->localName === 'r') {
				return new XlsxFastEditorRow($this->editor, $this->sheetNumber, $r);
			}
			$r = $this->r->previousElementSibling;
		}
		return null;
	}

	/**
	 * Access the next existing row, if any, null otherwise.
	 */
	public function getNextRow(): ?XlsxFastEditorRow
	{
		$r = $this->r->nextElementSibling;
		while ($r !== null) {
			if ($r->localName === 'r') {
				return new XlsxFastEditorRow($this->editor, $this->sheetNumber, $r);
			}
			$r = $this->r->nextElementSibling;
		}
		return null;
	}

	/**
	 * To iterate over all the existing cells of the row.
	 * @return \Traversable<XlsxFastEditorCell>
	 */
	public function cellsIterator(): \Traversable
	{
		$c = $this->r->firstElementChild;
		while ($c !== null) {
			if ($c->localName === 'c') {
				yield new XlsxFastEditorCell($this->editor, $this->sheetNumber, $c);
			}
			$c = $c->nextElementSibling;
		}
	}

	/**
	 * Get the first existing cell for a given line.
	 * @return XlsxFastEditorCell|null The first cell of the given line if it exists, null otherwise.
	 */
	public function getFirstCell(): ?XlsxFastEditorCell
	{
		$c = $this->r->firstElementChild;
		while ($c !== null) {
			if ($c->localName === 'c') {
				return new XlsxFastEditorCell($this->editor, $this->sheetNumber, $c);
			}
			$c = $c->nextElementSibling;
		}
		return null;
	}

	/**
	 * Get the cell of the given name if it exists.
	 * @param $cellName Cell name such as `B4`
	 * @param int $accessMode To control the behaviour when the cell does not exist:
	 * set to `XlsxFastEditor::ACCESS_MODE_NULL` to return `null` (default),
	 * set to `XlsxFastEditor::ACCESS_MODE_EXCEPTION` to raise an `XlsxFastEditorInputException` exception,
	 * set to `XlsxFastEditor::ACCESS_MODE_AUTOCREATE` to auto-create the cell.
	 * @return XlsxFastEditorCell|null A cell, potentially `null` if the cell does not exist and `$accessMode` is set to `XlsxFastEditor::ACCESS_MODE_NULL`
	 * @phpstan-return ($accessMode is XlsxFastEditor::ACCESS_MODE_NULL ? XlsxFastEditorCell|null : XlsxFastEditorCell)
	 */
	public function getCell(string $cellName, int $accessMode = XlsxFastEditor::ACCESS_MODE_NULL): ?XlsxFastEditorCell
	{
		$xpath = $this->getXPath();
		$cs = $xpath->query("./o:c[@r='$cellName'][1]", $this->r);
		$c = null;
		if ($cs !== false && $cs->length > 0) {
			$c = $cs[0];
			if (!($c instanceof \DOMElement)) {
				throw new XlsxFastEditorXmlException("Error querying XML fragment for cell {$this->number()}/{$cellName}!");
			}
		}

		if ($c === null && $accessMode === XlsxFastEditor::ACCESS_MODE_AUTOCREATE) {
			// The cell <c> was not found
			$dom = $xpath->document;
			$c = $dom->createElement('c');
			if ($c === false) {
				throw new XlsxFastEditorXmlException("Error creating cell {$this->sheetNumber}/{$cellName}!");
			}
			$c->setAttribute('r', $cellName);

			// Excel expects the cells to be sorted
			$sibling = $this->r->firstElementChild;
			while ($sibling !== null && XlsxFastEditor::cellOrderCompare($sibling->getAttribute('r'), $cellName) < 0) {
				$sibling = $sibling->nextElementSibling;
			}
			$this->r->insertBefore($c, $sibling);
		}

		if ($c === null) {
			if ($accessMode === XlsxFastEditor::ACCESS_MODE_EXCEPTION) {
				throw new XlsxFastEditorInputException("Cell {$this->sheetNumber}/{$cellName} not found!");
			}
			return null;
		}

		return new XlsxFastEditorCell($this->editor, $this->sheetNumber, $c);
	}

	/**
	 * Get the last existing cell for a given line.
	 * @return XlsxFastEditorCell|null The last cell of the given line if it exists, null otherwise.
	 */
	public function getLastCell(): ?XlsxFastEditorCell
	{
		$c = $this->r->lastElementChild;
		while ($c !== null) {
			if ($c->localName === 'c') {
				return new XlsxFastEditorCell($this->editor, $this->sheetNumber, $c);
			}
			$c = $c->previousElementSibling;
		}
		return null;
	}
}
