<?php

namespace alexandrainst\XlsxFastEditor;

/**
 * Class used to return some read-only raw information about a row.
 */
final class XlsxFastEditorRow
{
	private XlsxFastEditor $editor;
	private \DOMElement $r;

	public function __construct(XlsxFastEditor $editor, \DOMElement $r)
	{
		$this->editor = $editor;
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

	/**
	 * Access the previous existing row, if any, null otherwise.
	 */
	public function getPreviousRow(): ?XlsxFastEditorRow
	{
		$r = $this->r->previousElementSibling;
		while ($r !== null) {
			if ($r->localName === 'r') {
				return new XlsxFastEditorRow($this->editor, $r);
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
				return new XlsxFastEditorRow($this->editor, $r);
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
				yield new XlsxFastEditorCell($this->editor, $c);
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
				return new XlsxFastEditorCell($this->editor, $c);
			}
			$c = $c->nextElementSibling;
		}
		return null;
	}

	/**
	 * Get the cell of the given name.
	 * @param $cellName Cell name such as `B4`
	 */
	public function getCell(string $cellName): ?XlsxFastEditorCell
	{
		$dom = $this->r->ownerDocument;
		if ($dom === null) {
			throw new XlsxFastEditorInputException("Internal error accessing cell {$this->number()}/{$cellName}!");
		}
		$xpath = new \DOMXPath($dom);
		$xpath->registerNamespace('o', XlsxFastEditor::OXML_NAMESPACE);
		$cs = $xpath->query("./o:c[@r='$cellName'][1]", $this->r);
		if ($cs !== false && $cs->length > 0) {
			$c = $cs[0];
			if (!($c instanceof \DOMElement)) {
				throw new XlsxFastEditorXmlException("Error querying XML fragment for cell {$this->number()}/{$cellName}!");
			}
			return new XlsxFastEditorCell($this->editor, $c);
		}
		return null;
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
				return new XlsxFastEditorCell($this->editor, $c);
			}
			$c = $c->previousElementSibling;
		}
		return null;
	}
}
