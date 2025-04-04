<?php

declare(strict_types=1);

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

	/**
	 * @throws XlsxFastEditorXmlException
	 */
	private function getXPath(): \DOMXPath
	{
		if ($this->xpath === null) {
			$dom = $this->r->ownerDocument;
			if ($dom === null) {
				throw new XlsxFastEditorXmlException("Internal error accessing row {$this->number()}!");
			}
			$xpath = new \DOMXPath($dom);
			$xpath->registerNamespace('o', XlsxFastEditor::_OXML_NAMESPACE);
			$this->xpath = $xpath;
		}
		return $this->xpath;
	}

	/**
	 * Access the previous existing row, if any, `null` otherwise.
	 * ℹ️ This is a faster method than `XlsxFastEditor::getRow()`
	 */
	public function getPreviousRow(): ?XlsxFastEditorRow
	{
		$r = $this->r->previousElementSibling;
		while ($r instanceof \DOMElement) {
			if ($r->localName === 'row') {
				return new XlsxFastEditorRow($this->editor, $this->sheetNumber, $r);
			}
			$r = $r->previousElementSibling;
		}
		return null;
	}

	/**
	 * Access the next existing row, if any, `null` otherwise.
	 * ℹ️ This is a faster method than `XlsxFastEditor::getRow()`
	 */
	public function getNextRow(): ?XlsxFastEditorRow
	{
		$r = $this->r->nextElementSibling;
		while ($r instanceof \DOMElement) {
			if ($r->localName === 'row') {
				return new XlsxFastEditorRow($this->editor, $this->sheetNumber, $r);
			}
			$r = $r->nextElementSibling;
		}
		return null;
	}

	/**
	 * To iterate over all the existing cells of the row.
	 * ℹ️ This is a faster method than `XlsxFastEditorRow::getCellOrNull()`
	 * @return \Traversable<XlsxFastEditorCell>
	 */
	public function cellsIterator(): \Traversable
	{
		$c = $this->r->firstElementChild;
		while ($c instanceof \DOMElement) {
			if ($c->localName === 'c') {
				yield new XlsxFastEditorCell($this->editor, $this->sheetNumber, $c);
			}
			$c = $c->nextElementSibling;
		}
	}

	/**
	 * Get the first existing cell for a given line.
	 * ℹ️ This is a faster method than `XlsxFastEditorRow::getCellOrNull()`
	 * @return XlsxFastEditorCell|null The first cell of the given line if it exists, `null` otherwise.
	 */
	public function getFirstCell(): ?XlsxFastEditorCell
	{
		$c = $this->r->firstElementChild;
		while ($c instanceof \DOMElement) {
			if ($c->localName === 'c') {
				return new XlsxFastEditorCell($this->editor, $this->sheetNumber, $c);
			}
			$c = $c->nextElementSibling;
		}
		return null;
	}

	/**
	 * Get the cell of the given name if it exists.
	 *
	 * ℹ️ Instead of calling multiple times this function, consider the faster navigation methods
	 * `XlsxFastEditorRow::cellsIterator()`, `XlsxFastEditorRow::getFirstCell()`, `XlsxFastEditorCell::getNextCell()`, etc.
	 *
	 * @param string $cellName Column name such as `'B'` or full cell name such as `'B4'`
	 * @param int $accessMode To control the behaviour when the cell does not exist:
	 * set to `XlsxFastEditor::ACCESS_MODE_NULL` to return `null` (default),
	 * set to `XlsxFastEditor::ACCESS_MODE_EXCEPTION` to raise an `XlsxFastEditorInputException` exception,
	 * set to `XlsxFastEditor::ACCESS_MODE_AUTOCREATE` to auto-create the cell.
	 * @return XlsxFastEditorCell|null A cell, potentially `null` if the cell does not exist and `$accessMode` is set to `XlsxFastEditor::ACCESS_MODE_NULL`
	 * @phpstan-return ($accessMode is XlsxFastEditor::ACCESS_MODE_NULL ? XlsxFastEditorCell|null : XlsxFastEditorCell)
	 * @internal
	 * @throws \InvalidArgumentException if `$cellName` has an invalid format
	 * @throws XlsxFastEditorInputException optionally if the corresponding cell does not exist, depending on choice of `$accessMode`
	 * @throws XlsxFastEditorXmlException
	 */
	public function getCell(string $cellName, int $accessMode = XlsxFastEditor::ACCESS_MODE_NULL): ?XlsxFastEditorCell
	{
		if (ctype_alpha($cellName)) {
			$cellName .= $this->number();
		} elseif (!ctype_alnum($cellName)) {
			throw new \InvalidArgumentException("Invalid cell reference {$cellName}!");
		}
		$cellName = strtoupper($cellName);

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
			$rowNumber = (int)preg_replace('/[^\d]+/', '', $cellName);
			if ($rowNumber !== $this->number()) {
				throw new \InvalidArgumentException("Invalid line in cell reference {$cellName} for line {$this->number()}!");
			}

			$dom = $xpath->document;
			try {
				$c = $dom->createElement('c');
			} catch (\DOMException $dex) {
				throw new XlsxFastEditorXmlException("Error creating cell {$this->sheetNumber}/{$cellName}!", $dex->code, $dex);
			}
			if ($c === false) {
				throw new XlsxFastEditorXmlException("Error creating cell {$this->sheetNumber}/{$cellName}!");
			}
			$c->setAttribute('r', $cellName);

			// Excel expects the cells to be sorted
			$sibling = $this->r->firstElementChild;
			while ($sibling instanceof \DOMElement && XlsxFastEditor::cellOrderCompare($sibling->getAttribute('r'), $cellName) < 0) {
				$sibling = $sibling->nextElementSibling;
			}
			if ($sibling instanceof \DOMElement) {
				$this->r->insertBefore($c, $sibling);
			} else {
				$this->r->appendChild($c);
			}
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
	 * Get the cell of the given name, or `null` if if does not exist.
	 *
	 * ℹ️ Instead of calling multiple times this function, consider the faster navigation methods
	 * `XlsxFastEditorRow::cellsIterator()`, `XlsxFastEditorRow::getFirstCell()`, `XlsxFastEditorCell::getNextCell()`, etc.
	 *
	 * @param string $cellName Column name such as `'B'` or full cell name such as `'B4'`
	 * @return XlsxFastEditorCell|null A cell, potentially `null` if the cell does not exist
	 * @throws \InvalidArgumentException if `$cellName` has an invalid format
	 * @throws XlsxFastEditorXmlException
	 */
	public function getCellOrNull(string $cellName): ?XlsxFastEditorCell
	{
		try {
			return $this->getCell($cellName, XlsxFastEditor::ACCESS_MODE_NULL);
		} catch (XlsxFastEditorInputException $iex) {
			// Will not happen
			return null;
		}
	}

	/**
	 * Get the cell of the given name, or autocreate it if it does not already exist.
	 * @param string $cellName Column name such as `'B'` or full cell name such as `'B4'`
	 * @return XlsxFastEditorCell A cell
	 * @throws \InvalidArgumentException if `$cellName` has an invalid format
	 * @throws XlsxFastEditorXmlException
	 */
	public function getCellAutocreate(string $cellName): XlsxFastEditorCell
	{
		try {
			return $this->getCell($cellName, XlsxFastEditor::ACCESS_MODE_AUTOCREATE);
		} catch (XlsxFastEditorInputException $iex) {
			// Will not happen
			throw new XlsxFastEditorXmlException('Internal error with getCell!', $iex->getCode(), $iex);
		}
	}

	/**
	 * Get the last existing cell for a given line.
	 * @return XlsxFastEditorCell|null The last cell of the given line if it exists, `null` otherwise.
	 */
	public function getLastCell(): ?XlsxFastEditorCell
	{
		$c = $this->r->lastElementChild;
		while ($c instanceof \DOMElement) {
			if ($c->localName === 'c') {
				return new XlsxFastEditorCell($this->editor, $this->sheetNumber, $c);
			}
			$c = $c->previousElementSibling;
		}
		return null;
	}
}
