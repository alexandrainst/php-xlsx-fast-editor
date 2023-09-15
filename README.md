# php-xlsx-fast-editor

PHP library to make basic but fast read & write operations on existing Excel workbooks.

It handles XLSX/XLSM documents (Microsoft Excel 2007+, Office Open XML Workbook) using fast and simple low-level ZIP & XML manipulations,
without requiring any library dependency, while minimising unintended side-effects.

## Rationale

If you need advanced manipulation of Excel documents such as working with styles,
check the [PhpSpreadsheet](https://github.com/PHPOffice/PhpSpreadsheet) library
(previously [PHPExcel](https://github.com/PHPOffice/PHPExcel/)),
but for simply reading & writing basic values from existing Excel workbooks, `PhpSpreadsheet` is over an order of magnitude too slow,
and has the risk of breaking some unsupported Excel features such as some notes and charts.

There are also libraries to create new Excel documents from scratch, or for just reading some values, but not any obvious one for editing.

`php-xlsx-fast-editor` addresses the need of quickly reading & writing & editing existing Excel documents,
while reducing the risk of breaking anything.

Note that to create a new document, you can just provide a blank Excel document as input.

## Use

Via [Composer](https://packagist.org/packages/alexandrainst/php-xlsx-fast-editor):

```sh
composer require alexandrainst/php-xlsx-fast-editor
```

or manually:

```php
require 'vendor/alexandrainst/XlsxFastEditor/autoload.php';
```


## Examples

```php
<?php

use alexandrainst\XlsxFastEditor\XlsxFastEditor;
use alexandrainst\XlsxFastEditor\XlsxFastEditorException;

try {
	$xlsxFastEditor = new XlsxFastEditor('test.xlsx');

	// Workbook / worksheet methods
	$nbWorksheets = $xlsxFastEditor->getWorksheetCount();
	$worksheetName = $xlsxFastEditor->getWorksheetName(1);
	$worksheetId1 = $xlsxFastEditor->getWorksheetNumber('Sheet1');
	// If you want to force Excel to recalculate formulas on next load:
	$xlsxFastEditor->setFullCalcOnLoad($worksheetId1, true);

	// Direct read/write access
	$fx = $xlsxFastEditor->readFormula($worksheetId1, 'A1');
	$f = $xlsxFastEditor->readFloat($worksheetId1, 'B2');
	$i = $xlsxFastEditor->readInt($worksheetId1, 'C3');
	$s = $xlsxFastEditor->readString($worksheetId1, 'D4');
	$h = $xlsxFastEditor->readHyperlink($worksheetId1, 'B4');
	$d = $xlsxFastEditor->readDateTime($worksheetId1, 'F4');
	$xlsxFastEditor->deleteRow($worksheetId1, 5);
	$xlsxFastEditor->writeFormula($worksheetId1, 'A1', '=B2*3');
	$xlsxFastEditor->writeFloat($worksheetId1, 'B2', 3.14);
	$xlsxFastEditor->writeInt($worksheetId1, 'C3', 13);
	$xlsxFastEditor->writeString($worksheetId1, 'D4', 'Hello');
	$xlsxFastEditor->writeHyperlink($sheet1, 'B4', 'https://example.net/');	// Only for cells with an existing hyperlink

	// Read as array
	$table = $xlsxFastEditor->readArray($sheet1);
	$s = $table['B'][2];

	$table = $xlsxFastEditor->readArrayWithHeaders($sheet1);
	$s = $table['columnName'][2];

	// Regex search & replace operating globally on all the worksheets:
	$xlsxFastEditor->textReplace('/Hello/i', 'World');

	// Navigation methods for existing rows
	$row = $xlsxFastEditor->getFirstRow($worksheetId1);
	$row = $xlsxFastEditor->getRow($worksheetId1, 2);
	$row = $row->getPreviousRow();
	$row = $row->getNextRow();
	$row = $xlsxFastEditor->getLastRow($worksheetId1);

	// Methods for rows
	$rowNumber = $row->number();

	// Navigation methods for existing cells
	$cell = $row->getFirstCell();
	$cell = $row->getCellOrNull('D4');
	$cell = $cell->getPreviousCell();
	$cell = $cell->getNextCell();
	$cell = $row->getLastCell();

	// Methods for cells
	$cellName = $cell->name();
	$columnName = $cell->column();
	$fx = $cell->readFormula();
	$f = $cell->readFloat();
	$i = $cell->readInt();
	$s = $cell->readString();
	$h = $cell->readHyperlink();
	$d = $cell->readDateTime();
	$cell->writeFormula('=B2*3');
	$cell->writeFloat(3.14);
	$cell->writeInt(13);
	$cell->writeString('Hello');
	$cell->writeHyperlink('https://example.net/');	// Only for cells with an existing hyperlink

	// Iterators for existing rows and cells
	foreach ($xlsxFastEditor->rowsIterator($worksheetId1) as $row) {
		foreach ($row->cellsIterator() as $cell) {
			// $cell->...
		}
	}

	$xlsxFastEditor->save();
	// If you do not want to save, call `close()` instead:
	// $xlsxFastEditor->close();
} catch (XlsxFastEditorException $xlsxe) {
	die($xlsxe->getMessage());
}
```

## Requirements

PHP 7.4+ with ZIP and XML extensions.
Check [`composer.json`](./composer.json) for details.

## Credits

Originally written by [Alexandre Alapetite](https://github.com/Alkarex) for the [Alexandra Institute](https://alexandra.dk/), 2023.
License [GNU AGPL](https://gnu.org/licenses/agpl.html).
