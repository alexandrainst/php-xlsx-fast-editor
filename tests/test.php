<?php

declare(strict_types=1);

assert_options(ASSERT_ACTIVE, true);	// phpcs:ignore Generic.PHP.DeprecatedFunctions.Deprecated
assert_options(ASSERT_BAIL, true);	// phpcs:ignore Generic.PHP.DeprecatedFunctions.Deprecated
assert_options(ASSERT_EXCEPTION, true);	// phpcs:ignore Generic.PHP.DeprecatedFunctions.Deprecated
assert_options(ASSERT_WARNING, true);	// phpcs:ignore Generic.PHP.DeprecatedFunctions.Deprecated

require(__DIR__ . '/../autoload.php');

use alexandrainst\XlsxFastEditor\XlsxFastEditor;
use alexandrainst\XlsxFastEditor\XlsxFastEditorException;

copy(__DIR__ . '/test.xlsx', __DIR__ . '/_copy.xlsx');

try {
	$xlsxFastEditor = new XlsxFastEditor(__DIR__ . '/_copy.xlsx');

	assert($xlsxFastEditor->getWorksheetCount() === 3);

	date_default_timezone_set('UTC');
	assert($xlsxFastEditor->getWorkbookDateSystem() === 1900);
	assert(XlsxFastEditor::excelDateToDateTime(0.5, 1900)->format('c') === '1900-01-01T12:00:00+00:00');
	assert(XlsxFastEditor::excelDateToDateTime(32, 1900)->format('c') === '1900-02-01T00:00:00+00:00');
	assert(XlsxFastEditor::excelDateToDateTime(44865, 1904)->format('c') === '2026-11-01T00:00:00+00:00');

	$sheet1 = $xlsxFastEditor->getWorksheetNumber('Sheet1');
	assert($sheet1 === 3);

	assert($xlsxFastEditor->deleteRow($sheet1, 5) === true);

	assert($xlsxFastEditor->readFloat($sheet1, 'D2') === 3.14159);
	assert($xlsxFastEditor->readFloat($sheet1, 'D4') === -1.0);
	assert($xlsxFastEditor->readFloat($sheet1, 'e5') === null);
	assert($xlsxFastEditor->readInt($sheet1, 'c3') === -5);
	assert($xlsxFastEditor->readInt($sheet1, 'F6') === null);
	assert($xlsxFastEditor->readString($sheet1, 'B3') === 'déjà-vu');
	assert($xlsxFastEditor->readString($sheet1, 'b4') === 'naïveté');
	assert($xlsxFastEditor->readString($sheet1, 'F7') === null);
	assert($xlsxFastEditor->readHyperlink($sheet1, 'B4') === 'https://example.net/');
	assert($xlsxFastEditor->readHyperlink($sheet1, 'C3') === null);

	$sheet2 = $xlsxFastEditor->getWorksheetNumber('Sheet2');
	assert($xlsxFastEditor->getWorksheetName($sheet2) === 'Sheet2');

	assert($xlsxFastEditor->readFormula($sheet2, 'c2') === '=Sheet1!C2*2');
	assert($xlsxFastEditor->readFloat($sheet2, 'D2') === 3.14159 * 2);
	assert($xlsxFastEditor->readFloat($sheet2, 'D4') === -1.0 * 2);
	assert($xlsxFastEditor->readInt($sheet2, 'c3') === -5 * 2);
	assert($xlsxFastEditor->readString($sheet2, 'B3') === 'déjà-vu');

	assert($xlsxFastEditor->readDateTime($sheet1, 'F2')?->format('c') === '1980-11-24T00:00:00+00:00');
	assert($xlsxFastEditor->readDateTime($sheet1, 'F3')?->format('c') === '1980-11-24T10:20:30+00:00');
	assert($xlsxFastEditor->readDateTime($sheet1, 'F4')?->format('c') === '1900-01-01T10:20:30+00:00');

	assert($xlsxFastEditor->readArray($sheet1)['B'][2] === 'Hello');
	assert($xlsxFastEditor->readArrayWithHeaders($sheet1)['Strings'][2] === 'Hello');

	// Navigation
	assert($xlsxFastEditor->getFirstRow($sheet1)?->number() === 1);
	assert($xlsxFastEditor->getRow($sheet1, 1)?->getFirstCell()?->name() === 'A1');
	assert($xlsxFastEditor->getRow($sheet1, 2)?->number() === 2);
	assert($xlsxFastEditor->getRow($sheet1, 3)?->getLastCell()?->name() === 'F3');
	assert($xlsxFastEditor->getLastRow($sheet1)?->number() === 4);

	$sheet3 = $xlsxFastEditor->getWorksheetNumber('Sheet3');
	assert($xlsxFastEditor->getWorksheetName($sheet3) === 'Sheet3');
	assert($xlsxFastEditor->getHighestColumnName($sheet3) === 'G');

	$row4 = $xlsxFastEditor->getRow($sheet1, 4);
	assert($row4 !== null);
	assert($row4->getPreviousRow()?->getNextRow()?->number() === 4);
	assert($row4->getCellOrNull('D4')?->name() === 'D4');
	assert($row4->getCellOrNull('d4')?->name() === 'D4');
	assert($row4->getCellOrNull('D')?->name() === 'D4');
	$ex = null;
	try {
		assert($row4->getCellAutocreate('D5')->name() === 'D5');
	} catch (\InvalidArgumentException $aex) {
		$ex = $aex;
	}
	assert($ex instanceof \InvalidArgumentException);

	$cellD4 = $row4->getCell('D4');
	assert($cellD4 !== null);
	assert($cellD4->getPreviousCell()?->getNextCell()?->name() === 'D4');

	assert(XlsxFastEditor::cellOrderCompare('B3', 'AA23') < 0);
	assert(XlsxFastEditor::cellOrderCompare('AA23', 'AB23') < 0);
	assert(XlsxFastEditor::cellOrderCompare('BB22', 'BB123') < 0);
	assert(XlsxFastEditor::cellOrderCompare('AA23', 'AA23') === 0);
	assert(XlsxFastEditor::cellOrderCompare('AA23', 'B3') > 0);
	assert(XlsxFastEditor::cellOrderCompare('AB23', 'AA23') > 0);
	assert(XlsxFastEditor::cellOrderCompare('BB123', 'BB22') > 0);

	// Iterators
	$nb = 0;
	foreach ($xlsxFastEditor->rowsIterator($sheet1) as $row) {
		assert($row->number() > 0);
		foreach ($row->cellsIterator() as $cell) {
			assert($cell->name() !== null);
			$nb++;
		}
	}
	assert($nb === 24);

	// Writing existing cells
	$xlsxFastEditor->writeFormula($sheet1, 'c2', '=2*3');
	$xlsxFastEditor->writeString($sheet1, 'b4', 'α');
	$xlsxFastEditor->writeHyperlink($sheet1, 'B4', 'https://example.org/');
	assert($xlsxFastEditor->writeHyperlink($sheet1, 'C3', 'https://example.org/') === false);
	$xlsxFastEditor->writeInt($sheet1, 'c4', 15);
	$xlsxFastEditor->writeFloat($sheet1, 'd4', -66.6);

	// Writing existing cells with formulas
	$xlsxFastEditor->writeFormula($sheet2, 'c2', '=Sheet1!C2*3');
	$xlsxFastEditor->writeString($sheet2, 'B3', 'β');
	$xlsxFastEditor->writeInt($sheet2, 'C3', -7);
	$xlsxFastEditor->writeFloat($sheet2, 'D3', 273.15);

	// Writing special XML characters
	$xlsxFastEditor->writeString($sheet2, 'B5', '< " & \' >');
	$xlsxFastEditor->writeFormula($sheet2, 'C5', '=LEN("< & \' >")');

	// Writing non-existing cells but existing lines
	$xlsxFastEditor->writeFormula($sheet2, 'I2', '=7*3');
	$xlsxFastEditor->writeString($sheet2, 'F2', 'γ');
	$xlsxFastEditor->writeInt($sheet2, 'G3', -7);
	$xlsxFastEditor->writeFloat($sheet2, 'H4', 273.15);

	// Writing non-existing lines
	$xlsxFastEditor->writeFormula($sheet2, 'E11', '=7*5');
	$xlsxFastEditor->writeString($sheet2, 'B10', 'δ');
	$xlsxFastEditor->writeInt($sheet2, 'C9', 13);
	$xlsxFastEditor->writeFloat($sheet2, 'D10', -273.15);

	// Regex
	assert($xlsxFastEditor->textReplace('/Hello/i', 'World') > 0);

	assert($xlsxFastEditor->getFullCalcOnLoad($sheet1) == null);
	$xlsxFastEditor->setFullCalcOnLoad($sheet1, true);
	assert($xlsxFastEditor->getFullCalcOnLoad($sheet1) === true);

	$xlsxFastEditor->save();

	// Verify all the changes
	$xlsxFastEditor = new XlsxFastEditor(__DIR__ . '/_copy.xlsx');

	assert($xlsxFastEditor->readFormula($sheet1, 'c2') === '=2*3');
	assert($xlsxFastEditor->readString($sheet1, 'B4') === 'α');
	assert($xlsxFastEditor->readHyperlink($sheet1, 'B4') === 'https://example.org/');
	assert($xlsxFastEditor->readInt($sheet1, 'C4') === 15);
	assert($xlsxFastEditor->readFloat($sheet1, 'D4') === -66.6);

	assert($xlsxFastEditor->readFormula($sheet2, 'c2') === '=Sheet1!C2*3');
	assert($xlsxFastEditor->readString($sheet2, 'B3') === 'β');
	assert($xlsxFastEditor->readInt($sheet2, 'C3') === -7);
	assert($xlsxFastEditor->readFloat($sheet2, 'D3') === 273.15);

	assert($xlsxFastEditor->readFormula($sheet2, 'I2') === '=7*3');
	assert($xlsxFastEditor->readString($sheet2, 'F2') === 'γ');
	assert($xlsxFastEditor->readInt($sheet2, 'G3') === -7);
	assert($xlsxFastEditor->readFloat($sheet2, 'H4') === 273.15);

	assert($xlsxFastEditor->readFormula($sheet2, 'E11') === '=7*5');
	assert($xlsxFastEditor->readString($sheet2, 'B10') === 'δ');
	assert($xlsxFastEditor->readInt($sheet2, 'C9') === 13);
	assert($xlsxFastEditor->readFloat($sheet2, 'D10') === -273.15);

	assert($xlsxFastEditor->readString($sheet1, 'B2') === 'World');

	// Test special XML characters
	assert($xlsxFastEditor->readString($sheet2, 'B5') === '< " & \' >');
	assert($xlsxFastEditor->readFormula($sheet2, 'C5') === '=LEN("< & \' >")');

	$xlsxFastEditor->close();

	// Verify by hand that the resulting file opens without warning in Microsoft Excel.
	// Verify by hand that the cell Sheet1!E4 has its formula recalculated (to -999) when opening in Excel.
	// unlink(__DIR__ . '/_copy.xlsx');
} catch (XlsxFastEditorException $xlsxe) {
	die($xlsxe);
}
