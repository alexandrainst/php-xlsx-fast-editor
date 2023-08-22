<?php

assert_options(ASSERT_ACTIVE, true);
assert_options(ASSERT_BAIL, true);
assert_options(ASSERT_EXCEPTION, true);
assert_options(ASSERT_WARNING, true);

require(__DIR__ . '/../autoload.php');

use alexandrainst\XlsxFastEditor\XlsxFastEditor;
use alexandrainst\XlsxFastEditor\XlsxFastEditorException;

copy(__DIR__ . '/test.xlsx', __DIR__ . '/_copy.xlsx');

try {
	$xlsxFastEditor = new XlsxFastEditor(__DIR__ . '/_copy.xlsx');

	assert($xlsxFastEditor->getWorksheetCount() === 2);

	$sheet1 = $xlsxFastEditor->getWorksheetNumber('Sheet1');
	assert($sheet1 === 1);

	assert($xlsxFastEditor->readFloat($sheet1, 'D2') === 3.14159);
	assert($xlsxFastEditor->readFloat($sheet1, 'D4') === -1.0);
	assert($xlsxFastEditor->readFloat($sheet1, 'e5') === null);
	assert($xlsxFastEditor->readInt($sheet1, 'c3') === -5);
	assert($xlsxFastEditor->readInt($sheet1, 'F6') === null);
	assert($xlsxFastEditor->readString($sheet1, 'b4') === 'naïveté');
	assert($xlsxFastEditor->readString($sheet1, 'F7') === null);

	$sheet2 = $xlsxFastEditor->getWorksheetNumber('Sheet2');
	assert($sheet2 === 2);
	assert($xlsxFastEditor->getWorksheetName($sheet2) === 'Sheet2');

	assert($xlsxFastEditor->readFormula($sheet2, 'c2') === '=Sheet1!C2*2');
	assert($xlsxFastEditor->readFloat($sheet2, 'D2') === 3.14159 * 2);
	assert($xlsxFastEditor->readFloat($sheet2, 'D4') === -1.0 * 2);
	assert($xlsxFastEditor->readInt($sheet2, 'c3') === -5 * 2);
	assert($xlsxFastEditor->readString($sheet2, 'B3') === 'déjà-vu');

	// Navigation
	assert($xlsxFastEditor->getFirstRow($sheet1)?->number() === 1);
	assert($xlsxFastEditor->getRow($sheet1, 1)?->getFirstCell()?->name() === 'A1');
	assert($xlsxFastEditor->getRow($sheet1, 2)?->number() === 2);
	assert($xlsxFastEditor->getRow($sheet1, 3)?->getLastCell()?->name() === 'E3');
	assert($xlsxFastEditor->getRow($sheet1, 4)?->getCell('D4')?->name() === 'D4');
	assert($xlsxFastEditor->getLastRow($sheet1)?->number() === 4);

	// Iterators
	$nb = 0;
	foreach ($xlsxFastEditor->rowsIterator($sheet1) as $row) {
		assert($row->number() > 0);
		foreach ($row->cellsIterator() as $cell) {
			assert($cell->name() !== null);
			$nb++;
		}
	}
	assert($nb === 20);

	// Writing existing cells
	$xlsxFastEditor->writeFormula($sheet1, 'c2', '=2*3');
	$xlsxFastEditor->writeString($sheet1, 'b4', 'α');
	$xlsxFastEditor->writeInt($sheet1, 'c4', 15);
	$xlsxFastEditor->writeFloat($sheet1, 'd4', -66.6);

	// Writing existing cells with formulas
	$xlsxFastEditor->writeFormula($sheet2, 'c2', '=Sheet1!C2*3');
	$xlsxFastEditor->writeString($sheet2, 'B3', 'β');
	$xlsxFastEditor->writeInt($sheet2, 'C3', -7);
	$xlsxFastEditor->writeFloat($sheet2, 'D3', 273.15);

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

	$xlsxFastEditor->setFullCalcOnLoad($sheet1, true);

	$xlsxFastEditor->save();

	// Verify all the changes
	$xlsxFastEditor = new XlsxFastEditor(__DIR__ . '/_copy.xlsx');

	assert($xlsxFastEditor->readFormula($sheet1, 'c2') === '=2*3');
	assert($xlsxFastEditor->readString($sheet1, 'B4') === 'α');
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

	$xlsxFastEditor->close();

	// Verify by hand that the resulting file opens without warning in Microsoft Excel.
	// Verify by hand that the cell Sheet1!E4 has its formula recalculated (to -999) when opening in Excel.
	// unlink(__DIR__ . '/copy.xlsx');
} catch (XlsxFastEditorException $xlsxe) {
	die($xlsxe);
}
