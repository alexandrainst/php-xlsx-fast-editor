# php-xlsx-fast-editor

PHP library to make basic but fast read & write operations on existing Excel workbooks.

It handles XLSX/XLSM documents (Microsoft Excel 2007+, Office Open XML Workbook) using fast and simple low-level ZIP & XML manipulations,
without requiring any library dependency.

## Rationale

If you need advanced manipulation of Excel documents such as working with formulas and styles,
check the [PhpSpreadsheet](https://github.com/PHPOffice/PhpSpreadsheet) library
(previously [PHPExcel](https://github.com/PHPOffice/PHPExcel/)),
but for simply reading & writing basic values from existing Excel workbooks, `PhpSpreadsheet` is over an order of magnitude too slow.

There are also libraries to create new Excel documents from scratch, or for just reading some values, but not any obvious one for editing.

`php-xlsx-fast-editor` addresses the need of quickly reading & writing & editing existing Excel documents.

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

$xlsxFastEditor = new XlsxFastEditor('test.xlsx');

$worksheetId1 = $xlsxFastEditor->getWorksheetNumber('Sheet1');
$worksheetId2 = $xlsxFastEditor->getWorksheetNumber('Sheet2');

$f = $xlsxFastEditor->readFloat($worksheetId1, 'B2');
$i = $xlsxFastEditor->readInt($worksheetId1, 'C3');
$s = $xlsxFastEditor->readString($worksheetId2, 'D4');

// If you want to force Excel to recalculate formulas on next load:
$xlsxFastEditor->setFullCalcOnLoad($worksheetId2, true);

$xlsxFastEditor->writeFloat($worksheetId1, 'B2', 3.14);
$xlsxFastEditor->writeInt($worksheetId1, 'C3', 13);
$xlsxFastEditor->writeString($worksheetId2, 'D4', 'Hello');

// Regex search & replace operating globally on all the worksheets:
$xlsxFastEditor->textReplace('/Hello/i', 'World');

$xlsxFastEditor->save();
// If you do not want to save, call `close()` instead:
// $xlsxFastEditor->close();
```

## Requirements

PHP 7.4+ with ZIP and XML extensions.
Check [`composer.json`](./composer.json) for details.

## Credits

Originally written by [Alexandre Alapetite](https://github.com/Alkarex) for the [Alexandra Institute](https://alexandra.dk/), 2023.
License [GNU AGPL](https://gnu.org/licenses/agpl.html).
