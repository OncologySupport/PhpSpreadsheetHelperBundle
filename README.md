# PhpSpreadsheet Helper Bundle

`PhpSpreadsheet Helper Bundle` eases the use of the excellent
[PHPOffice/PhpSpreadsheet bundle](https://https://github.com/PHPOffice/PhpSpreadsheet)
in your Symfony application by adding helper methods

## Documentation

## What's new in version 2

WorksheetFactory has been re-factored to WorksheetUtils to allow the helper methods to be available on 
worksheets beyond the first tab of the spreadsheet. See the example below.

## License

PhpSpreadsheet Helper Bundle is released under the MIT License. See the bundled [LICENSE](LICENSE) file for details.

## Installation

Applications that use Symfony Flex
----------------------------------

Open a command console, enter your project directory and execute:

```console
composer require oncology-support/phpspreadsheet-helper-bundle
```

Applications that don't use Symfony Flex
----------------------------------------

### Step 1: Download the Bundle

Open a command console, enter your project directory and execute the
following command to download the latest stable version of this bundle:

```console
composer require oncology-support/phpspreadsheet-helper-bundle
```

### Step 2: Enable the Bundle

Then, enable the bundle by adding it to the list of registered bundles
in the `config/bundles.php` file of your project:

```php
// config/bundles.php

return [
    // ...
    OncologySupport\PhpSpreadsheetHelper\OncologySupportPhpSpreadsheetHelperBundle::class => ['all' => true],
];
```

### Step 3: Use it!

This bundle provides helper functions that you call from your controller.

```php
// your controller

use OncologySupport\PhpSpreadsheetHelper\Utilities\SpreadsheetFactory;
use OncologySupport\PhpSpreadsheetHelper\Utilities\WorksheetUtils;

public function createAndDownloadSpreadsheet()
{
    $spreadsheet = new SpreadsheetFactory();
    $worksheet = new WorksheetUtils($spreadsheet->getActiveSheet(), 'First Tab');
    
    $worksheet->addTitleRow('Very Important Data');
    
    $header = ['Order Date', 'Site', 'Item', 'Item Description'];
    $rows = [
        ['12/15/2023', 'Main site', 'Dog collar', 'A stunning red dog collar'],
        ['12/19/2023', 'Main site', 'Cat collar', 'Blue cat collar'],
        ['1/3/2024', 'Backup site', 'Turtle harness', 'Take your turtle for a walk!'],
    ];
    
    $worksheet->addDataRows[$rows];
    
    // add another tab on the spreadsheet
    $worksheet2 = new WorksheetUtils($spreadsheet->createSheet(), 'Second Tab');
    $spreadsheet->setActiveSheetIndex(1);

    $worksheet2->addTitleRow('Even More Important Data');
    
    $header = ['Order Date', 'Site', 'Item', 'Item Description'];
    $rows = [
        ['2/12/2024', 'Main site', 'Donkey collar', 'Big enough for the most stubborn donk'],
        ['4/3/2024', 'Backup site', 'Duck harness', 'Take your duck for a walk!'],
    ];
    
    $worksheet2->addDataRows[$rows];
    
    // download to an excel file    
    $spreadsheet->outputToWeb();
}

Enjoy!
