# PhpSpreadsheet Helper Bundle

`PhpSpreadsheet Helper Bundle` eases the use of the excellent
[PHPOffice/PhpSpreadsheet bundle](https://https://github.com/PHPOffice/PhpSpreadsheet)
in your Symfony application by adding helper methods

## Documentation

## License

PhpSpreadsheet Helper Bundle is released under the MIT License. See the bundled [LICENSE](LICENSE) file for details.

## Installation

Applications that use Symfony Flex
----------------------------------

Open a command console, enter your project directory and execute:

```console
$ composer require oncology-support/phpspreadsheet-helper-bundle
```

Applications that don't use Symfony Flex
----------------------------------------

### Step 1: Download the Bundle

Open a command console, enter your project directory and execute the
following command to download the latest stable version of this bundle:

```console
$ composer require oncology-support/phpspreadsheet-helper-bundle
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

... to be completed

Enjoy!
