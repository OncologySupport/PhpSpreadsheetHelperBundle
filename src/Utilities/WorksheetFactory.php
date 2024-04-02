<?php
/**
 * Created by PhpStorm.
 * User: ehymel
 * Date: 3/27/2018
 * Time: 7:33 PM.
 */

namespace OncologySupport\PhpSpreadsheetHelper\Utilities;

use PhpOffice\PhpSpreadsheet\Spreadsheet;
use PhpOffice\PhpSpreadsheet\Style\NumberFormat;
use PhpOffice\PhpSpreadsheet\Worksheet\Worksheet;

class WorksheetFactory
{
    private readonly Spreadsheet $spreadsheet;
    private readonly ?Worksheet $worksheet;

    private int $currentRow = 1;
    private int $columnCount;

    /**
     * WorksheetFactory constructor.
     *
     * @throws \PhpOffice\PhpSpreadsheet\Exception
     */
    public function __construct(SpreadsheetFactory $factory, string $worksheetTitle = '')
    {
        $this->spreadsheet = $factory->getSpreadsheet();
        $this->worksheet = $this->spreadsheet->getActiveSheet();
        $this->worksheet->setTitle($worksheetTitle);
    }

    /**
     * Add title to first row of worksheet.
     *
     * @throws \PhpOffice\PhpSpreadsheet\Exception
     */
    public function addTitleRow(string $title = null): void
    {
        $this->currentRow = 1;
        $cellCoordinate = [1, $this->currentRow];

        $this->worksheet
            ->setCellValue($cellCoordinate, $title)
            ->getStyle($cellCoordinate)
            ->getFont()->setBold(true)->setSize(14)
        ;

        // merge cells that contain this title for proper auto-sizing of columns
        $this->worksheet->mergeCells([1, $this->currentRow, 5, $this->currentRow]);
    }

    /**
     * Add a single row of data to worksheet.
     */
    public function addDataRow(array $rowElements): void
    {
        ++$this->currentRow;

        foreach ($rowElements as $key => $element) {
            $cellValue = is_array($element) ? $element[0] : $element;
            $cellCoordinate = [$key + 1, $this->currentRow];

            $this->worksheet->setCellValue($cellCoordinate, $cellValue);

            if (is_array($element) && $element[1] === 'date') {
                $this->worksheet
                    ->getStyle($cellCoordinate)
                    ->getNumberFormat()
                    ->setFormatCode(NumberFormat::FORMAT_DATE_XLSX14);
            }
        }
    }

    /**
     * Add array of row data to worksheet.
     */
    public function addDataRows(array $data): void
    {
        foreach ($data as $rowElements) {
            $this->addDataRow($rowElements);
        }

        $this->autosizeAllColumns();
    }

    /**
     * Add headers to columns for orders in this worksheet.
     *
     * @param array $headers Array of headers to add to worksheet
     *
     * @return void
     *
     * @throws \PhpOffice\PhpSpreadsheet\Exception
     */
    public function addHeaderRow(array $headers): void
    {
        ++$this->currentRow;
        $this->columnCount = count($headers);

        foreach ($headers as $key => $header) {
            $column_index = $key + 1;
            $cellCoordinate = [$column_index, $this->currentRow];

            $this->worksheet
                ->setCellValue($cellCoordinate, $header)
                ->getStyle($cellCoordinate)
                ->getFont()->setBold(true);

            // set wrap text
            $this->worksheet
                ->getStyle($cellCoordinate)
                ->getAlignment()->setWrapText(true);
        }

        // set filter on these rows
        $this->worksheet->setAutoFilter([1, $this->currentRow, $this->columnCount, $this->currentRow]);
        $this->worksheet->getAutoFilter()->setRangeToMaxRow();

        // Freeze pane just below header row
        $this->worksheet->freezePane([1, $this->currentRow + 1]);
    }

    private function autosizeAllColumns(): void
    {
        for ($column = 0; $column <= $this->columnCount; ++$column) {
            $this->worksheet->getColumnDimensionByColumn($column)->setAutoSize(true);
        }
    }
}
