<?php
/**
 * Created by PhpStorm.
 * User: ehymel
 * Date: 3/27/2018
 * Time: 7:33 PM.
 */

namespace OncologySupport\PhpSpreadsheetHelper\Utilities;

use PhpOffice\PhpSpreadsheet\Spreadsheet;
use PhpOffice\PhpSpreadsheet\Writer\Xlsx;

class SpreadsheetFactory
{
    private Spreadsheet $spreadsheet;

    /**
     * Spreadsheet constructor.
     *
     * @throws \PhpOffice\PhpSpreadsheet\Exception
     */
    public function __construct()
    {
        $this->createSpreadsheet();
    }

    /**
     * @return Spreadsheet
     */
    public function getSpreadsheet(): Spreadsheet
    {
        return $this->spreadsheet;
    }

    /**
     * Returns a new Spreadsheet object with reasonable default document properties.
     *
     * @throws \PhpOffice\PhpSpreadsheet\Exception
     */
    private function createSpreadsheet(): void
    {
        $this->spreadsheet = new Spreadsheet();

        $this->spreadsheet
            ->getProperties()
            ->setCreator('Oncology Support, PLLC');
    }

    /**
     * Save spreadsheet to local disk in Excel format.
     *
     * @throws \PhpOffice\PhpSpreadsheet\Exception
     * @throws \PhpOffice\PhpSpreadsheet\Writer\Exception
     */
    public function outputToWeb(string $filename = 'spreadsheet.xlsx'): never
    {
        // set active sheet index to the first sheet, so Excel opens this as the first sheet
        $this->spreadsheet->setActiveSheetIndex(0);

        // Redirect output to a client’s web browser (Xlsx)
        header('Content-Type: application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
        header('Content-Disposition: attachment;filename="'.$filename.'"');
        header('Cache-Control: max-age=0');

        $writer = new Xlsx($this->spreadsheet);
        $writer->save('php://output');
        exit;
    }
}
