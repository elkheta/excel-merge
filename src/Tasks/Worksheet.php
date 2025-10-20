<?php declare(strict_types=1);

namespace Nzalheart\ExcelMerge\Tasks;

use DOMDocument;
use DOMXPath;
use DOMNode;
use Symfony\Component\Filesystem\Exception\FileNotFoundException;

use function array_key_exists;
use function count;
use function dirname;

final class Worksheet extends MergeTask
{
    private ?DOMDocument $cachedFirstSheet = null;
    private ?DOMXPath $cachedFirstXpath = null;
    private ?DOMNode $cachedFirstSheetData = null;
    private int $cachedLastRowNum = 0;

    /**
     * Batch process multiple worksheets efficiently
     *
     * @param array $mergeTasks Array of merge task data
     * @param Styles $stylesTask
     */
    public function batchAppendToFirstSheet(array $mergeTasks, Styles $stylesTask): void
    {
        // Load destination sheet once
        $firstSheetFile = $this->getResultDir() . "/xl/worksheets/sheet1.xml";
        $this->cachedFirstSheet = new DOMDocument();
        $this->cachedFirstSheet->load($firstSheetFile);
        
        $this->cachedFirstXpath = new DOMXPath($this->cachedFirstSheet);
        $this->cachedFirstXpath->registerNamespace('m', 'http://schemas.openxmlformats.org/spreadsheetml/2006/main');
        
        $this->cachedFirstSheetData = $this->cachedFirstXpath->query('//m:sheetData')->item(0);
        
        // Calculate initial last row
        $firstSheetRows = $this->cachedFirstXpath->query('//m:sheetData/m:row');
        $this->cachedLastRowNum = 0;
        foreach ($firstSheetRows as $row) {
            $rowNum = (int) $row->getAttribute('r');
            if ($rowNum > $this->cachedLastRowNum) {
                $this->cachedLastRowNum = $rowNum;
            }
        }

        // Process all files
        foreach ($mergeTasks as $task) {
            list($styles, $conditionalStyles) = $stylesTask->merge($task['zipDir']);
            
            foreach ($task['worksheets'] as $worksheet) {
                $this->appendToFirstSheetOptimized($worksheet, $styles, $conditionalStyles);
            }
        }

        // Update dimension once at the end
        $this->updateDimension();
        
        // Save once
        $this->cachedFirstSheet->save($firstSheetFile);
        
        // Clear cache
        $this->cachedFirstSheet = null;
        $this->cachedFirstXpath = null;
        $this->cachedFirstSheetData = null;
    }

    /**
     * Optimized append using cached DOM
     */
    private function appendToFirstSheetOptimized(
        string $filename,
        array $stylesMapping,
        array $conditionalStylesMapping
    ): void {
        if (!file_exists($filename)) {
            throw new FileNotFoundException();
        }

        // Load source sheet
        $sourceSheet = new DOMDocument();
        $sourceSheet->load($filename);
        
        $sourceXpath = new DOMXPath($sourceSheet);
        $sourceXpath->registerNamespace('m', 'http://schemas.openxmlformats.org/spreadsheetml/2006/main');

        // Get all data rows from source (skip header row 1)
        $sourceRows = $sourceXpath->query('//m:sheetData/m:row');
        
        // Batch import rows
        $rowsToAppend = [];
        foreach ($sourceRows as $sourceRow) {
            $sourceRowNum = (int) $sourceRow->getAttribute('r');
            
            if ($sourceRowNum <= 1) {
                continue;
            }
            
            $this->cachedLastRowNum++;
            
            // Import the row
            $importedRow = $this->cachedFirstSheet->importNode($sourceRow, true);
            $importedRow->setAttribute('r', (string) $this->cachedLastRowNum);
            
            // Batch process cells
            $this->updateRowCells($importedRow, $stylesMapping);
            
            $rowsToAppend[] = $importedRow;
        }

        // Append all rows at once
        foreach ($rowsToAppend as $row) {
            $this->cachedFirstSheetData->appendChild($row);
        }
    }

    /**
     * Update cell references and styles in a row
     */
    private function updateRowCells(DOMNode $row, array $stylesMapping): void
    {
        $rowNum = $row->getAttribute('r');
        $cells = $row->getElementsByTagName('c');
        
        // Process cells in reverse to avoid live NodeList issues
        $cellArray = [];
        foreach ($cells as $cell) {
            $cellArray[] = $cell;
        }
        
        foreach ($cellArray as $cell) {
            $oldRef = $cell->getAttribute('r');
            
            // Update cell reference
            if (preg_match('/^([A-Z]+)\d+$/', $oldRef, $matches)) {
                $cell->setAttribute('r', $matches[1] . $rowNum);
            }
            
            // Remap style
            $styleId = $cell->getAttribute('s');
            if ($styleId !== '' && is_numeric($styleId)) {
                $oldStyleId = (int) $styleId;
                if (array_key_exists($oldStyleId, $stylesMapping)) {
                    $cell->setAttribute('s', (string) $stylesMapping[$oldStyleId]);
                }
            }
        }
    }

    /**
     * Update dimension at the end
     */
    private function updateDimension(): void
    {
        $dimensionNodes = $this->cachedFirstXpath->query('//m:dimension');
        if ($dimensionNodes->length > 0) {
            $dimensionNode = $dimensionNodes->item(0);
            
            // Find highest column efficiently
            $allCells = $this->cachedFirstXpath->query('//m:sheetData/m:row/m:c');
            $highestCol = 'A';
            
            foreach ($allCells as $cell) {
                $cellRef = $cell->getAttribute('r');
                if (preg_match('/^([A-Z]+)\d+$/', $cellRef, $matches)) {
                    if ($matches[1] > $highestCol) {
                        $highestCol = $matches[1];
                    }
                }
            }
            
            $dimensionNode->setAttribute('ref', "A1:{$highestCol}{$this->cachedLastRowNum}");
        }
    }

    /**
     * Appends worksheet data to the first sheet instead of creating new sheets
     *
     * @param string $filename                 The filename of the sheet to copy
     * @param int[]  $stylesMapping
     * @param int[]  $conditionalStylesMapping
     *
     * @return array{int, string}
     */
    public function appendToFirstSheet(
        string $filename,
        array $stylesMapping,
        array $conditionalStylesMapping,
    ): array {
        if (!file_exists($filename)) {
            throw new FileNotFoundException();
        }

        // Load source sheet
        $sourceSheet = new DOMDocument();
        $sourceSheet->load($filename);
        
        $sourceXpath = new DOMXPath($sourceSheet);
        $sourceXpath->registerNamespace('m', 'http://schemas.openxmlformats.org/spreadsheetml/2006/main');

        // Load first sheet (destination)
        $firstSheetFile = $this->getResultDir() . "/xl/worksheets/sheet1.xml";
        $firstSheet = new DOMDocument();
        $firstSheet->load($firstSheetFile);
        
        $firstXpath = new DOMXPath($firstSheet);
        $firstXpath->registerNamespace('m', 'http://schemas.openxmlformats.org/spreadsheetml/2006/main');

        // Get sheetData from both sheets
        $sourceSheetData = $sourceXpath->query('//m:sheetData')->item(0);
        $firstSheetData = $firstXpath->query('//m:sheetData')->item(0);

        // Find last row in first sheet
        $firstSheetRows = $firstXpath->query('//m:sheetData/m:row');
        $lastRowNum = 0;
        foreach ($firstSheetRows as $row) {
            $rowNum = (int) $row->getAttribute('r');
            if ($rowNum > $lastRowNum) {
                $lastRowNum = $rowNum;
            }
        }

        // Get all data rows from source (skip header row 1)
        $sourceRows = $sourceXpath->query('//m:sheetData/m:row');
        foreach ($sourceRows as $sourceRow) {
            $sourceRowNum = (int) $sourceRow->getAttribute('r');
            
            // Skip header row
            if ($sourceRowNum <= 1) {
                continue;
            }
            
            $lastRowNum++;
            
            // Import the row into first sheet
            $importedRow = $firstSheet->importNode($sourceRow, true);
            
            // Update row number
            $importedRow->setAttribute('r', (string) $lastRowNum);
            
            // Update cell references
            $cells = $importedRow->getElementsByTagName('c');
            foreach ($cells as $cell) {
                $oldRef = $cell->getAttribute('r');
                if (preg_match('/^([A-Z]+)(\d+)$/', $oldRef, $matches)) {
                    $newRef = $matches[1] . $lastRowNum;
                    $cell->setAttribute('r', $newRef);
                }
                
                // Remap styles
                $styleId = $cell->getAttribute('s');
                if ($styleId && is_numeric($styleId)) {
                    $oldStyleId = (int) $styleId;
                    if (array_key_exists($oldStyleId, $stylesMapping)) {
                        $cell->setAttribute('s', (string) $stylesMapping[$oldStyleId]);
                    }
                }
            }
            
            // Append to first sheet
            $firstSheetData->appendChild($importedRow);
        }

        // Update dimension
        $dimensionNodes = $firstXpath->query('//m:dimension');
        if ($dimensionNodes->length > 0) {
            $dimensionNode = $dimensionNodes->item(0);
            // Find highest column
            $allCells = $firstXpath->query('//m:sheetData/m:row/m:c');
            $highestCol = 'A';
            foreach ($allCells as $cell) {
                $cellRef = $cell->getAttribute('r');
                if (preg_match('/^([A-Z]+)(\d+)$/', $cellRef, $matches)) {
                    $col = $matches[1];
                    if ($col > $highestCol) {
                        $highestCol = $col;
                    }
                }
            }
            $dimensionNode->setAttribute('ref', "A1:{$highestCol}{$lastRowNum}");
        }

        // Save the modified first sheet
        $firstSheet->save($firstSheetFile);

        // Return sheet 1 info (we're always appending to sheet 1)
        return [1, 'Merged Data'];
    }

    /**
     * Adds a new worksheet to the merged Excel file.
     *
     * @param string $filename                 The filename of the sheet to copy
     * @param int[]  $stylesMapping
     * @param int[]  $conditionalStylesMapping
     *
     * @return array{int, string}
     */
    public function merge(
        string $filename,
        array $stylesMapping,
        array $conditionalStylesMapping,
    ): array {
        if (!file_exists($filename)) {
            throw new FileNotFoundException();
        }
        $newSheetNumber = $this->getSheetCount($this->getResultDir()) + 1;

        // copy file into place
        $newName = $this->getResultDir() . "/xl/worksheets/sheet{$newSheetNumber}.xml";
        if (!is_dir(dirname($newName))) {
            mkdir(dirname($newName));
        }
        copy($filename, $newName);

        // copy rels
        $relsFile = dirname($filename) . '/_rels/' . basename($filename) . '.rels';
        $relsDirectory = $this->getResultDir() . '/xl/worksheets/_rels';
        if (file_exists($relsFile)) {
            if (!is_dir($relsDirectory)) {
                mkdir($relsDirectory);
            }
            copy($relsFile, $relsDirectory . "/sheet{$newSheetNumber}.xml.rels");
        }

        // adjust references to any shared strings
        $sheet = new DOMDocument();
        $sheet->load($newName);

        $this->remapStyles($sheet, $stylesMapping);
        $this->remapConditionalStyles($sheet, $conditionalStylesMapping);
        $this->remapColsStyles($sheet, $stylesMapping);

        // save worksheet with adjustments
        $sheet->save($newName);

        // extract worksheet name
        $sheetName = $this->extractWorksheetName($filename);

        return [$newSheetNumber, $sheetName];
    }

    private function getSheetCount(string $dir): int
    {
        $existingSheets = glob("{$dir}/xl/worksheets/sheet*.xml");

        if (false !== $existingSheets && count($existingSheets) > 0) {
            natsort($existingSheets);
            $last = basename(end($existingSheets));

            if (null !== sscanf($last, 'sheet%d.xml', $number)) {
                return $number;
            }
        }

        return 0;
    }

    /**
     * @param int[] $mapping
     */
    private function remapStyles(DOMDocument $sheet, array $mapping): void
    {
        $this->doRemapping($sheet, '//m:c[@s]', 's', $mapping);
    }

    /**
     * @param int[] $mapping
     */
    private function remapConditionalStyles(DOMDocument $sheet, array $mapping): void
    {
        $this->doRemapping($sheet, '//m:conditionalFormatting/m:cfRule[@dxfId]', 'dxfId', $mapping);
    }

    /**
     * @param int[] $mapping
     */
    private function remapColsStyles(DOMDocument $sheet, array $mapping): void
    {
        $this->doRemapping($sheet, '//m:col[@Style]', 'style', $mapping);
    }

    /**
     * @param int[] $mapping
     */
    private function doRemapping(DOMDocument $sheet, string $xpathQuery, string $attribute, array $mapping): void
    {
        // adjust references to styles
        $xpath = new DOMXPath($sheet);
        $xpath->registerNamespace('m', 'http://schemas.openxmlformats.org/spreadsheetml/2006/main');
        $conditionalStyles = $xpath->query($xpathQuery);

        if (false !== $conditionalStyles) {
            foreach ($conditionalStyles as $tag) {
                $oldId = $tag->getAttribute($attribute);

                if (is_numeric($oldId)) {
                    $oldId = (int) $oldId;
                    if (array_key_exists($oldId, $mapping)) {
                        $tag->setAttribute($attribute, (string) $mapping[$oldId]);
                    }
                }
            }
        }
    }

    private function extractWorksheetName(string $filename): string
    {
        $workbook = new DOMDocument();
        $workbook->load(dirname($filename) . '/../workbook.xml');

        $xpath = new DOMXPath($workbook);
        $xpath->registerNamespace('m', 'http://schemas.openxmlformats.org/spreadsheetml/2006/main');
        sscanf(basename($filename), 'sheet%d.xml', $number);

        $sheetName = "Worksheet $number";
        $elems = $xpath->query("//m:sheets/m:sheet[@sheetId='" . $number . "']");
        if (false != $elems) {
            foreach ($elems as $e) {
                // should be one only
                $sheetName = $e->getAttribute('name');
                break;
            }
        }

        return $sheetName;
    }
}