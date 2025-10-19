<?php declare(strict_types=1);

/**
 * This file is part of Marketresponse.
 *
 * Unauthorized copying of this file, via any medium is strictly prohibited.
 *
 * @copyright Copyright (c) 2025 Marketresponse - All Rights Reserved
 * @license   Proprietary and confidential
 */

namespace Nzalheart\ExcelMerge\Tasks;

use DOMDocument;
use DOMXPath;
use Symfony\Component\Filesystem\Exception\FileNotFoundException;

use function array_key_exists;
use function count;
use function dirname;

final class Worksheet extends MergeTask
{
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
        //		$elems = $xpath->query("//m:sheets/m:sheet[@sheetId='" . $sheetNumber . "']");
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
