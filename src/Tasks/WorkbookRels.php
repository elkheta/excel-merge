<?php declare(strict_types=1);

namespace Nzalheart\ExcelMerge\Tasks;

use DOMDocument;
use DOMXPath;

/**
 * Modifies the "xl/_rels/workbook.xml.rels" file to contain one more worksheet.
 */
final class WorkbookRels extends MergeTask
{
    public function merge(): void
    {
        /**
         *  xl/_rels/workbook.xml.rels
         *  => in 'Relationships'
         *  => add
         *  <Relationship Id="rId{N}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet{N}.xml"/>.
         *
         *  => Renumber all rId{X} values to rId{X+1} where X >= N
         *
         * -> Re-order and re-number so that we first list all the sheets, and then the rest
         */
        $filename = "{$this->getResultDir()}/xl/_rels/workbook.xml.rels";
        $dom = new DOMDocument();
        $dom->load($filename);

        $xpath = new DOMXPath($dom);
        $xpath->registerNamespace('m', 'http://schemas.openxmlformats.org/package/2006/relationships');
        $elems = $xpath->query('//m:Relationship');

        $restId = $this->sheetNumber + 1;
        if (false != $elems) {
            foreach ($elems as $e) {
                $type = $e->getAttribute('Type');
                $isWorksheet = (false !== mb_strpos($type, 'worksheet'));

                if ($isWorksheet) {
                    sscanf($e->getAttribute('Target'), 'worksheets/sheet%d.xml', $sheetNr);
                    $e->setAttribute('Id', 'rId' . $sheetNr);
                } else {
                    $e->setAttribute('Id', 'rId' . ($restId++));
                }
            }
        }

        $newRid = 'rId' . $this->sheetNumber;
        $tag = $dom->createElement('Relationship');
        $tag->setAttribute('Id', $newRid);
        $tag->setAttribute('Type', 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet');
        $tag->setAttribute('Target', 'worksheets/sheet' . $this->sheetNumber . '.xml');

        if (null !== $dom->documentElement) {
            $dom->documentElement->appendChild($tag);
        }

        $dom->save($filename);
    }
}
