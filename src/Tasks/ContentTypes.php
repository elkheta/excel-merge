<?php declare(strict_types=1);

namespace Nzalheart\ExcelMerge\Tasks;

use DOMDocument;

final class ContentTypes extends MergeTask
{
    public function merge(): void
    {
        $filename = "{$this->getResultDir()}/[Content_Types].xml";

        $dom = new DOMDocument();
        $dom->load($filename);

        $tag = $dom->createElement('Override');
        $tag->setAttribute('PartName', "/xl/worksheets/sheet{$this->sheetNumber}.xml");
        $tag->setAttribute('ContentType', 'application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml');

        if (null !== $dom->documentElement) {
            $dom->documentElement->appendChild($tag);
        }

        $dom->save($filename);
    }
}
