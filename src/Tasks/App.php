<?php declare(strict_types=1);

namespace Nzalheart\ExcelMerge\Tasks;

use DOMDocument;
use DOMXPath;

final class App extends MergeTask
{
    public function merge(): void
    {
        $filename = "{$this->getResultDir()}/docProps/app.xml";

        $dom = new DOMDocument();
        $dom->load($filename);

        $xpath = new DOMXPath($dom);
        $xpath->registerNamespace('m', 'http://schemas.openxmlformats.org/officeDocument/2006/extended-properties');
        $xpath->registerNamespace('mvt', 'http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes');

        $elements = $xpath->query('//m:HeadingPairs/mvt:vector/mvt:variant[2]/mvt:i4');
        if (false !== $elements) {
            foreach ($elements as $element) {
                $element->nodeValue = (string) $this->sheetNumber;
            }
        }

        $elements = $xpath->query('//m:TitlesOfParts/mvt:vector');

        if (false !== $elements) {
            foreach ($elements as $element) {
                // Caroline Clep: Rename if already exists
                $nodes = $element->childNodes;
                foreach ($nodes as $node) {
                    if ($node->nodeValue === $this->sheetName) {
                        $node->nodeValue = 'Previous_' . $node->nodeValue;
                        break;
                    }
                }

                // Caroline Clep: sheets numbers incorrectly
                // $e->setAttribute('size', $this->sheetNumber);
                $element->setAttribute('size', (string) ($element->getAttribute('size') + 1));

                $element->setAttribute('size', (string) $this->sheetNumber);

                $tag = $dom->createElement('vt:lpstr');
                $tag->nodeValue = $this->sheetName;

                $element->appendChild($tag);
            }
        }

        $dom->save($filename);
    }
}
