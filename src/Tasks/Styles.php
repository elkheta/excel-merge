<?php declare(strict_types=1);

namespace Nzalheart\ExcelMerge\Tasks;

use DOMDocument;
use DOMXPath;
use Nzalheart\ExcelMerge\Data\Mappings;
use Nzalheart\ExcelMerge\Data\NodeStyles;
use Webmozart\Assert\Assert;

use function array_key_exists;
use function count;

use const DIRECTORY_SEPARATOR;

/**
 * Consolidates the contents of two 'xl/styles.xml' files into one, and
 * returns two mappings:.
 *
 * 1. a mapping of how old style IDs map onto new style IDs
 * 2. a mapping of how old 'conditional style' IDs map onto new 'conditional style' IDs
 */
final class Styles extends MergeTask
{
    /** @var string[] */
    private array $styleTags = ['numFmts', 'fonts', 'fills', 'borders', 'dxfs'];

    /** @return array{array<int>, array<int>} */
    public function merge(string $zipDir): array
    {
        $xmlFilename = DIRECTORY_SEPARATOR . 'xl' . DIRECTORY_SEPARATOR . 'styles.xml';
        $existingFilename = $this->getResultDir() . $xmlFilename;
        $sourceFilename = $zipDir . $xmlFilename;

        // get hash signature for each entry in 'numfmt', 'fonts', 'fills' and 'borders'
        // see if there are any new ones
        // - if so, add them and store the id. Make sure to update the 'count' attribute in the parent tag
        // - if it already existed, get the id
        $existingDom = new DOMDocument();
        $existingDom->load($existingFilename);

        $existingXpath = new DOMXPath($existingDom);
        $existingXpath->registerNamespace('m', 'http://schemas.openxmlformats.org/spreadsheetml/2006/main');

        $styles = $this->getStyles($existingXpath);

        $sourceDom = new DOMDocument();
        $sourceDom->load($sourceFilename);

        // re-assign xpath to work on source doc
        $sourceXpath = new DOMXPath($sourceDom);
        $sourceXpath->registerNamespace('m', 'http://schemas.openxmlformats.org/spreadsheetml/2006/main');

        // iterate all the style tags in document that we want to merge in
        list($mapping, $styles) = $this->addNewStyles($sourceXpath, $styles);

        // replace styles from existing styles.xml document with the merged styles
        $this->replaceStyleTags($styles, $existingXpath);

        // now go through the 'cellXfs' tags. Update the references to 'fontId', 'numFmtId',
        // 'fillId', and 'borderId'. Generate a tag for each style that we're importing.
        //
        // If it already existed, note the id. If it didn't exist, add it and store the id.
        // Return this mapping of ids
        list($definedStyles, $stylesMapping) = $this->rewriteCells($existingXpath, $sourceXpath, $mapping);

        // write the new styles list
        $this->replaceStylesList($definedStyles, $existingXpath);

        // save the merged style file
        $existingDom->save($existingFilename);

        // return a mapping of how style ids in this workbook relate to style ids in the merged workbook
        return [$stylesMapping, $mapping->getMapByKey('dxfs')];
    }

    /**
     * @return array<string, array<int, NodeStyles>>
     */
    private function getStyles(DOMXPath $existingXpath): array
    {
        $existingStyles = [];
        foreach ($this->styleTags as $tag) {
            $elems = $existingXpath->query("//m:{$tag}");
            $existingStyles[$tag] = [];
            if (false !== $elems && $elems->length > 0) {
                $elementItem = $elems->item(0);
                Assert::notNull($elementItem);
                if ($elementItem->hasChildNodes()) {
                    foreach ($elementItem->childNodes as $id => $style) {
                        $existingStyles[$tag][$id] = new NodeStyles(
                            $style,
                            $style->C14N(true, false),
                            $id
                        );
                    }
                }
            }
        }

        return $existingStyles;
    }

    /**
     * @param DOMXPath                              $sourceXpath    The document to add styles from
     * @param array<string, array<int, NodeStyles>> $existingStyles
     *
     * @return array{Mappings, array<string, array<int, NodeStyles>>}
     */
    private function addNewStyles(DOMXPath $sourceXpath, array $existingStyles): array
    {
        $mapping = [];
        foreach ($this->styleTags as $tag) {
            $elems = $sourceXpath->query("//m:{$tag}");

            $mapping[$tag] = [];
            if (false !== $elems) {
                $elementItem = $elems->item(0);
                if (null !== $elementItem && $elementItem->hasChildNodes()) {
                    foreach ($elementItem->childNodes as $id => $style) {
                        $string = $style->C14N(true, false);

                        foreach ($existingStyles[$tag] as $e) {
                            if ($e->getString() === $string) {
                                // this is an existing style
                                $mapping[$tag][$id] = $e->getId();
                                continue 2; // continue to next style
                            }
                        }

                        // this is a new style
                        $newId = count($existingStyles[$tag]);

                        $existingStyles[$tag][] = new NodeStyles(
                            $style,
                            $style->C14N(true, false),
                            $newId
                        );

                        $mapping[$tag][$id] = $newId;
                    }
                }
            }
        }

        return [new Mappings($mapping), $existingStyles];
    }

    /**
     * @param array<string, array<int, NodeStyles>> $existingStyles
     */
    private function replaceStyleTags(array $existingStyles, DOMXPath $xpath): void
    {
        foreach ($existingStyles as $tag => $styles) {
            $elems = $xpath->query("//m:{$tag}");

            if (false !== $elems && $elems->length > 0) {
                $elem = $elems->item(0);
                Assert::notNull($elem);

                while ($elem->hasChildNodes()) {
                    if (null !== $elem->firstChild) {
                        $elem->removeChild($elem->firstChild);
                    }
                }

                foreach ($styles as $style) {
                    $elem->appendChild($xpath->document->importNode($style->getNode(), true));
                }
                $elem->setAttribute('count', (string) count($styles));
            }
        }
    }

    /**
     * @return array{NodeStyles[], array<int>}
     */
    private function rewriteCells(DOMXPath $existingXpath, DOMXPath $sourceXpath, Mappings $mapping): array
    {
        $elems = $existingXpath->query('//m:cellXfs');
        $definedStyles = [];
        if (false !== $elems && $elems->length > 0) {
            $elemsItem = $elems->item(0);
            Assert::notNull($elemsItem);
            if ($elemsItem->hasChildNodes()) {
                foreach ($elemsItem->childNodes as $id => $style) {
                    $definedStyles[$id] = new NodeStyles(
                        $style,
                        $style->C14N(true, false),
                        $id
                    );
                }
            }
        }

        $stylesMapping = [];
        $elems = $sourceXpath->query('//m:cellXfs');
        if (false !== $elems && $elems->length > 0) {
            $elemsItem = $elems->item(0);
            Assert::notNull($elemsItem);
            if ($elemsItem->hasChildNodes()) {
                foreach ($elemsItem->childNodes as $id => $style) {
                    $fontId = (int) $style->getAttribute('fontId');
                    if (array_key_exists($fontId, $mapping->getMapByKey('fonts'))) {
                        $style->setAttribute('fontId', (string) (0 + $mapping->getMapByKey('fonts')[$fontId]));
                    }

                    $numFmtId = (int) $style->getAttribute('numFmtId');
                    if (array_key_exists($numFmtId, $mapping->getMapByKey('numFmts'))) {
                        $style->setAttribute('numFmtId', (string) (0 + $mapping->getMapByKey('numFmts')[$numFmtId]));
                    }

                    $fillId = (int) $style->getAttribute('fillId');
                    if (array_key_exists($fillId, $mapping->getMapByKey('fills'))) {
                        $style->setAttribute('fillId', (string) (0 + $mapping->getMapByKey('fills')[$fillId]));
                    }

                    $borderId = (int) $style->getAttribute('borderId');
                    if (array_key_exists($borderId, $mapping->getMapByKey('borders'))) {
                        $style->setAttribute('borderId', (string) (0 + $mapping->getMapByKey('borders')[$borderId]));
                    }

                    $string = $style->C14N(true, false);

                    foreach ($definedStyles as $definedStyle) {
                        if ($definedStyle->getString() == $string) {
                            // we found an existing style
                            $stylesMapping[$id] = $definedStyle->getId();
                            continue 2;
                        }
                    }

                    // this is a new style!
                    $newId = count($definedStyles);
                    $definedStyles[$newId] = new NodeStyles(
                        $style,
                        $style->C14N(true, false),
                        $newId
                    );

                    $stylesMapping[$id] = $newId;
                }
            }
        }

        return [$definedStyles, $stylesMapping];
    }

    /**
     * @param NodeStyles[] $definedStyles
     */
    private function replaceStylesList($definedStyles, DOMXPath $existingXpath): void
    {
        $elems = $existingXpath->query('//m:cellXfs');
        if (false !== $elems && $elems->length > 0) {
            $elem = $elems->item(0);
            Assert::notNull($elem);
            while ($elem->hasChildNodes()) {
                if (null !== $elem->firstChild) {
                    $elem->removeChild($elem->firstChild);
                }
            }
            foreach ($definedStyles as $definedStyle) {
                $elem->appendChild($existingXpath->document->importNode($definedStyle->getNode(), true));
            }
            $elem->setAttribute('count', (string) count($definedStyles));
        }
    }
}
