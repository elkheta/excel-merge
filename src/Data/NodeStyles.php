<?php
declare(strict_types=1);

namespace Nzalheart\ExcelMerge\Data;

use DOMNode;

final class NodeStyles
{
    private DOMNode $node;

    private string $string;

    private int $id;

    public function __construct(DOMNode $node, string $string, int $id)
    {
        $this->node = $node;
        $this->string = $string;
        $this->id = $id;
    }

    public function getNode(): DOMNode
    {
        return $this->node;
    }

    public function getString(): string
    {
        return $this->string;
    }

    public function getId(): int
    {
        return $this->id;
    }
}
