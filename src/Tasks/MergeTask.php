<?php declare(strict_types=1);

namespace Nzalheart\ExcelMerge\Tasks;

use Nzalheart\ExcelMerge\ExcelMerge;

abstract class MergeTask
{
    protected ExcelMerge $parent;

    protected int $sheetNumber;

    protected string $sheetName;

    public function __construct(ExcelMerge $parent)
    {
        $this->parent = $parent;
    }

    public function getResultDir(): string
    {
        return $this->parent->getResultDir();
    }

    public function set(int $sheetNumber, string $sheetName): void
    {
        $this->sheetNumber = $sheetNumber;
        $this->sheetName = $sheetName;
    }
}
