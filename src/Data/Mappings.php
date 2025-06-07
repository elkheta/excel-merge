<?php declare(strict_types=1);

namespace Nzalheart\ExcelMerge\Data;

use function array_key_exists;

final class Mappings
{
    /** @var array<string, array<int, int>> */
    private array $map;

    /** @param array<string, array<int, int>> $map*/
    public function __construct(array $map)
    {
        $this->map = $map;
    }

    /** @return array<string, array<int, int>> */
    public function getMap(): array
    {
        return $this->map;
    }

    /** @return int[] */
    public function getMapByKey(string $key): array
    {
        return array_key_exists($key, $this->map) ? $this->map[$key] : [];
    }
}
