<?php
declare(strict_types=1);

namespace Tests\ExcelMerge;

use Nzalheart\ExcelMerge\ExcelMerge;
use PHPUnit\Framework\TestCase;
use Symfony\Component\Filesystem\Filesystem;

use const DIRECTORY_SEPARATOR;

class ExcelMergeTest extends TestCase
{
    public function testFilesAreMergedWithoutExternalLinks(): void
    {
        $files = [
            __DIR__ . DIRECTORY_SEPARATOR . 'sampleExcelFiles' . DIRECTORY_SEPARATOR . 'sampleWithoutExternalLinks' . DIRECTORY_SEPARATOR . 'file1.xlsx',
            __DIR__ . DIRECTORY_SEPARATOR . 'sampleExcelFiles' . DIRECTORY_SEPARATOR . 'sampleWithoutExternalLinks' . DIRECTORY_SEPARATOR . 'file2.xlsx',
            __DIR__ . DIRECTORY_SEPARATOR . 'sampleExcelFiles' . DIRECTORY_SEPARATOR . 'sampleWithoutExternalLinks' . DIRECTORY_SEPARATOR . 'file3.xlsx'
        ];

        $excelMerger = new ExcelMerge($files);
        $destinationPath = sys_get_temp_dir() . DIRECTORY_SEPARATOR . 'excel-merge' . DIRECTORY_SEPARATOR . uniqid() . '.xlsx';
        $filesystem = new Filesystem();
        $filesystem->mkdir(sys_get_temp_dir() . DIRECTORY_SEPARATOR . 'excel-merge', 0755);
        $excelMerger->save($destinationPath);

        $this->assertFileExists($destinationPath);
    }

    public function testFilesAreMergedWithExternalLinks(): void
    {
        $files = [
            __DIR__ . DIRECTORY_SEPARATOR . 'sampleExcelFiles' . DIRECTORY_SEPARATOR . 'sampleWithExternalLinks' . DIRECTORY_SEPARATOR . 'file1.xlsx',
            __DIR__ . DIRECTORY_SEPARATOR . 'sampleExcelFiles' . DIRECTORY_SEPARATOR . 'sampleWithExternalLinks' . DIRECTORY_SEPARATOR . 'file2.xlsx',
            __DIR__ . DIRECTORY_SEPARATOR . 'sampleExcelFiles' . DIRECTORY_SEPARATOR . 'sampleWithExternalLinks' . DIRECTORY_SEPARATOR . 'file3.xlsx'
        ];

        $excelMerger = new ExcelMerge($files);
        $destinationPath = sys_get_temp_dir() . DIRECTORY_SEPARATOR . 'excel-merge' . DIRECTORY_SEPARATOR . uniqid() . '.xlsx';
        $filesystem = new Filesystem();
        $filesystem->mkdir(sys_get_temp_dir() . DIRECTORY_SEPARATOR . 'excel-merge', 0755);
        $excelMerger->save($destinationPath);

        $this->assertFileExists($destinationPath);
    }
}
