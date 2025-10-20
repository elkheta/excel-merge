<?php declare(strict_types=1);

namespace Nzalheart\ExcelMerge;

use Nzalheart\ExcelMerge\Tasks\App;
use Nzalheart\ExcelMerge\Tasks\ContentTypes;
use Nzalheart\ExcelMerge\Tasks\Styles;
use Nzalheart\ExcelMerge\Tasks\Workbook;
use Nzalheart\ExcelMerge\Tasks\WorkbookRels;
use Nzalheart\ExcelMerge\Tasks\Worksheet;
use RecursiveDirectoryIterator;
use RecursiveIteratorIterator;
use SplFileInfo;
use Webmozart\Assert\Assert;
use ZipArchive;

use function count;
use function in_array;

use const DIRECTORY_SEPARATOR;
use const E_USER_ERROR;
use const E_USER_WARNING;
use const PATHINFO_DIRNAME;
use const PATHINFO_EXTENSION;
use const PATHINFO_FILENAME;

/**
 * @property string $workingDir
 * @property string $resultDir
 */
final class ExcelMerge
{
    public bool $debug = false;

    /** @var array<string> */
    private array $files = [];

    private string $workingDir;

    private string $tmpDir;

    private string $resultDir;

    private Styles $stylesTask;

    private Worksheet $worksheetTask;

    private WorkbookRels $workbookRelsTask;

    private App $appTask;

    private ContentTypes $contentTypesTask;

    private Workbook $workbookTask;

    /** @param array<string> $files */
    public function __construct(array $files)
    {
        $this->initializeDirectories();
        $this->registerMergeTasks();

        // Batch process all files at once
        $this->processFiles($files);
    }

    public function __destruct()
    {
        if (!$this->debug) {
            $workingDirPath = realpath($this->workingDir);
            Assert::string($workingDirPath);
            $this->removeTree($workingDirPath);
        }
    }

    public function getResultDir(): string
    {
        return $this->resultDir;
    }

    public function getWorkingDir(): string
    {
        return $this->workingDir;
    }

    public function getTmpDir(): string
    {
        return $this->tmpDir;
    }

    public function addFile(string $filename): void
    {
        // This method is now handled by processFiles() batch processing
        // Keeping for backward compatibility but no longer used in constructor
        if ($this->isSupportedFile($filename)) {
            if ($this->resultsDirEmpty()) {
                $this->addFirstFile($filename);
            } else {
                $this->mergeWorksheets($filename);
            }
            $this->files[] = $filename;
        }
    }

    public function save(string $destinationPath): void
    {
        $zipfile = $this->zipContents();

        $destinationPath =
            pathinfo($destinationPath, PATHINFO_DIRNAME) .
            DIRECTORY_SEPARATOR .
            pathinfo($destinationPath, PATHINFO_FILENAME) . '.' .
            pathinfo($zipfile, PATHINFO_EXTENSION);

        rename($zipfile, $destinationPath);
    }

    public function download(string $downloadFilename): void
    {
        $zipfile = $this->zipContents();

        $downloadFilename =
            pathinfo($downloadFilename, PATHINFO_FILENAME) . '.' .
            pathinfo($zipfile, PATHINFO_EXTENSION);

        header('Content-Type: application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
        header('Content-Disposition: attachment;filename="' . $downloadFilename . '"');
        header('Cache-Control: max-age=0');
        echo file_get_contents($zipfile);
        unlink($zipfile);
        exit;
    }

    /** @return array<string> */
    public function getFiles(): array
    {
        return $this->files;
    }

    private function processFiles(array $files): void
    {
        $validFiles = [];
        foreach ($files as $file) {
            if ($this->isSupportedFile($file, false)) {
                $validFiles[] = $file;
            }
        }

        if (empty($validFiles)) {
            return;
        }

        // Add first file
        $this->addFirstFile($validFiles[0]);
        $this->files[] = $validFiles[0];

        // Batch merge remaining files
        if (count($validFiles) > 1) {
            $remainingFiles = array_slice($validFiles, 1);
            $this->batchMergeWorksheets($remainingFiles);
        }
    }

    private function initializeDirectories(): void
    {
        $this->initializeWorkingDirectory();
        $this->initializeTempDirectory();
        $this->initializeResultDirectory();
    }

    private function initializeWorkingDirectory(): void
    {
        $this->workingDir =
            sys_get_temp_dir() .
            DIRECTORY_SEPARATOR .
            'ExcelMerge-' .
            date('Ymd-His') .
            '-' .
            uniqid() .
            DIRECTORY_SEPARATOR;

        if (!is_dir($this->workingDir)) {
            mkdir($this->workingDir, 0755, true);
        }

        if (!is_dir($this->workingDir)) {
            trigger_error("Could not create temporary working directory {$this->workingDir}", E_USER_ERROR);
        }
    }

    private function initializeTempDirectory(): void
    {
        $this->tmpDir = $this->workingDir . 'tmp' . DIRECTORY_SEPARATOR;
        mkdir($this->tmpDir, 0755, true);
    }

    private function initializeResultDirectory(): void
    {
        $this->resultDir = $this->workingDir . 'result' . DIRECTORY_SEPARATOR;
        mkdir($this->resultDir, 0755, true);
    }

    private function addFirstFile(string $filename): void
    {
        if ($this->isSupportedFile($filename)) {
            $this->unzip($filename, $this->resultDir);
        }
    }

    private function batchMergeWorksheets(array $filenames): void
    {
        if (empty($filenames)) {
            return;
        }

        $allMergeTasks = [];

        // Prepare all files
        foreach ($filenames as $index => $filename) {
            if (!$this->isSupportedFile($filename, false)) {
                continue;
            }

            // Create unique directory for each file to avoid conflicts
            $zipDir = $this->tmpDir . 'file_' . $index . '_' . basename($filename);
            $this->unzip($filename, $zipDir);

            $worksheets = glob("{$zipDir}/xl/worksheets/sheet*.xml");
            Assert::isArray($worksheets);

            $allMergeTasks[] = [
                'zipDir' => $zipDir,
                'worksheets' => $worksheets,
                'filename' => $filename
            ];

            $this->files[] = $filename;
        }

        if (empty($allMergeTasks)) {
            return;
        }

        // Process all files in batch using the optimized method
        $this->worksheetTask->batchAppendToFirstSheet($allMergeTasks, $this->stylesTask);
    }

    private function mergeWorksheets(string $filename): void
    {
        if ($this->resultsDirEmpty()) {
            $this->addFirstFile($filename);
        } else {
            if ($this->isSupportedFile($filename)) {
                // Create unique directory to avoid conflicts
                $zipDir = $this->tmpDir . DIRECTORY_SEPARATOR . 'merge_' . uniqid() . '_' . basename($filename);
                $this->unzip($filename, $zipDir);

                list($styles, $conditionalStyles) = $this->stylesTask->merge($zipDir);

                $worksheets = glob("{$zipDir}/xl/worksheets/sheet*.xml");
                Assert::isArray($worksheets);
                // Only process the first worksheet (sheet1.xml) from each file
                foreach ($worksheets as $worksheet) {
                    if (basename($worksheet) === 'sheet1.xml') {
                        $this->worksheetTask->appendToFirstSheet($worksheet, $styles, $conditionalStyles);
                        break; // Only process the first sheet from each file
                    }
                }
            }
        }
    }

    private function registerMergeTasks(): void
    {
        $this->stylesTask = new Styles($this);
        $this->worksheetTask = new Worksheet($this);
        $this->workbookRelsTask = new WorkbookRels($this);
        $this->contentTypesTask = new ContentTypes($this);
        $this->appTask = new App($this);
        $this->workbookTask = new Workbook($this);
    }

    private function isSupportedFile(string $filename, bool $throwError = true): bool
    {
        $extension = pathinfo($filename, PATHINFO_EXTENSION);
        $isSupported = in_array(mb_strtolower($extension), ['xlsx', 'xlsm'], true);
        if (!$isSupported && $throwError) {
            trigger_error('Can only merge Excel files in .XLSX or .XLSM format. Skipping ' . $filename, E_USER_WARNING);
        }

        return $isSupported;
    }

    private function resultsDirEmpty(): bool
    {
        $files = scandir($this->resultDir);
        Assert::isArray($files);

        return 0 == count(array_diff($files, ['.', '..']));
    }

    private function unzip(string $filename, string $directory): void
    {
        $zip = new ZipArchive();
        $zip->open($filename);
        $zip->extractTo($directory);
        $zip->close();
    }

    private function removeTree(string $dir): bool
    {
        $result = false;

        $dir = realpath($dir);
        Assert::string($dir);
        $tmpDirPath = realpath(sys_get_temp_dir());
        Assert::string($tmpDirPath);
        if (0 === mb_strpos($dir, $tmpDirPath)) {
            $result = true;
            $files = scandir($dir);
            if (false !== $files) {
                $files = array_diff($files, ['.', '..']);
                foreach ($files as $file) {
                    if (is_dir("$dir/$file")) {
                        $result &= $this->removeTree("$dir/$file");
                    } else {
                        $result &= unlink("$dir/$file");
                    }
                }
            }

            $result &= rmdir($dir);
        }

        return (bool) $result;
    }

    private function zipContents(): string
    {
        $zipDirectory = realpath($this->resultDir);
        Assert::string($zipDirectory);
        $targetZip = $this->workingDir . DIRECTORY_SEPARATOR . 'merged-excel-file';
        $ext = 'xlsx';

        $delete = [];

        $zip = new ZipArchive();
        $zip->open($targetZip, ZipArchive::CREATE | ZipArchive::OVERWRITE);

        /** @var SplFileInfo[] $files */
        $files = new RecursiveIteratorIterator(
            new RecursiveDirectoryIterator($zipDirectory),
            RecursiveIteratorIterator::LEAVES_ONLY
        );

        foreach ($files as $name => $file) {
            if (!$file->isDir()) {
                $filePath = $file->getRealPath();
                if (basename($filePath) != $targetZip) {
                    $relativePath = mb_substr($filePath, mb_strlen($zipDirectory) + 1);
                    $zip->addFile($filePath, $relativePath);
                    $delete[] = $filePath;
                }
            }
        }

        $zip->close();

        if (!$this->debug) {
            foreach ($delete as $d) {
                unlink($d);
            }
        }

        rename($targetZip, "$targetZip.$ext");

        return "$targetZip.$ext";
    }
}