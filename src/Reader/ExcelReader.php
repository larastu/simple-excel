<?php

namespace Larastu\SimpleExcel\Reader;

use Exception;
use SimpleXMLElement;
use ZipArchive;


class ExcelReader
{
    private $tempDir = '';
    private $tempFiles = array();

    /**
     * @var array Data about separate sheets in the file
     */
    private $sheets = [];

    public function __construct($filepath)
    {
        $DefaultTZ = @date_default_timezone_get();
        if ($DefaultTZ) {
            date_default_timezone_set($DefaultTZ);
        }
        if (!is_readable($filepath)) {
            throw new Exception('没有读取文件权限');
        }

        $this->tempDir = sys_get_temp_dir() . DIRECTORY_SEPARATOR . uniqid() . DIRECTORY_SEPARATOR;

        $zip = new ZipArchive;
        if ($zip->open($filepath) !== true) {
            throw new Exception('文件打开失败');
        }

        if ($zip->locateName('xl/sharedStrings.xml') !== false) {
            $zip->extractTo($this->tempDir, 'xl/sharedStrings.xml');
            $this->tempFiles[] = $filePath = $this->tempDir . 'xl' . DIRECTORY_SEPARATOR . 'sharedStrings.xml';

            $sharedString = SharedString::getInstance();
            $sharedString->init($filePath);
        }

        if ($zip->locateName('xl/workbook.xml') === false) {
            throw new Exception('文件格式错误');
        }
        $workbookXML = new SimpleXMLElement($zip->getFromName('xl/workbook.xml'));
        foreach ($workbookXML->sheets->sheet as $sheet) {
            $attributes = $sheet->attributes('r', true);
            foreach ($attributes as $name => $value) {
                if ($name == 'id') {
                    $sheetID = (int)str_replace('rId', '', (string)$value);
                    if ($zip->locateName('xl/worksheets/sheet' . $sheetID . '.xml') === false) {
                        throw new Exception('Excel错误' . $sheetID);
                    }
                    $zip->extractTo($this->tempDir, 'xl/worksheets/sheet' . $sheetID . '.xml');
                    $this->tempFiles[] = $xmlPath = $this->tempDir . 'xl' . DIRECTORY_SEPARATOR . 'worksheets' . DIRECTORY_SEPARATOR . 'sheet' . $sheetID . '.xml';

                    $this->sheets[$sheetID] = new Sheet((string)$sheet['name'], $xmlPath);
                    break;
                }
            }
        }
        ksort($this->sheets);
        $this->sheets = array_values($this->sheets);

        // If worksheet is present and is OK, parse the styles already
        if ($zip->locateName('xl/styles.xml') !== false) {
            $styles = Styles::getInstance();
            $styles->init($zip->getFromName('xl/styles.xml'));
        }

        $zip->close();
    }

    /**
     * Destructor, destroys all that remains (closes and deletes temp files)
     */
    public function __destruct()
    {
        foreach ($this->tempFiles as $TempFile) {
            @unlink($TempFile);
        }

        // Better safe than sorry - shouldn't try deleting '.' or '/', or '..'.
        if (strlen($this->tempDir) > 2) {
            @rmdir($this->tempDir . 'xl' . DIRECTORY_SEPARATOR . 'worksheets');
            @rmdir($this->tempDir . 'xl');
            @rmdir($this->tempDir);
        }
    }

    public function getSheets(): array
    {
        return $this->sheets;
    }

    public function getSheet($index)
    {
        if (!isset($this->sheets[$index])) {
            throw new Exception('工作表不存在');
        }
        return $this->sheets[$index];
    }
}
