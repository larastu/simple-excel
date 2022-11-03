<?php

namespace Larastu\SimpleExcel\Writer;

use Exception;
use ZipArchive;

class ExcelWriter
{
    protected $title = '';
    protected $subject = '';
    protected $author = '';
    protected $company = '';
    protected $description = '';
    protected $keywords = [];

    /**
     * @var Sheet[] 工作表对照表
     */
    protected $sheetMap = [];

    public function __construct()
    {
        date_default_timezone_get() or date_default_timezone_set('UTC');//php.ini missing tz, avoid warning
        StylesXml::getInstance()->addCellStyle('GENERAL', null);
    }

    public function setTitle($title = '')
    {
        $this->title = $title;
    }

    public function setSubject($subject = '')
    {
        $this->subject = $subject;
    }

    public function setAuthor($author = '')
    {
        $this->author = $author;
    }

    public function setCompany($company = '')
    {
        $this->company = $company;
    }

    public function setKeywords($keywords = '')
    {
        $this->keywords = $keywords;
    }

    public function setDescription($description = '')
    {
        $this->description = $description;
    }

    /**
     * 获取工作表
     * @param string $name
     * @return Sheet
     * @throws Exception
     */
    public function createSheet(string $name): Sheet
    {
        $name = trim($name);
        if (empty($name)) {
            throw new Exception('工作表名称不能为空');
        }

        if (isset($this->sheetMap[$name])) {
            throw new Exception('工作表名称已经存在');
        }
        return $this->sheetMap[$name] = new Sheet($name);
    }

    /**
     * 获取工作表
     * @param $name
     * @return Sheet
     * @throws Exception
     */
    public function getSheet($name): Sheet
    {
        $name = trim($name);
        if (!isset($this->sheetMap[$name])) {
            throw new Exception('工作表不存在');
        }
        return $this->sheetMap[$name];
    }

    /**
     * 导出Excel
     * @param $title
     */
    public function writeToStdOut($title)
    {
        header('Content-Type: application/vnd.ms-excel');
        header('Content-Disposition: attachment;filename="' . $title . '.xlsx"');
        header('Cache-Control: max-age=1');
        header('Expires: Mon, 26 Jul 1997 05:00:00 GMT');
        header('Last-Modified: ' . gmdate('D, d M Y H:i:s') . ' GMT');
        header('Cache-Control: cache, must-revalidate');
        header('Pragma: public');

        $temp_file = $this->tempFilename();
        self::writeToFile($temp_file);
        readfile($temp_file);
        exit;
    }

    protected function tempFilename()
    {
        $tmpdir = sys_get_temp_dir();
        return tempnam($tmpdir, "temp_file_");
    }

    public function writeToFile($filename)
    {
        if (empty($this->sheetMap)) {
            Util::log('请先创建工作表');
            return;
        }
        foreach ($this->sheetMap as $sheet) {
            $sheet->finalizeSheet();
        }

        if (file_exists($filename)) {
            @unlink($filename);
        }
        $zip = new ZipArchive();
        if (!$zip->open($filename, ZipArchive::CREATE)) {
            Util::log('Excel文件创建失败');
            return;
        }

        $zip->addEmptyDir('docProps/');
        $zip->addFromString('docProps/app.xml', self::buildAppXML());
        $zip->addFromString('docProps/core.xml', self::buildCoreXML());

        $zip->addEmptyDir('_rels/');
        $zip->addFromString('_rels/.rels', self::buildRelationshipsXML());

        $zip->addEmptyDir('xl/');
        $zip->addFromString('xl/workbook.xml', self::buildWorkbookXML());
        $zip->addFile(StylesXml::getInstance()->writeStylesXML(), 'xl/styles.xml');  //$zip->addFromString("xl/styles.xml"           , self::buildStylesXML() );
        $zip->addFromString('[Content_Types].xml', self::buildContentTypesXML());

        $zip->addEmptyDir('xl/worksheets/');
        $sheetIndex = 1;
        foreach ($this->sheetMap as $sheet) {
            $zip->addFile($sheet->getFile(), "xl/worksheets/sheet{$sheetIndex}.xml");
            $sheetIndex++;
        }

        $zip->addEmptyDir('xl/_rels/');
        $zip->addFromString('xl/_rels/workbook.xml.rels', self::buildWorkbookRelsXML());
        $zip->close();
    }

    protected function buildAppXML(): string
    {
        $xml = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>';
        $xml .= '<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties" xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes">';
        $xml .= '<TotalTime>0</TotalTime>';
        $xml .= '<Company>' . Util::xmlSpecialChars($this->company) . '</Company>';
        $xml .= '</Properties>';
        return $xml;
    }

    protected function buildCoreXML(): string
    {
        $xml = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>';
        $xml .= '<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:dcmitype="http://purl.org/dc/dcmitype/" xmlns:dcterms="http://purl.org/dc/terms/" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">';
        $xml .= '<dcterms:created xsi:type="dcterms:W3CDTF">' . date("Y-m-d\TH:i:s.00\Z") . '</dcterms:created>';//$date_time = '2014-10-25T15:54:37.00Z';
        $xml .= '<dc:title>' . Util::xmlSpecialChars($this->title) . '</dc:title>';
        $xml .= '<dc:subject>' . Util::xmlSpecialChars($this->subject) . '</dc:subject>';
        $xml .= '<dc:creator>' . Util::xmlSpecialChars($this->author) . '</dc:creator>';
        if (!empty($this->keywords)) {
            $xml .= '<cp:keywords>' . Util::xmlSpecialChars(implode(", ", (array)$this->keywords)) . '</cp:keywords>';
        }
        $xml .= '<dc:description>' . Util::xmlSpecialChars($this->description) . '</dc:description>';
        $xml .= '<cp:revision>0</cp:revision>';
        $xml .= '</cp:coreProperties>';
        return $xml;
    }

    protected function buildRelationshipsXML(): string
    {
        $xml = '<?xml version="1.0" encoding="UTF-8"?>';
        $xml .= '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">';
        $xml .= '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>';
        $xml .= '<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties" Target="docProps/core.xml"/>';
        $xml .= '<Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties" Target="docProps/app.xml"/>';
        $xml .= '</Relationships>';
        return $xml;
    }

    protected function buildWorkbookXML(): string
    {
        $i = 0;
        $xml = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>';
        $xml .= '<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">';
        $xml .= '<fileVersion appName="Calc"/><workbookPr backupFile="false" showObjects="all" date1904="false"/><workbookProtection/>';
        $xml .= '<bookViews><workbookView activeTab="0" firstSheet="0" showHorizontalScroll="true" showSheetTabs="true" showVerticalScroll="true" tabRatio="212" windowHeight="8192" windowWidth="16384" xWindow="0" yWindow="0"/></bookViews>';
        $xml .= '<sheets>';
        foreach ($this->sheetMap as $sheet) {
            $xml .= '<sheet name="' . Util::xmlSpecialChars($sheet->getName()) . '" sheetId="' . ($i + 1) . '" state="visible" r:id="rId' . ($i + 2) . '"/>';
            $i++;
        }
        $xml .= '</sheets>';
        $xml .= '<definedNames>';
        foreach ($this->sheetMap as $sheet) {
            if ($sheet->autoFilter) {
                $xml .= '<definedName name="_xlnm._FilterDatabase" localSheetId="0" hidden="1">\''
                    . Util::xmlSpecialChars($sheet->getName()) . '\'!$A$1:'
                    . Sheet::xlsCell($sheet->getRowCount() - 1, $sheet->getColumnCount() - 1, true)
                    . '</definedName>';
                $i++;
            }
        }
        $xml .= '</definedNames>';
        $xml .= '<calcPr iterateCount="100" refMode="A1" iterate="false" iterateDelta="0.001"/></workbook>';
        return $xml;
    }

    protected function buildWorkbookRelsXML(): string
    {
        $xml = '<?xml version="1.0" encoding="UTF-8"?>';
        $xml .= '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">';
        $xml .= '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>';
        $sheetIndex = 1;
        foreach ($this->sheetMap as $sheet) {
            $xml .= '<Relationship Id="rId' . ($sheetIndex + 1) . '"'
                . ' Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet"'
                . ' Target="worksheets/sheet' . $sheetIndex . '.xml"/>';
            $sheetIndex++;
        }
        $xml .= '</Relationships>';
        return $xml;
    }

    protected function buildContentTypesXML(): string
    {
        $content_types_xml = "";
        $content_types_xml .= '<?xml version="1.0" encoding="UTF-8"?>';
        $content_types_xml .= '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">';
        $content_types_xml .= '<Override PartName="/_rels/.rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>';
        $content_types_xml .= '<Override PartName="/xl/_rels/workbook.xml.rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>';
        $sheetIndex = 1;
        foreach ($this->sheetMap as $sheet) {
            $content_types_xml .= '<Override PartName="/xl/worksheets/sheet' . $sheetIndex . '.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>';
            $sheetIndex++;
        }
        $content_types_xml .= '<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>';
        $content_types_xml .= '<Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>';
        $content_types_xml .= '<Override PartName="/docProps/app.xml" ContentType="application/vnd.openxmlformats-officedocument.extended-properties+xml"/>';
        $content_types_xml .= '<Override PartName="/docProps/core.xml" ContentType="application/vnd.openxmlformats-package.core-properties+xml"/>';
        $content_types_xml .= '</Types>';
        return $content_types_xml;
    }
}
