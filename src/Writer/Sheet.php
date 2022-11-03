<?php

namespace Larastu\SimpleExcel\Writer;

use Exception;

class Sheet
{
    //http://www.ecma-international.org/publications/standards/Ecma-376.htm
    //http://officeopenxml.com/SSstyles.php
    //------------------------------------------------------------------
    //http://office.microsoft.com/en-us/excel-help/excel-specifications-and-limits-HP010073849.aspx
    const EXCEL_2007_MAX_ROW = 1000000;
    const EXCEL_2007_MAX_COL = 10000;


    /**
     * @var string 工作表名称
     */
    private $name;
    /**
     * @var FileWriter
     */
    private $writer;
    /**
     * @var int
     */
    private $rowCount = 0;
    /**
     * @var int
     */
    private $max_cell_tag_start = 0;
    /**
     * @var int
     */
    private $max_cell_tag_end = 0;
    /**
     * @var array 列类型配置
     */
    private $columnTypes = [];
    /**
     * @var array
     */
    private $mergeCells = [];
    /**
     * @var bool
     */
    private $finalized = false;
    public $autoFilter = false;

    public function __construct($name)
    {
        $this->setName($name);

        $this->writer = new FileWriter();
        $this->init();
    }

    public function getFile()
    {
        return $this->writer->getTempFile();
    }

    public function setName($name)
    {
        $this->name = $name;
    }

    public function getName(): string
    {
        return $this->name;
    }

    public function getRowCount(): int
    {
        return $this->rowCount;
    }

    public function getColumnCount(): int
    {
        return count($this->columnTypes);
    }

    protected function init($col_widths = [], $auto_filter = false, $freeze_rows = false, $freeze_columns = false)
    {
        $this->writer->write('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' . "\n");
        $this->writer->write('<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">');

        $this->writer->write('<sheetPr filterMode="false">');
        $this->writer->write('<pageSetUpPr fitToPage="false"/>');
        $this->writer->write('</sheetPr>');
        $this->max_cell_tag_start = $this->writer->ftell();
        $max_cell = self::xlsCell(self::EXCEL_2007_MAX_ROW, self::EXCEL_2007_MAX_COL);//XFE1048577
        $this->writer->write('<dimension ref="A1:' . $max_cell . '"/>');
        $this->max_cell_tag_end = $this->writer->ftell();

        $this->writer->write('<sheetViews>');
        $this->writer->write('<sheetView colorId="64" defaultGridColor="true" rightToLeft="false" showFormulas="false" showGridLines="true" showOutlineSymbols="true" showRowColHeaders="true" showZeros="true" tabSelected="false" topLeftCell="A1" view="normal" windowProtection="false" workbookViewId="0" zoomScale="100" zoomScaleNormal="100" zoomScalePageLayoutView="100">');
        if ($freeze_rows && $freeze_columns) {
            $this->writer->write('<pane ySplit="' . $freeze_rows . '" xSplit="' . $freeze_columns . '" topLeftCell="' . self::xlsCell($freeze_rows, $freeze_columns) . '" activePane="bottomRight" state="frozen"/>');
            $this->writer->write('<selection activeCell="' . self::xlsCell($freeze_rows, 0) . '" activeCellId="0" pane="topRight" sqref="' . self::xlsCell($freeze_rows, 0) . '"/>');
            $this->writer->write('<selection activeCell="' . self::xlsCell(0, $freeze_columns) . '" activeCellId="0" pane="bottomLeft" sqref="' . self::xlsCell(0, $freeze_columns) . '"/>');
            $this->writer->write('<selection activeCell="' . self::xlsCell($freeze_rows, $freeze_columns) . '" activeCellId="0" pane="bottomRight" sqref="' . self::xlsCell($freeze_rows, $freeze_columns) . '"/>');
        } elseif ($freeze_rows) {
            $this->writer->write('<pane ySplit="' . $freeze_rows . '" topLeftCell="' . self::xlsCell($freeze_rows, 0) . '" activePane="bottomLeft" state="frozen"/>');
            $this->writer->write('<selection activeCell="' . self::xlsCell($freeze_rows, 0) . '" activeCellId="0" pane="bottomLeft" sqref="' . self::xlsCell($freeze_rows, 0) . '"/>');
        } elseif ($freeze_columns) {
            $this->writer->write('<pane xSplit="' . $freeze_columns . '" topLeftCell="' . self::xlsCell(0, $freeze_columns) . '" activePane="topRight" state="frozen"/>');
            $this->writer->write('<selection activeCell="' . self::xlsCell(0, $freeze_columns) . '" activeCellId="0" pane="topRight" sqref="' . self::xlsCell(0, $freeze_columns) . '"/>');
        } else { // not frozen
            $this->writer->write('<selection activeCell="A1" activeCellId="0" pane="topLeft" sqref="A1"/>');
        }
        $this->writer->write('</sheetView>');
        $this->writer->write('</sheetViews>');

        $this->writer->write('<cols>');
        $i = 0;
        if (!empty($col_widths)) {
            foreach ($col_widths as $column_width) {
                $this->writer->write('<col collapsed="false" hidden="false" max="' . ($i + 1) . '" min="' . ($i + 1) . '" style="0" customWidth="true" width="' . floatval($column_width) . '"/>');
                $i++;
            }
        }
        $this->writer->write('<col collapsed="false" hidden="false" max="1024" min="' . ($i + 1) . '" style="0" customWidth="false" width="11.5"/>');
        $this->writer->write('</cols>');

        $this->writer->write('<sheetData>');
    }

    /**
     * 设置列的文本格式
     * @param array|int $types
     */
    public function setColumnTypes($types)
    {
        if (!is_array($types)) {
            $count = intval($types);
            $types = [];
            while ($count > 0) {
                $types[] = '';
                $count--;
            }
        }
        foreach ($types as $k => $v) {
            $index = intval($k);
            if ($index != $k) {
                Util::log('参数错误');
            }
            $number_format = self::numberFormatStandardized($v);
            $number_format_type = self::determineNumberFormatType($number_format);
            $this->columnTypes[$index] = [
                'number_format' => $number_format,//contains excel format like 'YYYY-MM-DD HH:MM:SS'
                'number_format_type' => $number_format_type, //contains friendly format like 'datetime'
                'default_cell_style' => StylesXml::getInstance()->addCellStyle($number_format, null),
            ];
        }
    }

    /**
     * 合并单元格
     * @param string $coordinate
     * @throws Exception
     */
    public function merge(string $coordinate = 'A1:Z1')
    {
        if ($this->finalized)
            return;

        $coordinate = strtoupper($coordinate);
        preg_match_all('/([0-9]+|[A-Z]+)/', $coordinate, $matches);
        if (count($matches[0]) != 4) {
            throw new Exception('参数错误');
        }
        [$startColumn, $startRow, $endColumn, $endRow] = $matches[0];
        $columnEnumMap = $this->getColumnEnumMap();
        if (!isset($columnEnumMap[$startColumn]) || !isset($columnEnumMap[$endColumn])) {
            throw new Exception('参数错误');
        }
        $startRowNum = intval($startRow) - 1;
        $endRowNum = intval($endRow) - 1;
        if ($startRowNum < 0 || $endRowNum < 0) {
            throw new Exception('参数错误');
        }
        $this->mergeCells[] = $coordinate;
    }

    public function getColumnEnumMap(): array
    {
        $list = range('A', 'Z');
        $map = [];
        $index = 0;
        foreach (['', 'A', 'B', 'C'] as $prefix) {
            foreach ($list as $c) {
                $map[$prefix . $c] = $index++;
            }
        }
        return $map;
    }

    //------------------------------------------------------------------
    private static function determineNumberFormatType($num_format): string
    {
        $num_format = preg_replace("/\[(Black|Blue|Cyan|Green|Magenta|Red|White|Yellow)\]/i", "", $num_format);
        if ($num_format == 'GENERAL') return 'n_auto';
        if ($num_format == '@') return 'n_string';
        if ($num_format == '0') return 'n_numeric';
        if (preg_match('/[H]{1,2}:[M]{1,2}(?![^"]*+")/i', $num_format)) return 'n_datetime';
        if (preg_match('/[M]{1,2}:[S]{1,2}(?![^"]*+")/i', $num_format)) return 'n_datetime';
        if (preg_match('/[Y]{2,4}(?![^"]*+")/i', $num_format)) return 'n_date';
        if (preg_match('/[D]{1,2}(?![^"]*+")/i', $num_format)) return 'n_date';
        if (preg_match('/[M]{1,2}(?![^"]*+")/i', $num_format)) return 'n_date';
        if (preg_match('/$(?![^"]*+")/', $num_format)) return 'n_numeric';
        if (preg_match('/%(?![^"]*+")/', $num_format)) return 'n_numeric';
        if (preg_match('/0(?![^"]*+")/', $num_format)) return 'n_numeric';
        return 'n_auto';
    }

    //------------------------------------------------------------------
    private static function numberFormatStandardized($num_format): string
    {
        switch ($num_format) {
            case '':
                $num_format = 'GENERAL';
                break;
            case 'string':
                $num_format = '@';
                break;
            case 'money':
            case 'price':
                $num_format = '#,##0.00';
                break;
            case 'number':
            case 'integer':
                $num_format = '0';
                break;
            case 'date':
                $num_format = 'YYYY-MM-DD';
                break;
            case 'time':
                $num_format = 'HH:MM:SS';
                break;
            case 'datetime':
                $num_format = 'YYYY-MM-DD HH:MM:SS';
                break;
            case 'dollar':
                $num_format = '[$$-1009]#,##0.00;[RED]-[$$-1009]#,##0.00';
                break;
            case 'euro':
                $num_format = '#,##0.00 [$€-407];[RED]-#,##0.00 [$€-407]';
                break;
        }
        $ignore_until = '';
        $escaped = '';
        for ($i = 0, $ix = strlen($num_format); $i < $ix; $i++) {
            $c = $num_format[$i];
            if ($ignore_until == '' && $c == '[')
                $ignore_until = ']';
            else if ($ignore_until == '' && $c == '"')
                $ignore_until = '"';
            else if ($ignore_until == $c)
                $ignore_until = '';
            if ($ignore_until == '' && ($c == ' ' || $c == '-' || $c == '(' || $c == ')') && ($i == 0 || $num_format[$i - 1] != '_'))
                $escaped .= "\\" . $c;
            else
                $escaped .= $c;
        }
        return $escaped;
    }

    /**
     * 添加标题数据
     * @param array $row
     * @param null $colOptions
     * @return $this
     */
    public function addHeader(array $row, $colOptions = null)
    {
        $options = [
            'halign' => 'center',
            'valign' => 'center',
            'font-style' => 'bolder',
        ];
        if (!empty($colOptions)) {
            $options = array_merge($options, $colOptions);
        }
        $this->addRow($row, $options, true);
        return $this;
    }

    /**
     * 添加行数据
     * @param array $row
     * @param null $options
     * @param bool $is_header
     * @return $this
     */
    public function addRow(array $row, $options = null, bool $is_header = false): Sheet
    {
        if (!empty($options)) {
            $ht = isset($options['height']) ? floatval($options['height']) : 12.1;
            $customHt = isset($options['height']) ? true : false;
            $hidden = isset($options['hidden']) ? (bool)($options['hidden']) : false;
            $collapsed = isset($options['collapsed']) ? (bool)($options['collapsed']) : false;
            $this->writer->write('<row collapsed="' . ($collapsed) . '" customFormat="false" customHeight="' . ($customHt) . '" hidden="' . ($hidden) . '" ht="' . ($ht) . '" outlineLevel="0" r="' . ($this->rowCount + 1) . '">');
        } else {
            $this->writer->write('<row collapsed="false" customFormat="false" customHeight="false" hidden="false" ht="12.1" outlineLevel="0" r="' . ($this->rowCount + 1) . '">');
        }

        $style = &$options;
        $index = 0;
        foreach ($row as $v) {
            if (!isset($this->columnTypes[$index])) {
                $this->setColumnTypes([
                    $index => ''
                ]);
            }
            $number_format = $is_header ? 'GENERAL' : $this->columnTypes[$index]['number_format'];
            $number_format_type = $is_header ? 'n_string' : $this->columnTypes[$index]['number_format_type'];
            $cell_style_idx = empty($style)
                ? $this->columnTypes[$index]['default_cell_style']
                : StylesXml::getInstance()->addCellStyle($number_format, json_encode(isset($style[0]) ? $style[$index] : $style));
            $this->writeCell($this->rowCount, $index, $v, $number_format_type, $cell_style_idx);
            $index++;
        }
        $this->writer->write('</row>');
        $this->rowCount++;
        return $this;
    }

    protected function writeCell($row_number, $column_number, $value, $num_format_type, $cell_style_idx)
    {
        $cell_name = self::xlsCell($row_number, $column_number);

        if (!is_scalar($value) || $value === '') { //objects, array, empty
            $this->writer->write('<c r="' . $cell_name . '" s="' . $cell_style_idx . '"/>');
        } elseif (is_string($value) && $value[0] == '=') {
            $this->writer->write('<c r="' . $cell_name . '" s="' . $cell_style_idx . '" t="s"><f>' . Util::xmlSpecialChars($value) . '</f></c>');
        } elseif ($num_format_type == 'n_date') {
            $this->writer->write('<c r="' . $cell_name . '" s="' . $cell_style_idx . '" t="n"><v>' . intval(self::convertDateTime($value)) . '</v></c>');
        } elseif ($num_format_type == 'n_datetime') {
            $this->writer->write('<c r="' . $cell_name . '" s="' . $cell_style_idx . '" t="n"><v>' . self::convertDateTime($value) . '</v></c>');
        } elseif ($num_format_type == 'n_numeric') {
            $this->writer->write('<c r="' . $cell_name . '" s="' . $cell_style_idx . '" t="n"><v>' . Util::xmlSpecialChars($value) . '</v></c>');//int,float,currency
        } elseif ($num_format_type == 'n_string') {
            $this->writer->write('<c r="' . $cell_name . '" s="' . $cell_style_idx . '" t="inlineStr"><is><t>' . Util::xmlSpecialChars($value) . '</t></is></c>');
        } elseif ($num_format_type == 'n_auto' || 1) { //auto-detect unknown column types
            if (!is_string($value) || $value == '0' || ($value[0] != '0' && ctype_digit($value)) || preg_match("/^\-?(0|[1-9][0-9]*)(\.[0-9]+)?$/", $value)) {
                $this->writer->write('<c r="' . $cell_name . '" s="' . $cell_style_idx . '" t="n"><v>' . Util::xmlSpecialChars($value) . '</v></c>');//int,float,currency
            } else { //implied: ($cell_format=='string')
                $this->writer->write('<c r="' . $cell_name . '" s="' . $cell_style_idx . '" t="inlineStr"><is><t>' . Util::xmlSpecialChars($value) . '</t></is></c>');
            }
        }
    }

    //------------------------------------------------------------------
    public static function convertDateTime($date_input) //thanks to Excel::Writer::XLSX::Worksheet.pm (perl)
    {
        $seconds = 0;    # Time expressed as fraction of 24h hours in seconds
        $year = $month = $day = 0;

        $date_time = $date_input;
        if (preg_match("/(\d{4})\-(\d{2})\-(\d{2})/", $date_time, $matches)) {
            list($junk, $year, $month, $day) = $matches;
        }
        if (preg_match("/(\d+):(\d{2}):(\d{2})/", $date_time, $matches)) {
            list($junk, $hour, $min, $sec) = $matches;
            $seconds = ($hour * 60 * 60 + $min * 60 + $sec) / (24 * 60 * 60);
        }

        //using 1900 as epoch, not 1904, ignoring 1904 special case

        # Special cases for Excel.
        if ("$year-$month-$day" == '1899-12-31') return $seconds;    # Excel 1900 epoch
        if ("$year-$month-$day" == '1900-01-00') return $seconds;    # Excel 1900 epoch
        if ("$year-$month-$day" == '1900-02-29') return 60 + $seconds;    # Excel false leapday

        # We calculate the date by calculating the number of days since the epoch
        # and adjust for the number of leap days. We calculate the number of leap
        # days by normalising the year in relation to the epoch. Thus the year 2000
        # becomes 100 for 4 and 100 year leapdays and 400 for 400 year leapdays.
        $epoch = 1900;
        $offset = 0;
        $norm = 300;
        $range = $year - $epoch;

        # Set month days and check for leap year.
        $leap = (($year % 400 == 0) || (($year % 4 == 0) && ($year % 100))) ? 1 : 0;
        $mdays = array(31, ($leap ? 29 : 28), 31, 30, 31, 30, 31, 31, 30, 31, 30, 31);

        # Some boundary checks
        if ($year != 0 || $month != 0 || $day != 0) {
            if ($year < $epoch || $year > 9999) return 0;
            if ($month < 1 || $month > 12) return 0;
            if ($day < 1 || $day > $mdays[$month - 1]) return 0;
        }

        # Accumulate the number of days since the epoch.
        $days = $day;    # Add days for current month
        $days += array_sum(array_slice($mdays, 0, $month - 1));    # Add days for past months
        $days += $range * 365;                      # Add days for past years
        $days += intval(($range) / 4);             # Add leapdays
        $days -= intval(($range + $offset) / 100); # Subtract 100 year leapdays
        $days += intval(($range + $offset + $norm) / 400);  # Add 400 year leapdays
        $days -= $leap;                                      # Already counted above

        # Adjust for Excel erroneously treating 1900 as a leap year.
        if ($days > 59) {
            $days++;
        }

        return $days + $seconds;
    }

    public function finalizeSheet()
    {
        if ($this->finalized) {
            return;
        }
        $this->writer->write('</sheetData>');

        if (!empty($this->mergeCells)) {
            $this->writer->write('<mergeCells>');
            foreach ($this->mergeCells as $range) {
                $range = strtoupper($range);
                $this->writer->write('<mergeCell ref="' . $range . '"/>');
            }
            $this->writer->write('</mergeCells>');
        }

        $max_cell = self::xlsCell($this->rowCount - 1, count($this->columnTypes) - 1);

        if ($this->autoFilter) {
            $this->writer->write('<autoFilter ref="A1:' . $max_cell . '"/>');
        }

        $this->writer->write('<printOptions headings="false" gridLines="false" gridLinesSet="true" horizontalCentered="false" verticalCentered="false"/>');
        $this->writer->write('<pageMargins left="0.5" right="0.5" top="1.0" bottom="1.0" header="0.5" footer="0.5"/>');
        $this->writer->write('<pageSetup blackAndWhite="false" cellComments="none" copies="1" draft="false" firstPageNumber="1" fitToHeight="1" fitToWidth="1" horizontalDpi="300" orientation="portrait" pageOrder="downThenOver" paperSize="1" scale="100" useFirstPageNumber="true" usePrinterDefaults="false" verticalDpi="300"/>');
        $this->writer->write('<headerFooter differentFirst="false" differentOddEven="false">');
        $this->writer->write('<oddHeader>&amp;C&amp;&quot;Times New Roman,Regular&quot;&amp;12&amp;A</oddHeader>');
        $this->writer->write('<oddFooter>&amp;C&amp;&quot;Times New Roman,Regular&quot;&amp;12Page &amp;P</oddFooter>');
        $this->writer->write('</headerFooter>');
        $this->writer->write('</worksheet>');
        $max_cell_tag = '<dimension ref="A1:' . $max_cell . '"/>';
        $padding_length = $this->max_cell_tag_end - $this->max_cell_tag_start - strlen($max_cell_tag);
        $this->writer->fseek($this->max_cell_tag_start);
        $this->writer->write($max_cell_tag . str_repeat(" ", $padding_length));
        $this->writer->close();
        $this->finalized = true;
    }

    //------------------------------------------------------------------
    /*
     * @param $row_number int, zero based
     * @param $column_number int, zero based
     * @param $absolute bool
     * @return Cell label/coordinates, ex: A1, C3, AA42 (or if $absolute==true: $A$1, $C$3, $AA$42)
     * */
    public static function xlsCell($row_number, $column_number, $absolute = false)
    {
        $n = $column_number;
        for ($r = ""; $n >= 0; $n = intval($n / 26) - 1) {
            $r = chr($n % 26 + 0x41) . $r;
        }
        if ($absolute) {
            return '$' . $r . '$' . ($row_number + 1);
        }
        return $r . ($row_number + 1);
    }
}
