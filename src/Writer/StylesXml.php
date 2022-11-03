<?php

namespace Larastu\SimpleExcel\Writer;

class StylesXml
{
    private static $instance = null;

    /**
     * @var array 所使用的单元格样式列表
     */
    private $cellStyles = [];
    /**
     * @var array 所使用的文本格式列表
     */
    private $textFormats = [];

    public static function getInstance(): StylesXml
    {
        if (!is_null(self::$instance)) {
            return self::$instance;
        }

        return self::$instance = new self();
    }

    private function __construct()
    {
    }

    private function __clone()
    {
    }

    private function __wakeup()
    {
    }


    public function addCellStyle($textFormat, $cellStyleString)
    {
        $number_format_idx = Util::addGetIndex($this->textFormats, $textFormat);
        $lookup_string = $number_format_idx . ";" . $cellStyleString;
        return Util::addGetIndex($this->cellStyles, $lookup_string);
    }

    public function writeStylesXML()
    {
        $r = self::styleFontIndexes();
        $fills = $r['fills'];
        $fonts = $r['fonts'];
        $borders = $r['borders'];
        $style_indexes = $r['styles'];

        $fileWriter = new FileWriter();
        $fileWriter->write('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' . "\n");
        $fileWriter->write('<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">');
        $fileWriter->write('<numFmts count="' . count($this->textFormats) . '">');
        foreach ($this->textFormats as $i => $v) {
            $fileWriter->write('<numFmt numFmtId="' . (164 + $i) . '" formatCode="' . Util::xmlSpecialChars($v) . '" />');
        }
        $fileWriter->write('</numFmts>');

        $fileWriter->write('<fonts count="' . (count($fonts)) . '">');
        $fileWriter->write('<font><name val="Arial"/><charset val="1"/><family val="2"/><sz val="10"/></font>');
        $fileWriter->write('<font><name val="Arial"/><family val="0"/><sz val="10"/></font>');
        $fileWriter->write('<font><name val="Arial"/><family val="0"/><sz val="10"/></font>');
        $fileWriter->write('<font><name val="Arial"/><family val="0"/><sz val="10"/></font>');
        foreach ($fonts as $font) {
            if (!empty($font)) { //fonts have 4 empty placeholders in array to offset the 4 static xml entries above
                $f = json_decode($font, true);
                $fileWriter->write('<font>');
                $fileWriter->write('<name val="' . htmlspecialchars($f['name']) . '"/><charset val="1"/><family val="' . intval($f['family']) . '"/>');
                $fileWriter->write('<sz val="' . intval($f['size']) . '"/>');
                if (!empty($f['color'])) {
                    $fileWriter->write('<color rgb="' . strval($f['color']) . '"/>');
                }
                if (!empty($f['bold'])) {
                    $fileWriter->write('<b val="true"/>');
                }
                if (!empty($f['italic'])) {
                    $fileWriter->write('<i val="true"/>');
                }
                if (!empty($f['underline'])) {
                    $fileWriter->write('<u val="single"/>');
                }
                if (!empty($f['strike'])) {
                    $fileWriter->write('<strike val="true"/>');
                }
                $fileWriter->write('</font>');
            }
        }
        $fileWriter->write('</fonts>');

        $fileWriter->write('<fills count="' . (count($fills)) . '">');
        $fileWriter->write('<fill><patternFill patternType="none"/></fill>');
        $fileWriter->write('<fill><patternFill patternType="gray125"/></fill>');
        foreach ($fills as $fill) {
            if (!empty($fill)) { //fills have 2 empty placeholders in array to offset the 2 static xml entries above
                $fileWriter->write('<fill><patternFill patternType="solid"><fgColor rgb="' . strval($fill) . '"/><bgColor indexed="64"/></patternFill></fill>');
            }
        }
        $fileWriter->write('</fills>');

        $fileWriter->write('<borders count="' . (count($borders)) . '">');
        $fileWriter->write('<border diagonalDown="false" diagonalUp="false"><left/><right/><top/><bottom/><diagonal/></border>');
        foreach ($borders as $border) {
            if (!empty($border)) { //fonts have an empty placeholder in the array to offset the static xml entry above
                $pieces = json_decode($border, true);
                $border_style = !empty($pieces['style']) ? $pieces['style'] : 'hair';
                $border_color = !empty($pieces['color']) ? '<color rgb="' . strval($pieces['color']) . '"/>' : '';
                $fileWriter->write('<border diagonalDown="false" diagonalUp="false">');
                foreach (array('left', 'right', 'top', 'bottom') as $side) {
                    $show_side = in_array($side, $pieces['side']) ? true : false;
                    $fileWriter->write($show_side ? "<$side style=\"$border_style\">$border_color</$side>" : "<$side/>");
                }
                $fileWriter->write('<diagonal/>');
                $fileWriter->write('</border>');
            }
        }
        $fileWriter->write('</borders>');

        $fileWriter->write('<cellStyleXfs count="20">');
        $fileWriter->write('<xf applyAlignment="true" applyBorder="true" applyFont="true" applyProtection="true" borderId="0" fillId="0" fontId="0" numFmtId="164">');
        $fileWriter->write('<alignment horizontal="general" indent="0" shrinkToFit="false" textRotation="0" vertical="bottom" wrapText="false"/>');
        $fileWriter->write('<protection hidden="false" locked="true"/>');
        $fileWriter->write('</xf>');
        $fileWriter->write('<xf applyAlignment="false" applyBorder="false" applyFont="true" applyProtection="false" borderId="0" fillId="0" fontId="1" numFmtId="0"/>');
        $fileWriter->write('<xf applyAlignment="false" applyBorder="false" applyFont="true" applyProtection="false" borderId="0" fillId="0" fontId="1" numFmtId="0"/>');
        $fileWriter->write('<xf applyAlignment="false" applyBorder="false" applyFont="true" applyProtection="false" borderId="0" fillId="0" fontId="2" numFmtId="0"/>');
        $fileWriter->write('<xf applyAlignment="false" applyBorder="false" applyFont="true" applyProtection="false" borderId="0" fillId="0" fontId="2" numFmtId="0"/>');
        $fileWriter->write('<xf applyAlignment="false" applyBorder="false" applyFont="true" applyProtection="false" borderId="0" fillId="0" fontId="0" numFmtId="0"/>');
        $fileWriter->write('<xf applyAlignment="false" applyBorder="false" applyFont="true" applyProtection="false" borderId="0" fillId="0" fontId="0" numFmtId="0"/>');
        $fileWriter->write('<xf applyAlignment="false" applyBorder="false" applyFont="true" applyProtection="false" borderId="0" fillId="0" fontId="0" numFmtId="0"/>');
        $fileWriter->write('<xf applyAlignment="false" applyBorder="false" applyFont="true" applyProtection="false" borderId="0" fillId="0" fontId="0" numFmtId="0"/>');
        $fileWriter->write('<xf applyAlignment="false" applyBorder="false" applyFont="true" applyProtection="false" borderId="0" fillId="0" fontId="0" numFmtId="0"/>');
        $fileWriter->write('<xf applyAlignment="false" applyBorder="false" applyFont="true" applyProtection="false" borderId="0" fillId="0" fontId="0" numFmtId="0"/>');
        $fileWriter->write('<xf applyAlignment="false" applyBorder="false" applyFont="true" applyProtection="false" borderId="0" fillId="0" fontId="0" numFmtId="0"/>');
        $fileWriter->write('<xf applyAlignment="false" applyBorder="false" applyFont="true" applyProtection="false" borderId="0" fillId="0" fontId="0" numFmtId="0"/>');
        $fileWriter->write('<xf applyAlignment="false" applyBorder="false" applyFont="true" applyProtection="false" borderId="0" fillId="0" fontId="0" numFmtId="0"/>');
        $fileWriter->write('<xf applyAlignment="false" applyBorder="false" applyFont="true" applyProtection="false" borderId="0" fillId="0" fontId="0" numFmtId="0"/>');
        $fileWriter->write('<xf applyAlignment="false" applyBorder="false" applyFont="true" applyProtection="false" borderId="0" fillId="0" fontId="1" numFmtId="43"/>');
        $fileWriter->write('<xf applyAlignment="false" applyBorder="false" applyFont="true" applyProtection="false" borderId="0" fillId="0" fontId="1" numFmtId="41"/>');
        $fileWriter->write('<xf applyAlignment="false" applyBorder="false" applyFont="true" applyProtection="false" borderId="0" fillId="0" fontId="1" numFmtId="44"/>');
        $fileWriter->write('<xf applyAlignment="false" applyBorder="false" applyFont="true" applyProtection="false" borderId="0" fillId="0" fontId="1" numFmtId="42"/>');
        $fileWriter->write('<xf applyAlignment="false" applyBorder="false" applyFont="true" applyProtection="false" borderId="0" fillId="0" fontId="1" numFmtId="9"/>');
        $fileWriter->write('</cellStyleXfs>');

        $fileWriter->write('<cellXfs count="' . (count($style_indexes)) . '">');

        foreach ($style_indexes as $v) {
            $applyAlignment = isset($v['alignment']) ? 'true' : 'false';
            $wrapText = !empty($v['wrap_text']) ? 'true' : 'false';
            $horizAlignment = isset($v['halign']) ? $v['halign'] : 'general';
            $vertAlignment = isset($v['valign']) ? $v['valign'] : 'bottom';
            $applyBorder = isset($v['border_idx']) ? 'true' : 'false';
            $applyFont = 'true';
            $borderIdx = isset($v['border_idx']) ? intval($v['border_idx']) : 0;
            $fillIdx = isset($v['fill_idx']) ? intval($v['fill_idx']) : 0;
            $fontIdx = isset($v['font_idx']) ? intval($v['font_idx']) : 0;
            $fileWriter->write('<xf applyAlignment="' . $applyAlignment . '" applyBorder="' . $applyBorder . '" applyFont="' . $applyFont . '" applyProtection="false" borderId="' . ($borderIdx) . '" fillId="' . ($fillIdx) . '" fontId="' . ($fontIdx) . '" numFmtId="' . (164 + $v['num_fmt_idx']) . '" xfId="0">');
            $fileWriter->write('	<alignment horizontal="' . $horizAlignment . '" vertical="' . $vertAlignment . '" textRotation="0" wrapText="' . $wrapText . '" indent="0" shrinkToFit="false"/>');
            $fileWriter->write('	<protection locked="true" hidden="false"/>');
            $fileWriter->write('</xf>');
        }
        $fileWriter->write('</cellXfs>');
        $fileWriter->write('<cellStyles count="6">');
        $fileWriter->write('<cellStyle builtinId="0" customBuiltin="false" name="Normal" xfId="0"/>');
        $fileWriter->write('<cellStyle builtinId="3" customBuiltin="false" name="Comma" xfId="15"/>');
        $fileWriter->write('<cellStyle builtinId="6" customBuiltin="false" name="Comma [0]" xfId="16"/>');
        $fileWriter->write('<cellStyle builtinId="4" customBuiltin="false" name="Currency" xfId="17"/>');
        $fileWriter->write('<cellStyle builtinId="7" customBuiltin="false" name="Currency [0]" xfId="18"/>');
        $fileWriter->write('<cellStyle builtinId="5" customBuiltin="false" name="Percent" xfId="19"/>');
        $fileWriter->write('</cellStyles>');
        $fileWriter->write('</styleSheet>');
        $fileWriter->close();
        return $fileWriter->getTempFile();
    }

    protected function styleFontIndexes(): array
    {
        $border_allowed = ['left', 'right', 'top', 'bottom'];
        $border_style_allowed = [
            'thin', 'medium', 'thick', 'dashDot', 'dashDotDot',
            'dashed', 'dotted', 'double', 'hair', 'mediumDashDot',
            'mediumDashDotDot', 'mediumDashed', 'slantDashDot'
        ];
        $horizontal_allowed = ['general', 'left', 'right', 'justify', 'center'];
        $vertical_allowed = ['bottom', 'center', 'distributed', 'top'];
        $default_font = ['size' => '10', 'name' => 'Arial', 'family' => '2'];
        $fills = ['', ''];//2 placeholders for static xml later
        $fonts = ['', '', '', ''];//4 placeholders for static xml later
        $borders = [''];//1 placeholder for static xml later
        $style_indexes = [];
        foreach ($this->cellStyles as $i => $cell_style_string) {
            $semi_colon_pos = strpos($cell_style_string, ";");
            $number_format_idx = substr($cell_style_string, 0, $semi_colon_pos);
            $style_json_string = substr($cell_style_string, $semi_colon_pos + 1);
            $style = @json_decode($style_json_string, true);

            $style_indexes[$i] = array('num_fmt_idx' => $number_format_idx);//initialize entry
            if (isset($style['border']) && is_string($style['border']))//border is a comma delimited str
            {
                $border_value['side'] = array_intersect(explode(",", $style['border']), $border_allowed);
                if (isset($style['border-style']) && in_array($style['border-style'], $border_style_allowed)) {
                    $border_value['style'] = $style['border-style'];
                }
                if (isset($style['border-color']) && is_string($style['border-color']) && $style['border-color'][0] == '#') {
                    $v = substr($style['border-color'], 1, 6);
                    $v = strlen($v) == 3 ? $v[0] . $v[0] . $v[1] . $v[1] . $v[2] . $v[2] : $v;// expand cf0 => ccff00
                    $border_value['color'] = "FF" . strtoupper($v);
                }
                $style_indexes[$i]['border_idx'] = Util::addGetIndex($borders, json_encode($border_value));
            }
            if (isset($style['fill']) && is_string($style['fill']) && $style['fill'][0] == '#') {
                $v = substr($style['fill'], 1, 6);
                $v = strlen($v) == 3 ? $v[0] . $v[0] . $v[1] . $v[1] . $v[2] . $v[2] : $v;// expand cf0 => ccff00
                $style_indexes[$i]['fill_idx'] = Util::addGetIndex($fills, "FF" . strtoupper($v));
            }
            if (isset($style['halign']) && in_array($style['halign'], $horizontal_allowed)) {
                $style_indexes[$i]['alignment'] = true;
                $style_indexes[$i]['halign'] = $style['halign'];
            }
            if (isset($style['valign']) && in_array($style['valign'], $vertical_allowed)) {
                $style_indexes[$i]['alignment'] = true;
                $style_indexes[$i]['valign'] = $style['valign'];
            }
            if (isset($style['wrap_text'])) {
                $style_indexes[$i]['alignment'] = true;
                $style_indexes[$i]['wrap_text'] = (bool)$style['wrap_text'];
            }

            $font = $default_font;
            if (isset($style['font-size'])) {
                $font['size'] = floatval($style['font-size']);//floatval to allow "10.5" etc
            }
            if (isset($style['font']) && is_string($style['font'])) {
                if ($style['font'] == 'Comic Sans MS') {
                    $font['family'] = 4;
                }
                if ($style['font'] == 'Times New Roman') {
                    $font['family'] = 1;
                }
                if ($style['font'] == 'Courier New') {
                    $font['family'] = 3;
                }
                $font['name'] = strval($style['font']);
            }
            if (isset($style['font-style']) && is_string($style['font-style'])) {
                if (strpos($style['font-style'], 'bold') !== false) {
                    $font['bold'] = true;
                }
                if (strpos($style['font-style'], 'italic') !== false) {
                    $font['italic'] = true;
                }
                if (strpos($style['font-style'], 'strike') !== false) {
                    $font['strike'] = true;
                }
                if (strpos($style['font-style'], 'underline') !== false) {
                    $font['underline'] = true;
                }
            }
            if (isset($style['color']) && is_string($style['color']) && $style['color'][0] == '#') {
                $v = substr($style['color'], 1, 6);
                $v = strlen($v) == 3 ? $v[0] . $v[0] . $v[1] . $v[1] . $v[2] . $v[2] : $v;// expand cf0 => ccff00
                $font['color'] = "FF" . strtoupper($v);
            }
            if ($font != $default_font) {
                $style_indexes[$i]['font_idx'] = Util::addGetIndex($fonts, json_encode($font));
            }
        }
        return [
            'fills' => $fills,
            'fonts' => $fonts,
            'borders' => $borders,
            'styles' => $style_indexes
        ];
    }

}
