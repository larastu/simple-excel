<?php

namespace Larastu\SimpleExcel\Reader;

use DateInterval;
use DateTime;
use DateTimeZone;
use SimpleXMLElement;

class Styles
{
    private static $BuiltinFormats = [
        0 => '',
        1 => '0',
        2 => '0.00',
        3 => '#,##0',
        4 => '#,##0.00',

        9 => '0%',
        10 => '0.00%',
        11 => '0.00E+00',
        12 => '# ?/?',
        13 => '# ??/??',
        14 => 'mm-dd-yy',
        15 => 'd-mmm-yy',
        16 => 'd-mmm',
        17 => 'mmm-yy',
        18 => 'h:mm AM/PM',
        19 => 'h:mm:ss AM/PM',
        20 => 'h:mm',
        21 => 'h:mm:ss',
        22 => 'm/d/yy h:mm',

        37 => '#,##0 ;(#,##0)',
        38 => '#,##0 ;[Red](#,##0)',
        39 => '#,##0.00;(#,##0.00)',
        40 => '#,##0.00;[Red](#,##0.00)',

        45 => 'mm:ss',
        46 => '[h]:mm:ss',
        47 => 'mmss.0',
        48 => '##0.0E+0',
        49 => '@',

        // CHT & CHS
        27 => '[$-404]e/m/d',
        30 => 'm/d/yy',
        36 => '[$-404]e/m/d',
        50 => '[$-404]e/m/d',
        57 => '[$-404]e/m/d',

        // THA
        59 => 't0',
        60 => 't0.00',
        61 => 't#,##0',
        62 => 't#,##0.00',
        67 => 't0%',
        68 => 't0.00%',
        69 => 't# ?/?',
        70 => 't# ??/??'
    ];

    private static $DateReplacements = array(
        'All' => array(
            '\\' => '',
            'am/pm' => 'A',
            'yyyy' => 'Y',
            'yy' => 'y',
            'mmmmm' => 'M',
            'mmmm' => 'F',
            'mmm' => 'M',
            ':mm' => ':i',
            'mm' => 'm',
            'm' => 'n',
            'dddd' => 'l',
            'ddd' => 'D',
            'dd' => 'd',
            'd' => 'j',
            'ss' => 's',
            '.s' => ''
        ),
        '24H' => array(
            'hh' => 'H',
            'h' => 'G'
        ),
        '12H' => array(
            'hh' => 'h',
            'h' => 'G'
        )
    );

    /**
     * @var DateTime
     */
    private static $BaseDate;
    /**
     * @var mixed
     */
    private static $DecimalSeparator = '';
    /**
     * @var mixed
     */
    private static $ThousandSeparator = '';
    /**
     * @var mixed
     */
    private static $CurrencyCode = '';

    private static $RuntimeInfo = [
        'GMPSupported' => false
    ];

    private static $instance = null;

    private $Options = array(
        'TempDir' => '',
        'ReturnDateTimeObjects' => false
    );

    /**
     * @var array
     */
    private $parsedFormatCache = [];
    /**
     * @var array
     */
    private $styles = [];
    private $formats = [];

    public static function getInstance(): Styles
    {
        if (is_null(self::$instance)) {
            self::$instance = new static();
        }

        return self::$instance;
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

    private function __destruct()
    {
    }

    public function init($xmlString)
    {
        $stylesXML = new SimpleXMLElement($xmlString);
        if ($stylesXML->cellXfs && $stylesXML->cellXfs->xf) {
            foreach ($stylesXML->cellXfs->xf as $xf) {
                // Format #0 is a special case - it is the "General" format that is applied regardless of applyNumberFormat
                if ($xf->attributes()->applyNumberFormat || (0 == (int)$xf->attributes()->numFmtId)) {
                    $formatId = (int)$xf->attributes()->numFmtId;
                    // If format ID >= 164, it is a custom format and should be read from styleSheet\numFmts
                    $this->styles[] = $formatId;
                } else {
                    // 0 for "General" format
                    $this->styles[] = 0;
                }
            }
        }

        if ($stylesXML->numFmts && $stylesXML->numFmts->numFmt) {
            foreach ($stylesXML->numFmts->numFmt as $numFmt) {
                $this->formats[(int)$numFmt->attributes()->numFmtId] = (string)$numFmt->attributes()->formatCode;
            }
        }

        // Setting base date
        if (!self::$BaseDate) {
            self::$BaseDate = new DateTime;
            self::$BaseDate->setTimezone(new DateTimeZone('UTC'));
            self::$BaseDate->setDate(1900, 1, 0);
            self::$BaseDate->setTime(0, 0);
        }

        // Decimal and thousand separators
        if (!self::$DecimalSeparator && !self::$ThousandSeparator && !self::$CurrencyCode) {
            $Locale = localeconv();
            self::$DecimalSeparator = $Locale['decimal_point'];
            self::$ThousandSeparator = $Locale['thousands_sep'];
            self::$CurrencyCode = $Locale['int_curr_symbol'];
        }

        if (function_exists('gmp_gcd')) {
            self::$RuntimeInfo['GMPSupported'] = true;
        }
    }

    public function formatValue($value, $index)
    {
        if (!is_numeric($value)) {
            return $value;
        }

        if (isset($this->styles[$index]) && ($this->styles[$index] !== false)) {
            $index = $this->styles[$index];
        } else {
            return $value;
        }

        // A special case for the "General" format
        if ($index == 0) {
            return $this->generalFormat($value);
        }

        $format = [];

        if (isset($this->parsedFormatCache[$index])) {
            $format = $this->parsedFormatCache[$index];
        }

        if (!$format) {
            $format = array(
                'Code' => false,
                'Type' => false,
                'Scale' => 1,
                'Thousands' => false,
                'Currency' => false
            );

            if (isset(self::$BuiltinFormats[$index])) {
                $format['Code'] = self::$BuiltinFormats[$index];
            } elseif (isset($this->formats[$index])) {
                $format['Code'] = $this->formats[$index];
            }

            // Format code found, now parsing the format
            if ($format['Code']) {
                $Sections = explode(';', $format['Code']);
                $format['Code'] = $Sections[0];

                switch (count($Sections)) {
                    case 2:
                        if ($value < 0) {
                            $format['Code'] = $Sections[1];
                        }
                        break;
                    case 3:
                    case 4:
                        if ($value < 0) {
                            $format['Code'] = $Sections[1];
                        } elseif ($value == 0) {
                            $format['Code'] = $Sections[2];
                        }
                        break;
                }
            }

            // Stripping colors
            $format['Code'] = trim(preg_replace('{^\[[[:alpha:]]+\]}i', '', $format['Code']));

            // Percentages
            if (substr($format['Code'], -1) == '%') {
                $format['Type'] = 'Percentage';
            } elseif (preg_match('{^(\[\$[[:alpha:]]*-[0-9A-F]*\])*[hmsdy]}i', $format['Code'])) {
                $format['Type'] = 'DateTime';

                $format['Code'] = trim(preg_replace('{^(\[\$[[:alpha:]]*-[0-9A-F]*\])}i', '', $format['Code']));
                $format['Code'] = strtolower($format['Code']);

                $format['Code'] = strtr($format['Code'], self::$DateReplacements['All']);
                if (strpos($format['Code'], 'A') === false) {
                    $format['Code'] = strtr($format['Code'], self::$DateReplacements['24H']);
                } else {
                    $format['Code'] = strtr($format['Code'], self::$DateReplacements['12H']);
                }
            } elseif ($format['Code'] == '[$EUR ]#,##0.00_-') {
                $format['Type'] = 'Euro';
            } else {
                // Removing skipped characters
                $format['Code'] = preg_replace('{_.}', '', $format['Code']);
                // Removing unnecessary escaping
                $format['Code'] = preg_replace("{\\\\}", '', $format['Code']);
                // Removing string quotes
                $format['Code'] = str_replace(array('"', '*'), '', $format['Code']);
                // Removing thousands separator
                if (strpos($format['Code'], '0,0') !== false || strpos($format['Code'], '#,#') !== false) {
                    $format['Thousands'] = true;
                }
                $format['Code'] = str_replace(array('0,0', '#,#'), array('00', '##'), $format['Code']);

                // Scaling (Commas indicate the power)
                $Scale = 1;
                $Matches = array();
                if (preg_match('{(0|#)(,+)}', $format['Code'], $Matches)) {
                    $Scale = pow(1000, strlen($Matches[2]));
                    // Removing the commas
                    $format['Code'] = preg_replace(array('{0,+}', '{#,+}'), array('0', '#'), $format['Code']);
                }

                $format['Scale'] = $Scale;

                if (preg_match('{#?.*\?\/\?}', $format['Code'])) {
                    $format['Type'] = 'Fraction';
                } else {
                    $format['Code'] = str_replace('#', '', $format['Code']);

                    $Matches = array();
                    if (preg_match('{(0+)(\.?)(0*)}', preg_replace('{\[[^\]]+\]}', '', $format['Code']), $Matches)) {
                        $Integer = $Matches[1];
                        $DecimalPoint = $Matches[2];
                        $Decimals = $Matches[3];

                        $format['MinWidth'] = strlen($Integer) + strlen($DecimalPoint) + strlen($Decimals);
                        $format['Decimals'] = $Decimals;
                        $format['Precision'] = strlen($format['Decimals']);
                        $format['Pattern'] = '%0' . $format['MinWidth'] . '.' . $format['Precision'] . 'f';
                    }
                }

                $Matches = array();
                if (preg_match('{\[\$(.*)\]}u', $format['Code'], $Matches)) {
                    $CurrCode = $Matches[1];
                    $CurrCode = explode('-', $CurrCode);
                    if ($CurrCode) {
                        $CurrCode = $CurrCode[0];
                    }

                    if (!$CurrCode) {
                        $CurrCode = self::$CurrencyCode;
                    }

                    $format['Currency'] = $CurrCode;
                }
                $format['Code'] = trim($format['Code']);
            }

            $this->parsedFormatCache[$index] = $format;
        }

        // Applying format to value
        if ($format) {
            if ($format['Code'] == '@') {
                return (string)$value;
            } // Percentages
            elseif ($format['Type'] == 'Percentage') {
                if ($format['Code'] === '0%') {
                    $value = round(100 * $value) . '%';
                } else {
                    $value = sprintf('%.2f%%', round(100 * $value, 2));
                }
            } // Dates and times
            elseif ($format['Type'] == 'DateTime') {
                $Days = (int)$value;
                // Correcting for Feb 29, 1900
                if ($Days > 60) {
                    $Days--;
                }

                // At this point time is a fraction of a day
                $Time = ($value - (int)$value);
                $Seconds = 0;
                if ($Time) {
                    // Here time is converted to seconds
                    // Some loss of precision will occur
                    $Seconds = (int)($Time * 86400);
                }

                $value = clone self::$BaseDate;
                $value->add(new DateInterval('P' . $Days . 'D' . ($Seconds ? 'T' . $Seconds . 'S' : '')));

                if (!$this->Options['ReturnDateTimeObjects']) {
                    $value = $value->format($format['Code']);
                } else {
                    // A DateTime object is returned
                }
            } elseif ($format['Type'] == 'Euro') {
                $value = 'EUR ' . sprintf('%1.2f', $value);
            } else {
                // Fractional numbers
                if ($format['Type'] == 'Fraction' && ($value != (int)$value)) {
                    $Integer = floor(abs($value));
                    $Decimal = fmod(abs($value), 1);
                    // Removing the integer part and decimal point
                    $Decimal *= pow(10, strlen($Decimal) - 2);
                    $DecimalDivisor = pow(10, strlen($Decimal));

                    if (self::$RuntimeInfo['GMPSupported']) {
                        $GCD = gmp_strval(gmp_gcd($Decimal, $DecimalDivisor));
                    } else {
                        $GCD = self::GCD($Decimal, $DecimalDivisor);
                    }

                    $DecimalPart = 0; ##################################################
                    $AdjDecimal = $DecimalPart / $GCD;
                    $AdjDecimalDivisor = $DecimalDivisor / $GCD;

                    if (
                        strpos($format['Code'], '0') !== false ||
                        strpos($format['Code'], '#') !== false ||
                        substr($format['Code'], 0, 3) == '? ?'
                    ) {
                        // The integer part is shown separately apart from the fraction
                        $value = ($value < 0 ? '-' : '') .
                        $Integer ? $Integer . ' ' : '' .
                            $AdjDecimal . '/' .
                            $AdjDecimalDivisor;
                    } else {
                        // The fraction includes the integer part
                        $AdjDecimal += $Integer * $AdjDecimalDivisor;
                        $value = ($value < 0 ? '-' : '') .
                            $AdjDecimal . '/' .
                            $AdjDecimalDivisor;
                    }
                } else {
                    // Scaling
                    $value = $value / $format['Scale'];

                    if (!empty($format['MinWidth']) && $format['Decimals']) {
                        if ($format['Thousands']) {
                            $value = number_format($value, $format['Precision'],
                                self::$DecimalSeparator, self::$ThousandSeparator);
                        } else {
                            $value = sprintf($format['Pattern'], $value);
                        }

                        $value = preg_replace('{(0+)(\.?)(0*)}', $value, $format['Code']);
                    }
                }

                // Currency/Accounting
                if ($format['Currency']) {
                    $value = preg_replace('', $format['Currency'], $value);
                }
            }

        }

        return $value;
    }

    /**
     * Attempts to approximate Excel's "general" format.
     *
     * @param mixed Value
     *
     * @return mixed Result
     */
    public function generalFormat($value)
    {
        // Numeric format
        if (is_numeric($value)) {
            $value = (float)$value;
        }
        return $value;
    }

    /**
     * Helper function for greatest common divisor calculation in case GMP extension is
     *    not enabled
     *
     * @param int Number #1
     * @param int Number #2
     *
     * @param int Greatest common divisor
     */
    public static function GCD($A, $B)
    {
        $A = abs($A);
        $B = abs($B);
        if ($A + $B == 0) {
            return 0;
        } else {
            $C = 1;

            while ($A > 0) {
                $C = $A;
                $A = $B % $A;
                $B = $C;
            }

            return $C;
        }
    }
}
