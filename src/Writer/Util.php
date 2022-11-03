<?php

namespace Larastu\SimpleExcel\Writer;

class Util
{
    public static function log($message)
    {
        var_dump($message);
        exit;
    }

    public static function xmlSpecialChars($val): string
    {
        //note, badchars does not include \t\n\r (\x09\x0a\x0d)
        static $badchars = "\x00\x01\x02\x03\x04\x05\x06\x07\x08\x0b\x0c\x0e\x0f\x10\x11\x12\x13\x14\x15\x16\x17\x18\x19\x1a\x1b\x1c\x1d\x1e\x1f\x7f";
        static $goodchars = "                              ";
        return strtr(htmlspecialchars($val, ENT_QUOTES), $badchars, $goodchars);//strtr appears to be faster than str_replace
    }

    public static function addGetIndex(&$haystack, $needle)
    {
        $existing_idx = array_search($needle, $haystack, true);
        if ($existing_idx === false) {
            $existing_idx = count($haystack);
            $haystack[] = $needle;
        }
        return $existing_idx;
    }
}
