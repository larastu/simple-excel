<?php

namespace Larastu\SimpleExcel\Writer;

class FileWriter
{
    protected $fd = null;
    protected $buffer = '';
    protected $checkUtf8 = false;
    /**
     * @var false|string
     */
    private $tempFile;

    public function __construct($file = null, $checkUtf8 = false)
    {
        if (is_null($file)) {
            $file = $this->tempFilename();
        }
        $this->checkUtf8 = $checkUtf8;
        $this->fd = fopen($file, 'w');
        if ($this->fd === false) {
            Util::log("Unable to open $file for writing.");
        }
        $this->tempFile = $file;
    }

    public function getTempFile()
    {
        return $this->tempFile;
    }

    protected function tempFilename()
    {
        $tmpdir = sys_get_temp_dir();
        return tempnam($tmpdir, "temp_file_");
    }

    public function write($string)
    {
        $this->buffer .= $string;
        if (isset($this->buffer[8191])) {
            $this->purge();
        }
    }

    protected function purge()
    {
        if ($this->fd) {
            if ($this->checkUtf8 && !self::isValidUTF8($this->buffer)) {
                Util::log("Error, invalid UTF8 encoding detected.");
                $this->checkUtf8 = false;
            }
            fwrite($this->fd, $this->buffer);
            $this->buffer = '';
        }
    }

    public function close()
    {
        $this->purge();
        if ($this->fd) {
            fclose($this->fd);
            $this->fd = null;
        }
    }

    public function __destruct()
    {
        $this->close();
    }

    public function ftell()
    {
        if ($this->fd) {
            $this->purge();
            return ftell($this->fd);
        }
        return -1;
    }

    public function fseek($pos)
    {
        if ($this->fd) {
            $this->purge();
            return fseek($this->fd, $pos);
        }
        return -1;
    }

    protected static function isValidUTF8($string)
    {
        if (function_exists('mb_check_encoding')) {
            return mb_check_encoding($string, 'UTF-8') ? true : false;
        }
        return preg_match("//u", $string) ? true : false;
    }
}
