<?php

namespace Larastu\SimpleExcel\Reader;

use XMLReader;

class SharedString
{
    const SHARED_STRING_CACHE_LIMIT = 50000;

    private static $instance = null;
    /**
     * @var XMLReader
     */
    private $xmlReader = null;
    private $xmlPath;
    private $stringCache = null;

    private $sharedStringCount;
    /**
     * @var int
     */
    private $sharedStringIndex;
    /**
     * @var string|null
     */
    private $lastSharedStringValue;
    /**
     * @var bool
     */
    private $SSForwarded;
    /**
     * @var false
     */
    private $SSOpen;

    public static function getInstance(): SharedString
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
        if ($this->xmlReader instanceof XMLReader) {
            $this->xmlReader->close();
            unset($this->xmlReader);
        }
    }

    public function init($filepath)
    {
        $this->xmlPath = $filepath;
        $this->xmlReader = new XMLReader();
        $this->xmlReader->open($filepath);
        $this->prepareSharedStringCache();
    }

    private function prepareSharedStringCache(): void
    {
        while ($this->xmlReader->read()) {
            if ($this->xmlReader->name == 'sst') {
                $this->sharedStringCount = $this->xmlReader->getAttribute('count');
                break;
            }
        }

        if (!$this->sharedStringCount || (self::SHARED_STRING_CACHE_LIMIT < $this->sharedStringCount && self::SHARED_STRING_CACHE_LIMIT !== null)) {
            return;
        }

        $CacheIndex = 0;
        $CacheValue = '';
        while ($this->xmlReader->read()) {
            switch ($this->xmlReader->name) {
                case 'si':
                    if ($this->xmlReader->nodeType == XMLReader::END_ELEMENT) {
                        $this->stringCache[$CacheIndex] = $CacheValue;
                        $CacheIndex++;
                        $CacheValue = '';
                    }
                    break;
                case 't':
                    if ($this->xmlReader->nodeType != XMLReader::END_ELEMENT) {
                        $CacheValue .= $this->xmlReader->readString();
                    }
                    break;
            }
        }

        $this->xmlReader->close();
    }

    /**
     * Retrieves a shared string value by its index
     *
     * @param int Shared string index
     *
     * @return string Value
     */
    public function getSharedString($index): ?string
    {
        if (is_null($this->xmlReader)) {
            return '';
        }
        if ((self::SHARED_STRING_CACHE_LIMIT === null || self::SHARED_STRING_CACHE_LIMIT > 0) && !empty($this->stringCache)) {
            return $this->stringCache[$index] ?? '';
        }

        // If the desired index is before the current, rewind the XML
        if ($this->sharedStringIndex > $index) {
            $this->SSOpen = false;
            $this->xmlReader->close();
            $this->xmlReader->open($this->xmlPath);
            $this->sharedStringIndex = 0;
            $this->lastSharedStringValue = null;
            $this->SSForwarded = false;
        }

        // Finding the unique string count (if not already read)
        if ($this->sharedStringIndex == 0 && !$this->sharedStringCount) {
            while ($this->xmlReader->read()) {
                if ($this->xmlReader->name == 'sst') {
                    $this->sharedStringCount = $this->xmlReader->getAttribute('uniqueCount');
                    break;
                }
            }
        }

        // If index of the desired string is larger than possible, don't even bother.
        if ($this->sharedStringCount && ($index >= $this->sharedStringCount)) {
            return '';
        }

        // If an index with the same value as the last already fetched is requested
        // (any further traversing the tree would get us further away from the node)
        if (($index == $this->sharedStringIndex) && ($this->lastSharedStringValue !== null)) {
            return $this->lastSharedStringValue;
        }

        // Find the correct <si> node with the desired index
        while ($this->sharedStringIndex <= $index) {
            // SSForwarded is set further to avoid double reading in case nodes are skipped.
            if ($this->SSForwarded) {
                $this->SSForwarded = false;
            } else {
                if (!$this->xmlReader->read()) {
                    break;
                }
            }

            if ($this->xmlReader->name == 'si') {
                if ($this->xmlReader->nodeType == XMLReader::END_ELEMENT) {
                    $this->SSOpen = false;
                } else {
                    $this->SSOpen = true;
                    if ($this->sharedStringIndex >= $index) {
                        break;
                    }
                    $this->SSOpen = false;
                    $this->xmlReader->next('si');
                    $this->SSForwarded = true;
                }
                $this->sharedStringIndex++;
            }
        }

        $value = '';

        // Extract the value from the shared string
        if ($this->SSOpen && ($this->sharedStringIndex == $index)) {
            while ($this->xmlReader->read()) {
                switch ($this->xmlReader->name) {
                    case 't':
                        if ($this->xmlReader->nodeType != XMLReader::END_ELEMENT) {
                            $value .= $this->xmlReader->readString();
                        }
                        break;
                    case 'si':
                        if ($this->xmlReader->nodeType == XMLReader::END_ELEMENT) {
                            $this->SSOpen = false;
                            $this->SSForwarded = true;
                            break 2;
                        }
                        break;
                }
            }
        }

        if ($value) {
            $this->lastSharedStringValue = $value;
        }
        return $value;
    }
}
