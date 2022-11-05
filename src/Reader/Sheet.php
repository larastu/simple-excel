<?php

namespace Larastu\SimpleExcel\Reader;

use Iterator;
use XMLReader;

class Sheet implements Iterator
{
    /**
     * @var XMLReader
     */
    private $xmlReader;
    private $xmlPath;
    private $name;
    /**
     * @var bool
     */
    private $valid = false;

    /**
     * @var array 当前行数据
     */
    private $currentRow;
    /**
     * @var int 当前行号(从1开始)
     */
    private $currentRowNum = 0;
    /**
     * @var bool
     */
    private $rowOpen = false;

    public function __construct($name, $filePath)
    {
        $this->name = $name;
        $this->xmlPath = $filePath;
    }

    public function __destruct()
    {
        if ($this->xmlReader instanceof XMLReader) {
            $this->xmlReader->close();
            unset($this->xmlReader);
        }
    }

    public function getName()
    {
        return $this->name;
    }

    public function current()
    {
        if ($this->currentRowNum == 0) {
            $this->next();
        }
        return $this->currentRow;
    }

    public function next()
    {
        $this->currentRowNum++;

        $this->currentRow = array();

        if (!$this->rowOpen) {
            while ($this->valid = $this->xmlReader->read()) {
                if ($this->xmlReader->name == 'row') {
                    // Getting the row spanning area (stored as e.g., 1:12)
                    // so that the last cells will be present, even if empty
                    $RowSpans = $this->xmlReader->getAttribute('spans');
                    if ($RowSpans) {
                        $RowSpans = explode(':', $RowSpans);
                        $currentRowColumnCount = $RowSpans[1];
                    } else {
                        $currentRowColumnCount = 0;
                    }

                    if ($currentRowColumnCount > 0) {
                        $this->currentRow = array_fill(0, $currentRowColumnCount, '');
                    }

                    $this->rowOpen = true;
                    break;
                }
            }
        }

        // Reading the necessary row, if found
        if ($this->rowOpen) {
            // These two are needed to control for empty cells
            $MaxIndex = 0;
            $CellCount = 0;

            $cellHasSharedString = false;

            $index = 0;
            $styleId = 0;
            while ($this->valid = $this->xmlReader->read()) {
                switch ($this->xmlReader->name) {
                    // End of row
                    case 'row':
                        if ($this->xmlReader->nodeType == XMLReader::END_ELEMENT) {
                            $this->rowOpen = false;
                            break 2;
                        }
                        break;
                    // Cell
                    case 'c':
                        // If it is a closing tag, skip it
                        if ($this->xmlReader->nodeType == XMLReader::END_ELEMENT) {
                            continue;
                        }

                        $styleId = (int)$this->xmlReader->getAttribute('s');

                        // Get the index of the cell
                        $index = $this->xmlReader->getAttribute('r');
                        $letter = preg_replace('{[^[:alpha:]]}S', '', $index);
                        $index = self::indexFromColumnLetter($letter);

                        // Determine cell type
                        if ($this->xmlReader->getAttribute('t') == 's') {
                            $cellHasSharedString = true;
                        } else {
                            $cellHasSharedString = false;
                        }

                        $this->currentRow[$index] = '';

                        $CellCount++;
                        if ($index > $MaxIndex) {
                            $MaxIndex = $index;
                        }

                        break;
                    // Cell value
                    case 'v':
                    case 'is':
                        if ($this->xmlReader->nodeType == XMLReader::END_ELEMENT) {
                            continue;
                        }

                        $value = $this->xmlReader->readString();

                        if ($cellHasSharedString) {
                            $value = SharedString::getInstance()->getSharedString($value);
                        }

                        // Format value if necessary
                        if ($value !== '' && $styleId && isset($this->Styles[$styleId])) {
                            $value = Styles::getInstance()->formatValue($value, $styleId);
                        } elseif ($value) {
                            $value = Styles::getInstance()->generalFormat($value);
                        }

                        $this->currentRow[$index] = $value;
                        break;
                }
            }

            // Adding empty cells, if necessary
            // Only empty cells inbetween and on the left side are added
            if ($MaxIndex + 1 > $CellCount) {
                $this->currentRow = $this->currentRow + array_fill(0, $MaxIndex + 1, '');
                ksort($this->currentRow);
            }
        }

        return $this->currentRow;
    }

    /**
     * Takes the column letter and converts it to a numerical index (0-based)
     *
     * @param string Letter(s) to convert
     *
     * @return mixed Numeric index (0-based) or boolean false if it cannot be calculated
     */
    public static function indexFromColumnLetter($Letter)
    {
        $Letter = strtoupper($Letter);

        $Result = 0;
        for ($i = strlen($Letter) - 1, $j = 0; $i >= 0; $i--, $j++) {
            $Ord = ord($Letter[$i]) - 64;
            if ($Ord > 26) {
                // Something is very, very wrong
                return false;
            }
            $Result += $Ord * pow(26, $j);
        }
        return $Result - 1;
    }

    public function key()
    {
        return $this->currentRowNum;
    }

    public function valid()
    {
        return $this->valid;
    }

    public function rewind()
    {
        if ($this->xmlReader instanceof XMLReader) {
            $this->xmlReader->close();
        } else {
            $this->xmlReader = new XMLReader;
        }

        $this->xmlReader->open($this->xmlPath);

        $this->valid = true;
        $this->rowOpen = false;
        $this->currentRow = false;
        $this->currentRowNum = 0;
    }

}
