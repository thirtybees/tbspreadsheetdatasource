<?php

namespace TbSpreadsheetDatasource;

use PhpOffice\PhpSpreadsheet\Reader\Csv;
use PhpOffice\PhpSpreadsheet\Reader\IReader;
use PrestaShopException;
use Thirtybees\Core\Import\DataSourceInterface;

class DataSource implements DataSourceInterface
{

    /**
     * @var IReader
     */
    protected $reader;

    /**
     * @var string
     */
    private $filename;

    /**
     * @var array
     */
    private $worksheetInfo;

    /**
     * @var int
     */
    private $worksheetIndex = 0;

    /**
     * @var array[];
     */
    private $data;

    /**
     * @var int
     */
    private $index = -1;


    /**
     * @param string $filename
     * @throws \PhpOffice\PhpSpreadsheet\Reader\Exception
     */
    public function __construct($filename, $params)
    {
        $this->filename = $filename;
        $this->reader = \PhpOffice\PhpSpreadsheet\IOFactory::createReaderForFile($filename);
        if ($this->reader instanceof Csv) {
            $this->reader->setTestAutoDetect(false);
            if (isset($params['separator'])) {
                $this->reader->setDelimiter($params['separator']);
            }
        }
        $this->reader->setReadDataOnly(true);
        if (isset($params['worksheetIndex'])) {
            $this->worksheetIndex = (int)$params['worksheetIndex'];
        }
    }

    /**
     * @return array
     * @throws \PhpOffice\PhpSpreadsheet\Exception
     */
    public function getRow()
    {
        $data = $this->getData();
        if ($data) {
            $this->index++;
            if (isset($data[$this->index])) {
                return $data[$this->index];
            }
        }
        return [];
    }

    /**
     * @return int
     * @throws PrestaShopException
     */
    public function getNumberOfColumns()
    {
        return (int)$this->getSheetInfo()['totalColumns'];
    }

    /**
     * @return int
     * @throws PrestaShopException
     */
    public function getNumberOfRows()
    {
        return (int)$this->getSheetInfo()['totalRows'];
    }

    /**
     * @return bool
     */
    public function close()
    {
        $this->data = null;
        $this->worksheetInfo = null;
        return true;
    }

    /**
     * @return array
     * @throws \PhpOffice\PhpSpreadsheet\Exception
     */
    protected function getData()
    {
        if (is_null($this->data)) {
            $spreadsheet = $this->reader->load($this->filename);
            $spreadsheet->setActiveSheetIndex($this->worksheetIndex);
            $worksheet = $spreadsheet->getActiveSheet();
            $this->data = $worksheet->toArray();
            $this->index = -1;
        }
        return $this->data;
    }

    /**
     * @return array
     * @throws PrestaShopException
     */
    protected function getSheetInfo()
    {
        if (is_null($this->worksheetInfo)) {
            if (method_exists($this->reader, 'listWorksheetInfo')) {
                $infos = $this->reader->listWorksheetInfo($this->filename);
                if (isset($infos[$this->worksheetIndex])) {
                    $this->worksheetInfo = $infos[$this->worksheetIndex];
                } else {
                    throw new PrestaShopException("File does not contains sheet with index " . $this->worksheetIndex);
                }
            } else {
                throw new PrestaShopException("Reader does not support listWorksheetInfo method");
            }
        }
        return $this->worksheetInfo;
    }
}