<?php
/**
 * Copyright (C) 2017-2024 thirty bees
 *
 * NOTICE OF LICENSE
 *
 * This source file is subject to the Academic Free License (AFL 3.0)
 * that is bundled with this package in the file LICENSE.md.
 * It is also available through the world-wide-web at this URL:
 * https://opensource.org/licenses/afl-3.0.php
 * If you did not receive a copy of the license and are unable to
 * obtain it through the world-wide-web, please send an email
 * to license@thirtybees.com so we can send you a copy immediately.
 *
 * @author    thirty bees <modules@thirtybees.com>
 * @copyright 2017-2024 thirty bees
 * @license   Academic Free License (AFL 3.0)
 */


use PhpOffice\PhpSpreadsheet\Reader\Exception as PhpOfficeException;
use TbSpreadsheetDatasource\DataSource;
use Thirtybees\Core\Import\DataSourceInterface;

if (!defined('_TB_VERSION_')) {
    exit;
}

class TbSpreadsheetDatasource extends Module
{
    /**
     * TbSpreadsheetDatasource constructor.
     *
     * @throws PrestaShopException
     */
    public function __construct()
    {
        $this->name = 'tbspreadsheetdatasource';
        $this->tab = 'administration';
        $this->version = '1.0.0';
        $this->author = 'thirty bees';
        $this->controllers = [];
        $this->bootstrap = true;
        parent::__construct();
        $this->displayName = $this->l('Spreadsheet Datasource');
        $this->description = $this->l('Provides ability to import excel files using thirty bees CSV import.');
        $this->need_instance = 0;
        $this->tb_versions_compliancy = '>= 1.4.0';
        $this->tb_min_version = '1.4.0';
    }

    /**
     * Module installation process
     *
     * @return bool
     * @throws PrestaShopDatabaseException
     * @throws PrestaShopException
     */
    public function install()
    {
        if (version_compare(phpversion(), '7.2', '<')) {
            $this->_errors[] = Tools::displayError('Module '.$this->displayName.' requires PHP version 7.2 and higher');
            return false;
        }

        return (
            parent::install() &&
            $this->registerHook('actionRegisterImportDataSource')
        );
    }

    /**
     * Header hook handler
     */
    public function hookActionRegisterImportDataSource()
    {
        return [
            'name' => $this->displayName,
            'extensions' => ['csv', 'xls', 'xlsx', 'xlst', 'ods', 'ots'],
            'constructor' => [$this, 'createDataSource']
        ];
    }

    /**
     * @return DataSourceInterface
     * @throws PhpOfficeException
     */
    public function createDataSource($filepath, $parameters)
    {
        require_once __DIR__.'/vendor/autoload.php';
        return new DataSource($filepath, $parameters);
    }
}
