<?php
/**
 * excel大文件导出导入接口
 */
namespace App\Libs\InputOutputExcel;

interface InputOutputExcelContract
{
    public function export($excelData, $csvFileName,$head);
    public function import($sFilePath, $line, $offset);
}
