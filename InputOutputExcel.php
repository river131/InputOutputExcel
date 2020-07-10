<?php

/**
 * excel大文件导出导入
 * User: dajiang
 * Date: 2020/6/17
 * Time: 11:07
 */

namespace App\Libs\InputOutputExcel;

class InputOutputExcel implements InputOutputExcelContract
{
    const EXCEL_2007_MAX_ROW = 1048576;
    const EXCEL_2007_MAX_COL = 16384;
    private $exportFileName;//导出的文件名
//    public $excelData;//要导出的数据数组
//    public $exportPath;//导出文件路径

    public function __construct()
    {
        //设置默认导出文件名
        $this->exportFileName = date('Y-m-d') . uniqid();
    }

    //得到导出的文件名
    public function getExportFileName()
    {
        return $this->exportFileName;
    }

    /**
     * 设置导出路径
     * @return string
     */
    public function setExportPath()
    {
        $dir = storage_path('exports');
        if (!file_exists($dir)) {
            mkdir($dir, 0777, true);
        }
        return $dir;
    }

    /**
     * 使用生成器返回数据
     * @param $excelData
     * @return \Generator
     */
    public function getCsv($excelData)
    {
        foreach ($excelData as $key => $val) {
            yield $val;
        }
    }

    /**
     * 导出excel csv格式
     * @param $excelData 接受二维数组
     * @param string $fileName 导出文件名
     * @param string $head 导出的csv标头一维数组
     * @return string
     */
    public function export($excelData, $csvFileName = '',$head=[])
    {
        if (empty($excelData) || !is_array($excelData)) {
            throw new \Exception("要导出的数据参数不存在或格式有误");
        }
        if (count($excelData) > self::EXCEL_2007_MAX_ROW) {
            throw new \Exception("导出数据行数不能超过" . self::EXCEL_2007_MAX_ROW . '行');
        }
        if (empty($csvFileName)) {
            $csvFileName = $this->getExportFileName();
        }
        $csvFileName = $csvFileName . '.csv';
//print_r($excelData);die;
        if(empty($head)){
            $columns = array_shift($excelData);
        }else{
            $columns=$head;
        }
//        print_r($columns);
        $delimiter = ',';
        $enclosure = '"';
        //设置好告诉浏览器要下载excel文件的headers
        header('Content-Description: File Transfer');
        header('Content-Type: application/vnd.ms-excel');
        header('Content-Disposition: attachment; filename="' . $csvFileName . '"');
        header('Expires: 0');
        header('Cache-Control: must-revalidate');
        header('Pragma: public');
        $fp = fopen('php://output', 'a');//打开output流
        mb_convert_variables('GBK', 'UTF-8', $columns);
        fputcsv($fp, $columns, $delimiter, $enclosure);//将数据格式化为CSV格式并写入到output流中

        $data = $this->getCsv($excelData);//使用生成器
        foreach ($data as $k => $v) {
            mb_convert_variables('GBK', 'UTF-8', $v);
            fputcsv($fp, $v, $delimiter, $enclosure);
            //刷新输出缓冲到浏览器
//            ob_flush();
//            flush();//必须同时使用 ob_flush() 和flush() 函数来刷新输出缓冲。
        }
        fclose($fp) or die("can‘t close php://output");
        exit();

    }

    /**
     * 导入Excel csv数据表格
     * @param string $sFilePath 文件名(在服务器上绝对路径)
     * @param int $line 读取几行，默认全部读取
     * @param int $offset 从第几行开始读，默认从第一行读取
     * @return bool|array
     */
    public function import($sFilePath, $line = 0, $offset = 0)
    {
        $i = 0;
        $j = 0;
        $arr = [];
        try {
            foreach ($this->getRows($sFilePath) as $arow) {
                //跳过行数
                if ($i < $offset && $offset) {
                    $i++;
                    continue;
                }

                //大于读取行数则退出
                if ($i > $line && $line) {
                    break;
                }

                if ($arow) {
                    foreach ($arow as $k => &$value) {
                        $content = iconv("gbk", "utf-8//IGNORE", $value);//转化编码
                        $arr[$j][] = $content;
                    }
                    unset($arow);

                    $i++;
                    $j++;
                }
            }
        } catch (Exception $e) {
            print_r($e->getMessage());
        }
        return $arr;
    }

    /**
     * 用生成器取数据
     * @param $file 文件路径
     * @return \Generator
     * @throws \Exception
     */
    function getRows($file)
    {
        $delimiter = ',';
        $enclosure = '"';
        if (!file_exists($file)) {
            throw new \Exception("要导入文件不存在");
        }
        $aFilePath = explode('.', $file);
        $extension = array_pop($aFilePath);
        if ($extension !== 'csv') {
            throw new \Exception("导入文件只支持csv格式");
        }
        $handle = fopen($file, 'rb');
        if ($handle === false) {
            throw new \Exception("文件打开失败");
        }

        while (feof($handle) === false) {
            yield fgetcsv($handle, '', $delimiter, $enclosure);
        }
        fclose($handle);
//        exit;
    }


}

?>