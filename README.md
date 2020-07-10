# InputOutputExcel v1.0.0
PHP写的 excel csv大文件导出导入类。
对于网站后台系统常用的大文件导入导出功能，常用的有PHPExcel,PHP_XLSXWriter等，但这些对导出的大数据量支持并不友好，所以自己封装了一个基于csv格式类，引入了yield生成器函数，实现浏览器边写入边下载，所以高性能，支持百万级数据导入导出，仅需5秒左右。

