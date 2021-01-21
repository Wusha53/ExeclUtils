# ExeclUtils
这是一个使用POI来查看Execl的工具类.返回格式为List<String>.有需求可以自己更改.暂时只支持CSV,XLS XLSX 格式读取
使用方法
  1.使用xlsx数据流:ExeclUtils.ReadExcelUtilsXLSX(InputStream),
  2.使用xls数据流:ExeclUtils.ReadExcelUtilsXLS(InputStream),
  3.使用xlsx或者xls文件名字<带后缀>ExeclUtils.ReadExcelUtils(String fileName)
  4.使用CSV数据流或者文件名字ExeclUtils.ReadExcelUtilsCSV(InputStream  <或者> String fileName)
  后续功能可能会持续更新
