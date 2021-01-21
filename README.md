# ExeclUtils
这是一个使用POI来查看Execl的工具类.返回格式为List<String>.有需求可以自己更改.暂时只支持CSV,XLS XLSX 格式读取
使用方法
	
	
  1.使用xlsx数据流:ExeclUtils.ReadExcelUtilsXLSX(InputStream), 
  
  2.使用xls数据流:ExeclUtils.ReadExcelUtilsXLS(InputStream),
  
  3.使用xlsx或者xls文件名字<带后缀>ExeclUtils.ReadExcelUtils(String fileName)
  
  4.使用CSV数据流或者文件名字ExeclUtils.ReadExcelUtilsCSV(InputStream  <或者> String fileName)
   
 * @param fileName 文件地址     * @param startRow 开始的行数   * @param endRow   结束的行数
    
  5.新增删除功能removeColumn(String fileName, int startRow, int endRow)
 
 后续功能可能会持续更新


Add it in your root build.gradle at the end of repositories:

	allprojects {
		repositories {
			...
			maven { url 'https://jitpack.io' }
		}
	}
Step 2. Add the dependency

	dependencies {
	        implementation 'com.github.qq331617870:ExeclUtils:1.3'
	}
