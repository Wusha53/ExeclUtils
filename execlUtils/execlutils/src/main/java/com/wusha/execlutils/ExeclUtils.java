package com.wusha.execlutils;

import org.apache.poi.hssf.usermodel.HSSFDateUtil;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellValue;
import org.apache.poi.ss.usermodel.FormulaEvaluator;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.BufferedReader;
import java.io.FileInputStream;
import java.io.FileReader;
import java.io.InputStream;
import java.io.InputStreamReader;
import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.List;

public class ExeclUtils {

    /**
     * 创建一个csv文件的工具类
     * @param inputStream 数据流
     * @throws Exception
     */
    public static List<String> ReadExcelUtilsCSV(InputStream inputStream)throws Exception{
        BufferedReader br = new BufferedReader(new InputStreamReader(inputStream));
        List<String> list = new ArrayList<String>();
        String stemp;
        while ((stemp = br.readLine()) != null) {
            list.add(stemp);
        }
        inputStream.close();
        br.close();
        return  list;
    } /**
     * 创建一个csv文件的工具类
     * @param filePath 文件路径
     * @throws Exception
     */
    public static List<String> ReadExcelUtilsCSV(String filePath)throws Exception{
        BufferedReader br = new BufferedReader(new FileReader(filePath));
        List<String> list = new ArrayList<String>();
        String stemp;
        while ((stemp = br.readLine()) != null) {
            list.add(stemp);
        }
        br.close();
        return  list;
    }

    /**
     * XLS数据流格式
     * @param inputStream
     * @return
     * @throws Exception
     */
    public static String ReadExcelUtilsXLS (InputStream inputStream)throws Exception {
        HSSFWorkbook hssfWorkbook = new HSSFWorkbook(inputStream);
        HSSFSheet hssfSheet = hssfWorkbook.getSheetAt(0);
//            XSSFWorkbook workbook = new XSSFWorkbook(stream);
//            XSSFSheet sheet = workbook.getSheetAt(0);
        int rowsCount = hssfSheet.getPhysicalNumberOfRows();
        FormulaEvaluator formulaEvaluator = hssfWorkbook.getCreationHelper().createFormulaEvaluator();
        String cellInfo="";
        for (int r = 0; r < rowsCount; r++) {
            Row row = hssfSheet.getRow(r);
            int cellsCount = row.getPhysicalNumberOfCells();
            for (int c = 0; c < cellsCount; c++) {
                String value = getCellAsString(row, c, formulaEvaluator);
                if (cellInfo.equals("")){
                    cellInfo=value;
                }else {
                    cellInfo = cellInfo +","+ value;
                }

            }
        }
        return cellInfo;
    }
    /**
     * XLS文件名格式
     * @param fileName 文件名
     * @return
     * @throws Exception
     */
    public static String ReadExcelUtils (String fileName)throws Exception {
        InputStream inputStream = new FileInputStream(fileName);
        Workbook workbook=null;
        if (fileName.endsWith(".xls")) {
            workbook = new HSSFWorkbook(inputStream);
        } else if (fileName.endsWith(".xlsx")) {
            workbook = new XSSFWorkbook(inputStream);
        }
        Sheet sheet = workbook.getSheetAt(0);
        //            XSSFWorkbook workbook = new XSSFWorkbook(stream);
//            XSSFSheet sheet = workbook.getSheetAt(0);
        int rowsCount = sheet.getPhysicalNumberOfRows();
        FormulaEvaluator formulaEvaluator = workbook.getCreationHelper().createFormulaEvaluator();
        String cellInfo="";
        for (int r = 0; r < rowsCount; r++) {
            Row row = sheet.getRow(r);
            int cellsCount = row.getPhysicalNumberOfCells();
            for (int c = 0; c < cellsCount; c++) {
                String value = getCellAsString(row, c, formulaEvaluator);
                if (cellInfo.equals("")){
                    cellInfo=value;
                }else {
                    cellInfo = cellInfo +","+ value;
                }

            }
        }
        return cellInfo;
    }


    /**
     * XLS数据流格式
     * @param inputStream
     * @return
     * @throws Exception
     */
    public static String ReadExcelUtilsXLSX (InputStream inputStream)throws Exception {
        XSSFWorkbook workbook = new XSSFWorkbook(inputStream);
        XSSFSheet sheet = workbook.getSheetAt(0);
        int rowsCount = sheet.getPhysicalNumberOfRows();
        FormulaEvaluator formulaEvaluator = workbook.getCreationHelper().createFormulaEvaluator();
        String cellInfo="";
        for (int r = 0; r < rowsCount; r++) {
            Row row = sheet.getRow(r);
            int cellsCount = row.getPhysicalNumberOfCells();
            for (int c = 0; c < cellsCount; c++) {
                String value = getCellAsString(row, c, formulaEvaluator);
                if (cellInfo.equals("")){
                    cellInfo=value;
                }else {
                    cellInfo = cellInfo +","+ value;
                }


            }
        }
        return cellInfo;
    }
    protected static String getCellAsString(Row row, int c, FormulaEvaluator formulaEvaluator) {
        String value = "";
        try {
            Cell cell = row.getCell(c);
            CellValue cellValue = formulaEvaluator.evaluate(cell);
            switch (cellValue.getCellType()) {
                case Cell.CELL_TYPE_BOOLEAN:
                    value = ""+cellValue.getBooleanValue();
                    break;
                case Cell.CELL_TYPE_NUMERIC:
                    double numericValue = cellValue.getNumberValue();
                    if(HSSFDateUtil.isCellDateFormatted(cell)) {
                        double date = cellValue.getNumberValue();
                        SimpleDateFormat formatter =
                                new SimpleDateFormat("dd/MM/yy");
                        value = formatter.format(HSSFDateUtil.getJavaDate(date));
                    } else {
                        value = ""+numericValue;
                    }
                    break;
                case Cell.CELL_TYPE_STRING:
                    value = ""+cellValue.getStringValue();
                    break;
                default:
            }
        } catch (NullPointerException e) {
            /* proper error handling should be here */

        }
        return value;
    }
}
