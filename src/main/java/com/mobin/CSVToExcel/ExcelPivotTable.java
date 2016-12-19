package com.mobin.CSVToExcel;

import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.poifs.filesystem.POIFSFileSystem;
import org.apache.poi.ss.SpreadsheetVersion;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.AreaReference;
import org.apache.poi.ss.util.CellReference;
import org.apache.poi.xssf.usermodel.XSSFPivotTable;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.*;
import java.util.HashMap;
import java.util.Map;

/**
 * Created by Mobin on 2016/12/17.
 * 透视表
 * 源文件必须第一是表名，第二列是日期 ，第三列是值
 */
public class ExcelPivotTable {
    private String[] line;
    private static int columnLableIndex = 1;
    private int lineNum = 0;
    private int colNum = 0;
    private Map<String, Integer> mapRow = new HashMap<>();   //值所在行
    private Map<String, Integer> mapCol = new HashMap<>();   //表头所在列
    private String[] str;
    private Integer lineRow;
    private Integer lineCol;
    private int cellnum;
    private Row row;
    private Row curr_row;

    public void pivotTable(String originPath) {
        try (XSSFWorkbook wb = new XSSFWorkbook(new FileInputStream("/home/noce1/mobin/watchData.xlsx"));
             FileOutputStream out = new FileOutputStream("/home/noce1/mobin/watchData.xlsx")) {

            String startTime_year = originPath.substring(originPath.indexOf('.') - 8, originPath.indexOf('.'));
            XSSFSheet sheet = wb.createSheet(startTime_year);
            setCellData(sheet, originPath);//创建透视表所需的数据

            String reference = "A1:" + String.valueOf((char) (sheet.getRow(0).getLastCellNum() + 64))
                    + (sheet.getLastRowNum() + 1);
            AreaReference source = new AreaReference(reference, SpreadsheetVersion.EXCEL2007);
            CellReference position = new CellReference("A42");  //相对位置
            XSSFPivotTable pivotTable = sheet.createPivotTable(source, position);
            pivotTable.addRowLabel(0);  //使用一列作为行标签
            Integer startTime = Integer.valueOf(startTime_year + "00");
            Integer endTime = startTime + (sheet.getRow(0).getLastCellNum() - 2);

            for (int columnLable = startTime; columnLable <= endTime; columnLable++) {  //列标签
                pivotTable.addColumnLabel(DataConsolidateFunction.SUM, columnLableIndex++, String.valueOf(columnLable));
            }
            wb.write(out);
            wb.close();
        } catch (FileNotFoundException e) {
            e.printStackTrace();
        } catch (IOException e) {
            e.printStackTrace();
        }

    }

    public void setCellData(XSSFSheet sheet, String origin) {
        try (InputStream input = new FileInputStream(origin);//源转换文件如:f:\\test
             Reader reader = new InputStreamReader(input, "UTF-8");
             BufferedReader bufferedReader = new BufferedReader(reader)) {  //如：:f:\\test.xls
            sheet.createRow(0);                                //创建表头行
            //插入数据
            for (String data = bufferedReader.readLine(); data != null; data = bufferedReader.readLine()) {
                str = data.split(",");
                lineRow = mapRow.get(str[0]);   //获取该值所对应的行，如果为空表示是新行
                lineCol = mapCol.get(str[1]);   //获取表头所对应的列，如果为空表示是新表头字段
                if ((lineRow != null)) {                                 //同一行
                    curr_row = sheet.getRow(lineRow);
                    isaddHeader(sheet, lineCol, str[1]);
                    cellnum = mapCol.get(str[1]);
                    curr_row.createCell(cellnum).setCellValue(Double.valueOf(str[2]));  //将值转为数值类型，否则出错
                } else {                                                           //不同行
                    ++lineNum;
                    mapRow.put(str[0], lineNum);
                    row = sheet.createRow(lineNum);
                    row.createCell(0).setCellValue(str[0]);  //第一列值
                    isaddHeader(sheet, lineCol, str[1]);  //添加表头
                    cellnum = mapCol.get(str[1]);
                    row.createCell(cellnum).setCellValue(Double.valueOf(str[2]));
                }
            }
            sheet.getRow(0).createCell(0).setCellValue("name");
        } catch (FileNotFoundException e) {
            e.printStackTrace();
        } catch (IOException e) {
            e.printStackTrace();
        }
    }

    public void isaddHeader(Sheet sheet, Integer lineCol, String value_1) {
        if (lineCol == null) {                                                                         //说明该表头字段不存在
            mapCol.put(value_1, ++colNum);
            sheet.getRow(0).createCell(colNum).setCellValue(value_1);  //添加新表头字段
        }
    }

    public static void main(String[] args) {
        if (args.length != 1) {
            System.err.println("只需输入1个参数:源文件路径");
            return;
        } else {
            ExcelPivotTable excelPivotTable = new ExcelPivotTable();
            excelPivotTable.pivotTable(args[0]);
        }

    }
}
