package com.mobin.CSVToExcel;


import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.*;
import java.util.HashMap;
import java.util.Map;

/**
 * Created by MOBIN on 2016/7/29.
 */
public class CSVToExcel {
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

    public void convert(String originPath, String targetPath) {
        try (InputStream input = new FileInputStream(originPath);//源转换文件如:f:\\test
             Reader reader = new InputStreamReader(input, "UTF-8");
             BufferedReader bufferedReader = new BufferedReader(reader);
             FileOutputStream out = new FileOutputStream(targetPath)) {  //如：:f:\\test.xls

            Workbook wb = new XSSFWorkbook();  //Excel
            Sheet sheet = wb.createSheet("Sheet");//如：:xxxSheet
            sheet.createRow(0);                                //创建表头行
            //插入数据
            for (String data = bufferedReader.readLine(); data != null; data = bufferedReader.readLine()) {
                str = data.split(",");
                 lineRow = mapRow.get(str[0]);   //获取该值所对应的行，如果为空表示是新行
                 lineCol = mapCol.get(str[1]);   //获取表头所对应的列，如果为空表示是新表头字段
                if ((lineRow != null)) {                                 //同一行
                    curr_row = sheet.getRow(lineRow);
                    isaddHeader(sheet, lineCol,str[1]);
                    cellnum = mapCol.get(str[1]);
                    curr_row.createCell(cellnum).setCellValue(Double.valueOf(str[2]));
                } else {                                                           //不同行
                    ++lineNum;
                    mapRow.put(str[0], lineNum);
                    row = sheet.createRow(lineNum);
                    row.createCell(0).setCellValue(str[0]);  //第一列值
                    isaddHeader(sheet, lineCol,str[1]);  //添加表头
                    cellnum = mapCol.get(str[1]);
                    row.createCell(cellnum).setCellValue(Double.valueOf(str[2]));
                }
            }
            sheet.getRow(0).createCell(0).setCellValue("name");
            wb.write(out);
            wb.close();
        } catch (FileNotFoundException e) {
            e.printStackTrace();
        } catch (IOException e) {
            e.printStackTrace();
        }
        System.out.println("转换完成！");
    }

    public void isaddHeader(Sheet sheet, Integer lineCol, String value_1) {
        if (lineCol == null) {                                                                         //说明该表头字段不存在
            mapCol.put(value_1, ++colNum);
            sheet.getRow(0).createCell(colNum).setCellValue(value_1);  //添加新表头字段
        }
    }

    public static void main(String[] args) {
        if (args.length != 2) {
            System.err.println("请输入2个参数:源文件路径,目标文件路径");
            return;
        } else {
            CSVToExcel csvToExcel = new CSVToExcel();
            csvToExcel.convert(args[0],args[1]);
        }
    }
}

