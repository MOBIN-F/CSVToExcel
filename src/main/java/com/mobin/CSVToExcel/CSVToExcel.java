package com.mobin.CSVToExcel;


import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;

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
    private String[] value_1;

    public void convert(String originPath, String targetPath) {
        try (InputStream input = new FileInputStream(originPath);//源转换文件如:f:\\test
             Reader reader = new InputStreamReader(input, "UTF-8");
             BufferedReader bufferedReader = new BufferedReader(reader);
             FileOutputStream out = new FileOutputStream(targetPath)) {  //如：:f:\\test.xls

            Workbook wb = new HSSFWorkbook();  //Excel
            Sheet sheet = wb.createSheet("Sheet");//如：:xxxSheet
            sheet.createRow(0);                                //创建表头行
            //插入数据
            for (String data = bufferedReader.readLine(); data != null; data = bufferedReader.readLine()) {
                str = data.split(",");
                value_1 = str[1].split("~");
                int cellnum_1 = Integer.valueOf(value_1[1]);//截取拼接在后面的数字，该数字表示值和表头放在哪列

                Integer lineRow = mapRow.get(str[0]);   //获取该值所对应的行，如果为空表示是新行
                Integer lineCol = mapCol.get(value_1[0]);   //获取表头所对应的列，如果为空表示是新表头字段
                if ((lineRow != null)) {//同一行
                    Row curr_row = sheet.getRow(lineRow);
                    isaddHeader(sheet, lineCol, cellnum_1, value_1[0]);
                    curr_row.createCell(cellnum_1).setCellValue(str[2]);
                } else {                                                           //不同行
                    ++lineNum;
                    mapRow.put(str[0], lineNum);
                    Row row = sheet.createRow(lineNum);
                    row.createCell(0).setCellValue(str[0]);  //第一列值
                    isaddHeader(sheet, lineCol, cellnum_1, value_1[0]);
                    row.createCell(cellnum_1).setCellValue(str[2]);
                }
            }
            wb.write(out);
        } catch (FileNotFoundException e) {
            e.printStackTrace();
        } catch (IOException e) {
            e.printStackTrace();
        }
        System.out.println("转换完成！");
    }

    public void isaddHeader(Sheet sheet, Integer lineCol, int cellnum_1, String value_1) {
        if (lineCol == null) {                                                                         //说明该表头字段不存在
            mapCol.put(value_1, ++colNum);
            sheet.getRow(0).createCell(cellnum_1).setCellValue(value_1);  //添加新表头字段
        }
    }

    public static void main(String[] args) {
        if (args.length != 3) {
            System.err.println("请输入3个参数:源文件路径,目标文件路径");
            return;
        } else {
            CSVToExcel csvToExcel = new CSVToExcel();
            csvToExcel.convert(args[0], args[1]);
        }
    }
}

