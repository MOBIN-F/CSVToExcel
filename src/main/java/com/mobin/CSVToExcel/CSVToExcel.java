package com.mobin.CSVToExcel;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;

import java.io.*;

/**
 * Created by MOBIN on 2016/7/29.
 */
public class CSVToExcel {

    private String tmp = null;
    private int lineNum = -1;
    private int headerCellNum = 0;
    public void convert(String originPath,String targetPath,String headertype){
        try(InputStream input = new FileInputStream(originPath);//源转换文件如:f:\\test
            Reader reader = new InputStreamReader(input,"UTF-8");
            BufferedReader bufferedReader = new BufferedReader(reader);
            FileOutputStream out = new FileOutputStream(targetPath);

            InputStream headerInput = getClass().getResourceAsStream("/"+headertype.toUpperCase()+".txt");//获取表头文件
            Reader reader1 = new InputStreamReader(headerInput,"UTF-8");
            BufferedReader bufferedReader1 = new BufferedReader(reader1)){  //如：:f:\\test.xls

            Workbook wb = new HSSFWorkbook();  //Excel
            Sheet sheet = wb.createSheet("Sheet");//如：:xxxSheet
            //创建Excel表头
            Row headerRow = sheet.createRow(++lineNum);
            for(String header = bufferedReader1.readLine(); header != null; header = bufferedReader1.readLine()){
                headerRow.createCell(++headerCellNum).setCellValue(header);
            }
            //插入数据
            for(String data = bufferedReader.readLine(); data != null; data = bufferedReader.readLine()){
                String[] str = data.split(",");
                int cellnum = Integer.valueOf(str[1].split("~")[1]);
                if(str[0].equals(tmp)){//同一行
                    Row curr_row= sheet.getRow(lineNum);
                    curr_row.createCell(cellnum).setCellValue(str[2]);
                }else {//不同值需
                    Row row = sheet.createRow(++lineNum);
                    row.createCell(0).setCellValue(str[0]);
                    row.createCell(cellnum).setCellValue(str[2]);
                }
                tmp = str[0];
            }
            wb.write(out);
        } catch (FileNotFoundException e) {
            e.printStackTrace();
        } catch (IOException e) {
            e.printStackTrace();
        }
        System.out.println("转换完成！");
    }

    public static void main(String[] args) {
        if(args.length != 3){
            System.err.println("请输入3个参数:源文件路径,目标文件路径,Excel表头类型");
            return;
        }else if(!(args[2].toUpperCase().equals("CD")||args[2].toUpperCase().equals("S8b")||args[2].toUpperCase().equals("CX")
                 ||args[2].toUpperCase().equals("S3A")||args[2].toUpperCase().equals("GJ")||args[2].toUpperCase().equals("JK")
                 ||args[2].toUpperCase().equals("S8A")||args[2].toUpperCase().equals("RY") ||args[2].toUpperCase().equals("S2")
                 ||args[2].toUpperCase().equals("S7"))){
            System.err.println("第三个输入参数有误!");
            return;
        }else {
            CSVToExcel csvToExcel = new CSVToExcel();
            csvToExcel.convert(args[0], args[1], args[2]);
        }
    }
}

