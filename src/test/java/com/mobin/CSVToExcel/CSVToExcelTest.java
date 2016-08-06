package com.mobin.CSVToExcel;

import org.junit.Test;

/**
 * Created by MOBIN on 2016/8/5.
 */
public class CSVToExcelTest {
    @Test
    public void csvToExcel(){
        CSVToExcel csvToExcel = new CSVToExcel();
        csvToExcel.convert("f:\\test1","f:\\p.xls");
    }
}
