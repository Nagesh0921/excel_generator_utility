package com.nm.excel.utiity;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import java.util.ArrayList;
import java.util.Arrays;

public class ExcelGenerator {
    private static final Logger LOG = LoggerFactory.getLogger(ExcelGenerator.class);

    public XSSFWorkbook generate(){
        LOG.info("ExcelGenerator generate :: invoked");
        Excel excel = Excel.ExcelBuilder.newInstance()
                .setExcelName("Test.xlsx")
                .setNoOfSheets(1)
                .setNameOfSheets(new ArrayList<>(Arrays.asList("Overview", "Severity")))
                .setWorkbook()
                .build();
        XSSFWorkbook wb = excel.getWorkbook();
        LOG.info("Excel Sheet Name {}",wb.getSheetName(0));
        return wb;
    }
}
