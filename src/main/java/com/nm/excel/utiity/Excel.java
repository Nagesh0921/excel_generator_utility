package com.nm.excel.utiity;

import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.util.List;

/**
 * @author nagesh
 */
public final class Excel {

    private final String excelName;
    //TODO: Improvement is to remove no of sheets.
    private final Integer noOfSheets;
    private final List<String> nameofSheets;
    private final XSSFWorkbook wb;

    public Excel(ExcelBuilder builder){
        excelName = builder.excelName;
        noOfSheets = builder.noOfSheets;
        nameofSheets = builder.nameOfSheets;
        wb = builder.wb;
    }

    public XSSFWorkbook getWorkbook(){
        return this.wb;
    }

    public static class ExcelBuilder{
        private String excelName;
        private Integer noOfSheets;
        private List<String> nameOfSheets;
        private XSSFWorkbook wb;

        public static ExcelBuilder newInstance(){
            return new ExcelBuilder();
        }

        private ExcelBuilder(){

        }

        public ExcelBuilder setExcelName(String excelName){
            this.excelName = excelName;
            return this;
        }
        public ExcelBuilder setNoOfSheets(Integer noOfSheets){
            this.noOfSheets = noOfSheets;
            return this;
        }
        public ExcelBuilder setNameOfSheets(List<String> nameOfSheets){
            this.nameOfSheets = nameOfSheets;
            return this;
        }
        public ExcelBuilder setWorkbook(){
            this.wb = new XSSFWorkbook();
            for(String name : nameOfSheets){
                Sheet sheet = this.wb.createSheet(name);
            }
            return this;
        }

        public Excel build(){
            return new Excel(this);
        }
    }
}
