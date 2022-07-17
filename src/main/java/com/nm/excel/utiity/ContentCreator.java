package com.nm.excel.utiity;

import org.apache.poi.ss.usermodel.FillPatternType;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.xssf.usermodel.*;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import java.util.Map;

public class ContentCreator {
    private static final Logger LOG = LoggerFactory.getLogger(ContentCreator.class);

    private int rowIndex = 0, colIndex = 0, bufferValue = 0;

    public int[][] createHeaderContent(XSSFSheet sheet, ExcelInput input, int startRowIndex, int startColIndex){
        LOG.info("ContentCreator :: createHeaderContent invoked");
        //Create Row
        rowIndex = startRowIndex;
        colIndex = startColIndex;
        for(Map.Entry<String, String> set : input.getHeadersMap().entrySet()){
            XSSFRow row = sheet.createRow(rowIndex);
            XSSFCell cell = row.createCell(colIndex);
            cell.setCellValue(set.getKey());
            applyHeaderStyle(cell);
            colIndex++;
            XSSFCell valCell = row.createCell(colIndex);
            valCell.setCellValue(set.getValue());
            rowIndex++; colIndex = startColIndex;
        }
        return new int[][]{{rowIndex+1, colIndex}};
    }

    public void createContent(XSSFSheet sheet, ExcelInput input, int startRowIndex, int bufferValue){
        LOG.info("ContentCreator :: createContent invoked.");
        rowIndex = startRowIndex+bufferValue;
        XSSFRow row = sheet.createRow(rowIndex);
        for(String tableHeader : input.getHeaders()){
            XSSFCell cell = row.createCell(colIndex);
            cell.setCellValue(tableHeader);
            applyHeaderStyle(cell);
            colIndex++;
        }
    }

    public void createData(XSSFSheet sheet, ExcelInput input, int startRowIndex){
        LOG.info("ContentCreator :: createData invoked.");
        rowIndex = startRowIndex;
        Object[][] obj = input.getObj();
        for(int r = 0; r < obj.length; r++){
            XSSFRow row = sheet.createRow(rowIndex);
            for(int c = 0; c < obj[0].length; c++){
                XSSFCell cell = row.createCell(c);
                cell.setCellValue((obj[r][c]).toString());
            }
            rowIndex++;
        }
    }

    private void applyHeaderStyle(XSSFCell cell){
        XSSFCellStyle cellStyle = cell.getSheet().getWorkbook().createCellStyle();
        cellStyle.setFillBackgroundColor(IndexedColors.GREY_25_PERCENT.index);
        cellStyle.setFillPattern(FillPatternType.BIG_SPOTS);
        cellStyle.setFillForegroundColor(IndexedColors.GREY_40_PERCENT.getIndex());
        cell.setCellStyle(cellStyle);
    }
}
