package com.tranquyet.util.excel;

import com.tranquyet.constant.excel.TableExcelValueConstant;
import org.apache.poi.ss.usermodel.*;

/**
 * ComponentExcel uses to create style for cell and main border of table and graph in the List Student Sheet and the Chart Sheet
 * use Singleton Pattern to get instance
 */
public class ComponentExcel {

    private static ComponentExcel componentExcelInstance = null;

    private ComponentExcel() {

    }

    /**
     * @return the instance of class
     */
    public static ComponentExcel getInstance() {
        if (componentExcelInstance == null) {
            componentExcelInstance = new ComponentExcel();
        }
        return componentExcelInstance;
    }




    /**
     * create style header
     * @param cell current cell of Excel File
     * @param wb representation of an Excel workbook
     */
    public void createStyleHeader(Cell cell, Workbook wb) {
        Font font = wb.createFont();
        font.setFontHeightInPoints(TableExcelValueConstant.SIZE_HEADER_TABLE_EXCEL);
        font.setFontName(TableExcelValueConstant.FONT_STYLE_EXCEL);
        font.setBold(TableExcelValueConstant.BOLD_HEADER_EXCEL);
        CellStyle style = wb.createCellStyle();
        style.setFont(font);
        style.setAlignment(HorizontalAlignment.CENTER);
        style.setVerticalAlignment(VerticalAlignment.CENTER);
        cell.setCellStyle(style);
    }

    /**
     * create style header
     * @param cell current cell of Excel File
     * @param wb representation of an Excel workbook
     */
    public void createStylePage(Cell cell, Workbook wb) {
        Font font = wb.createFont();
        font.setFontHeightInPoints(TableExcelValueConstant.SIZE_WORD_EXCEL);
        font.setFontName(TableExcelValueConstant.FONT_STYLE_EXCEL);
        font.setBold(TableExcelValueConstant.BOLD_WORD_EXCEL);
        CellStyle style = wb.createCellStyle();
        style.setFont(font);
        style.setAlignment(HorizontalAlignment.LEFT);
        style.setVerticalAlignment(VerticalAlignment.CENTER);
        cell.setCellStyle(style);
    }
    /**
     * create style name field
     * @param cell current cell of Excel File
     * @param wb representation of an Excel workbook
     */
    public void createStyleNameField(Cell cell, Workbook wb) {
        Font font = wb.createFont();
        font.setFontHeightInPoints(TableExcelValueConstant.SIZE_WORD_EXCEL);
        font.setFontName(TableExcelValueConstant.FONT_STYLE_EXCEL);
        font.setBold(TableExcelValueConstant.BOLD_HEADER_EXCEL);
        CellStyle style = wb.createCellStyle();
        style.setFont(font);
        style.setAlignment(HorizontalAlignment.LEFT);
        style.setBorderTop(BorderStyle.THIN);
        style.setTopBorderColor(IndexedColors.BLACK.getIndex());
        style.setVerticalAlignment(VerticalAlignment.CENTER);
        cell.setCellStyle(style);
    }
    /**
     * create style name column
     * @param cell current cell of Excel File
     * @param wb representation of an Excel workbook
     */
    public void createStyleNameColumn(Cell cell, Workbook wb) {
        Font font = wb.createFont();
        font.setFontHeightInPoints(TableExcelValueConstant.SIZE_WORD_EXCEL);
        font.setFontName(TableExcelValueConstant.FONT_STYLE_EXCEL);
        font.setBold(TableExcelValueConstant.BOLD_WORD_EXCEL);
        CellStyle style = wb.createCellStyle();
        style.setFont(font);
        style.setAlignment(HorizontalAlignment.LEFT);
        style.setBorderTop(BorderStyle.THIN);
        style.setTopBorderColor(IndexedColors.BLACK.getIndex());
        style.setBorderBottom(BorderStyle.THIN);
        style.setBottomBorderColor(IndexedColors.BLACK.getIndex());
        style.setVerticalAlignment(VerticalAlignment.CENTER);
        cell.setCellStyle(style);
    }


    /**
     * create style for the field row of table
     * @param cell current cell of Excel File
     * @param wb representation of an Excel workbook
     */
    public void createStyleFieldTable(Cell cell, Workbook wb) {

        Font font = wb.createFont();
        font.setFontHeightInPoints(TableExcelValueConstant.SIZE_WORD_EXCEL);
        font.setFontName(TableExcelValueConstant.FONT_STYLE_EXCEL);
        font.setBold(TableExcelValueConstant.BOLD_HEADER_EXCEL);

        CellStyle style = wb.createCellStyle();
        style.setFont(font);
        style.setBorderBottom(BorderStyle.THIN);
        style.setBottomBorderColor(IndexedColors.BLACK.getIndex());
        style.setBorderLeft(BorderStyle.THIN);
        style.setLeftBorderColor(IndexedColors.BLACK.getIndex());
        style.setBorderRight(BorderStyle.THIN);
        style.setRightBorderColor(IndexedColors.BLACK.getIndex());
        style.setBorderTop(BorderStyle.THIN);
        style.setTopBorderColor(IndexedColors.BLACK.getIndex());
        style.setAlignment(HorizontalAlignment.CENTER);
        style.setVerticalAlignment(VerticalAlignment.CENTER);
        cell.setCellStyle(style);
    }
    /**
     * create style for cells of table
     * @param cell current cell of Excel File
     * @param wb representation of an Excel workbook
     */
    public void createStyleCellTable(Cell cell, Workbook wb) {

        Font font = wb.createFont();
        font.setFontHeightInPoints(TableExcelValueConstant.SIZE_WORD_EXCEL);
        font.setFontName(TableExcelValueConstant.FONT_STYLE_EXCEL);
        font.setBold(TableExcelValueConstant.BOLD_WORD_EXCEL);

        CellStyle style = wb.createCellStyle();
        style.setFont(font);
        style.setBorderBottom(BorderStyle.THIN);
        style.setBottomBorderColor(IndexedColors.BLACK.getIndex());
        style.setBorderLeft(BorderStyle.THIN);
        style.setLeftBorderColor(IndexedColors.BLACK.getIndex());
        style.setBorderRight(BorderStyle.THIN);
        style.setRightBorderColor(IndexedColors.BLACK.getIndex());
        style.setBorderTop(BorderStyle.THIN);
        style.setTopBorderColor(IndexedColors.BLACK.getIndex());
        style.setAlignment(HorizontalAlignment.CENTER);
        style.setVerticalAlignment(VerticalAlignment.CENTER);
        cell.setCellStyle(style);
    }

    public void createDateType(Cell cell, Workbook wb){

        Font font = wb.createFont();
        font.setFontHeightInPoints(TableExcelValueConstant.SIZE_WORD_EXCEL);
        font.setFontName(TableExcelValueConstant.FONT_STYLE_EXCEL);
        font.setBold(TableExcelValueConstant.BOLD_WORD_EXCEL);

        CellStyle style = wb.createCellStyle();
        style.setFont(font);
        CellStyle cellStyle = wb.createCellStyle();
        style.setAlignment(HorizontalAlignment.CENTER);
        style.setVerticalAlignment(VerticalAlignment.CENTER);
        cell.setCellStyle(style);
    }
}
