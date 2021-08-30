package com.tranquyet.util.excel.table;

import com.tranquyet.constant.excel.TableExcelValueConstant;
import com.tranquyet.service.StudentStorageService;
import com.tranquyet.util.excel.table.chart.ChartTableExcel;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFSheet;

/**
 * ComponentExcel uses to create style for cell, data and graph in the List Student Sheet
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
     *
     * @param cell current cell of Excel File
     * @param wb   representation of an Excel workbook
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
     *
     * @param cell current cell of Excel File
     * @param wb   representation of an Excel workbook
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
     *
     * @param cell current cell of Excel File
     * @param wb   representation of an Excel workbook
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
     *
     * @param cell current cell of Excel File
     * @param wb   representation of an Excel workbook
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
     *
     * @param cell current cell of Excel File
     * @param wb   representation of an Excel workbook
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
     *
     * @param cell current cell of Excel File
     * @param wb   representation of an Excel workbook
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

    /**
     * set break page for student sheet
     * <pre>
     *     1 page have a constant row is 41
     *     then after after create a page the cell will auto increase 41
     *     so, we use for loop to create n page where count to check n*20 % 20 == 0
     *     and check to count the cell dont use to bind value and n time will have n*20
     *
     * </pre>
     *
     * @param studentSheet list student sheet
     */
    protected static void createBreakPageOne(Sheet studentSheet) {
        int count = 0;
        int check = 20;
        ChartTableExcel.createChartTable((XSSFSheet) studentSheet, 0);
        studentSheet.setRowBreak(41);
        //studentSheet.setRowBreak(36);
        for (int i = 0; i < studentSheet.getLastRowNum(); i++) {
            if (count % 20 == 0 && count != 0) {
                studentSheet.setRowBreak(36 + count + check);
                check += 20;
            }
            count++;
        }
        studentSheet.getPrintSetup().setPaperSize(PrintSetup.A4_PAPERSIZE);
        studentSheet.getPrintSetup().setLandscape(false);
        studentSheet.setFitToPage(true);
        PrintSetup printSetup = studentSheet.getPrintSetup();
        printSetup.setFitHeight((short) 0);
        printSetup.setFitWidth((short) 1);
    }

    /**
     * set break page for student sheet
     * <pre>
     *     1 page have a constant row is 41
     *     then after after create a page the cell will auto increase 41
     *     so, we use for loop to create n page where count to check n*20 % 20 == 0
     *     and check to count the cell dont use to bind value and n time will have n*20
     *
     * </pre>
     *
     * @param studentSheet list student sheet
     */
    public static void createBreakPage(Sheet studentSheet, int pageNumber) {
        int count = 0;
        int check = 17;
        if (pageNumber == StudentStorageService.pagination(34).size() - 1) {
            studentSheet.setRowBreak(79);
        } else if (pageNumber == StudentStorageService.pagination(34).size() - 1
                &&
                StudentStorageService.getTotal() <= 20) {
            studentSheet.setRowBreak(40);
        } else {
            studentSheet.setRowBreak(36);
            for (int i = 0; i < studentSheet.getLastRowNum() - 20; i++) {
                if (count % 20 == 0 && count != 0) {
                    studentSheet.setRowBreak(36 + count + check);
                    check += 17;
                }
                count++;
            }
        }
        studentSheet.getPrintSetup().setPaperSize(PrintSetup.A4_PAPERSIZE);
        studentSheet.getPrintSetup().setLandscape(false);
        studentSheet.setFitToPage(true);
        PrintSetup printSetup = studentSheet.getPrintSetup();
        printSetup.setFitHeight((short) 0);
        printSetup.setFitWidth((short) 1);
    }

}
