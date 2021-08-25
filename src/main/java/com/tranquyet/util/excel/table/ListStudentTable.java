package com.tranquyet.util.excel.table;

import com.tranquyet.constant.excel.TableExcelValueConstant;
import com.tranquyet.model.ExcelModel;
import com.tranquyet.model.Student;
import com.tranquyet.util.excel.ComponentExcel;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellRangeAddress;

import java.util.List;

/**
 * ListStudentTable will create table, bind value to cell
 */
public class ListStudentTable {
    /**
     * active method:<pre>
     *              createFieldTable: generate the field for table
     *              createNameTableAndPage: generate the name of table, page number
     *              createCellAndBind: generate the cell of table and bind the value into them
     * </pre>
     *
     * @param studentSheet   create List Student Sheet
     * @param componentExcel getting instance to embed the style for cell and border
     * @param wb             representation of an Excel workbook
     * @param excelModel     embedding the name of table and current page and total page
     * @param pageNumber     current page, if any time it increase, a new page create
     * @param list           contain the student information, have no more than 20 student per/list
     */
    public static void createPage(Sheet studentSheet, ComponentExcel componentExcel, Workbook wb
            , ExcelModel excelModel, int pageNumber, List<Student> list) {


        studentSheet.setDefaultRowHeightInPoints(23);

        /**
         * Dynamic value which depend on pageNumber, any time page number increase will create a new page
         * and the location of cell, name of table will increase pageNumber time
         */
        int startRow = pageNumber * TableExcelValueConstant.ALL_ROW_PER_PAGE; // dynamic start row for cell which hold the value

        createFieldTable(studentSheet, wb, startRow, componentExcel); // generate the field for table

        createNameTableAndPage(studentSheet, wb, startRow, componentExcel, excelModel); //generate the name of table, page number

        createCellAndBind(studentSheet, componentExcel, wb, startRow, list); // generate the cell of table and bind the value into them

        createBreakPage(studentSheet);

    }

    /**
     * generate the field for table
     *
     * @param studentSheet   create List Student Sheet
     * @param wb             representation of an Excel workbook
     * @param startRow       Dynamic value which depend on pageNumber, any time page number increase will create a new page
     *                       and the location of cell, name of table will increase pageNumber time
     * @param componentExcel embedding the name of table and current page and total page
     */
    public static void createFieldTable(Sheet studentSheet, Workbook wb
            , int startRow, ComponentExcel componentExcel) {
        /**
         *  create field for excel table
         */
        Row row = null;
        Cell cell;
        /**
         * create area for field table
         */
        for (int i = 0; i < 1; i++) {
            int positionTemp = i + 2;
            row = studentSheet.createRow(positionTemp + startRow); // row
            /**
             *  bind value for field table
             */
            for (int j = 0; j < TableExcelValueConstant.NUMBER_FIELD_EXCEL; j++) {
                cell = row.createCell(j + 1);
                if (TableExcelValueConstant.ARRAY_FIELD_EXCEL[j].equals(TableExcelValueConstant.NAME_FIELD_TABLE)) {
                    studentSheet.setColumnWidth(cell.getColumnIndex(), 30 * 256);
                    cell.setCellValue(TableExcelValueConstant.ARRAY_FIELD_EXCEL[j]); // set name field
                    componentExcel.createStyleNameField(cell, wb); // set style for header
                } else if (TableExcelValueConstant.ARRAY_FIELD_EXCEL[j].equals(TableExcelValueConstant.CODE_FIELD_TABLE)) {
                    studentSheet.setColumnWidth(cell.getColumnIndex(), 25 * 256);
                    cell.setCellValue(TableExcelValueConstant.ARRAY_FIELD_EXCEL[j]); // set name field
                    componentExcel.createStyleFieldTable(cell, wb); // set style for header
                } else if (TableExcelValueConstant.ARRAY_FIELD_EXCEL[j].equals(TableExcelValueConstant.GENDER_FIELD_TABLE)) {
                    studentSheet.setColumnWidth(cell.getColumnIndex(), 24 * 256);
                    cell.setCellValue(TableExcelValueConstant.ARRAY_FIELD_EXCEL[j]); // set name field
                    componentExcel.createStyleFieldTable(cell, wb); // set style for header
                } else if (TableExcelValueConstant.ARRAY_FIELD_EXCEL[j].equals(TableExcelValueConstant.DOB_FIELD_TABLE)) {
                    studentSheet.setColumnWidth(cell.getColumnIndex(), 30 * 256);
                    cell.setCellValue(TableExcelValueConstant.ARRAY_FIELD_EXCEL[j]); // set name field
                    componentExcel.createStyleFieldTable(cell, wb); // set style for header
                } else if (TableExcelValueConstant.ARRAY_FIELD_EXCEL[j].equals(TableExcelValueConstant.PHONE_FIELD_TABLE)) {
                    studentSheet.setColumnWidth(cell.getColumnIndex(), 30 * 256);
                    cell.setCellValue(TableExcelValueConstant.ARRAY_FIELD_EXCEL[j]); // set name field
                    componentExcel.createStyleFieldTable(cell, wb); // set style for header
                } else if (TableExcelValueConstant.ARRAY_FIELD_EXCEL[j].equals(TableExcelValueConstant.AGE_FIELD_TABLE)) {
                    studentSheet.setColumnWidth(cell.getColumnIndex(), 10 * 256);
                    cell.setCellValue(TableExcelValueConstant.ARRAY_FIELD_EXCEL[j]); // set name field
                    componentExcel.createStyleFieldTable(cell, wb); // set style for header
                }
            }


        }
    }

    /**
     * generate the name of table, page number
     *
     * @param studentSheet   create List Student Sheet
     * @param wb             representation of an Excel workbook
     * @param startRow       Dynamic value which depend on pageNumber, any time page number increase will create a new page
     *                       and the location of cell, name of table will increase pageNumber time
     * @param componentExcel embedding the name of table and current page and total page
     * @param excelModel     embedding the name of table and current page and total page
     */
    public static void createNameTableAndPage(Sheet studentSheet, Workbook wb
            , int startRow
            , ComponentExcel componentExcel
            , ExcelModel excelModel) {
        /**
         * : create header
         */
        Cell headerCell = null;
        Row headerRow = studentSheet.createRow(0 + startRow); // row
        headerCell = headerRow.createCell(2);
        headerCell.setCellValue(TableExcelValueConstant.HEADER_NAME_TABLE);
        studentSheet.addMergedRegion(new CellRangeAddress(
                0 + startRow, //first row (0-based)
                0 + startRow, //last row  (0-based)
                2, //first column (0-based)
                4  //last column  (0-based)
        ));
        componentExcel.createStyleHeader(headerCell, wb);

/**
 *  create page number
 */
        Cell pageCell = null;
        pageCell = headerRow.createCell(6);
        pageCell.setCellValue(TableExcelValueConstant.PAGE_TABLE + excelModel.getCurrentPage() + "/" + excelModel.getTotalPage());
        studentSheet.addMergedRegion(new CellRangeAddress(
                0 + startRow, //row
                0 + startRow, //row
                6,
                7
        ));
        componentExcel.createStylePage(pageCell, wb);

    }

    /**
     * generate the cell of table and bind the value into them and draw the main bottom border for table
     *
     * @param studentSheet   create List Student Sheet
     * @param componentExcel embedding the name of table and current page and total page
     * @param wb             representation of an Excel workbook
     * @param startRow       Dynamic value which depend on pageNumber, any time page number increase will create a new page
     *                       and the location of cell, name of table will increase pageNumber time
     * @param list           contain the student information, have no more than 20 student per/list
     */
    public static void createCellAndBind(Sheet studentSheet, ComponentExcel componentExcel, Workbook wb, int startRow, List<Student> list) {
        /**
         *  create cell and bind the value
         */
        Cell cell = null;
        Row row = null;
        int numberRow = list.size();
        /**
         * (numberRow + 2) create the row to draw main bottom border
         */
        //CreationHelper createHelper = wb.getCreationHelper();
        //String formula = "=INT((TODAY()-E4)/365)";
        for (int i = 0; i < numberRow + 2; i++) {
            row = studentSheet.createRow(3 + i + startRow); // row start cell need +5 because 5 row before had used by field table and main top border
            /**
             * bind the value to cell
             */
            if (i < numberRow) {
                for (int j = 0; j < TableExcelValueConstant.NUMBER_FIELD_EXCEL; j++) {
                    cell = row.createCell(j + 1);
                    if (j == 1) {
                        cell.setCellValue(list.get(i).getArrayInformation()[j]);// set value from student
                        componentExcel.createStyleNameColumn(cell, wb); // set style for cell
                    } else {
                        cell.setCellValue(list.get(i).getArrayInformation()[j]);// set value from student
                        componentExcel.createStyleCellTable(cell, wb); // set style for cell
                    }
                }
            }
        }

    }

    /**
     * @param studentSheet
     */
    public static void createBreakPage(Sheet studentSheet) {
        int count = 0;
        int check = 3;
        studentSheet.setRowBreak(22);
        for (int i = 0; i < studentSheet.getLastRowNum(); i++) {
            if (count % 20 == 0 && count != 0) {
                studentSheet.setRowBreak(22 + count + check);
                check += 3;
            }
            count++;
        }
        studentSheet.getPrintSetup().setPaperSize(PrintSetup.A4_PAPERSIZE);
        studentSheet.getPrintSetup().setLandscape(true);
        studentSheet.setFitToPage(true);
        PrintSetup printSetup = studentSheet.getPrintSetup();
        printSetup.setFitHeight((short) 0);
        printSetup.setFitWidth((short) 1);
    }

}
