package com.tranquyet.util.excel.table;

import com.tranquyet.constant.excel.TableExcelValueConstant;
import com.tranquyet.model.ExcelModel;
import com.tranquyet.model.Student;
import com.tranquyet.service.StudentStorageService;
import com.tranquyet.util.excel.table.chart.ChartTableExcel;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.XSSFSheet;

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
     *              createBreakPage: set break page for student sheet
     * </pre>
     *
     * @param studentSheet   create List Student Sheet
     * @param componentExcel getting instance to embed the style for cell and border
     * @param wb             representation of an Excel workbook
     */
    public static void createPage(Sheet studentSheet, ComponentExcel componentExcel, Workbook wb
    ) {


        studentSheet.setDefaultRowHeightInPoints(19);
        ExcelModel excelModel = new ExcelModel();
        excelModel.setNameTable("Demo");
        int totalPage = StudentStorageService.pagination(TableExcelValueConstant.NUMBER_ROW_PER_PAGE).size();
        excelModel.setTotalPage(totalPage);
        List<List<Student>> listStudent = StudentStorageService.pagination(TableExcelValueConstant.NUMBER_ROW_PER_PAGE);

        for (int i = 0; i < totalPage; i++) {
            excelModel.setCurrentPage(i + 1);
            /**
             * Dynamic value which depend on pageNumber, any time page number increase will create a new page
             * and the location of cell, name of table will increase pageNumber time
             */
            int startRow = i * TableExcelValueConstant.ALL_ROW_PER_PAGE;
            if (i == totalPage - 1) {
                createFieldTable(studentSheet, wb, startRow, componentExcel); // generate the field for table
                createNameTableAndPage(studentSheet, wb, startRow, componentExcel, excelModel); //generate the name of table, page number
                createCellAndBindLastPage(studentSheet, componentExcel, wb, startRow, listStudent.get(i), excelModel);

                studentSheet.getPrintSetup().setPaperSize(PrintSetup.A4_PAPERSIZE);
                studentSheet.getPrintSetup().setLandscape(false);
                studentSheet.setFitToPage(true);
                PrintSetup printSetup = studentSheet.getPrintSetup();
                printSetup.setFitHeight((short) 0);
                printSetup.setFitWidth((short) 1);

            } else {
                createFieldTable(studentSheet, wb, startRow, componentExcel); // generate the field for table
                createNameTableAndPage(studentSheet, wb, startRow, componentExcel, excelModel); //generate the name of table, page number
                createCellAndBind(studentSheet, componentExcel, wb, startRow, listStudent.get(i)); // generate the cell of table and bind the value into them

            }
        }

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
        if (startRow == studentSheet.getLastRowNum()) {
            pageCell.setCellValue(TableExcelValueConstant.PAGE_TABLE + (excelModel.getCurrentPage() + 1) + "/" + (excelModel.getTotalPage() + 1));
        }else if(StudentStorageService.getTotal()==34){
            pageCell.setCellValue(TableExcelValueConstant.PAGE_TABLE + (excelModel.getCurrentPage()) + "/" + (excelModel.getTotalPage() + 1));
        }
        else {
            pageCell.setCellValue(TableExcelValueConstant.PAGE_TABLE + excelModel.getCurrentPage() + "/" + (excelModel.getTotalPage()));
        }
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
        for (int i = 0; i < numberRow; i++) {
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

            if (i == 33) {
                studentSheet.setRowBreak(studentSheet.getLastRowNum());
            }
        }
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
    public static void createCellAndBindLastPage(Sheet studentSheet, ComponentExcel componentExcel, Workbook wb, int startRow, List<Student> list, ExcelModel excelModel) {
        /**
         *  create cell and bind the value
         */
        Cell cell = null;
        Row row = null;
        int numberRow = list.size();
        /**
         * (numberRow + 2) create the row to draw main bottom border
         */
        for (int i = 0; i < numberRow + 20; i++) {
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
            if (i == numberRow) {
                if (list.size() > 20) {
                    studentSheet.setRowBreak(studentSheet.getLastRowNum() - 1);
                    createNameTableAndPage(studentSheet, wb, studentSheet.getLastRowNum(), componentExcel, excelModel);
                    ChartTableExcel.createChartTable((XSSFSheet) studentSheet, 5 + i + startRow);
                } else {
                    ChartTableExcel.createChartTable((XSSFSheet) studentSheet, 5 + i + startRow);
                }

            }

        }
    }

}
