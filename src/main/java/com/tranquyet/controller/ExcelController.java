package com.tranquyet.controller;

import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.List;
import java.util.Map;
import java.util.Scanner;

import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import com.tranquyet.constant.excel.TableExcelValueConstant;
import com.tranquyet.constant.message.ErrorMessageConstant;
import com.tranquyet.constant.student.StudentValueConstant;
import com.tranquyet.model.ExcelModel;
import com.tranquyet.model.Student;
import com.tranquyet.service.StudentStorageService;
import com.tranquyet.util.demo.data.DataDemo;
import com.tranquyet.util.excel.table.ComponentExcel;
import com.tranquyet.util.excel.table.ListStudentTable;
import com.tranquyet.util.screen.Screen;

/**
 * ExcelController will get the request from client to create Excel file consist of List Student Sheet and Graph Sheet
 *
 * @author tranquyet
 */
public class ExcelController {

    /**
     * receive the List student which contain sub_lists, having no more than 20 student per sub_list
     *
     * @return a List have sub_lists which have size <= 20
     */
    public List<List<Student>> pagingStudent() {
        return StudentStorageService.pagination(TableExcelValueConstant.NUMBER_ROW_PER_PAGE);
    }

    /**
     * @return acquire the map which have key - unique years and value - number of student have the same year of birth
     */
    public Map<Integer, Long> statisticSameYear() {
        return StudentStorageService.countSameYear();
    }

    /**
     * createDataExcel: create Sheet for Excel file comprise the List Student Sheet and Graph Student Sheet
     * where
     * List Student Sheet would draw the table contain the value of student
     * Chart Sheet would draw a bar chart describe the same year of birth of student
     *
     * @param name name of Excel file
     */
    public void createDataExcel(String name) {

        XSSFWorkbook wb = new XSSFWorkbook();

        XSSFSheet studentSheet = wb.createSheet(TableExcelValueConstant.LIST_STUDENT_SHEET_EXCEL); // create List Student Sheet

        //XSSFSheet chartSheet = wb.createSheet(GraphExcelValueConstant.CHART_SHEET_EXCEL); // create Chart Sheet


        ComponentExcel componentExcelInstance = ComponentExcel.getInstance(); // create instance to use style for cell and border of table in Excel file
        ExcelModel excelModel = new ExcelModel(); // create Excel model
        excelModel.setNameTable(name);

        ListStudentTable.createPage(studentSheet, componentExcelInstance, wb);

        //BarChartExcel.createGraphSheet(wb, chartSheet); // create the chart in Chart Sheet

        boolean checkStatus = createExcelFile(wb, excelModel); // create Excel file
        if (checkStatus == true) {
            Screen.showSuccess(TableExcelValueConstant.SUCCESSFUL_EXCEL_FILE);
        }
    }

    /**
     * create demo excel file with demo data
     */
    public void createDemoExcel() {
        XSSFWorkbook wb = new XSSFWorkbook();
        XSSFSheet studentSheet = wb.createSheet(TableExcelValueConstant.LIST_STUDENT_SHEET_EXCEL);
        XSSFSheet studentSheet_1 = wb.createSheet(TableExcelValueConstant.LIST_STUDENT_SHEET_EXCEL+"_(1)");
        XSSFSheet studentSheet_2 = wb.createSheet(TableExcelValueConstant.LIST_STUDENT_SHEET_EXCEL+"_(2)");
       // XSSFSheet chartSheet = wb.createSheet(GraphExcelValueConstant.CHART_SHEET_EXCEL);

        ComponentExcel componentExcelInstance = ComponentExcel.getInstance();

        DataDemo.test_1(19); //
        ExcelModel excelModel = new ExcelModel();
        excelModel.setNameTable("Demo");
        ListStudentTable.createPage(studentSheet, componentExcelInstance, wb);
        DataDemo.test_1(13); //
        ListStudentTable.createPage(studentSheet_1, componentExcelInstance, wb);
        DataDemo.test_1(19); //
        ListStudentTable.createPage(studentSheet_2, componentExcelInstance, wb);

        //BarChartExcel.createGraphSheet(wb, chartSheet);
        boolean checkStatus = createExcelFile(wb, excelModel);
        if (checkStatus == true) {
            Screen.showSuccess(TableExcelValueConstant.SUCCESSFUL_EXCEL_FILE);
        }
    }

    /**
     * create Excel file
     *
     * @param workbook   representation of an Excel workbook
     * @param excelModel representation of  Excel file comprise (name)
     * @return create excel file
     */
    public boolean createExcelFile(Workbook workbook, ExcelModel excelModel) {
        StringBuilder stringBuilder = new StringBuilder().append(excelModel.getNameTable()).append(".xlsx"); // name of Excel file following (excelModel.name + ".xlsx)"

        try (FileOutputStream outputStream = new FileOutputStream(stringBuilder.toString())) {
            workbook.write(outputStream);
        } catch (FileNotFoundException e) {
            Screen.showError(ErrorMessageConstant.FAIL_FILE_EXCEL);
            return false;
        } catch (IOException e) {
            Screen.showError(ErrorMessageConstant.FAIL_FILE_EXCEL);
            return false;
        }
        return true;
    }

    /**
     * create action respectively requirement as create Excel file or create Demo Excel file
     *
     * @param action representation of content action
     * @param input  receiving the information from client
     */
    public void excelAction(String action, Scanner input) {
        String value = null;
        switch (action) {
            case StudentValueConstant.CREATE_EXCEL_FILE:
                do {
                    System.out.println("Enter name of Excel file: ");
                    value = input.nextLine();
                } while (value.isEmpty());
                createDataExcel(value);
                break;
            case StudentValueConstant.DEMO_EXCEL_FILE:
                createDemoExcel();
                break;
        }

    }
}
