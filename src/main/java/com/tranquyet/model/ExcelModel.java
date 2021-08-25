package com.tranquyet.model;

import com.tranquyet.constant.message.ErrorMessageConstant;
import com.tranquyet.util.screen.Screen;

import java.util.Date;

/**
 * ExcelModel will create model to contain the information of Excel file
 */
public class ExcelModel {

    private String nameTable;

    private String createdDate;

    private int currentPage;

    private int totalPage;

    public ExcelModel() {
        this.createdDate = new Date().toString();
    }

    public String getNameTable() {
        return nameTable;
    }

    public void setNameTable(String nameTable) {
        this.nameTable = nameTable;
    }


    public int getCurrentPage() {
        return currentPage;
    }

    public void setCurrentPage(int currentPage) {
        this.currentPage = currentPage;
    }

    public int getTotalPage() {
        return totalPage;
    }

    public void setTotalPage(int totalPage) {
        this.totalPage = totalPage;
    }

    /**
     * check for the right form name of Excel file
     *
     * @param name - name of Excel file
     * @return<pre> if name is valid
     *          return true
     *          else return false
     *          </pre>
     */
    public boolean validateName(String name) {
        if (name.isEmpty()) {
            Screen.showError(ErrorMessageConstant.EMPTY_NAME_EXCEL_FILE);
            return false;
        }
        ;
        return true;
    }

    @Override
    public String toString() {
        return "ExcelModel{" +
                "nameTable='" + nameTable + '\'' +
                ", createdDate='" + createdDate + '\'' +
                ", totalPage=" + totalPage +
                '}';
    }

}
