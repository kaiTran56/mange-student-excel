package com.tranquyet.constant.excel;

/**
 * TableExcelValueConstant will hold the constant variable of List Student Sheet
 */
public class TableExcelValueConstant {
    /**
     * Name of List Student Sheet
     */
    public static final String LIST_STUDENT_SHEET_EXCEL = "List Students Sheet";

    /**
     * Successful message when create Excel file
     */
    public static final String SUCCESSFUL_EXCEL_FILE = "Create file ExcelModel, Successfully!";

    /**
     * font of word
     */
    public static final String FONT_STYLE_EXCEL = "Tahoma";

    /**
     * size of word
     */
    public static final short SIZE_WORD_EXCEL = 14;

    /**
     * size of header
     */
    public static final short SIZE_HEADER_TABLE_EXCEL = 20;

    /**
     * Bold style of word
     */
    public static final boolean BOLD_WORD_EXCEL = false;

    /**
     * Bold style of header
     */
    public static final boolean BOLD_HEADER_EXCEL = true;

    /**
     * constant number row of table to set value
     */
    public static final int NUMBER_ROW_PER_PAGE = 34;

    /**
     * name of table
     */
    public static final String HEADER_NAME_TABLE = "DANH SÁCH SINH VIÊN";

    /**
     * form of page show in sheet
     */
    public static final String PAGE_TABLE = "Page: ";

    /**
     * constant all row of a full page, have enough 20 student / page
     */
    public static final int ALL_ROW_PER_PAGE = 37;

    /**
     * All the field in the table
     */
    public static final String[] ARRAY_FIELD_EXCEL = {"mã số sinh viên", "tên sinh viên", "giới tính", "ngày tháng năm sinh", "số điện thoại"
            , "Tuổi"};

    public static final String NAME_FIELD_TABLE = ARRAY_FIELD_EXCEL[1];

    public static final String CODE_FIELD_TABLE = ARRAY_FIELD_EXCEL[0];

    public static final String GENDER_FIELD_TABLE = ARRAY_FIELD_EXCEL[2];

    public static final String DOB_FIELD_TABLE = ARRAY_FIELD_EXCEL[3];

    public static final String PHONE_FIELD_TABLE = ARRAY_FIELD_EXCEL[4];

    public static final String AGE_FIELD_TABLE = ARRAY_FIELD_EXCEL[5];
    /**
     * count number field
     */
    public static final int NUMBER_FIELD_EXCEL = ARRAY_FIELD_EXCEL.length;



}
