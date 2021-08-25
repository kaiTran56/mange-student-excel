package com.tranquyet.constant.student;
/**
 * constant value for student information
 */
public class StudentValueConstant {
    /**
     * action title to input name
     */
    public static final String FIELD_NAME = "name";

    /**
     * action title to input phone
     */
    public static final String FIELD_PHONE = "phone"; //
    /**
     * action title to input dob
     */
    public static final String FIELD_DOB = "dob"; //
    /**
     * action title to input gender
     */
    public static final String FIELD_GENDER = "gender"; //
    /**
     * action title to input email
     */
    public static final String FIELD_EMAIL = "email"; //
    /**
     * action title to input location
     */
    public static final String FIELD_LOCATION = "location"; //
    /**
     * action title to input codeOfClass
     */
    public static final String FIELD_CODE_OF_CLASS = "CodeOfClass"; //
    /**
     * check form or email
     */
    public static final String REGEX_EMAIL = "^(.+)@(\\S+)$"; //
    /**
     * check form of phone
     */
    public static final String REGEX_PHONE = "^[0-9]{10}$"; //
    /**
     * check form of date
     */
    public static final String FORMAT_DATE = "dd/MM/yyyy"; //

    /**
     * action title to create Excel File
     */
    public static final String CREATE_EXCEL_FILE = "createExcel";

    /**
     * action title to create demo Excel File
     */
    public static final String DEMO_EXCEL_FILE = "demoExcel";

    /**
     * constant year array to help create dob in DataDemo
     */
    public static final String[] DEMO_YEAR_DATA = {"2000", "2001", "2002", "2003", "2004", "2005", "2006"};



}
