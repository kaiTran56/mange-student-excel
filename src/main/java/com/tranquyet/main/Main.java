package com.tranquyet.main;

import com.tranquyet.constant.student.StudentValueConstant;
import com.tranquyet.constant.message.ErrorMessageConstant;
import com.tranquyet.controller.ExcelController;
import com.tranquyet.controller.ManagementStudentController;
import com.tranquyet.model.Student;
import com.tranquyet.util.student.AutoCodeGeneration;
import com.tranquyet.util.screen.Screen;

import java.util.Scanner;
/**
 * run application
 */
public class Main {
    public static Scanner scanner = new Scanner(System.in);

    public static void main(String... args) throws Exception {
        ManagementStudentController manage = new ManagementStudentController();
        ExcelController excelController = new ExcelController();
        int option;
        while (true) {
            try {
                Screen.showScreen();
                option = Integer.parseInt(scanner.nextLine());
                switch (option) {
                    case 1:// option 1: add student
                        Student student = new Main().enterInformation(scanner, manage);
                        manage.addStudent(student);
                        break;
                    case 2: // option 2: show student
                        manage.show();
                        break;
                    case 3 :// create Excel file
                        excelController.excelAction(StudentValueConstant.CREATE_EXCEL_FILE, scanner);
                        break;
                    case 4:// create demo Excel file
                        excelController.excelAction(StudentValueConstant.DEMO_EXCEL_FILE, null);
                        break;
                    case 5:// option 5: exit
                        manage.exit();
                        break;
                    default:
                        Screen.showError(ErrorMessageConstant.OPTION_NUMBER_ERROR);
                }
            } catch (Exception e) {
                Screen.showError(ErrorMessageConstant.OPTION_NUMBER_ERROR);
            }
        }
    }

    /**
     * receive the input value of student and checking the valid condition
     *
     * @param input  receive the information from client
     * @param manage call the checkValue method to retrieve the right form information
     * @return the valid student object after checking
     */
    public Student enterInformation(Scanner input, ManagementStudentController manage) {
        Student student = new Student();

        student.setCode(AutoCodeGeneration.generateCode()); // code field <automatically create>
        System.out.println("Code: " + student.getCode());

        String name = manage.checkValueStudent(input, StudentValueConstant.FIELD_NAME, student); // name field
        student.setName(name);

        String dob = manage.checkValueStudent(input, StudentValueConstant.FIELD_DOB, student); // dob field
        student.setDob(dob);

        String phone = manage.checkValueStudent(input, StudentValueConstant.FIELD_PHONE, student); // phone field
        student.setPhone(phone);

        String email = manage.checkValueStudent(input, StudentValueConstant.FIELD_EMAIL, student); // email field
        student.setEmail(email);

        String gender = manage.checkValueStudent(input, StudentValueConstant.FIELD_GENDER, student); // gender field
        student.setGender(gender);

        String location = manage.checkValueStudent(input, StudentValueConstant.FIELD_LOCATION, student); // location field
        student.setLocation(location);

        String codeOfClass = manage.checkValueStudent(input, StudentValueConstant.FIELD_CODE_OF_CLASS, student); // codeOfClass field
        student.setCodeOfClass(codeOfClass);

        return student;
    }

}
