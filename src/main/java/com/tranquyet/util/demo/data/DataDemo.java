package com.tranquyet.util.demo.data;

import com.tranquyet.constant.student.StudentValueConstant;
import com.tranquyet.controller.ManagementStudentController;
import com.tranquyet.model.Student;
import com.tranquyet.util.student.AutoCodeGeneration;

import java.util.Random;

/**
 * DataDemo will create a demo dataset which help to describe more clearly Excel file at the list Student Sheet and Graph Sheet.
 * After per loop, one student object would be created with value like those but expert dob of student, which would be created
 * randomly at year of date, it would help to clear more at the demo dataset, specially at the graph sheet of Excel file
 */
public class DataDemo {

    /**
     * create the demo dataset to test create excel file
     */
    public static void test() {
        ManagementStudentController managementController = new ManagementStudentController();

        Random random = new Random();

        StringBuilder dobTemp = new StringBuilder();

        Student student = new Student();
        student.setCode(AutoCodeGeneration.generateCode());
        student.setName("Tran Xuan Quyet");
        student.setGender("Nam");
        student.setDob("05/06/2001");
        student.setLocation("Ha Noi");
        student.setEmail("quyettran@gmail.com");
        student.setPhone("0969764184");
        student.setCodeOfClass("BSD.BIZ");
        managementController.addStudent(student);

        /**
         * After per loop, one student object would be created with value like those but expert dob of student, which would be created
         * randomly at year of date, it would help to clear more at the demo dataset, specially at the graph sheet of Excel file
         * @param numbRandom create random number from (0-5) that use as index of DEMO_YEAR_DATA array
         * @param dobTemp create form dobTemp to set to student object like form (dd/MM/yyyy)
         */
        for (int i = 0; i < 100; i++) {
            int numbRandom = random.nextInt(5); //random number
            Student student1 = new Student();
            student1.setName("Demo Name");
            student1.setGender("Nam");
            student1.setLocation("Ha Noi");
            student1.setEmail("demo@gmail.com");
            student1.setPhone("0969764184");
            student1.setCodeOfClass("HelloWorld");
            dobTemp.append("05/06/").append(StudentValueConstant.DEMO_YEAR_DATA[numbRandom]); // create random dob for student
            student1.setDob(dobTemp.toString());
            student1.setCode(AutoCodeGeneration.generateCode()); // set random code for student
            managementController.addStudent(student1);
            dobTemp.delete(0, dobTemp.length()); // clear dobTemp after loop
        }

    }
}
