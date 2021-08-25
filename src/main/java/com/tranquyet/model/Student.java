package com.tranquyet.model;

import com.tranquyet.constant.student.StudentValueConstant;
import com.tranquyet.constant.message.ErrorMessageConstant;
import com.tranquyet.util.screen.Screen;

import java.text.ParseException;
import java.text.SimpleDateFormat;
import java.util.Calendar;
import java.util.Date;
import java.util.regex.Pattern;

/**
 * Model of Student Information
 *
 * @attributes <pre>
 *         code         String
 *         name         String
 *         location     String
 *         dob          String
 *         gender       String
 *         phone        String
 *         email        String
 *         codeOfClass  String
 *     </pre>
 * A typical Student is c=(c,n,l,d,g,p,e,c) where code(c), name(n), location(l), dob(d), gender(g), phone(p), email(e), codeOfClass(c)
 */
public class Student {

    private String code;

    private String name;

    private String location;

    private String dob;

    private String gender;

    private String phone;

    private String email;

    private String codeOfClass;

    public void setCode(String code) {
        this.code = code;
    }

    public void setName(String name) {
        this.name = name;
    }

    public void setLocation(String location) {
        this.location = location;
    }

    /**
     * After binding the dob value for student, the dob will use to count for Age and split year of birth
     *
     * @param dob bind dob of student
     */
    public void setDob(String dob) {
        this.dob = dob;
        this.countCurrentAge(); // count Age
        this.getYear(); // split year of birth
    }

    public void setGender(String gender) {
        this.gender = gender;
    }

    public void setPhone(String phone) {
        this.phone = phone;
    }

    public void setEmail(String email) {
        this.email = email;
    }

    public void setCodeOfClass(String codeOfClass) {
        this.codeOfClass = codeOfClass;
    }

    public String getCode() {
        return code;
    }


    /**
     * split year of birth from dob
     *
     * @return year of dob
     */
    public int getYear() {
        int yearOfBirth = 0;
        SimpleDateFormat sdf = new SimpleDateFormat(StudentValueConstant.FORMAT_DATE);
        try {
            yearOfBirth = sdf.parse(dob).getYear() + 1900;
        } catch (ParseException e) {
            e.printStackTrace();
        }
        return yearOfBirth;
    }

    public String getName() {
        return name;
    }

    /**
     * check right form of name field
     *
     * @param name bind the name of student
     * @return <pre>
     *     if name valid
     *          return true
     *      else return false
     * </pre>
     */
    public boolean validateName(String name) {
        if (name == null || name.isEmpty()) {
            Screen.showError(ErrorMessageConstant.EMPTY_NAME_ERROR);
            return false;
        } else if (name.length() > 25) {
            Screen.showError(ErrorMessageConstant.NAME_FORMAT_ERROR);
            return false;
        }
        return true;
    }

    /**
     * check for right form of location
     *
     * @param location bind location of student
     * @return <pre>
     *     if location valid return true
     *     else return false
     * </pre>
     */
    public boolean validateLocation(String location) {
        return true;
    }

    public String getDob() {
        return dob;
    }

    /**
     * check for right form of dob
     *
     * @param dob bind dob of student
     * @return <pre>
     *     if dob valid return true
     *     else return false
     * </pre>
     */
    public boolean validateDob(String dob) {
        Date date = null;

        // check empty dob

        if (dob.isEmpty()) {
            Screen.showError(ErrorMessageConstant.EMPTY_DOB_ERROR);
            return false;
        }
        try {
            SimpleDateFormat sdf = new SimpleDateFormat(StudentValueConstant.FORMAT_DATE);
            date = sdf.parse(dob);
            // check format of dob
            if (!dob.equals(sdf.format(date))) {
                Screen.showError(ErrorMessageConstant.DATE_FORMAT_ERROR);
                return false;
            }
        } catch (ParseException ex) {
            Screen.showError(ErrorMessageConstant.DATE_FORMAT_ERROR);
            return false;
        }
        return true;
    }

    /**
     * count the current age by current year - year of birth
     *
     * @return return the current age
     */
    public int countCurrentAge() {
        int currentYear = Calendar.getInstance().get(Calendar.YEAR); // current year
        int yearOfBirth = 0;
        try {
            SimpleDateFormat sdf = new SimpleDateFormat(StudentValueConstant.FORMAT_DATE);
            yearOfBirth = sdf.parse(dob).getYear() + 1900; // year of dob
        } catch (ParseException e) {
            e.printStackTrace();
        }
        int currentAge = currentYear - yearOfBirth; // current age

        return currentAge;
    }

    public String getGender() {
        return gender;
    }

    /**
     * check for right form of gender Nam/nam or Nu/nu
     *
     * @param gender bind the gender of student
     * @return <pre>
     *     if gender is valid
     *          return true
     *     else
     *          return false
     * </pre>
     */
    public boolean validateGender(String gender) {
        if (gender.isEmpty()) { // check empty gender
            Screen.showError(ErrorMessageConstant.EMPTY_GENDER_ERROR);
            return false;
        }
        // check format gender
        else if (!gender.equals("Nam") && !gender.equals("nam") && !gender.equals("nu") && !gender.equals("Nu")) {
            Screen.showError(ErrorMessageConstant.GENDER_FORMAT_ERROR);
            return false;
        }
        return true;
    }

    public String getPhone() {
        return phone;
    }

    /**
     * check for right form of phone
     *
     * @param phone bind phone of student
     * @return <pre>
     *     if phone is valid
     *          return true
     *     else
     *          return false
     * </pre>
     */
    public boolean validatePhone(String phone) {
        if (Pattern.matches(StudentValueConstant.REGEX_PHONE, phone) == false) { // check format phone
            Screen.showError(ErrorMessageConstant.PHONE_FORMAT_ERROR);
            return false;
        } else if (phone.isEmpty() && phone == null) { // check empty phone
            Screen.showError(ErrorMessageConstant.EMPTY_PHONE_ERROR);
            return false;
        }
        return true;
    }

    public String getEmail() {
        return email;
    }

    /**
     * @return arr which contain all information of student
     */
    public String[] getArrayInformation() {
        String[] arrInfor = {code, name, gender, dob, phone, countCurrentAge() + ""};
        return arrInfor;
    }

    /**
     * check for right form of email
     *
     * @param email bind email of student
     * @return <pre>
     *     if email is valid
     *          return true
     *     else
     *          return false
     * </pre>
     */
    public boolean validateEmail(String email) {
        if (Pattern.matches(StudentValueConstant.REGEX_EMAIL, email) == false) { // check format email
            Screen.showError(ErrorMessageConstant.EMAIL_FORMAT_ERROR);
            return false;
        }
        return true;
    }

    public String getCodeOfClass() {
        return codeOfClass;
    }

    /**
     * check for right form of codeOfClass
     *
     * @param codeOfClass bind codeOfClass of student
     * @return true
     */
    public boolean validateCodeOfClass(String codeOfClass) {
        return true;
    }

    @Override
    public String toString() {
        return "Student{" +
                "code='" + code + '\'' +
                ", name='" + name + '\'' +
                ", location='" + location + '\'' +
                ", dob='" + dob + '\'' +
                ", gender='" + gender + '\'' +
                ", phone='" + phone + '\'' +
                ", email='" + email + '\'' +
                ", codeOfClass='" + codeOfClass + '\'' +
                '}';
    }
}
