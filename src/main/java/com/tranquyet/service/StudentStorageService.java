package com.tranquyet.service;

import com.tranquyet.constant.message.ErrorMessageConstant;
import com.tranquyet.model.Student;
import com.tranquyet.util.screen.Screen;

import java.util.*;
import java.util.concurrent.atomic.AtomicInteger;
import java.util.stream.Collectors;

/**
 * Student Storage use to save the information of student
 */
public class StudentStorageService {

    public static List<Student> listStudent = new ArrayList<Student>();// use to store all the student information

    /**
     * add the information of student
     *
     * @param student student object use to add to storage
     *                <p>
     *                <p>
     *                if student != null
     *                add to List
     *                else
     *                show error message
     */
    public static void addStudent(Student student) {
        if (student != null) {
            listStudent.add(student);
        } else {
            Screen.showError(ErrorMessageConstant.STUDENT_MODEL_ERROR);
        }
    }

    /**
     * @return the number of student in the list
     */
    public static List<Student> getAll() {
        return listStudent;
    }

    /**
     *
     */
    public static List<Student> getStudentForTableSheet(){
        List<Student> studentTableSheet = new ArrayList<>();
        getAll().forEach(p->{
            Student student = new Student();
            student.setCode(p.getCode());
            student.setName(p.getName());
            student.setGender(p.getGender());
            student.setDob(p.getDob());
            student.setPhone(p.getPhone());
            studentTableSheet.add(student);
        });
        return studentTableSheet;
    }

    /**
     * show the all the information of student
     */
    public static void showStudents() {
        System.out.println("Total: " + getTotal());
        AtomicInteger order = new AtomicInteger();//  the order number of student
        listStudent.forEach(p -> {
            order.getAndIncrement();// increase 1 per time
            System.out.println(order + ". " + p.toString());
        });
    }

    /**
     * @return Total student in database
     */
    public static int getTotal() {
        return listStudent.size();
    }

    /**
     * pagination of sheet 20 student/page
     *
     * @param studentPerPage number student per page
     *                       <p>
     *                       <p>
     *                       for loop at i = 0, i < total, step 20
     *                       paging add sub arraylist of all
     * @return paging - a list contain sub_list which have size 20 or smaller 20 at the last sub_list
     */
    public static List<List<Student>> pagination(final int studentPerPage) {
        List<List<Student>> paging = new ArrayList<List<Student>>();
        final int N = getTotal();
        for (int i = 0; i < N; i += studentPerPage) {
            paging.add(new ArrayList<Student>( // create sub_list
                    getStudentForTableSheet().subList(i, Math.min(N, i + studentPerPage)))
            );
        }
        return paging;
    }

    /**
     * Count the number of student who have the same year of birth
     *
     * @return <pre>
     *     A map which have key - unique year
     *                      value - number of student have the same year
     * </pre>
     */
    public static Map<Integer, Long> countSameYear() {
        List<Integer> listAge = new ArrayList<>();
        getAll().forEach(p -> {
            listAge.add(p.getYear());
        });
        Map<Integer, Long> temp = listAge.stream().sorted()
                .collect(Collectors.groupingBy(c -> c, Collectors.counting())); // create key map contain unique year
        // and value map contain the number of student have the same year
        Map<Integer, Long> map =  temp.entrySet().stream()
                .sorted(Map.Entry.comparingByValue(Comparator.naturalOrder()))
                .limit(10)
                .collect(Collectors.toMap(
                        Map.Entry::getKey, Map.Entry::getValue, (e1, e2) -> e2, LinkedHashMap::new));
        return map;
    }

}
