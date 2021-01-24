package com.jzy.demo.excel.input;

import lombok.Data;

/**
 * @ClassName Student
 * @Author JinZhiyun
 * @Description 学生信息封装
 * @Date 2020/4/1 21:47
 * @Version 1.0
 **/
@Data
public class Student {
    private String id;
    private Integer age;
    private String gender;

    public Student(String id, Integer age, String gender) {
        this.id = id;
        this.age = age;
        this.gender = gender;
    }

    public Student() {
    }
}