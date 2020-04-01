package com.jzy.demo.input;

import com.jzy.excel.AbstractInputExcel;
import com.jzy.excel.exception.ExcelColumnNotFoundException;
import com.jzy.excel.exception.InvalidFileTypeException;
import lombok.Getter;

import java.io.IOException;
import java.util.ArrayList;
import java.util.List;

/**
 * @ClassName Test1InputExcel
 * @Author JinZhiyun
 * @Description 从学生信息表test1.xlsx中读取学生信息，其中要求必须有“学号”、“年龄”两列，“性别”一列不是必须的。
 * 程序会从中寻找名为“学号”、“年龄”、“性别”的列，并将这些列的数据封装成Student对象。
 * @Date 2020/4/1 21:06
 * @Version 1.0
 **/
public class Test1InputExcel extends AbstractInputExcel {
    private static final String ID_COLUMN = "学号";
    private static final String AGE_COLUMN = "年龄";
    private static final String GENDER_COLUMN = "性别";

    /**
     * 读取到的学生信息
     */
    @Getter
    private List<Student> students;

    /**
     * 规定名称的列的索引位置，初始值为-1无效值，即表示还没找到
     */
    private int columnIndexOfId = -1;
    private int columnIndexOfAge = -1;
    private int columnIndexOfGender = -1;

    public Test1InputExcel(String inputFile) throws IOException, InvalidFileTypeException {
        super(inputFile);
    }

    /**
     * 非必须。重置所有暂存的读取结果（成员变量）
     */
    @Override
    public void resetOutput() {
        students = new ArrayList<>();
    }

    /**
     * 非必须。重置带读取的目标列索引值
     */
    @Override
    public void resetColumnIndex() {
        columnIndexOfId = -1;
        columnIndexOfAge = -1;
        columnIndexOfGender = -1;
    }

    /**
     * 非必须。找到要读取的列的位置索引值
     *
     * @param sheetIx 要处理的sheet的索引
     */
    @Override
    protected void findColumnIndexOfSpecifiedName(int sheetIx) {
        int row0ColumnCount = getColumnCount(sheetIx, DEFAULT_START_ROW); // 第startRow行的列数
        for (int i = 0; i < row0ColumnCount; i++) {
            String value = read(sheetIx, DEFAULT_START_ROW, i);
            if (value != null) {
                switch (value) {
                    case ID_COLUMN:
                        columnIndexOfId = i;
                        break;
                    case AGE_COLUMN:
                        columnIndexOfAge = i;
                        break;
                    case GENDER_COLUMN:
                        columnIndexOfGender = i;
                        break;
                    default:
                }
            }
        }
    }

    /**
     * 非必须。测试判断要读取的目标列是否找到（子类实现）。如果没有，抛出异常。
     *
     * @return
     * @throws ExcelColumnNotFoundException 规定名称的列未找到
     */
    @Override
    public boolean testColumnNameValidity() throws ExcelColumnNotFoundException {
        if (columnIndexOfId < 0) {
            throw new ExcelColumnNotFoundException(null, ID_COLUMN);
        }
        if (columnIndexOfAge < 0) {
            throw new ExcelColumnNotFoundException(null, AGE_COLUMN);
        }
        return true;
    }

    /**
     * 核心方法，具体的读取数据操作。将数据存储于类成员变量中。
     *
     * @param sheetIndex sheet索引
     * @return 当前sheet的有效行数
     */
    @Override
    public int readData(int sheetIndex) {
        int effectiveDataRowCount = 0;

        int rowCount = getRowCount(sheetIndex); // 表的总行数
        for (int i = DEFAULT_START_ROW + 1; i < rowCount; i++) {
            String id = read(sheetIndex, i, columnIndexOfId);
            String ageStr = read(sheetIndex, i, columnIndexOfAge);
            Integer age = Integer.parseInt(ageStr);
            String gender = read(sheetIndex, i, columnIndexOfGender);
            Student student = new Student(id, age, gender);
            students.add(student);
        }

        return effectiveDataRowCount;
    }
}
