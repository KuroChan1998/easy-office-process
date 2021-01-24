package com.jzy.demo.excel;

import com.jzy.demo.excel.input.Student;
import com.jzy.office.excel.AbstractInputExcel;
import com.jzy.office.excel.ExcelWriteable;
import com.jzy.office.exception.ExcelColumnNotFoundException;
import com.jzy.office.exception.ExcelSheetNameInvalidException;
import com.jzy.office.exception.ExcelTooManyRowsException;
import com.jzy.office.exception.InvalidFileTypeException;
import lombok.Getter;
import lombok.Setter;

import java.io.IOException;
import java.util.*;

/**
 * @ClassName ReadableAndWriteableTest1Excel
 * @Author JinZhiyun
 * @Description 读取学生信息表test1.xlsx中的学生信息，根据学号写入对应的注册情况
 * @Date 2020/4/1 22:31
 * @Version 1.0
 **/
public class ReadableAndWriteableTest1Excel extends AbstractInputExcel implements ExcelWriteable {
    private static final String ID_COLUMN = "学号";
    private static final String REGISTRATION_COLUMN = "是否已注册";

    public ReadableAndWriteableTest1Excel(String inputFile) throws IOException, InvalidFileTypeException {
        super(inputFile);
    }

    /**
     * 读取到的学生信息
     */
    @Getter
    private List<Student> students;

    /**
     * 规定名称的列的索引位置，初始值为-1无效值，即表示还没找到
     */
    private int columnIndexOfId = -1;
    private int columnIndexOfRegistrationStatus = -1;


    /**
     * 学生注册情况，键=学号，值=注册情况
     */
    @Getter
    @Setter
    private Map<String, String> studentRegistrationStatus;

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
        columnIndexOfRegistrationStatus = -1;
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
                    case REGISTRATION_COLUMN:
                        columnIndexOfRegistrationStatus = i;
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
        if (columnIndexOfRegistrationStatus < 0) {
            throw new ExcelColumnNotFoundException(null, REGISTRATION_COLUMN);
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
            Student student = new Student();
            student.setId(id);
            students.add(student);
        }

        return effectiveDataRowCount;
    }

    @Override
    public boolean writeData() {
        int sheetIndex = 0;
        int rowCount = getRowCount(sheetIndex); // 表的总行数
        for (int i = DEFAULT_START_ROW + 1; i < rowCount; i++) {
            String id = read(sheetIndex, i, columnIndexOfId);
            String registrationStatus = studentRegistrationStatus.get(id);
            //写入注册状态
            write(sheetIndex, i, columnIndexOfRegistrationStatus, registrationStatus);
        }
        return true;
    }

    public static void main(String[] args) throws IOException, InvalidFileTypeException {
        //模拟数据库中学生注册情况
        int studentTotal = 50;
        Map<String, String> studentRegistrationStatusAtDataBase = new HashMap<>(studentTotal);
        for (int id = 1; id <= studentTotal; id++) {
            String status = ReadableAndWriteableTest1Excel.oneRandomNumber(0, 1) == 0 ? "否" : "是";
            studentRegistrationStatusAtDataBase.put(id + "", status);
        }
        System.out.println("模拟数据库中学生注册情况为：" + studentRegistrationStatusAtDataBase);

        //改成你实际的文件路径
        String filePath = "E:\\Engineering\\java\\idea\\easy-office-process\\example\\test1.xlsx";
        //通过文件绝对路径构造excel对象
        ReadableAndWriteableTest1Excel excel = new ReadableAndWriteableTest1Excel(filePath);
        try {
            //读学生
            excel.testAndReadData();
        } catch (ExcelTooManyRowsException e) {
            e.printStackTrace();
        } catch (ExcelColumnNotFoundException e) {
            e.printStackTrace();
        } catch (ExcelSheetNameInvalidException e) {
            e.printStackTrace();
        }

        List<Student> students = excel.getStudents();
        Map<String, String> studentRegistrationStatus = new HashMap<>(students.size());
        for (Student student : students) {
            //从数据库中查注册信息
            String status = studentRegistrationStatusAtDataBase.get(student.getId());
            //存储到注册信息临时变量
            studentRegistrationStatus.put(student.getId(), status);
        }

        //设置要写入的注册信息
        excel.setStudentRegistrationStatus(studentRegistrationStatus);
        //写注册信息
        excel.writeData();

        //修改后另存为文件路径
        String savePath = "C:\\Users\\92970\\Desktop\\1.xlsx";
        excel.saveAndClose(savePath);

        System.out.println("读取和修改完成，文件另存为:" + savePath);
    }

    /**
     * 返回start~end间的随机整数
     *
     * @param start 开始（含）
     * @param end   结束（含）
     * @return 随机整数
     */
    public static Integer oneRandomNumber(int start, int end) {
        Random random = new Random();
        int r = random.nextInt(end - start + 1) + start; //每次随机出一个数字（1-3）
        return r;
    }
}
