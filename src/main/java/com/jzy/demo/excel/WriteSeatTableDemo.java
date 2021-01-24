package com.jzy.demo.excel;

import com.jzy.office.excel.DefaultExcel;
import com.jzy.office.exception.InvalidFileTypeException;
import org.apache.commons.lang3.StringUtils;

import java.io.IOException;
import java.util.ArrayList;
import java.util.List;

/**
 * @ClassName WriteSeatTableDemo
 * @Author JinZhiyun
 * @Description 修改座位表信息示例，处理的示例表格为项目example目录下的座位表.xlsx。
 * 给定一个教室，一组学生姓名列表，要求在座位表模板的基础上只输出该教室的sheet并将列表中的学生按座位顺序依次填充
 * @Date 2020/4/1 20:30
 * @Version 1.0
 **/
public class WriteSeatTableDemo {
    public static void main(String[] args) throws IOException, InvalidFileTypeException {
        //制作201教室座位表
        String classroom = "201";
        //学生列表
        int studentCount = 20;
        List<String> students = new ArrayList<>(studentCount);
        for (int i = 1; i <= studentCount; i++) {
            char c = (char) ((int) 'a' + (i - 1));
            students.add("学生" + c);
        }
        System.out.println("共有" + studentCount + "个学生。分别为" + students);


        //改成你实际的文件路径
        String filePath = "E:\\Engineering\\java\\idea\\easy-office-process\\example\\座位表.xlsx";
        //通过文件绝对路径构造excel对象
        DefaultExcel seatTableExcel = new DefaultExcel(filePath);

        //先把其他没用的教室删掉
        int start = 0;
        int totalSheetCount = seatTableExcel.getSheetCount();
        for (int i = 0; i < totalSheetCount; i++) {
            if (seatTableExcel.getSheetName(start).equals(classroom)) {
                start++;
            } else {
                seatTableExcel.removeSheetAt(start);
            }
        }

        //开始依序填座位表
        int targetSheetIndex = 0; //在第0张sheet找
        int rowCount = seatTableExcel.getRowCount(targetSheetIndex);
        for (int i = 0; i < rowCount; i++) {
            //遍历表格所有行
            for (int j = 0; j < seatTableExcel.getColumnCount(targetSheetIndex, i); j++) {
                //遍历所有列
                String value = seatTableExcel.read(targetSheetIndex, i, j);
                if (StringUtils.isNumeric(value)) {
                    //对所有为数字的单元格（即座位号）填充姓名
                    int index = Integer.parseInt(value) - 1;
                    if (index < students.size()) {
                        //座位号值大于学生数量的座位不填
                        seatTableExcel.write(targetSheetIndex, i, j, value + " " + students.get(Integer.parseInt(value) - 1));
                        System.out.println("将\"" + students.get(Integer.parseInt(value) - 1) + "\"填充到第" + (i + 1) + "行第" + (j + 1) + "列的" + (index + 1) + "号座位中...");
                    }
                }
            }
        }


        //修改后另存为文件路径
        String savePath = "C:\\Users\\92970\\Desktop\\1.xlsx";
        seatTableExcel.saveAndClose(savePath);

        System.out.println(classroom + "教室座位表制作完成，文件保存到" + savePath);
    }
}
