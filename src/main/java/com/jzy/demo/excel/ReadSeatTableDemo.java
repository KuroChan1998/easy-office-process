package com.jzy.demo.excel;

import com.jzy.office.excel.DefaultExcel;
import com.jzy.office.exception.InvalidFileTypeException;
import org.apache.commons.lang3.StringUtils;

import java.io.IOException;

/**
 * @ClassName ReadSeatTableDemo
 * @Author JinZhiyun
 * @Description 读取座位表信息示例，处理的示例表格为项目example目录下的座位表.xlsx
 * @Date 2020/4/1 20:22
 * @Version 1.0
 **/
public class ReadSeatTableDemo {
    public static void main(String[] args) throws IOException, InvalidFileTypeException {
        //改成你实际的文件路径
        String filePath = "E:\\Engineering\\java\\idea\\easy-office-process\\example\\座位表.xlsx";
        //通过文件绝对路径构造excel对象
        DefaultExcel seatTableExcel = new DefaultExcel(filePath);

        //总共多少张sheet
        int sheetCount = seatTableExcel.getSheetCount();
        System.out.println("共有" + sheetCount + "个教室。");
        for (int sheetIndex = 0; sheetIndex < sheetCount; sheetIndex++) {
            //遍历每个教室
            System.out.print("当前教室门牌号为：" + seatTableExcel.getSheetName(sheetIndex) + "；");

            int rowCount = seatTableExcel.getRowCount(sheetIndex);
            Integer maxCapacity = null;
            for (int j = 0; j < rowCount; j++) {
                for (int k = 0; k < seatTableExcel.getColumnCount(sheetIndex, j); k++) {
                    //遍历表格所有行
                    String value = seatTableExcel.read(sheetIndex, j, k);
                    if (StringUtils.isNumeric(value)) {
                        //对所有为数字的单元格找到最大的作为当前教室容量
                        Integer cap = Integer.parseInt(value);
                        if (maxCapacity == null || cap > maxCapacity) {
                            maxCapacity = cap;
                        }
                    }
                }
            }
            System.out.println("容量：" + maxCapacity);

        }


    }
}
