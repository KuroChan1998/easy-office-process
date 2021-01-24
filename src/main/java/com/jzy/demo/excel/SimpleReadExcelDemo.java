package com.jzy.demo.excel;

import com.jzy.office.excel.DefaultExcel;
import com.jzy.office.exception.InvalidFileTypeException;

import java.io.IOException;
import java.util.List;

/**
 * @ClassName SimpleReadWord2003Demo
 * @Author JinZhiyun
 * @Description 一个简单的读excel示例，处理的示例表格为项目example目录下的test1.xlsx
 * @Date 2020/4/1 10:59
 * @Version 1.0
 **/
public class SimpleReadExcelDemo {
    public static void main(String[] args) throws IOException, InvalidFileTypeException {
        //改成你实际的文件路径
        String filePath = "E:\\Engineering\\java\\idea\\easy-office-process\\example\\test1.xlsx";
        //通过文件绝对路径构造excel对象
        DefaultExcel excel = new DefaultExcel(filePath);

        //获得当前工作表sheet数量
        int sheetCount = excel.getSheetCount();
        System.out.println("当前工作表共有" + sheetCount + "张sheet");

        int sheetIndex = 0;
        //获取第一张sheet的总行数
        int rowCount = excel.getRowCount(sheetIndex);
        System.out.println("当前工作表第" + (sheetIndex + 1) + "张sheet共有" + rowCount + "行");

        //获取第一张sheet的名称
        String sheetName = excel.getSheetName(sheetIndex);
        System.out.println("当前工作表第" + (sheetIndex + 1) + "张sheet名为：" + sheetName);

        //根据sheet名称获取该名称对应sheet的索引
        int targetSheetIndex = excel.getSheetIndex(sheetName);
        System.out.println("当前工作表名为\"" + sheetName + "\"的sheet的索引值为" + targetSheetIndex);

        //获取指定sheet指定行的所有值
        int rowIndex = 0; //第1行
        List<String> rowValue = excel.readRow(sheetIndex, rowIndex);
        System.out.println("当前工作表第" + (sheetIndex + 1) + "张sheet的第" + (rowIndex + 1) + "行值为：" + rowValue);

        //获取指定sheet指定列的所有值（从指定行开始到最后）
        int columnIndex = 0; //第1列
        int startRow = 1; //第2行开始
        List<String> columnValue = excel.readColumn(sheetIndex, startRow, columnIndex);
        System.out.println("当前工作表第" + (sheetIndex + 1) + "张sheet的第" + (columnIndex + 1) + "列（从第" + (startRow + 1) + "行开始）的值为：" + columnValue);

        //获取指定sheet指定行指定列的单元格中的内容
        String cellValue = excel.read(sheetIndex, rowIndex, columnIndex);
        System.out.println("当前工作表第" + (sheetIndex + 1) + "张sheet的第" + (rowIndex + 1) + "行第" + (columnIndex + 1) + "列的值为：" + cellValue);

        //获取指定sheet从指定行开始到指定行结束所有行的内容
        int endRow = 3;
        List<List<String>> rowValues = excel.readRows(sheetIndex, startRow, endRow);
        System.out.println("当前工作表第" + (sheetIndex + 1) + "张sheet从第" + (rowIndex + 1) + "行开始到第" + endRow + "行的值为：");
        rowValues.forEach((rv) -> System.out.println("\t" + rv));

        //获取指定sheet的全部内容
        List<List<String>> allData = excel.read(sheetIndex);
        System.out.println("当前工作表第" + (sheetIndex + 1) + "张sheet的所有数据为：");
        allData.forEach((data) -> System.out.println("\t" + data));
    }
}
