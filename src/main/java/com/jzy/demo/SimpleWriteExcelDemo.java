package com.jzy.demo;

import com.jzy.excel.DefaultExcel;
import com.jzy.excel.exception.InvalidFileTypeException;

import java.io.IOException;
import java.util.Arrays;

/**
 * @ClassName SimpleWriteExcelDemo
 * @Author JinZhiyun
 * @Description 一个简单的写excel示例，处理的示例表格为项目example目录下的test1.xlsx
 * @Date 2020/4/1 19:42
 * @Version 1.0
 **/
public class SimpleWriteExcelDemo {
    public static void main(String[] args) throws IOException, InvalidFileTypeException {
        //改成你实际的文件路径
        String filePath = "E:\\Engineering\\java\\idea\\excel-processing\\example\\test1.xlsx";
        //通过文件绝对路径构造excel对象
        DefaultExcel excel = new DefaultExcel(filePath);


        int sheetIndex=0;
        //清除第指定sheet中所有内容
        excel.clearSheet(sheetIndex);

        //往指定sheet的指定行批量写入数据
        int rowIndex = 0;
        excel.writeRow(sheetIndex, rowIndex, Arrays.asList("我是", "第一行", "新的数据"));

        //往指定sheet的指定列批量写入数据
        int columnIndex=0;
        excel.writeColumn(sheetIndex, columnIndex, Arrays.asList("我是", "第一列", "新的数据"));

        //往指定sheet的指定行和指定列的单元格写入数据
        int targetRow = 1; //第2行
        int targetColumn = 1; //第2列
        excel.write(sheetIndex,targetRow, targetColumn, "你好");

        //往指定sheet的指定范围内的单元格写入同一个数据
        int startRow = 3; //第4行开始
        int endRow=5; //第6行结束
        int startColumn = 0; //第1列开始
        int endColumn=3; //第4列结束
        excel.write(sheetIndex, startColumn, endColumn, startRow, endRow, "重复值");

        //删除指定sheet的指定行
        int rowToRemove=5;
        excel.removeRow(sheetIndex, rowToRemove);

        //删除指定sheet的多行
        int rowToRemoveStart=3;
        int rowToRemoveEnd=4;
        //删除第4~5行
        excel.removeRows(sheetIndex, rowToRemoveStart, rowToRemoveEnd);

        //修改后另存为文件路径
        String savePath="C:\\Users\\92970\\Desktop\\1.xlsx";
//        excel.save();//这样会直接覆盖更新原文件
        excel.save(savePath);
    }
}
