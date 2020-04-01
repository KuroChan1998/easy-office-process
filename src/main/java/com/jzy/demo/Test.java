package com.jzy.demo;

import com.jzy.excel.DefaultExcel;
import com.jzy.excel.ExcelVersionEnum;
import com.jzy.excel.exception.InvalidFileTypeException;

/**
 * @ClassName Test
 * @Author JinZhiyun
 * @Description //TODO
 * @Date 2020/4/1 12:10
 * @Version 1.0
 **/
public class Test {
    public static void main(String[] args) throws InvalidFileTypeException {
        //创建一个新的excel2007文件
        DefaultExcel excel2007 =new DefaultExcel(ExcelVersionEnum.VERSION_2007);
        //创建一个新的excel2003文件
        DefaultExcel excel2003 =new DefaultExcel(ExcelVersionEnum.VERSION_2003);
    }
}
