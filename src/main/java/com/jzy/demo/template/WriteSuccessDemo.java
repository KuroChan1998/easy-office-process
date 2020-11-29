package com.jzy.demo.template;

import com.jzy.exception.InvalidFileTypeException;

import java.io.IOException;
import java.util.Date;

/**
 * @ClassName WriteSuccessDemo
 * @Author JinZhiyun
 * @Description 写入成功示例
 * @Date 2020/4/1 22:23
 * @Version 1.0
 **/
public class WriteSuccessDemo {
    public static void main(String[] args) throws IOException, InvalidFileTypeException {
        //改成你实际的文件路径
        String filePath = "E:\\Engineering\\java\\idea\\excel-processing\\example\\test1.xlsx";
        //通过文件绝对路径构造excel对象
        Test1TemplateExcel excel = new Test1TemplateExcel(filePath);

        String reviewer = "张三";
        Date today = new Date();

        //设置写入信息
        excel.setReviewer(reviewer);
        excel.setReviewDate(today);
        //写入
        excel.writeData();

        //修改后另存为文件路径
        String savePath = "C:\\Users\\92970\\Desktop\\1.xlsx";
        excel.save(savePath);

        System.out.println("审阅完成，文件保存到" + savePath);
    }
}
