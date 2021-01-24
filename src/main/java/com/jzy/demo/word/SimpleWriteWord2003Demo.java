package com.jzy.demo.word;

import com.jzy.office.exception.InvalidFileTypeException;
import com.jzy.office.word.DefaultWord2003;

import java.io.IOException;
import java.util.HashMap;

/**
 * @ClassName SimpleReadWord2003Demo
 * @Author JinZhiyun
 * @Description 一个简单的写word示例，演示针对word 2003的处理方法
 * @Date 2020/4/1 10:59
 * @Version 1.0
 **/
public class SimpleWriteWord2003Demo {
    public static void main(String[] args) throws IOException, InvalidFileTypeException {
        //改成你实际的文件路径
        String filePath = "D:\\CDesktop\\Engineering\\java\\idea\\easy-office-process\\example\\1.doc";

        //构造word 2003对象
        DefaultWord2003 word2003 = new DefaultWord2003(filePath);

        //创建准备替换的书签集
        HashMap<String, String> bookmark = new HashMap<>();
        bookmark.put("${table}", "0000");
        bookmark.put("${label1}", "1111");

        //使用bookmark替换第3段中内容
        int paraPos = 2;
        HashMap<String, String> replacedBookmark = word2003.replaceInPara(paraPos, bookmark);
        System.out.println("第" + (paraPos + 1) + "段所有被成功替换的书签为：" + replacedBookmark);

        //这样对全文所有段落进行替换
//        HashMap<String, String> replacedBookmark2 = word2003.replaceInParas(bookmark);

        //使用bookmark替换第1个表格中内容
        int tablePos = 0;
        HashMap<String, String> replacedBookmark3 = word2003.replaceInTable(tablePos, bookmark);
        System.out.println("第" + (tablePos + 1) + "个表格所有被成功替换的书签为：" + replacedBookmark3);

        //这样对全文所有表格进行替换
//        HashMap<String, String> replacedBookmark4 = word2003.replaceInTables(bookmark);

        //这样对全文所有段落和表格进行替换
//        HashMap<String, String> replacedBookmark5 =word2003.replaceInAll(bookmark);

        String savePath = "C:\\Users\\92970\\Desktop\\1.doc";
        //word2003.saveAndClose();//这样会直接覆盖更新原文件
        word2003.saveAndClose(savePath);
    }
}
