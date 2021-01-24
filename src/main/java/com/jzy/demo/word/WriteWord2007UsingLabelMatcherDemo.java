package com.jzy.demo.word;

import com.jzy.office.exception.InvalidFileTypeException;
import com.jzy.office.matcher.LabelMatcher;
import com.jzy.office.matcher.LabelMatchers;
import com.jzy.office.word.DefaultWord2007;

import java.io.IOException;
import java.util.HashMap;

/**
 * @ClassName SimpleReadWord2007Demo
 * @Author JinZhiyun
 * @Description 一个简单的写word示例，演示使用标签替换器对word 2007进行操作
 * @Date 2020/4/1 10:59
 * @Version 1.0
 **/
public class WriteWord2007UsingLabelMatcherDemo {
    public static void main(String[] args) throws IOException, InvalidFileTypeException {
        //改成你实际的文件路径
        String filePath = "D:\\CDesktop\\Engineering\\java\\idea\\easy-office-process\\example\\1.docx";

        //构造word 2007对象
        DefaultWord2007 word2007 = new DefaultWord2007(filePath);

        //创建准备替换的书签集
        HashMap<String, String> bookmark = new HashMap<>();
        bookmark.put("table", "0000");
        bookmark.put("label1", "1111");

        //使用bookmark使用默认标签匹配器替换第4段中内容
        int paraPos = 3;
        HashMap<String, String> replacedBookmark = word2007.replaceInParaUsingLabelMatcher(paraPos, bookmark, LabelMatchers.DEFAULT_LABEL_MATCHER);
        System.out.println("第" + (paraPos + 1) + "段所有被成功替换的书签为：" + replacedBookmark);

        //使用bookmark使用自定义标签匹配器替换第2段中内容
        int paraPos2 = 2;
        //使用匹配{{key}}的标签匹配器
        LabelMatcher lMatcher = LabelMatchers.getLabelMatcher("\\{\\{(.+?)\\}\\}");
        HashMap<String, String> replacedBookmark2 = word2007.replaceInParaUsingLabelMatcher(paraPos2, bookmark, lMatcher);
        System.out.println("第" + (paraPos2 + 1) + "段使用自定义标签匹配器后所有被成功替换的书签为：" + replacedBookmark2);

        String savePath = "C:\\Users\\92970\\Desktop\\1.docx";
        //word2007.saveAndClose();//这样会直接覆盖更新原文件
        word2007.saveAndClose(savePath);
    }
}
