package com.jzy.word;

import org.apache.poi.xwpf.usermodel.XWPFDocument;

import java.io.*;
import java.util.HashMap;

/**
 * @author JinZhiyun
 * @version 1.0
 * @ClassName Test
 * @description //TODO
 * @date 2020/11/27 17:17
 **/
public class Test {
    public static void main(String[] args) throws IOException {
        InputStream inp = new FileInputStream("C:\\Users\\92970\\Desktop\\1.doc");
        XWPFDocument docx = new XWPFDocument(inp);
        inp.close();

        HashMap<String, String> bookmark = new HashMap<String, String>();

        bookmark.put("table", "★凌霄宝殿★");

        bookmark.put("书签1", "★你好邢志云★");

        bookmark.put("书签2", "★这是★");

        bookmark.put("书签3", "★美丽新世界★");

        bookmark.put("替换文字1", "★猫猫★");

        bookmark.put("不存在的书签", "★呵呵★");//不存在的会被忽略

        bookmark.put("替换文字3", "★111你好这是美丽新世界，替换其中的内容，再写入目标文档中★");

        Word.replaceSS(docx, bookmark,false);

        OutputStream outStream = new FileOutputStream("C:\\Users\\92970\\Desktop\\2.doc");
        docx.write(outStream);
        outStream.flush();
        outStream.close();
        docx.close();
    }
}
