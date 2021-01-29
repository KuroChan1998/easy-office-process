package com.jzy.demo.word;

import com.jzy.office.exception.InvalidFileTypeException;
import com.jzy.office.word.DefaultWord2003;

import java.io.IOException;
import java.util.List;

/**
 * @ClassName SimpleReadWord2003Demo
 * @Author JinZhiyun
 * @Description 一个简单的读word示例，演示针对word 2003的处理方法
 * @Date 2020/4/1 10:59
 * @Version 1.0
 **/
public class SimpleReadWord2003Demo {
    public static void main(String[] args) throws IOException, InvalidFileTypeException {
        //改成你实际的文件路径
        String filePath = "E:\\Engineering\\java\\idea\\easy-office-process\\example\\1.doc";

        //构造word 2003对象
        DefaultWord2003 word2003 = new DefaultWord2003(filePath);

        //获得段落数量（包含表格中的段落）
        int paraNum = word2003.getParagraphNum();
        System.out.println("当前文档共有" + paraNum + "个段落（其中包含表格中的段落）。");

        //获得表格数量
        int tableNum = word2003.getTableNum();
        System.out.println("当前文档共有" + tableNum + "个表格。");

        //获取所有段落的文本，依次存于list中
//        List<String> paraStrings = word2003.readParagraphsToList();
//        paraStrings.forEach((data) -> System.out.println(data));

        //获取第paraPos+1段的文本
        int paraPos = 2;
        String para2Text = word2003.readParagraph(paraPos);
        System.out.println("第" + (paraPos + 1) + "段文本为：" + para2Text);

        //读取第tablePos+1个表格的内容
//        int tablePos = 0;
//        List<List<List<String>>> table1 = word2003.readTable(tablePos);
//        System.out.println("下面是第" + (tablePos + 1) + "个表格的内容展示...");
//        for (int i = 0; i < table1.size(); i++) {
//            System.out.println("\t第" + i + "行的内容是" + table1.get(i));
//        }

        //读取第1个表格第1列的文本
//        int tablePos = 0;
//        int columnPos = 0;
//        List<List<String>> columnStrings = word2003.readTableColumn(tablePos, colucolumnPosmn);
//        columnStrings.forEach((data) -> System.out.println(data));

        //读取第1个表格第1行的文本
//        int tablePos = 0;
//        int rowPos = 0;
//        List<List<String>> rowStrings = word2003.readTableRow(tablePos, rowPos);
//        rowStrings.forEach((data) -> System.out.println(data));

        //读取第1个表格第2行第3列的文本
        int tablePos = 0;
        int rowPos = 1;
        int columnPos = 2;
        List<String> cellStrings = word2003.readTable(tablePos, rowPos, columnPos);
        System.out.println("下面是第" + (tablePos + 1) + "个表格第" + (rowPos + 1) + "行第" + (columnPos + 1) + "列的单元格内容...当前单元格共" + cellStrings.size() + "段：");
        for (int i = 0; i < cellStrings.size(); i++) {
            System.out.println("\t当前单元格第" + (i + 1) + "段的文本为：" + cellStrings.get(i));
        }

//        rowStrings.forEach((data) -> System.out.println(data));
    }
}
