package com.jzy.word;

import java.util.*;
import java.util.regex.*;
import org.apache.poi.xwpf.usermodel.*;

public class Word
{
    //■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■
    /**
     替换Word中指定的字符串“${标签1}”，在Word中需要替换的地方写入这样的字符串，用下面代码实现标签替换（<a style='color:red;'>注：Word中的标签可以重复，都会被替换</a>）<br/>
     例子❶：替换标签内容<br/>
     HashMap&lt;String, String> bookmark = new HashHashMap&lt;String, String>();<br/>
     bookmark.put("标签1", "你好邢志云");<br/>
     bookmark.put("标签2", "这是");<br/>
     Word.replaceSS(docx, bookmark,false);<br/>
     例子❷：列出word中所有的标签（不替换标签），顺序是先段落中的标签，然后是表格中的标签。<br/>
     ArrayList&lt;String> al=Word.replaceSS(docx, null,false);<br/>
     for(int i=0;i&lt;al.size();i++)<br/>
     System.out.println(al.get(i));<br/>
     例子❸：检查Word中标签是否设置正确（方法是把标签直接替换掉标签的位置，用眼睛直观能看出哪些地方设置了标签）<br/>
     Word.replaceSS(docx, null,true);

     * @param docx
     * @param bookmark null时表示不替换标签。
     * @param ShowKey 默认为false，true的时候标签的地方会显示为    ▶ 标签名◀    ，能方便的看出标签是否设置正确。
     * @return 返回值是个ArrayList，存储了所有在文档中出现的标签名称，${标签1}→标签1，顺序是先段落中的标签，然后是表格中的标签。
     */
    public static ArrayList<String> replaceSS(XWPFDocument docx, HashMap<String, String> bookmark,boolean ShowKey)
    {
        ArrayList<String> returnlist= new ArrayList<String>();
        ArrayList<String> rl;
        //替换段落里面的变量
        rl=replaceInParas(docx, bookmark,ShowKey);
        for(int i=0;i<rl.size();i++)
            returnlist.add(rl.get(i));
        //替换表格里面的变量
        rl=replaceInTable(docx, bookmark,ShowKey);
        for(int i=0;i<rl.size();i++)
            returnlist.add(rl.get(i));
        return returnlist;
    }
    /**
     * 替换段落里面的变量
     * @param doc 要替换的文档
     * @param bookmark 参数
     */
    public static ArrayList<String> replaceInParas(XWPFDocument doc, HashMap<String, String> bookmark,boolean ShowKey)
    {
        ArrayList<String> returnlist= new ArrayList<String>();
        Iterator<XWPFParagraph> iterator = doc.getParagraphsIterator();
        XWPFParagraph para;
        while (iterator.hasNext()) {
            para = iterator.next();
            ArrayList<String> rl=replaceInPara(para, bookmark,ShowKey);
            for(int i=0;i<rl.size();i++)
                returnlist.add(rl.get(i));
        }
        return returnlist;
    }
    /**
     * 替换表格里面的变量
     * @param doc 要替换的文档
     * @param bookmark 参数
     */
    public static ArrayList<String> replaceInTable(XWPFDocument doc, HashMap<String, String> bookmark,boolean ShowKey) {
        ArrayList<String> returnlist= new ArrayList<String>();
        Iterator<XWPFTable> iterator = doc.getTablesIterator();
        XWPFTable table;
        List<XWPFTableRow> rows;
        List<XWPFTableCell> cells;
        List<XWPFParagraph> paras;
        while (iterator.hasNext()) {
            table = iterator.next();
            rows = table.getRows();
            for (XWPFTableRow row : rows) {
                cells = row.getTableCells();
                for (XWPFTableCell cell : cells) {
                    paras = cell.getParagraphs();
                    for (XWPFParagraph para : paras) {
                        ArrayList<String> rl=replaceInPara(para, bookmark,ShowKey);
                        for(int i=0;i<rl.size();i++)
                            returnlist.add(rl.get(i));
                    }
                }
            }
        }
        return returnlist;
    }
    /**
     * 替换段落里面的变量
     * @param para 要替换的段落
     * @param bookmark 参数
     */
    public static ArrayList<String> replaceInPara(XWPFParagraph para, HashMap<String, String> bookmark,boolean ShowKey)
    {
        ArrayList<String> returnlist= new ArrayList<String>();

        List<XWPFRun> runs;
        Matcher matcher;
        String pkey,pval;
        if (matcherL(para.getParagraphText()).find())
        {
            runs = para.getRuns();
            for (int i=0; i<runs.size(); i++)
            {
                pval = runs.get(i).toString();
                matcher = matcherL(pval);
                if (matcher.find())
                {
                    pkey="Error";
                    while ((matcher = matcherL(pval)).find())//多个标签在1个run中的时候，就会用到这里。
                    {
                        pkey=matcher.group(1);
                        returnlist.add(pkey);
                        if(bookmark==null)
                            pval=matcher.replaceFirst("");
                        else
                            pval =matcher.replaceFirst(String.valueOf(bookmark.get(pkey)));
                    }
                    //直接调用runs.get(i).setText(runText);方法设置文本时，在底层会重新创建一个XWPFRun，把文本附加在当前文本后面，
                    //所以我们不能直接设值，需要先删除当前run,然后再自己手动插入一个新的run。
                    //para.removeRun(i);
                    //para.insertNewRun(i).setText(runText);//邢志云2018-12-03 01:53:44但用这种方式无法保持原来的文字格式
                    if(ShowKey)
                        runs.get(i).setText("▶"+pkey+"◀", 0);
                    else if(pval!=null && !"null".equals(pval) && bookmark!=null)
                    {
                        runs.get(i).setText(pval, 0);//这个完美
                    }
                }
            }
        }
        return returnlist;
    }
    /**
     * 正则匹配字符串
     * @param str
     * @return
     */
    public static Matcher matcherL(String str) {
        Pattern pattern = Pattern.compile("\\$\\{(.+?)\\}", Pattern.CASE_INSENSITIVE);
        Matcher matcher = pattern.matcher(str);
        return matcher;
    }
    //■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■
    public static void main(String[] args) throws Exception
    {

    }
}


