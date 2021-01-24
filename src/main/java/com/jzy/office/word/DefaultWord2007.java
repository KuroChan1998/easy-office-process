package com.jzy.office.word;

import com.jzy.office.exception.InvalidFileTypeException;
import com.jzy.office.matcher.LabelMatcher;
import com.jzy.office.matcher.LabelMatchers;
import lombok.ToString;
import org.apache.commons.lang3.StringUtils;
import org.apache.poi.xwpf.usermodel.*;

import java.io.File;
import java.io.IOException;
import java.io.InputStream;
import java.util.HashMap;
import java.util.Iterator;
import java.util.List;

/**
 * @ClassName DefaultWord2007
 * @Author JinZhiyun
 * @Description 默认word 2007文档处理类。一般普通的处理word 2007文档可以继承此类
 * 该类支持使用标签匹配器
 * @Date 2020/11/28 19:20
 * @Version 1.0
 **/
@ToString(callSuper = true)
public class DefaultWord2007 extends Word2007 {
    private static final long serialVersionUID = -20072L;

    public DefaultWord2007(String inputFile) throws IOException, InvalidFileTypeException {
        super(inputFile);
    }

    public DefaultWord2007(File file) throws IOException, InvalidFileTypeException {
        super(file);
    }

    public DefaultWord2007(InputStream inputStream, WordVersionEnum version) throws IOException, InvalidFileTypeException {
        super(inputStream, version);
    }

    public DefaultWord2007(XWPFDocument document) {
        super(document);
    }

    /**
     * 使用默认标签匹配器，根据替换书签集的内容，替换当前文档所有段落和表格的对应标签。
     * 默认标签匹配器匹配标签格式为：${标签1}。参见{@link LabelMatchers#DEFAULT_LABEL_MATCHER}
     *
     * @param bookmark 替换书签集
     * @return 返回在bookmark和段落表格文本中共有的被正确替换掉的标签集合
     */
    public HashMap<String, String> replaceInAllUsingLabelMatcher(HashMap<String, String> bookmark) {
        return replaceInAllUsingLabelMatcher(bookmark, LabelMatchers.DEFAULT_LABEL_MATCHER);
    }

    /**
     * 使用给定的标签匹配器，根据替换书签集的内容，替换当前文档所有段落和表格的对应标签。
     *
     * @param bookmark 替换书签集
     * @return 返回在bookmark和段落表格文本中共有的被正确替换掉的标签集合
     */
    public HashMap<String, String> replaceInAllUsingLabelMatcher(HashMap<String, String> bookmark, LabelMatcher lMatcher) {
        HashMap<String, String> replacedBookmark = replaceInParasUsingLabelMatcher(bookmark, lMatcher);
        replacedBookmark.putAll(replaceInTablesUsingLabelMatcher(bookmark, lMatcher));
        return replacedBookmark;
    }


    /**
     * 使用默认标签匹配器，根据替换书签集的内容，替换当前文档所有表格的对应标签。
     * 默认标签匹配器匹配标签格式为：${标签1}。参见{@link LabelMatchers#DEFAULT_LABEL_MATCHER}
     *
     * @param bookmark 替换书签集
     * @return 返回在bookmark和表格文本中共有的被正确替换掉的标签集合
     */
    public HashMap<String, String> replaceInTablesUsingLabelMatcher(HashMap<String, String> bookmark) {
        return replaceInTablesUsingLabelMatcher(bookmark, LabelMatchers.DEFAULT_LABEL_MATCHER);
    }


    /**
     * 根据匹配器的规则，以及替换书签集的内容，替换当前文档所有表格的对应标签。
     *
     * @param bookmark 替换书签集
     * @param lMatcher 标签匹配器
     * @return 返回在bookmark和表格文本中共有的被正确替换掉的标签集合
     */
    public HashMap<String, String> replaceInTablesUsingLabelMatcher(HashMap<String, String> bookmark, LabelMatcher lMatcher) {
        HashMap<String, String> replacedBookmark = new HashMap<>();
        Iterator<XWPFTable> iterator = document.getTablesIterator();
        XWPFTable table;
        while (iterator.hasNext()) {
            //遍历每个表格
            table = iterator.next();
            //对每个表格进行替换
            replacedBookmark.putAll(replaceInTableUsingLabelMatcher(table, bookmark, lMatcher));
        }
        return replacedBookmark;
    }

    public HashMap<String, String> replaceInTableUsingLabelMatcher(int pos, HashMap<String, String> bookmark, LabelMatcher lMatcher) {
        return replaceInTableUsingLabelMatcher(getTable(pos), bookmark, lMatcher);
    }


    /**
     * 根据匹配器的规则，以及替换书签集的内容，替换当前指定表格文本的对应标签。
     * 如果文本中的标签bookmark中没有，将当前标签替换成空串
     * 如果bookmark中的标签当前文本中没有，不做任何处理（更多替换解释见：{@link LabelMatcher#replaceAllLabels(String, HashMap)}
     * 返回在bookmark和表格文本中共有的被正确替换掉的标签集合
     * <p>
     * 例子❶，如当前匹配标签为：${标签1}。当前表格为：
     * -------------------------------------------
     * |123	        |23	            |aaa${lab}a  |
     * -------------------------------------------
     * |asdad	    |dada${table}d  |            |
     * -------------------------------------------
     * ...
     * Map<String, String> bookmark = new HashMap<>();
     * bookmark.put("label0", "0000");
     * bookmark.put("label1", "1111");
     * replaceInTableUsingLabelMatcher(table, bookmark, LabelMatchers.DEFAULT_LABEL_MATCHER);
     * ...
     * 替换后的表格为：
     * -------------------------------------------
     * |123	        |23	            |aaaa        |
     * -------------------------------------------
     * |asdad	    |dada0000d      |            |
     * -------------------------------------------
     * <p>
     * 注意：这里会遍历当前表格的每个单元格文本，每个文本都是多个para的集合，并交给{@link DefaultWord2007#replaceInParaUsingLabelMatcher(XWPFParagraph, HashMap, LabelMatcher)}
     *
     * @param table    指定表格
     * @param bookmark 替换书签集
     * @param lMatcher 标签匹配器
     * @return 返回在bookmark和表格文本中共有的被正确替换掉的标签集合
     */
    public HashMap<String, String> replaceInTableUsingLabelMatcher(XWPFTable table, HashMap<String, String> bookmark, LabelMatcher lMatcher) {
        HashMap<String, String> replacedBookmark = new HashMap<>();
        if (table == null || bookmark == null || bookmark.size() == 0 || lMatcher == null) {
            return replacedBookmark;
        }
        List<XWPFTableRow> rows = table.getRows();
        List<XWPFTableCell> cells;
        List<XWPFParagraph> paras;
        for (XWPFTableRow row : rows) {
            cells = row.getTableCells();
            for (XWPFTableCell cell : cells) {
                paras = cell.getParagraphs();
                for (XWPFParagraph para : paras) {
                    //遍历每行每个单元的每个段落
                    replacedBookmark.putAll(replaceInParaUsingLabelMatcher(para, bookmark, lMatcher));
                }
            }
        }
        return replacedBookmark;
    }

    /**
     * 使用默认标签匹配器，根据替换书签集的内容，替换当前文档所有段落文本的对应标签。
     * 默认标签匹配器匹配标签格式为：${标签1}。参见{@link LabelMatchers#DEFAULT_LABEL_MATCHER}
     *
     * @param bookmark 替换书签集
     * @return 返回在bookmark和段落文本中共有的被正确替换掉的标签集合
     */
    public HashMap<String, String> replaceInParasUsingLabelMatcher(HashMap<String, String> bookmark) {
        return replaceInParasUsingLabelMatcher(bookmark, LabelMatchers.DEFAULT_LABEL_MATCHER);
    }

    /**
     * 根据匹配器的规则，以及替换书签集的内容，替换当前文档所有段落文本的对应标签。
     *
     * @param bookmark 替换书签集
     * @param lMatcher 标签匹配器
     * @return 返回在bookmark和段落文本中共有的被正确替换掉的标签集合
     */
    public HashMap<String, String> replaceInParasUsingLabelMatcher(HashMap<String, String> bookmark, LabelMatcher lMatcher) {
        HashMap<String, String> replacedBookmark = new HashMap<>();
        Iterator<XWPFParagraph> iterator = document.getParagraphsIterator();
        XWPFParagraph para;
        while (iterator.hasNext()) {
            //遍历所有段落
            para = iterator.next();
            //对每个段落进行替换
            replacedBookmark.putAll(replaceInParaUsingLabelMatcher(para, bookmark, lMatcher));
        }
        return replacedBookmark;
    }

    public HashMap<String, String> replaceInParaUsingLabelMatcher(int pos, HashMap<String, String> bookmark, LabelMatcher lMatcher) {
        return replaceInParaUsingLabelMatcher(getParagraph(pos), bookmark, lMatcher);
    }

    /**
     * 根据匹配器的规则，以及替换书签集的内容，替换当前指定段落文本的对应标签。
     * 如果文本中的标签bookmark中没有，将当前标签替换成空串
     * 如果bookmark中的标签当前文本中没有，不做任何处理（更多替换解释见：{@link LabelMatcher#replaceAllLabels(String, HashMap)}
     * 返回在bookmark和段落文本中共有的被正确替换掉的标签集合
     * <p>
     * 例子❶，如当前匹配标签为：${标签1}。当前段落文本为：习近平${label1}在讲话中${label2}强调
     * ...
     * Map<String, String> bookmark = new HashMap<>();
     * bookmark.put("label0", "0000");
     * bookmark.put("label1", "1111");
     * replaceInParaUsingLabelMatcher(para, bookmark, LabelMatchers.DEFAULT_LABEL_MATCHER);
     * ...
     * 替换后的段落文本为：习近平1111在讲话中强调
     * <p>
     * 需要注意的是：由于读取和修改段落文本时通过run的方式，那么${标签1}可能由于输入格式的原因可能读在两个run中，
     * 从而导致替换失败。解决方法为：保证${标签1}解析到同一个run中，尽量避免制作word模板时纯手动输入一个标签，
     * 可以通过现在另外的txt中写好${标签1}，再复制到word指定位置解决。
     *
     * @param para     指定段落
     * @param bookmark 替换书签集
     * @param lMatcher 标签匹配器
     * @return 返回在bookmark和段落文本中共有的被正确替换掉的标签集合
     */
    public HashMap<String, String> replaceInParaUsingLabelMatcher(XWPFParagraph para, HashMap<String, String> bookmark, LabelMatcher lMatcher) {
        HashMap<String, String> replacedBookmark = new HashMap<>();
        if (para == null || bookmark == null || bookmark.size() == 0 || lMatcher == null) {
            return replacedBookmark;
        }
        String paraText = para.getParagraphText();
        if (lMatcher.find(paraText)) {
            //如果标签匹配器匹配到了当前段落文本
            //获得当前段落的各种run
            List<XWPFRun> runs = para.getRuns();
            for (int i = 0; i < runs.size(); i++) {
                XWPFRun run = runs.get(i);
                String runString = run.toString();
                if (StringUtils.isNotEmpty(runString) && lMatcher.find(runString)) {
                    //如果标签匹配器匹配到了当前run的文本，使用bookmark对其对应标签替换文本内容
                    HashMap<String, String> replacedBookmarkWithOutput = lMatcher.replaceAllLabels(runString, bookmark);
                    //获得替换后的文本结果
                    String replacedText = replacedBookmarkWithOutput.get(LabelMatcher.OUTPUT);
                    /*
                     * 直接调用runs.get(i).setText(runText);方法设置文本时，在底层会重新创建一个XWPFRun，把文本附加在当前文本后面，
                     * 所以我们不能直接设值，需要先删除当前run,然后再自己手动插入一个新的run。
                     * 或者使用run.setText(runStringReplaced, 0);
                     */
                    //para.removeRun(i);
                    //但用这种方式无法保持原来的文字格式
                    //para.insertNewRun(i).setText(runText);
                    run.setText(replacedText, 0);

                    //移除替换后的文本结果，只将所有成功替换的标签集追加到输出的hashmap中
                    replacedBookmarkWithOutput.remove(LabelMatcher.OUTPUT);
                    replacedBookmark.putAll(replacedBookmarkWithOutput);

                }
            }
        }
        return replacedBookmark;
    }
}
