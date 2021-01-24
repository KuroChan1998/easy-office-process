package com.jzy.office.word;

import com.jzy.office.exception.InvalidFileTypeException;
import com.jzy.office.matcher.LabelMatcher;
import lombok.ToString;
import org.apache.commons.lang3.StringUtils;
import org.apache.poi.xwpf.usermodel.*;

import java.io.*;
import java.util.*;

/**
 * Word 2007版本的包装类，支持.docx文件的解析
 *
 * @author JinZhiyun
 * @version 1.0 2020/11/28
 */
@ToString(callSuper = true)
public class Word2007 extends AbstractWord<XWPFDocument> {
    private static final long serialVersionUID = -20071L;

    /**
     * 由输入文件路径构造word对象
     *
     * @param inputFile 输入文件路径
     * @throws IOException
     * @throws InvalidFileTypeException
     */
    public Word2007(String inputFile) throws IOException, InvalidFileTypeException {
        this(new File(inputFile));
    }

    /**
     * 由一个File构造word对象
     *
     * @param file 输入文件对象
     * @throws IOException
     * @throws InvalidFileTypeException
     */
    public Word2007(File file) throws IOException, InvalidFileTypeException {
        WordVersionEnum version = WordVersionEnum.getVersion(file.getAbsolutePath());
        if (WordVersionEnum.VERSION_2007.equals(version)) {
            document = new XWPFDocument(new FileInputStream(file));
            this.version = version;
            this.inputFilePath = file.getAbsolutePath();
        } else {
            throw new InvalidFileTypeException("错误的文件类型！" + Word2007.class + "及其子类仅支持文件格式：" + WordVersionEnum.VERSION_2007.getSuffix());
        }
    }

    /**
     * 由一个输入流构造Word2007对象
     *
     * @param inputStream 输入流对象
     * @param version     word版本的枚举对象
     * @throws IOException
     * @throws InvalidFileTypeException
     */
    public Word2007(InputStream inputStream, WordVersionEnum version) throws IOException, InvalidFileTypeException {
        if (WordVersionEnum.VERSION_2007.equals(version)) {
            this.version = WordVersionEnum.VERSION_2007;
            document = new XWPFDocument(inputStream);
        } else {
            throw new InvalidFileTypeException("错误的文件类型！" + Word2007.class + "及其子类仅支持文件格式：" + WordVersionEnum.VERSION_2007.getSuffix());
        }
    }

    /**
     * 由一个文档构造word对象
     *
     * @param document 文档对象
     */
    public Word2007(XWPFDocument document) {
        this.document = document;
    }

    @Override
    public void save(OutputStream outputStream) throws IOException {
        document.write(outputStream);
    }

    /**
     * 获取第pos+1个段落对象
     *
     * @param pos 段落索引
     * @return
     */
    public XWPFParagraph getParagraph(int pos) {
        return document.getParagraphArray(pos);
    }

    /**
     * 该接口返回所有段落数，但不含表格对象中的段落
     *
     * @return
     */
    @Override
    public int getParagraphNum() {
        int num = 0;
        Iterator<XWPFParagraph> iterator = document.getParagraphsIterator();
        while (iterator.hasNext()) {
            //遍历所有段落
            iterator.next();
            num++;
        }
        return num;
    }

    @Override
    public String readParagraph(int pos) {
        XWPFParagraph paragraph = getParagraph(pos);
        if (paragraph == null) {
            return "";
        }
        return paragraph.getText();
    }

    @Override
    public String readParagraphs() {
        StringBuilder result = new StringBuilder();
        Iterator<XWPFParagraph> iterator = document.getParagraphsIterator();
        XWPFParagraph para;
        while (iterator.hasNext()) {
            //遍历所有段落
            para = iterator.next();
            result.append(para.getText());
        }
        return result.toString();
    }

    @Override
    public List<String> readParagraphsToList() {
        List<String> str = new ArrayList<>();
        Iterator<XWPFParagraph> iterator = document.getParagraphsIterator();
        XWPFParagraph para;
        while (iterator.hasNext()) {
            //遍历所有段落
            para = iterator.next();
            str.add(para.getText());
        }
        return str;
    }


    @Override
    public HashMap<String, String> replaceInParas(HashMap<String, String> bookmark) {
        HashMap<String, String> replacedBookmark = new HashMap<>();
        Iterator<XWPFParagraph> iterator = document.getParagraphsIterator();
        XWPFParagraph para;
        while (iterator.hasNext()) {
            //遍历所有段落
            para = iterator.next();
            //对每个段落进行替换
            replacedBookmark.putAll(replaceInPara(para, bookmark));
        }
        return replacedBookmark;
    }

    @Override
    public HashMap<String, String> replaceInPara(int pos, HashMap<String, String> bookmark) {
        return replaceInPara(getParagraph(pos), bookmark);
    }

    /**
     * 根据替换书签集的内容，替换当前指定段落文本的对应标签。举例可参见 {@link AbstractWord#replaceInPara(int, HashMap)}
     * 如果需要使用标签匹配器，请参见 {@link DefaultWord2007#replaceInParaUsingLabelMatcher(XWPFParagraph, HashMap, LabelMatcher)}
     * <p>
     * 需要注意的是：由于读取和修改段落文本时通过run的方式，那么${标签1}可能由于输入格式的原因可能读在两个run中，
     * 从而导致替换失败。解决方法为：保证${标签1}解析到同一个run中，尽量避免制作word模板时纯手动输入一个标签，
     * 可以通过现在另外的txt中写好${标签1}，再复制到word指定位置解决。
     *
     * @param para     指定段落
     * @param bookmark 替换书签集
     * @return 返回在bookmark和段落文本中共有的被正确替换掉的标签集合
     */
    public HashMap<String, String> replaceInPara(XWPFParagraph para, HashMap<String, String> bookmark) {
        HashMap<String, String> replacedBookmark = new HashMap<>();
        if (para == null || bookmark == null || bookmark.size() == 0) {
            return replacedBookmark;
        }
        //获得当前段落的各种run
        List<XWPFRun> runs = para.getRuns();
        for (int i = 0; i < runs.size(); i++) {
            XWPFRun run = runs.get(i);
            String runString = run.toString();
            if (StringUtils.isNotEmpty(runString)) {
                //使用bookmark对其对应标签替换文本内容
                for (Map.Entry<String, String> bm : bookmark.entrySet()) {
                    //遍历所有书签
                    String labelKey = bm.getKey();
                    String replacement = bm.getValue();
                    if (runString.contains(labelKey)) {
                        //如果runString中含有当前键，替换
                        runString = runString.replaceAll(labelKey, replacement);
                        //将成功被替换掉的标签添加到输出结果集
                        replacedBookmark.put(labelKey, replacement);
                    }
                }
                /*
                 * 直接调用runs.get(i).setText(runText);方法设置文本时，在底层会重新创建一个XWPFRun，把文本附加在当前文本后面，
                 * 所以我们不能直接设值，需要先删除当前run,然后再自己手动插入一个新的run。
                 * 或者使用run.setText(runStringReplaced, 0);
                 */
                //para.removeRun(i);
                //但用这种方式无法保持原来的文字格式
                //para.insertNewRun(i).setText(runText);
                run.setText(runString, 0);/**/
            }
        }
        return replacedBookmark;
    }

    @Override
    public int getTableNum() {
        int num = 0;
        Iterator<XWPFTable> iterator = document.getTablesIterator();
        while (iterator.hasNext()) {
            //遍历每个表格
            iterator.next();
            //对每个表格进行替换
            num++;
        }
        return num;
    }

    /**
     * 获取第pos+1个表格对象
     *
     * @param pos 表格所有
     * @return
     */
    public XWPFTable getTable(int pos) {
        return document.getTableArray(pos);
    }

    @Override
    List<List<List<String>>> readTable0(int pos) {
        List<List<List<String>>> cellParaStrings = new ArrayList<>();
        //获得对应位置表格
        XWPFTable table = getTable(pos);
        List<XWPFTableRow> rows = table.getRows();
        List<XWPFTableCell> cells;
        for (XWPFTableRow row : rows) {
            //遍历所有行的所有单元格
            cells = row.getTableCells();
            cellParaStrings.add(cellsToStringList(cells));
        }
        return cellParaStrings;
    }

    @Override
    List<List<String>> readTableRow0(int tablePos, int rowPos) {
        //获得对应位置表格
        XWPFTable table = getTable(tablePos);
        List<XWPFTableRow> rows = table.getRows();
        //获得对应行
        XWPFTableRow row = rows.get(rowPos);
        List<XWPFTableCell> cells = row.getTableCells();
        return cellsToStringList(cells);
    }


    @Override
    List<List<String>> readTableColumn0(int tablePos, int columnPos) {
        List<List<String>> columnParaStrings = new ArrayList<>();
        //获得对应位置表格
        XWPFTable table = getTable(tablePos);
        List<XWPFTableRow> rows = table.getRows();
        for (XWPFTableRow row : rows) {
            //遍历所有行，获得第columnPos+1列的单元格
            List<XWPFTableCell> cells = row.getTableCells();
            XWPFTableCell cell = cells.get(columnPos);
            columnParaStrings.add(cellToStringList(cell));
        }
        return columnParaStrings;
    }

    @Override
    List<String> readTable0(int tablePos, int rowPos, int columnPos) {
        //获得对应位置表格
        XWPFTable table = getTable(tablePos);
        List<XWPFTableRow> rows = table.getRows();
        //获得对应行
        XWPFTableRow row = rows.get(rowPos);
        List<XWPFTableCell> cells = row.getTableCells();
        //获得对应单元格
        XWPFTableCell cell = cells.get(columnPos);
        List<XWPFParagraph> paras = cell.getParagraphs();
        return cellToStringList(cell);
    }

    private List<List<String>> cellsToStringList(List<XWPFTableCell> cells) {
        List<List<String>> list = new ArrayList<>();
        for (XWPFTableCell cell : cells) {
            list.add(cellToStringList(cell));
        }
        return list;
    }

    private List<String> cellToStringList(XWPFTableCell cell) {
        List<String> cellParaStrings = new ArrayList<>();
        List<XWPFParagraph> paras = cell.getParagraphs();
        for (XWPFParagraph para : paras) {
            //遍历单元格的每个段落
            cellParaStrings.add(para.getText());
        }
        return cellParaStrings;
    }

    @Override
    public HashMap<String, String> replaceInTables(HashMap<String, String> bookmark) {
        HashMap<String, String> replacedBookmark = new HashMap<>();
        Iterator<XWPFTable> iterator = document.getTablesIterator();
        XWPFTable table;
        while (iterator.hasNext()) {
            //遍历每个表格
            table = iterator.next();
            //对每个表格进行替换
            replacedBookmark.putAll(replaceInTable(table, bookmark));
        }
        return replacedBookmark;
    }

    /**
     * 根据替换书签集的内容，替换第pos+1个表格中的对应标签。
     *
     * @param pos      表格索引
     * @param bookmark 替换书签集
     * @return 返回在bookmark和表格文本中共有的被正确替换掉的标签集合
     */
    @Override
    public HashMap<String, String> replaceInTable(int pos, HashMap<String, String> bookmark) {
        return replaceInTable(getTable(pos), bookmark);
    }

    /**
     * 根据替换书签集的内容，替换当前指定表格文本的对应标签。举例可参见 {@link AbstractWord#replaceInTable(int, HashMap)}
     * 如果需要使用标签匹配器，请参见 {@link DefaultWord2007#replaceInTableUsingLabelMatcher(XWPFTable, HashMap, LabelMatcher)}
     * <p>
     * 注意：这里会遍历当前表格的每个单元格文本，每个文本都是多个para的集合，并交给{@link Word2007#replaceInPara(XWPFParagraph, HashMap)}
     *
     * @param table    指定表格
     * @param bookmark 替换书签集
     * @return 返回在bookmark和表格文本中共有的被正确替换掉的标签集合
     */
    public HashMap<String, String> replaceInTable(XWPFTable table, HashMap<String, String> bookmark) {
        HashMap<String, String> replacedBookmark = new HashMap<>();
        if (table == null || bookmark == null || bookmark.size() == 0) {
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
                    replacedBookmark.putAll(replaceInPara(para, bookmark));
                }
            }
        }
        return replacedBookmark;
    }


    /**
     * 根据替换书签集的内容，替换当前文档所有段落和表格的对应标签。
     *
     * @param bookmark 替换书签集
     * @return 返回在bookmark和段落表格文本中共有的被正确替换掉的标签集合
     */
    @Override
    public HashMap<String, String> replaceInAll(HashMap<String, String> bookmark) {
        HashMap<String, String> replacedBookmark = replaceInParas(bookmark);
        replacedBookmark.putAll(replaceInTables(bookmark));
        return replacedBookmark;
    }
}


