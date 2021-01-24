package com.jzy.office.word;

import com.jzy.office.exception.InvalidFileTypeException;
import lombok.Getter;
import lombok.ToString;
import org.apache.poi.hwpf.HWPFDocument;
import org.apache.poi.hwpf.usermodel.*;

import java.io.*;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

/**
 * Word 2003版本的包装类，支持.doc，.wps文件的解析
 * 由于结构以及poi提供接口的不同，该类对于段落的解析，会把表格对象的段落也作为文档段落
 *
 * @author JinZhiyun
 * @version 1.0 2020/11/28
 */
@ToString(callSuper = true)
public class Word2003 extends AbstractWord<HWPFDocument> implements Serializable {
    private static final long serialVersionUID = -20031L;

    /**
     * 文档读取范围
     */
    @Getter
    protected Range range;


    /**
     * 由输入文件路径构造word对象
     *
     * @param inputFile 输入文件路径
     * @throws IOException
     * @throws InvalidFileTypeException
     */
    public Word2003(String inputFile) throws IOException, InvalidFileTypeException {
        this(new File(inputFile));
    }

    /**
     * 由一个File构造word对象
     *
     * @param file 输入文件对象
     * @throws IOException
     * @throws InvalidFileTypeException
     */
    public Word2003(File file) throws IOException, InvalidFileTypeException {
        this(new FileInputStream(file), WordVersionEnum.getVersion(file.getAbsolutePath()));
        this.inputFilePath = file.getAbsolutePath();
    }

    /**
     * 由一个输入流和版本枚举对象构造word对象
     *
     * @param inputStream 输入流对象
     * @param version     word版本的枚举对象
     * @throws IOException
     * @throws InvalidFileTypeException
     */
    public Word2003(InputStream inputStream, WordVersionEnum version) throws IOException, InvalidFileTypeException {
        if (WordVersionEnum.VERSION_2003.equals(version)
                || WordVersionEnum.VERSION_WPS.equals(version)) {
            this.version = version;
            this.document = new HWPFDocument(inputStream);
            this.range = document.getRange();
        } else {
            throw new InvalidFileTypeException("错误的文件类型！" + Word2003.class + "及其子类仅支持文件格式：" + WordVersionEnum.VERSION_2003.getSuffix()
                    + ", " + WordVersionEnum.VERSION_WPS.getSuffix());
        }
    }

    @Override
    public void save(OutputStream outputStream) throws IOException {
        document.write(outputStream);
    }

    /**
     * 由一个文档构造word对象
     *
     * @param document 文档对象
     */
    public Word2003(HWPFDocument document) {
        this.document = document;
    }


    /**
     * 获得第pos+1段的段落对象
     *
     * @param pos 段索引
     * @return
     */
    public Paragraph getParagraph(int pos) {
        return range.getParagraph(pos);
    }


    /**
     * 该接口返回的段落数包含了表格中的段落
     *
     * @return
     */
    @Override
    public int getParagraphNum() {
        return range.numParagraphs();
    }

    @Override
    public String readParagraph(int pos) {
        return getParagraph(pos).text();
    }

    @Override
    public String readParagraphs() {
        StringBuilder s = new StringBuilder();
        //读取word文本内容
        for (int i = 0; i < getParagraphNum(); i++) {
            s.append(readParagraph(i));
        }
        return s.toString();
    }

    @Override
    public List<String> readParagraphsToList() {
        List<String> paras = new ArrayList<>();
        //读取word文本内容
        for (int i = 0; i < getParagraphNum(); i++) {
            paras.add(readParagraph(i));
        }
        return paras;
    }

    /**
     * 必须要注意的是：poi对word 2003段落的处理，不会区分所有表格中的段落。
     * 该接口会替换当前文档包括表格在内的所有段落中的对应文本
     *
     * @param bookmark 替换书签集
     * @return 返回在bookmark和段落文本中共有的被正确替换掉的标签集合
     */
    @Override
    public HashMap<String, String> replaceInParas(HashMap<String, String> bookmark) {
        HashMap<String, String> replacedBookmark = new HashMap<>();
        for (int i = 0; i < getParagraphNum(); i++) {
            replacedBookmark.putAll(replaceInPara(i, bookmark));
        }
        return replacedBookmark;
    }

    @Override
    public HashMap<String, String> replaceInPara(int pos, HashMap<String, String> bookmark) {
        return replaceInPara(getParagraph(pos), bookmark);
    }

    /**
     * 根据替换书签集的内容，替换当前指定段落文本的对应标签。举例可参见 {@link AbstractWord#replaceInPara(int, HashMap)}
     * 这里通过直接调用poi内置方法{@link Paragraph#replaceText(String, String)}实现替换
     *
     * @param para     指定段落
     * @param bookmark 替换书签集
     * @return 返回在bookmark和段落文本中共有的被正确替换掉的标签集合
     */
    public HashMap<String, String> replaceInPara(Paragraph para, HashMap<String, String> bookmark) {
        HashMap<String, String> replacedBookmark = new HashMap<>();
        if (para == null || bookmark == null || bookmark.size() == 0) {
            return replacedBookmark;
        }
        String text = para.text();

        //使用bookmark对其对应标签替换文本内容
        for (Map.Entry<String, String> bm : bookmark.entrySet()) {
            //遍历所有书签
            String labelKey = bm.getKey();
            String replacement = bm.getValue();
            if (text.contains(labelKey)) {
                //如果text中含有当前键，替换
                para.replaceText(labelKey, replacement);
                //将成功被替换掉的标签添加到输出结果集
                replacedBookmark.put(labelKey, replacement);
            }
        }
        return replacedBookmark;
    }

    @Override
    public int getTableNum() {
        int num = 0;
        TableIterator it = new TableIterator(range);
        while (it.hasNext()) {
            //迭代文档中的表格
            it.next();
            num++;
        }
        return num;
    }


    /**
     * 获得第pos+1个表格对象
     *
     * @param pos 表格索引
     * @return
     */
    public Table getTable(int pos) {
        int num = 0;
        TableIterator it = new TableIterator(range);
        Table table;
        while (it.hasNext()) {
            //迭代文档中的表格
            table = it.next();
            if (num == pos) {
                return table;
            }
            num++;
        }
        return null;
    }

    @Override
    List<List<List<String>>> readTable0(int pos) {
        List<List<List<String>>> cellParaStrings = new ArrayList<>();
        //获得对应位置表格
        Table table = getTable(pos);
        //迭代行，默认从0开始
        for (int i = 0; i < table.numRows(); i++) {
            TableRow tr = table.getRow(i);
            List<List<String>> rowParaStrings = new ArrayList<>();
            for (int j = 0; j < tr.numCells(); j++) {
                //遍历每个单元
                TableCell td = tr.getCell(j);
                rowParaStrings.add(cellToStringList(td));
            }
            cellParaStrings.add(rowParaStrings);
        }
        return cellParaStrings;
    }

    @Override
    List<List<String>> readTableRow0(int tablePos, int rowPos) {
        List<List<String>> rowParaStrings = new ArrayList<>();
        //获得对应位置表格
        Table table = getTable(tablePos);
        //获得对应行
        TableRow tr = table.getRow(rowPos);
        for (int j = 0; j < tr.numCells(); j++) {
            //遍历每个单元
            TableCell td = tr.getCell(j);
            rowParaStrings.add(cellToStringList(td));
        }
        return rowParaStrings;
    }

    @Override
    List<List<String>> readTableColumn0(int tablePos, int columnPos) {
        List<List<String>> columnParaStrings = new ArrayList<>();
        //获得对应位置表格
        Table table = getTable(tablePos);
        //迭代行，默认从0开始
        for (int i = 0; i < table.numRows(); i++) {
            TableRow tr = table.getRow(i);
            //获得当前行对应列单元格
            TableCell td = tr.getCell(columnPos);
            columnParaStrings.add(cellToStringList(td));
        }
        return columnParaStrings;
    }

    @Override
    List<String> readTable0(int tablePos, int rowPos, int columnPos) {
        //获得表格
        Table table = getTable(tablePos);
        //获得特定行
        TableRow tr = table.getRow(rowPos);
        //获得特定单元格
        TableCell td = tr.getCell(columnPos);
        return cellToStringList(td);
    }

    private List<List<String>> cellsToStringList(List<TableCell> cells) {
        List<List<String>> list = new ArrayList<>();
        for (TableCell cell : cells) {
            list.add(cellToStringList(cell));
        }
        return list;
    }

    private List<String> cellToStringList(TableCell cell) {
        List<String> cellParaStrings = new ArrayList<>();
        for (int k = 0; k < cell.numParagraphs(); k++) {
            //遍历每个单元的每个段落
            Paragraph para = cell.getParagraph(k);
            cellParaStrings.add(para.text());
        }
        return cellParaStrings;
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
     * 根据替换书签集的内容，替换当前指定段落文本的对应标签。举例可参见 {@link AbstractWord#replaceInTable(int, HashMap)}
     * <p>
     * 注意：这里会遍历当前表格的每个单元格文本，每个文本都是多个para的集合，并交给{@link Word2003#replaceInPara(Paragraph, HashMap)}
     *
     * @param table    指定表格
     * @param bookmark 替换书签集
     * @return 返回在bookmark和表格文本中共有的被正确替换掉的标签集合
     */
    public HashMap<String, String> replaceInTable(Table table, HashMap<String, String> bookmark) {
        HashMap<String, String> replacedBookmark = new HashMap<>();
        if (table == null || bookmark == null || bookmark.size() == 0) {
            return replacedBookmark;
        }
        //迭代行，默认从0开始
        for (int i = 0; i < table.numRows(); i++) {
            TableRow tr = table.getRow(i);
            //迭代列，默认从0开始
            for (int j = 0; j < tr.numCells(); j++) {
                //遍历每个单元
                TableCell td = tr.getCell(j);
                for (int k = 0; k < td.numParagraphs(); k++) {
                    //遍历每个单元的每个段落
                    Paragraph para = td.getParagraph(k);
                    replacedBookmark.putAll(replaceInPara(para, bookmark));
                }
            }
        }
        return replacedBookmark;
    }

    @Override
    public HashMap<String, String> replaceInTables(HashMap<String, String> bookmark) {
        HashMap<String, String> replacedBookmark = new HashMap<>();

        TableIterator it = new TableIterator(range);
        Table table;
        while (it.hasNext()) {
            //迭代文档中的表格
            table = it.next();
            //替换
            replacedBookmark.putAll(replaceInTable(table, bookmark));
        }
        return replacedBookmark;
    }

    /**
     * 必须要注意的是：poi对word 2003段落的处理，不会区分所有表格中的段落
     * 因此直接调用{@link Word2003#replaceInParas(HashMap)}即可达到效果
     *
     * @param bookmark 替换书签集
     * @return 返回在bookmark和当前文档中共有的被正确替换掉的标签集合
     */
    @Override
    public HashMap<String, String> replaceInAll(HashMap<String, String> bookmark) {
        return replaceInParas(bookmark);
    }

}


