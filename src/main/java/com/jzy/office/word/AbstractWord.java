package com.jzy.office.word;

import com.jzy.office.AbstractOffice;
import lombok.Getter;
import org.apache.poi.xwpf.usermodel.XWPFParagraph;

import java.io.Closeable;
import java.io.IOException;
import java.io.OutputStream;
import java.io.Serializable;
import java.util.HashMap;
import java.util.List;

/**
 * Word 包装类基类
 *
 * @author JinZhiyun
 * @version 1.0 2020/11/28
 */
public abstract class AbstractWord<D extends Closeable> extends AbstractOffice implements Serializable {
    private static final long serialVersionUID = -1L;

    /**
     * word版本枚举对象
     */
    @Getter
    protected WordVersionEnum version;

    /**
     * 文档对象
     */
    @Getter
    protected D document;

    @Override
    public void close() throws IOException {
        super.close();
        document.close();
    }

    /**
     * 将当前文档对象的修改保存到输出流
     *
     * @param outputStream 输出流
     * @throws IOException
     */
    @Override
    public abstract void save(OutputStream outputStream) throws IOException;

    /**
     * 获得总段落数量
     *
     * @return
     */
    public abstract int getParagraphNum();

    /**
     * 返回第pos+1段的文本
     *
     * @param pos 段落索引
     * @return
     */
    public abstract String readParagraph(int pos);

    /**
     * 返回当前文档所有段落的文本
     *
     * @return 返回一整个字符串
     */
    public abstract String readParagraphs();

    /**
     * 返回当前文档所有段落的文本，依次装入list中
     * 由于结构以及poi提供接口的不同：
     * 对word 2003子类实现对于段落的解析时，会把表格对象的段落也作为文档段落
     * 对word 2007子类实现对于段落的解析时，不会把表格对象的段落作为文档段落
     *
     * @return 返回所有段落的文本的list
     */
    public abstract List<String> readParagraphsToList();

    /**
     * 根据替换书签集的内容，替换当前指定段落文本的对应标签。
     * 更多细节的解释参见 {@link Word2007#replaceInPara(XWPFParagraph, HashMap)}
     *
     * <p>
     * 例子❶，第1段的文本为：习近平${label1}在讲话中${label2}强调
     * ...
     * Map<String, String> bookmark = new HashMap<>();
     * bookmark.put("${label0}", "0000");
     * bookmark.put("${label1}", "1111");
     * replaceInPara(0, bookmark);
     * ...
     * 替换后的第1段文本为：习近平1111在讲话中${label2}强调
     * <p>
     *
     * @param pos      指定段落索引
     * @param bookmark 替换书签集
     * @return 返回在bookmark和段落文本中共有的被正确替换掉的标签集合
     */
    public abstract HashMap<String, String> replaceInPara(int pos, HashMap<String, String> bookmark);

    /**
     * 根据替换书签集的内容，替换当前文档所有段落文本的对应标签。
     *
     * @param bookmark 替换书签集
     * @return 返回在bookmark和段落文本中共有的被正确替换掉的标签集合
     */
    public abstract HashMap<String, String> replaceInParas(HashMap<String, String> bookmark);

    /**
     * 检查给定的表格索引是否大于当前文档的最大表格索引值。如果大于，抛出异常
     *
     * @param tablePos 给定的表格索引
     */
    protected void checkTableRange(int tablePos) {
        int maxPos = getTableNum() - 1;
        if (maxPos < tablePos) {
            throw new IndexOutOfBoundsException("表格索引不存在。输入：" + tablePos + "，最大：" + maxPos);
        }
    }

    /**
     * 获得总表格数量
     *
     * @return
     */
    public abstract int getTableNum();

    /**
     * 返回第pos+1个表格的文本。
     * 按表格二维结构，返回对应结构的二维数组。每个二维数组元素都是对应单元格段落的list
     * -------------------------------------------------------------
     * |   List<String>    |   List<String>    |   List<String>    |
     * -------------------------------------------------------------
     * |   List<String>    |   List<String>    |   List<String>    |
     * -------------------------------------------------------------
     * ...
     *
     * @param pos 表格索引
     * @return 表格二维数组
     */
    public List<List<List<String>>> readTable(int pos) {
        checkTableRange(pos);
        return readTable0(pos);
    }

    abstract List<List<List<String>>> readTable0(int pos);

    /**
     * 返回第pos+1个表格的第rowPos+1行的所有元素集合。每个元素都是对应单元格段落的list
     *
     * @param tablePos 表格索引
     * @param rowPos   表格行索引
     * @return
     */
    public List<List<String>> readTableRow(int tablePos, int rowPos) {
        checkTableRange(tablePos);
        return readTableRow0(tablePos, rowPos);
    }

    abstract List<List<String>> readTableRow0(int tablePos, int rowPos);

    /**
     * 返回第pos+1个表格的第columnPos+1列的所有元素集合。每个元素都是对应单元格段落的list
     *
     * @param tablePos  表格索引
     * @param columnPos 表格列索引
     * @return
     */
    public List<List<String>> readTableColumn(int tablePos, int columnPos) {
        checkTableRange(tablePos);
        return readTableColumn0(tablePos, columnPos);
    }

    abstract List<List<String>> readTableColumn0(int tablePos, int columnPos);

    public List<String> readTable(int tablePos, int rowPos, int columnPos) {
        checkTableRange(tablePos);
        return readTable0(tablePos, rowPos, columnPos);
    }

    /**
     * 返回第pos+1个表格的第rowPos+1行第columnPos+1列单元格的元素。该元素是对应单元格段落的list
     *
     * @param tablePos  表格索引
     * @param rowPos    表格行索引
     * @param columnPos 表格列索引
     * @return 表格对应单元格的段落的list
     */
    abstract List<String> readTable0(int tablePos, int rowPos, int columnPos);

    /**
     * 根据替换书签集的内容，替换当前指定表格文本的对应标签。
     * 如果文本中的标签bookmark中没有，不做任何处理
     * 如果bookmark中的标签当前文本中没有，不做任何处理
     * 返回在bookmark和表格文本中共有的被正确替换掉的标签集合
     *
     * <p>
     * 例子❶，如当前表格为：
     * -------------------------------------------
     * |123	        |23	            |aaa${lab}a  |
     * -------------------------------------------
     * |asdad	    |dada${table}d  |            |
     * -------------------------------------------
     * ...
     * Map<String, String> bookmark = new HashMap<>();
     * bookmark.put("${label0}", "0000");
     * bookmark.put("${label1}", "1111");
     * replaceInTable(table, bookmark);
     * ...
     * 替换后的表格为：
     * -------------------------------------------
     * |123	        |23	            |aaa${lab}a   |
     * -------------------------------------------
     * |asdad	    |dada0000d      |            |
     * -------------------------------------------
     * <p>
     *
     * @param pos      指定表格索引
     * @param bookmark 替换书签集
     * @return 返回在bookmark和表格文本中共有的被正确替换掉的标签集合
     */
    public abstract HashMap<String, String> replaceInTable(int pos, HashMap<String, String> bookmark);

    /**
     * 根据替换书签集的内容，替换当前文档所有表格的对应标签。
     *
     * @param bookmark 替换书签集
     * @return 返回在bookmark和表格文本中共有的被正确替换掉的标签集合
     */
    public abstract HashMap<String, String> replaceInTables(HashMap<String, String> bookmark);

    /**
     * 根据替换书签集的内容，替换当前文档所有文本中的对应标签。详情参见子类实现
     *
     * @param bookmark 替换书签集
     * @return 返回在bookmark被正确替换掉的标签集合
     */
    public abstract HashMap<String, String> replaceInAll(HashMap<String, String> bookmark);

    @Override
    public String toString() {
        return "共有 " + getParagraphNum() + "个段落！" + getTableNum() + "个表格！";
    }
}


