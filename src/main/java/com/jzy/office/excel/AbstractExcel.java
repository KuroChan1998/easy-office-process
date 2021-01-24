package com.jzy.office.excel;

import com.jzy.office.AbstractOffice;
import lombok.Getter;

import java.io.Closeable;
import java.io.IOException;
import java.io.OutputStream;
import java.io.Serializable;
import java.util.ArrayList;
import java.util.List;

/**
 * Excel 包装类基类
 *
 * @author JinZhiyun
 * @version 1.0 2020/01/27
 */
public abstract class AbstractExcel<W extends Closeable> extends AbstractOffice implements Serializable {
    private static final long serialVersionUID = 5628415838137969509L;

    /**
     * 工作簿对象
     */
    @Getter
    protected W workbook;

    /**
     * excel版本枚举对象
     */
    @Getter
    protected ExcelVersionEnum version;

    /**
     * 根据后缀判断是否为 AbstractExcel 文件，后缀匹配xls、et和xlsx
     *
     * @param pathname 输入excel路径
     * @return
     */
    public static boolean isExcel(String pathname) {
        return ExcelVersionEnum.getVersion(pathname) != null;
    }

    @Override
    public void close() throws IOException {
        super.close();
        workbook.close();
    }

    /**
     * 将当前工作簿对象的修改保存到输出流
     *
     * @param outputStream 输出流
     * @throws IOException
     */
    @Override
    public abstract void save(OutputStream outputStream) throws IOException;

    /**
     * 获取excel 中sheet 总页数
     *
     * @return
     */
    public abstract int getSheetCount();

    /**
     * 返回sheet 中的行数
     *
     * @param sheetIndex 指定 Sheet 页，从 0 开始
     * @return
     */
    public abstract int getRowCount(int sheetIndex);

    /**
     * 返回所在行的列数
     *
     * @param sheetIndex 指定 Sheet 页，从 0 开始
     * @param rowIndex   指定行，从0开始
     * @return 返回-1 表示所在行为空
     */
    public abstract int getColumnCount(int sheetIndex, int rowIndex);

    /**
     * 获取sheet的索引，从0开始
     *
     * @param name sheet 名称
     * @return -1表示该未找到名称对应的sheet
     */
    public abstract int getSheetIndex(String name);

    /**
     * 获取 sheet名称
     *
     * @param sheetIndex 指定 Sheet 页，从 0 开始
     * @return
     */
    public abstract String getSheetName(int sheetIndex);

    /**
     * 设置sheet名称，长度为1-31，不能包含后面任一字符: ：\ / ? * [ ]
     *
     * @param sheetIndex 指定 Sheet 页，从 0 开始，//
     * @param name
     * @return
     */
    public abstract boolean setSheetName(int sheetIndex, String name);

    /**
     * 设置sheet 页的索引
     *
     * @param sheetName  Sheet 名称
     * @param sheetIndex Sheet 索引，从0开始
     */
    public abstract void setSheetOrder(String sheetName, int sheetIndex);

    /**
     * 清空指定sheet页（先删除后添加并指定sheetIndex）
     *
     * @param sheetIndex 指定 Sheet 页，从 0 开始
     * @return
     */
    public abstract boolean clearSheet(int sheetIndex);

    /**
     * 删除指定sheet
     *
     * @param sheetIndex 指定 Sheet 页，从 0 开始
     * @return
     */
    public abstract boolean removeSheetAt(int sheetIndex);

    /**
     * 删除指定名称的sheet
     *
     * @param sheetName sheet名称
     * @return
     */
    public boolean removeSheetByName(String sheetName) {
        removeSheetAt(getSheetIndex(sheetName));
        return true;
    }

    /**
     * 删除指定sheet中行，改变该行之后行的索引
     *
     * @param sheetIndex 指定 Sheet 页，从 0 开始
     * @param rowIndex   指定行，从0开始
     * @return
     */
    public abstract boolean removeRow(int sheetIndex, int rowIndex);

    /**
     * 删除指定sheet中rowIndexStart（含）到rowIndexEnd（含）的行，改变该行之后行的索引
     *
     * @param sheetIndex    指定 Sheet 页，从 0 开始
     * @param rowIndexStart 起始行（含）
     * @param rowIndexEnd   结束行（含）
     * @return
     */
    public boolean removeRows(int sheetIndex, int rowIndexStart, int rowIndexEnd) {
        boolean r = false;
        if (rowIndexEnd > getRowCount(sheetIndex) - 1) {
            rowIndexEnd = getRowCount(sheetIndex) - 1;
        }
        for (int i = rowIndexEnd; i >= rowIndexStart; i--) {
            r = removeRow(sheetIndex, i) || r;
        }
        return r;
    }

    /**
     * 指定单元格是否为空
     *
     * @param sheetIndex 指定 Sheet 页，从 0 开始
     * @param rowIndex   指定开始行，从 0 开始
     * @param colIndex   指定开始列，从 0 开始
     * @return true 行不为空，false 行为空
     */
    public abstract boolean isNullCell(int sheetIndex, int rowIndex, int colIndex);

    /**
     * 创建单元格
     *
     * @param sheetIndex 指定 sheet 页，从 0 开始
     * @param rowIndex   指定行，从 0 开始
     * @param colIndex   指定创建列，从 0 开始
     * @return true 列为空，false 行不为空
     */
    public abstract boolean createCell(int sheetIndex, int rowIndex, int colIndex);

    /**
     * 指定行是否为空
     *
     * @param sheetIndex 指定 Sheet 页，从 0 开始
     * @param rowIndex   指定开始行，从 0 开始
     * @return true 不为空，false 不行为空
     */
    public abstract boolean isNullRow(int sheetIndex, int rowIndex);

    /**
     * 创建行，若行存在，则清空
     *
     * @param sheetIndex 指定 sheet 页，从 0 开始
     * @param rowIndex   指定创建行，从 0 开始
     * @return
     */
    public abstract boolean createRow(int sheetIndex, int rowIndex);

    /**
     * 重置指定行的值。从第0列开始
     *
     * @param rowData    数据
     * @param sheetIndex 指定 Sheet 页，从 0 开始
     * @param rowIndex   指定行，从0开始
     * @return
     */
    public boolean writeRow(int sheetIndex, int rowIndex, List<String> rowData) {
        return writeRow(sheetIndex, rowIndex, 0, rowData);
    }

    /**
     * 重置指定行的值。从第startColumnIndex列开始
     *
     * @param rowData          数据
     * @param sheetIndex       指定 Sheet 页，从 0 开始
     * @param rowIndex         指定行，从0开始
     * @param startColumnIndex 从第几列开始写该行数据
     * @return
     */
    public boolean writeRow(int sheetIndex, int rowIndex, int startColumnIndex, List<String> rowData) {
        for (int i = 0; i < rowData.size(); i++) {
            write(sheetIndex, rowIndex, i + startColumnIndex, rowData.get(i));
        }
        return true;
    }


    /**
     * 重置指定列的值，从第0行开始写
     *
     * @param columnData  数据
     * @param sheetIndex  指定 Sheet 页，从 0 开始
     * @param columnIndex 指定行，从0开始
     * @return
     */
    public boolean writeColumn(int sheetIndex, int columnIndex, List<String> columnData) {
        return writeColumn(sheetIndex, 0, columnIndex, columnData);
    }

    /**
     * 重置指定列的值，从第0行开始写
     *
     * @param columnData    数据
     * @param sheetIndex    指定 Sheet 页，从 0 开始
     * @param startRowIndex 从第几行开始写该行数据
     * @param columnIndex   指定行，从0开始
     * @return
     */
    public boolean writeColumn(int sheetIndex, int startRowIndex, int columnIndex, List<String> columnData) {
        for (int i = 0; i < columnData.size(); i++) {
            write(sheetIndex, i + startRowIndex, columnIndex, columnData.get(i));
        }
        return true;
    }

    /**
     * 将同一个值value设置在指定区域内的每一个单元格
     *
     * @param sheetIndex  sheet号
     * @param value       值
     * @param startColumn 起始列（含）
     * @param endColumn   结束列（含）
     * @param startRow    起始行（含）
     * @param endRow      结束行（含）
     * @return
     */
    public boolean write(int sheetIndex, int startColumn, int endColumn, int startRow, int endRow, String value) {
        for (int i = startColumn; i <= endColumn; i++) {
            for (int j = startRow; j <= endRow; j++) {
                write(sheetIndex, j, i, value);
            }
        }
        return true;
    }


    /**
     * 设置row 和 column 位置的单元格值
     *
     * @param sheetIndex 指定 Sheet 页，从 0 开始
     * @param rowIndex   指定行，从0开始
     * @param colIndex   指定列，从0开始
     * @param value      值
     * @return
     */
    public abstract boolean write(int sheetIndex, int rowIndex, int colIndex, String value);

    /**
     * 按行读取 AbstractExcel 第一页所有数据
     *
     * @return
     */
    public List<List<String>> read() {
        return read(0);
    }

    /**
     * 按行读取指定sheet 页所有数据
     *
     * @param sheetIndex 指定 sheet 页，从 0 开始
     * @return
     */
    public List<List<String>> read(int sheetIndex) {
        return readRows(sheetIndex, 0);
    }

    /**
     * 读取指定sheet 页指定行数据，第startRowIndex~最后有效行
     *
     * @param sheetIndex    指定 sheet 页，从 0 开始
     * @param startRowIndex 指定开始行（含）
     * @return
     */
    public List<List<String>> readRows(int sheetIndex, int startRowIndex) {
        return readRows(sheetIndex, startRowIndex, Integer.MAX_VALUE);
    }

    /**
     * 读取指定sheet 页指定行数据，第startRowIndex~endRowIndex行
     *
     * @param sheetIndex    指定 sheet 页，从 0 开始
     * @param startRowIndex 指定开始行（含）
     * @param endRowIndex   指定结束行（含）
     * @return
     */
    public List<List<String>> readRows(int sheetIndex, int startRowIndex, int endRowIndex) {
        List<List<String>> list = new ArrayList<>();
        int rowCount = getRowCount(sheetIndex);
        if (endRowIndex > rowCount - 1) {
            endRowIndex = rowCount - 1;
        }

        for (int i = startRowIndex; i <= endRowIndex; i++) {
            list.add(readRow(sheetIndex, i));
        }

        return list;
    }

    /**
     * 返回指定行的值的集合。
     *
     * @param sheetIndex 指定 Sheet 页，从 0 开始
     * @param rowIndex   指定行，从0开始
     * @return
     */
    public List<String> readRow(int sheetIndex, int rowIndex) {
        return readRow(sheetIndex, rowIndex, 0);
    }

    /**
     * 返回指定行的值的集合，从startColumnIndex列开始到当前行最后的有效列
     *
     * @param sheetIndex       指定 Sheet 页，从 0 开始
     * @param rowIndex         指定行，从0开始
     * @param startColumnIndex 起始列（含）
     * @return
     */
    public List<String> readRow(int sheetIndex, int rowIndex, int startColumnIndex) {
        return readRow(sheetIndex, rowIndex, startColumnIndex, Integer.MAX_VALUE);
    }

    /**
     * 返回指定行从startColumnIndex列~endColumnIndex列的值的集合
     *
     * @param sheetIndex       指定 Sheet 页，从 0 开始
     * @param rowIndex         指定行，从0开始
     * @param startColumnIndex 起始列（含）
     * @param endColumnIndex   结束列（含）
     * @return
     */
    public abstract List<String> readRow(int sheetIndex, int rowIndex, int startColumnIndex, int endColumnIndex);

    /**
     * 返回 row 和 column 位置的单元格值
     *
     * @param sheetIndex 指定 Sheet 页，从 0 开始
     * @param rowIndex   指定行，从0开始
     * @param colIndex   指定列，从0开始
     * @return
     */
    public abstract String read(int sheetIndex, int rowIndex, int colIndex);

    /**
     * 返回列的值的集合
     *
     * @param sheetIndex 指定 Sheet 页，从 0 开始
     * @param colIndex   指定列，从0开始
     * @return
     */
    public List<String> readColumn(int sheetIndex, int colIndex) {
        return readColumn(sheetIndex, 0, colIndex);
    }

    /**
     * 返回列的值的集合，从startRowIndex行~表格最大有效行
     *
     * @param sheetIndex    指定 Sheet 页，从 0 开始
     * @param startRowIndex 起始行（含）
     * @param colIndex      指定列，从0开始
     * @return
     */
    public List<String> readColumn(int sheetIndex, int startRowIndex, int colIndex) {
        return readColumn(sheetIndex, startRowIndex, Integer.MAX_VALUE, colIndex);
    }

    /**
     * 返回列的值的集合，从startRowIndex行~endRowIndex行
     *
     * @param sheetIndex    指定 Sheet 页，从 0 开始
     * @param startRowIndex 起始行（含）
     * @param endRowIndex   结束行（含）
     * @param colIndex      指定列，从0开始
     * @return
     */
    public abstract List<String> readColumn(int sheetIndex, int startRowIndex, int endRowIndex, int colIndex);
}