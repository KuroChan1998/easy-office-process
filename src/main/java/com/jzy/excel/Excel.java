package com.jzy.excel;

import com.jzy.excel.exception.InvalidFileTypeException;
import com.jzy.util.MyTimeUtils;
import lombok.Getter;
import lombok.Setter;
import org.apache.commons.lang3.StringUtils;
import org.apache.poi.hssf.usermodel.HSSFDateUtil;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.*;
import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.Date;
import java.util.List;

/**
 * Excel 包装类，对poi的二次封装
 *
 * @author JinZhiyun
 * @version 1.0 2020/01/27
 */
public abstract class Excel implements Serializable {
    private static final long serialVersionUID = 5628415838137969509L;

    /**
     * 工作簿对象
     */
    @Getter
    @Setter
    protected Workbook workbook;

    /**
     * excel版本枚举对象
     */
    @Getter
    protected ExcelVersionEnum version;

    /**
     * 输入文件路径
     */
    @Getter
    private String inputFilePath;

    /**
     * 输出流
     */
    private OutputStream os;

    /**
     * 日期格式
     */
    @Getter
    @Setter
    private String datePattern = MyTimeUtils.FORMAT_YMDHMS_BACKUP;

    /**
     * 由输入文件路径构造excel对象
     *
     * @param inputFile 输入文件路径
     * @throws IOException
     * @throws InvalidFileTypeException
     */
    public Excel(String inputFile) throws IOException, InvalidFileTypeException {
        this(new File(inputFile));
    }

    /**
     * 由一个File构造excel对象
     *
     * @param file 输入文件对象
     * @throws IOException
     * @throws InvalidFileTypeException
     */
    public Excel(File file) throws IOException, InvalidFileTypeException {
        String inputFile = file.getAbsolutePath();
        ExcelVersionEnum versionEnum = ExcelVersionEnum.getVersion(inputFile);
        if (ExcelVersionEnum.VERSION_2003.equals(versionEnum)) {
            version = ExcelVersionEnum.VERSION_2003;
            workbook = new HSSFWorkbook(new FileInputStream(file));
        } else if (ExcelVersionEnum.VERSION_2007.equals(versionEnum)) {
            version = ExcelVersionEnum.VERSION_2007;
            workbook = new XSSFWorkbook(new FileInputStream(file));
        } else if (ExcelVersionEnum.VERSION_ET.equals(versionEnum)) {
            version = ExcelVersionEnum.VERSION_ET;
            workbook = new HSSFWorkbook(new FileInputStream(file));
        } else {
            throw new InvalidFileTypeException("错误的文件类型！文件类型仅支持：" + ExcelVersionEnum.listAllVersionSuffix());
        }
        this.inputFilePath = inputFile;
    }

    /**
     * 由一个输入流和版本枚举对象构造excel对象
     *
     * @param inputStream 输入流对象
     * @param version     excel版本的枚举对象
     * @throws IOException
     * @throws InvalidFileTypeException
     */
    public Excel(InputStream inputStream, ExcelVersionEnum version) throws IOException, InvalidFileTypeException {
        if (ExcelVersionEnum.VERSION_2003.equals(version) || ExcelVersionEnum.VERSION_ET.equals(version)) {
            this.version = version;
            workbook = new HSSFWorkbook(inputStream);
        } else if (ExcelVersionEnum.VERSION_2007.equals(version)) {
            this.version = version;
            workbook = new XSSFWorkbook(inputStream);
        } else {
            throw new InvalidFileTypeException("错误的文件类型！文件类型仅支持：" + ExcelVersionEnum.listAllVersionSuffix());
        }
    }

    /**
     * 由一个工作簿构造excel对象
     *
     * @param workbook 工作簿对象
     */
    public Excel(Workbook workbook) {
        this.workbook = workbook;
    }

    /**
     * 构建指定excel版本的新表格
     *
     * @param version excel版本的枚举对象
     * @throws InvalidFileTypeException 不合法的入参excel版本枚举异常
     */
    public Excel(ExcelVersionEnum version) throws InvalidFileTypeException {
        if (ExcelVersionEnum.VERSION_2003.equals(version) || ExcelVersionEnum.VERSION_ET.equals(version)) {
            this.version = version;
            workbook = new HSSFWorkbook();
        } else if (ExcelVersionEnum.VERSION_2007.equals(version)) {
            this.version = version;
            workbook = new XSSFWorkbook();
        } else {
            throw new InvalidFileTypeException("错误的文件类型！文件类型仅支持：" + ExcelVersionEnum.listAllVersionSuffix());
        }
    }

    @Override
    public String toString() {
        return "共有 " + getSheetCount() + "个sheet 页！";
    }

    public String toString(int sheetIndex) {
        return "第 " + (sheetIndex + 1) + "个sheet 页，名称： " + getSheetName(sheetIndex) + "，共 " + getRowCount(sheetIndex) + "行！";
    }

    /**
     * 根据后缀判断是否为 Excel 文件，后缀匹配xls、et和xlsx
     *
     * @param pathname 输入excel路径
     * @return
     */
    public static boolean isExcel(String pathname) {
        return ExcelVersionEnum.getVersion(pathname) != null;
    }

    /**
     * 按行读取 Excel 第一页所有数据
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
     * 返回 row 和 column 位置的单元格值
     *
     * @param sheetIndex 指定 Sheet 页，从 0 开始
     * @param rowIndex   指定行，从0开始
     * @param colIndex   指定列，从0开始
     * @return
     */
    public String read(int sheetIndex, int rowIndex, int colIndex) {
        if (rowIndex < 0 || colIndex < 0) {
            return null;
        }
        Sheet sheet = workbook.getSheetAt(sheetIndex);
        Row row = sheet.getRow(rowIndex);
        if (row == null) {
            return null;
        }
        return getCellValueToString(row.getCell(colIndex));
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
    public List<String> readRow(int sheetIndex, int rowIndex, int startColumnIndex, int endColumnIndex) {
        Sheet sheet = workbook.getSheetAt(sheetIndex);
        List<String> list = new ArrayList<String>();
        Row row = sheet.getRow(rowIndex);
        if (row == null) {
            list.add(null);
        } else {
            if (endColumnIndex > row.getLastCellNum() - 1) {
                //结束列超过当前行列最大索引
                endColumnIndex = row.getLastCellNum() - 1;
            }
            for (int i = startColumnIndex; i <= endColumnIndex; i++) {
                list.add(getCellValueToString(row.getCell(i)));
            }
        }
        return list;
    }

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
    public List<String> readColumn(int sheetIndex, int startRowIndex, int endRowIndex, int colIndex) {
        Sheet sheet = workbook.getSheetAt(sheetIndex);
        List<String> list = new ArrayList<String>();
        if (sheet == null) {
            return list;
        }
        int rowCount = getRowCount(sheetIndex);
        if (endRowIndex > rowCount - 1) {
            endRowIndex = rowCount - 1;
        }
        for (int i = startRowIndex; i <= endRowIndex; i++) {
            Row row = sheet.getRow(i);
            if (row == null) {
                list.add(null);
                continue;
            }
            String value = getCellValueToString(row.getCell(colIndex));
            list.add(value);
        }
        return list;
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
     * 设置cell 样式
     *
     * @param sheetIndex 指定 Sheet 页，从 0 开始
     * @param rowIndex   指定行，从 0 开始
     * @param colIndex   指定列，从 0 开始
     * @param style      要设置样式
     * @return
     */
    public boolean setStyle(int sheetIndex, int rowIndex, int colIndex, CellStyle style) {
        Sheet sheet = workbook.getSheetAt(sheetIndex);
        // sheet.autoSizeColumn(colIndex, true);// 设置列宽度自适应
//        sheet.setColumnWidth(colIndex, 4000);
        if (isNullRow(sheetIndex, rowIndex) || isNullCell(sheetIndex, rowIndex, colIndex)) {
            return false;
        }
        Cell cell = sheet.getRow(rowIndex).getCell(colIndex);
        cell.setCellStyle(style);
        return true;
    }

    /**
     * 获得cell样式
     *
     * @param sheetIndex 指定 Sheet 页，从 0 开始
     * @param rowIndex   行索引
     * @param colIndex   列索引
     * @return cell样式
     */
    public CellStyle getStyle(int sheetIndex, int rowIndex, int colIndex) {
        Sheet sheet = workbook.getSheetAt(sheetIndex);
        Row row = sheet.getRow(rowIndex);
        if (row == null) {
            return null;
        }
        Cell cell = row.getCell(colIndex);
        if (cell == null) {
            return null;
        }
        return cell.getCellStyle();
    }

    /**
     * 设置单元格背景颜色，但不改变单元格原有样式
     *
     * @param sheetIndex 指定 Sheet 页，从 0 开始
     * @param rowIndex   行索引
     * @param colIndex   列索引
     * @param colorIndex 颜色的索引值
     * @return
     */
    public boolean updateCellBackgroundColor(int sheetIndex, int rowIndex, int colIndex, short colorIndex) {
        CellStyle cellStyle = workbook.createCellStyle();
        cellStyle.cloneStyleFrom(getStyle(sheetIndex, rowIndex, colIndex));
        cellStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);  //填充单元格
        cellStyle.setFillForegroundColor(colorIndex);
        setStyle(sheetIndex, rowIndex, colIndex, cellStyle);
        return true;
    }

    /**
     * 合并单元格
     *
     * @param sheetIndex 指定 Sheet 页，从 0 开始
     * @param firstRow   开始行
     * @param lastRow    结束行
     * @param firstCol   开始列
     * @param lastCol    结束列
     */
    public void region(int sheetIndex, int firstRow, int lastRow, int firstCol, int lastCol) {
        Sheet sheet = workbook.getSheetAt(sheetIndex);
        sheet.addMergedRegion(new CellRangeAddress(firstRow, lastRow, firstCol, lastCol));
    }

    /**
     * 指定行是否为空
     *
     * @param sheetIndex 指定 Sheet 页，从 0 开始
     * @param rowIndex   指定开始行，从 0 开始
     * @return true 不为空，false 不行为空
     */
    public boolean isNullRow(int sheetIndex, int rowIndex) {
        Sheet sheet = workbook.getSheetAt(sheetIndex);
        return sheet.getRow(rowIndex) == null;
    }

    /**
     * 创建行，若行存在，则清空
     *
     * @param sheetIndex 指定 sheet 页，从 0 开始
     * @param rowIndex   指定创建行，从 0 开始
     * @return
     */
    public boolean createRow(int sheetIndex, int rowIndex) {
        Sheet sheet = workbook.getSheetAt(sheetIndex);
        sheet.createRow(rowIndex);
        return true;
    }

    /**
     * 指定单元格是否为空
     *
     * @param sheetIndex 指定 Sheet 页，从 0 开始
     * @param rowIndex   指定开始行，从 0 开始
     * @param colIndex   指定开始列，从 0 开始
     * @return true 行不为空，false 行为空
     */
    public boolean isNullCell(int sheetIndex, int rowIndex, int colIndex) {
        Sheet sheet = workbook.getSheetAt(sheetIndex);
        if (isNullRow(sheetIndex, rowIndex)) {
            return true;
        }
        Row row = sheet.getRow(rowIndex);
        return row.getCell(colIndex) == null;
    }

    /**
     * 创建单元格
     *
     * @param sheetIndex 指定 sheet 页，从 0 开始
     * @param rowIndex   指定行，从 0 开始
     * @param colIndex   指定创建列，从 0 开始
     * @return true 列为空，false 行不为空
     */
    public boolean createCell(int sheetIndex, int rowIndex, int colIndex) {
        Sheet sheet = workbook.getSheetAt(sheetIndex);
        if (isNullRow(sheetIndex, rowIndex)) {
            createRow(sheetIndex, rowIndex);
        }
        sheet.getRow(rowIndex).createCell(colIndex);
        return true;
    }

    /**
     * 返回sheet 中的行数
     *
     * @param sheetIndex 指定 Sheet 页，从 0 开始
     * @return
     */
    public int getRowCount(int sheetIndex) {
        Sheet sheet = workbook.getSheetAt(sheetIndex);
        if (sheet.getPhysicalNumberOfRows() == 0) {
            return 0;
        }
        return sheet.getLastRowNum() + 1;

    }

    /**
     * 返回所在行的列数
     *
     * @param sheetIndex 指定 Sheet 页，从 0 开始
     * @param rowIndex   指定行，从0开始
     * @return 返回-1 表示所在行为空
     */
    public int getColumnCount(int sheetIndex, int rowIndex) {
        Sheet sheet = workbook.getSheetAt(sheetIndex);
        Row row = sheet.getRow(rowIndex);
        return row == null ? -1 : row.getLastCellNum();

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
    public boolean write(int sheetIndex, int rowIndex, int colIndex, String value) {
        Sheet sheet = workbook.getSheetAt(sheetIndex);
        if (isNullRow(sheetIndex, rowIndex)) {
            createRow(sheetIndex, rowIndex);
        }
        if (isNullCell(sheetIndex, rowIndex, colIndex)) {
            createCell(sheetIndex, rowIndex, colIndex);
        }
        sheet.getRow(rowIndex).getCell(colIndex).setCellValue(value);
        return true;
    }

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
     * 获取excel 中sheet 总页数
     *
     * @return
     */
    public int getSheetCount() {
        return workbook.getNumberOfSheets();
    }

    public Sheet createSheet() {
        return workbook.createSheet();
    }

    public Sheet createSheet(String sheetName) {
        return workbook.createSheet(sheetName);
    }

    /**
     * 设置sheet名称，长度为1-31，不能包含后面任一字符: ：\ / ? * [ ]
     *
     * @param sheetIndex 指定 Sheet 页，从 0 开始，//
     * @param name
     * @return
     */
    public boolean setSheetName(int sheetIndex, String name) {
        workbook.setSheetName(sheetIndex, name);
        return true;
    }

    /**
     * 获取 sheet名称
     *
     * @param sheetIndex 指定 Sheet 页，从 0 开始
     * @return
     */
    public String getSheetName(int sheetIndex) {
        Sheet sheet = workbook.getSheetAt(sheetIndex);
        if (sheet == null) {
            return null;
        }
        return sheet.getSheetName();
    }

    /**
     * 获取sheet的索引，从0开始
     *
     * @param name sheet 名称
     * @return -1表示该未找到名称对应的sheet
     */
    public int getSheetIndex(String name) {
        return workbook.getSheetIndex(name);
    }

    /**
     * 删除指定sheet
     *
     * @param sheetIndex 指定 Sheet 页，从 0 开始
     * @return
     */
    public boolean removeSheetAt(int sheetIndex) {
        workbook.removeSheetAt(sheetIndex);
        return true;
    }

    /**
     * 删除指定名称的sheet
     *
     * @param sheetName sheet名称
     * @return
     */
    public boolean removeSheetByName(String sheetName) {
        workbook.removeSheetAt(getSheetIndex(sheetName));
        return true;
    }

    /**
     * 删除指定sheet中行，改变该行之后行的索引
     *
     * @param sheetIndex 指定 Sheet 页，从 0 开始
     * @param rowIndex   指定行，从0开始
     * @return
     */
    public boolean removeRow(int sheetIndex, int rowIndex) {
        Sheet sheet = workbook.getSheetAt(sheetIndex);
        int lastRowNum = sheet.getLastRowNum();
        if (rowIndex < 0 || rowIndex > lastRowNum) {
            return false;
        }
        if (rowIndex != lastRowNum) {
            sheet.shiftRows(rowIndex + 1, lastRowNum, -1);// 将行号为rowIndex+1一直到行号为lastRowNum的单元格全部上移一行，以便删除rowIndex行
        } else {
            Row removingRow = sheet.getRow(rowIndex);
            if (removingRow != null) {
                sheet.removeRow(removingRow);
            }
        }
        return true;
    }

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
     * 设置sheet 页的索引
     *
     * @param sheetName  Sheet 名称
     * @param sheetIndex Sheet 索引，从0开始
     */
    public void setSheetOrder(String sheetName, int sheetIndex) {
        workbook.setSheetOrder(sheetName, sheetIndex);
    }

    /**
     * 清空指定sheet页（先删除后添加并指定sheetIndex）
     *
     * @param sheetIndex 指定 Sheet 页，从 0 开始
     * @return
     */
    public boolean clearSheet(int sheetIndex) {
        String sheetName = getSheetName(sheetIndex);
        removeSheetAt(sheetIndex);
        workbook.createSheet(sheetName);
        setSheetOrder(sheetName, sheetIndex);
        return true;
    }

    /**
     * 将当前修改保存覆盖至输入文件inputFilePath中
     *
     * @throws IOException
     */
    public void save() throws IOException {
        if (StringUtils.isNotEmpty(inputFilePath)) {
            os = new FileOutputStream(new File(inputFilePath));
            save(os);
        } else {
            throw new IOException("文件的默认路径（源文件路径）不存在");
        }
    }


    /**
     * 将当前修改保存到输出流
     *
     * @param outputStream 输出流
     * @throws IOException
     */
    public void save(OutputStream outputStream) throws IOException {
        this.workbook.write(outputStream);
    }

    /**
     * 将当前修改保存到outputPath对应的文件中
     *
     * @param outputPath 输出文件的路径
     * @throws IOException
     */
    public void save(String outputPath) throws IOException {
        os = new FileOutputStream(new File(outputPath));
        save(os);
    }

    /**
     * 关闭流
     *
     * @throws IOException
     */
    public void close() throws IOException {
        if (os != null) {
            os.close();
        }
        workbook.close();
    }

    /**
     * 转换单元格的类型为String 默认的 <br>
     * 默认的数据类型：CELL_TYPE_BLANK(3), CELL_TYPE_BOOLEAN(4), CELL_TYPE_ERROR(5),CELL_TYPE_FORMULA(2), CELL_TYPE_NUMERIC(0),
     * CELL_TYPE_STRING(1)
     *
     * @param cell
     * @return
     */
    private String getCellValueToString(Cell cell) {
        String strCell = "";
        if (cell == null) {
            return null;
        }
        switch (cell.getCellTypeEnum()) {
            case BOOLEAN:
                strCell = String.valueOf(cell.getBooleanCellValue());
                break;
            case NUMERIC:
                if (HSSFDateUtil.isCellDateFormatted(cell)) {
                    Date date = cell.getDateCellValue();
                    if (datePattern != null) {
                        SimpleDateFormat sdf = new SimpleDateFormat(datePattern);
                        strCell = sdf.format(date);
                    } else {
                        strCell = date.toString();
                    }
                    break;
                }
                // 不是日期格式，则防止当数字过长时以科学计数法显示
                cell.setCellType(CellType.STRING);
                strCell = cell.toString();
                break;
            case STRING:
                strCell = cell.getStringCellValue();
                break;
            default:
                break;
        }
        return strCell;
    }
}