package com.jzy.office.excel;

import com.jzy.office.exception.InvalidFileTypeException;
import com.jzy.util.MyTimeUtils;
import lombok.Getter;
import lombok.Setter;
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
 * @ClassName CommonExcel
 * @Author JinZhiyun
 * @Description Excel 2003 2007 版本的通用包装类，支持.xls, .xlsx, .et文件的解析
 * @Date 2021/1/23 22:06
 * @Version 1.0
 **/
public class CommonExcel extends AbstractExcel<Workbook> {
    private static final long serialVersionUID = 7769992970075361131L;

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
    public CommonExcel(String inputFile) throws IOException, InvalidFileTypeException {
        this(new File(inputFile));
    }

    /**
     * 由一个File构造excel对象
     *
     * @param file 输入文件对象
     * @throws IOException
     * @throws InvalidFileTypeException
     */
    public CommonExcel(File file) throws IOException, InvalidFileTypeException {
        this(new FileInputStream(file), ExcelVersionEnum.getVersion(file.getAbsolutePath()));
        this.inputFilePath = file.getAbsolutePath();
    }

    /**
     * 由一个输入流和版本枚举对象构造excel对象
     *
     * @param inputStream 输入流对象
     * @param version     excel版本的枚举对象
     * @throws IOException
     * @throws InvalidFileTypeException
     */
    public CommonExcel(InputStream inputStream, ExcelVersionEnum version) throws IOException, InvalidFileTypeException {
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
    public CommonExcel(Workbook workbook) {
        this.workbook = workbook;
    }

    /**
     * 构建指定excel版本的新表格
     *
     * @param version excel版本的枚举对象
     * @throws InvalidFileTypeException 不合法的入参excel版本枚举异常
     */
    public CommonExcel(ExcelVersionEnum version) throws InvalidFileTypeException {
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

    @Override
    public void save(OutputStream outputStream) throws IOException {
        workbook.write(outputStream);
    }

    @Override
    public int getSheetCount() {
        return workbook.getNumberOfSheets();
    }

    @Override
    public int getRowCount(int sheetIndex) {
        Sheet sheet = workbook.getSheetAt(sheetIndex);
        if (sheet.getPhysicalNumberOfRows() == 0) {
            return 0;
        }
        return sheet.getLastRowNum() + 1;

    }

    @Override
    public int getColumnCount(int sheetIndex, int rowIndex) {
        Sheet sheet = workbook.getSheetAt(sheetIndex);
        Row row = sheet.getRow(rowIndex);
        return row == null ? -1 : row.getLastCellNum();
    }

    @Override
    public int getSheetIndex(String name) {
        return workbook.getSheetIndex(name);
    }

    @Override
    public String getSheetName(int sheetIndex) {
        Sheet sheet = workbook.getSheetAt(sheetIndex);
        if (sheet == null) {
            return null;
        }
        return sheet.getSheetName();
    }

    @Override
    public boolean setSheetName(int sheetIndex, String name) {
        workbook.setSheetName(sheetIndex, name);
        return true;
    }

    @Override
    public void setSheetOrder(String sheetName, int sheetIndex) {
        workbook.setSheetOrder(sheetName, sheetIndex);
    }

    @Override
    public boolean clearSheet(int sheetIndex) {
        String sheetName = getSheetName(sheetIndex);
        removeSheetAt(sheetIndex);
        workbook.createSheet(sheetName);
        setSheetOrder(sheetName, sheetIndex);
        return true;
    }

    @Override
    public boolean removeSheetAt(int sheetIndex) {
        workbook.removeSheetAt(sheetIndex);
        return true;
    }

    public Sheet createSheet() {
        return workbook.createSheet();
    }

    public Sheet createSheet(String sheetName) {
        return workbook.createSheet(sheetName);
    }


    @Override
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

    @Override
    public boolean isNullCell(int sheetIndex, int rowIndex, int colIndex) {
        Sheet sheet = workbook.getSheetAt(sheetIndex);
        if (isNullRow(sheetIndex, rowIndex)) {
            return true;
        }
        Row row = sheet.getRow(rowIndex);
        return row.getCell(colIndex) == null;
    }

    @Override
    public boolean createCell(int sheetIndex, int rowIndex, int colIndex) {
        Sheet sheet = workbook.getSheetAt(sheetIndex);
        if (isNullRow(sheetIndex, rowIndex)) {
            createRow(sheetIndex, rowIndex);
        }
        sheet.getRow(rowIndex).createCell(colIndex);
        return true;
    }

    @Override
    public boolean isNullRow(int sheetIndex, int rowIndex) {
        Sheet sheet = workbook.getSheetAt(sheetIndex);
        return sheet.getRow(rowIndex) == null;
    }

    @Override
    public boolean createRow(int sheetIndex, int rowIndex) {
        Sheet sheet = workbook.getSheetAt(sheetIndex);
        sheet.createRow(rowIndex);
        return true;
    }


    @Override
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

    @Override
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

    @Override
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

    @Override
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
