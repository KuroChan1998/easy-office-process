package com.jzy.excel;


import com.jzy.exception.ExcelColumnNotFoundException;
import com.jzy.exception.ExcelSheetNameInvalidException;
import com.jzy.exception.ExcelTooManyRowsException;
import com.jzy.exception.InvalidFileTypeException;
import lombok.ToString;
import org.apache.poi.ss.usermodel.Workbook;

import java.io.File;
import java.io.IOException;
import java.io.InputStream;

/**
 * @ClassName AbstractInputExcel
 * @Author JinZhiyun
 * @Description 待读取的输入表格的抽象类。input类别的表格目的就是为了要从表格中读取信息，封装成对象至成员变量中。
 * @Date 2020/1/12 15:13
 * @Version 1.0
 **/
@ToString(callSuper = true)
public abstract class AbstractInputExcel extends DefaultExcel implements Readable {
    private static final long serialVersionUID = 7458303551368759495L;

    /**
     * 有效信息开始的行，默认——从第1行开始
     */
    protected static final int DEFAULT_START_ROW = 0;

    public AbstractInputExcel(String inputFile) throws IOException, InvalidFileTypeException {
        super(inputFile);
    }

    public AbstractInputExcel(File file) throws IOException, InvalidFileTypeException {
        super(file);
    }

    public AbstractInputExcel(InputStream inputStream, ExcelVersionEnum version) throws IOException, InvalidFileTypeException {
        super(inputStream, version);
    }

    public AbstractInputExcel(Workbook workbook) {
        super(workbook);
    }

    public AbstractInputExcel(ExcelVersionEnum version) throws InvalidFileTypeException {
        super(version);
    }

    /**
     * input类别的表格读取信息，封装成对象至成员变量中。此方法重置所有表示读取结果的成员变量
     */
    @Override
    public abstract void resetOutput();

    /**
     * 重置所有表示规定列的索引值的成员变量。
     */
    @Override
    public abstract void resetColumnIndex();

    /**
     * 找到当前sheet指定列名称对应的列的索引
     *
     * @param sheetIx 要处理的sheet的索引
     */
    protected abstract void findColumnIndexOfSpecifiedName(int sheetIx);

    /**
     * 表格的列属性名是否符合要求
     *
     * @return 是否符合的布尔值
     * @throws ExcelColumnNotFoundException 规定名称的列未找到
     */
    @Override
    public abstract boolean testColumnNameValidity() throws ExcelColumnNotFoundException;

    /**
     * 测试批量读取的前提条件是否满足，然后执行读取，默认第一张sheet的数据。
     *
     * @return 返回表格目标sheet有效数据的行数
     * @throws ExcelColumnNotFoundException   列属性中有未匹配的属性名
     * @throws ExcelTooManyRowsException      行数超过规定值，将规定的上限值和实际值都传给异常对象
     * @throws ExcelSheetNameInvalidException 不合法的sheet名
     */
    public int testAndReadData() throws ExcelTooManyRowsException, ExcelColumnNotFoundException, ExcelSheetNameInvalidException {
        return testAndReadData(0);
    }

    /**
     * 测试批量读取的前提条件是否满足，然后执行读取目标sheet。
     *
     * @param sheetIndex 目标sheet的索引
     * @return 返回表格目标sheet有效数据的行数
     * @throws ExcelColumnNotFoundException 列属性中有未匹配的属性名
     * @throws ExcelTooManyRowsException    行数超过规定值，将规定的上限值和实际值都传给异常对象
     */
    public int testAndReadData(int sheetIndex) throws ExcelTooManyRowsException, ExcelColumnNotFoundException {
        //读取数据前的准备工作
        prepareForReading(sheetIndex);

        //批量读取具体数据（子类实现）
        return readData(sheetIndex);
    }

    public void prepareForReading(int sheetIndex) throws ExcelTooManyRowsException, ExcelColumnNotFoundException {
        //重置所有暂存的读取结果（成员变量），子类实现
        resetOutput();

        //当前表格指定索引的sheet是否超过默认最大行数限制（30000行）。此为父类中方法，如果要更改此标准请重写testRowCountValidityOfSheet
        testRowCountValidityOfSheet(sheetIndex);

        /*=========================
            开始尝试寻找要读取的目标列（是否有必要？如何读取？——取决于子类的实现）
                                    ==========================*/
        //重置目标列索引值（子类实现）
        resetColumnIndex();

        //找到要读取的列的位置（子类实现）
        findColumnIndexOfSpecifiedName(sheetIndex);

        //测试判断要读取的目标列是否找到（子类实现）。如果没有，抛出异常。
        testColumnNameValidity();
        /*=======================测试结束=====================*/
    }

    /**
     * 从excel的某张sheet中批量读取数据，数据存放于成员变量。子类实现细节
     *
     * @param sheetIndex sheet索引
     * @return 当前sheet的有效行数
     */
    @Override
    public abstract int readData(int sheetIndex);
}
