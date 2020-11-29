package com.jzy.demo.template;

import com.jzy.excel.AbstractTemplateExcel;
import com.jzy.exception.InvalidFileTypeException;
import com.jzy.util.MyTimeUtils;
import lombok.Getter;
import lombok.Setter;

import java.io.IOException;
import java.util.Date;

/**
 * @ClassName Test1TemplateExcel
 * @Author JinZhiyun
 * @Description 往学生信息表test1.xlsx的末尾追加上检阅人和检阅日期信息
 * @Date 2020/4/1 22:03
 * @Version 1.0
 **/
public class Test1TemplateExcel extends AbstractTemplateExcel {
    public Test1TemplateExcel(String inputFile) throws IOException, InvalidFileTypeException {
        super(inputFile);
    }

    /**
     * 检阅人
     */
    @Getter
    @Setter
    private String reviewer;

    /**
     * 检阅日期
     */
    @Getter
    @Setter
    private Date reviewDate;

    /**
     * 将预先存放好的数据执行写入到当前excel中。
     *
     * @return
     */
    @Override
    public boolean writeData() {
        int sheetIndex = 0;
        int rowCount = getRowCount(sheetIndex);
        //往尾行写信息
        write(sheetIndex, rowCount, 0, "审阅人:" + reviewer);
        write(sheetIndex, rowCount, 1, "审阅日期:" + MyTimeUtils.dateToStringYMD(reviewDate));
        return true;
    }
}
