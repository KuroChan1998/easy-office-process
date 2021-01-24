package com.jzy.office.excel;

import com.jzy.office.Readable;

/**
 * @InterfaceName ExcelReadable
 * @Author JinZhiyun
 * @Description excel的批量数据可读特性描述
 * @Date 2020/3/31 15:36
 * @Version 1.0
 **/
public interface ExcelReadable extends Readable {
    /**
     * 从excel的某张sheet中批量读取数据，数据存放于成员变量
     *
     * @param sheetIndex sheet索引
     * @return 当前sheet的有效行数
     */
    int readData(int sheetIndex);
}
