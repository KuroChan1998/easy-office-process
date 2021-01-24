package com.jzy.office.excel;

/**
 * @author JinZhiyun
 * @version 1.0
 * @IntefaceName ExcelResettable
 * @description 可重置特性的接口
 * @date 2019/11/1 15:05
 **/
public interface ExcelResettable {
    /**
     * 重置所有表示读取结果的成员变量
     */
    default void resetOutput() {
    }

    /**
     * 重置所有表示规定列的索引值的成员变量
     */
    default void resetColumnIndex() {
    }
}
