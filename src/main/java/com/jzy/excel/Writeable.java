package com.jzy.excel;

/**
 * @ClassName Writeable
 * @Author JinZhiyun
 * @Description excel的批量数据可写特性描述
 * @Date 2020/3/31 15:44
 * @Version 1.0
 **/
public interface Writeable {
    /**
     * 往excel的某张sheet中批量写入数据，要写入的数据预先存放于成员变量
     *
     * @return 写入成功与否
     * @throws Exception  导致写入失败的异常
     */
    boolean writeData() throws Exception;
}
