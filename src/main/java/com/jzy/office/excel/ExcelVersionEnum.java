package com.jzy.office.excel;

import java.util.ArrayList;
import java.util.List;

/**
 * @author JinZhiyun
 * @version 1.0
 * @EnumName ExcelVersionEnum
 * @description Excel版本枚举类
 * @date 2019/10/30 12:55
 **/
public enum ExcelVersionEnum {
    /**
     * AbstractExcel 2003 版本
     */
    VERSION_2003(".xls"),

    /**
     * AbstractExcel 2007 版本
     */
    VERSION_2007(".xlsx"),

    /**
     * wps et 格式
     */
    VERSION_ET(".et");

    private String suffix;

    ExcelVersionEnum(String suffix) {
        this.suffix = suffix;
    }

    public String getSuffix() {
        return suffix;
    }

    /**
     * 返回所有合法的文件版本后缀
     *
     * @return
     */
    public static List<String> listAllVersionSuffix() {
        List<String> list = new ArrayList<>();
        for (ExcelVersionEnum versionEnum : ExcelVersionEnum.values()) {
            list.add(versionEnum.suffix);
        }
        return list;
    }

    /**
     * 根据输入的文件名返回相应的excel版本枚举对象
     *
     * @param pathname 文件名
     * @return excel版本
     */
    public static ExcelVersionEnum getVersion(String pathname) {
        if (pathname == null) {
            return null;
        }
        if (pathname.endsWith(ExcelVersionEnum.VERSION_2003.getSuffix())) {
            return ExcelVersionEnum.VERSION_2003;
        }
        if (pathname.endsWith(ExcelVersionEnum.VERSION_2007.getSuffix())) {
            return VERSION_2007;
        }

        if (pathname.endsWith(ExcelVersionEnum.VERSION_ET.getSuffix())) {
            return VERSION_ET;
        }
        return null;
    }
}
