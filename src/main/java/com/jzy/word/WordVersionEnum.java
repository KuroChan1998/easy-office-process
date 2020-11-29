package com.jzy.word;

import com.jzy.excel.ExcelVersionEnum;

import java.util.ArrayList;
import java.util.List;

/**
 * @author JinZhiyun
 * @version 1.0
 * @EnumName WordVersionEnum
 * @description Word版本枚举类
 * @date 2020/11/28 19:55
 **/
public enum WordVersionEnum {
    /**
     * word 2003 版本
     */
    VERSION_2003(".doc"),

    /**
     * word 2007 版本
     */
    VERSION_2007(".docx"),

    /**
     * wps格式
     */
    VERSION_WPS(".wps");

    private String suffix;

    WordVersionEnum(String suffix) {
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
    public static List<String> listAllVersionSuffix(){
        List<String> list = new ArrayList<>();
        for (WordVersionEnum versionEnum : WordVersionEnum.values()) {
            list.add(versionEnum.suffix);
        }
        return list;
    }

    /**
     * 根据输入的文件名返回相应的word版本枚举对象
     *
     * @param pathname 文件名
     * @return word版本
     */
    public static WordVersionEnum getVersion(String pathname) {
        if (pathname == null) {
            return null;
        }
        if (pathname.endsWith(WordVersionEnum.VERSION_2003.getSuffix())) {
            return WordVersionEnum.VERSION_2003;
        }
        if (pathname.endsWith(WordVersionEnum.VERSION_2007.getSuffix())) {
            return VERSION_2007;
        }

        if (pathname.endsWith(WordVersionEnum.VERSION_WPS.getSuffix())) {
            return VERSION_WPS;
        }
        return null;
    }
}
