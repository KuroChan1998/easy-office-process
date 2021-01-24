package com.jzy.office.matcher;

/**
 * @ClassName DefaultLabelMatcher
 * @Author JinZhiyun
 * @Description 默认的标签匹配器，默认对象匹配${}形式的标签，匹配大小写
 * @Date 2021/1/16 16:52
 * @Version 1.0
 **/
public class DefaultLabelMatcher extends LabelMatcher {
    public static final String DEFAULT_LABEL_REGEX = "\\$\\{(.+?)\\}";

    public DefaultLabelMatcher() {
        this(DEFAULT_LABEL_REGEX);
    }

    public DefaultLabelMatcher(String regex) {
        super(regex);
    }
}
