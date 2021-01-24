package com.jzy.office.matcher;

import org.apache.commons.lang3.StringUtils;

/**
 * @ClassName LabelMatchers
 * @Author JinZhiyun
 * @Description 标签匹配器池
 * @Date 2021/1/16 18:56
 * @Version 1.0
 **/
public class LabelMatchers extends Matchers<LabelMatcher> {
    public static final LabelMatcher DEFAULT_LABEL_MATCHER = new DefaultLabelMatcher();

    public static LabelMatcher getLabelMatcher(String regex) {
        if (StringUtils.isEmpty(regex)) {
            return null;
        }
        return new DefaultLabelMatcher(regex);
    }

    public static LabelMatchers getInstance() {
        LabelMatchers matchers = new LabelMatchers();
        matchers.putMatcher(DEFAULT_LABEL_MATCHER);
        return matchers;
    }
}
