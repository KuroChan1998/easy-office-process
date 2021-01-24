package com.jzy.office.matcher;

import com.jzy.office.exception.NotInitException;
import lombok.Getter;
import org.apache.commons.lang3.StringUtils;

import java.util.ArrayList;
import java.util.List;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

/**
 * @ClassName StringMatcher
 * @Author JinZhiyun
 * @Description 字符串匹配器
 * @Date 2021/1/16 18:37
 * @Version 1.0
 **/
public class StringMatcher implements StringMatchable {
    /**
     * 正则表达式
     */
    @Getter
    protected String regex;

    /**
     * 匹配模式
     */
    @Getter
    protected Pattern pattern = Pattern.compile("", Pattern.CASE_INSENSITIVE);

    public StringMatcher() {
    }

    public StringMatcher(String regex) {
        this.regex = regex;
        this.pattern = Pattern.compile(regex, Pattern.CASE_INSENSITIVE);
    }

    public StringMatcher(Pattern pattern) {
        this.pattern = pattern;
        this.regex = pattern.toString();
    }

    private void assertInitRegexOrPattern() throws NotInitException {
        if (pattern == null || StringUtils.isEmpty(regex)) {
            throw new NotInitException("匹配模式或正则表达式未指定！请指定正则表达式regex或匹配模式regex。");
        }
    }

    public Matcher getMatcher(String target) {
        try {
            assertInitRegexOrPattern();
        } catch (NotInitException e) {
            e.printStackTrace();
            return null;
        }
        return pattern.matcher(target);
    }

    @Override
    public String replaceFirst(String target, String replacement) {
        return getMatcher(target).replaceFirst(replacement);
    }

    @Override
    public String replaceAll(String target, String replacement) {
        return getMatcher(target).replaceAll(replacement);
    }

    @Override
    public boolean match(String target) {
        return getMatcher(target).matches();
    }

    @Override
    public boolean find(String target) {
        return getMatcher(target).find();
    }

    @Override
    public String group(String target, int group) {
        Matcher matcher = getMatcher(target);
        if (matcher.find()) {
            return matcher.group(group);
        }
        return "";
    }

    @Override
    public List<String> groupAll(String target) {
        List<String> r = new ArrayList<>();
        Matcher matcher = getMatcher(target);
        while (matcher.find()) {
            r.add(matcher.group());
        }
        return r;
    }

    @Override
    public String toString() {
        return "StringMatcher{" +
                "regex='" + regex + '\'' +
                ", pattern=" + pattern +
                '}';
    }
}
