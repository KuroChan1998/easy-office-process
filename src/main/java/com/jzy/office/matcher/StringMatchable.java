package com.jzy.office.matcher;

import java.util.List;

/**
 * @ClassName StringMatchable
 * @Author JinZhiyun
 * @Description
 * @Date 2021/1/16 17:03
 * @Version 1.0
 **/
public interface StringMatchable extends Matchable<String, String> {
    /**
     * 获得当前的匹配正则表达式
     *
     * @return
     */
    @Override
    String getRegex();

    /**
     * 使用正则表达式匹配目标字符串
     *
     * @param target
     * @return
     */
    @Override
    boolean match(String target);

    /**
     * 使用正则表达式在目标字符串中寻找是否有匹配的子串
     *
     * @param target
     * @return
     */
    @Override
    boolean find(String target);

    /**
     * 根据当前匹配的情况，将target中第一个匹配的子串替换为replacement
     *
     * @param target
     * @param replacement
     * @return 返回替换后的字符串
     */
    String replaceFirst(String target, String replacement);

    /**
     * 根据当前匹配的情况，将target中所有匹配的子串替换为replacement
     *
     * @param target
     * @param replacement
     * @return 返回替换后的字符串
     */
    String replaceAll(String target, String replacement);

    /**
     * 返回整个目标字符串中的匹配子串
     *
     * @param target 目标字符串
     * @return
     */
    default String group(String target) {
        return group(target, 0);
    }

    /**
     * 返回在目标字符串中正则表达式第几个匹配的()中的子串，0即整个字符串
     *
     * @param target 目标字符串
     * @param group  第几个()
     * @return 匹配的分组结果
     */
    String group(String target, int group);

    /**
     * 获得目标字符串中所有正则所匹配的子串集合
     *
     * @param target 目标字符串
     * @return
     */
    List<String> groupAll(String target);
}
