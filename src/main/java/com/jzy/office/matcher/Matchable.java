package com.jzy.office.matcher;

/**
 * @ClassName Matchable
 * @Author JinZhiyun
 * @Description 匹配器接口
 * @Date 2021/1/16 16:34
 * @Version 1.0
 **/
public interface Matchable<R, T> {
    /**
     * 返回具体匹配规则
     *
     * @return
     */
    R getRegex();

    /**
     * 使用R的规则对整个T进行匹配
     *
     * @param target 匹配目标
     * @return 是否匹配
     */
    boolean match(T target);

    /**
     * 使用R的规则在T的子集进行匹配
     *
     * @param target
     * @return
     */
    boolean find(T target);
}
