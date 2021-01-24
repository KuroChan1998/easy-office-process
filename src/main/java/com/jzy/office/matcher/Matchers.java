package com.jzy.office.matcher;

import java.util.*;

/**
 * @ClassName Matchers
 * @Author JinZhiyun
 * @Description 抽象匹配器池
 * @Date 2021/1/16 16:47
 * @Version 1.0
 **/
public abstract class Matchers<M extends Matchable> {
    /**
     * 所有匹配器对应序号，从1开始
     */
    private List<Integer> sequences = new LinkedList<>();

    /**
     * <序号, 匹配器>
     */
    private Map<Integer, M> matchers = new HashMap<>();

    /**
     * 返回主匹配器，即序号最小最高
     *
     * @return
     */
    public M getMajorMatcher() {
        return getMatcher(getMinSequence());
    }

    /**
     * 返回指定序号的匹配器
     *
     * @param sequence 输入序号
     * @return
     */
    public M getMatcher(Integer sequence) {
        return matchers.get(sequence);
    }

    /**
     * 向最后添加一个匹配器
     *
     * @param matcher 匹配器对象
     */
    public void putMatcher(M matcher) {
        Integer seq = getMaxSequence() + 1;
        sequences.add(seq);
        matchers.put(seq, matcher);
    }

    /**
     * 添加指定序号的匹配器
     *
     * @param sequence 输入序号
     * @param matcher  匹配器对象
     */
    public void putMatcher(Integer sequence, M matcher) {
        sequences.add(sequence);
        Collections.sort(sequences);
        matchers.put(sequence, matcher);
    }

    /**
     * 返回最小序号
     *
     * @return
     */
    public Integer getMinSequence() {
        return sequences.size() == 0 ? 0 : sequences.get(0);
    }

    /**
     * 返回最大序号
     *
     * @return
     */
    public Integer getMaxSequence() {
        return sequences.size() == 0 ? 0 : sequences.get(sequences.size() - 1);
    }

}
