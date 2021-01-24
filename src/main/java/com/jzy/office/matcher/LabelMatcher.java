package com.jzy.office.matcher;

import java.util.HashMap;
import java.util.regex.Matcher;

/**
 * @ClassName LabelMatcher
 * @Author JinZhiyun
 * @Description 抽象标签匹配器
 * @Date 2021/1/16 16:44
 * @Version 1.0
 **/
public abstract class LabelMatcher extends StringMatcher {
    /**
     * replaceAllLabels方法所返回的最终结果中，替换的target字符串的值对应存于该键
     */
    public static final String OUTPUT = "{{OUTPUT}}";

    public LabelMatcher(String regex) {
        super(regex);
    }

    /**
     * 如果字串中的标签bookmark中没有，将当前标签替换成空串
     * 如果bookmark中的标签当前字串中没有，不做任何处理
     * 返回在bookmark和target中共有的被正确替换掉的标签集合，以及最终替换结果（最终结果的键为{@link LabelMatcher#OUTPUT}）
     * <p>
     * 例子❶：如当前匹配标签为：${标签1}。替换中指定的字符串中的“${label1}”
     * LabelMatcher matcher= LabelMatchers.DEFAULT_LABEL_MATCHER;
     * String target = "习近平${label1}在讲话中${label2}强调";
     * Map<String, String> bookmark = new HashMap<>();
     * bookmark.put("label0", "0000");
     * bookmark.put("label1", "1111");
     * System.out.println(matcher.replaceAllLabels(target, bookmark));
     * 输出：习近平1111在讲话中强调
     *
     * @param target   目标字串
     * @param bookmark 准备替换的标签键值对集合
     * @return 返回在bookmark和target中共有的被正确替换掉的标签集合，以及最终替换结果
     */
    public HashMap<String, String> replaceAllLabels(String target, HashMap<String, String> bookmark) {
        //记录最后target中所有被正确替换掉的bookmark中的key
        HashMap<String, String> replacedBookmarkWithOutput = new HashMap<>();
        if (bookmark == null || bookmark.size() == 0) {
            return replacedBookmarkWithOutput;
        }
        Matcher matcher;
        while ((matcher = getMatcher(target)).find()) {
            //依次获得${labelKey}中的labelKey
            String labelKey = matcher.group(1);
            if (bookmark.containsKey(labelKey)) {
                //书签集中有当前找到的labelKey
                String replacement = bookmark.get(labelKey);
                //将当前${labelKey}替换为replacement
                target = matcher.replaceFirst(replacement);
                //添加到成功替换的labelKey集合
                replacedBookmarkWithOutput.put(labelKey, replacement);
            } else {
                target = matcher.replaceFirst("");
            }
        }
        //添加整体替换后的结果
        replacedBookmarkWithOutput.put(OUTPUT, target);
        return replacedBookmarkWithOutput;
    }
}
