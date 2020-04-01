package com.jzy.util;


import com.jzy.excel.exception.InvalidParameterException;
import lombok.Data;

import java.util.ArrayList;
import java.util.List;

/**
 * @ClassName VCardUtils
 * @Author JinZhiyun
 * @Description 通讯录vcard处理工具类
 * @Date 2020/3/16 10:38
 * @Version 1.0
 **/
public class VCardUtils {
    @Data
    public static class VCard {
        /**
         * 联系人姓名
         */
        private String n;

        /**
         * 联系人电话
         */
        private String tel;

        /**
         * 联系人邮箱
         */
        private String email;

        public VCard(String n, String tel, String email) {
            this.n = n;
            this.tel = tel;
            this.email = email;
        }

        /**
         * 转成Vcard文件格式的字符串形式
         * 如：
         * BEGIN:VCARD
         * VERSION:3.0
         * N:张三
         * TEL:13088886666
         * EMAIL:aa
         * END:VCARD
         *
         * @return vcard字符串
         */
        public String toVCardString() {
            StringBuilder stringBuilder = new StringBuilder();
            stringBuilder.append("BEGIN:VCARD").append("\r\n")
                    .append("VERSION:3.0").append("\r\n")
                    .append("N:").append(n == null ? "" : n).append("\r\n")
                    .append("TEL:").append(tel == null ? "" : tel).append("\r\n")
                    .append("EMAIL:").append(email == null ? "" : email).append("\r\n")
                    .append("END:VCARD").append("\r\n");
            return stringBuilder.toString();
        }
    }

    public static VCard newVCardInstance(String n, String tel, String email) {
        return new VCard(n, tel, email);
    }

    /**
     * 输入一系列vcard对象，输出vard格式的字串
     *
     * @param vCards vcard对象的列表
     * @return vard格式的字串
     */
    public static String format(List<VCard> vCards) {
        StringBuilder stringBuilder = new StringBuilder();
        if (vCards != null) {
            for (VCard vCard : vCards) {
                stringBuilder.append(vCard.toVCardString());
            }
        }
        return stringBuilder.toString();
    }

    /**
     * 输入姓名、电话的列表，输出vard格式的字串
     *
     * @param ns   姓名集合
     * @param tels 电话集合
     * @return vard格式的字串
     */
    public static String format(List<String> ns, List<String> tels) {
        if (ns == null || tels == null) {
            throw new InvalidParameterException("输入的list不能为空");
        }
        return format(ns, tels, new ArrayList<>(ns.size()));
    }

    /**
     * 输入姓名、电话、邮箱的列表，输出vard格式的字串
     *
     * @param ns     姓名集合
     * @param tels   电话集合
     * @param emails 邮箱集合
     * @return vard格式的字串
     */
    public static String format(List<String> ns, List<String> tels, List<String> emails) {
        StringBuilder stringBuilder = new StringBuilder();
        if (ns == null || tels == null || emails == null) {
            throw new InvalidParameterException("输入的list不能为空");
        }
        if (ns.size() != tels.size() || tels.size() != emails.size()) {
            throw new InvalidParameterException("输入的list长度必须都相等");
        }
        int size = ns.size();
        for (int i = 0; i < size; i++) {
            VCard vCard = new VCard(ns.get(i), tels.get(i), emails.get(i));
            stringBuilder.append(vCard.toVCardString());
        }
        return stringBuilder.toString();
    }
}
