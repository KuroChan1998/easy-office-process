package com.jzy.util;

import org.apache.commons.lang3.time.DateUtils;

import java.lang.management.ManagementFactory;
import java.text.ParseException;
import java.text.ParsePosition;
import java.text.SimpleDateFormat;
import java.util.*;

/**
 * @author JinZhiyun
 * @author ruoyi
 * @ClassName MyTimeUtils
 * @Description 计算时间工具类
 * @Date 2019/6/6 9:17
 * @Version 1.0
 **/
public class MyTimeUtils {
    public static final String FORMAT_YMD = "yyyy-MM-dd";

    public static final String FORMAT_YMDHMS = "yyyy-MM-dd HH:mm:ss";

    public static final String FORMAT_YMDHMS_BACKUP = "yyyy/MM/dd HH:mm:ss";

    public static final long VALID_TIME_3_MIN = 180000;  //3分钟,180s
    public static final long VALID_TIME_5_MIN = 300000;  //5分钟,500s
    public static final long VALID_TIME_10_MIN = 600000;  //10分钟,600s

    private MyTimeUtils() {
    }

    /**
     * @return long
     * @author JinZhiyun
     * @Description 获取当前时间，返回毫秒级时间
     * @Date 10:20 2019/6/6
     * @Param []
     **/
    public static long getTime() {
        //在获取现在的时间
        Calendar calendar = Calendar.getInstance();
        long date = calendar.getTime().getTime();            //获取毫秒时间
        return date;
    }

    /**
     * @return boolean
     * @author JinZhiyun
     * @Description 比较当前时间与指定时间timeComparedTo差值是否超过了validTime
     * @Date 10:20 2019/6/6
     * @Param [timeComparedTo, validTimeMilliseconds]
     **/
    public static boolean cmpTime(long timeComparedTo, long validTime) {
        //在获取现在的时间
        Calendar calendar = Calendar.getInstance();
        long timeNow = calendar.getTime().getTime();            //获取毫秒时间
        if (timeNow - timeComparedTo > validTime) {
            return false;
        } else {
            return true;
        }
    }

    /**
     * @return int
     * @author JinZhiyun
     * @Description 根据生日计算年龄
     * @Date 22:48 2019/6/19
     * @Param [birthDay]
     **/
    public static int getAgeByBirth(Date birthDay) {
        if (birthDay == null) {
            return -1;
        }
        int age = 0;
        Calendar cal = Calendar.getInstance();
        if (cal.before(birthDay)) {
            //出生日期晚于当前时间，无法计算
            return -1;
        }
        int yearNow = cal.get(Calendar.YEAR);  //当前年份
        int monthNow = cal.get(Calendar.MONTH);  //当前月份
        int dayOfMonthNow = cal.get(Calendar.DAY_OF_MONTH); //当前日期
        cal.setTime(birthDay);
        int yearBirth = cal.get(Calendar.YEAR);
        int monthBirth = cal.get(Calendar.MONTH);
        int dayOfMonthBirth = cal.get(Calendar.DAY_OF_MONTH);
        age = yearNow - yearBirth;   //计算整岁数
        if (monthNow <= monthBirth) {
            if (monthNow == monthBirth) {
                if (dayOfMonthNow < dayOfMonthBirth) {
                    age--;//当前日期在生日之前，年龄减一
                }
            } else {
                age--;//当前月份在生日之前，年龄减一
            }
        }
        return age;
    }

    /**
     * @return java.lang.String
     * @author JinZhiyun
     * @Description 将短时间格式时间转换为字符串 yyyy-MM-dd
     * @Date 17:26 2019/6/23
     * @Param [dateDate]
     **/
    public static String dateToStringYMD(Date date) {
        return dateToString(date, FORMAT_YMD);
    }

    /**
     * @return java.lang.String
     * @author JinZhiyun
     * @Description 将短时间格式时间转换为字符串 yyyy-MM-dd HH:mm:ss
     * @Date 17:26 2019/6/23
     * @Param [dateDate]
     **/
    public static String dateToStringYMDHMS(Date date) {
        return dateToString(date, FORMAT_YMDHMS);
    }

    /**
     * 将短时间格式时间转换为字符串，手动指定格式
     *
     * @param date      Date对象
     * @param formatStr format格式化字符串
     * @return 字符串形式的date
     */
    public static String dateToString(Date date, String formatStr) {
        SimpleDateFormat formatter = new SimpleDateFormat(formatStr);
        String dateString = formatter.format(date);
        return dateString;
    }

    /**
     * @return java.util.Date
     * @author JinZhiyun
     * @Description 将短时间格式字符串转换为时间 yyyy-MM-dd
     * @Date 17:27 2019/6/23
     * @Param [strDate]
     **/
    public static Date stringToDateYMD(String strDate) {
        return stringToDate(strDate, FORMAT_YMD);
    }

    /**
     * @return java.util.Date
     * @author JinZhiyun
     * @Description 将短时间格式字符串转换为时间 yyyy-MM-dd HH:mm:ss
     * @Date 17:27 2019/6/23
     * @Param [strDate]
     **/
    public static Date stringToDateYMDHMS(String strDate) {
        return stringToDate(strDate, FORMAT_YMDHMS);
    }

    /**
     * 将短时间格式字符串转换为时间，手动指定格式
     *
     * @param strDate   短时间格式字符串
     * @param formatStr format格式化字符串
     * @return date时间
     */
    public static Date stringToDate(String strDate, String formatStr) {
        SimpleDateFormat formatter = new SimpleDateFormat(formatStr);
        ParsePosition pos = new ParsePosition(0);
        Date strtodate = formatter.parse(strDate, pos);
        return strtodate;
    }

    /**
     * 将cst时间转为Date
     *
     * @param cst cst时间字符串
     * @return Date对象
     * @throws ParseException 格式化失败
     */
    public static Date cstToDate(String cst) throws ParseException {
        SimpleDateFormat sdf = new SimpleDateFormat("EEE MMM dd HH:mm:ss zzz yyyy", Locale.US);
        return sdf.parse(cst);
    }

    /**
     * 获取当前年份
     *
     * @return 当前年份-整数
     */
    public static int getCurrentYear() {
        Calendar cal = Calendar.getInstance();
        int year = cal.get(Calendar.YEAR);
        return year;
    }

    /**
     * 获取当前月份
     *
     * @return 当前月份-整数
     */
    public static int getCurrentMonth() {
        Calendar cal = Calendar.getInstance();
        int month = cal.get(Calendar.MONTH) + 1;
        return month;
    }

    /**
     * 获取指定日期所在周的开始时间
     *
     * @return 对应周的开始时间
     */
    public static Date getBeginDayOfWeek(Date date) {
        // 使用Calendar类进行时间的计算
        Calendar cal = Calendar.getInstance();
        cal.setTime(stringToDateYMD(dateToStringYMD(date)));
        // 获得当前日期是一个星期的第几天
        // 判断要计算的日期是否是周日，如果是则减一天计算周六的，否则会计算到下一周去。
        // dayWeek值（1、2、3...）对应周日，周一，周二...
        int dayWeek = cal.get(Calendar.DAY_OF_WEEK);
        if (dayWeek == 1) {
            dayWeek = 7;
        } else {
            dayWeek -= 1;
        }
        // 计算本周开始的时间
        cal.add(Calendar.DAY_OF_MONTH, 1 - dayWeek);
        Date startDate = cal.getTime();
        return startDate;
    }

    /**
     * 获取本周的开始时间
     *
     * @return 本周的开始时间
     */
    public static Date getBeginDayOfCurrentWeek() {
        return getBeginDayOfWeek(new Date());
    }

    /**
     * 获取指定日期所在月的开始时间
     *
     * @return 对应月的开始时间
     */
    public static Date getBeginDayOfMonth(Date date) {
        Calendar cal = Calendar.getInstance();
        cal.setTime(stringToDateYMD(dateToStringYMD(date)));
        // 设置为1号,当前日期既为本月第一天
        cal.set(Calendar.DAY_OF_MONTH, 1);
        Date startDate = cal.getTime();
        return startDate;
    }

    /**
     * 获取本月的开始时间
     *
     * @return 本月的开始时间
     */
    public static Date getBeginDayOfCurrentMonth() {
        return getBeginDayOfMonth(new Date());
    }


    /**
     * 获取当前日
     *
     * @return 当前日期-整数
     */
    public static int getCurrentDay() {
        Calendar cal = Calendar.getInstance();
        int day = cal.get(Calendar.DATE);
        return day;
    }

    /**
     * 获取当前日
     *
     * @return 当前日期-整数
     */
    public static int getCurrentHour() {
        Calendar cal = Calendar.getInstance();
        int hour = cal.get(Calendar.HOUR_OF_DAY);
        return hour;
    }


    /**
     * 获得date下一天的日期
     *
     * @param date 指定日期
     * @return 指定日期下一天的日期
     */
    public static Date getNextDay(Date date) {
        return getFutureDay(date, 1);
    }

    /**
     * 获取date的daysInterval天后的日期
     *
     * @param date         指定date
     * @param daysInterval 多少天后
     * @return daysInterval天后的日期
     */
    public static Date getFutureDay(Date date, int daysInterval) {
        if (date == null) {
            return null;
        }
        Calendar c = Calendar.getInstance();
        c.setTime(date);
        int day1 = c.get(Calendar.DATE);
        c.set(Calendar.DATE, day1 + daysInterval);

        return c.getTime();
    }

    /**
     * 获取date的daysInterval天后的日期到date之间的所有date
     *
     * @param date         指定date
     * @param daysInterval 多少天后
     * @return daysInterval天后的日期到date之间的所有date
     */
    public static List<Date> getFutureDays(Date date, int daysInterval) {
        Date mostFutureDay = getFutureDay(date, daysInterval);
        return getDaysBetween(date, mostFutureDay);
    }


    /**
     * 获取指定天的前一天
     *
     * @param date 指定日期
     * @return 指定日期前一天的日期
     */
    public static Date getLastDay(Date date) {
        return getFutureDay(date, -1);
    }

    /**
     * 获取date的daysInterval天前的日期
     *
     * @param date         指定date
     * @param daysInterval 多少天前
     * @return daysInterval天前的日期
     */
    public static Date getPastDay(Date date, int daysInterval) {
        return getFutureDay(date, 0 - daysInterval);
    }

    /**
     * 获取date的daysInterval天前的日期到date之间的所有date
     *
     * @param date         指定date
     * @param daysInterval 多少天前
     * @return daysInterval天前的日期到date之间的所有date
     */
    public static List<Date> getPastDays(Date date, int daysInterval) {
        Date mostPastDay = getFutureDay(date, 0 - daysInterval);
        return getDaysBetween(mostPastDay, date);
    }

    /**
     * 判断d1和d2是否是同一天
     *
     * @param d1 日期1
     * @param d2 日期2
     * @return 是否是同一天
     */
    public static boolean isSameDay(Date d1, Date d2) {
        return DateUtils.isSameDay(d1, d2);
    }

    /**
     * 获取startDay和endDay之间的所有日期
     *
     * @param startDay 起始天
     * @param endDay   结束天
     * @return 区间的所有日期
     */
    public static List<Date> getDaysBetween(Date startDay, Date endDay) {
        List<Date> days = new ArrayList<>();

        if (startDay == null) {
            return days;
        }

        if (endDay == null) {
            //endDay为空，默认只返回startDay
            days.add(startDay);
            return days;
        }

        if (endDay.before(startDay)) {
            //如果终止日期小于起始日期，返回空列表
            return days;
        }

        Date d = startDay;
        do {
            days.add(d);
            d = getNextDay(d);
        } while (!d.after(endDay));

        return days;
    }

    /**
     * 获取当前时间的secondsInterval秒后的时间
     *
     * @param secondsInterval 多少秒后
     * @return secondsInterval秒后的时间
     */
    public static Date getSecondsAfter(int secondsInterval) {
        return getSecondsAfter(new Date(), secondsInterval);
    }

    /**
     * 获取date的secondsInterval秒后的时间
     *
     * @param date            指定date
     * @param secondsInterval 多少秒后
     * @return secondsInterval秒后的时间
     */
    public static Date getSecondsAfter(Date date, int secondsInterval) {
        if (date == null) {
            return null;
        }
        Calendar calendar = Calendar.getInstance();
        calendar.setTime(date);
        calendar.add(Calendar.SECOND, secondsInterval);
        return calendar.getTime();
    }

    /**
     * 获取当前时间的secondsInterval秒前的时间
     *
     * @param secondsInterval 多少秒前
     * @return secondsInterval秒前的时间
     */
    public static Date getSecondsBefore(int secondsInterval) {
        return getSecondsBefore(new Date(), secondsInterval);
    }

    /**
     * 获取date的secondsInterval秒前的时间
     *
     * @param date            指定date
     * @param secondsInterval 多少秒前
     * @return secondsInterval秒前的时间
     */
    public static Date getSecondsBefore(Date date, int secondsInterval) {
        return getSecondsAfter(date, 0 - secondsInterval);
    }

    /**
     * 获取服务器启动时间
     *
     * @return
     */
    public static Date getServerStartDate() {
        long time = ManagementFactory.getRuntimeMXBean().getStartTime();
        return new Date(time);
    }

    /**
     * 计算两个时间差
     *
     * @param endDate 结束时间
     * @param nowDate 起始时间
     * @return 字符串描述的时间差
     */
    public static String getDatePoor(Date endDate, Date nowDate) {
        long nd = 1000 * 24 * 60 * 60;
        long nh = 1000 * 60 * 60;
        long nm = 1000 * 60;
        // long ns = 1000;
        // 获得两个时间的毫秒时间差异
        long diff = endDate.getTime() - nowDate.getTime();
        // 计算差多少天
        long day = diff / nd;
        // 计算差多少小时
        long hour = diff % nd / nh;
        // 计算差多少分钟
        long min = diff % nd % nh / nm;
        // 计算差多少秒//输出结果
        // long sec = diff % nd % nh % nm / ns;
        return day + "天" + hour + "小时" + min + "分钟";
    }

    public static void main(String[] args) {
        System.out.println(dateToStringYMDHMS(getPastDay(new Date(), 7)));
    }
}
