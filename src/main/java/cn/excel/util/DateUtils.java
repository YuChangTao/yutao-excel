package cn.excel.util;

import java.text.SimpleDateFormat;
import java.util.Date;

/**
 * @author yutyi
 * 日期工具类
 */
public class DateUtils extends org.apache.commons.lang.time.DateUtils {

	/**
	 * 转换成日期格式
	 * 
	 * @param str
	 * @return
	 */
	public static Date toTimestamp(String str) {
		String[] patterns = new String[] { "yyyy-MM-dd HH:mm:ss", "yyyy-MM-dd HH:mm", "yyyy-MM-dd", "HH:mm:ss", "HH:mm", "yyyy/M/d H:m:s", "HH:mm", "yyyy/M/d H:m", "yyyy/M/d",
				"H:m:s", "H:m", "yyyy-MM-dd HH:mm:ss.SSS", "yyyy年M月d日" };
		try {
			Date date = org.apache.commons.lang.time.DateUtils.parseDate(str, patterns);
			return new java.sql.Timestamp(date.getTime());
		} catch (Exception ex) {
			return null;
		}
	}

    /**
     * 字符串转java.sql.Date
     * @param str
     * @return
     */
	public static java.sql.Date toDate(String str) {
		Date date = toTimestamp(str);
		return date == null ? null : new java.sql.Date(date.getTime());
	}

    /**
     * 字符串转java.sql.Time
     * @param str
     * @return
     */
	public static java.sql.Time toTime(String str) {
		Date date = toTimestamp(str);
		return date == null ? null : new java.sql.Time(date.getTime());
	}

    /**
     * 日期格式化成指定字符串格式
     * @param date
     * @param pattern
     * @return
     */
    public String format(Date date, String pattern) {
        return new SimpleDateFormat(pattern).format(date);
    }
}
