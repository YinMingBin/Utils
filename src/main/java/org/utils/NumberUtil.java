package org.utils;

/**
 * 操作数字工具类
 *
 * @author YinMingBin
 */
public class NumberUtil {

    /**
     * 字符串转 int
     * @param s 字符串
     * @return int
     */
    public static int toInt(String s) {
        try {
            return Integer.parseInt(s);
        }catch (NumberFormatException numberFormatException){
            System.err.println("Int-数字转换异常：" + s);
        }
        return 0;
    }

    /**
     * 字符串转 long
     * @param s 字符串
     * @return long
     */
    public static long toLong(String s) {
        try {
            return Long.parseLong(s);
        }catch (NumberFormatException numberFormatException){
            System.err.println("Long-数字转换异常：" + s);
        }
        return 0L;
    }

    /**
     * 字符串转 double
     * @param s 字符串
     * @return double
     */
    public static double toDouble(String s) {
        try {
            return Double.parseDouble(s);
        }catch (NumberFormatException numberFormatException){
            System.err.println("Double-数字转换异常：" + s);
        }
        return 0.0;
    }

    /**
     * 字符串转 float
     * @param s 字符串
     * @return float
     */
    public static double toFloat(String s) {
        try {
            return Float.parseFloat(s);
        }catch (NumberFormatException numberFormatException){
            System.err.println("Float-数字转换异常：" + s);
        }
        return 0.0;
    }

}
