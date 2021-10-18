package cn.sanchar.excel.util;

import java.util.List;

/**
 * description: 数组工具类
 *
 * @author shencai.huang@hand-china.com
 * @date 2021/10/14 11:32 上午
 * lastUpdateBy: shencai.huang@hand-china.com
 * lastUpdateDate: 2021/10/14
 */
public class ArrayUtil {

    /**
     * 判断数组是否为空 - Object
     *
     * @param arr
     * @return 是否为空
     */
    public static boolean isEmpty(Object[] arr) {
        return arr == null || arr.length == 0;
    }

    /**
     * 判断数组是否不为空 - Object
     *
     * @param arr
     * @return 是否不为空
     */
    public static boolean isNotEmpty(Object[] arr) {
        return arr != null && arr.length > 0;
    }

    /**
     * 判断数组是否为空 - int
     *
     * @param arr
     * @return 是否为空
     */
    public static boolean isEmpty(int[] arr) {
        return arr.length == 0;
    }

    /**
     * 判断数组是否不为空 - int
     *
     * @param arr
     * @return 是否不为空
     */
    public static boolean isNotEmpty(int[] arr) {
        return arr.length > 0;
    }

    /**
     * copy 数据到新的数组 - String
     *
     * @param source 原数组
     * @param length 新数组大小
     * @return 新数组
     */
    public static String[] copyNewArr(String[] source, int length) {
        String[] newArr = new String[length];
        for (int i = 0; i < newArr.length; i++) {
            if (i > source.length - 1) {
                break;
            }
            newArr[i] = source[i];
        }
        return newArr;
    }

    /**
     * copy 数据到新的数组 - boolean
     *
     * @param source 原数组
     * @param length 新数组大小
     * @return 新数组
     */
    public static boolean[] copyNewArr(boolean[] source, int length) {
        boolean[] newArr = new boolean[length];
        for (int i = 0; i < newArr.length; i++) {
            if (i > source.length - 1) {
                break;
            }
            newArr[i] = source[i];
        }
        return newArr;
    }

    /**
     * copy 数据到新的数组 - int
     *
     * @param source 原数组
     * @param length 新数组大小
     * @return 新数组
     */
    public static int[] copyNewArr(int[] source, int length) {
        int[] newArr = new int[length];
        for (int i = 0; i < newArr.length; i++) {
            if (i > source.length - 1) {
                break;
            }
            newArr[i] = source[i];
        }
        return newArr;
    }

    /**
     * copy 数据到新的数组 - short
     *
     * @param source 原数组
     * @param length 新数组大小
     * @return 新数组
     */
    public static short[] copyNewArr(short[] source, int length) {
        short[] newArr = new short[length];
        for (int i = 0; i < newArr.length; i++) {
            if (i > source.length - 1) {
                break;
            }
            newArr[i] = source[i];
        }
        return newArr;
    }

    /**
     * 集合转数组
     *
     * @param list 转换的集合
     * @param <T>  集合对象类型
     * @return 数组
     */
    public static <T> T[] parseArray(List<T> list) {
        return (T[]) list.toArray();
    }
}
