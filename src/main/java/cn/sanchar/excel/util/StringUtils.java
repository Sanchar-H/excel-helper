package cn.sanchar.excel.util;

import org.checkerframework.checker.nullness.qual.Nullable;

/**
 * description: 字符串工具类
 *
 * @author shencai.huang@hand-china.com
 * @date 2021/10/14 10:21 下午
 * lastUpdateBy: shencai.huang@hand-china.com
 * lastUpdateDate: 2021/10/14
 */
public abstract class StringUtils {

    public StringUtils() {
    }

    public static boolean isEmpty(@Nullable Object str) {
        return str == null || "".equals(str);
    }

    public static boolean isNotEmpty(@Nullable Object str) {
        return str != null && !"".equals(str);
    }

    public static int calStringArrTotalLength(String[] arr) {
        if (ArrayUtil.isEmpty(arr)) {
            return 0;
        }
        StringBuilder sb = new StringBuilder();
        for (String s : arr) {
            sb.append(s).append(" ");
        }
        return sb.length() - 1;
    }
}
