package cn.sanchar.excel.util;

import org.checkerframework.checker.nullness.qual.Nullable;

/**
 * description:
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
}
