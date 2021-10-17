package cn.sanchar.excel.annotation;

import java.lang.annotation.ElementType;
import java.lang.annotation.Retention;
import java.lang.annotation.RetentionPolicy;
import java.lang.annotation.Target;

/**
 * description: 启用Excel导入
 *
 * @author shencai.huang@hand-china.com
 * @date 2021/9/27 4:42 下午
 * lastUpdateBy: shencai.huang@hand-china.com
 * lastUpdateDate: 2021/9/27
 */
@Target(ElementType.TYPE)
@Retention(RetentionPolicy.RUNTIME)
public @interface EnableExcelImport {
    /**
     * sheet 页名-列表
     */
    String[] sheetNames() default {};

    /**
     * sheet 索引-列表
     */
    int[] sheetIndexes() default {};

    /**
     * 开始行-列表
     */
    int[] startRowIndexes() default {};
}
