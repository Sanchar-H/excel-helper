package per.sanchar.excel.annotation;

import java.lang.annotation.ElementType;
import java.lang.annotation.Retention;
import java.lang.annotation.RetentionPolicy;
import java.lang.annotation.Target;

/**
 * description: 启用Excel导出
 *
 * @author shencai.huang@hand-china.com
 * @date 2021/9/27 4:42 下午
 * lastUpdateBy: shencai.huang@hand-china.com
 * lastUpdateDate: 2021/9/27
 */
@Target(ElementType.TYPE)
@Retention(RetentionPolicy.RUNTIME)
public @interface EnableExcelExport {
    /**
     * 文件格式
     */
    String fileType() default "xlsx";
    /**
     * sheet 页名-列表
     */
    String[] sheetNames() default {};
    /**
     * 是否需要表头-列表
     */
    boolean[] isIncludeHeaders() default {};
    /**
     * 开始行-列表
     */
    int[] startRowIndexes() default {};
    /**
     * 开始列-列表
     */
    short[] startColumnIndexes() default {};
}
