package per.sanchar.excel.annotation;

import java.lang.annotation.ElementType;
import java.lang.annotation.Retention;
import java.lang.annotation.RetentionPolicy;
import java.lang.annotation.Target;

/**
 * description: Excel 导入导出列
 *
 * @author shencai.huang@hand-china.com
 * @date 2021/9/28 11:23 上午
 * lastUpdateBy: shencai.huang@hand-china.com
 * lastUpdateDate: 2021/9/28
 */
@Target(ElementType.FIELD)
@Retention(RetentionPolicy.RUNTIME)
public @interface SheetColumn {
    /**
     * 表头列名
     */
    String name();
    /**
     * 列索引
     */
    int index() default -1;
    /**
     * 列宽
     */
    int width() default 4;
    /**
     * 必输
     */
    boolean required() default false;
    /**
     * 导入字段
     */
    boolean imported() default true;
    /**
     * 导出字段
     */
    boolean exported() default true;
}
