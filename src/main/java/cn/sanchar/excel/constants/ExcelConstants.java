package cn.sanchar.excel.constants;

/**
 * description: 公用静态常量
 *
 * @author shencai.huang@hand-china.com
 * @date 2021/10/13 5:45 下午
 * lastUpdateBy: shencai.huang@hand-china.com
 * lastUpdateDate: 2021/10/13
 */
public class ExcelConstants {

    /**
     * Excel 文件格式 xls
     */
    public static final String EXCEL_XLS = "xls";

    /**
     * Excel Sheet 页名前缀
     */
    public static final String SHEET_PREFIX = "Sheet";

    /**
     * 隐藏Sheet页默认名称
     */
    public static final String HIDDEN_SHEET_NAME = "@<excel.hidden.data>";

    /**
     * excel下拉框取之范围表达式格式
     */
    public static final String FORMULA_FORMAT = "'%s'!$%s$2:$%s$%d";

    /**
     * 字符 A
     */
    public static final Character A = 'A';

    /**
     * 下拉框默认提示标题
     */
    public static final String DEFAULT_TIPS_TITLE = "提示";

    /**
     * 下拉框默认提示消息
     */
    public static final String DEFAULT_TIPS_MESSAGE = "只能选择下拉框里面的数据";

    /**
     * 下拉框默认错误提示标题
     */
    public static final String DEFAULT_ERROR_TITLE = "错误";

    /**
     * 下拉框默认错误提示消息
     */
    public static final String DEFAULT_ERROR_MESSAGE = "请选择下拉框的值";

}
