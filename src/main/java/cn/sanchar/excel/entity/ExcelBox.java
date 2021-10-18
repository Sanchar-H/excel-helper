package cn.sanchar.excel.entity;

import cn.sanchar.excel.util.StringUtils;

import java.util.Arrays;
import java.util.Objects;

/**
 * description: excel 下拉框实体
 *
 * @author shencai.huang@hand-china.com
 * @date 2021/10/18 12:06 下午
 * lastUpdateBy: shencai.huang@hand-china.com
 * lastUpdateDate: 2021/10/18
 */
public class ExcelBox {

    /**
     * 数据集标题名称
     */
    private String boxName;
    /**
     * sheet页名
     */
    private String sheetName;
    /**
     * sheet页索引
     */
    private Integer sheetIndex;
    /**
     * 下拉框的取值表达式
     */
    private String formula;
    /**
     * 下拉框的选项值
     */
    private String[] values;
    /**
     * 起始行
     */
    private Integer firstRow;
    /**
     * 终止行
     */
    private Integer lastRow;
    /**
     * 起始列
     */
    private Integer firstCol;
    /**
     * 终止列
     */
    private Integer lastCol;
    /**
     * 是否显示提示
     */
    private Boolean isShowTips;
    /**
     * 提示标题
     */
    private String tipsTitle;
    /**
     * 提示内容
     */
    private String tipsMessage;
    /**
     * 错误提示标题
     */
    private String errorTitle;
    /**
     * 错误提示内容
     */
    private String errorMessage;

    public ExcelBox() {
    }

    public String getBoxName() {
        return boxName;
    }

    public void setBoxName(String boxName) {
        this.boxName = boxName;
    }

    public String getSheetName() {
        return sheetName;
    }

    public void setSheetName(String sheetName) {
        this.sheetName = sheetName;
    }

    public Integer getSheetIndex() {
        return sheetIndex;
    }

    public void setSheetIndex(Integer sheetIndex) {
        this.sheetIndex = sheetIndex;
    }

    public String getFormula() {
        return formula;
    }

    public void setFormula(String formula) {
        this.formula = formula;
    }

    public String[] getValues() {
        return values;
    }

    public void setValues(String[] values) {
        this.values = values;
    }

    public Integer getFirstRow() {
        return firstRow;
    }

    public void setFirstRow(Integer firstRow) {
        this.firstRow = firstRow;
    }

    public Integer getLastRow() {
        return lastRow;
    }

    public void setLastRow(Integer lastRow) {
        this.lastRow = lastRow;
    }

    public Integer getFirstCol() {
        return firstCol;
    }

    public void setFirstCol(Integer firstCol) {
        this.firstCol = firstCol;
    }

    public Integer getLastCol() {
        return lastCol;
    }

    public void setLastCol(Integer lastCol) {
        this.lastCol = lastCol;
    }

    public Boolean getShowTips() {
        return isShowTips;
    }

    public void setShowTips(Boolean showTips) {
        isShowTips = showTips;
    }

    public String getTipsTitle() {
        return tipsTitle;
    }

    public void setTipsTitle(String tipsTitle) {
        this.tipsTitle = tipsTitle;
    }

    public String getTipsMessage() {
        return tipsMessage;
    }

    public void setTipsMessage(String tipsMessage) {
        this.tipsMessage = tipsMessage;
    }

    public String getErrorTitle() {
        return errorTitle;
    }

    public void setErrorTitle(String errorTitle) {
        this.errorTitle = errorTitle;
    }

    public String getErrorMessage() {
        return errorMessage;
    }

    public void setErrorMessage(String errorMessage) {
        this.errorMessage = errorMessage;
    }

    public ExcelBox boxName(String boxName) {
        this.boxName = boxName;
        return this;
    }

    public ExcelBox sheetName(String sheetName) {
        this.sheetName = sheetName;
        return this;
    }

    public ExcelBox sheetIndex(Integer sheetIndex) {
        this.sheetIndex = sheetIndex;
        return this;
    }

    public ExcelBox formula(String formula) {
        this.formula = formula;
        return this;
    }

    public ExcelBox values(String[] values) {
        this.values = values;
        return this;
    }

    public ExcelBox firstRow(Integer firstRow) {
        this.firstRow = firstRow;
        return this;
    }

    public ExcelBox lastRow(Integer lastRow) {
        this.lastRow = lastRow;
        return this;
    }

    public ExcelBox firstCol(Integer firstCol) {
        this.firstCol = firstCol;
        return this;
    }

    public ExcelBox lastCol(Integer lastCol) {
        this.lastCol = lastCol;
        return this;
    }

    public ExcelBox showTips(Boolean showTips) {
        isShowTips = showTips;
        return this;
    }

    public ExcelBox tipsTitle(String tipsTitle) {
        this.tipsTitle = tipsTitle;
        return this;
    }

    public ExcelBox tipsMessage(String tipsMessage) {
        this.tipsMessage = tipsMessage;
        return this;
    }

    public ExcelBox errorTitle(String errorTitle) {
        this.errorTitle = errorTitle;
        return this;
    }

    public ExcelBox errorMessage(String errorMessage) {
        this.errorMessage = errorMessage;
        return this;
    }

    @Override
    public boolean equals(Object o) {
        if (this == o) {
            return true;
        }
        if (o == null || getClass() != o.getClass()) {
            return false;
        }
        ExcelBox excelBox = (ExcelBox) o;
        return Objects.equals(boxName, excelBox.boxName)
                && Objects.equals(sheetName, excelBox.sheetName)
                && Objects.equals(sheetIndex, excelBox.sheetIndex)
                && Objects.equals(formula, excelBox.formula)
                && Arrays.equals(values, excelBox.values)
                && Objects.equals(firstRow, excelBox.firstRow)
                && Objects.equals(lastRow, excelBox.lastRow)
                && Objects.equals(firstCol, excelBox.firstCol)
                && Objects.equals(lastCol, excelBox.lastCol)
                && Objects.equals(isShowTips, excelBox.isShowTips)
                && Objects.equals(tipsTitle, excelBox.tipsTitle)
                && Objects.equals(tipsMessage, excelBox.tipsMessage)
                && Objects.equals(errorTitle, excelBox.errorTitle)
                && Objects.equals(errorMessage, excelBox.errorMessage);
    }

    @Override
    public int hashCode() {
        int result = Objects.hash(boxName, sheetName, sheetIndex, formula, firstRow, lastRow, firstCol, lastCol, isShowTips, tipsTitle, tipsMessage, errorTitle, errorMessage);
        result = 31 * result + Arrays.hashCode(values);
        return result;
    }

    @Override
    public String toString() {
        return "ExcelBox{" +
                "boxName='" + boxName + '\'' +
                ", sheetName='" + sheetName + '\'' +
                ", sheetIndex=" + sheetIndex +
                ", formula='" + formula + '\'' +
                ", values=" + Arrays.toString(values) +
                ", firstRow=" + firstRow +
                ", lastRow=" + lastRow +
                ", firstCol=" + firstCol +
                ", lastCol=" + lastCol +
                ", isShowTips=" + isShowTips +
                ", tipsTitle='" + tipsTitle + '\'' +
                ", tipsMessage='" + tipsMessage + '\'' +
                ", errorTitle='" + errorTitle + '\'' +
                ", errorMessage='" + errorMessage + '\'' +
                '}';
    }

    public String generateUniqueKey() {
        return boxName + (StringUtils.isEmpty(values) ? 0 : values.length) + Arrays.toString(values).length();
    }
}
