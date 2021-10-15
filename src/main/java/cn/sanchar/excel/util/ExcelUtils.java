package cn.sanchar.excel.util;

import cn.sanchar.excel.annotation.EnableExcelExport;
import cn.sanchar.excel.annotation.EnableExcelImport;
import cn.sanchar.excel.annotation.SheetColumn;
import cn.sanchar.excel.constants.ExcelConstants;
import cn.sanchar.excel.exception.ExcelHandleException;
import com.alibaba.fastjson.util.TypeUtils;
import org.apache.commons.compress.utils.Lists;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.FillPatternType;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.apache.poi.xssf.usermodel.XSSFRichTextString;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.checkerframework.checker.nullness.qual.NonNull;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import java.io.IOException;
import java.io.InputStream;
import java.io.OutputStream;
import java.lang.reflect.Field;
import java.util.Arrays;
import java.util.Calendar;
import java.util.Date;
import java.util.Iterator;
import java.util.List;
import java.util.Objects;
import java.util.Optional;
import java.util.stream.Collectors;
import java.util.stream.IntStream;

/**
 * description: Excel 导入导出工具类
 *
 * @author shencai.huang@hand-china.com
 * @date 2021/9/27 5:15 下午
 * lastUpdateBy: shencai.huang@hand-china.com
 * lastUpdateDate: 2021/9/27
 */
public abstract class ExcelUtils {

    private static final Logger LOGGER = LoggerFactory.getLogger(ExcelUtils.class);

    public ExcelUtils() {}

    /**
     * 将excel的指定单个sheet页数据转换为对象集合
     *
     * @param inputStream excel文件流
     * @param clazz       pojo类型
     * @param <T>         泛型对象
     * @return 对象集合 List<T>
     */
    public static <T> List<T> singleParseToList(InputStream inputStream, @NonNull Class<T> clazz) {
        Workbook workbook;
        try {
            workbook = WorkbookFactory.create(inputStream);
        } catch (IOException e) {
            LOGGER.error("error.", e);
            throw new ExcelHandleException("inputStream can not to create a workbook.");
        }
        return singleParseToList(workbook, clazz);
    }

    /**
     * 将excel的指定单个sheet页数据转换为对象集合
     *
     * @param workbook 工作簿
     * @param clazz    pojo类型
     * @param <T>      泛型对象
     * @return 对象集合 List<T>
     */
    public static <T> List<T> singleParseToList(@NonNull Workbook workbook, @NonNull Class<T> clazz) {
        validateObject(workbook, clazz);
        EnableExcelImport enableExcelImport = clazz.getAnnotation(EnableExcelImport.class);
        if (Objects.isNull(enableExcelImport)) {
            throw new ExcelHandleException(String.format("class '%s' not support excel import.", clazz.getName()));
        }
        Sheet[] sheets = new Sheet[1];
        // 先通过sheet页名匹配sheet页
        String[] sheetNames = enableExcelImport.sheetNames();
        if (ArrayUtil.isNotEmpty(sheetNames)) {
            String sheetName = sheetNames[0];
            sheets[0] = Optional.ofNullable(workbook.getSheet(sheetName)).orElseThrow(()
                    -> new ExcelHandleException(String.format("sheetName [%s] can not be matched a available sheet.", sheetName)));
        }
        int numberOfSheets = workbook.getNumberOfSheets();
        // 未配置sheet页名则通过sheet索引匹配sheet页
        int[] sheetAts = enableExcelImport.sheetAts();
        if (ArrayUtil.isEmpty(sheetNames) && ArrayUtil.isNotEmpty(sheetAts)) {
            int sheetAt = sheetAts[0];
            if (sheetAt > numberOfSheets - 1) {
                throw new ExcelHandleException(String.format("sheetIndex [%d] is out of range[0, %d).", sheetAt, numberOfSheets));
            }
            sheets[0] = workbook.getSheetAt(sheetAt);

        }
        // 如果sheet页名和sheet索引均未配置有，则默认取第一个sheet页
        if (ArrayUtil.isEmpty(sheetNames) && ArrayUtil.isEmpty(sheetAts)) {
            sheets[0] = workbook.getSheetAt(0);
        }
        // 未配置起始行的sheet页默认第一行开始读数据
        int[] startRowIndexes = enableExcelImport.startRowIndexes();
        startRowIndexes = ArrayUtil.isEmpty(startRowIndexes) ? new int[1] : startRowIndexes;
        return parseSheetToList(sheets, startRowIndexes, clazz).get(0);
    }

    /**
     * 将excel的指定多个sheet页数据转换为对象集合列表
     *
     * @param inputStream excel文件流
     * @param clazz       pojo类型
     * @param <T>         泛型对象
     * @return 对象集合 List<List<T>>
     */
    public static <T> List<List<T>> parseToList(InputStream inputStream, Class<T> clazz) {
        Workbook workbook;
        try {
            workbook = WorkbookFactory.create(inputStream);
        } catch (IOException e) {
            LOGGER.error("error.", e);
            throw new ExcelHandleException("inputStream can not to create a workbook.");
        }
        return parseToList(workbook, clazz);
    }

    /**
     * 将excel的指定多个sheet页数据转换为对象集合列表
     *
     * @param workbook 工作簿
     * @param clazz    pojo类型
     * @param <T>      泛型对象
     * @return 对象集合 List<List<T>>
     */
    public static <T> List<List<T>> parseToList(@NonNull Workbook workbook, @NonNull Class<T> clazz) {
        validateObject(workbook, clazz);
        EnableExcelImport enableExcelImport = clazz.getAnnotation(EnableExcelImport.class);
        if (Objects.isNull(enableExcelImport)) {
            throw new ExcelHandleException(String.format("class '%s' not support excel import.", clazz.getName()));
        }
        Sheet[] sheets = null;
        // 先通过sheet页名匹配sheet页
        String[] sheetNames = enableExcelImport.sheetNames();
        if (ArrayUtil.isNotEmpty(sheetNames)) {
            sheets = new Sheet[sheetNames.length];
            for (int i = 0; i < sheetNames.length; i++) {
                String sheetName = sheetNames[i];
                sheets[i] = Optional.ofNullable(workbook.getSheet(sheetName)).orElseThrow(()
                        -> new ExcelHandleException(String.format("sheetName [%s] can not be matched a available sheet.", sheetName)));
            }
        }
        int numberOfSheets = workbook.getNumberOfSheets();
        // 未配置sheet页名则通过sheet索引匹配sheet页
        int[] sheetAts = enableExcelImport.sheetAts();
        if (ArrayUtil.isEmpty(sheets) && ArrayUtil.isNotEmpty(sheetAts)) {
            sheets = new Sheet[sheetAts.length];
            for (int i = 0; i < sheetAts.length; i++) {
                int sheetAt = sheetAts[i];
                if (sheetAt > numberOfSheets - 1) {
                    throw new ExcelHandleException(String.format("sheetIndex [%d] is out of range[0, %d).", sheetAt, numberOfSheets));
                }
                sheets[i] = workbook.getSheetAt(sheetAt);
            }
        }
        // 如果sheet页名和sheet索引均未配置有，则默认取第一个sheet页
        if (ArrayUtil.isEmpty(sheets)) {
            sheets = new Sheet[1];
            sheets[0] = workbook.getSheetAt(0);
        }
        // 未配置起始行的sheet页默认第一行开始读数据
        int[] startRowIndexes = ArrayUtil.copyNewArr(enableExcelImport.startRowIndexes(), sheets.length);

        return parseSheetToList(sheets, startRowIndexes, clazz);
    }

    /**
     * 校验对象空属性
     *
     * @param workbook 工作簿
     * @param clazz    pojo类型
     */
    private static <T> void validateObject(Workbook workbook, Class<T> clazz) {
        if (Objects.isNull(workbook)) {
            throw new ExcelHandleException("param [workbook] can not be null.");
        }
        if (Objects.isNull(clazz)) {
            throw new ExcelHandleException("param [clazz] can not be null.");
        }
    }

    /**
     * 将sheet页数据转换为对象集合列表
     *
     * @param sheets          sheet页数组
     * @param startRowIndexes 对应sheet页开始行索引
     * @param clazz           pojo类型
     * @param <T>             泛型对象
     * @return 对象集合 List<List<T>>
     */
    private static <T> List<List<T>> parseSheetToList(Sheet[] sheets, int[] startRowIndexes, Class<T> clazz) {
        return IntStream.range(0, sheets.length).mapToObj(sheetIndex -> {
            List<T> res = Lists.newArrayList();
            Sheet st;
            Row row;
            Cell cell;
            try {
                st = sheets[sheetIndex];
                // 实体类的属性
                List<Field> fields = Arrays.stream(clazz.getDeclaredFields())
                        .filter(field -> field.isAnnotationPresent(SheetColumn.class)
                                && field.getAnnotation(SheetColumn.class).imported())
                        .collect(Collectors.toList());
                // 所有属性注解中列索引index最大值
                int maxIndex = fields.stream().mapToInt(item
                        -> item.getAnnotation(SheetColumn.class).index()).max().getAsInt();
                for (int i = startRowIndexes[sheetIndex]; i < st.getPhysicalNumberOfRows(); i++) {
                    // 空行跳过
                    if (Objects.isNull(row = st.getRow(i))) {
                        continue;
                    }
                    T object = clazz.newInstance();
                    int currentIndex = maxIndex;
                    for (Field field : fields) {
                        // 通过属性注解SheetColumn的列索引index取单元格的值
                        // 未指定index，即index为-1，则按照所有索引中最大索引+1取值，最大索引+1即为新的最大索引
                        SheetColumn sheetColumn = field.getAnnotation(SheetColumn.class);
                        int index = sheetColumn.index();
                        cell = index == -1 ? row.getCell(++currentIndex) : row.getCell(index);
                        String value = getCellValue(cell);
                        field.setAccessible(true);
                        Object val = TypeUtils.cast(value, field.getType(), null);
                        field.set(object, val);
                    }
                    res.add(object);
                }
            } catch (Exception e) {
                LOGGER.error("error.", e);
                throw new ExcelHandleException("excel parse unknown error.");
            }
            return res;
        }).collect(Collectors.toList());
    }

    /**
     * 获取单元格转成字符串后的值
     *
     * @param cell 单元格
     * @return 单元格的值
     */
    private static String getCellValue(Cell cell) {
        if (Objects.isNull(cell)) {
            return null;
        }
        cell.setCellType(CellType.STRING);
        return cell.getStringCellValue();
    }

    /**
     * 将对象集合写入Excel-单sheet页
     *
     * @param data  对象数据
     * @param clazz pojo类型
     * @param <T>   泛型对象
     * @return 工作簿
     */
    public static <T> void singleListToWorkbook(List<T> data, @NonNull OutputStream outputStream, @NonNull Class<T> clazz) {
        listToWorkbook(Arrays.asList(data), outputStream, clazz);
    }

    /**
     * 将对象集合写入Excel工作簿-单sheet页
     *
     * @param data  对象数据
     * @param clazz pojo类型
     * @param <T>   泛型对象
     * @return 工作簿
     */
    public static <T> Workbook singleListToWorkbook(List<T> data, @NonNull Class<T> clazz) {
        return listToWorkbook(Arrays.asList(data), clazz);
    }

    /**
     * 将对象集合列表写入Excel-多sheet页
     *
     * @param dataList     对象数据集合
     * @param clazz        pojo类型
     * @param outputStream 输出流
     * @param <T>          泛型对象
     */
    public static <T> void listToWorkbook(List<List<T>> dataList, @NonNull OutputStream outputStream, @NonNull Class<T> clazz) {
        Optional.ofNullable(listToWorkbook(dataList, clazz)).ifPresent(workbook -> {
            try {
                workbook.write(outputStream);
            } catch (IOException e) {
                LOGGER.error("error.", e);
                throw new ExcelHandleException("excel write unknown error.");
            }
        });
    }

    /**
     * 将对象集合列表写入Excel工作簿-多sheet页
     *
     * @param dataList 对象数据集合
     * @param clazz    pojo类型
     * @param <T>      泛型对象
     * @return 工作簿
     */
    public static <T> Workbook listToWorkbook(List<List<T>> dataList, @NonNull Class<T> clazz) {
        EnableExcelExport enableExcelExport = clazz.getAnnotation(EnableExcelExport.class);
        if (Objects.isNull(enableExcelExport)) {
            throw new ExcelHandleException("class type not support excel export.");
        }

        // sheet页参数，一个索引对应一个sheet页
        String[] sheetNames = ArrayUtil.copyNewArr(enableExcelExport.sheetNames(), dataList.size());
        boolean[] includeHeaders = ArrayUtil.copyNewArr(enableExcelExport.isIncludeHeaders(), dataList.size());
        int[] startRowIndexes = ArrayUtil.copyNewArr(enableExcelExport.startRowIndexes(), dataList.size());
        short[] startColumnIndexes = ArrayUtil.copyNewArr(enableExcelExport.startColumnIndexes(), dataList.size());

        // 创建 xls 或者 xlsx 格式多工作簿，默认 xlsx 格式
        Workbook workbook = ExcelConstants.EXCEL_XLS.equals(enableExcelExport.fileType()) ? new HSSFWorkbook() : new XSSFWorkbook();

        Iterator<List<T>> iterator = dataList.iterator();
        int i = 0;
        int sheetIndex = 1;
        List<T> data;
        Sheet st;
        Row row;
        Cell cell;
        while (iterator.hasNext()) {
            try {
                String sheetName = sheetNames[i];
                boolean includeHeader = includeHeaders[i];
                int rowNum = startRowIndexes[i];
                short startColumn = startColumnIndexes[i];
                data = iterator.next();
                // 实体类的属性
                List<Field> fields = Arrays.stream(clazz.getDeclaredFields())
                        .filter(field -> field.isAnnotationPresent(SheetColumn.class)
                                && field.getAnnotation(SheetColumn.class).exported())
                        .collect(Collectors.toList());
                // 所有属性注解中列索引index最大值
                int maxIndex = fields.stream().mapToInt(item -> item.getAnnotation(SheetColumn.class).index()).max().getAsInt();
                st = workbook.createSheet(StringUtils.isEmpty(sheetName) ? ExcelConstants.SHEET_PREFIX + sheetIndex++ : sheetName);
                row = null;
                int currentIndex = maxIndex;
                if (includeHeader) {
                    row = st.createRow(rowNum++);
                }
                // 初始化表头和列宽
                for (Field field : fields) {
                    SheetColumn sheetColumn = field.getAnnotation(SheetColumn.class);
                    int index = sheetColumn.index();
                    int width = sheetColumn.width();
                    // 通过属性注解SheetColumn的列索引index创建单元格
                    // 未指定index，即index为-1，则按照所有索引中最大索引+1取值，最大索引+1即为新的最大索引
                    index = (index == -1 ? ++currentIndex : index) + startColumn;
                    // row已经被初始化说明需要创建表头
                    if (Objects.nonNull(row)) {
                        String name = sheetColumn.name();
                        boolean required = sheetColumn.required();
                        cell = row.createCell(index);
                        fillCell(workbook, cell, name, required);
                    }
                    // 列宽为-1为自适应列宽
                    if (width == -1) {
                        st.autoSizeColumn(index);
                    } else {
                        st.setColumnWidth(index, width * 2 * 256);
                    }
                }
                // 业务数据填充
                for (T datum : data) {
                    currentIndex = maxIndex;
                    row = st.createRow(rowNum++);
                    for (Field field : fields) {
                        SheetColumn sheetColumn = field.getAnnotation(SheetColumn.class);
                        int index = sheetColumn.index();
                        index = (index == -1 ? ++currentIndex : index) + startColumn;
                        cell = row.createCell(index);
                        field.setAccessible(true);
                        Object value = field.get(datum);
                        fillCell(workbook, cell, value, false);
                    }
                }
                i++;
            } catch (Exception e) {
                LOGGER.error("error.", e);
                throw new ExcelHandleException("excel write unknown error.");
            }
        }
        return workbook;
    }

    /**
     * 填充单元格数据
     *
     * @param workbook 工作簿
     * @param cell     单元格
     * @param value    填充值
     * @param required 是否必输
     */
    private static void fillCell(Workbook workbook, Cell cell, Object value, boolean required) {
        // 必输列表头填充-浅黄色
        if (required) {
            CellStyle yStyle = workbook.createCellStyle();
            yStyle.setFillForegroundColor(IndexedColors.LIGHT_YELLOW.index);
            yStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);
            cell.setCellStyle(yStyle);
        }
        if (Objects.isNull(value)) {
            return;
        }
        switch (value.getClass().getName()) {
            case "java.lang.String":
                cell.setCellValue((String) value);
                break;
            case "java.lang.Integer":
                cell.setCellValue((Integer) value);
                break;
            case "java.lang.Long":
                cell.setCellValue((Long) value);
                break;
            case "java.lang.Float":
                cell.setCellValue((Float) value);
                break;
            case "java.lang.Double":
                cell.setCellValue((Double) value);
                break;
            case "java.math.BigDecimal":
                cell.setCellValue(Double.valueOf(String.valueOf(value)));
                break;
            case "java.lang.Boolean":
                cell.setCellValue((Boolean) value);
                break;
            case "java.util.Date":
                cell.setCellValue((Date) value);
                break;
            case "java.util.GregorianCalendar":
                cell.setCellValue((Calendar) value);
                break;
            case "org.apache.poi.xssf.usermodel.XSSFRichTextString":
                cell.setCellValue((XSSFRichTextString) value);
                break;
            default:
                return;
        }
    }
}
