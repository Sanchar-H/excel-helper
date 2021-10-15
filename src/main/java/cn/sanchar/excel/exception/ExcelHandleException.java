package cn.sanchar.excel.exception;

import java.io.ByteArrayOutputStream;
import java.io.PrintStream;
import java.util.HashMap;
import java.util.LinkedHashMap;
import java.util.Map;

/**
 * description: Excel 处理自定义异常类
 *
 * @author shencai.huang@hand-china.com
 * @date 2021/10/13 3:58 下午
 * lastUpdateBy: shencai.huang@hand-china.com
 * lastUpdateDate: 2021/10/13
 */
public class ExcelHandleException extends RuntimeException{
    private final transient Object[] parameters;
    private String code;

    public ExcelHandleException(String code, Object... parameters) {
        super(code);
        this.parameters = parameters;
        this.code = code;
    }

    public ExcelHandleException(String code, Throwable cause, Object... parameters) {
        super(code, cause);
        this.parameters = parameters;
        this.code = code;
    }

    public ExcelHandleException(String code, Throwable cause) {
        super(code, cause);
        this.code = code;
        this.parameters = new Object[0];
    }

    public ExcelHandleException(Throwable cause, Object... parameters) {
        super(cause);
        this.parameters = parameters;
    }

    public Object[] getParameters() {
        return this.parameters;
    }

    public String getCode() {
        return this.code;
    }

    public String getTrace() {
        ByteArrayOutputStream baos = new ByteArrayOutputStream();
        PrintStream ps = new PrintStream(baos);
        this.printStackTrace(ps);
        ps.flush();
        return new String(baos.toByteArray());
    }

    public Map<String, Object> toMap() {
        HashMap<String, Object> map = new LinkedHashMap();
        map.put("code", this.code);
        map.put("message", super.getMessage());
        return map;
    }
}
