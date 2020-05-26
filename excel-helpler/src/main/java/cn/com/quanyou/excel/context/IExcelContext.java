package cn.com.quanyou.excel.context;

import cn.com.quanyou.excel.context.impl.ErrorMeta;
import cn.com.quanyou.excel.context.impl.SheetMeta;
import org.springframework.context.ApplicationContext;

import java.util.List;
import java.util.Map;

public interface IExcelContext {
    Long getTaskId();
    String getReportType();
    Map<String, Object> getParamMap();
    List<SheetMeta> getSheetMetas();
    String getFileName();
    String getTemplatePath();
    String getExcelType();
    Integer getBatchRows();

    ApplicationContext getApplicationContext();
    Throwable getThrowable();
    void setThrowable(Throwable throwable);
    List<ErrorMeta> getErrorMetas();

    SheetMeta getCurrentSheet();
    Map<String, String> getCurrentSheetColumnMap();
    String getFieldNameToColumnName(String fieldName, String defaultValue);
}
