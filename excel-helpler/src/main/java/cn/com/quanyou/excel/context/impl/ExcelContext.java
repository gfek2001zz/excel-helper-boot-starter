package cn.com.quanyou.excel.context.impl;

import cn.com.quanyou.excel.context.IExcelContext;
import lombok.Builder;
import lombok.Data;
import org.apache.commons.collections4.CollectionUtils;
import org.apache.commons.collections4.MapUtils;
import org.springframework.context.ApplicationContext;

import java.io.InputStream;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

@Data
@Builder
public class ExcelContext implements IExcelContext {
    private Long taskId;
    private String reportType;
    private String fileName;
    private String excelType;
    private Integer batchRows;
    private String templatePath;
    private InputStream xmlFileStream;
    private ApplicationContext applicationContext;
    private Throwable throwable;
    private String exportPath;

    private Map<String, Object> paramMap;
    private List<SheetMeta> sheetMetas;

    private List<ErrorMeta> errorMetas;

    private SheetMeta currentSheet;

    private Map<String, String> currentSheetColumnMap;

    public List<ErrorMeta> getErrorMetas() {
        if (errorMetas == null) {
            errorMetas = new ArrayList<>();
        }

        return errorMetas;
    }

    @Override
    public String getFieldNameToColumnName(String fieldName, String defaultValue) {
        String columnName = defaultValue;
        if (MapUtils.isNotEmpty(currentSheetColumnMap)) {
            columnName = currentSheetColumnMap.get(fieldName);
        }

        return columnName;
    }
}
