package cn.com.quanyou.excel.entity;

import lombok.Data;

import java.io.Serializable;

@Data
public class ExcelColumnCfgModel implements Serializable {
    private Long id;
    private Long excelSheetCfgId;
    private String columnName;
    private String fieldName;
    private String fieldType;
    private String columnFormat;
    private String validatorBean;
    private String defaultValue;
    private Integer columnSort;
}
