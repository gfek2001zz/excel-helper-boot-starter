package cn.com.quanyou.excel.entity;

import lombok.Data;

import java.io.Serializable;

@Data
public class ExcelCfgModel implements Serializable {
    private Long id;
    private String reportName;
    private String reportType;
    private String excelType;
    private String fileName;
    private Integer rowsNum;
    private String templatePath;
    private String exportPath;
}
