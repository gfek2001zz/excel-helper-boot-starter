package cn.com.quanyou.excel.entity;

import lombok.Builder;
import lombok.Data;

import java.io.Serializable;

@Data
@Builder
public class ExcelSheetCfgModel implements Serializable {
    private Long id;
    private Long excelCfgId;
    private String sheetName;
    private Integer titleRow;
    private Integer dataRow;
    private String providerBean;
    private String consumerBean;
    private String voBean;
}
