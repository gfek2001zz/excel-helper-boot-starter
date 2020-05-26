package cn.com.quanyou.excel.context.impl;

import lombok.Builder;
import lombok.Data;
import org.springframework.util.CollectionUtils;

import java.util.ArrayList;
import java.util.List;

@Data
@Builder
public class SheetMeta {
    private String sheetName;
    private int titleRow;
    private int dataRow;
    private String providerBean;
    private String consumerBean;
    private String expressionBean;
    private String voBean;

    private List<ColumnMeta> columnMetas;
    private List<RowData> rowDataList;

    public List<ColumnMeta> getColumnMetas() {
        if (CollectionUtils.isEmpty(columnMetas)) {
            return new ArrayList<>();
        }

        return columnMetas;
    }

    public List<RowData> getRowDataList() {
        if (CollectionUtils.isEmpty(rowDataList)) {
            return new ArrayList<>();
        }

        return rowDataList;
    }


}
