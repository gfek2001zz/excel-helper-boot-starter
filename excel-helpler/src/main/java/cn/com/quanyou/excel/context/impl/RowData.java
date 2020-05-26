package cn.com.quanyou.excel.context.impl;

import lombok.Data;

import java.util.HashMap;
import java.util.Map;

@Data
public class RowData {
    private Map<String, Object> columnToData;
    private Map<String, String> columnIdxToColumn;
    private int RowIdx;

    public Map<String, Object> getColumnToData() {
        if (null == columnToData) {
            columnToData = new HashMap<>();
        }

        return columnToData;
    }

    public Map<String, String> getColumnIdxToColumn() {
        if (null == columnIdxToColumn) {
            columnIdxToColumn = new HashMap<>();
        }

        return columnIdxToColumn;
    }



}
