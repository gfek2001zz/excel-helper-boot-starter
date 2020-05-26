package cn.com.quanyou.excel.context.impl;

import lombok.Builder;
import lombok.Data;
import org.springframework.util.CollectionUtils;

import java.util.ArrayList;
import java.util.List;

@Data
@Builder
public class ColumnMeta {
    private String fieldName;
    private String displayName;
    private String type;
    private Integer width;
    private String defaultValue;
    private String format;
    private List<String> validatorBean; //校验器
    private List<Integer> cellValueLength;
    private Integer colIdx;


    public List<Integer> getCellValueLength() {
        if (CollectionUtils.isEmpty(cellValueLength)) {
            cellValueLength = new ArrayList<>();
        }

        return cellValueLength;
    }


}
