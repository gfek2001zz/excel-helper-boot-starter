package cn.com.quanyou.excel.context.impl;

import lombok.AllArgsConstructor;
import lombok.Builder;
import lombok.Data;
import lombok.NoArgsConstructor;

@Data
@Builder
@AllArgsConstructor
@NoArgsConstructor
public class ErrorMeta {

    private String fieldName;
    private String rowName;
    private String colName;
    private String message;
}
