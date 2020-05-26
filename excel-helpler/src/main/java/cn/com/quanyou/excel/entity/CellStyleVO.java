package cn.com.quanyou.excel.entity;

import lombok.Builder;
import lombok.Data;
import org.apache.poi.ss.usermodel.CellStyle;

@Data
@Builder
public class CellStyleVO {

    private String fieldName;

    private CellStyle cellStyle;
}
