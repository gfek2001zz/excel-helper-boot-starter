package cn.com.quanyou.excel.context;

import cn.com.quanyou.excel.context.impl.ColumnMeta;
import org.apache.poi.ss.usermodel.Sheet;

import java.util.List;

public interface IExcelCellExpression {
    int setExpression(IExcelContext context, Sheet sheet, List<?> data, List<ColumnMeta> columnMetas, int rowIdx);
}
