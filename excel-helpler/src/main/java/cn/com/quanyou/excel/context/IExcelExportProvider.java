package cn.com.quanyou.excel.context;
import cn.com.quanyou.excel.entity.Page;

public interface IExcelExportProvider {
    void begin(IExcelContext context);
    Page<?> batchData(Object obj, Integer page, Integer pageSize);
    void fail(IExcelContext context);
    void end(IExcelContext context);
}
