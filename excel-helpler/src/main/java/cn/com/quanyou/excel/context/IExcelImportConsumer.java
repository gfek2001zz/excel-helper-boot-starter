package cn.com.quanyou.excel.context;

import java.util.List;

public interface IExcelImportConsumer {
    void begin(IExcelContext context);

    void useBatchData(IExcelContext context, List<?> data);

    void fail(IExcelContext context);

    void end(IExcelContext context);
}
