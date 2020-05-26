package cn.com.quanyou.excel.exported;

import cn.com.quanyou.excel.context.IExcelContext;

import java.io.IOException;

public interface IExcelWriteStream {

    void generateExcel(IExcelContext context) throws IOException;
}
