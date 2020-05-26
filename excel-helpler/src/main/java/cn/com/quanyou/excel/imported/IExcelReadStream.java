package cn.com.quanyou.excel.imported;

import cn.com.quanyou.excel.context.IExcelContext;

import java.io.ByteArrayOutputStream;

public interface IExcelReadStream {
    void readExcel(IExcelContext context, ByteArrayOutputStream fileOutputStream) throws Exception;
}
