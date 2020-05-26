package cn.com.quanyou.excel.imported;

import java.io.ByteArrayOutputStream;
import java.io.IOException;
import java.util.Map;

public interface IExcelImportTask {
    void startTask(String excelType, String currentUser, String fileName, ByteArrayOutputStream fileOutputStream) throws IOException;
    void startTask(String excelType, String currentUser, String fileName, ByteArrayOutputStream fileOutputStream, Map<String, Object> params) throws IOException;

}
