package cn.com.quanyou.excel.exported;

import java.io.IOException;
import java.util.Map;

public interface IExcelExportTask {
    void startTask(String excelType, String currentUser);
    void startTask(String excelType, String currentUser, Map<String, Object> params) throws IOException;
}
