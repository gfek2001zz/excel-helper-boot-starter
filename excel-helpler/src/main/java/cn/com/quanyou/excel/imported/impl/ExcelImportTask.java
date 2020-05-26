package cn.com.quanyou.excel.imported.impl;

import cn.com.quanyou.excel.context.IExcelContext;
import cn.com.quanyou.excel.entity.TaskModel;
import cn.com.quanyou.excel.entity.TypeEnum;
import cn.com.quanyou.excel.imported.IExcelImportTask;
import cn.com.quanyou.excel.imported.IExcelReadStream;
import cn.com.quanyou.excel.service.IExcelTaskService;
import cn.com.quanyou.excel.starter.properties.ExcelProperties;
import lombok.Builder;
import lombok.Data;
import lombok.extern.slf4j.Slf4j;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.context.ApplicationContext;

import java.io.ByteArrayOutputStream;
import java.util.Map;

/**
 * Excel导入任务
 *
 * @author 高强
 */
@Slf4j
@Data
@Builder
public class ExcelImportTask implements IExcelImportTask {

    @Autowired
    private IExcelTaskService taskService;

    @Autowired
    private IExcelReadStream excelReadStream;

    private final ApplicationContext applicationContext;
    private final ExcelProperties excelProperties;

    @Override
    public void startTask(String excelType, String currentUser, String fileName, ByteArrayOutputStream fileOutputStream) {
        this.startTask(excelType, currentUser, fileName, fileOutputStream, null);
    }

    @Override
    public void startTask(String excelType, String currentUser, String fileName, ByteArrayOutputStream fileOutputStream, Map<String, Object> params)  {
        try {
            TaskModel task = taskService.initTask(excelType, "2", currentUser);
            String configType = excelProperties.getConfigType();
            IExcelContext excelContext = taskService.generateContext(applicationContext, excelType, TypeEnum.IMPORT, configType, task, fileName, params);
            excelReadStream.readExcel(excelContext, fileOutputStream);
        } catch (Exception ex) {
            log.error("导出任务执行失败：{}", ex.getMessage(), ex);
        }

    }
}
