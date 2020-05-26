package cn.com.quanyou.excel.exported.impl;

import cn.com.quanyou.excel.context.IExcelContext;
import cn.com.quanyou.excel.entity.TaskModel;
import cn.com.quanyou.excel.entity.TypeEnum;
import cn.com.quanyou.excel.exported.IExcelExportTask;
import cn.com.quanyou.excel.exported.IExcelWriteStream;
import cn.com.quanyou.excel.service.IExcelTaskService;
import cn.com.quanyou.excel.starter.properties.ExcelProperties;
import lombok.Builder;
import lombok.Data;
import lombok.extern.slf4j.Slf4j;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.context.ApplicationContext;

import java.util.Map;

@Slf4j
@Data
@Builder
public class ExcelExportTask implements IExcelExportTask {
    private final ApplicationContext applicationContext;
    private final ExcelProperties excelProperties;

    @Autowired
    private IExcelTaskService taskService;

    @Autowired
    private IExcelWriteStream excelWriteStream;


    @Override
    public void startTask(String excelType, String currentUser) {
        this.startTask(excelType, currentUser, null);
    }

    /**
     * 执行导出任务
     *
     * @param excelType
     * @param currentUser
     * @param params
     */
    public void startTask(String excelType, String currentUser, Map<String, Object> params) {
        try {
            TaskModel task = taskService.initTask(excelType, "1", currentUser);
            String configType = excelProperties.getConfigType();
            IExcelContext excelContext = taskService.generateContext(applicationContext, excelType, TypeEnum.EXPORT, configType, task, params);
            excelWriteStream.generateExcel(excelContext);
        } catch (Exception e) {
            log.error("导出任务执行失败：{}", e.getMessage(), e);
        }
    }

}
