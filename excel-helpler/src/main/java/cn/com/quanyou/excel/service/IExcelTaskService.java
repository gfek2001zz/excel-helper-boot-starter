package cn.com.quanyou.excel.service;

import cn.com.quanyou.excel.context.IExcelContext;
import cn.com.quanyou.excel.entity.TaskModel;
import cn.com.quanyou.excel.entity.TypeEnum;
import org.springframework.context.ApplicationContext;

import java.io.IOException;
import java.util.Map;

public interface IExcelTaskService {
    TaskModel initTask(String excelType, String operationType, String currentUser);

    IExcelContext generateContext(ApplicationContext context, String reportType, TypeEnum typeEnum, String configType, TaskModel task, Map<String, Object> paramMap) throws IOException;

    IExcelContext generateContext(ApplicationContext context, String reportType, TypeEnum typeEnum, String configType, TaskModel task, String fileName, Map<String, Object> paramMap) throws IOException;
}
