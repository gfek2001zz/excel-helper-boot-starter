package cn.com.quanyou.excel.exported.impl;

import cn.com.quanyou.excel.context.IExcelContext;
import cn.com.quanyou.excel.context.impl.SheetMeta;
import cn.com.quanyou.excel.dao.IExcelTaskDao;
import cn.com.quanyou.excel.entity.ExcelTypeEnum;
import cn.com.quanyou.excel.entity.TaskModel;
import cn.com.quanyou.excel.entity.TaskStatusEnum;
import cn.com.quanyou.excel.exception.ExcelException;
import cn.com.quanyou.excel.exported.IExcelWriteStream;
import com.github.tobato.fastdfs.domain.proto.storage.DownloadByteArray;
import com.github.tobato.fastdfs.service.FastFileStorageClient;
import lombok.extern.slf4j.Slf4j;
import org.apache.commons.lang3.StringUtils;
import org.apache.commons.lang3.time.DateFormatUtils;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.scheduling.annotation.Async;
import org.springframework.stereotype.Component;

import java.io.*;
import java.util.Date;
import java.util.List;

@Slf4j
@Component
public class ExcelWriteStream implements IExcelWriteStream {

    @Autowired
    private FastFileStorageClient fastFileStorageClient;

    @Autowired
    private IExcelTaskDao excelTaskDao;

    @Autowired
    private ExcelSheetHandler excelSheetHandler;

    @Autowired
    private ExcelFileUploadHandler fileUploadHandler;


    /**
     * 生成Excel
     *
     * @param context
     * @throws IOException
     */
    @Async
    @Override
    public void generateExcel(IExcelContext context) throws IOException {
        excelTaskDao.updateTaskStatus(TaskStatusEnum.PROCESSING, context.getTaskId());
        Workbook wb = getWorkbook(context);

        TaskModel taskModel = excelTaskDao.findExcelTaskByTaskId(context.getTaskId());
        try {
            if (wb != null) {
                List<SheetMeta> sheetMetas = context.getSheetMetas();
                sheetMetas.stream().forEach(excelSheetHandler.generateSheet(context, wb));
                String downloadUrl = fileUploadHandler.uploadExcelFile(context, wb);
                taskModel.setDownloadUrl(downloadUrl);
                if (StringUtils.isNotEmpty(context.getFileName())) {
                    taskModel.setFileName(context.getFileName().concat(DateFormatUtils.format(new Date(), "yyyyMMdd")));
                } else {
                    taskModel.setFileName(context.getReportType().concat(DateFormatUtils.format(new Date(), "yyyyMMdd")));
                }
                taskModel.setStatus(TaskStatusEnum.PROCESSED);
                taskModel.setEndTime(new Date());
            } else {
                context.setThrowable(new ExcelException("导出失败，未找到模板配置！"));
            }
        } catch (Exception excelEx) {
            context.setThrowable(new ExcelException(excelEx));
            log.error(excelEx.getMessage(), excelEx);
        } finally {
            if (context.getThrowable() != null) {
                StringWriter stringWriter = new StringWriter();
                context.getThrowable().printStackTrace(new PrintWriter(stringWriter));
                taskModel.setExecResults(stringWriter.getBuffer().toString());
                taskModel.setStatus(TaskStatusEnum.ERROR);
            }

            excelTaskDao.updateTaskInfo(taskModel);
        }
    }


    /**
     * 创建模板
     *
     * @param context
     * @return
     * @throws IOException
     */
    private Workbook getWorkbook(IExcelContext context) throws IOException {
        Workbook wb = null;

        if (StringUtils.isNotEmpty(context.getTemplatePath())) {
            String templatePath = context.getTemplatePath();
            String filePath = templatePath.substring(templatePath.lastIndexOf("group1/"), templatePath.lastIndexOf("group1/") + 7);

            DownloadByteArray callback = new DownloadByteArray();
            byte[] fileBytes = fastFileStorageClient.downloadFile("group1", filePath, callback);
            InputStream is = new ByteArrayInputStream(fileBytes);

            if (ExcelTypeEnum.XLS.equals(ExcelTypeEnum.valueOf(context.getExcelType()))) {
                wb = new HSSFWorkbook(is);
            } else if (ExcelTypeEnum.XLSX.equals(ExcelTypeEnum.valueOf(context.getExcelType()))) {
                wb = new XSSFWorkbook(is);
            }
        } else {
            if (ExcelTypeEnum.XLS.equals(ExcelTypeEnum.valueOf(context.getExcelType()))) {
                wb = new HSSFWorkbook();
            } else if (ExcelTypeEnum.XLSX.equals(ExcelTypeEnum.valueOf(context.getExcelType()))) {
                wb = new XSSFWorkbook();
            }
        }

        return wb;
    }
}
