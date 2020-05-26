package cn.com.quanyou.excel.service.impl;

import cn.com.quanyou.excel.context.IExcelContext;
import cn.com.quanyou.excel.context.impl.ColumnMeta;
import cn.com.quanyou.excel.context.impl.ExcelContext;
import cn.com.quanyou.excel.context.impl.ExcelXmlParser;
import cn.com.quanyou.excel.context.impl.SheetMeta;
import cn.com.quanyou.excel.dao.IExcelTaskDao;
import cn.com.quanyou.excel.entity.*;
import cn.com.quanyou.excel.service.IExcelTaskService;
import org.apache.commons.lang3.StringUtils;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.context.ApplicationContext;
import org.springframework.core.io.ClassPathResource;
import org.springframework.stereotype.Service;

import java.io.IOException;
import java.io.InputStream;
import java.util.List;
import java.util.Map;
import java.util.function.Function;
import java.util.stream.Collectors;
import java.util.stream.Stream;

@Service
public class ExcelTaskServiceImpl implements IExcelTaskService {

    @Autowired
    private IExcelTaskDao excelTaskDao;

    /**
     * 初始化导出任务
     *
     * @param excelType
     * @param currentUser
     * @return
     */
    @Override
    public TaskModel initTask(String excelType, String operationType, String currentUser) {
        TaskModel.TaskModelBuilder taskModelBuilder = TaskModel.builder();
        taskModelBuilder.reportType(excelType);
        taskModelBuilder.operationType(operationType);
        taskModelBuilder.status(TaskStatusEnum.PENDING);
        taskModelBuilder.createUser(currentUser);
        taskModelBuilder.lastUpdateUser(currentUser);

        TaskModel taskModel = taskModelBuilder.build();
        excelTaskDao.initTask(taskModel);
        return taskModel;
    }

    @Override
    public IExcelContext generateContext(ApplicationContext context, String reportType, TypeEnum typeEnum, String configType, TaskModel task, Map<String, Object> paramMap) throws IOException {
        return generateContext(context, reportType, typeEnum, configType, task, null, paramMap);
    }

    /**
     * 构建上下文
     *
     * @param context
     * @param reportType
     * @param configType
     * @param task
     * @param paramMap
     * @return
     * @throws IOException
     */
    @Override
    public IExcelContext generateContext(ApplicationContext context, String reportType, TypeEnum typeEnum, String configType, TaskModel task, String fileName, Map<String, Object> paramMap)
            throws IOException {
        ExcelContext.ExcelContextBuilder contextBuilder = ExcelContext.builder();
        contextBuilder.reportType(reportType).paramMap(paramMap)
                .applicationContext(context).taskId(task.getTaskId()).fileName(fileName);

        if (ConfigTypeEnum.DATABASE.equals(ConfigTypeEnum.valueOf(configType))) {
            ExcelCfgModel excelCfgModel =
                    excelTaskDao.findExcelCfgByReportType(reportType, typeEnum.getValue());
            List<ExcelSheetCfgModel> sheetCfgModel =
                    excelTaskDao.findExcelSheetCfgByReportType(reportType, typeEnum.getValue());
            List<ExcelColumnCfgModel> columnCfgModel =
                    excelTaskDao.findExcelColumnCfgByReportType(reportType, typeEnum.getValue());
            Map<Long, List<ExcelColumnCfgModel>> dataMap = columnCfgModel.parallelStream()
                    .collect(Collectors.groupingBy(ExcelColumnCfgModel::getExcelSheetCfgId));

            if (StringUtils.isEmpty(fileName)) {
                contextBuilder.fileName(excelCfgModel.getFileName());
            }

            contextBuilder.batchRows(excelCfgModel.getRowsNum());
            contextBuilder.exportPath(excelCfgModel.getExportPath());
            contextBuilder.excelType(excelCfgModel.getExcelType());

            contextBuilder.sheetMetas(sheetCfgModel.parallelStream().map(buildSheetMeta(excelCfgModel, dataMap)).collect(Collectors.toList()));
        } else if (ConfigTypeEnum.XML.equals(ConfigTypeEnum.valueOf(configType))) {
            ClassPathResource resource = new ClassPathResource(String.format(typeEnum.getTemplate(), reportType));
            InputStream xmlFileStream = resource.getInputStream();
            contextBuilder.xmlFileStream(xmlFileStream);

            ExcelXmlParser parser = new ExcelXmlParser();
            List<SheetMeta> sheetMetas = parser.parse(contextBuilder);
            contextBuilder.sheetMetas(sheetMetas);
        }

        return contextBuilder.build();
    }


    /**
     * Sheet页配置
     *
     * @param excelCfgModel
     * @param dataMap
     * @return
     */
    private Function<ExcelSheetCfgModel, SheetMeta> buildSheetMeta(ExcelCfgModel excelCfgModel, Map<Long, List<ExcelColumnCfgModel>> dataMap) {

        return excelSheetCfgModel -> {
            SheetMeta.SheetMetaBuilder builder = SheetMeta.builder();

            builder.sheetName(excelSheetCfgModel.getSheetName());
            builder.titleRow(excelSheetCfgModel.getTitleRow());
            builder.dataRow(excelSheetCfgModel.getDataRow());
            if (excelCfgModel.getReportType().equals(TypeEnum.EXPORT.getValue())) {
                builder.providerBean(excelSheetCfgModel.getProviderBean());
            } else if (excelCfgModel.getReportType().equals(TypeEnum.IMPORT.getValue())) {
                builder.consumerBean(excelSheetCfgModel.getConsumerBean());
            }

            builder.voBean(excelSheetCfgModel.getVoBean());
            List<ExcelColumnCfgModel> columnCfgModel = dataMap.get(excelSheetCfgModel.getId());
            builder.columnMetas(columnCfgModel.stream().map(buildColumnMeta()).collect(Collectors.toList()));

            return builder.build();
        };
    }

    /**
     * Column配置
     *
     * @return
     */
    private Function<ExcelColumnCfgModel, ColumnMeta> buildColumnMeta() {
        return excelColumnCfgModel -> {
            ColumnMeta.ColumnMetaBuilder builder = ColumnMeta.builder();
            builder.displayName(excelColumnCfgModel.getColumnName());
            builder.fieldName(excelColumnCfgModel.getFieldName());
            builder.format(excelColumnCfgModel.getColumnFormat());
            builder.type(excelColumnCfgModel.getFieldType());
            builder.defaultValue(excelColumnCfgModel.getDefaultValue());
            if (null != excelColumnCfgModel.getValidatorBean()) {
                builder.validatorBean(Stream.of(excelColumnCfgModel.getValidatorBean().split(",")).collect(Collectors.toList()));
            }
            builder.colIdx(excelColumnCfgModel.getColumnSort());

            return builder.build();
        };
    }
}
