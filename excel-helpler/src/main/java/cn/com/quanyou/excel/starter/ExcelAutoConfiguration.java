package cn.com.quanyou.excel.starter;

import cn.com.quanyou.excel.exported.impl.ExcelExportTask;
import cn.com.quanyou.excel.imported.impl.ExcelImportTask;
import cn.com.quanyou.excel.starter.properties.ExcelProperties;
import lombok.extern.slf4j.Slf4j;
import org.springframework.boot.context.properties.EnableConfigurationProperties;
import org.springframework.context.ApplicationContext;
import org.springframework.context.annotation.Bean;
import org.springframework.context.annotation.Configuration;


@Slf4j
@Configuration
@EnableConfigurationProperties({ExcelProperties.class})
public class ExcelAutoConfiguration {

    private ExcelProperties excelProperties;
    private ApplicationContext applicationContext;

    public ExcelAutoConfiguration(ExcelProperties properties, ApplicationContext context) {
        this.excelProperties = properties;
        this.applicationContext = context;
    }

    @Bean
    public ExcelImportTask getExcelImportTask() {
        ExcelImportTask.ExcelImportTaskBuilder excelImportTaskBuilder = ExcelImportTask.builder();
        excelImportTaskBuilder.excelProperties(excelProperties);
        excelImportTaskBuilder.applicationContext(applicationContext);

        return excelImportTaskBuilder.build();
    }

    @Bean
    public ExcelExportTask getExcelExportTask() {
        ExcelExportTask.ExcelExportTaskBuilder taskBuilder = ExcelExportTask.builder();
        taskBuilder.excelProperties(excelProperties);
        taskBuilder.applicationContext(applicationContext);

        return taskBuilder.build();
    }
}
