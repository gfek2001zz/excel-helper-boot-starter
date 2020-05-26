package cn.com.quanyou.excel.exported.impl;

import cn.com.quanyou.excel.context.IExcelContext;
import cn.com.quanyou.excel.context.IExcelExportProvider;
import cn.com.quanyou.excel.context.impl.ColumnMeta;
import cn.com.quanyou.excel.context.impl.SheetMeta;
import cn.com.quanyou.excel.entity.CellStyleVO;
import cn.com.quanyou.excel.entity.Page;
import cn.com.quanyou.excel.exception.ExcelException;
import lombok.extern.slf4j.Slf4j;
import org.apache.commons.lang3.StringUtils;
import org.apache.poi.ss.usermodel.*;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.context.ApplicationContext;
import org.springframework.stereotype.Component;
import org.springframework.util.CollectionUtils;

import java.util.List;
import java.util.Map;
import java.util.function.Consumer;
import java.util.stream.Collectors;

@Slf4j
@Component
public class ExcelSheetHandler {

    @Autowired
    private ExcelColumnHandler excelColumnHandler;

    @Autowired
    private ExcelRowHandler excelRowHandler;

    /**
     * 生成Sheet页
     *
     * @param context
     * @param wb
     * @return
     */
    public Consumer<SheetMeta> generateSheet(IExcelContext context, Workbook wb) {
        return sheetMeta -> {
            List<ColumnMeta> columnMetas = sheetMeta.getColumnMetas();

            Map<String, CellStyle> cellStyleMap = columnMetas.stream().map(excelColumnHandler.createCellStyle(wb))
                    .collect(Collectors.toConcurrentMap(CellStyleVO::getFieldName, CellStyleVO::getCellStyle));

            Sheet sheet = wb.getSheet(sheetMeta.getSheetName());
            if (sheet == null) {
                sheet = wb.createSheet(sheetMeta.getSheetName());
            }

            //套用模板则不生成标题行
            if (StringUtils.isEmpty(context.getTemplatePath())) {
                excelRowHandler.createTitleRow(columnMetas, sheetMeta, sheet);
            }

            IExcelExportProvider provider = getExcelExportProvider(context, sheetMeta.getProviderBean());
            try {
                provider.begin(context);

                Page<?> firstData = provider.batchData(context.getParamMap(), 1, context.getBatchRows());
                int total = firstData.getTotal();
                if (total > 0) {
                    Integer totalPage = total / context.getBatchRows();
                    if (total % context.getBatchRows() > 0)
                        totalPage = totalPage + 1;

                    int rowIdx = 0;
                    rowIdx = excelRowHandler.createDataRow(firstData.getRecords(), sheetMeta.getDataRow(), columnMetas, sheet, rowIdx, cellStyleMap);

                    for (int page = 2; page <= totalPage; page++) {
                        Page<?> dataList = provider.batchData(context.getParamMap(), page, context.getBatchRows());
                        if (!CollectionUtils.isEmpty(firstData.getRecords())) {
                            rowIdx = excelRowHandler.createDataRow(dataList.getRecords(), sheetMeta.getDataRow(), columnMetas, sheet, rowIdx, cellStyleMap);
                        }
                    }
                } else {
                    context.setThrowable(new ExcelException("没有导出数据！"));
                    provider.fail(context);
                }


            } catch (Exception ex) {
                log.error("导出Excel发生错误，{}", ex.getMessage(), ex);

                context.setThrowable(new ExcelException(ex));
                provider.fail(context);
            }

            if (context.getThrowable() == null) {
                columnMetas.stream().forEach(excelColumnHandler.autoColumnLength(sheetMeta, sheet));
            }
            provider.end(context);

        };
    }


    /**
     * 获取Provider
     * @param context
     * @param providerBean
     * @return
     */
    private IExcelExportProvider getExcelExportProvider(IExcelContext context, String providerBean) {
        if (StringUtils.isEmpty(providerBean)) {
            providerBean = "DefaultExportProvider";
        }

        ApplicationContext applicationContext = context.getApplicationContext();
        return applicationContext.getBean(providerBean, IExcelExportProvider.class);
    }
}
